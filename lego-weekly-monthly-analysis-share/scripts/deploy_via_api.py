#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通过 GitHub Contents API 部署静态文件到 GitHub Pages（gh-pages 分支）。
绕开被企业网络墙掉的 github.com:443 git 协议，改用 api.github.com（REST API，网络可达）。

环境变量：
  GH_TOKEN    GitHub Personal Access Token（需 Contents: Read and Write 权限）
  GH_OWNER    用户名
  LOCAL_DIR   本地待部署目录（绝对路径，必填）
  SUBPATH     远端子目录（可选，默认空=站点根目录）。如 analysis -> 部署到 /analysis
  REPO        仓库名（默认 lego-su-zhong-dashboard）

行为：镜像同步 LOCAL_DIR -> gh-pages/SUBPATH
  - 本地有、远程无 -> 新建
  - 本地有、远程有 -> 更新（带 sha）
  - 远程有、本地无 -> 删除（仅 SUBPATH 范围内）
  - 安全保护：根目录(SUBPATH 为空)部署时，绝不触碰其他顶层目录（如 /analysis），避免误删
  - 内置排除规则：跳过非主看板交付物（.DS_Store / .docx / .zip / guide_*.html / monthly_analysis* / weekly_analysis*）
"""
import os
import sys
import json
import base64
import time
import hashlib
import urllib.request
import urllib.error
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

API = "https://api.github.com"

# ---------- 配置（环境变量） ----------
TOKEN = os.environ.get("GH_TOKEN", "")
OWNER = os.environ.get("GH_OWNER", "")
REPO = os.environ.get("REPO", "lego-su-zhong-dashboard")
LOCAL_DIR = os.environ.get("LOCAL_DIR", "")
SUBPATH = os.environ.get("SUBPATH", "").strip("/")
DELETE_STALE = os.environ.get("DELETE_STALE", "false").strip().lower() in ("1", "true", "yes")

if not TOKEN:
    print("❌ 请设置环境变量 GH_TOKEN", file=sys.stderr); sys.exit(1)
if not OWNER:
    print("❌ 请设置环境变量 GH_OWNER", file=sys.stderr); sys.exit(1)
if not LOCAL_DIR or not os.path.isdir(LOCAL_DIR):
    print(f"❌ LOCAL_DIR 无效: {LOCAL_DIR!r}", file=sys.stderr); sys.exit(1)

# ---------- 排除规则（basename 匹配，大小写不敏感）----------
EXCLUDE_SUFFIX = (".ds_store", ".docx", ".zip")
EXCLUDE_PREFIX_ALWAYS = ("guide_",)                                  # 任何场景都排除
PROTECT_FILES = ("CNAME", "404.html", ".nojekyll")                   # 任何场景都不删除（自定义域名/配置）
# 周月部署时跳过 index.html，避免覆盖日看板主 index.html（两个目录同名不同内容）
EXCLUDE_INDEX = os.environ.get("EXCLUDE_INDEX", "false").strip().lower() in ("1", "true", "yes")

def is_excluded(name: str) -> bool:
    n = name.lower()
    if n.endswith(EXCLUDE_SUFFIX):
        return True
    if n.startswith(EXCLUDE_PREFIX_ALWAYS):
        return True
    if EXCLUDE_INDEX and n == "index.html":
        return True
    return False

# 根目录部署时，保护这些顶层目录不被触碰（不递归、不删除）
PROTECT_TOP_DIRS = ("analysis",) if not SUBPATH else ()

# ---------- HTTP 工具 ----------
def _req(method, path, data=None, retries=4):
    url = f"{API}{path}"
    for attempt in range(1, retries + 1):
        try:
            body = json.dumps(data).encode("utf-8") if data is not None else None
            req = urllib.request.Request(url, data=body, method=method)
            req.add_header("Authorization", f"Bearer {TOKEN}")
            req.add_header("Accept", "application/vnd.github+json")
            req.add_header("User-Agent", "lego-deploy")
            if body:
                req.add_header("Content-Type", "application/json")
            with urllib.request.urlopen(req, timeout=120) as resp:
                raw = resp.read().decode("utf-8", "replace")
                return resp.status, (json.loads(raw) if raw else {})
        except urllib.error.HTTPError as e:
            try:
                txt = e.read().decode("utf-8", "replace")
            except Exception:
                txt = ""
            if e.code in (400, 403, 409, 422, 429):  # 偶发 Bad Request / 限流(403,429) / 冲突（并发）/ 校验失败，重试
                print(f"   ⚠️ {method} {path} -> HTTP {e.code}，重试 {attempt}/{retries}")
                time.sleep(2 * attempt)
                continue
            print(f"   ❌ {method} {path} -> HTTP {e.code}: {txt[:200]}", file=sys.stderr)
            return e.code, {}
        except Exception as e:
            if attempt < retries:
                print(f"   ⚠️ {method} {path} 网络异常({e})，重试 {attempt}/{retries}")
                time.sleep(3 * attempt)
                continue
            print(f"   ❌ {method} {path} 失败: {e}", file=sys.stderr)
            return 0, {}
    return 0, {}

def api_put(path, content_b64, sha=None):
    data = {"message": f"deploy {SUBPATH or 'root'} via api",
            "content": content_b64,
            "branch": "gh-pages"}
    if sha:
        data["sha"] = sha
    return _req("PUT", path, data)

def api_delete(path, sha):
    data = {"message": f"remove stale {path}",
            "sha": sha,
            "branch": "gh-pages"}
    return _req("DELETE", path, data)

def git_blob_sha(data: bytes) -> str:
    """计算与 GitHub Contents API 一致的 git blob SHA（用于跳过未变文件）。"""
    return hashlib.sha1(b"blob " + str(len(data)).encode() + b"\x00" + data).hexdigest()

# ---------- 远程文件树（递归，但跳过受保护的顶层目录）----------
def remote_tree(prefix, skip_dirs=()):
    """返回 {path: sha} 所有文件（不含 skip_dirs 内）"""
    tree = {}
    st, js = _req("GET", f"/repos/{OWNER}/{REPO}/contents/{prefix}?ref=gh-pages")
    if st == 404:
        return tree
    if st != 200 or not isinstance(js, list):
        if isinstance(js, dict) and js.get("type") == "file":
            tree[prefix] = js.get("sha")
        return tree
    for it in js:
        if it.get("type") == "dir":
            if it.get("name") in skip_dirs:
                continue  # 受保护目录：不递归、不列文件
            tree.update(remote_tree(it["path"], skip_dirs))
        else:
            tree[it["path"]] = it.get("sha")
    return tree

# ---------- 主流程 ----------
def main():
    print(f"GitHub 账号: {OWNER} | 仓库: {REPO} | 子目录: /{SUBPATH}")
    print(f"本地目录: {LOCAL_DIR}")
    if PROTECT_TOP_DIRS:
        print(f"🛡️ 保护顶层目录（不触碰）: {', '.join(PROTECT_TOP_DIRS)}")

    # 1. 扫描本地文件
    local = {}  # remote_path -> local_abspath
    for root, dirs, files in os.walk(LOCAL_DIR):
        for fn in files:
            if is_excluded(fn):
                continue
            ab = os.path.join(root, fn)
            rel = os.path.relpath(ab, LOCAL_DIR).replace(os.sep, "/")
            rpath = f"{SUBPATH}/{rel}" if SUBPATH else rel
            local[rpath] = ab

    # 2. 拉取远程树（仅 SUBPATH 范围，跳过受保护目录）
    base = SUBPATH if SUBPATH else ""
    remote = remote_tree(base, skip_dirs=PROTECT_TOP_DIRS) if base == "" else remote_tree(base)

    # 3. 计算本地 blob sha，跳过未变文件，并行上传变更项（提速核心）
    #    - 未变文件（本地 blob sha == 远程 sha）直接跳过，省掉一次网络往返
    #    - 新增/变更文件用线程池并发 PUT，默认 8 路（DEPLOY_WORKERS 可调）
    workers = int(os.environ.get("DEPLOY_WORKERS", "8"))
    changed = []   # (rpath, b64, remote_sha_or_None)
    skip_count = 0
    for rpath, ab in local.items():
        with open(ab, "rb") as f:
            raw = f.read()
        rsha = remote.get(rpath)
        if rsha and rsha == git_blob_sha(raw):
            skip_count += 1
            continue
        changed.append((rpath, base64.b64encode(raw).decode("ascii"), rsha))

    def _do_put(item):
        rpath, b64, sha = item
        st, _ = api_put(
            f"/repos/{OWNER}/{REPO}/contents/{urllib.parse.quote(rpath, safe='/')}", b64, sha)
        return rpath, st

    put_count = 0
    fail_count = 0
    if changed:
        print(f"   🚀 并发上传 {len(changed)} 个变更文件（{workers} 路）...")
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = [ex.submit(_do_put, it) for it in changed]
            for fut in as_completed(futs):
                rpath, st = fut.result()
                if st in (200, 201):
                    put_count += 1
                    print(f"   ✅ PUT {rpath}")
                else:
                    fail_count += 1
                    print(f"   ❌ PUT {rpath} 失败 HTTP {st}")
    else:
        print("   ✨ 无变更文件，跳过上传")
    if skip_count:
        print(f"   ⏭️ 跳过未变文件: {skip_count} 个")

    # 4. 删除远程多余（仅 DELETE_STALE=true 时；日/周脚本默认共存不删，避免互删）
    del_count = 0
    if DELETE_STALE:
        for rpath, sha in remote.items():
            if rpath not in local:
                if os.path.basename(rpath) in PROTECT_FILES:
                    print(f"   🛡️ 保留受保护文件(不删): {rpath}")
                    continue
                st, _ = api_delete(f"/repos/{OWNER}/{REPO}/contents/{urllib.parse.quote(rpath, safe='/')}", sha)
                if st in (200, 204):
                    del_count += 1
                    print(f"   🗑️ DEL {rpath}")
                else:
                    print(f"   ⚠️ DEL {rpath} 失败 HTTP {st}")
    else:
        print("   ℹ️ 默认不删除远程多余文件（日/周脚本共存，避免互删；如需镜像清理请设 DELETE_STALE=true）")

    print(f"\n✅ 已同步：新增/更新 {put_count} 个，删除 {del_count} 个")

    # 5. 确保 Pages 已开启
    _ensure_pages()

    url = f"https://{OWNER}.github.io/{REPO}/"
    if SUBPATH:
        url = f"{url}{SUBPATH}/"
    print(f"✅ 固定链接: {url}")
    print("（内容每次更新，地址不变）")

def _ensure_pages():
    payload = json.dumps({"source": {"branch": "gh-pages", "path": "/"}}).encode()
    for method in ("POST", "PUT"):
        req = urllib.request.Request(
            f"{API}/repos/{OWNER}/{REPO}/pages", data=payload, method=method)
        req.add_header("Authorization", f"Bearer {TOKEN}")
        req.add_header("Content-Type", "application/json")
        req.add_header("User-Agent", "lego-deploy")
        try:
            urllib.request.urlopen(req, timeout=120)
            return
        except Exception:
            continue

if __name__ == "__main__":
    main()
