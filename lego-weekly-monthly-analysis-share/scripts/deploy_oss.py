#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通过阿里云 OSS REST API 部署静态文件到 OSS bucket（静态网站托管）。
绕开被企业网络墙掉的 GitHub Pages(github.io)/Gitee Pages(gitee.io) CDN 域名——
OSS 默认 bucket 访问域名 <bucket>.<region>.aliyuncs.com 在企业网可访问（已验证 403=网络通）。

环境变量：
  OSS_AK        阿里云 AccessKeyId（需该 bucket 的 PutObject 权限）
  OSS_SK        阿里云 AccessKeySecret
  OSS_BUCKET    bucket 名称
  OSS_REGION    区域，默认 oss-cn-hangzhou
  OSS_ENDPOINT  完整 endpoint（可选，默认 {REGION}.aliyuncs.com）
  LOCAL_DIR     本地待部署目录（必填）
  SUBPATH       远端子目录（可选，默认空=根目录）。如 analysis -> /analysis
  DELETE_STALE  是否删除远程 SUBPATH 内多余文件（可选，默认 1=是）

签名：OSS Signature V1 (HMAC-SHA1)
依赖：仅 Python 标准库
"""
import os
import sys
import json
import base64
import time
import datetime
import hmac
import hashlib
import urllib.request
import urllib.error

# ---------- 配置（环境变量） ----------
AK = os.environ.get("OSS_AK", "")
SK = os.environ.get("OSS_SK", "")
BUCKET = os.environ.get("OSS_BUCKET", "")
REGION = os.environ.get("OSS_REGION", "oss-cn-hangzhou")
ENDPOINT = os.environ.get("OSS_ENDPOINT", f"{REGION}.aliyuncs.com")
LOCAL_DIR = os.environ.get("LOCAL_DIR", "")
SUBPATH = os.environ.get("SUBPATH", "").strip("/")
DELETE_STALE = os.environ.get("DELETE_STALE", "1") == "1"

if not (AK and SK and BUCKET):
    print("❌ 请设置环境变量 OSS_AK / OSS_SK / OSS_BUCKET", file=sys.stderr); sys.exit(1)
if not LOCAL_DIR or not os.path.isdir(LOCAL_DIR):
    print(f"❌ LOCAL_DIR 无效: {LOCAL_DIR!r}", file=sys.stderr); sys.exit(1)

# 根目录部署时保护其他顶层目录（如 analysis 不被根目录部署误删）
PROTECT_TOP_DIRS = ("analysis",) if not SUBPATH else ()

CONTENT_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".htm": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
    ".txt": "text/plain; charset=utf-8",
}
EXCLUDE_SUFFIX = (".ds_store", ".docx", ".zip")
EXCLUDE_PREFIX_ALWAYS = ("guide_",)
EXCLUDE_PREFIX_ROOTONLY = ()

def is_excluded(name):
    n = name.lower()
    if n.endswith(EXCLUDE_SUFFIX):
        return True
    if n.startswith(EXCLUDE_PREFIX_ALWAYS):
        return True
    if not SUBPATH and n.startswith(EXCLUDE_PREFIX_ROOTONLY):
        return True
    return False

def sign(verb, content_type, date_str, resource, oss_headers=None):
    # OSS Signature V1
    # StringToSign = VERB + "\n" + MD5 + "\n" + TYPE + "\n" + DATE + "\n" + CanonicalizedOSSHeaders + Resource
    md5 = ""
    oss_headers = oss_headers or {}
    canon = ""
    for k in sorted(oss_headers.keys()):
        canon += f"{k.lower()}:{oss_headers[k]}\n"
    string_to_sign = f"{verb}\n{md5}\n{content_type}\n{date_str}\n{canon}{resource}"
    sig = base64.b64encode(
        hmac.new(SK.encode("utf-8"), string_to_sign.encode("utf-8"), hashlib.sha1).digest()
    ).decode("ascii")
    return sig

def _req(verb, obj, body=None, content_type="application/octet-stream", retries=4, oss_headers=None, content_disposition=None):
    url = f"https://{BUCKET}.{ENDPOINT}/{obj}"
    date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%a, %d %b %Y %H:%M:%S GMT")
    resource = f"/{BUCKET}/{obj}"
    oss_headers = oss_headers or {}
    sig = sign(verb, content_type, date_str, resource, oss_headers)
    for attempt in range(1, retries + 1):
        try:
            req = urllib.request.Request(url, data=body, method=verb)
            req.add_header("Date", date_str)
            req.add_header("Authorization", f"OSS {AK}:{sig}")
            req.add_header("Content-Type", content_type)
            for k, v in oss_headers.items():
                req.add_header(k, v)
            if content_disposition:
                # 显式 inline，覆盖 OSS 可能的强制下载(Content-Disposition: attachment / x-oss-force-download)
                req.add_header("Content-Disposition", content_disposition)
            if body is not None:
                req.add_header("Content-Length", str(len(body)))
            with urllib.request.urlopen(req, timeout=40) as resp:
                return resp.status, resp.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            if e.code in (403,):
                # 403 可能是签名错或权限错，重试一次后报错
                if attempt == 1:
                    print(f"   ⚠️ {verb} {obj} -> 403，检查 AK/SK/权限", file=sys.stderr)
                return e.code, e.read().decode("utf-8", "replace")
            print(f"   ❌ {verb} {obj} -> HTTP {e.code}", file=sys.stderr)
            return e.code, ""
        except Exception as e:
            if attempt < retries:
                print(f"   ⚠️ {verb} {obj} 网络异常({e})，重试 {attempt}/{retries}")
                time.sleep(3 * attempt)
                continue
            print(f"   ❌ {verb} {obj} 失败: {e}", file=sys.stderr)
            return 0, ""
    return 0, ""

def list_objects(prefix):
    """列出 bucket 中 prefix 下的对象 key（含子目录，跳过受保护顶层目录）"""
    keys = []
    marker = ""
    while True:
        url = f"https://{BUCKET}.{ENDPOINT}/?prefix={prefix}&marker={marker}&max-keys=1000"
        date_str = datetime.datetime.now(datetime.timezone.utc).strftime("%a, %d %b %Y %H:%M:%S GMT")
        resource = f"/{BUCKET}/"
        sig = sign("GET", "", date_str, resource)
        req = urllib.request.Request(url, method="GET")
        req.add_header("Date", date_str)
        req.add_header("Authorization", f"OSS {AK}:{sig}")
        try:
            with urllib.request.urlopen(req, timeout=40) as resp:
                xml = resp.read().decode("utf-8", "replace")
        except Exception as e:
            print(f"   ⚠️ 列举对象失败: {e}", file=sys.stderr)
            break
        import xml.etree.ElementTree as ET
        try:
            root = ET.fromstring(xml)
        except Exception:
            break
        ns = "{http://doc.oss-cn-hangzhou.aliyuncs.com/xmlns/}"
        for c in root.findall(f"{ns}Contents"):
            k = c.find(f"{ns}Key").text
            keys.append(k)
        nxt = root.find(f"{ns}NextMarker")
        if nxt is not None and nxt.text:
            marker = nxt.text
        else:
            break
    return keys

def main():
    print(f"OSS Bucket: {BUCKET} | Endpoint: {ENDPOINT} | 子目录: /{SUBPATH}")
    if PROTECT_TOP_DIRS:
        print(f"🛡️ 保护顶层目录（不触碰）: {', '.join(PROTECT_TOP_DIRS)}")

    # 1. 本地文件
    local = {}
    for root, dirs, files in os.walk(LOCAL_DIR):
        for fn in files:
            if is_excluded(fn):
                continue
            ab = os.path.join(root, fn)
            rel = os.path.relpath(ab, LOCAL_DIR).replace(os.sep, "/")
            obj = f"{SUBPATH}/{rel}" if SUBPATH else rel
            local[obj] = ab

    # 2. 远程已有（仅 SUBPATH 范围，跳过受保护目录）
    prefix = f"{SUBPATH}/" if SUBPATH else ""
    remote_all = list_objects(prefix) if prefix else list_objects("")
    remote = []
    for k in remote_all:
        top = k.split("/")[0]
        if top in PROTECT_TOP_DIRS and not SUBPATH:
            continue
        if k in local or k.startswith(prefix):
            remote.append(k)
    # 仅保留 SUBPATH 范围内的 key
    remote = [k for k in remote if (k.startswith(prefix) if prefix else True)]

    # 3. 上传/更新
    put = 0
    for obj, ab in local.items():
        ext = os.path.splitext(ab)[1].lower()
        ct = CONTENT_TYPES.get(ext, "application/octet-stream")
        with open(ab, "rb") as f:
            body = f.read()
        st, _ = _req("PUT", obj, body, ct, oss_headers={"x-oss-object-acl": "public-read"}, content_disposition="inline")
        if 200 <= st < 300:
            put += 1
            print(f"   ✅ PUT {obj}")
        else:
            print(f"   ❌ PUT {obj} 失败 HTTP {st}")

    # 4. 删除多余
    if DELETE_STALE:
        deln = 0
        for k in remote:
            if k not in local:
                st, _ = _req("DELETE", k)
                if 200 <= st < 300:
                    deln += 1
                    print(f"   🗑️ DEL {k}")
                else:
                    print(f"   ⚠️ DEL {k} 失败 HTTP {st}")
        print(f"\n✅ 已同步：上传/更新 {put} 个，删除 {deln} 个")
    else:
        print(f"\n✅ 已上传/更新 {put} 个（未删除远程多余文件）")

    url = f"https://{BUCKET}.{ENDPOINT}/"
    if SUBPATH:
        url = f"{url}{SUBPATH}/"
    print(f"✅ 访问地址: {url}")

if __name__ == "__main__":
    main()
