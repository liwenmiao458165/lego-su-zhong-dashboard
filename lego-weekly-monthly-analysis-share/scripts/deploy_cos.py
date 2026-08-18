#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
腾讯云 COS 部署脚本（固定链接主通道，2026-07-26 启用）
将 LOCAL_DIR 下静态文件推送到 COS 桶，对象级设为 public-read，
支持 SUBPATH 子目录。

⚠️ 关键事实（已实测坐实）：
  - COS *默认域名* *.myqcloud.com 在 GET 时强制加 Content-Disposition: attachment
    + x-cos-force-download: true → 浏览器直接下载，不渲染。
  - 只有「绑定自定义域名」才能逃脱该策略（自定义域名端点不附加 force-download 头）。
  - 因此本脚本在配置了自定义域名(.cos_custom_domain / COS_CUSTOM_DOMAIN)时，
    输出链接一律用 https://<自定义域名>/...，绝不输出默认 myqcloud 链接当成品。

环境变量：
  COS_SECRET_ID   (必需，或从 .cos_secret_id 读取) 腾讯云 SecretId
  COS_SECRET_KEY  (必需，或从 .cos_secret_key 读取) 腾讯云 SecretKey
  COS_BUCKET      (可选，或从 .cos_bucket 读取) 存储桶名；默认 suzhong-hk-1458473628
  COS_REGION      (可选，或从 .cos_region 读取) 地域；默认 ap-hongkong
  COS_CUSTOM_DOMAIN (可选，或从 .cos_custom_domain 读取) 已绑定的自定义域名，如 suzhong-dashboard.com
  LOCAL_DIR       (必需) 本地待上传目录
  SUBPATH         (可选) 远程子路径前缀，默认 ''（根）
  DELETE_STALE    (可选) true 时删除远程 SUBPATH 下本地不存在的文件
"""
import os
import sys
import mimetypes

try:
    from qcloud_cos import CosConfig, CosS3Client
    from qcloud_cos.cos_exception import CosServiceError, CosClientError
except ImportError:
    sys.stderr.write("❌ 缺少依赖 qcloud_cos，请先: pip install cos-python-sdk-v5\n")
    sys.exit(2)


def _read_secret(p):
    if os.path.exists(p):
        with open(p) as _f:
            return _f.read().strip()
    return ""


BASE = os.path.dirname(os.path.abspath(__file__))
COS_SECRET_ID = os.environ.get("COS_SECRET_ID", "").strip() or _read_secret(os.path.join(BASE, '.cos_secret_id'))
COS_SECRET_KEY = os.environ.get("COS_SECRET_KEY", "").strip() or _read_secret(os.path.join(BASE, '.cos_secret_key'))
COS_BUCKET = os.environ.get("COS_BUCKET", "").strip() or _read_secret(os.path.join(BASE, '.cos_bucket')) or "suzhong-hk-1458473628"
COS_REGION = os.environ.get("COS_REGION", "").strip() or _read_secret(os.path.join(BASE, '.cos_region')) or "ap-hongkong"
COS_CUSTOM_DOMAIN = os.environ.get("COS_CUSTOM_DOMAIN", "").strip() or _read_secret(os.path.join(BASE, '.cos_custom_domain'))
LOCAL_DIR = os.environ.get("LOCAL_DIR", "").strip()
SUBPATH = os.environ.get("SUBPATH", "").strip().strip("/")
DELETE_STALE = os.environ.get("DELETE_STALE", "false").strip().lower() in ("1", "true", "yes")
# 周月部署时跳过 index.html，避免覆盖日看板主 index.html（两个目录同名不同内容）
EXCLUDE_INDEX = os.environ.get("EXCLUDE_INDEX", "false").strip().lower() in ("1", "true", "yes")

if not (COS_SECRET_ID and COS_SECRET_KEY and COS_BUCKET and LOCAL_DIR):
    sys.stderr.write("❌ 缺少必需项：COS_SECRET_ID / COS_SECRET_KEY / COS_BUCKET / LOCAL_DIR\n")
    sys.exit(2)

if not os.path.isdir(LOCAL_DIR):
    sys.stderr.write(f"❌ LOCAL_DIR 不是目录: {LOCAL_DIR}\n")
    sys.exit(2)

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
    ".webp": "image/webp",
    ".woff": "font/woff",
    ".woff2": "font/woff2",
    ".ttf": "font/ttf",
    ".csv": "text/csv; charset=utf-8",
    ".txt": "text/plain; charset=utf-8",
    ".xml": "application/xml",
    ".pdf": "application/pdf",
}

# 不应上传的附加件
EXCLUDE_SUFFIX = (".docx", ".zip", ".tmp", ".part", ".ds_store")
EXCLUDE_NAMES = ("guide_",)  # 以 guide_ 开头的临时/辅助文件不传


def content_type_for(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in CONTENT_TYPES:
        return CONTENT_TYPES[ext]
    guessed, _ = mimetypes.guess_type(path)
    return guessed or "application/octet-stream"


def should_exclude(rel_path, fname):
    if fname.endswith(EXCLUDE_SUFFIX):
        return True
    if any(fname.startswith(p) for p in EXCLUDE_NAMES):
        return True
    if EXCLUDE_INDEX and fname.lower() == "index.html":
        return True
    return False


def main():
    config = CosConfig(Region=COS_REGION, SecretId=COS_SECRET_ID, SecretKey=COS_SECRET_KEY)
    client = CosS3Client(config)

    # 校验桶可访问
    try:
        client.head_bucket(Bucket=COS_BUCKET)
    except CosServiceError as e:
        if e.get_status_code() in (403, 404):
            sys.stderr.write(f"❌ 无法访问存储桶 {COS_BUCKET}（状态码 {e.get_status_code()}）："
                             f"请确认桶名/地域正确且密钥有访问权限\n")
            sys.exit(1)
        raise

    uploaded = []
    for root, dirs, files in os.walk(LOCAL_DIR):
        for fname in sorted(files):
            abs_path = os.path.join(root, fname)
            rel = os.path.relpath(abs_path, LOCAL_DIR)
            rel_norm = rel.replace(os.sep, "/")
            if should_exclude(rel_norm, fname):
                print(f"  ⏭ 跳过(排除规则): {rel_norm}")
                continue
            key = f"{SUBPATH}/{rel_norm}" if SUBPATH else rel_norm
            ct = content_type_for(fname)
            with open(abs_path, "rb") as f:
                body = f.read()
            try:
                client.put_object(
                    Bucket=COS_BUCKET,
                    Body=body,
                    Key=key,
                    ACL="public-read",
                    ContentType=ct,
                    ContentDisposition="inline",  # 显式 inline：自定义域名下浏览器直接渲染（双保险）
                    EnableMD5=True,
                )
                uploaded.append((key, len(body), ct))
                print(f"  ✅ 上传 {key} ({len(body)}B, {ct})")
            except (CosServiceError, CosClientError) as e:
                sys.stderr.write(f"  ❌ 上传失败 {key}: {e}\n")
                sys.exit(1)

    print(f"\n✅ 共上传 {len(uploaded)} 个文件到 COS 桶 {COS_BUCKET} (地域 {COS_REGION})"
          + (f" 子路径 /{SUBPATH}" if SUBPATH else " 根目录"))

    if DELETE_STALE:
        prefix = f"{SUBPATH}/" if SUBPATH else ""
        try:
            resp = client.list_objects(Bucket=COS_BUCKET, Prefix=prefix, MaxKeys=1000)
            remote = [c["Key"] for c in resp.get("Contents", [])]
            local_keys = {k for k, _, _ in uploaded}
            for rk in remote:
                if rk not in local_keys and (not prefix or rk.startswith(prefix)):
                    client.delete_object(Bucket=COS_BUCKET, Key=rk)
                    print(f"  🗑 删除陈旧对象: {rk}")
        except (CosServiceError, CosClientError) as e:
            sys.stderr.write(f"  ⚠️ 清理陈旧对象失败(可忽略): {e}\n")

    # ---- 输出最终链接 ----
    if COS_CUSTOM_DOMAIN:
        base = f"https://{COS_CUSTOM_DOMAIN}"
        print(f"\n🔗 自定义域名访问基址: {base}")
        print(f"   首页: {base}/index.html" if not SUBPATH else f"   子目录: {base}/{SUBPATH}/index.html")
        print("   ✅ 自定义域名端点不强制下载，浏览器直接渲染")
        print(f"FINAL_LINK={base}/index.html" if not SUBPATH else f"FINAL_LINK={base}/{SUBPATH}/index.html")
    else:
        print(f"\n⚠️ 未配置自定义域名(.cos_custom_domain)。默认域名会强制下载，不可作分享链接！")
        print(f"   默认端点(仅供自检, 会下载): https://{COS_BUCKET}.cos.{COS_REGION}.myqcloud.com/index.html")
        print(f"   请绑定自定义域名后重跑，链接自动切换为 https://<你的域名>/index.html")


if __name__ == "__main__":
    main()
