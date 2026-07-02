"""
WB 全链路图片元数据管理模块 (v1.0)
==================================
为 AI 生图 → 去背 → 贴图 → 上款 提供 UID/group_id 绑定能力，
不依赖文件名，可承受重命名、重名、跨目录移动。

核心机制：
  1. 每个图片文件旁边放置同名 .meta.json sidecar（如 DX0255_B.png.meta.json）。
  2. 每个 DX 文件夹维护 uid_map.json，汇总该 DX 下所有图片的 UID 关系。
  3. 所有下游脚本读取 sidecar/uid_map，而不是解析文件名。

约定：
  - uid: 全局唯一，如 UID_20250702_0001
  - group_id: 同一组原图共享，如 G_00001
  - stage: inbox | ai | rembg | sticker | bw | upload
  - role: B | W | BW | 黑B | 黑W | 黑BW | 白BW | 黑BW | ...
  - 同一逻辑图在不同 stage 保留同一 uid（通过 stage 区分），子 stage 用 parent_uid 指向父阶段。
"""

import json
import hashlib
import os
import time
import re
from pathlib import Path
from typing import Dict, List, Optional, Any

BASE_DIR = Path("D:/Semems WB")
PROJECTS_DIR = BASE_DIR / "02_PROJECTS"

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------

def compute_md5(path: str | Path) -> str:
    """计算文件 MD5"""
    h = hashlib.md5()
    p = Path(path)
    if not p.exists():
        return ""
    with open(p, "rb") as f:
        for chunk in iter(lambda: f.read(64 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _atomic_write_json(path: Path, data: dict):
    """原子写入 JSON：先写 .tmp 再 rename"""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    tmp.replace(path)


def _read_json(path: Path, default=None) -> dict:
    """安全读取 JSON"""
    if not path.exists():
        return default if default is not None else {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default if default is not None else {}


# ---------------------------------------------------------------------------
# Sidecar 操作（每张图片一个 .meta.json）
# ---------------------------------------------------------------------------

def meta_path(path: str | Path) -> Path:
    """返回图片 path 对应的 sidecar 路径"""
    p = Path(path)
    return p.with_suffix(p.suffix + ".meta.json")


def read_meta(path: str | Path) -> Optional[dict]:
    """读取图片 sidecar，不存在返回 None"""
    mp = meta_path(path)
    data = _read_json(mp)
    return data if data else None


def write_meta(path: str | Path, data: dict):
    """写入图片 sidecar"""
    mp = meta_path(path)
    mp.parent.mkdir(parents=True, exist_ok=True)
    _atomic_write_json(mp, data)


def ensure_meta(path: str | Path, **kwargs) -> dict:
    """读取或创建 sidecar，并用 kwargs 更新字段"""
    meta = read_meta(path) or {}
    meta.update(kwargs)
    if "md5" not in meta or not meta["md5"]:
        meta["md5"] = compute_md5(path)
    if "file" not in meta or not meta["file"]:
        meta["file"] = str(Path(path).name)
    write_meta(path, meta)
    return meta


def update_meta(path: str | Path, **kwargs) -> Optional[dict]:
    """更新 sidecar 字段，不存在则创建"""
    return ensure_meta(path, **kwargs)


# ---------------------------------------------------------------------------
# UID Map 操作（每个 DX 一个 uid_map.json）
# ---------------------------------------------------------------------------

def uid_map_path(dx_dir: str | Path) -> Path:
    """返回 DX 目录的 uid_map.json 路径"""
    return Path(dx_dir) / "uid_map.json"


def read_uid_map(dx_dir: str | Path) -> dict:
    """读取 DX 的 uid_map.json"""
    return _read_json(uid_map_path(dx_dir), default={"dx": "", "version": 1, "groups": {}, "images": {}})


def write_uid_map(dx_dir: str | Path, data: dict):
    """写入 DX 的 uid_map.json"""
    p = uid_map_path(dx_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    data.setdefault("version", 1)
    if not data.get("dx"):
        data["dx"] = Path(dx_dir).name
    _atomic_write_json(p, data)


def ensure_uid_map(dx_dir: str | Path) -> dict:
    """确保 uid_map.json 存在并返回"""
    data = read_uid_map(dx_dir)
    if not data.get("images"):
        data["images"] = {}
    if not data.get("groups"):
        data["groups"] = {}
    return data


def register_image_in_map(dx_dir: str | Path, uid: str, group_id: str, stage: str,
                          role: str, file_path: str, parent_uid: Optional[str] = None,
                          source_file: Optional[str] = None, **extra) -> dict:
    """在 uid_map.json 中注册/更新一张图片"""
    data = ensure_uid_map(dx_dir)
    dx = Path(dx_dir).name
    data["dx"] = dx

    # 组注册
    if group_id:
        data["groups"].setdefault(group_id, {
            "group_id": group_id,
            "created": time.strftime("%Y-%m-%d %H:%M:%S"),
            "images": []
        })
        if uid not in data["groups"][group_id]["images"]:
            data["groups"][group_id]["images"].append(uid)

    rel_path = file_path
    try:
        rel_path = str(Path(file_path).relative_to(Path(dx_dir)))
    except Exception:
        pass

    entry = data["images"].get(uid, {})
    entry.update({
        "uid": uid,
        "group_id": group_id,
        "stage": stage,
        "role": role,
        "file": rel_path,
        "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
    })
    if parent_uid:
        entry["parent_uid"] = parent_uid
    if source_file:
        entry["source_file"] = source_file
    entry.update(extra)

    # 计算 md5（如果文件存在）
    full_path = Path(dx_dir) / rel_path
    if full_path.exists():
        entry["md5"] = compute_md5(full_path)

    data["images"][uid] = entry
    write_uid_map(dx_dir, data)
    return entry


def find_in_map(dx_dir: str | Path, **filters) -> List[dict]:
    """按字段过滤 uid_map 中的图片"""
    data = read_uid_map(dx_dir)
    results = []
    for uid, entry in data.get("images", {}).items():
        if all(entry.get(k) == v for k, v in filters.items()):
            results.append(entry)
    return results


def resolve_uid(dx_dir: str | Path, uid: str) -> Optional[dict]:
    """根据 uid 查找 uid_map 中的条目"""
    data = read_uid_map(dx_dir)
    return data.get("images", {}).get(uid)


def find_children(dx_dir: str | Path, parent_uid: str, stage: Optional[str] = None) -> List[dict]:
    """查找某个 uid 在指定 stage 下的子图片"""
    data = read_uid_map(dx_dir)
    results = []
    for entry in data.get("images", {}).values():
        if entry.get("parent_uid") == parent_uid:
            if stage is None or entry.get("stage") == stage:
                results.append(entry)
    return results


# ---------------------------------------------------------------------------
# Stage 专用便捷函数
# ---------------------------------------------------------------------------

def register_inbox(inbox_path: str | Path, uid: str, group_id: str, role: str):
    """注册 INBOX 原图"""
    p = Path(inbox_path)
    dx = p.name
    update_meta(p, uid=uid, group_id=group_id, stage="inbox", role=role)
    # INBOX 没有 dx 文件夹，不写 uid_map


def register_ai(ai_path: str | Path, uid: str, group_id: str, role: str,
                parent_uid: Optional[str] = None, inbox_file: Optional[str] = None):
    """注册 AI 生成图"""
    p = Path(ai_path)
    dx_dir = p.parent.parent  # 02_PROJECTS/DXxxxx/01_AI/file.png
    update_meta(p, uid=uid, group_id=group_id, stage="ai", role=role,
                parent_uid=parent_uid, source_file=inbox_file)
    register_image_in_map(dx_dir, uid, group_id, "ai", role, str(p),
                          parent_uid=parent_uid, source_file=inbox_file)


def register_rembg(cut_path: str | Path, uid: str, group_id: str, role: str,
                   parent_uid: str, ai_file: str):
    """注册去背输出图（_cut.png）"""
    p = Path(cut_path)
    dx_dir = p.parent.parent
    update_meta(p, uid=uid, group_id=group_id, stage="rembg", role=role,
                parent_uid=parent_uid, source_file=ai_file)
    register_image_in_map(dx_dir, uid, group_id, "rembg", role, str(p),
                          parent_uid=parent_uid, source_file=ai_file)


def register_sticker(upload_path: str | Path, uid: str, group_id: str, role: str,
                     parent_uid: str, cut_file: str):
    """注册贴图成品"""
    p = Path(upload_path)
    dx_dir = p.parent.parent
    update_meta(p, uid=uid, group_id=group_id, stage="sticker", role=role,
                parent_uid=parent_uid, source_file=cut_file)
    register_image_in_map(dx_dir, uid, group_id, "sticker", role, str(p),
                          parent_uid=parent_uid, source_file=cut_file)


def register_bw(bw_path: str | Path, uid: str, group_id: str, role: str,
                source_uids: List[str], source_files: List[str]):
    """注册 BW 合成图"""
    p = Path(bw_path)
    dx_dir = p.parent.parent
    update_meta(p, uid=uid, group_id=group_id, stage="bw", role=role,
                source_uids=source_uids, source_files=source_files)
    register_image_in_map(dx_dir, uid, group_id, "bw", role, str(p),
                          source_uids=source_uids, source_files=source_files)


def register_upload(upload_path: str | Path, uid: str, group_id: str, role: str,
                    parent_uid: str, source_file: str):
    """注册上款最终图（兼容 register_sticker）"""
    register_sticker(upload_path, uid, group_id, role, parent_uid, source_file)


# ---------------------------------------------------------------------------
# 迁移：从现有 source_map.json + 文件名构建 uid_map
# ---------------------------------------------------------------------------

def _extract_role_from_name(filename: str) -> str:
    """从文件名推断 role"""
    stem = Path(filename).stem
    # 去背图：DXxxxx_B_cut.png → B；DXxxxx_黑B_cut.png → 黑B
    if stem.endswith("_cut"):
        stem = stem[:-4]
    # 贴图图：DXxxxx_B_白T.jpg → B；DXxxxx_黑B_黑T.jpg → 黑B
    if "_白T" in filename or "_黑T" in filename:
        m = re.search(r"_([黑]?[BW])_", filename)
        if m:
            return m.group(1)
    # BW 合成图
    if "_白BW" in filename:
        return "白BW"
    if "_黑BW" in filename:
        return "黑BW"
    if "_BW" in filename:
        return "BW"
    # AI 图：DXxxxx_B.png → B
    if "_" in stem:
        last = stem.split("_")[-1]
        if last in ("B", "W", "BW", "WB"):
            return last
        if last in ("黑B", "黑W", "黑BW"):
            return last
    return "?"


def migrate_dx(dx_dir: str | Path, fallback_group_prefix="G") -> dict:
    """扫描 DX 目录，基于现有文件和 source_map.json 构建 uid_map。
    用于旧项目迁移；已有 uid 的文件会保留。
    旧项目无法准确知道 B/W 配对关系，默认按 DX  folder 聚合为同一 group。
    """
    dx_dir = Path(dx_dir)
    dx = dx_dir.name
    data = ensure_uid_map(dx_dir)
    data["dx"] = dx

    smap = _read_json(dx_dir / "source_map.json", default={})
    src_by_file = {s.get("file", ""): s for s in smap.get("sources", [])}

    # 旧项目按 DX 聚合为同一个 group（比每文件一个 group 更符合展示需求）
    default_group_id = f"{fallback_group_prefix}_{dx}"

    # 扫描各 stage 目录
    stages = {
        "ai": dx_dir / "01_AI",
        "rembg": dx_dir / "02_REM_BG",
        "sticker": dx_dir / "03_UPLOAD",
    }

    for stage, dir_path in stages.items():
        if not dir_path.exists():
            continue
        for f in sorted(dir_path.iterdir()):
            if not f.is_file():
                continue
            if f.name.endswith(".meta.json"):
                continue
            if f.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
                continue

            # 已有 meta 则优先使用
            meta = read_meta(f)
            uid = meta.get("uid") if meta else None
            group_id = meta.get("group_id") if meta else None
            role = meta.get("role") if meta else None

            if not uid:
                # 从 source_map 找 src_id
                src_info = src_by_file.get(f.name, {})
                src_id = src_info.get("src_id", "")
                if src_id:
                    uid = f"UID_{src_id}"
                else:
                    # 按文件内容 MD5 生成稳定 UID
                    md5 = compute_md5(f)
                    uid = f"UID_MIGRATED_{md5[:16]}"

            if not group_id:
                group_id = default_group_id

            if not role:
                role = _extract_role_from_name(f.name)

            rel = str(f.relative_to(dx_dir))
            register_image_in_map(dx_dir, uid, group_id, stage, role, str(f))
            ensure_meta(f, uid=uid, group_id=group_id, stage=stage, role=role)

    return data


def migrate_all_projects():
    """迁移所有 DX 项目"""
    if not PROJECTS_DIR.exists():
        return
    for d in sorted(PROJECTS_DIR.iterdir()):
        if d.is_dir() and re.match(r"^DX\d+$", d.name):
            try:
                migrate_dx(d)
                print(f"[wb_meta] OK migrate: {d.name}")
            except Exception as e:
                print(f"[wb_meta] ERR migrate {d.name}: {e}")


# ---------------------------------------------------------------------------
# 调试 / CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "migrate":
        migrate_all_projects()
    elif len(sys.argv) > 2 and sys.argv[1] == "migrate-dx":
        migrate_dx(sys.argv[2])
    else:
        print("用法:")
        print("  python wb_meta.py migrate")
        print("  python wb_meta.py migrate-dx <DX_DIR>")
