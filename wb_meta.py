"""
WB 全链路图片元数据管理模块 (v2.0)
==================================
为 AI 生图 → 去背 → 贴图 → 上款 提供 UID/group_id 绑定能力。

核心升级（v2.0）：
  - 以 **MD5 为主键**，文件路径为辅助。
  - 图片改名、移动、复制后，只要内容没变，仍可通过 MD5 找到对应元数据。
  - sidecar 改为按 UID 命名，不再依赖原文件名：
      05_META/DXxxxx/sidecars/UID_xxx.meta.json
  - uid_map.json 维护 md5_index: {md5: uid}

核心机制：
  1. 所有元数据统一放在 D:/Semems WB/05_META/ 下。
  2. 每个 DX 的 uid_map.json 放在 05_META/DXxxxx/uid_map.json。
  3. sidecar 放在 05_META/DXxxxx/sidecars/UID_xxx.meta.json。
  4. 01_AI / 02_REM_BG / 03_UPLOAD 只放图片，不放文档。

约定：
  - uid: 全局唯一，如 UID_20250702_0001
  - group_id: 同一组原图共享，如 G_00001
  - stage: inbox | ai | rembg | sticker | bw | upload
  - role: B | W | BW | 黑B | 黑W | 黑BW | 白BW | 黑BW | ...
"""

import json
import hashlib
import os
import time
import re
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Any

BASE_DIR = Path("D:/Semems WB")
PROJECTS_DIR = BASE_DIR / "02_PROJECTS"
META_DIR = BASE_DIR / "05_META"

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


def _dx_from_path(path: str | Path) -> Optional[str]:
    """从图片路径推断 DX 名（如 02_PROJECTS/DX0274/01_AI/xxx.png -> DX0274）"""
    p = Path(path).resolve()
    try:
        rel = p.relative_to(PROJECTS_DIR)
        parts = rel.parts
        if parts and re.match(r"^DX\d+$", parts[0]):
            return parts[0]
    except ValueError:
        pass
    return None


# ---------------------------------------------------------------------------
# Sidecar 操作（按 UID 命名，MD5 为主键）
# ---------------------------------------------------------------------------

def sidecar_dir(dx_dir: str | Path) -> Path:
    """返回某 DX 的 sidecars 目录"""
    dx = Path(dx_dir).name
    return META_DIR / dx / "sidecars"


def sidecar_path_by_uid(dx_dir: str | Path, uid: str) -> Path:
    """返回 UID 对应的 sidecar 路径"""
    return sidecar_dir(dx_dir) / f"{uid}.meta.json"


def meta_path(path: str | Path) -> Path:
    """返回图片 path 对应的 sidecar 路径（按 UID 查找；若找不到则返回猜测路径）"""
    p = Path(path).resolve()
    # 如果 path 已经是 meta 文件，直接返回
    if str(p).endswith(".meta.json"):
        return p

    dx = _dx_from_path(p)
    if not dx:
        # 无法推断 DX，fallback 到旧式路径（理论上不会发生）
        try:
            rel = p.relative_to(PROJECTS_DIR)
        except ValueError:
            rel = p
        return META_DIR / rel.with_suffix(rel.suffix + ".meta.json")

    dx_dir = PROJECTS_DIR / dx
    md5 = compute_md5(p)
    data = read_uid_map(dx_dir)
    uid = data.get("md5_index", {}).get(md5)
    if uid:
        return sidecar_path_by_uid(dx_dir, uid)

    # 没有 MD5 记录：fallback 到旧式路径镜像
    try:
        rel = p.relative_to(PROJECTS_DIR)
    except ValueError:
        rel = p
    return META_DIR / rel.with_suffix(rel.suffix + ".meta.json")


def read_meta(path: str | Path) -> Optional[dict]:
    """读取图片 sidecar。

    优先按 MD5 在 uid_map.md5_index 里找 UID，再读 05_META/DXxxxx/sidecars/UID.meta.json。
    这样即使图片被改名或移动，只要内容没变就能找到元数据。
    """
    p = Path(path).resolve()
    dx = _dx_from_path(p)
    if not dx:
        return _read_json(meta_path(p)) or None

    dx_dir = PROJECTS_DIR / dx
    md5 = compute_md5(p)
    if not md5:
        return None

    data = read_uid_map(dx_dir)
    uid = data.get("md5_index", {}).get(md5)
    if not uid:
        return None

    sp = sidecar_path_by_uid(dx_dir, uid)
    meta = _read_json(sp)
    if meta:
        # 如果图片路径变了，同步更新 file 字段（不写入，仅返回时修正）
        try:
            rel = p.relative_to(PROJECTS_DIR / dx)
            if meta.get("file") != str(rel):
                meta["file"] = str(rel)
        except ValueError:
            meta["file"] = str(p.name)
    return meta


def write_meta(path: str | Path, data: dict):
    """写入图片 sidecar（必须包含 uid）"""
    uid = data.get("uid")
    if not uid:
        raise ValueError("write_meta 需要 data 中包含 uid")
    dx = _dx_from_path(path)
    if not dx:
        # 无法推断 DX 时 fallback
        mp = meta_path(path)
    else:
        mp = sidecar_path_by_uid(PROJECTS_DIR / dx, uid)
    mp.parent.mkdir(parents=True, exist_ok=True)
    _atomic_write_json(mp, data)


def ensure_meta(path: str | Path, **kwargs) -> dict:
    """读取或创建 sidecar，并用 kwargs 更新字段"""
    meta = read_meta(path) or {}
    meta.update(kwargs)
    if "md5" not in meta or not meta["md5"]:
        meta["md5"] = compute_md5(path)
    if "file" not in meta or not meta["file"]:
        try:
            meta["file"] = str(Path(path).relative_to(PROJECTS_DIR / _dx_from_path(path)))
        except ValueError:
            meta["file"] = str(Path(path).name)
    write_meta(path, meta)
    return meta


def update_meta(path: str | Path, **kwargs) -> Optional[dict]:
    """更新 sidecar 字段，不存在则创建"""
    return ensure_meta(path, **kwargs)


# ---------------------------------------------------------------------------
# UID Map 操作（MD5 主键 + UID 索引）
# ---------------------------------------------------------------------------

def uid_map_path(dx_dir: str | Path) -> Path:
    """返回 DX 目录的 uid_map.json 路径（位于 05_META/DXxxxx/）"""
    dx = Path(dx_dir).name
    return META_DIR / dx / "uid_map.json"


def read_uid_map(dx_dir: str | Path) -> dict:
    """读取 DX 的 uid_map.json"""
    return _read_json(uid_map_path(dx_dir), default={
        "dx": "", "version": 2, "groups": {}, "images": {}, "md5_index": {}
    })


def write_uid_map(dx_dir: str | Path, data: dict):
    """写入 DX 的 uid_map.json"""
    p = uid_map_path(dx_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    data.setdefault("version", 2)
    data.setdefault("md5_index", {})
    if not data.get("dx"):
        data["dx"] = Path(dx_dir).name
    _atomic_write_json(p, data)


def ensure_uid_map(dx_dir: str | Path) -> dict:
    """确保 uid_map.json 存在并返回"""
    data = read_uid_map(dx_dir)
    for key in ("images", "groups", "md5_index"):
        if not data.get(key):
            data[key] = {}
    return data


def register_image_in_map(dx_dir: str | Path, uid: str, group_id: str, stage: str,
                          role: str, file_path: str, parent_uid: Optional[str] = None,
                          source_file: Optional[str] = None, md5: Optional[str] = None,
                          **extra) -> dict:
    """在 uid_map.json 中注册/更新一张图片。

    如果 md5 已存在，则更新该条目的路径（支持改名/移动后重新注册）。
    """
    data = ensure_uid_map(dx_dir)
    dx = Path(dx_dir).name
    data["dx"] = dx

    # 计算 MD5
    full_path = Path(file_path)
    if not full_path.exists() and PROJECTS_DIR in full_path.parents:
        pass
    if not full_path.exists() and dx:
        alt = PROJECTS_DIR / dx / file_path
        if alt.exists():
            full_path = alt

    md5_val = md5 or compute_md5(full_path)

    # MD5 主键：已存在则复用 UID，只更新路径
    existing_uid = data.get("md5_index", {}).get(md5_val)
    if existing_uid and existing_uid in data.get("images", {}):
        uid = existing_uid

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
    if md5_val:
        entry["md5"] = md5_val
    if parent_uid:
        entry["parent_uid"] = parent_uid
    if source_file:
        entry["source_file"] = source_file
    entry.update(extra)

    data["images"][uid] = entry
    data["md5_index"][md5_val] = uid

    write_uid_map(dx_dir, data)
    return entry


def find_in_map(dx_dir: str | Path, **filters) -> List[dict]:
    """按字段过滤 uid_map 中的图片"""
    data = read_uid_map(dx_dir)
    results = []
    for entry in data.get("images", {}).values():
        if all(entry.get(k) == v for k, v in filters.items()):
            results.append(entry)
    return results


def resolve_uid(dx_dir: str | Path, uid: str) -> Optional[dict]:
    """根据 uid 查找 uid_map 中的条目"""
    data = read_uid_map(dx_dir)
    return data.get("images", {}).get(uid)


def resolve_md5(dx_dir: str | Path, md5: str) -> Optional[dict]:
    """根据 MD5 查找 uid_map 中的条目"""
    data = read_uid_map(dx_dir)
    uid = data.get("md5_index", {}).get(md5)
    if uid:
        return data.get("images", {}).get(uid)
    return None


def find_children(dx_dir: str | Path, parent_uid: str, stage: Optional[str] = None) -> List[dict]:
    """查找某个 uid 在指定 stage 下的子图片"""
    data = read_uid_map(dx_dir)
    results = []
    for entry in data.get("images", {}).values():
        if entry.get("parent_uid") == parent_uid:
            if stage is None or entry.get("stage") == stage:
                results.append(entry)
    return results


def reconcile_file(dx_dir: str | Path, path: str | Path) -> Optional[dict]:
    """根据 MD5 主键查找/修正单张图片的元数据（改名/移动后仍能找到）"""
    p = Path(path).resolve()
    md5 = compute_md5(p)
    if not md5:
        return None
    data = read_uid_map(dx_dir)
    uid = data.get("md5_index", {}).get(md5)
    if not uid:
        return None

    entry = data.get("images", {}).get(uid)
    if not entry:
        return None

    # 路径变了就更新
    try:
        rel = str(p.relative_to(Path(dx_dir)))
    except ValueError:
        rel = str(p.name)

    if entry.get("file") != rel:
        entry["file"] = rel
        entry["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
        data["images"][uid] = entry
        write_uid_map(dx_dir, data)
    return entry


def reconcile_dx(dx_dir: str | Path) -> Dict[str, int]:
    """扫描 DX 下所有图片，用 MD5 修正 uid_map 中的路径。
    优化：每个 DX 只读一次 uid_map，按路径快速匹配；只有路径找不到时才算 MD5。
    返回统计：{'found': N, 'updated': N, 'missing': N}
    """
    dx_dir = Path(dx_dir)
    stats = {"found": 0, "updated": 0, "missing": 0}
    stages = {
        "ai": dx_dir / "01_AI",
        "rembg": dx_dir / "02_REM_BG",
        "sticker": dx_dir / "03_UPLOAD",
    }

    data = read_uid_map(dx_dir)
    images = data.get("images", {})
    md5_index = data.get("md5_index", {})
    # 路径 -> uid 快速索引
    path_to_uid = {}
    for uid, entry in images.items():
        rel = entry.get("file")
        if rel:
            path_to_uid[rel] = uid

    changed = False
    for stage, dir_path in stages.items():
        if not dir_path.exists():
            continue
        for f in sorted(dir_path.iterdir()):
            if not f.is_file():
                continue
            if f.suffix.lower() not in (".png", ".jpg", ".jpeg", ".webp"):
                continue
            try:
                rel = str(f.relative_to(dx_dir))
            except ValueError:
                rel = str(f.name)

            uid = path_to_uid.get(rel)
            entry = images.get(uid) if uid else None
            if entry is None:
                # 路径对不上，用 MD5 找回旧记录
                md5 = compute_md5(f)
                uid = md5_index.get(md5)
                if uid:
                    entry = images.get(uid)
                    if entry is not None:
                        old_rel = entry.get("file")
                        if old_rel and old_rel in path_to_uid:
                            del path_to_uid[old_rel]
                        entry["file"] = rel
                        entry["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                        path_to_uid[rel] = uid
                        stats["updated"] += 1
                        changed = True

            if entry is None:
                stats["missing"] += 1
                continue

            stats["found"] += 1
            if entry.get("stage") != stage:
                entry["stage"] = stage
                stats["updated"] += 1
                changed = True

    if changed:
        write_uid_map(dx_dir, data)
    return stats


# ---------------------------------------------------------------------------
# Stage 专用便捷函数
# ---------------------------------------------------------------------------

def register_inbox(inbox_path: str | Path, uid: str, group_id: str, role: str):
    """注册 INBOX 原图"""
    p = Path(inbox_path)
    update_meta(p, uid=uid, group_id=group_id, stage="inbox", role=role,
                md5=compute_md5(p))
    # INBOX 没有 dx 文件夹，不写 uid_map


def register_ai(ai_path: str | Path, uid: str, group_id: str, role: str,
                parent_uid: Optional[str] = None, inbox_file: Optional[str] = None):
    """注册 AI 生成图"""
    p = Path(ai_path)
    dx_dir = p.parent.parent
    md5 = compute_md5(p)
    update_meta(p, uid=uid, group_id=group_id, stage="ai", role=role,
                parent_uid=parent_uid, source_file=inbox_file, md5=md5)
    register_image_in_map(dx_dir, uid, group_id, "ai", role, str(p),
                          parent_uid=parent_uid, source_file=inbox_file, md5=md5)


def register_rembg(cut_path: str | Path, uid: str, group_id: str, role: str,
                   parent_uid: str, ai_file: str):
    """注册去背输出图（_cut.png）"""
    p = Path(cut_path)
    dx_dir = p.parent.parent
    md5 = compute_md5(p)
    update_meta(p, uid=uid, group_id=group_id, stage="rembg", role=role,
                parent_uid=parent_uid, source_file=ai_file, md5=md5)
    register_image_in_map(dx_dir, uid, group_id, "rembg", role, str(p),
                          parent_uid=parent_uid, source_file=ai_file, md5=md5)


def register_sticker(upload_path: str | Path, uid: str, group_id: str, role: str,
                     parent_uid: str, cut_file: str):
    """注册贴图成品"""
    p = Path(upload_path)
    dx_dir = p.parent.parent
    md5 = compute_md5(p)
    update_meta(p, uid=uid, group_id=group_id, stage="sticker", role=role,
                parent_uid=parent_uid, source_file=cut_file, md5=md5)
    register_image_in_map(dx_dir, uid, group_id, "sticker", role, str(p),
                          parent_uid=parent_uid, source_file=cut_file, md5=md5)


def register_bw(bw_path: str | Path, uid: str, group_id: str, role: str,
                source_uids: List[str], source_files: List[str]):
    """注册 BW 合成图"""
    p = Path(bw_path)
    dx_dir = p.parent.parent
    md5 = compute_md5(p)
    update_meta(p, uid=uid, group_id=group_id, stage="bw", role=role,
                source_uids=source_uids, source_files=source_files, md5=md5)
    register_image_in_map(dx_dir, uid, group_id, "bw", role, str(p),
                          source_uids=source_uids, source_files=source_files, md5=md5)


def register_upload(upload_path: str | Path, uid: str, group_id: str, role: str,
                    parent_uid: str, source_file: str):
    """注册上款最终图（兼容 register_sticker）"""
    register_sticker(upload_path, uid, group_id, role, parent_uid, source_file)


# ---------------------------------------------------------------------------
# 迁移：从现有 source_map.json + 文件名构建 uid_map
# ---------------------------------------------------------------------------

def _extract_role_from_name(filename: str) -> str:
    """从文件名推断 role（新命名规则见 wb_naming；失败回退旧逻辑）"""
    try:
        import wb_naming
        role = wb_naming.role_from_name(filename)
        if role != "?":
            return role
    except Exception:
        pass
    stem = Path(filename).stem
    if stem.endswith("_cut"):
        stem = stem[:-4]
    if "_白T" in filename or "_黑T" in filename:
        m = re.search(r"_([黑]?[BW])_", filename)
        if m:
            return m.group(1)
    if "_白BW" in filename:
        return "白BW"
    if "_黑BW" in filename:
        return "黑BW"
    if "_BW" in filename:
        return "BW"
    if "_" in stem:
        last = stem.split("_")[-1]
        if last in ("B", "W", "BW", "WB"):
            return last
        if last in ("黑B", "黑W", "黑BW"):
            return last
    return "?"


def migrate_dx(dx_dir: str | Path, fallback_group_prefix="G") -> dict:
    """扫描 DX 目录，基于 MD5 构建 uid_map 和 sidecars。
    用于旧项目迁移；同名/同内容文件会合并为同一条元数据。
    """
    dx_dir = Path(dx_dir)
    dx = dx_dir.name
    data = ensure_uid_map(dx_dir)
    data["dx"] = dx

    # 迁移是重建过程：清空旧的 groups/images/md5_index，避免残留
    data["groups"] = {}
    data["images"] = {}
    data["md5_index"] = {}
    write_uid_map(dx_dir, data)

    smap = _read_json(dx_dir / "source_map.json", default={})
    src_by_file = {s.get("file", ""): s for s in smap.get("sources", [])}

    default_group_id = f"{fallback_group_prefix}_{dx}"

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

            md5 = compute_md5(f)

            # 如果 MD5 已注册过，只更新路径（复本/移动场景）
            existing_uid = data.get("md5_index", {}).get(md5)
            if existing_uid:
                entry = data["images"].get(existing_uid, {})
                try:
                    entry["file"] = str(f.relative_to(dx_dir))
                except ValueError:
                    entry["file"] = str(f.name)
                entry["updated_at"] = time.strftime("%Y-%m-%d %H:%M:%S")
                # stage 以实际所在目录为准
                entry["stage"] = stage
                data["images"][existing_uid] = entry
                write_uid_map(dx_dir, data)
                continue

            # 读取旧 sidecar（兼容老位置）
            legacy_meta = f.with_suffix(f.suffix + ".meta.json")
            meta = _read_json(legacy_meta) if legacy_meta.exists() else None
            uid = meta.get("uid") if meta else None
            role = meta.get("role") if meta else None

            if not uid:
                src_info = src_by_file.get(f.name, {})
                src_id = src_info.get("src_id", "")
                if src_id:
                    uid = f"UID_{src_id}"
                else:
                    uid = f"UID_MIGRATED_{md5[:16]}"

            group_id = default_group_id
            if not role:
                role = _extract_role_from_name(f.name)

            rel = str(f.relative_to(dx_dir))
            register_image_in_map(dx_dir, uid, group_id, stage, role, str(f), md5=md5)
            ensure_meta(f, uid=uid, group_id=group_id, stage=stage, role=role, md5=md5)

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
    elif len(sys.argv) > 2 and sys.argv[1] == "reconcile":
        stats = reconcile_dx(sys.argv[2])
        print(f"[reconcile] {sys.argv[2]}: {stats}")
    else:
        print("用法:")
        print("  python wb_meta.py migrate")
        print("  python wb_meta.py migrate-dx <DX_DIR>")
        print("  python wb_meta.py reconcile <DX_DIR>")
