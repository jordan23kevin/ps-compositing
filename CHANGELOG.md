# CHANGELOG — PS 贴图流水线

## v2.2 — 2026-07-03

### 新增：UID/group_id 元数据传播

- 引入共享模块 `wb_meta.py`
- `wb_sticker_ps.py`
  - 读取 `_cut.png` sidecar，为每个上传成品写入 `.meta.json`
  - 调用 `wb_meta.register_sticker()` 更新 `uid_map.json`
- `process_black.py`
  - 黑版贴图完成后注册 `黑B`/`黑W` 元数据
  - 黑 BW 合成后注册 `黑BW` 元数据
- `ps_batch.py`
  - BW 合成图使用派生 UID `{源uid}_{role}` 注册，避免覆盖 sticker 条目
  - 记录 `source_uids`/`source_files` 保留 B/W 来源关系
- `ps_sticker_one.py` / `ps_batch_one.py`
  - 增加 `wb_meta` 导入，由被调用脚本完成具体注册
- 缺失 sidecar 时自动调用 `wb_meta.migrate_dx()` 兜底
- 所有元数据操作异常均被捕获，不影响 PS 自动化流程
