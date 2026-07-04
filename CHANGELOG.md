# CHANGELOG — PS 贴图流水线

## v2.3 — 2026-07-04

### 修复：黑T贴图失败导致 03_UPLOAD 无输出

- `jsx/place_design.jsx`
  - 改为路径模板：JSX 内部用 `app.open()` 打开胚衣与设计图，执行 `duplicate()` 后关闭两者。
  - 解决跨 JSX 调用传递 Photoshop 文档引用失败的问题（错误 8192 / 1233）。
- `wb_sticker_ps.py`
  - `StickerSession` 不再缓存胚衣/设计图文档，改为每次贴图都通过 JSX 重新打开并关闭文件。
  - 单 DX 内仍复用同一个 Photoshop COM 会话，避免每贴一张图都重启 PS。
- `process_black.py` v2.3
  - 复用 `wb_sticker_ps.StickerSession` 进行黑T贴图，统一使用路径模板 JSX。
  - 修复 `{{TORSO_DOC_NAME}}` / `{{DESIGN_DOC_NAME}}` 占位符与当前 JSX 不匹配导致的执行失败。
  - BW 合成仍在本脚本内完成，直接操作 `StickerSession` 的 COM 会话。
- `ps_batch_one.py` v2.2
  - 修复从 `ps_batch` 导入不存在的 `process_color` / `close_all_docs` 导致的 `ImportError`。
  - 改为直接调用 `ps_batch.process_dx` 完成白T/黑T BW 合成。

### 验证

- DX0319（含 `_黑BW_cut.png`）完整流水线通过，03_UPLOAD 生成 6 张成品：
  - `DX0319_B_白T.jpg`、`DX0319_W_白T.jpg`
  - `DX0319_B_黑T.jpg`、`DX0319_W_黑T.jpg`
  - `DX0319_白BW.jpg`、`DX0319_黑BW.jpg`
- DX0321（含 `_黑B_cut.png`、`_黑W_cut.png`）同样通过。

## v2.2 — 2026-07-03

### 新增：UID/group_id 元数据传播

- 引入共享模块 `wb_meta.py` v2.0（MD5 主键 + UID sidecar）
- 元数据统一写入 `D:\Semems WB\05_META\DXxxxx\`，与 `01_AI/02_REM_BG/03_UPLOAD` 图片分离
- `wb_sticker_ps.py`
  - 读取 `_cut.png` sidecar，为每个上传成品写入 `05_META` 下的 `.meta.json`
  - 调用 `wb_meta.register_sticker()` 更新 `05_META` 下的 `uid_map.json`
- `process_black.py`
  - 黑版贴图完成后注册 `黑B`/`黑W` 元数据到 `05_META`
  - 黑 BW 合成后注册 `黑BW` 元数据到 `05_META`
- `ps_batch.py`
  - BW 合成图使用派生 UID `{源uid}_{role}` 注册，避免覆盖 sticker 条目
  - 记录 `source_uids`/`source_files` 保留 B/W 来源关系
- `ps_sticker_one.py` / `ps_batch_one.py`
  - 增加 `wb_meta` 导入，由被调用脚本完成具体注册
- 缺失 sidecar 时自动调用 `wb_meta.migrate_dx()` 兜底到 `05_META`
- 所有元数据操作异常均被捕获，不影响 PS 自动化流程
