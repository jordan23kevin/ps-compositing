# CHANGELOG — 贴图流水线（纯软件）

## v2.5.2 — 2026-07-13（正面胚衣双组五参：单面款 / 双面款W+B）

- **wb_sticker_ps.py v2.5.2**
  - 正面胚衣（W白/白W11、W黑/黑W11）支持「双组五参」：素材库 `.meta.json` 顶层五参用于 **① 单面款（只有W贴图）**，新增 `"bw"` 子块五参用于 **② 双面款（有W+B贴图）**。背面胚衣（B白/B黑）只一组，行为不变。
  - `process_dx_folder` 按产品类型自动选组：`design_type=="W"`（单面款）用顶层第一组；`design_type=="BW"`（双面款）**正面**用 `bw` 第二套、**背面**用顶层。
  - **强制双组都填**：双面款正面若缺 `bw` 第二组（缺失/非数字/为0），直接抛 `ValueError`（如 `W白 胚衣: 缺少双面款(W+B)第二组参数「bw」…`），不偷偷用第一组兜底。
  - 新增 `_select_meta(full_meta, meta_set)` 负责选组与校验。
- **素材库页面 peiyi.html**：W白/W黑 卡片渲染两组输入框（① 单面款 / ② 双面款W+B），第二组未填显示「⚠ 双面款未填」；保存时把第二组归入 `payload.bw`，后端 `/api/peiyi/meta` 合并写入顶层 + `bw` 两套。B白/B黑 保持一组。
- 验证：`_select_meta` 单测通过（单面取顶层 / 双面取bw / 缺bw报错 / bw含0报错）；bridge 实时接口回读 `bw` 块 + `bw_missing` 标记正常；临时素材写回 `.meta.json` 双组并存。

## v2.4.1 — 2026-07-12（黑T 平铺图自动优化）

- **wb_sticker_ps.py v2.4.1**
  - 新增 `black_optimize` 开关：当通用/白版设计图贴在黑胚衣上时，自动调用 `black_opt.black_shirt_print_optimize`（白墨打底 + 暗部提亮 + 饱和补偿），解决黑T平铺图「边缘/文字发暗、发脏」问题。
  - `process_dx_folder` 中 4 个黑胚衣落点（W黑T×2、B黑T×2）均启用该优化；白胚衣落点与已有 `_黑*_cut.png` 专用款不受影响。
  - 新增 `black_opt.py`：自 `check_rem.py` 抽取黑衫优化函数，带 `cv2` 兜底，缺 `cv2` 时自动跳过、不中断贴图。
  - 已验证：DX0648 重新生成后黑T不发暗，与白T几何完全一致。

## v2.0.0 — 2026-07-12（重大：去除 Photoshop 依赖）

- **ps_batch.py v2.0.0**
  - 彻底移除 `get_ps` / `open_doc` / `wait_docs` / `close_docs` / `export_bw` 等 COM 函数与 `win32com`/`pythoncom` 导入；`main()` 不再 `CoInitialize`。
  - BW 合成全程纯 PIL；整条「贴图 + BW」流水线零 PS 依赖。
- **wb_sticker_ps.py v2.4.0**：`StickerSession.place_design` 改为纯 PIL（trim→缩放→平移→绕中心旋转→normal 合成），复刻原 `place_design.jsx` 定位；移除 win32com。
- **process_black.py / process_white.py v2.5.0**：
  - 贴花复用纯软件 `StickerSession`（不再连 PS）。
  - `bw_synth` 改用 `ps_batch.process_dx` 纯软件合成 BW，移除对缺失的 PS 动作集「正反图」的依赖（否则点贴图会崩溃）。
- **BW 合成规格（v1.8→v2.0）**：依据 DX0481 参考图，圆直径 595、正面图宽度贴合圆圈（≈44.4%）、圆心 (1014,1449)、白边 5px、无阴影。
- 验证：纯软件贴花与旧 PS 版像素差 0.3–0.9；单款 BW 合成 ≈0.3s；全流程无 Photoshop 进程。

## v2.4 — 2026-07-10（历史，PS 版）

### 新增：单面款旧产物兜底清理（配合 04_OS check_rem v2.2.7 分流修复）

- `wb_sticker_ps.py` v2.3
  - 新增 `real_sides(dx_folder)`：遍历 `02_REM_BG/*_cut.png`，去掉 `黑`/`白` 前缀后解析 `B/W/BW/WB`（`BW/WB` 展开为 `B+W`），返回真实面集合。
  - 新增 `cleanup_stale_uploads(dx_folder)`：当真实面严格为 `{B}` 或 `{W}`（单面款）时，清理 `03_UPLOAD` 中已不存在的互补面胚衣图（`_W_白T/_W_黑T` 或 `_B_*`）与旧 BW 平铺图（`_白BW/_黑BW`）。双面款/未知不动。
  - `process_dx_folder()` 开头调用 `cleanup_stale_uploads()`。
- `process_black.py` v2.4 / `process_white.py` v1.1
  - 入口 `main()` 同样调用 `cleanup_stale_uploads()`（直接跑黑/白版专用贴图时也生效）。
- `ps_batch.py` v1.4.1
  - `process_dx()` 合成 BW 前对单面款跳过并删除残留 `_白BW/_黑BW` 平铺图（防止「新单面图 + 旧互补面图」被误拼成平铺）。

### 背景

- 04_OS `check_rem.py` v2.2.7 已让单面款改走模特图贴图（不再进本平铺流程），本仓的清理作为「直接跑 PS 脚本 / 历史残留」场景的兜底。

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
