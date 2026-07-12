# Semems WB 印花贴图流水线 — 架构 / 方法 / 问题方案 / 复现回滚

> **同步基线**：`pipeline-2026-07-12`（四个仓库均已打此基线 tag；其中 `white_t_mockup` 因与 ZCodeProject 同仓不同分支，使用 `pipeline-2026-07-12-wtm`，详见第八章）
> **最后更新**：2026-07-12
> **关键结论**：贴图流水线（平铺图贴花 + BW 合成 + 模特图）**已完全不依赖 Photoshop**，全部为纯软件（PIL / OpenCV / numpy）。重装或迁移 Photoshop 不再影响贴图。

---

## 一、系统总览

```
┌──────────────────────────────────────────────────────────────────────────┐
│  浏览器 (端口 8765 / 8766)                                                 │
│    Y2 控制台 / 去背预览 / 贴图按钮                                          │
└───────────────────────────┬──────────────────────────────────────────────┘
                              │ HTTP/子进程
                              ▼
┌──────────────────────────────────────────────────────────────────────────┐
│  ZCodeProject (大脑 / 编排)                                                 │
│    lovart_bridge.py  → 启动器（双击 D:\Semems WB\01_INBOX\lovart_bridge.bat）│
│    engine/check_rem.py (端口 8766 常驻) → 「贴图」按钮触发贴图流水线          │
└──────────┬───────────────────────────────────────┬───────────────────────┘
           │ 子进程调用                              │ 子进程调用
           ▼                                        ▼
┌─────────────────────────────┐      ┌────────────────────────────────────────┐
│  ps-compositing 仓库          │      │  white_t_mockup 仓库 (E:\Kimi Code)     │
│  (E:\Claude code\ps)          │      │  模特图（人穿着、有褶皱）贴图引擎        │
│  ── 平铺图（衣服摊平拍）──     │      │  v1.8.0：gradient 位移 + 布料同步明度   │
│  wb_sticker_ps.StickerSession │      └───────────────────┬────────────────────┘
│    (纯 PIL affine 贴花)        │                          │ subprocess 调用
│  ps_batch.compose_bw_pil      │                          │ python -m white_t_mockup
│    (纯 PIL BW 合成)            │                          ▼
└──────────────┬───────────────┘              ┌────────────────────────────────┐
                 │ 命名/胚衣/去背参数          │  D:\Semems WB\04_OS\engine        │
                 ▼                            │    wb_naming.py（命名规则唯一出处） │
┌─────────────────────────────┐              │    w_mockup_extra.py（生产参数） │
│  D:\Semems WB (数据根)        │              └────────────────────────────────┘
│    01_INBOX 原图入口          │
│    02_PROJECTS\DXxxxx\        │
│       01_AI / 02_REM_BG / 03_UPLOAD
│    03_MATERIAL 胚衣+meta.json │
│    _tpl 扭曲素材 (D:\Semems\1胚衣\_tpl)
│    05_META UID/元数据          │
└─────────────────────────────┘
```

### 两条贴图链

| 胚衣类型 | 触发 | 引擎 | 产物 |
|---|---|---|---|
| **平铺图**（衣服摊平拍，无褶皱） | 贴图按钮 → `ps_sticker_one.py` / `process_black.py` / `process_white.py` | `ps-compositing` 纯 PIL | `DX_W白T / W黑T / B白T / B黑T.jpg` + `DX_白BW / 黑BW.jpg` |
| **模特图**（人穿着，有褶皱） | 单面款（只有 W 或只有 B）→ `w_mockup_extra.py` | `white_t_mockup` 纯软件 | `DX_白W / 黑W / 白B / 黑B.jpg` |

> 命名规则（平铺 `DX_W白T`、模特 `DX_白W` 等）唯一出处：`D:\Semems WB\04_OS\engine\wb_naming.py`。改命名只改这一处，全链路自动跟随。

---

## 二、仓库清单与版本

| 仓库 | 本地路径 | 远程 | 分支 | 本次版本 | 角色 |
|---|---|---|---|---|---|
| ZCodeProject | `C:\Users\Administrator\ZCodeProject` | `github.com/jordan23kevin/ZCodeProject.git` | master | bridge v2.3.23 | 大脑/编排（lovart_bridge + check_rem） |
| ps-compositing | `E:\Claude code\ps` | `github.com/jordan23kevin/ps-compositing.git` | master | **v2.5.0** | 平铺图贴花 + BW 合成（纯 PIL） |
| white_t_mockup | `E:\Kimi Code\white_t_mockup` | （本地） | white-t-mockup | **v1.8.0** | 模特图贴图引擎（gradient + 布料同步明度） |
| 04_OS | `D:\Semems WB\04_OS` | `github.com/.../semems-wb-04os.git` | master | w_mockup_extra v2.4 | 命名规则 + 生产参数 |

> 同步点 tag：四个仓库均打 `pipeline-2026-07-12`。

---

## 三、环境与依赖

### 3.1 Python 运行时

| 组件 | 路径/版本 | 用途 |
|---|---|---|
| 系统 Python | `C:\Users\Administrator\AppData\Local\hermes\hermes-agent\venv\Scripts\python.exe`（Python 3.11.15） | 运行 ps-compositing / 04_OS / check_rem 子进程 |
| 托管 Python | `C:\Users\Administrator\.workbuddy\binaries\python\...`（3.13） | WorkBuddy 工具内部，非本流水线 |
| `E:/python_packages` | cv2 4.13.0 / numpy 2.4.6 / torch 等 | white_t_mockup 需要 cv2/numpy；通过 `PYTHONPATH` 注入 |

### 3.2 关键包版本（实测）

| 包 | 版本 | 备注 |
|---|---|---|
| Pillow | 12.2.0 | ps-compositing 唯一硬依赖 |
| numpy | 2.4.3（venv）/ 2.4.6（E:/python_packages） | 两边都有 |
| opencv-python (cv2) | 4.13.0 | 仅在 `E:/python_packages`，需 `PYTHONPATH=E:/python_packages` 才能 import |

### 3.3 PYTHONPATH 约定（重要，否则 cv2 找不到）

- **ps-compositing / 04_OS**：直接用 hermes venv python 即可（仅需 Pillow/numpy，venv 自带）。
- **white_t_mockup**：必须
  ```
  PYTHONPATH=E:/python_packages;E:/Kimi Code
  ```
  （前者提供 cv2/numpy/torch，后者提供 white_t_mockup 包）。check_rem 子进程调用时已显式设置。

### 3.4 requirements.txt

- `ps-compositing/requirements.txt`：`Pillow>=10`, `numpy`
- `white_t_mockup/requirements.txt`：`Pillow`, `numpy`, `opencv-python`；并注明 cv2/torch 实际来自 `E:/python_packages`

---

## 四、已完成方法

### 4.1 命名规则（单一出处）
`04_OS/engine/wb_naming.py` 定义 `FLAT_FMT` / `MODEL_FMT` / `BW_FMT` / `FLAT_STEMS`。所有读图/写图方（check_rem、ps_batch、w_mockup_extra、wb_meta）都走它解析，改命名只改这一处。

平铺图：`DX0650_W白T.jpg` / `DX0650_B黑T.jpg`；模特图：`DX0650_白W.jpg` / `DX0650_黑B.jpg`；BW：`DX0650_白BW.jpg` / `DX0650_黑BW.jpg`。

### 4.2 平铺图纯软件贴花（v2.5.0）
`wb_sticker_ps.StickerSession.place_design`（纯 PIL）：
1. 按 `ALPHA_THRESHOLD=20` 裁剪设计图透明边距；
2. 读取胚衣同名 `.meta.json` 五参：`width/height/rotation/highest_y/center_x`；
3. 按 `width/设计图宽度` 缩放，以「顶边中点」为参考点做平移与旋转；
4. 绕图层中心旋转 `-rotation` 度（PS 顺时针为正 → PIL 取负匹配）；
5. `alpha_composite` 正常混合到胚衣。
定位参数来自 `D:\Semems WB\03_MATERIAL\{W白,W黑,B白,B黑}\*.meta.json`（胚衣真实尺寸），`config.py` 中的 `FRONT_NEW/BACK_NEW` 仅保留为 `process_black.py/process_white.py` 旧线回退。与早期 place_meta 参考图平均像素差约 **5.5/255**（仅边缘抗锯齿/JPEG 差异，肉眼不可辨）。

### 4.3 BW 合成（v2.0.0，依据 DX0481 参考图）
`ps_batch.compose_bw_pil`：
- 底图 1340×1785，圆形插图直径 **595**；
- 正面图缩放为「**宽度 = 圆圈直径**」= 实测比例 **44.4%**（595/1340）；高度自然溢出后垂直居中裁切，只保留圆圈内部分；
- 圆心 `(1014, 1449)` = 宽 75.67% / 高 81.18%（参考图实测）；
- **白边 5px**，**无阴影**；
- 单面款（02_REM_BG 只有 B 或只有 W）自动删除残留 BW，避免"新单面+旧互补面"误拼。

### 4.4 模特图 gradient 位移贴合（v1.7.0，用户选定 s=90）
`white_t_mockup` `apply_displacement(disp_mode="gradient")`：把 disp 当高度场，沿褶皱切线做 2D 梯度位移，印花真正"裹"在褶皱上。生产参数（`w_mockup_extra.py`）：
`--disp-strength 90 --disp-smooth 40 --disp-dead-zone 15 --preserve-color`。
- `smooth=40` 只保留大褶皱；`dead_zone=15` 软斜坡让小起伏趋零并消除硬台阶撕裂；
- 位移场经 `_limit_gradient_2d`（|∇off|≤0.45）消除尖锐褶皱脊处的镜像重影；
- 手部/前景遮挡物区位移压平到 128（`occluder`），避免"引力场"假位移；
- `--preserve-color`：白/黑模特图都不走 multiply/阴影/高光/降饱和，贴图颜色与源文件一致。

### 4.5 布料同步明度（v1.8.0，2026-07-12）
`white_t_mockup` `apply_fabric_synced_shading`：
- **仅缩放印花 HSV 的 V（明度），H/S 零偏差** → 印花固有色与源文件完全一致；
- 明度严格跟随同位置布料（shading = 布料局部亮度 / 平整处亮度）：布料暗处印花同比例变暗，杜绝"布料暗了印花仍高亮"；
- 判据用**局部对比**（非全局亮度），只压"深褶皱窄缝"，平整/均匀区（含黑衫整片暗布）保持原始亮度 → 零压暗；
- 深褶处 shading→0，印花与布料阴影自然融合、细节消融（配合 occlusion 隐藏，杜绝悬浮感）；
- 手部/前景遮挡区 shading 强制置 1（不调制）。

### 4.6 黑衫白墨打底（v1.6.0 默认）
`white_t_mockup` `white_underbase`：自适应浓度白墨打底（越暗白墨越厚，max0.9/min0.05），解决近黑像素贴在黑衫上物理不可见的问题。两套黑衫流程（white_t_mockup 引擎 + check_rem「反黑」按钮）已接同款算法。

### 4.7 occlusion 保守策略
`04_OS/tpl_generator.py` 的 `compute_occlusion_map` 默认保守：印花区深褶皱仅隐藏约 2–6%。模特图**禁用** `strengthen_occlusion`（曾误把位移场偏离当深褶，导致印花顶部/侧边大块消失）。如需加强只对平铺胚衣，用更高阈值/更低 disp-scale。

### 4.8 单面款清理与元数据注册
`wb_sticker_ps.cleanup_stale_uploads`：单面款贴图前清理 03_UPLOAD 中已不存在的互补面/旧 BW 残留。`wb_meta.register_sticker` / `register_bw` 为每张产物写 sidecar（UID/group_id），溯源不依赖文件名。

### 4.9 平铺图通用黑衫自动优化（v2.4.1）
`wb_sticker_ps.place_design` 新增 `black_optimize` 开关：当通用/白版设计图没有对应 `_黑*_cut.png` 专用文件、必须直接贴在黑胚衣上时，自动调用 `black_opt.black_shirt_print_optimize`（自适应白墨打底 + 暗部提亮 + 饱和补偿）。这能避免通用图的半透明边缘/纹理与黑色胚衣混合导致的「边缘发暗、文字发脏」。
- 仅对 `process_dx_folder` 的 4 个黑胚衣落点（W黑T×2、B黑T×2）开启；
- 白胚衣落点、已有 `_黑B/_黑W/_黑BW_cut.png` 专用款均不受影响（专用款由 `process_black.py` 处理，已优化）。

---

## 五、遇到的问题与解决方案

| # | 问题 | 根因 | 解决方案 | 状态/落点 |
|---|---|---|---|---|
| 1 | BW 款贴图后只有 4 张平铺图、缺白BW/黑BW | 旧 BW 依赖 PS 动作集「正反图」，新装 PS 2025 该动作集缺失（`app.actionSets` undefined、无 .atn） | 路线B：用 DX0481 样张 PIL 像素级复刻 | ps_batch v1.6→v2.0 |
| 2 | 复刻的 BW 圆形小图被填成衣服色（白/黑块） | 自作主张加 `_remove_shirt_background` 把背景填成衣服色 | 去掉该函数，圆形小图**保留正面图完整背景**（木地板/报纸/鞋子/衣架） | ps_batch v1.7 |
| 3 | 比例 50% 还是 44.4%？ | 用户口述 50%，但参考图反向测量不符 | 反向测量正面图在圆圈内的最佳缩放 = **0.444**（=595/1340），用户确认用 44.4% | ps_batch v1.8 |
| 4 | 模特图文字"上下撕裂/重影" | displacement 硬死区 `|g-128|<dead_zone` 在轮廓线位移突跳 → 重映射折叠 | 软斜坡 `ramp=clip(d/dz,0,1)` 使 off 平滑趋零 | white_t_mockup 8541da3 |
| 5 | gradient 模式重影（尖锐褶皱脊镜像） | 原始梯度在脊处重映射折叠 | `_limit_gradient_2d` 自适应高斯模糊到 `|∇off|≤0.45`（雅可比 det≥0.1） | white_t_mockup v1.7 |
| 6 | 手部"引力场"假位移 | disp.png 含手/前景明暗，在遮挡区造出位移 | smooth 前把 occluder 覆盖区位移压平到 128 | white_t_mockup v1.7 |
| 7 | 黑衫贴图暗/近黑不可见 | 设计图近黑像素 + 默认混合 | `white_underbase` 白墨打底（v1.6 默认）+ `--preserve-color` | white_t_mockup v1.6 |
| 8 | 平铺图贴花去 PS | PS 启动慢、动作集缺失会崩、批量累加速度慢 | 纯 PIL affine 复刻（diff 0.3–0.9），彻底移除 win32com | ps-compositing v2.0/2.4/2.5 |
| 9 | `strengthen_occlusion` 在模特图大块消失 | 把位移场偏离当深褶，误判背景/身体边缘为深褶 | 还原保守 occlusion，新款不再自动加强 | 04_OS tpl_generator |
| 10 | 平铺图定位参数写死导致图案过小/偏位 | 旧 `FRONT_NEW/BACK_NEW` 的 `scale_percent=13.33%` 把图缩成胸口小标 | 改用素材库 `.meta.json` 五参（width/height/rotation/highest_y/center_x），按胚衣真实尺寸定位 | wb_sticker_ps v2.5.0 |
| 11 | 黑色平铺图贴花部分发暗/发脏 | 通用/白版 `_W_cut.png` 直接贴黑胚衣，设计图含大量半透明边缘/纹理，与黑色混合后变暗 | `wb_sticker_ps.place_design(black_optimize=True)` 自动做白墨打底+暗部提亮；无 `_黑*_cut.png` 时生效 | wb_sticker_ps v2.4.1 |

---

## 六、触发与运行

### 6.1 贴图按钮（check_rem 编排）
`check_rem.py`「📎 贴图」→ `_run_one_sticker` → 按序跑（均为纯软件子进程）：
1. `process_black.py`（黑T 专用贴花 + BW 合成，若本款有黑版专用文件）
2. `process_white.py`（白T 专用贴花 + BW 合成，若本款有白版专用文件）
3. `ps_sticker_one.py` → `wb_sticker_ps.process_dx_folder`（通用 B/W 平铺贴花）
4. `ps_batch_one.py` → `ps_batch.process_dx`（合成白BW/黑BW）
5. `quit_ps.py`（检测 PS 未运行则直接返回，安全无副作用）

> **不再打开 Photoshop**。BW 合成只读已生成的平铺图；若某款平铺图本身缺失或命名未更新，BW 会跳过，需先重走「贴图」。

### 6.2 独立命令（调试用）
```bash
# 纯软件平铺贴花（单款）
C:/Users/.../hermes-agent/venv/Scripts/python.exe "E:\Claude code\ps\ps_sticker_one.py" DX0641

# 纯软件 BW 合成（单款）
C:/Users/.../hermes-agent/venv/Scripts/python.exe "E:\Claude code\ps\ps_batch_one.py" DX0641

# 模特图（单面款）贴图
PYTHONPATH=E:/python_packages;E:/Kimi Code python -m white_t_mockup ...
```

---

## 七、100% 复现步骤

```bash
# 1) 拉取四个仓库到约定路径（见第二章）
git -C "C:\Users\Administrator\ZCodeProject" pull origin master
git -C "E:\Claude code\ps" pull origin master
git -C "E:\Kimi Code\white_t_mockup" pull origin white-t-mockup
git -C "D:\Semems WB\04_OS" pull origin master

# 2) 回滚到同步基线（见第八章）
git -C <repo> checkout pipeline-2026-07-12

# 3) 确认依赖（Pillow/numpy 在 venv；cv2 在 E:/python_packages）
C:/Users/.../hermes-agent/venv/Scripts/python.exe -c "import PIL,numpy;print(PIL.__version__,numpy.__version__)"

# 4) 启动大脑
双击 D:\Semems WB\01_INBOX\lovart_bridge.bat   # 启动 bridge + 守护 check_rem(8766)

# 5) 打开浏览器
#    Y2 控制台: http://127.0.0.1:8765
#    去背预览: http://127.0.0.1:8766
# 在去背预览页点「📎 贴图」即触发完整纯软件流水线。
```

---

## 八、回滚（快速回到任意历史版本）

四个仓库均已打基线 tag（ZCodeProject / ps-compositing / 04_OS 为 `pipeline-2026-07-12`；white_t_mockup 因与 ZCodeProject 同仓不同分支，使用 `pipeline-2026-07-12-wtm`）。回滚到本次同步点：

```bash
# 注意：white_t_mockup 与 ZCodeProject 同仓（ZCodeProject.git 的 white-t-mockup 分支），
# 因同仓不同分支，其里程碑 tag 用 pipeline-2026-07-12-wtm 以免与 master 的 pipeline-2026-07-12 冲突。
declare -A TAGS=(
  ["C:\Users\Administrator\ZCodeProject"]="pipeline-2026-07-12"
  ["E:\Claude code\ps"]="pipeline-2026-07-12"
  ["E:\Kimi Code"]="pipeline-2026-07-12-wtm"
  ["D:\Semems WB\04_OS"]="pipeline-2026-07-12"
)
for r in "${!TAGS[@]}"; do
  git -C "$r" fetch origin --tags
  git -C "$r" checkout "${TAGS[$r]}"
done
# 重启大脑使常驻进程(8766)生效：
双击 D:\Semems WB\01_INBOX\lovart_bridge.bat
```

如需任意更早版本：`git -C <repo> log --oneline` 找 SHA，`git -C <repo> checkout <SHA>`。
改 `ps_batch.py` / `wb_sticker_ps.py` 等纯软件脚本**存盘即生效**，无需重启 bridge；改 `white_t_mockup` / `w_mockup_extra.py` 需 kill 8766 进程后重启 `lovart_bridge.bat`。

### 已知良好基线（本次同步点）
| 仓库 | tag | SHA |
|---|---|---|
| ZCodeProject | pipeline-2026-07-12 | 5d20b23fe3b5b87673ab5d54b7a040d995ae80b4 |
| ps-compositing | pipeline-2026-07-12 | pipeline-2026-07-12（本仓库里程碑 tag，即本次提交） |
| white_t_mockup | pipeline-2026-07-12-wtm | e61cc5a16ece8691f7f716c4d42215a96528e239 |
| 04_OS | pipeline-2026-07-12 | bfdf71f8af7a803aa895b5708f724231083ae32e |

---

## 九、已知边界 / 遗留

- 整个 `02_PROJECTS` 从 DX0001 起有数百个**旧命名**（带下划线 `DXxxxx_B_白T.jpg`）平铺图；本次仅统一了用户指定的 DX0605/DX0645，其余未动，需要时再批量改。
- `white_t_mockup` 与 ZCodeProject **同仓**（均推送至 `github.com/jordan23kevin/ZCodeProject.git`），只是分支不同（master / `white-t-mockup`）；因此流水线里程碑 tag 为 `pipeline-2026-07-12-wtm`，避免与 master 的 `pipeline-2026-07-12` 重名冲突。
- `Photoshop` 仍可用于人工修图，但**流水线已不再调用它**。
