# 白边裁剪工具 v0.5.1

一款本地运行的批量图片处理工具，主打两件事：

- 自动裁掉 PPT、PDF、SVG 和常见图片周围的多余白边
- 将纯色或近似纯色背景处理成透明 PNG/WebP

适合处理论文插图、课件素材、截图、电商产品图、扫描件和设计素材。文件只在本机处理，不上传到任何服务器。

## 功能概览

| 功能 | 说明 |
|------|------|
| 白边裁剪 | 自动检测内容区域，去除四周空白 |
| PPT 处理 | 连接正在运行的 PowerPoint/WPS，处理当前页或全部页面 |
| PDF 处理 | 支持矢量裁剪，可输出 PDF/SVG，也可导出位图 |
| SVG 支持 | 支持 SVG 输入，支持 SVG/PDF/PNG 等输出 |
| 图片批处理 | 支持 PNG、JPG、JPEG、BMP、TIFF、WebP、GIF |
| 背景透明 | 支持自动取色、自定义颜色、容差、边缘连通和羽化 |
| 本地打包 | 一条 PowerShell 命令生成 Windows exe |

## 快速开始

安装依赖：

```bash
pip install -r requirement.txt
```

运行源码版：

```bash
python main.py
```

Windows 下如果项目中已有 `venv_clean`，请明确使用该环境启动，避免 MSYS/Git Bash 的 Python 缺少 `pywin32`：

```powershell
.\venv_clean\Scripts\python.exe .\main.py
```

打包 Windows exe：

```powershell
.\build.ps1
```

打包结果：

```text
dist/cropper-v0.5.1.exe
```

如果要覆盖已有 exe，请先关闭正在运行的 `cropper-v0.5.1.exe`。

## 界面说明

程序包含两个标签页。

### 白边裁剪

用于去除 PPT、PDF、SVG 或图片周围的空白区域。

处理来源：

| 来源 | 支持范围 |
|------|----------|
| PowerPoint | 当前幻灯片或全部幻灯片 |
| 本地文件 | PDF、SVG、PNG、JPG、BMP、TIFF、WebP、GIF |

输出格式：

| 格式 | 说明 |
|------|------|
| PDF | PPT/PDF/SVG 保留矢量；图片按原始像素嵌入，不重新采样 |
| SVG | 保留页面的视觉矢量内容；PDF 链接、批注、表单、书签和图层控制不保留 |
| PNG | 无损位图，支持透明 |
| TIFF | 无损位图，适合印刷或归档 |
| JPEG | 有损压缩，适合照片和分享 |
| WebP | 体积较小，适合网页 |
| GIF | 支持保留动画帧、时长和循环；输出会规范化帧处置以保持逐帧视觉一致 |

检测参数：

| 参数 | 说明 |
|------|------|
| 智能 | 综合亮度、颜色变化和边缘，推荐默认使用 |
| 简单 | 仅按亮度判断，速度最快，适合纯白背景 |
| 边缘敏感 | 更重视浅色边缘，适合内容接近白色的图 |
| 白色阈值 | 越高越只裁纯白，越低越容易裁掉浅灰 |
| 敏感度 | 越高越保守，越低裁剪越激进 |
| 边缘留白 | 先精确去白边，再额外保留的像素边距；输入 0 表示严格贴合内容 |
| 输出 DPI | 仅在 PPT/PDF/SVG 渲染为位图时生效，默认 300 DPI |
| 图片 DPI | PPT 导出 PDF/SVG 时限制内嵌位图的最高 DPI，默认 300；文字和形状仍为矢量 |

### 背景透明

用于把纯色或近似纯色背景变成透明背景。

典型用法：

1. 选择一张或多张图片
2. 选择目标颜色来源
3. 调整容差和羽化
4. 输出 PNG 或 WebP

参数说明：

| 参数 | 说明 |
|------|------|
| 四角平均 | 自动取图片四个角的平均颜色，适合常见白底图 |
| 左上角 | 只取左上角颜色，适合背景非常稳定的图 |
| 自定义 | 手动输入 `#RRGGBB`，例如 `#FFFFFF` |
| 容差 | 越高会移除更多相近颜色 |
| 边缘羽化 | 让透明边缘更柔和，减少硬边 |
| 只处理连到图片边缘的背景 | 默认开启，避免误删主体内部同色区域 |

建议默认保持“只处理连到图片边缘的背景”开启。只有在确实需要全图删除某种颜色时，再关闭它。

## 常见场景

| 场景 | 推荐功能 |
|------|----------|
| PPT 图表插入 Word/LaTeX | 白边裁剪，输出 PDF/SVG |
| PDF 论文插图提取 | 白边裁剪，输出 SVG/PNG |
| 网页截图去边 | 白边裁剪，输出 PNG/WebP |
| 产品白底图抠透明 | 背景透明，输出 PNG |
| 课件素材批处理 | 白边裁剪或背景透明，批量输出 |

## 打包说明

`build.ps1` 会自动：

- 优先使用项目里的 `venv_clean`
- 如果没有 `venv_clean`，则创建 `.venv`
- 安装/更新 `requirement.txt` 和 PyInstaller
- 收集 PyMuPDF、Pillow、Tkinter、pywin32 等运行所需依赖
- 输出单文件 exe 到 `dist/`

当前脚本默认不使用 UPX，以减少杀毒软件误报和原生扩展兼容问题；构建时会排除未使用的 AVIF、Pillow 字体/CMS/形态学扩展、Pythonwin UI、NumPy 延迟子包和网络 SSL 原生加速。`hashlib` 仍使用标准回退实现，本工具不发起 HTTPS 请求。单文件包仍包含 Python 运行时、Tkinter、Pillow、NumPy 核心、PyMuPDF 和 Windows COM 依赖，当前约 36 MB。

## 依赖

核心依赖：

- Pillow
- numpy
- PyMuPDF
- pywin32（仅 Windows PPT/WPS 连接需要）
- PyInstaller（仅打包需要，脚本会自动安装）

## 注意事项

- JPEG 不支持透明背景，透明处理请输出 PNG 或 WebP
- 动画 GIF/WebP 输出同格式、多页 TIFF 输出 TIFF 时会保留全部帧/页并使用联合裁剪框；转换为不兼容的单页格式会明确报错，不会静默丢帧
- 透明背景处理暂不支持多帧图片；程序会明确提示先拆分帧，不会只处理首帧
- 已有同名文件或多页输出目录不会被覆盖，程序会自动追加 `_2`、`_3`
- SVG 输入可用于裁剪和透明处理，但复杂 SVG 的渲染效果取决于 PyMuPDF
- PPT 导出 PDF/SVG 时保留文字和形状矢量；嵌入位图 DPI 可在界面中选择 300/450/600 或自行输入，默认 300
- PDF/SVG 矢量去白边不会重采样文字、矢量图或源 PDF 位图；普通 PDF 会同步硬裁 MediaBox。检测到 OCG/OCMD 可选图层时会保留 MediaBox、只写可逆 CropBox 并明确提示，避免隐藏内容被永久裁除
- 矢量输出的“边缘留白”统一按 300 DPI 换算像素，输入 2 表示约保留 2 px，输入 0 表示不主动留白
- PDF/PPT 页面尺寸使用 point，规范固定为 1 英寸 = 72 point；程序会结合每张图片的显示变换和所选 DPI 自动计算像素上限
- 打包时如果提示 exe 正在运行，请关闭已打开的程序后重试
- OneDrive/SharePoint PPT 没有本地源目录时，默认输出到桌面；也可手动选择输出目录
- PPT 原始图片无法可靠匹配或存在歧义时会明确警告并保留 PowerPoint 导出图像，不会强行替换成可能错误的媒体
- Mac 下 PPT 自动控制需要授予终端/脚本控制 PowerPoint 的权限；当前页会从整份 PDF 精确抽取，位图输出按所选 DPI 重新渲染

## 项目结构

```text
├── main.py          # 主入口与版本号
├── ui.py            # Tkinter 用户界面
├── processor.py     # 批处理与格式输出逻辑
├── detector.py      # 白边检测算法
├── units.py         # point/DPI 尺寸单位换算
├── transparency.py  # 背景透明算法
├── controllers.py   # PPT/PDF/SVG/图片读取控制器
├── build.ps1        # Windows 一键打包脚本
├── requirement.txt  # 运行依赖
└── README.md        # 项目说明
```

## License

MIT
