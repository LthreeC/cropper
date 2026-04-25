# 白边裁剪工具 v0.3

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

打包 Windows exe：

```powershell
.\build.ps1
```

打包结果：

```text
dist/cropper-v0.3.exe
```

如果要覆盖已有 exe，请先关闭正在运行的 `cropper-v0.3.exe`。

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
| PDF | 矢量输出，适合文档和打印 |
| SVG | 矢量输出，适合网页、设计软件和后续编辑 |
| PNG | 无损位图，支持透明 |
| TIFF | 无损位图，适合印刷或归档 |
| JPEG | 有损压缩，适合照片和分享 |
| WebP | 体积较小，适合网页 |

检测参数：

| 参数 | 说明 |
|------|------|
| 智能 | 综合亮度、颜色变化和边缘，推荐默认使用 |
| 简单 | 仅按亮度判断，速度最快，适合纯白背景 |
| 边缘敏感 | 更重视浅色边缘，适合内容接近白色的图 |
| 白色阈值 | 越高越只裁纯白，越低越容易裁掉浅灰 |
| 敏感度 | 越高越保守，越低裁剪越激进 |
| 边缘留白 | 裁剪后额外保留的像素边距 |
| DPI | 仅影响 PNG/TIFF/JPEG/WebP 等位图输出 |

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

当前脚本默认不使用 UPX，因此 exe 会比旧版大一些，但兼容性更稳。单文件包内包含 Python 运行时、Tkinter、Pillow、numpy、PyMuPDF 和 Windows COM 依赖，几十 MB 属于正常范围。

## 依赖

核心依赖：

- Pillow
- numpy
- PyMuPDF
- pywin32（仅 Windows PPT/WPS 连接需要）
- PyInstaller（仅打包需要，脚本会自动安装）

## 注意事项

- JPEG 不支持透明背景，透明处理请输出 PNG 或 WebP
- SVG 输入可用于裁剪和透明处理，但复杂 SVG 的渲染效果取决于 PyMuPDF
- PDF/SVG 矢量输出不受 DPI 影响，位图输出才受 DPI 影响
- 打包时如果提示 exe 正在运行，请关闭已打开的程序后重试
- Mac 下 PPT 自动控制需要授予终端/脚本控制 PowerPoint 的权限

## 项目结构

```text
├── main.py          # 主入口与版本号
├── ui.py            # Tkinter 用户界面
├── processor.py     # 批处理与格式输出逻辑
├── detector.py      # 白边检测算法
├── transparency.py  # 背景透明算法
├── controllers.py   # PPT/PDF/SVG/图片读取控制器
├── build.ps1        # Windows 一键打包脚本
├── requirement.txt  # 运行依赖
└── README.md        # 项目说明
```

## License

MIT
