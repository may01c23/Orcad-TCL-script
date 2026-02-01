# Orcad SCH 自动化脚本集 (Orcad-TCL-Scripts)

本仓库包含一组用于Orcad 原理图设计的 TCL 脚本，旨在通过自动化操作减少重复劳动，提升设计准确性。

## 🛠 功能模块

### 1. 格式刷工具 (`FormatPainter.tcl`)

**功能：** 能够抓取选定参考对象的特定属性（如字体大小、层信息、颜色等）。

### 2. NC 标记工具 (`NCMarker.tcl`)

**功能：** 批量给元器件添加NC标记。

### 3. 网络名放置助手 (`NetPlacement.tcl`)

**功能：** 批量按照规则放置网络名

---

## 🚀 安装与加载

### 加载脚本

在 Cadence 的控制台（Command Console）中，使用以下命令加载：

```tcl
source <脚本的绝对路径>/FileName.tcl

```

---

## 📂 文件清单

| 脚本名称 | 主要用途 | 适用对象 |
| --- | --- | --- |
| **FormatPainter.tcl** | 属性同步/格式复制 | Text, Pins, Shapes |
| **NCMarker.tcl** | NC 节点自动处理 | Unconnected Pins |
| **NetPlacement.tcl** | 基于网络的组件布局 | Components, Footprints |

---

## ⚠️ 注意事项

* 在执行批量修改属性（如使用 `FormatPainter`）前，建议先保存当前设计快照。
* 请确保脚本运行环境具有足够的执行权限。

---

## 🤝 贡献
欢迎提交 Issue 或 Pull Request 来改进这些工具！
