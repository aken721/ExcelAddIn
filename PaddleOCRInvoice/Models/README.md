# PaddleOCR 模型文件目录

## 目录结构

请将下载的模型文件放置在此目录下，结构如下：

```
Models/
├── det_infer/                      # 检测模型（通用文件夹名，便于版本升级）
│   ├── inference.pdmodel
│   ├── inference.pdiparams
│   └── inference.pdiparams.info
├── rec_infer/                      # 识别模型（通用文件夹名，便于版本升级）
│   ├── inference.pdmodel
│   ├── inference.pdiparams
│   └── inference.pdiparams.info
├── cls_infer/                      # 方向分类模型（通用文件夹名，便于版本升级）
│   ├── inference.pdmodel
│   ├── inference.pdiparams
│   └── inference.pdiparams.info
└── ppocr_keys.txt                  # 字典文件
```

> **为什么使用通用文件夹名称？**
> 
> 使用不带版本号的通用名称（`det_infer`、`rec_infer`、`cls_infer`）有以下好处：
> - ✅ **无缝升级**：从 v4 升级到 v5 只需替换文件，无需修改代码
> - ✅ **简化配置**：不用担心版本号变化导致路径错误
> - ✅ **统一管理**：所有版本使用相同的目录结构

## 模型下载

### 方式1: 使用自动下载脚本（推荐）

```powershell
cd Models
.\下载模型文件.ps1
```

脚本会自动下载 PP-OCRv4 模型并重命名为通用文件夹名称。

### 方式2: 手动下载

详细步骤请参考：[手动下载指南.md](./手动下载指南.md)

**当前使用模型版本：PP-OCRv4**

| 模型 | 下载地址 | 大小 | 重命名为 |
|------|---------|------|----------|
| 检测模型 | [ch_PP-OCRv4_det_infer.tar](https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_det_infer.tar) | ~4.5 MB | `det_infer` |
| 识别模型 | [ch_PP-OCRv4_rec_infer.tar](https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_rec_infer.tar) | ~11 MB | `rec_infer` |
| 分类模型 | [ch_ppocr_mobile_v2.0_cls_infer.tar](https://paddleocr.bj.bcebos.com/dygraph_v2.0/ch/ch_ppocr_mobile_v2.0_cls_infer.tar) | ~1.4 MB | `cls_infer` |
| 字典文件 | [ppocr_keys_v1.txt](https://raw.githubusercontent.com/PaddlePaddle/PaddleOCR/main/ppocr/utils/ppocr_keys_v1.txt) | ~120 KB | `ppocr_keys.txt` |

## 模型说明

### PP-OCRv4 模型特点

- ✅ **高精度**：在通用场景下识别准确率高
- ✅ **服务端模型**：适用于服务器端部署，精度优先
- ✅ **支持中英文**：同时支持中文和英文识别
- ✅ **工业级**：可用于实际生产环境
- ✅ **成熟稳定**：经过大量实际应用验证

### 文件大小参考

- 检测模型：~4.5 MB (解压后 ~13 MB)
- 识别模型：~11 MB (解压后 ~32 MB)
- 分类模型：~1.4 MB (解压后 ~4 MB)
- 字典文件：~120 KB

**总计约**：17-50 MB

## 快速下载命令

### Windows PowerShell

```powershell
cd Models

# 使用脚本自动下载
.\下载模型文件.ps1
```

### Linux/macOS

```bash
cd Models

# 下载检测模型
wget https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_det_infer.tar
tar -xf ch_PP-OCRv4_det_infer.tar
mv ch_PP-OCRv4_det_infer det_infer

# 下载识别模型
wget https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_rec_infer.tar
tar -xf ch_PP-OCRv4_rec_infer.tar
mv ch_PP-OCRv4_rec_infer rec_infer

# 下载分类模型
wget https://paddleocr.bj.bcebos.com/dygraph_v2.0/ch/ch_ppocr_mobile_v2.0_cls_infer.tar
tar -xf ch_ppocr_mobile_v2.0_cls_infer.tar
mv ch_ppocr_mobile_v2.0_cls_infer cls_infer

# 下载字典文件
wget https://raw.githubusercontent.com/PaddlePaddle/PaddleOCR/main/ppocr/utils/ppocr_keys_v1.txt -O ppocr_keys.txt

echo "模型文件下载完成！"
```

## 版本升级指南

### 从 PP-OCRv4 升级到 PP-OCRv5（未来）

当 PP-OCRv5 发布后，升级步骤如下：

1. 下载新版本模型
2. 解压并重命名为通用文件夹名称
3. 替换 Models 目录下的对应文件夹
4. **无需修改任何代码**

```powershell
# 示例：升级检测模型
cd Models

# 下载 v5 模型
Invoke-WebRequest -Uri "https://paddleocr.bj.bcebos.com/PP-OCRv5/chinese/ch_PP-OCRv5_det_infer.tar" -OutFile "det_v5.tar"

# 解压
tar -xf det_v5.tar

# 删除旧模型
Remove-Item det_infer -Recurse -Force

# 重命名并使用新模型
Rename-Item -Path "ch_PP-OCRv5_det_infer" -NewName "det_infer"
```

## 验证安装

下载完成后，运行以下命令验证：

```powershell
# 检查必需文件
Get-ChildItem Models -Recurse -Filter "inference.pdmodel"
Get-ChildItem Models -Filter "ppocr_keys.txt"
```

应该看到 3 个 `.pdmodel` 文件和 1 个 `.txt` 文件。

正确的目录结构：

```
Models/
├── README.md (本文件)
├── 手动下载指南.md
├── 下载模型文件.ps1
├── det_infer/
│   ├── inference.pdmodel
│   ├── inference.pdiparams
│   └── inference.pdiparams.info
├── rec_infer/
│   ├── inference.pdmodel
│   ├── inference.pdiparams
│   └── inference.pdiparams.info
├── cls_infer/
│   ├── inference.pdmodel
│   ├── inference.pdiparams
│   └── inference.pdiparams.info
└── ppocr_keys.txt
```

## 常见问题

### Q: 模型文件下载很慢怎么办？

A: 可以尝试：
1. 使用下载工具（如 IDM、迅雷等）
2. 从 GitHub Releases 下载
3. 使用 Gitee 镜像（如果有）
4. 从百度网盘下载（查看官方文档）

### Q: 解压出错怎么办？

A: 
1. 确保下载的文件完整（检查文件大小）
2. Windows 使用 7-Zip 或系统自带的 tar 命令
3. Linux/macOS 使用 `tar -xf` 命令

### Q: 能否使用其他版本的模型？

A: 可以，但需要注意：
1. 确保模型版本兼容（推荐使用 PP-OCRv3 或 v4）
2. 下载后务必重命名为通用文件夹名称
3. 测试识别效果是否满足需求

### Q: 为什么不直接使用版本号命名？

A: 使用通用名称的好处：
- 代码中无需硬编码版本号
- 升级模型时无需修改代码
- 便于团队协作和部署
- 简化配置文件管理

## 参考资料

- **PaddleOCR 官方仓库**：https://github.com/PaddlePaddle/PaddleOCR
- **模型列表**：https://github.com/PaddlePaddle/PaddleOCR/blob/main/doc/doc_ch/models_list.md
- **PP-OCR 技术文档**：https://github.com/PaddlePaddle/PaddleOCR/blob/main/doc/doc_ch/ppocr_introduction.md
- **在线文档**：https://paddlepaddle.github.io/PaddleOCR/

---

**最后更新**：2024-11-07  
**当前模型版本**：PP-OCRv4
