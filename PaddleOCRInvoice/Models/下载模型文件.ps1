# PaddleOCR 模型下载脚本
# 更新日期: 2024-11-06

Write-Host "=== PaddleOCR 模型文件下载 ===" -ForegroundColor Green
Write-Host ""

# 模型下载配置 (使用通用文件夹名称，便于版本升级)
$models = @{
    "det" = @{
        "name" = "检测模型 (Detection Model)"
        "url" = "https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_det_infer.tar"
        "alt_url" = "https://github.com/PaddlePaddle/PaddleOCR/releases/download/v2.9/ch_PP-OCRv4_det_infer.tar"
        "size" = "约 4.5 MB"
        "original_folder" = "ch_PP-OCRv4_det_infer"
        "target_folder" = "det_infer"
    }
    "rec" = @{
        "name" = "识别模型 (Recognition Model)"
        "url" = "https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_rec_infer.tar"
        "alt_url" = "https://github.com/PaddlePaddle/PaddleOCR/releases/download/v2.9/ch_PP-OCRv4_rec_infer.tar"
        "size" = "约 11 MB"
        "original_folder" = "ch_PP-OCRv4_rec_infer"
        "target_folder" = "rec_infer"
    }
    "cls" = @{
        "name" = "方向分类模型 (Classification Model)"
        "url" = "https://paddleocr.bj.bcebos.com/dygraph_v2.0/ch/ch_ppocr_mobile_v2.0_cls_infer.tar"
        "alt_url" = "https://github.com/PaddlePaddle/PaddleOCR/releases/download/v2.7/ch_ppocr_mobile_v2.0_cls_infer.tar"
        "size" = "约 1.4 MB"
        "original_folder" = "ch_ppocr_mobile_v2.0_cls_infer"
        "target_folder" = "cls_infer"
    }
}

# 字典文件
$keysUrl = "https://raw.githubusercontent.com/PaddlePaddle/PaddleOCR/main/ppocr/utils/ppocr_keys_v1.txt"
$keysAltUrl = "https://gitee.com/paddlepaddle/PaddleOCR/raw/main/ppocr/utils/ppocr_keys_v1.txt"

# 创建临时目录
$tempDir = Join-Path $PSScriptRoot "temp"
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir | Out-Null
}

# 下载函数（支持重试）
function Download-File {
    param(
        [string]$Url,
        [string]$AltUrl,
        [string]$OutFile,
        [string]$Name
    )
    
    Write-Host "正在下载: $Name" -ForegroundColor Yellow
    Write-Host "  目标: $OutFile"
    
    # 尝试主URL
    try {
        Write-Host "  尝试主链接..." -ForegroundColor Gray
        Invoke-WebRequest -Uri $Url -OutFile $OutFile -UseBasicParsing
        Write-Host "  ✓ 下载成功！" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ✗ 主链接失败: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    # 尝试备用URL
    if ($AltUrl) {
        try {
            Write-Host "  尝试备用链接..." -ForegroundColor Gray
            Invoke-WebRequest -Uri $AltUrl -OutFile $OutFile -UseBasicParsing
            Write-Host "  ✓ 下载成功！" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "  ✗ 备用链接也失败: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    return $false
}

# 下载模型文件
Write-Host "步骤 1/3: 下载模型文件" -ForegroundColor Cyan
Write-Host "------------------------------------------------"

$downloadSuccess = $true

foreach ($key in $models.Keys) {
    $model = $models[$key]
    $tarFile = Join-Path $tempDir "$key.tar"
    
    Write-Host ""
    Write-Host "[$($model.name)] ($($model.size))" -ForegroundColor White
    
    if (Download-File -Url $model.url -AltUrl $model.alt_url -OutFile $tarFile -Name $model.name) {
        # 解压
        Write-Host "  正在解压..." -ForegroundColor Gray
        try {
            tar -xf $tarFile -C $PSScriptRoot
            
            # 重命名为通用文件夹名称
            $originalPath = Join-Path $PSScriptRoot $model.original_folder
            $targetPath = Join-Path $PSScriptRoot $model.target_folder
            
            if (Test-Path $originalPath) {
                # 如果目标文件夹已存在，先删除
                if (Test-Path $targetPath) {
                    Write-Host "  清理旧模型文件..." -ForegroundColor Gray
                    Remove-Item $targetPath -Recurse -Force
                }
                
                # 重命名为通用名称
                Rename-Item -Path $originalPath -NewName $model.target_folder
                Write-Host "  ✓ 已重命名为通用文件夹: $($model.target_folder)" -ForegroundColor Green
            }
            
            Write-Host "  ✓ 解压完成" -ForegroundColor Green
            Remove-Item $tarFile -Force
        }
        catch {
            Write-Host "  ✗ 解压失败: $($_.Exception.Message)" -ForegroundColor Red
            $downloadSuccess = $false
        }
    }
    else {
        $downloadSuccess = $false
        Write-Host "  ⚠ 跳过此文件" -ForegroundColor Yellow
    }
}

# 下载字典文件
Write-Host ""
Write-Host "步骤 2/3: 下载字典文件" -ForegroundColor Cyan
Write-Host "------------------------------------------------"
$keysFile = Join-Path $PSScriptRoot "ppocr_keys.txt"

if (Download-File -Url $keysUrl -AltUrl $keysAltUrl -OutFile $keysFile -Name "字典文件") {
    Write-Host ""
} else {
    Write-Host "  ⚠ 字典文件下载失败" -ForegroundColor Yellow
    $downloadSuccess = $false
}

# 清理临时目录
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}

# 验证文件
Write-Host ""
Write-Host "步骤 3/3: 验证模型文件" -ForegroundColor Cyan
Write-Host "------------------------------------------------"

$requiredFiles = @(
    "det_infer\inference.pdmodel",
    "rec_infer\inference.pdmodel",
    "cls_infer\inference.pdmodel",
    "ppocr_keys.txt"
)

$allExists = $true
foreach ($file in $requiredFiles) {
    $fullPath = Join-Path $PSScriptRoot $file
    if (Test-Path $fullPath) {
        Write-Host "  ✓ $file" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $file (缺失)" -ForegroundColor Red
        $allExists = $false
    }
}

# 总结
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
if ($allExists -and $downloadSuccess) {
    Write-Host "✓ 模型文件下载完成！" -ForegroundColor Green
    Write-Host ""
    Write-Host "下一步: 运行示例程序测试" -ForegroundColor Yellow
    Write-Host "  cd ..\PaddleOCRInvoice.Sample" -ForegroundColor Gray
    Write-Host "  dotnet run" -ForegroundColor Gray
} else {
    Write-Host "⚠ 部分文件下载失败" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "请尝试以下方案:" -ForegroundColor White
    Write-Host "  1. 手动下载模型文件（见下方链接）" -ForegroundColor Gray
    Write-Host "  2. 使用科学上网工具" -ForegroundColor Gray
    Write-Host "  3. 从国内镜像下载" -ForegroundColor Gray
    Write-Host ""
    Write-Host "官方模型列表:" -ForegroundColor White
    Write-Host "  https://github.com/PaddlePaddle/PaddleOCR/blob/main/doc/doc_ch/models_list.md" -ForegroundColor Cyan
    Write-Host "  https://paddlepaddle.github.io/PaddleOCR/main/ppocr/model_list.html" -ForegroundColor Cyan
}
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

