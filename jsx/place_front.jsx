// 放置正面设计（简化版 - 假设设计已是透明背景）
var torsoFile = new File("{{TORSO_PATH}}");
var designFile = new File("{{DESIGN_PATH}}");
var outputFile = new File("{{OUTPUT_PATH}}");

var targetHeight = parseFloat("{{TARGET_HEIGHT}}");
var centerX = parseFloat("{{CENTER_X}}");
var centerY = parseFloat("{{CENTER_Y}}");
var rotation = parseFloat("{{ROTATION}}");

// 打开T恤模板
var doc = app.open(torsoFile);

// 打开设计图
var designDoc = app.open(designFile);

// 全选并复制
designDoc.selection.selectAll();
designDoc.selection.copy();
designDoc.close(SaveOptions.DONOTSAVECHANGES);

// 粘贴到T恤文档
var designLayer = doc.paste();

// 调整大小
var originalWidth = designLayer.bounds[2] - designLayer.bounds[0];
var originalHeight = designLayer.bounds[3] - designLayer.bounds[1];
var scale = targetHeight / originalHeight * 100;
designLayer.resize(scale, scale, AnchorPosition.MIDDLECENTER);

// 移动到指定位置
var currentCenterX = (designLayer.bounds[0] + designLayer.bounds[2]) / 2;
var currentCenterY = (designLayer.bounds[1] + designLayer.bounds[3]) / 2;
var deltaX = centerX - currentCenterX;
var deltaY = centerY - currentCenterY;
designLayer.translate(deltaX, deltaY);

// 旋转
designLayer.rotate(rotation);

// 保存JPG
var jpgOptions = new JPEGSaveOptions();
jpgOptions.quality = 12;
doc.saveAs(outputFile, jpgOptions, true, Extension.LOWERCASE);

// 关闭文档
doc.close(SaveOptions.DONOTSAVECHANGES);
