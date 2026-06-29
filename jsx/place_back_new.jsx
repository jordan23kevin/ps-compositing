// 新的背面贴图脚本 - 修正单位和位移
var torsoFile = new File("{{TORSO_PATH}}");
var designFile = new File("{{DESIGN_PATH}}");
var outputFile = new File("{{OUTPUT_PATH}}");

// 目标位置
var targetCenterX = 680;
var targetTopY = 570;
var scalePercent = 30; // 30%
var rotationAngle = 1; // 向右1度

// ===== Python 传入的参数 =====
var scaledTopY = parseFloat("{{SCALED_TOP_Y}}");
var scaledCenterX = parseFloat("{{SCALED_CENTER_X}}");

// 设置单位为像素
app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.PIXELS;

// 打开胚衣
var doc = app.open(torsoFile);

// 打开设计图
var designDoc = app.open(designFile);

// 用图层复制的方式，保留透明度
var designLayer = designDoc.artLayers[0].duplicate(doc);

// 关闭设计图文档
designDoc.close(SaveOptions.DONOTSAVECHANGES);

// 步骤1：缩放 30%
designLayer.resize(scalePercent, scalePercent, AnchorPosition.TOPLEFT);

// 步骤2：旋转 1度（向右）
designLayer.rotate(rotationAngle);

// 获取缩放旋转后的图层实际bounds
var bounds = designLayer.bounds;
var layerLeft = bounds[0].value;
var layerTop = bounds[1].value;
var layerRight = bounds[2].value;
var layerBottom = bounds[3].value;

// 计算需要移动的距离
var currentCenterX = layerLeft + scaledCenterX;
var currentTopY = layerTop + scaledTopY;
var moveX = targetCenterX - currentCenterX;
var moveY = targetTopY - currentTopY;

// 调试信息
var logFile = new File("C:/Users/Administrator/AppData/Local/Temp/ps_debug.log");
logFile.open("w");
logFile.writeln("=== PS 调试信息 ===");
logFile.writeln("layerLeft: " + layerLeft);
logFile.writeln("layerTop: " + layerTop);
logFile.writeln("layerRight: " + layerRight);
logFile.writeln("layerBottom: " + layerBottom);
logFile.writeln("scaledTopY: " + scaledTopY);
logFile.writeln("scaledCenterX: " + scaledCenterX);
logFile.writeln("currentCenterX: " + currentCenterX);
logFile.writeln("currentTopY: " + currentTopY);
logFile.writeln("moveX: " + moveX);
logFile.writeln("moveY: " + moveY);
logFile.close();

// 移动图层
designLayer.translate(moveX, moveY);

// 保存JPG
var jpgOptions = new JPEGSaveOptions();
jpgOptions.quality = 12;
doc.saveAs(outputFile, jpgOptions, true, Extension.LOWERCASE);

// 关闭文档
doc.close(SaveOptions.DONOTSAVECHANGES);
