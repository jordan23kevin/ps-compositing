// 通用贴图脚本（正图/背图共用）- duplicate方式，与背图完全一致
var torsoFile = new File("{{TORSO_PATH}}");
var designFile = new File("{{DESIGN_PATH}}");
var outputFile = new File("{{OUTPUT_PATH}}");

var targetCenterX = parseFloat("{{TARGET_CENTER_X}}");
var targetTopY = parseFloat("{{TARGET_TOP_Y}}");
var scalePercent = parseFloat("{{SCALE_PERCENT}}");
var rotationAngle = parseFloat("{{ROTATION}}");
var scaledTopY = parseFloat("{{SCALED_TOP_Y}}");
var scaledCenterX = parseFloat("{{SCALED_CENTER_X}}");

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

// 步骤1：缩放（百分比，相对于原始尺寸）
designLayer.resize(scalePercent, scalePercent, AnchorPosition.TOPLEFT);

// 步骤2：旋转
designLayer.rotate(rotationAngle);

// 获取缩放旋转后的图层实际bounds
var bounds = designLayer.bounds;
var layerLeft = bounds[0].value;
var layerTop = bounds[1].value;

// 计算需要移动的距离
var moveX = targetCenterX - (layerLeft + scaledCenterX);
var moveY = targetTopY - (layerTop + scaledTopY);

// 移动图层
designLayer.translate(moveX, moveY);

// 保存JPG
var jpgOptions = new JPEGSaveOptions();
jpgOptions.quality = 12;
doc.saveAs(outputFile, jpgOptions, true, Extension.LOWERCASE);

doc.close(SaveOptions.DONOTSAVECHANGES);
