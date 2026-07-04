// 通用贴图脚本（正图/背图共用）
// Python预先trim透明边距 + 缩放贴图，保存为临时PNG
// JSX打开胚衣和设计图，duplicate后贴图、保存
var torsoFile = new File("{{TORSO_PATH}}");
var designFile = new File("{{DESIGN_PATH}}");
var outputFile = new File("{{OUTPUT_PATH}}");

var rotationAngle = parseFloat("{{ROTATION}}");
// Python算好的移动量（图层左上角从0,0移动到目标）
var moveX = parseFloat("{{MOVE_X}}");
var moveY = parseFloat("{{MOVE_Y}}");

app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.PIXELS;

// 打开胚衣
var doc = app.open(torsoFile);

// 打开设计图（Python已trim+缩放，无透明边距）
var designDoc = app.open(designFile);

// 用图层复制的方式，保留透明度
var designLayer = designDoc.artLayers[0].duplicate(doc);

// 关闭设计图文档
designDoc.close(SaveOptions.DONOTSAVECHANGES);

// 移动到目标位置（图层左上角在0,0，移动moveX/moveY）
designLayer.translate(moveX, moveY);

// 旋转
designLayer.rotate(rotationAngle);

// 保存JPG
var jpgOptions = new JPEGSaveOptions();
jpgOptions.quality = 12;
doc.saveAs(outputFile, jpgOptions, true, Extension.LOWERCASE);

doc.close(SaveOptions.DONOTSAVECHANGES);
