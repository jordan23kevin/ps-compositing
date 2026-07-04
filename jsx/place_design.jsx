// 通用贴图脚本（正图/背图共用）
// 调用方已激活目标胚衣文档并打开设计图文档，JSX 只需 duplicate + 移动 + 旋转 + 保存
var designDocName = "{{DESIGN_DOC_NAME}}";
var outputFile = new File("{{OUTPUT_PATH}}");

var rotationAngle = parseFloat("{{ROTATION}}");
// Python算好的移动量（图层左上角从0,0移动到目标）
var moveX = parseFloat("{{MOVE_X}}");
var moveY = parseFloat("{{MOVE_Y}}");

app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.PIXELS;

// 使用调用方激活的胚衣文档
var doc = app.activeDocument;

// 使用已打开的设计图文档（Python已trim+缩放，无透明边距）
var designDoc = app.documents.getByName(designDocName);

// 用图层复制的方式，保留透明度
var designLayer = designDoc.artLayers[0].duplicate(doc);

// 移动到目标位置（图层左上角在0,0，移动moveX/moveY）
designLayer.translate(moveX, moveY);

// 旋转
designLayer.rotate(rotationAngle);

// 保存JPG
var jpgOptions = new JPEGSaveOptions();
jpgOptions.quality = 12;
doc.saveAs(outputFile, jpgOptions, true, Extension.LOWERCASE);

// 清理贴图图层，避免胚衣文档累积图层
designLayer.remove();
