var doc = app.activeDocument;
var file = new File("D:/Semems/1AI/DX0001/白BW.jpg");
var opts = new JPEGSaveOptions();
opts.quality = 12;
opts.embedColorProfile = true;
doc.saveAs(file, opts, true, Extension.LOWERCASE);
alert("保存完成: 白BW.jpg");
