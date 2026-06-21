#target photoshop
var doc = app.activeDocument;
if (doc) {
    var file = new File("D:/Semems/1AI/DX0001/白BW.jpg");
    if (file.exists) file.remove();
    var opts = new JPEGSaveOptions();
    opts.quality = 12;
    opts.embedColorProfile = true;
    doc.saveAs(file, opts, true);
}
