const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const Excel = require("exceljs")
const path = require("path")

const app = express();
app.set("view engine", "ejs");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static("public"));
app.use(fileUpload());

app.listen(process.env.PORT || 3003, () => {
  console.log("listening on 3003");
});

app.get("/", (req, res) => {
  res.render("index.ejs");
});

app.post("/fileupload", function(req, res) {
  console.log(req.files);
  if (Object.keys(req.files).length == 0) {
    res.redirect("/fileupload");
  }
  mylength = 0;
  fs.readdir("./public/files", function(err, items) {
    mylength = items.length;
    let sampleFile = req.files.sampleFile;
    let sampleFileExt = path.extname(sampleFile.name)
    sampleFile.mv("./public/files/" + mylength + sampleFileExt , function(err) {
      if (err) return res.status(500).send(err);
      res.redirect("/");
    });
  });
});

fs.readdir("./public/files", function(err, items) {
  mypath = "./public/files/" + items[0];
  var workbook = new Excel.Workbook();
  workbook.xlsx.readFile(mypath)
    .then(function() {
        var worksheet = workbook.getWorksheet(1)
        var row = worksheet.getRow(1)
        for(var i = 1; i<worksheet.columnCount + 1; i++){
          console.log(i)
          console.log(row.getCell(i).value)
        }
    });
});
