const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const app = express();
const readXlsxFile = require("read-excel-file/node");

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
  let sampleFile = req.files.sampleFile;
  sampleFile.mv("./public/files/" + req.files.sampleFile.name, function(err) {
    if (err) return res.status(500).send(err);
    res.redirect("/");
  });
});

fs.readdir("./public/files", function(err, items) {
  mypath = "./public/files/" + items[0];
  console.log(mypath);
  readXlsxFile(fs.createReadStream(mypath)).then(rows => {
    console.log(rows[0]);
  });
});
