const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const Excel = require("exceljs");
const path = require("path");
const NodeGeocoder = require("node-geocoder");
const async = require("async");

const app = express();
app.set("view engine", "ejs");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static("public"));
app.use(fileUpload());

// ENTER YOUR GOOGLE API KEY HERE
var myapi = "AIzaSyDkG702RFFEEm08CP87sLK_amm-ru_eUVs";

var options = {
  provider: "google",
  apiKey: myapi
};

var geocoder = NodeGeocoder(options);

app.listen(process.env.PORT || 3003, () => {
  console.log("listening on 3003");
});

app.get("/", (req, res) => {
  res.render("index.ejs");
});

app.get("/error", (req, res) => {
  res.render("error.ejs");
});

app.get("/done", (req, res) => {
  fs.readdir("./public/completed", function(err, items) {
    if (items.length > 0) {
      finalpath = "completed/" + items[items.length - 1];
      console.log(finalpath);
      res.render("done.ejs", { finalpath });
    }
  });
});

app.get("/select", (req, res) => {
  fs.readdir("./public/files", function(err, items) {
    if (err) {
      res.redirect("/error");
    }
    if (items.length > 0) {
      mypath = "./public/files/" + items[items.length - 1];
      console.log(mypath);
      console.log(path.extname(mypath));
      if (path.extname(mypath) != ".xlsx") {
        myerror = "FILE MUST BE .XLSX";
        res.render("error.ejs", { myerror });
      }
      var workbook = new Excel.Workbook();

      workbook.xlsx.readFile(mypath).then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(1);
        var addressColumn = worksheet.getColumn(6);
        filename = worksheet.name;
        columns = row.values;
        columnCount = worksheet.columnCount;
        alphabet = [
          "A",
          "B",
          "C",
          "D",
          "E",
          "F",
          "G",
          "H",
          "I",
          "J",
          "K",
          "L",
          "M",
          "N",
          "O",
          "P",
          "Q",
          "R",
          "S",
          "T",
          "U",
          "V",
          "W",
          "X",
          "Y",
          "Z",
          "AA",
          "AB",
          "AC",
          "AD",
          "AE",
          "AF",
          "AG",
          "AH",
          "AI",
          "AJ",
          "AK",
          "AL",
          "AM",
          "AN",
          "AO",
          "AP",
          "AQ",
          "AR",
          "AS",
          "AT",
          "AU",
          "AV",
          "AW",
          "AX",
          "AY",
          "AZ"
        ];
        //console.log(columns);
        res.render("select.ejs", { columns, columnCount, filename, alphabet });
      });
    }
  });
});

app.post("/fileupload", function(req, res) {
  console.log("FILENAME: " + req.files.sampleFile.name);
  if (Object.keys(req.files).length == 0) {
    res.redirect("/fileupload");
  }
  mylength = 0;
  fs.readdir("./public/files", function(err, items) {
    mylength = items.length;
    let sampleFile = req.files.sampleFile;
    let sampleFileExt = path.extname(sampleFile.name);
    if (sampleFileExt == ".xlsx") {
      sampleFile.mv("./public/files/" + mylength + sampleFileExt, function(
        err
      ) {
        if (err) return res.status(500).send(err);
        res.redirect("/select");
      });
    } else {
      res.render("error.ejs", { myerror: "File extension must be .XLSX" });
    }
  });
});

app.post("/geocode", function(req, res) {
  fs.readdir("./public/files", function(err, items) {
    if (items.length > 0) {
      mypath = "./public/files/" + items[items.length - 1];
      console.log(mypath);
      var workbook = new Excel.Workbook();

      workbook.xlsx.readFile(mypath).then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(1);
        var addressColumn = worksheet.getColumn(parseInt(req.body.col));
        filename = worksheet.name;
        columns = row.values;
        columnCount = worksheet.columnCount;
        addressList = [];
        if (addressColumn.values[3].result != undefined) {
          addressColumn.eachCell(function(cell, rowNumber) {
            addressList.push(cell.result);
          });
        } else {
          addressList = addressColumn.values;
        }
        console.log(addressList.slice(0, 10));

        var tempx = [];
        var tempy = [];
        async.each(
          addressList,
          function(addr, callback) {
            //console.log(addr);
            setTimeout(function() {
              if (addr != undefined) {
                geocoder.geocode(addr, function(err, geocoded) {
                  if (err) {
                    callback(err);
                    return;
                  }
                  if (geocoded) {
                    if (geocoded != undefined) {
                      tempx.push(geocoded[0].latitude);
                      tempy.push(geocoded[0].longitude);
                      callback();
                    } else {
                      tempx.push("undefined");
                      tempy.push("undefined");
                    }
                  }
                });
              }
            }, 50);
          },
          function(err) {
            if (err) {
              console.log(err);
            }

            console.log(tempx.length);
            console.log(tempy.length);
            tempy.length = 0;
            if (tempx.length == 0 || tempy.length == 0) {
              res.render("error.ejs", {
                myerror:
                  "There was an error geocoding this column. Please try again."
              });
            }
            worksheet.getColumn("X").values = tempx;
            worksheet.getColumn("Y").values = tempy;
            fs.readdir("./public/completed", function(err, items) {
              placeholder = 0;
              if (items != undefined) {
                placeholder = items.length;
              }
              mypath2 = "./public/completed/" + placeholder + ".xlsx";
              workbook.xlsx.writeFile(mypath2).then(function() {
                console.log(mypath2 + " -- file is written.");
                res.redirect("/done");
              });
            });
          }
        );
      });
    }
  });
});
