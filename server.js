const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const Excel = require("exceljs");
const path = require("path");
const NodeGeocoder = require("node-geocoder");

const app = express();
app.set("view engine", "ejs");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static("public"));
app.use(fileUpload());

// ENTER YOUR GOOGLE API KEY HERE
var myapi = "AIzaSyDXY_aO0xDGZX4BSOkw8w88wLpp0Q7HTIQ";

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

app.get("/select", (req, res) => {
  fs.readdir("./public/files", function(err, items) {
    if (items.length > 0) {
      mypath = "./public/files/" + items[items.length - 1];
      console.log(mypath);
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
        console.log(columns);
        res.render("select.ejs", { columns, columnCount, filename, alphabet });
      });
    }
  });
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
    let sampleFileExt = path.extname(sampleFile.name);
    sampleFile.mv("./public/files/" + mylength + sampleFileExt, function(err) {
      if (err) return res.status(500).send(err);
      res.redirect("/select");
    });
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
        console.log(addressColumn.values);
      });
    }
  });
  console.log(req.body);
  res.redirect("/select");
});

// HERE I AM READING THE 6TH COLUMN IN THE WORKSHEET WHICH IS HOLDING THE ADDRESSES,
// LOOPING THROUGH THEM, AND GEOCODING THEM. THE RESULTING X AND Y ARE BEING SAVED IN
// TEMPX AND TEMPY ARRAYS. I THEN INSERT THEM BACK INTO THE SHEET IN NEW COLUMNS AND SAVE.

fs.readdir("./public/files", function(err, items) {
  if (items.length > 0) {
    mypath = "./public/files/" + items[items.length - 1];
    console.log(mypath);
    var workbook = new Excel.Workbook();

    workbook.xlsx.readFile(mypath).then(function() {
      var worksheet = workbook.getWorksheet(1);
      var row = worksheet.getRow(1);
      var addressColumn = worksheet.getColumn(6);

      //console.log(row.values);
      //console.log(worksheet.columnCount);

      var addressList = [];
      addressColumn.eachCell(function(cell, rowNumber) {
        addressList.push(cell.value.result);
      });

      console.log(addressList);

      // FAILED ATTEMPT AT PROMISE ASYNC AWAIT

      /*function myGeocode(myAddress) {
        return new Promise(function(resolve, reject) {
          let myobj;
          geocoder.geocode(myAddress, function(err, res) {
            if (res != undefined) {
              resolve(res[0].latitude);
            }
          });
        });
      }

      Promise.all(addressList.map(myGeocode))
        .then(function(geocodedList) {
          console.log(geocodedList);
        })
        .catch(function(error) {
          console.error(error);
        });*/

      var tempx = [];
      var tempy = [];
      for (var j = 0; j < addressList.length; j++) {
        geocoder.geocode(addressList[j], function(err, res) {
          //console.log(res);
          if (res != undefined) {
            console.log(res[0].latitude);
            tempx.push(res[0].latitude);
            tempy.push(res[0].longitude);
          } else {
            tempx.push("undefined");
            tempy.push("undefined");
          }
        });
      }
      console.log(tempx);
      console.log(tempy);

      // INSERTING TEMPX & TEMPY ARRAY INTO 'X' & 'Y' COLUMNS

      worksheet.getColumn("X").values = tempx;
      worksheet.getColumn("Y").values = tempy;

      // WRITING TO NEW FILE

      /*workbook.xlsx.writeFile("./public/files/0test.xlsx").then(function() {
        console.log("xlsx file is written.");
      });*/

      // BATCH GEOCODING ISN'T ALLOWED WITH GOOGLE API

      /*geocoder.batchGeocode(addressList, function(err, res) {
        console.log(res);
      });*/
    });
  }
});
