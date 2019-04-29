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

var myapi = "Enter Google API Key Here";
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
      res.redirect("/");
    });
  });
});

// HERE I AM READING THE COLUMN 6TH COLUMN IN THE WORKSHEET WHICH IS HOLDING THE ADDRESSES,
// LOOPING THROUGH THEM, AND GEOCODING THEM. THE RESULTING X AND Y ARE BEING SAVED IN
// TEMPX AND TEMPY ARRAYS. I THEN SAVE THE

fs.readdir("./public/files", function(err, items) {
  if (items.length > 0) {
    mypath = "./public/files/" + items[0];
    var workbook = new Excel.Workbook();

    workbook.xlsx.readFile(mypath).then(function() {
      var worksheet = workbook.getWorksheet(1);
      var row = worksheet.getRow(1);
      var addressColumn = worksheet.getColumn(6);

      var addressList = [];
      addressColumn.eachCell(function(cell, rowNumber) {
        addressList.push(cell.value.result);
      });

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
          if (res != undefined) {
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
