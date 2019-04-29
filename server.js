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
      console.log(addressList);

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

      /*var geocodedx = [];
      var geocodedy = [];
      for (var j = 0; j < addressList.length; j++) {
        geocoder.geocode(addressList[j], function(err, res) {
          tempx = [];
          tempy = [];
          if (res != undefined) {
            geocodedx.push(res[0].latitude);
            geocodedy.push(res[0].longitude);
          } else {
            geocodedx.push("undefined");
            geocodedy.push("undefined");
          }
          if (j == addressList.length) {
            geocodedx = tempx;
            geocodedy = tempy;
          }
        });
      }
      console.log(geocodedx);
      console.log(geocodedy);*/

      /*worksheet.getColumn("X").values = addressList;
      worksheet.commit();*/

      /*geocoder.batchGeocode(addressList, function(err, res) {
        console.log(res);
      });*/
    });
  }
});
