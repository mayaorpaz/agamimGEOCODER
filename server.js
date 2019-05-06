const express = require("express");
const bodyParser = require("body-parser");
const fileUpload = require("express-fileupload");
const fs = require("fs");
const Excel = require("exceljs");
const path = require("path");
const NodeGeocoder = require("node-geocoder");
const async = require("async");
const flash = require("express-flash-messages");
var session = require("express-session");

const app = express();
app.set("view engine", "ejs");

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static("public"));
app.use(fileUpload());
app.use(flash());
app.use(
  session({
    cookie: { maxAge: 60000 },
    secret: "woot",
    resave: false,
    saveUninitialized: false
  })
);

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

// ENTER YOUR GOOGLE API KEY HERE
var myapi = "ENTER YOUR GOOGLE API KEY HERE";

var bigcolumn = 1

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
      var workbook = new Excel.Workbook();

      workbook.xlsx.readFile("./public/" + finalpath).then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(1);
        var addresscolumn = worksheet.getColumn(bigcolumn)
        var addressvalues = addresscolumn.values
        let xcolumn;
        let ycolumn;
        myrows = [];
        for (var i = 0; i < row.values.length; i++) {
          if (row.values[i] != undefined) myrows.push(i);
          if (row.values[i] == "Latitude"){
            xcolumn = worksheet.getColumn(i)
            ycolumn = worksheet.getColumn(i + 1)
          }
        }
        var xvalues = xcolumn.values
        var yvalues = ycolumn.values
        cells = [];
        /*for (var i = 1; i < myrows.length + 1; i++) {
          for (var j = 1; j < worksheet.getRow(myrows[i]).values.length; j++) {
            if (typeof worksheet.getRow(myrows[i]).values[j] == "object") {
              cells.push(worksheet.getRow(myrows[i]).values[j].result);
            } else {
              if (worksheet.getRow(myrows[i]).values[j] != undefined) {
                cells.push(worksheet.getRow(myrows[i]).values[j]);
              }
            }
          }
        }*/
        for(var i = 2; i<addressvalues.length; i++){
          cells.push(addressvalues[i])
          cells.push(xvalues[i])
          cells.push(yvalues[i])
          cells.push(xvalues[i] + ", " + yvalues[i])
        }
        var addressColumn = worksheet.getColumn(6);
        filename = worksheet.name;
        columns = row.values;
        columnCount = worksheet.columnCount;
        //console.log(columns);
        res.render("done.ejs", {
          columns,
          columnCount,
          filename,
          alphabet,
          cells
        });
      });
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
        var column1 = worksheet.getColumn(1).values;
        totalrows = 0;
        for (var k = 0; k < column1.length; k++) {
          if (column1[k] != undefined) {
            totalrows++;
          }
        }
        if (totalrows > 10000) {
          res.render("error.ejs", {
            myerror: "MAXIMUM ROW COUNT: 10000; YOUR ROW COUNT: " + totalrows
          });
        } else {
          cells = [];
          for (var i = 2; i < worksheet.rowCount + 1; i++) {
            for (var j = 1; j < worksheet.getRow(i).values.length; j++) {
              if (typeof worksheet.getRow(i).values[j] == "object") {
                cells.push(worksheet.getRow(i).values[j].result);
              } else {
                cells.push(worksheet.getRow(i).values[j]);
              }
            }
          }
          var addressColumn = worksheet.getColumn(6);
          filename = worksheet.name;
          columns = row.values;
          columnCount = worksheet.columnCount;
          //console.log(columns);
          res.render("select.ejs", {
            columns,
            columnCount,
            filename,
            alphabet,
            cells
          });
        }
      });
    }
  });
});

app.post("/fileupload", function(req, res) {
  console.log("FILENAME: " + req.files.sampleFile.name.slice(0, -5));
  if (Object.keys(req.files).length == 0) {
    res.redirect("/fileupload");
  }
  mylength = 0;
  fs.readdir("./public/files", function(err, items) {
    mylength = items.length;
    let sampleFile = req.files.sampleFile;
    let sampleFileExt = path.extname(sampleFile.name);
    if (sampleFile.size > 50000000) {
      res.render("error.ejs", { myerror: "Maximum file size: 50mb" });
    } else {
      if (sampleFileExt == ".xlsx") {
        sampleFile.mv(
          "./public/files/" +
            new Date().getTime() +
            "-" +
            req.files.sampleFile.name.replace(/\s/g, ""),
          function(err) {
            if (err) return res.status(500).send(err);
            res.redirect("/select");
          }
        );
      } else {
        res.render("error.ejs", { myerror: "File extension must be .XLSX" });
      }
    }
  });
});

app.post("/geocodeone", (req, res) => {
    geocoder.geocode(req.body.col, function(err, geocoded) {
      if (geocoded != undefined && geocoded[0] != undefined) {
        console.log(geocoded[0].latitude)
        console.log(geocoded[0].longitude)
        console.log(geocoded[0].formattedAddress)
        req.flash('notify1', geocoded[0].formattedAddress)
        req.flash('notify2', geocoded[0].latitude + ", " + geocoded[0].longitude)
        res.redirect('/')
      } else {
        req.flash('error', "Error. Try again.");
        res.redirect('/')
      }
    })
});

app.post("/geocode", function(req, res) {
  handler = 0;
  fs.readdir("./public/files", function(err, items) {
    if (items.length > 0) {
      mypath = "./public/files/" + items[items.length - 1];
      console.log(mypath);
      var workbook = new Excel.Workbook();

      workbook.xlsx.readFile(mypath).then(function() {
        var worksheet = workbook.getWorksheet(1);
        var row = worksheet.getRow(1);
        var addressColumn = worksheet.getColumn(parseInt(req.body.col));
        bigcolumn = parseInt(req.body.col)
        filename = worksheet.name;
        columns = row.values;
        columnCount = worksheet.columnCount;
        addressList = [];
        if (addressColumn.values[2].result != undefined) {
          addressColumn.eachCell(function(cell, rowNumber) {
            addressList.push({ address: cell.result, row: rowNumber });
          });
        } else {
          addressColumn.eachCell(function(cell, rowNumber) {
            addressList.push({ address: cell.value, row: rowNumber });
          });
        }
        console.log(addressList.slice(0, 10));

        var tempx = [];
        var tempy = [];
        async.each(
          addressList.slice(1),
          function(addr, callback) {
            if (addr.address != undefined) {
              geocoder.geocode(addr.address, function(err, geocoded) {
                if (err) {
                  callback(err);
                  return;
                }
                if (geocoded) {
                  if (geocoded != undefined && geocoded[0] != undefined) {
                    tempx.push({
                      resultx: geocoded[0].latitude,
                      row: addr.row
                    });
                    tempy.push({
                      resulty: geocoded[0].longitude,
                      row: addr.row
                    });
                    callback();
                  } else {
                    handler = 1;
                  }
                }
              });
            }
          },
          function(err) {
            console.log("Handler: " + handler);
            if (err) {
              console.log(err);
            }
            if (handler == 1) {
              handler = 2;
              console.log("Handler: " + handler);
              res.render("error.ejs", {
                myerror:
                  "Could not geocode this column accurately. Please make sure it contains a full address including CITY."
              });
            } else {
              console.log(tempx.length);
              console.log(tempy.length);
              if (tempx.length == 0 || (tempy.length == 0 && handler == 0)) {
                res.render("error.ejs", {
                  myerror:
                    "There was an error geocoding this column. Please try again."
                });
              }
              xcol = worksheet.getColumn(columnCount + 1);
              ycol = worksheet.getColumn(columnCount + 2);
              xcol.header = "Latitude";
              ycol.header = "Longitude";
              xcol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
                for (var i = 0; i < tempx.length; i++) {
                  if (tempx[i].row == rowNumber) {
                    cell.value = tempx[i].resultx;
                  }
                }
              });
              ycol.eachCell({ includeEmpty: true }, function(cell, rowNumber) {
                for (var i = 0; i < tempy.length; i++) {
                  if (tempy[i].row == rowNumber) {
                    cell.value = tempy[i].resulty;
                  }
                }
              });
              fs.readdir("./public/completed", function(err, items) {
                placeholder = 0;
                if (items != undefined) {
                  placeholder = items.length;
                }
                mypath2 =
                  "./public/completed/" + new Date().getTime() + ".xlsx";
                workbook.xlsx.writeFile(mypath2).then(function() {
                  console.log(mypath2 + " -- file is written.");
                  prerow = [];
                  getme = worksheet.getColumn(1).values;
                  totalrows = getme.length - 2;
                  geocodedrows = tempx.length;
                  fractiontotal = geocodedrows / totalrows;
                  total = Math.floor(fractiontotal * 100);
                  req.flash("notify", "Geocoded " + total + "% of addresses!");
                  res.redirect("/done");
                });
              });
            }
          }
        );
      });
    }
  });
});
