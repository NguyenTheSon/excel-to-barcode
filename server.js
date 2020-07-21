const path = require("path")
const fs = require('fs');
const JsBarcode = require('jsbarcode');
const { DOMImplementation, XMLSerializer } = require('xmldom');
const xmlSerializer = new XMLSerializer();
const document = new DOMImplementation().createDocument('http://www.w3.org/1999/xhtml', 'html', null);
const svgNode = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
var XLSX = require('xlsx');

var workbook = XLSX.readFile('Serials300box20072020.xls');
var sheet_name_list = workbook.SheetNames;
var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

for (let barcode of xlData) {
  JsBarcode(svgNode, barcode['Key'], {
    xmlDocument: document,
    // displayValue: false,
    textPosition: "top",
    width: 3.31,
    height: 83
  });

  let xml = xmlSerializer.serializeToString(svgNode);
  if (!fs.existsSync('./svg/')) {
    fs.mkdirSync('./svg/', { recursive: true });
  }
  fs.writeFile(`./svg/${barcode['Key']}.svg`, xml, (err) => {
    if (err) throw err;
    console.log(`SVG written! ${barcode['Key']}`);
  });
}
