var aspose = aspose || {};
aspose.cells = require("aspose.cells");

// Load workbook
var workbook = aspose.cells.Workbook("./public/excel.xls");

// Save XLSX as HTML
workbook.save("/xlsx-to-html.html");
var workbook = aspose.cells.Workbook("./public/excel.xls");
var saveOptions = aspose.cells.PdfSaveOptions();
saveOptions.setOnePagePerSheet(true);
// convert Excel to PDF
console.log(workbook)
workbook.save("ExceltoPDF.pdf", saveOptions);
var objExcel = objExcel.Workbooks.Open(docPath);
var objWorkbook = objExcel.Workbooks.Open('./public/excel.xls');

var wdFormatPdf = 57;
objWorkbook.SaveAs(pdfPath, wdFormatPdf);
objWorkbook.Close();