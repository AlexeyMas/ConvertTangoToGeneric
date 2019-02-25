var Excel = require('exceljs');
var fs = require('fs');

var inputDirectory = 'C:\\Project\\ConvertTangoToGeneric\\Input\\';
var outputDirectory = 'C:\\Project\\ConvertTangoToGeneric\\Output\\';
//var inputFilename = 'C:\\Project\\ConvertTangoToGeneric\\Input\\15811643.xlsx';
//var outputFilename = 'C:\\Project\\ConvertTangoToGeneric\\Output\\res.xlsx';



readFiles(inputDirectory, function(filename, content) {

    //console.log('Read file ' + filename);
}, function(err) {
    throw err;
});

function readFiles(dirname, onFileContent, onError) {
    fs.readdir(dirname, function(err, filenames) {
        if (err) {
            onError(err);
            return;
        }
        filenames.forEach(function(filename) {
            console.log(`TEST: ${filename}`);
            this.processExcel(inputDirectory + filename);

        });
    });
}

processExcel = function(inputFilename, filename) {
    // read from a file
    debugger;
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(inputFilename)
        .then(function() {
            // use workbook
        });

    // write to a file
    var workbook1 = new Excel.Workbook();
    var sheet1 = workbook1.addWorksheet('Sheet1');
    var reColumns=[
        {header:'FirstName',key:'firstname'},
        {header:'LastName',key:'lastname'},
        {header:'Other Name',key:'othername'}
    ];
    sheet1.columns = reColumns;
    workbook1.xlsx.writeFile(outputDirectory + filename)
        .then(function() {
            // done
        });
}




//read all xlls files in directory



