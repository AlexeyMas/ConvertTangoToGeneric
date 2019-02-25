var Excel = require('exceljs');
var fs = require('fs');

var inputDirectory = 'C:\\Project\\ConvertTangoToGeneric\\Input\\';
var outputDirectory = 'C:\\Project\\ConvertTangoToGeneric\\Output\\';
//var inputFilename = 'C:\\Project\\ConvertTangoToGeneric\\Input\\15811643.xlsx';
//var outputFilename = 'C:\\Project\\ConvertTangoToGeneric\\Output\\res.xlsx';

var usageMapping = {
    'Service Charges':	'rucurring_charges',
    'Usage Charges':	'domestic_voice_charges',
    'Itemized Charges':	'other_charges',
    'Equipment Charges':	'Equipment_Charges',
    'PICC Charges':	'other_charges',
    'Access Charges':	'other_charges',
    'Other Charges':	'other_charges',
    'One Time Charges':	'other_charges',
    'Advertising':	'other_charges',
    'Feature Charges':	'other_charges',
    'DA Charges':	'other_charges',
    'Credits':	'adjustments',
    'Late Charges':	'late_paymnet_charges',
    'Taxes/Sur':	'taxes',
    'Advertising Charges':	'other_charges',
    'Mileage Charges':	'other_charges'
};



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
            this.processExcel(inputDirectory + filename, filename);

        });
    });
}

processExcel = function(inputFilename, filename) {
    // read from a file
    //console.log(`TEST: ${filename}`);
    var workbook = new Excel.Workbook();
    workbook.xlsx.readFile(inputFilename)
        .then(function() {
            // read sheet
            /*workbook.eachSheet(function(worksheet, sheetId) {
                console.log(worksheet.name);
            });*/
            
            var invoiceSheet = workbook.getWorksheet('Invoice Info');
            invoiceSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
                //console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
            });

            //read column
            //create new sheet and map columns
            var vendorName = invoiceSheet.getRow(1).values[2];
            var global_account = invoiceSheet.getRow(5).values[2];
            var billing_account = global_account;
            var invoice_number = invoiceSheet.getRow(2).values[2];
            var invoice_date = invoiceSheet.getRow(3).values[2];
            var total = invoiceSheet.getRow(7).values[5];
            var currency = invoiceSheet.getRow(2).values[6];
            var total_amount_due = invoiceSheet.getRow(9).values[5];
            var due_date = invoiceSheet.getRow(4).values[2];
            var date_from = '';
            var date_to = '';



            var workbook1 = new Excel.Workbook();

            //GENERAL
            var sheet1 = workbook1.addWorksheet('General');
            var reColumns=[
                {header:'Contract_Number',key:'contract_number'},
                {header:'Provider',key:'provider'},
                {header:'Global_Account',key:'global_account'},
                {header:'Invoice_Number',key:'invoice_number'},
                {header:'Invoice_Date',key:'invoice_date'},
                {header:'Total',key:'total'},
                {header:'Currency',key:'currency'},
                {header:'Total_Amount_Due',key:'total_amount_due'},
                {header:'Due_Date',key:'due_date'},
                {header:'Date_From',key:'date_from'},
                {header:'Date_To',key:'date_to'}
            ];
            sheet1.columns = reColumns;
            sheet1.addRow({contract_number: '', provider: vendorName, global_account: global_account,
                invoice_number: invoice_number, invoice_date: invoice_date, total: total,
                currency:currency, total_amount_due: total_amount_due, due_date: due_date,
                date_from:date_from, date_to:date_to});
            //PLANS
            var sheet2 = workbook1.addWorksheet('Plans');
            var reColumns=[
                {header:'Name',key:'name'},
                {header:'Code',key:'code'},
                {header:'Type',key:'type'},
                {header:'Category',key:'category'},
                {header:'Price',key:'price'},
                {header:'Currency',key:'currency'},
                {header:'Charge_Period',key:'charge_Period'},
                {header:'Discount',key:'discount'},
                {header:'Column_description',key:'column_description'}
            ];
            sheet2.columns = reColumns;

            //SERVICES
            var sheet3 = workbook1.addWorksheet('Services');
            var reColumns=[
                {header:'Global_Account',key:'global_account'},
                {header:'Billing_Account',key:'billing_account'},
                {header:'Service_ID',key:'serviceID'},
                {header:'Name',key:'name'},
                {header:'Description',key:'description'},
                {header:'Service_Type',key:'serviceType'},
                {header:'Category',key:'category'},
                {header:'Assigned_User',key:'assigned_user'},
                {header:'Cost_Center',key:'cost_center'},
                {header:'Location_Name',key:'location_name'},
                {header:'Location_Country',key:'location_country'},
                {header:'Location_State',key:'location_state'},
                {header:'Location_City',key:'location_city'},
                {header:'Location_Address',key:'location_address'},
                {header:'Location_Zip',key:'location_zip'},
                {header:'Parent_Service_ID',key:'parent_service_ID'},
                {header:'Parent_Service_Type',key:'parent_Service_Type'},
                {header:'Parent_Relation',key:'parent_Relation'},
                {header:'Status',key:'status'},
                {header:'Activation_Date',key:'activation_Date'},
                {header:'Deactivation_Date',key:'deactivation_Date'},
                {header:'Column_region',key:'column_region'},
                {header:'Column_usage_name',key:'column_usage_name'}

            ];
            sheet3.columns = reColumns;

            workbook.eachSheet(function(worksheet, sheetId) {
             //console.log(worksheet.name);
                worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
                    
                    if (worksheet.getRow(rowNumber).values[3] !== 'Total' && rowNumber != 1 && worksheet.name != 'Invoice Info') {
                        var serviceID = worksheet.getRow(rowNumber).values[4];
                        //console.log("ShetName= "+worksheet.name+" Row " + rowNumber + ' serviceID='+serviceID + " = " + JSON.stringify(row.values) );
                        var serviceType = worksheet.name.indexOf('wireless') > -1 ? 'x_mobi_c_telecom_line' : 'x_mobi_wm_wired_service';
                        var assigned_to = worksheet.getRow(rowNumber).values[5];
                        var cost_center = worksheet.getRow(rowNumber).values[3];
                        sheet3.addRow({
                            global_account: global_account, billing_account: billing_account,
                            serviceID: serviceID, serviceType: serviceType, status: '1', assigned_user: assigned_to,
                            cost_center: cost_center
                        });
                    }
                });
            });


            //CHARGES
            var sheet4 = workbook1.addWorksheet('Charges');
            var reColumns=[
                {header:'Global_Account',key:'global_account'},
                {header:'Billing_Account',key:'billing_account'},
                {header:'Invoice_Number',key:'invoice_number'},
                {header:'Service_ID',key:'serviceID'},
                {header:'Service_Type',key:'serviceType'},
                {header:'Summary_Type',key:'summaryType'},
                {header:'Charge_Type',key:'chargeType'},
                {header:'Name',key:'name'},
                {header:'Description',key:'description'},
                {header:'Service_Plan',key:'service_plan'},
                {header:'Amount',key:'amount'},
                {header:'Currency',key:'currency'}
            ];
            sheet4.columns = reColumns;

            workbook.eachSheet(function(worksheet, sheetId) {

                var firstRow = worksheet.getRow(1);
                firstRow.eachCell(function(cell, colNumber) {
                    if (usageMapping[cell.value]) {
                        var name = usageMapping[cell.value];
                        var description = usageMapping[cell.value];
                        var chargeType = usageMapping[cell.value];
                        

                        worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {

                            if (worksheet.getRow(rowNumber).values[3] !== 'Total' && rowNumber != 1 && worksheet.name != 'Invoice Info') {
                                var serviceID = worksheet.getRow(rowNumber).values[4];
                                //console.log("ShetName= "+worksheet.name+" Row " + rowNumber + ' serviceID='+serviceID + " = " + JSON.stringify(row.values) );
                                var serviceType = worksheet.name.indexOf('wireless') > -1 ? 'x_mobi_c_telecom_line' : 'x_mobi_wm_wired_service';
                                var summaryType = worksheet.name.indexOf('wireless') > -1 ? 'x_mobi_c_telecom_line_summary' : 'x_mobi_wm_wired_service_summary';
                                var amount = worksheet.getRow(rowNumber).values[colNumber];
                                
                                if (amount > 0) {
                                    sheet4.addRow({
                                        global_account: global_account,
                                        billing_account: billing_account,
                                        invoice_number: invoice_number,
                                        serviceID: serviceID,
                                        serviceType: serviceType,
                                        summaryType: summaryType,
                                        currency: currency,
                                        name: name,
                                        description: description,
                                        amount: amount,
                                        chargeType: chargeType
                                    });
                                }
                            }
                        });

                    }
                });
                
            });

            //USAGES
            var sheet5 = workbook1.addWorksheet('Usages');
            var reColumns=[
                {header:'Global_Account',key:'global_account'},
                {header:'Billing_Account',key:'billing_account'},
                {header:'Invoice_Number',key:'invoice_number'},
                {header:'Service_ID',key:'serviceID'},
                {header:'Service_Type',key:'serviceType'},
                {header:'Summary_Type',key:'summaryType'},
                {header:'Usage_Type',key:'usageType'},
                {header:'Value',key:'value'},
                {header:'Unit',key:'unit'}
            ];
            sheet5.columns = reColumns;

            workbook.eachSheet(function(worksheet, sheetId) {

                var firstRow = worksheet.getRow(1);
                firstRow.eachCell(function(cell, colNumber) {
                    if (cell.value == 'Num of Minutes') {
                        var usageType = 'domestic_voice_usage';

                        worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {

                            if (worksheet.getRow(rowNumber).values[3] !== 'Total' && rowNumber != 1 && worksheet.name != 'Invoice Info') {
                                var serviceID = worksheet.getRow(rowNumber).values[4];

                                var serviceType = worksheet.name.indexOf('wireless') > -1 ? 'x_mobi_c_telecom_line' : 'x_mobi_wm_wired_service';
                                var summaryType = worksheet.name.indexOf('wireless') > -1 ? 'x_mobi_c_telecom_line_summary' : 'x_mobi_wm_wired_service_summary';
                                var amount = worksheet.getRow(rowNumber).values[colNumber];
                                if (amount > 0) {
                                    sheet5.addRow({
                                        global_account: global_account,
                                        billing_account: billing_account,
                                        invoice_number: invoice_number,
                                        serviceID: serviceID,
                                        serviceType: serviceType,
                                        summaryType: summaryType,
                                        usageType: usageType,
                                        value: amount,
                                        unit: 'Min'
                                    });
                                }
                            }
                        });

                    }
                });

            });


            saveExcelData(filename, workbook1);
        });
}


saveExcelData = function(filename, workbook) {
    // write to a file

    workbook.xlsx.writeFile(outputDirectory + filename)
        .then(function() {
            // done
        });
}



