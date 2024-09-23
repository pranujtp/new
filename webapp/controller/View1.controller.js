sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageBox",
    "sap/ui/core/format/DateFormat",
    "sap/ui/core/format/NumberFormat"
], function (Controller, JSONModel, MessageBox, DateFormat, NumberFormat) {
    "use strict";

    return Controller.extend("demoapp.controller.View1", {
        onInit: function () {
            var oModel = new JSONModel("model/data.json");
            this.getView().setModel(oModel, "datamodel");
        },

        ondownload: function () {
            var oModel = this.getView().getModel("datamodel");
            var aData = oModel.getProperty("/employees");

            // Prompt user to confirm download
            MessageBox.show(
                "Do you want to download the data?",
                MessageBox.Icon.QUESTION,
                "Confirmation",
                [MessageBox.Action.YES, MessageBox.Action.NO],
                function (oAction) {
                    if (oAction === MessageBox.Action.YES) {
                        var workBook = XLSX.utils.book_new();

                        // Process data to ensure correct formatting
                        var processedData = aData.map(function (employee) {
                            return {
                                ID: employee.id,
                                Firstname: employee.firstname,
                                Lastname: employee.Lastname,
                                Phonenumber: employee.phonenumber.toString(), // Ensure phone number is string
                                Role: employee.role,
                                Joindate: employee.Joindate ? new Date(employee.Joindate).toLocaleDateString('en-GB').replace(/\//g, '/') : "" // Convert date to "dd/mm/yyyy" format
                            };
                        });

                        // Define headers for all columns
                        var headers = ["ID", "Firstname", "Lastname", "Phonenumber", "Role", "Joindate"];

                        // Create worksheet with headers
                        var workSheet = XLSX.utils.aoa_to_sheet([headers]);

                        // Append processed data to the worksheet
                        XLSX.utils.sheet_add_json(workSheet, processedData, { skipHeader: true, origin: 'A2' });

                        // Adjust column width for all columns
                        workSheet['!cols'] = headers.map(function () {
                            return { width: 15 };
                        });

                        // Apply formatting if necessary (optional)
                        var range = XLSX.utils.decode_range(workSheet["!ref"]);
                        for (var rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
                            for (var colNum = range.s.c; colNum <= range.e.c; colNum++) {
                                var cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
                                var cell = workSheet[cellAddress];
                                if (cell && cell.v) {
                                    var colLetter = cellAddress.replace(/[0-9]/g, ''); // Extract column letter

                                    if (colLetter === 'D') { // Phone number column
                                        cell.t = 's'; // Keep phone number as string
                                        cell.z = '0';
                                    } else if (colLetter === 'F') { // Join date column
                                        if (cell.v) {
                                            var dateParts = cell.v.split('/');
                                            if (dateParts.length === 3) { // Check if the date is in dd/mm/yyyy format
                                                var day = parseInt(dateParts[0], 10);
                                                var month = parseInt(dateParts[1], 10) - 1; // Month is 0-based
                                                var year = parseInt(dateParts[2], 10);

                                                var date = new Date(Date.UTC(year, month, day));
                                                if (!isNaN(date.getTime())) { // Check if date is valid
                                                    cell.t = 'n'; // Set cell type to number
                                                    cell.z = 'dd/mm/yyyy'; // Set the desired date format
                                                    cell.v = (date.getTime() / 86400000) + 25569; // Convert to Excel date
                                                } else {
                                                    cell.v = ""; // Clear invalid date
                                                }
                                            } else {
                                                cell.v = ""; // Clear invalid date if not in expected format
                                            }
                                        }
                                    } else { // Other columns
                                        cell.t = 's';
                                        cell.z = '@';
                                    }
                                }
                            }
                        }

                        XLSX.utils.book_append_sheet(workBook, workSheet, "Employees");
                        XLSX.writeFile(workBook, "Employees.xlsx");
                    }
                }
            );
        }
    });
});
