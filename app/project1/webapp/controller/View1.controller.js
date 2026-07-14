sap.ui.define(
    [
        "sap/ui/core/mvc/Controller",
        "sap/ui/model/json/JSONModel",
        "sap/m/MessageToast",
        "sap/m/Column",
        "sap/m/Text",
        "sap/m/ColumnListItem",
        "sap/m/ObjectStatus",
    ],
    function (Controller, JSONModel, MessageToast, Column, Text, ColumnListItem, ObjectStatus) {
        "use strict";

        return Controller.extend("excelupload.project1.controller.View1", {
            onInit: function () {},

            onFileUpload: function (oEvent) {
                var oFile = oEvent.getParameter("files")[0];

                if (!oFile) {
                    return;
                }

                var reader = new FileReader();

                reader.onload = async function (e) {
                    var data = e.target.result;

                    var workbook = XLSX.read(data, {
                        type: "binary",
                    });

                    var sheetName = workbook.SheetNames[0];

                    var worksheet = workbook.Sheets[sheetName];

                    var aData = XLSX.utils.sheet_to_json(worksheet);

                    if (!aData.length) {
                        MessageToast.show("Excel contains no data");
                        return;
                    }

                    // Enhancement: Validate uploaded Excel data before binding to UI
                    this._validateData(aData);

                    // Enhancement: Check existing records in DB after local validation
                    await this._checkDatabaseDuplicates(aData);

                    var oModel = new JSONModel({
                        excelData: aData,
                        summary: this._summary,
                    });

                    this.getView().setModel(oModel, "excel");

                    this._createDynamicTable(aData);

                    MessageToast.show("Excel loaded successfully");
                }.bind(this);

                reader.readAsBinaryString(oFile);
            },

            _validateData: function (aData) {
                // Enhancement: Mandatory field validation rules
                var aMandatoryFields = ["EMPID", "NAME", "LOCATION"];

                // Enhancement: Track EMPIDs to identify duplicate records
                var oEmpIdTracker = {};

                // Enhancement: Counters used for summary dashboard
                var iValid = 0;
                var iInvalid = 0;
                var iDuplicate = 0;

                aData.forEach(function (oRow) {
                    var aErrors = [];

                    aMandatoryFields.forEach(function (sField) {
                        if (
                            !Object.prototype.hasOwnProperty.call(oRow, sField) ||
                            oRow[sField] === null ||
                            oRow[sField] === undefined ||
                            String(oRow[sField]).trim() === ""
                        ) {
                            aErrors.push("Missing " + sField);
                        }
                    });

                    // Enhancement: Duplicate EMPID validation within uploaded Excel
                    var sEmpId = String(oRow.EMPID || "").trim();

                    if (sEmpId) {
                        if (oEmpIdTracker[sEmpId]) {
                            // Enhancement: Highlight duplicate records within uploaded Excel
                            oRow.STATUS_STATE = "Warning";

                            aErrors.push("DUPLICATE EMPID");

                            iDuplicate++;
                        }

                        oEmpIdTracker[sEmpId] = true;
                    }

                    if (aErrors.length === 0) {
                        // Enhancement: Green indicator for valid records
                        oRow.STATUS = "VALID";
                        oRow.STATUS_STATE = "Success";

                        iValid++;
                    } else {
                        oRow.STATUS = aErrors.join(", ");

                        // Enhancement: Red indicator for validation failures
                        oRow.STATUS_STATE = "Error";

                        iInvalid++;
                    }
                });

                // Enhancement: Generate validation summary statistics
                this._summary = {
                    total: aData.length,
                    valid: iValid,
                    invalid: iInvalid,
                    duplicate: iDuplicate,
                    dbDuplicates: 0,
                };
            },

            // Enhancement: Check uploaded EMPIDs against existing database records
            _checkDatabaseDuplicates: async function (aData) {
                try {
                    var aEmpIds = [];

                    aData.forEach(function (oRow) {
                        if (oRow.EMPID) {
                            aEmpIds.push(String(oRow.EMPID));
                        }
                    });

                    const response = await fetch("/odata/v4/excel/checkDuplicates", {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                        },
                        body: JSON.stringify({
                            empIds: aEmpIds,
                        }),
                    });

                    if (!response.ok) {
                        throw new Error(await response.text());
                    }

                    const result = await response.json();

                    const aExisting = result.value || [];

                    const aExistingIds = aExisting.map(function (item) {
                        return item.EMPID;
                    });

                    var iDbDuplicates = 0;

                    aData.forEach(function (oRow) {
                        if (oRow.STATUS === "VALID" && aExistingIds.includes(String(oRow.EMPID))) {
                            oRow.STATUS = "ALREADY EXISTS IN DB";

                            // Enhancement: Mark records already available in database
                            oRow.STATUS_STATE = "Warning";

                            iDbDuplicates++;
                        }
                    });

                    this._summary.dbDuplicates = iDbDuplicates;
                } catch (error) {
                    console.error("Database duplicate validation error", error);
                }
            },

            _createDynamicTable: function (aData) {
                var oTable = this.byId("idTable");

                oTable.removeAllColumns();

                // Enhancement: Dynamically create table columns from Excel headers
                var aKeys = Object.keys(aData[0]);

                var oTemplate = new ColumnListItem();

                aKeys.forEach(function (sKey) {
                    oTable.addColumn(
                        new Column({
                            header: new Text({
                                text: sKey,
                            }),
                        })
                    );

                    // Enhancement: Render STATUS column with color-coded ObjectStatus control
                    if (sKey === "STATUS") {
                        oTemplate.addCell(
                            new ObjectStatus({
                                text: "{excel>STATUS}",
                                state: "{excel>STATUS_STATE}",
                            })
                        );
                    } else {
                        oTemplate.addCell(
                            new Text({
                                text: "{excel>" + sKey + "}",
                            })
                        );
                    }
                });

                // Enhancement: Dynamic row binding based on uploaded Excel structure
                oTable.bindItems({
                    path: "excel>/excelData",
                    template: oTemplate,
                });
            },
            // Enhancement: Export invalid records into Excel error report
            onDownloadErrorReport: function () {
                var oModel = this.getView().getModel("excel");

                var aData = oModel.getProperty("/excelData");

                var aInvalidRows = aData.filter(function (oRow) {
                    return oRow.STATUS !== "VALID";
                });

                if (!aInvalidRows.length) {
                    MessageToast.show("No invalid records found");

                    return;
                }

                var oWorkbook = XLSX.utils.book_new();

                var oWorksheet = XLSX.utils.json_to_sheet(aInvalidRows);

                XLSX.utils.book_append_sheet(oWorkbook, oWorksheet, "ValidationErrors");

                XLSX.writeFile(oWorkbook, "ValidationErrors.xlsx");

                MessageToast.show("Error report downloaded");
            },
            onSubmit: async function () {
                var oTable = this.byId("idTable");

                var aSelectedItems = oTable.getSelectedItems();

                var aSelectedData = [];

                aSelectedItems.forEach(function (oItem) {
                    var oData = oItem.getBindingContext("excel").getObject();

                    // Enhancement: Only valid records are allowed to be uploaded
                    if (oData.STATUS === "VALID") {
                        aSelectedData.push({
                            EMPID: String(oData.EMPID || ""),
                            NAME: String(oData.NAME || ""),
                            LOCATION: String(oData.LOCATION || ""),
                        });
                    }
                });

                // Validation check before backend submission
                if (!aSelectedData.length) {
                    MessageToast.show("No valid records selected");

                    return;
                }

                console.log("Valid Selected Rows");
                console.log(aSelectedData);

                try {
                    const response = await fetch("/odata/v4/excel/uploadEmployees", {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json",
                        },
                        body: JSON.stringify({
                            employees: aSelectedData,
                        }),
                    });

                    if (!response.ok) {
                        throw new Error(await response.text());
                    }

                    const result = await response.text();

                    console.log(result);

                    MessageToast.show("Data sent successfully");
                } catch (error) {
                    console.error(error);

                    MessageToast.show("Backend call failed");
                }
            },
        });
    }
);
