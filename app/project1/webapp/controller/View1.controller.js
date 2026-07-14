// =====================================================================================================================

sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "sap/m/MessageToast",
    "sap/m/Column",
    "sap/m/Text",
    "sap/m/ColumnListItem"
], function (
    Controller,
    JSONModel,
    MessageToast,
    Column,
    Text,
    ColumnListItem
) {

    "use strict";

    return Controller.extend("excelupload.project1.controller.View1", {

        onInit: function () {

        },

        onFileUpload: function (oEvent) {

            var oFile = oEvent.getParameter("files")[0];

            if (!oFile) {
                return;
            }

            var reader = new FileReader();

            reader.onload = function (e) {

                var data = e.target.result;

                var workbook = XLSX.read(data, {
                    type: "binary"
                });

                var sheetName = workbook.SheetNames[0];

                var worksheet = workbook.Sheets[sheetName];

                var aData = XLSX.utils.sheet_to_json(worksheet);

                if (!aData.length) {
                    MessageToast.show("Excel contains no data");
                    return;
                }

                var oModel = new JSONModel({
                    excelData: aData
                });

                this.getView().setModel(oModel, "excel");

                this._createDynamicTable(aData);

                MessageToast.show("Excel loaded successfully");

            }.bind(this);

            reader.readAsBinaryString(oFile);
        },

        _createDynamicTable: function (aData) {

            var oTable = this.byId("idTable");

            oTable.removeAllColumns();

            var aKeys = Object.keys(aData[0]);

            var oTemplate = new ColumnListItem();

            aKeys.forEach(function (sKey) {

                oTable.addColumn(
                    new Column({
                        header: new Text({
                            text: sKey
                        })
                    })
                );

                oTemplate.addCell(
                    new Text({
                        text: "{excel>" + sKey + "}"
                    })
                );

            });

            oTable.bindItems({
                path: "excel>/excelData",
                template: oTemplate
            });

        },

        onSubmit: async function () {

            var oTable = this.byId("idTable");

            var aSelectedItems = oTable.getSelectedItems();

            var aSelectedData = [];

            aSelectedItems.forEach(function (oItem) {

                aSelectedData.push(
                    oItem.getBindingContext("excel").getObject()
                );

            });

            console.log("Selected Rows");
            console.log(aSelectedData);

            try {

                const response = await fetch(
                    "/odata/v4/excel/uploadEmployees",
                    {
                        method: "POST",
                        headers: {
                            "Content-Type": "application/json"
                        },
                        body: JSON.stringify({
                            employees: aSelectedData
                        })
                    }
                );

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

        }

    });

});