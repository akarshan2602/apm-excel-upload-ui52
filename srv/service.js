const cds = require("@sap/cds");

module.exports = cds.service.impl(async function () {
    this.on("uploadEmployees", async (req) => {
        const employees = req.data.employees;

        const db = await cds.connect.to("db");

        await db.run(
            INSERT.into("excel.Employees").entries(
                employees.map((emp) => ({
                    EMPID: String(emp.EMPID),
                    NAME: String(emp.NAME),
                    LOCATION: String(emp.LOCATION),
                }))
            )
        );

        return "Success";
    });

    // Enhancement: Update existing employee records
    this.on("updateEmployees", async (req) => {
        const employees = req.data.employees;

        const db = await cds.connect.to("db");

        for (const emp of employees) {
            await db.run(
                UPDATE("excel.Employees")
                    .set({
                        NAME: String(emp.NAME),
                        LOCATION: String(emp.LOCATION),
                    })
                    .where({
                        EMPID: String(emp.EMPID),
                    })
            );
        }

        return "Employees Updated";
    });
    // ---------------------------------------Enhancement: Upload product records
    this.on("uploadProducts", async (req) => {
        const products = req.data.products;

        const db = await cds.connect.to("db");

        await db.run(
            INSERT.into("excel.Products").entries(
                products.map((product) => ({
                    ProductID: String(product.ProductID),
                    ProductName: String(product.ProductName),
                    Price: String(product.Price),
                }))
            )
        );

        return "Products Uploaded";
    });
    // ---------------------------------------Enhancement: Upload store records
    this.on("uploadStores", async (req) => {
        const stores = req.data.stores;

        const db = await cds.connect.to("db");

        await db.run(
            INSERT.into("excel.Stores").entries(
                stores.map((store) => ({
                    StoreID: String(store.StoreID),
                    StoreName: String(store.StoreName),
                    City: String(store.City),
                }))
            )
        );

        return "Stores Uploaded";
    });

    // Enhancement: Check whether uploaded EMPIDs already exist
    this.on("checkDuplicates", async (req) => {
        const empIds = req.data.empIds;

        const db = await cds.connect.to("db");

        const existingEmployees = await db.run(
            SELECT.from("excel.Employees")
                .columns("EMPID")
                .where({
                    EMPID: {
                        in: empIds,
                    },
                })
        );

        return existingEmployees;
    });

    // Enhancement: Fetch complete employee details for update-candidate detection
    this.on("getEmployeesByIds", async (req) => {
        const empIds = req.data.empIds;

        const db = await cds.connect.to("db");

        const employees = await db.run(
            SELECT.from("excel.Employees")
                .columns("EMPID", "NAME", "LOCATION")
                .where({
                    EMPID: {
                        in: empIds,
                    },
                })
        );

        return employees;
    });
});
