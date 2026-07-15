using excel from '../db/schema';

service ExcelService {

    entity Products as projection on excel.Products;

    entity Stores as projection on excel.Stores;
    
    entity Employees as projection on excel.Employees;

    action uploadEmployees(
        employees : many EmployeeInput
    ) returns String;

    action checkDuplicates(
        empIds : many String
    ) returns many ExistingEmployee;

    // Enhancement: Fetch complete employee details for update-candidate analysis
    action getEmployeesByIds(
        empIds : many String
    ) returns many EmployeeDetails;
        // Enhancement: Update existing employee records
    action updateEmployees(
        employees : many EmployeeInput
    ) returns String;
}

type EmployeeInput {
    EMPID    : String;
    NAME     : String;
    LOCATION : String;
}

type ExistingEmployee {
    EMPID : String;
}

type EmployeeDetails {
    EMPID    : String;
    NAME     : String;
    LOCATION : String;
}