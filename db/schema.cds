namespace excel;

entity Employees {
    key ID : UUID;

    EMPID : String;
    NAME : String;
    LOCATION : String;
}

entity Products {
    key ID : UUID;

    ProductID : String;
    ProductName : String;
    Price : String;
}

entity Stores {
    key ID : UUID;

    StoreID : String;
    StoreName : String;
    City : String;
}