-- Create the database
CREATE DATABASE testvbsproject;

USE testvbsproject;

-- Create Customers table with autogenerated CustomerID
CREATE TABLE Customers (
    CustomerID INT IDENTITY(1, 1) PRIMARY KEY,
    CustomerName VARCHAR(100),
    CustomerEmail VARCHAR(100)
);

-- Create Contacts table with foreign key to Customers
CREATE TABLE Contacts (
    ContactID INT IDENTITY(1, 1) PRIMARY KEY,
    CustomerID INT,
    ContactName VARCHAR(100),
    ContactEmail VARCHAR(100),
    CONSTRAINT FK_Contacts_Customers FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID)
);

-- Create Addresses table with foreign key to Customers
CREATE TABLE Addresses (
    AddressID INT IDENTITY(1, 1) PRIMARY KEY,
    CustomerID INT,
    AddressLine1 VARCHAR(100),
    AddressLine2 VARCHAR(100),
    City VARCHAR(100),
    State VARCHAR(100),
    ZipCode VARCHAR(20),
    CONSTRAINT FK_Addresses_Customers FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID)
);

-- Create Orders table with foreign keys to Customers, Contacts, and Addresses
CREATE TABLE Orders (
    OrderID INT IDENTITY(1, 1) PRIMARY KEY,
    CustomerID INT,
    ContactID INT,
    AddressID INT,
    OrderDate DATE,
    TotalAmount DECIMAL(10, 2),
    CONSTRAINT FK_Orders_Customers FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID),
    CONSTRAINT FK_Orders_Contacts FOREIGN KEY (ContactID) REFERENCES Contacts(ContactID),
    CONSTRAINT FK_Orders_Addresses FOREIGN KEY (AddressID) REFERENCES Addresses(AddressID)
);