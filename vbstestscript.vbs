' Create an ADO connection to the SQL Server
Dim conn
Set conn = CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=SQLOLEDB;Data Source=LAPTOP-NBAKBDVU;Initial Catalog=testvbsproject;User ID=sweet; Password=Laserguy1;"
conn.Open

' Function to display a message box with multi-line input prompt
Function MultiLinePrompt(prompt)
    Dim objShell : Set objShell = CreateObject("WScript.Shell")
    Dim input
    input = InputBox(prompt, "Database Management System", "", -1, -1)
    MultiLinePrompt = input
End Function

' Function to display a list of customers with their IDs and names
Sub ViewCustomers()
    Dim sql, rs, customerList
    sql = "SELECT CustomerID, CustomerName FROM Customers"
    Set rs = conn.Execute(sql)
    
    customerList = "Customer List:" & vbNewLine
    While Not rs.EOF
        customerList = customerList & rs("CustomerID") & ". " & rs("CustomerName") & vbNewLine
        rs.MoveNext
    Wend
    
    MsgBox customerList
    rs.Close
End Sub

' Function to validate customer ID
Function ValidateCustomerID(customerID)
    Dim sql, rs
    sql = "SELECT COUNT(*) AS RecordCount FROM Customers WHERE CustomerID = " & customerID
    Set rs = conn.Execute(sql)
    
    If rs("RecordCount") > 0 Then
        ValidateCustomerID = True
    Else
        ValidateCustomerID = False
    End If
    
    rs.Close
End Function

' Function to delete related records in Contacts table for a given customer ID
Sub DeleteContacts(customerID)
    Dim sql
    sql = "DELETE FROM Contacts WHERE CustomerID = " & customerID
    conn.Execute sql
End Sub

' Function to delete related records in Addresses table for a given customer ID
Sub DeleteAddresses(customerID)
    Dim sql
    sql = "DELETE FROM Addresses WHERE CustomerID = " & customerID
    conn.Execute sql
End Sub

' Function to add a customer
Sub AddCustomer()
    ' Prompt for customer details
    Dim customerName, customerEmail
    customerName = InputBox("Enter the customer name:")
    customerEmail = InputBox("Enter the customer email:")
    
    ' Execute SQL query to insert the customer into the database
    Dim sql
    sql = "INSERT INTO Customers (CustomerName, CustomerEmail) VALUES ('" & customerName & "', '" & customerEmail & "')"
    conn.Execute sql
    
    MsgBox "Customer added successfully."
End Sub

' Function to add a contact
Sub AddContact()
    ' Prompt for contact details
    Dim customerID, contactName, contactEmail
    customerID = InputBox("Enter the customer ID:")
    contactName = InputBox("Enter the contact name:")
    contactEmail = InputBox("Enter the contact email:")
    
    ' Execute SQL query to insert the contact into the database
    Dim sql
    sql = "INSERT INTO Contacts (CustomerID, ContactName, ContactEmail) VALUES (" & customerID & ", '" & contactName & "', '" & contactEmail & "')"
    conn.Execute sql
    
    MsgBox "Contact added successfully."
End Sub

' Function to add an address
Sub AddAddress()
    ' Prompt for address details
    Dim customerID, addressLine1, addressLine2, city, state, zipCode
    customerID = InputBox("Enter the customer ID:")
    addressLine1 = InputBox("Enter the address line 1:")
    addressLine2 = InputBox("Enter the address line 2:")
    city = InputBox("Enter the city:")
    state = InputBox("Enter the state:")
    zipCode = InputBox("Enter the ZIP code:")
    
    ' Execute SQL query to insert the address into the database
    Dim sql
    sql = "INSERT INTO Addresses (CustomerID, AddressLine1, AddressLine2, City, State, ZipCode) VALUES (" & customerID & ", '" & addressLine1 & "', '" & addressLine2 & "', '" & city & "', '" & state & "', '" & zipCode & "')"
    conn.Execute sql
    
    MsgBox "Address added successfully."
End Sub

' Menu loop
Do While True
    ' Display the menu options
    menuOptions = "Menu:" & vbNewLine & _
                  "1. Add a customer" & vbNewLine & _
                  "2. Add a contact" & vbNewLine & _
                  "3. Add an address" & vbNewLine & _
                  "4. View customer list" & vbNewLine & _
                  "5. Delete a customer" & vbNewLine & _
                  "6. Exit"

    ' Get user's choice using a multi-line prompt input box
    userChoice = MultiLinePrompt(menuOptions)

    ' Perform the selected operation based on user input
    Select Case userChoice
        Case "1"
            AddCustomer
        Case "2"
            AddContact
        Case "3"
            AddAddress
        Case "4"
            ViewCustomers
        Case "5"
            ' Delete a customer
            Dim customerID
            customerID = InputBox("Enter the customer ID to delete:")
            If IsNumeric(customerID) Then
                If ValidateCustomerID(customerID) Then
                    ' Delete related contacts and addresses first
                    DeleteContacts customerID
                    DeleteAddresses customerID
                    ' Then delete the customer
                    DeleteCustomer customerID
                Else
                    MsgBox "Invalid customer ID. Please try again."
                End If
            Else
                MsgBox "Invalid input. Please enter a valid customer ID."
            End If
        Case "6"
            ' Exit the program
            conn.Close
            WScript.Quit
        Case Else
            ' Invalid choice
            MsgBox "Invalid choice. Please try again."
    End Select
Loop

' Function to delete a customer
Sub DeleteCustomer(customerID)
    Dim sql
    sql = "DELETE FROM Customers WHERE CustomerID = " & customerID
    conn.Execute sql
    MsgBox "Customer with ID " & customerID & " successfully deleted."
End Sub
