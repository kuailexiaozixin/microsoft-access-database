Operation =1
Option =0
Begin InputTables
    Name ="tOrders"
    Name ="tCustomers"
End
Begin OutputColumns
    Expression ="tOrders.*"
    Expression ="tCustomers.FirstName"
    Expression ="tCustomers.LastName"
End
Begin Joins
    LeftTable ="tOrders"
    RightTable ="tCustomers"
    Expression ="tOrders.CustomerID = tCustomers.ID"
    Flag =1
End
