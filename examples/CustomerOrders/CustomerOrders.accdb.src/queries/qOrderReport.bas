Operation =1
Option =0
Begin InputTables
    Name ="tOrders"
    Name ="tCustomers"
    Name ="tOrderItems"
End
Begin OutputColumns
    Alias ="OrderID"
    Expression ="tOrders.ID"
    Expression ="tOrders.OrderNumber"
    Expression ="tOrders.OrderDate"
    Expression ="tCustomers.FirstName"
    Expression ="tCustomers.LastName"
    Expression ="tCustomers.Address"
    Expression ="tCustomers.City"
    Expression ="tCustomers.State"
    Expression ="tCustomers.PostCode"
    Expression ="tCustomers.Phone"
    Expression ="tOrderItems.ItemDescription"
    Expression ="tOrderItems.UnitPriceNet"
    Expression ="tOrderItems.Quantity"
    Expression ="tOrderItems.SubTotalNet"
End
Begin Joins
    LeftTable ="tOrders"
    RightTable ="tCustomers"
    Expression ="tOrders.CustomerID = tCustomers.ID"
    Flag =1
    LeftTable ="tOrders"
    RightTable ="tOrderItems"
    Expression ="tOrders.ID = tOrderItems.OrderID"
    Flag =1
End
