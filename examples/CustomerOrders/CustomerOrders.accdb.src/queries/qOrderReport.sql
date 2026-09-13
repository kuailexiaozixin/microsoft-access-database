SELECT
  tOrders.ID AS OrderID,
  tOrders.OrderNumber,
  tOrders.OrderDate,
  tCustomers.FirstName,
  tCustomers.LastName,
  tCustomers.Address,
  tCustomers.City,
  tCustomers.State,
  tCustomers.PostCode,
  tCustomers.Phone,
  tOrderItems.ItemDescription,
  tOrderItems.UnitPriceNet,
  tOrderItems.Quantity,
  tOrderItems.SubTotalNet
FROM
  (
    tOrders
    INNER JOIN tCustomers ON tOrders.CustomerID = tCustomers.ID
  )
  INNER JOIN tOrderItems ON tOrders.ID = tOrderItems.OrderID;
