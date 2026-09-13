SELECT
  tOrders.*,
  tCustomers.FirstName,
  tCustomers.LastName
FROM
  tOrders
  INNER JOIN tCustomers ON tOrders.CustomerID = tCustomers.ID;
