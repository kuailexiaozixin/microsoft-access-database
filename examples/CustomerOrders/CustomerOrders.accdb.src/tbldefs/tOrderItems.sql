CREATE TABLE [tOrderItems] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [OrderID] LONG,
  [ItemDescription] VARCHAR (255),
  [UnitPriceNet] CURRENCY,
  [Quantity] LONG,
  [SubTotalNet] CURRENCY
)
