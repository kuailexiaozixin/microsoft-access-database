CREATE TABLE [tCustomers] (
  [ID] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [FirstName] VARCHAR (255),
  [LastName] VARCHAR (255),
  [Address] VARCHAR (255),
  [City] VARCHAR (255),
  [State] VARCHAR (255),
  [PostCode] VARCHAR (255),
  [Phone] VARCHAR (255),
  [Email] VARCHAR (255)
)
