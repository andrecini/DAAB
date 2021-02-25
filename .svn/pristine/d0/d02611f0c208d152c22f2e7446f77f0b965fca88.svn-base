use Northwind

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[getProductsByCategory]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[getProductsByCategory]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[getProductDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[getProductDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[getProductName]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[getProductName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[debit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[debit]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[credit]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[credit]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Credits]') and OBJECTPROPERTY(id, N'IsTable') = 1)
drop table [dbo].[Credits]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Debits]') and OBJECTPROPERTY(id, N'IsTable') = 1)
drop table [dbo].[Debits]
GO

-- RetrieveDataReader() and RetrieveDataset() samples
CREATE PROCEDURE getProductsByCategory @CategoryID INTEGER
AS
SELECT ProductID, ProductName, QuantityPerUnit, UnitPrice
FROM Products
WHERE CategoryID = @CategoryID
GO

-- RetrieveSingleRow() sample
CREATE PROCEDURE getProductDetails
	@ProductID int,
	@ProductName nvarchar(40) OUTPUT,
	@UnitPrice money OUTPUT,
	@QtyPerUnit nvarchar(20) OUTPUT
AS
SELECT 	@ProductName = ProductName, 
       	@UnitPrice = UnitPrice,
       	@QtyPerUnit = QuantityPerUnit
FROM Products 
WHERE ProductID = @ProductID
GO

-- LookupSingleItem() sample
CREATE PROCEDURE getProductName @ProductID int
AS
SELECT ProductName
FROM Products
WHERE ProductID = @ProductID
GO


-- PerformTransactionalUpdate() sample
CREATE TABLE Credits
(CreditNo INTEGER IDENTITY,
 AccountNo CHAR(20),
 Amount SMALLMONEY
)
GO
CREATE TABLE Debits
(CreditNo INTEGER IDENTITY,
 AccountNo CHAR(20),
 Amount SMALLMONEY
)
GO
CREATE PROCEDURE credit 
	@AccountNo CHAR(20),
	@Amount SMALLMONEY
AS
INSERT Credits
VALUES
(@AccountNo, @Amount)
GO
CREATE PROCEDURE debit 
	@AccountNo CHAR(20),
	@Amount SMALLMONEY
AS
INSERT Debits
VALUES
(@AccountNo, @Amount)
GO