-- CREATE 2019 SHIPVPOR

CREATE TABLE [dbo].[2019ShipvPOR](
	[Planning_Wk] [varchar](20) NOT NULL,
	[Product_Line] [varchar](2) NOT NULL,
	[Platform] [varchar](50) NOT NULL,
	[Program] [varchar](50) NOT NULL,
	[MPA] [varchar](50) NOT NULL,
	[Target_Location] [varchar](10) NOT NULL,
	[SKU] [varchar](20) NOT NULL,
	[WkDate] [datetime] NOT NULL,
	[PORQty] [int] NOT NULL,
	[Region] [varchar](20) NOT NULL,
	[ShipWkNo] [int] NOT NULL,
	[ShipTotal] [int] NOT NULL,
	[FK] [varchar](100) NOT NULL
)

INSERT INTO [MJ_SHIPVPOR].dbo.[2019ShipvPOR] (Planning_Wk, Product_Line, Platform, Program, MPA, Target_Location, SKU, WkDate, PORQty, Region, ShipWkNo, ShipTotal, FK)
SELECT  POR.Planning_Wk, POR.Product_Line, POR.Platform, POR.Program, POR.MPA, POR.Target_Location, POR.SKU, POR.WkDate, POR.Qty, Ship.Region, Ship.ISO_WK_NR, Ship.Total, POR.FK
FROM [MJ_DB2019].dbo.POR POR Inner join [MJ_SHIPMENT].dbo.SHIPMENT Ship on POR.FK = Ship.PK 

SELECT * FROM [MJ_SHIPVPOR].dbo.[2019ShipvPOR]

-- DELETE THOSE SHIPTOTAL 0
DELETE 
From [MJ_SHIPVPOR].dbo.[2019ShipvPOR]
WHERE [MJ_SHIPVPOR].dbo.[2019ShipvPOR].ShipTotal = '0'

-- CREATE 2018 SHIPVPOR

CREATE TABLE [dbo].[2018ShipvPOR](
	[Planning_Wk] [varchar](20) NOT NULL,
	[Product_Line] [varchar](2) NOT NULL,
	[Platform] [varchar](50) NOT NULL,
	[Program] [varchar](50) NOT NULL,
	[MPA] [varchar](50) NOT NULL,
	[Target_Location] [varchar](10) NOT NULL,
	[SKU] [varchar](20) NOT NULL,
	[WkDate] [datetime] NOT NULL,
	[PORQty] [int] NOT NULL,
	[Region] [varchar](20) NOT NULL,
	[ShipWkNo] [int] NOT NULL,
	[ShipTotal] [int] NOT NULL,
	[FK] [varchar](100) NOT NULL
)

INSERT INTO [MJ_SHIPVPOR].dbo.[2018ShipvPOR] (Planning_Wk, Product_Line, Platform, Program, MPA, Target_Location, SKU, WkDate, PORQty, Region, ShipWkNo, ShipTotal, FK)
SELECT  POR.Planning_Wk, POR.Product_Line, POR.Platform, POR.Program, POR.MPA, POR.Target_Location, POR.SKU, POR.WkDate, POR.Qty, Ship.Region, Ship.ISO_WK_NR, Ship.Total, POR.FK
FROM [MJ_DB2018].dbo.POR POR Inner join [MJ_SHIPMENT].dbo.SHIPMENT Ship on POR.FK = Ship.PK 

SELECT * FROM [MJ_SHIPVPOR].dbo.[2018ShipvPOR]

DELETE 
From [MJ_SHIPVPOR].dbo.[2018ShipvPOR]
WHERE [MJ_SHIPVPOR].dbo.[2018ShipvPOR].ShipTotal = '0'

--SHIPMENT ADD PROGRAM
INSERT INTO [MJ_SHIPMENT].dbo.SHIPMENT(Program)
SELECT POR.Program
FROM [MJ_DB2019].dbo.POR POR Inner join [MJ_SHIPMENT].dbo.SHIPMENT Ship on POR.FK = Ship.PK 