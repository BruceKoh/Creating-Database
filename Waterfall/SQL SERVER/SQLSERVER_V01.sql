-- POR2019
SELECT * FROM [MJDB2019].dbo.POR2019

-- POR2018
SELECT * FROM [MJDB2018].dbo.POR2018

-- SHIPMENT
SELECT * FROM [MJSHIPMENT].dbo.SHIPMENT

-- GET PLATFORM FROM SHIP
SELECT DISTINCT SHIP.Platform FROM [MJSHIPMENT].dbo.SHIPMENT AS SHIP

-- GET POR DATA (MAKING THE PIVOT)
SELECT POR2019.YYYYWW, POR2019.Planning_Wk, POR2019.Qty 
FROM [MJDB2019].dbo.POR2019 AS POR2019 
WHERE POR2019.Platform = 'RUBY MOBILE '

-- GET SHIP DATA (FOR D)
SELECT SHIP.YYYYWW,SUM(SHIP.Total) AS SUM FROM [MJSHIPMENT].dbo.SHIPMENT AS SHIP
WHERE SHIP.Platform = 'RUBY MOBILE '
GROUP BY SHIP.YYYYWW

-- GET REQUIRED POR TABLE
SELECT POR2019.Planning_Wk, POR2019.YYYYWW, POR2019.Region, POR2019.MPA, SUM(POR2019.Qty) AS Qty, POR2019.QtyType
FROM [MJDB2019].dbo.POR2019 AS POR2019
WHERE POR2019.Program = 'RUBY MOBILE '
GROUP BY POR2019.Planning_Wk,POR2019.YYYYWW, POR2019.Region, POR2019.MPA, POR2019.QtyType
ORDER BY POR2019.Planning_Wk , POR2019.YYYYWW

-- GET REQUIRED SHIP TABLE
SELECT SHIP.YYYYWW, SHIP.Region,SHIP.MPA,SHIP.Total AS Qty, SHIP.QtyType FROM [MJSHIPMENT].dbo.SHIPMENT AS SHIP
WHERE SHIP.Platform = 'RUBY MOBILE ' 

-- GET DISTINCT POR2019.YYYYWW
SELECT DISTINCT POR2019.Planning_Wk FROM MJDB2019.dbo.POR2019
WHERE POR2019.Platform = 'RUBY MOBILE ' 
ORDER BY POR2019.Planning_Wk

--GET DISTINCT POR2019.Planning_wk STRIP
SELECT DISTINCT REPLACE(POR2019.Planning_Wk, 'W', '') AS Plan_Wk
FROM MJDB2019.dbo.POR2019 AS POR2019
WHERE POR2019.Platform = 'RUBY MOBILE ' 
GROUP BY POR2019.Planning_Wk

--GETTING SHIP INFO
SELECT SHIP.YYYYWW,SHIP.Region, SHIP.MPA, SHIP.Total, Ship.QtyType FROM [MJSHIPMENT].dbo.SHIPMENT AS SHIP
WHERE SHIP.Platform = 'RUBY MOBILE ' AND SHIP.YYYYWW <= 201902
ORDER BY SHIP.YYYYWW