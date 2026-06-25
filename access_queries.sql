-- MS Access SQL. Вариант 13.
-- Производство электроэнергии в развитых странах мира (млрд. кВт/час).
-- Нормализованная схема: Regions 1 -> M EnergyProduction.

CREATE TABLE Regions (
    RegionId AUTOINCREMENT CONSTRAINT PK_Regions PRIMARY KEY,
    RegionName TEXT(100) NOT NULL
);

CREATE UNIQUE INDEX UX_Regions_RegionName ON Regions (RegionName);

CREATE TABLE EnergyProduction (
    ProductionId AUTOINCREMENT CONSTRAINT PK_EnergyProduction PRIMARY KEY,
    RegionId LONG NOT NULL,
    CountryName TEXT(100) NOT NULL,
    Production2010 DOUBLE NOT NULL,
    Production2015 DOUBLE NOT NULL
);

ALTER TABLE EnergyProduction
    ADD CONSTRAINT FK_EnergyProduction_Regions
    FOREIGN KEY (RegionId) REFERENCES Regions(RegionId);

CREATE UNIQUE INDEX UX_EnergyProduction_CountryName ON EnergyProduction (CountryName);
CREATE INDEX IX_EnergyProduction_RegionId ON EnergyProduction (RegionId);

INSERT INTO Regions (RegionName) VALUES ('Северная Америка');
INSERT INTO Regions (RegionName) VALUES ('Европа');
INSERT INTO Regions (RegionName) VALUES ('Азия');

INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'США', 629.0, 724.0 FROM Regions WHERE RegionName='Северная Америка';
INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'Англия', 93.9, 112.9 FROM Regions WHERE RegionName='Европа';
INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'Канада', 82.8, 96.7 FROM Regions WHERE RegionName='Северная Америка';
INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'Германия', 76.5, 95.1 FROM Regions WHERE RegionName='Европа';
INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'Япония', 65.2, 81.2 FROM Regions WHERE RegionName='Азия';
INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'Франция', 49.6, 61.8 FROM Regions WHERE RegionName='Европа';
INSERT INTO EnergyProduction (RegionId, CountryName, Production2010, Production2015)
SELECT RegionId, 'Швеция', 24.7, 32.5 FROM Regions WHERE RegionName='Европа';

-- Запрос 1. Найти страну, которая в 2015 г. произвела больше всех электроэнергии.
SELECT TOP 1 EP.CountryName, EP.Production2015
FROM EnergyProduction AS EP
ORDER BY EP.Production2015 DESC;

-- Запрос 2. Найти страны, в которых в 2015 г. производство превысило 70 млрд. кВт/час.
SELECT R.RegionName, EP.CountryName, EP.Production2015
FROM Regions AS R
INNER JOIN EnergyProduction AS EP ON R.RegionId = EP.RegionId
WHERE EP.Production2015 > 70
ORDER BY EP.CountryName;

-- Запрос 3. Найти страны, в которых в 2010 г. производство не превышало 100 млрд. кВт/час.
SELECT R.RegionName, EP.CountryName, EP.Production2010
FROM Regions AS R
INNER JOIN EnergyProduction AS EP ON R.RegionId = EP.RegionId
WHERE EP.Production2010 <= 100
ORDER BY EP.CountryName;
