//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace AccountingPL
//{
//    class RebuildDatabase
//    {





// -- > Done 5/20/20
//        /****** Object:  Table [tb_Address]    Script Date: 3/18/2020 8:56:00 PM ******/
//        SET ANSI_NULLS ON
//        GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_Address]
//        (

//   [AddressID][int] NOT NULL,

//   [AddressLine1] [nvarchar] (80) NULL,
//	[AddressLine2] [nvarchar] (80) NULL,
//	[City] [nvarchar] (40) NULL,
//	[StateProvince] [Name] NULL,  [nvarchar] (80)
//	[Country] [Name] NULL,  [nvarchar] (80)
//	[PostalCode] [nvarchar] (15) NULL,
//	[Phone] [nvarchar] (20) NULL,
//	[Fax] [nvarchar] (20) NULL,
// CONSTRAINT[PK_tb_Locations] PRIMARY KEY CLUSTERED
//(
//   [AddressID] ASC
//)WITH(STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF) ON[PRIMARY]
//) ON[PRIMARY]
//GO







// -- > Done 5/20/20
//            /****** Object:  Table [tb_Category]    Script Date: 3/18/2020 8:57:31 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_Category]
//        (

//   [Header][nvarchar](20) NULL,
//	[Category] [nvarchar] (100) NULL
//) ON[PRIMARY]
//GO







// No need
//            /****** Object:  Table [tb_Config]    Script Date: 3/18/2020 8:57:51 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_Config]
//        (

//   [Year][varchar](4) NULL
//) ON[PRIMARY]
//GO







// --> Done 5/20/20
//            /****** Object:  Table [tb_Dates]    Script Date: 3/18/2020 8:58:45 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_Dates]
//        (

//   [WeekEnd][date] NOT NULL,

//   [Year] [nvarchar] (4) NULL,
//	[WeekNumb] [nvarchar] (2) NULL
//) ON[PRIMARY]
//GO








// -- > Done 5/20/20
//			/****** Object:  Table [tb_ExpenseCost]    Script Date: 3/18/2020 8:59:15 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_ExpenseCost]
//		(

//   [Week][date] NULL,

//   [IDS][int] NULL,

//   [ExpenseCost][money] NULL,

//   [Accounting][money] NULL,

//   [Bank][money] NULL,

//   [CreditCard][money] NULL,

//   [Fuel][money] NULL,

//   [Legal][money] NULL,

//   [License][money] NULL,

//   [PayrollProc][money] NULL,

//   [Insurance][money] NULL,

//   [WorkersComp][money] NULL,

//   [Advertising][money] NULL,

//   [Charitable][money] NULL,

//   [Auto][money] NULL,

//   [CashShortage][money] NULL,

//   [Electrical][money] NULL,

//   [General][money] NULL,

//   [HVAC][money] NULL,

//   [Lawn][money] NULL,

//   [Painting][money] NULL,

//   [Plumbing][money] NULL,

//   [Remodeling][money] NULL,

//   [Structural][money] NULL,

//   [DishMachine][money] NULL,

//   [Janitorial][money] NULL,

//   [Office][money] NULL,

//   [Restaurant][money] NULL,

//   [Uniforms][money] NULL,

//   [Data][money] NULL,

//   [Electricity][money] NULL,

//   [Music][money] NULL,

//   [NaturalGas][money] NULL,

//   [Security][money] NULL,

//   [Trash][money] NULL,

//   [WaterSewer][money] NULL
//) ON[PRIMARY]
//GO







// -- > 5/20/20
///****** Object:  Table [tb_FoodCost]    Script Date: 3/18/2020 9:00:08 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_FoodCost]
//        (

//   [Week][date] NULL,

//   [IDs][int] NULL,

//   [FoodCost][money] NULL,

//   [PrimSupp][money] NULL,

//   [OthSupp][money] NULL,

//   [Bread][money] NULL,

//   [Beverage][money] NULL,

//   [Produce][money] NULL,

//   [CarbonDioxide][money] NULL
//) ON[PRIMARY]
//GO







// -- > 5/20/20
///****** Object:  Table [tb_LaborCost]    Script Date: 3/18/2020 9:00:23 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_LaborCost]
//        (

//   [Week][date] NOT NULL,

//   [IDS] [int] NULL,
//	[LaborCost] [money] NULL,
//	[HostCashier] [money] NULL,
//	[Cooks] [money] NULL,
//	[Servers] [money] NULL,
//	[DMO] [money] NULL,
//	[Supervisor] [money] NULL,
//	[Overtime] [money] NULL,
//	[GeneralManager] [money] NULL,
//	[Manager] [money] NULL,
//	[Bonus] [money] NULL,
//	[PayrollTax] [money] NULL
//) ON[PRIMARY]
//GO







// -- > 5/20/20
//            /****** Object:  Table [tb_NetSales]    Script Date: 3/18/2020 9:00:38 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_NetSales]
//        (

//   [Week][date] NOT NULL,

//   [IDs] [int] NULL,
//	[NetSales] [money] NULL,
//	[Healthcare] [money] NULL,
//	[Retirement] [money] NULL,
//	[TotalCost] [money] NULL,
//	[ReturnonRev] [money] NULL
//) ON[PRIMARY]
//GO







// -- > 5/20/20
//            /****** Object:  Table [tb_OrderDetail]    Script Date: 3/18/2020 9:00:53 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_OrderDetail]
//        (

//   [SalesOrderID][int] NOT NULL,

//   [SalesOrderDetailID] [int] IDENTITY(1,1) NOT NULL,

//   [OrderQty] [int] NOT NULL,

//   [ProductID] [int] NOT NULL,

//   [UnitPrice] [money]
//        NOT NULL,

//   [UnitPriceDiscount] [money]
//        NOT NULL,

//   [LineTotal]  AS(isnull(([UnitPrice]*((1.0)-[UnitPriceDiscount]))*[OrderQty],(0.0)))
//) ON[PRIMARY]
//GO






// --> Done 5/20/20
// -- > Error
//            /****** Object:  Table [tb_OrderHeader]    Script Date: 3/18/2020 9:01:12 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_OrderHeader]
//        (

//   [SalesOrderID][int] NOT NULL,

//   [OrderDate] [datetime]
//        NOT NULL,

//   [DueDate] [datetime]
//        NOT NULL,

//   [ShipDate] [datetime] NULL,
//	[SalesOrderNumber] AS(isnull(N'SO'+CONVERT([nvarchar](23),[SalesOrderID],(0)),N'*** ERROR ***')),
//	[PurchaseOrderNumber] [OrderNumber] NULL,  [nvarchar] (80)
//	[AccountNumber] [AccountNumber] NULL,  [nvarchar] (80)
//	[SubTotal]
//        [money]
//        NOT NULL,

//    [TaxAmt] [money]
//        NOT NULL,

//    [Freight] [money]
//        NOT NULL,

//    [TotalDue]  AS(isnull(([SubTotal]+[TaxAmt])+[Freight],(0))),
//	[Comment]
//        [nvarchar]
//        (max) NULL
//) ON[PRIMARY] TEXTIMAGE_ON[PRIMARY]
//GO







// -- > 5/20/20
///****** Object:  Table [tb_OverheadCost]    Script Date: 3/18/2020 9:01:27 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_OverheadCost]
//        (

//   [Week][date] NULL,

//   [IDS][int] NULL,

//   [OverheadCost][money] NULL,

//   [Mortgage][money] NULL,

//   [LoanPayment][money] NULL,

//   [Association][money] NULL,

//   [PropertyTax][money] NULL,

//   [AdvertisingCoop][money] NULL,

//   [NationalAdvertise][money] NULL,

//   [LicensingFee][money] NULL
//) ON[PRIMARY]
//GO







// -- > 5/20/20
///****** Object:  Table [tb_VendorInv]    Script Date: 3/18/2020 9:01:41 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_VendorInv]
//        (

//   [Week][date] NOT NULL,

//   [IDS] [int] NULL,
//	[InvDate] [date] NULL,
//	[VendorID] [nvarchar] (100) NULL,
//	[InvNumber] [nvarchar] (50) NULL,
//	[Category] [nvarchar] (40) NULL,
//	[Item] [nvarchar] (40) NULL,
//	[Amount] [money] NULL
//) ON[PRIMARY]
//GO







// -- > 5/20/20
//            /****** Object:  Table [tb_Vendors]    Script Date: 3/18/2020 9:01:56 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE TABLE[tb_Vendors]
//        (

//   [ID][int] IDENTITY(1,1) NOT NULL,

//  [VendorID] [nvarchar] (100) NULL,
//	[VendorName] [nvarchar] (100) NULL,
//	[SalesPerson] [nvarchar] (50) NULL,
//	[Phone] [nvarchar] (30) NULL,
//	[AddressLine1] [nvarchar] (100) NULL,
//	[AddressLine2] [nvarchar] (100) NULL,
//	[City] [nvarchar] (50) NULL,
//	[StateProvince] [nvarchar] (50) NULL,
//	[CountryRegion] [nvarchar] (50) NULL,
//	[PostalCode] [nvarchar] (20) NULL,
//	[ModifiedDate] [datetime2] (7) NOT NULL
//) ON[PRIMARY]
//GO

//ALTER TABLE[tb_Vendors] ADD CONSTRAINT[DF_tb_Vendors_ModifiedDate]  DEFAULT(getdate()) FOR[ModifiedDate]
//GO







// -- > 5/20/20
//            /****** Object:  View [vw_OrderLogs]    Script Date: 3/18/2020 9:02:23 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE VIEW[vw_OrderLogs]
//        AS
//SELECT        dbo.tb_NetSales.Week, dbo.tb_Dates.Year, dbo.tb_Address.AddressID, dbo.tb_Dates.WeekNumb, dbo.tb_NetSales.NetSales, dbo.tb_NetSales.Healthcare, dbo.tb_NetSales.Retirement, FDCost.PrimSupp, FDCost.OthSupp, 
//                         FDCost.Bread, FDCost.Beverage, FDCost.Produce, FDCost.CarbonDioxide, FDCost.FoodCost, LBRCost.HostCashier, LBRCost.Cooks, LBRCost.Servers, LBRCost.DMO, LBRCost.Supervisor, LBRCost.Overtime, 
//                         LBRCost.GeneralManager, LBRCost.Manager, LBRCost.Bonus, LBRCost.PayrollTax, LBRCost.LaborCost, ExpCost.Accounting, ExpCost.Bank, ExpCost.CreditCard, ExpCost.Fuel, ExpCost.Legal, ExpCost.License, 
//                         ExpCost.PayrollProc, ExpCost.Insurance, ExpCost.WorkersComp, ExpCost.Advertising, ExpCost.Charitable, ExpCost.Auto, ExpCost.CashShortage, ExpCost.Electrical, ExpCost.General, ExpCost.HVAC, ExpCost.Lawn, 
//                         ExpCost.Painting, ExpCost.Plumbing, ExpCost.Remodeling, ExpCost.Structural, ExpCost.DishMachine, ExpCost.Janitorial, ExpCost.Office, ExpCost.Restaurant, ExpCost.Uniforms, ExpCost.Data, ExpCost.Electricity, 
//                         ExpCost.Music, ExpCost.NaturalGas, ExpCost.Security, ExpCost.Trash, ExpCost.WaterSewer, ExpCost.ExpenseCost, OHCost.Mortgage, OHCost.LoanPayment, OHCost.Association, OHCost.PropertyTax, 
//                         OHCost.AdvertisingCoop, OHCost.NationalAdvertise, OHCost.LicensingFee, OHCost.OverheadCost, dbo.tb_NetSales.TotalCost, dbo.tb_NetSales.ReturnonRev
//FROM            dbo.tb_NetSales INNER JOIN
//                         dbo.tb_Dates ON dbo.tb_NetSales.Week = dbo.tb_Dates.WeekEnd INNER JOIN
//                             (SELECT        'Overhead' AS Header, Week, IDS, OverheadCost, Mortgage, LoanPayment, Association, PropertyTax, AdvertisingCoop, NationalAdvertise, LicensingFee
//                               FROM            dbo.tb_OverheadCost) AS OHCost ON dbo.tb_NetSales.Week = OHCost.Week AND dbo.tb_NetSales.IDs = OHCost.IDS INNER JOIN
//                             (SELECT        'Expense' AS Header, Week, IDS, ExpenseCost, Accounting, Bank, CreditCard, Fuel, Legal, License, PayrollProc, Insurance, WorkersComp, Advertising, Charitable, Auto, CashShortage, Electrical, General, HVAC,
//                                                         Lawn, Painting, Plumbing, Remodeling, Structural, DishMachine, Janitorial, Office, Restaurant, Uniforms, Data, Electricity, Music, NaturalGas, Security, Trash, WaterSewer
//                               FROM            dbo.tb_ExpenseCost AS tb_ExpenseCost_1) AS ExpCost ON dbo.tb_NetSales.Week = ExpCost.Week AND dbo.tb_NetSales.IDs = ExpCost.IDS INNER JOIN
//                             (SELECT        'Food' AS Header, Week, IDs, FoodCost, PrimSupp, OthSupp, Bread, Beverage, Produce, CarbonDioxide
//                               FROM            dbo.tb_FoodCost AS tb_FoodCost_1) AS FDCost ON dbo.tb_NetSales.Week = FDCost.Week AND dbo.tb_NetSales.IDs = FDCost.IDs INNER JOIN
//                             (SELECT        'Labor' AS Header, Week, IDS, LaborCost, HostCashier, Cooks, Servers, DMO, Supervisor, Overtime, GeneralManager, Manager, Bonus, PayrollTax
//                               FROM            dbo.tb_LaborCost AS tb_LaborCost_1) AS LBRCost ON dbo.tb_NetSales.Week = LBRCost.Week AND dbo.tb_NetSales.IDs = LBRCost.IDS INNER JOIN
//                         dbo.tb_Address ON dbo.tb_NetSales.IDs = dbo.tb_Address.AddressID
//GO










// -- > 5/20/20
///****** Object:  View [vw_VendInvCat]    Script Date: 3/18/2020 9:02:43 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//CREATE VIEW[vw_VendInvCat]
//AS
//SELECT        dbo.tb_VendorInv.Week, dbo.tb_VendorInv.IDS, dbo.tb_VendorInv.InvDate, dbo.tb_VendorInv.InvNumber, dbo.tb_VendorInv.Amount, dbo.tb_Category.Header, dbo.tb_VendorInv.Category, dbo.tb_VendorInv.VendorID,
//                         dbo.tb_Vendors.VendorName, dbo.tb_Vendors.SalesPerson, dbo.tb_Vendors.Phone, dbo.tb_Vendors.AddressLine1, dbo.tb_Vendors.AddressLine2, dbo.tb_Vendors.City, dbo.tb_Vendors.StateProvince,
//                         dbo.tb_Vendors.CountryRegion, dbo.tb_Vendors.PostalCode
//FROM            dbo.tb_Category INNER JOIN
//                         dbo.tb_VendorInv ON dbo.tb_Category.Category = dbo.tb_VendorInv.Category INNER JOIN
//                         dbo.tb_Vendors ON dbo.tb_VendorInv.VendorID = dbo.tb_Vendors.VendorID
//GO









// -- > 5/20/20
///****** Object:  StoredProcedure [CheckRecord]    Script Date: 3/18/2020 9:03:07 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//-- =============================================
//-- Author:      <Author, , Name>
//-- Create Date: <Create Date, , >
//-- Description: <Description, , >
//-- =============================================
//CREATE PROCEDURE [CheckRecord] @IDs nvarchar(3)
//-- (  -- Add the parameters for the stored procedure here<@Param1, sysname, @p1> <Datatype_For_Param1, , int> = <Default_Value_For_Param1, , 0>, <@Param2, sysname, @p2> <Datatype_For_Param2, , int> = <Default_Value_For_Param2, , 0> )
//AS
//BEGIN
//    -- SET NOCOUNT ON added to prevent extra result sets from
//    -- interfering with SELECT statements.
//    SET NOCOUNT ON

//    -- Insert statements for procedure here


//SELECT tb_ExpenseCost.Week, tb_Address.AddressID, tb_Dates.WeekEnd, tb_Dates.Year, tb_FoodCost.FoodCost, tb_ExpenseCost.ExpenseCost, tb_LaborCost.LaborCost, tb_NetSales.NetSales, tb_OverheadCost.OverheadCost
//FROM            tb_ExpenseCost INNER JOIN
//                         tb_Dates ON tb_ExpenseCost.Week = tb_Dates.WeekEnd INNER JOIN
//                         tb_Address ON tb_ExpenseCost.IDS = tb_Address.AddressID INNER JOIN
//                         tb_FoodCost ON tb_Dates.WeekEnd = tb_FoodCost.Week AND tb_Address.AddressID = tb_FoodCost.IDs INNER JOIN
//                         tb_LaborCost ON tb_Dates.WeekEnd = tb_LaborCost.Week AND tb_Address.AddressID = tb_LaborCost.IDS INNER JOIN
//                         tb_NetSales ON tb_Dates.WeekEnd = tb_NetSales.Week AND tb_Address.AddressID = tb_NetSales.IDs INNER JOIN
//                         tb_OverheadCost ON tb_Dates.WeekEnd = tb_OverheadCost.Week AND tb_Address.AddressID = tb_OverheadCost.IDS
//WHERE (tb_NetSales.Week = CONVERT(DATE, DATEADD([day], ((DATEDIFF([day], '20000102', getdate()) / 7) * 7) + 7, '20000102'), 102)) and tb_Address.AddressID=158

//if @@ROWCOUNT= 0
//begin
//INSERT INTO tb_FoodCost (Week, IDs, FoodCost, PrimSupp, OthSupp, Bread, Beverage, Produce, CarbonDioxide)
//VALUES
//(FORMAT(DATEADD([day], ((DATEDIFF([day], '20000102', getdate()) / 7) * 7) + 7, '20000102'), 'd', 'en-US' ),158,0.00,0.00,0.00,0.00,0.00,0.00,0.00)

//INSERT INTO tb_ExpenseCost(Week, IDs, ExpenseCost, Accounting, Bank, CreditCard, Fuel, Legal, License, PayrollProc, Insurance, WorkersComp, Advertising, Charitable, Auto, CashShortage, Electrical, General, HVAC, Lawn, Painting, Plumbing, Remodeling, Structural, DishMachine, Janitorial, Office, Restaurant, Uniforms, Data, Electricity, Music, NaturalGas, Security, Trash, WaterSewer)
//VALUES
//(FORMAT(DATEADD([day], ((DATEDIFF([day], '20000102', getdate()) / 7) * 7) + 7, '20000102'), 'd', 'en-US' ),158,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00)

//INSERT INTO tb_LaborCost(Week, IDs, LaborCost, HostCashier, Cooks, Servers, DMO, Supervisor, Overtime, GeneralManager, Manager, Bonus, PayrollTax)
//VALUES
//(FORMAT(DATEADD([day], ((DATEDIFF([day], '20000102', getdate()) / 7) * 7) + 7, '20000102'), 'd', 'en-US' ),158,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00)

//INSERT INTO tb_NetSales(Week, IDs, NetSales, Healthcare, Retirement, TotalCost, ReturnonRev)
//VALUES
//(FORMAT(DATEADD([day], ((DATEDIFF([day], '20000102', getdate()) / 7) * 7) + 7, '20000102'), 'd', 'en-US' ),158,0.00,0.00,0.00,0.00,0.00)

//INSERT INTO tb_OverheadCost(Week, IDs, OverheadCost, Mortgage, LoanPayment, Association, PropertyTax, AdvertisingCoop, NationalAdvertise, LicensingFee)
//VALUES
//(FORMAT(DATEADD([day], ((DATEDIFF([day], '20000102', getdate()) / 7) * 7) + 7, '20000102'), 'd', 'en-US' ),158,0.00,0.00,0.00,0.00,0.00,0.00,0.00,0.00)

//END
//END
//GO










// -- > 5/20/20
///****** Object:  StoredProcedure [MakePNL]    Script Date: 3/18/2020 9:03:19 PM ******/
//SET ANSI_NULLS ON
//GO

//SET QUOTED_IDENTIFIER ON
//GO

//-- =============================================
//-- Author:      <Author, , Name>
//-- Create Date: <Create Date, , >
//-- Description: <Description, , >
//-- =============================================
//CREATE PROCEDURE[MakePNL] @Year nvarchar(4), @AddressID nvarchar(3)
//-- (
//    -- Add the parameters for the stored procedure here
//   -- <@Param1, sysname, @p1> <Datatype_For_Param1, , int> = <Default_Value_For_Param1, , 0>,
//   -- <@Param2, sysname, @p2> <Datatype_For_Param2, , int> = <Default_Value_For_Param2, , 0>
//-- )
//AS
//BEGIN
//    -- SET NOCOUNT ON added to prevent extra result sets from
//    -- interfering with SELECT statements.
//    SET NOCOUNT ON

//    -- Insert statements for procedure here
//	-- Year,

//    SELECT Week, WeekNumb, NetSales, PrimSupp, OthSupp, Bread, Beverage, Produce, CarbonDioxide, FoodCost, HostCashier, Cooks, Servers, DMO, Supervisor, Overtime, GeneralManager, Manager, Bonus, PayrollTax, LaborCost, Accounting, Bank, CreditCard, Fuel, Legal, License, PayrollProc, Insurance, WorkersComp, Advertising, Charitable, Auto, CashShortage, Electrical, General, HVAC, Lawn, Painting, Plumbing, Remodeling, Structural, DishMachine, Janitorial, Office, Restaurant, Uniforms, Data, Electricity, Music, NaturalGas, Security, Trash, WaterSewer, ExpenseCost, Mortgage, LoanPayment, Association, PropertyTax, AdvertisingCoop, NationalAdvertise, LicensingFee, OverheadCost, TotalCost, ReturnonRev
//    from dynamicelements..vw_OrderLogs where year= @Year and AddressID = @AddressID order by week

//END
//GO






//    }
//}
