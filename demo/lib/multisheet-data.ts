import { FormulaEngine } from "../../src/core/engine";

export const createEngineWithMultiSheetData = () => {
  const engine = FormulaEngine.buildEmpty();

  // Create three sheets
  const salesSheet = engine.addSheet("Sales");
  const productsSheet = engine.addSheet("Products");
  const dashboardSheet = engine.addSheet("Dashboard");

  const salesSheetId = engine.getSheetId(salesSheet);
  const productsSheetId = engine.getSheetId(productsSheet);
  const dashboardSheetId = engine.getSheetId(dashboardSheet);

  // Products Sheet - Master product data (input only)
  const productsData = new Map<string, any>([
    // Headers
    ["A1", "Product ID"],
    ["B1", "Product Name"],
    ["C1", "Category"],
    ["D1", "Unit Price"],
    ["E1", "Cost"],
    ["F1", "Margin"],

    // Product data
    ["A2", "P001"],
    ["B2", "Gaming Laptop"],
    ["C2", "Electronics"],
    ["D2", 1200],
    ["E2", 800],
    ["F2", "=(D2-E2)/D2"],

    ["A3", "P002"],
    ["B3", "Wireless Mouse"],
    ["C3", "Accessories"],
    ["D3", 45],
    ["E3", 20],
    ["F3", "=(D3-E3)/D3"],

    ["A4", "P003"],
    ["B4", "Mechanical Keyboard"],
    ["C4", "Accessories"],
    ["D4", 120],
    ["E4", 60],
    ["F4", "=(D4-E4)/D4"],

    ["A5", "P004"],
    ["B5", "4K Monitor"],
    ["C5", "Electronics"],
    ["D5", 350],
    ["E5", 200],
    ["F5", "=(D5-E5)/D5"],

    ["A6", "P005"],
    ["B6", "Tablet"],
    ["C6", "Electronics"],
    ["D6", 600],
    ["E6", 400],
    ["F6", "=(D6-E6)/D6"],

    ["A7", "P006"],
    ["B7", "Headphones"],
    ["C7", "Accessories"],
    ["D7", 80],
    ["E7", 35],
    ["F7", "=(D7-E7)/D7"],

    ["A8", "P007"],
    ["B8", "Smart Watch"],
    ["C8", "Electronics"],
    ["D8", 250],
    ["E8", 150],
    ["F8", "=(D8-E8)/D8"],
  ]);

  // Sales Sheet - Transaction data (input only)
  const salesData = new Map<string, any>([
    // Headers
    ["A1", "Sale ID"],
    ["B1", "Date"],
    ["C1", "Product ID"],
    ["D1", "Product Name"],
    ["E1", "Quantity"],
    ["F1", "Unit Price"],
    ["G1", "Total"],
    ["H1", "Category"],

    // Sales transactions
    ["A2", "S001"],
    ["B2", "2024-01-15"],
    ["C2", "P001"],
    ["D2", "=INDEX(Products!B:B,MATCH(C2,Products!A:A,0))"],
    ["E2", 2],
    ["F2", "=INDEX(Products!D:D,MATCH(C2,Products!A:A,0))"],
    ["G2", "=E2*F2"],
    ["H2", "=INDEX(Products!C:C,MATCH(C2,Products!A:A,0))"],

    ["A3", "S002"],
    ["B3", "2024-01-16"],
    ["C3", "P002"],
    ["D3", "=INDEX(Products!B:B,MATCH(C3,Products!A:A,0))"],
    ["E3", 5],
    ["F3", "=INDEX(Products!D:D,MATCH(C3,Products!A:A,0))"],
    ["G3", "=E3*F3"],
    ["H3", "=INDEX(Products!C:C,MATCH(C3,Products!A:A,0))"],

    ["A4", "S003"],
    ["B4", "2024-01-17"],
    ["C4", "P003"],
    ["D4", "=INDEX(Products!B:B,MATCH(C4,Products!A:A,0))"],
    ["E4", 3],
    ["F4", "=INDEX(Products!D:D,MATCH(C4,Products!A:A,0))"],
    ["G4", "=E4*F4"],
    ["H4", "=INDEX(Products!C:C,MATCH(C4,Products!A:A,0))"],

    ["A5", "S004"],
    ["B5", "2024-01-18"],
    ["C5", "P004"],
    ["D5", "=INDEX(Products!B:B,MATCH(C5,Products!A:A,0))"],
    ["E5", 1],
    ["F5", "=INDEX(Products!D:D,MATCH(C5,Products!A:A,0))"],
    ["G5", "=E5*F5"],
    ["H5", "=INDEX(Products!C:C,MATCH(C5,Products!A:A,0))"],

    ["A6", "S005"],
    ["B6", "2024-01-19"],
    ["C6", "P005"],
    ["D6", "=INDEX(Products!B:B,MATCH(C6,Products!A:A,0))"],
    ["E6", 2],
    ["F6", "=INDEX(Products!D:D,MATCH(C6,Products!A:A,0))"],
    ["G6", "=E6*F6"],
    ["H6", "=INDEX(Products!C:C,MATCH(C6,Products!A:A,0))"],

    ["A7", "S006"],
    ["B7", "2024-01-20"],
    ["C7", "P001"],
    ["D7", "=INDEX(Products!B:B,MATCH(C7,Products!A:A,0))"],
    ["E7", 1],
    ["F7", "=INDEX(Products!D:D,MATCH(C7,Products!A:A,0))"],
    ["G7", "=E7*F7"],
    ["H7", "=INDEX(Products!C:C,MATCH(C7,Products!A:A,0))"],

    ["A8", "S007"],
    ["B8", "2024-01-21"],
    ["C8", "P006"],
    ["D8", "=INDEX(Products!B:B,MATCH(C8,Products!A:A,0))"],
    ["E8", 4],
    ["F8", "=INDEX(Products!D:D,MATCH(C8,Products!A:A,0))"],
    ["G8", "=E8*F8"],
    ["H8", "=INDEX(Products!C:C,MATCH(C8,Products!A:A,0))"],

    ["A9", "S008"],
    ["B9", "2024-01-22"],
    ["C9", "P007"],
    ["D9", "=INDEX(Products!B:B,MATCH(C9,Products!A:A,0))"],
    ["E9", 2],
    ["F9", "=INDEX(Products!D:D,MATCH(C9,Products!A:A,0))"],
    ["G9", "=E9*F9"],
    ["H9", "=INDEX(Products!C:C,MATCH(C9,Products!A:A,0))"],

    ["A10", "S009"],
    ["B10", "2024-01-23"],
    ["C10", "P003"],
    ["D10", "=INDEX(Products!B:B,MATCH(C10,Products!A:A,0))"],
    ["E10", 1],
    ["F10", "=INDEX(Products!D:D,MATCH(C10,Products!A:A,0))"],
    ["G10", "=E10*F10"],
    ["H10", "=INDEX(Products!C:C,MATCH(C10,Products!A:A,0))"],

    ["A11", "S010"],
    ["B11", "2024-01-24"],
    ["C11", "P004"],
    ["D11", "=INDEX(Products!B:B,MATCH(C11,Products!A:A,0))"],
    ["E11", 2],
    ["F11", "=INDEX(Products!D:D,MATCH(C11,Products!A:A,0))"],
    ["G11", "=E11*F11"],
    ["H11", "=INDEX(Products!C:C,MATCH(C11,Products!A:A,0))"],
  ]);

  // Dashboard Sheet - Comprehensive analytics and output
  const dashboardData = new Map<string, any>([
    ["A1", "📊 BUSINESS DASHBOARD & ANALYTICS"],
    ["A2", "Generated from Products and Sales data"],

    // === PRODUCT ANALYSIS ===
    ["A4", "📦 PRODUCT ANALYSIS"],
    ["A5", "Total Products"],
    ["B5", '=COUNTIF(Products!A2:A10000,"<>")'],
    ["A6", "Average Price"],
    ["B6", "=AVERAGE(Products!D2:D10000)"],
    ["A7", "Average Cost"],
    ["B7", "=AVERAGE(Products!E2:E10000)"],
    ["A8", "Average Margin"],
    ["B8", "=AVERAGE(Products!F2:F10000)"],
    ["A9", "Highest Priced Product"],
    ["B9", "=INDEX(Products!B2:B10000,MATCH(MAX(Products!D2:D10000),Products!D2:D10000,0))"],
    ["A10", "Price Range"],
    ["B10", '=CONCATENATE("$",MIN(Products!D2:D10000)," - $",MAX(Products!D2:D10000))'],

    // Product category breakdown
    ["D4", "📊 PRODUCT CATEGORIES"],
    ["D5", "Electronics Count"],
    ["E5", '=COUNTIF(Products!C2:C10000,"Electronics")'],
    ["D6", "Accessories Count"],
    ["E6", '=COUNTIF(Products!C2:C10000,"Accessories")'],
    ["D7", "Electronics Avg Price"],
    ["E7", '=AVERAGEIF(Products!C2:C10000,"Electronics",Products!D2:D10000)'],
    ["D8", "Accessories Avg Price"],
    ["E8", '=AVERAGEIF(Products!C2:C10000,"Accessories",Products!D2:D10000)'],
    ["D9", "Electronics Inventory Value"],
    ["E9", '=SUMIF(Products!C2:C10000,"Electronics",Products!D2:D10000)'],
    ["D10", "Accessories Inventory Value"],
    ["E10", '=SUMIF(Products!C2:C10000,"Accessories",Products!D2:D10000)'],

    // === SALES ANALYSIS ===
    ["A12", "💰 SALES ANALYSIS"],
    ["A13", "Total Sales Revenue"],
    ["B13", "=SUM(Sales!G2:G10000)"],
    ["A14", "Total Units Sold"],
    ["B14", "=SUM(Sales!E2:E10000)"],
    ["A15", "Total Transactions"],
    ["B15", '=COUNTIF(Sales!A2:A10000,"<>")'],
    ["A16", "Average Sale Value"],
    ["B16", "=AVERAGE(Sales!G2:G10000)"],
    ["A17", "Average Units per Sale"],
    ["B17", "=AVERAGE(Sales!E2:E10000)"],
    ["A18", "Largest Single Sale"],
    ["B18", "=MAX(Sales!G2:G10000)"],
    ["A19", "Sales Date Range"],
    ["B19", '=CONCATENATE("From ",INDEX(Sales!B2:B10000,1)," to latest")'],

    // Sales by category
    ["D12", "💳 SALES BY CATEGORY"],
    ["D13", "Electronics Revenue"],
    ["E13", '=SUMIF(Sales!H2:H10000,"Electronics",Sales!G2:G10000)'],
    ["D14", "Accessories Revenue"],
    ["E14", '=SUMIF(Sales!H2:H10000,"Accessories",Sales!G2:G10000)'],
    ["D15", "Electronics Units"],
    ["E15", '=SUMIF(Sales!H2:H10000,"Electronics",Sales!E2:E10000)'],
    ["D16", "Accessories Units"],
    ["E16", '=SUMIF(Sales!H2:H10000,"Accessories",Sales!E2:E10000)'],
    ["D17", "Electronics Avg Sale"],
    ["E17", '=AVERAGEIF(Sales!H2:H10000,"Electronics",Sales!G2:G10000)'],
    ["D18", "Accessories Avg Sale"],
    ["E18", '=AVERAGEIF(Sales!H2:H10000,"Accessories",Sales!G2:G10000)'],

    // === PRODUCT PERFORMANCE ===
    ["A21", "🏆 TOP PERFORMING PRODUCTS"],
    ["A22", "Gaming Laptop (P001)"],
    ["B22", '=SUMIF(Sales!C2:C10000,"P001",Sales!G2:G10000)'],
    ["A23", "Wireless Mouse (P002)"],
    ["B23", '=SUMIF(Sales!C2:C10000,"P002",Sales!G2:G10000)'],
    ["A24", "Keyboard (P003)"],
    ["B24", '=SUMIF(Sales!C2:C10000,"P003",Sales!G2:G10000)'],
    ["A25", "Monitor (P004)"],
    ["B25", '=SUMIF(Sales!C2:C10000,"P004",Sales!G2:G10000)'],
    ["A26", "Tablet (P005)"],
    ["B26", '=SUMIF(Sales!C2:C10000,"P005",Sales!G2:G10000)'],
    ["A27", "Headphones (P006)"],
    ["B27", '=SUMIF(Sales!C2:C10000,"P006",Sales!G2:G10000)'],
    ["A28", "Smart Watch (P007)"],
    ["B28", '=SUMIF(Sales!C2:C10000,"P007",Sales!G2:G10000)'],

    // Product units sold
    ["D21", "📈 UNITS SOLD BY PRODUCT"],
    ["D22", "Gaming Laptop Units"],
    ["E22", '=SUMIF(Sales!C2:C10000,"P001",Sales!E2:E10000)'],
    ["D23", "Wireless Mouse Units"],
    ["E23", '=SUMIF(Sales!C2:C10000,"P002",Sales!E2:E10000)'],
    ["D24", "Keyboard Units"],
    ["E24", '=SUMIF(Sales!C2:C10000,"P003",Sales!E2:E10000)'],
    ["D25", "Monitor Units"],
    ["E25", '=SUMIF(Sales!C2:C10000,"P004",Sales!E2:E10000)'],
    ["D26", "Tablet Units"],
    ["E26", '=SUMIF(Sales!C2:C10000,"P005",Sales!E2:E10000)'],
    ["D27", "Headphones Units"],
    ["E27", '=SUMIF(Sales!C2:C10000,"P006",Sales!E2:E10000)'],
    ["D28", "Smart Watch Units"],
    ["E28", '=SUMIF(Sales!C2:C10000,"P007",Sales!E2:E10000)'],

    // === BUSINESS METRICS ===
    ["A30", "📊 KEY BUSINESS METRICS"],
    ["A31", "Revenue per Product Type"],
    ["B31", "=B13/B5"],
    ["A32", "Market Share - Electronics"],
    ["B32", '=CONCATENATE(ROUND(E13/B13*100,1),"%")'],
    ["A33", "Market Share - Accessories"],
    ["B33", '=CONCATENATE(ROUND(E14/B13*100,1),"%")'],
    ["A34", "Avg Margin %"],
    ["B34", '=CONCATENATE(ROUND(B8*100,1),"%")'],
    ["A35", "Inventory Efficiency"],
    ["B35", "=B14/B5"], // Units sold per product type
    ["A36", "Sales Velocity"],
    ["B36", "=B13/B15"], // Revenue per transaction

    // === INTERACTIVE LOOKUP ===
    ["A38", "🔍 INTERACTIVE PRODUCT LOOKUP"],
    ["A39", "Enter Product ID:"],
    ["B39", "P001"],
    ["A40", "Product Name:"],
    ["B40", "=INDEX(Products!B2:B10000,MATCH(B39,Products!A2:A10000,0))"],
    ["A41", "Category:"],
    ["B41", "=INDEX(Products!C2:C10000,MATCH(B39,Products!A2:A10000,0))"],
    ["A42", "Unit Price:"],
    ["B42", "=INDEX(Products!D2:D10000,MATCH(B39,Products!A2:A10000,0))"],
    ["A43", "Cost:"],
    ["B43", "=INDEX(Products!E2:E10000,MATCH(B39,Products!A2:A10000,0))"],
    ["A44", "Margin:"],
    ["B44", '=CONCATENATE(ROUND(INDEX(Products!F2:F10000,MATCH(B39,Products!A2:A10000,0))*100,1),"%")'],
    ["A45", "Total Revenue from Sales:"],
    ["B45", "=SUMIF(Sales!C2:C10000,B39,Sales!G2:G10000)"],
    ["A46", "Total Units Sold:"],
    ["B46", "=SUMIF(Sales!C2:C10000,B39,Sales!E2:E10000)"],

    // === TEXT ANALYSIS & INSIGHTS ===
    ["D30", "📝 BUSINESS INSIGHTS"],
    ["D31", "Best Performing Category:"],
    ["E31", '=IF(E13>E14,"Electronics","Accessories")'],
    ["D32", "Category Performance Gap:"],
    ["E32", '=CONCATENATE("$",ABS(E13-E14))'],
    ["D33", "Most Popular Product:"],
    ["E33", "=INDEX(Products!B:B,MATCH(MAX(E22:E28),E22:E28,0))"],
    ["D34", "Highest Revenue Product:"],
    ["E34", "=INDEX(Products!B:B,MATCH(MAX(B22:B28),B22:B28,0))"],
    ["D35", "Report Summary:"],
    ["E35", '=CONCATENATE("$",ROUND(B13,0)," revenue from ",B14," units across ",B5," products")'],
    ["D36", "Category Leader:"],
    ["E36", '=CONCATENATE(UPPER(E31)," leads with $",ROUND(MAX(E13,E14),0))'],
  ]);

  // Populate all sheets
  engine.setSheetContents(productsSheetId, productsData);
  engine.setSheetContents(salesSheetId, salesData);
  engine.setSheetContents(dashboardSheetId, dashboardData);

  return {
    engine,
    sheets: {
      sales: { name: salesSheet, id: salesSheetId },
      products: { name: productsSheet, id: productsSheetId },
      dashboard: { name: dashboardSheet, id: dashboardSheetId },
    },
  };
};
