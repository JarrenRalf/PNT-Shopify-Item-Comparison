function onChange(e)
{  
  if (e.changeType === 'INSERT_GRID') // A new sheet has been created
  {
    try
    {
      var spreadsheet = e.source;
      var sheets = spreadsheet.getSheets();
      var info, numRows = 0, numCols = 1, maxRow = 2, maxCol = 3;

      for (var sheet = 0; sheet < sheets.length; sheet++) // Loop through all of the sheets in this spreadsheet and find the new one
      {
        info = [
          sheets[sheet].getLastRow(),
          sheets[sheet].getLastColumn(),
          sheets[sheet].getMaxRows(),
          sheets[sheet].getMaxColumns()
        ]

        // A new sheet is imported by File -> Import -> Insert new sheet(s) - The left disjunct is for a csv and the right disjunct is for an excel file
        if ((info[maxRow] - info[numRows] === 2 && info[maxCol] - info[numCols] === 2) || (info[maxRow] === 1000 && info[maxCol] >= 26 && info[numRows] !== 0 && info[numCols] !== 0)) 
        {
          const values = sheets[sheet].getRange(1, 1, info[numRows], info[numCols]).getValues();
          var dataSheet = (values[0][2] === 'Body (HTML)') ? sheets[sheets.map(sh => sh.getSheetName()).indexOf('FromShopify')] : (values[0][2] === 'disabled Body') ? 
                                                             sheets[sheets.map(sh => sh.getSheetName()).indexOf('FromAdagio' )] : (values[0][2] === 'Price Type'   ) ? 
                                                             sheets[sheets.map(sh => sh.getSheetName()).indexOf('FromWebsiteDiscount')] : null;

          if (dataSheet !== null)
          {
            const dataSheetName = dataSheet.getSheetName();
            
            if (dataSheetName !== 'FromWebsiteDiscount')
            {
              dataSheet.clearContents().getRange(1, 1, info[numRows], info[numCols]).setValues(values)

              if (dataSheetName == 'FromShopify')
                replaceLeadingApostrophesOnVariantSKUs(dataSheet, info[numCols], info[numRows], spreadsheet);
              else 
                ss.getSheetByName("Dashboard").getRange(9, 6).setValue(timeStamp()).activate(); // Timestamp on dashboard
            }
            else // The discounts are being uploaded
            {
              spreadsheet.toast('Accessing the Discount Percentages...', '', -1);
              const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages');
              const discounts = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 5);
              var discountedItem;

              spreadsheet.toast('Discount Percentages acquired. Updating the Shopify discounts (Approx 80 seconds)...', '', -1);

              dataSheet.clearContents().getRange(1, 1, info[numRows], info[numCols]).setValues(values.map(webItem => {
                discountedItem = discounts.find(item => item[0].split(' - ').pop().toString().toUpperCase() === webItem[1].toString().toUpperCase());
                return (discountedItem != null) ? [webItem[0], webItem[1], webItem[2], discountedItem[2], discountedItem[3], discountedItem[4]] : webItem;
              })).activate();

              dataSheet.setFrozenRows(1)

              spreadsheet.toast('Shopify discounts successfully updated.', 'Complete', 20)
            }
          }      
          
          const recentlyImportedSheetName = sheets[sheet].getSheetName();
          
           // Don't delete the sheets that are duplicates
          if (recentlyImportedSheetName.substring(0, 7) !== "Copy Of" && recentlyImportedSheetName !== 'FromAdagio' && recentlyImportedSheetName !== 'FromShopify' && recentlyImportedSheetName !== 'Discounts')
            spreadsheet.deleteSheet(sheets[sheet]) // Delete the new sheet that was created

          break;
        }
      }
    }
    catch (err)
    {
      var error = err['stack'];
      Logger.log(error);
      Browser.msgBox('Please contact the spreadsheet owner and let them know what action you were performing that lead to the following error: ' + error)
      throw new Error(error);
    }
  }
}

function installedOnEdit(e)
{  
  const ss  = e.source;
  const rng = e.range;
  const row = rng.rowStart;
  const isSingleRow = row === rng.rowEnd
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getSheetName();
  
  try
  {
    if (sheetName == 'Dashboard' && rng.columnStart == 5 && rng.columnEnd == 5 && isSingleRow && rng.isChecked())
    {
      rng.uncheck();
      SpreadsheetApp.flush();

      if (row === 10)
        PriceUpdatesGrouped(ss)
      else if (row === 11)
        SaleItems(ss)
    }
    else if (sheetName == 'FromShopify' && rng.columnEnd > 40)
      replaceLeadingApostrophesOnVariantSKUs(sheet, rng.columnEnd, rng.rowEnd, ss);
    else if (sheetName == 'FromAdagio' && rng.columnEnd > 24)
      ss.getSheetByName("Dashboard").getRange(9, 6).setValue(timeStamp()).activate(); // Timestamp on dashboard
    else if (sheetName == 'FromWebsiteDiscount' && rng.columnEnd > 5)
    {
      ss.toast('Accessing the Discount Percentages...', '', -1);
      const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages');
      const totalNumDiscounts = discountSheet.getLastRow() - 1;
      const discounts = discountSheet.getSheetValues(2, 11, totalNumDiscounts, 5);
      var discountedItem;

      ss.toast('Discount Percentages acquired. Updating the Shopify discounts (Approx 110 seconds)...', '', -1);

      const range = sheet.getRange(1, 1, rng.rowEnd, rng.columnEnd);
      const values = range.getValues().map(webItem => {
        discountedItem = discounts.find(item => item[0].split(' - ').pop().toString().toUpperCase() === webItem[1].toString().toUpperCase());
        return (discountedItem != null) ? [webItem[0], webItem[1], webItem[2], discountedItem[2], discountedItem[3], discountedItem[4]] : webItem.slice(0, 6);
      })

      ss.toast('Shopify discounts updated. Writing data to Discounts sheet...', '', -1);
      
      sheet.clearContents().getRange(1, 1, values.length, values[0].length).setValues(values).activate();

      ss.toast('Discounts sheet successfully updated.', 'Complete', 20)

      sheet.setFrozenRows(1);

      ss.getSheetByName('Dashboard').getRange(15, 6).setValue(totalNumDiscounts);
    }
  }
  catch (error)
  {
    Logger.log(error);
    Browser.msgBox(error);
  }
}

/**
 * 
 */
function PriceUpdatesGrouped(ss)
{
  ss.toast('Price updates beginning...','')
  var startTime = new Date().getTime();
  
  const   BLUE = "#e8ecf9";
  const  GREEN = "#93c47d";
  const        MASTER_SKU = 0;
  const               SKU = 6;
  const             PRICE = 7;
  const  COMPARE_AT_PRICE = 8;
  const DASHBOARD_QTY_COL = 6;
  const     DASHBOARD_ROW = 10;
  
  const sheet = ss.getSheetByName('Prices');
  const dashboard = ss.getSheetByName('Dashboard');
  const adagioSheet = ss.getSheetByName('FromAdagio');
  const shopifySheet = ss.getSheetByName('FromShopify');
  var numItems_Adagio, numItems_Shopify, adagioData = [], shopifyData = [], highlightItems_Red = [], highlightItems_Green = [], masterSkuList = [], shopifyPrices = [[], []],
  fontColours = [], fontWeights = [], numberFormats = [];

  [ adagioData, numItems_Adagio ] = generateData( adagioSheet);
  [shopifyData, numItems_Shopify] = generateData(shopifySheet,  "Variant Compare At Price");

  for (var i = 1; i < numItems_Shopify; i++)
  {
    for (var j = 1; j < numItems_Adagio; j++)
    {
      // Shopify item SKU is not blank (i.e. it is not a Picture line) and the SKUs match and the prices are different and the item is not on sale and the web price is not $0.00.
      if ((shopifyData[i][SKU] !== '') && (adagioData[j][SKU].toString().toLowerCase() == shopifyData[i][SKU].toString().toLowerCase()) && 
          (Number(adagioData[j][PRICE]) != Number(shopifyData[i][PRICE])) && (shopifyData[i][COMPARE_AT_PRICE] === '') && shopifyData[i][PRICE] != 0)
      {
        // Determine which items are price increase and which are decrease, then put the Adagio price into the shopify data (except if the item is on sale)
        if (Number(adagioData[j][PRICE]) > Number(shopifyData[i][PRICE]))
          highlightItems_Red.push(shopifyData[i][SKU])
        else
          highlightItems_Green.push(shopifyData[i][SKU]);

        shopifyPrices[0].push(shopifyData[i][SKU])
        shopifyPrices[1].push(shopifyData[i][PRICE])
        shopifyData[i][PRICE] = adagioData[j][PRICE]

        if (!masterSkuList.includes(shopifyData[i][MASTER_SKU])) masterSkuList.push(shopifyData[i][MASTER_SKU]); // Add the master sku to the list (if it is not already there)

        break; // Break the Adagio for-loop because SKUs are unique, so once you have found a matching SKU in the Adagio DB, then there are NO more
      }
    }
  }

  if (masterSkuList.length !== 0)
  {
    const data = shopifyData.filter(value => masterSkuList.includes(value[MASTER_SKU]) && value[SKU] !== '')
    data.map(u => (shopifyPrices[0].includes(u[SKU])) ? u.splice(8, 1, '', shopifyPrices[1][shopifyPrices[0].indexOf(u[SKU])]) : u.splice(8, 1, '', ''))
    var items_TwoOptions = [], items_OneOption = [], items_ZeroOptions;
    shopifyData[0].splice(8, 1, '', '')

    items_ZeroOptions = data.filter(item => {
      if (item[5] !== '') // Option2 Value is not blank
        items_TwoOptions.push(item);
      else if (item[3] !== 'Default Title') // Option1 Value is not 'Default Title'
        items_OneOption.push(item)
      else
        return true;

      return false
    });

    var groupedData = [['', '', '', '', '', '', '', '', '', ''], 
                      shopifyData[0]]
                      .concat((items_TwoOptions.length !== 0) ? items_TwoOptions : [['No items found that require a price change in this category.', '', '', '', '', '', 'Two Options', '', '', '']],
                        [['', '', '', '', '', '', '', '', '', '']],
                        [shopifyData[0]],
                        (items_OneOption.length !== 0) ? items_OneOption : [['No items found that require a price change in this category.', '', '', '', '', '', 'One Option', '', '', '']],
                        [['', '', '', '', '', '', '', '', '', '']],
                        [shopifyData[0]],
                        (items_ZeroOptions.length !== 0) ? items_ZeroOptions : [['No items found that require a price change in this category.', '', '', '', '', '', 'Zero Options', '', '', '']])
  }
  else
  {
    var groupedData = [ ['', '', '', '', '', '', '', '', '', ''],                                                                             // Blank row
                          shopifyData[0],                                                                                                     // Shopify headers
                          ['No items found that require a price change in this category.', '', '', '', '', '', 'Two Options', '', '', ''],    // Blank row
                          ['', '', '', '', '', '', '', '', '', ''],                                                                           // Blank row
                          shopifyData[0],                                                                                                     // Shopify headers
                          ['No items found that require a price change in this category.', '', '', '', '', '', 'One Option', '', '', ''],     // Blank row
                          ['', '', '', '', '', '', '', '', '', ''],                                                                           // Blank row
                          shopifyData[0],                                                                                                     // Shopify headers
                          ['No items found that require a price change in this category.', '', '', '', '', '', 'Zero Options', '', '', '']  ];// Blank row
  }

  var backgroundColours = groupedData.map(sku => {
    if (highlightItems_Red.includes(sku[6])) // Price Increase
    {
      fontColours.push(["Black", "Black", "Black", "Black", "Black", "Black", "Black", "Red", "Black", "#b7b7b7"])         // Make the Variant Price font Red
      fontWeights.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'normal', 'bold']) // Make the prices bold (variant and compare at)
      numberFormats.push(['@', '@', '@', '@', '@', '@', '@', '0.00', '@', '0.00'])                                       // Make prices have 2 decimals
      return [BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, 'white', 'white']                                          // Highlight entire row Blue
    }
    else if (highlightItems_Green.includes(sku[6])) // Price Decrease
    {
      fontColours.push(["Black", "Black", "Black", "Black", "Black", "Black", "Black", "Green", "Black", "#b7b7b7"])       // Make the Variant Price font Green
      fontWeights.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold', 'normal', 'bold']) // Make the prices bold (variant and compare at)
      numberFormats.push(['@', '@', '@', '@', '@', '@', '@', '0.00', '@', '0.00'])                                       // Make prices have 2 decimals
      return [BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, 'white', 'white']                                          // Highlight entire row Blue
    }
    else if (sku[6] === '' || sku[6] === 'Variant SKU') // Repeated Header Row
    {
      fontColours.push(["Black", "Black", "Black", "Black", "Black", "Black", "Black", "Black", "Black", "Black"])           // Black font colour for entire row
      fontWeights.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal']) // Make the Compare At Price bold
      numberFormats.push(['@', '@', '@', '@', '@', '@', '@', '@', '@', '0.00'])                                              // Make prices have 2 decimals
      return [GREEN, GREEN, GREEN, GREEN, GREEN, GREEN, GREEN, GREEN, 'white', 'white']                                      // No highlighting (white background colour)
    }
    else // No change to the price
    {
      fontColours.push(["Black", "Black", "Black", "Black", "Black", "Black", "Black", "Black", "Black", "#b7b7b7"])         // Black font colour for entire row
      fontWeights.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold']) // Make the Compare At Price bold
      numberFormats.push(['@', '@', '@', '@', '@', '@', '@', '0.00', '@', '0.00'])                                         // Make prices have 2 decimals
      return ["White", "White", "White", "White", "White", "White", "White", "White", 'white', 'white']                    // No highlighting (white background colour)
    }
  });

  const numHighlights = highlightItems_Red.length + highlightItems_Green.length;
  const formattedDate = timeStamp();
  const runTime = elapsedTime(startTime);
  const numRows = groupedData.length;
  const numCols = groupedData[0].length;
  
  sheet.clearContents().getRange(2, 2, 3, numCols).setValues([['', 'Price Updates (Grouped)', '', '', '', '', 'Price Increase', 'Price Decrease', '', 'Shopify Price'],
                                                              [numHighlights + ' Highlighted', 'Elapsed Time:', runTime, '', 'Timestamp:', '', formattedDate, '', '', ''],
                                                              shopifyData[0]])
  sheet.getRange(5, 2, numRows, numCols).clearFormat().setFontWeights(fontWeights).setFontColors(fontColours).setNumberFormats(numberFormats).setBackgrounds(backgroundColours)
    .setBorder(false, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setValues(groupedData); 

  sheet.getRange(5, 10, numRows).setBorder(false, true, false, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)

  if (sheet.getMaxRows() > numRows + 5)
    sheet.deleteRows(numRows + 5, sheet.getMaxRows() - numRows - 4); // Delete extra rows if there are any

  dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 2).setValues([[numHighlights, formattedDate]]);
  ss.toast('','Price updates complete.')
}

function SaleItems(ss)
{
  ss.toast('Sale Items beginning...','')
  var startTime = new Date().getTime();
  
  const YELLOW = "#ffffbf";
  const        MASTER_SKU = 0;
  const               SKU = 6;
  const             PRICE = 7;
  const  COMPARE_AT_PRICE = 8;
  const  spreadsheet = SpreadsheetApp.getActive();
  const        sheet = spreadsheet.getSheetByName('Sale Items');
  const  adagioSheet = spreadsheet.getSheetByName('FromAdagio');
  const shopifySheet = spreadsheet.getSheetByName('FromShopify');
  var numItems_Adagio, numItems_Shopify, adagioData = [], shopifyData = [], saleItems = [], highlightItems_Yellow = [], masterSkuList = [], fontWeights = [], numberFormats = [];

  [ adagioData, numItems_Adagio ] = generateData( adagioSheet);
  [shopifyData, numItems_Shopify] = generateData(shopifySheet,  "Variant Compare At Price");

  for (var i = 1; i < numItems_Shopify; i++)
  {
    for (var j = 1; j < numItems_Adagio; j++)
    {
      // Shopify item SKU is not blank (i.e. it is not a Picture line) and the SKUs match and the prices are different
      if (((shopifyData[i][SKU] !== '') && (adagioData[j][SKU].toString().toLowerCase() == shopifyData[i][SKU].toString().toLowerCase()) 
                                        && (Number(adagioData[j][PRICE]) != Number(shopifyData[i][PRICE]))) 
                                        && (shopifyData[i][PRICE] == 0 || shopifyData[i][COMPARE_AT_PRICE] !== '')) // And one of either the web price is $0.00 or the items is on sale
      {
        (shopifyData[i][PRICE] == 0) ? highlightItems_Yellow.push(shopifyData[i][SKU]) : saleItems.push(shopifyData[i][SKU]); 

        if (!masterSkuList.includes(shopifyData[i][MASTER_SKU])) masterSkuList.push(shopifyData[i][MASTER_SKU]); // Add the master sku to the list (if it is not already there)

        break; // Break the Adagio for-loop because SKUs are unique, so once you have found a matching SKU in the Adagio DB, then there are NO more
      }
    }
  }

  const data = shopifyData.filter(value => masterSkuList.includes(value[MASTER_SKU]) && value[SKU] !== ''); // Keep the items that belong to the master sku list and are not picture lines

  var backgroundColours = data.map(sku => {
    
    fontWeights.push(['normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'normal', 'bold'])  // Make the Compare At Price bold
    numberFormats.push(['@', '@', '@', '@', '@', '@', '@', '0.00', '0.00'])                                     // Make prices have 2 decimals

    return (highlightItems_Yellow.includes(sku[6])) ? [YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW, YELLOW] 
                                                    : ["White", "White", "White", "White", "White", "White", "White", "White", "White"];
  });

  shopifyData[0][8] = 'Compare At Price'; // Remove the word variant from the description to reduce size of string
  sheet.clearContents().getRange(2, 2, 3, 9).setValues([[saleItems.length + ' Items On Sale', 'Items on Sale', '', '', '', '', '', 'Shopify Price = $0.00', ''],
                                                        [highlightItems_Yellow.length + ' Highlighted', 'Elapsed Time:', elapsedTime(startTime), '', 'Timestamp:', '', timeStamp(), '', ''],
                                                        shopifyData[0]])
  sheet.getRange(5, 2, data.length, 9).clearFormat().setFontWeights(fontWeights).setFontColor('black').setNumberFormats(numberFormats).setBackgrounds(backgroundColours)
    .setBorder(false, true, true, true, false, false,'black',SpreadsheetApp.BorderStyle.SOLID_THICK).setValues(data);

  if (sheet.getMaxRows() > data.length + 5)
    sheet.deleteRows(data.length + 5, sheet.getMaxRows() - data.length - 4); // Delete extra rows if there are any

  ss.getSheetByName('Dashboard').getRange(11, 6, 1, 2).setValues([[saleItems.length, timeStamp()]]);

  ss.toast('','Sale Items complete.')
}

function DuplicateSKUs()
{
  var startTime = new Date().getTime();
  
  const         SKU = 6;
  const NUM_HEADERS = 3;
  const DASHBOARD_QTY_COL = 7;
  const     DASHBOARD_ROW = 8;
  
  var spreadsheet = SpreadsheetApp.getActive();
  var     sheet = spreadsheet.getSheetByName("Duplicates");
  var dashboard = spreadsheet.getSheetByName("Dashboard");
  var shopifyData = [], dupSKUsIndices = [], uniqueSKUs = [];
  var outputData = [[]];
  var sku, numItems_Shopify;
  
  var shopifySheet = spreadsheet.getSheetByName("FromShopify");
  [shopifyData, numItems_Shopify] = generateData(shopifySheet);
  
  // Find all of the duplicate SKUs
  for (var j = 1; j < numItems_Shopify; j++)
  {
    sku = shopifyData[j][SKU];
    if (sku !== '') // Some SKUs are empty becasue they represent addiional pictures
      (uniqueSKUs.indexOf(sku) === -1) ? uniqueSKUs.push(sku) : dupSKUsIndices.push(j);
  }
  
  // Set up all of the Data
  var numItems = dupSKUsIndices.length;
  outputData[0] = ['Duplicate SKUs on Shopify',              '', '', '',           '', '', '', ''];
  outputData[1] = [                         '', 'Elapsed Time:', '', '', 'Timestamp:', '', '', ''];
  outputData[2] = shopifyData[0];
  for (var u = 0; u < numItems; u++)
    outputData[u + NUM_HEADERS] = shopifyData[dupSKUsIndices[u]];
  
  sheet.clearContents(); // Clears the content of the entire sheet
  
  // Set all of the info on the sheet
  var formattedDate = timeStamp();
  var runTime = elapsedTime(startTime)
  outputData[1][0] = numItems;
  outputData[1][6] = formattedDate;
  outputData[1][2] = runTime;
  sheet.getRange(2, 2, numItems + NUM_HEADERS, shopifyData[0].length).setValues(outputData);
  //dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems, null, null, null, formattedDate, null, runTime]]);
}

function ItemsMissingFromShopify()
{
  var startTime = new Date().getTime();
  
  const BLUE = "#e8ecf9";
  const  MASTER_SKU = 0;
  const         SKU = 6;
  const DASHBOARD_QTY_COL =  7;
  const     DASHBOARD_ROW = 10;

  const sheets = SpreadsheetApp.getActive().getSheets(); // Get all of the sheets in the spreadsheet
  const sheetNames = sheets.map(s => s.getSheetName());  // Get all of the names of the sheets in the spreasheet
  const      sheet = sheets[sheetNames.indexOf("Missing Items")]; // Use the sheet names to retrieve the desired sheets
  const  dashboard = sheets[sheetNames.indexOf("Dashboard")];
  const  adagioSheet = sheets[sheetNames.indexOf("FromAdagio")];
  const shopifySheet = sheets[sheetNames.indexOf("FromShopify")];

  var numItems_Adagio, numItems_Shopify, adagioData = [], shopifyData = [], highlightItems = [], masterSkuList = [], mstrSKU, itemIndices_Adagio = [], outputData = [[]]; 

  [ adagioData, numItems_Adagio ] = generateData( adagioSheet);
  [shopifyData, numItems_Shopify] = generateData(shopifySheet);
  
  for (var i = 1; i < numItems_Adagio; i++)
  {
    for (var j = 1; j < numItems_Shopify; j++)
    {
      if ((shopifyData[j][SKU] !== '') && (adagioData[i][SKU].toLowerCase() == shopifyData[j][SKU].toLowerCase()))
        break; // Break the Shopify for-loop because SKUs are unique, so once you have found a matching SKU in the Shopify DB, then there are NO more
    }
    
    if (j == numItems_Shopify) // This means the second for loop went till completion without finding the item
    {
      highlightItems.push(adagioData[i][SKU]); // Will be used below to determine which lines to highlight
      
      if (!masterSkuList.includes(adagioData[i][MASTER_SKU])) masterSkuList.push(adagioData[i][MASTER_SKU]); // Add the master sku to the list (if it is not already there)
    }
  }
  
  const data = adagioData.filter(value => masterSkuList.includes(value[MASTER_SKU]) && value[SKU] !== ''); // Keep the items that belong to the master sku list and are not picture lines

  const backgroundColours = data.map(sku => {
    if (highlightItems.includes(sku[6])) // Price Increase
      return [BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE] // Highlight entire row Blue
    else // No change to the price
      return ["White", "White", "White", "White", "White", "White", "White", "White"]  // No highlighting (white background colour)
  });

  const numHighlights = highlightItems.length;
  const formattedDate = timeStamp();
  const runTime = elapsedTime(startTime);
  
  sheet.clearContents().getRange(2, 2, 3, 8).setValues([['Items Not On Shopify', '', '', '', '', '', '', ''],
                                                        [numHighlights + ' Highlighted', 'Elapsed Time:', runTime, '', 'Timestamp:', '', formattedDate, ''],
                                                        adagioData[0]])
  sheet.getRange(5, 2, data.length, data[0].length).clearFormat().setBackgrounds(backgroundColours)
    .setBorder(false, true, true, true, false, false,'black',SpreadsheetApp.BorderStyle.SOLID_THICK).setValues(data);

  if (sheet.getMaxRows() > data.length + 5)
    sheet.deleteRows(data.length + 5, sheet.getMaxRows() - data.length - 4); // Delete extra rows if there are any

  //dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems_NOT_Sale, null, null, null, formattedDate, null, runTime]]);
}

function MissingImages()
{
  var startTime = new Date().getTime();
  
  const   MASTER_SKU = 0;
  const  DESCRIPTION = 1;
  const    PUBLISHED = 2;
  const IMAGE_SOURCE = 9;
  const  NUM_HEADERS = 3;
  const DASHBOARD_QTY_COL =  7;
  const     DASHBOARD_ROW = 12;
  
  var spreadsheet = SpreadsheetApp.getActive();
  var     sheet = spreadsheet.getSheetByName("Images");
  var dashboard = spreadsheet.getSheetByName("Dashboard");
  var numItems_Shopify;
  var shopifyData_ = [];
  var headers = [[]]; // The headers array is needed in this case because the outputData is going to get sorted
  
  var shopifySheet = spreadsheet.getSheetByName("FromShopify");
  [shopifyData_, numItems_Shopify] = generateData(shopifySheet, "Published", "Image Src");

  var shopData = shopifyData_.filter(value => value[DESCRIPTION] != '' && value[PUBLISHED] == "TRUE" && value[IMAGE_SOURCE] == '');
  var shopifyData__  = shopData.filter(value => value.splice(IMAGE_SOURCE)); // Remove the 'Image Source' column
  var data = shopifyData__.filter(value => value.splice(PUBLISHED, 1));      // Remove the 'Published' column
  var numItems = data.length;

  sheet.clearContents(); // Clears the content of the entire sheet

  // Set up all of the Data
  headers[0] = ['Images Missing On Shopify',              '', '', '',           '', '', '', ''];
  headers[1] = [                         '', 'Elapsed Time:', '', '', 'Timestamp:', '', '', ''];
  var head  = shopifyData_.shift();
  head.splice(IMAGE_SOURCE); // Remove the 'Image Source' column
  head.splice(PUBLISHED, 1); // Remove the 'Published' column
  headers[2] = head;
  data.sort(sortBySecondColumn); // Sort the data by the Title column
  
  // Set all of the info on the sheet
  var outputData = headers.concat(data);
  var formattedDate = timeStamp();
  var runTime = elapsedTime(startTime)
  outputData[1][0] = numItems;
  outputData[1][6] = formattedDate;
  outputData[1][2] = runTime;
  sheet.getRange(2, 2, numItems + NUM_HEADERS, shopData[0].length).setValues(outputData);
  //dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems, null, null, null, formattedDate, null, runTime]]);
}

function MissingWeights()
{
  var startTime = new Date().getTime();
  
  const   MASTER_SKU = 0;
  const  DESCRIPTION = 1;
  const    PUBLISHED = 2;
  const       WEIGHT = 8;
  const  NUM_HEADERS = 3;
  const DASHBOARD_QTY_COL =  7;
  const     DASHBOARD_ROW = 14;
  
  var spreadsheet = SpreadsheetApp.getActive();
  var     sheet = spreadsheet.getSheetByName("Weights");
  var dashboard = spreadsheet.getSheetByName("Dashboard");
  var numItems_Shopify;
  var shopifyData_ = [];
  var headers = [[]]; // The headers array is needed in this case because the outputData is going to get sorted
  
  var shopifySheet = spreadsheet.getSheetByName("FromShopify");
  [shopifyData_, numItems_Shopify] = generateData(shopifySheet, "Published", "Variant Grams");

  var shopData = shopifyData_.filter(value => value[DESCRIPTION] != '' && value[PUBLISHED] == "TRUE" && value[WEIGHT] == '0');
  var shopifyData__  = shopData.filter(value => value.splice(WEIGHT, 1)); // Remove the 'Weight' column
  var data = shopifyData__.filter(value => value.splice(PUBLISHED, 1));   // Remove the 'Published' column
  var numItems = data.length;

  sheet.getRange("A:J").clearContent();// Clears the content of the "A through J"

  // Set up all of the Data
  headers[0] = ['Dumb Jarrens Missing On Shopify',              '', '', '',           '', '', '', ''];
  headers[1] = [                         '', 'Elapsed Time:', '', '', 'Timestamp:', '', '', ''];
  var head  = shopifyData_.shift();
  head.splice(WEIGHT, 1);    // Remove the "Weight" header
  head.splice(PUBLISHED, 1); // Remove the "Published" header
  headers[2] = head;
  data.sort(sortBySecondColumn); // Sort the data by the Title column
  
  // Set all of the info on the sheet
  var outputData = headers.concat(data);
  var formattedDate = timeStamp();
  var runTime = elapsedTime(startTime)
  outputData[1][0] = numItems;
  outputData[1][6] = formattedDate;
  outputData[1][2] = runTime;
  sheet.getRange(5, 2, numItems + NUM_HEADERS, shopData[0].length).setValues(outputData);
  //dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems, null, null, null, formattedDate, null, runTime]]);
}

function ItemsMissingFromShopify_NonGrouped()
{
  const startTime = new Date().getTime();
  const  ADAGIO_SKU = 7;
  const SHOPIFY_SKU = 6;
  const NUM_HEADERS = 3;
  const DASHBOARD_QTY_COL =  7;
  const     DASHBOARD_ROW = 16;
  const sheets = SpreadsheetApp.getActive().getSheets();
  const sheetNames = sheets.map(sheet => sheet.getSheetName());
  const [ adagioData] = generateData(sheets[sheetNames.indexOf( "FromAdagio")], 'Active Item', 'disabled Body');
  const [shopifyData] = generateData(sheets[sheetNames.indexOf("FromShopify")]);
  const data = adagioData.filter(arr1 => shopifyData.filter(arr2 => arr1[ADAGIO_SKU] == arr2[SHOPIFY_SKU]).length == 0);
  const formattedDate = timeStamp();
  const numItems = data.length;
  const runTime = elapsedTime(startTime)
  const outputData =  [ ['Items not on Shopify - Non-Grouped (Missing Items ONLY)',               '',       '', '',           '', '',            '', '', '', ''],
                        [                                      numItems + ' Items',  'Elapsed Time:',  runTime, '', 'Timestamp:', '', formattedDate, '', '', ''],
                        adagioData[0]
  ].concat(data);
  sheets[sheetNames.indexOf("Missing - Non-Grouped")].clearContents().getRange(2, 2, numItems + NUM_HEADERS, adagioData[0].length).setNumberFormat('@').setValues(outputData);
  //sheets[sheetNames.indexOf("Dashboard")].getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems, null, null, null, formattedDate, null, runTime]]);
}

function ItemsMissingFromShopify_Grouped()
{
  const startTime = new Date().getTime();
  
  const  MASTER_SKU = 0;
  const  ADAGIO_SKU = 7;
  const SHOPIFY_SKU = 6;
  const NUM_HEADERS = 3;
  const DASHBOARD_QTY_COL =  7;
  const     DASHBOARD_ROW = 18;
  const   BLUE = "#e8ecf9";
  const sheets = SpreadsheetApp.getActive().getSheets();
  const sheetNames = sheets.map(sheet => sheet.getSheetName());
  const sheet = sheets[sheetNames.indexOf("Missing - Grouped")];
  const [ adagioData] = generateData(sheets[sheetNames.indexOf( "FromAdagio")], 'Active Item', 'disabled Body');
  const [shopifyData] = generateData(sheets[sheetNames.indexOf("FromShopify")]);
  const nonGrouped_MissingItems = adagioData.filter(arr1 => shopifyData.filter(arr2 => arr1[ADAGIO_SKU] == arr2[SHOPIFY_SKU]).length == 0);
  var highlightedRows = [], index = 3, isGroupedItem, isNonGroupedItem;

  const data = adagioData.filter(item1 => {
    isNonGroupedItem = nonGrouped_MissingItems.filter(item2 => {
      isGroupedItem = item1[MASTER_SKU] == item2[MASTER_SKU];
      if (isGroupedItem)
      {
        if (item1[ADAGIO_SKU] == item2[ADAGIO_SKU])
          highlightedRows.push(index);
      }
      return isGroupedItem;
    }).length != 0;
    if (isNonGroupedItem) index++;
    return isNonGroupedItem;
  })

  const formattedDate = timeStamp();
  const numItems = data   .length;
  const numCols  = data[0].length;
  const colours = new Array(numItems + NUM_HEADERS).fill(null).map((_, idx) => (highlightedRows.indexOf(idx) === -1) ? [...new Array(numCols).fill(null)] : [...new Array(numCols).fill(BLUE)]);
  
  const runTime = elapsedTime(startTime)
  const outputData =  [ ['Items not on Shopify - Non-Grouped (Missing Items ONLY)',               '',       '', '',           '', '',            '', '', '', ''],
                        [                                      numItems + ' Items',  'Elapsed Time:',  runTime, '', 'Timestamp:', '', formattedDate, '', '', ''],
                        adagioData[0]
  ].concat(data);
  sheet.getRange('B:K').clearContent()//.setBackground(null);
  sheet.getRange(2, 2, numItems + NUM_HEADERS, numCols).setBackgrounds(colours).setNumberFormat('@').setValues(outputData);
  //sheets[sheetNames.indexOf("Dashboard")].getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems, null, null, null, formattedDate, null, runTime]]);
}

function DescripNotMatching()
{
  var startTime = new Date().getTime();
  
  const  MASTER_SKU = 0;
  const DESCRIPTION = 1;
  const         SKU = 6;
  const NUM_HEADERS = 3;
  const       DASHBOARD_QTY_COL =  7;
  const           DASHBOARD_ROW = 20;
  
  var spreadsheet = SpreadsheetApp.getActive();
  var     sheet = spreadsheet.getSheetByName("Descrips Not Matching");
  var dashboard = spreadsheet.getSheetByName("Dashboard");
  var mstrSKU, numItems_Adagio, numItems_Shopify;
  var adagioData = [], shopifyData = [], itemIndices_Shopify = [];
  var headers = [[]], data = [[]]; // Initialize a double array representing the output data
  
  var  adagioSheet = spreadsheet.getSheetByName("FromAdagio");
  var shopifySheet = spreadsheet.getSheetByName("FromShopify");
  [ adagioData, numItems_Adagio ] = generateData( adagioSheet);
  [shopifyData, numItems_Shopify] = generateData(shopifySheet);
  
  var blankDescriptions = shopifyData.filter(value => value[DESCRIPTION] != ''); // Remove items with blank description
  
  for (var j = 1; j < numItems_Shopify; j++)
  {
    for (var i = 1; i < numItems_Adagio; i++)
    {
      if ((shopifyData[j][SKU] !== '') && (adagioData[i][SKU].toLowerCase() == shopifyData[j][SKU].toLowerCase())) 
      {
        mstrSKU = shopifyData[j][MASTER_SKU].toLowerCase(); // The master SKU is never blank
        
        for (var k = 0; k < blankDescriptions.length; k++)
        {
          if ((mstrSKU == blankDescriptions[k][MASTER_SKU].toLowerCase()) && (blankDescriptions[k][DESCRIPTION].toLowerCase() != adagioData[i][DESCRIPTION].toLowerCase()))
          {
            shopifyData[j][DESCRIPTION] = blankDescriptions[k][DESCRIPTION];
            itemIndices_Shopify.push([j, i]);
          }
        }
      }
    }
  }
  
  sheet.clearContents(); // Clears the content of the entire sheet
  
  var uniqueIndices = itemIndices_Shopify.filter((value, index, arr) => arr.indexOf(value) === index); // Removes the duplicates from the original list 
  var uniqueIndices = uniqByKeepFirst(itemIndices_Shopify, value => value[0])
  var numItems = uniqueIndices.length;

  // Set up all of the Data
  headers[0] = ['Descriptions Not Matching',              '', '',           '', '', '', '', '', ''];
  headers[1] = [                         '', 'Elapsed Time:', '', 'Timestamp:', '', '', '', '', ''];
  shopifyData[0].splice(2, 0, "Adagio Description");
  shopifyData[0][DESCRIPTION] = "Shopify Description"
  headers[2] = shopifyData[0];
  for (var u = 0; u < numItems; u++)
  {
    shopifyData[uniqueIndices[u][0]].splice(2, 0, adagioData[uniqueIndices[u][1]][DESCRIPTION]);
    data[u] = shopifyData[uniqueIndices[u][0]];
  }

  data.sort(sortBySecondColumn); // Sort the data by the Title column
  
  // Set all of the info on the sheet
  var outputData = headers.concat(data);
  var formattedDate = timeStamp();
  var runTime = elapsedTime(startTime)
  outputData[1][0] = numItems;
  outputData[1][6] = formattedDate;
  outputData[1][2] = runTime;
  sheet.getRange(2, 2, numItems + NUM_HEADERS, shopifyData[0].length).setValues(outputData);
  //dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 7).setValues([[numItems, null, null, null, formattedDate, null, runTime]]);
}

function OnWebWithNoInventory()
{
  var startTime = new Date().getTime();
  
  const BLUE = "#e8ecf9";
  const MASTER_SKU = 0;
  const        QTY = 2;
  const        SKU = 6;
  const  spreadsheet = SpreadsheetApp.getActive();
  const        sheet = spreadsheet.getSheetByName('On Web No Inventory');
  const shopifySheet = spreadsheet.getSheetByName('FromShopify');
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const itemNum = adagioData[0].indexOf('Item #')
  var numItems_Shopify, data = [], shopifyData = [], numItems_Adagio = adagioData.length, adagioQty = [[], []], masterSkuList = [], backgroundColours = [];

  [shopifyData, numItems_Shopify] = generateData(shopifySheet);

  for (var i = 1; i < numItems_Shopify; i++)
  {
    for (var j = 1; j < numItems_Adagio; j++)
    {
      // Shopify item SKU is not blank (i.e. it is not a Picture line) and the SKUs match
      if (((shopifyData[i][SKU] !== '') && (adagioData[j][itemNum].toString().toLowerCase() == shopifyData[i][SKU].toString().toLowerCase()))
                                        && adagioData[j][QTY] <= 0)
      {
        adagioQty[0].push(shopifyData[i][SKU]);
        (adagioData[j][QTY] == 0) ? adagioQty[1].push(0) : adagioQty[1].push(adagioData[j][QTY]);

        if (!masterSkuList.includes(shopifyData[i][MASTER_SKU])) masterSkuList.push(shopifyData[i][MASTER_SKU]); // Add the master sku to the list (if it is not already there)

        break; // Break the Adagio for-loop because SKUs are unique, so once you have found a matching SKU in the Adagio DB, then there are NO more
      }
    }
  }

  shopifyData[0].splice(8, 0, '', '')

  if (masterSkuList.length !== 0)
  {
    var data = shopifyData.filter(value => masterSkuList.includes(value[MASTER_SKU]) && value[SKU] !== '')
      .map(u => {

        if (adagioQty[0].includes(u[SKU]))
        {
          backgroundColours.push([BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, 'white', BLUE])
          u.splice(8, 0, '', adagioQty[1][adagioQty[0].indexOf(u[SKU])])

          return u
        }
        else
        {
          backgroundColours.push(['white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
          u.splice(8, 0, '', '')

          return u
        }
      })
  }

  const numRows = data.length;
  const numCols = 10;

  sheet.clearContents().getRange(2, 2, 3, numCols).setValues([['Items on Web with No Inventory', '', '', '', '', '', '', '', '', 'Qty'],
                                                        [adagioQty[0].length + ' Highlighted', 'Elapsed Time:', elapsedTime(startTime), '', 'Timestamp:', '', timeStamp(), '', '', ''],
                                                        shopifyData[0]])

  sheet.getRange(5, 2, numRows, numCols).clearFormat()
    .setBorder(false, true, true, true, false, false,'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackgrounds(backgroundColours).setValues(data);

  sheet.getRange(5, numCols, numRows).setBorder(false, true, false, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)

  if (sheet.getMaxRows() > numRows + 5)
    sheet.deleteRows(numRows + 5, sheet.getMaxRows() - numRows - 4); // Delete extra rows if there are any
}

function NotOnWebWithInventory()
{
  var startTime = new Date().getTime();
  
  const BLUE = "#e8ecf9";
  const MASTER_SKU = 0;
  const        QTY = 2;
  const        SKU = 6;
  const  spreadsheet = SpreadsheetApp.getActive();
  const        sheet = spreadsheet.getSheetByName('Not On Web With Inventory');
  const shopifySheet = spreadsheet.getSheetByName('FromShopify');
  const adagioData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const itemNum = adagioData[0].indexOf('Item #')
  var data = [], shopifyData = [], numItems_Adagio = adagioData.length, adagioQty = [[], []], masterSkuList = [], backgroundColours = [];

  var fullData = shopifySheet.getDataRange().getValues();
  var sheetName = 'FromShopify';
  var str = "Following Header Titles Not Found On The " + sheetName + " Sheet:";
  var columnHeaderTitles = ["Handle", "Title", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU", "Variant Price"];
  var columnsToKeep = [];
  const STATUS = fullData[0].indexOf('Status');
  var numColHeaderTitles = columnHeaderTitles.length;
  
  for (var j = 0; j < numColHeaderTitles; j++)
  {
    for (var i = 0; i < fullData[0].length; i++)
    {
      if (fullData[0][i] == columnHeaderTitles[j])
      {
        columnsToKeep.push(i);
        break;
      }
      else if (i == fullData[0].length - 1) // We have reached the end of the list and haven't found the Header Title in the data
        str += ' ' + columnHeaderTitles[j] + ',';
    }
  }

  var header = fullData.shift().filter((_, index) => columnsToKeep.indexOf(index) !== -1);
  var shopifyData = fullData.filter((_, index, array) => {
    var n = index;
    while (!array[n][STATUS])
      n--;
    return array[n][STATUS] == 'archived'; // Keep the items that have a status of active or draft
  }).map(value => value.filter((_, index) => columnsToKeep.indexOf(index) !== -1)) // Keep only the columns that match the chosen headers
  shopifyData.unshift(header)

  var numItems_Shopify = shopifyData.length;
  var errorMessage = str.substring(0, str.length - 1); // Remove the last comma in order to replace it with a period
  errorMessage += ". To troubleshoot this issue: 1) Make sure the data was imported as expected."
               +  "\n    2) Make sure the column header is spelt exactly correct inside the function you just ran, and in the generateData() function."
               +  "\n\nOtherwise, consider making adjustments to the generateData() function."
  
  // If one of the headers couldn't be found, throw the error message
  if (numColHeaderTitles !== columnsToKeep.length)
    throw new Error(errorMessage);


  for (var i = 1; i < numItems_Shopify; i++)
  {
    for (var j = 1; j < numItems_Adagio; j++)
    {
      // Shopify item SKU is not blank (i.e. it is not a Picture line) and the SKUs match
      if (((shopifyData[i][SKU] !== '') && (adagioData[j][itemNum].toString().toLowerCase() == shopifyData[i][SKU].toString().toLowerCase()))
                                        && adagioData[j][QTY] > 0)
      {
        adagioQty[0].push(shopifyData[i][SKU]);
        adagioQty[1].push(adagioData[j][QTY]);

        if (!masterSkuList.includes(shopifyData[i][MASTER_SKU])) masterSkuList.push(shopifyData[i][MASTER_SKU]); // Add the master sku to the list (if it is not already there)

        break; // Break the Adagio for-loop because SKUs are unique, so once you have found a matching SKU in the Adagio DB, then there are NO more
      }
    }
  }

  shopifyData[0].splice(8, 0, '', '')

  if (masterSkuList.length !== 0)
  {
    var data = shopifyData.filter(value => masterSkuList.includes(value[MASTER_SKU]) && value[SKU] !== '')
      .map(u => {

        if (adagioQty[0].includes(u[SKU]))
        {
          backgroundColours.push([BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, BLUE, 'white', BLUE])
          u.splice(8, 0, '', adagioQty[1][adagioQty[0].indexOf(u[SKU])])

          return u
        }
        else
        {
          backgroundColours.push(['white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'])
          u.splice(8, 0, '', '')

          return u
        }
      })
  }

  const numRows = data.length;
  const numCols = 10;

  sheet.clearContents().getRange(2, 2, 3, numCols).setValues([['Items Not on Web with Inventory', '', '', '', '', '', '', '', '', 'Qty'],
                                                        [adagioQty[0].length + ' Highlighted', 'Elapsed Time:', elapsedTime(startTime), '', 'Timestamp:', '', timeStamp(), '', '', ''],
                                                        shopifyData[0]])

  sheet.getRange(5, 2, numRows, numCols).clearFormat()
    .setBorder(false, true, true, true, false, false,'black', SpreadsheetApp.BorderStyle.SOLID_THICK).setBackgrounds(backgroundColours).setValues(data);

  sheet.getRange(5, numCols, numRows).setBorder(false, true, false, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK)

  if (sheet.getMaxRows() > numRows + 5)
    sheet.deleteRows(numRows + 5, sheet.getMaxRows() - numRows - 4); // Delete extra rows if there are any
}

/**
* Reset the shopify data.
*/
function resetShopifyData()
{
  const DASHBOARD_ROW = 7;
  var sheetName = 'FromShopify';
  var csvFileName = "products_export_1.csv";
  resetData(sheetName, csvFileName, DASHBOARD_ROW);
}

/**
* Reset the adagio data.
*/
function resetAdagioData()
{
  const DASHBOARD_ROW = 9;
  var sheetName = 'FromAdagio';
  var csvFileName = "For Shopify.csv";
  resetData(sheetName, csvFileName, DASHBOARD_ROW);
}

/**
* This function resets the data of the chosen sheet.
*
* @param {String} sheetName     The name of the sheet which the data will be written to
* @param {String} csvFileName   The name of the csv file
* @param {Number} DASHBOARD_ROW The dashboard row number to paste values to.
*/
function resetData(sheetName, csvFileName, DASHBOARD_ROW)
{
  const DASHBOARD_QTY_COL = 6;
  const       NUM_HEADERS = 1;
  
  var spreadsheet = SpreadsheetApp.getActive();
  var   dashboard = spreadsheet.getSheetByName("Dashboard");
  var       sheet = spreadsheet.getSheetByName(sheetName);
  var data = Utilities.parseCsv(DriveApp.getFilesByName(csvFileName).next().getBlob().getDataAsString());
  var formattedDate = timeStamp();
  var numItems = data.length;

  sheet.clearContents();
  sheet.getRange(1, 1, numItems, data[0].length).setNumberFormat('@').setValues(data);
  dashboard.activate(); 
  dashboard.getRange(DASHBOARD_ROW, DASHBOARD_QTY_COL, 1, 2).setValues([[numItems - NUM_HEADERS, formattedDate]]);
}

function UpdateAllBeta()
{
  var startTime = new Date().getTime();
  PriceUpdates();
  DuplicateSKUs();
  ItemsMissingFromShopify();
  MissingImages();
  MissingWeights();
  DisabledOnWeb();
  DisabledInAdagio();
  DescripNotMatching();
  //SpreadsheetApp.getActive().getSheetByName("Dashboard").getRange(4, 11, 1, 3).setValues([[timeStamp(), null, elapsedTime(startTime)]]);
}

/**
* This function generates the data used to derive all of the sheets, including additional headers sent as a String representing the additional columns needed.
*
* @param   {Sheet}      sheet   The sheet that the imported Data will come from
* @param  {String[]} ...varArgs A variable number of arguments which will represent additional header titles
* @throws  errorMessage   If the headers in the data do not match what is expected. 
* @return {Object[][], Number} [data, numRows] The chosen (and relevant) set of data, along with the number of rows in the data
* @author Jarren Ralf
*/
function generateData(sheet, ...varArgs)
{
  var fullData = sheet.getDataRange().getValues();
  var sheetName = sheet.getSheetName();
  var str = "Following Header Titles Not Found On The " + sheetName + " Sheet:";
  var columnHeaderTitles = ["Handle", "Title", "Option1 Name", "Option1 Value", "Option2 Name", "Option2 Value", "Variant SKU", "Variant Price"];
  var columnsToKeep = [];
  const STATUS = fullData[0].indexOf('Status');
  
  // Add the additional arguments as column headers
  if (varArgs.length != 0)
    columnHeaderTitles.push(...varArgs);
  
  var numColHeaderTitles = columnHeaderTitles.length;
  
  for (var j = 0; j < numColHeaderTitles; j++)
  {
    for (var i = 0; i < fullData[0].length; i++)
    {
      if (fullData[0][i] == columnHeaderTitles[j])
      {
        columnsToKeep.push(i);
        break;
      }
      else if (i == fullData[0].length - 1) // We have reached the end of the list and haven't found the Header Title in the data
        str += ' ' + columnHeaderTitles[j] + ',';
    }
  }

  if (sheetName === 'FromShopify')
  {
    var header = fullData.shift().filter((_, index) => columnsToKeep.indexOf(index) !== -1);
    var data = fullData.filter((_, index, array) => {
      var n = index;
      while (!array[n][STATUS])
        n--;
      return array[n][STATUS] != 'archived'; // Keep the items that have a status of active or draft
    }).map(value => value.filter((_, index) => columnsToKeep.indexOf(index) !== -1)) // Keep only the columns that match the chosen headers
    data.unshift(header)
  }
  else
    var data = fullData.map(value => value.filter((_, index) => columnsToKeep.indexOf(index) !== -1)); // Keep only the columns that match the chosen headers

  var numRows = data.length;
  var errorMessage = str.substring(0, str.length - 1); // Remove the last comma in order to replace it with a period
  errorMessage += ". To troubleshoot this issue: 1) Make sure the data was imported as expected."
               +  "\n    2) Make sure the column header is spelt exactly correct inside the function you just ran, and in the generateData() function."
               +  "\n\nOtherwise, consider making adjustments to the generateData() function."
  
  // If one of the headers couldn't be found, throw the error message
  if (numColHeaderTitles !== columnsToKeep.length)
    throw new Error(errorMessage);

  return [data, numRows];
}

/**
 * This function is run when there is a File -> Import -> Insert new sheet(s) or Replace current sheet and it's purpose is to replace leading apostrophes on the variant skus.
 * 
 * @param {Sheet}          sheet    : The sheet containing the item Shopify item data
 * @param {Number}        numCols   : The number of columns in the data
 * @param {Number}        numRows   : The number of rows in the data
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @param {Object[][]}     data     : The dat containing item information
 * @author Jarren Ralf
 */
function replaceLeadingApostrophesOnVariantSKUs(sheet, numCols, numRows, spreadsheet, data)
{
  if (data != null)
  {
    var skuIndex = data[0].indexOf('Variant SKU');
    var range = sheet.getRange(1, skuIndex + 1, numRows);
    var values = data.map(sku => (sku[skuIndex][0] !== '\'') ? [sku[skuIndex]] : [sku[skuIndex][0].substring(1)]);
  }
  else
  {
    var header = sheet.getSheetValues(1, 1, 1, numCols);
    var range = sheet.getRange(1, header[0].indexOf('Variant SKU') + 1, numRows);
    var data = range.getValues();
    var values = data.map(sku => (sku[0][0] !== '\'') ? sku : [sku[0].substring(1)]);
  }

  range.setNumberFormat('@').setValues(values)
  spreadsheet.getSheetByName("Dashboard").getRange(7, 6).setValue(timeStamp()).activate();
}

/**
*
*
* @param {Object[][]} multiArr
* @param {Object}      value
* @return {Number}  The number of occurences of a particular value within a double array
*/
function countOccurences2D(multiArr, value)
{
  var arr = [].concat.apply([], multiArr);
  
  return arr.reduce((acc, element) => { return (value === element ? acc + 1 : acc) }, 0);
}

function selectGroupOne ()
{
  selectGroupedItems(2)
}
function selectGroupTwo ()
{
  selectGroupedItems(1)
}
function selectGroupThree ()
{
  selectGroupedItems(0)
}

function selectGroupedItems(n)
{
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  var rowStart = [], rowEnd = [data.length + 1];

  for (var i = data.length - 1; i >= 6; i--)
  {
    if (data[i][7] === 'Variant SKU')
      rowStart.push(i + 1);
    else if (data[i][7] === '')
      rowEnd.push(i + 1)
  }
  rowStart.push(6)
  const numRows = rowEnd.map((row, index) => row - rowStart[index]);

  sheet.getRange(rowStart[n], 2, numRows[n], 8).activate();
}

/**
* Sorts data by the second column (ignores capitals).
*
* @return {Object[][]} The data sorted by it's second column.
*/
function sortBySecondColumn(a, b)
{
  return (a[1].toLowerCase() === b[1].toLowerCase()) ? 0 : (a[1].toLowerCase() < b[1].toLowerCase()) ? -1 : 1;
}

/**
* This function sets the ellapsed time of a function and prints it on the Test page.
*
* @param  {Number} startTime The start time that the script began running at represented by a number in milliseconds
* @return {Number} Returns the elapsed time or run time
* @author Jarren Ralf
*/
function elapsedTime(startTime)
{
  var timeNow = new Date().getTime(); // Get milliseconds from a date in past
  var elapsedTime = (timeNow - startTime)/1000;
  
  return elapsedTime;
}

/**
* This function creates a formatted date string for the current time and places the timestamp on the Adagio page.
*
* @return {String} The formatted data object
* @author Jarren Ralf
*/
function timeStamp()
{
  var   spreadsheet = SpreadsheetApp.getActive()
  var      timeZone = spreadsheet.getSpreadsheetTimeZone();
  var         today = new Date();
  var        format = "EEE, dd MMM yyyy HH:mm:ss";
  var formattedDate = Utilities.formatDate(today, timeZone, format);
  
  return formattedDate;
}

/**
* This is a function I found and modified to keep the first instance of an item in a muli-array based on the uniqueness of one of the values.
*
* @param      {Object[][]}    arr The given array
* @param  {Callback Function} key A function that chooses one of the elements of the object or array
* @return     {Object[][]}    The reduced array containing only unique items based on the key
*/
function uniqByKeepFirst(arr, key)
{
    let seen = new Set();
  
    return arr.filter(item => {
        let k = key(item);
        return seen.has(k) ? false : seen.add(k);
    });
}