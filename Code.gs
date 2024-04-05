/**
 * This function handles the on edit events in this spreadsheet. This function is looking for when a user changes information in the Customer name column,
 * or if a checkbox becomes check, signifying that order has been submitted.
 * 
 * @param {Event Object} e : The event object
 */
function installedOnEdit(e)
{
  const spreadsheet = e.source;
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const value = e.value;
  const sheetName = spreadsheet.getActiveSheet().getSheetName();

  try
  {
    if (sheetName === 'Dashboard' && row > 1)
      if (col === 2) // Changing or adding a customer's name
        updateCustomerName(range, value, spreadsheet);
      else if (col === 6) // Changing or adding a customer's email
        updateSharedStatusOfCustomerSS(range);
    else if (sheetName === 'Export' && col === 3) // The user may be editting the pricing
      updatePrice(e, range, value)
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function handles the on change events in this spreadsheet. Specifically, it is trying to identify when a submission checkbox for a customer is changed
 * so that an email can be sent to the relevant PNT employees.
 * 
 * @param {Event Object} e : The event object
 */
function onChange(e)
{
  try
  {
    if (e.changeType === 'OTHER')
    {
      const today = new Date().getTime();
      const dashboard = e.source.getSheetByName('Dashboard')
      
      const customerInfo = dashboard.getSheetValues(2, 2, dashboard.getLastRow() - 1, 4).map(date => {
        date[2] = (Number(date[2]) !== 0) ? Math.abs(Number(date[2]) - today) < 5000 : '';
        return date;
      });

      for (var i = 0; i < customerInfo.length; i++)
        if (customerInfo[i][2]) // Check if the change has been made in the last 5 seconds
          if (dashboard.getRange(i + 2, 3).isChecked())
            sendConfirmationEmail(customerInfo[i][0], customerInfo[i][3]);
          else
            sendCancelationEmail(customerInfo[i][0], customerInfo[i][3]);
    }
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function runs when the spreadsheet is opened or refreshed. It places a custom menu at the top of Browser which has quick access to running some important functions.
 */
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('PNT Menu')
      .addItem('Create Spreadsheets for Selected Customers', 'createSSforSelectedCustomers')
      .addItem('Email (and Share) Spreadsheets with Selected Customers', 'emailAndShareSpreadsheetsWithSelectedUsers')
      .addItem('Share Spreadsheets with Selected Customers', 'shareSpreadsheetsWithSelectedUsers')
    .addSeparator()
      .addItem('Create onChange Trigger (with pntnoreply)', 'createTrigger_OnChange_ByPntNoReplyGmail')
    .addSeparator()
      .addItem('Convert Selected Items to Wholesale Pricing', 'convertToWholeSalePricing')
      .addItem('Clear Export', 'clearExport')
      .addItem('Get Export Data', 'getExportData')
    .addToUi();
}

/**
 * This function clears the export sheet and then sends Adrian a courtesy email letting him know that the import template for Adagio OrderEntry has changed.
 * 
 * @author Jarren Ralf
 */
function clearExport()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const exportSheet = spreadsheet.getActiveSheet();

  try
  {
    if (exportSheet.getSheetName() !== 'Export')
    {
      spreadsheet.getSheetByName('Export').activate();
      Browser.msgBox('You must be on the Export sheet in order to clear it.')
    }
    else
    {
      exportSheet.clear();
      MailApp.sendEmail('adrian@pacificnetandtwine.com', 'The Template for Importing into OE has changed!', 'Remember to change the import template from LodgeImport to ShopifyImport next time you use it.') 
    }
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error);
  }
}

/**
 * This function takes the selection that the user has made on the export page and it converts those specific ranges to Wholesale pricing.
 * 
 * @author Jarren Ralf
 */
function convertToWholeSalePricing()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const exportSheet = spreadsheet.getActiveSheet();

  try
  {
    if (exportSheet.getSheetName() !== 'Export')
    {
      spreadsheet.getSheetByName('Export').activate();
      Browser.msgBox('Please select items on the Export sheet to convert to Wholesale prcing.');
    }
    else
    {
      var firstRows = [], lastRows = [], exportData = [];
      const ranges = exportSheet.getActiveRangeList().getRanges();

      ranges.map((rng, r) => {
        firstRows.push(rng.getRow());
        lastRows.push(rng.getLastRow());
        exportData.push(exportSheet.getSheetValues(firstRows[r], 1, lastRows[r] - firstRows[r] + 1, 4))
      })

      if (Math.max( ...lastRows) <= exportSheet.getLastRow()) // If the user has not selected an item, alert them with an error message
      { 
        const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages');
        const discounts = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 5);
        const BASE_PRICE = 1, WHOLESALE_PRICE = 4;

        const exportData_WithWholesalePrices = exportData.map(data => 
          data.map(item => {
            if (item[0] !== 'H')
            {
              itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]);

              if (itemPricing != undefined && itemPricing[BASE_PRICE] != 0 && itemPricing[WHOLESALE_PRICE] != 0)
                item[2] = (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2);
            }
            
            return item;
          })
        )

        ranges.map((range, r) => exportSheet.getRange(range.getRow(), 1, range.getNumRows(), 4).setValues(exportData_WithWholesalePrices[r]));
      }
      else
        SpreadsheetApp.getUi().alert('Please select an items from the list only.');
    }
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error);
  }
}

/**
 * This function creates new spreadsheets for the customers on the Dashboard that don't already have a spreadsheet in the fourth column.
 * 
 * @author Jarren Ralf
 */
function createSSforSelectedCustomers()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const dashboard = spreadsheet.getActiveSheet();

  try
  {
    if (dashboard.getSheetName() !== 'Dashboard')
    {
      spreadsheet.getSheetByName('Dashboard').activate();
      Browser.msgBox('Please select the customers that you wish to create spreadsheets for.')
    }
    else
    {
      var ss, url, velocityReportSheet, velocityReportSheetName, customerInvoiceData, invoiceSheet, splitDescription, sku, uom, colours = [], numRows, horizontalAligns, colourSelector = true;
      const templateSS = SpreadsheetApp.openById('1SN4H5_eEIYGvT2MrDIpusazpRePDVOdgI2hJlqEzULQ');
      const lodgeSalesSS = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0');
      const invoiceDataSheet = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('All Data');

      const invoiceData = invoiceDataSheet.getSheetValues(2, 1, invoiceDataSheet.getLastRow() - 1, 8).map(item => {
        item[4] = (item[4] === '100') ? 'Richmond' : (item[4] === '200') ? 'Parksville' : 'Prince Rupert'; // Convert 100, 200, and 300 location codes to the appropriate names for the customers
        splitDescription = item[0].split(' - ');
        sku = splitDescription.pop();
        uom = splitDescription.pop();
        splitDescription.pop();
        splitDescription.push(uom, sku);
        item[0] = splitDescription.join(' - ');

        return item;
      })

      invoiceData.shift() // Remove the header

      const customerListSheet = lodgeSalesSS.getSheetByName('Customer List');
      const customerList = customerListSheet.getSheetValues(3, 1, customerListSheet.getLastRow() - 2, 3);
      const numYears = new Date().getFullYear() - 2011;
      const CUST_NAME = 0, SALES_TOTAL = 2;
      const white = ['white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'], blue = ['#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8'];

      dashboard.getActiveRangeList().getRanges().map(rng => {
        rng.offset(0, 1 - rng.getColumn(), rng.getNumRows(), 5).getValues().map((customer, i) => {
          if (isNotBlank(customer[0]) && isNotBlank(customer[1]) && isBlank(customer[4])) // Both customer # and name are not blank, and the spreadsheet URL is blank
          {
            ss = templateSS.copy('PNT Order Sheet - ' + customer[1]); // Create the customers spreadsheet from the template spreadsheet
            ss.addEditor('pntnoreply@gmail.com'); // Add the pntnoreply email so that the emails can come from this account
            ss.getSheetByName('Item Search').getRange(1, 2).setValue(customer[1]).offset(3, 2).setValue(customer[0]); // Set the customer name and customer #
            velocityReportSheetName = customerList.find(custNum => custNum[0] === customer[0]); 
            lodgeSalesSS.getSheetByName(velocityReportSheetName[2]).copyTo(ss); // Take the "velocity report" from the Lodge Sales spreadsheet and put it on the customer's sheet
            velocityReportSheet = ss.getSheetByName('Copy of ' + velocityReportSheetName[2]).setName('Yearly Purchase Report');
            chartTitleInfo = velocityReportSheet.getRange(1, 2, 1, 3).getDisplayValues()[0];

            chart = velocityReportSheet.newChart()
              .asColumnChart()
              .addRange(velocityReportSheet.getRange(3, 5, numYears, 2))
              .setNumHeaders(0)
              .setXAxisTitle('Year')
              .setYAxisTitle('Sales Total')
              .setTransposeRowsAndColumns(false)
              .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
              .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
              .setOption('title', chartTitleInfo[CUST_NAME])
              .setOption('subtitle', 'Total: ' + chartTitleInfo[SALES_TOTAL])
              .setOption('isStacked', 'false')
              .setOption('bubble.stroke', '#000000')
              .setOption('textStyle.color', '#000000')
              .setOption('useFirstColumnAsDomain', true)
              .setOption('titleTextStyle.color', '#757575')
              .setOption('legend.textStyle.color', '#1a1a1a')
              .setOption('subtitleTextStyle.color', '#999999')
              .setOption('series', {0: {hasAnnotations: true, dataLabel: 'value'}})
              .setOption('trendlines', {0: {lineWidth: 4, type: 'linear', color: '#6aa84f'}})
              .setOption('hAxis', {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}})
              .setOption('annotations', {domain: {textStyle: {color: '#808080'}}, total: {textStyle : {color: '#808080'}}})
              .setOption('vAxes', {0: {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}, minorGridlines: {count: 2}}})
              .setPosition(1, 1, 0, 0)
              .build();

            velocityReportSheet.insertChart(chart);
            velocityReportSheet.protect();
            chartSheet = ss.moveChartToObjectSheet(chart).setName('Chart').setTabColor('#f1c232');
            colours.length = 0; // Clear the background colouyrs array

            customerInvoiceData = invoiceData.filter(name => name[1] === velocityReportSheetName[1]) // Customer invoice data
              .map((line, i, arr) => {

                if (i === 0)
                  colourSelector = true;
                else if (line[2].toString().trim() != arr[i - 1][2].toString().trim()) // If the current invoice number does not match the current one, then switch the background colours
                  colourSelector = !colourSelector;

                colours.push((colourSelector) ? white : blue);
                
                return line;
              })

            invoiceSheet = ss.insertSheet('Past Invoices', {template: ss.getSheetByName('Template')}).showSheet()
            numRows = customerInvoiceData.length;
            horizontalAligns = new Array(numRows).fill(['left', 'right', 'right', 'center', 'center', 'center', 'right', 'right']);

            invoiceSheet.getRange(2, 1, customerInvoiceData.length, 8).setNumberFormat('@').setBackgrounds(colours).setHorizontalAlignments(horizontalAligns).setValues(customerInvoiceData);
            invoiceSheet.deleteColumn(2).protect();
            url = ss.getUrl();
            customer[2] = "=IMPORTRANGE(INDIRECT(ADDRESS(ROW(), 5, 4)),\"Item Search!F2\")";
            customer[3] = "=IMPORTRANGE(INDIRECT(ADDRESS(ROW(), 5, 4)),\"Item Search!B4\")";
            customer[4] = url;

            rng.offset(i, 1 - rng.getColumn(), 1, 5).setValues([customer]);
            return customer;
          }
        })
      })
    }
  }
  catch (e)
  {
    var error = e['stack'];
    sendErrorEmail(error);
  }
}

/**
 * Creates triggers only if the jarrencralf account runs this function.
 * 
 * @author Jarren Ralf
 */
function createTriggers()
{
  if (Session.getActiveUser().getEmail() !== 'jarrencralf@gmail.com')
    Browser.msgBox('This function can only be run by the jarrencralf@gmail account.');
  else
  {
    ScriptApp.newTrigger('updateCustomerSpreadsheets').timeBased().atHour(23).everyDays(1).create();
    ScriptApp.newTrigger('removeUnapprovedEditorsFromCustomerSpreadsheet').timeBased().everyHours(1).create();
    ScriptApp.newTrigger('formatAllCustomerSpreadsheets').timeBased().everyDays(1).atHour(3).create();
    ScriptApp.newTrigger('installedOnEdit').forSpreadsheet('1MVL3lRDKrTa1peqBCjlS9GMAysdOD13_Sl0ygYb8VpE').onEdit().create();
  }
}

/**
 * Creates the onChange trigger only if the pntnoreply account runs this function.
 * 
 * @author Jarren Ralf
 */
function createTrigger_OnChange_ByPntNoReplyGmail()
{
  if (Session.getActiveUser().getEmail() !== 'pntnoreply@gmail.com')
    Browser.msgBox('This function can only be run by the pntnoreply@gmail account.');
  else
    ScriptApp.newTrigger('onChange').forSpreadsheet('1MVL3lRDKrTa1peqBCjlS9GMAysdOD13_Sl0ygYb8VpE').onChange().create();
}

/**
 * This function deletes all of the trigger for this spreadsheet
 * 
 * @author Jarren Ralf
 */
function deleteTriggers()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger))
}

/**
 * This function gets the selected cells from the user on the Dashboard and emails (and shares) the selected spreadsheets with the email addresses provided.
 * 
 * @author Jarren Ralf
 */
function emailAndShareSpreadsheetsWithSelectedUsers()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const dashboard = spreadsheet.getActiveSheet();

  try
  {
    if (dashboard.getSheetName() !== 'Dashboard')
    {
      spreadsheet.getSheetByName('Dashboard').activate();
      Browser.msgBox('Please return to the Dashboard to run this function.')
    }
    else
    {
      // dashboard.getActiveRangeList().getRanges().map(rng => {
      //   rng.offset(0, 5 - rng.getColumn(), rng.getNumRows(), 2).getValues().map((custSS, i) => {
      //       if (isNotBlank(custSS[0]) && isNotBlank(custSS[1]))
      //       {
      //         ss = SpreadsheetApp.openByUrl(custSS[0]);
      //         ss.addEditors(custSS[1].split(',').map(email => email.trim()));
      //         ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).map(protection => protection.removeEditors(protection.getEditors()))
      //         rng.offset(i, 7 - rng.getColumn(), 1, 1).check()
      //       }
      //     })
      // })
    }
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function formats the customers spreadsheets. A trigger runs this function daily. Due to the amount of possible customer spreadsheets, in order to avoid maxing out the runtime,
 * each day of the week it formats different customers. The result is each customer gets their spreadsheet formatted once a week.
 * 
 * @author Jarren Ralf
 */
function formatAllCustomerSpreadsheets()
{
  const dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  const dayOfWeek = new Date().getDay();
  var itemSearchSheet, maxRows, startTime;

  try
  {
    dashboard.getSheetValues(2, 5, dashboard.getLastRow() - 1, 1).map((custSS, i) => {

      startTime = new Date().getTime(); // Reset the function runtime
      
      if (i % 6 === dayOfWeek && isNotBlank(custSS[0])) // If spreadsheet URL is not blank, and use the day of the week (Sun - Sat => 0 - 6) in order to decide which spreadsheets to format
      {
        itemSearchSheet = SpreadsheetApp.openByUrl(custSS[0]).getSheetByName('Item Search');
        maxRows = itemSearchSheet.getMaxRows() - 4;

        itemSearchSheet.getRange(5, 1, maxRows, itemSearchSheet.getMaxColumns()).setBorder(false, false, false, false, false, false) // The full range below the header
            .setBackgrounds(new Array(maxRows).fill(['#cccccc', '#4a86e8', '#cccccc', '#cccccc', '#cccccc', 'white', '#cccccc', 'white', 'white']))
            .setFontColors(new Array(maxRows).fill(['#434343', '#4a86e8', '#434343', '#434343', '#434343', 'black', '#434343', 'black', 'black']))
            .setFontFamily('Arial')
            .setFontLine('none')
            .setFontSize(10)
            .setFontStyle('normal')
            .setFontWeight('bold')
            .setHorizontalAlignments(new Array(maxRows).fill(['left', 'center', 'center', 'center', 'center', 'center', 'center', 'left', 'left']))
            .setNumberFormat('@')
            .setVerticalAlignment('middle')
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .offset( 0, 1, maxRows, 1).setBorder(false, true, false, true, null, false, '#1155cc', SpreadsheetApp.BorderStyle.SOLID_THICK) // The vertical blue line below the header
          .offset(-3, 4, 1, 3).setBorder(false, false, false, false, false, null) // The checkbox, timestamp, and item information display
            .setBackground('#4a86e8')
            .setFontColors([['white', 'white', '#ffff00']])
            .setFontFamily('Arial')
            .setFontLine('none')
            .setFontSizes([[34, 11, 16]])
            .setFontStyle('normal')
            .setFontWeight('bold')
            .setHorizontalAlignment('center')
            .setNumberFormats([['#', '@', '@']])
            .setVerticalAlignment('middle')
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .offset(0, 3, 1, 1).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // The delivery instructions
            .setBackground('white')
            .setFontColor('black')
            .setFontFamily('Arial')
            .setFontLine('none')
            .setFontSize(11)
            .setFontStyle('normal')
            .setFontWeight('bold')
            .setHorizontalAlignment('left')
            .setNumberFormat('@')
            .setVerticalAlignment('middle')
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .offset(-1, -1, 1, 1).setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // The PO number
            .setBackground('white')
            .setFontColor('black')
            .setFontFamily('Arial')
            .setFontLine('none')
            .setFontSize(11)
            .setFontStyle('normal')
            .setFontWeight('bold')
            .setHorizontalAlignment('left')
            .setNumberFormat('@')
            .setVerticalAlignment('middle')
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)

        Logger.log(itemSearchSheet.getSheetValues(1, 2, 1, 1)[0][0] + '\'s spreadsheet has been successfully formatted in ' + (new Date().getTime() - startTime)/1000 + ' seconds.')
      }
    })
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function gets the export data from all of the customer's spreadsheets that have submitted their order.
 * 
 * @author Jarren Ralf
 */
function getExportData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const dashboard = spreadsheet.getActiveSheet();

  try
  {
    if (dashboard.getSheetName() !== 'Dashboard')
    {
      spreadsheet.getSheetByName('Dashboard').activate();
      Browser.msgBox('Please return to the Dashboard to run this function.')
    }
    else
    {
      var ss, itemSearchSheet, range, numRows, numItems, recentlyCreatedSheet, recentlyCreatedItems, deliveryInstructions, itemPricing, isQCL = false, exportData = [], exportData_WithDiscountedPrices = [];
      const numSS = dashboard.getLastRow() - 1;
      const dashboardRange = dashboard.getRange(2, 1, numSS, 5);

      dashboardRange.getValues().map((customer, c) => {
        if (dashboard.getRange(c + 2, 3).isChecked())
        {
          ss = SpreadsheetApp.openByUrl(customer[4]);
          itemSearchSheet = ss.getSheetByName('Item Search');
          recentlyCreatedSheet = ss.getSheetByName('Recently Created');
          numItems = recentlyCreatedSheet.getLastRow();
          recentlyCreatedItems = recentlyCreatedSheet.getSheetValues(1, 1, numItems, 1);
          numRows = Math.max(getLastRowSpecial(itemSearchSheet.getSheetValues(1, 8, itemSearchSheet.getMaxRows(), 1)), // Description column
                            getLastRowSpecial(itemSearchSheet.getSheetValues(1, 9, itemSearchSheet.getMaxRows(), 1))) // Item / General Order Notes column
                    - 3; 
          range = itemSearchSheet.getRange(4, 3, numRows, 7);
          deliveryInstructions = itemSearchSheet.getSheetValues(2, 9, 1, 1)[0][0];

          /* If there are delivery instructions, make them the final line of the order.
          * If necessary, make multiple comment lines if comments are > 75 characters long.
          */
          exportData.push(...range.getValues(), // The SKUs and quantities
            ['I', 'Provide your preferred delivery / pick up date and location below:', '', ''],
            ...(isNotBlank(deliveryInstructions)) ? deliveryInstructions.match(/.{1,75}/g).map(c => ['I', c, '', '']) : [['I', '**Customer left this field blank**', '', '']]);

          // range.offset(1, 0, numRows - 1).clearContent() // Clear the customers order, including notes
          //   .offset(-4, 5, 1, 1).setValue('')            // Remove the Customer PO #
          //   .offset( 1, 0, 1, 2).setValues([['Items displayed in order of newest to oldest.', '']]) // Remove the Delivery / Pick Up instructions
          //   .offset(0, -2).uncheck()                     // Uncheck the submit order checkbox
          //   .offset(-1, -5, 2, 1).setValue('')           // Remove the words from the search box
          //   .offset( 3,  1, 1, 1).setValue('')           // Remove the hidden timestamp
          //   .offset(1, -1, itemSearchSheet.getMaxRows() - 4, 1).clearContent() // Clear the full search range
          //   .offset(0, 0, numItems).setValues(recentlyCreatedItems); // Place the recently created items on the search page

          spreadsheet.toast(customer[1] + '\'s spreadsheet has been reset.')
        }
      })

      const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages')
      const discounts = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 5);
      const BASE_PRICE = 1, LODGE_PRICE = 3, WHOLESALE_PRICE = 4;

      exportData.map(item => {
        if (item[0] === 'H')
        {
          isQCL = (item[1] === 'DL1015') ? true : false;
          exportData_WithDiscountedPrices.push(['H', item[1], item[2], item[3]])
        }
        else if (item[0] === 'I')
          exportData_WithDiscountedPrices.push(['I', item[1], '', ''])
        else if (item[0] === 'D')
        {
          item[1] = item[1].toString().trim().toUpperCase(); // Make the SKU uppercase

          if (isNotBlank(item[1])) // SKU is not blank
          {
            if (isNotBlank(item[3])) // Order quantity is not blank
            {
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
              {
                itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

                if (itemPricing == undefined) // SKU is assumed to be invalid
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, item[3]], 
                    ...('SKU Not Found: ' + item[1] + ' - ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                  )
                else // SKU is assumed to be valid
                {
                  if (itemPricing[BASE_PRICE] != 0 && ((isQCL && itemPricing[WHOLESALE_PRICE] != 0) || itemPricing[LODGE_PRICE] != 0))
                    item[2] = (isQCL) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2) : (itemPricing[BASE_PRICE]*(100 - itemPricing[LODGE_PRICE])/100).toFixed(2); // Set the pricing

                  exportData_WithDiscountedPrices.push(['D', item[1], item[2], item[3]])
                }
              }
              else // Order quantity is not a valid number
              {
                itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

                if (itemPricing == undefined) // SKU is assumed to be invalid
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, 0], 
                    ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', ''], 
                    ...('SKU Not Found: ' + item[1] + ' - ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                  )
                else // SKU is assumed to be valid
                {
                  if (itemPricing[BASE_PRICE] != 0 && ((isQCL && itemPricing[WHOLESALE_PRICE] != 0) || itemPricing[LODGE_PRICE] != 0))
                    item[2] = (isQCL) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2) : (itemPricing[BASE_PRICE]*(100 - itemPricing[LODGE_PRICE])/100).toFixed(2); // Set the pricing

                  exportData_WithDiscountedPrices.push(
                    ['D', item[1], item[2], 0], 
                    ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', '']
                  )
                }
              }
            }
            else // The order quantity is blank (while SKU is not)
            {
              itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

              if (itemPricing == undefined) // SKU is assumed to be invalid
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0], 
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', ''],
                  ...('SKU Not Found: ' + item[1] + ' - ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              else // SKU is assumed to be valid
              {
                if (itemPricing[BASE_PRICE] != 0 && ((isQCL && itemPricing[WHOLESALE_PRICE] != 0) || itemPricing[LODGE_PRICE] != 0))
                  item[2] = (isQCL) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2) : (itemPricing[BASE_PRICE]*(100 - itemPricing[LODGE_PRICE])/100).toFixed(2); // Set the pricing

                exportData_WithDiscountedPrices.push(
                  ['D', item[1], item[2], 0],
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '']
                )
              }
            }
          }
          else // The SKU is blank
          {
            if (isNotBlank(item[3])) // Order quantity is not blank
            {
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
              {
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, item[3]], 
                  ...('Description: ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              }
              else // Order quantity is not a valid number
              {
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0], 
                  ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', ''], 
                  ...('Description: ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              }
            }
            else // The order quantity is blank 
            {
              if (isNotBlank(item[5])) // Description is not blank (but SKU and quantity are)
              {
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0], 
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', ''],
                  ...('Description: ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              }
            }
          }

          if (isNotBlank(item[6])) // There are notes for the current line
            exportData_WithDiscountedPrices.push(...('Notes: ' + item[6]).match(/.{1,75}/g).map(c => ['C', c, '', '']))
        }
        else // There was no line indicator
        {
          item[1] = item[1].toString().trim().toUpperCase(); // Make the SKU uppercase

          if (isNotBlank(item[1])) // SKU is not blank
          {
            if (isNotBlank(item[3])) // Order quantity is not blank
            {
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
              {
                itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

                if (itemPricing == undefined) // SKU is assumed to be invalid
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, item[3]], 
                    ...('SKU Not Found: ' + item[1] + ' - ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                  )
                else // SKU is assumed to be valid
                {
                  if (itemPricing[BASE_PRICE] != 0 && ((isQCL && itemPricing[WHOLESALE_PRICE] != 0) || itemPricing[LODGE_PRICE] != 0))
                    item[2] = (isQCL) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2) : (itemPricing[BASE_PRICE]*(100 - itemPricing[LODGE_PRICE])/100).toFixed(2); // Set the pricing

                  exportData_WithDiscountedPrices.push(['D', item[1], item[2], item[3]])
                }
              }
              else // Order quantity is not a valid number
              {
                itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

                if (itemPricing == undefined) // SKU is assumed to be invalid
                  exportData_WithDiscountedPrices.push(
                    ['D', 'MISCITEM', 0, 0], 
                    ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', ''], 
                    ...('SKU Not Found: ' + item[1] + ' - ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                  )
                else // SKU is assumed to be valid
                {
                  if (itemPricing[BASE_PRICE] != 0 && ((isQCL && itemPricing[WHOLESALE_PRICE] != 0) || itemPricing[LODGE_PRICE] != 0))
                    item[2] = (isQCL) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2) : (itemPricing[BASE_PRICE]*(100 - itemPricing[LODGE_PRICE])/100).toFixed(2); // Set the pricing

                  exportData_WithDiscountedPrices.push(
                    ['D', item[1], item[2], 0], 
                    ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', '']
                  )
                }
              }
            }
            else // The order quantity is blank (while SKU is not)
            {
              itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

              if (itemPricing == undefined) // SKU is assumed to be invalid
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0], 
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', ''],
                  ...('SKU Not Found: ' + item[1] + ' - ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              else // SKU is assumed to be valid
              {
                if (itemPricing[BASE_PRICE] != 0 && ((isQCL && itemPricing[WHOLESALE_PRICE] != 0) || itemPricing[LODGE_PRICE] != 0))
                  item[2] = (isQCL) ? (itemPricing[BASE_PRICE]*(100 - itemPricing[WHOLESALE_PRICE])/100).toFixed(2) : (itemPricing[BASE_PRICE]*(100 - itemPricing[LODGE_PRICE])/100).toFixed(2); // Set the pricing

                exportData_WithDiscountedPrices.push(
                  ['D', item[1], item[2], 0],
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', '']
                )
              }
            }
          }
          else // The SKU is blank
          {
            if (isNotBlank(item[3])) // Order quantity is not blank
            {
              if (Number(item[3]).toString() !== 'NaN') // Order number is a valid number
              {
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, item[3]], 
                  ...('Description: ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              }
              else // Order quantity is not a valid number
              {
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0], 
                  ['C', 'Invalid order QTY: "' + item[3] + '" for above item, therefore it was replaced with 0', '', ''], 
                  ...('Description: ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              }
            }
            else // The order quantity is blank 
            {
              if (isNotBlank(item[5])) // Description is not blank (but SKU and quantity are)
              {
                exportData_WithDiscountedPrices.push(
                  ['D', 'MISCITEM', 0, 0], 
                  ['C', 'Order quantity was blank for the above item, therefore it was replaced with 0', '', ''],
                  ...('Description: ' + item[5] + ' - ' + item[4]).toString().match(/.{1,75}/g).map(c => ['C', c, '', ''])
                )
              }
            }
          }

          if (isNotBlank(item[6])) // There are notes for the current line
            exportData_WithDiscountedPrices.push(...('Notes: ' + item[6]).match(/.{1,75}/g).map(c => ['C', c, '', '']))
        }
      })

      const exportSheet = spreadsheet.getSheetByName('Export').clear();
      const ranges = [[],[],[]];
      const backgroundColours = [
        '#c9daf8', // Make the header rows blue
        '#fcefe1', // Make the comment rows orange
        '#e0d5fd'  // Make the instruction comment rows purple
      ];

      exportData_WithDiscountedPrices.map((h, r) => 
        h = (h[0] !== 'H') ? (h[0] !== 'C') ? (h[0] !== 'I') ? false : 
        ranges[2].push('A' + (r + 1) + ':D' + (r + 1)) : // Instruction comment rows purple
        ranges[1].push('A' + (r + 1) + ':D' + (r + 1)) : // Comment rows orange
        ranges[0].push('A' + (r + 1) + ':D' + (r + 1))   // Header rows blue
      )

      ranges.map((rngs, r) => exportSheet.getRangeList(rngs).setBackground(backgroundColours[r])); // Set the appropriate background colours
      exportSheet.getRange(1, 1, exportData_WithDiscountedPrices.length, 4).setNumberFormat('@').setValues(exportData_WithDiscountedPrices).activate();
    }
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * Gets the last row number based on a selected column range values
 *
 * @param {array} range : takes a 2d array of a single column's values
 * @returns {number} : the last row number with a value. 
 */ 
function getLastRowSpecial(range)
{
  for (var row = 0, rowNum = 0, blank = false; row < range.length; row++)
  {
    if (isBlank(range[row][0]) && !blank)
    {
      rowNum = row;
      blank = true;
    }
    else if(isNotBlank(range[row][0]))
      blank = false;
  }
  return rowNum;
}

/**
 * This function checks if the given string is blank.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is blank.
 * @author Jarren Ralf
 */
function isBlank(str)
{
  return str === '';
}

/**
 * This function checks if the given string is not blank.
 * 
 * @param {String} str : The given string.
 * @returns {Boolean} Whether the given string is not blank.
 * @author Jarren Ralf
 */
function isNotBlank(str)
{
  return str !== '';
}

/**
 * This function runs on a trigger every X and it removes any editors from the each customer spreadsheet that are not contained in the corresponding Customer Email(s) column on the Dashboard.
 * 
 * @author Jarren Ralf
 */
function removeUnapprovedEditorsFromCustomerSpreadsheet()
{
  const dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var approvedEditors, email;

  try
  {
    dashboard.getSheetValues(2, 5, dashboard.getLastRow() - 1, 2).map((custSS, i) => {
      if (isNotBlank(custSS[0])) // Spreadsheet URL is not blank
      {
        ss = SpreadsheetApp.openByUrl(custSS[0]).addEditor('pntnoreply@gmail.com'); // Make sure pntnoreply@gmail.com is an editor
        approvedEditors = ['jarrencralf@gmail.com', 'pntnoreply@gmail.com'];
        approvedEditors.push(...custSS[1].split(',').map(email => email.trim())); // Get the list of approved editors and add it to jarrencralf@gmail and pntnoreply@gmail
        
        currentEditors = ss.getEditors().map(user => {
          email = user.getEmail();

          if (!approvedEditors.includes(email)) // If an editor is not on the approved email list, then remove them
            ss.removeEditor(email);
        });

        ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).map(protection => protection.removeEditors(protection.getEditors()));
        dashboard.getRange(2 + i, 7).check(); // Check the box that signals if the spreadsheet is appropriately shared with the relevant emails
      }
    })
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function sends an email to all of the relevant PNT Lodge employees with the items and notes that the customer ordered.
 * 
 * @author Jarren Ralf
 */
function sendCancelationEmail(name, ssUrl)
{
  const spreadsheet = SpreadsheetApp.openByUrl(ssUrl);
  const itemSearchSheet = spreadsheet.getSheetByName('Item Search');
  const poNum = itemSearchSheet.getSheetValues(1, 8, 1, 1)[0][0];
  const isPoNotBlank = isNotBlank(poNum);
  const customerEmails = spreadsheet.getEditors().map(editor => editor.getEmail()).filter(email => email !== 'jarrencralf@gmail.com' && email !== 'pntnoreply@gmail.com').join(', ');

  // Send an email to the PNT employees with the new order
  MailApp.sendEmail({
    to: "lb_blitz_allstar@hotmail.com", // "deryk@pacificnetandtwine.com, scottnakashima@hotmail.com, eryn@pacificnetandtwine.com, triteswarehouse@pacificnetandtwine.com"
    replyTo: customerEmails,
    name: 'PNT Sales',
    subject: (isPoNotBlank) ? name + " has cancelled their order, PO # " + poNum : name + " has cancelled their order",
    htmlBody: "<p>Reply to this email if you want to contact the customer.</p>"
  });

  // Send an email confirmation to the customer
  MailApp.sendEmail({
    to: customerEmails, // Send a confirmation to all of the editors, except me
    replyTo: "deryk@pacificnetandtwine.com, scottnakashima@hotmail.com, eryn@pacificnetandtwine.com, triteswarehouse@pacificnetandtwine.com",
    name: 'PNT Sales',
    subject: (isPoNotBlank) ? "Order Cancellation for PO # " + poNum : "Order Cancellation",
    htmlBody: "<p>Your order has been successfully cancelled.</p><p><br></p><p>Reply to this email if you would like to contact the Lodge Sales team at Pacific Net & Twine.</p><p>Thank you.</p>"
  });
}

/**
 * This function sends an email to all of the relevant PNT Lodge employees with the items and notes that the customer ordered.
 * 
 * @author Jarren Ralf
 */
function sendConfirmationEmail(name, ssUrl)
{
  const spreadsheet = SpreadsheetApp.openByUrl(ssUrl);
  const itemSearchSheet = spreadsheet.getSheetByName('Item Search');
  const poNum = itemSearchSheet.getSheetValues(1, 8, 1, 1)[0][0];
  const isPoNotBlank = isNotBlank(poNum);
  const dateAndLocationForDelivery = itemSearchSheet.getSheetValues(2, 9, 1, 1)[0][0];
  const isDateAndLocationNotBlank = isNotBlank(dateAndLocationForDelivery);
  const numRows = Math.max(getLastRowSpecial(itemSearchSheet.getSheetValues(1, 8, itemSearchSheet.getMaxRows(), 1)), // Description column
                           getLastRowSpecial(itemSearchSheet.getSheetValues(1, 9, itemSearchSheet.getMaxRows(), 1))) // Item / General Order Notes column
                  - 4;
  var values = itemSearchSheet.getSheetValues(5, 4, numRows, 6)
  values.map(arr => {arr.splice(1, 1)}) // Remove the pricing (it's all $0.00 anyways)
  const numCols = values[0].length;
  const customerEmails = spreadsheet.getEditors().map(editor => editor.getEmail()).filter(email => email !== 'jarrencralf@gmail.com' && email !== 'pntnoreply@gmail.com').join(', ');

  var body = "<table><tr><th colspan=\"" 
    + numCols + "\">Provide your preferred delivery / pick up date and location below:</th></tr><tr><th colspan=\"" 
    + numCols + "\">" 
    + ((isDateAndLocationNotBlank) ? dateAndLocationForDelivery : "**Customer left this field blank**") 
    + "</th></tr><tr><th colspan=\"" 
    + numCols + "\"><br></th></tr><tr><th>Item Number</th><th>Qty</th><th>UoM</th><th>Description</th><th>Item / General Order Notes</th></tr>";

  // Build the html table, which will be the body of the email, from the multi-array of skus and descriptions
  for (var r = 0; r < numRows; r++)
  {
    body += "<tr>";

    for (var c = 0; c < numCols; c++)
      body += "<td>" + values[r][c] + "</td>";

    body += "</tr>";
  }

  body += "<tr><th colspan=\"" 
    + numCols + "\"><br></th></tr><tr><th colspan=\"" 
    + numCols + "\">If you have any questions or concerns about this order, then replying to this email will send a message directly to the customer.</th></tr></table>";

  // Send an email to the PNT employees with the new order
  MailApp.sendEmail({
    to: "lb_blitz_allstar@hotmail.com", // "deryk@pacificnetandtwine.com, scottnakashima@hotmail.com, eryn@pacificnetandtwine.com, triteswarehouse@pacificnetandtwine.com"
    replyTo: customerEmails,
    name: 'PNT Sales',
    subject: (isPoNotBlank) ? name + " has placed an order, PO # " + poNum : name + " has placed an order!",
    htmlBody: body
  });

  body = body.slice(0, -202); // Remove the closing remarks
  body += "<tr><th colspan=\"" 
    + numCols + "\"><br></th></tr><tr><th colspan=\"" 
    + numCols + "\">If you have any additional comments, questions, or problems with your order, reply to this email and one of our team members will get back to you as soon as they can.</th></tr></table>";

  // Send an email confirmation to the customer
  MailApp.sendEmail({
    to: customerEmails, // Send a confirmation to all of the editors, except me
    replyTo: "deryk@pacificnetandtwine.com, scottnakashima@hotmail.com, eryn@pacificnetandtwine.com, triteswarehouse@pacificnetandtwine.com",
    name: 'PNT Sales',
    subject: (isPoNotBlank) ? "Order Confirmation for PO # " + poNum : "Order Confirmation",
    htmlBody: body
  });
}

/**
 * This function sends an email to Jarren to give a heads up that a function in apps script has failed to run.
 * 
 * @param {String} error : The property of the error object that displays the functions and line numbers that the error occurs at.
 * @author Jarren Ralf
 */
function sendErrorEmail(error)
{
  if (MailApp.getRemainingDailyQuota() > 3) // Don't try and send an email if the daily quota of emails has been sent
  {
    var today = new Date()
    var formattedError = '<p>' + error.replaceAll(' at ', '<br /> &emsp;&emsp;&emsp;') + '</p>';
    var templateHtml = HtmlService.createTemplateFromFile('FunctionFailedToRun');
    templateHtml.dateAndTime = today.toLocaleTimeString() + ' on ' + today.toDateString();
    templateHtml.scriptURL   = "https://script.google.com/home/projects/178jXC1SLz1GQpIOiNLgRAzE4j4A-F1jt4OatEQ3BLLwaO3nH4rZrRDRm/edit";
    var emailBody = templateHtml.evaluate().append(formattedError).getContent();
    
    MailApp.sendEmail({      to: 'lb_blitz_allstar@hotmail.com',
                        subject: 'Lodge Order Processor Script Failure', 
                       htmlBody: emailBody
    });
  }
  else
    Logger.log('No email sent because it appears that the daily quota of emails has been met!')
}

/**
 * This function gets the selected cells from the user on the Dashboard and shares the selected spreadsheets with the email addresses provided.
 * 
 * @author Jarren Ralf
 */
function shareSpreadsheetsWithSelectedUsers()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const dashboard = spreadsheet.getActiveSheet();

  try
  {
    if (dashboard.getSheetName() !== 'Dashboard')
    {
      spreadsheet.getSheetByName('Dashboard').activate();
      Browser.msgBox('Please return to the Dashboard to run this function.')
    }
    else
    {
      dashboard.getActiveRangeList().getRanges().map(rng => {
        rng.offset(0, 5 - rng.getColumn(), rng.getNumRows(), 2).getValues().map((custSS, i) => {
            if (isNotBlank(custSS[0]) && isNotBlank(custSS[1]))
            {
              ss = SpreadsheetApp.openByUrl(custSS[0]);
              ss.addEditors(custSS[1].split(',').map(email => email.trim()));
              ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).map(protection => protection.removeEditors(protection.getEditors()))
              rng.offset(i, 7 - rng.getColumn(), 1, 1).check()
            }
          })
      })
    }
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
* Sorts data by the created date of the product for the richmond spreadsheet.
*
* @param  {String[]} a : The current array value to compare
* @param  {String[]} b : The next array value to compare
* @return {String[][]} The output data.
* @author Jarren Ralf
*/
function sortByCreatedDate(a, b)
{
  return (a[1] === b[1]) ? 0 : (a[1] < b[1]) ? 1 : -1;
}

/**
 * This function updates the customer name when the second column of the Dashboard is editted that contains the customer's name.
 * 
 * @param   {Range}       range     : The range that the edit just occured at.
 * @param   {String}      value     : The value of the edit that just occured.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet.
 * @author Jarren Ralf
 */
function updateCustomerName(range, value, spreadsheet)
{
  if (value !== undefined) // The value in the customer name column is not blank
  {
    const listRange = range.getDataValidation().copy().getCriteriaValues()[0]; // The range of the data validation that contains the list of all Lodge customers
    const customerList = listRange.offset(0, -1, listRange.getNumRows(), 2).getValues(); // The name and customer number of all lodge customers
    const idx = customerList.findIndex(name => name[1] === value); // The index position of the lodge customer that matches the users input

    if (idx !== -1) // Lodge customer was found in list
    {
      range.offset(0, -1, 1, 2).setValues([customerList[idx]]); // Add the customer number to the left cell
      spreadsheet.toast('Customer was found.');
    }
    else if (isNotBlank(range.offset(0, -1).getValue())) // A customer name does not match one in the list, but if there is an account number to the left, then change the customers name to the new input
    {
      customerList[customerList.findIndex(custNum => custNum[0] === range.offset(0, -1).getValue())][1] = value;
      listRange.setValues(customerList.map(col => [col[1]]))
      spreadsheet.toast('Customer name was updated in list.')
    }
    else
      range.setValue('') // There was no customer account number to the left, so assume that the user made a mistake in their input
  }
  else // The value in the customer name column is blank
    range.offset(0, -1, 1, 7).setValues([['', '', false, '', '', '', false]]); // Remove the customer number if there is one
}

/**
 * This function updates the customers spreadsheets. A trigger runs this function daily. The list of items are updated, as well as the customer's
 * velocity report and invoice data.
 * 
 * @author Jarren Ralf
 */
function updateCustomerSpreadsheets()
{
  var splitDescription, newDescription, ss, d, velocityReportSheet, velocityReportSheetName, invoiceSheet, customerInvoiceData, itemList = [];

  try
  {
    const lodgeSalesSS = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0');
    const invoiceDataSheet = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('All Data');
    const invoiceData = invoiceDataSheet.getSheetValues(2, 1, invoiceDataSheet.getLastRow() - 1, 8).map(item => {
      item[4] = (item[4] === '100') ? 'Richmond' : (item[4] === '200') ? 'Parksville' : 'Prince Rupert';
      splitDescription = item[0].split(' - ');
      sku = splitDescription.pop();
      uom = splitDescription.pop();
      splitDescription.pop();
      splitDescription.push(uom, sku);
      item[0] = splitDescription.join(' - ');

      return item;
    })
    
    const dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
    const customerListSheet = lodgeSalesSS.getSheetByName('Customer List');
    const customerList = customerListSheet.getSheetValues(3, 1, customerListSheet.getLastRow() - 2, 3);
    const numYears = new Date().getFullYear() - 2011;
    const CUST_NAME = 0, SALES_TOTAL = 2;

    invoiceData.shift() // Remove the header

    const sortedItems = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString()).map(item => {
      splitDescription = item[1].split(' - ');
      splitDescription.splice(-4, 1);
      newDescription = splitDescription.join(' - ');
      itemList.push([newDescription]);

      d = item[6].split('.');                           // Split the date at the "."
      item[6] = new Date(d[2],d[1] - 1,d[0]).getTime(); // Convert the date sting to a striong object for sorting purposes
        
      return [newDescription, item[6]];
    }).sort(sortByCreatedDate).sort(sortByCreatedDate).map(descrip => [descrip[0]])

    // Remove the headers
    itemList.shift();
    sortedItems.shift();

    dashboard.getRange(2, 1, dashboard.getLastRow() - 1, 5).getValues()
      .map(customer => {
        if (isNotBlank(customer[4]))
        {
          ss = SpreadsheetApp.openByUrl(customer[4])
          ss.getSheetByName('Item List').getRange(1, 1, itemList.length).setValues(itemList);
          ss.getSheetByName('Recently Created').getRange(1, 1, itemList.length).setValues(sortedItems);
          ss.deleteSheet(ss.getSheetByName('Yearly Purchase Report'));
          ss.deleteSheet(ss.getSheetByName('Chart'));
          ss.deleteSheet(ss.getSheetByName('Past Invoices'));

          velocityReportSheetName = customerList.find(custNum => custNum[0] === customer[0]);
          lodgeSalesSS.getSheetByName(velocityReportSheetName[2]).copyTo(ss);
          velocityReportSheet = ss.getSheetByName('Copy of ' + velocityReportSheetName[2]).setName('Yearly Purchase Report');
          chartTitleInfo = velocityReportSheet.getRange(1, 2, 1, 3).getDisplayValues()[0];

          chart = velocityReportSheet.newChart()
            .asColumnChart()
            .addRange(velocityReportSheet.getRange(3, 5, numYears, 2))
            .setNumHeaders(0)
            .setXAxisTitle('Year')
            .setYAxisTitle('Sales Total')
            .setTransposeRowsAndColumns(false)
            .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
            .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
            .setOption('title', chartTitleInfo[CUST_NAME])
            .setOption('subtitle', 'Total: ' + chartTitleInfo[SALES_TOTAL])
            .setOption('isStacked', 'false')
            .setOption('bubble.stroke', '#000000')
            .setOption('textStyle.color', '#000000')
            .setOption('useFirstColumnAsDomain', true)
            .setOption('titleTextStyle.color', '#757575')
            .setOption('legend.textStyle.color', '#1a1a1a')
            .setOption('subtitleTextStyle.color', '#999999')
            .setOption('series', {0: {hasAnnotations: true, dataLabel: 'value'}})
            .setOption('trendlines', {0: {lineWidth: 4, type: 'linear', color: '#6aa84f'}})
            .setOption('hAxis', {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}})
            .setOption('annotations', {domain: {textStyle: {color: '#808080'}}, total: {textStyle : {color: '#808080'}}})
            .setOption('vAxes', {0: {textStyle: {color: '#000000'}, titleTextStyle: {color: '#000000'}, minorGridlines: {count: 2}}})
            .setPosition(1, 1, 0, 0)
            .build();

          velocityReportSheet.insertChart(chart);
          ss.moveChartToObjectSheet(chart).setName('Chart').setTabColor('#f1c232').getSheetId();

          customerInvoiceData = invoiceData.filter(name => name[1] === velocityReportSheetName[1]);
            
          invoiceSheet = ss.insertSheet('Past Invoices', {template: ss.getSheetByName('Template')})
          invoiceSheet.getRange(2, 1, customerInvoiceData.length, 8).setValues(customerInvoiceData);
          invoiceSheet.deleteColumn(2);
        }
    })
  }
  catch (e)
  {
    var error = e['stack'];
    throw new Error(error);
  }
}

/**
 * This function checks for a price change on the export sheet. The function notices whether a user simple changes the price by inputting a number that contains a decimal, OR 
 * if the number is a whole number (with no decimal ** not even 15.00) and therefore a percentage discount change.
 * 
 * @param {Event Object} e   : The event object
 * @param     {Range}  range : The range that was editted
 * @param    {String}  value : The value in the cell that was inputted
 */
function updatePrice(e, range, value)
{
  if (value == undefined) // The user has pressed delete
  {
    range.setValue(e.oldValue)
    SpreadsheetApp.flush();
    Browser.msgBox('Invalid price or discount.')
  }
  else if (isNaN(Number(value))) // The inputted value is not a valid number
  {
    range.setValue(e.oldValue)
    SpreadsheetApp.flush();
    Browser.msgBox('Invalid price or discount.')
  }
  else if (!value.toString().includes('.')) // The inputted value is a valid number without a decimal, therefore assumed to be a discount percentage
  {
    const BASE_PRICE = 1;
    const skuAndPriceRange = range.offset(0, -1, 1, 2);
    const skuAndPrice = skuAndPriceRange.getValues()[0];
    skuAndPrice[0] = skuAndPrice[0].toString().toUpperCase(); // SKU
    const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages')
    const itemPricing = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 2).find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === skuAndPrice[0]);

    if (itemPricing != undefined) // The sku was found on the discounts page
    {
      if (itemPricing[BASE_PRICE] != 0) // The base price is NOT zero on the discounts page
        skuAndPriceRange.setValues([[skuAndPrice[0], (itemPricing[BASE_PRICE]*(100 - Number(skuAndPrice[1]))/100).toFixed(2)]]); // Change the price with the desired discount
      else // The base price is zero on the discounts page
      {
        range.setValue(e.oldValue)
        SpreadsheetApp.flush();
        Browser.msgBox('Base price was $0.00 on the discounts spreadsheet.')
      }
    }
    else // The sku was NOT found on the discounts page
    {
      range.setValue(e.oldValue)
      SpreadsheetApp.flush();
      Browser.msgBox('SKU was not found on the discounts spreadsheet.')
    }
  }
}

/**
 * This function updates the shared status of the customer's spreadsheet when the sixth column of the Dashboard is editted that contains the customer's email(s).
 * 
 * @param  {Range} range : The range that the edit just occured at.
 * @author Jarren Ralf
 */
function updateSharedStatusOfCustomerSS(range)
{
  range.offset(0, 1, 1, 1).uncheck();
}