/**
 * This function handles the on edit events in this spreadsheet. This function is looking for when a user changes information in the Customer name column,
 * or if a checkbox becomes check, signifying that order has been submitted.
 * 
 * @param {Event Object} e : The event object
 * @throws General error if anything goes wrong
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
 * @throws General error if anything goes wrong
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
        {
          dashboard.getRange(i + 2, 3).check(); // Check the checkbox to indicate that the user has an outstanding (unimported) order
          sendConfirmationEmail(customerInfo[i][0], customerInfo[i][3]);
        }
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
      var ss, url, velocityReportSheet, velocityReportSheetName, customerInvoiceData, invoiceSheet, templateSheet, ordersSheet,
      chart, chartTitleInfo, splitDescription, colours = [], numRows, horizontalAligns, colourSelector = true;
      const templateSS = SpreadsheetApp.openById('1hhKeKpoheS71KuVCb9k1yPLq8PV41tIFsIj8U8VZUa4');
      const lodgeSalesSS = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0');
      const invoiceDataSheet = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('All Data');

      const invoiceData = invoiceDataSheet.getSheetValues(2, 1, invoiceDataSheet.getLastRow() - 1, 8).map(item => {
        item[4] = (item[4] === '100') ? 'Richmond' : (item[4] === '200') ? 'Parksville' : 'Prince Rupert'; // Convert 100, 200, and 300 location codes to the appropriate names for the customers
        splitDescription = item[0].split(' - ');
        splitDescription.splice(-4, 1);
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
            dashboard.showColumns(4); // Show the column to make sure that the dashboard has authorization to use importrange on that spreadsheet 
            ss = templateSS.copy('PNT Order Sheet - ' + customer[1]); // Create the customers spreadsheet from the template spreadsheet
            ss.addEditor('pntnoreply@gmail.com'); // Add the pntnoreply email so that the emails can come from this account
            ss.getSheetByName('Item Search').getRange(1, 2).setValue(customer[1]).offset(3, 1).setValue(customer[0]).offset(0, -1).setValue(''); // Set the customer name and customer # and remove timestamp (if there is one)
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
            ss.moveChartToObjectSheet(chart).setName('Chart').setTabColor('#f1c232');
            colours.length = 0; // Clear the background colours array

            customerInvoiceData = invoiceData.filter(name => name[1] === velocityReportSheetName[1]) // Customer invoice data
              .map((line, i, arr) => {

                if (i === 0)
                  colourSelector = true;
                else if (line[2].toString().trim() != arr[i - 1][2].toString().trim()) // If the current invoice number does not match the current one, then switch the background colours
                  colourSelector = !colourSelector;

                colours.push((colourSelector) ? white : blue);
                
                return line;
            })

            numRows = customerInvoiceData.length;
            templateSheet = ss.getSheetByName('Template');
            invoiceSheet = ss.insertSheet('Past Invoices', {template: templateSheet}).showSheet();
            ss.getSheetByName('Export').clear(); // Just incase there is any data left on this page from the template SS
            ordersSheet = ss.getSheetByName('Submitted Orders');
            ordersSheet.getRange(2, 1, ordersSheet.getMaxRows() - 1, ordersSheet.getLastColumn()).clear();
            //spreadsheet.deleteSheet(ss.getSheetByName('Template'));
            horizontalAligns = new Array(numRows).fill(['left', 'right', 'right', 'center', 'center', 'center', 'right', 'right']);

            if (numRows > 0)
              invoiceSheet.getRange(2, 1, numRows, 8).setNumberFormat('@').setBackgrounds(colours).setHorizontalAlignments(horizontalAligns).setValues(customerInvoiceData);

            invoiceSheet.deleteColumn(2).protect();
            
            url = ss.getUrl();
            customer[3] = "=IMPORTRANGE(INDIRECT(ADDRESS(ROW(), 5, 4)),\"Item Search!B4\")"; // Completed Date
            customer[4] = url;

            rng.offset(i, 1 - rng.getColumn(), 1, 5).setValues([customer]).offset(0, 6, 1, 2).uncheck();
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
    ScriptApp.newTrigger('updateOrderSheet_TEMPLATE').timeBased().atHour(23).everyDays(1).create();
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
 * It also sends an email to each address listed with an set of instructions for how to use the spreadsheet.
 * 
 * @throws General error if anything goes wrong
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
      const files = DriveApp.getFolderById('1oHuZbunXp4RcvKTi7IOVDy9-bfxcwf-y').getFiles(); // The inline gif images for the instructional email
      var htmlTemplate = HtmlService.createTemplateFromFile("Instructional Email"); // The email template
      var ss, emails, file, emailImages = {1: null, 2: null, 3: null, 4: null, 5: null, 6: null, 7: null, 8: null, 9: null, 10: null, 11: null, 12: null};
      
      while (files.hasNext()) // Loop through the gifs
      {
        file = files.next();
        emailImages[file.getName().split('_', 1)[0]] = file.getAs("image/gif"); // Sort them correctly into their place in the object
      }
      
      dashboard.getActiveRangeList().getRanges().map(rng => {
        rng.offset(0, 2 - rng.getColumn(), rng.getNumRows(), 5).getValues().map((custSS, i) => {
            if (isNotBlank(custSS[3]) && isNotBlank(custSS[4])) // The URL and emails are not blank
            {
              ss = SpreadsheetApp.openByUrl(custSS[3]);
              emails = custSS[4].split(',').map(email => email.trim());
              Logger.log('Sharing spreadsheet with ' + custSS[0] + ' employee emails:')

              for (var email = 0; email < emails.length; email++) // Loop through all of the employee emails and share with them individually to avoid errors
              {
                try
                {
                  ss.addEditor(emails[email]);
                  Logger.log('Successfully shared with: ' + emails[email])
                }
                catch (e)
                {
                  var error = e['stack'];
                  Logger.log('*Unsuccessful*')
                  Logger.log(error);
                }
              }

              protectSpreadsheet(ss);

              htmlTemplate.lodgeName = custSS[0];
              htmlTemplate.pntOrderFormURL = custSS[3];
              htmlTemplate.invoiceDataURL = custSS[3] + '#gid=' + ss.getSheetByName('Past Invoices').getSheetId();
              htmlTemplate.velocityReportURL = custSS[3] + '#gid=' + ss.getSheetByName('Yearly Purchase Report').getSheetId();
              htmlTemplate.createGoogleAccountURL = "https://support.google.com/accounts/answer/27441?hl=en";

              MailApp.sendEmail({
                to: emails.join(","),
                name: "Jarren Ralf",
                subject: "Pacific Net & Twine (PNT) Order Form",
                htmlBody: htmlTemplate.evaluate().getContent(),
                inlineImages: emailImages
              });

              spreadsheet.toast('The instructional email has been sent to ' + custSS[0], 'Email Sent')
              rng.offset(i, 7 - rng.getColumn(), 1, 2).check();
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
 * This function formats the customers spreadsheets. A trigger runs this function daily. Due to the amount of possible customer spreadsheets, in order to avoid maxing out the runtime,
 * each day of the week it formats different customers. The result is each customer gets their spreadsheet formatted once a week.
 * 
 * @throws General error if anything goes wrong
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
            .setBackgrounds(new Array(maxRows).fill(['#cccccc', '#4a86e8', '#cccccc', '#cccccc', '#cccccc', 'white', 'white']))
            .setFontColors(new Array(maxRows).fill(['#434343', '#4a86e8', '#434343', '#434343', '#434343', 'black', 'black']))
            .setFontFamily('Arial')
            .setFontLine('none')
            .setFontSize(10)
            .setFontStyle('normal')
            .setFontWeight('bold')
            .setHorizontalAlignments(new Array(maxRows).fill(['left', 'center', 'center', 'center', 'center', 'left', 'left']))
            .setNumberFormat('@')
            .setVerticalAlignment('middle')
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
          .offset(-1, 1, maxRows, 1).setBorder(false, true, false, true, null, false, '#1155cc', SpreadsheetApp.BorderStyle.SOLID_THICK) // The vertical blue line below the header
          .offset(-2, 2, 1, 3).setBorder(false, false, false, false, false, null) // The checkbox, timestamp, and item information display
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
 * @throws General error if anything goes wrong
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
      var exportRange, itemPricing, exportData_WithDiscountedPrices = [];
      const numSS = dashboard.getLastRow() - 1;
      const dashboardRange = dashboard.getRange(2, 1, numSS, 5);

      const discountSheet = SpreadsheetApp.openById('1gXQ7uKEYPtyvFGZVmlcbaY6n6QicPBhnCBxk-xqwcFs').getSheetByName('Discount Percentages')
      const discounts = discountSheet.getSheetValues(2, 11, discountSheet.getLastRow() - 1, 5);
      const BASE_PRICE = 1;
      var PRICE, index;

      dashboardRange.getValues().map((customer, c) => {
        if (dashboard.getRange(c + 2, 3).isChecked())
        {
          PRICE = (customer[0] !== 'DL1015') ? 4 : 3;
          exportRange = SpreadsheetApp.openByUrl(customer[4]).getSheetByName('Export').getDataRange();

          exportRange.getValues().map(item => {
            if (item[0] === 'H') // Header line
              exportData_WithDiscountedPrices.push(['H', item[1], item[2], item[3]])
            else if (item[0] === 'I') // Comment line
              exportData_WithDiscountedPrices.push(['I', item[1], '', ''])
            else if (item[0] === 'D') // Detail line
            {
              itemPricing = discounts.find(sku => sku[0].split(' - ').pop().toString().toUpperCase() === item[1]); // Find the item pricing on the discount sheet

              if (itemPricing == undefined) // SKU is assumed to be invalid
                exportData_WithDiscountedPrices.push(['D', 'MISCITEM', 0, item[3]], ['C', 'SKU Not Found: ' + item[1] + ' therefore it was replaced with MISCITEM', '', ''])
              else // SKU is assumed to be valid
              {
                if (itemPricing[BASE_PRICE] != 0 && itemPricing[PRICE] != 0)
                  item[2] = (itemPricing[BASE_PRICE]*(100 - itemPricing[PRICE])/100).toFixed(2); // Set the pricing

                exportData_WithDiscountedPrices.push(['D', item[1], item[2], item[3]])
              }
            }
            else if (item[0] === 'C') // Comment line
            {
              if (item[1].split('Description: ').length === 2)
              {
                index = 1;

                while (exportData_WithDiscountedPrices[exportData_WithDiscountedPrices.length - index][0] === 'C')
                  index++;

                if (exportData_WithDiscountedPrices[exportData_WithDiscountedPrices.length - index][1] === 'MISCITEM')
                  exportData_WithDiscountedPrices.push(['C', item[1], '', '']) // Only keep the Description comment if the item is a MISCITEM
              }
              else
                exportData_WithDiscountedPrices.push(['C', item[1], '', ''])
            }
          })

          exportRange.clear();
          dashboard.getRange(c + 2, 3).uncheck();
          SpreadsheetApp.flush();
          spreadsheet.toast(customer[1] + '\'s export page has been cleared.')
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

      ranges.map((rngs, r) => (rngs.length !== 0) ? exportSheet.getRangeList(rngs).setBackground(backgroundColours[r]) : false); // Set the appropriate background colours
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
 * This function sets all of the protections necessary to keep as much of the data and functionality safe in this sheet.
 * 
 * @param {Spreadsheeet} ss : The customer's spreadsheet.
 * @author Jarren Ralf
 */
function protectSpreadsheet(ss)
{
  // Since the number of items change, we need to adjust some of the protected and unprotected ranges
  var unprotectedRanges, lastRow = ss.getSheetByName('Item List').getLastRow();

  ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).map(protection => {
    unprotectedRanges = protection.getUnprotectedRanges();

    if (unprotectedRanges.length > 0)
      protection.setUnprotectedRanges(unprotectedRanges.map(range => (range.getLastRow() > 5) ? range.offset(0, 0, lastRow, range.getNumColumns()) : range));

    protection.removeEditors(protection.getEditors());
    protection.addEditor('pntnoreply@gmail.com');
  });

  ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).map(protection => {protection.setRange(protection.getRange().offset(0, 0, lastRow))});
}

/**
 * This function runs on a trigger every X and it removes any editors from the each customer spreadsheet that are not contained in the corresponding Customer Email(s) column on the Dashboard.
 * It also checks if the drawings are missing an assigned script, if they are, then it reassigns them.
 * 
 * @throws General error if anything goes wrong
 * @author Jarren Ralf
 */
function removeUnapprovedEditorsFromCustomerSpreadsheet()
{
  const dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var approvedEditors, email, drawings;

  try
  {
    dashboard.getSheetValues(2, 2, dashboard.getLastRow() - 1, 7).map((custSS, i) => {
      if (isNotBlank(custSS[3]) && custSS[6]) // Spreadsheet URL is not blank and that the customer has already received their instructional email
      {
        ss = SpreadsheetApp.openByUrl(custSS[3]).addEditor('pntnoreply@gmail.com'); // Make sure pntnoreply@gmail.com is an editor

        // These are the drawings that are used as buttons for the users
        drawings = ss.getSheetByName('Item Search').getDrawings().map(drawing => {
          return {'button': drawing, 'x': drawing.getContainerInfo().getOffsetX(), 'w': drawing.getWidth(), 'isScriptNotAssigned': drawing.getOnAction() === ''}
        })

        if (drawings[0].isScriptNotAssigned) // If Script is missing, then assign it back to the button
          drawings[0].button.setOnAction((drawings[0].x < drawings[1].x && drawings[0].w < drawings[1].w) ? 'allItems' : 'addSelectedItemsToOrder');
        
        if (drawings[1].isScriptNotAssigned) // If Script is missing, then assign it back to the button
          drawings[1].button.setOnAction((drawings[1].x < drawings[0].x && drawings[1].w < drawings[0].w) ? 'allItems' : 'addSelectedItemsToOrder');

        approvedEditors = ['jarrencralf@gmail.com', 'pntnoreply@gmail.com'];
        approvedEditors.push(...custSS[4].split(',').map(email => email.trim())); // Get the list of approved editors and add it to jarrencralf@gmail and pntnoreply@gmail

        Logger.log(custSS[0] + ' approved editors:')
        Logger.log(approvedEditors)

        for (var editor = 0; editor < approvedEditors.length; editor++) // Loop through all of the employee emails and share with them individually to avoid errors
        {
          try
          {
            ss.addEditor(approvedEditors[editor]);
            Logger.log('Successfully shared with: ' + approvedEditors[editor])
          }
          catch (e)
          {
            var error = e['stack'];
            Logger.log('*Unsuccessful*')
            Logger.log(error);
          }
        }

        currentEditors = ss.getEditors().map(user => {
          email = user.getEmail();

          if (!approvedEditors.includes(email)) // If an editor is not on the approved email list, then remove them
          {
            Logger.log('**Editor removed from ' + custSS[0] + '\'s Order sheet: ' + email)
            ss.removeEditor(email);
          }
            
        });

        protectSpreadsheet(ss);
        dashboard.getRange(2 + i, 7).check(); // Check the box that signals if the spreadsheet is appropriately shared with the relevant emails
        Logger.log('------------------------------------------------------------------------------------------------------------------------------------------------------------')
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
function sendConfirmationEmail(name, ssUrl)
{
  const spreadsheet = SpreadsheetApp.openByUrl(ssUrl);
  const lastExportSheet = spreadsheet.getSheetByName('Last Export');
  const numRows = lastExportSheet.getLastRow() - 1;
  const values = lastExportSheet.getSheetValues(1, 1, numRows + 1, 5)
  const orderInfo = values.shift(); // Includes PO number and Delivery Instructions
  const isPoNotBlank = isNotBlank(orderInfo[3]);
  const numCols = values[0].length;
  const customerEmails = spreadsheet.getEditors().map(editor => editor.getEmail()).filter(email => email !== 'jarrencralf@gmail.com' && email !== 'pntnoreply@gmail.com').join(', ');

  var body = "<table><tr><th colspan=\"" 
    + numCols + "\">Provide your preferred delivery / pick up date and location below:</th></tr><tr><th colspan=\"" 
    + numCols + "\">" 
    + ((isNotBlank(orderInfo[4])) ? orderInfo[4] : "**Customer left this field blank**") 
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
    to: "deryk@pacificnetandtwine.com, scottnakashima@hotmail.com, eryn@pacificnetandtwine.com, triteswarehouse@pacificnetandtwine.com",
    replyTo: customerEmails,
    name: 'PNT Sales',
    subject: (isPoNotBlank) ? name + " has placed an order, PO # " + orderInfo[3] : name + " has placed an order!",
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
    subject: (isPoNotBlank) ? "Order Confirmation for PO # " + orderInfo[3] : "Order Confirmation",
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
    templateHtml.scriptURL   = "https://script.google.com/u/0/home/projects/1tuY0zWpu_kZtb6TQDsxgYligCOs159qQvQ5bj_nhZTq1sNR8T8LC--Wz/edit";
    var emailBody = templateHtml.evaluate().append(formattedError).getContent();
    
    MailApp.sendEmail({      to: 'lb_blitz_allstar@hotmail.com',
                           name: 'Jarren Ralf',
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
 * @throws General error if anything goes wrong
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
      var ss, emails;

      dashboard.getActiveRangeList().getRanges().map(rng => {
        rng.offset(0, 2 - rng.getColumn(), rng.getNumRows(), 5).getValues().map((custSS, i) => {
            if (isNotBlank(custSS[3]) && isNotBlank(custSS[4]))
            {
              ss = SpreadsheetApp.openByUrl(custSS[3]);
              emails = custSS[4].split(',').map(email => email.trim())
              Logger.log('Sharing spreadsheet with ' + custSS[0] + ' employee emails:')

              for (var i = 0; i < emails.length; i++) // Loop through all of the employee emails and share with them individually to avoid errors
              {
                try
                {
                  ss.addEditor(emails[i]);
                  Logger.log('Successfully shared with: ' + emails[i])
                }
                catch (e)
                {
                  var error = e['stack'];
                  Logger.log('*Unsuccessful*')
                  Logger.log(error);
                }
              }

              protectSpreadsheet(ss);
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
    range.offset(0, -1, 1, 8).setValues([['', '', false, '', '', '', false, false]]); // Remove the customer number if there is one
}

/**
 * This function updates the customers spreadsheets. A trigger runs this function daily. The list of items are updated, as well as the customer's
 * velocity report and invoice data.
 * 
 * @throws General error if anything goes wrong
 * @author Jarren Ralf
 */
function updateCustomerSpreadsheets()
{
  var splitDescription, newDescription, ss, d, numRows, velocityReportSheet, velocityReportSheetName, horizontalAligns, 
  customerName, chart, chartTitleInfo, invoiceSheet, customerInvoiceData, itemList = [], colours = [], colourSelector = true;

  try
  {
    const lodgeSalesSS = SpreadsheetApp.openById('1o8BB1RWkxK1uo81tBjuxGc3VWArvCdhaBctQDssPDJ0');
    const invoiceDataSheet = SpreadsheetApp.openById('1xKw4GAtNbAsTEodCDmCMbPCbXUlK9OHv0rt5gYzqx9c').getSheetByName('All Data');
    const invoiceData = invoiceDataSheet.getSheetValues(2, 1, invoiceDataSheet.getLastRow() - 1, 8).map(item => {
      item[4] = (item[4] === '100') ? 'Richmond' : (item[4] === '200') ? 'Parksville' : 'Prince Rupert';
      splitDescription = item[0].split(' - ');
      splitDescription.splice(-4, 1);
      item[0] = splitDescription.join(' - ');

      return item;
    })
    
    const dashboard = SpreadsheetApp.getActive().getSheetByName('Dashboard');
    const customerListSheet = lodgeSalesSS.getSheetByName('Customer List');
    const customerList = customerListSheet.getSheetValues(3, 1, customerListSheet.getLastRow() - 2, 3);
    const numYears = new Date().getFullYear() - 2011;
    const CUST_NAME = 0, SALES_TOTAL = 2;
    const white = ['white', 'white', 'white', 'white', 'white', 'white', 'white', 'white'], blue = ['#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8', '#c9daf8'];

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
    const numItems = itemList.length;

    dashboard.getRange(2, 1, dashboard.getLastRow() - 1, 5).getValues()
      .map(customer => {
        if (isNotBlank(customer[4]))
        {
          Logger.log('Updating ' + customer[1] + '\'s spreadsheet...')
          ss = SpreadsheetApp.openByUrl(customer[4])
          ss.getSheetByName('Item List').hideSheet().getRange(1, 1, numItems).setValues(itemList);
          ss.getSheetByName('Recently Created').hideSheet().getRange(1, 1, numItems).setValues(sortedItems);
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
          velocityReportSheet.protect();
          ss.moveChartToObjectSheet(chart).setName('Chart').setTabColor('#f1c232');
          colours.length = 0; // Clear the background colours array
          customerName = velocityReportSheetName[1].toUpperCase();

          customerInvoiceData = invoiceData.filter(name => name[1].toUpperCase() === customerName) // Customer invoice data
            .map((line, i, arr) => {

              if (i === 0)
                colourSelector = true;
              else if (line[2].toString().trim() != arr[i - 1][2].toString().trim()) // If the current invoice number does not match the current one, then switch the background colours
                colourSelector = !colourSelector;

              colours.push((colourSelector) ? white : blue);
              
              return line;
            })

          invoiceSheet = ss.insertSheet('Past Invoices', {template: ss.getSheetByName('Template').hideSheet()}).showSheet()
          ss.getSheetByName('Export').hideSheet();
          ss.getSheetByName('Last Export').hideSheet();
          ss.getSheetByName('Template').hideSheet();
          ss.getSheetByName('Submitted Orders').showSheet();

          numRows = customerInvoiceData.length;
          horizontalAligns = new Array(numRows).fill(['left', 'right', 'right', 'center', 'center', 'center', 'right', 'right']);

          invoiceSheet.getRange(2, 1, numRows, 8).setNumberFormat('@').setBackgrounds(colours).setHorizontalAlignments(horizontalAligns).setValues(customerInvoiceData);
          invoiceSheet.deleteColumn(2).protect();
          Logger.log(customer[1] + ' spreadsheet update complete.')
          Logger.log('------------------------------------------------------------------------------------------------------------------------------------------------------------')
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
 * This function updates the item list and recently created items sheet on the template spreadsheet.
 * 
 * @throws General error if anything goes wrong
 * @author Jarren Ralf
 */
function updateOrderSheet_TEMPLATE()
{
  var splitDescription, newDescription, d,  itemList = [];

  try
  {
    const spreadsheet = SpreadsheetApp.openById('1hhKeKpoheS71KuVCb9k1yPLq8PV41tIFsIj8U8VZUa4') // The template spreadsheet
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
    const numItems = itemList.length;
    const itemSearchSheet = spreadsheet.getSheetByName('Item Search');
    const ordersSheet = ss.getSheetByName('Submitted Orders').showSheet();
    spreadsheet.getSheetByName('Export').clearContents().hideSheet();
    spreadsheet.getSheetByName('Last Export').clearContents().hideSheet();
    spreadsheet.getSheetByName('Item List').hideSheet().getRange(1, 1, numItems).setValues(itemList);
    spreadsheet.getSheetByName('Recently Created').hideSheet().getRange(1, 1, numItems).setValues(sortedItems);
    spreadsheet.getSheetByName('Template').hideSheet();
    ordersSheet.getRange(2, 1, ordersSheet.getMaxRows() - 1, ordersSheet.getLastColumn()).clear();
    itemSearchSheet.getRange(1, 1).setValue('') // The search box
      .offset(0,  5).setValue('') // PO #
      .offset(1,  0).setValue('Items displayed in order of newest to oldest.') // Display message
      .offset(0,  1).setValue('') // Delivery Instructions
      .offset(0, -3).uncheck()    // Uncheck the Submission Box
      .offset(2, -2).setValue('') // Remove the hidden timestamp
      .offset(1, -1, itemSearchSheet.getMaxRows() - 4, 7).clearContent() // Clear the previous items and any order information
      .offset(0,  0, numItems, 1).setValues(sortedItems) // Set the recent items on the sheet
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
 * @author Jarren Ralf
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