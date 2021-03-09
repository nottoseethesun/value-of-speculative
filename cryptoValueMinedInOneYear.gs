/**
 * Google Add-on Title: "Mined Crypto Tax Basis"
 *
 * What: This is a crypto mining accounting utility in the form of a Google Apps Script program.
 *
 * Disclaimer: This is not tax advice and is not financial advice.
 *             Speak to a registered financial advisor and registered tax advisor.
 *
 * Purpose: This is for finding what is in many jurisdictions the taxable amount of crypto mined for a single given year.
 *
 * How:
 *
 * It reads data in from a Google Sheet and adds up this value for a Gregorian calendar year.
 *
 * It's pretty easy to use because you just paste mined amounts and prices into a Google Sheet,
 * make a few customizations here, and run this script.
 *
 * For a single given year, this script gets the prices from a tab in a Google Docs spreadsheet,
 * as exported for example by your favorite crypto listing service, such as https://www.nomics.com,
 * and gets the amounts and date mined from a different tab in the Google Docs spreadsheet.
 * Then it computes the total value mined for the Gregorian calendar year.
 *
 * Required Input:
 *
 *   a) From a tab named "mined" on the sheet: Rows that each have the date, an amount,
 *      and the transaction type (such as "mined", "transferred", and so forth);
 *
 *   b) From a tab named "prices" on the sheet: Rows that each have the date and
 *      the day's price for the particular crypto asset.
 *
 * Usage:
 *
 *   a) Make a backup of  your spreadsheet before using this!  Not responsible for any loss of data.
 *   b) Not responsible for incorrect results or loss or damage from those.
 *   c) Set up your spreadsheet with the Required Input as described above.
 *   d) From the Sheet, use menu "Tools -> <> Script editor" and paste in this code.
 *      That makes this a "bound script".
 *   e) Note that the code expects dates in the spreadsheet tab with the mined data to be in the format,
 *      `yyyy-mm-dd` or `yyyy-mm-ddT...*`.
 *   f) Enter the manual config data in the section thus marked below, in the code.
 *      To find it, search on "Start: Manual Config Data Section", below.
 *   g) At the time of writing, you'll need to enable Advanced Google Services for
 *      the dependencies, such as `Sheets`, to be loaded, as described
 *      at https://stackoverflow.com/a/47309054/566260
 *
 *  Help:
 *
 *   Google Apps Scripts API Doc: https://developers.google.com/apps-script/reference
 *
 */
function cryptoValueMinedInOneYear() {
  // - - - Start: Manual Config Data Section.
  /*-
   * Below, you'll  need to manually change/fill in the various id's, columns, and
   * ranges for the spreadsheet tabs:
   */
  const year = 2020;
  const spreadSheetId = '1FiIRjHO0fHMhC7fDMhJRc-WmVKU8N86VytnjTTHIC1k';
  const pricesSheetTabId = 'prices';
  const minedSheetTabId = 'mined';

  const minedDateColumn = 'A';
  const minedAmountColumn = 'F';

  const priceStartRowIndex = 2;
  // Note: The end row is `priceStartRowIndex` plus the number of days in your year.
  const pricesColumn = 'B';

  const miningStartRowIndex = 2; // First row that has mining data
  const miningEndRowIndex = 321; // Last row that has mining data

  const transactionTypeColumn = 'H';
  /*-
   * Set this to the value you need: Otherwise, ordinary transactions
   * could be counted.
   * Optionally add options such as "pos" like so: /^(mined|pos)/i
   */
  const acceptableTransactionTypes = /^(mined)/i;
  // - - - End: Manual Config Data Section.

  const RESULTS_MSG = 'Total Value Mined';
  const alertWidth = 400;
  const alertHeight = 400;

  // Now start the execution:
  const totalValueMined = getTotalValueMined();
  displayResult(totalValueMined);

  //  - - - Supporting Functions:

  function getTotalValueMined() {
    const miningRecords = getMiningRecords();
    console.log(miningRecords);
    const total = miningRecords.reduce(function(acc, record) {
      return acc + record.minedValue;
    }, 0);

    return total;
  }

  function getMiningRecords() {
    const minedDates = getMinedDates();
    const minedAmounts = getMinedAmounts();
    const transactionTypes = getTransactionTypes();
    const prices = getPrices();
    const records = minedDates.reduce(function(acc, date, index) {
      const transactionType = transactionTypes[index];
      if (!acceptableTransactionTypes.test(transactionType)) {
        return acc;
      }
      const dayNumber = getDayOfYear(date);
      const price = prices[dayNumber];
      const minedAmount = minedAmounts[index];
      const minedValue = price * minedAmount;
      const record = {
        minedDate: date,
        dayNumber: dayNumber,
        price: price,
        minedAmount: minedAmount,
        minedValue: minedValue,
      };
      acc.push(record);
      return acc;
    }, []);

    return records;
  }

  function getPrices() {
    const DAYS_IN_YEAR = getDaysInYear();
    const pricesEndRowIndex = priceStartRowIndex + DAYS_IN_YEAR;
    const pricesRange = // Google Scripts doesn't do multi-line js templates yet.
      `${pricesSheetTabId}!${pricesColumn}${priceStartRowIndex}:${pricesColumn}${pricesEndRowIndex}`;
    const pricesValue = Sheets.Spreadsheets.Values.get(spreadSheetId, pricesRange);
    const prices = pricesValue.values.map(function (stringPrice) {
      return parseFloat(stringPrice[0]);
    });

    return prices;
  }

  function getMinedDates() {
    const DATE_STRIPPER = /T.*$/; // For transforming from '2019-12-28T12:16:16' to '2019-12-28'.
    const datesValue = Sheets.Spreadsheets.Values.get(
      spreadSheetId, getMinedRange(minedDateColumn));
    const dates = datesValue.values.map(function (date) {
      return date[0].replace(DATE_STRIPPER, '');
   });

    return dates;
  }

  function getMinedAmounts() {
    const amountsValue = Sheets.Spreadsheets.Values.get(
      spreadSheetId, getMinedRange(minedAmountColumn));
    const amounts = amountsValue.values.map(function (amount) {
      return parseFloat(amount[0]);
    });

    return amounts;
  }

  function getTransactionTypes() {
    const transactionTypesValue = Sheets.Spreadsheets.Values.get(
      spreadSheetId, getMinedRange(transactionTypeColumn));
    const transactionTypes = transactionTypesValue.values.map(function (transactionType) {
      return transactionType[0];
   });

    return transactionTypes;
  }

  function getMinedRange(column) {
    const mr =
      `${minedSheetTabId}!${column}${miningStartRowIndex}:${column}${miningEndRowIndex}`;
    return mr;
  }

  /**
   * @param yearMonthDay string E.g., 2019-01-01
   * @return 0-365 (zero-based index).
   */
  function getDayOfYear(yearMonthDay /* year-month-day */) {
    const MS_IN_SEC = 1000;
    const MIN_IN_HR = 60;
    const SEC_IN_MIN = 60;
    const HR_IN_DAY = 24;

    const parts = yearMonthDay.split(/\-/g);
    const year = parts[0];
    const month = parts[1];
    const day = parts[2];

    const specified = new Date(year, month - 1, day - 1);
    const start = new Date(year, 0, 0);
    const diff = (specified - start) +
      ((start.getTimezoneOffset() - specified.getTimezoneOffset()) *
      SEC_IN_MIN * MS_IN_SEC);
    const oneDay = MS_IN_SEC * SEC_IN_MIN * MIN_IN_HR * HR_IN_DAY;
    const numberOfDayInYear = Math.floor(diff / oneDay);

    return numberOfDayInYear;
  }

  function getDaysInYear() {
    const COMMON_YEAR_DAYS = 365;
    const LEAP_YEAR_DAYS = 366;

    if (year % 4 !== 0) {
      return COMMON_YEAR_DAYS;
    }
    if (year % 100 !== 0) {
      return LEAP_YEAR_DAYS;
    }
    if (year % 400 !== 0) {
      return COMMON_YEAR_DAYS;
    }
    return LEAP_YEAR_DAYS;
  }

  function displayResult(result) {
    const title = `Total Mined Value as Sum of Value Each Day`;
    const html = HtmlService.createHtmlOutput(`<h1>${title}</h1><h2>${result}</h2>`);
    console.info(`${title}: ${result}`);

    try {
      SpreadsheetApp.getUi().showModalDialog(html, RESULTS_MSG);
    } catch (err) {
      console.error(err);
      console.warn('You got that error since this is meant to be run as a bound, not stand-alone, spreadsheet.');
    }
  }

}
