/**
 * This is a crypto mining accounting utility in the form of a Google Apps Script program, for example for
 * finding how much the $USD amount is that one mined of a given single crypto in a Gregorian calendar year.
 *
 * Note that Apps Script at the time of this writing does not have a lot of es2015+ and NodeJS features, so the code is a bit clunkier. :)
 *
 * For a single given year, gets the prices from a tab in a Google Docs spreadsheet, as exported for example by https://coingecko.com,
 * and gets the amounts and date mined from a different tab in the Google Docs spreadsheet.
 * Then it computes the total value mined for the Gregorian calendar year.
 *
 * Usage:
 *
 *   a) Make a backup of  your spreadsheet before using this!  Not responsible for any loss of data.
 *   b) Note that the code expects dates in the spreadsheet tab with the mined data to be in the format,
 *      `yyyy-mm-dd` or `yyyy-mm-ddT...*`.
 *   c) Enter the manual config data in the section thus marked below, in the code.
 *   d) At the time of writing, you'll need to enable Advanced Google Services for
 *      the dependencies, such as `Sheets`, to be loaded, as described at https://stackoverflow.com/a/47309054/566260
 *   e) Optionally, deploy via a Manifest: The i.d.  of Version 1.1.0 is:
 *        AKfycbwXfr83gRL_rdXfQ7mioHiF41HLEBDRXgC45zdFEoKSZCQXVIhtDxrPtXGjVKyhpo4 .
 */
(function cryptoValueMinedInOneYear() {
  // - - - Start: Manual Config Data Section.
  // You'll  need to manually fill in the various id's, columns, and ranges for the spreadsheet tabs:
  const year = 2019;
  const spreadSheetId = '1Qcvml2r94LJavCCG08tpPLbGc73yhiM7bzcKfum-FVY';
  const pricesSheetTabId = 'grc-usd-max-2019';
  const minedSheetTabId = 'grc-transactions-2020-01-01';

  const minedDateColumn = 'B';
  const minedAmountColumn = 'F';

  const priceStartRowIndex = 1404; // End row is + the number of days in your year.
  const pricesColumn = 'B';

  const miningStartRowIndex = 2;
  const miningEndRowIndex = 463;
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
    const prices = getPrices();
    const records = minedDates.map(function(date, index) {
      const dayNumber = getDayOfYear(date);
      const price = prices[dayNumber];
      const minedAmount = minedAmounts[index];
      const minedValue = price * minedAmount;
      return {
        minedDate: date,
        dayNumber: dayNumber,
        price: price,
        minedAmount: minedAmount,
        minedValue: minedValue,
      };
    });

    return records;
  }

  function getPrices() {
    const DAYS_IN_YEAR = getDaysInYear();
    const pricesEndRowIndex = priceStartRowIndex + DAYS_IN_YEAR;
    const pricesRange = '' + pricesSheetTabId + '!' + pricesColumn + priceStartRowIndex + ':' + pricesColumn + pricesEndRowIndex;
    const pricesValue = Sheets.Spreadsheets.Values.get(spreadSheetId, pricesRange);
    const prices = pricesValue.values.map(function (stringPrice) {
      return parseFloat(stringPrice[0]);
    });

    return prices;
  }

  function getMinedDates() {
    const DATE_STRIPPER = /T.*$/; // For transforming from '2019-12-28T12:16:16' to '2019-12-28'.
    const datesValue = Sheets.Spreadsheets.Values.get(spreadSheetId, getMinedRange(minedDateColumn));
    const dates = datesValue.values.map(function (date) {
      return date[0].replace(/T.*$/, '');
   });

    return dates;
  }

  function getMinedAmounts() {
    const amountsValue = Sheets.Spreadsheets.Values.get(spreadSheetId, getMinedRange(minedAmountColumn));
    const amounts = amountsValue.values.map(function (amount) {
      return parseFloat(amount[0]);
    });

    return amounts;
  }

  function getMinedRange(column) {
    return '' + minedSheetTabId + '!' + column + miningStartRowIndex + ':' + column + miningEndRowIndex;
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
    const diff = (specified - start) + ((start.getTimezoneOffset() - specified.getTimezoneOffset()) * SEC_IN_MIN * MS_IN_SEC);
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
    const html = HtmlService.createHtmlOutput('<h2>' + result + '</h2>');
    SpreadsheetApp.getUi().showModalDialog(html, RESULTS_MSG);
  }

})();
