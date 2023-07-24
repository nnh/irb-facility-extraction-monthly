/**
 * Main function to run the script.
 */
function main() {
  const [filteredData, indexes] = extractFacilitiesData_();

  const outputText = filteredData
    .map((row, idx) => (idx > 0 ? createOutputText_(row, indexes) : null))
    .filter(x => x !== null)
    .join('');
  outputFile_(outputText);
}

/**
 * Extracts facilities data from the spreadsheet.
 * @returns {Array[]} Array containing filtered data and indexes.
 */
function extractFacilitiesData_() {
  /** @type {Map<string, number>} */
  const columnIndexes = new Map([
    ['code', 0],
    ['name', 1],
    ['deptCode', 3],
    ['deptName', 4],
    ['prefectureName', 7],
    ['responsiblePerson', 10],
    ['j_sanka', 11],
    ['irb', 19],
  ]);

  /** @type {Map<string, number>} */
  const indexes = new Map();
  Array.from(columnIndexes).forEach(([name, _], idx) => indexes.set(name, idx));

  const inputSpreadsheetId =
    PropertiesService.getScriptProperties().getProperty('inputSpreadsheetId');
  const inputSpreadsheet = SpreadsheetApp.openById(inputSpreadsheetId);
  const facilitiesSheet = inputSpreadsheet.getSheetByName('施設一覧');
  const facilitiesSheetValues = facilitiesSheet.getDataRange().getValues();
  /** @type {Array<Array<*>>} */
  const extractedData = facilitiesSheetValues.map(row =>
    Array.from(columnIndexes).map(([_, index]) => row[index])
  );
  const filteredData = extractedData.filter(
    (row, idx) =>
      idx === 0 ||
      (row[indexes.get('j_sanka')] === '1' && row[indexes.get('irb')] === '1')
  );
  const removeDuplicateRecords = removeDuplicateRecords_(filteredData, indexes);
  const removeNewlinesFromFilteredData = removeNewlinesFromFilteredData_(
    removeDuplicateRecords
  );
  indexes.set('prefectureCode', indexes.size);
  const prefectureDataArray = getPrefecturesInOrder_(
    removeNewlinesFromFilteredData,
    inputSpreadsheet,
    indexes
  );
  const sortFilteredDataWithPrefecture = sortArray_(
    prefectureDataArray,
    indexes
  );
  return [sortFilteredDataWithPrefecture, indexes];
}
/**
 * Retrieves the prefectures in order and appends the prefecture code to the two-dimensional array.
 *
 * @param {Array<Array<string>>} arr - The input two-dimensional array.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} inputSpreadsheet - The input spreadsheet.
 * @param {Map<string, number>} indexes - The map containing column indexes.
 * @returns {Array<Array<string|number>>} - The modified two-dimensional array with prefecture codes appended.
 */
function getPrefecturesInOrder_(arr, inputSpreadsheet, indexes) {
  const prefectureSheet = inputSpreadsheet
    .getSheetByName('JPLSG_SV2 MST_県')
    .getDataRange()
    .getValues()
    .map((x, idx) => {
      const code = idx === 0 ? '県コード' : idx;
      return [x[0], code];
    });
  const prefectureMap = new Map(prefectureSheet);
  return arr.map(x => {
    const code = prefectureMap.has(x[indexes.get('prefectureName')])
      ? prefectureMap.get(x[indexes.get('prefectureName')])
      : -1;
    return [...x, code];
  });
}

/**
 * Sorts the array based on prefecture code and facility code.
 * @param {Array<Array<*>>} arr - Array to be sorted.
 * @param {Map<string, number>} indexes - Map containing column indexes.
 * @returns {Array<Array<*>>} Sorted array.
 */
function sortArray_(arr, indexes) {
  const res = arr.slice(1).sort((a, b) => {
    if (a[indexes.get('prefectureCode')] === b[indexes.get('prefectureCode')]) {
      return a[indexes.get('code')].localeCompare(b[indexes.get('code')]);
    }
    return a[indexes.get('prefectureCode')] - b[indexes.get('prefectureCode')];
  });
  return [arr[0], ...res];
}

/**
 * Creates the output text based on row data.
 * @param {Array<*>} row - Row data.
 * @param {Map<string, number>} indexes - Map containing column indexes.
 * @returns {string} Output text.
 */
function createOutputText_(row, indexes) {
  return `        <TR><TD>${row[indexes.get('prefectureName')]}</TD><TD>${
    row[indexes.get('name')]
  }</TD><TD>${row[indexes.get('deptName')]}</TD><TD>${
    row[indexes.get('responsiblePerson')]
  }</TD></TR>\n`;
}

/**
 * Outputs the text to a file.
 * @param {Array<string>} outputText - Array of output text.
 */
function outputFile_(outputText) {
  const outputFolderId =
    PropertiesService.getScriptProperties().getProperty('outputFolderId');
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const fileName = 'CHM14sankasisetu.txt';
  moveFileToFolder_(outputFolder, fileName);
  outputFolder.createFile(fileName, outputText, 'text/plain');
}

/**
 * Moves the file to a destination folder.
 * @param {GoogleAppsScript.Drive.Folder} sourceFolder - Source folder.
 * @param {string} fileName - File name to be moved.
 */
function moveFileToFolder_(sourceFolder, fileName) {
  const files = sourceFolder.getFilesByName(fileName);
  const destinationFolderId =
    PropertiesService.getScriptProperties().getProperty('saveFolderId');
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);
  while (files.hasNext()) {
    const file = files.next();
    destinationFolder.createFile(file);
    sourceFolder.removeFile(file);
  }
}
/**
 * Remove newline characters from all elements of a two-dimensional array.
 * @param {Array<Array<string>>} data - The two-dimensional array containing the data.
 * @returns {Array<Array<string>>} The modified two-dimensional array with newline characters removed.
 */
function removeNewlinesFromFilteredData_(data) {
  const modifiedData = data.map(row => {
    return row.map(element => element.replace(/\n/g, ''));
  });
  return modifiedData;
}
/**
 * Remove duplicate records based on facility code and keep the record with the smallest department code.
 * @param {Array<Array<string>>} data - The two-dimensional array containing facility data.
 * @param {Map<string, number>} indexes - The map containing column indexes for facility code and department code.
 * @returns {Array<Array<string>>} The modified two-dimensional array with duplicate records removed.
 */
function removeDuplicateRecords_(data, indexes) {
  const facilityCodeIndex = indexes.get('code');
  const departmentCodeIndex = indexes.get('deptCode');

  const uniqueRecords = [];
  const facilityCodes = new Set();

  for (const record of data) {
    const facilityCode = record[facilityCodeIndex];
    const departmentCode = record[departmentCodeIndex];

    if (!facilityCodes.has(facilityCode)) {
      uniqueRecords.push(record);
      facilityCodes.add(facilityCode);
    } else {
      const existingRecord = uniqueRecords.find(
        rec => rec[facilityCodeIndex] === facilityCode
      );
      if (existingRecord[departmentCodeIndex] > departmentCode) {
        uniqueRecords.splice(uniqueRecords.indexOf(existingRecord), 1);
        uniqueRecords.push(record);
      }
    }
  }

  return uniqueRecords;
}
