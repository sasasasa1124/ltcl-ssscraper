const {google} = require("googleapis");
var conf = require('config');

/**
 *
 * @param {number} row - The row number of the cell reference. Row 1 is row number 0.
 * @param {number} column - The column number of the cell reference. A is column number 0.
 * @returns {string} Returns a cell reference as a string using A1 Notation
 *
 * @example
 *
 *   getA1Notation(2, 4) returns "E3"
 *   getA1Notation(2, 4) returns "E3"
 *
 */
 const getA1Notation = (row, column) => {
    const a1Notation = [`${row + 1}`];
    const totalAlphabets = "Z".charCodeAt() - "A".charCodeAt() + 1;
    let block = column;
    while (block >= 0) {
      a1Notation.unshift(
        String.fromCharCode((block % totalAlphabets) + "A".charCodeAt())
      );
      block = Math.floor(block / totalAlphabets) - 1;
    }
    return a1Notation.join("");
  };

/**
 *
 * @param {string} cell -  The cell address in A1 notation
 * @returns {object} The row number and column number of the cell (0-based)
 *
 * @example
 *
 *   fromA1Notation("A2") returns {row: 1, column: 3}
 *
 */
 const fromA1Notation = (cell) => {
    let [, columnName, row] = cell.toUpperCase().match(/([A-Z]+)([0-9]+)/);
    row = parseInt(row);
    const characters = "Z".charCodeAt() - "A".charCodeAt() + 1;
  
    let column = 0;
    columnName.split("").forEach((char) => {
      column *= characters;
      column += char.charCodeAt() - "A".charCodeAt() + 1;
    });
  
    return { row, column };
  };
  
// Returns an array of dates between the two dates
function getDates (startDate, endDate) {
    const dates = []
    let currentDate = startDate
    const addDays = function (days) {
      const date = new Date(this.valueOf())
      date.setDate(date.getDate() + days)
      return date
    }
    while (currentDate <= endDate) {
      dates.push(currentDate)
      currentDate = addDays.call(currentDate, 1)
    }
    return dates
  }

/**
 * change name into corresponding slack user id
 * @param  {String} name each mentor's name in the spreadsheet
 * @return {String} corresponding slack user id
 */
 const nameToId = function (name)
 {
     const table = conf.name_id_table;
     return table[name];
 }
 
 /**
  * change name into corresponding slack mention id
  * @param  {String} name each mentor's name in the spreadsheet
  * @return {String} corresponding slack mention id
  */
  const nameToMention = function (name)
 {
     return (conf.name_table.includes(name) ? `<@${nameToId(name)}>` : name);
 }
 
 /**
  * change name into corresponding spreadsheet's row number
  * it should search corresponding row number by searching in spreadsheet objects ** with constant column B;
  * 
  * @param  {String} name each mentor's name in the spreadsheet
  * @return {String} corresponding ss row number
  */
  const nameToRow = async function (name)
 {
     const todayObject = new Date();
     var corresponding_row = 0
     const client = await auth.getClient();
     const googleSheets = await google.sheets({version:"v4", auth:client});
     const rows = await googleSheets.spreadsheets.values.get({
         auth: auth,
         spreadsheetId: spreadsheetsId,
         range: `${todayObject.getFullYear()}年${todayObject.getMonth() + 1}月!B:B`
     });
     rows.data.values.forEach((item,key) => {
         if (item == name) {
             corresponding_row = key;
         }
     });
     return corresponding_row + 1;
 }
 
 const googleSheets = async function ()
 {
     const client = await auth.getClient();
     const googleSheets = await google.sheets({version:"v4", auth:client});
 }
 
 /**
  * mapping from days to ss columns
  * @param  {String} day each mentor's name in the spreadsheet
  * @return {String} corresponding spreadsheet column id
  */
  const dayToColumn = function (day){
     return ['', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG'][day];
 }
 
 /**
  * formatting datetime object into an array of [year,month,date];
  * @param  {Object} obj datetime Object of JavaScript
  * @return {Array} the corresponding array of datetime object
  */
  const datetimeToArray = function (obj)
 {
     return {
         'year': obj.getFullYear(),
         'month': obj.getMonth() + 1,
         'date': obj.getDate(),
         'string': `${('0' + (obj.getMonth()+1)).slice(-2)}/${('0' + obj.getDate()).slice(-2)}`
     };
 }
 
 /**
  * formatting datetime object into an array of [year,month,date];
  * @param {Object} obj datetime Object of JavaScript
  * @param {int} days difference of date in integer format
  * @return {Object} datetime Object
  */
  const deltaDate = function (todayObj,days)
 {
     obj = new Date(todayObj.getTime() + days * 60 * 24 * 60000);
     return obj
 }

   /**
  * formatting datetime object into an array of [year,month,date];
  * @param {Array} array array object
  * @return {Array} array object transposed
  */
 const transposeArray = function (rawArray)
 {
     const array = rawArray[0].map((col,i) => rawArray.map(row => row[i]));
     return array;
 }
 
  /**
  * formatting datetime object into an array of [year,month,date];
  * @param {Object} rowObj row Object of Google Sheets
  * @return {Array} Array of Array, which include row object 
  */
 const rowObjToArray = function (rowObj)
 {
     var rawArray = new Array();
     console.log(rowObj.data.valueRanges);
     rowObj.data.valueRanges.map((obj) => {
         rawArray.push(obj.values);
     });
    //  const array = transposeArray(rawArray);
     return rawArray;
 }

  /**
  * formatting datetime object into an array of [year,month,date];
  * @param {Object} authorization object of google apis
  * @param {int} spreadsheetId
  * @param {Array} ranges query for google sheets
  * @return {Object} row Ojbect
  */
 const fetchGoogleSheets = async function (auth, spreadsheetsId, ranges)
 {
    const client = await auth.getClient();
    const googleSheets = await google.sheets({version:"v4", auth:client});
    const row =  await googleSheets.spreadsheets.values.batchGet({
        auth: client,
        spreadsheetId: spreadsheetsId,
        ranges: ranges,
    }); 
    return row;
 }

   /**
  * fetching date in format of '01/01'
  * @param {Array} datetime array
  * @return {Array} string array, each element is like '01/01'
  */
const dateFetch = function(datetime)
{
    const result = new Array();
    const month = datetime['month'];
    const datetimes =  getDates(new Date(datetime['year'], datetime['month']-1, 1), new Date(datetime['year'], datetime['month'], 1));
    datetimes.map((datetime) => {
        result.push(`${('0' + (datetime.getMonth()+1)).slice(-2)}/${('0' + datetime.getDate()).slice(-2)}`);
    });
    return result.slice(0,-1);
}

   /**
  * 
  * @param {String} date string array with formated date in '01/01'
  * @param {Array} matrix 2d arrays
  * @return {Object} the key is date and the value is cell index, like {'01/01':"C12"}
  */
const dateCell = function(date, matrix)
{
    result = {};
    matrix.map((array,i) => array.map((element,j) => {
        if (date == element) {
            result[element] = getA1Notation(i,j);
        }
    }));
    return result;
}

   /**
  * 
  * @param {Array} mentors string array with formated mentor in '01/01'
  * @param {Array} matrix 2d arrays
  * @return {Object} the key is mentor and the value is cell index, like {'01/01':"C12"}
  */
const mentorCell = function(mentors, matrix)
{
    result = {};
    matrix.map((array,i) => array.map((element,j) => {
        if (mentors.includes(element)) {
            if (typeof result[element] === 'undefined') {
                result[element] = getA1Notation(i,j);
            }
        }
    }));
    return result;
}

const fetchMentorDate = function(mentorCell,dateCell,matrix)
{
    const column = fromA1Notation(dateCell)['column'];
    const row = fromA1Notation(mentorCell)['row'];
    return matrix[row-1][column-1];
}

const fetchColumnDate = function(columnCell,dateCell,matrix)
{
    return matrix[fromA1Notation(columnCell)['row']-1][fromA1Notation(dateCell)[['column']]-1];
}

const formatRawRow = function(row)
{
    let matrix = [];
    row.data.valueRanges.map(obj => obj['values'].map(array => matrix.push(array)));
    return matrix;
}

const searchColumnNum = function(columnNames, matrix, headerRowNum = conf.spreadsheet.headerRow)
{
    result = [];
    headerRow = matrix.map(array => array[0]);
    columnNames.map((columnName) => {
        headerRow.map((element,index) => {
            if (columnName == element) {
                result.push({columnName : getA1Notation(index,headerRowNum)});
            }
        });
    })
    return result;
}

// const dateKeys = function (date,keyNames,matrix)
// {
//     dateCell = dateCell(date['string'],matrix)[date[['string']]];
//     keyCells.map((element,index) => {
//         todayKeys.push({
//             "key": keyNames[index],
//             "value" : utils.fetchColumnDate(element['columnName'],todayCell,matrix)
//         });
//     });
// }

const keymsgGenerator = function(today,todayKeys,changes)
{
    return (
        `
        \n
        ${today['month']}/${today['date']}\n
        本日の鍵所持者:\n
        ${todayKeys.map((element) => {
            return element['value'];
        })}\n
        本日→明日の鍵の移動\n
        ${changes ? changes : '移動なし\n'}
        \n
        詳細確認:${conf.spreadsheet.url}\n
        `
    )
}

const self_studymsgGenerator = function(today,todayManagers)
{
    const managerA = todayManagers.find((o) => o.Room == '受付兼オンライン質問部屋常駐\n（Aスペース）');
    const managerB = todayManagers.find((o) => o.Room == 'オンライン質問部屋常駐\n（Bスペースorオンライン）');
    return (
    `
    \n
    ${today['month']}/${today['date']}\n
    本日の自習室勤務\n
    Aスペース(受付兼オンライン質問部屋常駐):${nameToMention(managerA.value)}\n
    Bスペース(オンライン質問部屋常駐):${nameToMention(managerB.value)}\n
    詳細確認:${conf.spreadsheet.url}\n
    本日勤務状況確認:${conf.sites.url}
    `
    )
}

 const today = datetimeToArray(new Date());
 const tommorow = datetimeToArray((deltaDate(new Date(), 1)));
 const yesterday = datetimeToArray((deltaDate(new Date(), 1)));

 module.exports = {
     nameToId,
     nameToMention,
     nameToRow,
     dayToColumn,
     datetimeToArray,
     deltaDate,
     rowObjToArray,
     fetchGoogleSheets,
     dateFetch,
     dateCell,
     mentorCell,
     fetchMentorDate,
     fetchColumnDate,
     formatRawRow,
     searchColumnNum,
     keymsgGenerator,
     self_studymsgGenerator,
     today,
     tommorow,
     yesterday
 }