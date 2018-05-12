//! (u)tility
//! version : 1.01
//! authors : Cody Henderson
//! license : MIT
//! https://beetstech.com
/**
*  
* 
* CONTENTS:
* - getFirstEmptyRowWholeRow()
* - prepareStatementVars()
* - convertMillis()
* - formatSheet()
* - readSelect()
* - write2GSheets()
* - write2Sql()
* - deDuplicateArr()
* - readFromTable()
* - clearPropServ()
*
*/


/**
* Given an active sheet, returns the first empty whole row
* 
* @param  {string} activeSheet
* @return {integer} row
* 
*/
function getFirstEmptyRowWholeRow(activeSheet){
  var range = activeSheet.getDataRange();
  var values = range.getValues(),
      row = 0;
  for(var row = 0; row<values.length; row++){
    if(!values[row].join('')) break;
  }
  return (row+1);
}


/**
* Performs a basic query in order to retrieve table meta, which is then used to
* get the column names and prepare the column name variable string required by
* prepareStatement().
* 
* @param    {object} vPack - Container for write-directive variables.
* - @param  {array}  colNames - (Optional) Array of the column name strings for the INSERT statement. If not provided, table will be queried,
*                               and INSERT statement will included all columns (except auto-incrementing).
* @param    {object} stmt - SQL connection object created with createStatement()
* @return   {object} stmtVars - Container for return variables.
* - @return {string} sqlTableCols - Comma separated string of column names.
* - @return {string} sqlColVars - Comma separated string of "?". One for each column name.
* 
*/

function prepareStatementVars(vPack, stmt){
  
  var results = stmt.executeQuery('SELECT * FROM '+vPack.sqlTable+' LIMIT 10');
  var numCols = results.getMetaData().getColumnCount(),
      isAutoIncrement = [],
      colNameArr = [];
  
  if(vPack.hasOwnProperty('colNames') === false){
    // Increment through each column to check if auto-incrementing.
    // Cannot put the isAutoIncrement() function directly in the following "if" statement. Error message,
    // and general problems all around (wrong values). isAutoIncrement() is finicky.
    for(var col=0; col<numCols; col++){
      var isAuto = results.getMetaData().isAutoIncrement(col+1);
      isAutoIncrement.push(isAuto);
    }
    
    for(var col=0; col<numCols; col++){
      if(isAutoIncrement[col] === false){
        colNameArr.push(results.getMetaData().getColumnName(col+1));  
      }
    }
    var stmtVars = {sqlTableCols: colNameArr.toString()};
  }else{
    colNameArr   = vPack.colNames;
    var stmtVars = {sqlTableCols: vPack.colNames.toString()};
  }
  
  var colVarsString = '?';
  for(var i=1; i<colNameArr.length; i++){colVarsString += ',?';}
  stmtVars.sqlColVars = colVarsString;
  
  return stmtVars;
}


/**
* Converts milliseconds to MM:SS
* 
* @param  {integer} millis
* @return {string} duration
* 
*/

function convertMillis(millis){
  var minutes  = Math.floor(millis/60000);
  var seconds  = ((millis%60000)/1000).toFixed(0);
  var duration = minutes+":"+(seconds<10 ? '0' : '')+seconds;
  return duration;
}


/**
* Applies standard Beetstech styling to a sheet
* 
* @param {string}  activeSheet - The active sheet object returned by getActiveSheet() or setActiveSheet()
* @param {string}  sheetPurpose (optional) - "script" (yellow), "dataval" (black), "importrange" (purple), "work" (green)
* @param {integer} rowsLimit (optional) - Number of rows from the last row that should receive height formatting.
*                                                  Useful when sheet has tens of thousands of rows and this process takes too long.
*/

function formatSheet(sheet, sheetPurpose, rowsLimit){
  
  var numCols   = sheet.getLastColumn();
  var numRows   = sheet.getLastRow();
  var headerRow = sheet.getRange(1, 1, 1, numCols);
  var allCells  = sheet.getRange(1, 1, numRows, numCols);
  
  if(typeof sheetPurpose === 'undefined' || sheetPurpose === 'script'){
    var tabColor = '#ffd966';
  }else if(sheetPurpose === 'dataval'){
    var tabColor = '#000000';
  }else if(sheetPurpose === 'importrange'){
    var tabColor = '#8e7cc3';
  }else if(sheetPurpose === 'manual'){
    var tabColor = '#cc4125';
  }else if(sheetPurpose === 'utility'){
    var tabColor = '#f6b26b';
  }else if(sheetPurpose === 'work'){
    var tabColor = '#57bb8a';
  }
  
  sheet.setTabColor(tabColor);
  sheet.setFrozenRows(1);
  allCells.setVerticalAlignment('middle');
  allCells.setFontFamily('Roboto Mono');
  allCells.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  allCells.setFontSize(10);
  
  headerRow.setBackground('#d9d9d9');
  headerRow.setHorizontalAlignment('center');
  headerRow.setFontWeight('bold');
  
  // This task can take a while for sheets with many rows
  // Use optional "rowsLimit" param to affect only the last {int} rows
  if(typeof rowsLimit === 'undefined' || rowsLimit === false){
    var rowStart = 1;
  }else if(typeof rowsLimit === 'integer'){
    var rowStart = (numRows-(numRows-rowsLimit));
  }else{
    return;
  }
  for(var i=rowStart; i <= numRows; i++){
     sheet.setRowHeight(i, 30)
  };
}


/**
* Reads values from a Google Sheet
* 
* @param   {object} readPackage
* - @param {string} readSheetID - The Google spreadsheet ID
* - @param {string} readSheetName - The individual sheet ID to read
* @return  {array}  values - Two-dimensional array of values, indexed by row, then by column
* 
*/

function readSelect(readPackage){
  
  var ss = SpreadsheetApp.openById(readPackage.readSheetID);
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(readPackage.readSheetName);
  var values = activeSheet.getDataRange().getValues();
  
  return values;
}


/**
* Writes pre-formatted values to Google Sheets
* Handles one or more "write-directive objects", each a container w/ details for one G Sheet write
* Inserts the sheet if not found
* Performs Stack Driver time logging
* 
* @param {array} vPack - Array of container Objects with write-directive variables inside.
* - @param {object}
* - - @param {string}  functionName: Name of the function that initiated the script. Used for Stack Driver endTime logging
* - - @param {boolean} append: True if data should be appended to bottom of existing data, False if existing data should be erased
* - - @param {array of array} dataToWrite: 2D array with values to write to destination
* - - @param {boolean} sheetWrite: True if data should be written to Google Sheets [DEPRECATED]
* - - @param {string}  sheetId: Google Sheets ID of destination spreadsheet
* - - @param {string}  sheetName: Name of destination sheet within spreadsheet
* - - @param {boolean} sqlWrite: True if data should be written to Cloud SQL afterwards [DEPRECATED]
* - - @param {string}  sqlTable: Name of destination Cloud SQL table [DEPRECATED]
*/
function write2GSheets(vPack){
  console.time(vPack[0].functionName+' [write2GSheets(): Write to sheets - time]');
  
  for(var count=0; count<vPack.length; count++){
    
    var ss = SpreadsheetApp.openById(vPack[count].sheetId);
    SpreadsheetApp.setActiveSpreadsheet(ss);
    
    try{
      ss.setActiveSheet(ss.getSheetByName(vPack[count].sheetName));
    }catch(e){
      ss.insertSheet(vPack[count].sheetName);
    }
    
    var activeSheet = ss.getSheetByName(vPack[count].sheetName);
    var row = getFirstEmptyRowWholeRow(activeSheet);
    
    if(vPack[count].append === true){
      if(vPack[count].dataToWrite.length>1){
        vPack[count].dataToWrite.shift();
      }
      activeSheet.insertRowsAfter(row - 1, vPack[count].dataToWrite.length);
      
    }else{
      activeSheet.clearContents();
      row = 1;
    }
    
    activeSheet.getRange(row, 1, vPack[count].dataToWrite.length, vPack[count].dataToWrite[0].length).setValues(vPack[count].dataToWrite);
    
    //formatSheet(activeSheet);
  }
  
  console.timeEnd(vPack[0].functionName+' [write2GSheets(): Write to sheets - time]');
  
  if(vPack[0].sqlWrite === true){
    write2Sql(vPack);
  }else{
    sheetsPostWrite(vPack);
  }
}


/**  
   *  Writes pre-formatted data to a Cloud SQL table.
   * 1. Option for appending rows, or clean write which uses TRUNCATE command
   * TRUNCATE deletes tables and recreates, which is significantly faster than deleting rows,
   * but may not be desired in all situations
   * 2. Column names are determined from table meta
   *
   * @param {array} vPack
   * - @param {object} data
   * - - @param {string} functionName: Name of the function that initiated the script. Used for Stack Driver endTime logging
   * - - @param {boolean} append: True if data should be appended to bottom of existing data, False if existing data should be erased
   * - - @param {array of array} dataToWrite: 2D array with values to write to destination
   * - - @param {boolean} sheetWrite: True if data should be written to Google Sheets
   * - - @param {string} sheetId: Google Sheets ID of destination spreadsheet
   * - - @param {string} sheetName: Name of destination sheet within spreadsheet
   * - - @param {boolean} sqlWrite: True if data should be written to Cloud SQL afterwards
   * - - @param {string} sqlTable: Name of destination Cloud SQL table
   *
   */

function write2Sql(vPack){
  console.time(vPack[0].functionName+' [write2Sql(): SQL prepare inserts - time]');
  
  for(var count=0; count<vPack.length; count++){
    if(vPack[count].sqlWrite === true){
      var conn = Jdbc.getCloudSqlConnection(
        'jdbc:google:mysql://data-warehouse-197302:us-west1:beetstech-sql-instance/Beetstech',
        'cody',
        'yb>JFux/C8K76=bFzdVY'
      );
      conn.setAutoCommit(false);
      var stmt = conn.createStatement();
      
      if(vPack[count].append === false){
        var sql = "TRUNCATE "+vPack[count].sqlTable;
        stmt.executeUpdate(sql);
        conn.commit();
        // Remove header row, because it shouldn't have been removed at this point
        vPack[count].dataToWrite.shift();
      }
      
      var stmtVars = prepareStatementVars(vPack[count], stmt);
      var stmt = conn.prepareStatement('INSERT INTO '+vPack[count].sqlTable+' ('+stmtVars.sqlTableCols+') VALUES ('+stmtVars.sqlColVars+')');
      
      for(var i=0; i<vPack[count].dataToWrite.length; i++){
        for(var j=0; j<vPack[count].dataToWrite[i].length; j++){
          // If value is empty string, write null. It's an easy way to accomodate different data type restrictions
          var value = vPack[count].dataToWrite[i][j] === "" ? null : vPack[count].dataToWrite[i][j];
          stmt.setObject((j+1), value);
        }
        stmt.addBatch();
      }
      
      console.timeEnd(vPack[count].functionName+' [write2Sql(): SQL prepare inserts - time]');
      console.time(vPack[count].functionName+' [write2Sql(): SQL execute inserts - time]');
      
      var batch = stmt.executeBatch();
      conn.commit();
      conn.close();
      
      console.timeEnd(vPack[count].functionName+' [write2Sql(): SQL execute inserts - time]');
    }
  }
}


/**
 * Inserts an object into Cache Services.
 * - Can insert nearly any size object by converting to string,
 *   zipping the string, base64 encoding the zip, and splitting
 *   the base64 blob.
 * - Uses LockService to prevent changes to data from other users.
 *
 * @param {String}   key     - Identifying name for the property keys, and used for later retrieval
 * @param {String}   cache   - Name of the Cache Service to use ("Document", "Script", "User")
 * @param {Object}   obj     - The object to be inserted
 * @param {Integer}  minutes - Number of minutes until cache expiration. Max is 360 (21600 secs) (6 hours).
 *
 */

function cacheManager(key, obj, minutes){
  
  var lock = LockService.getScriptLock(); 
  lock.tryLock(3000); // get the lock, error otherwise  
  
  if(!obj){ // read
    // lock needed to ensure reading latest data    
    var out = uncacheData_(key);
    lock.releaseLock(); // don't need the lock  
    return out;
  }
  else{// write
    cacheData_(key, obj, minutes);
    lock.releaseLock();
  }
}  


/**
 * The function that actually transforms and caches the data.
 * Trigger by running cacheManager().
 *
 * @param {String}   key     - Identifying name for the property keys, and used for later retrieval
 * @param {String}   cache   - Name of the Cache Service to use ("Document", "Script", "User")
 * @param {Object}   obj     - The object to be inserted
 * @param {Integer}  minutes - Number of minutes until cache expiration. Max is 360 (21600 secs) (6 hours).
 *
 */

function cacheData_(key, obj, minutes){
  if(!minutes){
    var minutes = 3600;
  }else{
    minutes = minutes * 60;
  }
  var json = JSON.stringify(obj);
  
  //cache as JSON if small enough
  if(json.length < 1E5){
    CacheService.getScriptCache().put(key, json, minutes);
  }
  else{// try and cache as zip
    var blob = Utilities.newBlob(json,'string/json');
    var zip = Utilities.zip([blob]);
    
    var encoded = Utilities.base64Encode(zip.getBytes());
    if(encoded.length < 1E5){ // 100kb    
      CacheService.getScriptCache().put(key, encoded, minutes);
    }
    else{      
      // create checksum
      var md5 = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,encoded,Utilities.Charset.US_ASCII));
      
      // cache large data sets
      var split = {md5:md5, keys:[]};
            
      // Generate uids
      var numOfKeys = encoded.length / 1E5;
      Logger.log(numOfKeys);
      for(var i=0; i<numOfKeys; i++){
        split.keys.push(Utilities.getUuid());
      }
      // Cache index first so it expires first
      CacheService.getScriptCache().put('multipartCache-'+key, JSON.stringify(split), minutes);
      
      // Cache substrings
      for(var i=0; i<split.keys.length; i++){
        var encodedSection = encoded.slice(i*1E5, (i+1)*1E5);
        CacheService.getScriptCache().put(split.keys[i], encodedSection, minutes+1);
      }     
      
    }
  }
}


/**
 * Retrieves an object into Cache Services.
 * Reverses all procedures performed for cache insertion,
 * and returns the original object.
 *
 * @param {String}   key - Identifying name used when caching the object.
 * @return {Object}      - The original object that was cached.
 *
 */

function uncacheData_(key){  
  var encoded = CacheService.getScriptCache().get(key);
  if(!encoded){
    
    // check for multipart
    var encoded = CacheService.getScriptCache().get('multipartCache-'+key);
    if(encoded){
      var encoded = JSON.parse(encoded); // decode data
      var assemble = [];
  
      for(var i=0; i<encoded.keys.length; i++){        
        assemble.push(CacheService.getScriptCache().get(encoded.keys[i]));      
      }
      var assembleString = assemble.join('');
      
      // get md5 checksum for assembled data string
      var md5 = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, assembleString, Utilities.Charset.US_ASCII));
      
      if(encoded.md5 == md5){ 
        return deCacheZip(assembleString);
      }
      else{
        return null;
      }
    }
    else{ // no data
      return null;
    }
  }
  else{ // data exists
    try{ //is this plain JSON?      
      var decoded = JSON.parse(encoded);   
      return decoded;
      }
    catch(e){ // try to decode/ unzip and parse
      return deCacheZip(encoded);
    }
  }
  
  //-----------------------------------------------------------------------------------------------------------------
  
  function deCacheZip(encoded){
    var zip = Utilities.base64Decode(encoded);
    var blob = Utilities.newBlob(zip).setContentType('application/zip');
    var unzip = Utilities.unzip(blob);
    
    var string = unzip[0].getDataAsString();
    var decoded = JSON.parse(string);
    return decoded;
  }
}


/**
 * Inserts an object into PropertiesService
 *
 * @param {String}  prop - Name of the PropertiesService to use ("Document", "Script", "User")
 * @param {Object}  obj  - The object to be inserted
 * @param {String}  name - The name of the property keys, and the name used for later retrieval
 *
 */

function putInProps(prop, obj, name){
  const chunkSize = 4999;
  var chunks = chunkSubstr(JSON.stringify(obj), chunkSize);
  
  switch(prop){
    case 'Document':
      var props = PropertiesService.getDocumentProperties();
      break;
    case 'Script':
      var props = PropertiesService.getScriptProperties();
      break;
    case 'User':
      var props = PropertiesService.getUserProperties();
      break;
  }
  
  for(var i=0; i<chunks.length; i++){
    props.setProperty(name+'-CHUNK_'+(i+1), chunks[i]);
  }
  props.setProperty(name+'-CHUNK_INDEX', chunks.length);
}


/**
 * Creates an array of strings w/ a specified length from a longer string
 *
 * @param {String} str - The string to be "chunked"
 * @param {Integer} size - The desired chunk size
 * @return {Array} chunks - The array of strings
 *
 */

function chunkSubstr(str, size){
  const numChunks = Math.ceil(str.length / size);
  const chunks = new Array(numChunks);

  for(var i=0, o=0; i<numChunks; ++i, o += size){
    chunks[i] = str.substr(o, size);
  }

  return chunks;
}


/**
* Compares 2D array on a single column (which should have unique values), and returns 2D array free of duplicates
* 
* @param {array of array} possibleDups
* @param {object} unqCriteria
* - @param {string} format - Possible options are "Array", "ArrayOfArray", "ArrOfObj"
* - @param {integer} unqCol - (OPTIONAL) Column index if unique values are in an array.
* - @param {string} unqProp - (OPTIONAL) Property key if unique values are in an object.
* @return {Array} noDups
*/

function deDuplicateArr(possibleDups, unqCriteria){
  var noDups = [],
      tempObj = {};
  
  // Add rows to tempObj, each time checking if the unique value already exists, and only adding if it does not
  for(var i=0; i<possibleDups.length; i++){
    if(!tempObj.hasOwnProperty( possibleDups[i][unqCriteria.unqProp] )){
      tempObj[possibleDups[i][unqCriteria.unqProp]] = possibleDups[i];
    }
  }
  
  for(var id in tempObj){
    noDups.push(tempObj[id]);
  }
  
  return noDups;
}


/**
* Read data from Cloud SQL table
*
* @param   {Object} sqlReadPackage - Container with string variables.
* - @param {Integer} maxRows - Number of rows after which query is limited.
* - @param {String} query - The MySQL query to execute.
* - @param {Boolean} - headerRow - Whether or not to use the generated column names (true) or provide own later (false).
* @return  {Array} tableArr - 2D array of values, indexed by row, then by column.
* 
*/

function readTable(sqlReadPackage){
  var conn = Jdbc.getCloudSqlConnection(
    'jdbc:google:mysql://data-warehouse-197302:us-west1:beetstech-sql-instance/Beetstech',
    'cody',
    'yb>JFux/C8K76=bFzdVY'
  );
  var stmt = conn.createStatement();
  stmt.setMaxRows(sqlReadPackage.maxRows);
  var results = stmt.executeQuery(sqlReadPackage.query);
  var numCols = results.getMetaData().getColumnCount(),
      colNameArr = [],
      tableArr  = [];
  
  for(var col=0; col<numCols; col++){
    colNameArr.push(results.getMetaData().getColumnName(col+1));      
  }
  
  while(results.next()){
    var rowArr = [],
        rowString = '';
    
    for(var col=0; col<numCols; col++){
      rowString += results.getString(col+1)+'\t';
      rowArr.push(results.getString(col+1));
    }
    tableArr.push(rowArr);
  }
  
  if(sqlReadPackage.headerRow === true){
    tableArr.unshift(colNameArr); 
  }
  
  results.close();
  stmt.close();
  
  return tableArr;
}


/**
* Deletes all script, document, and user properties
* 
*/

function clearPropServ(){
  Logger.log('Delete Local Script Properties:')  
  if(PropertiesService.getScriptProperties() !== null){
    Logger.log(PropertiesService.getScriptProperties().deleteAllProperties())
    Logger.log('  Deleted')
  }else{
    Logger.log('  None')
  }

  Logger.log('Delete Local NPT Doc Properties:')  
  if(PropertiesService.getDocumentProperties() !== null){
    Logger.log(PropertiesService.getDocumentProperties().deleteAllProperties())
    Logger.log('  Deleted')    
  }else{
    Logger.log('  None')
  }
  
  Logger.log('Delete Local NPT User Properties:')    
  if(PropertiesService.getUserProperties() !== null){  
    Logger.log(PropertiesService.getUserProperties().deleteAllProperties())
    Logger.log('  Deleted')    
  }else{
    Logger.log('  None')
  }
}


/**
* Sets the new execution time as the 'runtimeCount' script property.
*
* @param {Date} runtimeCountStart - Date value created with "new Date()"
* 
*/
	
function runtimeCountStop(start){

  var props = PropertiesService.getScriptProperties();
  var currentRuntime = props.getProperty("runtimeCount");
  var stop = new Date();
  var newRuntime = Number(stop) - Number(start)+Number(currentRuntime);
  var setRuntime = {
    runtimeCount: newRuntime,
  }
  props.setProperties(setRuntime);

}

/**
* Records the project's total execution time in a sheet.
* Set a daily time-based trigger for this function.
* After being recorded in the sheet, the script property is reset.
* 
*/

function recordRuntime(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Runtime";
  try{
    ss.setActiveSheet(ss.getSheetByName("Runtime"));
  } catch (e){
    ss.insertSheet(sheetName);
  }
  var sheet = ss.getSheetByName("Runtime");
  var props = PropertiesService.getScriptProperties();
  var runtimeCount = props.getProperty("runtimeCount");
  var recordTime = new Date();

  sheet.appendRow([recordTime, runtimeCount]);
  props.deleteProperty("runtimeCount");

}


