// Pull Kisi users
function kisi() {
  // Position in sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var kisi_A = SpreadsheetApp.setActiveSheet(ss.getSheetByName("kisi_A"));

  // Clear content except header all the way to "O" column. TODO: make it find cells with content and cleare those.
  kisi_A.getRange("A2:O").clearContent();

  //************************
  // Define main parameters for API call
  var BASE_URL = "https://api.kisi.io/members";
  var auth = Utilities.base64Encode(kisi_cred); // https://api.kisi.io/docs#section/Setup/Authentication
  var headers = {
    Authorization: "Basic " + auth + ", KISI-LOGIN " + kisi_secret,
    //"Authorization" : kisi_secret,
    Accept: "application/json",
  };
  var options = {
    method: "GET",
    headers: headers,
  };

  
  //************************
  // Loop preperations
  const MAX_NUMBER = 500;    // Max offset we can use, Number of records will be this + offset used
  const START_OFFSET = 0;
  const OFFSET = 100;        // Records per call ( max 100 )

  let allResponseData = []; //  google "let" this :)

  
  //************************
  // Start the "for" loop
  // We will loop thru offset and "concat" details
  // () is responsible for "confif of loop
  for (
       var currentOffset = START_OFFSET;        // Start value
       currentOffset <= MAX_NUMBER;             // untill we offser all records we need
       currentOffset = currentOffset + OFFSET   // update current offset  ( may need adjustment by one offset? )
      )
    // {} is the actuall loop happening
  {
    const URL = `${BASE_URL}?limit=${OFFSET}&offset=${currentOffset}`;   // our Live URL that get's updated as we make consecutive requests
    const response = UrlFetchApp.fetch(URL, options);                    // Our actuall API call
    const responseData = JSON.parse(response.getContentText());          // Parse JSON

    allResponseData = allResponseData.concat(responseData);              // concat arrays we get from every call into one
  }
  
  const data = allResponseData;
  
  //************************
  // Assemble User's data
  // This decided where to post. Starts after header.
  var lastRow = Math.max(kisi_A.getRange(2, 1).getLastRow(), 1);
  var index = 0;   // Mark start of process

  // Populate sheet by looping thru records in out list of dictonaries and pulling data we need into correct columns.
  for (
       var i = 0; 
       i < data.length; i++  // Run the loop as long as "i" is less then amount of records in "data"
      ) 
  {
    kisi_A.getRange(index + lastRow + i, 1).setValue(data[i].id);
    kisi_A.getRange(index + lastRow + i, 2).setValue(data[i].user_id);
    kisi_A.getRange(index + lastRow + i, 3).setValue(data[i].name);
    var email = (data[i] && data[i].user && data[i].user.email) || "";
    kisi_A.getRange(index + lastRow + i, 4).setValue(email);
    kisi_A.getRange(index + lastRow + i, 5).setValue(data[i].created_at);
    kisi_A.getRange(index + lastRow + i, 3).setValue(data[i].name);

    //debug >> Full answer
    // kisi_A.getRange(index + lastRow + i, 10).setValue(data);
  }

  kisi_A.sort(1); // sort by column 1
  SpreadsheetApp.flush(); // This actually posts data when it's ready instead of making many changes one at a time.
}
