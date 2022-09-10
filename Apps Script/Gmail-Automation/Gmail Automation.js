import 'google-apps-script';

//This will be used to check in a DOA PreAlert is in the Excel Sheet, and if not, add it to the tracker and send the Trap Request Email

function processDOAMessages()
{
    //------------------------GLobal Constants------------------------------------------------------------------------------//


    // Mailing Lists
    const mailingList = {};
    mailingList.doa.incoming = ['balvinder.mann@cokeva.com','harbans.kaur@cokeva.com','lillie.lee@cokeva.com']; //Mailing List
    //mailingList.doa.outgoing = [];

    //Regex Strings
    const regexString = {};
    regexString.pebbles = new RegExp('I have \d \w{3}|PN \S*|SN \S*|PR[T|D] \S*","gm'); 
    // regexString.doa = new RegExp('I have \d \w{3}|PN \S*|SN \S*|PR[T|D] \S*","gm');
    // regexString.rtv = new RegExp('I have \d \w{3}|PN \S*|SN \S*|PR[T|D] \S*","gm');

    // search strings
    const searchString = {};
    searchString.doa = 'is:unread AND label:received-doa';

    // Spreadsheet URL's
    const spreadsheetUrl = {};
    spreadsheetUrl.doa = 'https://docs.google.com/spreadsheets/d/18-e2beGK9JiP2PT3nJXkAe2o-TGwYBOnaKjHz79ZHcA/edit#gid=600901128';
    spreadsheetUrl.doaTesting = 'https://docs.google.com/spreadsheets/d/1DvtLZgbrp27dVH6otw7CS2TZ7pAj0hDJQQ38HnnrwS0/edit#gid=600901128';

    
    
    //========================================================Main Function====================================================================//
    
    
    //--------------------------------------------------Process Emails -> parts---------------------------------------------------------------//
    
    // Relevant emails. using Gmail Filters to put the right emails in the bucket, and only processing unread ones.
    // Pull out relevant strings. The structure of strings is an array of objects
    const parts = parseRelevantInformation(getRelevantMessages(searchString.doa),regexString.pebbles); //data is structured parts[i].sn || .pn || .prt


    //---------------------------------------------parts -> spreadsheet + error messages-----------------------------------------------------//

    //The target google sheet
    const ss = SpreadsheetApp.openByUrl(spreadsheetUrl.doaTesting);
    
    //Process messages
    const partsProcessed = {};
    let partsPassed = [];
    let partsFailed = [];
    let addToTrap = [];
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        try {
            IncomingtoOpen(ss,searchSheet(part));
            partsPassed.push(part.successMessage());
            if (part.trapped === false) {
                addToTrap.push(part);
            }
        } catch (error) {
            partsFailed.push(`${part.failureMessgae} w/ error ${error.message}`);
        }
    }
    const heading = ['parts added successfully:','parts failed to be added','parts that need to be added to a serial trap'];
    partsProcessed.passed = partsPassed;
    partsProcessed.failed = partsFailed;
    //partsProcessed.addToTrap = addToTrap;
    if(addToTrap.length !== 0){partsProcessed.addToTrap = addToTrap;}

    //-------------------------------------------------------error messages -> email to self----------------------------------------------------//

    const sendAddress = 'kelly.gruber@cokeva.com';
    const subject = 'DOA EMAILS PROCESSED';
    const plainBody = emailBodyCompose(heading, partsProcessed);

    GmailApp.sendEmail(sendAddress,subject,plainBody);
}



function getRelevantMessages(searchString)// takes a search string and returns the body of the relevant emails
//searchString needs to be a gmail search string. docs are found here -> https://support.google.com/mail/answer/7190?hl=en
{
    const threads = GmailApp.search(searchString);      // pull down the relevent emails
    let messages = [];                                  //initialize an empty array
    threads.forEach(function(thread)                    //for each email in the thread, get just the body of the message and push it onto the array
    {
        messages.push((thread.isUnread()) ? thread.getMessages()[0] : null);        //pushes the messages into the array if they are unread
        thread.markRead();                                                          // marks them as read
    });
    return messages;
}



function parseRelevantInformation(messages, regexString) // Takes the relevant messages from the getRelevantMessages function and parses the relevant information into an array of strings
//messages => an array of strings
//regexString => a regex string to search with
//returns an array of strings
{
    let parts = [];                                    // Initilize an empty array
    
    // strings.parts[0].sn <- the expected data structure

    for(let m = o; m < messages.length; m++)            // for each message in the message array...
    {
        const text = messages[m].getPlainBody();          // get the plain text body of the message
        const matches = text.match(regexString);          // perform the regex extraction of the relevant text
        if(!matches || matches.length < 3)              // if the variable doesn't exsist OR the match is too short, thorugh an error
        {
            //No matches; couldn't parse continue with the next message
            continue;
        }
        
        const count = parseInt(matches[1].substring(7));  // convert from a string to an int

                
        for (let i = 0; i < count; i++)
        {
            const part = new Part(matches[3*i+1].substring(3),matches[3*i+2].substring(3),matches[3*i+3].substring(4),messages[m]);
            parts.push(part);
        }
    }
    return parts;
}



function IncomingtoOpen(spreadsheet, incoming) {
    //get the target data range for the selected row
    //var sheet = SpreadsheetApp.getActiveSheet();
    //var row = sheet.getActiveRange().getRow();
    let sheet = spreadsheet.getSheetByName('Incoming');
    let data = [];
    if(typeof incoming !== 'object'){
        let row = incoming;
        data = sheet.getRange(row,1,1,11).getValues(); // target aquired
  
        //delete the row
        sheet.deleteRow(row);
    }else{
        let part = incoming;
        data[0] = 'DOA';
        data[1] = part.prt;
        data[2] = part.pn;
        data[3] = part.sn;
        data[10] = Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yyyy").toString()+' No Pre-Alert Email Sent';
    }
    //change active sheet
    sheet = spreadsheet.getSheetByName("Open DOA FA Aging");
    //sheet.activate();
  
    //add blank row to the bottom of the spreadsheet
    const lastRow = sheet.getLastRow();
    sheet.insertRowAfter(lastRow);
  
    //Copy in Relevant Data
    const neededData = [1,2,3,4,11];
    for(let i = 0; i < neededData.length; i++){
      const neededDataLoop = neededData[i];
      const targetCol = (neededDataLoop == 11) ? 27 : neededDataLoop;
      const targetData = data[0][neededData[i] - 1];
      const targetRow = sheet.getLastRow();
      sheet.getRange(targetRow,targetCol,1,1).setValue(targetData);
    }
  
    //autofill down formulas
    const formulaCells = [5,6,7,29,30];
    for (let i = 0; i < formulaCells.length; i++){
      const sourceRange = sheet.getRange(sheet.getLastRow() - 1,formulaCells[i]);
      //var destination = sheet.getRange(sheet.getLastRow(),formulaCells[i]);
      sourceRange.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }
  
    //set staus to Open
    sheet.getRange(sheet.getLastRow(),8).setValue("Open");
  
    //copy formatting
    const source = sheet.getRange(sheet.getLastRow() - 1,1,1,30);
    source.copyFormatToRange(sheet,1,30,sheet.getLastRow(),sheet.getLastRow());
  
    //Set Recieved @ Cokeva to True
    const recieved = sheet.getRange(sheet.getLastRow(),9,1,1);
    recieved.setValue("True"); 
    sheet.getRange(sheet.getLastRow(),10).setValue(e.value ? Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yyyy") : null);
    return;
}


// expected input is an part object
function searchSheet(part){
    let out;
    const textFinder = range.createTextFinder(part.prt);
    const foundRange = textFinder.findAll();
    if(foundRange !== undefined || foundRange.length != 0){//if the prt is not in the sheet, this is where it would add it to the sheet
        const row = ((foundRange.length > 1) ? foundRange[0] : foundRange).getRow();// if it's in the sheet, do the following
        out = row;
    }else{
        part.trapped = false;
        out = part;
    }
    return out;
}

function emailBodyCompose(headings, processed) {
    //composes the success and fail email of the program, so that what is tracked can be used later to verify if it was done
    const sections = Object.keys(processed); //sections are the property names of the object
    let messageArray = [];
    if (headings.length !== sections.length) {//if the length of the headings array and the processed array are of different lengths, through an error
        return console.error('size of heading array !== the size of the sections array');
    }
    for (let i = 0; i < headings.length; i++) {//for eacg heading in the heading section, add a block to the email chain
        messageArray.push(headings[i] + '\\n');
        processed[sections[i]].forEach(part =>// jshint ignore:line 
        {    
            messageArray.push(`prt: ${part.prt} sn: ${part.sn} pn: ${part.pn}`);
        });
    }
    return messageArray.join('\\n');//returns strings
}

class Part {
    constructor(partNumber, serialNumber, prt, message, trapped = true) {
        this.pn = partNumber;
        this.sn = serialNumber;
        this.prt = prt;
        this.message = message;
        this.trapped = trapped;
    }

    #message(state) {// jshint ignore:line
        //const message = this.prt.toString() + state.toString() + 'w/ sn: ' + this.sn.toString() + ' and pn: ' + this.sn.toString();
        const message = `${this.prt} ${state} w/sn: ${this.sn} and pn: ${this.sn}`; 
        return message;
    }

    successMessage() {
        const state = 'processed successfully!';
        return this.#message(state);// jshint ignore:line

    }

    failureMessgae() {
        const state = 'PROCESSED FAILED!!';
        return this.#message(state);// jshint ignore:line
    }
}
