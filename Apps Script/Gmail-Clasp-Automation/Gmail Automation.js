//import 'google-apps-script';
//This will be used to check in a DOA PreAlert is in the Excel Sheet, and if not, add it to the tracker and send the Trap Request Email
//===================================================Classes====================================================================//
class Part {
    constructor(pn, serialNumber, prt, trapped = true) {
        this.pn = pn;
        this.sn = serialNumber;
        this.prt = prt;
        this.trapped = trapped;
    }
    messageState(state) {
        //const message = this.prt.toString() + state.toString() + 'w/ sn: ' + this.sn.toString() + ' and pn: ' + this.sn.toString();
        const message = `${this.prt} ${state} w/sn: ${this.sn} and pn: ${this.sn}`;
        return message;
    }
    successMessage() {
        const state = 'processed successfully!';
        return this.messageState(state);
    }
    failureMessgae() {
        const state = 'PROCESSED FAILED!!';
        return this.messageState(state);
    }
}
class MailingList {
    constructor(incoming, outgoing) {
        this.incoming = incoming;
        this.outgoing = outgoing;
    }
}
//========================================================Main Function====================================================================//
function processDOAMessages() {
    //------------------------GLobal Constants------------------------------------------------------------------------------//
    // Mailing Lists
    let mailingList = {
        doa: new MailingList(['balvinder.mann@cokeva.com', 'harbans.kaur@cokeva.com', 'lillie.lee@cokeva.com'])
    };
    //Regex Strings
    let regexString = {
        pebbles: new RegExp('I have [0-9]* [A-Z]{3}|PN [0-9\-R]*|SN [A-Z][0-9]*|PR[D|T] [0-9]*', 'g')
    };
    // search strings
    let searchString = {
        doaRecieved: 'is:unread AND label:received-doa',
        doaIncoming: 'is:unread AND label:incoming-doa'
    };
    // Spreadsheet URL's
    let spreadsheetUrl = {
        doa: 'https://docs.google.com/spreadsheets/d/18-e2beGK9JiP2PT3nJXkAe2o-TGwYBOnaKjHz79ZHcA/edit#gid=600901128',
        doaTesting: 'https://docs.google.com/spreadsheets/d/145PzZZ7weiAhcTn2uRlPgcPiDhOWhT6BPa4XIo3CHxQ/edit#gid=600901128'
    };
    pebblesRevieved(searchString, regexString, spreadsheetUrl);
}
//=========================================================Recieved Function==============================================================//
//--------------------------------------------------Process Emails -> parts---------------------------------------------------------------//
// Relevant emails. using Gmail Filters to put the right emails in the bucket, and only processing unread ones.
// Pull out relevant strings. The structure of strings is an array of objects
function pebblesRevieved(searchString, regexString, spreadsheetUrl) {
    let parts;
    try {
        parts = parseRelInfoPebbles(getRelevantMessages(searchString.doaRecieved), regexString.pebbles); //data is structured parts[i].sn || .pn || .prt
    }
    catch (error) {
        console.log(error.message);
        return error.message;
    }
    //---------------------------------------------parts -> spreadsheet + error messages-----------------------------------------------------//
    //The target google sheet
    const ss = SpreadsheetApp.openByUrl(spreadsheetUrl.doaTesting);
    //Process messages
    let partsProcessed = {
        passed: [],
        failed: [],
        addToTrap: [],
    };
    let partsPassed = [];
    let partsFailed = [];
    let addToTrap = [];
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        try {
            IncomingtoOpen(ss, searchSheet(part, ss));
            partsPassed.push(part.successMessage());
            if (part.trapped === false) {
                addToTrap.push(part);
            }
        }
        catch (error) {
            partsFailed.push(`${part.failureMessgae} w/ error ${error.message}`);
        }
    }
    //==========================================================FIX LATER==========================================================================//
    // const heading = ['parts added successfully:','parts failed to be added','parts that need to be added to a serial trap'];
    // partsProcessed.passed = partsPassed;
    // partsProcessed.failed = partsFailed;
    // //partsProcessed.addToTrap = addToTrap;
    // if(addToTrap.length !== 0){partsProcessed.addToTrap = addToTrap;}
    // //-------------------------------------------------------error messages -> email to self----------------------------------------------------//
    // const sendAddress: string = 'kelly.gruber@cokeva.com';
    // const subject: string = 'DOA EMAILS PROCESSED';
    // const plainBody: string = emailBodyCompose(heading, partsProcessed);
    // GmailApp.sendEmail(sendAddress,subject,plainBody);
}
function getRelevantMessages(searchString) {
    const threads = GmailApp.search(searchString); // pull down the relevent emails
    if (threads.length < 1) {
        console.log(new Error('no messages found'));
        return;
    }
    let messages = []; //initialize an empty array
    threads.forEach(function (thread) {
        messages.push((thread.isUnread()) ? thread.getMessages()[0] : null); //pushes the messages into the array if they are unread
        thread.markRead(); // marks them as read
    });
    return messages;
}
function parseRelInfoPebbles(messages, regexString) {
    let parts = []; // Initilize an empty array
    // strings.parts[0].sn <- the expected data structure
    for (let m = 0; m < messages.length; m++) // for each message in the message array...
     {
        const text = messages[m].getPlainBody(); // get the plain text body of the message
        const matches = text.match(regexString); // perform the regex extraction of the relevant text
        if (!matches || matches.length < 3) // if the variable doesn't exsist OR the match is too short, thorugh an error
         {
            //No matches; couldn't parse continue with the next message
            continue;
        }
        const count = parseInt(matches[0].substring(7)); // convert from a string to an int
        for (let i = 0; i < count; i++) {
            const part = new Part(matches[3 * i + 1].substring(3), matches[3 * i + 2].substring(3), matches[3 * i + 3].substring(4));
            parts.push(part);
        }
    }
    return parts;
}
function IncomingtoOpen(ss, incoming) {
    //get the target data range for the selected row
    //var sheet = SpreadsheetApp.getActiveSheet();
    //var row = sheet.getActiveRange().getRow();
    let sheet = ss.getSheetByName('Incoming');
    let data = [];
    if (typeof incoming !== 'object') {
        let row = incoming;
        data = sheet.getRange(row, 1, 1, 11).getValues(); // target aquired
        //delete the row
        sheet.deleteRow(row);
    }
    else {
        let part = incoming;
        data[0] = 'DOA';
        data[1] = part.prt;
        data[2] = part.pn;
        data[3] = part.sn;
        data[10] = Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yyyy").toString() + ' No Pre-Alert Email Sent';
    }
    //change active sheet
    sheet = ss.getSheetByName("Open DOA FA Aging");
    //sheet.activate();
    //add blank row to the bottom of the spreadsheet
    const lastRow = sheet.getLastRow();
    sheet.insertRowAfter(lastRow);
    //Copy in Relevant Data
    const neededData = [1, 2, 3, 4, 11];
    for (let i = 0; i < neededData.length; i++) {
        const neededDataLoop = neededData[i];
        const targetCol = (neededDataLoop == 11) ? 27 : neededDataLoop;
        const targetData = data[0][neededData[i] - 1];
        const targetRow = sheet.getLastRow();
        sheet.getRange(targetRow, targetCol, 1, 1).setValue(targetData);
    }
    //autofill down formulas
    const formulaCells = [5, 6, 7, 29, 30];
    for (let i = 0; i < formulaCells.length; i++) {
        const sourceRange = sheet.getRange(sheet.getLastRow() - 1, formulaCells[i]);
        //var destination = sheet.getRange(sheet.getLastRow(),formulaCells[i]);
        sourceRange.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }
    //set staus to Open
    sheet.getRange(sheet.getLastRow(), 8).setValue("Open");
    //copy formatting
    const source = sheet.getRange(sheet.getLastRow() - 1, 1, 1, 30);
    source.copyFormatToRange(sheet, 1, 30, sheet.getLastRow(), sheet.getLastRow());
    //Set Recieved @ Cokeva to True
    const recieved = sheet.getRange(sheet.getLastRow(), 9, 1, 1);
    recieved.setValue("True");
    const target = sheet.getRange(sheet.getLastRow(), 10);
    target.setValue(Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yyyy"));
    return;
}
// expected input is an part object
function searchSheet(part, ss) {
    const sheet = ss.getSheetByName('Incoming');
    let out;
    const textFinder = sheet.createTextFinder(part.prt);
    const foundRange = textFinder.findAll();
    if (foundRange !== undefined || foundRange.length != 0) { //if the prt is not in the sheet, this is where it would add it to the sheet
        const range = foundRange[0]; // if it's in the sheet, do the following
        const row = range.getRow();
        out = row;
    }
    else {
        part.trapped = false;
        out = part;
    }
    return out;
}
function emailBodyCompose(headings, processed) {
    //composes the success and fail email of the program, so that what is tracked can be used later to verify if it was done
    const sections = Object.keys(processed); //sections are the property names of the object
    let messageArray = [];
    if (headings.length !== sections.length) { //if the length of the headings array and the processed array are of different lengths, through an error
        console.error('size of heading array !== the size of the sections array');
        return;
    }
    for (let i = 0; i < headings.length; i++) { //for eacg heading in the heading section, add a block to the email chain
        messageArray.push(headings[i] + '\\n');
        processed[sections[i]].forEach((message) => // jshint ignore:line 
         {
            messageArray.push(message);
        });
    }
    let body;
    body = messageArray.join('\\n');
    return body; //returns strings
}
function emailToIncoming(ss, incoming, sender) {
    //initilize the correct sheet
    const sheet = ss.getSheetByName('Incoming');
    //structure the data
    let data = [];
    data[0] = 'DOA';
    data[1] = incoming.prt;
    data[2] = incoming.pn;
    data[3] = incoming.sn;
    data[7] = 'INCOMING';
    data[8] = Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yy");
    data[10] = `${data[8]} Pre Alert from ${sender}.`;
    //add blank row to the bottom of the spreadsheet
    sheet.insertRowAfter(sheet.getLastRow());
    //Copy in Relevant Data
    const neededData = [1, 2, 3, 4, 11];
    for (let i = 0; i < neededData.length; i++) {
        const targetCol = neededData[i];
        // const targetCol = neededDataLoop;
        const targetData = data[0][neededData[i] - 1];
        const targetRow = sheet.getLastRow();
        sheet.getRange(targetRow, targetCol, 1, 1).setValue(targetData);
    }
    //autofill down formulas
    const formulaCells = [5, 6, 10];
    for (let i = 0; i < formulaCells.length; i++) {
        const sourceRange = sheet.getRange(sheet.getLastRow() - 1, formulaCells[i]);
        //var destination = sheet.getRange(sheet.getLastRow(),formulaCells[i]);
        sourceRange.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }
    return;
}
