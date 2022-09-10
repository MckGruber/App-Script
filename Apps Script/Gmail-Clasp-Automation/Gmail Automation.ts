//import 'google-apps-script';

//This will be used to check in a DOA PreAlert is in the Excel Sheet, and if not, add it to the tracker and send the Trap Request Email

//===================================================Classes====================================================================//

class Part {
    pn: string;
    sn: string;
    prt: string;
    trapped: boolean;
    message: string;
    constructor(pn: string, serialNumber:string, prt:string, trapped: boolean = true) {
        this.pn = pn;
        this.sn = serialNumber;
        this.prt = prt;
        this.trapped = trapped;
    }

    messageState(state: string) {
        //const message = this.prt.toString() + state.toString() + 'w/ sn: ' + this.sn.toString() + ' and pn: ' + this.sn.toString();
        const message: string = `${this.prt} ${state} w/sn: ${this.sn} and pn: ${this.sn}`; 
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
    incoming: string[];
    outgoing: string[];
    constructor(incoming?:string[], outgoing?:string[]) {
        this.incoming = incoming;
        this.outgoing = outgoing;
    }
}

//========================================================Main Function====================================================================//
function processDOAMessages()
{
    //------------------------GLobal Constants------------------------------------------------------------------------------//


    // Mailing Lists
    let mailingList= {
         doa: new MailingList(['balvinder.mann@cokeva.com','harbans.kaur@cokeva.com','lillie.lee@cokeva.com'])
    }

    //Regex Strings
    let regexString = {
        //pebbles: new RegExp('I have [0-9]* [A-Z]{3}|PN [0-9\-R]*|SN [A-Z][0-9]*|PR[D|T] [0-9]*','g')
        pebbles: ['I have [0-9]* [A-Z]{3}','PN [0-9\-R]*','SN [A-Z][0-9]*','PR[D|T] [0-9]*'],
        klaPurchasing: ['PN [0-9\-]*','PRT: [0-9]*','SN: [XRW][0-9]*']
    }

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

    pebblesRevieved(searchString.doaIncoming, regexString.pebbles, spreadsheetUrl.doa);

}


//=========================================================Recieved Function==============================================================//
//--------------------------------------------------Process Emails -> parts---------------------------------------------------------------//

// Relevant emails. using Gmail Filters to put the right emails in the bucket, and only processing unread ones.
// Pull out relevant strings. The structure of strings is an array of objects


function pebblesRevieved(searchString: string,regexString: string[],spreadsheetUrl: string) {
    let parts: Part[];
    try{
      parts = parseReleventInfo(getRelevantMessages(searchString),regexString); //data is structured parts[i].sn || .pn || .prt
    }
    catch (error) {
      console.log(error.message);
      return error.message;
    }


    //---------------------------------------------parts -> spreadsheet + error messages-----------------------------------------------------//

    //The target google sheet
    const ss = SpreadsheetApp.openByUrl(spreadsheetUrl);
    
    //Process messages
    let partsProcessed = {
        passed: [],
        failed: [],
        addToTrap: [],
    }
    let partsPassed: string[] = [];
    let partsFailed: string[] = [];
    let addToTrap: Part[] = [];
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        try {
            IncomingtoOpen(ss,searchSheet(part,ss));
            partsPassed.push(part.successMessage());
            if (part.trapped === false) {
                addToTrap.push(part);
            }
        } catch (error) {
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



function getRelevantMessages(searchString: string): GoogleAppsScript.Gmail.GmailMessage[]// takes a search string and returns the body of the relevant emails
//searchString needs to be a gmail search string. docs are found here -> https://support.google.com/mail/answer/7190?hl=en
{
    const threads = GmailApp.search(searchString);      // pull down the relevent emails
    if(threads.length < 1){
        console.log(new Error('no messages found'))
      return;
      }
    let messages: GoogleAppsScript.Gmail.GmailMessage[] = [];                                  //initialize an empty array
    threads.forEach(function(thread)                    //for each email in the thread, get just the body of the message and push it onto the array
    {
        messages.push((thread.isUnread()) ? thread.getMessages()[0] : null);        //pushes the messages into the array if they are unread
        thread.markRead();                                                          // marks them as read
    });
    return messages;
}



function parseRelInfoPebbles(messages: GoogleAppsScript.Gmail.GmailMessage[], regexString: RegExp): Part[] // Takes the relevant messages from the getRelevantMessages function and parses the relevant information into an array of strings
//messages => an array of strings
//regexString => a regex string to search with
//returns an array of strings
{
    let parts: Part[] = [];                                    // Initilize an empty array
    
    // strings.parts[0].sn <- the expected data structure

    for(let m = 0; m < messages.length; m++)            // for each message in the message array...
    {
        const text = messages[m].getPlainBody();          // get the plain text body of the message
        const matches = text.match(regexString);          // perform the regex extraction of the relevant text
        if(!matches || matches.length < 3)              // if the variable doesn't exsist OR the match is too short, thorugh an error
        {
            //No matches; couldn't parse continue with the next message
            continue;
        }
        
        const count = parseInt(matches[0].substring(7));  // convert from a string to an int

                
        for (let i = 0; i < count; i++)
        {
            const part = new Part(matches[3*i+1].substring(3),matches[3*i+2].substring(3),matches[3*i+3].substring(4));
            parts.push(part);
        }
    }
    return parts;
}

function parseReleventInfo(messages:GoogleAppsScript.Gmail.GmailMessage[], regexStrings: string[]) {
    let parts: Part[] = [];
    for (let m = 0; m < messages.length; m++) {
        // get plain text of message
        const text: string = messages[m].getPlainBody();
        // determine the count by checking the first regex string, and if it doesn't start with the prefixes of the defined attributes for a part, 
        // use it to match the first number in a Pebbles email. else it's one
        const count = ((regexStrings[0].substring(0,1))!==('PN'||'SN'||'PR') ? parseInt(text.match(new RegExp(regexStrings[0]))[0].substring(7)): 1);
        // for the number of parts in the email, loop through the regex list and pull out the relevant data
        for(let c = 0; c < count; c++) {
            let pn: string[] = [];
            let prt: string[] = [];
            let sn: string[] = [];

            for (let i = 0; i < regexStrings.length; i++) {
                const regexString: string = regexStrings[i];
                const match: string = text.match(new RegExp(regexString, 'g'))[c];
                let start: number;
                switch (regexString.substring(0,1)) {
                    case 'PN':
                        start = (match.substring(2) === ' ') ? 3 : 4;
                        pn.push(match.substring(start));
                        break;

                    case 'SN':
                        start = (match.substring(2) === ' ') ? 3 : 4;
                        sn.push(match.substring(start));
                        break;

                    case 'PR':
                        start = (match.substring(3) === ' ') ? 4 : 5;
                        prt.push(match.substring(start));
                        break;
                
                    default:
                        break;
                }
            }
            parts.push(new Part(pn[c],sn[c],prt[c]));
        }
    }
    return parts;
}

function IncomingtoOpen(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, incoming: Part | number) {
    //get the target data range for the selected row
    //var sheet = SpreadsheetApp.getActiveSheet();
    //var row = sheet.getActiveRange().getRow();
    let sheet = ss.getSheetByName('Incoming');
    let data = [];
    if(typeof incoming !== 'object'){
        let row: number = incoming;
        data = sheet.getRange(row,1,1,11).getValues(); // target aquired
  
        //delete the row
        sheet.deleteRow(row);
    }else{
        let part: Part = incoming;
        data[0] = 'DOA';
        data[1] = part.prt;
        data[2] = part.pn;
        data[3] = part.sn;
        data[10] = Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yyyy").toString()+' No Pre-Alert Email Sent';
    }
    //change active sheet
    sheet = ss.getSheetByName("Open DOA FA Aging");
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
    const target = sheet.getRange(sheet.getLastRow(),10);
    target.setValue(Utilities.formatDate(new Date(), "GMT-07:00", "MM-dd-yyyy"));
    return;
}


// expected input is an part object
function searchSheet(part: Part, ss: GoogleAppsScript.Spreadsheet.Spreadsheet){
    const sheet = ss.getSheetByName('Incoming');
    let out: number | Part; 
    const textFinder = sheet.createTextFinder(part.prt);
    const foundRange = textFinder.findAll();
    if(foundRange !== undefined || foundRange.length != 0){//if the prt is not in the sheet, this is where it would add it to the sheet
        const range: GoogleAppsScript.Spreadsheet.Range = foundRange[0];// if it's in the sheet, do the following
        const row: number = range.getRow();
        out = row;
    }else{
        part.trapped = false;
        out = part;
    }
    return out;
}

function emailBodyCompose(headings: string[], processed: object): string{
    //composes the success and fail email of the program, so that what is tracked can be used later to verify if it was done
    const sections = Object.keys(processed); //sections are the property names of the object
    let messageArray: string[] = [];
    if (headings.length !== sections.length) {//if the length of the headings array and the processed array are of different lengths, through an error
        console.error('size of heading array !== the size of the sections array');
        return
    }
    for (let i = 0; i < headings.length; i++) {//for eacg heading in the heading section, add a block to the email chain
        messageArray.push(headings[i] + '\\n');
        processed[sections[i]].forEach((message: string) =>// jshint ignore:line 
        {    
            messageArray.push(message);
        });
    }
    let body: string;
    body = messageArray.join('\\n');
    return body;//returns strings
}

function emailToIncoming(ss: GoogleAppsScript.Spreadsheet.Spreadsheet, incoming: Part, sender: String) {
    //initilize the correct sheet
    const sheet: GoogleAppsScript.Spreadsheet.Sheet = ss.getSheetByName('Incoming');
    
    //structure the data
    let data: String[] = [];
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
    const neededData = [1,2,3,4,11];
    for (let i = 0; i < neededData.length; i++) {
        const targetCol: number = neededData[i];
        // const targetCol = neededDataLoop;
        const targetData: number | string = data[0][neededData[i] - 1];
        const targetRow: number = sheet.getLastRow();
        sheet.getRange(targetRow,targetCol,1,1).setValue(targetData);
    } 

    //autofill down formulas
    const formulaCells = [5,6,10];
    for (let i = 0; i < formulaCells.length; i++){
        const sourceRange = sheet.getRange(sheet.getLastRow() - 1,formulaCells[i]);
        //var destination = sheet.getRange(sheet.getLastRow(),formulaCells[i]);
        sourceRange.autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    }
    return;
}
