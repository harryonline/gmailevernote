/**
* ---GMail to Evernote---
*
*  Copyright (c) 2012 Harry Oosterveen
*
*  Licensed under the Apache License, Version 2.0 (the "License");
*  you may not use this file except in compliance with the License.
*  You may obtain a copy of the License at
*
*      http://www.apache.org/licenses/LICENSE-2.0
*
*  Unless required by applicable law or agreed to in writing, software
*  distributed under the License is distributed on an "AS IS" BASIS,
*  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
*  See the License for the specific language governing permissions and
*  limitations under the License.
*/

/**
*  @author  Harry Oosterveen <mail@harryonline.net>
*  @version 1.1
*  @since   2012-11-12
*/

/**
*  Find threads labeled with Evernote or sublabel
*  Forward last message to evernote mail address and removes Evernote label or sublabel
*  See http://www.harryonline.net/evernote/send-google-mail-to-evernote/226
*  See https://support.evernote.com/ics/support/KBAnswer.asp?questionID=547
*/

/**
*
*  ChangeLog - After updates, users have to reauthorize!
*  2012-12-04: Bugfix, script now accepts multiple messages for one notebook (thanks, Mike Weiss)
*  2012-11-28: Improved setup by further automating it
*/

/**
*  The script reads User properties for the settings. Go to File - Project properties, and click
*  on the User properties tab to set or review these.
*
*  The following properties are used:
*
*  evernoteMail: your Evernote email address, like [your username].abc123@m.evernote.com
*                If this is not given, the script will search your mail for previous messages to
*                an email address like @m.evernote.com and uses that address
*
*  defaultTag:   (optional) a tag that will be assigned to all forwarded messages. You can use this
*                to recognize notes that were sent by email, or to distinguish between different
*                email accounts from which you forward messages
*
*  logSheet:     the spreadsheet ID where you want to create a log of forwarded messages. If not 
*                given, the script will create a log sheet. You can find the logsheet on Google
*                Drive, it is called 'Gmail to Evernote Log'. To find out more about the spreadsheet 
*                ID, see https://developers.google.com/apps-script/class_spreadsheetapp#openById
*  
*  Set up a trigger to run this function every 15 minutes (more or less often as you like).
*  When saving the trigger, you will be asked to provide authorization and grant access.
*/

var evernoteMail = UserProperties.getProperty('evernoteMail');
var defaultTag = UserProperties.getProperty('defaultTag');
var logSheet = UserProperties.getProperty('logSheet');

var userMail = Session.getEffectiveUser().getEmail();
var mailInfo = "\r\n\r\nTo change these settings, visit http://bit.ly/gmailevernote, go to File - Project properties, and click on the User properties tab." +
  "\r\n\r\nSee also http://www.harryonline.net/evernote/send-google-mail-to-evernote/226\r\n";

function readEvernote() {
  // Create logsheet if no sheet defined
  if( logSheet == undefined ) {
    createLogSheet_();
  }
  // Find evernote mail if no evernoteMail defined
  if( evernoteMail == undefined ) {
    findEvernoteMail_();
  }
  // Find GMail userlabels and filter those that start with 'Evernote'
  var labels = GmailApp.getUserLabels();
  for( var i=0; i < labels.length; i ++ ) {
    var labelPath = labels[i].getName().split('/');
    if( labelPath[0] == 'Evernote' ) {
      // Get threads
      var threads = labels[i].getThreads();
      for( var j=0;  j < threads.length;j ++ ) {
        // Read last mail
        var subject = threads[j].getFirstMessageSubject();
        var lastMsg = threads[j].getMessages().pop();
        if( labelPath.length > 1 ) {
          // Add label name (last element) as inbox
          subject += ' @' + labelPath[labelPath.length-1];
        }
        if( defaultTag != undefined && defaultTag != '' ) {
            subject += ' #' + defaultTag;
        }
        // Read tags from mail
        var msgLabels = threads[j].getLabels();
        for( var k=0; k < msgLabels.length; k ++ ) {
          var msgLabelPath = msgLabels[k].getName().split('/');
          if( msgLabelPath[0] != 'Evernote' ) {
            subject += ' #' + msgLabelPath.pop();
          }
        }
        lastMsg.forward( evernoteMail, {subject:subject} );
        log_( subject );
        // Remove label from thread
        threads[j].removeLabel(labels[i]);
      }
    }
  }
}

function log_(subject)
{
  if( logSheet != '' ) {
    logSheetParts = logSheet.split( ':' );
    var ss = SpreadsheetApp.openById(logSheetParts[0]);
    if( logSheetParts.length > 1 ) {
      var sheet = ss.getSheetByName(logSheetParts[1]);
    } else {
      var sheet = ss.getSheets()[0];
    }
    sheet.appendRow( [new Date(), 
                      'Evernote:' + userMail, 
                      subject ] );
  } else {
    Logger.log( subject );
  }
}

function findEvernoteMail_()
{
  var threads = GmailApp.search('to:@m.evernote.com');
  for( var i=0;  i < threads.length; i ++ ) {
    var messages = threads[i].getMessages();
    for( var j=0; j < messages.length;  j ++ ) {
      var to = messages[j].getTo();
      var mail = to.match( /[\w\.]+@m\.evernote\.com/ );
      if( mail ) {
        evernoteMail = mail[0];
        UserProperties.setProperty('evernoteMail', evernoteMail);
        Logger.log( 'Evernote mail: %s', evernoteMail );
        GmailApp.sendEmail( userMail, 'Gmail to Evernote: evernoteMail changed',
                           'The Gmail to Evernote script will forward messages to ' + evernoteMail + mailInfo,
                           {noReply:true});
        return;
      }
    }
  }
  GmailApp.sendEmail(userMail,'Gmail to Evernote: evernoteMail not found',
                     'The Gmail to Evernote script could not find an Evernote email address. ' 
                     + 'Please send a message to your Evernote email address and try running the script again.',
                     {noReply:true});
  
}

function createLogSheet_()
{
  var ss = SpreadsheetApp.create("Gmail to Evernote Log");
  logSheet = ss.getId();
  UserProperties.setProperty('logSheet', logSheet);
  Logger.log( 'Create log sheet %s', ss.getId());
  GmailApp.sendEmail( userMail, 'Gmail to Evernote: logsheet created',
                     'The Gmail to Evernote script has created a logsheet with ID: ' + logSheet + "\r\nURL: " + ss.getUrl() + mailInfo,
                     {noReply:true});
  var sheet = ss.getSheets()[0];
  sheet.appendRow(['Date', 'Source','Message']);
}
