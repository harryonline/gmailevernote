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
*  @version 2
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
*  ChangeLog - After updates, users may have to reauthorize
*  2012-12-29: Add email header fields to note
*  2012-12-22: Add email URL to note
*  2012-12-05: Added delay to avoid 'Service invoked too many times in a short time:' error
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
*  gm2en_tag:    (optional) a tag that will be assigned to all forwarded messages. You can use this
*                to recognize notes that were sent by email, or to distinguish between different
*                email accounts from which you forward messages
*
*  gm2en_log:    the spreadsheet ID where you want to create a log of forwarded messages. If not 
*                given, the script will create a log sheet. You can find the logsheet on Google
*                Drive, it is called 'Gmail to Evernote Log'. To find out more about the spreadsheet 
*                ID, see https://developers.google.com/apps-script/class_spreadsheetapp#openById
*
*  gm2en_nacct:  if you use multiple accounts, set this to the number of simulanuously used accounts.
*                The script will provide additional links, so there is also a link if the mail is 
*                not on the account of the default user. See
*                http://support.google.com/accounts/bin/answer.py?hl=en&answer=1721977
*                If set to 0, no link to the mail will be added
*  
*  gm2en_fields: Comma-separated list of email header fields to be added to the Evernote note. 
*                The script will provide additional links, so there is also a link if the mail is 
*                not on the account of the default user. See
*                http://support.google.com/accounts/bin/answer.py?hl=en&answer=1721977
*
*  gm2en_hdrcss: CSS style for the header DIV that is added to the note. By default, it has a light-grey
*                bottom border to separate if from the body.
*
*  gm2en_version: Latest version for which user has received notification
*  
*  Set up a trigger to run this function every 15 minutes (more or less often as you like).
*  When saving the trigger, you will be asked to provide authorization and grant access.
*/

var version = 2;
// Variables for user properties
var evernoteMail, gm2en_tag, gm2en_log, gm2en_nacct, gm2en_fields, gm2en_version;
// Sheet for log entries
var logSheet;
// Email of user
var userMail = Session.getEffectiveUser().getEmail();

function readEvernote()
{
  getUserProperties_();
  checkVersion_();

  // Field that can be copied to Evernote
  var hdrFields = { bcc:'Bcc', cc:'Cc', date:'Date', from:'From', replyto:'ReplyTo', subject:'Subject', to:'To' };

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
        if( gm2en_tag != undefined && gm2en_tag != '' ) {
            subject += ' #' + gm2en_tag;
        }
        // Read tags from mail
        var msgLabels = threads[j].getLabels();
        for( var k=0; k < msgLabels.length; k ++ ) {
          var msgLabelPath = msgLabels[k].getName().split('/');
          if( msgLabelPath[0] != 'Evernote' ) {
            subject += ' #' + msgLabelPath.pop();
          }
        }
        // Put header info and link(s) before message body
        var fields = gm2en_fields.toLowerCase().split(/[, ]+/);
        var header = '';
        for( var f=0; f < fields.length;  f ++ ) {
          var fldName = hdrFields[fields[f]];
          if( fldName != undefined ) {
            eval( sprintf_( 'var fldValue = lastMsg.get%s();', fldName));
            if( fldValue != '' ) {
              header += sprintf_( "%s: %s\n", fldName, fldValue );
            }
          }
        }
        header = htmlEncode_( header );

        var msgId = lastMsg.getId();
        if( gm2en_nacct > 0 ) {
          header += createLink_(mailUrl_( msgId, 0 )) ;
          for( var u=1; u < gm2en_nacct; u ++ ) {
            header += ' - ' + createLink_( mailUrl_( msgId, u ), 'user '+ u );
          }
          header += sprintf_( " (%s)\n", userMail );
        }
        if( header != '' ) {
          body = sprintf_( '<div style="%s">%s</div>%s', gm2en_hdrcss, header.replace( /\n/g, "<br/>" ), lastMsg.getBody());
        } else {
          body = lastMsg.getBody();
        }
        GmailApp.sendEmail(evernoteMail, subject, '', {htmlBody:body, attachments:lastMsg.getAttachments() })
        log_( subject );
        // Remove label from thread
        threads[j].removeLabel(labels[i]);
        Utilities.sleep(1000);
      }
    }
  }
}

/**
*  Check version with latest version, and send information email when update is available
*  - version ( global var): currently running version 
*  - userVersion {version, lastCheck}: last version that user has been informed about 
*  - newVersion {version, info}: current version as read from Internet
*/
function checkVersion_()
{
  var mSecDay = 86400000;
  var curDate = new Date();
  var curTime = curDate.getTime();
  if( gm2en_version == undefined ) {
    // Spread version checking over the day to avoid all users checking at the same time
    var nextCheck = curTime - Math.floor( Math.random() * mSecDay );
    var userVersion = { nextCheck:nextCheck, version:0 };
  } else {
    var userVersion = Utilities.jsonParse( gm2en_version );
  }
  if( curTime > userVersion.nextCheck ) {
    // Read latest version from Internet
    var response = UrlFetchApp.fetch("http://gs.harryonline.net/gm2en.json");
    var newVersion = Utilities.jsonParse( response.getContentText() );
    if( userVersion.version < newVersion.version  ) {
      // User has not been informed of new version yet
      var message = "A new version of the Gmail to Evernote script has been released.";
      var doSend = false;
      if( newVersion.info != "" ) {
        // Relevant information for all users
        message += "\n\n" + newVersion.info;
        doSend = true;
      }
      if( version < newVersion.version ) {
        // Old version, user running own copy
        message += "\n\nYou can find the latest version at http://bit.ly/gmailevernote or https://github.com/harryonline/gmailevernote.";
        doSend = true;
      }
      if( doSend ) {
        sendEmail_( 'new version', message );
      }
      userVersion.version = newVersion.version;
    }
    userVersion.nextCheck += mSecDay;
    UserProperties.setProperty('gm2en_version', Utilities.jsonStringify( userVersion ));
  }
}

/**
*  Read user properties, check and create default values if necessary
*/
function getUserProperties_()
{
  var userProps = UserProperties.getProperties();

  // specific user property keys have changed to name with prefix, update these
  var changeKeys = { defaultTag:'gm2en_tag', logSheet:'gm2en_log', nAccounts:'gm2en_nacct' } ;
  for( var oldKey in changeKeys ) {
    var newKey = changeKeys[oldKey];
    if( userProps[oldKey] != undefined && userProps[newKey] == undefined ) {
      userProps[newKey] = userProps[oldKey];
      UserProperties.setProperty( newKey, userProps[newKey] );
      UserProperties.deleteProperty( oldKey );
    }
  }
  
  evernoteMail = userProps['evernoteMail'];
  if( evernoteMail == undefined ) {
    findEvernoteMail_();
  }
  
  gm2en_tag = userProps['gm2en_tag'];
  
  gm2en_log = userProps['gm2en_log'];
  if( gm2en_log == undefined ) {
    createLogSheet_();
  }
  
  gm2en_nacct = userProps['gm2en_nacct'];
  if( gm2en_nacct == undefined ) {
    gm2en_nacct = 1;
    UserProperties.setProperty('gm2en_nacct', 1);
  }
  if( gm2en_nacct > 10 ) {
    gm2en_nacct = 10;
    UserProperties.setProperty('gm2en_nacct', 10 );
  }

  gm2en_fields = userProps['gm2en_fields'];
  if( gm2en_fields == undefined ) {
    gm2en_fields = 'From, To, Cc, Date';
    UserProperties.setProperty('gm2en_fields', gm2en_fields );
  }

  gm2en_hdrcss = userProps['gm2en_hdrcss'];
  if( gm2en_hdrcss == undefined ) {
    gm2en_hdrcss = 'border-bottom:1px solid #ccc;padding-bottom:1em;margin-bottom:1em;';
    UserProperties.setProperty('gm2en_hdrcss', gm2en_hdrcss );
  }
  
  gm2en_version = userProps['gm2en_version'];
}  


/**
*  Find Evernote email address by searching in previous mails
*/
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
        sendEmail_( 'evernoteMail changed', 'The Gmail to Evernote script will forward messages to ' + evernoteMail + mailInfo );
        return;
      }
    }
  }
  sendEmail_( 'evernoteMail not found', 'The Gmail to Evernote script could not find an Evernote email address. ' 
             + 'Please send a message to your Evernote email address and run the script again.' );
  
}

/**
*  Add an entry to the logsheet
*  @param {string} subject
*
*/
function log_( subject )
{
  if( gm2en_log != '' ) {
    if( logSheet == undefined ) {
      logSheetParts = gm2en_log.split( ':' );
      var ss = SpreadsheetApp.openById(logSheetParts[0]);
      if( logSheetParts.length > 1 ) {
        logSheet = ss.getSheetByName(logSheetParts[1]);
      } else {
        logSheet = ss.getSheets()[0];
      }
    }
    logSheet.appendRow( [new Date(),
                      'Evernote:' + userMail, 
                      subject ] );
  } else {
    Logger.log( subject );
  }
}

/**
*  Create a new logsheet if it does not exist
*  Sets the userProperty and varuiable gm2en_log
*/
function createLogSheet_()
{
  var ss = SpreadsheetApp.create("Gmail to Evernote Log");
  gm2en_log = ss.getId();
  UserProperties.setProperty('gm2en_log', gm2en_log);
  sendEmail_( 'logsheet created', 
             sprintf_( "The Gmail to Evernote script has created a logsheet with ID: %s\nURL: %s" +
                      "\n\nTo change these settings, visit http://bit.ly/gmailevernote, " +
                      "go to File - Project properties, and click on the User properties tab.",
                      gm2en_log, ss.getUrl() ));
  var sheet = ss.getSheets()[0];
  sheet.appendRow(['Date', 'Source','Message']);
}

/**
*  Send an email message to the user
*/

function sendEmail_( subject, body )
{
  var postText = "\n\nThis is an automated message from the Gmail to Evernote script.\nFor more information, visit http://www.harryonline.net/tag/gm2en .";
  GmailApp.sendEmail(userMail, 'Gmail to Evernote: ' + subject, body + postText, {noReply:true});
}

/**
*  Create a HTML link
*  @param {string} url 
*  @param {string} text, url will be used if text not given
*  @return {string} completed HTML <a> tag
*/

function createLink_( url, text )
{
  if( text == undefined ) {
    text = url;
  }
  return sprintf_( '<a href="%s">%s</a>', url, text );
}

/**
*  Create a Gmail url
*  @param {string} msgId message ID
*  @param {string} user user no, 0 or higher if multiple sign-in
*  @return {string} url of GMail message
*/
function mailUrl_( msgId, user ) 
{
  return sprintf_( 'https://mail.google.com/mail/u/%s/#inbox/%s', user, msgId );
}

/**
*  Simple alternative for php sprintf-like text replacements
*  Each '%s' in the format is replaced by an additional parameter
*  E.g. sprintf_( '<a href="%s">%s</a>', url, text ) results in '<a href="url">text</a>'
*  @param {string} format text with '%s' as placeholders for extra parameters
*  @return {string} format with '%s' replaced by additional parameters
*/
function sprintf_( format )
{
  for( var i=1; i < arguments.length; i++ ) {
    format = format.replace( /%s/, arguments[i] );
  }
  return format;
}

/**
*  Encode special HTML characters
*  From: http://jsperf.com/htmlencoderegex
*/
function htmlEncode_( html )
{
  return html.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
