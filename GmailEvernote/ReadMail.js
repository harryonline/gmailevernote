/**
* ---GMail to Evernote---
*
*  Copyright (c) 2012,2013,2018 Harry Oosterveen
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
*  @version 8
*  @since   2012-11-12
*/

/**
*  Find threads labeled with Evernote or sublabel
*  Forward last message to evernote mail address and removes Evernote label or sublabel
*  See http://www.harryonline.net/evernote/send-google-mail-to-evernote/226
*  See https://support.evernote.com/ics/support/KBAnswer.asp?questionID=547
*/

var version = 8;
// Variables for user properties
var evernoteMail, cfg={};
// Sheet for log entries
var logSheet;
// Email of user
var userMail = Session.getActiveUser().getEmail();
var webInterface = false;

function readEvernote()
{
  getUserProperties_();
  checkVersion_();

  // Field that can be copied to Evernote
  var hdrFields = { bcc:'Bcc', cc:'Cc', date:'Date', from:'From', replyto:'ReplyTo', subject:'Subject', to:'To' };
  
  try {  
    // Catch errors, usually connecting to Google mail
    
    // Label for messages sent to Evernote
    if( cfg.sent_label != '' ) {
      var sentLabel = GmailApp.getUserLabelByName(cfg.sent_label );
    }

    // Find GMail userlabels and filter those that start with 'Evernote'
    var labels = GmailApp.getUserLabels();
    for( var i=0; i < labels.length; i ++ ) {
      var labelPath = labels[i].getName().split('/');
      if( labelPath[0] == cfg.nbk_label && labels[i].getName() != cfg.sent_label ) {
        // Get threads
        var threads = labels[i].getThreads();
        for( var j=0;  j < threads.length;j ++ ) {
          var subject = threads[j].getFirstMessageSubject();
          // Read last mail
          var lastMsg = threads[j].getMessages().pop();
          if( lastMsg.getTo() != evernoteMail ) {
            if( labelPath.length > 1 ) {
              // Add label name (last element) as inbox
              subject += ' @' + labelPath[labelPath.length-1];
            }
            if( cfg.tag != '' ) {
              subject += ' #' + cfg.tag;
            }
            // Read tags from mail
            var msgLabels = threads[j].getLabels();
            for( var k=0; k < msgLabels.length; k ++ ) {
              var msgLabelPath = msgLabels[k].getName().split('/');
              if( msgLabelPath[0] != cfg.nbk_label && 
                 msgLabels[k].getName() != cfg.sent_label && 
                ( cfg.tag_label == '' || msgLabelPath[0] == cfg.tag_label ) ) {
                  subject += ' #' + msgLabelPath.pop();
                }
            }
            // Put header info and link(s) before message body
            var fields = cfg.fields.toLowerCase().split(/[, ]+/);
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
            if( cfg.nacct > 0 ) {
              header += createLink_(mailUrl_( msgId, 0 )) ;
              for( var u=1; u < cfg.nacct; u ++ ) {
                header += ' - ' + createLink_( mailUrl_( msgId, u ), 'user '+ u );
              }
              header += sprintf_( " (%s)\n", userMail );
            }
            if( header != '' ) {
              header = sprintf_( '<div style="%s">%s</div>', cfg.hdrcss, header.replace( /\n/g, "<br/>" ));
            }
            var body = header + getBody(lastMsg)
            GmailApp.sendEmail(evernoteMail, subject, '', {htmlBody:body, attachments:lastMsg.getAttachments() })
            log_( subject );
            if( sentLabel != undefined ) {
              threads[j].addLabel(sentLabel);
            }
            if( cfg.keep_sent == 'off' ) {
              // Remove message just sent to Evernote
              var sentResults = GmailApp.search("label:sent " + subject, 0, 1);
              if( sentResults.length > 0 ) {
                var sentMsg = sentResults[0].getMessages().pop();
                if( sentMsg.getTo() == evernoteMail ) {
                  sentMsg.moveToTrash();
                }
              }
            }
          }
          // Remove label from thread
          threads[j].removeLabel(labels[i]);
          Utilities.sleep(1000);
        }
      }
    }
  } catch( e ) {
    Logger.log( '%s: %s %s', e.name, e.message, e.stack );
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
  if( cfg.version == 0 ) {
    // Spread version checking over the day to avoid all users checking at the same time
    var nextCheck = curTime - Math.floor( Math.random() * mSecDay );
    var userVersion = { nextCheck:nextCheck, version:0 };
  } else {
    var userVersion = JSON.parse( cfg.version );
  }
  
  if( curTime > userVersion.nextCheck ) {
    // Read latest version from Internet
    try {
      var response = UrlFetchApp.fetch("http://gs.harryonline.net/gm2en.json");
      var newVersion = JSON.parse( response.getContentText() );
      if( userVersion.version < newVersion.version  ) {
        // User has not been informed of new version yet
        var message = [];
        if( userVersion.version < Math.floor( newVersion.version )) {
          message.push( "A new version of the Gmail to Evernote script has been released." );
        }
        var doSend = false;
        if( newVersion.info != "" ) {
          // Relevant information for all users
          message.push( newVersion.info );
          doSend = true;
        }
        if( version < Math.floor( newVersion.version )) {
          // Old version, user running own copy, inform only on updates; fractions are used to send announcements
          message.push( "You can find the latest version at http://bit.ly/gmailevernote or https://github.com/harryonline/gmailevernote." );
          doSend = true;
        }
        if( doSend ) {
          sendEmail_( 'update', message.join( "\n\n" ));
        }
        userVersion.version = newVersion.version;
      }
    } catch(e) {
      Logger.log( '%s: %s %s', e.name, e.message, e.stack );
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
  
  evernoteMail = userProps['evernoteMail'];
  if( evernoteMail == undefined ) {
    findEvernoteMail_();
  }
  
  // Global variables and default settings
  // Read from user properties, otherwise use defaults
  
  var props = {
	'tag' : '',
	'log' : '',
	'nacct' : 1,
	'fields' : 'From, To, Cc, Date',
	'hdrcss' : 'border-bottom:1px solid #ccc;padding-bottom:1em;margin-bottom:1em;',
	'version' : 0,
	'nbk_label' : 'Evernote',
	'tag_label' : '',
	'sent_label' : '',
    'keep_sent' : 'off',
    'version' : 0
  };
  
  for( var p in props ) {
    var key = 'gm2en_' + p;
    if( userProps[key] == undefined ) {
      switch( p ) {
        case 'log' :
          // Create new log sheet
          cfg[p] = createLogSheet_();
          break;
         
        default:
          // Use default value
          cfg[p] = props[p];
      }
      UserProperties.setProperty(key, cfg[p]);
    } else {
      cfg[p] = userProps[key];
    }
  }
  cfg.evernoteMail = evernoteMail;
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
        sendEmail_( 'evernoteMail changed', 'The Gmail to Evernote script will forward messages to ' + evernoteMail );
        return;
      }
    }
  }
  evernoteMail = '';
  if( !webInterface ) {
    sendEmail_( 'evernoteMail not found', 'The Gmail to Evernote script could not find an Evernote email address. '  
               + 'Please send a message to your Evernote email address and run the script again.' );
  }
}

/**
*  Add an entry to the logsheet
*  @param {string} subject
*
*/
function log_( subject )
{
  if( cfg.log != '' ) {
    try {
      if( logSheet == undefined ) {
        logSheetParts = cfg.log.split( ':' );
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
    } catch(e) {
      Logger.log( '%s at line %s: %s', e.name, e.lineNumber, e.message );
    }
  } else {
    Logger.log( subject );
  }
}

/**
*  Create a new logsheet if it does not exist
*/
function createLogSheet_()
{
  var ss = SpreadsheetApp.create("Gmail to Evernote Log");
  var log = ss.getId();
  sendEmail_( 'logsheet created', 
             sprintf_( "The Gmail to Evernote script has created a logsheet with ID: %s\nURL: %s" +
                      "\n\nTo change these settings, visit http://bit.ly/gmailevernote, " +
                      "go to File - Project properties, and click on the User properties tab.",
                      log, ss.getUrl() ));
  var sheet = ss.getSheets()[0];
  sheet.appendRow(['Date', 'Source','Message']);
  return log;
}

/**
*  Send an email message to the user
*/

function sendEmail_( subject, body )
{
  if( webInterface ) {
    messages.push( body );
  } else {
    var postText = "\n\nThis is an automated message from the Gmail to Evernote script.\nFor more information, visit http://www.harryonline.net/tag/gm2en .";
    GmailApp.sendEmail(userMail, 'Gmail to Evernote: ' + subject, body + postText, {noReply:true});
  }
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

/**
*  Get message body, in pre-text if plain text
*/
function getBody( message ) {
  var htmlBody = message.getBody()
  var plainBody = message.getPlainBody()
  if( htmlBody !== plainBody ) {
    return htmlBody
  }
  return '<pre style="white-space:pre-wrap">' + plainBody + '</pre>'
}