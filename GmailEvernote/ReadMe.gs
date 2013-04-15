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
*  @version 3
*  @since   2012-11-12
*/


/**
*
*  ChangeLog - After updates, users may have to reauthorize
*  To do: update only on major updates
*  2013-04-13: Checkversion: Minor versions (.x) used just for information sharing, major versions (x.) need updating software and possibly reauthorization
*  2013-04-10: Catch gmail server errors and version check server errors
*  2013-04-10: Set up and change settings via web interface
*  2013-04-10: Additional setting: select sent label (none by default, can use any label)
*  2013-04-10: Additional setting: select tag label (none by default, can use any top label)
*  2013-04-10: Additional setting: select notebook label (Evernote by default, can use any other top label)
*  2013-04-09: Bugfix, use of unset variable mailInfo in function findEvernoteMail_
*  2013-04-03: Catch logsheet errors, and proceed without logging
*  2013-04-03: Added custom labels, for notebook, tags, and sent items
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
*  New in version 3:
*  -----------------
*  gm2en_nbk_label: if this label, or a sublabel, is used, it will be used for the selection of a notebook
*
*  gm2en_tag_label: if this label, or a sublabel, is used, it will become a tag in Evernote
*
*  gm2en_sent_label: if this ia an existing label, it will replace the notebook label after sending
*  
*  How to use:
*  -----------
*  Set up a trigger to run this function every 15 minutes (more or less often as you like).
*  When saving the trigger, you will be asked to provide authorization and grant access.
*/
