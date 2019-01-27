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
*  @version 5
*  @since   2012-11-12
*/

/**
*  Show the configuration screen
*/

var old = false;

function doGet() {
  if (old) {
    return oldGet();
  }
  webInterface = true;
  return HtmlService
  .createTemplateFromFile('Page')
  .evaluate()
  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getData() {
  getUserProperties_()
  return {
    values: cfg,
    nbk_label: getLabelOptions(false, true),
    tag_label: getLabelOptions(true, true),
    sent_label: getLabelOptions(true, false),
  }
}

/**
*  Read the gmail labels
*  - gmailAllLabels contains all labels
*  - gmailTopLabels contains toplevel labels only
*/

var gmailAllLabels, gmailTopLabels;

function getLabels() {
  gmailAllLabels = [];
  gmailTopLabels = [];
  var labels = GmailApp.getUserLabels();
  for( var i =0;  i < labels.length; i ++ ) {
    var name = labels[i].getName()
    gmailAllLabels.push( name);
    if( !name.match( /\// )) {
      gmailTopLabels.push( name);
    }
  }
  gmailAllLabels.sort();
  gmailTopLabels.sort();
}

/**
*  Create a list with options from gmail labels
*  @param {boolean} canEmpty, allow empty label, no label selected
*  @param {boolean} topOnly, show only toplevel labels
*/

function getLabelOptions(canEmpty, topOnly) {
  if( gmailAllLabels == undefined ) {
    getLabels();
  }
  var result = [];
  if (canEmpty) {
    result.push('');
  }
  var idx = 0;
  var gmailLabels = topOnly? gmailTopLabels : gmailAllLabels;
  for( var i =0;  i < gmailLabels.length; i ++ ) {
    result.push(gmailLabels[i]);
    idx ++;
  }
  return result;
}

/**
*  Handle response from configuration form
*  @param {object} data, where data is the list of variables
*/ 

function update(data) {
  setTrigger( data.trigger );
  
  var changes = [];
  getUserProperties_();

  if( data.evernoteMail != evernoteMail ) {
    changes.push( 'evernoteMail' );
    UserProperties.setProperty('evernoteMail', data.evernoteMail );
  }
  for( var p in cfg ) {
    if( data[p] != undefined && data[p] != cfg[p] ) {
      changes.push( p );
      UserProperties.setProperty('gm2en_' + p, data[p] );
    }
  }

  var result = {
    trigger: data.trigger,
    changes: changes,
    log: data.log,
  }
  return result;
}

/**
*  Set trigger to run readmail function
*  @param {integer} freq, interval in minutes
*/ 

function setTrigger( freq ) {
  if( freq == undefined ) {
    freq = 15;
  }
  // Delete existing timed-based triggers for readEvernote
  var ENfunction = 'readEvernote';
  var triggers = ScriptApp.getProjectTriggers();
  for( var i=0; i < triggers.length; i ++ ) {
    if( triggers[i].getHandlerFunction() == ENfunction && triggers[i].getEventType() == ScriptApp.EventType.CLOCK ) {
      ScriptApp.deleteTrigger(triggers[0]);
    }
  }
  if( freq > 0 ) {
    ScriptApp.newTrigger( ENfunction )
    .timeBased()
    .everyMinutes(freq)
    .create();
  }
}

