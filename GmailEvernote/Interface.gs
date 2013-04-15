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
*  Show the configuration screen
*/

var app, messages=[];

function doGet() {
  webInterface = true;
  
  app = UiApp.createApplication().setStyleAttribute('background-color', '#e6f0e6;');

  var body = app.createFlowPanel().setStyleAttributes(css.body);
  body.setId('body');
  app.add(body);
  body.add( app.createLabel('Gmail to Evernote').setStyleAttributes(css.header));
  body.add( app.createLabel('The Gmail to Evernote script will send all Gmail messages with a given label to Evernote. For more details, see the website.', true));
  body.add( app.createAnchor('www.harryonline.net/tag/gm2en', 'http://www.harryonline.net/tag/gm2en').setStyleAttributes(css.spaceBelow));
  getUserProperties_();
  // Number of rows is #cfg + evernote + trigger + submit

  var form = app.createFormPanel();
  body.add(form);
  var vp = app.createVerticalPanel();
  form.add(vp);
  
  vp.add( app.createLabel('Before you can complete this form, a notebook label must have been created in Gmail. This can be any label. ' 
                          + 'A message that has this label, or one of its sublabels, will be send to Evernote.', true).setStyleAttributes(css.spaceBelow));
  
  for( var i=0;  i < messages.length; i ++ ) {
    vp.add( app.createLabel( messages[i].replace( "\n", '<br>' ), true ).setStyleAttributes(css.warning));
  }
  
  var vp1 = app.createVerticalPanel().setStyleAttributes(css.formPanel);
  vp.add( vp1 );

  vp1.add( app.createLabel('Please check and complete, then click Submit').setStyleAttributes(css.formHeader));
  
  var evernoteTest = new RegExp( /@m\.evernote\.com$/i );
  var evernoteString = '@m.evernote.com$';
  var mailBox = app.createTextBox();
  var submitButton = app.createSubmitButton('Submit');
  
  if( !evernoteTest.test( evernoteMail )) {
    mailBox.setStyleAttributes(css.invalid);
    submitButton.setEnabled( false );
  }
  
  var validMail = app.createClientHandler()
  .validateMatches(mailBox,evernoteString, 'i')
  .forEventSource().setStyleAttributes(css.valid)
  .forTargets(submitButton).setEnabled(true);

  var invalidMail = app.createClientHandler()
  .validateNotMatches(mailBox,evernoteString,'i')
  .forEventSource().setStyleAttributes(css.invalid)
  .forTargets(submitButton).setEnabled(false);
  
  mailBox.addKeyUpHandler(validMail);
  mailBox.addKeyUpHandler(invalidMail);
  mailBox.addValueChangeHandler(validMail);
  mailBox.addValueChangeHandler(invalidMail);

  vp1.add( createRow_( 'Evernote address','evernoteMail',
                     mailBox.setValue(evernoteMail).setStyleAttributes( css.inputText ),
                     'Your Evernote email address, like [your username].abc123@m.evernote.com' ));

  vp1.add( createRow_( 'Notebook label','nbk_label',
                     getLabelListBox_( cfg.nbk_label, false, true ),
                     'This label, or a sublabel, will be used for the selection of a notebook' ));

  vp1.add( createRow_( 'Trigger every','trigger',
                     app.createTextBox().setValue(15).setStyleAttributes( css.inputNumber ),
                     'Set how often the mail will be checked', 'minutes' ) );

  var showButton = app.createButton('Show advanced options' ).setStyleAttributes(css.spaceBelow);
  var hideButton = app.createButton('Hide advanced options' ).setVisible(false).setStyleAttributes(css.spaceBelow);
  vp1.add( showButton ).add( hideButton );
  
  var vp2 = app.createVerticalPanel().setVisible(false).setId( 'advPanel' );
  vp1.add( vp2 );

  vp2.add( createRow_( 'Default tag','tag',
                     app.createTextBox().setValue(cfg.tag).setStyleAttributes( css.inputText ),
                     'A tag that will be assigned to all forwarded messages' ) );
  
  vp2.add( createRow_( 'Log sheet ID','log',
                     app.createTextBox().setValue(cfg.log).setStyleAttributes( css.inputText ),
                     'ID of the spreadsheet in which a log of forwarded messages is kept' ) );

  vp2.add( createRow_( 'Nr. of accounts','nacct',
                     app.createTextBox().setValue(cfg.nacct).setStyleAttributes( css.inputNumber ),
                     'If you use multiple accounts, set this to the number of simultaneously used accounts' ) );

  vp2.add( createRow_( 'Email fields','fields',
                     app.createTextBox().setValue(cfg.fields).setStyleAttributes( css.inputText ),
                     'Comma-separated list of email header fields, these will be added to the Evernote note' ) );

  vp2.add( createRow_( 'Header CSS','hdrcss',
                     app.createTextBox().setValue(cfg.hdrcss).setStyleAttributes( css.inputText ),
                     'CSS style for the header DIV that is added to the note' ) );

  vp2.add( createRow_( 'Tag label','tag_label',
                     getLabelListBox_( cfg.tag_label, true, true ),
                       'If set, only this label and sublabel will become tags in Evernote' ));

  vp2.add( createRow_( 'Sent label','sent_label',
                       getLabelListBox_( cfg.sent_label, true, false ),
                       'This label will replace the notebook label after sending' ));
  
  vp1.add(submitButton );

  
  var hideHandler = app.createClientHandler()
  .forEventSource().setVisible(false)
  .forTargets(showButton).setVisible(true)
  .forTargets(vp2).setVisible(false)

  var showHandler = app.createClientHandler()
  .forEventSource().setVisible(false)
  .forTargets(hideButton).setVisible(true)
  .forTargets(vp2).setVisible(true)
  
  showButton.addClickHandler(showHandler);
  hideButton.addClickHandler(hideHandler);
  
  return app;
}

/**
*  Create a row for a configuration variable
*  @param {string} label, text before widget
*  @param {string} id, id and name for widget
*  @param {string} widget, for entry of data
*  @param {string} description, title text on label
*  @param {string} extra, optional extra information
*/

function createRow_( label, id, widget, description, extra ) {
  var hp = app.createHorizontalPanel().setStyleAttributes(css.block);
  hp.add( app.createLabel( label ).setTitle( description ).setStyleAttributes(css.label));
  if( id != '' ) {
    widget.setId( id ).setName( id );
  }
  hp.add( widget );
  if( extra !== undefined ) {
    hp.add( app.createLabel( extra ).setStyleAttributes(css.extra));
  }
  return hp;
}

/**
*  Read the gmail labels
*  - gmailAllLabels contains all labels
*  - gmailTopLabels contains toplevel labels only
*/

var gmailAllLabels, gmailTopLabels;

function getLabels_() {
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
}

/**
*  Create a listbox with gmail labels
*  @param {string} name of current label, set as default choice
*  @param {boolean} canEmpty, allow empty label, no label selected
*  @param {boolean} topOnly, show only toplevel labels
*/

function getLabelListBox_( current, canEmpty, topOnly  ) {
  if( gmailAllLabels == undefined ) {
    getLabels_();
  }
  var idx = 0;
  var listBox = app.createListBox().setStyleAttributes(css.listBox);
  if( canEmpty ) {
    listBox.addItem('');
    idx ++;
  }
  var gmailLabels = topOnly? gmailTopLabels : gmailAllLabels;
  for( var i =0;  i < gmailLabels.length; i ++ ) {
    listBox.addItem( gmailLabels[i]);
    if( current == gmailLabels[i] ) {
      listBox.setSelectedIndex(idx);
    }
    idx ++;
  }
  return listBox;
}

/**
*  Handle response from configuration form
*  @param {object} e, where e.parameter is the list of variables
*/ 

function doPost(e) {
  app = UiApp.getActiveApplication();
  var body  = app.getElementById('body');
  
  setTrigger_( e.parameter.trigger );
  
  var changes = [];
  getUserProperties_();
  if( e.parameter.evernoteMail != evernoteMail ) {
    changes.push( 'evernoteMail' );
    UserProperties.setProperty('evernoteMail', e.parameter.evernoteMail );
  }
  for( var p in cfg ) {
    if( e.parameter[p] != undefined && e.parameter[p] != cfg[p] ) {
      changes.push( p );
      UserProperties.setProperty('gm2en_' + p, e.parameter[p] );
    }
  }
  var triggerMsg = sprintf_("Your mail will be checked every %s minutes.<br>",  e.parameter.trigger );
  if( changes.length == 0 ) {
    body.add( app.createHTML( triggerMsg + "No further changes have been made").setStyleAttributes(css.notification));
  }  else {
    body.add( app.createHTML( triggerMsg + 'The following properties have been updated: ' + changes.join( ', ' ), true).setStyleAttributes(css.notification));
  }
  
  body.add( app.createLabel( 'You can view the log sheet at the link below:' ).setStyleAttributes(css.spaceAbove));
  var link = 'https://docs.google.com/spreadsheet/ccc?key=' + e.parameter.log;
  body.add( app.createAnchor( link, true, link ));
  return app;
}

/**
*  Set trigger to run readmail function
*  @param {integer} freq, interval in minutes
*/ 

function setTrigger_( freq ) {
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
  ScriptApp.newTrigger( ENfunction )
   .timeBased()
   .everyMinutes(freq)
   .create();
}
