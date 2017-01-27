// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
var server = require('./server');
var router = require('./router');
var authHelper = require('./authHelper_daemon');
var outlook = require('node-outlook');

var handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/mail'] = getAllMails;
handle['/calendar'] = getAllEvents;

server.start(router.route, handle);

function home(response, request) {
  console.log('Request handler \'home\' was called.');
  // response.writeHead(200, {'Content-Type': 'text/html'});
  // response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
  // response.end();
  // 
  authorize(response, request);
}

var url = require('url');
function authorize(response, request) {
  console.log('Request handler \'authorize\' was called.');
  
  // authHelper.getToken().then(getUserEmail);
    // .then((token) => {
    //   getUserEmail(token);

    // });
  
  authHelper.getToken().then(getAllEvents);
}

const myuser = {email: 'xu@cloudnativeltd.onmicrosoft.com'};
function getUserEmail(token) {
  // Set the API endpoint to use the v2.0 endpoint
  // 
  // 
  outlook.base.setApiEndpoint('https://graph.microsoft.com/v1.0');

  // Set up oData parameters
  var queryParams = {
    '$select': 'DisplayName, mail',
  };

  outlook.base.getUser({user: myuser, token: token.token.access_token, odataParams: queryParams}, function(error, user){
  // outlook.base.getUser({user: myuser, token: token.token.access_token}, function(error, user){
      if (error) {
        // callback(error, null);
        console.log(error);
      } else {
        // callback(null, user.EmailAddress);
        console.log(user);
      }
    });
}

function getAllMails(token) {


  // Set up oData parameters
  var queryParams = {
    '$select': 'Subject,ReceivedDateTime,From,IsRead',
    '$orderby': 'ReceivedDateTime desc',
    // '$top': 10
  };

  // Set the API endpoint to use the v2.0 endpoint
  outlook.base.setApiEndpoint('https://graph.microsoft.com/v1.0');
  // Set the anchor mailbox to the user's SMTP address
  outlook.base.setAnchorMailbox(myuser.email);

  outlook.mail.getMessages({user: myuser, token: token.token.access_token, folderId: 'inbox', odataParams: queryParams},
    function(error, result){
      if (error) {
        console.log('getMessages returned an error: ' + error);
        // response.write('<p>ERROR: ' + error + '</p>');
        // response.end();
      }
      else if (result) {
        console.log('getMessages returned ' + result.value.length + ' messages.');
        console.log(result);
        result.value.forEach(function(message) {
          console.log('  Subject: ' + message.subject);
          var from = message.from ? message.from.emailAddress.name : 'NONE';
          console.log('  From: ' + from);
          console.log('  IsRead: ' + message.isRead);
        });
      }
    });
}

function getAllEvents(token) {


  // Set up oData parameters
    var queryParams = {
      '$select': 'Subject,Start,End,Attendees',
      '$orderby': 'Start/DateTime desc',
      // '$top': 10
    };

  // Set the API endpoint to use the v2.0 endpoint
  outlook.base.setApiEndpoint('https://graph.microsoft.com/v1.0');
  // Set the anchor mailbox to the user's SMTP address
  outlook.base.setAnchorMailbox(myuser.email);

  outlook.calendar.getEvents({user: myuser, token: token.token.access_token, odataParams: queryParams},
    function(error, result){
      if (error) {
        console.log('getEvents returned an error: ' + error);
      } else if (result) {
        console.log('getEvents returned ' + result.value.length + ' events.');
        console.log(result);
        result.value.forEach(function(event) {
          console.log('  Subject: ' + event.subject);
          console.log('  Event dump: ' + JSON.stringify(event));
        });
      }
    });
}

