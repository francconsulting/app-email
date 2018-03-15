const server = require('./server');
const router = require('./router');
const authHelper = require('./authHelper');
const url = require('url');
const microsoftGraph = require("@microsoft/microsoft-graph-client");

const handle = {};
handle['/'] = home;
handle['/authorize'] = authorize;
handle['/mail'] = mail;

server.start(router.route, handle);

function home(response, request) {
  console.log('Request handler \'home\' was called.');
  response.writeHead(200, {'Content-Type': 'text/html'});
  response.write(`<p>Please <a href="${authHelper.getAuthUrl()}">sign in</a> with your Office 365 or Outlook.com account.</p>`);
  response.end();
}


function authorize(response, request) {
  console.log('Request handler \'authorize\' was called.');

  // The authorization code is passed as a query parameter
  const url_parts = url.parse(request.url, true);
  const code = url_parts.query.code;
  console.log(`Code: ${code}`);
  processAuthCode(response, code);
  /*response.writeHead(200, {'Content-Type': 'text/html'});
  response.write(`<p>Received auth code: ${code}</p>`);
  response.end();*/
}


async function processAuthCode(response, code) {
  let token,email;

  try {
    token = await authHelper.getTokenFromCode(code);
  } catch(error){
    console.log('Access token error: ', error.message);
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  try {
    email = await getUserEmail(token.token.access_token);
  } catch(error){
    console.log(`getUserEmail returned an error: ${error}`);
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  const cookies = [`node-tutorial-token=${token.token.access_token};Max-Age=4000`,
                   `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
                   `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`,
                   `node-tutorial-email=${email ? email : ''}';Max-Age=4000`];
  response.setHeader('Set-Cookie', cookies);
  response.writeHead(302, {'Location': 'http://localhost:8000/mail'});
  response.end();
}

async function getUserEmail(token) {
  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  // Get the Graph /Me endpoint to get user email address
  const res = await client
    .api('/me')
    .get();

  // Office 365 users have a mail attribute
  // Outlook.com users do not, instead they have
  // userPrincipalName
  return res.mail ? res.mail : res.userPrincipalName;
}

async function tokenReceived(response, error, token) {
  let token,email;

  try {
    token = await authHelper.getTokenFromCode(code);
  } catch(error){
    console.log('Access token error: ', error.message);
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  try {
    email = await getUserEmail(token.token.access_token);
  } catch(error){
    console.log(`getUserEmail returned an error: ${error}`);
    response.write(`<p>ERROR: ${error}</p>`);
    response.end();
    return;
  }

  const cookies = [`node-tutorial-token=${token.token.access_token};Max-Age=4000`,
                   `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
                   `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`,
                   `node-tutorial-email=${email ? email : ''}';Max-Age=4000`];
  response.setHeader('Set-Cookie', cookies);
  response.writeHead(302, {'Location': 'http://localhost:3000/mail'});
  response.end();
}


function getValueFromCookie(valueName, cookie) {
  if (cookie.includes(valueName)) {
    let start = cookie.indexOf(valueName) + valueName.length + 1;
    let end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}


async function getAccessToken(request, response) {
  const expiration = new Date(parseFloat(getValueFromCookie('node-tutorial-token-expires', request.headers.cookie)));

  if (expiration <= new Date()) {
    // refresh token
    console.log('TOKEN EXPIRED, REFRESHING');
    const refresh_token = getValueFromCookie('node-tutorial-refresh-token', request.headers.cookie);
    const newToken = await authHelper.refreshAccessToken(refresh_token);

    const cookies = [`node-tutorial-token=${token.token.access_token};Max-Age=4000`,
                     `node-tutorial-refresh-token=${token.token.refresh_token};Max-Age=4000`,
                     `node-tutorial-token-expires=${token.token.expires_at.getTime()};Max-Age=4000`];
    response.setHeader('Set-Cookie', cookies);
    return newToken.token.access_token;
  }

  // Return cached token
  return getValueFromCookie('node-tutorial-token', request.headers.cookie);
}


async function mail(response, request) {
  let token;

  try {
    token = await getAccessToken(request, response);
  } catch (error){
    response.writeHead(200, {'Content-Type': 'text/html'});
    response.write('<p> No token found in cookie!</p>');
    response.end();
    return;
  }

  console.log('Token found in cookie: ', token);
  const email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log('Email found in cookie: ', email);

  response.writeHead(200, {'Content-Type': 'text/html'});
  response.write('<div><h1>Your inbox</h1></div>');

  // Create a Graph client
  const client = microsoftGraph.Client.init({
    authProvider: (done) => {
      // Just return the token
      done(null, token);
    }
  });

  try {
    // Get the 10 newest messages
    const res = await client
      .api('/me/mailfolders/inbox/messages/')
      .header('X-AnchorMailbox', email)
      .top(5)
      .select('subject,from,receivedDateTime,isRead')
      .orderby('receivedDateTime DESC')
      .get();

    console.log(`getMessages returned ${res.value.length} messages.`);
    response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
    res.value.forEach(message => {
      console.log('  Subject: ' + message.subject);
      const from = message.from ? message.from.emailAddress.name : 'NONE';
      response.write(`<tr><td>${from}` +
        `</td><td>${message.isRead ? '' : '<b>'} ${message.subject} ${message.isRead ? '' : '</b>'}` +
        `</td><td>${message.receivedDateTime.toString()}
        </td><td>${message.id}
        </td><td>${message['@odata.etag']}</td></tr>`);
    });

    response.write('</table>');
    
    //response.write(JSON.stringify(res))
    response.write("</ br> </ br>");
    for (var key in res){
      response.write(key + ": "+ JSON.stringify(res[key]) + "</br> </br></br> </br></br> </br>");
    }
  } catch (err) {
    console.log(`getMessages returned an error: ${err}`);
    response.write(`<p>ERROR: ${err}</p>`);
  }

  response.end();
}


/*
https://docs.microsoft.com/es-es/outlook/rest/node-tutorial#create-the-app

https://docs.microsoft.com/en-us/outlook/rest/node-tutorial

https://msdn.microsoft.com/en-us/office/office365/api/mail-rest-operations#all-mail-api-operations

https://docs.microsoft.com/en-us/outlook/rest/
*/
//qttybHY246!bmXWZNJ44)*)

//Id. de aplicación
//b00bf3ed-3c3b-451c-9daa-2a7646de1728


