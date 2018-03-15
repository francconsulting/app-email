const credentials = {
  client: {
    id: '04867c03-1f42-4357-aac0-7ebd49a788db',
    secret: 'dwYBL59341^ebathTYBD@^;',

    //    dwYBL59341^ebathTYBD@^;

//  04867c03-1f42-4357-aac0-7ebd49a788db
  },
  auth: {
    tokenHost: 'https://login.microsoftonline.com',
    authorizePath: 'common/oauth2/v2.0/authorize',
    tokenPath: 'common/oauth2/v2.0/token'
  }
};
const oauth2 = require('simple-oauth2').create(credentials);

const redirectUri = 'http://localhost:3000/authorize';

// The scopes the app requires
const scopes = [ 'openid',
                'offline_access',
                 'User.Read',
                 'Mail.Read' ];

function getAuthUrl() {
  const returnVal = oauth2.authorizationCode.authorizeURL({
    redirect_uri: redirectUri,
    scope: scopes.join(' ')
  });
  console.log(`Generated auth url: ${returnVal}`);
  return returnVal;
}

exports.getAuthUrl = getAuthUrl;


async function getTokenFromCode(auth_code, callback, response) {
  let result = await oauth2.authorizationCode.getToken({
    code: auth_code,
    redirect_uri: redirectUri,
    scope: scopes.join(' ')
  });

  const token = oauth2.accessToken.create(result);
  console.log('Token created: ', token.token);
  return token;
}

exports.getTokenFromCode = getTokenFromCode;


function refreshAccessToken(refreshToken, callback) {
  return oauth2.accessToken.create({refresh_token: refreshToken}).refresh();
}

exports.refreshAccessToken = refreshAccessToken;