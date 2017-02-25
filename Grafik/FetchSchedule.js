var fs = require('fs');
var readline = require('readline');
var google = require('googleapis');
var googleAuth = require('google-auth-library');
// If modifying these scopes, delete your previously saved credentials
// at ~/.credentials/gmail-nodejs-quickstart.json
var SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify'];
var TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH || process.env.USERPROFILE) + '/.credentials/';
var TOKEN_PATH = TOKEN_DIR + 'gmail-nodejs-quickstart.json';
// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
    if (err) {
        console.log('Error loading client secret file: ' + err);
        return;
    }
    // Authorize a client with the loaded credentials, then call the
    // Gmail API.
    authorize(JSON.parse(content), getMessage);
});
/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    var clientSecret = credentials.installed.client_secret;
    var clientId = credentials.installed.client_id;
    var redirectUrl = credentials.installed.redirect_uris[0];
    var auth = new googleAuth();
    var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);
    // Check if we have previously stored a token.
    fs.readFile(TOKEN_PATH, function (err, token) {
        if (err) {
            getNewToken(oauth2Client, callback);
        }
        else {
            oauth2Client.credentials = JSON.parse(token);
            callback(oauth2Client);
        }
    });
}
/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
    var authUrl = oauth2Client.generateAuthUrl({
        access_type: 'offline'
        , scope: SCOPES
    });
    console.log('Authorize this app by visiting this url: ', authUrl);
    var rl = readline.createInterface({
        input: process.stdin
        , output: process.stdout
    });
    rl.question('Enter the code from that page here: ', function (code) {
        rl.close();
        oauth2Client.getToken(code, function (err, token) {
            if (err) {
                console.log('Error while trying to retrieve access token', err);
                return;
            }
            oauth2Client.credentials = token;
            storeToken(token);
            callback(oauth2Client);
        });
    });
}
/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
    try {
        fs.mkdirSync(TOKEN_DIR);
    }
    catch (err) {
        if (err.code != 'EEXIST') {
            throw err;
        }
    }
    fs.writeFile(TOKEN_PATH, JSON.stringify(token));
    console.log('Token stored to ' + TOKEN_PATH);
}
/**
 * Lists the labels in the user's account.
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function getMessage(auth) {
    var gmail = google.gmail('v1');
    gmail.users.messages.list({
        auth: auth
        , userId: 'me'
        , q: 'Subject: grafik AND From: WroclawArkady@multikino.pl AND is: unread'
    , }, function (err, response) {
        if (err) {
            console.log('The API returned an error: ' + err);
            return;
        }
        try {
            var messageId = response.messages[0].id
        }
        catch (err) {
            console.log("no new schedules");
            return false;
        }
        gmail.users.messages.get({
            auth: auth
            , id: response.messages[0].id
            , userId: 'me'
        , }, function (err, response) {
            if (err) {
                console.log('The API returned an error: ' + err);
                return;
            }
            gmail.users.messages.attachments.get({
                auth: auth
                , userId: 'me'
                , messageId: messageId
                , id: response.payload.parts[1].body.attachmentId
            , }, function (err, response) {
                if (err) {
                    console.log('The API returned an error: ' + err);
                    return;
                }
                fs.writeFile('Grafik.xlsx', response.data, 'base64', (err) => {
                    if (err) throw err;
                    console.log('It\'s saved!');
                });
                gmail.users.messages.modify({
                    auth: auth
                    , id: messageId
                    , userId: 'me'
                    , resource: {
                        "addLabelIds": []
                        , "removeLabelIds": ["UNREAD"]
                    }
                }, function (err, response) {
                    if (err) {
                        console.log('The API returned an error: ' + err);
                        return;
                    }
                });
            });
        });
    });
    return true;
}
