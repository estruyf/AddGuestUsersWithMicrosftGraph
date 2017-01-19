var adal = require('adal-node');
var request = require('request');

const TENANT = "tenant.onmicrosoft.com";
const GRAPH_URL = "https://graph.microsoft.com";
const CLIENT_ID = "copy and paste your client id";
const CLIENT_SECRET = "create a secret key and add it here";
const GROUP_ID = "enter-the-group-id";

function getToken() {
    return new Promise((resolve, reject) => {
        const authContext = new adal.AuthenticationContext(`https://login.microsoftonline.com/${TENANT}`);
        authContext.acquireTokenWithClientCredentials(GRAPH_URL, CLIENT_ID, CLIENT_SECRET, (err, tokenRes) => {
            if (err) {
                reject(err);
            }
            var accesstoken = tokenRes.accessToken;
            resolve(accesstoken);
        });
    });
}


getToken().then(token => {
    /* INVITE A USER TO YOUR TENANT */
    var options = {
        method: 'POST',
        url: 'https://graph.microsoft.com/beta/invitations',
        headers: {
            'Authorization': 'Bearer ' + token,
            'content-type': 'application/json'
        },
        body: JSON.stringify({
            "invitedUserDisplayName": "display name",
            "invitedUserEmailAddress": "mail@address.be",
            "inviteRedirectUrl": "https://URL-TO-SITE",
            "sendInvitationMessage": false
        })
    };

    request(options, (error, response, body) => {
        if (!error && response.statusCode == 201) {
            var result = JSON.parse(body);
            // Log all the keys and values
            for (var key in result) {
                console.log(`${key}: ${JSON.stringify(result[key])}`);
            }

            /* ADD USER TO A GROUP */
            var options = {
                method: 'POST',
                url: 'https://graph.microsoft.com/v1.0/groups/${GROUP_ID}/members/$ref',
                headers: {
                    'Authorization': 'Bearer ' + token,
                    'content-type': 'application/json'
                },
                body: JSON.stringify({
                    "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${result.invitedUser.id}`
                })
            };

            request(options, (error, response, body) => {
                console.log(body);
                if (!error && response.statusCode == 204) {
                    console.log('OK');
                } else {
                    console.log('NOK');
                }
            });
        }
    });
});