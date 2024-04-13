/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
const express = require("express");
const msal = require('@azure/msal-node');
var fetch = require('./fetch');
var path = require('path');

const SERVER_PORT = process.env.PORT || 3000;
const REDIRECT_URI = "http://localhost:3000/redirect";

// Before running the sample, you will need to replace the values in the config, 
// including the clientSecret
const config = {
    auth: {
        clientId: "xxxxx",
        authority: "xxxxx",
        clientSecret: "xxxxxx"
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

// Create msal application object
const pca = new msal.ConfidentialClientApplication(config);

// Create Express App and Routes
const app = express();

app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');
var indexRouter = require('./routes/index');

app.use('/d', indexRouter);

app.get('/', (req, res) => {
        const authCodeUrlParameters = {
            scopes: ["user.read","Mail.Read","offline_access"],
            redirectUri: REDIRECT_URI,
            prompt:'consent'
        };
        // get url to sign user in and consent to scopes needed for application

        //https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/samples/msal-node-samples/refresh-token/README.md
        pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
            res.redirect(response);
        }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read","Mail.Read","offline_access"],
        redirectUri: REDIRECT_URI,
        accessType: 'offline',
    };
    pca.acquireTokenByCode(tokenRequest).then(async (response) => {
        const accessToken = response.accessToken;
        const refreshToken = () => {
            const tokenCache = pca.getTokenCache().serialize();
            const refreshTokenObject = (JSON.parse(tokenCache)).RefreshToken
            const refreshToken = refreshTokenObject[Object.keys(refreshTokenObject)[0]].secret;
            return refreshToken;
        }
        const tokens = {
            accessToken,
            refreshToken:refreshToken()
        }
        console.log(tokens)
        const graphResponse = await fetch("https://graph.microsoft.com/v1.0/me/messages", tokens.accessToken);
        mails = []
        console.log(graphResponse)
        graphResponse.value.forEach((email) => {
            mails.push({"subject":email.subject})
        })
        
        res.render('index', { mails });
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});

module.exports = app;