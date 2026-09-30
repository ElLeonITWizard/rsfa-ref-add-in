const path = require("path");

require("dotenv").config({
  path: path.join(__dirname, "../.env"),
});

const msal = require("@azure/msal-node");

const getAccessToken = async () => {
  try {
    const msalConfig = {
      auth: {
        clientId: process.env.CLIENT_ID,
        authority:
          "https://login.microsoftonline.com/" + process.env.TENANT_ID,
        clientSecret: process.env.CLIENT_SECRET,
      },
    };

    const cca = new msal.ConfidentialClientApplication(msalConfig);

    const tokenRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };

    const response =
      await cca.acquireTokenByClientCredential(tokenRequest);

    return response.accessToken;
  } catch (error) {
    return error;
  }
};

module.exports = getAccessToken;