// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
require('dotenv').config({
  path: `${__dirname}/.env`
});
const express = require("express");
var router = express.Router();
const authHelper = require("../server-helpers/obo-auth-helper");
const getGraphData = require("../server-helpers/msgraph-helper");
const jwt = require("jsonwebtoken");
const msal = require("@azure/msal-node");
const getAccessToken = require("../server-helpers/helpers");

var router = express.Router();

router.get("/sprsfa/getTkn", async function (req, res) {
  const authHeader = req.headers.authorization;

  if (!authHeader) {
    return res
      .status(200)
      .send({ success: false, errorDetails: "Api key mismatch" });
  }

  let apiKey = authHeader.split(" ")[1];

  if (apiKey !== process.env.API_KEY) {
    console.log("Api key mismatch");
    return res
      .status(200)
      .send({ success: false, errorDetails: "Api key mismatch" });
  }

  try {
    const accessToken = await getAccessToken();
    res.status(200).send({ success: true, data: accessToken });
  } catch (err) {
    // On rare occasions the SSO access token is unexpired when Office validates it,
    // but expires by the time it is used in the OBO flow. Microsoft identity platform will respond
    // with "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Construct an error message to return to the client so it can refresh the SSO token.
    if (
      err.errorMessage !== undefined &&
      err.errorMessage.indexOf("AADSTS500133") !== -1
    ) {
      res.status(401).send({
        type: "TokenExpiredError",
        errorDetails: JSON.stringify(err),
      });
    } else {
      res.status(403).send({ type: "Unknown", errorDetails: err });
    }
  }
});

module.exports = router;
