// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
require('dotenv').config({
  path: `${__dirname}/.env`
});
const express = require("express");
var router = express.Router();
const authHelper = require("../server-helpers/obo-auth-helper");
const getAccessToken = require("../server-helpers/helpers");
const getGraphData = require("../server-helpers/msgraph-helper");
const jwt = require("jsonwebtoken");
const msal = require("@azure/msal-node");

router.get(
  "/getSharePointSubFolderData",
  //authHelper.validateJwt,
  async function (req, res) {
    //res.status(200).send("test");
    //   let itemNames = ["test new 22"];
    // res.status(200).send(itemNames);
    console.log("req.query.id", req.query.folderId, req.query.subFolderName);

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

      let path = "";
      if (!req.query.subFolderName) {
        // path = `/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703/drives/b!OW6bf17v-Ea6wbCcm1onA45SWzuNT4VDlSq7dVb3pCKTi3G_nWO5TI8YRD8b3e3n/items/${req.query.folderId}/children`;
        path = `/sites/${process.env.SITE_ID}/drives/${process.env.DRIVE_ID}/items/${req.query.folderId}/children?top=1000`;
      } else {
        // path = `/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703/drives/b!OW6bf17v-Ea6wbCcm1onA45SWzuNT4VDlSq7dVb3pCKTi3G_nWO5TI8YRD8b3e3n/items/${req.query.folderId}:/${req.query.subFolderName}:/children`;
        path = `/sites/${process.env.SITE_ID}/drives/${process.env.DRIVE_ID}/items/${req.query.folderId}:/${req.query.subFolderName}:/children?top=1000`;
      }

      // const rootUrl =
      //   "/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/drives/b!nhKI4bZeMkasCW83mNQRqbc7xjFNbW1Ck86ID6D2PaDVXFawhv_ORox8N2k5jk1J/items/01NTBFYXAVNZ5DGIF2DJCYLA5Y7LOXAVFP/children?top=1000";
      //drives/{document libraryid}/root/children
      const encodedURI = encodeURI(path);
      const params = "";
      const graphData = await getGraphData(accessToken, encodedURI, params);

      // console.log("data----", graphData);
      let itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        // itemNames.push(item["webUrl"]);
        if (item.hasOwnProperty("folder")) {
          itemNames.push({ id: item["id"], name: item["name"] });
        }
      }

      res.status(200).send(itemNames);
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
        res
          .status(403)
          .send({ type: "Unknown", errorDetails: JSON.stringify(err) });
      }
    }
  }
);

module.exports = router;
