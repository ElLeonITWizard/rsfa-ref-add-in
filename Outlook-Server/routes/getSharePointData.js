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

router.get(
  "/getSharePointData",
  //authHelper.validateJwt,
  async function (req, res) {
    //res.status(200).send("test");
    //   let itemNames = ["test new 22"];
    // res.status(200).send(itemNames);

    const authHeader = req.headers.authorization;

    if(!authHeader) {
      return res
        .status(200)
        .send({ success: false, errorDetails: "Api key mismatch" });
    }

    let apiKey = authHeader.split(" ")[1];
    console.log("Api key", process.env.API_KEY, process.env.CLIENT_ID);

    if (apiKey !== process.env.API_KEY) {
      console.log("Api key mismatch inside");
      return res
        .status(200)
        .send({ success: false, errorDetails: "Api key mismatch" });
    }

    try {
      const accessToken = await getAccessToken();

      // const authHeader = req.headers.authorization;
      // let oboRequest = {
      //   oboAssertion: authHeader.split(" ")[1],
      //   scopes: ["https://graph.microsoft.com/.default"],
      // };

      // // The Scope claim tells you what permissions the client application has in the service.
      // // In this case we look for a scope value of access_as_user, or full access to the service as the user.
      // const tokenScopes = jwt.decode(oboRequest.oboAssertion).scp.split(" ");
      // let accessAsUserScope = tokenScopes.find(
      //   (scope) => scope === "access_as_user"
      // );
      // if (accessAsUserScope !== "access_as_user") {
      //   res.status(401).send({ type: "Missing access_as_user" });
      //   return;
      // }
      // const cca = authHelper.getConfidentialClientApplication();
      // const response = await cca.acquireTokenOnBehalfOf(oboRequest);

      //const rootUrl = `https://[domain-name].sharepoint.com/sites/allcompany/_api/Web/lists/GetByTitle('testlist')/items`;
      // const params = '?$select=name&$top=10';
      ///sites/{site-id}/lists/{list-id}/items
      // const rootUrl = "/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703,3b5b528e-4f8d-4385-952a-bb7556f7a422/lists/bf718b93-639d-4cb9-8f18-443f1bddede7/items";
      // const rootUrl = '/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/lists/b0565cd5-ff86-46ce-8c7c-3769398e4d49/items'

      const rootUrl = `/sites/${process.env.SITE_ID}/drives/${process.env.DRIVE_ID}/root/children?top=10000`;
      // const rootUrl = "/sites/7f9b6e39-ef5e-46f8-bac1-b09c9b5a2703/drives/b!OW6bf17v-Ea6wbCcm1onA45SWzuNT4VDlSq7dVb3pCKTi3G_nWO5TI8YRD8b3e3n/root/children";
      // const rootUrl =
      //   "/sites/e188129e-5eb6-4632-ac09-6f3798d411a9/drives/b!nhKI4bZeMkasCW83mNQRqbc7xjFNbW1Ck86ID6D2PaDVXFawhv_ORox8N2k5jk1J/root/children?top=1000";
      //drives/{document libraryid}/root/children

      const params = "";
      const graphData = await getGraphData(
        accessToken,
        rootUrl,
        params
      );

      // console.log("data----", graphData);
      let itemNames = [];
      const oneDriveItems = graphData["value"];
      for (let item of oneDriveItems) {
        // itemNames.push(item["webUrl"]);
        // itemNames.push(item["name"]);
        itemNames.push({id: item["id"], name: item["name"]});
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
