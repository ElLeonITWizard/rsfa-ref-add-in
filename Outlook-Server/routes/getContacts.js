require('dotenv').config({
  path: `${__dirname}/.env`
});
const express = require("express");
var router = express.Router();
const fetch = require("node-fetch");

const MONDAY_BOARD_ID = 2074688484;
const MONDAY_COLUMN_IDS = ["contact_email", "email_mkxd41p7", "contact_phone"];

const MONDAY_ACCESS_TOKEN =
  process.env.MONDAY_API_TOKEN ||
  "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjI3OTM2ODQ0OCwiYWFpIjoxMSwidWlkIjo0Njc1OTgxNywiaWFkIjoiMjAyMy0wOS0wNFQwOTo0NDowMi4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTgxNTA1NjIsInJnbiI6ImFwc2UyIn0.axJ3pNzUM5ZnMOa2mp4OcOvpD9zqx2Ramu3FewxJ9Qs";

async function fetchContactsPage(cursor = null) {
  const cursorPart = cursor
    ? `cursor: "${cursor}"`
    : `columns: []`; 

  const query = `
    query {
      items_page_by_column_values(
        limit: 500,
        board_id: ${MONDAY_BOARD_ID},
        ${cursorPart}
      ) {
        cursor
        items {
          id
          name
          column_values(ids: ${JSON.stringify(MONDAY_COLUMN_IDS)}) {
            text
          }
        }
      }
    }
  `;

  const response = await fetch("https://api.monday.com/v2/", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + MONDAY_ACCESS_TOKEN,
    },
    body: JSON.stringify({ query })
  });

  const json = await response.json();
  if (json.errors || json.status_code === 500) {
    console.error("Monday error:", JSON.stringify(json, null, 2));
    throw new Error("Oops, something went wrong while calling Monday API");
  }

  return json;
}

router.get("/getContacts", async function (req, res) {
  try {
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

    // Pagination added - LM
    let allItems = [];
    let cursor = null;
    let firstResponse = null;

    do {
      const jsonResponse = await fetchContactsPage(cursor);

      if (!firstResponse) {
        firstResponse = jsonResponse;
      }

      const page =
        jsonResponse?.data?.items_page_by_column_values || { cursor: null, items: [] };

      allItems = allItems.concat(page.items || []);
      cursor = page.cursor || null;
    } while (cursor);

	const filteredItems = allItems.filter((item) => {
	  const name = (item.name || "").trim();
	  const cv = item.column_values || [];
	  const email = (cv[0]?.text || "").trim();
	  const email2 = (cv[1]?.text || "").trim();
	  const phone = (cv[2]?.text || "").trim();

	  if (/multiple agents found/i.test(name)) {
		return false;
	  }
	  return email || email2 || phone;
	});


    // 🔁 2) Deduplicar por (name + email + email2 + phone)
    const seen = new Set();
    const uniqueItems = [];

    for (const item of filteredItems) {
      const cv = item.column_values || [];
      const sig =
        (item.name || "").trim().toLowerCase() +
        "|" +
        (cv[0]?.text || "").trim().toLowerCase() +
        "|" +
        (cv[1]?.text || "").trim().toLowerCase() +
        "|" +
        (cv[2]?.text || "").trim().toLowerCase();

      if (!seen.has(sig)) {
        seen.add(sig);
        uniqueItems.push(item);
      }
    }

    if (!firstResponse.data) {
      firstResponse.data = {};
    }

    firstResponse.data.items_page_by_column_values = {
      cursor: null,
      items: uniqueItems,
    };

    res.status(200).send({ success: true, data: firstResponse });
  } catch (err) {
    console.error("getContacts error:", err);
    res
      .status(err.statusCode || 500)
      .send({ type: "Unknown", errorDetails: JSON.stringify(err) });
  }
});

module.exports = router;
