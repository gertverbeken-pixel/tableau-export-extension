// api/export.js
require('dotenv').config();
const express = require('express');
const serverless = require('serverless-http');
const axios = require('axios');
const app = express();

app.get('/export', async (req, res) => {
  try {
    // 1. Read query params
    const { viewId, dashboard, user } = req.query;
    const ts = new Date().toISOString();

    // 2. Sign in to Tableau REST API
    const signinResp = await axios.post(
      `${process.env.TABLEAU_SERVER}/api/3.13/auth/signin`,
      {
        credentials: {
          personalAccessTokenName: process.env.PAT_NAME,
          personalAccessTokenSecret: process.env.PAT_SECRET,
          site: { contentUrl: process.env.SITE_CONTENT_URL }
        }
      }
    );
    const token = signinResp.data.credentials.token;
    const siteId = signinResp.data.credentials.site.id;

    // 3. Fetch CSV of the full view
    const dataResp = await axios.get(
      `${process.env.TABLEAU_SERVER}/api/3.13/sites/${siteId}/views/${viewId}/data?maxAge=0`,
      {
        responseType: 'text',
        headers: { 'X-Tableau-Auth': token }
      }
    );
    // 4. Sign out
    await axios.post(
      `${process.env.TABLEAU_SERVER}/api/3.13/auth/signout`,
      {},
      { headers: { 'X-Tableau-Auth': token } }
    );

    // 5. Append tracker row
    let csv = dataResp.data;
    csv += `"TRACKER","${user}","${dashboard}","${ts}"\n`;

    // 6. Send as downloadable CSV
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', `attachment; filename="${dashboard}_${ts}.csv"`);
    return res.send(csv);

  } catch (err) {
    console.error(err);
    return res.status(500).send('Export failed: ' + err.message);
  }
});

module.exports.handler = serverless(app);