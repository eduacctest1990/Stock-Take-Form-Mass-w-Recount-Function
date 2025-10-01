// /api/archive.js

import { ConfidentialClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import "isomorphic-fetch";

// --- SharePoint Configuration ---
const SHAREPOINT_SITE_NAME = "Operations Stock Count";
const SHAREPOINT_LIBRARY_NAME = "Documents";

// Main function to handle the API request
export default async function handler(req, res) {
  // We only accept POST requests for this action
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method Not Allowed' });
  }

  try {
    const graphClient = await getAuthenticatedGraphClient();
    
    // --- 1. GET DATA FROM THE FRONTEND ---
    // The comparisonResults array is sent from the browser
    const stockTakeData = req.body.data;
    if (!stockTakeData || stockTakeData.length === 0) {
      return res.status(400).json({ message: "No data received to archive." });
    }
    
    // --- 2. GENERATE CSV AND UPLOAD ---
    const siteId = await getSharePointSiteId(graphClient, SHAREPOINT_SITE_NAME);
    const csvData = generateCsv(stockTakeData);
    const fileName = createFileName();

    await uploadFileToSharePoint(graphClient, siteId, SHAREPOINT_LIBRARY_NAME, fileName, csvData);

    console.log(`✅ Successfully uploaded ${fileName}.`);
    res.status(200).json({ 
      message: `Successfully archived ${fileName} to SharePoint.`
    });

  } catch (error) {
    console.error("❌ An error occurred:", error);
    res.status(500).json({ message: "An error occurred.", error: error.message });
  }
}

// --- HELPER FUNCTIONS ---

async function getAuthenticatedGraphClient() {
  const { SHAREPOINT_TENANT_ID, SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET } = process.env;
  const msalConfig = {
    auth: {
      clientId: SHAREPOINT_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${SHAREPOINT_TENANT_ID}`,
      clientSecret: SHAREPOINT_CLIENT_SECRET,
    },
  };
  const cca = new ConfidentialClientApplication(msalConfig);
  const authResponse = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  if (!authResponse.accessToken) throw new Error("Failed to acquire access token.");
  return Client.init({ authProvider: (done) => done(null, authResponse.accessToken) });
}

async function getSharePointSiteId(graphClient, siteName) {
  const site = await graphClient.api(`/sites`).filter(`displayName eq '${siteName}'`).get();
  if (site.value.length === 0) throw new Error(`SharePoint site '${siteName}' not found.`);
  return site.value[0].id;
}

function generateCsv(records) {
  const header = 'ItemID,SystemBalance,InitialPhysical,FinalPhysical,Difference,Status,RecountHistory\n';
  const rows = records.map(r => {
    // Join recount history array into a single string, e.g., "150|145"
    const recountHistoryStr = r.recountHistory.join('|'); 
    return `"${r.itemId}",${r.systemQty},${r.initialPhysicalQty},${r.finalPhysicalQty},${r.difference},"${r.status}","${recountHistoryStr}"`;
  });
  return header + rows.join('\n');
}

function createFileName() {
  const date = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
  const uniqueId = Math.random().toString(36).substr(2, 9);
  return `Inventory-Comparison-${date}-${uniqueId}.csv`;
}

async function uploadFileToSharePoint(graphClient, siteId, libraryName, fileName, fileContent) {
  const uploadUrl = `/sites/${siteId}/drive/root:/${libraryName}/${fileName}:/content`;
  await graphClient.api(uploadUrl).put(fileContent);
  console.log(`  -> Uploaded ${fileName}`);
}