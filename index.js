import axios from "axios";
import OAuthClient from "intuit-oauth";
import cron from "node-cron";
import { google } from "googleapis";

// ============================
// ENV VARS YOU MUST SET
// ============================
// --- Google ---
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const GOOGLE_CLIENT_SECRET = process.env.GOOGLE_CLIENT_SECRET;
const GOOGLE_REFRESH_TOKEN = process.env.GOOGLE_REFRESH_TOKEN;
const GOOGLE_REDIRECT_URI = process.env.GOOGLE_REDIRECT_URI; // same one you used in OAuth playground
const RECEIPTS_MASTER_FOLDER_ID = process.env.RECEIPTS_MASTER_FOLDER_ID; // your master receipts folder
const PROCESSED_FOLDER_NAME = process.env.PROCESSED_FOLDER_NAME || "Processed";

// --- QuickBooks Online ---
const QBO_CLIENT_ID = process.env.QBO_CLIENT_ID;
const QBO_CLIENT_SECRET = process.env.QBO_CLIENT_SECRET;
const QBO_REDIRECT_URI = process.env.QBO_REDIRECT_URI;
const QBO_ENVIRONMENT = process.env.QBO_ENVIRONMENT || "production"; // "sandbox" or "production"
const QBO_REALM_ID = process.env.QBO_REALM_ID; // company ID
const QBO_REFRESH_TOKEN = process.env.QBO_REFRESH_TOKEN;

// Optional mapping of jobName -> QBO CustomerRef id
// Example: {"81":"123","mike":"456"}
const JOB_MAP = process.env.JOB_MAP ? JSON.parse(process.env.JOB_MAP) : {};

// How close a QBO expense date must be to receipt date (days)
const DATE_WINDOW_DAYS = Number(process.env.DATE_WINDOW_DAYS || 2);

// ============================
// GOOGLE AUTH
// ============================
function getGoogleDriveClient() {
  const oauth2Client = new google.auth.OAuth2(
    GOOGLE_CLIENT_ID,
    GOOGLE_CLIENT_SECRET,
    GOOGLE_REDIRECT_URI
  );

  oauth2Client.setCredentials({
    refresh_token: GOOGLE_REFRESH_TOKEN,
  });

  return google.drive({ version: "v3", auth: oauth2Client });
}

// ============================
// QBO AUTH
// ============================
const qboAuth = new OAuthClient({
  clientId: QBO_CLIENT_ID,
  clientSecret: QBO_CLIENT_SECRET,
  environment: QBO_ENVIRONMENT,
  redirectUri: QBO_REDIRECT_URI,
});

async function getQboAccessToken() {
  const tokenObj = await qboAuth.refreshUsingToken(QBO_REFRESH_TOKEN);
  return tokenObj.getJson().access_token;
}

function qboBaseUrl() {
  return QBO_ENVIRONMENT === "sandbox"
    ? "https://sandbox-quickbooks.api.intuit.com"
    : "https://quickbooks.api.intuit.com";
}

// ============================
// HELPERS
// ============================
function parseReceiptFilename(name) {
  // Expected: Vendor_jobname_date_$amount.pdf
  // Example: HomeDepot_81_11-21-2025_$298.00.pdf
  const clean = name.replace(".pdf", "");

  // vendor = first chunk, remainder joined for safety
  const parts = clean.split("_");
  if (parts.length < 4) return null;

  const vendor = parts[0];

  const jobName = parts[1] || "NoJob";
  const dateStr = parts[2] || "NoDate";
  const amountStr = parts.slice(3).join("_"); // in case amount has underscores

  // normalize amount like "$298.00" -> 298.00
  const amtMatch = amountStr.match(/([0-9]+(?:\.[0-9]{1,2})?)/);
  const amount = amtMatch ? Number(amtMatch[1]) : null;

  // normalize date "11-21-2025"
  const dParts = dateStr.split("-");
  if (dParts.length !== 3) return null;
  const [mm, dd, yyyy] = dParts.map((x) => Number(x));
  const date = new Date(yyyy, mm - 1, dd);

  return { vendor, jobName, amount, date };
}

function dateToQbo(date) {
  // QBO uses YYYY-MM-DD
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function withinWindow(receiptDate, expenseDate) {
  const diffMs = Math.abs(receiptDate.getTime() - expenseDate.getTime());
  const diffDays = diffMs / (1000 * 60 * 60 * 24);
  return diffDays <= DATE_WINDOW_DAYS;
}

// ============================
// GOOGLE DRIVE: LIST RECEIPTS
// ============================
async function getOrCreateProcessedFolder(drive) {
  const res = await drive.files.list({
    q: `'${RECEIPTS_MASTER_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and name='${PROCESSED_FOLDER_NAME}' and trashed=false`,
    fields: "files(id,name)",
  });

  if (res.data.files?.length) return res.data.files[0].id;

  const created = await drive.files.create({
    requestBody: {
      name: PROCESSED_FOLDER_NAME,
      mimeType: "application/vnd.google-apps.folder",
      parents: [RECEIPTS_MASTER_FOLDER_ID],
    },
    fields: "id",
  });

  return created.data.id;
}

async function listAllPdfReceipts(drive) {
  // Find PDFs under master folder (including vendor subfolders)
  const res = await drive.files.list({
    q: `'${RECEIPTS_MASTER_FOLDER_ID}' in parents and trashed=false`,
    fields: "files(id,name,mimeType)",
  });

  const direct = res.data.files || [];
  let all = [...direct];

  // include subfolders
  for (const f of direct) {
    if (f.mimeType === "application/vnd.google-apps.folder") {
      const sub = await drive.files.list({
        q: `'${f.id}' in parents and mimeType='application/pdf' and trashed=false`,
        fields: "files(id,name,mimeType,parents)",
      });
      all.push(...(sub.data.files || []));
    }
  }

  // filter pdfs only
  return all.filter((x) => x.mimeType === "application/pdf");
}

// ============================
// QBO: FIND & UPDATE EXPENSE
// ============================
async function findMatchingExpense(accessToken, { vendor, amount, date }) {
  // Query last ~30 days expenses of that amount
  const dateStart = new Date(date);
  dateStart.setDate(dateStart.getDate() - DATE_WINDOW_DAYS);
  const dateEnd = new Date(date);
  dateEnd.setDate(dateEnd.getDate() + DATE_WINDOW_DAYS);

  const query = `
    select * from Purchase
    where TxnDate >= '${dateToQbo(dateStart)}'
      and TxnDate <= '${dateToQbo(dateEnd)}'
      and TotalAmt = '${amount}'
    maxresults 20
  `;

  const url = `${qboBaseUrl()}/v3/company/${QBO_REALM_ID}/query?query=${encodeURIComponent(query)}`;

  const res = await axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json",
    },
  });

  const purchases = res.data?.QueryResponse?.Purchase || [];
  if (!purchases.length) return null;

  // If multiple, pick first within window (extra safety)
  for (const p of purchases) {
    const pDate = new Date(p.TxnDate);
    if (withinWindow(date, pDate)) return p;
  }

  return purchases[0];
}

async function markBillableTaxable(accessToken, purchase, jobName) {
  // Clone the object and patch Line[] fields
  const updated = JSON.parse(JSON.stringify(purchase));

  if (updated.Line && Array.isArray(updated.Line)) {
    updated.Line = updated.Line.map((line) => {
      if (line.DetailType === "AccountBasedExpenseLineDetail") {
        line.AccountBasedExpenseLineDetail =
          line.AccountBasedExpenseLineDetail || {};

        // Billable
        line.AccountBasedExpenseLineDetail.BillableStatus = "Billable";

        // Taxable: set TaxCodeRef "TAX" (works for most QBO setups)
        line.AccountBasedExpenseLineDetail.TaxCodeRef = { value: "TAX" };

        // Customer/job if known
        const custId = JOB_MAP[jobName];
        if (custId) {
          line.AccountBasedExpenseLineDetail.CustomerRef = { value: custId };
        }
      }
      return line;
    });
  }

  const url = `${qboBaseUrl()}/v3/company/${QBO_REALM_ID}/purchase?minorversion=75`;

  const res = await axios.post(url, updated, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      Accept: "application/json",
    },
  });

  return res.data?.Purchase || null;
}

async function attachReceiptToPurchase(accessToken, purchaseId, fileBlobBase64, fileName) {
  // QBO Attachments endpoint is multipart/form-data
  const boundary = "----rltboundary" + Date.now();

  const metadata = {
    AttachableRef: [
      {
        EntityRef: { type: "Purchase", value: purchaseId },
      },
    ],
    FileName: fileName,
    ContentType: "application/pdf",
  };

  const metaPart =
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="file_metadata_01"\r\n` +
    `Content-Type: application/json; charset=UTF-8\r\n\r\n` +
    `${JSON.stringify(metadata)}\r\n`;

  const filePart =
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="file_content_01"; filename="${fileName}"\r\n` +
    `Content-Type: application/pdf\r\n` +
    `Content-Transfer-Encoding: base64\r\n\r\n` +
    `${fileBlobBase64}\r\n`;

  const endPart = `--${boundary}--\r\n`;

  const body = metaPart + filePart + endPart;

  const url = `${qboBaseUrl()}/v3/company/${QBO_REALM_ID}/upload`;

  await axios.post(url, body, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": `multipart/form-data; boundary=${boundary}`,
      Accept: "application/json",
    },
    maxBodyLength: Infinity,
  });
}

// ============================
// MAIN WORKER
// ============================
async function processReceiptsOnce() {
  const drive = getGoogleDriveClient();
  const processedFolderId = await getOrCreateProcessedFolder(drive);

  const pdfs = await listAllPdfReceipts(drive);

  if (!pdfs.length) {
    console.log("No PDFs found.");
    return;
  }

  const accessToken = await getQboAccessToken();

  for (const file of pdfs) {
    try {
      const parsed = parseReceiptFilename(file.name);

      if (!parsed || !parsed.amount || parsed.date.toString() === "Invalid Date") {
        console.log(`Skipping (bad filename): ${file.name}`);
        continue;
      }

      console.log(`Processing: ${file.name}`, parsed);

      // download pdf
      const dl = await drive.files.get(
        { fileId: file.id, alt: "media" },
        { responseType: "arraybuffer" }
      );

      const buffer = Buffer.from(dl.data);
      const base64 = buffer.toString("base64");

      // find matching expense
      const match = await findMatchingExpense(accessToken, parsed);
      if (!match) {
        console.log(`No QBO match for ${file.name}`);
        continue;
      }

      // attach
      await attachReceiptToPurchase(accessToken, match.Id, base64, file.name);
      console.log(`Attached to Purchase ${match.Id}`);

      // billable + taxable + job
      const updated = await markBillableTaxable(accessToken, match, parsed.jobName);
      if (updated) console.log(`Updated billable/taxable for ${match.Id}`);

      // move to processed folder
      await drive.files.update({
        fileId: file.id,
        addParents: processedFolderId,
        removeParents: (file.parents || []).join(","),
        fields: "id, parents",
      });

      console.log(`Moved to Processed: ${file.name}`);
    } catch (err) {
      console.error(`Error on file ${file.name}:`, err?.response?.data || err);
    }
  }
}

// ============================
// SCHEDULE (every 5 minutes)
// ============================
cron.schedule("*/5 * * * *", async () => {
  console.log("⏳ Receipt matcher tick:", new Date().toISOString());
  await processReceiptsOnce();
});

// run immediately on boot
processReceiptsOnce().catch(console.error);
