/**
 * ============================================================
 *  INDIANA STRIP STEEL — STEEL INQUIRY FORM BACKEND
 *  Google Apps Script — logs submissions to a Google Sheet and
 *  sends an automated confirmation email to the requester.
 * ============================================================
 *
 * SETUP INSTRUCTIONS
 * -------------------
 * 1. Create a new Google Sheet (or open an existing one).
 *    Rename Sheet1 to "Inquiries" (or leave it — the script will
 *    create an "Inquiries" tab automatically if it doesn't exist).
 *
 * 2. In the Sheet, go to Extensions > Apps Script.
 *
 * 3. Delete any placeholder code and paste this entire file in.
 *
 * 4. Update the CONFIG section below (company name, reply-from
 *    name, etc.) if you'd like to customize it.
 *
 * 5. Click Deploy > New deployment.
 *      - Select type: "Web app"
 *      - Description: "Steel Inquiry Form backend"
 *      - Execute as: "Me"
 *      - Who has access: "Anyone"
 *    Click Deploy, then authorize the script when prompted
 *    (you'll see a Google warning screen since it's your own
 *    unverified script — click "Advanced" > "Go to project (unsafe)"
 *    then allow the permissions. This is required so the script
 *    can write to your Sheet and send email on your behalf.)
 *
 * 6. Copy the Web App URL it gives you
 *    (looks like: https://script.google.com/macros/s/XXXX.../exec)
 *
 * 7. Paste that URL into the SCRIPT_URL constant near the top of
 *    the <script> block in steel_inquiry_form.html.
 *
 * 8. Every time you edit this script, you must create a NEW
 *    deployment version (Deploy > Manage deployments > Edit >
 *    New version) for the changes to take effect on the live URL.
 * ============================================================
 */

// ---------------------- CONFIG ----------------------
const CONFIG = {
  sheetName: "Inquiries",
  companyName: "Indiana Strip Steel & Consulting Services, LLC.",
  replyFromName: "Indiana Strip Steel — Sales Team",
  internalNotifyEmail: "" // optional: put a staff email here to also get notified of every inquiry, or leave blank
};

// Column order written to the sheet (also used as the header row)
const COLUMNS = [
  "Timestamp",
  "Company Name",
  "Your Name",
  "Your Title",
  "Mobile",
  "Office",
  "Contact Email",
  "Steel Type",
  "Chemistry",
  "Hardness Min",
  "Hardness Max",
  "Steel Form",
  "Thickness",
  "Thickness Tolerance",
  "Width",
  "Width Tolerance",
  "ID Min",
  "ID Max",
  "OD Min",
  "OD Max",
  "Quantity",
  "Unit",
  "Delivery"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    appendToSheet(data);
    sendConfirmationEmail(data);
    if (CONFIG.internalNotifyEmail) {
      sendInternalNotification(data);
    }
    return jsonResponse({ result: "success" });
  } catch (err) {
    return jsonResponse({ result: "error", message: err.message });
  }
}

// Optional: lets you sanity-check the deployment by visiting the
// web app URL directly in a browser (GET request).
function doGet(e) {
  return jsonResponse({ status: "Steel Inquiry Form backend is running." });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendToSheet(data) {
  const sheet = getSheet();
  sheet.appendRow([
    new Date(),
    data.companyName || "",
    data.yourName || "",
    data.yourTitle || "",
    data.mobile || "",
    data.office || "",
    data.contactEmail || "",
    data.steelType || "",
    data.chemistry || "",
    data.hardnessMin || "",
    data.hardnessMax || "",
    data.steelForm || "",
    data.thickness || "",
    data.thicknessTol || "",
    data.width || "",
    data.widthTol || "",
    data.idMin || "",
    data.idMax || "",
    data.odMin || "",
    data.odMax || "",
    data.quantity || "",
    data.unit || "",
    data.delivery || ""
  ]);
}

function sendConfirmationEmail(data) {
  if (!data.contactEmail) return;

  const subject = "We received your steel inquiry — " + CONFIG.companyName;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif; color:#1c2b45; max-width:600px;">
      <h2 style="color:#1f3a6e; margin-bottom:4px;">${CONFIG.companyName}</h2>
      <p>Hi ${escapeHtml(data.yourName || "there")},</p>
      <p>Thank you for submitting a steel inquiry. Here's a summary of what we received:</p>
      <table style="border-collapse:collapse; width:100%; font-size:14px;">
        ${row("Company", data.companyName)}
        ${row("Steel Type", data.steelType)}
        ${row("Chemistry", data.chemistry)}
        ${row("Hardness", joinRange(data.hardnessMin, data.hardnessMax))}
        ${row("Form", data.steelForm)}
        ${row("Thickness", data.thickness)}
        ${row("Width", data.width)}
        ${row("ID Range", joinRange(data.idMin, data.idMax))}
        ${row("OD Range", joinRange(data.odMin, data.odMax))}
        ${row("Quantity", (data.quantity || "") + " " + (data.unit || ""))}
        ${row("Delivery", data.delivery)}
      </table>
      <p style="margin-top:20px;">A member of our sales team will follow up with you shortly. If anything above needs correcting, just reply to this email.</p>
      <p style="color:#5a6b85; font-size:12px; margin-top:30px;">This is an automated confirmation from ${CONFIG.companyName}.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: data.contactEmail,
    subject: subject,
    htmlBody: htmlBody,
    name: CONFIG.replyFromName
  });
}

function sendInternalNotification(data) {
  const subject = "New Steel Inquiry — " + (data.companyName || "Unknown company");
  const body =
    "New inquiry submitted:\n\n" +
    COLUMNS.slice(1).map(function(label, i) {
      const keys = Object.keys(data);
      return "";
    }).join("") +
    JSON.stringify(data, null, 2);

  MailApp.sendEmail({
    to: CONFIG.internalNotifyEmail,
    subject: subject,
    body: body
  });
}

function row(label, value) {
  if (!value) return "";
  return (
    '<tr>' +
    '<td style="padding:6px 10px; border:1px solid #e2e6ee; font-weight:bold; background:#f4f6fb;">' + escapeHtml(label) + '</td>' +
    '<td style="padding:6px 10px; border:1px solid #e2e6ee;">' + escapeHtml(value) + '</td>' +
    '</tr>'
  );
}

function joinRange(min, max) {
  if (!min && !max) return "";
  return (min || "?") + " – " + (max || "?");
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
