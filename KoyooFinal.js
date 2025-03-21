/**
 * This script processes funding applications from a Google Form linked to Google Sheets.
 * It checks eligibility, calculates the amount each beneficiary should receive,
 * and generates a summary report for the admin in PDF format.
 */

function processFundingApplications() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("responses1"); // Adjust if sheet name is different
    var data = sheet.getDataRange().getValues(); // Fetch all data from the sheet
    var headers = data[0]; // Store the header row
    var applicants = data.slice(1); // Store applicant data (excluding headers)
    
    var eligibleApplicants = [];
    var rejectedApplicants = [];
    var totalFundsAvailable = promptForFunds(); // Admin manually inputs the available amount

    if (totalFundsAvailable === null) {
        Logger.log("Funding input canceled by admin.");
        return; // Stop execution if admin cancels input
    }
    
    // Column indexes based on the provided Google Sheet structure
    var educationLevelIndex = headers.findIndex(h => h.trim().toLowerCase() === "education level");
    var membershipDurationIndex = headers.findIndex(h => h.trim().toLowerCase() === "how long has the parent/guardian been a contributing member?");
    var initiativesIndex = headers.findIndex(h => h.trim().toLowerCase() === "is the guardian actively involved in the group's (ketf) initiatives?");
    var membershipNumberIndex = headers.findIndex(h => h.trim().toLowerCase() === "parent/guardian's membership number");
    
    applicants.forEach(function(row) {
        var educationLevel = row[educationLevelIndex];
        var membershipDuration = parseInt(row[membershipDurationIndex]); // Convert months to number
        var activeInInitiatives = String(row[initiativesIndex]).toLowerCase() === "yes";
        var hasMembership = String(row[membershipNumberIndex]).trim() !== "";

        // Eligibility check based on rules
        if (!hasMembership || membershipDuration < 6 || !activeInInitiatives) {
            rejectedApplicants.push({ name: row[1], reason: "Eligibility criteria not met" });
            return;
        }
        
        // Check funding limits
        var maxAmount = 0;
        if (!educationLevel) {
        rejectedApplicants.push({ name: row[1], reason: "Education level missing." });
        return;
        }

        if (educationLevel.toString().toLowerCase() === "primary school") {
            rejectedApplicants.push({ name: row[1], reason: "Primary school students are not eligible for funding." });
            return;
        } else if (educationLevel.toString().toLowerCase() === "junior secondary") {
            maxAmount = 1000;
        } else if (educationLevel.toString().toLowerCase() === "senior secondary") {
            maxAmount = 10000;
        } else if (educationLevel.toString().toLowerCase() === "tertiary") {
            maxAmount = 15000;
        }
        
        eligibleApplicants.push({ name: row[1], maxAmount: maxAmount });
    });
    
    // Allocate funds proportionally
    distributeFunds(eligibleApplicants, totalFundsAvailable);
    
    // Generate PDF summary and send it to the admin
    generateSummaryReport(eligibleApplicants, rejectedApplicants, totalFundsAvailable);
}

/**
 * Prompts the admin to manually enter available funds.
 * Returns the entered amount or null if canceled.
 */
function promptForFunds() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter available funding amount (Ksh):");
    if (response.getSelectedButton() == ui.Button.CANCEL) {
        return null;
    }
    var funds = parseFloat(response.getResponseText());
    return isNaN(funds) ? 0 : funds;
}

/**
 * Distributes funds proportionally among eligible applicants based on available funds.
 */
function distributeFunds(eligibleApplicants, totalFunds) {
    var totalRequested = eligibleApplicants.reduce((sum, app) => sum + app.maxAmount, 0);
    if (totalRequested === 0) return;
    
    var scalingFactor = Math.min(1, totalFunds / totalRequested);
    eligibleApplicants.forEach(app => {
        app.finalAmount = Math.floor(app.maxAmount * scalingFactor); // Round down to avoid exceeding total funds
    });
}

/**
 * Generates a summary report of accepted and rejected applicants as a PDF and emails it to the admin.
 */
function generateSummaryReport(eligibleApplicants, rejectedApplicants, totalFunds) {
    var doc = DocumentApp.create("Funding Summary Report");
    var body = doc.getBody();
    
    body.appendParagraph("Funding Summary Report").setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph("Total Funds Available: Ksh " + totalFunds);
    
    body.appendParagraph("\nAccepted Applicants:").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    eligibleApplicants.forEach(app => {
        body.appendParagraph(app.name + " - Ksh " + app.finalAmount);
    });
    
    body.appendParagraph("\nRejected Applicants:").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    rejectedApplicants.forEach(app => {
        body.appendParagraph(app.name + " - " + app.reason);
    });

    // Save and retrieve document URL
    doc.saveAndClose();
    
    var docId = doc.getId(); // Get the document ID
    var pdf = DriveApp.getFileById(docId).getAs("application/pdf"); // Fetch PDF from Google Drive

    // Send the email with the PDF attachment
    MailApp.sendEmail({
        to: "izakonyach@gmail.com",
        cc: "ccemail@example.com",
        bcc: "bccemail@example.com",
        subject: "KOYOO Trust Fund's Funding Report",
        body: "See attached PDF.",
        attachments: [pdf]
    });
}