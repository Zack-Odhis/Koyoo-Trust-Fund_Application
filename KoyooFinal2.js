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
    
    Logger.log(headers); // Log headers to debug
    
    var eligibleApplicants = [];
    var rejectedApplicants = [];
    var totalFundsAvailable = promptForFunds(); // Admin manually inputs the available amount

    if (totalFundsAvailable === null) {
        Logger.log("Funding input canceled by admin.");
        return; // Stop execution if admin cancels input
    }
    
    // Column indexes based on the provided Google Sheet structure
    var educationLevelIndex = headers.findIndex(h => h.trim().toLowerCase() === "education level");
    var membershipDurationIndex = headers.findIndex(h => h.trim().toLowerCase().includes("how long"));
    var initiativesIndex = headers.findIndex(h => h.trim().toLowerCase().includes("guardian actively involved"));
    var membershipNumberIndex = headers.findIndex(h => h.trim().toLowerCase().includes("membership number"));
    var parentGuardianIndex = headers.findIndex(h => h.trim().toLowerCase().includes("parent/guardian"));
    
    Logger.log("Parent Guardian Index: " + parentGuardianIndex);
    
    // Process applications
    applicants.forEach(function(row) {
        var educationLevel = row[educationLevelIndex];
        var membershipResponse = String(row[membershipDurationIndex] || "").trim().toLowerCase();
        var activeInInitiatives = String(row[initiativesIndex] || "").toLowerCase() === "yes";
        var hasMembership = String(row[membershipNumberIndex] || "").trim() !== "";
        var parentGuardian = row[parentGuardianIndex] || "Unknown";
        
        Logger.log("Raw Parent Name Data: " + row[parentGuardianIndex]);
        
        var membershipDuration = 0;
        if (membershipResponse.includes("less than 6")) {
            membershipDuration = 0;
        } else if (membershipResponse.includes("6-12")) {
            membershipDuration = 6;
        } else if (membershipResponse.includes("over 1 year")) {
            membershipDuration = 12;
        }

        var rejectionReasons = [];

        if (!hasMembership) {
            rejectionReasons.push("No valid membership number.");
        }
        if (membershipDuration < 6) {
            rejectionReasons.push("Membership duration is less than 6 months.");
        }
        if (!activeInInitiatives) {
            rejectionReasons.push("Not actively involved in KETF initiatives.");
        }

        if (rejectionReasons.length > 0) {
            rejectedApplicants.push({ name: row[1], parent: parentGuardian, reason: rejectionReasons.join(" ") });
            return;
        }

        var maxAmount = 0;
        if (!educationLevel) {
            rejectedApplicants.push({ name: row[1], parent: parentGuardian, reason: "Education level missing." });
            return;
        }

        if (educationLevel.toString().toLowerCase() === "primary school") {
            rejectedApplicants.push({ name: row[1], parent: parentGuardian, reason: "Primary school students are not eligible for funding." });
            return;
        } else if (educationLevel.toString().toLowerCase() === "junior secondary") {
            maxAmount = 1000;
        } else if (educationLevel.toString().toLowerCase() === "senior secondary") {
            maxAmount = 10000;
        } else if (educationLevel.toString().toLowerCase() === "tertiary") {
            maxAmount = 15000;
        }

        eligibleApplicants.push({ name: row[1], parent: parentGuardian, maxAmount: maxAmount });
    });

    distributeFunds(eligibleApplicants, totalFundsAvailable);
    generateSummaryReport(eligibleApplicants, rejectedApplicants, totalFundsAvailable);
}

function promptForFunds() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt("Enter available funding amount (Ksh):");
    if (response.getSelectedButton() == ui.Button.CANCEL) {
        return null;
    }
    var funds = parseFloat(response.getResponseText());
    return isNaN(funds) ? 0 : funds;
}

function distributeFunds(eligibleApplicants, totalFunds) {
    var totalRequested = eligibleApplicants.reduce((sum, app) => sum + app.maxAmount, 0);
    if (totalRequested === 0) return;
    
    var scalingFactor = Math.min(1, totalFunds / totalRequested);
    eligibleApplicants.forEach(app => {
        app.finalAmount = Math.floor(app.maxAmount * scalingFactor);
    });
}

function generateSummaryReport(eligibleApplicants, rejectedApplicants, totalFunds) {
    var doc = DocumentApp.create("Funding Summary Report");
    var body = doc.getBody();
    
    body.appendParagraph("Funding Summary Report").setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph("Total Funds Available: Ksh " + totalFunds);
    
    body.appendParagraph("\nAccepted Applicants:").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    eligibleApplicants.forEach(app => {
        body.appendParagraph(app.name + " (Parent: " + app.parent + ") - Ksh " + app.finalAmount);
    });
    
    body.appendParagraph("\nRejected Applicants:").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    rejectedApplicants.forEach(app => {
        body.appendParagraph(app.name + " (Parent: " + app.parent + ") - " + app.reason);
    });

    doc.saveAndClose();
    
    var docId = doc.getId();
    var pdf = DriveApp.getFileById(docId).getAs("application/pdf");

    // Get active spreadsheet and convert it to CSV
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("responses1");
    var csvContent = convertSheetToCSV(sheet);
    var csvFile = DriveApp.createFile("Funding_Data.csv", csvContent, MimeType.CSV);

    // Send email with both attachments
    MailApp.sendEmail({
        to: "gopiyo77@gmail.com",
        cc: "izakonyach@gmail.com",
        subject: "KOYOO Trust Fund's Funding Report",
        body: "Hello, I have included parent next to applicant's name in the report, the system also now attaches the csv file from google sheets containing all relevant information obtained from the google forms.",
        attachments: [pdf, csvFile.getAs(MimeType.CSV)]
    });

    // Optional: Delete the CSV file from Google Drive after sending
    DriveApp.getFileById(csvFile.getId()).setTrashed(true);
}

function convertSheetToCSV(sheet) {
    var data = sheet.getDataRange().getValues();
    return data.map(row => row.join(",")).join("\n");
}