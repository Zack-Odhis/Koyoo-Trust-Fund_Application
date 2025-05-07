# Funding Application Processing System

## Overview
This project automates the processing of funding applications submitted via Google Forms and stored in Google Sheets. It evaluates applicant eligibility, calculates fund distribution, and generates a summary report in PDF format. The system also exports relevant data as a CSV file and sends it via email.

## Features
- **Eligibility Check**: Evaluates applicants based on membership duration, initiative participation, and education level.
- **Fund Distribution**: Allocates funds proportionally based on the total available funding.
- **Report Generation**: Creates a summary report in Google Docs and converts it to PDF.
- **Email Notification**: Sends the report and CSV data to specified recipients.

## How It Works
1. The script retrieves applicant data from the Google Sheet.
2. It verifies eligibility based on:
   - Membership status and duration
   - Involvement in initiatives
   - Education level (primary school students are ineligible)
3. Eligible applicants receive funding based on a proportional distribution model.
4. A summary report is generated in Google Docs, saved as a PDF, and emailed along with a CSV file.

## Setup Instructions
### Prerequisites
To use this system, you need:
- **A Google Sheet** linked to a Google Form for collecting applications.
- **Google Apps Script API** enabled for script execution.
- **Google Drive & Mail Access** for generating and sending reports.

### Deployment Steps
1. Open your Google Sheet linked to the funding application form.
2. Go to **Extensions > Apps Script**.
3. Copy and paste the provided script into the Apps Script editor.
4. Save and deploy the script.
5. Grant the necessary permissions for the script to run.

## File Structure
```
funding-application/
│── src/
│   ├── fundingProcessor.gs  # Main script for processing applications
│   ├── reportGenerator.gs   # Handles report creation and emailing
│   ├── googleAPI.gs         # Google Sheets and Docs API integration
│── README.md                # Project documentation
```

## Usage
1. Run the script in Google Apps Script by executing:
   ```js
   processFundingApplications();
   ```
2. Enter the total available funding when prompted.
3. Review the generated summary report and email notifications.

## Dependencies
- Google Apps Script
- Google Sheets API
- Google Docs API
- Google Drive API
- Google Mail API (MailApp)

## Contact
For any issues or improvements, please reach out to:
- Email: luisconstancio200gmail.com
- CC: izakonyach@gmail.com

