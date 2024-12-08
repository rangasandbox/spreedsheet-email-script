function sendEmailsBasedOnDepth() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contactsSheet = ss.getSheetByName('Email');
    const templatesSheet = ss.getSheetByName('Template');
    
    if (!contactsSheet || !templatesSheet) {
      throw new Error('Could not find sheets named "Email" and "Template"');
    }
    
    const contactsData = contactsSheet.getDataRange().getValues();
    const templatesData = templatesSheet.getDataRange().getValues();
    
    // Find column indexes
    const headers = contactsData[0];
    const firstNameCol = findColumnIndex(headers, 'first name');
    const lastNameCol = findColumnIndex(headers, 'last name');
    const emailCol = findColumnIndex(headers, 'email');
    const companyCol = findColumnIndex(headers, 'company');
    const investorNamesCol = findColumnIndex(headers, 'investor names');
    const depthCol = findColumnIndex(headers, 'depth');
    const sentCol = findColumnIndex(headers, 'sent');
    
    // Create template map with both subject and body
    const templateMap = new Map();
    for (let i = 1; i < templatesData.length; i++) {
      const depth = templatesData[i][0];
      const subject = templatesData[i][1];
      const template = templatesData[i][2];
      if (depth !== '' && template !== '') {
        templateMap.set(depth, { subject, template });
      }
    }
    
    // Email signature
    const signature = `
Best regards,
Sumit
Founder
example@company.com
+1 (123) 456-7890`;
    
    // Process each contact
    let emailsSent = 0;
    for (let i = 1; i < contactsData.length; i++) {
      const row = contactsData[i];
      const depth = row[depthCol];
      const email = row[emailCol];
      
      if (!email || (depth !== 0 && depth !== 1)) {
        continue;
      }
      
      const templateData = templateMap.get(depth);
      if (!templateData) {
        continue;
      }
      
      const from = "Rin";
      // Replace placeholders and add signature
      let emailBody = templateData.template
        .replace('[Name]', `${row[firstNameCol]} ${row[lastNameCol]}`)
        .replace('[Company Name]', row[companyCol] || '')
        .replace('[Investor Names]', row[investorNamesCol] || '') + 
        '\n\n' + signature;

      // Create HTML version with proper formatting
      const htmlBody = `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
  body {
    margin: 0;
    padding: 0;
    width: 100% !important;
    -webkit-text-size-adjust: 100%;
    -ms-text-size-adjust: 100%;
  }
  p {
    margin: 0;
    padding: 0;
    width: 100% !important;
    line-height: 1.5;
    white-space: pre-wrap;
  }
  .signature {
    margin-top: 20px;
    color: #666;
    font-size: 0.9em;
  }
</style>
</head>
<body>
<p>${emailBody}</p>
</body>
</html>`;
      
      try {
        MailApp.sendEmail({
          to: email,
          subject: templateData.subject,
          htmlBody: htmlBody,
          name: from
        });
        
        emailsSent++;
        
        if (sentCol !== -1) {
          contactsSheet.getRange(i + 1, sentCol + 1).setValue('Yes');
        }
      } catch (error) {
        Logger.log(`Failed to send email to ${email}: ${error.toString()}`);
      }
    }
    
    SpreadsheetApp.getUi().alert(`Process complete. ${emailsSent} emails sent.`);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
  }
}

function findColumnIndex(headers, columnName) {
  return headers.findIndex(header => 
    header.toString().toLowerCase().replace(/\s+/g, '') === 
    columnName.toLowerCase().replace(/\s+/g, '')
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Sender')
    .addItem('Send Emails Based on Depth', 'sendEmailsBasedOnDepth')
    .addToUi();
}
