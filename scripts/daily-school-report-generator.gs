/**
 * Generates a daily report by finding a specific lesson plan file,
 * and a weekly meal plan file, extracting relevant content from both,
 * and writing it to the current Google Docs file.
 */
function generateDailyReport() {
  // IMPORTANT: Replace this with the ID of your Google Drive folder.
  // The folder ID is the long string in the URL after '/folders/'.
  const LESSON_PLAN_FOLDER_ID = '0BxIRVU3Uu0TqYzBhM2VkNDQtNWQxZS00NDA2LWJkMjctNjQ3NDMxZTVjMGIx'; 
  const MEAL_PLAN_FOLDER_ID = '0BxIRVU3Uu0TqYWQ1YTRjZjEtZDk4ZS00MjUxLTgyYmItMzBlNjZjNjc5YzMy';
  
  const doc = DocumentApp.getActiveDocument();
  const docUrl = doc.getUrl();
  const docName = doc.getName();
  
  try {
    const today = new Date();
    const extractedElements = [];
    let mealPlanImageBlob = null;
    let mealPlanFileUrl = null;
    
    // --- Step 1: Extract the Daily Classroom Report from an email ---
    const classroomReportContent = extractClassroomReportContent();
    if (classroomReportContent) {
      extractedElements.push({ type: 'HEADING', data: 'Daily Classroom Report' });
      // The `classroomReportContent.text` is a single string.
      // We will create a table from the text and add it to the elements.
      if (classroomReportContent.text) {
        const reportTableData = parseReportTextToTable(classroomReportContent.text);
        if (reportTableData.length > 0) {
          extractedElements.push({ type: 'TABLE', data: reportTableData, isClassroomReport: true });
        }
      }
    }
    
    // --- Step 2: Extract the Daily Lesson Plan ---
    const day = today.getDay();
    const diff = today.getDate() - (day === 0 ? 6 : day - 1); // Adjust for Sunday (0) and Monday (1)
    const monday = new Date(today.getFullYear(), today.getMonth(), diff);
    
    const lessonPlanFileName = `R.C. Lesson Plan ${
      (monday.getMonth() + 1).toString().padStart(2, '0')}-${
      monday.getDate().toString().padStart(2, '0')}-${
      monday.getFullYear().toString().slice(-2)}`;
    
    Logger.log(`Searching for lesson plan file: "${lessonPlanFileName}"`);
    
    try {
      const folder = DriveApp.getFolderById(LESSON_PLAN_FOLDER_ID);
      const files = folder.getFilesByName(lessonPlanFileName);
      
      if (files.hasNext()) {
        const file = files.next();
        extractedElements.push({ type: 'HEADING', data: 'R.C. Lesson Plan' });
        extractedElements.push(...extractInformation(file));
      } else {
        extractedElements.push({ type: 'PARAGRAPH', data: `File not found: A lesson plan for today (${lessonPlanFileName}) does not exist in the specified folder.` });
      }
    } catch (e) {
      extractedElements.push({ type: 'PARAGRAPH', data: `Error accessing folder or file: ${e.message}` });
    }
  
    // --- Step 3: Extract the Daily Meal Plan from a PDF and get its URL or image ---
    const mealPlanFileName = `${today.toLocaleString('en-US', { month: 'long' })} ${today.getFullYear()} MAC Menu NV & V PDF.pdf`;
    Logger.log(`Searching for meal plan file: "${mealPlanFileName}"`);
  
    try {
      const folder = DriveApp.getFolderById(MEAL_PLAN_FOLDER_ID);
      const files = folder.searchFiles(
        `title = '${mealPlanFileName}' and mimeType = '${MimeType.PDF}'`
      );
  
      if (files.hasNext()) {
        const file = files.next();
        Logger.log(`Found file: ${file.getName()}`);
        Logger.log(`Found file MIME type: ${file.getMimeType()}`);
        
        // Try to get the thumbnail image first. If it fails, get the URL as a fallback.
        try {
          mealPlanImageBlob = file.getThumbnail();
        } catch (e) {
          Logger.log(`Warning: Failed to get thumbnail image from PDF. Falling back to URL. Error: ${e.message}`);
          // Set to null to explicitly trigger the fallback
          mealPlanImageBlob = null;
        }
        
        // Always get the URL as a fallback or for a direct link
        mealPlanFileUrl = file.getUrl();
      } else {
        extractedElements.push({ type: 'PARAGRAPH', data: `File not found: The "${mealPlanFileName}" file does not exist in the specified folder.` });
      }
    } catch (e) {
      extractedElements.push({ type: 'PARAGRAPH', data: `Error accessing meal plan file: ${e.message}` });
    }
  
    // --- Step 4: Write the report content to the current document ---
    writeToGoogleDoc(extractedElements, mealPlanImageBlob, mealPlanFileUrl);
    
    sendNotificationEmail(
      `Daily School Report - ${today.toLocaleDateString()}`, 
      `Your daily report has been updated.`,
      docName,
      extractedElements,
      mealPlanImageBlob,
      mealPlanFileUrl
    );
    
  } catch (e) {
    sendNotificationEmail(
      `Daily School Report (Failed) - ${new Date().toLocaleDateString()}`, 
      `An error occurred while generating your daily report: ${e.message}`, 
      docName
    );
  }
}

/**
 * Extracts all text and table data from a Google Docs file.
 * This is the core logic that handles different document elements.
 * @param {GoogleAppsScript.Drive.File} file The file to extract data from.
 * @return {Array<Object>} An array of objects representing the document's elements.
 */
function extractInformation(file) {
  const extractedElements = [];
  const docId = file.getId();
  
  try {
    Logger.log(`Attempting to open file with ID: ${docId}`);
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    
    // Iterate through all the elements in the document body
    for (let i = 0; i < body.getNumChildren(); i++) {
      const element = body.getChild(i);
      
      // Check the type of the element
      if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
        extractedElements.push({ type: 'PARAGRAPH', data: element.asParagraph().getText() });
      } else if (element.getType() === DocumentApp.ElementType.TABLE) {
        const table = element.asTable();
        const tableData = [];
        let highlightRowIndex = -1;
        const today = new Date();
        const todayShortName = today.toLocaleString('en-US', { weekday: 'short' }).toUpperCase();
        
        for (let r = 0; r < table.getNumRows(); r++) {
          const row = table.getRow(r);
          const rowContent = [];
          for (let c = 0; c < row.getNumCells(); c++) {
            const cell = row.getCell(c);
            rowContent.push(cell.getText());
          }
          // Check if the first cell of the row starts with the current day's three-letter abbreviation
          if (rowContent[0].toUpperCase().trim().startsWith(todayShortName)) {
            highlightRowIndex = r;
          }
          tableData.push(rowContent);
        }
        
        // Push the table data with the determined highlight row
        extractedElements.push({ type: 'TABLE', data: tableData, highlightRow: highlightRowIndex });
      } else if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
        extractedElements.push({ type: 'LIST_ITEM', data: `- ${element.asListItem().getText()}` });
      }
      // Add other element types as needed
    }
    
  } catch (e) {
    Logger.log(`Error during extraction from file ${file.getName()}: ${e.message}`);
    extractedElements.push({ type: 'PARAGRAPH', data: `Error during extraction: ${e.message}` });
  }
  
  return extractedElements;
}

/**
 * Reads the latest email with a specific subject line, extracts the text
 * and the first HTML table from its body, and returns the content.
 * @return {object | null} An object with 'text' and 'table' properties, or null if not found.
 */
function extractClassroomReportContent() {
  const today = new Date();
  const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd MMM yyyy");
  
  const subjectQuery = `subject:"Classroom Report for" newer_than:1d`;
  Logger.log(`Searching for email with subject: "${subjectQuery}"`);

  const threads = GmailApp.search(subjectQuery, 0, 10);
  
  for (const thread of threads) {
    const message = thread.getMessages()[0];
    const subject = message.getSubject();
    
    const expectedSubjectRegex = new RegExp(`^Classroom Report for [A-Z][a-z]{3} \\[${dateString}\\]$`);
    
    if (subject.match(expectedSubjectRegex)) {
      // Use getPlainBody() for reliable text extraction, and getBody() for tables
      const plainBody = message.getPlainBody();
      const htmlBody = message.getBody();
      
      let tableData = null;
      let textContent = plainBody;
      
      // Still try to extract the table from the HTML body if it exists
      if (htmlBody) {
        const tableRegex = /<table.*?>(.*?)<\/table>/s;
        const tableMatch = htmlBody.match(tableRegex);

        if (tableMatch) {
          const tableHtml = tableMatch[0];
          tableData = [];
          const rowRegex = /<tr.*?>(.*?)<\/tr>/gs;
          let rowMatch;
          while ((rowMatch = rowRegex.exec(tableHtml)) !== null) {
            const rowContent = [];
            const cellRegex = /<td.*?>(.*?)<\/td>/gs;
            let cellMatch;
            while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
              const cleanContent = cleanHtmlText(cellMatch[1].trim());
              rowContent.push(cleanContent);
            }
            tableData.push(rowContent);
          }
        }
      }
      
      Logger.log('Successfully extracted content from email.');
      return { text: textContent, table: tableData };
    }
  }

  Logger.log('No matching email thread found.');
  return null;
}

/**
 * Parses the raw text from the classroom report into a two-column table format.
 * This function uses a more robust approach to handle variable line breaks and concatenated data.
 * @param {string} reportText The plain text content of the report.
 * @return {Array<Array<string>>} A 2D array representing the table data.
 */
function parseReportTextToTable(reportText) {
  const tableData = [['Time', 'Event']];
  // Regex to find all instances of a time string (e.g., "8:57 AM")
  const timeRegex = /(\d{1,2}:\d{2} (?:AM|PM))/gi;
  // Use matchAll to get an iterator of all matches with their indices
  const matches = [...reportText.matchAll(timeRegex)];

  if (matches.length === 0) {
    return tableData; // Return just the header if no times are found.
  }
  
  // Regex to find and remove the unwanted "Powered by NeatSchool" text
  const neatSchoolRegex = /Powered by NeatSchool - https:\/\/www\.neatschool\.net\s*$/i;

  // Iterate through the matches to extract each time and the corresponding event text
  for (let i = 0; i < matches.length; i++) {
    const currentTimeMatch = matches[i];
    const time = currentTimeMatch[1];
    
    // Determine the start index of the event text (immediately after the time string)
    const eventStartIndex = currentTimeMatch.index + currentTimeMatch[0].length;
    
    // Determine the end index of the event text (the start of the next time string, or end of text)
    let eventEndIndex;
    if (i < matches.length - 1) {
      eventEndIndex = matches[i + 1].index;
    } else {
      eventEndIndex = reportText.length;
    }

    // Extract the event text and clean it up
    let eventText = reportText.substring(eventStartIndex, eventEndIndex).trim();
    
    // Remove the unwanted "Powered by NeatSchool" text from the last entry
    eventText = eventText.replace(neatSchoolRegex, '').trim();

    tableData.push([time, eventText]);
  }

  return tableData;
}


/**
 * Cleans up HTML text by replacing tags with newlines and
 * removing redundant whitespace.
 * @param {string} html The HTML string to clean.
 * @return {string} The cleaned text.
 */
function cleanHtmlText(html) {
  // Replace block-level and break tags with a newline.
  let cleanedText = html.replace(/<br\s*\/?>|<\/p>|<\/div>|<\/li>|<\/tr>|<\/td>/gi, '\n');
  
  // Remove all other remaining HTML tags.
  cleanedText = cleanedText.replace(/<[^>]*>/g, '');
  
  // Decode HTML entities like &nbsp;.
  cleanedText = cleanedText.replace(/&nbsp;|\u00a0/gi, ' ');
  
  // Clean up the text.
  // First, replace multiple spaces and tabs with a single space.
  cleanedText = cleanedText.replace(/[ \t]{2,}/g, ' ');
  // Next, replace multiple newlines with a single newline.
  cleanedText = cleanedText.replace(/\n{2,}/g, '\n');
  
  // Trim leading/trailing whitespace.
  return cleanedText.trim();
}

/**
 * Writes the provided content to the current active Google Docs file,
 * clearing any existing content first and preserving element types.
 * @param {Array<Object>} elements The array of structured document elements.
 * @param {GoogleAppsScript.Base.Blob} mealPlanImageBlob The PDF image blob to insert.
 * @param {string} mealPlanFileUrl The URL of the meal plan PDF file (fallback).
 */
function writeToGoogleDoc(elements, mealPlanImageBlob, mealPlanFileUrl) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  // Clear existing content
  body.clear();
  
  // Add a new title and the report content
  body.appendParagraph(`Daily Report: ${new Date().toLocaleDateString()}`)
      .setBold(true)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph(''); // Add a blank line for spacing
  
  // Iterate through the structured elements and write them to the document
  elements.forEach(element => {
    switch (element.type) {
      case 'HEADING':
        body.appendParagraph(element.data).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        break;
      case 'PARAGRAPH':
        // Split the text by newline characters and append each line as a new paragraph.
        const lines = element.data.split('\n');
        lines.forEach(line => {
          body.appendParagraph(line);
        });
        break;
      case 'TABLE':
        const newTable = body.appendTable(element.data);
        if (element.highlightRow !== undefined && element.highlightRow !== -1) {
          const row = newTable.getRow(element.highlightRow);
          const attributes = {};
          attributes[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFF2CC';
          for (let c = 0; c < row.getNumCells(); c++) {
            row.getCell(c).setAttributes(attributes);
          }
        }
        break;
      case 'LIST_ITEM':
        body.appendListItem(element.data);
        break;
    }
  });

  // Append the meal plan content
  if (mealPlanImageBlob) {
    body.appendParagraph('');
    body.appendParagraph('Today\'s Menu').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendImage(mealPlanImageBlob);
  } else if (mealPlanFileUrl) {
    body.appendParagraph('');
    const paragraph = body.appendParagraph('Click here to view the full meal plan:')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    paragraph.setLinkUrl(mealPlanFileUrl);
  }
  
  Logger.log('Report successfully written to Google Doc.');
}

/**
 * Sends a notification email to the current user with HTML content.
 * @param {string} subject The subject line of the email.
 * @param {string} body The body content of the email.
 * @param {string} docName The name of the Google Doc.
 * @param {Array<Object>} elements The extracted document elements.
 * @param {GoogleAppsScript.Base.Blob} mealPlanImageBlob The meal plan image blob.
 * @param {string} mealPlanFileUrl The URL of the meal plan PDF file.
 */
function sendNotificationEmail(subject, body, docName, elements, mealPlanImageBlob, mealPlanFileUrl) {
  const userEmail = Session.getActiveUser().getEmail();
  
  if (!userEmail) {
    Logger.log('Could not send notification email. No active user email found.');
    return;
  }
  
  // Set limits for content to prevent overly large emails
  const MAX_TEXT_LENGTH = 1000;
  const MAX_TABLE_ROWS = 10;

  // Build the HTML body of the email
  let htmlBody = `
    <p>Hello,</p>
    <p>${body}</p>
    <hr>
    <h3>Daily Report: ${new Date().toLocaleDateString()}</h3>
  `;
  
  // Add the report elements
  if (elements) {
    elements.forEach(element => {
      switch (element.type) {
        case 'HEADING':
          htmlBody += `<h2>${element.data}</h2>`;
          break;
        case 'PARAGRAPH':
          // Truncate long text to prevent large email bodies
          let truncatedText = element.data;
          if (truncatedText.length > MAX_TEXT_LENGTH) {
            truncatedText = truncatedText.substring(0, MAX_TEXT_LENGTH) + '...';
          }
          // Replace newlines with <br> for HTML rendering
          const paragraphHtml = truncatedText.replace(/\n/g, '<br>');
          htmlBody += `<p>${paragraphHtml}</p>`;
          break;
        case 'TABLE':
          htmlBody += '<table style="border-collapse: collapse; width: 100%;">';
          if (element.data) {
            // Limit the number of rows to prevent large email bodies
            const tableData = element.data.slice(0, MAX_TABLE_ROWS);
            if (element.data.length > MAX_TABLE_ROWS) {
              htmlBody += `<p><i>(Table truncated to ${MAX_TABLES_ROWS} rows)</i></p>`;
            }
            tableData.forEach((row, rowIndex) => {
              let rowStyle = '';
              if (element.highlightRow === rowIndex) {
                rowStyle = 'background-color: #FFF2CC;';
              }
              htmlBody += `<tr style="${rowStyle}">`;
              
              const startCellIndex = element.isClassroomReport ? 0 : 1;
              
              for (let i = startCellIndex; i < row.length; i++) {
                const cell = row[i];
                const cellHtml = cell.replace(/\n/g, '<br>');
                htmlBody += `<td style="border: 1px solid #ccc; padding: 8px;">${cellHtml}</td>`;
              }
              htmlBody += '</tr>';
            });
          }
          htmlBody += '</table>';
          
          break;
        case 'LIST_ITEM':
          htmlBody += `<p>${element.data}</p>`;
          break;
      }
    });
  }
  
  let imageUrl = null;
  if (mealPlanImageBlob) {
    try {
      // Create a temporary file in Drive from the image blob.
      const tempFile = DriveApp.createFile(mealPlanImageBlob);
      // Set sharing permissions to 'Anyone with the link' to make it publicly viewable.
      tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      // Get the file ID to construct a direct download link.
      const fileId = tempFile.getId();
      // Construct the direct download URL. This is the most reliable way to display images in email clients.
      imageUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
      
      // Add the image tag to the HTML body
      htmlBody += `<h3>Today's Menu</h3>`;
      htmlBody += `<img src="${imageUrl}" alt="Meal Plan Image" style="max-width: 100%; height: auto;">`;
      
      // Immediately trash the temporary file to keep the user's Drive clean.
      // This is important because the URL will still be active.
      tempFile.setTrashed(true);
      
    } catch (e) {
      Logger.log(`Failed to create public image file in Drive: ${e.message}. Falling back to URL.`);
      // If saving the image to Drive fails, we'll fall back to the direct PDF link.
      mealPlanImageBlob = null;
    }
  }

  // Fallback to the direct link if the image failed to load or be hosted.
  if (!mealPlanImageBlob && mealPlanFileUrl) {
    htmlBody += `<h3>Today's Menu</h3>`;
    htmlBody += `<p>View the full meal plan here: <a href="${mealPlanFileUrl}">Meal Plan PDF</a></p>`;
  }
  
  MailApp.sendEmail({
    to: userEmail,
    subject: subject,
    htmlBody: htmlBody
  });
  
  Logger.log(`Notification email sent to ${userEmail} with subject: ${subject}`);
}

