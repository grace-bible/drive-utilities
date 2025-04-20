/**
 * @OnlyCurrentDoc
*/

// Copyright by Josh McKenna 2025
// Original inspiration by Martin Hawksey 2020
//
// Licensed under the Apache License, Version 2.0 (the "License"); you may not
// use this file except in compliance with the License.  You may obtain a copy
// of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
// WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
// License for the specific language governing permissions and limitations under
// the License.

/**
 * Sends emails to recipients based on data from a specified Google Sheet.
 * 
 * Retrieves recipient email addresses and the associated email sent status and 
 * uses a Gmail draft message as a template. Logs the status of sent emails directly
 * into the specified column. Prompts the user for missing header names if not provided.
 * 
 * @param {string} [subjectLine] - Subject line for the email draft message. Optional; prompts if not provided.
 * @param {string} thisSheet - The Google Sheet file URL or ID with mail merge data.
 * @param {string} thisTab - The name of the tab within the Google Sheet with mail merge data.
 * @param {string} [emailRecipients] - Header of the column containing recipient email addresses. Optional; prompts if not provided.
 * @param {string} [emailSent] - Header of the column where email sent dates are logged. Optional; prompts if not provided.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified sheet, tab, or required headers do not exist.
 * 
 * @example
 * sendEmails("Weekly Update", "1B2c...xyr", "Mail Merge");
 */
function sendEmails(subjectLine, thisSheet, thisTab, emailRecipients, emailSent) {
  console.log(`Start sendEmails('${subjectLine}', '${thisSheet}', '${thisTab}', '${emailRecipients}', '${emailSent}')`);
  console.time(`sendEmails() processing time`);

  let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const RECIPIENT_COL = emailRecipients || getUserInput("Enter the HEADER NAME of the column of recipient email addresses:");
  const EMAIL_SENT_COL = emailSent || getUserInput("Enter the HEADER NAME of the column of email sent status:");

  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge",
      "Type or copy/paste the SUBJECT LINE of the Gmail " +
      "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine === "") {
      console.error(`Abort script due to prompt response: '${subjectLine}'`);
      return;
    }
  }

  if (!thisSheet || !thisTab) {
    sheet = SpreadsheetApp.getActiveSheet();
  } else {
    sheet = SpreadsheetApp.openById(getIdFromUrl(thisSheet)).getSheetByName(String(thisTab));
    if (!sheet) {
      console.warn(`Sheet named '${thisTab}' was not found in '${thisSheet}'.`);
      sheet = SpreadsheetApp.getActiveSheet();
      console.log(`Proceeding with the current active sheet as default: '${sheet}'`);
    }
  }

  console.info(`Sending mail merge from '${sheet.getName()}' with subject: '${subjectLine}'`);

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  const heads = data.shift();

  if (!isValidHeaderRow(heads)) {
    activeSpreadsheet.toast(`Email Send failed: Header row must contain at least two unique headers.`);
    console.error(`Header row must contain at least two unique headers. Invalid headers: '${heads}'`);
    return;
  }

  console.log(`Headers array: '${heads}'`);

  if (!heads.includes(RECIPIENT_COL)) {
    activeSpreadsheet.toast(`Email send failed due to missing column header: '${RECIPIENT_COL}'`);
    console.error(`Abort script due to missing column header: '${RECIPIENT_COL}'`);
    return;
  }

  if (!heads.includes(EMAIL_SENT_COL)) {
    activeSpreadsheet.toast(`Email send failed due to due to missing column header: '${EMAIL_SENT_COL}'`);
    console.error(`Abort script due to missing column header: '${EMAIL_SENT_COL}'`);
    return;
  }

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  console.log(`Email sent column: '${emailSentColIdx}'`);

  const emails = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const sendResult = [];

  console.time("Total row processing time");
  emails.forEach(function (row, rowIdx) {
    if (row[RECIPIENT_COL] !== '' && row[EMAIL_SENT_COL] === '' && !sheet.isRowHiddenByFilter(rowIdx + 2)) {
      console.time(`Row '${rowIdx + 2}' processing time `);
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        MailApp.sendEmail(
          row[RECIPIENT_COL],
          msgObj.subject,
          msgObj.text,
          {
            htmlBody: msgObj.html,
            // bcc: 'a.bbc@email.com',
            // cc: 'a.cc@email.com',
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          }
        );
        sendResult.push([new Date()]);
        console.info(`INFO: Email sent to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2})`);
      } catch (e) {
        sendResult.push([e.message]);
        console.error(`Failed to send email to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2}). Error: ${e.message}`);
      } finally {
        console.timeEnd(`Row '${rowIdx + 2}' processing time `);
      }
    } else {
      if (row[EMAIL_SENT_COL] !== '') {
        console.log(`Skipping Row ${rowIdx + 2} - Email already sent.`);
      }
      if (sheet.isRowHiddenByFilter(rowIdx + 2)) {
        console.log(`Skipping Row ${rowIdx + 2} - Row hidden by filter.`);
      }
      sendResult.push([row[EMAIL_SENT_COL]]);
    }
  });
  console.timeEnd("Total row processing time");

  sheet.getRange(2, emailSentColIdx + 1, sendResult.length).setValues(sendResult);
  console.log(`Finished writing outputs to rows.`);

  console.timeEnd(`sendEmails() processing time`);
}

/**
 * Retrieves a Gmail draft message by matching the subject line.
 * 
 * This function searches through the user's Gmail drafts for a message with a subject
 * line that matches the specified parameter. If a matching draft is found, it extracts
 * the subject, plain and HTML message body, and any attached files.
 *
 * @param {string} subject_line The subject line to search for the draft message.
 * @returns {{ message: { subject: string, text: string, html: string }, attachments: GoogleAppsScript.Gmail.GmailAttachment[], inlineImages: Object }} An object containing the subject, plain and HTML message body, and any attachments.
 * 
 * @throws {Error} Throws an error if no matching draft is found or if there is an issue accessing the drafts.
*/
function getGmailTemplateFromDrafts_(subject_line) {
  console.info(`Searching for draft with subject line: '${subject_line}'`);
  try {
    const drafts = GmailApp.getDrafts();
    console.log(`Total drafts retrieved: ${drafts.length}`);
    const draft = drafts.filter(subjectFilter_(subject_line))[0];
    if (!draft) {
      console.warn(`No draft found matching the subject line: '${subject_line}'`);
      return;
    } else {
      console.info(`Draft found with subject: '${draft.getMessage().getSubject()}'`);
    }
    const msg = draft.getMessage();

    const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
    console.log(`Total inline images found: ${allInlineImages.length}`);
    const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
    const htmlBody = msg.getBody();
    console.log(`Draft HTML body length: ${htmlBody.length}`);

    const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

    const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
    const matches = [...htmlBody.matchAll(imgexp)];
    console.log(`Total inline image matches found in HTML body: ${matches.length}`);

    const inlineImagesObj = {};
    matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
    console.info(`Returning message details for subject: '${subject_line}'`);

    return {
      message: {
        subject: subject_line,
        text: msg.getPlainBody(),
        html: htmlBody
      },
      attachments: attachments,
      inlineImages: inlineImagesObj
    };
  } catch (e) {
    console.error(`No Gmail draft found: '${e.message}'`);
    return;
  }
}

/**
 * Filters draft objects by matching the subject line.
 * 
 * This function returns a filter function that checks if a draft's subject 
 * matches the specified subject line.
 *
 * @param {string} subject_line The subject line to search for in the draft messages.
 * @returns {function(GoogleAppsScript.Gmail.GmailDraft): boolean} A function that takes a draft 
 * object and returns true if the draft's subject matches the subject line, otherwise false.
*/
function subjectFilter_(subject_line) {
  return function (element) {
    return element.getMessage().getSubject() === subject_line;
  }
}