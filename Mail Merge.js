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

  let RECIPIENT_COL = emailRecipients || getUserInput("Enter the header name for recipient email addresses:");
  let EMAIL_SENT_COL = emailSent || getUserInput("Enter the header name for email sent status:");

  if (!subjectLine) {
    subjectLine = Browser.inputBox("Mail Merge",
      "Type or copy/paste the subject line of the Gmail " +
      "draft message you would like to mail merge with:",
      Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine === "") {
      console.error(`ERROR: Abort script due to prompt response: '${subjectLine}'`);
      return;
    }
  }

  if (!thisSheet || !thisTab) {
    sheet = SpreadsheetApp.getActiveSheet();
  } else {
    sheet = SpreadsheetApp.openById(getIdFromUrl(thisSheet)).getSheetByName(String(thisTab));
    if (!sheet) {
      console.warn(`WARNING: Sheet named '${thisTab}' was not found in '${thisSheet}'.`);
      sheet = SpreadsheetApp.getActiveSheet();
      console.log(`DEBUG: Proceeding with the current active sheet as default: '${sheet}'`);
    }
  }

  console.info(`INFO: Sending mail merge from '${sheet.getName()}' with subject: '${subjectLine}'`);

  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  const heads = data.shift();

  if (!isValidHeaderRow(heads)) {
    console.error(`ERROR: Header row must contain at least two unique headers. Invalid headers: '${heads}'`);
    return;
  }

  console.log(`DEBUG: Headers array: '${heads}'`);

  if (!heads.includes(RECIPIENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${RECIPIENT_COL}'`);
    return;
  }

  if (!heads.includes(EMAIL_SENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${EMAIL_SENT_COL}'`);
    return;
  }

  console.log(`DEBUG: Ready to define emailSentColIdx: '${heads.indexOf(EMAIL_SENT_COL)}'`);

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  console.log(`DEBUG: Email sent column: '${emailSentColIdx}'`);

  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const out = [];

  console.time("Total row processing time");
  obj.forEach(function (row, rowIdx) {
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
        out.push([new Date()]);
        console.info(`INFO: Email sent to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2})`);
      } catch (e) {
        out.push([e.message]);
        console.error(`ERROR: Failed to send email to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2}). Error: ${e.message}`);
      } finally {
        console.timeEnd(`Row '${rowIdx + 2}' processing time `);
      }
    } else {
      if (row[EMAIL_SENT_COL] !== '') {
        console.log(`DEBUG: Skipping Row ${rowIdx + 2} - Email already sent.`);
      }
      if (sheet.isRowHiddenByFilter(rowIdx + 2)) {
        console.log(`DEBUG: Skipping Row ${rowIdx + 2} - Row hidden by filter.`);
      }
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  console.timeEnd("Total row processing time");

  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  console.log(`DEBUG: Finished writing outputs to rows.`);

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
    console.info(`INFO: Searching for draft with subject line: '${subject_line}'`);
    try {
      const drafts = GmailApp.getDrafts();
      console.log(`DEBUG: Total drafts retrieved: ${drafts.length}`);
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      if (!draft) {
        console.warn(`WARNING: No draft found matching the subject line: '${subject_line}'`);
        return;
      } else {
        console.info(`INFO: Draft found with subject: '${draft.getMessage().getSubject()}'`);
      }
      const msg = draft.getMessage();

      const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
      console.log(`DEBUG: Total inline images found: ${allInlineImages.length}`);
      const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
      const htmlBody = msg.getBody();
      console.log(`DEBUG: Draft HTML body length: ${htmlBody.length}`);

      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];
      console.log(`DEBUG: Total inline image matches found in HTML body: ${matches.length}`);

      const inlineImagesObj = {};
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
      console.info(`INFO: Returning message details for subject: '${subject_line}'`);

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
      console.error(`ERROR: No Gmail draft found: '${e.message}'`);
      return;
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
  }

  /**
   * Fills a template string with values from the provided data object.
   * 
   * This function replaces all occurrences of {{key}} in the template string 
   * with corresponding values from the data object. If a key does not exist in 
   * the object, it will be replaced with an empty string.
   * 
   * @param {string} template The template string containing {{}} markers for replacement.
   * @param {object} data An object with key-value pairs to replace the {{}} markers.
   * @returns {string} The filled template string with the values from the data object.
   *
   * @example
   * const result = fillInTemplateFromObject_('Hello, {{name}}!', { name: 'Alice' });
   * console.log(result); // Output: 'Hello, Alice!'
   */
  function fillInTemplateFromObject_(template, data) {
    let template_string = JSON.stringify(template);

    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      const text = data[key.replace(/[{}]+/g, "")] || "";
      return escapeData_(text.replace(/\n/g, '<br>'));
    });
    return JSON.parse(template_string);
  }

  /**
   * Escapes cell data to make it JSON-safe.
   * 
   * This function replaces special characters in a string with their escaped 
   * equivalents to ensure the string can be safely used in JSON contexts.
   *
   * @param {string} str The string to escape special characters.
   * @returns {string} The escaped string.
   *
   * @example
   * const safeString = escapeData_('This "string" contains special characters: \n and \t.');
   * console.log(safeString); // Output: This \"string\" contains special characters: \\n and \\t.
   */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  }

  /**
   * Checks if the provided header row is valid.
   * 
   * This function verifies that the header row contains at least two unique,
   * non-empty headers. It returns true if the conditions are met, otherwise false.
   *
   * @param {Array<string>} headers An array of header strings to validate.
   * @returns {boolean} Returns true if the header row is valid; otherwise false.
   */
  function isValidHeaderRow(headers) {
    const uniqueHeaders = new Set(headers.filter(header => header.trim() !== ""));
    return uniqueHeaders.size >= 2;
  }

  console.timeEnd(`sendEmails() processing time`);
}