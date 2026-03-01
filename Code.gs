Skip to content
balarmeet
bulk-email-sender-apps-script
Repository navigation
Code
Issues
Pull requests
Actions
Projects
Wiki
Security
Insights
Settings
bulk-email-sender-apps-script
/
Code.gs
in
main

Edit

Preview
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
45
46
47
48
49
50
51
52
53
54
55
56
57
58
59
60
61
62
63
64
65
66
67
68
69
70
71
72
73
74
75
76
77
78
79
80
81
82
83
84
85
86
87
88
89
90
91
92
93
94
95
96
97
98
99
100
101
102
103
104
105
106
107
108
109
110
111
112
113
114
115
116
117
118
119
120
121
122
123
124
125
126
127
128
129
130
131
132
133
134
135
136
137
138
139
140
141
142
143
144
145
146
147
148
149
150
151
152
153
154
155
156
157
158
159
160
161
162
163
164
165
166
167
168
169
170
171
172
173
174
175
176
177
178
179
180
181
182
183
184
185
186
187
188
189
190
191
192
193
194
195
196
197
198
199
200
201
202
203
204
205
206
207
208
209
210
211
212
213
214
215
216
217
218
219
220
221
222
223
224
225
226
227
228
229
230
231
232
233
234
235
236
237
238
239
240
241
242
243
244
245
246
247
248
249
250
251
252
253
254
255
256
257
258
259
260
261
262
263
264
265
266
267
268
269
270
271
272
273
274
275
276
277
278
279
280
281
282
283
284
285
286
287
288
289
290
291
292
293
294
295
296
297
298
299
300
301
302
303
304
305
306
307
308
309
310
311
312
313
314
315
316
317
318
319
320
321
322
323
324
325
326
327
328
329
330
331
332
333
334
335
336
337
338
339
340
341
342
343
344
345
346
347
348
349
350
351
352
353
354
355
356
357
358
359
360
361
362
363
364
365
366
367
368
369
370
371
372
373
374
375
376
377
378
379
380
381
382
383
384
385
386
387
388
389
390
391
392
393
394
395
396
397
398
399
400
401
402
403
404
405
406
407
408
409
410
411
412
413
414
415
416
417
418
419
420
421
422
423
424
425
426
427
428
429
430
431
432
433
434
435
436
437
438
439
440
441
442
443
444
445
446
447
448
449
450
451
452
453
454
455
456
457
458
459
460
461
462
463
464
465
466
467
468
469
470
471
472
473
474
475
476
477
478
479
480
481
482
483
484
485
486
487
488
489
490
491
492
493
494
495
496
497
498
499
500
501
502
503
504
505
506
507
508
509
510
511
512
513
514
515
516
517
518
519
520
521
522
523
524
525
526
527
528
529
530
531
532
533
534
535
536
537
538
539
540
541
542
543
544
545
546
547
548
549
550
551
552
553
554
555
556
557
558
559
560
561
562
563
564
565
566
567
568
569
570
/**
 * BULK EMAIL SENDER FOR GOOGLE SHEETS
 * 
 * This script sends personalized bulk emails using Google Docs as templates
 * and Google Sheets as the data source. Features include:
 * - HTML email templates with inline images
 * - Personalized placeholders ({{Name}}, {{v1}}, {{v2}}, {{v3}})
 * - PDF attachments support
 * - Error handling and status tracking
 * - Test email functionality
 * - Comprehensive UI notifications
 * 
 */

// ============================
// CONFIGURATION CONSTANTS
// ============================

var SPREADSHEET_ID = '1QYNR4BCAbzyjNoWvNMsFrh18w1ZLADQeIJJaWwkzEOk'; // Google Sheet ID
var MAIL_SHEET = 'mail'; // Google Sheet tab name
var START_ROW = 10; // First data row for bulk sending (Row 7 is for testing)

// ============================
// MENU SETUP
// ============================

/**
 * Creates custom menu in Google Sheets when opened
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Email Tools')
    .addItem('Send Bulk Emails', 'sendMailsFromSheet')
    .addItem('Send Test Email', 'test_mail')
    .addToUi();
}

// ============================
// UI NOTIFICATION FUNCTIONS
// ============================

/**
 * Shows a success alert with custom message
 * @param {string} title - Alert title
 * @param {string} message - Alert message
 */
function showSuccessAlert(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('✅ ' + title, message, ui.ButtonSet.OK);
}

/**
 * Shows an error alert with custom message
 * @param {string} title - Alert title
 * @param {string} message - Alert message
 */
function showErrorAlert(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('❌ ' + title, message, ui.ButtonSet.OK);
}

/**
 * Shows a warning alert with custom message
 * @param {string} title - Alert title
 * @param {string} message - Alert message
 */
function showWarningAlert(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('⚠️ ' + title, message, ui.ButtonSet.OK);
}

/**
 * Shows a progress dialog with custom message
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 */
function showProgressDialog(title, message) {
  var ui = SpreadsheetApp.getUi();
  ui.alert('⏳ ' + title, message, ui.ButtonSet.OK);
}

/**
 * Shows a confirmation dialog with custom message
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 * @returns {boolean} True if user confirms, false otherwise
 */
function showConfirmationDialog(title, message) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('❓ ' + title, message, ui.ButtonSet.YES_NO);
  return response == ui.Button.YES;
}

/**
 * Shows a detailed results dialog after bulk sending
 * @param {number} successCount - Number of successful emails
 * @param {number} errorCount - Number of failed emails
 * @param {number} totalCount - Total number of emails processed
 * @param {Array<string>} errorDetails - Array of error messages
 */
function showResultsDialog(successCount, errorCount, totalCount, errorDetails) {
  var ui = SpreadsheetApp.getUi();
  var message = '📊 Email Sending Results:\n\n';
  message += '✅ Successfully sent: ' + successCount + '\n';
  message += '❌ Failed: ' + errorCount + '\n';
  message += '📋 Total processed: ' + totalCount + '\n\n';
  
  if (errorCount > 0) {
    message += 'Error Details:\n';
    errorDetails.forEach(function(error, index) {
      message += (index + 1) + '. ' + error + '\n';
    });
    message += '\nCheck column I for detailed error messages.';
  }
  
  ui.alert('🎯 Bulk Send Complete', message, ui.ButtonSet.OK);
}

// ============================
// UTILITY FUNCTIONS
// ============================

/**
 * Marks a row as successfully sent in the spreadsheet
 * @param {Sheet} sheet - The spreadsheet sheet object
 * @param {number} row - Row number to mark as sent
 */
function markAsSent(sheet, row) {
  sheet.getRange(row, 7).setValue("YES"); // Column G: Sent Status
  sheet.getRange(row, 8).setValue(getFormattedDateTime()); // Column H: Date Sent
  sheet.getRange(row, 9).clearContent(); // Column I: Clear previous errors
}

/**
 * Marks a row with an error message in the spreadsheet
 * @param {Sheet} sheet - The spreadsheet sheet object
 * @param {number} row - Row number to mark as error
 * @param {string} errorMsg - Error message to display
 */
function markAsError(sheet, row, errorMsg) {
  sheet.getRange(row, 7).setValue("ERROR"); // Column G: Sent Status
  sheet.getRange(row, 8).setValue(getFormattedDateTime()); // Column H: Date Sent
  sheet.getRange(row, 9).setValue(errorMsg); // Column I: Error Message
}

/**
 * Returns formatted current date and time
 * @returns {string} Formatted date and time string
 */
function getFormattedDateTime() {
  return Utilities.formatDate(
    new Date(), 
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 
    "yyyy-MM-dd HH:mm:ss"
  );
}

// ============================
// DOCUMENT PROCESSING FUNCTIONS
// ============================

/**
 * Processes Google Doc to HTML, converts images to inline attachments,
 * and removes document title
 * @param {string} docId - Google Document ID
 * @returns {Object} Processed HTML content and inline images
 */
function getProcessedDocHtml(docId) {
  try {
    // Get OAuth token for API calls
    var token = ScriptApp.getOAuthToken();
    
    // Export Google Doc as HTML using Drive API
    var exportUrl = 'https://www.googleapis.com/drive/v3/files/' + 
                   encodeURIComponent(docId) + '/export?mimeType=text/html';
    
    var response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    // Check if export was successful
    if (response.getResponseCode() !== 200) {
      throw new Error('Failed to export document: ' + response.getResponseCode());
    }

    var htmlContent = response.getContentText();
    
    // Extract body content only (remove HTML head, styles, etc.)
    var bodyMatch = htmlContent.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    if (bodyMatch && bodyMatch[1]) {
      htmlContent = bodyMatch[1];
    }
    
    // Remove document title (first H1 element)
    htmlContent = htmlContent.replace(/^\s*<h1[^>]*>[\s\S]*?<\/h1>\s*/i, '');
    htmlContent = htmlContent.replace(/^\s+/, ''); // Trim leading whitespace
    
    // Process inline images: replace public URLs with cid (Content ID)
    var inlineImages = {};
    var imageIndex = 0;
    
    var processedHtml = htmlContent.replace(
      /<img[^>]+src=(["'])([^"']+)\1[^>]*>/gi, 
      function(match, quote, src) {
        // Skip data URIs (base64 encoded images)
        if (/^\s*data:/i.test(src)) return match;
        
        try {
          var fetchOptions = { 
            muteHttpExceptions: true,
            followRedirects: true
          };
          
          // Add authorization for Google domains
          if (/^(https?:\/\/)?(lh3\.googleusercontent\.com|drive\.google\.com|docs\.google\.com)/i.test(src)) {
            fetchOptions.headers = { Authorization: 'Bearer ' + token };
          }
          
          var imageResponse = UrlFetchApp.fetch(src, fetchOptions);
          if (imageResponse.getResponseCode() !== 200) {
            Logger.log('Failed to fetch image: ' + src);
            return match;
          }
          
          var blob = imageResponse.getBlob();
          var cid = 'img_' + (++imageIndex);
          inlineImages[cid] = blob;
          
          // Replace src with cid reference for inline attachment
          return match.replace(src, 'cid:' + cid);
        } catch (e) {
          Logger.log('Error inlining image: ' + e.toString());
          return match;
        }
      }
    );

    Logger.log('✅ Successfully processed document. Images inlined: ' + imageIndex);
    
    return {
      htmlBody: processedHtml,
      inlineImages: inlineImages,
      success: true,
      message: 'Document processed successfully with ' + imageIndex + ' images'
    };
    
  } catch (error) {
    Logger.log('❌ Error in getProcessedDocHtml: ' + error.toString());
    
    // Fallback: return plain text version if HTML processing fails
    try {
      var doc = DocumentApp.openById(docId);
      var plainText = doc.getBody().getText();
      
      return {
        htmlBody: '<div style="font-family: Arial, sans-serif; line-height: 1.6;">' + 
                  plainText.replace(/\n/g, '<br>') + '</div>',
        inlineImages: {},
        success: false,
        message: 'Used fallback plain text: ' + error.toString()
      };
    } catch (fallbackError) {
      return {
        htmlBody: '<p style="color: red;">Error loading document content.</p>',
        inlineImages: {},
        success: false,
        message: 'Complete failure: ' + fallbackError.toString()
      };
    }
  }
}

// ============================
// MAIN BULK EMAIL FUNCTION
// ============================

/**
 * Main function to send bulk emails from spreadsheet data
 * Processes all rows starting from START_ROW
 */
function sendMailsFromSheet() {
  var successCount = 0, errorCount = 0;
  var errorDetails = [];
  
  try {
    var spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = spreadsheet.getSheetByName(MAIL_SHEET);
    if (!sheet) {
      showErrorAlert('Sheet Not Found', 'Sheet "' + MAIL_SHEET + '" not found in the spreadsheet.');
      return;
    }

    // Get configuration values from sheet
    var subject = sheet.getRange("B1").getValue();
    var docId = sheet.getRange("B2").getValue();
    
    // Validate required fields
    if (!subject) {
      showErrorAlert('Missing Subject', 'Subject is empty in cell B1. Please add an email subject.');
      return;
    }
    if (!docId) {
      showErrorAlert('Missing Document ID', 'Document ID is empty in cell B2. Please add a Google Doc ID.');
      return;
    }

    // Show confirmation dialog before starting bulk send
    var confirmation = showConfirmationDialog(
      'Confirm Bulk Email Send',
      'This will send emails to all recipients starting from row ' + START_ROW + '.\n\n' +
      '• Subject: ' + subject + '\n' +
      '• Document ID: ' + docId + '\n\n' +
      'Do you want to continue?'
    );
    
    if (!confirmation) {
      showWarningAlert('Operation Cancelled', 'Bulk email sending was cancelled by user.');
      return;
    }

    // Show progress dialog
    showProgressDialog('Processing Template', 'Processing Google Doc template and preparing emails...');

    // Process Google Doc template to HTML
    var docContent = getProcessedDocHtml(docId);
    if (!docContent.success) {
      showWarningAlert('Template Processing Warning', docContent.message);
    }

    var lastRow = sheet.getLastRow();
    var totalEmails = 0;

    // Count total emails to be sent
    for (var row = START_ROW; row <= lastRow; row++) {
      var email = sheet.getRange(row, 1).getValue();
      var sentStatus = sheet.getRange(row, 7).getValue();
      if (email && sentStatus !== "YES") {
        totalEmails++;
      }
    }

    if (totalEmails === 0) {
      showWarningAlert('No Emails to Send', 'No pending emails found to send. All rows may already be marked as sent.');
      return;
    }

    // Process each row in the spreadsheet
    for (var row = START_ROW; row <= lastRow; row++) {
      try {
        // Get row data from columns A-F
        var email = sheet.getRange(row, 1).getValue();
        var name = sheet.getRange(row, 2).getValue();
        var attachmentId = sheet.getRange(row, 3).getValue(); 
        var v1 = sheet.getRange(row, 4).getValue(); 
        var v2 = sheet.getRange(row, 5).getValue(); 
        var v3 = sheet.getRange(row, 6).getValue();               
        var sentStatus = sheet.getRange(row, 7).getValue();

        // Skip if no email or already sent
        if (!email || sentStatus === "YES") continue;

        // Prepare attachment if provided
        var file = null;
        if (attachmentId) {
          try {
            file = DriveApp.getFileById(attachmentId);
          } catch (fileError) {
            Logger.log('Error accessing attachment for row ' + row + ': ' + fileError.toString());
            // Continue without attachment
          }
        }  

        // Personalize HTML template with recipient data
        var personalizedHtml = personalizeTemplate(
          docContent.htmlBody, 
          name, 
          v1, 
          v2, 
          v3
        );

        // Build email options
        var mailOptions = {
          to: email,
          subject: subject,
          htmlBody: personalizedHtml,
          inlineImages: docContent.inlineImages
        };

        // Add attachment if available
        if (file) {
          mailOptions.attachments = [file.getAs(MimeType.PDF)];
        }

        // Send email
        MailApp.sendEmail(mailOptions);

        // Mark as sent and increment success count
        markAsSent(sheet, row);
        successCount++;
        
        // Respect Gmail rate limits (2 second pause between emails)
        Utilities.sleep(2000); 
        
      } catch (rowError) {
        // Handle individual row errors
        errorCount++;
        var errorMsg = 'Row ' + row + ': ' + rowError.toString();
        Logger.log('Error sending row ' + row + ': ' + errorMsg);
        errorDetails.push(errorMsg);
        markAsError(sheet, row, rowError.toString());
      }
    }
    
    // Show final results
    showResultsDialog(successCount, errorCount, successCount + errorCount, errorDetails);
    
    // Log final results
    Logger.log("Bulk send completed: " + successCount + " sent, " + errorCount + " failed.");
    
  } catch (err) {
    Logger.log("Critical Error in sendMailsFromSheet: " + err);
    showErrorAlert('Critical Error', 'A critical error occurred: ' + err.toString());
  }
}

// ============================
// TEMPLATE PERSONALIZATION
// ============================

/**
 * Replaces placeholder variables in template with actual values
 * @param {string} template - HTML template string
 * @param {string} name - Recipient name
 * @param {string} v1 - Variable 1 value
 * @param {string} v2 - Variable 2 value
 * @param {string} v3 - Variable 3 value
 * @returns {string} Personalized HTML content
 */
function personalizeTemplate(template, name, v1, v2, v3) {
  var personalized = template;
  
  // Replace all placeholders with their values
  if (personalized.includes("{{Name}}")) {
    personalized = personalized.replace(/{{\s*Name\s*}}/gi, name || "");
  }
  if (personalized.includes("{{v1}}")) {
    personalized = personalized.replace(/{{\s*v1\s*}}/gi, v1 || "");
  }
  if (personalized.includes("{{v2}}")) {
    personalized = personalized.replace(/{{\s*v2\s*}}/gi, v2 || "");
  }
  if (personalized.includes("{{v3}}")) {
    personalized = personalized.replace(/{{\s*v3\s*}}/gi, v3 || "");
  }
  
  return personalized;
}

// ============================
// TEST EMAIL FUNCTION
// ============================

/**
 * Sends a test email using data from Row 7
 * Useful for testing template and configuration before bulk send
 */
function test_mail() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MAIL_SHEET);
    if (!sheet) {
      showErrorAlert('Sheet Not Found', 'Sheet "' + MAIL_SHEET + '" not found.');
      return;
    }

    // Get configuration values
    var subject = sheet.getRange("B1").getValue();
    var docId = sheet.getRange("B2").getValue();
    
    // Validate required fields
    if (!subject) {
      showErrorAlert('Missing Subject', 'Subject is empty in cell B1. Please add an email subject.');
      return;
    }
    if (!docId) {
      showErrorAlert('Missing Document ID', 'Document ID is empty in cell B2. Please add a Google Doc ID.');
      return;
    }

    // Get test data from Row 7
    var email = sheet.getRange("A7").getValue();
    var name = sheet.getRange("B7").getValue();
    var attachmentId = sheet.getRange("C7").getValue();
    var v1 = sheet.getRange("D7").getValue();
    var v2 = sheet.getRange("E7").getValue();
    var v3 = sheet.getRange("F7").getValue();

    if (!email) {
      showErrorAlert('Test Email Required', 'Please enter a test email address in cell A7.');
      return;
    }

    // Show progress dialog
    showProgressDialog('Sending Test Email', 'Processing template and sending test email to: ' + email);

    // Process Google Doc template
    var docContent = getProcessedDocHtml(docId);
    if (!docContent.success) {
      showWarningAlert('Template Processing Warning', docContent.message);
    }

    // Prepare attachment if provided
    var file = null;
    if (attachmentId) {
      try {
        file = DriveApp.getFileById(attachmentId);
      } catch (fileError) {
        showWarningAlert('Attachment Warning', 'Could not access attachment: ' + fileError.toString());
      }
    }  

    // Personalize template
    var personalizedHtml = personalizeTemplate(
      docContent.htmlBody, 
      name, 
      v1, 
      v2, 
      v3
    );

    // Build email options
    var mailOptions = {
      to: email,
      subject: subject,
      htmlBody: personalizedHtml,
      inlineImages: docContent.inlineImages
    };

    // Add attachment if available
    if (file) {
      mailOptions.attachments = [file.getAs(MimeType.PDF)];
    }

    // Send test email
    MailApp.sendEmail(mailOptions);

    // Mark test as sent in sheet
    sheet.getRange("G7").setValue("YES");
    sheet.getRange("H7").setValue(getFormattedDateTime());

    // Show success message
    showSuccessAlert(
      'Test Email Sent Successfully', 
      '✅ Test email sent to: ' + email + '\n\n' +
      '📧 Subject: ' + subject + '\n' +
      '📄 Document: ' + docId + '\n' +
      '🖼️ Images processed: ' + Object.keys(docContent.inlineImages).length + '\n' +
      '📎 Attachment: ' + (file ? 'Yes' : 'No') + '\n\n' +
      'Template processing: ' + docContent.message
    );
    
    Logger.log("✅ Test email sent successfully to: " + email);
    Logger.log("✅ Document processing: " + docContent.message);
    
  } catch (err) {
    Logger.log("❌ Error in test_mail: " + err);
    showErrorAlert('Test Email Failed', 'Failed to send test email: ' + err.toString());
  }
}
 
Copied!
