/**
 * Master orchestrator function to process new emails from Gmail.
 */
function processNewEmails() {
  const ss = getSpreadsheet();
  const inboundLogsSheet = ss.getSheetByName(DATABASE.INBOUND_LOGS_SHEET_NAME);
  const approvalQueueSheet = ss.getSheetByName(DATABASE.APPROVAL_QUEUE_SHEET_NAME);

  if (!inboundLogsSheet || !approvalQueueSheet) {
    console.error('Database sheets not initialized. Please run initializeSheets() first.');
    return;
  }

  const inboundLabel = GmailApp.getUserLabelByName(CONFIG.INBOX_LABEL);
  const processedLabel = GmailApp.getUserLabelByName(CONFIG.PROCESSED_LABEL);
  const reviewLabel = GmailApp.getUserLabelByName(CONFIG.REVIEW_LABEL);

  if (!inboundLabel || !processedLabel || !reviewLabel) {
    console.error('Required Gmail labels are missing. Create: Inbound-Pending, Processed-By-AI, AI-Review-Required.');
    return;
  }

  const threads = GmailApp.search(`label:${CONFIG.INBOX_LABEL} -label:${CONFIG.PROCESSED_LABEL} is:unread`, 0, CONFIG.BATCH_SIZE);

  if (threads.length === 0) {
    console.log('No new emails to process.');
    return;
  }

  threads.forEach(thread => {
    const startTime = new Date();
    let aiResponse = null;
    let firstMessage = null;
    let txnId = `TXN-${Utilities.getUuid().replace(/-/g, '').substring(0, 8)}`;

    try {
      const messages = thread.getMessages();
      if (messages.length === 0) return;
      firstMessage = messages[0];

      const sender = firstMessage.getFrom();
      const subject = firstMessage.getSubject();
      const messageId = firstMessage.getId();
      const threadId = thread.getId();
      const cleanBody = firstMessage.getPlainBody();

      // Get Structured AI Analysis 
      aiResponse = analyzeEmailWithGemini(sender, subject, cleanBody);

      // --- APPLY BUSINESS POLICY AND RULES Safely ---
      const policyResult = applyOrganizationPolicies(sender, subject, aiResponse);
      let finalAction = aiResponse.recommended_action;
      
      if (policyResult.overriddenAction) {
        console.warn(`Policy Override: ${policyResult.ruleId}. Bypassing AI response to: ${policyResult.overriddenAction}.`);
        finalAction = policyResult.overriddenAction;
        aiResponse.justification = `[OVERRIDDEN BY ${policyResult.ruleId}]: ${policyResult.reason} | Original AI Logic: ${aiResponse.justification}`;
      }

      const endTime = new Date();
      const processingLatencyMs = endTime.getTime() - startTime.getTime();

      // Resolved Java native byte array mapping issues using standard helper functions
      const senderHash = computeSHA256(sender);
      const bodyHash = computeSHA256(cleanBody);

      const logEntry = [
        txnId,
        startTime.toISOString(),
        messageId,
        senderHash,
        subject,
        subject,
        bodyHash,
        cleanBody,
        policyResult.ruleId, 
        CONFIG.GEMINI_MODEL,
        aiResponse.category,
        aiResponse.confidence_score,
        JSON.stringify(aiResponse),
        '',
        '',
        'PENDING',
        processingLatencyMs,
        aiResponse.error_type || 'NONE',
        aiResponse.error_stack || '',
        0,
        'SYSTEM',
        endTime.toISOString()
      ];

      inboundLogsSheet.appendRow(logEntry);
      const currentRowInboundLogs = inboundLogsSheet.getLastRow();
      const statusColumnIndexInboundLogs = 16; 

      if (aiResponse.confidence_score >= CONFIG.CONFIDENCE_THRESHOLD && finalAction === 'AUTO_REPLY') {
        thread.replyAll(aiResponse.draft_response);
        thread.markRead();
        thread.removeLabel(inboundLabel);
        thread.addLabel(processedLabel);

        inboundLogsSheet.getRange(currentRowInboundLogs, statusColumnIndexInboundLogs).setValue('RESOLVED_AUTO');
      } else {
        const approvalEntry = [
          `Q-${Utilities.getUuid().replace(/-/g, '').substring(0, 8)}`,
          txnId,
          'PENDING_REVIEW',
          aiResponse.justification,
          finalAction,
          aiResponse.confidence_score,
          cleanBody.substring(0, Math.min(cleanBody.length, 500)),
          '',
          '',
          '',
          JSON.stringify(aiResponse),
          '',
          endTime.toISOString(),
          '',
          '',
          '',
          '' // assigned_to column starts empty
        ];
        approvalQueueSheet.appendRow(approvalEntry);

        thread.addLabel(reviewLabel);
        thread.removeLabel(inboundLabel);

        try {
          MailApp.sendEmail({
            to: CONFIG.OPERATIONAL_ALERT_EMAIL,
            subject: `HIGH PRIORITY: Email Review Required for Thread ID: ${threadId}`,
            body: `An email with subject '${subject}' from '${sender}' requires human review.\n\nReason: ${aiResponse.justification}\nAction: ${finalAction}\nConfidence: ${aiResponse.confidence_score}`
          });
        } catch(mailError) {
          console.error("Alert email failed to send: " + mailError.message);
        }

        inboundLogsSheet.getRange(currentRowInboundLogs, statusColumnIndexInboundLogs).setValue('QUEUED_FOR_HUMAN');
      }

      thread.removeLabel(inboundLabel);
      thread.addLabel(processedLabel);

    } catch (e) {
      console.error(`Error processing thread ${thread.getId()}:`, e.message, e.stack);
      const errorEndTime = new Date();
      const errorProcessingLatencyMs = errorEndTime.getTime() - startTime.getTime();

      let senderHashFallback = 'N/A';
      let bodyHashFallback = 'N/A';

      if (firstMessage) {
        try {
          senderHashFallback = computeSHA256(firstMessage.getFrom());
          bodyHashFallback = computeSHA256(firstMessage.getPlainBody());
        } catch(digestError) {
          console.error("Fallback digest failed:", digestError.message);
        }
      }

      const errorLogEntry = [
        txnId,
        startTime.toISOString(),
        firstMessage ? firstMessage.getId() : 'N/A',
        senderHashFallback,
        firstMessage ? firstMessage.getSubject() : 'N/A',
        firstMessage ? firstMessage.getSubject() : 'N/A',
        bodyHashFallback,
        firstMessage ? firstMessage.getPlainBody() : 'N/A',
        'SYSTEM_ERROR',
        'N/A',
        aiResponse ? aiResponse.category : 'Error',
        aiResponse ? aiResponse.confidence_score : 0.0,
        aiResponse ? JSON.stringify(aiResponse) : JSON.stringify({ error: e.message, stack: e.stack }),
        '',
        '',
        'ERROR',
        errorProcessingLatencyMs,
        'ORCHESTRATOR_ERROR',
        e.stack,
        0,
        'SYSTEM',
        errorEndTime.toISOString()
      ];
      inboundLogsSheet.appendRow(errorLogEntry);

      if (thread) {
        thread.removeLabel(inboundLabel);
        thread.addLabel(processedLabel);
      }
    }
  });

  console.log('Finished processing email batch.');
}