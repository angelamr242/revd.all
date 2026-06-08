/**
 * Automatically provisions database sheets and registers the creator.
 */
function initializeSheets() {
  const ss = getSpreadsheet();
  
  // 1. Inbound_Logs Setup
  let inboundSheet = ss.getSheetByName(DATABASE.INBOUND_LOGS_SHEET_NAME);
  if (!inboundSheet) {
    inboundSheet = ss.insertSheet(DATABASE.INBOUND_LOGS_SHEET_NAME);
  }
  const inboundHeaders = [
    "txn_id", "ingestion_ts", "gmail_message_id", "sender_hash", "subject_raw", 
    "subject_sanitized", "body_raw_hash", "body_sanitized", "policy_rule_id", 
    "ai_model_version", "ai_classification", "ai_confidence_score", "ai_raw_response_json", 
    "prompt_tokens", "completion_tokens", "execution_status", "execution_latency_ms", 
    "error_type", "error_stack_trace", "retry_count", "created_by", "last_updated_ts"
  ];
  if (inboundSheet.getLastRow() === 0) {
    inboundSheet.appendRow(inboundHeaders);
    inboundSheet.getRange(1, 1, 1, inboundHeaders.length).setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
  }

  // 2. Approval_Queue Setup
  let approvalSheet = ss.getSheetByName(DATABASE.APPROVAL_QUEUE_SHEET_NAME);
  if (!approvalSheet) {
    approvalSheet = ss.insertSheet(DATABASE.APPROVAL_QUEUE_SHEET_NAME);
  }
  const approvalHeaders = [
    "queue_id", "txn_id", "queue_status", "reason_for_queue", "ai_suggested_action", 
    "ai_confidence_score", "body_sanitized_preview", "human_decision", "human_override_action", 
    "reviewer_notes", "original_ai_response_json", "edited_prompt_used", "queued_ts", 
    "reviewed_ts", "reviewer_email", "review_latency_ms", "assigned_to"
  ];
  if (approvalSheet.getLastRow() === 0) {
    approvalSheet.appendRow(approvalHeaders);
    approvalSheet.getRange(1, 1, 1, approvalHeaders.length).setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
  }

  // 3. Users_Roles Directory
  let userSheet = ss.getSheetByName('Users_Roles');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users_Roles');
  }
  const userHeaders = ["user_email", "display_name", "role", "status"];
  if (userSheet.getLastRow() === 0) {
    userSheet.appendRow(userHeaders);
    userSheet.getRange(1, 1, 1, userHeaders.length).setFontWeight("bold").setBackground("#1E293B").setFontColor("#FFFFFF");
    
    // Auto-prepopulate project deployer as the initial system administrator
    const ownerEmail = Session.getEffectiveUser().getEmail();
    userSheet.appendRow([ownerEmail, "System Administrator", "ADMIN", "ACTIVE"]);
  }

  // Pre-trigger creation of Drives configuration keys
  getOrCreateDocTemplateId();
  getOrCreateReportsFolder();

  console.log("Database initialized successfully.");
}