# REVD.ALL — Workspace Optimizer & AI Operations Agent

`REVD.ALL` is a highly secure, multi-user AI-powered Operations Agent built inside Google Workspace using the Google Apps Script ecosystem and the Gemini 2.5 Flash API. 

Designed for student clubs, startups, and hectic workspaces, the system reads incoming emails, extracts user intents and key entities, enforces custom organization-defined policy constraints, and routes transactions safely to automated actions or an immersive human-in-the-loop review dashboard.

The frontend is styled after the **Immersive Minimalist WebGL aesthetic** (inspired by the Northgarden design studio), featuring slow-drifting liquid background mesh gradients, a 3D rotating cryptographic node engine, frosted glass cards (backdrop-blur), and responsive light/dark color configurations.

---

## 📸 Dashboard Preview

Here is a live look at the REVD.ALL Operator and Admin Control Deck:

<div align="center">
  <!-- REPLACE these paths with your committed image files once uploaded to your repo -->
  <img src="assets/dashboard-dark-view.png" width="800" alt="REVD.ALL Dark Mode Console View" style="border: 1px solid rgba(255,255,255,0.1); border-radius: 12px; margin-bottom: 12px;">
  <p><em>Figure 1: Immersive Dark Mode featuring the active review console, PII masked body, and the transaction log feed.</em></p>
</div>

<div align="center">
  <img src="assets/dashboard-light-view.png" width="800" alt="REVD.ALL Light Mode Paper View" style="border: 1px solid rgba(0,0,0,0.1); border-radius: 12px;">
  <p><em>Figure 2: Minimalist Bone-Cream Light Mode designed for high legibility under natural lighting conditions.</em></p>
</div>

---
<img width="1600" height="888" alt="image" src="https://github.com/user-attachments/assets/6e812b62-be05-475c-8f54-5e39d25d5217" />

<img width="1600" height="888" alt="image" src="https://github.com/user-attachments/assets/ce2da1cb-2853-4382-8523-16eebd5cf547" />

## 🛠️ System Architecture

```text
               [ Gmail Inbox ] ──( Labeled: Inbound-Pending )
                      │
                      ▼
             [ Orchestrator Engine ]
                      │
                      ▼
         [ Gemini 2.5 API Classifier ]  ◄──( Injects System Prompts )
                      │
                      ▼
        [ Organization Policies Engine ] ◄──( Evaluates POL-001 / POL-002 / POL-003 )
                      │
             ┌────────┴────────┐
             ▼                 ▼
     ( Auto-Reply )     ( Escalated / Overridden )
    [ Gmail API Send ]         │
                               ▼
                    [ Sheets Database Logs ]
                               │
                               ▼
                   [ Interactive Dashboard ] ◄──( Secured via Google OAuth SSO )
                               │
             ┌─────────────────┴─────────────────┐
             ▼                                   ▼
   [ Collaborative Steering ]          [ Silent Resolve Command ]
   ( Re-drafts via Gemini REST )         ( Archives thread & updates logs )
```
📋 Database Schemas

The database layer runs on Google Sheets, separated into three structural tabs:
1. Inbound_Logs

Maintains a detailed, immutable transaction log of all classified emails:
txn_id | ingestion_ts | gmail_message_id | sender_hash | subject_raw | subject_sanitized | body_raw_hash | body_sanitized | policy_rule_id | ai_model_version | ai_classification | ai_confidence_score | ai_raw_response_json | prompt_tokens | completion_tokens | execution_status | execution_latency_ms | error_type | error_stack_trace | retry_count | created_by | last_updated_ts
2. Approval_Queue

Handles manual verification records and assignment loads:
queue_id | txn_id | queue_status | reason_for_queue | ai_suggested_action | ai_confidence_score | body_sanitized_preview | human_decision | human_override_action | reviewer_notes | original_ai_response_json | edited_prompt_used | queued_ts | reviewed_ts | reviewer_email | review_latency_ms | assigned_to
3. Users_Roles

Controls workspace access permissions:
user_email | display_name | role (ADMIN, OPERATOR) | status (ACTIVE, INACTIVE)
⚙️ Installation & Configuration

Follow this sequential procedure to stand up the Operations Agent in your Workspace:
1. Setup Your Google Sheet

    Create a new Google Spreadsheet and copy its Spreadsheet ID from the address bar.

    Under Extensions, click Apps Script to open your cloud workspace editor.

2. Load the Repository Files

Create separate files in your Apps Script project matching the files in the src/ directory of this repository:
`
    Config.gs (Replace DATABASE.SPREADSHEET_ID with your sheet ID, and update CONFIG.OPERATIONAL_ALERT_EMAIL)

`    Database.gs

`    PolicyEngine.gs

 `   GeminiAPI.gs

`    Orchestrator.gs

`    RPCHandlers.gs

`    Sidebar.html (Paste the HTML design code)

3. Configure Credentials

    In the Apps Script sidebar, click the gear icon (⚙️ Project Settings).

    Scroll to Script Properties and click Add script property.

    Set the Property key to GEMINI_API_KEY and the value to your copied Google AI Studio key.

4. Create Gmail Folders

Create the following folders/labels in your Gmail client:
``
    Inbound-Pending

 `   Processed-By-AI

  `  AI-Review-Required

5. Initialize the Sheets

    Return to the code editor. Select the function initializeSheets from the toolbar dropdown menu.

    Click Run.

    Authorize Google permissions. This will generate your sheets and register your email as the first active ADMIN.

6. Deploy the Web App Console

    Click Deploy > New deployment (top right).

    Choose type: Web app.

    Set Execute as to: User accessing the web app (enabling secure Single-Sign-On).

    Set Who has access to: Anyone (or restricted to your Workspace domain).

    Deploy and copy the Web App URL. Open it in your browser.

🧪 Testing Playbook

To demonstrate and verify the system's operational integrity, execute these scenarios:
``
    High-Confidence Auto-Reply: Send an email about a Sponsorship under $1,000. It will automatically draft, reply, and log as RESOLVED_AUTO without human intervention.

`    `Policy Spending Limits (POL-001): Send an email asking for a Sponsorship with a value of $1,500. The agent will flag it, bypass auto-replies, and route it to the Approval_Queue with a "VIOLATED" limit trace.

 `   Silent Resolution (Opposite of Flagging): Open the escalated item on the web dashboard. Click "Resolve Without Reply" to close the case, update the sheet status to RESOLVED_DB, and remove labels from Gmail silently.

 `   Collaborative AI Steering: Open a pending item, write a feedback instruction in the steering panel, and click "Adjust Draft with AI". The draft textarea will update with an adjusted response incorporating your directions
<img width="1600" height="842" alt="image" src="https://github.com/user-attachments/assets/1e04535a-a88c-4cc0-ad1d-9eb769bedc37" />
<img width="1600" height="842" alt="image" src="https://github.com/user-attachments/assets/cc186c90-1537-47bb-b35e-2a05f87a2504" />

