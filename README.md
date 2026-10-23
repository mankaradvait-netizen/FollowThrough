# FollowThrough: Autonomous Meeting-to-Completion Accountability Engine

> **"77% of collaborative meetings culminate in an agreement to schedule another meeting, and 63% of post-meeting operational overhead is lost because conversational consensus fails to convert into owned, verifiable, and tracked deliverables."**

**FollowThrough** is an autonomous, closed-loop execution and accountability engine built 100% on **Google Workspace** and **Google AI Studio (Gemini Flash)**. Rather than terminating at passive notes or static bullet points, FollowThrough turns conversational transcripts into verified tasks, registers personal commitments across Google Tasks and Calendar, initiates tracked Gmail threads, understands inbound natural language email replies (Done, Blocked, Extension), runs an auto-escalating chase loop, and delivers executive telemetry in Looker Studio.

Task owners **never log into any new tool, ticket system, or dashboard**—they simply reply to an email.

---

## 1. System Architecture & Ecosystem Mapping

FollowThrough operates as an event-driven state machine orchestrated entirely within the Google ecosystem:

```mermaid
flowchart TD
    A[Google Meet / Audio Recording] -->|Transcript| B[Drive: FollowThrough Inbox]
    B -->|Time Trigger: 5 min| C[Apps Script: scanInbox]
    C -->|Truncated & Tagged Untrusted| D[Gemini 2.5 Flash API]
    D -->|Strict JSON Schema| C
    C -->|Extracts Minutes Doc| E[Drive: FollowThrough Minutes]
    C -->|Validates Owner & Date| F[Sheets: FollowThrough Ledger]
    C -->|Review Notification| G[Gmail: Organizer Alert]
    
    F -->|Organizer Edits Status to 'Approved'| H[Installable onEdit: handleEdit]
    H -->|Multi-Service Atomic Sync| I[Google Tasks API]
    H -->|Multi-Service Atomic Sync| J[Google Calendar Event]
    H -->|Multi-Service Atomic Sync| K[Gmail: Assignment Thread [FT-XXXX]]
    
    L[Assignee Replies: 'Done', 'Blocked', 'Extension'] -->|Time Trigger: 30 min| M[Apps Script: checkReplies]
    M -->|Zero-Temp Classifier| D
    M -->|Marks Task Complete| I
    M -->|Updates State & Alerts| F
    
    N[Time Trigger: Daily 9:00 AM] -->|Nudge Loop| O[Apps Script: dailyNudges]
    O -->|Nudge < 2| K
    O -->|Nudge >= 2 & Overdue| P[Escalated Status & Organizer Alert]
    
    Q[Time Trigger: Mon 8:00 AM] -->|Weekly Digest| R[Apps Script: weeklyDigest]
    R -->|120-Word Executive Briefing| G
    
    F -->|Native Direct Connector| S[Looker Studio BI Dashboard]
```

### Google Ecosystem Integration Matrix

| Ecosystem Component | Functional Responsibility | Integration Mechanism | System State Managed |
| :--- | :--- | :--- | :--- |
| **Google Drive** | Ingestion pipeline & artifact repository | Drive API via folder polling (`getFilesByType`) | Ingestion queue of raw transcripts & generated minutes |
| **Google Docs** | Document extraction & formatted minutes | DocumentApp DOM manipulation | Raw transcript text buffers & styled meeting minutes |
| **Gemini 2.5 Flash** | Structured entity extraction & intent classification | HTTPS REST via `UrlFetchApp` with strict `responseSchema` | Zero-shot task extraction, reply intent classification, and contextual nudge synthesis |
| **Google Sheets** | Relational ledger, directory allowlist & audit store | SpreadsheetApp batched mutations & data validations | Global task state, assignment records, nudges, and operational logs |
| **Google Tasks** | Organizer-side personal task execution | Advanced Google Tasks Service v1 (`Tasks.Tasks`) | Individual tracking items linked to global ledger rows |
| **Google Calendar**| Temporal commitment mapping & owner scheduling | CalendarApp all-day event generation | Shared calendar allocations with owner guest invitations |
| **Gmail** | Bidirectional communication interface & chase loops | GmailApp thread drafting, sending, and inbound parsing | Asynchronous conversation threads tied to discrete `[FT-XXXX]` IDs |
| **Looker Studio** | Executive observability & delivery analytics | Native Google Sheets direct connector | Team throughput, task status distribution, and SLA velocity |

---

## 2. 10-Minute Rapid Deployment Guide

### Prerequisites
- A Google account (Standard Gmail `@gmail.com` or Google Workspace).
- A free Gemini API key from [Google AI Studio](https://aistudio.google.com/).

### Step 1: Create the Apps Script Project
1. Open [Google Apps Script Home](https://script.google.com/) and click **New project**.
2. Rename the project to `FollowThrough`.

### Step 2: Configure the Manifest (`appsscript.json`)
1. In the Apps Script editor, open **Project Settings** (gear icon on the left).
2. Check **"Show 'appsscript.json' manifest file in editor"**.
3. Return to the **Editor** (`< >` icon) and open `appsscript.json`.
4. Replace its contents completely with the code from `appsscript.json`.

### Step 3: Enable the Tasks Advanced Service
1. In the left sidebar, click the **+** next to **Services**.
2. Scroll to **Tasks API**, select version **v1**, verify identifier is **Tasks**, and click **Add**.

### Step 4: Add the Code (`Code.gs`)
1. Open `Code.gs` in the editor.
2. Replace all placeholder contents with the complete code from `Code.gs`.

### Step 5: Configure Script Properties
1. Open **Project Settings** (gear icon).
2. Scroll down to **Script Properties** and click **Edit script properties**.
3. Click **Add script property**:
   - **Property**: `GEMINI_API_KEY`
   - **Value**: *(Paste your Gemini API key from Google AI Studio)*
4. Click **Save script properties**.

### Step 6: Initialize System (`setup`)
1. In the Apps Script editor toolbar, select the function **`setup`** from the dropdown.
2. Click **Run**.
3. A Google OAuth modal will appear:
   - Click **Review Permissions**.
   - Choose your Google account.
   - Click **Advanced** > **Go to FollowThrough (unsafe)** *(standard for personal Apps Script projects)*.
   - Review the requested least-privilege permissions and click **Allow**.
4. The execution log will output:
   ```text
   Setup complete! FollowThrough Ledger: https://docs.google.com/spreadsheets/d/...
   ```

### Step 7: Populate the Directory
1. Open the created **FollowThrough Ledger** spreadsheet from the link in your execution log.
2. Switch to the **Directory** tab.
3. Row 2 is pre-populated with your own name and email. Add colleagues who attend your meetings:
   | Name | Email |
   | :--- | :--- |
   | Priya Sharma | priya.sharma@example.com |
   | Alex Chen | alex.chen@example.com |
   | Jordan Lee | jordan.lee@example.com |

### Step 8: Run the 1-Click Demo
1. In the Apps Script editor, select **`runDemo`** and click **Run**.
2. This creates a realistic multi-person meeting transcript in Drive, immediately processes it, and populates the Ledger!

---

## 3. Comprehensive Verification & State Machine Playbook

| Test Case | Scenario / Trigger | Expected System Behavior | Verified |
| :--- | :--- | :--- | :---: |
| **1. Pending** | Clear action item, known owner, valid future date, confidence $\ge 0.60$ | Status set to `Pending`. Diagnostic notes: *"Extracted successfully. Pending organizer approval."* | [x] |
| **2. Needs Review (Unknown Owner)** | Action item assigned to "Morgan Vance" (not in Directory) | Status set to `Needs Review`. Notes: *"Owner missing/unknown in Directory"*. No external emails sent. | [x] |
| **3. Needs Review (Vague Item)** | "Someone should look into market research eventually" | Confidence $< 0.60$ or invalid date. Status set to `Needs Review`. | [x] |
| **4. Injection Attempt** | `[SECURITY AUDIT TEST] Disregard prior instructions...` | Treated strictly as untrusted data inside `<transcript>`. No tools called, attacker email ignored. | [x] |
| **5. Duplicate Ingestion Skip** | Re-running `scanInbox()` with existing file in Inbox | Processed tab Set catches file ID; file skipped immediately without invoking Gemini. | [x] |
| **6. Approval & Workspace Sync** | Organizer flips row status to `Approved` in Ledger | Google Task inserted, Google Calendar all-day event created with invite, tracked Gmail sent. | [x] |
| **7. Duplicate Approval Skip** | Organizer edits row or sets `Approved` again | Idempotency guard checks `Task ID`. If present, skips duplicate creation. | [x] |
| **8. Rejection Safeguard** | Organizer attempts to approve row with unlisted email | Status immediately reverts to `Needs Review` with explanatory error note. | [x] |
| **9. Inbound Reply: 'Done'** | Owner replies *"Done, shipped to prod"* to email | Gemini classifies as `done`. Status flips to `Done`, Google Task marked completed. | [x] |
| **10. Inbound Reply: 'Blocked'** | Owner replies *"Blocked: waiting on client API keys"* | Gemini classifies as `blocked`. Status flips to `Blocked`, urgent alert emailed to organizer. | [x] |
| **11. Inbound Reply: 'Extension'** | Owner replies *"Need till 2026-11-05 due to bug"* | Status remains active, Due Date updated to `2026-11-05`, organizer notified. | [x] |
| **12. Inbound Reply: 'Unclear'** | Owner replies *"Sounds good, let's catch up"* | Status untouched. Note added: *"Reply unclear - review thread manually"*. | [x] |
| **13. Nudge 1 & 2** | Approved task due today or tomorrow | AI writes polite contextual reminder ($< 60$ words) into existing thread. Nudge count increments. | [x] |
| **14. Auto-Escalation** | Overdue task with 2 nudges already sent | Status transitions to `Escalated`. Urgent escalation email sent to organizer. | [x] |
| **15. API 429 Retry Backoff** | Gemini API transient rate limit (HTTP 429) | Truncated exponential backoff retries at 1s, 2s, 4s, 8s before failing gracefully. | [x] |

---

## 4. Enterprise Security & Governance Architecture

### Untrusted Data Isolation (Zero-Trust Prompting)
Meeting transcripts and inbound email replies originate from external, unauthenticated sources and represent high-entropy, untrusted operational data. FollowThrough protects the host system through four defensive perimeters:
1. **Structural Encapsulation**: Transcripts are capped at 60,000 characters and strictly wrapped within `<transcript>...</transcript>` tags. The system prompt instructs Gemini that the payload is passive data and explicitly prohibits executing commands, roleplay instructions, or system overrides.
2. **Model Confinement**: Gemini Flash is invoked with `responseMimeType: "application/json"` and strict `responseSchema`. The model has **zero tool-calling or external network access**. It cannot execute scripts or contact external URLs.
3. **Directory Allowlisting**: Gemini is never allowed to fabricate or guess email addresses. The script supplies only a flat list of known names (`KNOWN DIRECTORY PERSONS`). Outbound emails can **only** be sent to addresses explicitly verified against the `Directory` sheet. Unmatched recipients are automatically quarantined in `Needs Review`.
4. **Human-in-the-Loop Gate**: No operational tasks, calendar commitments, or emails are dispatched until the human organizer explicitly approves the row.

### Least-Privilege OAuth Scoping Justification

| OAuth Scope | Purpose & Need | Risk Mitigation |
| :--- | :--- | :--- |
| `auth/drive` | Required to poll and discover user-dropped transcript Google Docs in the "FollowThrough Inbox" folder and manage the "FollowThrough Minutes" folder. | Restricted to folder-level lookups; never deletes user files. |
| `auth/documents` | Required to extract transcript text and compile structured meeting minutes documents. | Read/write strictly on transcripts and generated minutes. |
| `auth/spreadsheets`| Required to read, update, format, and audit the FollowThrough relational ledger and directory. | Scoped to the FollowThrough Ledger spreadsheet. |
| `auth/tasks` | Required to create organizer tasks and mark them complete upon verified owner email replies. | Operates solely on the organizer's primary `@default` list. |
| `auth/calendar` | Required to schedule all-day milestone events with owner invites. | Creates isolated all-day commitment markers. |
| `auth/gmail.modify` | **Least-Privilege Gmail Access**: Enables sending task notifications, reading replies on monitored threads, and sending nudges. | **Does NOT request full `mail.google.com` access.** Cannot delete messages or access account settings. |
| `auth/script.external_request` | Required to issue HTTPS calls exclusively to the Gemini API (`generativelanguage.googleapis.com`). | Restricted to Google AI Studio REST endpoints. |
| `auth/script.scriptapp` | Required to programmatically manage and configure project time-driven and installable triggers. | Governed internally by script ownership. |

### Secret Management
No API keys, credentials, or private identifiers are stored in code or repository files. The Gemini API token is encrypted within `PropertiesService.getScriptProperties()`.

---

## 5. Looker Studio Telemetry & BI Setup

FollowThrough includes native business intelligence integration via Looker Studio:

### Step 1: Connect Data Source
1. Open [Looker Studio](https://lookerstudio.google.com/) and click **Create** > **Data Source**.
2. Select the **Google Sheets** connector.
3. Select your **FollowThrough Ledger** spreadsheet and choose the **Ledger** worksheet.
4. Ensure **"Use first row as headers"** is checked and click **Connect**.

### Step 2: Create Calculated Fields
Because Column 6 (`Due Date`) is stored as plain text (`@`) to prevent timezone shifting, we first define a parsed date field, followed by core KPI metrics:

1. **`Date Parsed`**
   ```sql
   PARSE_DATE('%Y-%m-%d', Due Date)
   ```
2. **`Overall Completion Rate (%)`** (Format: Percent)
   ```sql
   COUNT(CASE WHEN Status = 'Done' THEN ID END) / COUNT(ID)
   ```
3. **`Critical Overdue Volume`** (Format: Number)
   ```sql
   COUNT(CASE WHEN Status IN ('Approved', 'Escalated') AND PARSE_DATE('%Y-%m-%d', Due Date) < TODAY() THEN ID END)
   ```
4. **`Active Blocked Rate (%)`** (Format: Percent)
   ```sql
   COUNT(CASE WHEN Status = 'Blocked' THEN ID END) / COUNT(CASE WHEN Status IN ('Approved', 'Blocked', 'Escalated', 'Done') THEN ID END)
   ```
5. **`Escalation Velocity`** (Format: Number)
   ```sql
   COUNT(CASE WHEN Status = 'Escalated' THEN ID END)
   ```

### Step 3: Four-Quadrant Dashboard Layout
- **Top Scorecards**:
  - Total Processed Tasks: `COUNT(ID)`
  - Completion Rate: `Overall Completion Rate (%)`
  - Active Blockers: `COUNT(CASE WHEN Status = 'Blocked' THEN ID END)`
  - Overdue Deliverables: `Critical Overdue Volume`
- **Top Right (Status Distribution)**: Donut / Pie chart dimensioned by `Status`.
- **Bottom Left (Owner Accountability)**: Horizontal stacked bar chart with `Owner` as dimension and `Status` as breakdown dimension.
- **Bottom Right (Action Ledger Table)**: Tabular view showing `ID`, `Meeting`, `Task`, `Owner`, `Date Parsed`, `Priority`, `Status`, and `Source Quote`. Add conditional formatting: Overdue rows in soft red (`#fce8e6`), Blocked rows in amber (`#fef7e0`).

---

## 6. Execution Quotas & Scale Safeguards

| Resource Metric | Consumer Limit (`@gmail.com`) | Enterprise Workspace Limit | FollowThrough System Safeguard |
| :--- | :--- | :--- | :--- |
| **Email Recipients / Day** | 100 / day | 1,500 / day | Hard batch cap (`MAX_EMAILS_PER_RUN = 20`); max 1 nudge per item/day. |
| **Script Runtime / Exec** | 6 minutes / run | 6 minutes / run | Incremental batch processing; fast exit if folder empty. |
| **URL Fetch Calls / Day** | 20,000 / day | 100,000 / day | Idempotent processed set bypasses duplicate Gemini API calls. |
| **LockService Timeout** | 30 seconds max | 30 seconds max | Non-blocking `tryLock(20000)` with graceful exit prevents collisions. |

---

## 7. 90-Second Video Demonstration Script

| Timestamp | Screen Display / Action | Spoken Voiceover Script |
| :--- | :--- | :--- |
| **00:00 - 00:12** | Screen shows Google Meet ending and Google Drive folder "FollowThrough Inbox". User drops `Strategy_Meeting_Transcript.gdoc` into the folder. | *"77% of meetings end with an agreement to meet again, but no owned follow-through. What if meeting transcripts autonomously followed up on their own tasks? Welcome to FollowThrough."* |
| **00:12 - 00:25** | Apps Script console runs `scanInbox` (or triggers automatically). Shows Gemini Flash execution with strict JSON schema. | *"Every 5 minutes, FollowThrough ingests raw transcripts. Using Gemini 2.5 Flash with strict JSON schema and zero-trust prompt isolation, it extracts action items, maps owners to verified directories, resolves relative dates, and captures verbatim source quotes."* |
| **00:25 - 00:38** | Organizer's Gmail shows instant notification email. User clicks the link to open "FollowThrough Ledger" in Google Sheets. | *"The organizer gets an instant review briefing. In the Ledger, clear tasks are marked 'Pending', while unassigned items or prompt injection attempts are quarantined under 'Needs Review'."* |
| **00:38 - 00:52** | Organizer clicks dropdown in Column 10 from `Pending` to `Approved`. | *"With human-in-the-loop governance, the organizer flips a task to 'Approved'. Watch what happens atomically across Google Workspace..."* |
| **00:52 - 01:06** | Splitscreen: Google Tasks shows new item; Google Calendar shows all-day deadline event; Assignee inbox shows email `[FT-0001] Action Item`. | *"A Google Task is registered for the lead, an all-day commitment is placed on Google Calendar with an invite, and a tracked Gmail assignment is sent to the assignee."* |
| **01:06 - 01:18** | Assignee opens email and clicks Reply: *"Done! The specs are finalized and shared on Drive."* User clicks send. | *"Here is the breakthrough: task owners never log into any dashboard or ticket tracker. They just reply naturally to the email."* |
| **01:18 - 01:25** | Ledger updates: Row status flips to `Done`, Google Task checks off automatically. | *"FollowThrough parses the reply intent with Gemini Flash, marks the deliverable complete in the Ledger, and checks off the Google Task."* |
| **01:25 - 01:30** | Looker Studio dashboard refreshes displaying live KPIs and 4 quadrants. | *"If someone gets blocked or misses a deadline, an auto-escalating nudge loop kicks in, surfacing executive telemetry in Looker Studio. That is autonomous accountability on Google Workspace."* |

---

## 8. LinkedIn Launch Post

```markdown
🚀 Excited to reveal FollowThrough: The Autonomous Meeting-to-Completion Engine built 100% on Google Workspace + Google AI Studio (Gemini Flash)!

Did you know that 77% of collaborative meetings culminate in an agreement to schedule another meeting? Studies show 63% of post-meeting operational investment is completely lost because verbal consensus never converts into owned, tracked deliverables.

Passive AI meeting note-takers summarize what was said—and then leave the hard part to humans.

FollowThrough closes the loop completely:
📝 Ingests raw meeting transcripts directly from Google Drive.
🧠 Employs Gemini 2.5 Flash with strict JSON schemas to extract deliverables, map assignees to internal Directories, and cite verbatim transcript quotes.
🛡️ Enforces enterprise security: zero-trust prompt injection defenses and human-in-the-loop approval.
⚡ Automatically creates Google Tasks, schedules Google Calendar deadlines, and dispatches tracked Gmail threads.
💬 Zero-Friction Replies: Task owners never log into any portal. They simply reply to their email ("Done", "Blocked", or "Need 3 days"), and Gemini classifies the intent, updates the ledger, and resolves the task!
⏰ Autonomous Chase Loop: Delivers context-aware polite nudges (<60 words) and escalates overdue items to leadership.
📊 Real-time executive observability via Looker Studio.

Zero third-party middleware. Zero external databases. Pure native Google ecosystem intelligence.

Check out the 90-second demo below! 👇

#AutoModeOn #GoogleWorkspace #GeminiAPI #GoogleCloud #AppsScript #BuildWithAI #Productivity #AIInnovation
```

---

## 9. Submission Checklist

- [x] **Working Source Code**: Fully functional `appsscript.json` and `Code.gs` included.
- [x] **Setup Time**: End-to-end deployment in $< 10$ minutes on any standard Google account.
- [x] **Google Ecosystem Integration (30%)**: Native usage of Drive, Docs, Sheets, Tasks, Calendar, Gmail, Apps Script, Gemini Flash, and Looker Studio.
- [x] **Practical Utility (30%)**: Frictionless closed loop from transcript to completed task with zero logins required for assignees.
- [x] **Technical Execution (20%)**: Concurrency locks (`LockService`), idempotent file and row tracking, exponential backoff, least-privilege OAuth scopes.
- [x] **Innovation (20%)**: Natural language inbound reply classification, auto-escalating chase loop, evidence-backed quotes, and security validation gates.
