This project is an end-to-end automation built in Microsoft Excel using **Power Query** for data preparation and **VBA** for orchestration and email distribution.

It was developed in April 2024.

# Overview 🔍

The solution automates the full weekly reporting cycle:

1. Opens country-specific report files stored in a structured folder.
2. Refreshes multiple Power Query connections.
3. Updates report metadata (e.g. refresh date).
4. Saves and renames the refreshed file with the current date.
5. Sends the report via Outlook to predefined recipients.
6. Logs execution details (date, time, user) in a journal sheet.

The macro dynamically reads configuration (recipients, subject, body, send flag, scheduled day) from a control sheet, enabling non-technical users to manage distribution logic without modifying the code

# Architecture 💡

## Data Layer

- Power Query connections:
    - Query - Special Requests
    - Query - Open Requests
    - Query - Closed Requests
    - Query - Clarifications
- Refresh handled programmatically via VBA.
- FastCombine enabled to suppress privacy-level prompts and improve refresh performance.

## Control Layer

- List sheet: email metadata and scheduling logic.
- Dashboard sheet: sender mailbox configuration.
- Journal sheet: execution logging.

## Execution Layer (VBA)

- Iterates over configured recipients.
- Validates:
    - Attachment existence
    - Send flag = “YES”
    - Day-of-week condition
- Refreshes report files.
- Renames outputs with timestamp.
- Sends emails via Outlook (Outlook.Application).
- Writes audit entry after completion.

# Key Features ✔️

- Fully automated weekly reporting workflow.
- Config-driven email distribution.
- Automatic query refresh before dispatch.
- Dynamic file renaming with current date.
- Basic execution logging (date, time, username).
- Performance optimization (manual calculation, screen updating disabled).

# Business Impact 💬

- Eliminated manual refresh and distribution process.
- Reduced risk of sending outdated data.
- Standardized weekly reporting across multiple stakeholders.
- Enabled scalable distribution without increasing operational workload.

# Technical Stack 🛠️

- Microsoft Excel
- Power Query (M)
- VBA
- Microsoft Outlook (COM automation)

# Notes 📑

- The project requires a predefined folder structure (/Countries directory).
- Outlook must be installed and accessible via COM.
- Future improvement: structured error handling and logging of failures.

This was one of my first professional automation projects and represents an early example of integrating data transformation (Power Query) with process automation (VBA) to deliver a complete reporting pipeline inside Excel.
