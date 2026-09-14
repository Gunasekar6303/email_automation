# Job Application Email Automation

A Python-based job application email automation system that sends personalized emails to recruiters and HR professionals using the Gmail API.

The system reads recruiter information from an Excel spreadsheet, creates personalized emails, attaches a resume automatically, sends emails through Gmail API, and updates the application status in Excel.

---

## Project Overview

Applying for multiple jobs often involves repeatedly sending emails to recruiters and HR professionals with the same resume and application message.

To automate this repetitive process, this project provides a simple Python-based email automation workflow.

The system connects:

Excel → Python → Google OAuth 2.0 → Gmail API → Recruiter

After a successful email delivery request, the corresponding Excel status is updated to `SENT`.

If an error occurs, the status is updated to `FAILED`.

Already processed `SENT` entries are skipped to prevent duplicate emails.

---

## Features

- Read recruiter information from Excel
- Send personalized emails
- Automatically attach resume PDF
- Gmail API integration
- Google OAuth 2.0 authentication
- Track email status
- Prevent duplicate emails
- Skip already sent recipients
- Prevent sending emails to the sender's own email address
- Handle failed email attempts
- Save Excel status after every email
- Support batch email sending
- Local Python execution

---

## System Architecture

```text
                    ┌─────────────────────┐
                    │     Excel File      │
                    │                     │
                    │  HR Email           │
                    │  HR Name            │
                    │  Status             │
                    └──────────┬──────────┘
                               │
                               │ OpenPyXL
                               ▼
                    ┌─────────────────────┐
                    │    Python Script    │
                    │                     │
                    │  Read Excel         │
                    │  Validate Data      │
                    │  Check Status       │
                    │  Personalization    │
                    │  Create Email       │
                    │  Attach Resume      │
                    └──────────┬──────────┘
                               │
                               │ OAuth 2.0
                               ▼
                    ┌─────────────────────┐
                    │    Google Cloud     │
                    │                     │
                    │     Gmail API       │
                    │     OAuth 2.0       │
                    └──────────┬──────────┘
                               │
                               ▼
                    ┌─────────────────────┐
                    │       Gmail         │
                    └──────────┬──────────┘
                               │
                               ▼
                    ┌─────────────────────┐
                    │   HR / Recruiter    │
                    │                     │
                    │ Email + Resume      │
                    └─────────────────────┘

                               │
                               ▼
                    ┌─────────────────────┐
                    │    Excel Status     │
                    │                     │
                    │  SENT / FAILED      │
                    └─────────────────────┘



## End-to-End Workflow

1. Read HR details from Excel
                ↓
2. Check email address
                ↓
3. Check current status
                ↓
4. Skip if status = SENT
                ↓
5. Skip if recipient is sender's own email
                ↓
6. Create personalized email
                ↓
7. Attach resume PDF
                ↓
8. Encode email
                ↓
9. Send through Gmail API
                ↓
10. Update Excel status
                ↓
       ┌────────┴────────┐
       ↓                 ↓
    SUCCESS            ERROR
       ↓                 ↓
      SENT             FAILED



