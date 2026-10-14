# MAHY KHOORY PMS Workflow Web App

This is a server-backed PMS application. Unlike the earlier standalone HTML prototype, it stores evaluations in SQLite, emails each approver in sequence, records approval history, creates a filled PDF after final approval, and exports all captured fields to Excel.

## Setup

1. Install Node.js 20 or later.
2. Copy `.env.example` to `.env` and set the Microsoft Entra application values.
3. In Entra ID, create an app registration and grant Microsoft Graph **Mail.Send application permission**, then grant admin consent.
4. Set `MAIL_SENDER` to the licensed/shared mailbox used to send PMS notifications.
5. Set `BASE_URL` to the HTTPS address reachable by approvers. Approval links cannot work for other users when BASE_URL is localhost.
6. Run:

   npm install
   npm start

7. Open the BASE_URL in a browser.

## Production notes

- Put the app behind HTTPS and Microsoft Entra sign-in before production use.
- Restrict Graph application mailbox access to the PMS sender mailbox using your tenant's Exchange application access controls.
- Back up `data/pms.db` and the `uploads` folder.
- The app accepts up to 10 attachments, 15 MB each.
- Approval links are random one-time-stage tokens. Once a step advances, the previous link no longer works.

## Outputs

- Final PDF: shown in the register after every approver approves and also emailed to the initiator.
- Excel: **Export all fields to Excel** includes employee name, every form field, ratings/comments, workflow, current status, and full approval history.
