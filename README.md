# Team Case Counter (Outlook add-in)

An Outlook add-in for a support team working out of a shared Office 365 mailbox. It shows how many cases each team member (anyone with an address in your company's email domain, which you set in the add-in) handled:

- **Last 7 days**, **last 30 days**, or **a calendar month** of your choice
- A **case** is an email thread (Outlook conversation)
- If a person replies to the same thread several times, it counts as **1 case** for them
- If several people reply on the same thread, **each of them** gets 1 case

**No admin approval?** If your organisation won't approve the add-in's access to mail, use the [Outlook desktop macro](outlook-desktop-macro/README.md) instead. It works in classic Outlook for Windows with no app registration or hosting.

The pane lists each person with their case count and total messages sent. Click a name to see the threads behind the count, and use **Download CSV** to export the table.

## How it works

The add-in uses Microsoft Graph to read every message in the shared mailbox (all folders, including Sent Items) in the chosen period. It keeps only the messages sent by a team member and groups them by `conversationId`, so each person gets one case per distinct thread they replied on. Customer messages and drafts are ignored. A thread counts in a period if the person sent at least one message on it during that period.

The counting rules are in `src/core/counting.js` and are covered by tests in `test/`.

### How replies must be sent

The add-in can only credit a reply to a person if Outlook records who sent it:

| How the team replies | Counted? |
| --- | --- |
| From their own address, with the shared mailbox in To/CC | Yes |
| **Send on Behalf** of the shared mailbox | Yes (credited to the person in the `Sender` field) |
| **Send As** the shared mailbox | **No.** Exchange records the shared mailbox as the sender, so the person can't be identified. The pane shows how many such replies it skipped. |

If your team uses Send As, an Exchange admin can switch them to Send on Behalf. For replies sent from personal mailboxes to show up, they need to copy the shared mailbox. For Send on Behalf replies to be saved in the shared mailbox's Sent Items, the admin should run:

```powershell
Set-Mailbox support@yourcompany.com -MessageCopyForSendOnBehalfEnabled $true
```

## Install for yourself only (no Node.js needed)

A ready-built copy of the add-in is in `docs/` and is served by GitHub Pages at `https://urzahil.github.io/Outlook-Monitoring-Plugin/`. Adding it from a file installs it in your Outlook only; nobody else sees it. The hosted page holds no data, IDs or settings. Your mail only travels between Outlook and Microsoft.

1. **Turn on GitHub Pages (once):** in this repo on GitHub, open **Settings → Pages**. Under **Build and deployment**, choose **Deploy from a branch**, pick this branch and the **/docs** folder, and click **Save**. After a minute or two, `https://urzahil.github.io/Outlook-Monitoring-Plugin/taskpane.html` should load.
2. **Register the app in Entra ID:** follow [step 1 below](#1-register-the-app-in-microsoft-entra-id), using these redirect URIs:
   - `https://urzahil.github.io/Outlook-Monitoring-Plugin/taskpane.html`
   - `brk-multihub://urzahil.github.io`
3. **Add it to Outlook:** download [`manifest-github-pages.xml`](manifest-github-pages.xml) (on GitHub, open the file and click **Download raw file**). In Outlook on the web or the new Outlook, go to **Apps → Add apps** (or **Get Add-ins**), then **My add-ins → Add a custom add-in → Add from file**, and choose that file. In classic Outlook for Windows, **File → Manage Add-ins** opens the same screen.
4. **First run:** open any email and click **Team cases**. Open **Settings**, enter the app (client) ID, the directory (tenant) ID, the shared mailbox address and your team's email domain, then click **Save settings**.

After changing anything in `src/`, run `npm run build:site` and commit `docs/` and `manifest-github-pages.xml` to update the hosted copy.

## Setup

### 1. Register the app in Microsoft Entra ID

1. In the [Entra admin center](https://entra.microsoft.com), go to **App registrations → New registration**. Name it "Team Case Counter" and choose **Accounts in this organizational directory only**.
2. Under **Authentication → Add a platform → Single-page application**, add these redirect URIs (replace the host with wherever you host the add-in):
   - `brk-multihub://localhost:3000`
   - `https://localhost:3000/taskpane.html`
3. Under **API permissions**, add the Microsoft Graph **delegated** permission `Mail.Read.Shared`, then click **Grant admin consent**.
4. Copy the **Application (client) ID** and **Directory (tenant) ID**. You enter them under **Settings** in the add-in's panel.

Each person using the add-in needs **Full Access** to the shared mailbox, which your support team members normally already have.

### 2. Run it locally

Requires Node.js 18 or later.

```sh
npm install
npm start          # serves https://localhost:3000 (installs a trusted dev certificate on first run)
```

Then sideload `manifest.xml` in Outlook: in Outlook on the web or the new Outlook, open **Get Add-ins → My add-ins → Add a custom add-in → Add from file**.

Open any email and click **Team cases** in the ribbon (or the **Apps** menu). The first time, open **Settings**, enter the app (client) ID, the directory (tenant) ID, the shared mailbox address (for example `support@yourcompany.com`) and your team's email domain (for example `yourcompany.com`), and click **Save settings**.

### 3. Deploy it for the team

1. Run `npm run build` and host the `dist/` folder on any HTTPS static host (Azure Static Web Apps, an internal web server, and so on).
2. In `manifest.xml`, replace every `https://localhost:3000` with your host's URL. Also add the matching redirect URIs (`brk-multihub://<your-host>` and `https://<your-host>/taskpane.html`) to the app registration.
3. In the Microsoft 365 admin center, go to **Settings → Integrated apps → Upload custom apps** and deploy `manifest.xml` to your team.

## Development

```sh
npm test           # unit tests for counting and date ranges
npm run build      # production bundle in dist/
```

## Notes

- Dates use your local time zone. A calendar month runs from the 1st at 00:00 to the 1st of the next month; the current month stops at today.
- A large mailbox may take a few seconds to scan a 30-day window. The pane shows progress, and Graph throttling is retried automatically.
- To leave out team addresses that aren't people (for example `noreply@yourcompany.com`), add them under **Settings → Addresses to ignore**. The shared mailbox address is always ignored.
