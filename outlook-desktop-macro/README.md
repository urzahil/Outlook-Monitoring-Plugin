# Team Case Counter: Outlook desktop macro

A version of the case counter that runs as a macro inside **classic Outlook for Windows**. It reads the shared mailbox through your own Outlook, so it needs no app registration, no admin approval and no web hosting.

It counts cases the same way as the add-in:

- A **case** is an email thread (Outlook conversation).
- If a person replies to the same thread several times, it counts as **1 case** for them.
- If several people reply on the same thread, **each of them** gets 1 case.

The results open in your web browser as a page with one row per person. Click a name to see their threads. A link on the page opens the same table as a spreadsheet (CSV).

## Requirements

- Classic Outlook for Windows. Macros don't run in the new Outlook, Outlook on the web or Outlook for Mac.
- The shared mailbox shows in your Outlook folder list. It does if you have Full Access to it.
- Macros allowed in Outlook (see step 1). If your company has locked macro settings, ask IT.

## Install (once)

1. **Allow macros.** In Outlook, go to **File → Options → Trust Center → Trust Center Settings… → Macro Settings**. Choose **Notifications for all macros**, click **OK**, and restart Outlook.
2. **Download the macro.** Open [`TeamCaseCounter.bas`](TeamCaseCounter.bas) on GitHub and click **Download raw file**.
3. **Import it.** In Outlook, press **Alt + F11** to open the VBA editor. Go to **File → Import File…**, choose `TeamCaseCounter.bas`, then **File → Save** and close the editor.
4. **Add a button (optional).** Go to **File → Options → Customize Ribbon**. On the right, select the **Home (Mail)** tab and click **New Group**. On the left, choose **Macros** from the **Choose commands from** list. Select `Project1.CountTeamCases`, click **Add >>**, then **OK**. You can use **Rename…** to give the button a friendly name and icon.

## Use

Click your button, or press **Alt + F8**, choose `CountTeamCases` and click **Run**. If Outlook asks whether to enable macros, choose **Enable Macros**.

The first time, it asks for:

1. **Which mailbox** is the shared support mailbox, from a numbered list of the mailboxes in your Outlook.
2. The **shared mailbox's email address**, so its own messages aren't counted as a person.
3. Your **team's email domain**, for example `yourcompany.com`.
4. Any **team addresses that aren't people** to ignore, such as `noreply@yourcompany.com`. You can leave this empty.

It then asks for the period: `7` for the last 7 days, `30` for the last 30 days, or a calendar month such as `2026-09`. Outlook may stop responding for a few seconds while it counts a large mailbox.

Your settings are saved in the Windows registry for your user only, not in this repository. To change them, run `ResetTeamCaseSettings` the same way, then run `CountTeamCases` again.

## How replies must be sent

| How the team replies | Counted? |
| --- | --- |
| From their own address, with the shared mailbox in To/CC | Yes |
| **Send on Behalf** of the shared mailbox | Yes, credited to the person who sent it |
| **Send As** the shared mailbox | **No.** Outlook records only the shared mailbox, so the person can't be identified. The report shows how many such replies it skipped. |

Send on Behalf replies are only in the shared mailbox if your admin has turned on saving sent items there (`MessageCopyForSendOnBehalfEnabled`). Otherwise the copy stays in the sender's own Sent Items and isn't counted.
