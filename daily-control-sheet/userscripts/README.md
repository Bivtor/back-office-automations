# Remittance upload pipeline (userscripts)

Four Tampermonkey stages that carry the remittance files the control sheet matched into the deposit system. Each stage runs on one page of the deposit app, asks the local control-sheet server (`http://127.0.0.1:8777`) whether a job is armed, and does nothing on an ordinary visit when no server is running.

| Stage | Page | What it does |
|---|---|---|
| [1 — find the session](stage1-find-the-session.user.js) | Daily Deposit Summary | Finds today's ACH report upload by name and date, refuses to guess when more than one row matches, records the session, and opens its details page. |
| [2 — read the deposit list](stage2-read-the-deposit-list.user.js) | Daily Deposit Details | Reads every deposit row, sends them to the server to be matched against the sheet's figures, and shows the plan in a draggable panel: what to delete, split, upload or leave, and why. A two-click Go runs it one action per page load, since each delete reloads the page. |
| [3 — attach the file](stage3-attach-the-file.user.js) | Deposit View | Fetches the approved file's bytes from the server, loads them into the form's own file input, checks deposit ID, folder name, total amount and deleted state before touching anything, submits, and confirms from the page's own attachment table on reload. Walks the rest of the queue in the same tab. |
| [4 — do the split](stage4-do-the-split.user.js) | Split ACH Deposit | Enters the lockbox ACH amounts in the sheet's order, adds rows through the page's own button so its counter stays in step, requires the page's remaining amount to agree with the script's arithmetic to the cent, then saves. |

## Server endpoints the stages expect

```
GET  /api/upload/job                 armed? files, report pattern, plan, run state
POST /api/upload/session             {session_id}
POST /api/upload/deposits            {deposits:[...]}  -> annotated plan
POST /api/upload/approve             {ids:[...], auto_submit}
GET  /api/upload/for/<deposit_id>    job for one deposit (file url, size, type, expected amount)
POST /api/upload/report              {deposit_id, status, message}
GET  /api/upload/next?after=<id>     next deposit in the queue
POST /api/upload/run                 {active, reason}
POST /api/upload/step                {action, deposit_id, ok?, message?}
POST /api/upload/split/approve       {deposit_id}
POST /api/upload/split/report        {status, message}
```

The `server.py` in this folder's parent predates these endpoints; the server side of the pipeline ships in a later update.
