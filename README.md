# Event QR Check-in (100% free)

Unique QR code per guest. Staff scan it with a cellphone; the first scan registers
the entry (green **WELCOME! + name**), any later scan shows red **QR CODE ALREADY USED (at hh:mm)**.

**Cost: R$ 0 / $0.** Uses only a free Google account (Sheets + Apps Script + Gmail) and
free GitHub Pages for the scanner page.

**Quick test, no setup:** open the scanner on your phone, tap **Try demo mode**, then scan
the codes on `demo-tickets.html` (shown on your computer screen).

Files:
- `Code.gs` – backend (paste into Google Apps Script)
- `index.html` – scanner page for the staff phones
- `demo-tickets.html` – sample QR tickets for demo mode

---

## Step 1 – Google Sheet + backend (5 min)

1. Create a new Google Sheet (sheets.new).
2. Menu **Extensions > Apps Script**. Delete the sample code, paste all of `Code.gs`.
   Change `EVENT_NAME` at the top. Click **Save**.
3. Click **Deploy > New deployment** > gear icon > **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**, authorize with your Google account
     (if you see "Google hasn't verified this app": *Advanced > Go to ... (unsafe)* — it's your own script).
   - **Copy the Web app URL** (ends with `/exec`).
4. Back in the sheet, reload the page. A new menu **Event QR** appears:
   - **1) Setup sheet** – creates the headers.
   - **Set staff PIN** – choose a password for the scanner phones (e.g. `festa2026`).

## Step 2 – Add guests and generate QR codes

1. Fill **Name** and **Email** from row 2 down (you can paste from a form/Excel).
2. **Event QR > 2) Generate QR codes** – each guest gets a unique code + QR image.
   (Run it again whenever you add more guests; existing codes are never changed.)
3. Send the tickets:
   - **Event QR > 3) Email QR codes to guests** – each guest gets an email with their QR
     (free Gmail limit: ~100 emails/day; for more, run it again the next day), **or**
   - send manually by WhatsApp: QR image link = `https://quickchart.io/qr?size=300&text=CODE`.

## Step 3 – Scanner page on the phones (free hosting)

The camera only works on HTTPS pages, so host `index.html` on GitHub Pages:

1. github.com > **New repository** (e.g. `event-checkin`, Public) > upload `index.html`.
2. Repo **Settings > Pages** > Source: *Deploy from a branch*, branch `main`, folder `/root` > Save.
3. After ~1 min your page is at `https://YOUR_USER.github.io/event-checkin/`.

On each staff phone open the page, tap **Settings**, paste the Web App URL + staff PIN,
**Save & start scanning**, allow the camera. Done — it's remembered on that phone.

Tip – one-tap setup link for staff (send privately, it contains the PIN):
`https://YOUR_USER.github.io/event-checkin/#api=WEB_APP_URL&pin=YOUR_PIN`

## At the event

- Point the phone at the guest's QR → big **green** screen = let in, **red** = already used,
  **orange** = invalid/fake code. Beep + vibration. Tap the screen for the next guest.
- Damaged QR / screen cracked? Type the code (printed under the QR in the email) in the box.
- Several phones can scan at the same time; the backend uses a lock, so the same code can
  never be accepted twice.
- The sheet updates live: **Status = USED** and **Checked-in at** time — your attendance report.
- **Event QR > Reset ALL check-ins** to clear test scans before the event.

## Notes / limits

- Needs internet on the scanner phones (4G or Wi-Fi). Each check-in takes ~1–2 s.
- Codes are 12 random characters (impossible to guess); guests scanning their own QR with
  their camera only see the code text, so they can't "use up" their ticket.
- If you edit `Code.gs` later: **Deploy > Manage deployments > Edit > Version: New version**
  (keeps the same URL).

## Troubleshooting "Cannot reach the Web App"

Open `YOUR_WEB_APP_URL?action=ping&pin=YOUR_PIN` in a browser. It should show `{"status":"ok"}`.
- "Script function not found: doGet" → Deploy > Manage deployments > Edit > Version: **New version** > Deploy.
- Google login / access denied → "Who has access" must be **Anyone**.
- URL ends with `/dev` → use the `/exec` URL.
