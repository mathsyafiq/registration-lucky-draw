# Event Registration & Lucky Draw

- `docs/` — frontend, served by GitHub Pages
  - `index.html` — home page linking to the three pages
  - `register.html` — **Participants**: scan QR → enter Staff ID → registered (staff not in the list can register as walk-ins)
  - `draw.html` — **Lucky Draw**: wheel of fortune, drum roll, announces the winner's Staff ID twice (PIN)
  - `admin.html` — **Admin**: live attendance, upload participant list, QR poster, winners, new event (PIN)
  - `assets/app.js` — shared helpers + **API URL setting**
- `apps-script/Code.gs` — backend API (Google Apps Script + Google Sheet)

## 1. Deploy the backend (Apps Script)

1. Create a new Google Sheet → **Extensions → Apps Script**.
2. Replace the contents of `Code.gs` with `apps-script/Code.gs`.
3. **Project Settings (⚙)** → Time zone: **(GMT+08:00) Kuala Lumpur**.
4. Run `initialize()` once (authorise when asked). It creates the `Settings`, `Participants` and `Winners` tabs.
   It creates the `Departments` tab too.
   Your **admin PIN** is printed in the execution log and is also in the `Settings` tab. Change it from the Admin page.
5. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the Web App URL (ends in `/exec`).

> After changing `Code.gs` later, use **Deploy → Manage deployments → Edit → New version**; the URL stays the same.

## 2. Configure the frontend

In `docs/assets/app.js`, set:

```js
var DEFAULT_API_URL = 'https://script.google.com/macros/s/XXXX/exec';
```

## 3. Publish on GitHub Pages

```bash
git init && git add . && git commit -m "Registration and lucky draw system"
```

Then create a GitHub repo, push, and on GitHub: **Settings → Pages → Source: Deploy from a branch → `main` / `/docs`**.
Your site will be at `https://<user>.github.io/<repo>/`.

## Running an event

1. **Admin → Event setup**: set the event name and date, and paste the participant list (copy the columns
   *Staff ID, Name, Department* from Excel/Google Sheets; a header row is detected automatically).
   You can also type the list directly into the `Participants` tab of the Google Sheet.
2. **Admin → Event setup → Print poster**: put the QR code at the entrance. For a shared tablet at the door, use
   the kiosk link (it resets the form after each person).
3. Watch arrivals on **Admin → Overview** (refreshes every 10 s). Staff without a phone can be registered
   manually from **Admin → Participants**.
   - **Walk-ins:** if a Staff ID is not in the list, the participant can tap *Register as a new participant*, enter
     their name and choose their department. They are added to the list (marked *Walk-in*) and join the lucky draw.
     Turn this on/off in **Admin → Event setup → Allow walk-in registration**.
   - **Departments** for the walk-in dropdown: **Admin → Event setup → Departments** (one per line), or type them into
     the `Departments` tab of the Google Sheet. Until a list is set, the departments in the participant list are used.
4. When it's time for the draw, open **Lucky Draw** on the projector laptop, enter the PIN, open **⚙ Settings**
   and set the number of lucky draws. Click **Full screen**. Press **▶ Test sound & voice** once to check the speakers.
5. The VIP clicks **SPIN** (or press Space). After the drum roll the winner's Staff ID is shown and announced twice,
   digit by digit (e.g. "dua, empat, kosong, satu, satu, kosong"). **⏹ Stop announcement** stops the voice.
   - **Continue** → the winner is removed from the wheel; spin for the next prize.
   - **Winner absent — void** → the prize goes back into play; that person is not drawn again.
6. You may close registration from the Admin header when the draw starts.

### Next event

**Admin → Event setup → Start a new event**. The current attendance and winners are copied to new archive tabs in
the Google Sheet, then cleared. Tick *Keep the same participant list* if the staff list hasn't changed; otherwise
upload the new list afterwards.

## Letting other organisers use it

Anyone can run their own copy with their own Google Sheet:

- **Simplest:** they make a copy of your Google Sheet (File → Make a copy, which copies the script too), run
  `initialize()`, and deploy it as in step 1.
- They can then use **your** GitHub Pages site with **their** backend by adding `?api=<their /exec URL>` to the link,
  e.g. `https://<user>.github.io/<repo>/admin.html?api=https://script.google.com/macros/s/AKfy.../exec`.
  The QR code generated in their Admin page keeps their `?api=`, so their participants register into their sheet.

## Notes

- The winner is picked on the server (random among registered participants who have not been drawn) and saved
  before the wheel stops, so refreshing the page or double-clicking cannot change or duplicate a result.
- With a large crowd the wheel shows a random 40 Staff IDs so the labels stay readable; the draw itself always
  uses everyone eligible, and the winner's ID is placed on the wheel where it stops.
- Sounds are generated in the browser (no audio files). The voice uses the browser's text-to-speech: Chrome or
  Safari on a laptop works best. It picks a Malay (ms-MY) voice, or Indonesian if none is installed; you can choose
  a different voice in ⚙ Settings.
- Registration returns the participant's name and department so they can confirm it's them. Anyone who knows a
  Staff ID can see that person's name and department.
- 20 wrong PIN attempts lock the Admin/Lucky Draw pages for 10 minutes. Use a PIN of 6+ characters.
- Settings edited directly in the Google Sheet take up to 30 seconds to apply.
