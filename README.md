# INCOLORSHOP – Distributor ERP Portal (Google Sheets + Apps Script)

Handover notes for the development team. Current version: **v1.5**

---

## 1. What this is

A web app where each distributor logs in and sees **only their own stock** (IMS), styled like the
company's IMS sheet. There is no separate database – the Google Sheet is the database and
Google Apps Script is the backend + hosting.

```
Distributor browser  ──►  Apps Script web app (index.html)  ──►  google.script.run  ──►  Code.gs / IMS.gs
                                                                                            │
                                            WEBAPP sheet (LOGIN PAGE, SETTINGS, IMS_SNAPSHOT) ◄──┘
                                                                                            │
                                            MIS MEETING sheet (URL in LOGIN PAGE!D1) ◄───────┘  (read-only, every 5 min)
```

Key behaviours
- Login validated against the `LOGIN PAGE` tab.
- A time trigger copies the distributor data from the MIS MEETING sheet into a hidden `IMS_SNAPSHOT`
  tab **every 5 minutes** and warms a memory cache, so the app opens instantly.
- The open page re-checks every 5 minutes; if nothing changed the server replies "unchanged".
- Each user's appearance (colour, day/night, text size) is saved on the server against their ID.
- Sidebar buttons are permission-driven (a YES/NO column per module on `LOGIN PAGE`).

---

## 2. Files in the Apps Script project

| File | Type | Purpose |
|---|---|---|
| `Code.gs` | Script | Config (company name, version, modules), `doGet`, `include()`, login/session, permissions, appearance prefs |
| `IMS.gs` | Script | Reads MIS MEETING sheet → snapshot + cache; `getIMS`, `forceSync`, `installTriggers` |
| `index.html` | HTML | Shell: pulls in every HTML part below in order |
| `Styles.html` | HTML | All CSS, 6 colour themes, dark mode, 3 text sizes, IMS sheet-style table |
| `Login.html` | HTML | Login screen (left hero + sign-in card) |
| `Layout.html` | HTML | Sidebar + top bar (A− A A+, day/night, profile menu, log out) |
| `Scripts.html` | HTML | Shared JS: session, router `go()`, theme engine, IMS cache, filters, error toast |
| `Page_Dashboard.html` | HTML | Dashboard page |
| `Page_IMS.html` | HTML | Distributor IMS page (sheet-style layout) |
| `Page_Profile.html` | HTML | Profile page |
| `Page_Appearance.html` | HTML | Colour picker / presets / day-night / text size |

In the Apps Script editor create each file with the **same name** (the editor adds `.gs` / `.html`).
If using `clasp` (see §6) the `.gs` files are stored locally as `.js`.

Load order matters: `index.html` defines `var Pages = {}` **before** including the `Page_*.html`
files, and includes `Scripts.html` last.

---

## 3. Sheet structure

### 3.1 WEBAPP sheet (the one the script is bound to)

**`LOGIN PAGE`** – header row can be anywhere (script finds the row containing `DIST ID`)

| Col | Header | Notes |
|---|---|---|
| A | PERSON NAME | Display name |
| B | DIST ID | Login ID (case-insensitive) |
| C | DIST PASSWORD | Plain text (spelling `DIST PASSOWRD` is also accepted) |
| D | IMS DIST | `YES` → IMS button visible and data allowed |
| E | PARTY NAME | *Optional.* Distributor's header name on `DISTRIBUTOR CLOSING BALANCE` if it differs from PERSON NAME |
| — | ROLE | *Optional.* Shown under the name (default "Distributor") |

Cell **D1** = full URL of the MIS MEETING sheet. **The URL lives only here, never in code.**

**`SETTINGS`** – optional, `KEY | VALUE` in columns A/B

| KEY | Effect |
|---|---|
| COMPANY NAME | Overrides `CONFIG.companyName` |
| TAGLINE | Overrides tagline |
| DEFAULT THEME | `ruby` / `ocean` / `forest` / `violet` / `amber` / `slate` |
| LOW STOCK LIMIT | Stock ≤ this = low (default 5) |

**`IMS_SNAPSHOT`** – auto-created, hidden. Do not edit; rebuilt every sync.

### 3.2 MIS MEETING sheet (read-only source, URL in D1)

**`DISTRIBUTOR CLOSING BALANCE`**
- Columns A–D: `CATEGORY | ITEM CODE | ITEM NAME | SKU` (item master)
- Then for each distributor a **pair** of columns: `Opening Date | Stock`
  - Row above the pair: distributor name (may be a merged cell)
  - Row between: opening date value and total stock
- Header row auto-detected = the row containing `ITEM CODE` and `SKU`.
- Distributor name is matched to the login by `PARTY NAME` → `PERSON NAME` → `DIST ID`.

**`PURCHASE RATE`**
- Needs `ITEM CODE` and `PURCHASE PRICE` (or `PURCHASE RATE` / `RATE`) columns. Joined by item code; first match wins.

Run `listPartiesFound()` in the editor to log the distributor names the parser detected.

---

## 4. IMS page – column mapping

| Column | Source | Status |
|---|---|---|
| Category, Item code, Item name, SKU | DISTRIBUTOR CLOSING BALANCE A–D | ✅ |
| Opening stock, Opening date | that distributor's `Stock` / `Opening Date` pair | ✅ |
| Rate, Total amount | PURCHASE RATE × closing stock | ✅ |
| Closing stock | calculated: opening + inward − outward | ✅ (calc) |
| % Stock | closing ÷ max level, colour bands ≤33 red / 33–66 yellow / 66–100 green / ≥100 blue | ✅ (needs max level) |
| Max level, Today stock, Inward, Outward, Lead time | **not wired** – show 0 / — | ⏳ pending source |

Candidates seen in the sheet: `SALE FORM`, `IMPORT_TRANFER`, `DISTRIBUTOR LEAD TIME`.
To wire one: add a loader in `IMS.gs` (like `loadRates`), add the field to `FIELDS` / `parseClosingBalance`
/ `rowsToItems`, and drop the placeholder line in `Page_IMS.html` (`i.max=i.max||0 …`).

---

## 5. First-time setup

1. Open the WEBAPP sheet → **Extensions → Apps Script**.
2. Create the 11 files from §2 and paste the code.
3. Put the MIS MEETING URL in `LOGIN PAGE!D1`. Make sure the account running the script can open it.
4. In the editor run **`installTriggers`** once (authorise). This creates the 5-min trigger and does the first sync.
5. **Deploy → New deployment → Web app** → Execute as *Me*, Who has access *Anyone* → Deploy. Share the URL.

---

## 6. Updating the code

### Manual (no tools)
Paste the changed files → **Deploy → Manage deployments → ✏️ → Version: New version → Deploy.**
Apps Script keeps serving the OLD code until a new version is deployed. Check the version stamp at the
bottom-left of the login screen (e.g. `v1.5`) to confirm what is live. Bump `APP_VERSION` in `Code.gs` on every release.

### With clasp (recommended)
```bash
npm install -g @google/clasp
clasp login
clasp clone <SCRIPT_ID>          # Project Settings → Script ID (one time)
# ...copy updated files into the folder (.gs → .js)...
clasp push
clasp deploy -i <DEPLOYMENT_ID>  # Manage deployments → ID; updates the same web-app URL
```

---

## 7. Backend API (called from the browser via `google.script.run`)

| Function | Args | Returns |
|---|---|---|
| `login(id, pw)` | | `{ok, token, user{name,id,role,party,perms{ims}}, prefs}` |
| `logout(token)` | | `{ok}` |
| `savePrefs(token, {theme,mode,font})` | theme = preset name or `custom:#rrggbb` | `{ok, prefs}` |
| `getIMS(token, since)` | `since` = last synced stamp | `{ok, allowed, synced, items[], meta{openDate,total}, lowLimit}` or `{unchanged:true}` |
| `forceSync(token)` | | same as `getIMS` after a fresh sync |

Sessions: UUID token in `CacheService`, 8 h (`CONFIG.sessionHours`).
Prefs: `PropertiesService` key `prefs_<DISTID>`.

---

## 8. Adding a new module (e.g. Orders)

1. `Code.gs` → `MODULES`: add `orders: ['ORDERS']` (column header on LOGIN PAGE that must say `YES`).
2. Create `Orders.gs` with its server functions (use `userFromToken(token)` for auth).
3. Create `Page_Orders.html`:
   ```html
   <section class="page" id="page-orders"></section>
   <script>
   Pages.orders = { title:'Orders', render:function(el){ el.innerHTML='...'; } };
   </script>
   ```
4. `index.html` → add `<?!= include('Page_Orders') ?>` next to the other pages.
5. `Layout.html` → add a nav button:
   `<button class="nav" data-p="orders" data-perm="orders" onclick="go('orders')">…Orders</button>`
6. Add an `ORDERS` column on LOGIN PAGE with YES for allowed users. Deploy a new version.

---

## 9. Troubleshooting

| Symptom | Check |
|---|---|
| Clicking a menu does nothing | New version not deployed – look at the version stamp on the login page |
| Red "Error: …" toast at bottom | JS error text; report it with the page name |
| "Tab DISTRIBUTOR CLOSING BALANCE not found" | D1 URL wrong, or the script account can't open the sheet |
| Distributor sees "Nothing here" | Their name above the Stock column ≠ PARTY NAME / PERSON NAME / DIST ID. Run `listPartiesFound()` |
| Rate shows "—" | Item code missing in PURCHASE RATE |
| Data not refreshing | Trigger missing – run `installTriggers()` again; check Executions log |
| Sync slow / cache miss | Payload > 95 KB per distributor skips the cache and reads `IMS_SNAPSHOT` (still works) |

---

## 10. Known limits / notes

- Passwords are plain text in the sheet – fine for a portal, don't reuse sensitive passwords.
- `CacheService` value limit is 100 KB per key; very large distributors fall back to the snapshot tab.
- Apps Script quota: `openByUrl` on the linked sheet every 5 min is well within limits.
- Design tokens live in `Styles.html` (`:root` and `[data-theme=…]`); a custom colour derives all tokens from one hex in `Scripts.html → applyCustom()`.
