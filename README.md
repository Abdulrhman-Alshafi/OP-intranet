# OP Intranet — SPFx Web Parts

A collection of production-ready SharePoint Framework web parts built for the OP intranet.

---

## Tech Stack

| | |
|---|---|
| **SPFx** | 1.22.2 |
| **React** | 17.0.1 |
| **TypeScript** | ~5.8.0 |
| **Fluent UI** | v8 |
| **Build system** | Heft (Rush Stack) |
| **PnP Controls** | `@pnp/spfx-controls-react` v3 |

---

## Quick Start

```bash
npm install -g @rushstack/heft   # one-time global install
npm install                      # install project dependencies
npm start                        # local workbench (https://localhost:4321/temp/workbench.html)
npm run build                    # production build + package (.sppkg)
```

The packaged solution file is output to:
```
sharepoint/solution/op-intranet.sppkg
```
Upload this file to your SharePoint App Catalog to deploy.

---

## Web Parts

### 1. OP Announcement Banner

**What it does**

Displays dismissable, color-coded announcement banners at the top of a page. Supports four severity levels — Info, Warning, Success, and Urgent — each with its own color and icon. Users can dismiss individual banners; dismissed banners are remembered in local browser storage.

**Property pane settings**

| Setting | Description |
|---|---|
| Announcements | Collection editor — add/edit banners. Each banner has: Message, Type (Info / Warning / Success / Urgent), optional CTA link text, optional CTA URL |
| Stack direction | Vertical (banners stacked) or Horizontal (side by side) |
| Show icons | Toggle banner icons on or off |
| Show dismiss button | Allow users to close individual banners |

**SharePoint requirements**

None — all banner data is stored directly in web part properties. No list or library needed.

---

### 2. OP Announcements

**What it does**

A social-style announcements feed where users can read, create, react to (emoji), and comment on company announcements. Supports pagination and role-based creation controls.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |
| List name | Internal name of the SharePoint announcements list (default: `Announcements`) |
| Items per page | 1–20 announcements per page |
| Enable create | Allow regular users to author new announcements |
| Enable reactions | Show emoji reaction buttons |
| Enable comments | Show threaded comments on each announcement |

**SharePoint requirements**

1. Create a **Custom list** named `Announcements` (name is configurable) in the target site.
2. Add the following columns:

| Column name | Type | Notes |
|---|---|---|
| Title | Single line of text | Built-in — used as the announcement headline |
| Body | Multiple lines of text | Rich text recommended |
| Reactions | Multiple lines (JSON) | Stores emoji reaction data as JSON |
| Comments | Multiple lines (JSON) | Stores comment threads as JSON |

3. Set list permissions so all site members can read items; restrict creation if needed.

---

### 3. OP Countdown Timer

**What it does**

A live countdown clock for events or deadlines. Shows days, hours, minutes, and seconds ticking down in real time. Supports a background image with an overlay, three layout modes, and a custom message shown when the countdown reaches zero.

**Property pane settings**

| Setting | Description |
|---|---|
| Event title | Text shown above the countdown |
| Event description | Optional subtitle |
| Target date | Date and time the countdown counts toward |
| Completed message | Text shown once the timer hits zero |
| Layout | Horizontal, Compact, or Banner |
| Accent color | Hex color for unit labels and accents |
| Background image | URL or local image upload |
| Overlay opacity | 0–100 — darkness of the image overlay |
| Show Days / Hours / Minutes / Seconds | Toggle each time unit individually |

**SharePoint requirements**

None — all configuration is stored in web part properties.

---

### 4. OP Knowledge Base

**What it does**

A searchable, filterable FAQ and knowledge base viewer. Reads articles from a SharePoint list and displays them with full-text search and category filter pills. Supports pagination.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |
| List name | Internal name of the KB list (default: `KnowledgeBase`) |
| Default category | Pre-selected filter on load: All / HR / IT / Policies / Finance / General |
| Items per page | 5–50 articles per page |
| Enable search | Toggle the search box |
| Enable category filter | Toggle the category filter pills |

**SharePoint requirements**

1. Create a **Custom list** named `KnowledgeBase` (name is configurable) in the target site.
2. Add the following columns:

| Column name | Type | Notes |
|---|---|---|
| Title | Single line of text | Built-in — article question/title |
| Answer | Multiple lines of text | The answer body (rich text recommended) |
| Category | Choice | Values: HR, IT, Policies, Finance, General |

3. Populate the list with articles.

---

### 5. OP Polls & Quick Surveys

**What it does**

An interactive polling web part. Employees vote on active polls and immediately see live bar-chart results. Results auto-refresh on a configurable interval. Supports anonymous and attributed voting.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |
| Refresh interval | Auto-refresh results every 10–300 seconds |
| Allow anonymous | Enable voting without user attribution |
| Polls per page | 1–20 polls displayed per page |

**SharePoint requirements**

1. Create a **Custom list** named `Polls` in the target site with these columns:

| Column name | Type | Notes |
|---|---|---|
| Title | Single line | The poll question |
| Options | Multiple lines (JSON) | Poll answer options stored as a JSON array |
| ExpiryDate | Date and Time | Date after which the poll closes |
| IsActive | Yes/No | Controls whether the poll is shown |

2. Create a **Custom list** named `PollVotes` in the same site with these columns:

| Column name | Type | Notes |
|---|---|---|
| PollId | Number | ID of the parent poll item |
| SelectedOption | Single line | The chosen option text |
| VoterId | Number | SP user ID (0 if anonymous) |

3. Grant all site members Contribute access to both lists so they can vote and read results.

---

### 6. OP Recognition Wall

**What it does**

An employee recognition board for submitting and viewing kudos posts. Optionally highlights an Employee of the Month. Site admins can pin or moderate posts. Supports reactions and pagination.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |
| Refresh interval | Auto-refresh wall every 1–60 minutes |
| Show Employee of Month | Toggle the Employee of the Month highlight panel |
| Posts per page | 5–50 recognition posts per page |

**SharePoint requirements**

1. Create a **Custom list** named `Recognition` in the target site with these columns:

| Column name | Type | Notes |
|---|---|---|
| Title | Single line | Short kudos headline |
| Message | Multiple lines | Full recognition message |
| Recipient | Person or Group | Employee being recognized |
| RecognizedBy | Person or Group | Person giving the kudos (auto-fills with current user) |
| Category | Choice | E.g. Teamwork, Innovation, Leadership, Customer Focus |
| Likes | Number | Reaction/like count |

2. Optionally create a separate item or column to designate the **Employee of the Month** (e.g., a `FeaturedMonth` Date column or a separate single-item list).
3. Grant all site members Contribute access so they can submit kudos.

---

### 7. OP Services Grid

**What it does**

A premium card grid showcasing services, links, or products. Each card supports a logo/image, title, description, CTA button, and a custom accent color. The number of columns is configurable (3, 4, or 5).

**Property pane settings**

| Setting | Description |
|---|---|
| Section title | Header text shown above the grid |
| Services | Collection editor — each card has: Image URL / upload, Title (required), Description, CTA button label, CTA URL, Card accent color (hex) |
| Columns | 3, 4, or 5 columns |

**SharePoint requirements**

None — all card data is stored directly in web part properties.

---

### 8. OP Swiper

**What it does**

A full-width hero image carousel with optional side tiles. Each slide has a background image, headline, description, and a CTA button. Up to 4 side tiles can be shown alongside the swiper. Supports autoplay, pagination dots, and arrow navigation.

**Property pane settings**

| Setting | Description |
|---|---|
| Slides | Collection editor — each slide has: Background image (URL or upload), Headline, Description, CTA button text, CTA URL |
| Side tiles | Collection editor (up to 4) — each tile has: Image (URL or upload), Fallback background color, Title, Link URL |
| Container height | Height in pixels (default: 500) |
| Autoplay delay | Milliseconds between slides; 0 to disable |
| Button style | Solid, Outline, or Transparent |
| Show pagination | Toggle pagination bullet dots |
| Show navigation | Toggle previous / next arrow buttons |

**SharePoint requirements**

None — all slide and tile data is stored directly in web part properties.

---

### 9. OP Task Dashboard

**What it does**

A unified task view aggregating tasks from **Microsoft Planner** (via Graph API) and/or a **SharePoint Tasks list**. Displays progress bars, overdue indicators, color-coded status badges, and source labels. Supports grouping by status/source/due date and column sorting.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |
| Enable Planner | Pull tasks from Microsoft Planner |
| Enable SharePoint | Pull tasks from a SharePoint Tasks list |
| SharePoint list | List picker — select the Tasks list when SP source is on |
| Selected plans | Collection editor — enter Planner Plan IDs and display names |
| Group by | Status / Source / Due Date / None |
| Sort by | Due Date / Priority / Title |
| Show completed | Toggle completed tasks visibility |
| Highlight overdue | Toggle overdue task highlighting |
| Refresh interval | Auto-refresh every 1–60 minutes |
| Overdue color | Hex color for overdue indicators (default: `#d13438`) |
| In-progress color | Hex color for in-progress indicators (default: `#0078d4`) |
| Completed color | Hex color for completed indicators (default: `#107c10`) |
| Show progress bars | Toggle plan-level progress bars |
| Show source badges | Toggle Planner / SharePoint source labels on tasks |

**SharePoint requirements**

1. **For Planner tasks**: The Microsoft 365 app registration used by the SPFx app must have the **Tasks.Read** (or **Tasks.ReadWrite**) delegated Graph permission approved by a tenant admin via the SharePoint API access page (`/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement`).

2. **For SharePoint tasks** (optional): Create a **Tasks list** (list template 171 — "Tasks" app from the classic list templates) in the target site. The list picker in the property pane will find it automatically.

---

### 10. Salary Documents

**What it does**

A secure, role-based salary document management system with payslip period tracking.

- **Employees** see only their own salary documents in a single table, filterable by **Year**, **Month**, and a search box — all on one responsive row.
- **Members of the Accountants group** and **site collection administrators** see a three-tab interface:
  - **My Salary** — the accountant's own salary files with period filter + search
  - **Documents** — all employees' documents with Year/Month filter, search, pagination, and delete
  - **Upload** — single-file upload (employee picker + year/month selector) and bulk Excel-driven upload

Key behaviours:
- Every uploaded document has its SharePoint inheritance broken automatically. The target employee receives **Read** access; the Accountants group receives **Edit** access on that item.
- Documents display a **Pay Period** column (e.g. *Jan 2026*) in all tables.
- **Duplicate protection** — uploading a file whose name already exists in the library is blocked in both single and bulk upload modes.
- **Bulk validation** checks for missing files, extra files, duplicate names in the Excel sheet, unresolvable emails, and already-existing files before any upload starts.
- Site collection administrators can delete documents even if they are not in the Accountants group.

**Property pane settings**

| Setting | Description |
|---|---|
| Document Library Name | Internal name of the document library (default: `SalaryDocuments`) |
| Accountant Group Name | Exact SharePoint group display name for accountants (default: `Accountants`) |

**SharePoint requirements — step by step**

#### Step 1 — Create the document library

1. Go to the target site → **New → Document library**
2. Name it `SalaryDocuments` (or whatever you configure in the property pane)
3. Open **Library settings → Advanced settings** → set **Item-level permissions** to:
   - Read access: **Only their own items**
   - Create/edit access: **None** (the web part handles it programmatically)

#### Step 2 — Add metadata columns to the library

1. **Add column → Person** → name it exactly `Employee`
   - Allow only one person (not groups), single value
   - Go to **Library settings → Indexed columns → Add a column → select Employee** (indexing is required for the filter query to work)

2. **Add column → Number** → name it exactly `PayPeriodYear`
   - No decimal places, not required

3. **Add column → Number** → name it exactly `PayPeriodMonth`
   - No decimal places, not required, valid values 1–12

#### Step 3 — Set library-level permissions

The upload code calls `breakroleinheritance` on each uploaded item, which requires **Full Control** on the library itself.

1. **Library settings → Permissions for this document library**
2. Click **Stop Inheriting Permissions**
3. Remove all existing groups
4. Add back only:
   - `Site Owners` → **Full Control**
   - `Accountants` (your group) → **Full Control** ← required for `breakroleinheritance`; individual items will be downgraded to Edit automatically

#### Step 4 — Create the Accountants SharePoint group

1. **Site settings → People and groups → New → New Group**
2. Name it `Accountants` (must match the web part property exactly)
3. Add all accountant users to this group

#### Step 5 — Graph API permissions (for PeoplePicker)

The upload panel's employee picker uses Microsoft Graph. A tenant admin must approve the `User.ReadBasic.All` delegated permission:

1. Go to **SharePoint admin center → Advanced → API access**
2. Approve the pending `User.ReadBasic.All` request (created automatically on first load)

#### Bulk upload Excel format

The Excel manifest must contain exactly **four columns** in the first sheet:

| FileName | EmployeeEmail | Year | Month |
|---|---|---|---|
| JohnSmith_Jan2026.pdf | john.smith@company.com | 2026 | 1 |
| SaraLee_Feb2026.pdf | sara.lee@company.com | 2026 | 2 |

- `FileName` — must exactly match the uploaded file name (including extension); case-insensitive match
- `EmployeeEmail` — must be a valid Azure AD user email
- `Year` — 4-digit number, e.g. `2026`
- `Month` — number `1` (January) through `12` (December)

---

### 11. HR Documents

**What it does**

A secure, role-based HR document management system for employment contracts, offer letters, and other personnel correspondence. Regular employees see **only their own** documents, filterable by **Document Type** (Offer Letter, Contract, Warning Letter, Promotion Letter, Termination Letter, NDA). Members of the HR group see a three-tab interface:

- **My Documents** — the HR member's own personal HR letters with type filter
- **Documents** — all employees' documents with search, document-type filter, pagination, and delete
- **Upload** — single-file upload (with employee picker + document type selector) and bulk Excel-driven upload

Every uploaded document has its SharePoint inheritance broken automatically and item-level permissions granted only to the target employee (Read) and the HR group (Edit). A **Document Type** badge column is displayed in all tables.

**Property pane settings**

| Setting | Description |
|---|---|
| Document Library Name | Internal name of the document library (default: `HRDocuments`) |
| HR Group Name | Exact SharePoint group display name for HR staff (default: `HR`) |

**SharePoint requirements — step by step**

#### Step 1 — Create the document library

1. Go to the target site → **New → Document library**
2. Name it `HRDocuments` (or whatever you configure in the property pane)
3. Open **Library settings → Advanced settings** → set **Item-level permissions** to:
   - Read access: **Only their own items**
   - Create/edit access: **None** (the web part handles it programmatically)

#### Step 2 — Add metadata columns to the library

1. In the library → **Add column → Person** → name it exactly `Employee`
   - Allow only one person (not groups), single value
   - After creation go to **Library settings → Indexed columns → Add a column → select Employee** (indexing required for filter queries)

2. In the library → **Add column → Single line of text** → name it exactly `DocumentType`
   - Not required, max 255 characters

#### Step 3 — Break library-level permissions

1. **Library settings → Permissions for this document library**
2. Click **Stop Inheriting Permissions**
3. Remove all existing groups
4. Add back only:
   - `Site Owners` → **Full Control**
   - `HR` (your group) → **Full Control** ← must be Full Control (not Edit) so the upload code can call `breakroleinheritance` on individual items

#### Step 4 — Create the HR SharePoint group

1. **Site settings → People and groups → New → New Group**
2. Name it `HR` (must match the web part property exactly)
3. Add all HR staff to this group

#### Step 5 — Graph API permissions (for PeoplePicker)

The upload panel's employee picker uses Microsoft Graph. A tenant admin must approve the `User.ReadBasic.All` delegated permission:

1. Go to **SharePoint admin center → Advanced → API access**
2. Approve the pending `User.ReadBasic.All` request (created automatically on first load)

#### Bulk upload Excel format

The Excel manifest used for bulk upload must contain exactly **three columns** in the first sheet:

| FileName | EmployeeEmail | DocumentType |
|---|---|---|
| JohnSmith_OfferLetter.pdf | john.smith@company.com | Offer Letter |
| SaraLee_Contract.pdf | sara.lee@company.com | Contract |
| Ahmed_Warning.pdf | ahmed.ali@company.com | Warning Letter |

- `FileName` must exactly match the file name of the uploaded document (including extension)
- `EmployeeEmail` must be a valid user email that exists in Azure AD
- `DocumentType` must be one of: `Offer Letter`, `Contract`, `Warning Letter`, `Promotion Letter`, `Termination Letter`, `NDA`

---

### 12. IT Support Tickets

**What it does**

An IT support ticket viewer and submission form integrated with **iTop** via a secure **Azure Function proxy**. No iTop credentials are ever exposed to the browser — all iTop REST API calls are made server-side by the Azure Function.

- **All employees** see a **My Tickets** tab showing only their own open/closed tickets, filterable by status and a search box.
- **Members of the IT Admin group** see an additional **All Tickets** tab with every ticket across the organization, with pagination and the same filter controls.
- A **New Ticket** tab lets any employee submit a support request (title, description, service, urgency). The service dropdown is populated live from iTop.
- Each ticket row includes a deep-link that opens the ticket directly in the iTop portal.
- Status badges are color-coded (New = blue, Assigned = orange, Resolved = green, Closed = grey, Pending = yellow) and urgency badges are also color-coded (Critical = red, High = orange, Medium = blue, Low = grey).

**Architecture**

```
SPFx WebPart  ──HttpClient──►  Azure Function  ──REST──►  iTop
              (x-functions-key)   (server-side)         (/webservices/rest.php)
```

The SPFx web part uses `HttpClient` (not `SPHttpClient`) for all Azure Function calls, passing the Function Key via the `x-functions-key` HTTP header. `SPHttpClient` is used **only** for the SharePoint group membership check (admin role-gating).

**Property pane settings**

| Setting | Description |
|---|---|
| Azure Function Base URL | Root URL of the deployed Azure Function app (e.g. `https://my-func-app.azurewebsites.net`) |
| Azure Function Key | Function-level key for authentication |
| IT Admin Group Name | SharePoint group whose members see the All Tickets tab (default: `IT Support`) |

**Azure Function endpoints required**

The Azure Function must expose these three endpoints:

| Method | Path | Description |
|---|---|---|
| GET | `/api/itop/tickets?email={email}` | Returns the caller's own tickets (iTop `core/get` with OQL `SELECT UserRequest WHERE caller_id.email = ':email'`) |
| GET | `/api/itop/tickets?all=true&page={n}&limit={n}` | Returns all tickets with pagination (IT Admin view) |
| GET | `/api/itop/services` | Returns the list of iTop services for the New Ticket dropdown (cached on the client after first fetch) |
| POST | `/api/itop/tickets` | Creates a new UserRequest in iTop (`core/create`) |

**SharePoint requirements**

#### Step 1 — Create the IT Admin SharePoint group

1. **Site settings → People and groups → New → New Group**
2. Name it `IT Support` (or whatever you configure in the property pane — must match exactly)
3. Add all IT support staff who should see the **All Tickets** tab

#### Step 2 — Deploy the Azure Function

1. Create an Azure Function App (Node.js or C#) with the four endpoints listed above
2. Configure the Function App with your iTop instance URL and API credentials as **Application Settings** (environment variables — never in code)
3. Copy the **Function Key** from the Azure portal (Host Keys or Function Keys)
4. In the web part property pane, paste the Function Base URL and Function Key

#### Step 3 — CORS

In the Azure Function App → **CORS settings**, add your SharePoint tenant origin:
```
https://yourtenant.sharepoint.com
```

No additional Graph API permissions or SharePoint lists are required for this web part.

---

## Build & Deploy

```bash
# Development (local workbench with hot reload)
npm start

# Production build
npx heft test --clean --production

# Package for SharePoint
npx heft package-solution --production
# → outputs: sharepoint/solution/op-intranet.sppkg
```

**Deploying:**
1. Upload `op-intranet.sppkg` to the SharePoint App Catalog
2. Choose **Make this solution available to all sites** if you want tenant-wide deployment
3. Add individual web parts to pages via the modern page editor

