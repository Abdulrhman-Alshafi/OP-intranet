# OP Intranet — SPFx Web Parts

A collection of production-ready SharePoint Framework (SPFx) web parts built for the OP intranet portal. All web parts follow a consistent design language using Fluent UI v8, SharePoint theme tokens, and a shared service pattern.

---

## Tech Stack

| | |
|---|---|
| **SPFx** | 1.22.2 |
| **React** | 17.0.1 (Functional components + Hooks) |
| **TypeScript** | ~5.8.0 |
| **Fluent UI** | v8 (`@fluentui/react`) |
| **Build system** | Heft (Rush Stack) |
| **PnP Controls** | `@pnp/spfx-controls-react` v3 · `@pnp/spfx-property-controls` |

---

## Quick Start

```bash
npm install -g @rushstack/heft   # one-time global install
npm install                      # install project dependencies
npm start                        # local workbench → https://localhost:4321/temp/workbench.html
npm run build                    # production build + package (.sppkg)
```

The packaged solution file is output to:
```
sharepoint/solution/op-intranet.sppkg
```
Upload this file to your SharePoint App Catalog to deploy.

---

## Web Parts Overview

| # | Web Part | SharePoint Backend |
|---|---|---|
| 1 | [OP Countdown Timer](#1-op-countdown-timer) | None — web part properties only |
| 2 | [OP Knowledge Base](#2-op-knowledge-base) | `KnowledgeBase` list |
| 3 | [OP Polls & Quick Surveys](#3-op-polls--quick-surveys) | `Polls` + `PollVotes` lists |
| 4 | [OP Recognition Wall](#4-op-recognition-wall) | `Recognition` list |
| 5 | [OP Services Grid](#5-op-services-grid) | None — web part properties only |
| 6 | [OP Swiper](#6-op-swiper) | None — web part properties only |
| 7 | [OP Task Dashboard](#7-op-task-dashboard) | Planner (Graph) + optional SP Tasks list |
| 8 | [OP Salary Documents](#8-op-salary-documents) | `SalaryDocuments` library |
| 9 | [OP HR Documents](#9-op-hr-documents) | `HRDocuments` library |
| 10 | [OP FAQ Section](#10-op-faq-section) | `FAQs` list |
| 11 | [OP Announcements Hub](#11-op-announcements-hub) | `OPAnnouncements` + `OPAnnouncementsDismissed` lists |
| 12 | [OP Help Desk Devices](#12-op-help-desk-devices) | SharePoint list (auto-provisioned) |
| 13 | [OP Devices Catalog](#13-op-devices-catalog) | SharePoint list (auto-provisioned) |
| 14 | [OP Hero Section](#14-op-hero-section) | None — web part properties only |
| 15 | [OP Hero Section V2](#15-op-hero-section-v2) | None — web part properties only |
| 16 | [OP Intranet Hub](#16-op-intranet-hub) | Multiple SP lists (auto-provisioned) |
| 17 | [OP IT Help Desk Hero](#17-op-it-help-desk-hero) | None — web part properties only |
| 18 | [OP Team Members](#18-op-team-members) | None — web part properties only |
| 19 | [OP Sales Performance Dashboard](#19-op-sales-performance-dashboard) | SharePoint list |
| 20 | [OP CV Recommendation](#20-op-cv-recommendation) | `CVRecommendations` list + `CVRecommendationFiles` library |

---

## Web Parts

### 1. OP Countdown Timer

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

### 2. OP Knowledge Base

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

### 3. OP Polls & Quick Surveys

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

### 4. OP Recognition Wall

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

### 5. OP Services Grid

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

### 6. OP Swiper

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

### 7. OP Task Dashboard

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

### 8. OP Salary Documents

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

### 9. OP HR Documents

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

### 10. OP FAQ Section

**What it does**

A collapsible accordion-style FAQ web part that reads questions and answers from a SharePoint list. FAQs are grouped by **Category** and sorted by **Order**. Regular users can only view; site collection administrators see **Add**, **Edit**, and **Delete** controls.

**Features**

- Smooth CSS animated expand/collapse for each item (chevron rotates 90°)
- Category grouping — items are rendered under their category label, falling back to "General"
- Site-admin gate — add/edit/delete controls are hidden from non-admins
- Inline form panel (Fluent UI `Panel`) for creating and editing FAQs
- Loading spinner while fetching, error `MessageBar` on failure, empty-state message when the list has no items

**Permission model**

The web part calls `/_api/web/currentuser?$select=IsSiteAdmin`. Only users with the **Site Collection Administrator** flag can add, edit, or delete FAQ items. All other users have read-only access.

**Property pane settings**

| Setting | Default | Description |
|---|---|---|
| FAQ List Name | `FAQs` | Internal name of the SharePoint list storing the FAQ items |

**SharePoint requirements — step by step**

#### Step 1 — Create the FAQs list

1. Go to the target site → **New → List**
2. Name it `FAQs` (must match the property pane setting exactly)

#### Step 2 — Add columns to the list

The `Title` column (built-in) stores the **Question** text. Add the following additional columns:

| Column display name | Column type | Notes |
|---|---|---|
| `Answer` | Multiple lines of plain text | Stores the FAQ answer |
| `Order` | Number | Controls display order (ascending); SharePoint stores this internally as `Order0` |
| `Category` | Single line of text | Groups FAQs visually (leave blank to fall back to "General") |

> **Important:** When you name a Number column `Order` via the SharePoint UI, the internal/REST API field name becomes `Order0` (SharePoint reserves `Order` as a system keyword). The web part uses `Order0` in all REST calls automatically.

#### Step 3 — Set list permissions (optional)

By default the list inherits site permissions. If you want to restrict who can update FAQs beyond site-admin status, break inheritance and configure item-level permissions manually. The web part permission check (`IsSiteAdmin`) is independent of SharePoint list permissions.

**Source files**

```
src/services/FAQService.ts                           ← data access layer
src/webparts/faqSection/
├── FAQSectionWebPart.ts                             ← web part class
├── FAQSectionWebPart.manifest.json                  ← component manifest
├── loc/
│   ├── en-us.js
│   └── mystrings.d.ts
└── components/
    ├── IFAQSectionProps.ts                          ← root props interface
    ├── FAQSection.tsx                               ← root component (data + state)
    ├── FAQList.tsx                                  ← category-grouped list renderer
    ├── FAQItem.tsx                                  ← single accordion item
    ├── FAQForm.tsx                                  ← add/edit panel form
    └── FAQSection.module.scss                       ← all styles (CSS Modules + SP theme tokens)
```

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

---

### 11. OP Announcements Hub

**What it does**

A dynamic announcements hub with card grid and list views, category filter pills, and admin management. Employees can dismiss announcements; admins can create, edit, and delete them. Supports configurable layout mode, number of items, sort order, and accent color.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |
| Layout mode | Grid or List |
| Number of items to display | 1–50 |
| Sort order | Newest first / Oldest first |
| Show images | Toggle announcement images |
| Enable animations | Toggle card hover animations |
| Enable category colors | Color-coded category badges |
| Accent color | Hex color for primary accents |

**SharePoint requirements**

Lists are **auto-provisioned** on first load.

1. `OPAnnouncements` — stores announcement items:

| Column | Type |
|---|---|
| Title | Single line |
| Description | Multiple lines |
| Category | Choice: General, HR, IT, Finance, Events, Policy, Urgent, Press release |
| ImageUrl | Single line |
| IsImportant | Yes/No |

2. `OPAnnouncementsDismissed` — tracks per-user dismissals:

| Column | Type |
|---|---|
| AnnouncementId | Number |
| UserId | Number |

Grant site members **Contribute** on both lists.

---

### 12. OP Help Desk Devices

**What it does**

Allows IT Help Desk staff to manage device inventory and device request workflows. Admins can add, assign, and track devices. Employees can view and request available devices.

**SharePoint requirements**

The list and columns are **auto-provisioned** on first load by the service layer. Ensure the logged-in user has at least **Contribute** permissions on the target site.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |

---

### 13. OP Devices Catalog

**What it does**

A self-service device catalog for employees to browse available IT devices and submit requests. Displays device details such as model, category, availability, and specs.

**SharePoint requirements**

The list and columns are **auto-provisioned** on first load. Ensure site members have **Contribute** access.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title of the web part |

---

### 14. OP Hero Section

**What it does**

A customizable hero section with a heading, description, and optional images. Designed to be placed at the top of intranet home pages.

**Property pane settings**

| Setting | Description |
|---|---|
| Heading | Main hero headline |
| Description | Supporting body text |
| Image URL | Background or side image |
| CTA button label | Call-to-action button text |
| CTA button URL | Call-to-action link target |
| Accent color | Hex color for button and accents |

**SharePoint requirements**

None — all configuration is stored in web part properties.

---

### 15. OP Hero Section V2

**What it does**

A cinematic full-width hero section with a full-bleed background image, overlay, and an optional CTA button. Upgraded visual style over V1 with parallax-ready layout and overlay opacity control.

**Property pane settings**

| Setting | Description |
|---|---|
| Heading | Main hero headline |
| Description | Subtitle / supporting text |
| Background image URL | Full-width background image |
| Overlay opacity | 0–100% darkness of image overlay |
| CTA button label | Call-to-action button text |
| CTA button URL | Call-to-action link |
| Accent color | Button and highlight hex color |

**SharePoint requirements**

None — all configuration is stored in web part properties.

---

### 16. OP Intranet Hub

**What it does**

A fully dynamic intranet dashboard with multiple configurable sections: Quick Links, Stats, Departments, Tools, Employee Spotlight, and Materials. Acts as the main landing page hub for the intranet.

**SharePoint requirements**

Multiple lists are **auto-provisioned** on first load. Ensure the user has at least **Contribute** access on the target site. A site collection administrator is required for the initial provisioning run.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Display title |
| Visible sections | Toggle Quick Links, Stats, Departments, Tools, Spotlight, Materials individually |
| Accent color | Global hex accent color |

---

### 17. OP IT Help Desk Hero

**What it does**

A hero section purpose-built for the IT Help Desk page. Includes a pre-configured headline, description, and up to 4 action image tiles (e.g. "New devices in stock", "Download latest security update"). Each tile links to a relevant page.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Hero headline |
| Description | Supporting body text |
| Tiles | Collection editor — each tile has: Title, Image URL, Link URL |
| Background image | Hero section background |
| Overlay opacity | 0–100% |
| Accent color | Hex color for button |

**SharePoint requirements**

None — all configuration is stored in web part properties.

---

### 18. OP Team Members

**What it does**

Displays a curated set of employees as profile cards. Each card shows a photo, name, job title, department, and a contact link. Ideal for Department or About pages.

**Property pane settings**

| Setting | Description |
|---|---|
| Title | Section heading |
| Members | Collection editor — each member has: Display name, Job title, Department, Photo URL, Email |
| Columns | 2, 3, or 4 cards per row |
| Card style | Default or Compact |

**SharePoint requirements**

None — all configuration is stored in web part properties.

---

### 19. OP Sales Performance Dashboard

**What it does**

Visualizes sales targets vs. achievements with interactive charts (bar or pie). Reads data from a SharePoint list and renders progress indicators per sales rep or team.

**Property pane settings**

| Setting | Description |
|---|---|
| List ID | GUID of the SharePoint list containing sales data |
| Chart type | Bar or Pie |

**SharePoint requirements**

Create a SharePoint list with the following columns:

| Column | Type |
|---|---|
| Title | Single line — sales rep or team name |
| Target | Number — sales target value |
| Achievement | Number — actual achieved value |
| Period | Single line — e.g. `Q1 2026` |

---

### 20. OP CV Recommendation

**What it does**

An HR portal web part that allows employees to recommend candidate CVs to the HR department. Includes a submission form, role-based list view, status tracking, and an HR dashboard.

**Role-based behaviour**

| Role | Access |
|---|---|
| **Employees** | Submit CVs; view, and delete only their own submissions |
| **HR Team** (SP group `HR Team`) | View all submissions; update status; delete any entry |

**Features**

- Dashboard stats cards: Total / Submitted / Under Review / Accepted / Rejected
- CV submission panel form with validation and file upload (PDF / DOC / DOCX)
- Searchable, filterable, sortable table view with pagination
- Status workflow: Submitted → Under Review → Accepted / Rejected (HR only)
- Detail panel with CV file download link
- Delete confirmation dialog

**Property pane settings**

| Setting | Default | Description |
|---|---|---|
| Title | `CV Recommendation` | Display title of the web part |
| Items per page | `10` | Rows shown per page (5–50) |
| Accent color | `#0078d4` | Hex color for primary button and accents |

**SharePoint requirements — step by step**

#### Step 1 — Lists and library (auto-provisioned)

Both the list and the document library are created automatically on first load if they don't exist:

- **`CVRecommendations`** (Custom list) — stores submission metadata
- **`CVRecommendationFiles`** (Document library) — stores uploaded CV files

No manual provisioning is required, but the user performing the first load must have **Contribute** permissions (or higher) on the site.

#### Step 2 — CVRecommendations list columns

These columns are auto-created by the service. For reference:

| Column | Type | Notes |
|---|---|---|
| Title | Single line | Candidate full name (built-in) |
| CandidateEmail | Single line | Candidate email address |
| PhoneNumber | Single line | Candidate phone number |
| Position | Single line | Role being applied for |
| Notes | Multiple lines | Recommender's comments |
| Status | Choice | Submitted · Under Review · Accepted · Rejected |
| CVFileUrl | Single line | Server-relative URL of uploaded file |
| CVFileName | Single line | Original file name |
| Author | Person (built-in) | Automatically set to submitting user |
| Created | Date (built-in) | Auto-stamped by SharePoint |

#### Step 3 — Create the SharePoint groups

**HR Team group** (full access to all records):

1. Go to **Site settings → People and groups → New → New Group**
2. Name it exactly `HR Team`
3. Add all HR staff to this group
4. Grant this group **Full Control** on the `CVRecommendations` list and the `CVRecommendationFiles` library

**Employees group** (submit and read own items only):

1. The default `Site Members` group (Contribute level) is sufficient
2. To restrict employees to **read only their own items**, configure item-level permissions on the list:
   - Go to **List settings → Advanced settings**
   - Set **Read access**: *Read items that were created by the user*
   - Set **Create and edit access**: *Create items and edit items that were created by the user*

#### Step 4 — Item-level permissions on the list

1. **List settings → Advanced settings**
2. Under **Item-Level Permissions**:
   - Read access → **Read items that were created by the user**
   - Create and edit access → **Create items and edit items that were created by the user**
3. Click **OK**

This ensures SharePoint enforces access at the data layer — not only the frontend filter.

#### Step 5 — Document library permissions

1. Navigate to the `CVRecommendationFiles` library
2. **Library settings → Permissions for this document library → Stop Inheriting Permissions**
3. Grant:
   - `HR Team` → **Full Control**
   - `Site Members` → **Contribute** (needed for employees to upload their own CV files)
   - `Site Owners` → **Full Control**

#### Step 6 — Optional: Power Automate notification flow

To send an email/Teams notification to HR when a new CV is submitted:

1. Go to **Power Automate → New flow → Automated cloud flow**
2. Trigger: **When an item is created** → select site + `CVRecommendations` list
3. Action: **Send an email (V2)** or **Post a message in a Teams channel**
4. Address the notification to your HR Team distribution list or Teams channel

**Source files**

```
src/services/CvRecommendationService.ts            ← data access layer
src/webparts/cvRecommendation/
├── CvRecommendationWebPart.ts                     ← web part class
├── CvRecommendationWebPart.manifest.json          ← component manifest
├── loc/
│   ├── en-us.js
│   └── mystrings.d.ts
└── components/
    ├── ICvRecommendationProps.ts                  ← root props interface
    ├── CvRecommendation.tsx                       ← root component (data + state)
    ├── CvDashboard.tsx                            ← stats cards
    ├── CvSubmitForm.tsx                           ← submission panel form
    ├── CvList.tsx                                 ← filterable/sortable table + detail panel
    └── CvRecommendation.module.scss               ← styles (CSS Modules + SP theme tokens)
```

---

