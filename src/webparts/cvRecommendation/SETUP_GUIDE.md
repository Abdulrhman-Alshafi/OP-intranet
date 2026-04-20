# CV Recommendation Web Part — Setup & Deployment Guide

## Overview

This web part allows employees to recommend candidate CVs to HR.  
It is built with **SPFx + React + Fluent UI** and uses **SharePoint Lists** as the backend.

---

## 1. SharePoint List Configuration

### List: `CVRecommendations`

1. Go to your SharePoint site → **Site Contents → New → List**.
2. Name it exactly: **`CVRecommendations`**
3. Add the following site columns (or let the web part auto-provision them on first load):

| Column Name        | Type                    | Details                                      |
|--------------------|-------------------------|----------------------------------------------|
| `Title`            | Single line of text     | Built-in — used for Candidate Name           |
| `CandidateEmail`   | Single line of text     | Required                                     |
| `PhoneNumber`      | Single line of text     | Optional                                     |
| `Position`         | Single line of text     | Required                                     |
| `Notes`            | Multiple lines of text  | Optional                                     |
| `Status`           | Choice                  | Choices: `Submitted`, `Under Review`, `Accepted`, `Rejected`. Default: `Submitted` |
| `CVFileUrl`        | Single line of text     | Stores URL of uploaded CV file               |
| `CVFileName`       | Single line of text     | Stores original filename                     |

> **Note:** The web part includes an `ensureLists()` method that automatically creates the  
> `CVRecommendations` list and `CVRecommendationFiles` library if they don't already exist when  
> the web part first loads. This requires the user to have Site Owner or Designer permissions.

---

### Document Library: `CVRecommendationFiles`

1. Go to **Site Contents → New → Document Library**.
2. Name it exactly: **`CVRecommendationFiles`**
3. This library stores all uploaded CV files (PDF, DOC, DOCX).

---

## 2. SharePoint Groups & Permissions

### Groups to Create

#### A. `HR Team` (Full Access to CVs)

1. Go to **Site Settings → People and Groups → New Group**.
2. Group name: **`HR Team`**
3. Settings:
   - Who can view the membership: **Everyone**
   - Who can edit the membership: **Group Owner**
4. Add HR department members to this group.

#### B. `Employees` (Contribute — their own items only)

This is usually the standard **"Site Members"** group or a custom group.  
Members gets item-level access to only their own submissions through a combination  
of SharePoint permissions + frontend filtering in the web part.

---

### Permission Levels

#### On `CVRecommendations` List:

1. Go to the list → **List Settings → Permissions for this list**.
2. Click **Stop Inheriting Permissions**.
3. Assign:
   | Group           | Permission Level |
   |-----------------|-----------------|
   | HR Team         | `Full Control` or `Edit` |
   | Site Members    | `Contribute` |
   | Site Visitors   | No access (remove) |

#### Enable Item-Level Permissions (Employees can only read their own):

1. Go to **List Settings → Advanced Settings**.
2. Under **Item-level Permissions**:
   - **Read access**: `Read items that were created by the user`
   - **Create and Edit access**: `Create items and edit items that were created by the user`
3. Click **OK**.

> **HR Team** members bypass item-level permissions because they have `Full Control`.  
> The web part also applies frontend filtering (`AuthorId eq {currentUserId}`) for non-HR users as a secondary layer.

#### On `CVRecommendationFiles` Library:

Apply the same permission break as above:
- HR Team: `Full Control`
- Site Members: `Contribute` (need to upload files)

---

## 3. Determining HR User

The web part calls:
```
/_api/web/currentuser/groups
```
and checks if any group has the title **`HR Team`**.  
Make sure the SharePoint group is named exactly `HR Team`.

To use a different group name, update the constant in:
```
src/services/CvRecommendationService.ts
```
```typescript
const HR_GROUP_NAME = 'HR Team';
```

---

## 4. Building & Deploying

### Prerequisites
- Node.js 16.x or 18.x LTS
- SPFx dependencies installed (`npm install`)

### Build

```bash
# Development build
gulp build

# Production bundle
gulp bundle --ship

# Package solution
gulp package-solution --ship
```

The `.sppkg` file will be created at:
```
sharepoint/solution/op-intranet.sppkg
```

### Deploy to App Catalog

1. Go to your tenant **App Catalog** (`/sites/appcatalog/AppCatalog`).
2. Upload the `.sppkg` file.
3. Click **Deploy** when prompted to trust the solution.
4. If deploying tenant-wide, check **Make this solution available to all sites**.

### Add to SharePoint Page

1. Go to the target SharePoint page.
2. Edit the page → Click **+** to add a new web part.
3. Search for **"CV Recommendation"**.
4. Add it to the page.
5. Optionally configure via the Property Pane:
   - **Title** — displayed as the web part heading
   - **Items Per Page** — how many rows to show per page (default 10)
   - **Accent Color** — hex colour for primary actions (default `#0078d4`)

---

## 5. Feature Summary

| Feature                           | Employee View          | HR View                |
|-----------------------------------|------------------------|------------------------|
| Submit CV Recommendation          | ✅                     | ✅                     |
| View own submissions              | ✅                     | ✅ (all submissions)   |
| Search / filter                   | ✅ (own items)         | ✅ (all items)         |
| Pagination                        | ✅                     | ✅                     |
| View full details                 | ✅ (own items)         | ✅ (all items)         |
| Download CV file                  | ✅ (own items)         | ✅ (all items)         |
| Delete submission                 | ✅ (own items)         | ✅ (any item)          |
| Update status                     | ❌                     | ✅                     |
| Dashboard stats                   | Own stats only         | All stats              |

---

## 6. Status Lifecycle

```
Submitted → Under Review → Accepted
                         ↘ Rejected
```

Only HR can change the status from the **detail panel** of any submission.

---

## 7. Optional: Power Automate Notifications

To send email/Teams notifications when a CV is submitted:

1. Create a **Power Automate flow**.
2. Trigger: **When an item is created** (List: `CVRecommendations`).
3. Action: **Send an email (V2)** or **Post a message in Teams**.
4. Recipients: HR Team SharePoint group or a distribution list.

---

## 8. File Structure

```
src/
  services/
    CvRecommendationService.ts          ← SharePoint API layer
  webparts/
    cvRecommendation/
      CvRecommendationWebPart.ts        ← SPFx WebPart class
      CvRecommendationWebPart.manifest.json
      loc/
        en-us.js                        ← Localizable strings
        mystrings.d.ts
      components/
        ICvRecommendationProps.ts       ← Props interface
        CvRecommendation.tsx            ← Main orchestrator component
        CvDashboard.tsx                 ← Stats dashboard
        CvSubmitForm.tsx                ← CV submission panel/form
        CvList.tsx                      ← Filterable, sortable, paginated list
        CvRecommendation.module.scss    ← Scoped styles (Fluent UI tokens)
```
