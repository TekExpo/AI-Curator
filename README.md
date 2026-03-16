# AI Curator – Article Recommender

> A SharePoint Framework (SPFx) web part — and **Microsoft Teams personal & channel tab** — that delivers **AI-powered, personalised article recommendations** directly inside Microsoft 365, with one-click sharing to **Viva Engage** communities and **LinkedIn**.

---

## Overview

Knowledge workers are overwhelmed by information. AI Curator solves this by learning each user's interests and surfacing the most relevant articles — right inside the Microsoft 365 tools they already use every day: SharePoint and Microsoft Teams.

Users select topics they care about, and the web part calls an intelligent Azure-hosted API to recommend curated articles matching those interests. They can save articles for later and share them directly to Viva Engage communities or to LinkedIn — all without leaving Microsoft 365.

---

## Screenshots

| Recommended Articles | My Saved Links | Share to Viva Engage |
|---|---|---|
| AI-ranked articles matching user interests | Personal reading list with one-click remove | Rich-text Viva Engage post with live preview |

---

## Key Features

### 🎯 Personalised Article Recommendations
- Users search for and select topics of interest via a live **AI-powered topic suggestion API**
- Interests are persisted per-user in a SharePoint list (`UserPersonalization`)
- The web part calls an external **recommendations API** (`/articles`) to fetch articles ranked by relevance to the user's saved topics
- Each article card displays: **title**, **summary** (with bullet-point formatting), **source**, and **published date**

### ⭐ My Saved Links
- Users can bookmark any recommended article with a single click
- Saved links are stored in the user's `UserPersonalization` record
- Displayed in a dedicated **My Saved Links** tab with one-click removal

### 🏷️ My Interests Management
- **My Current Interests** section shows the user's saved topics as green chips
- Topics can be deselected or restored inline before saving
- A search box powered by the `/suggest-topics` API helps users discover new topics

### 💬 Share to Viva Engage
- Each article has a **Share to Viva Engage** button that opens a slide-in panel
- The panel pre-populates a **rich-text editor** (Quill) with the article summary
- Users choose a **Viva Engage community** from a dropdown populated via Microsoft Graph
- The post includes: formatted description, article hyperlink, and attribution
- Live **post preview** shows exactly what will be posted before submitting

### 🔗 Share to LinkedIn
- Each article also has a **Share to LinkedIn** button that opens a slide-in panel
- The panel pre-populates a rich-text description from the article summary
- Clicking **Share on LinkedIn** opens LinkedIn's share composer pre-filled with the article title, description, and URL
- No OAuth registration required — uses LinkedIn's standard `shareArticle` URL scheme

### 🟣 Microsoft Teams App
- Ships as a **personal tab** — pinned to the Teams sidebar for individual access
- Ships as a **configurable channel tab** — added to any team or group chat for shared access
- Packaged inside the same `.sppkg` via a Teams manifest — no separate app registration required
- Full feature parity with the SharePoint web part: recommendations, saved links, interests management, and sharing

### 🎨 Polished UI
- Built with **Fluent UI v8** — consistent with the Microsoft 365 design language
- Green accent theme (`#107C10`) throughout
- Responsive layout, dark theme support, Microsoft Teams context support
- Accessible: ARIA labels, keyboard navigation, focus management

---

## Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│          Microsoft 365 Surface (SharePoint or Teams)             │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │            AI Curator SPFx Web Part (React 17)             │  │
│  │                                                            │  │
│  │  Tab 1: Recommended Articles  (default view)               │  │
│  │  Tab 2: My Saved Links                                     │  │
│  │  Tab 3: My Interests                                       │  │
│  └──────────┬────────────────┬───────────────────┬───────────┘  │
│             │                │                   │               │
│             ▼                ▼                   ▼               │
│    External AI API    SharePoint List     Microsoft Graph        │
│    (Azure App Svc)   (UserPersonaliz.)    (Viva Engage)          │
│                                                                  │
│  Surfaces: SharePoint Modern Page · Teams Personal Tab ·        │
│            Teams Configurable Channel Tab                        │
└──────────────────────────────────────────────────────────────────┘
```

### External API (Azure App Service)
| Endpoint | Purpose |
|---|---|
| `GET /suggest-topics?query=` | Live topic suggestions as the user types |
| `GET /articles?topics=&limit=` | AI-ranked article recommendations |

### SharePoint List: `UserPersonalization`
| Column | Type | Purpose |
|---|---|---|
| `UserId` | Number | SharePoint user ID |
| `LoginName` | Text | User login name |
| `SelectedTags` | Note | Comma-separated topic interests |
| `SavedLinks` | Note | Comma-separated saved article URLs |

### Microsoft Graph Permissions
| Permission | Purpose |
|---|---|
| `User.Read` | Get current user identity |
| `Group.Read.All` | List Viva Engage communities |
| `GroupMember.Read.All` | Filter groups by Yammer provisioning |
| `ChannelMessage.Send` | Post to Viva Engage group threads |

---

## Technology Stack

| Technology | Version | Role |
|---|---|---|
| SharePoint Framework (SPFx) | 1.22.1 | Web part framework |
| React | 17.0.1 | UI rendering |
| Fluent UI React | 8.106.4 | Microsoft 365-consistent components |
| TypeScript | ~5.8 | Type safety |
| Heft | 1.1.2 | Build toolchain |
| SPHttpClient | (SPFx built-in) | SharePoint REST operations |
| Microsoft Graph (MSGraphClientV3) | — | Viva Engage integration |
| React Quill | 2.0.0 | Rich-text editor for Viva Engage / LinkedIn posts |

---

## Prerequisites

- **Node.js** 18.17.1 (use nvm to switch: `nvm use 18`)
- **SharePoint Online** tenant with modern pages
- **SharePoint Administrator** role (to deploy the `.sppkg` and approve Graph permissions)
- The `UserPersonalization` list created on the target site (see setup below)

---

## Local Development

```powershell
# 1. Install dependencies
cd SharePointWebPart
npm install

# 2. Start local workbench
npx heft start --clean
# Opens: https://localhost:4321/temp/workbench.html

# 3. Or use the hosted workbench (recommended for Graph API testing)
# Open: https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx
```

---

## Production Build & Deployment

### Step 1 — Build and package

```powershell
cd SharePointWebPart

# Clean production build
npx heft clean
npx heft build --production

# Create the .sppkg package
npx heft package-solution --production
```

The package is created at:
```
SharePointWebPart/sharepoint/solution/ai-curator-article-recommender.sppkg
```

### Step 2 — Upload to App Catalog

1. Go to **SharePoint Admin Center** → **More features** → **Apps** → **App Catalog**
2. Navigate to **Apps for SharePoint**
3. Upload `ai-curator-article-recommender.sppkg`
4. Check **"Make this solution available to all sites"** → click **Deploy**

### Step 3 — Approve Microsoft Graph permissions

1. **SharePoint Admin Center** → **Advanced** → **API access**
2. Approve all four pending Graph permission requests:
   - `Microsoft Graph – User.Read`
   - `Microsoft Graph – Group.Read.All`
   - `Microsoft Graph – GroupMember.Read.All`
   - `Microsoft Graph – ChannelMessage.Send`

### Step 4 — Create the UserPersonalization list

Run the following PnP PowerShell commands on the target site:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -Interactive

New-PnPList -Title "UserPersonalization" -Template GenericList
Add-PnPField -List "UserPersonalization" -DisplayName "UserId"       -InternalName "UserId"       -Type Number
Add-PnPField -List "UserPersonalization" -DisplayName "LoginName"    -InternalName "LoginName"    -Type Text
Add-PnPField -List "UserPersonalization" -DisplayName "SelectedTags" -InternalName "SelectedTags" -Type Note
Add-PnPField -List "UserPersonalization" -DisplayName "SavedLinks"   -InternalName "SavedLinks"   -Type Note
```

### Step 5 — Add to a page

1. Edit any modern SharePoint page
2. Click **+** → search for **"AI Curator"**
3. Add the web part → open the property pane (pencil icon)
4. Set **User Personalization List** to `UserPersonalization`
5. Publish the page

---

## Configuration (Property Pane)

| Setting | Default | Description |
|---|---|---|
| Articles List Name | `Articles` | Display name of the articles reference list |
| User Personalization List | `UserPersonalization` | List storing user interests and saved links |
| Articles Limit | `10` | Maximum number of articles to fetch per request (1–20) |

---

## Project Structure

```
AI-Curator/
├── backend/                                    # Azure App Service — FastAPI AI backend
│   ├── main.py                                 # /suggest-topics and /articles endpoints
│   ├── summarizer.py                           # Azure OpenAI / OpenAI summarisation
│   ├── content_fetcher.py                      # Web content fetching
│   └── requirements.txt
└── SharePointWebPart/
    ├── teams/
    │   └── manifest.json                       # Teams personal + channel tab manifest
    └── src/webparts/aiCuratorArticleRecommender/
        ├── AiCuratorArticleRecommenderWebPart.ts   # Web part entry point & property pane
        ├── components/
        │   ├── AiCuratorArticleRecommender.tsx     # Root component — 3-tab Pivot
        │   ├── AiCuratorArticleRecommender.module.scss
        │   ├── ArticleList/
        │   │   ├── ArticleList.tsx                 # Recommended Articles tab — article grid
        │   │   └── ArticleCard.tsx                 # Individual article card (save + share actions)
        │   ├── TagSelector/
        │   │   └── TagSelector.tsx                 # My Interests tab — AI topic selector
        │   ├── SharePanel/
        │   │   ├── SharePanel.tsx                  # Viva Engage share panel (rich-text + preview)
        │   │   └── ISharePanelProps.ts
        │   └── LinkedInPanel/
        │       ├── LinkedInPanel.tsx               # LinkedIn share panel
        │       └── ILinkedInPanelProps.ts
        └── services/
            ├── TopicsService.ts                    # External AI API calls (/suggest-topics, /articles)
            ├── SharePointService.ts                # SharePoint list operations (SPHttpClient)
            └── VivaEngageService.ts                # Microsoft Graph / Viva Engage
```

---

## How It Works

1. **On load** — the web part fetches article recommendations based on the current user's saved topics
2. **My Interests tab** — the user types a keyword; the `/suggest-topics` API returns matching topics in real time; the user selects topics and saves them
3. **Recommended Articles tab** — calls `/articles?topics=<saved-tags>` and renders ranked article cards with summary, source, and date
4. **Save** — clicking the bookmark icon appends the article URL to `UserPersonalization.SavedLinks` via SPHttpClient
5. **Share to Viva Engage** — clicking the Viva Engage share icon opens a panel; Microsoft Graph fetches the user's Viva Engage communities; the user edits a rich-text description (pre-filled from the article summary) and posts to the selected community
6. **Share to LinkedIn** — clicking the LinkedIn share icon opens a panel; the user edits the pre-filled description, then clicks **Share on LinkedIn** to open LinkedIn's share composer in a new tab
7. **My Saved Links tab** — reads `SavedLinks` from the user's personalization record and renders them with a delete button

---

## Security

- **No API keys stored in the web part** — all AI processing is performed server-side by the external Azure-hosted API
- SharePoint permissions use **SPHttpClient with the web part context** — no elevated permissions
- Microsoft Graph calls use **delegated permissions** — the web part acts as the signed-in user
- Viva Engage post HTML is **sanitised** — article URLs are HTML-escaped before insertion

---

## Hackathon Submission Notes

This project demonstrates:

- ✅ **AI integration** — live topic suggestions and article recommendations via an Azure-hosted AI API (Azure OpenAI)
- ✅ **Microsoft 365 integration** — SharePoint lists for personalisation, Viva Engage for social sharing
- ✅ **Microsoft Teams app** — personal tab and configurable channel tab packaged in the same `.sppkg`
- ✅ **Microsoft Graph API** — community discovery and posting via MSGraphClientV3
- ✅ **LinkedIn sharing** — one-click sharing to professional communities via LinkedIn's `shareArticle` URL scheme
- ✅ **Modern SPFx patterns** — SPFx 1.22.1, Heft build toolchain, React 17, Fluent UI v8
- ✅ **Rich user experience** — three-tab layout, rich-text editor, live post preview, responsive design
- ✅ **Production ready** — deployed as a tenant-wide `.sppkg`, proper Graph permission declarations, no API keys in the front end

---

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Copy `AiCuratorArticleRecommender.config.template.ts` → `AiCuratorArticleRecommender.config.ts` and fill in any local values
4. Commit your changes: `git commit -m "feat: your feature"`
5. Push and open a Pull Request

> **Note:** `AiCuratorArticleRecommender.config.ts` is gitignored. Never commit real keys or secrets.

---

## License

MIT © TrnDigital / TekExpo
