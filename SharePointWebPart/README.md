# AI Curator – Article Recommender (SPFx Web Part)

A production-ready SharePoint Framework (SPFx) web part with AI-powered article recommendations, tag-based personalization, and Viva Engage sharing. Uses HEFT as the build toolchain (no Gulp).

## Features

- **Tag-Based Personalization** – Users select interest tags from the Articles list on the "My Interests" tab and save them to their personal record in SharePoint.
- **AI-Powered Recommendations** – Saved tags are sent to the AI Curator API to fetch tailored article recommendations on the "Recommended Articles" tab.
- **Save & Share Articles** – Each recommendation can be bookmarked to the user's saved links, shared to a Viva Engage (Yammer) group, or shared to LinkedIn directly from the web part.
- **Configurable via Property Pane** – List names and Viva Engage settings are editable without code changes.
- **Responsive & Accessible** – Built with Fluent UI v8 React components. Supports dark/light themes and Teams contexts.
- **Session Caching** – Optional sessionStorage caching with 15-minute TTL to reduce redundant API calls.
- **Teams Compatible** – Supports SharePointWebPart, TeamsTab, and TeamsPersonalApp hosts.

## Keyword / Tag Data Flow

> **Important:** Interest tags are **never** entered in the property pane. The flow is:
>
> 1. Tags are read from the **Keywords** column of the **Articles** list (Tab 1 chip options).
> 2. The user selects tags and clicks **Save My Interests** → stored in their **User Personalization** list record (matched by numeric SharePoint User ID).
> 3. On Tab 2 activation, saved `SelectedTags` are fetched from the User Personalization list and sent to the AI Curator API to generate recommendations.
> 4. The property pane never contains any keyword or tag fields.

---

## SharePoint List Requirements

### Articles List

This list already exists and provides both the tag vocabulary and (optionally) article links.

| Column        | Type         | Notes                                       |
| ------------- | ------------ | ------------------------------------------- |
| Title         | Single line  | Article title                               |
| Keywords      | Single line  | Comma-separated tags, e.g. `AI, DevOps, Security` |

> The web part reads **Keywords** from all items in this list, flattens and deduplicates them, and presents them as selectable chips on Tab 1.

### User Personalization List

Create this list manually in the same SharePoint site before deploying the web part.

| Column        | Internal Name | Type                               | Notes                                                      |
| ------------- | ------------- | ---------------------------------- | ---------------------------------------------------------- |
| Title         | Title         | Single line of text                | Set to the user's login name (e.g. `i:0#.f\|membership\|user@tenant.com`) |
| UserId        | UserId        | Number (no decimals)               | **Primary match key** – set to the user's numeric SP ID (`currentUser.Id`) |
| SelectedTags  | SelectedTags  | Multiple lines of text (plain)     | JSON array string of selected tag names, e.g. `["AI","DevOps"]` |
| SavedLinks    | SavedLinks    | Multiple lines of text (plain)     | JSON array string of saved article URLs                    |

> **Important:** `UserId` is used to look up and update the user's record. The `SelectedTags` value is sent to the AI Curator API to fetch recommendations.

---

## Prerequisites

| Tool    | Version         |
| ------- | --------------- |
| Node.js | 18.17.1 – 18.x  |
| npm     | 9.x+            |
| SPFx    | 1.22.1          |

> This project uses **HEFT** as the build toolchain. Gulp is **not** used or required.

## Getting Started

### 1. Install Dependencies

```bash
cd SharePointWebPart
npm install
```

### 2. Local Development (Workbench)

```bash
npm start
```

This opens the SharePoint hosted workbench. Set the workbench URL in `config/serve.json` to your tenant:  
`https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx`

### 3. Build for Production

```bash
npx heft build --production
npm run bundle
npm run package-solution
```

The packaged solution is output to:  
`sharepoint/solution/ai-curator-article-recommender.sppkg`

### 4. Deploy

1. Upload the `.sppkg` file to your **SharePoint App Catalog**.
2. Click **Deploy** when prompted (enable tenant-wide deployment if desired).
3. Add the web part titled **"AI Curator - Article Recommender"** to any modern SharePoint page.

### 5. Approve Graph API Permissions

After deployment, go to **SharePoint Admin Center → Advanced → API Access** and approve:

| Resource          | Scope               |
| ----------------- | ------------------- |
| Microsoft Graph   | Group.Read.All      |
| Microsoft Graph   | User.Read           |
| Microsoft Graph   | ChannelMessage.Send |
| Yammer            | user_impersonation  |

---

## Property Pane Configuration

### Group 1: SharePoint Data Source

| Property           | Type | Default    | Description                                           |
| ------------------ | ---- | ---------- | ----------------------------------------------------- |
| Articles List Name | Text | `Articles` | Display name of the list containing article keywords  |

### Group 2: Personalization & Sharing

| Property                   | Type   | Default               | Description                                                     |
| -------------------------- | ------ | --------------------- | --------------------------------------------------------------- |
| User Personalization List  | Text   | `UserPersonalization` | SP list where user tag selections and saved links are stored    |
| Enable Viva Engage Sharing | Toggle | `false`               | Shows Share button on article cards                             |
| Yammer App Client ID       | Text   | (empty)               | Client ID from Yammer app registration (for OAuth)              |

> **Note:** There are no keyword or tag fields in the property pane. Tags are sourced at runtime from the Articles list, and user selections are persisted in the User Personalization list.

---

## Viva Engage (Yammer) Setup

To enable sharing to Viva Engage groups:

1. Register an app at [https://www.yammer.com/client_applications](https://www.yammer.com/client_applications).
2. Set **Redirect URI** to `https://<tenant>.sharepoint.com`.
3. Copy the **Client ID** and paste it into the property pane **Yammer App Client ID** field.
4. Enable the **Enable Viva Engage Sharing** toggle in the property pane.
5. Ensure the `Yammer: user_impersonation` permission is approved in SharePoint Admin Center → API Access.

---

## Project Structure

```
SharePointWebPart/
├── config/
│   ├── config.json                      # Bundle & localized resources
│   ├── package-solution.json            # Solution packaging + Graph/Yammer permissions
│   └── serve.json                       # Local serve settings
├── src/
│   └── webparts/
│       └── aiCuratorArticleRecommender/
│           ├── AiCuratorArticleRecommender.config.ts   # API configuration
│           ├── AiCuratorArticleRecommenderWebPart.ts   # Web part class & property pane
│           ├── components/
│           │   ├── AiCuratorArticleRecommender.tsx           # Root React component (tabs)
│           │   ├── AiCuratorArticleRecommender.module.scss   # Scoped green-themed styles
│           │   ├── IAiCuratorArticleRecommenderProps.ts      # Web part → component props
│           │   ├── IAiCuratorArticleRecommenderState.ts      # Component state interface
│           │   ├── TagSelector/
│           │   │   ├── TagSelector.tsx          # Tab 1 chip-based tag selector
│           │   │   └── ITagSelectorProps.ts
│           │   ├── ArticleList/
│           │   │   ├── ArticleList.tsx           # Tab 2 article list wrapper
│           │   │   ├── ArticleCard.tsx           # Individual article card (save/share)
│           │   │   └── IArticleListProps.ts
│           │   ├── SharePanel/
│           │   │   ├── SharePanel.tsx            # Fluent UI Panel for Viva Engage sharing
│           │   │   └── ISharePanelProps.ts
│           │   └── LinkedInPanel/
│           │       ├── LinkedInPanel.tsx         # Fluent UI Panel for LinkedIn sharing
│           │       └── ILinkedInPanelProps.ts
│           ├── services/
│           │   ├── SharePointService.ts          # SPHttpClient-based SP REST calls
│           │   ├── TopicsService.ts              # AI Curator API calls + caching
│           │   └── VivaEngageService.ts          # MSGraphClientV3 Yammer group/post
│           └── loc/
│               ├── en-us.js                     # English strings
│               └── mystrings.d.ts               # String type definitions
├── package.json
├── tsconfig.json
└── README.md
```

---

## Build Commands Reference

| Command                               | Description                              |
| ------------------------------------- | ---------------------------------------- |
| `npm install`                         | Install all dependencies                 |
| `npm start`                           | Serve with hot reload (hosted workbench) |
| `npx heft build`                      | Development build                        |
| `npx heft build --production`         | Production build (minified)              |
| `npm run bundle`                      | Bundle for packaging (production)        |
| `npm run package-solution`            | Create `.sppkg` file                     |

---

## Troubleshooting

| Issue | Solution |
| ----- | -------- |
| Build errors with Node 20+ | Use Node 18.x (required by SPFx 1.22) |
| "List does not exist" | Verify Articles and UserPersonalization list names match exactly (case-sensitive) |
| Tags not loading on Tab 1 | Check the Articles list has a Keywords column with comma-separated values |
| Recommendations not loading | Ensure the user has at least one saved tag in their UserPersonalization record |
| Viva Engage groups not loading | Approve Graph permissions in Admin Center; verify Yammer Client ID |
| Share button not visible | Enable **Viva Engage Sharing** toggle in property pane |
| Blank web part | Open browser DevTools → Console for error details |
| `UserId` mismatch | The UserId column must store the numeric SP ID (`currentUser.Id`), not the login name |

## License

This project is provided as-is for internal use. Modify and distribute as needed within your organization.
