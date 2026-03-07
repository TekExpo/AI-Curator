# AI Curator – Article Recommender (SPFx Web Part)

A production-ready SharePoint Framework (SPFx) web part that uses keywords from a SharePoint list to query an LLM endpoint and display curated article recommendations.

## Features

- **AI-Powered Recommendations** – Sends keywords to your LLM endpoint and renders article links with summaries.
- **Configurable via Property Pane** – All settings (endpoint URL, list name, column name, max articles, caching) are editable without code changes.
- **Responsive & Accessible** – Built with Fluent UI React components; works on desktop, tablet, and mobile. Supports dark/light themes.
- **Session Caching** – Optional sessionStorage caching with 15-minute TTL to reduce redundant API calls.
- **Error Handling** – Friendly messages for empty lists, network failures, and invalid JSON responses.
- **Teams Compatible** – Supports SharePointWebPart, TeamsTab, and TeamsPersonalApp hosts.

## Prerequisites

| Tool           | Version          |
| -------------- | ---------------- |
| Node.js        | 18.17.1 – 18.x  |
| npm            | 9.x+             |
| Gulp CLI       | 4.x              |
| Yeoman         | 4.x (optional)   |
| SPFx           | 1.20.2           |

```bash
# Install Gulp CLI globally (if not already installed)
npm install -g gulp-cli
```

## Getting Started

### 1. Install Dependencies

```bash
cd SharePointWebPart
npm install
```

### 2. Local Development (Workbench)

```bash
gulp serve
```

This opens the SharePoint Workbench where you can add the web part and test it locally.  
**Note:** To test against a live SharePoint list, use the hosted workbench at:  
`https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx`

### 3. Build for Production

```bash
gulp bundle --ship
gulp package-solution --ship
```

The `.sppkg` package will be at:  
`sharepoint/solution/ai-curator-article-recommender.sppkg`

### 4. Deploy

1. Upload the `.sppkg` file to your **SharePoint App Catalog**.
2. Deploy when prompted.
3. Add the web part titled **"AI Curator - Article Recommender"** to any modern SharePoint page.

## Property Pane Configuration

| Property            | Type    | Default     | Description                                                         |
| ------------------- | ------- | ----------- | ------------------------------------------------------------------- |
| LLM Endpoint URL    | Text    | (empty)     | Full URL of the LLM API that accepts keyword-based article requests |
| SharePoint List Name| Text    | `Articles`  | Display name of the SP list containing keywords                     |
| Keyword Column Name | Text    | `Keywords`  | Internal name of the column holding keyword values                  |
| Site URL             | Text    | (empty)     | Optional. Leave empty to use the current site                       |
| Max Number of Articles| Slider| `5`         | Max articles the LLM should return (1–25)                           |
| Enable Caching       | Toggle | `true`      | Cache LLM responses in sessionStorage (15-min TTL)                  |

## LLM Endpoint Contract

### Request (POST)

```json
{
  "keywords": "keyword1, keyword2, keyword3",
  "maxResults": 5
}
```

### Expected Response

```json
[
  {
    "title": "Article Title 1",
    "url": "https://example.com/article-1",
    "summary": "A short description of the article."
  },
  {
    "title": "Article Title 2",
    "url": "https://example.com/article-2",
    "summary": "Another brief summary."
  }
]
```

### Authentication

If your LLM endpoint requires a Bearer token, add the `Authorization` header in:  
`src/webparts/aiCuratorArticleRecommender/components/AiCuratorArticleRecommender.tsx`

Look for the `// AUTHENTICATION:` comment block inside the `callLlmEndpoint()` function.

## Extending the LLM Payload

To add custom fields (e.g., `language`, `userId`, `category`):

1. **Add a Property Pane field** in `AiCuratorArticleRecommenderWebPart.ts` → `getPropertyPaneConfiguration()`.
2. **Add the prop** in `IAiCuratorArticleRecommenderProps.ts`.
3. **Pass the prop** from the web part class → `render()` method → React.createElement.
4. **Include it in the payload** in `AiCuratorArticleRecommender.tsx` → `callLlmEndpoint()` → `payload` object.

## SharePoint List Setup

Create a SharePoint list with at minimum:

| Column       | Type                | Notes                                    |
| ------------ | ------------------- | ---------------------------------------- |
| Title        | Single line of text | Default column (can be left as-is)       |
| Keywords     | Single line of text | Comma-separated keywords per item        |

**Example items:**

| Title                  | Keywords                              |
| ---------------------- | ------------------------------------- |
| Machine Learning Topics | machine learning, neural networks, AI |
| Cloud Architecture      | azure, microservices, kubernetes      |

## Project Structure

```
SharePointWebPart/
├── config/
│   ├── config.json                   # Bundle & localized resources
│   ├── deploy-azure-storage.json     # Azure CDN deployment
│   ├── package-solution.json         # Solution packaging
│   ├── serve.json                    # Local serve settings
│   └── write-manifests.json          # CDN manifest path
├── src/
│   └── webparts/
│       └── aiCuratorArticleRecommender/
│           ├── AiCuratorArticleRecommenderWebPart.manifest.json
│           ├── AiCuratorArticleRecommenderWebPart.ts       # Web part class
│           ├── components/
│           │   ├── AiCuratorArticleRecommender.tsx          # React component
│           │   ├── AiCuratorArticleRecommender.module.scss  # Scoped styles
│           │   └── IAiCuratorArticleRecommenderProps.ts     # TypeScript interfaces
│           └── loc/
│               ├── en-us.js          # English strings
│               └── mystrings.d.ts    # String type definitions
├── .gitignore
├── .yo-rc.json
├── gulpfile.js
├── package.json
├── tsconfig.json
└── README.md
```

## Troubleshooting

| Issue | Solution |
| ----- | -------- |
| `gulp` not found | Run `npm install -g gulp-cli` |
| Build errors with Node 20+ | Use Node 18.x (required by SPFx 1.20) |
| "List does not exist" | Verify the list name matches exactly (case-sensitive) |
| LLM returns error | Check the endpoint URL and ensure CORS is configured |
| Blank web part | Open browser DevTools → Console for error details |

## License

This project is provided as-is for internal use. Modify and distribute as needed within your organization.
