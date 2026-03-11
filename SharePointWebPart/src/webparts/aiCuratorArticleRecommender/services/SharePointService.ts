import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICurrentUser {
  Id: number;
  LoginName: string;
  Title: string;
}

export interface IUserPersonalizationRecord {
  itemId: number;
  SelectedTags: string;
  SavedLinks: string;
}

/**
 * Service for all SharePoint REST API operations.
 * Uses SPHttpClient directly — no PnPjs dependency required.
 */
export class SharePointService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _webUrl: string;

  constructor(context: WebPartContext) {
    this._spHttpClient = context.spHttpClient;
    this._webUrl = context.pageContext.web.absoluteUrl;
  }

  /**
   * Returns numeric Id and LoginName of the currently logged-in user.
   */
  public async getCurrentUser(): Promise<ICurrentUser> {
    const url = `${this._webUrl}/_api/web/currentUser?$select=Id,LoginName,Title`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`Failed to get current user: HTTP ${response.status}`);
    }
    const data = await response.json() as { Id: number; LoginName: string; Title: string };
    return { Id: data.Id, LoginName: data.LoginName, Title: data.Title };
  }

  /**
   * Fetches the Articles list, extracts all Keywords values,
   * and returns a sorted, deduplicated flat array of tag strings.
   */
  public async getTagsFromArticlesList(listName: string): Promise<string[]> {
    const url =
      `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items` +
      `?$select=Keywords&$top=5000`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(
        `Failed to fetch list "${listName}": HTTP ${response.status}`
      );
    }
    const data = await response.json() as { value?: Array<{ Keywords?: string }> };
    const items = data.value ?? [];
    const tagSet = new Set<string>();
    for (const item of items) {
      const kw = item.Keywords ?? '';
      kw.split(',')
        .map((t) => t.trim().toLowerCase())
        .filter((t) => t.length > 0)
        .forEach((t) => tagSet.add(t));
    }
    return Array.from(tagSet).sort();
  }

  /**
   * Queries userPersonalization by numeric UserId.
   * Returns null if no record exists for this user.
   */
  public async getUserPersonalizationByUserId(
    listName: string,
    userId: number
  ): Promise<IUserPersonalizationRecord | null> {
    const url =
      `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items` +
      `?$select=Id,SelectedTags,SavedLinks&$filter=UserId eq ${userId}&$top=1`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(
        `Failed to query "${listName}" for userId ${userId}: HTTP ${response.status}`
      );
    }
    const data = await response.json() as {
      value?: Array<{ Id: number; SelectedTags?: string; SavedLinks?: string }>;
    };
    const items = data.value ?? [];
    if (items.length === 0) return null;
    return {
      itemId: items[0].Id,
      SelectedTags: items[0].SelectedTags ?? '',
      SavedLinks: items[0].SavedLinks ?? ''
    };
  }

  /**
   * Creates a new userPersonalization record for the current user.
   */
  public async createUserPersonalization(
    listName: string,
    loginName: string,
    userId: number,
    selectedTags: string
  ): Promise<{ itemId: number }> {
    const url =
      `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`;
    const body = JSON.stringify({
      Title: loginName,
      UserId: userId,
      SelectedTags: selectedTags,
      SavedLinks: ''
    });
    const response: SPHttpClientResponse = await this._spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body
      }
    );
    if (!response.ok) {
      const errorText = await response.text().catch(() => '');
      throw new Error(
        `Failed to create record in "${listName}": HTTP ${response.status} — ${errorText.substring(0, 300)}`
      );
    }
    const created = await response.json() as { Id: number };
    return { itemId: created.Id };
  }

  /**
   * Updates SelectedTags for an existing userPersonalization record by item ID.
   */
  public async updateSelectedTags(
    listName: string,
    itemId: number,
    selectedTags: string
  ): Promise<void> {
    const url =
      `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`;
    const body = JSON.stringify({ SelectedTags: selectedTags });
    const response: SPHttpClientResponse = await this._spHttpClient.fetch(
      url,
      SPHttpClient.configurations.v1,
      {
        method: 'PATCH',
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body
      }
    );
    if (!response.ok && response.status !== 204) {
      throw new Error(
        `Failed to update SelectedTags in "${listName}" item ${itemId}: HTTP ${response.status}`
      );
    }
  }

  /**
   * Appends an article URL to SavedLinks (comma-separated, no duplicates).
   */
  public async saveArticleLink(
    listName: string,
    itemId: number,
    currentSavedLinks: string,
    newArticleUrl: string
  ): Promise<void> {
    const existing = currentSavedLinks
      .split(',')
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    if (existing.includes(newArticleUrl.trim())) {
      return; // already saved — no-op
    }
    existing.push(newArticleUrl.trim());
    const updatedLinks = existing.join(',');

    const url =
      `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`;
    const body = JSON.stringify({ SavedLinks: updatedLinks });
    const response: SPHttpClientResponse = await this._spHttpClient.fetch(
      url,
      SPHttpClient.configurations.v1,
      {
        method: 'PATCH',
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body
      }
    );
    if (!response.ok && response.status !== 204) {
      throw new Error(
        `Failed to save link in "${listName}" item ${itemId}: HTTP ${response.status}`
      );
    }
  }
}
