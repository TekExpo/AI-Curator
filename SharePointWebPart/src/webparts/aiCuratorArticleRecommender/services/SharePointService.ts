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
    if (existing.indexOf(newArticleUrl.trim()) !== -1) {
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

  /**
   * Checks whether a list with the given title exists on the site.
   */
  private async listExists(listName: string): Promise<boolean> {
    const url = `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')?$select=Id`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    return response.ok;
  }

  /**
   * Creates a SharePoint Generic list with the given title and fields.
   * Fields: array of { FieldTypeKind, InternalName, Title }
   */
  private async createList(
    listName: string,
    fields: Array<{ InternalName: string; Title: string; FieldTypeKind: number }>
  ): Promise<void> {
    // 1. Create the list (100 = Generic List)
    const createUrl = `${this._webUrl}/_api/web/lists`;
    const createBody = JSON.stringify({
      Title: listName,
      BaseTemplate: 100,
      AllowContentTypes: false,
      ContentTypesEnabled: false
    });
    const createResp: SPHttpClientResponse = await this._spHttpClient.post(
      createUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: createBody
      }
    );
    if (!createResp.ok) {
      const errText = await createResp.text().catch(() => '');
      throw new Error(`Failed to create list "${listName}": HTTP ${createResp.status} — ${errText.substring(0, 200)}`);
    }

    // 2. Add each custom field
    const fieldsUrl = `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields`;
    for (const field of fields) {
      const fieldBody = JSON.stringify({
        Title: field.Title,
        FieldTypeKind: field.FieldTypeKind,
        InternalName: field.InternalName,
        StaticName: field.InternalName
      });
      await this._spHttpClient.post(
        fieldsUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: fieldBody
        }
      );
    }
  }

  /**
   * Ensures the Articles and UserPersonalization lists exist on the site.
   * Creates them with the required columns if missing.
   */
  public async ensureSiteLists(
    articlesListName: string,
    userPersonalizationListName: string
  ): Promise<void> {
    const [articlesExists, personalizationExists] = await Promise.all([
      this.listExists(articlesListName),
      this.listExists(userPersonalizationListName)
    ]);

    const tasks: Promise<void>[] = [];

    if (!articlesExists) {
      tasks.push(
        this.createList(articlesListName, [
          { InternalName: 'Keywords', Title: 'Keywords', FieldTypeKind: 2 },  // Text
          { InternalName: 'ArticleUrl', Title: 'Article URL', FieldTypeKind: 2 },
          { InternalName: 'Source', Title: 'Source', FieldTypeKind: 2 }
        ])
      );
    }

    if (!personalizationExists) {
      tasks.push(
        this.createList(userPersonalizationListName, [
          { InternalName: 'UserId', Title: 'UserId', FieldTypeKind: 9 },       // Number
          { InternalName: 'SelectedTags', Title: 'SelectedTags', FieldTypeKind: 3 }, // Note
          { InternalName: 'SavedLinks', Title: 'SavedLinks', FieldTypeKind: 3 }      // Note
        ])
      );
    }

    await Promise.all(tasks);
  }

  /**
   * Removes an article URL from SavedLinks (comma-separated).
   */
  public async removeSavedLink(
    listName: string,
    itemId: number,
    currentSavedLinks: string,
    urlToRemove: string
  ): Promise<void> {
    const updated = currentSavedLinks
      .split(',')
      .map((l) => l.trim())
      .filter((l) => l.length > 0 && l !== urlToRemove.trim())
      .join(',');

    const url =
      `${this._webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`;
    const body = JSON.stringify({ SavedLinks: updated });
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
        `Failed to remove link from "${listName}" item ${itemId}: HTTP ${response.status}`
      );
    }
  }
}
