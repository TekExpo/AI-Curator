import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IYammerGroup {
  id: string;
  name: string;
}

/**
 * Service for Viva Engage (Yammer) operations via Microsoft Graph API.
 * Uses MSGraphClientV3 for all calls.
 */
export class VivaEngageService {
  private readonly _graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this._graphClient = graphClient;
  }

  /**
   * Returns Viva Engage communities the current user belongs to.
   * Filters M365 groups to those provisioned with a Yammer/Viva Engage backend.
   */
  public async getYammerGroups(): Promise<IYammerGroup[]> {
    try {
      // Filter M365 groups to only those backed by Viva Engage (Yammer)
      const result = await this._graphClient
        .api('/groups')
        .filter("resourceProvisioningOptions/Any(x:x eq 'Yammer')")
        .select('id,displayName')
        .top(100)
        .get() as {
          value?: Array<{ id: string; displayName: string }>;
        };

      const groups = result.value ?? [];
      return groups.map((g) => ({ id: g.id, name: g.displayName }));
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(`Failed to fetch Viva Engage communities: ${msg}`);
    }
  }

  /**
   * Posts a message to a Viva Engage / Yammer group via Microsoft Graph.
   * Message body: userComments + articleUrl + attribution.
   */
  public async postToGroup(
    groupId: string,
    articleUrl: string,
    descriptionHtml: string
  ): Promise<void> {
    // Build the HTML body: rich-text description + article link + attribution
    const safeUrl = articleUrl
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');

    const htmlContent =
      (descriptionHtml?.trim() || '') +
      `<p><a href="${safeUrl}">${safeUrl}</a></p>` +
      '<p><em>Shared via AI Curator \u2013 Article Recommender</em></p>';

    try {
      await this._graphClient
        .api(`/groups/${groupId}/threads`)
        .post({
          topic: 'AI Curator \u2013 Shared Article',
          posts: [
            {
              post: {
                body: {
                  contentType: 'html',
                  content: htmlContent
                }
              }
            }
          ]
        });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(`Failed to post to Viva Engage group: ${msg}`);
    }
  }
}
