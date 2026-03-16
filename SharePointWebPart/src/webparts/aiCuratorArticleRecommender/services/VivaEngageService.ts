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
   * Returns all Viva Engage communities via the employeeExperience API.
   * Uses the community's associated M365 groupId for posting.
   */
  public async getYammerGroups(): Promise<IYammerGroup[]> {
    try {
      const result = await this._graphClient
        .api('/employeeExperience/communities')
        .get() as {
          value?: Array<{ id: string; displayName: string; groupId: string }>;
        };

      const communities = result.value ?? [];
      return communities.map((c) => ({ id: c.groupId, name: c.displayName }));
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
              body: {
                contentType: 'html',
                content: htmlContent
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
