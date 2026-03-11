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
   * Returns Yammer / Viva Engage groups the current user belongs to.
   * Uses the Microsoft Graph /me/joinedGroups endpoint.
   */
  public async getYammerGroups(): Promise<IYammerGroup[]> {
    try {
      // Fetch groups that the user is a member of via Graph
      const result = await this._graphClient
        .api('/me/joinedGroups')
        .select('id,displayName')
        .top(50)
        .get() as {
          value?: Array<{ id: string; displayName: string }>;
        };

      const groups = result.value ?? [];
      return groups.map((g) => ({ id: g.id, name: g.displayName }));
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new Error(`Failed to fetch Viva Engage groups: ${msg}`);
    }
  }

  /**
   * Posts a message to a Viva Engage / Yammer group via Microsoft Graph.
   * Message body: userComments + articleUrl + attribution.
   */
  public async postToGroup(
    groupId: string,
    articleUrl: string,
    userComments: string
  ): Promise<void> {
    const bodyContent = [
      userComments?.trim().length > 0 ? userComments.trim() : '',
      articleUrl,
      'Shared via AI Curator – Article Recommender'
    ]
      .filter((line) => line.length > 0)
      .join('\n\n');

    try {
      await this._graphClient
        .api(`/groups/${groupId}/threads`)
        .post({
          topic: 'AI Curator – Shared Article',
          posts: [
            {
              post: {
                body: {
                  contentType: 'text',
                  content: bodyContent
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
