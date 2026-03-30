import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

/**
 * Interface for a Kudos post item from SharePoint
 */
export interface IKudosPost {
  Id: number;
  Title: string;
  SenderId: number;
  SenderName: string;
  SenderEmail: string;
  RecipientId: number;
  RecipientName: string;
  RecipientEmail: string;
  LikesCount: number;
  Category: string;
  Created: string;
  CommentsCount?: number;
  HasLiked?: boolean;
}

/**
 * Interface for a comment item from SharePoint
 */
export interface IComment {
  Id: number;
  PostId: number;
  Comment: string;
  AuthorId: number;
  AuthorName: string;
  Created: string;
}

/**
 * Interface for the top employee highlight
 */
export interface ITopEmployee {
  name: string;
  email: string;
  kudosCount: number;
  userId: number;
}

/**
 * Interface for a manually-selected Employee of the Month
 */
export interface IEmployeeOfMonth {
  Id: number;
  EmployeeId?: number;
  EmployeeName: string;
  EmployeeEmail: string;
  Month: string;
  SelectedById: number;
}

/**
 * List names used by the Recognition Wall
 */
const LIST_KUDOS_POSTS = 'KudosPosts';
const LIST_KUDOS_COMMENTS = 'KudosComments';
const LIST_KUDOS_LIKES = 'KudosLikes';
const LIST_EMPLOYEE_OF_MONTH = 'EmployeeOfMonth';

/**
 * Service class to handle all SharePoint list operations for the Recognition Wall.
 * Uses SPHttpClient following the same pattern as TaskService.ts.
 */
export class RecognitionService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;
  private _listsEnsured: boolean = false;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  // ─── List Provisioning ───────────────────────────────────────────────

  /**
   * Ensure all required SharePoint lists exist. Creates them if missing.
   * Called once on first data fetch.
   */
  public async ensureLists(): Promise<void> {
    if (this._listsEnsured) return;

    try {
      await this._ensureList(LIST_KUDOS_POSTS, 'Stores kudos/recognition posts', [
        { type: 'User', title: 'Sender' },
        { type: 'User', title: 'Recipient' },
        { type: 'Number', title: 'LikesCount' },
        { type: 'Choice', title: 'Category', choices: ['Teamwork', 'Innovation', 'Leadership', 'Customer Focus', 'Above & Beyond'] }
      ]);

      await this._ensureList(LIST_KUDOS_COMMENTS, 'Stores comments on kudos posts', [
        { type: 'Note', title: 'Comment' },
        { type: 'Number', title: 'PostId' }
      ]);

      await this._ensureList(LIST_KUDOS_LIKES, 'Tracks who liked which post', [
        { type: 'Number', title: 'PostId' },
        { type: 'User', title: 'UserId' }
      ]);

      await this._ensureList(LIST_EMPLOYEE_OF_MONTH, 'Admin-selected Employee of the Month', [
        { type: 'User', title: 'Employee' },
        { type: 'Text', title: 'EmployeeName' },
        { type: 'Text', title: 'EmployeeEmail' },
        { type: 'Text', title: 'Month' },
        { type: 'Number', title: 'SelectedById' }
      ]);

      this._listsEnsured = true;
    } catch (error) {
      console.error('RecognitionService: Failed to ensure lists', error);
      // Don't block — lists may already exist
      this._listsEnsured = true;
    }
  }

  /**
   * Create a list if it doesn't already exist.
   */
  private async _ensureList(
    listTitle: string,
    description: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    additionalFields: any[]
  ): Promise<void> {
    // Check if list exists
    const checkUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')`;
    const checkResponse = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkResponse.ok) return; // List already exists

    // Create the list (100 = generic list template)
    const createUrl = `${this._siteUrl}/_api/web/lists`;
    const createOptions: ISPHttpClientOptions = {
      body: JSON.stringify({
        Title: listTitle,
        Description: description,
        BaseTemplate: 100,
        AllowContentTypes: false,
        ContentTypesEnabled: false
      })
    };

    // We let SPHttpClient.configurations.v1 automatically handle the Accept and Content-Type headers
    const createResponse = await this._spHttpClient.post(createUrl, SPHttpClient.configurations.v1, createOptions);
    if (!createResponse.ok) {
      const err = await createResponse.json().catch(() => ({}));
      console.warn(`RecognitionService: Could not create list "${listTitle}"`, err);
      return;
    }

    // Create custom fields for the list
    for (const field of additionalFields) {
      let schemaXml = '';
      if (field.type === 'User') {
        schemaXml = `<Field Type="User" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Number') {
        schemaXml = `<Field Type="Number" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Note') {
        schemaXml = `<Field Type="Note" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Text') {
        schemaXml = `<Field Type="Text" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Choice') {
        const choicesXml = field.choices.map((c: string) => `<CHOICE>${c.replace(/&/g, '&amp;')}</CHOICE>`).join('');
        schemaXml = `<Field Type="Choice" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}"><CHOICES>${choicesXml}</CHOICES></Field>`;
      }

      if (schemaXml) {
        const addFieldUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/createfieldasxml`;
        const addFieldBody = JSON.stringify({
          parameters: {
            "__metadata": { "type": "SP.XmlSchemaFieldCreationInformation" },
            SchemaXml: schemaXml,
            Options: 8 // AddFieldToDefaultView
          }
        });

        // This specific endpoint occasionally rejects nometadata, so we use strictly verbose
        await this._spHttpClient.post(addFieldUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Content-Type': 'application/json;odata=verbose',
            'Accept': 'application/json;odata=verbose'
          },
          body: addFieldBody
        });
      }
    }
  }

  // ─── Posts CRUD ──────────────────────────────────────────────────────

  /**
   * Fetch kudos posts with sender/recipient details, ordered by newest first.
   * Also calculates comments count per post.
   */
  public async getPosts(top: number = 20, skip: number = 0, searchQuery?: string): Promise<{ posts: IKudosPost[]; total: number }> {
    await this.ensureLists();

    let filter = '';
    if (searchQuery) {
      filter = `&$filter=substringof('${encodeURIComponent(searchQuery)}',Title) or substringof('${encodeURIComponent(searchQuery)}',Recipient/Title)`;
    }

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_POSTS}')/items` +
      `?$select=Id,Title,Sender/Id,Sender/Title,Sender/EMail,Recipient/Id,Recipient/Title,Recipient/EMail,LikesCount,Category,Created` +
      `&$expand=Sender,Recipient` +
      `&$orderby=Created desc` +
      `&$top=${top}&$skip=${skip}` +
      filter;

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Failed to fetch posts: ${response.statusText}`);
    }

    const json = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const posts: IKudosPost[] = (json.value || []).map((item: any) => ({
      Id: item.Id,
      Title: item.Title || '',
      SenderId: item.Sender?.Id || 0,
      SenderName: item.Sender?.Title || 'Unknown',
      SenderEmail: item.Sender?.EMail || '',
      RecipientId: item.Recipient?.Id || 0,
      RecipientName: item.Recipient?.Title || 'Unknown',
      RecipientEmail: item.Recipient?.EMail || '',
      LikesCount: item.LikesCount || 0,
      Category: item.Category || '',
      Created: item.Created
    }));

    // Get comments counts for all posts in this batch
    await this._enrichPostsWithCommentCounts(posts);

    return { posts, total: posts.length };
  }

  /**
   * Enrich posts array with comment counts (batch approach).
   */
  private async _enrichPostsWithCommentCounts(posts: IKudosPost[]): Promise<void> {
    if (posts.length === 0) return;

    try {
      // Fetch all comments and group by PostId
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_COMMENTS}')/items?$select=PostId&$top=5000`;
      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return;

      const json = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const commentCounts: { [postId: number]: number } = {};
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (json.value || []).forEach((c: any) => {
        const pid = c.PostId;
        commentCounts[pid] = (commentCounts[pid] || 0) + 1;
      });

      posts.forEach(p => {
        p.CommentsCount = commentCounts[p.Id] || 0;
      });
    } catch (error) {
      console.warn('RecognitionService: Could not fetch comment counts', error);
    }
  }

  /**
   * Create a new kudos post.
   * @param recipientEmail - email of the recipient to resolve their SP user ID
   * @param message - the kudos message (stored in Title field)
   * @param category - category label
   */
  public async createPost(recipientEmail: string, message: string, category: string): Promise<void> {
    await this.ensureLists();

    // Resolve recipient user ID
    const recipientId = await this._ensureUser(recipientEmail);

    // Get current user ID
    const currentUserId = await this._getCurrentUserId();

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_POSTS}')/items`;
    const body = JSON.stringify({
      Title: message,
      SenderId: currentUserId,
      RecipientId: recipientId,
      LikesCount: 0,
      Category: category
    });
    const options: ISPHttpClientOptions = {
      body: body
    };

    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(`Failed to create post: ${JSON.stringify(err)}`);
    }
  }

  // ─── Comments ────────────────────────────────────────────────────────

  /**
   * Get comments for a specific post.
   */
  public async getComments(postId: number): Promise<IComment[]> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_COMMENTS}')/items` +
      `?$select=Id,PostId,Comment,Author/Id,Author/Title,Created` +
      `&$expand=Author` +
      `&$filter=PostId eq ${postId}` +
      `&$orderby=Created asc`;

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Failed to fetch comments: ${response.statusText}`);
    }

    const json = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (json.value || []).map((item: any) => ({
      Id: item.Id,
      PostId: item.PostId,
      Comment: item.Comment || '',
      AuthorId: item.Author?.Id || 0,
      AuthorName: item.Author?.Title || 'Unknown',
      Created: item.Created
    }));
  }

  /**
   * Add a comment to a post.
   */
  public async addComment(postId: number, commentText: string): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_COMMENTS}')/items`;
    const body = JSON.stringify({
      Title: `Comment on post ${postId}`,
      PostId: postId,
      Comment: commentText
    });
    const options: ISPHttpClientOptions = {
      headers: { 'Content-Type': 'application/json;odata=nometadata', 'Accept': 'application/json;odata=nometadata' },
      body: body
    };

    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!response.ok) {
      throw new Error(`Failed to add comment: ${response.statusText}`);
    }
  }

  // ─── Likes ───────────────────────────────────────────────────────────

  /**
   * Check if the current user has liked a specific post.
   */
  public async hasUserLiked(postId: number, userId: number): Promise<boolean> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_LIKES}')/items` +
      `?$select=Id&$filter=PostId eq ${postId} and UserIdId eq ${userId}&$top=1`;

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return false;

    const json = await response.json();
    return (json.value || []).length > 0;
  }

  /**
   * Get like statuses for multiple posts at once (batch optimization).
   */
  public async getUserLikeStatuses(postIds: number[], userId: number): Promise<{ [postId: number]: boolean }> {
    const result: { [postId: number]: boolean } = {};
    postIds.forEach(id => { result[id] = false; });

    if (postIds.length === 0) return result;

    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_LIKES}')/items` +
        `?$select=PostId&$filter=UserIdId eq ${userId}&$top=5000`;

      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return result;

      const json = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (json.value || []).forEach((item: any) => {
        if (result[item.PostId] !== undefined) {
          result[item.PostId] = true;
        }
      });
    } catch (error) {
      console.warn('RecognitionService: Could not fetch like statuses', error);
    }

    return result;
  }

  /**
   * Toggle like on a post. If user already liked, remove the like; otherwise add it.
   * Also updates the LikesCount on the post item.
   */
  public async toggleLike(postId: number, userId: number): Promise<boolean> {
    // Check if already liked
    const likeUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_LIKES}')/items` +
      `?$select=Id&$filter=PostId eq ${postId} and UserIdId eq ${userId}&$top=1`;

    const checkResponse = await this._spHttpClient.get(likeUrl, SPHttpClient.configurations.v1);
    const checkJson = await checkResponse.json();
    const existingLikes = checkJson.value || [];

    if (existingLikes.length > 0) {
      // Unlike: delete the like item
      const likeItemId = existingLikes[0].Id;
      await this._deleteItem(LIST_KUDOS_LIKES, likeItemId);
      await this._incrementLikesCount(postId, -1);
      return false; // no longer liked
    } else {
      // Like: create a new like item
      const createUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_LIKES}')/items`;
      const body = JSON.stringify({
        Title: `Like on post ${postId}`,
        PostId: postId,
        UserIdId: userId
      });
      const options: ISPHttpClientOptions = {
        body: body
      };
      await this._spHttpClient.post(createUrl, SPHttpClient.configurations.v1, options);
      await this._incrementLikesCount(postId, 1);
      return true; // now liked
    }
  }

  /**
   * Update the LikesCount field on a post by a delta (+1 or -1).
   */
  private async _incrementLikesCount(postId: number, delta: number): Promise<void> {
    // First get the current LikesCount
    const getUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_POSTS}')/items(${postId})?$select=LikesCount`;
    const getResponse = await this._spHttpClient.get(getUrl, SPHttpClient.configurations.v1);
    if (!getResponse.ok) return;

    const item = await getResponse.json();
    const newCount = Math.max(0, (item.LikesCount || 0) + delta);

    // Update the count
    const updateUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_POSTS}')/items(${postId})`;
    const body = JSON.stringify({ LikesCount: newCount });
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: body
    };
    await this._spHttpClient.post(updateUrl, SPHttpClient.configurations.v1, options);
  }

  // ─── Employee of the Month (Admin-Selected) ─────────────────────────

  /**
   * Get the admin-selected Employee of the Month for the current month.
   */
  public async getEmployeeOfMonth(): Promise<IEmployeeOfMonth | null> {
    await this.ensureLists();

    const now = new Date();
    const monthKey = `${now.getFullYear()}-${((now.getMonth() + 1) < 10 ? '0' : '') + (now.getMonth() + 1)}`;

    try {
      const personUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_EMPLOYEE_OF_MONTH}')/items` +
        `?$select=Id,Employee/Id,Employee/Title,Employee/EMail,EmployeeName,EmployeeEmail,Month,SelectedById` +
        `&$expand=Employee` +
        `&$filter=Month eq '${monthKey}'` +
        `&$top=1&$orderby=Created desc`;

      const personResponse = await this._spHttpClient.get(personUrl, SPHttpClient.configurations.v1);
      if (!personResponse.ok) return null;

      const personJson = await personResponse.json();
      const personItems = personJson.value || [];
      if (personItems.length === 0) return null;

      const personItem = personItems[0];
      return {
        Id: personItem.Id,
        EmployeeId: personItem.Employee?.Id || 0,
        EmployeeName: personItem.Employee?.Title || personItem.EmployeeName || '',
        EmployeeEmail: personItem.Employee?.EMail || personItem.EmployeeEmail || '',
        Month: personItem.Month || '',
        SelectedById: personItem.SelectedById || 0
      };
    } catch {
      const legacyUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_EMPLOYEE_OF_MONTH}')/items` +
        `?$select=Id,EmployeeName,EmployeeEmail,Month,SelectedById` +
        `&$filter=Month eq '${monthKey}'` +
        `&$top=1&$orderby=Created desc`;

      const legacyResponse = await this._spHttpClient.get(legacyUrl, SPHttpClient.configurations.v1);
      if (!legacyResponse.ok) return null;

      const legacyJson = await legacyResponse.json();
      const legacyItems = legacyJson.value || [];
      if (legacyItems.length === 0) return null;

      const legacyItem = legacyItems[0];
      return {
        Id: legacyItem.Id,
        EmployeeName: legacyItem.EmployeeName || '',
        EmployeeEmail: legacyItem.EmployeeEmail || '',
        Month: legacyItem.Month || '',
        SelectedById: legacyItem.SelectedById || 0
      };
    }
  }

  /**
   * Set Employee of the Month for the current month (admin-only).
   * Removes any previous entry for the same month first.
   */
  public async setEmployeeOfMonth(email: string, name: string): Promise<void> {
    await this.ensureLists();

    const now = new Date();
    const monthKey = `${now.getFullYear()}-${((now.getMonth() + 1) < 10 ? '0' : '') + (now.getMonth() + 1)}`;
    const currentUserId = await this._getCurrentUserId();
    const employeeId = await this._ensureUser(email);

    // Remove existing entry for this month
    await this.removeEmployeeOfMonth();

    // Create new entry
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_EMPLOYEE_OF_MONTH}')/items`;

    const modernBody = JSON.stringify({
      Title: `${name} - ${monthKey}`,
      EmployeeId: employeeId,
      EmployeeName: name,
      EmployeeEmail: email,
      Month: monthKey,
      SelectedById: currentUserId
    });

    let response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, { body: modernBody });

    if (!response.ok) {
      const legacyBody = JSON.stringify({
        Title: `${name} - ${monthKey}`,
        EmployeeName: name,
        EmployeeEmail: email,
        Month: monthKey,
        SelectedById: currentUserId
      });

      response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, { body: legacyBody });
    }

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(`Failed to set Employee of the Month: ${JSON.stringify(err)}`);
    }
  }

  /**
   * Remove Employee of the Month for the current month (admin-only).
   */
  public async removeEmployeeOfMonth(): Promise<void> {
    await this.ensureLists();

    const now = new Date();
    const monthKey = `${now.getFullYear()}-${((now.getMonth() + 1) < 10 ? '0' : '') + (now.getMonth() + 1)}`;

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_EMPLOYEE_OF_MONTH}')/items` +
      `?$select=Id&$filter=Month eq '${monthKey}'`;

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return;

    const json = await response.json();
    const items = json.value || [];
    for (const item of items) {
      await this._deleteItem(LIST_EMPLOYEE_OF_MONTH, item.Id);
    }
  }

  // ─── Admin Check ────────────────────────────────────────────────────

  /**
   * Check if the current user is a Site Collection Admin.
   */
  public async checkIsAdmin(): Promise<boolean> {
    try {
      const url = `${this._siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`;
      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return false;
      const json = await response.json();
      return !!json.IsSiteAdmin;
    } catch {
      return false;
    }
  }

  // ─── Delete Post ────────────────────────────────────────────────────

  /**
   * Delete a kudos post and all its related comments and likes.
   */
  public async deletePost(postId: number): Promise<void> {
    // Delete related comments
    const commentsUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_COMMENTS}')/items?$select=Id&$filter=PostId eq ${postId}`;
    const commentsResp = await this._spHttpClient.get(commentsUrl, SPHttpClient.configurations.v1);
    if (commentsResp.ok) {
      const commentsJson = await commentsResp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const c of (commentsJson.value || [])) {
        await this._deleteItem(LIST_KUDOS_COMMENTS, c.Id);
      }
    }

    // Delete related likes
    const likesUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_KUDOS_LIKES}')/items?$select=Id&$filter=PostId eq ${postId}`;
    const likesResp = await this._spHttpClient.get(likesUrl, SPHttpClient.configurations.v1);
    if (likesResp.ok) {
      const likesJson = await likesResp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const l of (likesJson.value || [])) {
        await this._deleteItem(LIST_KUDOS_LIKES, l.Id);
      }
    }

    // Delete the post itself
    await this._deleteItem(LIST_KUDOS_POSTS, postId);
  }

  // ─── Helpers ─────────────────────────────────────────────────────────

  /**
   * Resolve an email to a SharePoint user ID using EnsureUser.
   */
  private async _ensureUser(email: string): Promise<number> {
    const url = `${this._siteUrl}/_api/web/ensureuser`;
    const body = JSON.stringify({ logonName: email });
    const options: ISPHttpClientOptions = {
      body: body
    };

    const response: SPHttpClientResponse = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, options);
    if (!response.ok) {
      throw new Error(`Could not resolve user: ${email}`);
    }

    const json = await response.json();
    return json.Id;
  }

  /**
   * Get the current user's SharePoint ID.
   */
  private async _getCurrentUserId(): Promise<number> {
    const url = `${this._siteUrl}/_api/web/currentuser?$select=Id`;
    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error('Could not get current user');
    }
    const json = await response.json();
    return json.Id;
  }

  /**
   * Delete a list item by ID.
   */
  private async _deleteItem(listTitle: string, itemId: number): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };
    await this._spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }
}
