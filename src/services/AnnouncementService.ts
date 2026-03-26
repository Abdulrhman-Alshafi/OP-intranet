import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Author/User person field structure
 */
export interface IPerson {
  Id: number;
  Title: string;
  EMail: string;
}

/**
 * Announcement item
 */
export interface IAnnouncement {
  Id: number;
  Title: string;
  Content: string;
  Created: string;
  Author: IPerson;
  AuthorId: number;
}

/**
 * Reaction types
 */
export type ReactionType = 'Like' | 'Love' | 'Fire';

/**
 * User reaction to an announcement
 */
export interface IReaction {
  Id: number;
  AnnouncementId: number;
  ReactionType: ReactionType;
  UserId: number;
  UserEmail: string;
  Created: string;
}

/**
 * Aggregated reaction count and user status
 */
export interface IReactionSummary {
  Like: { count: number; userReacted: boolean };
  Love: { count: number; userReacted: boolean };
  Fire: { count: number; userReacted: boolean };
}

/**
 * Comment or reply item
 */
export interface IComment {
  Id: number;
  AnnouncementId: number;
  ParentCommentId?: number;
  Text: string;
  Created: string;
  Author: IPerson;
  AuthorId: number;
  Replies?: IComment[];
}

/**
 * Service for announcement data operations
 */
export class AnnouncementService {
  /**
   * Get all announcements
   */
  public static async getAnnouncements(
    siteUrl: string,
    listName: string,
    spHttpClient: SPHttpClient
  ): Promise<IAnnouncement[]> {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      listName
    )}')/items?$select=Id,Title,Content,Created,AuthorId,Author/Id,Author/Title,Author/EMail&$expand=Author&$orderby=Created desc&$top=500`;

    const response: SPHttpClientResponse = await spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch announcements: ${response.status}`);
    }

    const data = await response.json();
    return data.value || [];
  }

  /**
   * Create a new announcement
   */
  public static async createAnnouncement(
    siteUrl: string,
    listName: string,
    title: string,
    content: string,
    spHttpClient: SPHttpClient
  ): Promise<IAnnouncement> {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      listName
    )}')/items`;

    const body = JSON.stringify({
      Title: title,
      Content: content
    });

    const response: SPHttpClientResponse = await spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        },
        body
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to create announcement: ${response.status}`);
    }

    return await response.json();
  }

  /**
   * Get reactions for an announcement
   */
  public static async getReactions(
    siteUrl: string,
    listName: string,
    announcementId: number,
    spHttpClient: SPHttpClient
  ): Promise<IReaction[]> {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      listName
    )}')/items?$filter=AnnouncementId eq ${announcementId}&$top=500`;

    const response: SPHttpClientResponse = await spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    return data.value || [];
  }

  /**
   * Add or toggle a reaction
   */
  public static async setReaction(
    siteUrl: string,
    listName: string,
    announcementId: number,
    reactionType: ReactionType,
    userId: number,
    userEmail: string,
    spHttpClient: SPHttpClient,
    existingReactionId?: number
  ): Promise<void> {
    if (existingReactionId) {
      // Delete existing reaction
      const deleteEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
        listName
      )}')/items(${existingReactionId})`;

      const response = await spHttpClient.post(
        deleteEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to remove reaction: ${response.status}`);
      }
    } else {
      // Create new reaction
      const createEndpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
        listName
      )}')/items`;

      const body = JSON.stringify({
        AnnouncementId: announcementId,
        ReactionType: reactionType,
        UserId: userId,
        UserEmail: userEmail
      });

      const response = await spHttpClient.post(
        createEndpoint,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json'
          },
          body
        }
      );

      if (!response.ok) {
        throw new Error(`Failed to add reaction: ${response.status}`);
      }
    }
  }

  /**
   * Get comments for an announcement
   */
  public static async getComments(
    siteUrl: string,
    listName: string,
    announcementId: number,
    spHttpClient: SPHttpClient
  ): Promise<IComment[]> {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      listName
    )}')/items?$filter=AnnouncementId eq ${announcementId}&$select=Id,AnnouncementId,ParentCommentId,Text,Created,AuthorId,Author/Id,Author/Title,Author/EMail&$expand=Author&$orderby=Created asc&$top=500`;

    const response: SPHttpClientResponse = await spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    const comments = (data.value || []) as IComment[];

    // Build nested structure
    return this._buildCommentTree(comments);
  }

  /**
   * Add a comment or reply
   */
  public static async createComment(
    siteUrl: string,
    listName: string,
    announcementId: number,
    text: string,
    spHttpClient: SPHttpClient,
    parentCommentId?: number
  ): Promise<IComment> {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
      listName
    )}')/items`;

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const body: Record<string, any> = {
      AnnouncementId: announcementId,
      Text: text
    };

    if (parentCommentId) {
      body.ParentCommentId = parentCommentId;
    }

    const response: SPHttpClientResponse = await spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        },
        body: JSON.stringify(body)
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to create comment: ${response.status}`);
    }

    return await response.json();
  }

  /**
   * Build nested comment tree from flat array
   */
  private static _buildCommentTree(comments: IComment[]): IComment[] {
    const commentMap = new Map<number, IComment>();
    const rootComments: IComment[] = [];

    // First pass: create map
    comments.forEach((comment) => {
      commentMap.set(comment.Id, { ...comment, Replies: [] });
    });

    // Second pass: build tree
    comments.forEach((comment) => {
      const mappedComment = commentMap.get(comment.Id)!;
      if (comment.ParentCommentId) {
        const parent = commentMap.get(comment.ParentCommentId);
        if (parent) {
          if (!parent.Replies) parent.Replies = [];
          parent.Replies.push(mappedComment);
        }
      } else {
        rootComments.push(mappedComment);
      }
    });

    return rootComments;
  }

  /**
   * Build reaction summary for an announcement
   */
  public static buildReactionSummary(
    reactions: IReaction[],
    currentUserEmail: string,
    currentUserId?: number
  ): IReactionSummary {
    const summary: IReactionSummary = {
      Like: { count: 0, userReacted: false },
      Love: { count: 0, userReacted: false },
      Fire: { count: 0, userReacted: false }
    };

    reactions.forEach((reaction) => {
      const type = reaction.ReactionType as ReactionType;
      if (summary[type]) {
        summary[type].count++;
        if (
          reaction.UserEmail === currentUserEmail ||
          (!!currentUserId && reaction.UserId === currentUserId)
        ) {
          summary[type].userReacted = true;
        }
      }
    });

    return summary;
  }
}
