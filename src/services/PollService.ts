import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

// ─── Interfaces ─────────────────────────────────────────────────────────

/**
 * A poll item from the PollsQuestions SharePoint list.
 */
export interface IPoll {
  Id: number;
  Title: string;            // The poll question (stored in built-in Title field)
  PollChoices: string;      // JSON stringified array of choice strings
  CreatedBy: string;        // Display name of creator (from built-in Author)
  CreatedByEmail: string;   // Email of creator (from built-in Author)
  ClosingDate: string | null; // When the poll closes (null = never)
  Created: string;
  IsActive: boolean;        // Computed: not expired
}

/**
 * A vote item from the PollsVotes SharePoint list.
 */
export interface IVote {
  Id: number;
  PollQuestionId: number;   // Reference to PollsQuestions item ID
  SelectedChoice: string;   // The choice text selected by voter
  VoterId: number;          // SharePoint user ID of voter
  Created: string;
}

/**
 * Aggregated results for a single poll.
 */
export interface IPollResults {
  pollId: number;
  totalVotes: number;
  /** Map of choice text => vote count */
  optionCounts: { [option: string]: number };
  /** Map of choice text => percentage (0-100) */
  optionPercentages: { [option: string]: number };
  /** The choice the current user voted for, or null */
  currentUserVote: string | null;
}

// ─── List Names ─────────────────────────────────────────────────────────
const LIST_POLLS = 'PollsQuestions';
const LIST_VOTES = 'PollsVotes';

/**
 * Service for all SharePoint list operations related to Polls & Quick Surveys.
 * Follows the same patterns as RecognitionService.ts — uses SPHttpClient,
 * auto-provisions lists on first use, and exposes clean async methods.
 */
export class PollService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;
  private _listsEnsured: boolean = false;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  // ═══════════════════════════════════════════════════════════════════════
  // LIST PROVISIONING
  // ═══════════════════════════════════════════════════════════════════════

  /**
   * Ensure the PollsQuestions and PollsVotes lists exist.
   * Creates them with the required columns if missing.
   * Safe to call multiple times — only runs once.
   *
   * UPDATED FIELD NAMES (user-friendly):
   * - PollsQuestions: PollChoices (Note), ClosingDate (DateTime)
   *   (CreatedByEmail removed — uses built-in Author field instead)
   * - PollsVotes: PollQuestionId (Number), SelectedChoice (Text), Voter (User/Person)
   */
  public async ensureLists(): Promise<void> {
    if (this._listsEnsured) return;

    try {
      // PollsQuestions list stores poll definitions
      // Title = the question, Author = who created it (built-in)
      await this._ensureList(LIST_POLLS, 'Stores poll questions and choices', [
        { type: 'Note', title: 'PollChoices' },       // JSON array: ["Option A","Option B",...]
        { type: 'DateTime', title: 'ClosingDate' }     // When poll closes (optional)
      ]);

      // PollsVotes list stores individual votes
      await this._ensureList(LIST_VOTES, 'Stores individual poll votes', [
        { type: 'Number', title: 'PollQuestionId' },   // Links to PollsQuestions.Id
        { type: 'Text', title: 'SelectedChoice' },     // The choice text the user picked
        { type: 'User', title: 'Voter' }                // Person field — who voted
      ]);

      this._listsEnsured = true;
    } catch (error) {
      console.error('PollService: Failed to ensure lists', error);
      this._listsEnsured = true; // Don't block — lists may already exist
    }
  }

  /**
   * Create a SharePoint list if it doesn't exist, then add custom fields.
   */
  private async _ensureList(
    listTitle: string,
    description: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    additionalFields: any[]
  ): Promise<void> {
    // Check if list already exists
    const checkUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')`;
    const checkResponse = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkResponse.ok) return;

    // Create the list (BaseTemplate 100 = generic list)
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

    const createResponse = await this._spHttpClient.post(createUrl, SPHttpClient.configurations.v1, createOptions);
    if (!createResponse.ok) {
      const err = await createResponse.json().catch(() => ({}));
      console.warn(`PollService: Could not create list "${listTitle}"`, err);
      return;
    }

    // Add custom fields to the new list
    for (const field of additionalFields) {
      let schemaXml = '';
      if (field.type === 'Text') {
        schemaXml = `<Field Type="Text" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" MaxLength="255" />`;
      } else if (field.type === 'Note') {
        schemaXml = `<Field Type="Note" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Number') {
        schemaXml = `<Field Type="Number" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'DateTime') {
        schemaXml = `<Field Type="DateTime" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" Format="DateOnly" />`;
      } else if (field.type === 'User') {
        schemaXml = `<Field Type="User" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
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

  // ═══════════════════════════════════════════════════════════════════════
  // POLLS CRUD
  // ═══════════════════════════════════════════════════════════════════════

  /**
   * Fetch all polls, optionally filtered to only active (non-expired) polls.
   * Polls are ordered newest first.
   * Uses built-in Author field for creator info (no custom CreatedByEmail needed).
   */
  public async getPolls(activeOnly: boolean = true): Promise<IPoll[]> {
    await this.ensureLists();

    let filter = '';
    if (activeOnly) {
      const now = new Date().toISOString();
      // Active = ClosingDate is null (no expiry) OR ClosingDate >= now
      filter = `&$filter=(ClosingDate eq null) or (ClosingDate ge datetime'${now}')`;
    }

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_POLLS}')/items` +
      `?$select=Id,Title,PollChoices,Author/Title,Author/EMail,ClosingDate,Created` +
      `&$expand=Author` +
      `&$orderby=Created desc` +
      `&$top=50` +
      filter +
      `&_t=${new Date().getTime()}`;

    const requestOptions: ISPHttpClientOptions = {
      headers: {
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
      }
    };

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, requestOptions);
    if (!response.ok) {
      throw new Error(`Failed to fetch polls: ${response.statusText}`);
    }

    const json = await response.json();
    const now = new Date();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (json.value || []).map((item: any) => {
      const closingDate = item.ClosingDate ? new Date(item.ClosingDate) : null;
      return {
        Id: item.Id,
        Title: item.Title || '',
        PollChoices: item.PollChoices || '[]',
        CreatedBy: item.Author?.Title || 'Unknown',
        CreatedByEmail: item.Author?.EMail || '',
        ClosingDate: item.ClosingDate || null,
        Created: item.Created,
        IsActive: closingDate === null || closingDate > now
      } as IPoll;
    });
  }

  /**
   * Create a new poll.
   * @param question - The poll question (stored in Title field)
   * @param choices - Array of choice strings (2-6 items)
   * @param closingDate - Optional closing date ISO string
   */
  public async createPoll(
    question: string,
    choices: string[],
    closingDate: string | null
  ): Promise<void> {
    await this.ensureLists();

    // Validate: must have between 2 and 6 choices
    if (choices.length < 2 || choices.length > 6) {
      throw new Error('A poll must have between 2 and 6 choices.');
    }

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_POLLS}')/items`;
    const body = JSON.stringify({
      Title: question,
      PollChoices: JSON.stringify(choices),
      ClosingDate: closingDate || null
    });

    const requestOptions: ISPHttpClientOptions = { body };
    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, requestOptions);
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(`Failed to create poll: ${JSON.stringify(err)}`);
    }
  }

  /**
   * Delete a poll and all its votes.
   * Only the creator should call this (enforced in UI).
   */
  public async deletePoll(pollId: number): Promise<void> {
    // First delete all votes for this poll
    const votesUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_VOTES}')/items` +
      `?$select=Id&$filter=PollQuestionId eq ${pollId}&$top=5000`;
    const votesResponse = await this._spHttpClient.get(votesUrl, SPHttpClient.configurations.v1);
    if (votesResponse.ok) {
      const votesJson = await votesResponse.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const vote of (votesJson.value || []) as any[]) {
        await this._deleteItem(LIST_VOTES, vote.Id);
      }
    }

    // Then delete the poll itself
    await this._deleteItem(LIST_POLLS, pollId);
  }

  // ═══════════════════════════════════════════════════════════════════════
  // VOTING
  // ═══════════════════════════════════════════════════════════════════════

  /**
   * Check if the current user has already voted on a specific poll.
   * Uses the Voter (Person) field's ID for duplicate detection.
   *
   * @param pollId - The poll item ID
   * @param currentUserId - The current user's SharePoint user ID
   * @returns The choice they voted for, or null if they haven't voted
   */
  public async getUserVote(pollId: number, currentUserId: number): Promise<string | null> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_VOTES}')/items` +
      `?$select=SelectedChoice` +
      `&$filter=PollQuestionId eq ${pollId} and VoterId eq ${currentUserId}` +
      `&$top=1` +
      `&_t=${new Date().getTime()}`;

    const requestOptions: ISPHttpClientOptions = {
      headers: {
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
      }
    };

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, requestOptions);
    if (!response.ok) return null;

    const json = await response.json();
    const items = json.value || [];
    return items.length > 0 ? items[0].SelectedChoice : null;
  }

  /**
   * Submit a vote for a poll.
   *
   * VALIDATION FLOW:
   * 1. First checks if user has already voted (duplicate prevention via Voter Person field)
   * 2. If not, creates a new vote item in PollsVotes list
   * 3. Sets the Voter field to the current user's SP ID to prevent future duplicates
   *
   * @param pollId - The poll item ID
   * @param selectedChoice - The choice text selected
   * @param currentUserId - The current user's SharePoint user ID
   */
  public async submitVote(
    pollId: number,
    selectedChoice: string,
    currentUserId: number
  ): Promise<void> {
    await this.ensureLists();

    // STEP 1: Duplicate vote prevention
    // Query the Votes list by PollQuestionId + VoterId (Person field) to check if user already voted
    const existingVote = await this.getUserVote(pollId, currentUserId);
    if (existingVote !== null) {
      throw new Error('You have already voted on this poll.');
    }

    // STEP 2: Create the vote item with Voter as a Person field
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_VOTES}')/items`;
    const body = JSON.stringify({
      Title: `Vote on poll ${pollId}`,
      PollQuestionId: pollId,
      SelectedChoice: selectedChoice,
      VoterId: currentUserId  // Person field — stores the SP user ID
    });

    const requestOptions: ISPHttpClientOptions = { body };
    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, requestOptions);
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(`Failed to submit vote: ${JSON.stringify(err)}`);
    }
  }

  // ═══════════════════════════════════════════════════════════════════════
  // RESULTS CALCULATION
  // ═══════════════════════════════════════════════════════════════════════

  /**
   * Get aggregated results for a poll.
   *
   * CALCULATION LOGIC:
   * 1. Fetches all votes for the given poll from PollsVotes list
   * 2. Counts votes per choice
   * 3. Calculates percentages (vote count / total * 100)
   * 4. Checks if the current user has voted using VoterId (Person field)
   *
   * @param pollId - The poll item ID
   * @param choices - Array of choice strings (from the poll definition)
   * @param currentUserId - The current user's SharePoint user ID
   */
  public async getPollResults(
    pollId: number,
    choices: string[],
    currentUserId: number
  ): Promise<IPollResults> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_VOTES}')/items` +
      `?$select=SelectedChoice,VoterId` +
      `&$filter=PollQuestionId eq ${pollId}` +
      `&$top=5000` +
      `&_t=${new Date().getTime()}`;

    const requestOptions: ISPHttpClientOptions = {
      headers: {
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
      }
    };

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, requestOptions);

    // Initialize counts for all choices to 0
    const optionCounts: { [option: string]: number } = {};
    choices.forEach(opt => { optionCounts[opt] = 0; });

    let totalVotes = 0;
    let currentUserVote: string | null = null;

    if (response.ok) {
      const json = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      (json.value || []).forEach((vote: any) => {
        const selected = vote.SelectedChoice;
        if (optionCounts[selected] !== undefined) {
          optionCounts[selected]++;
        }
        totalVotes++;

        // Check if this vote belongs to the current user (Person field comparison)
        if (vote.VoterId === currentUserId) {
          currentUserVote = selected;
        }
      });
    }

    // Calculate percentages
    const optionPercentages: { [option: string]: number } = {};
    choices.forEach(opt => {
      optionPercentages[opt] = totalVotes > 0
        ? Math.round((optionCounts[opt] / totalVotes) * 100)
        : 0;
    });

    return {
      pollId,
      totalVotes,
      optionCounts,
      optionPercentages,
      currentUserVote
    };
  }

  /**
   * Batch-fetch results for multiple polls at once (optimization).
   * Returns a map of pollId => IPollResults.
   */
  public async getBatchResults(
    polls: IPoll[],
    currentUserId: number
  ): Promise<{ [pollId: number]: IPollResults }> {
    const resultsMap: { [pollId: number]: IPollResults } = {};

    // Fetch all votes in one call and group by PollQuestionId
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_VOTES}')/items` +
      `?$select=PollQuestionId,SelectedChoice,VoterId` +
      `&$top=5000` +
      `&_t=${new Date().getTime()}`;

    const requestOptions: ISPHttpClientOptions = {
      headers: {
        'Cache-Control': 'no-cache',
        'Pragma': 'no-cache'
      }
    };

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1, requestOptions);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let allVotes: any[] = [];
    if (response.ok) {
      const json = await response.json();
      allVotes = json.value || [];
    }

    // Group votes by PollQuestionId
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const votesByPoll: { [pollId: number]: any[] } = {};
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    allVotes.forEach((vote: any) => {
      const pid = vote.PollQuestionId;
      if (!votesByPoll[pid]) votesByPoll[pid] = [];
      votesByPoll[pid].push(vote);
    });

    // Calculate results for each poll
    for (const poll of polls) {
      let choices: string[] = [];
      try {
        choices = JSON.parse(poll.PollChoices);
      } catch { choices = []; }

      const optionCounts: { [option: string]: number } = {};
      choices.forEach(opt => { optionCounts[opt] = 0; });

      let totalVotes = 0;
      let currentUserVote: string | null = null;

      const pollVotes = votesByPoll[poll.Id] || [];
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      pollVotes.forEach((vote: any) => {
        const selected = vote.SelectedChoice;
        if (optionCounts[selected] !== undefined) {
          optionCounts[selected]++;
        }
        totalVotes++;
        if (vote.VoterId === currentUserId) {
          currentUserVote = selected;
        }
      });

      const optionPercentages: { [option: string]: number } = {};
      choices.forEach(opt => {
        optionPercentages[opt] = totalVotes > 0
          ? Math.round((optionCounts[opt] / totalVotes) * 100)
          : 0;
      });

      resultsMap[poll.Id] = {
        pollId: poll.Id,
        totalVotes,
        optionCounts,
        optionPercentages,
        currentUserVote
      };
    }

    return resultsMap;
  }

  // ═══════════════════════════════════════════════════════════════════════
  // EXPORT
  // ═══════════════════════════════════════════════════════════════════════

  /**
   * Generate CSV content for poll results (for Excel export).
   */
  public generateResultsCsv(poll: IPoll, results: IPollResults): string {
    let choices: string[] = [];
    try {
      choices = JSON.parse(poll.PollChoices);
    } catch { choices = []; }

    const rows: string[] = [];
    rows.push('Choice,Votes,Percentage');
    choices.forEach(opt => {
      const count = results.optionCounts[opt] || 0;
      const pct = results.optionPercentages[opt] || 0;
      // Escape commas in choice text
      const escaped = opt.includes(',') ? `"${opt}"` : opt;
      rows.push(`${escaped},${count},${pct}%`);
    });
    rows.push('');
    rows.push(`Total Votes,${results.totalVotes},`);

    return rows.join('\n');
  }

  // ═══════════════════════════════════════════════════════════════════════
  // HELPERS
  // ═══════════════════════════════════════════════════════════════════════

  /**
   * Delete a list item by ID.
   */
  private async _deleteItem(listTitle: string, itemId: number): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;
    const deleteOptions: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };
    await this._spHttpClient.post(url, SPHttpClient.configurations.v1, deleteOptions);
  }
}
