import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * Common task interface that normalizes tasks from different sources
 */
export interface ITask {
  id: string;
  title: string;
  source: 'Planner' | 'SharePoint';
  dueDate: Date | null;
  status: 'Not Started' | 'In Progress' | 'Completed';
  progress: number; // 0-100
  planId?: string;
  planName?: string;
  listName?: string;
  assignedTo?: string[];
  priority?: 'Low' | 'Medium' | 'High' | 'Urgent';
  hasDescription?: boolean;
}

/**
 * Result object for service methods
 */
export interface ITaskResult {
  tasks: ITask[];
  error: string | null;
}

/**
 * Service class to fetch tasks from multiple sources
 */
export class TaskService {
  private graphClient: MSGraphClientV3;
  private spHttpClient: SPHttpClient;

  constructor(graphClient: MSGraphClientV3, spHttpClient: SPHttpClient) {
    this.graphClient = graphClient;
    this.spHttpClient = spHttpClient;
  }

  /**
   * Fetch Planner tasks assigned to the current user
   */
  public async getPlannerTasks(): Promise<ITaskResult> {
    try {
      // Get all tasks assigned to the user
      const response = await this.graphClient
        .api('/me/planner/tasks')
        .version('v1.0')
        .select('id,title,percentComplete,dueDateTime,planId,hasDescription,priority')
        .get();

      if (!response || !response.value) {
        return { tasks: [], error: null };
      }

      const tasks: ITask[] = [];
      const planCache: { [key: string]: string } = {};

      // Fetch plan names for all tasks
      for (const task of response.value) {
        let planName = 'Unknown Plan';
        
        if (task.planId && !planCache[task.planId]) {
          try {
            const planResponse = await this.graphClient
              .api(`/planner/plans/${task.planId}`)
              .version('v1.0')
              .select('title')
              .get();
            planCache[task.planId] = planResponse.title || 'Unknown Plan';
          } catch (planError) {
            console.warn('Could not fetch plan name:', planError);
            planCache[task.planId] = 'Unknown Plan';
          }
        }
        
        planName = planCache[task.planId] || 'Unknown Plan';

        tasks.push({
          id: task.id,
          title: task.title || 'Untitled Task',
          source: 'Planner',
          dueDate: task.dueDateTime ? new Date(task.dueDateTime) : null,
          status: this._mapPlannerStatus(task.percentComplete),
          progress: task.percentComplete || 0,
          planId: task.planId,
          planName: planName,
          priority: this._mapPlannerPriority(task.priority),
          hasDescription: task.hasDescription || false
        });
      }

      return { tasks, error: null };
    } catch (error) {
      console.error('Error fetching Planner tasks:', error);
      const errorMessage = (error as Error).message || 'Unknown error';
      return { 
        tasks: [], 
        error: `Failed to fetch Planner tasks: ${errorMessage}` 
      };
    }
  }



  /**
   * Fetch tasks from a SharePoint Tasks list
   */
  public async getSharePointTasks(listId: string, siteUrl: string): Promise<ITaskResult> {
    if (!listId || !siteUrl) {
      return { tasks: [], error: null };
    }

    try {
      const endpoint = `${siteUrl}/_api/web/lists(guid'${listId}')/items?$select=Id,Title,DueDate,Status,PercentComplete,Priority,AssignedTo/Title&$expand=AssignedTo&$orderby=DueDate asc`;
      
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error?.message?.value || 'Failed to fetch SharePoint tasks');
      }

      const data = await response.json();
      const tasks: ITask[] = [];

      if (data.value && Array.isArray(data.value)) {
        for (const item of data.value) {
          tasks.push({
            id: item.Id.toString(),
            title: item.Title || 'Untitled Task',
            source: 'SharePoint',
            dueDate: item.DueDate ? new Date(item.DueDate) : null,
            status: this._mapSharePointStatus(item.Status),
            progress: item.PercentComplete || 0,
            assignedTo: item.AssignedTo?.Title ? [item.AssignedTo.Title] : [],
            priority: this._mapSharePointPriority(item.Priority)
          });
        }
      }

      return { tasks, error: null };
    } catch (error) {
      console.error('Error fetching SharePoint tasks:', error);
      const errorMessage = (error as Error).message || 'Unknown error';
      return { 
        tasks: [], 
        error: `Failed to fetch SharePoint tasks: ${errorMessage}` 
      };
    }
  }

  /**
   * Fetch tasks from multiple selected Planner plans
   */
  public async getTasksFromPlans(planIds: string[]): Promise<ITaskResult> {
    if (!planIds || planIds.length === 0) {
      return { tasks: [], error: null };
    }

    try {
      const allTasks: ITask[] = [];
      const planCache: { [key: string]: string } = {};

      for (const planId of planIds) {
        try {
          // Get plan name
          const planResponse = await this.graphClient
            .api(`/planner/plans/${planId}`)
            .version('v1.0')
            .select('title')
            .get();
          
          planCache[planId] = planResponse.title || 'Unknown Plan';

          // Get tasks for this plan
          const tasksResponse = await this.graphClient
            .api(`/planner/plans/${planId}/tasks`)
            .version('v1.0')
            .select('id,title,percentComplete,dueDateTime,planId,hasDescription,priority')
            .get();

          if (tasksResponse && tasksResponse.value) {
            for (const task of tasksResponse.value) {
              allTasks.push({
                id: task.id,
                title: task.title || 'Untitled Task',
                source: 'Planner',
                dueDate: task.dueDateTime ? new Date(task.dueDateTime) : null,
                status: this._mapPlannerStatus(task.percentComplete),
                progress: task.percentComplete || 0,
                planId: task.planId,
                planName: planCache[planId],
                priority: this._mapPlannerPriority(task.priority),
                hasDescription: task.hasDescription || false
              });
            }
          }
        } catch (planError) {
          console.warn(`Could not fetch tasks from plan ${planId}:`, planError);
        }
      }

      return { tasks: allTasks, error: null };
    } catch (error) {
      console.error('Error fetching tasks from plans:', error);
      const errorMessage = (error as Error).message || 'Unknown error';
      return { 
        tasks: [], 
        error: `Failed to fetch tasks from selected plans: ${errorMessage}` 
      };
    }
  }

  /**
   * Get all user's available Planner plans
   */
  public async getUserPlans(): Promise<{ plans: Array<{ id: string; title: string }>; error: string | null }> {
    try {
      // Get all groups the user is a member of
      const groupsResponse = await this.graphClient
        .api('/me/memberOf/microsoft.graph.group')
        .version('v1.0')
        .select('id,displayName')
        .get();

      if (!groupsResponse || !groupsResponse.value) {
        return { plans: [], error: null };
      }

      const plans: Array<{ id: string; title: string }> = [];

      // For each group, try to get its plan
      for (const group of groupsResponse.value) {
        try {
          const planResponse = await this.graphClient
            .api(`/groups/${group.id}/planner/plans`)
            .version('v1.0')
            .select('id,title')
            .get();

          if (planResponse && planResponse.value) {
            for (const plan of planResponse.value) {
              plans.push({
                id: plan.id,
                title: plan.title || 'Untitled Plan'
              });
            }
          }
        } catch (groupError) {
          // Some groups may not have plans, which is fine
          console.debug(`No plans for group ${group.displayName}`);
        }
      }

      return { plans, error: null };
    } catch (error) {
      console.error('Error fetching user plans:', error);
      const errorMessage = (error as Error).message || 'Unknown error';
      return { 
        plans: [], 
        error: `Failed to fetch plans: ${errorMessage}` 
      };
    }
  }

  /**
   * Map Planner status (percentComplete) to common status
   */
  private _mapPlannerStatus(percentComplete: number): 'Not Started' | 'In Progress' | 'Completed' {
    if (percentComplete === 100) {
      return 'Completed';
    } else if (percentComplete > 0) {
      return 'In Progress';
    } else {
      return 'Not Started';
    }
  }



  /**
   * Map SharePoint status to common status
   */
  private _mapSharePointStatus(status: string): 'Not Started' | 'In Progress' | 'Completed' {
    if (!status) return 'Not Started';
    
    switch (status.toLowerCase()) {
      case 'completed':
        return 'Completed';
      case 'in progress':
      case 'inprogress':
        return 'In Progress';
      default:
        return 'Not Started';
    }
  }

  /**
   * Map Planner priority to common priority
   */
  private _mapPlannerPriority(priority: number): 'Low' | 'Medium' | 'High' | 'Urgent' {
    if (priority === 1) return 'Urgent';
    if (priority <= 3) return 'High';
    if (priority <= 5) return 'Medium';
    return 'Low';
  }



  /**
   * Map SharePoint priority to common priority
   */
  private _mapSharePointPriority(priority: string): 'Low' | 'Medium' | 'High' | 'Urgent' {
    if (!priority) return 'Medium';
    
    switch (priority.toLowerCase()) {
      case '(1) high':
      case 'high':
        return 'High';
      case '(3) low':
      case 'low':
        return 'Low';
      default:
        return 'Medium';
    }
  }
}
