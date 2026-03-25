import * as React from 'react';
import importedStyles from './TaskDashboard.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { ITaskDashboardProps, ITaskFilter } from './ITaskDashboardProps';
import { ITask, TaskService } from '../../../services/TaskService';
import { escape } from '@microsoft/sp-lodash-subset';
import TaskCard from './TaskCard';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ITaskDashboardState {
  tasks: ITask[];
  loading: boolean;
  error: string | null;
  filter: ITaskFilter;
  lastUpdated: Date | null;
}

/**
 * Main TaskDashboard component
 */
export default class TaskDashboard extends React.Component<ITaskDashboardProps, ITaskDashboardState> {
  private _taskService: TaskService;

  constructor(props: ITaskDashboardProps) {
    super(props);
    
    this.state = {
      tasks: [],
      loading: false,
      error: null,
      filter: {
        sources: ['Planner', 'SharePoint'],
        statuses: ['Not Started', 'In Progress', 'Completed']
      },
      lastUpdated: null
    };

    this._taskService = new TaskService(this.props.graphClient, this.props.spHttpClient);
  }

  public async componentDidMount(): Promise<void> {
    await this._fetchAllTasks();
  }

  public async componentDidUpdate(prevProps: ITaskDashboardProps): Promise<void> {
    // Re-fetch if data source toggles change
    if (
      prevProps.enablePlanner !== this.props.enablePlanner ||
      prevProps.enableSharePoint !== this.props.enableSharePoint ||
      prevProps.sharePointListId !== this.props.sharePointListId ||
      JSON.stringify(prevProps.selectedPlans) !== JSON.stringify(this.props.selectedPlans)
    ) {
      await this._fetchAllTasks();
    }
  }

  /**
   * Fetch tasks from all enabled sources
   */
  private async _fetchAllTasks(): Promise<void> {
    this.setState({ loading: true, error: null });

    try {
      const allTasks: ITask[] = [];
      const errors: string[] = [];

      // Fetch from Planner
      if (this.props.enablePlanner) {
        if (this.props.selectedPlans && this.props.selectedPlans.length > 0) {
          // Fetch from selected plans
          const planIds = this.props.selectedPlans.map(p => p.id);
          const plannerResult = await this._taskService.getTasksFromPlans(planIds);
          if (plannerResult.error) {
            errors.push(plannerResult.error);
          } else {
            allTasks.push(...plannerResult.tasks);
          }
        } else {
          // Fetch all user's Planner tasks
          const plannerResult = await this._taskService.getPlannerTasks();
          if (plannerResult.error) {
            errors.push(plannerResult.error);
          } else {
            allTasks.push(...plannerResult.tasks);
          }
        }
      }

      // Fetch from SharePoint
      if (this.props.enableSharePoint && this.props.sharePointListId) {
        const spResult = await this._taskService.getSharePointTasks(
          this.props.sharePointListId,
          this.props.siteUrl
        );
        if (spResult.error) {
          errors.push(spResult.error);
        } else {
          allTasks.push(...spResult.tasks);
        }
      }

      // Sort tasks
      const sortedTasks = this._sortTasks(allTasks);

      this.setState({
        tasks: sortedTasks,
        loading: false,
        error: errors.length > 0 ? errors.join('; ') : null,
        lastUpdated: new Date()
      });
    } catch (error) {
      console.error('Error fetching tasks:', error);
      const errorMessage = (error as Error).message || 'Unknown error';
      this.setState({
        loading: false,
        error: `Failed to fetch tasks: ${errorMessage}`
      });
    }
  }

  /**
   * Sort tasks based on sortBy prop
   */
  private _sortTasks(tasks: ITask[]): ITask[] {
    const sorted = [...tasks];

    // Always prioritize overdue tasks if highlighting is enabled
    if (this.props.highlightOverdue) {
      sorted.sort((a, b) => {
        const aOverdue = this._isOverdue(a);
        const bOverdue = this._isOverdue(b);
        if (aOverdue && !bOverdue) return -1;
        if (!aOverdue && bOverdue) return 1;
        return 0;
      });
    }

    // Then apply the selected sort
    switch (this.props.sortBy) {
      case 'dueDate':
        sorted.sort((a, b) => {
          if (!a.dueDate && !b.dueDate) return 0;
          if (!a.dueDate) return 1;
          if (!b.dueDate) return -1;
          return new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime();
        });
        break;
      case 'priority': {
        const priorityOrder = { 'Urgent': 1, 'High': 2, 'Medium': 3, 'Low': 4 };
        sorted.sort((a, b) => {
          const aPriority = priorityOrder[a.priority || 'Medium'];
          const bPriority = priorityOrder[b.priority || 'Medium'];
          return aPriority - bPriority;
        });
        break;
      }
      case 'title':
        sorted.sort((a, b) => a.title.localeCompare(b.title));
        break;
    }

    return sorted;
  }

  /**
   * Filter tasks based on current filter state
   */
  private _filterTasks(tasks: ITask[]): ITask[] {
    let filtered = [...tasks];

    // Filter by source
    filtered = filtered.filter(task => this.state.filter.sources.indexOf(task.source) !== -1);

    // Filter by status
    filtered = filtered.filter(task => this.state.filter.statuses.indexOf(task.status) !== -1);

    // Filter completed tasks if showCompleted is false
    if (!this.props.showCompleted) {
      filtered = filtered.filter(task => task.status !== 'Completed');
    }

    return filtered;
  }

  /**
   * Group tasks based on groupBy prop
   */
  private _groupTasks(tasks: ITask[]): { [key: string]: ITask[] } {
    if (this.props.groupBy === 'none') {
      return { 'All Tasks': tasks };
    }

    const groups: { [key: string]: ITask[] } = {};

    tasks.forEach(task => {
      let groupKey: string;

      switch (this.props.groupBy) {
        case 'status':
          groupKey = task.status;
          break;
        case 'source':
          groupKey = task.source;
          break;
        case 'dueDate':
          if (!task.dueDate) {
            groupKey = 'No Due Date';
          } else if (this._isOverdue(task)) {
            groupKey = 'Overdue';
          } else if (this._isDueToday(task)) {
            groupKey = 'Due Today';
          } else if (this._isDueThisWeek(task)) {
            groupKey = 'Due This Week';
          } else {
            groupKey = 'Later';
          }
          break;
        default:
          groupKey = 'All Tasks';
      }

      if (!groups[groupKey]) {
        groups[groupKey] = [];
      }
      groups[groupKey].push(task);
    });

    return groups;
  }

  /**
   * Check if task is overdue
   */
  private _isOverdue(task: ITask): boolean {
    if (!task.dueDate || task.status === 'Completed') {
      return false;
    }
    return new Date(task.dueDate) < new Date();
  }

  /**
   * Check if task is due today
   */
  private _isDueToday(task: ITask): boolean {
    if (!task.dueDate) return false;
    
    const today = new Date();
    const dueDate = new Date(task.dueDate);
    
    return (
      today.getDate() === dueDate.getDate() &&
      today.getMonth() === dueDate.getMonth() &&
      today.getFullYear() === dueDate.getFullYear()
    );
  }

  /**
   * Check if task is due this week
   */
  private _isDueThisWeek(task: ITask): boolean {
    if (!task.dueDate) return false;
    
    const today = new Date();
    const dueDate = new Date(task.dueDate);
    const weekFromNow = new Date();
    weekFromNow.setDate(today.getDate() + 7);
    
    return dueDate > today && dueDate <= weekFromNow;
  }

  /**
   * Handle refresh button click
   */
  private _handleRefresh = async (): Promise<void> => {
    await this._fetchAllTasks();
  };

  /**
   * Handle source filter change
   */
  private _handleSourceFilterChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;

    const source = option.key as 'Planner' | 'SharePoint';
    const currentSources = [...this.state.filter.sources];

    if (option.selected) {
      if (currentSources.indexOf(source) === -1) {
        currentSources.push(source);
      }
    } else {
      const index = currentSources.indexOf(source);
      if (index > -1) {
        currentSources.splice(index, 1);
      }
    }

    this.setState({
      filter: {
        ...this.state.filter,
        sources: currentSources
      }
    });
  };

  /**
   * Handle status filter change
   */
  private _handleStatusFilterChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;

    const status = option.key as 'Not Started' | 'In Progress' | 'Completed';
    const currentStatuses = [...this.state.filter.statuses];

    if (option.selected) {
      if (currentStatuses.indexOf(status) === -1) {
        currentStatuses.push(status);
      }
    } else {
      const index = currentStatuses.indexOf(status);
      if (index > -1) {
        currentStatuses.splice(index, 1);
      }
    }

    this.setState({
      filter: {
        ...this.state.filter,
        statuses: currentStatuses
      }
    });
  };

  /**
   * Get task count summary
   */
  private _getTaskSummary(tasks: ITask[]): { total: number; overdue: number; dueToday: number } {
    return {
      total: tasks.length,
      overdue: tasks.filter(t => this._isOverdue(t)).length,
      dueToday: tasks.filter(t => this._isDueToday(t)).length
    };
  }

  /**
   * Render method
   */
  public render(): React.ReactElement<ITaskDashboardProps> {
    const { loading, error, tasks, lastUpdated } = this.state;
    const { title, hasTeamsContext } = this.props;

    const filteredTasks = this._filterTasks(tasks);
    const groupedTasks = this._groupTasks(filteredTasks);
    const summary = this._getTaskSummary(filteredTasks);

    // Source filter options
    const sourceOptions: IDropdownOption[] = [
      { key: 'Planner', text: 'Planner', selected: this.state.filter.sources.indexOf('Planner') !== -1 },
      { key: 'SharePoint', text: 'SharePoint', selected: this.state.filter.sources.indexOf('SharePoint') !== -1 }
    ];

    // Status filter options
    const statusOptions: IDropdownOption[] = [
      { key: 'Not Started', text: 'Not Started', selected: this.state.filter.statuses.indexOf('Not Started') !== -1 },
      { key: 'In Progress', text: 'In Progress', selected: this.state.filter.statuses.indexOf('In Progress') !== -1 },
      { key: 'Completed', text: 'Completed', selected: this.state.filter.statuses.indexOf('Completed') !== -1 }
    ];

    return (
      <section className={`${styles.taskDashboard} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* Header */}
        <div className={styles.header}>
          <div className={styles.titleSection}>
            <h2>{escape(title)}</h2>
            {!loading && summary.total > 0 && (
              <div className={styles.summary}>
                <div className={styles.summaryItem}>
                  <Icon iconName="TaskManager" className={styles.icon} />
                  <span>{summary.total} {summary.total === 1 ? 'task' : 'tasks'}</span>
                </div>
                {summary.dueToday > 0 && (
                  <div className={`${styles.summaryItem} ${styles.dueToday}`}>
                    <Icon iconName="Calendar" className={styles.icon} />
                    <span>{summary.dueToday} due today</span>
                  </div>
                )}
                {summary.overdue > 0 && (
                  <div className={`${styles.summaryItem} ${styles.overdue}`}>
                    <Icon iconName="Warning" className={styles.icon} />
                    <span>{summary.overdue} overdue</span>
                  </div>
                )}
              </div>
            )}
          </div>

          <div className={styles.controls}>
            <Dropdown
              placeholder="Filter by Source"
              multiSelect
              options={sourceOptions}
              onChange={this._handleSourceFilterChange}
              className={styles.filterDropdown}
              styles={{ dropdown: { minWidth: 150 } }}
            />
            <Dropdown
              placeholder="Filter by Status"
              multiSelect
              options={statusOptions}
              onChange={this._handleStatusFilterChange}
              className={styles.filterDropdown}
              styles={{ dropdown: { minWidth: 150 } }}
            />
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              title="Refresh tasks"
              ariaLabel="Refresh tasks"
              onClick={this._handleRefresh}
              disabled={loading}
            />
            {lastUpdated && (
              <div className={styles.lastUpdated}>
                Updated: {lastUpdated.toLocaleTimeString()}
              </div>
            )}
          </div>
        </div>

        {/* Error message */}
        {error && (
          <div className={styles.errorContainer}>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={() => this.setState({ error: null })}
              dismissButtonAriaLabel="Close"
            >
              {error}
            </MessageBar>
          </div>
        )}

        {/* Loading state */}
        {loading && (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.large} label="Fetching tasks..." />
            <div className={styles.loadingText}>
              Loading tasks from {this.props.enablePlanner && 'Planner'}
              {this.props.enableSharePoint && ', SharePoint'}...
            </div>
          </div>
        )}

        {/* Empty state */}
        {!loading && tasks.length === 0 && !error && (
          <div className={styles.emptyState}>
            <Icon iconName="TaskManager" className={styles.icon} />
            <h3>No tasks found</h3>
            <p>
              {!this.props.enablePlanner && !this.props.enableSharePoint
                ? 'Please enable at least one data source in the web part properties.'
                : 'You don\'t have any tasks assigned. Create tasks in Planner or SharePoint to see them here.'}
            </p>
          </div>
        )}

        {/* No results after filtering */}
        {!loading && tasks.length > 0 && filteredTasks.length === 0 && (
          <div className={styles.noTasksMessage}>
            No tasks match the current filters. Try adjusting your filter settings.
          </div>
        )}

        {/* Task groups */}
        {!loading && filteredTasks.length > 0 && (
          <div className={styles.taskContainer}>
            {Object.keys(groupedTasks).map((groupKey, index) => (
              <div key={index} className={styles.groupSection}>
                {this.props.groupBy !== 'none' && (
                  <div className={styles.groupHeader}>
                    <span>{groupKey}</span>
                    <span className={styles.groupCount}>
                      ({groupedTasks[groupKey].length} {groupedTasks[groupKey].length === 1 ? 'task' : 'tasks'})
                    </span>
                  </div>
                )}
                <div className={styles.taskGrid}>
                  {groupedTasks[groupKey].map((task) => (
                    <TaskCard
                      key={task.id}
                      task={task}
                      showProgress={this.props.showProgressBars}
                      showSource={this.props.showSourceBadges}
                      overdueColor={this.props.overdueColor}
                      inProgressColor={this.props.inProgressColor}
                      completedColor={this.props.completedColor}
                      highlightOverdue={this.props.highlightOverdue}
                    />
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
      </section>
    );
  }
}
