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

/**
 * Main TaskDashboard component
 */
const TaskDashboard = (props: ITaskDashboardProps): React.ReactElement => {
  const [tasks, setTasks] = React.useState<ITask[]>([]);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [filter, setFilter] = React.useState<ITaskFilter>({
    sources: ['Planner', 'SharePoint'],
    statuses: ['Not Started', 'In Progress', 'Completed']
  });
  const [lastUpdated, setLastUpdated] = React.useState<Date | null>(null);
  const selectedPlansKey = JSON.stringify(props.selectedPlans || []);

  const isOverdue = (task: ITask): boolean => {
    if (!task.dueDate || task.status === 'Completed') {
      return false;
    }
    return new Date(task.dueDate) < new Date();
  };

  const isDueToday = (task: ITask): boolean => {
    if (!task.dueDate) {
      return false;
    }

    const today = new Date();
    const dueDate = new Date(task.dueDate);

    return (
      today.getDate() === dueDate.getDate() &&
      today.getMonth() === dueDate.getMonth() &&
      today.getFullYear() === dueDate.getFullYear()
    );
  };

  const isDueThisWeek = (task: ITask): boolean => {
    if (!task.dueDate) {
      return false;
    }

    const today = new Date();
    const dueDate = new Date(task.dueDate);
    const weekFromNow = new Date();
    weekFromNow.setDate(today.getDate() + 7);

    return dueDate > today && dueDate <= weekFromNow;
  };

  const sortTasks = (sourceTasks: ITask[]): ITask[] => {
    const sorted = [...sourceTasks];

    if (props.highlightOverdue) {
      sorted.sort((a, b) => {
        const aOverdue = isOverdue(a);
        const bOverdue = isOverdue(b);
        if (aOverdue && !bOverdue) return -1;
        if (!aOverdue && bOverdue) return 1;
        return 0;
      });
    }

    switch (props.sortBy) {
      case 'dueDate':
        sorted.sort((a, b) => {
          if (!a.dueDate && !b.dueDate) return 0;
          if (!a.dueDate) return 1;
          if (!b.dueDate) return -1;
          return new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime();
        });
        break;
      case 'priority': {
        const priorityOrder: Record<string, number> = { Urgent: 1, High: 2, Medium: 3, Low: 4 };
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
  };

  const fetchAllTasks = async (): Promise<void> => {
    setLoading(true);
    setError(null);

    try {
      const allTasks: ITask[] = [];
      const errors: string[] = [];
      const taskService = new TaskService(props.graphClient, props.spHttpClient);

      if (props.enablePlanner) {
        if (props.selectedPlans && props.selectedPlans.length > 0) {
          const planIds = props.selectedPlans.map((plan) => plan.id);
          const plannerResult = await taskService.getTasksFromPlans(planIds);
          if (plannerResult.error) {
            errors.push(plannerResult.error);
          } else {
            allTasks.push(...plannerResult.tasks);
          }
        } else {
          const plannerResult = await taskService.getPlannerTasks();
          if (plannerResult.error) {
            errors.push(plannerResult.error);
          } else {
            allTasks.push(...plannerResult.tasks);
          }
        }
      }

      if (props.enableSharePoint && props.sharePointListId) {
        const spResult = await taskService.getSharePointTasks(props.sharePointListId, props.siteUrl);
        if (spResult.error) {
          errors.push(spResult.error);
        } else {
          allTasks.push(...spResult.tasks);
        }
      }

      setTasks(sortTasks(allTasks));
      setLoading(false);
      setError(errors.length > 0 ? errors.join('; ') : null);
      setLastUpdated(new Date());
    } catch (fetchError) {
      console.error('Error fetching tasks:', fetchError);
      const errorMessage = (fetchError as Error).message || 'Unknown error';
      setLoading(false);
      setError(`Failed to fetch tasks: ${errorMessage}`);
    }
  };

  React.useEffect(() => {
    fetchAllTasks().catch((fetchError) => {
      console.error('Error fetching tasks:', fetchError);
    });
  }, [props.enablePlanner, props.enableSharePoint, props.sharePointListId, selectedPlansKey]);

  const filterTasks = (sourceTasks: ITask[]): ITask[] => {
    let filtered = [...sourceTasks];

    filtered = filtered.filter((task) => filter.sources.indexOf(task.source) !== -1);
    filtered = filtered.filter((task) => filter.statuses.indexOf(task.status) !== -1);

    if (!props.showCompleted) {
      filtered = filtered.filter((task) => task.status !== 'Completed');
    }

    return filtered;
  };

  const groupTasks = (sourceTasks: ITask[]): { [key: string]: ITask[] } => {
    if (props.groupBy === 'none') {
      return { 'All Tasks': sourceTasks };
    }

    const groups: { [key: string]: ITask[] } = {};

    sourceTasks.forEach((task) => {
      let groupKey: string;

      switch (props.groupBy) {
        case 'status':
          groupKey = task.status;
          break;
        case 'source':
          groupKey = task.source;
          break;
        case 'dueDate':
          if (!task.dueDate) {
            groupKey = 'No Due Date';
          } else if (isOverdue(task)) {
            groupKey = 'Overdue';
          } else if (isDueToday(task)) {
            groupKey = 'Due Today';
          } else if (isDueThisWeek(task)) {
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
  };

  const handleRefresh = async (): Promise<void> => {
    await fetchAllTasks();
  };

  const handleSourceFilterChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }

    const source = option.key as 'Planner' | 'SharePoint';

    setFilter((currentFilter) => {
      const currentSources = [...currentFilter.sources];

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

      return {
        ...currentFilter,
        sources: currentSources
      };
    });
  };

  const handleStatusFilterChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) {
      return;
    }

    const status = option.key as 'Not Started' | 'In Progress' | 'Completed';

    setFilter((currentFilter) => {
      const currentStatuses = [...currentFilter.statuses];

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

      return {
        ...currentFilter,
        statuses: currentStatuses
      };
    });
  };

  const getTaskSummary = (sourceTasks: ITask[]): { total: number; overdue: number; dueToday: number } => ({
    total: sourceTasks.length,
    overdue: sourceTasks.filter((task) => isOverdue(task)).length,
    dueToday: sourceTasks.filter((task) => isDueToday(task)).length
  });

  const filteredTasks = filterTasks(tasks);
  const groupedTasks = groupTasks(filteredTasks);
  const summary = getTaskSummary(filteredTasks);
  const { title, hasTeamsContext } = props;
  const sourceOptions: IDropdownOption[] = [
    { key: 'Planner', text: 'Planner', selected: filter.sources.indexOf('Planner') !== -1 },
    { key: 'SharePoint', text: 'SharePoint', selected: filter.sources.indexOf('SharePoint') !== -1 }
  ];
  const statusOptions: IDropdownOption[] = [
    { key: 'Not Started', text: 'Not Started', selected: filter.statuses.indexOf('Not Started') !== -1 },
    { key: 'In Progress', text: 'In Progress', selected: filter.statuses.indexOf('In Progress') !== -1 },
    { key: 'Completed', text: 'Completed', selected: filter.statuses.indexOf('Completed') !== -1 }
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
              onChange={handleSourceFilterChange}
              className={styles.filterDropdown}
              styles={{ dropdown: { minWidth: 150 } }}
            />
            <Dropdown
              placeholder="Filter by Status"
              multiSelect
              options={statusOptions}
              onChange={handleStatusFilterChange}
              className={styles.filterDropdown}
              styles={{ dropdown: { minWidth: 150 } }}
            />
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              title="Refresh tasks"
              ariaLabel="Refresh tasks"
              onClick={handleRefresh}
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
              onDismiss={() => setError(null)}
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
              Loading tasks from {props.enablePlanner && 'Planner'}
              {props.enableSharePoint && ', SharePoint'}...
            </div>
          </div>
        )}

        {/* Empty state */}
        {!loading && tasks.length === 0 && !error && (
          <div className={styles.emptyState}>
            <Icon iconName="TaskManager" className={styles.icon} />
            <h3>No tasks found</h3>
            <p>
              {!props.enablePlanner && !props.enableSharePoint
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
                {props.groupBy !== 'none' && (
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
                      showProgress={props.showProgressBars}
                      showSource={props.showSourceBadges}
                      overdueColor={props.overdueColor}
                      inProgressColor={props.inProgressColor}
                      completedColor={props.completedColor}
                      highlightOverdue={props.highlightOverdue}
                    />
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
      </section>
  );
};

export default TaskDashboard;
