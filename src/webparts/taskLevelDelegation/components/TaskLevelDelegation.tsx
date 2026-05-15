import * as React from 'react';
import type { ITaskLevelDelegationProps } from './ITaskLevelDelegationProps';
import {
  NormalPeoplePicker,
  IPersonaProps,
  SearchBox,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  IColumn,
  Selection,
  SelectionMode,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  ProgressIndicator,
  IconButton,
  Icon,
  Dropdown,
  IDropdownOption,
  DatePicker
} from '@fluentui/react';
import { TokenService } from '../../../services/TokenService';
import { NintexApiService, INintexUser } from '../../../services/NintexApiService';
import { SPHttpClient } from '@microsoft/sp-http';
import { INintexTask } from '../../../models/INintexTask';
import styles from './TaskLevelDelegation.module.scss';

export const TaskLevelDelegation: React.FC<ITaskLevelDelegationProps> = (props) => {
  const [nintexToken, setNintexToken] = React.useState<string>('');
  const [assigneeUser, setAssigneeUser] = React.useState<IPersonaProps | undefined>(undefined);
  const [taskSearchText, setTaskSearchText] = React.useState<string>('');
  const [tasks, setTasks] = React.useState<INintexTask[]>([]);
  const [selectedTasks, setSelectedTasks] = React.useState<INintexTask[]>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [isInitializing, setIsInitializing] = React.useState<boolean>(true);
  const [errorMsg, setErrorMsg] = React.useState<string>('');
  const [successMsg, setSuccessMsg] = React.useState<string>('');
  const [hasSearched, setHasSearched] = React.useState<boolean>(false);

  // Delegation dialog state
  const [isDelegateDialogOpen, setIsDelegateDialogOpen] = React.useState<boolean>(false);
  const [delegateToUser, setDelegateToUser] = React.useState<IPersonaProps | undefined>(undefined);
  const [delegationMessage, setDelegationMessage] = React.useState<string>('');
  const [isDelegating, setIsDelegating] = React.useState<boolean>(false);
  const [delegateProgress, setDelegateProgress] = React.useState<number>(0);
  const [delegateProgressDesc, setDelegateProgressDesc] = React.useState<string>('');
  const [dialogErrorMsg, setDialogErrorMsg] = React.useState<string>('');
  const [dialogSuccessMsg, setDialogSuccessMsg] = React.useState<string>('');

  // Advanced filter state
  const [isFilterPanelOpen, setIsFilterPanelOpen] = React.useState<boolean>(false);
  // Applied filters (active in search)
  const [appliedFilters, setAppliedFilters] = React.useState<{
    status: string;
    workflowName: string;
    createdFrom: Date | undefined;
    createdTo: Date | undefined;
    dueDateFrom: Date | undefined;
    dueDateTo: Date | undefined;
  }>({ status: 'active', workflowName: '', createdFrom: undefined, createdTo: undefined, dueDateFrom: undefined, dueDateTo: undefined });
  // Draft filters (editing in panel)
  const [draftFilters, setDraftFilters] = React.useState<{
    status: string;
    workflowName: string;
    createdFrom: Date | undefined;
    createdTo: Date | undefined;
    dueDateFrom: Date | undefined;
    dueDateTo: Date | undefined;
  }>({ status: 'active', workflowName: '', createdFrom: undefined, createdTo: undefined, dueDateFrom: undefined, dueDateTo: undefined });

  const selectionRef = React.useRef(
    new Selection({
      onSelectionChanged: () => {
        setSelectedTasks(selectionRef.current.getSelection() as INintexTask[]);
      }
    })
  );

  React.useEffect(() => {
    const fetchToken = async (): Promise<void> => {
      try {
        setIsInitializing(true);
        const spHttpClient: SPHttpClient = props.context.spHttpClient;
        const tokenService = new TokenService(
          spHttpClient,
          props.tokenListUrl,
          props.tokenTitleValue,
          props.tokenColumnName
        );
        const token = await tokenService.getToken();
        setNintexToken(token);
      } catch (err) {
        console.error('Error fetching initial token:', err);
        setErrorMsg('Failed to initialize Nintex API token.');
      } finally {
        setIsInitializing(false);
      }
    };
    fetchToken().catch(console.error);
  }, [props.tokenListUrl, props.tokenTitleValue, props.tokenColumnName]);

  const onResolveSuggestions = async (filterText: string): Promise<IPersonaProps[]> => {
    if (!filterText || filterText.length < 2 || !nintexToken) return [];
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
      const users = await nintexApiService.searchNintexUsers(filterText, nintexToken);
      return users.map((u: INintexUser) => ({
        text: u.firstName ? `${u.firstName} ${u.lastName}` : u.email,
        secondaryText: u.email,
        id: u.id,
        imageUrl: undefined
      }));
    } catch (e) {
      console.error('Error fetching Nintex users:', e);
      return [];
    }
  };

  const onResolveDelegateSuggestions = async (filterText: string): Promise<IPersonaProps[]> => {
    const results = await onResolveSuggestions(filterText);
    if (assigneeUser && assigneeUser.secondaryText) {
      return results.filter(u => u.secondaryText?.toLowerCase() !== assigneeUser.secondaryText?.toLowerCase());
    }
    return results;
  };

  const handleSearch = async (overrideSearchText?: string): Promise<void> => {
    const currentSearchText = overrideSearchText !== undefined ? overrideSearchText : taskSearchText;
    setErrorMsg('');
    setSuccessMsg('');
    if (!assigneeUser || !assigneeUser.secondaryText) {
      setErrorMsg('Please select an Assignee.');
      return;
    }
    setIsLoading(true);
    setHasSearched(true);
    try {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
      const fetchedTasks = await nintexApiService.searchTasks(
        assigneeUser.secondaryText,
        currentSearchText,
        nintexToken,
        {
          status: appliedFilters.status || 'active',
          workflowName: appliedFilters.workflowName || undefined,
          dateFrom: appliedFilters.createdFrom ? appliedFilters.createdFrom.toISOString() : undefined,
          dateTo: appliedFilters.createdTo ? appliedFilters.createdTo.toISOString() : undefined,
          dueDateFrom: appliedFilters.dueDateFrom ? appliedFilters.dueDateFrom.toISOString() : undefined,
          dueDateTo: appliedFilters.dueDateTo ? appliedFilters.dueDateTo.toISOString() : undefined
        }
      );
      selectionRef.current.setAllSelected(false);
      setTasks(fetchedTasks);
      setSelectedTasks([]);
    } catch (err) {
      console.error('Error searching tasks:', err);
      setErrorMsg(err.message || 'An unexpected error occurred while searching tasks.');
    } finally {
      setIsLoading(false);
    }
  };

  // ── Advanced filter helpers ──
  const statusOptions: IDropdownOption[] = [
    { key: 'active', text: 'Active' },
    { key: 'expired', text: 'Expired' },
    { key: 'complete', text: 'Complete' },
    { key: 'overridden', text: 'Overridden' },
    { key: 'terminated', text: 'Terminated' },
    { key: 'all', text: 'All Statuses' }
  ];

  const getActiveFilterCount = (): number => {
    let count = 0;
    if (appliedFilters.status && appliedFilters.status !== 'active') count++;
    if (appliedFilters.workflowName) count++;
    if (appliedFilters.createdFrom) count++;
    if (appliedFilters.createdTo) count++;
    if (appliedFilters.dueDateFrom) count++;
    if (appliedFilters.dueDateTo) count++;
    return count;
  };

  const openFilterPanel = (): void => {
    setDraftFilters({ ...appliedFilters });
    setIsFilterPanelOpen(true);
  };

  const applyFilters = (): void => {
    setAppliedFilters({ ...draftFilters });
    setIsFilterPanelOpen(false);
  };

  const clearAllFilters = (): void => {
    const defaults = { status: 'active', workflowName: '', createdFrom: undefined, createdTo: undefined, dueDateFrom: undefined, dueDateTo: undefined };
    setDraftFilters(defaults);
    setAppliedFilters(defaults);
    setIsFilterPanelOpen(false);
  };

  const removeAppliedFilter = (key: string): void => {
    const updated = { ...appliedFilters };
    switch (key) {
      case 'status': updated.status = 'active'; break;
      case 'workflowName': updated.workflowName = ''; break;
      case 'createdFrom': updated.createdFrom = undefined; break;
      case 'createdTo': updated.createdTo = undefined; break;
      case 'dueDateFrom': updated.dueDateFrom = undefined; break;
      case 'dueDateTo': updated.dueDateTo = undefined; break;
    }
    setAppliedFilters(updated);
  };

  const formatDate = (d: Date): string => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });

  const openDelegateDialog = (): void => {
    setDelegateToUser(undefined);
    setDelegationMessage('');
    setDialogErrorMsg('');
    setDialogSuccessMsg('');
    setDelegateProgress(0);
    setDelegateProgressDesc('');
    setIsDelegateDialogOpen(true);
  };

  const handleDelegate = async (): Promise<void> => {
    setDialogErrorMsg('');
    setDialogSuccessMsg('');

    if (!delegateToUser || !delegateToUser.secondaryText) {
      setDialogErrorMsg('Please select a user to delegate to.');
      return;
    }

    if (assigneeUser && assigneeUser.secondaryText && 
        delegateToUser.secondaryText.toLowerCase() === assigneeUser.secondaryText.toLowerCase()) {
      setDialogErrorMsg('The delegate cannot be the same as the current assignee.');
      return;
    }

    if (selectedTasks.length === 0) {
      setDialogErrorMsg('No tasks selected.');
      return;
    }

    setIsDelegating(true);
    setDelegateProgress(0);
    setDelegateProgressDesc('Starting delegation...');

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const nintexApiService = new NintexApiService(props.context.httpClient as any, props.nintexApiBaseUrl);
    const delegateeEmail = delegateToUser.secondaryText;
    const errors: string[] = [];
    let completed = 0;

    for (const task of selectedTasks) {
      const assignments = task.taskAssignments || [];
      if (assignments.length === 0) {
        errors.push(`Task "${task.name}" has no assignments to delegate.`);
        completed++;
        setDelegateProgress(completed / selectedTasks.length);
        continue;
      }

      for (const assignment of assignments) {
        // Skip if this specific assignment is already assigned to the delegate
        if (assignment.assignee.toLowerCase() === delegateeEmail.toLowerCase()) {
          continue;
        }

        // Only delegate the assignment of the user we originally searched for
        if (assigneeUser && assigneeUser.secondaryText && 
            assignment.assignee.toLowerCase() !== assigneeUser.secondaryText.toLowerCase()) {
          continue;
        }

        setDelegateProgressDesc(`Delegating "${task.name}"...`);
        try {
          await nintexApiService.delegateTaskAssignment(
            task.id,
            assignment.id,
            [delegateeEmail],
            delegationMessage,
            nintexToken
          );
        } catch (err) {
          errors.push(`"${task.name}": ${err.message}`);
        }
      }

      completed++;
      setDelegateProgress(completed / selectedTasks.length);
    }

    setIsDelegating(false);

    if (errors.length > 0) {
      setDialogErrorMsg(`Some delegations failed:\n${errors.join('\n')}`);
    } else {
      setIsDelegateDialogOpen(false);
      const msg = `Successfully delegated ${selectedTasks.length} task(s) to ${delegateToUser.text}.`;
      setTaskSearchText('');
      // Refresh task list with cleared search
      await handleSearch('');
      setSuccessMsg(msg);
      // Auto-dismiss after 5 seconds
      setTimeout(() => setSuccessMsg(''), 5000);
    }
  };

  // ── Sorting state ──
  const [sortColumn, setSortColumn] = React.useState<string>('');
  const [isSortDescending, setIsSortDescending] = React.useState<boolean>(false);

  const onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    const newIsSortedDescending = column.key === sortColumn ? !isSortDescending : false;
    setSortColumn(column.key);
    setIsSortDescending(newIsSortedDescending);

    const fieldName = column.fieldName as keyof INintexTask;
    const sorted = [...tasks].sort((a, b) => {
      let valA = a[fieldName] || '';
      let valB = b[fieldName] || '';

      // Date fields: compare as dates
      if (fieldName === 'createdDate' || fieldName === 'dueDate') {
        const dateA = valA ? new Date(valA as string).getTime() : 0;
        const dateB = valB ? new Date(valB as string).getTime() : 0;
        return newIsSortedDescending ? dateB - dateA : dateA - dateB;
      }

      // String fields: compare case-insensitively
      valA = (valA as string).toString().toLowerCase();
      valB = (valB as string).toString().toLowerCase();
      if (valA < valB) return newIsSortedDescending ? 1 : -1;
      if (valA > valB) return newIsSortedDescending ? -1 : 1;
      return 0;
    });

    setTasks(sorted);
  };

  const taskColumns: IColumn[] = [
    {
      key: 'colName',
      name: 'Task Title',
      fieldName: 'name',
      minWidth: 160,
      maxWidth: 300,
      isResizable: true,
      isSorted: sortColumn === 'colName',
      isSortedDescending: sortColumn === 'colName' && isSortDescending,
      onColumnClick: onColumnClick,
      onRender: (item: INintexTask) => (
        <span style={{ fontWeight: 500, color: '#323130' }}>{item.name}</span>
      )
    },
    {
      key: 'colWorkflow',
      name: 'Workflow',
      fieldName: 'workflowName',
      minWidth: 130,
      maxWidth: 220,
      isResizable: true,
      isSorted: sortColumn === 'colWorkflow',
      isSortedDescending: sortColumn === 'colWorkflow' && isSortDescending,
      onColumnClick: onColumnClick,
      onRender: (item: INintexTask) => (
        <span style={{ color: '#605e5c' }}>{item.workflowName}</span>
      )
    },
    {
      key: 'colStatus',
      name: 'Status',
      fieldName: 'status',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      isSorted: sortColumn === 'colStatus',
      isSortedDescending: sortColumn === 'colStatus' && isSortDescending,
      onColumnClick: onColumnClick,
      onRender: (item: INintexTask) => {
        const status = item.status || '';
        const isActive = status.toLowerCase() === 'active';
        return (
          <span className={`${styles.statusBadge} ${isActive ? styles.statusActive : styles.statusOther}`}>
            {status}
          </span>
        );
      }
    },
    {
      key: 'colCreated',
      name: 'Created',
      fieldName: 'createdDate',
      minWidth: 100,
      maxWidth: 130,
      isResizable: true,
      isSorted: sortColumn === 'colCreated',
      isSortedDescending: sortColumn === 'colCreated' && isSortDescending,
      onColumnClick: onColumnClick,
      onRender: (item: INintexTask) => {
        if (!item.createdDate) return '';
        const d = new Date(item.createdDate);
        return (
          <span style={{ color: '#605e5c', fontSize: '12px' }}>
            {d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}
          </span>
        );
      }
    },
    {
      key: 'colDueDate',
      name: 'Due Date',
      fieldName: 'dueDate',
      minWidth: 100,
      maxWidth: 130,
      isResizable: true,
      isSorted: sortColumn === 'colDueDate',
      isSortedDescending: sortColumn === 'colDueDate' && isSortDescending,
      onColumnClick: onColumnClick,
      onRender: (item: INintexTask) => {
        if (!item.dueDate) return <span style={{ color: '#a19f9d', fontSize: '12px' }}>—</span>;
        const d = new Date(item.dueDate);
        const isPast = d < new Date();
        return (
          <span className={isPast ? styles.dueDatePast : undefined} style={{ fontSize: '12px' }}>
            {d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}
          </span>
        );
      }
    }
  ];

  return (
    <div className={styles.dashboardRoot}>

      {/* ── Branded Header ── */}
      <div className={styles.headerBar}>
        <Icon iconName="TaskGroup" className={styles.headerIcon} />
        <span className={styles.headerTitle}>Task Delegation</span>
        <span className={styles.headerSubtitle}>Nintex Workflow</span>
      </div>

      {isInitializing ? (
        <div className={styles.initState}>
          <Spinner size={SpinnerSize.large} label="Initializing component..." />
        </div>
      ) : (
        <>
          {/* ── Search Controls Card ── */}
          <div className={styles.searchCard}>
            <div className={styles.searchRow}>
              <div className={styles.searchField}>
                <label className={styles.searchFieldLabel}>Assignee</label>
                <NormalPeoplePicker
                  onResolveSuggestions={onResolveSuggestions}
                  itemLimit={1}
                  disabled={!nintexToken || isLoading}
                  onChange={(items) => setAssigneeUser(items && items.length > 0 ? items[0] : undefined)}
                  selectedItems={assigneeUser ? [assigneeUser] : []}
                  resolveDelay={500}
                  styles={{ root: { maxWidth: '100%' } }}
                  inputProps={{ placeholder: 'Search Nintex user...' }}
                />
              </div>

              <div className={styles.searchField}>
                <label className={styles.searchFieldLabel}>Task Search</label>
                <SearchBox
                  placeholder="Filter by task name..."
                  value={taskSearchText}
                  onChange={(e, val) => setTaskSearchText(val || '')}
                  onSearch={() => handleSearch()}
                  disabled={!nintexToken || isLoading}
                />
              </div>

              <div style={{ position: 'relative' }}>
                <IconButton
                  iconProps={{ iconName: 'FilterSettings' }}
                  title="Advanced Filters"
                  ariaLabel="Advanced Filters"
                  onClick={openFilterPanel}
                  disabled={!nintexToken || isLoading}
                  className={`${styles.filterButton} ${getActiveFilterCount() > 0 ? styles.filterButtonActive : ''}`}
                />
                {getActiveFilterCount() > 0 && (
                  <span className={styles.filterBadge}>{getActiveFilterCount()}</span>
                )}
              </div>

              <IconButton
                iconProps={{ iconName: 'Search' }}
                title="Search Tasks"
                ariaLabel="Search Tasks"
                onClick={() => handleSearch()}
                disabled={!nintexToken || isLoading}
                className={styles.searchButton}
              />
            </div>
          </div>

          {/* ── Active Filters Chips ── */}
          {getActiveFilterCount() > 0 && (
            <div className={styles.activeFiltersBar}>
              <span className={styles.activeFiltersLabel}>Filters:</span>
              {appliedFilters.status && appliedFilters.status !== 'active' && (
                <span className={styles.filterChip}>
                  Status: {statusOptions.find((o: IDropdownOption) => o.key === appliedFilters.status)?.text || appliedFilters.status}
                  <Icon iconName="Cancel" className={styles.chipClose} onClick={() => removeAppliedFilter('status')} />
                </span>
              )}
              {appliedFilters.workflowName && (
                <span className={styles.filterChip}>
                  Workflow: {appliedFilters.workflowName}
                  <Icon iconName="Cancel" className={styles.chipClose} onClick={() => removeAppliedFilter('workflowName')} />
                </span>
              )}
              {appliedFilters.createdFrom && (
                <span className={styles.filterChip}>
                  Created from: {formatDate(appliedFilters.createdFrom)}
                  <Icon iconName="Cancel" className={styles.chipClose} onClick={() => removeAppliedFilter('createdFrom')} />
                </span>
              )}
              {appliedFilters.createdTo && (
                <span className={styles.filterChip}>
                  Created to: {formatDate(appliedFilters.createdTo)}
                  <Icon iconName="Cancel" className={styles.chipClose} onClick={() => removeAppliedFilter('createdTo')} />
                </span>
              )}
              {appliedFilters.dueDateFrom && (
                <span className={styles.filterChip}>
                  Due from: {formatDate(appliedFilters.dueDateFrom)}
                  <Icon iconName="Cancel" className={styles.chipClose} onClick={() => removeAppliedFilter('dueDateFrom')} />
                </span>
              )}
              {appliedFilters.dueDateTo && (
                <span className={styles.filterChip}>
                  Due to: {formatDate(appliedFilters.dueDateTo)}
                  <Icon iconName="Cancel" className={styles.chipClose} onClick={() => removeAppliedFilter('dueDateTo')} />
                </span>
              )}
            </div>
          )}

          {/* ── Messages ── */}
          {(errorMsg || successMsg) && (
            <div className={styles.messagesArea}>
              {errorMsg && (
                <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg('')}>
                  {errorMsg}
                </MessageBar>
              )}
              {successMsg && (
                <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMsg('')}>
                  {successMsg}
                </MessageBar>
              )}
            </div>
          )}

          {/* ── Results Area (scrollable) ── */}
          {isLoading ? (
            <div className={styles.loadingState}>
              <Spinner size={SpinnerSize.large} label="Searching tasks..." />
            </div>
          ) : tasks.length > 0 ? (
            <div className={styles.resultsArea}>
              {/* Results header toolbar */}
              <div className={styles.resultsHeader}>
                <span className={styles.resultsCount}>
                  {tasks.length} task{tasks.length !== 1 ? 's' : ''} found
                  {selectedTasks.length > 0 && (
                    <span className={styles.selectedBadge}>
                      {selectedTasks.length} selected
                    </span>
                  )}
                </span>
                {selectedTasks.length > 0 && (
                  <PrimaryButton
                    text={`Delegate (${selectedTasks.length})`}
                    iconProps={{ iconName: 'People' }}
                    onClick={openDelegateDialog}
                    className={styles.delegateButton}
                  />
                )}
              </div>

              {/* Scrollable task list */}
              <div className={styles.taskListWrapper}>
                <DetailsList
                  items={tasks}
                  columns={taskColumns}
                  setKey="taskSet"
                  selection={selectionRef.current}
                  selectionMode={SelectionMode.multiple}
                  selectionPreservedOnEmptyClick={true}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                  checkButtonAriaLabel="Row checkbox"
                />
              </div>
            </div>
          ) : hasSearched ? (
            <div className={styles.resultsArea}>
              <div className={styles.emptyState}>
                <Icon iconName="SearchIssue" className={styles.emptyIcon} />
                <span className={styles.emptyText}>No tasks found</span>
                <span className={styles.emptyHint}>Try adjusting your search criteria or selecting a different assignee.</span>
              </div>
            </div>
          ) : (
            <div className={styles.resultsArea}>
              <div className={styles.emptyState}>
                <Icon iconName="Search" className={styles.emptyIcon} />
                <span className={styles.emptyText}>Search for tasks</span>
                <span className={styles.emptyHint}>Select an assignee and click search to find delegatable tasks.</span>
              </div>
            </div>
          )}
        </>
      )}

      {/* Delegate Dialog */}
      <Dialog
        hidden={!isDelegateDialogOpen}
        onDismiss={() => {
          if (!isDelegating) setIsDelegateDialogOpen(false);
        }}
        dialogContentProps={{
          type: DialogType.normal,
          title: <span style={{ color: '#d83b01' }}>Delegate Selected Tasks</span>,
          showCloseButton: !isDelegating
        }}
        modalProps={{
          isBlocking: false,
          styles: { main: { maxWidth: 540, width: '100%' } }
        }}
      >
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', padding: '0 5px' }}>

          <MessageBar messageBarType={MessageBarType.info}>
            Delegating <strong>{selectedTasks.length}</strong> task{selectedTasks.length !== 1 ? 's' : ''} to a new user.
          </MessageBar>

          <div>
            <label style={{ display: 'block', fontWeight: 600, fontSize: '12px', marginBottom: '4px' }}>
              Delegate to <span style={{ color: '#d13438' }}>*</span>
            </label>
            <NormalPeoplePicker
              onResolveSuggestions={onResolveDelegateSuggestions}
              itemLimit={1}
              disabled={isDelegating}
              onChange={(items) => setDelegateToUser(items && items.length > 0 ? items[0] : undefined)}
              selectedItems={delegateToUser ? [delegateToUser] : []}
              resolveDelay={500}
              styles={{ root: { maxWidth: '100%' } }}
              inputProps={{ placeholder: 'Search for a Nintex user...' }}
            />
          </div>

          <div>
            <label style={{ display: 'block', fontWeight: 600, fontSize: '12px', marginBottom: '4px' }}>
              Message (optional)
            </label>
            <TextField
              multiline
              rows={3}
              value={delegationMessage}
              onChange={(e, val) => setDelegationMessage(val || '')}
              disabled={isDelegating}
              placeholder="Optional reason for delegation..."
            />
          </div>

          {isDelegating && (
            <ProgressIndicator
              label="Delegating tasks..."
              description={delegateProgressDesc}
              percentComplete={delegateProgress}
            />
          )}

          {dialogErrorMsg && (
            <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
              {dialogErrorMsg}
            </MessageBar>
          )}

          {dialogSuccessMsg && (
            <MessageBar messageBarType={MessageBarType.success}>
              {dialogSuccessMsg}
            </MessageBar>
          )}
        </div>

        <DialogFooter>
          {isDelegating && (
            <span style={{ display: 'inline-block', verticalAlign: 'middle', marginRight: '10px' }}>
              <Spinner size={SpinnerSize.small} />
            </span>
          )}
          <DefaultButton
            onClick={() => setIsDelegateDialogOpen(false)}
            text="Cancel"
            disabled={isDelegating}
          />
          <PrimaryButton
            onClick={handleDelegate}
            text="Delegate"
            disabled={isDelegating || !delegateToUser}
          />
        </DialogFooter>
      </Dialog>

      {/* ── Advanced Filters Panel ── */}
      {isFilterPanelOpen && (
        <div className={styles.filterOverlay}>
          <div className={styles.filterBackdrop} onClick={() => setIsFilterPanelOpen(false)} />
          <div className={styles.filterPanel}>
            {/* Panel Header */}
            <div className={styles.filterPanelHeader}>
              <span className={styles.filterPanelTitle}>Advanced Filters</span>
              <IconButton
                iconProps={{ iconName: 'Cancel' }}
                ariaLabel="Close"
                onClick={() => setIsFilterPanelOpen(false)}
              />
            </div>

            {/* Panel Body */}
            <div className={styles.filterPanelBody}>
              {/* Status */}
              <div className={styles.filterSection}>
                <span className={styles.filterSectionTitle}>Task Status</span>
                <Dropdown
                  selectedKey={draftFilters.status}
                  options={statusOptions}
                  onChange={(e, option) => option && setDraftFilters({ ...draftFilters, status: option.key as string })}
                />
              </div>

              {/* Workflow Name */}
              <div className={styles.filterSection}>
                <span className={styles.filterSectionTitle}>Workflow Name</span>
                <TextField
                  value={draftFilters.workflowName}
                  onChange={(e, val) => setDraftFilters({ ...draftFilters, workflowName: val || '' })}
                  placeholder="Filter by workflow name..."
                />
              </div>

              {/* Created Date Range */}
              <div className={styles.filterSection}>
                <span className={styles.filterSectionTitle}>Created Date Range</span>
                <div className={styles.filterDateRow}>
                  <div>
                    <DatePicker
                      placeholder="From..."
                      value={draftFilters.createdFrom}
                      onSelectDate={(date) => setDraftFilters({ ...draftFilters, createdFrom: date || undefined })}
                      maxDate={draftFilters.createdTo || new Date()}
                    />
                  </div>
                  <div>
                    <DatePicker
                      placeholder="To..."
                      value={draftFilters.createdTo}
                      onSelectDate={(date) => setDraftFilters({ ...draftFilters, createdTo: date || undefined })}
                      minDate={draftFilters.createdFrom || undefined}
                      maxDate={new Date()}
                    />
                  </div>
                </div>
              </div>

              {/* Due Date Range */}
              <div className={styles.filterSection}>
                <span className={styles.filterSectionTitle}>Due Date Range</span>
                <div className={styles.filterDateRow}>
                  <div>
                    <DatePicker
                      placeholder="From..."
                      value={draftFilters.dueDateFrom}
                      onSelectDate={(date) => setDraftFilters({ ...draftFilters, dueDateFrom: date || undefined })}
                      maxDate={draftFilters.dueDateTo || undefined}
                    />
                  </div>
                  <div>
                    <DatePicker
                      placeholder="To..."
                      value={draftFilters.dueDateTo}
                      onSelectDate={(date) => setDraftFilters({ ...draftFilters, dueDateTo: date || undefined })}
                      minDate={draftFilters.dueDateFrom || undefined}
                    />
                  </div>
                </div>
              </div>
            </div>

            {/* Panel Footer */}
            <div className={styles.filterPanelFooter}>
              <DefaultButton text="Clear All" onClick={clearAllFilters} />
              <PrimaryButton text="Apply Filters" onClick={applyFilters} />
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default TaskLevelDelegation;
