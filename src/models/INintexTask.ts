export interface INintexTaskAssignment {
  id: string;  // assignmentId
  assignee: string;
}

export interface INintexTask {
  key?: string;          // required by IObjectWithKey for key-based selection tracking
  id: string;            // taskId
  name: string;
  status: string;
  createdDate: string;
  dueDate: string;
  description: string;
  workflowName: string;
  taskAssignments: INintexTaskAssignment[];
}
