import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { INintexTask } from '../models/INintexTask';

export interface INintexUser {
  id: string;
  email: string;
  firstName?: string;
  lastName?: string;
}

export interface INintexAutoDelegation {
  id: string;
  userId: string;
  effectiveFrom: string;
  effectiveTo: string;
  message?: string;
  standIns: Array<{
    id: string;
    firstName?: string;
    lastName?: string;
    emails?: string[];
  }>;
  createdDate?: string;
  updatedDate?: string;
  fromUserDisplay?: string;
}

export class NintexApiService {
  private httpClient: HttpClient;
  private baseUrl: string;

  constructor(httpClient: HttpClient, baseUrl: string) {
    this.httpClient = httpClient;
    this.baseUrl = baseUrl ? baseUrl.replace(/\/$/, "") : "";
  }

  public async getPendingTasksForUser(userEmail: string, token: string): Promise<INintexTask[]> {
    if (!this.baseUrl) {
      throw new Error("Nintex API Base URL is not configured.");
    }

    const endpoint = `${this.baseUrl}/workflows/v2/tasks?assignee=${encodeURIComponent(userEmail)}&status=Active`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to fetch tasks from Nintex: ${errorMsg}`);
    }

    const data = await response.json();
    return data.tasks || [];
  }

  public async searchTasks(
    assigneeEmail: string,
    taskName: string,
    token: string,
    filters?: {
      status?: string;
      workflowName?: string;
      dateFrom?: string;   // ISO 8601 – used as the API 'from' param (created date range)
      dateTo?: string;     // ISO 8601 – used as the API 'to' param (created date range)
    }
  ): Promise<INintexTask[]> {
    if (!this.baseUrl) {
      throw new Error("Nintex API Base URL is not configured.");
    }

    const status = (filters && filters.status) ? filters.status : 'active';
    let endpoint = `${this.baseUrl}/workflows/v2/tasks?assignee=${encodeURIComponent(assigneeEmail)}&status=${encodeURIComponent(status)}`;

    if (filters && filters.workflowName) {
      endpoint += `&workflowName=${encodeURIComponent(filters.workflowName)}`;
    }
    if (filters && filters.dateFrom) {
      endpoint += `&from=${encodeURIComponent(filters.dateFrom)}`;
    }
    if (filters && filters.dateTo) {
      endpoint += `&to=${encodeURIComponent(filters.dateTo)}`;
    }

    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to fetch tasks from Nintex: ${errorMsg}`);
    }

    const data = await response.json();
    let tasks: INintexTask[] = data.tasks || [];
    
    // Client-side filter: task name
    if (taskName) {
      const lowerTaskName = taskName.toLowerCase();
      tasks = tasks.filter((t: INintexTask) => t.name && t.name.toLowerCase().indexOf(lowerTaskName) !== -1);
    }

    return tasks;
  }

  public async delegateTaskAssignment(taskId: string, assignmentId: string, delegateeEmails: string[], message: string, token: string): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/${taskId}/assignments/${assignmentId}/delegate`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        assignees: delegateeEmails,
        message: message || undefined
      }),
      method: 'PUT'
    };

    const response: HttpClientResponse = await this.httpClient.fetch(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to delegate task ${taskId} assignment ${assignmentId}: ${errorMsg}`);
    }

    return true;
  }



  public async getNintexUserId(email: string, token: string): Promise<string> {
    const endpoint = `${this.baseUrl}/tenants/v1/users/${encodeURIComponent(email)}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to find Nintex user for email ${email}: ${errorMsg}`);
    }

    const data = await response.json();
    if (!data.id) {
      throw new Error(`Nintex user ID not found in response for email ${email}.`);
    }

    return data.id;
  }

  public async createAutoDelegation(delegatorId: string, delegateId: string, startDateTime: Date, endDateTime: Date, token: string, message: string = ""): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        userId: delegatorId,
        effectiveFrom: startDateTime.toISOString(),
        effectiveTo: endDateTime.toISOString(),
        message: message,
        standIns: [
          {
            id: delegateId
          }
        ]
      })
    };

    const response: HttpClientResponse = await this.httpClient.post(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to create auto delegation: ${errorMsg}`);
    }

    return true;
  }

  public async deleteAutoDelegation(delegationId: string, token: string): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations/${delegationId}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      },
      method: 'DELETE'
    };

    const response: HttpClientResponse = await this.httpClient.fetch(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to delete auto delegation: ${errorMsg}`);
    }

    return true;
  }

  public async updateAutoDelegation(delegationId: string, delegatorId: string, delegateId: string, startDateTime: Date, endDateTime: Date, token: string, message: string = ""): Promise<boolean> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations/${delegationId}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        userId: delegatorId,
        effectiveFrom: startDateTime.toISOString(),
        effectiveTo: endDateTime.toISOString(),
        message: message,
        standIns: [
          {
            id: delegateId
          }
        ]
      }),
      method: 'PUT'
    };

    const response: HttpClientResponse = await this.httpClient.fetch(endpoint, HttpClient.configurations.v1, options);
    
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to update auto delegation: ${errorMsg}`);
    }

    return true;
  }

  public async listAutoDelegations(token: string): Promise<INintexAutoDelegation[]> {
    const endpoint = `${this.baseUrl}/workflows/v2/tasks/autodelegations`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      const errorMsg = await response.text();
      throw new Error(`Failed to fetch auto delegations: ${errorMsg}`);
    }

    const data = await response.json();
    return Array.isArray(data) ? data : data.taskAutoDelegations || data.data || [];
  }

  public async searchNintexUsers(filterText: string, token: string): Promise<INintexUser[]> {
    if (!filterText) return [];
    const endpoint = `${this.baseUrl}/tenants/v1/users?filter=${encodeURIComponent(filterText)}&limit=20`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      console.error("Failed to fetch Nintex users", await response.text());
      return [];
    }

    const data = await response.json();
    const rawUsers: INintexUser[] = Array.isArray(data) ? data : data.users || [];

    const lower = filterText.trim().toLowerCase();
    if (!lower) return rawUsers;
    const terms = lower.split(/\s+/).filter(t => t.length > 0);

    return rawUsers.filter((u: INintexUser) => {
      const fullName = `${u.firstName || ''} ${u.lastName || ''}`.trim().toLowerCase();
      const email = (u.email || '').toLowerCase();
      return terms.every(term => fullName.indexOf(term) !== -1 || email.indexOf(term) !== -1);
    });
  }

  public async getNintexUserById(id: string, token: string): Promise<INintexUser | undefined> {
    if (!id) return undefined;
    const endpoint = `${this.baseUrl}/tenants/v1/users/${encodeURIComponent(id)}`;
    
    const options: IHttpClientOptions = {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/json'
      }
    };

    const response: HttpClientResponse = await this.httpClient.get(endpoint, HttpClient.configurations.v1, options);
    if (!response.ok) {
      console.error(`Failed to fetch Nintex user by id ${id}`, await response.text());
      return undefined;
    }

    return await response.json();
  }
}
