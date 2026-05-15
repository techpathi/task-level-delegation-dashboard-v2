import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITaskLevelDelegationProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  nintexApiBaseUrl: string;
  tokenListUrl: string;
  tokenColumnName: string;
  tokenTitleValue: string;
}
