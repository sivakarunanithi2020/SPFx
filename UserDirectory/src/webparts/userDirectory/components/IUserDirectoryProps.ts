import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IUserDirectoryProps {
  title: string;
  searchFirstName: boolean;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  clearTextSearchProps?: string;
  pageSize?: number;
  searchProps?: string;
  specficLoc?:string;
  exclude:string[];
}
