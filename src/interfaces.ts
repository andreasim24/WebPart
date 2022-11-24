import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISPListItem {
  Title: string;
  Id: string;
  Status: string;
  DueDate: string;
  Description: string;
}

export interface IListItemState {
  items: ISPListItem[];
  errorMessage: any;
}

export interface IToDoListProps {
  websiteUrl: string;
  spHttpClient: SPHttpClient;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}

export interface IItemListProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
  context: WebPartContext;
}
