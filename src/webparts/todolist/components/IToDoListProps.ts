import { SPHttpClient } from "@microsoft/sp-http";

export interface IToDoListProps {
  websiteUrl: string;
  spHttpClient: SPHttpClient;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
