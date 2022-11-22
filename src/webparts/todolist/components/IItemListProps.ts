import { SPHttpClient } from "@microsoft/sp-http";

export interface IItemListProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
}
