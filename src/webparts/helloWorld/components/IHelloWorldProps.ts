import { SPHttpClient } from "@microsoft/sp-http";

export interface IHelloWorldProps {
  productName: string;
  productDescription: string;
  productQuantity: number;
  isCertified: boolean;
  title: string;
  rating: number;
  processorType: string;
  websiteUrl: string;
  spHttpClient: SPHttpClient;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
