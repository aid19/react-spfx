import { SPHttpClient } from "@microsoft/sp-http";

export interface IHelloWorldProps {
  listName: string;
  spHttpClient: SPHttpClient;
  siteUrl:string;
}
