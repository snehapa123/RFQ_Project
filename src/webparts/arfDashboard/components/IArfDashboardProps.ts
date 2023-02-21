import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";  
export interface IArfDashboardProps {
  description: string;
  spHttpClient: SPHttpClient;
  // context: any;
  siteUrl:string;
  context: WebPartContext; 
}

