import { HttpClient } from "@microsoft/sp-http";

export interface IMyMailsProps {
  applicationID: string;
  redirectUri: string;
  tenantUrl: string;
  userMail: string;
  httpClient: HttpClient;
}
