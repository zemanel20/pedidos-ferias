import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface IPedidosProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
