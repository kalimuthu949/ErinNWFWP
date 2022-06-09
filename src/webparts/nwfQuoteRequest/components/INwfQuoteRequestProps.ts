import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INwfQuoteRequestProps {
  description: string;
  context: WebPartContext;
  spcontext: any;
}
