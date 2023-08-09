import { WebPartContext } from "@microsoft/sp-webpart-base"; 

export interface IFlashnewsProps {
  description: string;
  SiteUrl: string;
  ListName: string;
  FilterBy: string;
  condition: string;
  FilterValue:string;
  webUrl: string;
  Title:string;

  context: WebPartContext;  
  list: string;
  column: string;
  fields: string[];
  speed: number;
  direction: string;
  bgcolor:string;
  fgcolor:string;
  fontname:string;
  fontsize:string; 
  height:string;
  width:string;

  descbgcolor:string;
  descfgcolor:string;
  descfontsize:string;
  descfontname:string;
}
