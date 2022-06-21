import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IAllItems } from "../Models/IAllItems";
export interface IBbsCpaAllItemsProps {
  description: string;
  listitem: IAllItems[];
  Context: WebPartContext;
  SeeAllUrl: string;
  listId: string;
  WebUrl: string;
  view: string;}
