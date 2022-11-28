import { SPFI } from "@pnp/sp";

export interface ISPListItem {
  Title: string;
  Id: string;
  Status: string;
  DueDate: string;
  Description: string;
}

export interface IToDoListProps {
  userDisplayName: string;
  sp: SPFI;
  listName: string;
}

export interface IItemListProps {
  sp: SPFI;
  listName: string;
}
