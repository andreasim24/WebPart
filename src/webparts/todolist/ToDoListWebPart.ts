import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "ToDoListWebPartStrings";
import ToDoList from "./components/ToDoList";
import { IToDoListProps } from "../../interfaces";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/lists";

export interface IToDoListWebPartProps {
  listName: string;
}

export default class ToDoListWebPart extends BaseClientSideWebPart<IToDoListWebPartProps> {
  private _sp: SPFI;

  public render(): void {
    const element: React.ReactElement<IToDoListProps> = React.createElement(
      ToDoList,
      {
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this._sp,
        listName: this.properties.listName
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "To Do List Description"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown("listName", {
                  label: strings.ListNameFieldLabel,
                  options: [
                    { key: "To do list", text: "To do list" },
                    { key: "Task List", text: "Task List" }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
