import * as React from "react";
import styles from "./ToDoList.module.scss";
import { IToDoListProps } from "./IToDoListProps";
import { IItemListProps } from "./IItemListProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient } from "@microsoft/sp-http";
import { Icon, List } from "office-ui-fabric-react";
import * as strings from "ToDoListWebPartStrings";

export interface ISPListItem {
  Title: string;
  Id: string;
  Status: string;
}

export interface IListItemState {
  items: ISPListItem[];
  errorMessage: any;
}

class ItemList extends React.Component<IItemListProps, IListItemState> {
  constructor(props: IItemListProps, state: IListItemState) {
    super(props);

    this.state = {
      items: [],
      errorMessage: null
    };
  }

  private async _getListData(): Promise<ISPListItem[]> {
    try {
      const response = await this.props.spHttpClient.get(
        `${this.props.webUrl}/_api/web/lists/getByTitle('To do list')/items?$select=Title,Status`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const responseText = await response.text();
        throw new Error(responseText);
      }

      const data = await response.json();

      this.setState({ items: data.value });
      return data.value;
    } catch (error) {
      this.setState({ errorMessage: error.message });
    }
  }

  public componentDidMount(): void {
    this._getListData();
  }

  public render(): React.ReactElement<IItemListProps> {
    return (
      <>
        <List items={this.state.items} onRenderCell={this._onRenderListItem} />
        {this.state.errorMessage && <span>{this.state.errorMessage}</span>}
      </>
    );
  }

  public _itemStatus = (status: string): string => {
    switch (status) {
      case "Pending":
        return styles.itemPending;

      case "Completed":
        return styles.itemCompleted;

      case "Active":
        return styles.itemActive;

      case "Overdue":
        return styles.itemOverdue;

      default:
        return styles.itemStatus;
    }
  };

  public _onRenderListItem = (
    item: ISPListItem,
    index: number
  ): JSX.Element => {
    return (
      <div key={index} data-is-focusable={true}>
        <ul className={styles.list}>
          <li className={styles.listItem}>
            <span>{item.Title}</span>
            <span
              className={`${styles.itemStatus} ${this._itemStatus(
                item.Status
              )}`}
            >
              {item.Status}
              <Icon
                className={styles.itemIcon}
                iconName={`${item.Status === "Completed" ? "Completed" : null}`}
              />
            </span>
          </li>
        </ul>
      </div>
    );
  };
}

export default class ToDoList extends React.Component<IToDoListProps, {}> {
  public render(): React.ReactElement<IToDoListProps> {
    const { hasTeamsContext, userDisplayName } = this.props;

    return (
      <section
        className={`${styles.todoList} ${hasTeamsContext ? styles.teams : ""}`}
      >
        <h2>
          {strings.ToDoListHeading} {escape(userDisplayName)}
        </h2>
        <ItemList
          spHttpClient={this.props.spHttpClient}
          webUrl={this.props.websiteUrl}
        />
      </section>
    );
  }
}
