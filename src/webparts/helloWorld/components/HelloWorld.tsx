import * as React from "react";
import styles from "./HelloWorld.module.scss";
import { IHelloWorldProps } from "./IHelloWorldProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient } from "@microsoft/sp-http";
import { Icon, Label, List } from "office-ui-fabric-react";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";

export interface ISPListItem {
  Title: string;
  Id: string;
  Status: string;
}

export interface IHelloWorldState {
  items: ISPListItem[];
}

export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  IHelloWorldState
> {
  constructor(props: IHelloWorldProps, state: IHelloWorldState) {
    super(props);

    this.state = {
      items: []
    };
  }

  private async _getListData(): Promise<ISPListItem[]> {
    try {
      const response = await this.props.spHttpClient.get(
        "https://xq0nb.sharepoint.com/sites/TestSite/_api/web/lists/getByTitle('To do list')/items",
        SPHttpClient.configurations.v1
      );
      const data = await response.json();
      this.setState({ items: data.value });
      return data.value;
    } catch (error) {
      console.log(error);
    }
  }

  public componentDidMount(): void {
    this._getListData();
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      productName,
      productDescription,
      productQuantity,
      isCertified,
      rating,
      title,
      processorType,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section
        className={`${styles.helloWorld} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <Label>This is product catalog for, {escape(userDisplayName)}!</Label>
          <div>
            Product Name: <strong>{escape(productName)}</strong>
          </div>
          <div>Product Description: {escape(productDescription)}</div>
          <div>Product Quantity : {productQuantity}</div>
          <div>
            Certified <input type={"checkbox"} checked={isCertified} />
          </div>
          <div>Rating : {rating}</div>
          <div>Processor Type: {processorType}</div>
        </div>
        <h2>This is my Sharepoint List Items :</h2>
        {/* Try List Component from fabric-ui */}
        <List items={this.state.items} onRenderCell={this._onRenderList} />
        <div>{environmentMessage}</div>
        <div>
          Loading from: <strong>{escape(title)}</strong>
        </div>
        {/* Try Button Component from fabric-ui */}
        <PrimaryButton>Primary Button</PrimaryButton>
      </section>
    );
  }

  private _onRenderList(item: ISPListItem, index: number): JSX.Element {
    return (
      <div className="ms-ListBasicExample-itemCell" data-is-focusable={true}>
        <ul className={styles.list}>
          <li className={styles.listItem}>
            <span className="ms-font-l">{item.Title}</span>
            <span
              className={`${styles.itemStatus} ${
                item.Status === "Pending"
                  ? styles.itemPending
                  : item.Status === "Completed"
                  ? styles.itemCompleted
                  : styles.itemActive
              }`}
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
  }
}
