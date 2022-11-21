import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "HelloWorldWebPartStrings";
import HelloWorld from "./components/HelloWorld";
import { IHelloWorldProps } from "./components/IHelloWorldProps";

export interface IHelloWorldWebPartProps {
  productName: string;
  productDescription: string;
  productQuantity: number;
  isCertified: boolean;
  rating: number;
  processorType: string;
  title: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<IHelloWorldProps> = React.createElement(
      HelloWorld,
      {
        productName: this.properties.productName,
        productDescription: this.properties.productDescription,
        productQuantity: this.properties.productQuantity,
        isCertified: this.properties.isCertified,
        title: this.context.pageContext.web.title,
        rating: this.properties.rating,
        processorType: this.properties.processorType,
        websiteUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this.properties.processorType = "I7";
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  // private _renderList(items: ISPListItem[]): void {
  //   let html: string = "";
  //   items.forEach((item: ISPListItem) => {
  //     html += `
  //   <ul class="${styles.list}">
  //     <li class="${styles.listItem}">
  //       <span class="ms-font-l">${item.Title}</span>
  //     </li>
  //   </ul>`;
  //   });

  //   const listContainer: Element =
  //     this.domElement.querySelector("#spListContainer");
  //   listContainer.innerHTML = html;
  // }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then(context => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
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
            description: "Product Catalog"
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("productName", {
                  label: "Product Name"
                }),
                PropertyPaneTextField("productDescription", {
                  label: "Product Description",
                  multiline: true
                }),
                PropertyPaneTextField("productQuantity", {
                  label: "Product Quantity",
                  resizable: false,
                  deferredValidationTime: 3000,
                  placeholder: "Please enter product Quantity"
                }),
                PropertyPaneToggle("isCertified", {
                  key: "isCertified",
                  offText: "Not certified",
                  onText: "Certified!",
                  label: "Please certified first"
                }),
                PropertyPaneSlider("rating", {
                  label: "Rating",
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 5
                }),
                PropertyPaneChoiceGroup("processorType", {
                  label: "Processor",
                  options: [
                    { key: "I5", text: "Intel I5" },
                    { key: "I7", text: "Intel I7", checked: true },
                    { key: "I9", text: "Intel I9" }
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
