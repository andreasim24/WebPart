import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import * as strings from "ToDoListWebPartStrings";
import {
  ISPListItem,
  IItemListProps,
  IToDoListProps
} from "../../../interfaces";
import { getSP } from "../../../pnpjsConfig";
import { SPFI } from "@pnp/sp";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn
} from "@fluentui/react/lib/DetailsList";
import { Text } from "@fluentui/react/lib/Text";
import { IStackTokens, Stack, StackItem } from "@fluentui/react/lib/Stack";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { PrimaryButton } from "@fluentui/react/lib/Button";

const ItemListStackTokens: IStackTokens = {
  childrenGap: 15,
  padding: 10
};

const ItemList: React.FC<IItemListProps> = (props: IItemListProps) => {
  const [items, setItems] = React.useState<ISPListItem[]>([]);
  const [errorMessage, setErrorMessage] = React.useState(null);
  const columns: IColumn[] = [
    {
      key: "column1",
      name: "Title",
      fieldName: "Title",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: "column2",
      name: "Due Date",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      onRender: item => <Stack>{item.FieldValuesAsText.DueDate}</Stack>
    },
    {
      key: "column3",
      name: "Status",
      fieldName: "Status",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: ISPListItem) => {
        switch (item.Status) {
          case "Pending":
            return (
              <Stack
                horizontalAlign="center"
                verticalAlign="center"
                className={mergeStyles({
                  backgroundColor: "#ffe092",
                  color: "#997825",
                  borderRadius: "16px",
                  height: "24px",
                  whiteSpace: "nowrap",
                  padding: "4px 8px",
                  maxWidth: "max-content"
                })}
              >
                {item.Status}
              </Stack>
            );

          case "Completed":
            return (
              <Stack
                horizontalAlign="center"
                verticalAlign="center"
                className={mergeStyles({
                  backgroundColor: "#5dd4c0",
                  color: "#006b59",
                  borderRadius: "16px",
                  height: "24px",
                  whiteSpace: "nowrap",
                  padding: "4px 8px",
                  maxWidth: "max-content"
                })}
              >
                {item.Status}
              </Stack>
            );

          case "Active":
            return (
              <Stack
                horizontalAlign="center"
                verticalAlign="center"
                className={mergeStyles({
                  backgroundColor: "#8ac2ec",
                  color: "#235a85",
                  borderRadius: "16px",
                  height: "24px",
                  whiteSpace: "nowrap",
                  padding: "4px 8px",
                  maxWidth: "max-content"
                })}
              >
                {item.Status}
              </Stack>
            );

          case "Overdue":
            return (
              <Stack
                horizontalAlign="center"
                verticalAlign="center"
                className={mergeStyles({
                  backgroundColor: "#f6a89a",
                  color: "#903e2f",
                  borderRadius: "16px",
                  height: "24px",
                  whiteSpace: "nowrap",
                  padding: "4px 8px",
                  maxWidth: "max-content"
                })}
              >
                {item.Status}
              </Stack>
            );

          default:
            return <Stack>{item.Status}</Stack>;
        }
      }
    },
    {
      key: "column4",
      name: "Description",
      fieldName: "Description",
      isMultiline: true,
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    }
  ];

  const getListData = async (): Promise<void> => {
    const LIST_NAME = "To do list";
    const _sp: SPFI = getSP(props.context);

    try {
      const response = await _sp.web.lists
        .getByTitle(LIST_NAME)
        .items.select(
          "Title",
          "Status",
          "DueDate",
          "FieldValuesAsText/DueDate",
          "Description"
        )
        .expand("FieldValuesAsText")();
      setItems(response);
    } catch (error) {
      setErrorMessage(error.message);
    }
  };

  React.useEffect(() => {
    getListData();
  }, []);

  return (
    <>
      <DetailsList
        items={items}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionPreservedOnEmptyClick={true}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="select row"
      />
      {errorMessage && <span>{errorMessage}</span>}
    </>
  );
};

const ToDoList: React.FC<IToDoListProps> = (props: IToDoListProps) => {
  const { userDisplayName, spHttpClient, websiteUrl, context } = props;

  return (
    <Stack enableScopedSelectors tokens={ItemListStackTokens}>
      <StackItem>
        <Text variant={"xLarge"} nowrap block>
          {escape(userDisplayName)} {strings.ToDoListHeading}
        </Text>
      </StackItem>
      <StackItem>
        <ItemList
          spHttpClient={spHttpClient}
          webUrl={websiteUrl}
          context={context}
        />
      </StackItem>
      <StackItem>
        <PrimaryButton text="Primary" />
      </StackItem>
    </Stack>
  );
};

export default ToDoList;
