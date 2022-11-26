import * as React from "react";
import { escape } from "@microsoft/sp-lodash-subset";
import * as strings from "ToDoListWebPartStrings";
import styles from "./ToDoList.module.scss";
import {
  ISPListItem,
  IItemListProps,
  IToDoListProps
} from "../../../interfaces";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn
} from "@fluentui/react/lib/DetailsList";
import { Text } from "@fluentui/react/lib/Text";
import { IStackTokens, Stack, StackItem } from "@fluentui/react/lib/Stack";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import "@pnp/sp/lists";

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
              <span className={`${styles.ItemList} ${styles.statusPending}`}>
                {item.Status}
              </span>
            );

          case "Completed":
            return (
              <span className={`${styles.ItemList} ${styles.statusCompleted}`}>
                {item.Status}
              </span>
            );

          case "Active":
            return (
              <span className={`${styles.ItemList} ${styles.statusActive}`}>
                {item.Status}
              </span>
            );

          case "Overdue":
            return (
              <span className={`${styles.ItemList} ${styles.statusOverdue}`}>
                {item.Status}
              </span>
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
    try {
      const response = await props.sp.web.lists
        .getByTitle(props.listName)
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
  }, [props.listName]);

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
  const { userDisplayName, sp, listName } = props;

  return (
    <Stack enableScopedSelectors tokens={ItemListStackTokens}>
      <StackItem>
        <Text variant={"xLarge"} nowrap block>
          {escape(userDisplayName)} {strings.ToDoListHeading}
        </Text>
      </StackItem>
      <StackItem>
        <ItemList sp={sp} listName={listName} />
      </StackItem>
      <StackItem>
        <PrimaryButton text="Primary" />
      </StackItem>
    </Stack>
  );
};

export default ToDoList;
