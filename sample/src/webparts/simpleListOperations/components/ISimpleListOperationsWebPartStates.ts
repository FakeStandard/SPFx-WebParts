export interface ISimpleListOperationsStates {
  addText: string;
  updateText: IListItem[];
}

export interface IListItem {
  id: number;
  title: string;
}
