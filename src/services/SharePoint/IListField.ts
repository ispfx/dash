export interface IListField {
  Id: string;
  Title: string;
  InternalName: string;
  TypeAsString: string;
  [index: string]: any;
}

export interface IListFieldCollection {
  value: IListField[];
}
