import { IUserItem } from '../UserModalWebPart';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserTileProps {
  item: IUserItem;
  onOpenModal: (item: IUserItem) => void;
  context: WebPartContext;
}