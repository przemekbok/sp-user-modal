import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserItem } from '../UserModalWebPart';

export interface IUserModalProps {
  webPartTitle: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userItems: IUserItem[];
  isLoading: boolean;
  itemsPerPage: number;
  context: WebPartContext;
}