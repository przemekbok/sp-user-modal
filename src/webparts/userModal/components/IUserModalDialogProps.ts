import { IUserItem } from '../UserModalWebPart';

export interface IUserModalDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  userData: IUserItem | undefined;
  isDarkTheme: boolean;
}