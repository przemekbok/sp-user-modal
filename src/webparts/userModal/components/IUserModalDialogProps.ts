import { IUserItem } from '../UserModalWebPart';

export interface IUserModalDialogProps {
  isOpen: boolean;
  onDismiss: () => void;
  userData: IUserItem | null;
  isDarkTheme: boolean;
}