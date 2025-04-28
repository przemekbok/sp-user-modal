import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneHorizontalRule
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/profiles";
import { IDropdownOption } from '@fluentui/react/lib/Dropdown';

import * as strings from 'UserModalWebPartStrings';
import UserModal from './components/UserModal';
import { IUserModalProps } from './components/IUserModalProps';

export interface IUserModalWebPartProps {
  title: string;
  listName: string;
  itemsPerPage: number;
  userFieldName: string;
  descriptionFieldName: string;
  certificationFieldName: string;
}

export interface IUserItem {
  id: number;
  title: string; // User's display name
  position: string; // User's job title
  photoUrl: string; // User's profile photo
  description: string; // User description
  certification: string; // User certifications
  email: string; // User email
}

export default class UserModalWebPart extends BaseClientSideWebPart<IUserModalWebPartProps> {
  private _sp: SPFI;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _userItems: IUserItem[] = [];
  private _isLoading: boolean = true;
  private _availableLists: IDropdownOption[] = []; // Store available lists

  public onInit(): Promise<void> {
    // Initialize PnP JS
    this._sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));
    
    return Promise.all([
      this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      }),
      this._fetchAvailableLists()
    ]).then(() => {
      return Promise.resolve();
    });
  }

  public render(): void {
    this._isLoading = true;
    this._fetchUsersFromList().then(() => {
      this._isLoading = false;
      this._renderWebPart();
    }).catch(error => {
      console.error("Error fetching users:", error);
      this._isLoading = false;
      this._renderWebPart();
    });
  }

  private _renderWebPart(): void {
    const element: React.ReactElement<IUserModalProps> = React.createElement(
      UserModal,
      {
        webPartTitle: this.properties.title || 'Subject Matter Experts',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userItems: this._userItems,
        isLoading: this._isLoading,
        itemsPerPage: this.properties.itemsPerPage || 4,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private async _fetchAvailableLists(): Promise<void> {
    try {
      // Get all non-hidden lists from the site
      const lists = await this._sp.web.lists
        .filter("Hidden eq false and BaseTemplate eq 100")
        .select("Title, Id")
        .orderBy("Title")
        (); // Changed from .get() to () for PnP JS v3+
      
      // Convert lists to dropdown options
      this._availableLists = lists.map(list => ({
        key: list.Title,
        text: list.Title
      }));
      
      // Add default empty option
      this._availableLists.unshift({
        key: '',
        text: '- Select a list -'
      });

    } catch (error) {
      console.error("Error fetching SharePoint lists:", error);
      this._availableLists = [{
        key: '',
        text: '- Error loading lists -'
      }];
    }
  }

  private async _fetchUsersFromList(): Promise<void> {
    if (!this.properties.listName) {
      this._userItems = []; // No list selected, empty the array
      return;
    }

    try {
      // Set field names - with defaults if not provided
      const userFieldName = this.properties.userFieldName || 'User';
      const descFieldName = this.properties.descriptionFieldName || 'Description';
      const certFieldName = this.properties.certificationFieldName || 'Certification';

      // Get items from the list
      const items = await this._sp.web.lists.getByTitle(this.properties.listName).items.select(
        'ID',
        `${userFieldName}/Title`,
        `${userFieldName}/EMail`,
        `${userFieldName}/JobTitle`,
        `${userFieldName}/Name`,
        `${userFieldName}/Id`,
        descFieldName,
        certFieldName
      ).expand(userFieldName)();

      // Process items
      const processedItems: IUserItem[] = await Promise.all(
        items.map(async (item: any) => {
          const userId = item[userFieldName]?.Id;
          let photoUrl = '';
          
          // Get user's profile photo if userId exists
          if (userId) {
            try {
              // Try to get the profile photo URL
              const userProperties = await this._sp.profiles.getPropertiesFor(`i:0#.f|membership|${item[userFieldName].EMail}`);
              const pictureUrl = userProperties.PictureUrl;
              photoUrl = pictureUrl || require('./assets/person.png');
            } catch (error) {
              console.warn(`Error getting profile photo for user ${item[userFieldName].Title}:`, error);
              photoUrl = require('./assets/person.png');
            }
          } else {
            photoUrl = require('./assets/person.png');
          }
          
          return {
            id: item.ID,
            title: item[userFieldName]?.Title || 'Unknown User',
            position: item[userFieldName]?.JobTitle || '',
            photoUrl: photoUrl,
            description: item[descFieldName] || '',
            certification: item[certFieldName] || '',
            email: item[userFieldName]?.EMail || ''
          };
        })
      );
      
      this._userItems = processedItems;
    } catch (error) {
      console.error("Error fetching items from SharePoint list:", error);
      this._userItems = [];
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Get the list creation URL for the instructional text
    const listCreationUrl = `${this.context.pageContext.web.absoluteUrl}/_layouts/15/createlist.aspx`;
    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Web Part Title',
                  value: 'Subject Matter Experts'
                }),
                PropertyPaneDropdown('listName', {
                  label: 'SharePoint List',
                  options: this._availableLists,
                  selectedKey: this.properties.listName
                }),
                PropertyPaneLabel('createListNote', {
                  text: `Can't find your list? Create a new list in SharePoint at: ${this.context.pageContext.web.absoluteUrl}/_layouts/15/createlist.aspx`
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneDropdown('itemsPerPage', {
                  label: 'Tiles Per View',
                  options: [
                    { key: 1, text: '1' },
                    { key: 2, text: '2' },
                    { key: 3, text: '3' },
                    { key: 4, text: '4' }
                  ],
                  selectedKey: 4
                }),
                PropertyPaneTextField('userFieldName', {
                  label: 'User Field Name',
                  description: 'Default: User'
                }),
                PropertyPaneTextField('descriptionFieldName', {
                  label: 'Description Field Name',
                  description: 'Default: Description'
                }),
                PropertyPaneTextField('certificationFieldName', {
                  label: 'Certification Field Name',
                  description: 'Default: Certification'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}