import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'UserDirectoryWebPartStrings';
import UserDirectory from './components/UserDirectory';
import { IUserDirectoryProps } from './components/IUserDirectoryProps';

export interface IUserDirectoryWebPartProps {
  title: string;
  searchFirstName: boolean;
  searchProps: string;
  pageSize: number;
  specficLoc:string;
  exclude:string;
}

export default class UserDirectoryWebPart extends BaseClientSideWebPart<IUserDirectoryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IUserDirectoryProps> = React.createElement(
      UserDirectory,
      {
        title: this.properties.title,
        context: this.context,
        searchFirstName: this.properties.searchFirstName,
        searchProps: this.properties.searchProps,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        pageSize: this.properties.pageSize,
        specficLoc:this.properties.specficLoc,
        exclude:this.properties.exclude
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField("title", {
                  label: "Web Part Title"
              }),
              PropertyPaneToggle("searchFirstName", {
                  checked: false,
                  label: "Search on First Name ?"
              }),
            PropertyPaneTextField('searchProps', {
              label: "Properties to search",
              description: "Enter the properties separated by comma to be used for search (FirstName,LastName,PreferredName,WorkEmail,Department)",
              value: this.properties.searchProps,
              multiline: false,
              resizable: false
            }),
            PropertyPaneTextField('specficLoc', {
              label: "Specify Location",
              description: "Enter the location separated by comma or enter '*' for all location",
              value: this.properties.specficLoc,
              multiline: false,
              resizable: false
            }),
            PropertyPaneTextField('exclude', {
              label: "Exclude Items",
              description: "Enter the exclude word separated by comma",
              value: this.properties.exclude,
              multiline: false,
              resizable: false
            }),
            PropertyPaneSlider('pageSize', {
            label: 'Results per page',
            showValue: true,
            max: 20,
            min: 2,
            step: 2,
            value: this.properties.pageSize
          })
              ]
            }
          ]
        }
      ]
    };
  }
}
