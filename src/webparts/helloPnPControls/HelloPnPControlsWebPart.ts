import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

// Lsn 3.6.1 Add the People Picker control to the web part
import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

// Lsn 3.6.6 Add the Collection Data control to the web part
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData'

import styles from './HelloPnPControlsWebPart.module.scss';
import * as strings from 'HelloPnPControlsWebPartStrings';

export interface IHelloPnPControlsWebPartProps {
  description: string;
  //Lsn 3.6.2 Add the following property to the interface to store the people selected by the new control you're about to add to the property pane:
  people: IPropertyFieldGroupOrPerson[];

  // Lsn 3.6.7 Add the property to the interface to store the collection of data entered in the new control you're about to add to the property pane:
  expansionOptions: any[];
}

export default class HelloPnPControlsWebPart extends BaseClientSideWebPart<IHelloPnPControlsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }


  //Lsn 3.6.3 Display the selected people. Replace the lines you located with the following: <div class="selectedPeople"></div>

  // Lsn 3.6.8 add the following immediately after the <div class="selectedPeople"></div> element that you added previously: <div class="expansionOptions"></div>

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloPnPControls} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      
      <div class="selectedPeople"></div>
      <div class="expansionOptions"></div>

    </section>`;

    // Lsn 3.6.4 this will display their full names and email addresses in the <div> element you just added to the rendering:
    if (this.properties.people && this.properties.people.length > 0) {
      let peopleList: string = '';
      this.properties.people.forEach((person) => {
        peopleList = peopleList + `<li>${ person.fullName } (${ person.email })</li>`;
      });
    
      this.domElement.getElementsByClassName('selectedPeople')[0].innerHTML = `<ul>${ peopleList }</ul>`;
    }

    // Lsn 3.6.9 If any regions have been added, this will display their names and associated comments in the <div> element you just added to the rendering:
    if (this.properties.expansionOptions && this.properties.expansionOptions.length > 0) {
      let expansionOptions: string = '';
      this.properties.expansionOptions.forEach((option) => {
        expansionOptions = expansionOptions + `<li>${option['Region']}: ${option['Comment']}</li>`;
      });
      if (expansionOptions.length > 0) {
        this.domElement.getElementsByClassName('expansionOptions')[0].innerHTML = `<ul>${ expansionOptions }</ul>`;
      }
    }
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                
                // Lsn 3.6.5 Add the property pane field control to the property pane. Within the groupFields array, add the following people picker field control. This will bind the field control to the people property previously added to the web part's properties:

                PropertyFieldPeoplePicker('people', {
                  label: 'Property Pane Field People Picker PnP Reusable Control',
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),

                // Lsn 3.6.10 This will bind the field control to the expansionOptions property previously added to the web part's properties:
                PropertyFieldCollectionData('expansionOptions', {
                  key: 'collectionData',
                  label: 'Possible expansion options',
                  panelHeader: 'Possible expansion options',
                  manageBtnLabel: 'Manage expansion options',
                  value: this.properties.expansionOptions,
                  fields: [
                    {
                      id: 'Region',
                      title: 'Region',
                      required: true,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        { key: 'Northeast', text: 'Northeast' },
                        { key: 'Northwest', text: 'Northwest' },
                        { key: 'Southeast', text: 'Southeast' },
                        { key: 'Southwest', text: 'Southwest' },
                        { key: 'North', text: 'North' },
                        { key: 'South', text: 'South' }
                      ]
                    },
                    {
                      id: 'Comment',
                      title: 'Comment',
                      type: CustomCollectionFieldType.string
                    }
                  ]
                }),

              ]
            }
          ]
        }
      ]
    };
  }
}
