// All Imports
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

import styles from './ToDoWebPart.module.scss';
import * as strings from 'ToDoWebPartStrings';

// ----------------------------------- //

// define values for use later
export interface IToDoWebPartProps {
  todolist: string;
  displaycount: number;
  createlistbutton: string;
}

export interface SPtodolist {
  value: SPtodolist[];
}

export interface SPtodolistparts {
  Title: string;
  Id: string;
}

// build web part
export default class ToDoWebPartWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {

  // Get Method
  private _fetchListItems(todolist: string): Promise<SPtodolist> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${todolist}')?`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {return response.json();
    });
  }

  // Render Method
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.toDo}">
        <p>Potato</p>
        <p>${this.properties.displaycount}
      </div>`;

      // this._fetchListItems('Test').then((response) => {
      //   console.log(response);
      // })
  }
  // Version
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  //configure edit mode panel
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'A To-Do list web part.'
          },
          groups: [
            {     // list name field
              groupName: 'Enter List Name',
              groupFields: [
              PropertyPaneTextField('todolist', {
                label: 'To-Do List Name'
              }), //display item field
              PropertyPaneSlider('displaycount', {
                value: 5,
                max: 15,
                min: 3,
                label: 'Number of Items to Display',
                showValue: true

              }), // Button method
              PropertyPaneButton('todolist', {
                onClick: function(createlistbutton){
                  createlistbutton = createlistbutton+"1";
                  console.log(createlistbutton);
                  return createlistbutton;
                },
                buttonType: PropertyPaneButtonType.Primary,
                text: 'Create List'
              })
            ]}
          ]
        }
      ]
    };
  }
}
