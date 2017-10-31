// All Imports
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneLabel,
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
  listexists: string;
}

export interface SPtodolist {
  value: SPtodolist[];
}

export interface SPtodolistparts {
  Title: string;
  Id: string;
  Finished: boolean;
  Cancelled: boolean;
}

// Insert warning after list Name field.
  function insertAfter (newNode, targetElement) {
    targetElement.insertAdjacentElement('afterend', newNode);
  }

// build web part
export default class ToDoWebPartWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {

  // Get Method, fetching lists.
  public _fetchLists(todolist: string): Promise<SPtodolist> {
    let firstquery: string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists`;
    return this.context.spHttpClient.get(firstquery, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {return response.json();
    });
  }

  // Post Method, making a new list
  public _postList(todolist: string): Promise<SPtodolist> {
    let postlist: string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists`;
    let options: string = "";
    return this.context.spHttpClient.post(postlist, SPHttpClient.configurations.v1, options)
    // not done this bit yet
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
                max: 15,
                min: 3,
                label: 'Number of Items to Display',
                showValue: true
              }), // Button method - Gets, then decides if it needs to Post
              PropertyPaneButton('todolist', {
                onClick: (createlistbutton) => {
                  // Run Get Request on All Lists and return title array, then check for list already existing
                  this._fetchLists(createlistbutton).then((response) => {
                    let createlistflag: boolean = true;
                    let Titles: string[] = (response.value.map(function(listobject: any){return listobject.Title}));
                    for (let ite = 0; ite < Titles.length; ite++){
                      if(Titles[ite] == createlistbutton) {
                        if (document.getElementById('warningtextID') == null){
                        let warningtext = document.createElement('span');
                        warningtext.innerHTML = "<br> That List Already Exists.";
                        warningtext.id = 'warningtextID';
                        warningtext.style.color = '#ff0000';
                        insertAfter(warningtext, this);
                      }
                        createlistflag = false;
                        break;
                      }
                    }
                    if (createlistflag == true) {
                      // make Post request

                    }
                  })
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
