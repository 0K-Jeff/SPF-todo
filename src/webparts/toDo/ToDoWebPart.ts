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
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

import styles from './ToDoWebPart.module.scss';
import * as strings from 'ToDoWebPartStrings';

// ----------------------------------- //

// define values for use later
let todolist: string = 'Default';

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

// ToDo List ID ---- EDIT HERE ON INITIAL CONFIGURATION --- //

  let TODOLISTID: string = '0x0100CDA8167196CB8248AACFC2DA5E5291DC';

//  --------------------------- //

// build web part
export default class ToDoWebPartWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {

  // Entry validation method, commpares with existing lists and informs if list already exists // THIS CAN CAUSE ERRORS IN CURRENT FRAMEWORK //
  // public ListExistCheck(ListTitle: String): Promise<string> {
  //   return this._fetchLists().then((response) => {
  //     let errorMsg: string = '';
  //     let listTitles: string[] = response.value.map(function(listobject: any){return listobject.Title});
  //       for (let itr = 0; itr < listTitles.length; itr++){
  //         if (listTitles[itr] == ListTitle){
  //         errorMsg = 'This List is now loaded.';
  //       }
  //     };
  //     return Promise.resolve(errorMsg);
  //   });
  // }

  // Get Method, fetching lists.
  public _fetchLists(): Promise<SPtodolist> {
    let firstquery: string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists`;
    return this.context.spHttpClient.get(firstquery, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {return response.json();
    });
  }

  // Post Method, making a new list, configuring
  public _postList(todolist: string): void {
    let restURL: string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists`;

    // Configuration of list settings for POST
      let listoptions: string = JSON.stringify({
        Title: todolist,
        BaseTemplate: 100,
        ContentTypesEnabled: true,
        //ContentType: 'ToDo List',
        //ContentTypeId: TODOLISTID
      });

    // Interface for new List Post Request
      let newList: ISPHttpClientOptions = {
        body: listoptions
      };

    // place request
    this.context.spHttpClient.post(restURL, SPHttpClient.configurations.v1, newList)
    // handle post response
    .then((response: SPHttpClientResponse) => {
      console.log(`Status code: ${response.status}`);
      console.log(`Status Text: ${response.statusText}`);

      //
      response.json().then((responseJSON: JSON) => {
        console.log(responseJSON);
      });
    });
  }

  // Render Method
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.toDo}">
        <p>Potato</p>
        <p>${this.properties.displaycount}</p>
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
              PropertyPaneTextField(todolist, {
                label: 'To-Do List Name'
                // onGetErrorMessage: this.ListExistCheck.bind(this),
                // errorMessage: ''
              }), //display item field
              PropertyPaneSlider('displaycount', {
                max: 15,
                min: 3,
                label: 'Number of Items to Display',
                showValue: true
              }), // Button method - Gets, then decides if it needs to Post
              PropertyPaneButton(todolist, {
                onClick: (todolist) => {
                  // Run Get Request on All Lists and return title array, then check for list already existing
                  this._fetchLists().then((response) => {
                    let createlistflag: boolean = true;
                    let Titles: string[] = (response.value.map(function(listobject: any){return listobject.Title}));
                    // run array and check if list Title already exists
                    for (let itr = 0; itr < Titles.length; itr++){
                      if(Titles[itr] == todolist) {
                        todolist = todolist;
                        createlistflag = false;
                        alert('List already exists.');
                        break;
                      }
                    }
                    if (createlistflag == true) {
                      // make Post request
                      this._postList(todolist);
                      alert('List Created.');
                    }
                  })
                  return todolist;
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
