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
require('./ToDoStyles.css');
// ------------------------------------------------------------------------- //

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

// ToDo List ID ---- EDIT HERE ON INITIAL CONFIGURATION -------------------- //

  let TODOLISTID: string = '0x0100CDA8167196CB8248AACFC2DA5E5291DC';

//  ------------------------------------------------------------------------ //

// build web part
export default class ToDoWebPartWebPart extends BaseClientSideWebPart<IToDoWebPartProps> {

  // Get Method, fetching lists.
  public _fetchLists(): Promise<SPtodolist> {
    let firstquery: string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists`;
    return this.context.spHttpClient.get(firstquery, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) =>
    {return response.json();
    });
  }

  // Get Method, fetching Items
  public _fetchListItems(title:string): Promise<SPHttpClientResponse> {
    let query:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${title}')/items`
    return this.context.spHttpClient.get(query, SPHttpClient.configurations.v1)
  }

  // internal methods for List Post Requests (configuration) --------------- //
  public _addContentType(title:string): Promise<SPHttpClientResponse> {
    let url:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${title}')/ContentTypes/AddAvailableContentType`;
    //Config new Item ID and assign to object
    let postbody: string = JSON.stringify({
      'contentTypeId': TODOLISTID
    })
    let postbodyblip: ISPHttpClientOptions = {
      body: postbody
    }
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, postbodyblip)
  }

  public _getItemId(title:string): Promise<SPHttpClientResponse> {
    let url:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${title}')/ContentTypes?$select=Name,id&$filter=Name+eq'Item'`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
  }

  public _deleteItemType(title:string, id:string): void {
    let url:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${title}')/ContentTypes('${id}')/deleteObject()`;
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, null)
  }
  // ----------------------------------------------------------------------- //

  // Post Method, making a new list, configuring
  public _postList(todolist: string): void {
    let restURL: string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists`;
    // Configuration of list settings for POST
      let listoptions: string = JSON.stringify({
        Title: todolist,
        BaseTemplate: 100,
        ContentTypesEnabled: true,
      });
    // Interface for new List Post Request
      let newList: ISPHttpClientOptions = {
        body: listoptions
      };
    // place request
    this.context.spHttpClient.post(restURL, SPHttpClient.configurations.v1, newList)
    // handle post response
    .then((response: SPHttpClientResponse) => {
      this._addContentType(todolist)
      .then((response: SPHttpClientResponse) => {
        this._getItemId(todolist)
        .then((response: SPHttpClientResponse) => {
          response.json().then((responseJSON: any) => {
            this._deleteItemType(todolist, responseJSON.value[0].Id.StringValue)
            console.log(responseJSON);
          })
        })
      })
    });
  }

  // LIST ITEM POST METHODS ------------------------------------------------ //
  // Post Method, Creating a new List Item
  public _postListItem(listTitle:string, ItemName:string): void{
    let capstitle:string = listTitle.toUpperCase();
    let ETFN:string = `SP.data.${capstitle}ListItem`
    let targetURL:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${listTitle}')/items`
    let newitemops:ISPHttpClientOptions = {body: {'type': `${ETFN}`, 'Title': `${ItemName}`}};
    this.context.spHttpClient.post(targetURL, SPHttpClient.configurations.v1, newitemops)
    .then((response: SPHttpClientResponse) => {
      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
      })
    })
  }
  // Post Method, Updating list items.
  public _updateListItem(listTitle:string, ItemName:string, itemId:string): void{
    let capstitle:string = listTitle.toUpperCase();
    let ETFN:string = `SP.data.${capstitle}ListItem`
    let targetURL:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`
    let newitemops:ISPHttpClientOptions = {headers: {'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*'}, body: {'type': `${ETFN}`, 'Title': `${ItemName}`}};
    this.context.spHttpClient.post(targetURL, SPHttpClient.configurations.v1, newitemops)
    .then((response: SPHttpClientResponse) => {
      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
      })
    })
  }
  // Post Method, Deleting list items
  public _deleteListItem(listTitle:string, itemId:string): void{
    let targetURL:string = this.context.pageContext.web.absoluteUrl+`/_api/web/lists/getbytitle('${listTitle}')/items/items(${itemId})`
    let newitemops:ISPHttpClientOptions = {headers: {'X-HTTP-Method':'DELETE', 'IF-MATCH': '*'}};
    this.context.spHttpClient.post(targetURL, SPHttpClient.configurations.v1, newitemops)
    .then((response: SPHttpClientResponse) => {
      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
      })
    })
  }
  // ----------------------------------------------------------------------- //

  // Front End Item Manipulation methods ------------------------------------//

  // Create a new list item -- attach to button


  // Delete list item -- attach to button


  // Modify a list item -- attach to something


  // ------------------------------------------------------------------------//

  // Render Method
  public render(): void {
    this.domElement.innerHTML = `Please Create a List in Edit Mode.`;
    // ------------------ //
    this._fetchListItems(this.properties.todolist).then((response: any) => {
      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
        let listbodyhtml:string = '';
        for (let iter = 0; iter < responseJSON.value.length; iter++){
          let listHtml:string = // TODO Add data tracking for Complete and IsActive variables and store in DOM for use.
          `<div class='ToDoItem'>
            <div class='MarkDone ${responseJSON.value[iter].Complete ? 'ItemDone' : ''}'> </div>
            <p class="ItemTitle">${responseJSON.value[iter].Title}</p>
            <div class='MarkInactive datahidden ${responseJSON.value[iter].IsActive1 ? 'IsNotInactive' : ''}'> </div>
            <span class='datahidden'>${responseJSON.value[iter].ID}</span>
          </div>`;
          listbodyhtml = listbodyhtml + listHtml;
        }
        let newItemButton:string = `<div class='newItemButton'>BUTTON</div>`;
        listbodyhtml = listbodyhtml + newItemButton;
        this.domElement.innerHTML = listbodyhtml;
      })
    })
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
                label: 'To-Do List Name',
                deferredValidationTime: 2000
              }), //display item field
              PropertyPaneSlider('displaycount', {
                max: 15,
                min: 3,
                label: 'Number of Items to Display',
                showValue: true
              }), // Button method - Gets, then decides if it needs to Post
              PropertyPaneButton('todolist', {
                onClick: (todolist) => {
                  // Run Get Request on All Lists and return title array, then check for list already existing
                  this._fetchLists().then((response) => {
                    console.log(response);
                    let createlistflag: boolean = true;
                    let Titles: string[] = (response.value.map(function(listobject: any){return listobject.Title}));
                    // run array and check if list Title already exists
                    for (let itr = 0; itr < Titles.length; itr++){
                      let todolistUpper:string = todolist.toUpperCase();
                      let titleUpper:string = Titles[itr].toUpperCase();
                      if(titleUpper == todolistUpper) {
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
