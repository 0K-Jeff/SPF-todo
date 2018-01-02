// All Imports
import { Version } from '@microsoft/sp-core-library';
import * as React from '/Users/jeffrey/Sites/spf-todo/node_modules/react';
import * as ReactDOM from '/Users/jeffrey/Sites/spf-todo/node_modules/react-dom'
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
  filterSelected: string;
  currentpage: number;
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

  // Constructor based Pagination refresh
  protected onInit(): Promise<void> {
  this.properties.currentpage = 0;
  console.log(this.properties.currentpage);
  return super.onInit();
}

  // Get Method, fetching lists.
  public _fetchLists(): Promise<SPtodolist> {
    let firstquery: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists`;
    return this.context.spHttpClient.get(firstquery, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  // Get Method, fetching Items
  public _fetchListItems(title: string): Promise<SPHttpClientResponse> {
    let query: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${title}')/items`
    return this.context.spHttpClient.get(query, SPHttpClient.configurations.v1)
  }

  // internal methods for List Post Requests (configuration) --------------- //
  public _addContentType(title: string): Promise<SPHttpClientResponse> {
    let url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${title}')/ContentTypes/AddAvailableContentType`;
    //Config new Item ID and assign to object
    let postbody: string = JSON.stringify({
      'contentTypeId': TODOLISTID
    })
    let postbodyblip: ISPHttpClientOptions = {
      body: postbody
    }
    return this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, postbodyblip)
  }

  public _getItemId(title: string): Promise<SPHttpClientResponse> {
    let url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${title}')/ContentTypes?$select=Name,id&$filter=Name+eq'Item'`;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
  }

  public _deleteItemType(title: string, id: string): void {
    let url: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${title}')/ContentTypes('${id}')/deleteObject()`;
    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, null)
  }
  // ----------------------------------------------------------------------- //

  // Post Method, making a new list, configuring
  public _postList(todolist: string): void {
    let restURL: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists`;
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
                  this._deleteItemType(todolist, responseJSON.value[0].Id.StringValue);
                  console.log(responseJSON);
                })
              })
          })
      });
  }

  // LIST ITEM POST METHODS ------------------------------------------------ //

  // Post Request Function accepts (Target URL and Config object)
  public postItemFunction(targetURL: string, itemOptions: any) {
    this.context.spHttpClient.post(targetURL, SPHttpClient.configurations.v1, itemOptions).then(this.render);
  }

  // Post Method, creating a new item
  public _postListItem(me: any): void {
    let targetURL: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${me.target.dataset.contentlistname}')/items`;
    let newitemops: ISPHttpClientOptions = { body:JSON.stringify({ 'Title': `New ToDo Item` }) };
    this.postItemFunction(targetURL, newitemops);
  }

  // Post Method, Updating list item complete status.
  public _updateListItemComplete(me: any): void {
    let targetURL: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${me.target.dataset.contentlistname}')/items(${me.target.dataset.itemid})`;
    if(me.target.dataset.itemcomplete == "false"){
      me.target.dataset.itemcomplete = true;
    } else {
      me.target.dataset.itemcomplete = false;
    }
    let newitemops: ISPHttpClientOptions = { headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body:JSON.stringify({ 'Complete': `${me.target.dataset.itemcomplete}` }) };
    this.postItemFunction(targetURL, newitemops);
  }

  // Post Method, Updating list item Text.
  public _updateListItemTitle(me: any): void {
    let targetURL: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${me.target.dataset.contentlistname}')/items(${me.target.dataset.itemid})`;
    let newitemops: ISPHttpClientOptions = { headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' }, body:JSON.stringify({ 'Title': `${me.target.value}` }) };
    this.postItemFunction(targetURL, newitemops);
  }

  // Post Method, Deleting list items
  public _deleteListItem(me: any): void {
    console.log("potatosdesd");
    let targetURL: string = this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${me.target.dataset.contentlistname}')/items(${me.target.dataset.itemid})`;
    let newitemops: ISPHttpClientOptions = { headers: { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' } };
    this.postItemFunction(targetURL, newitemops);
  }

  // ----------------------------------------------------------------------- //

  // Utility Functions

  public _filterSwap(me: any): void {
    this.properties.filterSelected = me.target.textContent;
    console.log(this.properties.filterSelected);
    this.properties.currentpage = 0;
    this.render();
  }

  public _pageForward(): void {
    this.properties.currentpage++;
    this.render();
  }

  public _pageBack(): void {
    if (this.properties.currentpage > 0) {
    this.properties.currentpage--;
    this.render();
    }
  }

  public _pageSelect(me: any):void {
    this.properties.currentpage = (parseInt(me.target.innerHTML)-1);
    this.render();
  }

  // ------------------------------------------------------------------------//

  // Render Method
  public render(): void {
    this._fetchListItems(this.properties.todolist).then((response: any) => {
      response.json().then((responseJSON: any) => {
        console.log(responseJSON);
        if (responseJSON.value) {
          // Clear for repopulation
          this.domElement.innerHTML = '';
          let uniqueid: string = this.properties.todolist;
          let pageLeft: boolean = false;
          let pageRight: boolean = false;
          let displayedTotal:number = 0;
          let skippedTotal:number = 0;
          let filteredTotal:number = 0;

          // Add filters
          var filterSet = document.createElement("div");

          var filterItem = document.createElement("div");
          filterItem.className = 'filter';
          filterItem.innerHTML = 'All';
          if (this.properties.filterSelected === 'All') {filterItem.classList.add('active')};
          filterItem.addEventListener('click', this._filterSwap.bind(this))
          filterSet.appendChild(filterItem);

          var filterItem2 = document.createElement('div');
          filterItem2.className = 'filter';
          filterItem2.innerHTML = `Incomplete`
          if (this.properties.filterSelected === 'Incomplete') {filterItem2.classList.add('active')};
          filterItem2.addEventListener('click', this._filterSwap.bind(this))
          filterSet.appendChild(filterItem2);

          var filterItem3 = document.createElement("div");
          filterItem3.className = 'filter';
          filterItem3.innerHTML = `Complete`
          if (this.properties.filterSelected === 'Complete') {filterItem3.classList.add('active')};
          filterItem3.addEventListener('click', this._filterSwap.bind(this))
          filterSet.appendChild(filterItem3);

          this.domElement.appendChild(filterSet);

          // ---- //


        // Run through list items and render based on filter and pagination
          for (let iter = 0; iter <responseJSON.value.length; iter++){
            // filter based on completion
            let shouldRender: boolean = true;
            let completedItem: boolean = responseJSON.value[iter].Complete;
            let filterValue: string = this.properties.filterSelected;
            if (filterValue === 'All' || undefined) {
              shouldRender = true;
            } else if (filterValue === 'Incomplete' && completedItem === true) {
              shouldRender = false;
            } else if (filterValue === 'Complete' && completedItem === false) {
              shouldRender = false;
            }

            // filter based on pagination
            if (shouldRender === true) {
              filteredTotal++;
              if (skippedTotal < (this.properties.currentpage*this.properties.displaycount)) {
                skippedTotal++;
                pageLeft = true;
                shouldRender = false;
              } else if (displayedTotal < this.properties.displaycount) {
                displayedTotal++;
              } else if (displayedTotal >= this.properties.displaycount) {
                pageRight = true;
                shouldRender = false;
              }
            }


            if (shouldRender === true) {
            // create an empty row
            let ToDoItemRow: any = document.createElement("div");
            ToDoItemRow.className = "TodoRow";

            // Create Item Complete Button
            var ItemCompleteButton = document.createElement("div");
            ItemCompleteButton.innerHTML = `<div class='MarkDone ${responseJSON.value[iter].Complete ? 'ItemDone' : ''}' data-itemid=${responseJSON.value[iter].ID} data-contentlistname='${this.properties.todolist}' data-itemcomplete='${responseJSON.value[iter].Complete}'> </div>`;
            ItemCompleteButton.addEventListener('click', this._updateListItemComplete.bind(this));
            ToDoItemRow.appendChild(ItemCompleteButton);

            // Create Item form
            var ListItemText = document.createElement("form");
            ListItemText.innerHTML = `<input type='text' class="ItemTitle" value='${responseJSON.value[iter].Title}' data-itemid='${responseJSON.value[iter].ID}' data-contentlistname='${this.properties.todolist}'></input>`;
            ListItemText.addEventListener('change', this._updateListItemTitle.bind(this));
            ToDoItemRow.appendChild(ListItemText);

            // Create Delete button
            var ItemDeleteButton = document.createElement("div");
            ItemDeleteButton.innerHTML = `<div class='CancelItem' data-contentlistname='${this.properties.todolist}' data-itemid='${responseJSON.value[iter].ID}'>X</div>`;
            ItemDeleteButton.addEventListener('click', this._deleteListItem.bind(this));
            ToDoItemRow.appendChild(ItemDeleteButton);

            // Append Row to doc
            this.domElement.appendChild(ToDoItemRow);
            }
          }

        // Create 'new todo button'
        let newItemButton = document.createElement("div");
        newItemButton.innerHTML = `<div class='newItemButton' data-contentlistname='${this.properties.todolist}'>New Item</div>`;
        newItemButton.addEventListener('click', this._postListItem.bind(this));
        this.domElement.appendChild(newItemButton);

        // add pagination buttons
        if (pageRight == true){
          let rightButton = document.createElement('div');
          rightButton.className = 'pageRight';
          rightButton.innerHTML = '>';
          rightButton.addEventListener('click', this._pageForward.bind(this));
          this.domElement.appendChild(rightButton);
        }
        if (filteredTotal > this.properties.displaycount){
          let pageTotal:number = Math.ceil((filteredTotal/this.properties.displaycount));
          for (let i = pageTotal; i > 0; i--) {
            let pageNumber = document.createElement('div');
            pageNumber.className = 'pageNumber';
            if (this.properties.currentpage+1 == i) {
              pageNumber.classList.add('currentPageNumber')
            }
            pageNumber.innerHTML = `${i}`;
            pageNumber.addEventListener('click', this._pageSelect.bind(this));
            this.domElement.appendChild(pageNumber);
          }
        }
        if (pageLeft == true){
          let leftButton = document.createElement('div');
          leftButton.className = 'pageLeft';
          leftButton.innerHTML = '<';
          leftButton.addEventListener('click', this._pageBack.bind(this));
          this.domElement.appendChild(leftButton);
        }
        // Enter a value if no list exists
      } else {
      this.domElement.innerHTML = `<div>Please Enter a List for your ToDo.</div>`;
      }

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
                      let Titles: string[] = (response.value.map(function(listobject: any) { return listobject.Title }));
                      // run array and check if list Title already exists
                      for (let itr = 0; itr < Titles.length; itr++) {
                        let todolistUpper: string = todolist.toUpperCase();
                        let titleUpper: string = Titles[itr].toUpperCase();
                        if (titleUpper == todolistUpper) {
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
              ]
            }
          ]
        }
      ]
    };
  }
}
