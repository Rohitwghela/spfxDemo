import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, IWebPartContext, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './SpfxPnpCrud.module.scss';
import * as strings from 'spfxPnpCrudStrings';
import { ISpfxPnpCrudWebPartProps } from './ISpfxPnpCrudWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import pnp from 'sp-pnp-js';

interface IListItem {
    Title ? : string;
    Id: number;
}

export default class SpfxPnpCrudWebPart extends BaseClientSideWebPart < ISpfxPnpCrudWebPartProps > {

    constructor(context: IWebPartContext) {
        super();
        pnp.setup({
            headers: {
                'Accept': 'application/json;odata=nometadata'
            }
        });
        SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
        SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');
        //SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/css/toastr.css');
        //SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/2.2.2/jquery.js');
        //SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.js');
    }

    public render(): void {
        this.domElement.innerHTML = `
      <div class="">
        <div class="${styles.spfxPnpCrud}">
          <div class="${styles.container}">
            <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
              <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <p class="ms-font-l ms-fontColor-white">SharePoint CRUD operation using pnp and No framework</p>
                <p class="ms-font-l ms-fontColor-white">${escape(this.properties.description)}</p>
                <p class="ms-font-l ms-fontColor-white">${escape(this.context.pageContext.web.title)}</p>
                <p id="testID"></p>

                <div class="form-group">
                    <label for="idItemId">Enter Item ID</label>
                    <input type="text" class="form-control" id="idItemId" placeholder="Enter Item ID to be retrieve">
                </div>
                <div class="form-group">
                    <button class="btn btn-success" id="btnReadAllItems">
                        <span class="ms-Button-label">Read All Items</span>
                    </button>
                    <button class="btn btn-success" id="btnReadItemById">
                        <span class="ms-Button-label">Read Item By Id</span>
                    </button>
                    <div id="test" class="btn btn-success">TEST</div>
                    <button btn btn-success onClick={() => this.button_click() }>Click Here</button>
                </div>
                <div id="spListContainer"></div>

              </div> 
            </div>  
          </div> 
        </div>
      </div>
      `;

        //this.temp = ["a", "b", "c"];
        this._setButtonEventHandlers();

        
        //this.domElement.querySelector('#test').addEventListener('click', this.callMe);
    }

    //public temp;
    //public myTerm: ITerm;
    //public callMe(): void{
    //    debugger;
    //   alert(this.temp);
    //}
    private _setButtonEventHandlers(): void {
        debugger;
        const webPart: SpfxPnpCrudWebPart = this;
        this.domElement.querySelector('#btnReadAllItems').addEventListener('click', () => {
            this._GetListItemsNF();
        });
        this.domElement.querySelector('#btnReadItemById').addEventListener('click', () => {
            this._GetListItemById();
        });
        pnp.sp.site.getWebUrlFromPageUrl(window.location.href)
            .then(res => {
                alert(res);
        });

        //this.domElement.querySelector('#test').addEventListener('click', () => {
        //    this.callMe();
        //});

    }

    // Read all items from list PNP
    public _GetListItems(): void {
        pnp.sp.web.lists.getByTitle('TestList')
            .items.select('Title', 'Id').get()
            .then((items: IListItem[]): void => {
                console.log(items.length);
            }, (error: any): void => {
                console.log(error);
            });
    }

    // Real all items No Framework
    /*public _GetListItemsNF(): void{
        this.context;
        let listName = "";
        let url = "";
        this.context.spHttpClient.get("${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('testList')/items?$select=Title,Id", SPHttpClient.configurations.v1)
            .then((response: Response): Promise<{ value: IListItem[] }> => {
                return response.json();
            })
    }
    private _GetListItemsNF() : void{
        this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('testList')/items?$select=Title,Id", SPHttpClient.configurations.v1)
            .then((response: Response): Promise<{ value: IListItem[] }> => {
            return response.json();
        })
        .then((response: { value: IListItem[] }): void => {
            alert(response.value.length);
        }, (error: any): void => {
            alert(error);
        });
    }*/

    private _GetListItemsNF() : void{
        let url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('testList')/items?$select=Title,Id";
        this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((response:Response) : Promise<{ value: IListItem[] }> => {
                return response.json();
            })
            .then((response:{ value: IListItem[] }) =>{
                let html = '';
                response.value.forEach( (item:IListItem) =>  {
                    html += `
                        <ul class="${styles.list}">
                            <li class="${styles.listItem}">
                                <span>${item.Title}</span>
                            </li>
                        </ul>
                    `;
                });
                const responseView : Element = this.domElement.querySelector('#spListContainer');
                responseView.innerHTML = html;

            }, (error: any) => {
                let html = '';
                html += `
                        <ul class="${styles.list}">
                            <span>${error}</span>
                        </ul>
                    `;
                const responseView : Element = this.domElement.querySelector('#spListContainer');
                responseView.innerHTML = html;
            });
    }
    
    private _GetListItemById(): void{
        let itemId = this.domElement.querySelector('#idItemId');
        let url = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('testList')/items(" + itemId + ")";
    }
    

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [{
                header: {
                    description: strings.PropertyPaneDescription
                },
                groups: [{
                    groupName: strings.BasicGroupName,
                    groupFields: [
                        PropertyPaneTextField('description', {
                            label: strings.DescriptionFieldLabel
                        })
                    ]
                }]
            }]
        };
    }
}