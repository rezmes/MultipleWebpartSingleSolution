import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,

  // P025 Working with PropertyPane Toggle
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './P016SitePropertiesWebPart.module.scss';
import * as strings from 'P016SitePropertiesWebPartStrings';

// P019 SharePoint Lists name
import {
  SPHttpClient,
  SPHttpClientConfiguration,
  SPHttpClientResponse
} from '@microsoft/sp-http'

export interface ISharePointList {
  Title: string;
  Id: string;
}
export interface ISharePointLists {
  value: ISharePointList[];
}
export interface IP016SitePropertiesWebPartProps {
  description: string;

  // ///////////////////////////////////////////////P021 Textboxes
  productname: string;
  productdescription: string;
  productcost: number;
  quantity: number;
  billamount: number;
  discount: number;
  netbillamount: number;


  // //////////////////////////////////////////P025 Working with PropertyPane Toggle
currentTime: Date;
IsCertified: boolean;

}
// /////////////////////////////////////////////////////






export default class P016SitePropertiesWebPart extends BaseClientSideWebPart<IP016SitePropertiesWebPartProps> {
// ///////////////////////////////////////P023 Working with onInit Function
protected onInit(): Promise<void> {
  return new Promise<void>((resolve, _reject) => {
    this.properties.productname= "Mouse";
    this.properties.productdescription= "Wireless A4Tech";
    this.properties.productcost= 300;
    this.properties.quantity= 3;
    resolve(undefined)
  })
}

// ///////////////////////////////////////P024 Disabling Reactive Change
protected get disableReactivePropertyChanges(): boolean {
  return true;
}


// //////////////////////////////////////P019 SharePoint Lists name
private _getListOfLists(): Promise<ISharePointLists> {
  return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl +`/_api/web/lists?$filter=Hidden eq false`,SPHttpClient.configurations.v1)
  .then((response: SPHttpClientResponse) => {
    return response.json();
  });
}
private _getAndRenderLists(): void {
  if(Environment.type === EnvironmentType.Local) {

  }
  else if (Environment.type == EnvironmentType.SharePoint ||
    Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListOfLists()
      .then((response) => {
        this._renderListOfLists(response.value);
      });
    }
}
private _renderListOfLists(items: ISharePointList[]): void {
  let html: string = '';
  items.forEach((item: ISharePointList) => {
    html += `
    <ul class="${styles.list}">
                <li class="${styles.listItem}">
                <span class="ms-font-l">${item.Title}</span>
                </li>
                <li class="${styles.listItem}">
                <span class="ms-font-l">${item.Id}</span>
                </li>
            </ul>`;
  });
  const listsPlaceholder: Element = this.domElement.querySelector('#SPListPlaceHolder');
  listsPlaceholder.innerHTML =html;
}

// ///////////////////////////////////////////////P025 Working with PropertyPane Toggle
private _addCertPrice(): void {
  if (this.properties.IsCertified) {
    this.properties.netbillamount = this.properties.billamount * 1.2;
  } else {
    this.properties.discount = this.properties.billamount * 0.1;
    this.properties.netbillamount = this.properties.billamount - this.properties.discount;
  }
}



// ///////////////////////////////////////////////RENDER
public render(): void {
  // Calculate the net bill amount based on certification
  this._addCertPrice();

  this.domElement.innerHTML = `
    <div class="${styles.p016SiteProperties}">
      <div class="${styles.container}">
        <div class="${styles.row}">
          <div class="${styles.column}">
            <span class="${styles.title}">Welcome to SharePoint!</span>
            <p class="${styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
            <p class="${styles.description}">${escape(this.properties.description)}</p>
            <!-- P021 TextBoxes -->
            <table>
              <tr>
                <td>Product Name</td>
                <td>${this.properties.productname}</td>
              </tr>
              <tr>
                <td>Product Description</td>
                <td>${this.properties.productdescription}</td>
              </tr>
              <tr>
                <td>Product Cost</td>
                <td>${this.properties.productcost}</td>
              </tr>
              <tr>
                <td>Product Quantity</td>
                <td>${this.properties.quantity}</td>
              </tr>
              <tr>
                <td>Total Price</td>
                <td>${this.properties.billamount = this.properties.quantity * this.properties.productcost}</td>
              </tr>
              <tr>
                <td>Discount</td>
                <td>${this.properties.discount = this.properties.billamount * 0.1}</td>
              </tr>
              <tr>
                <td>Net Price</td>
                <td>${this.properties.netbillamount}</td>
              </tr>
            </table>
            <!-- P017 Site Info/Properties -->
            <p class="${styles.description}">${escape(this.context.pageContext.web.absoluteUrl)}</p>
            <p class="${styles.description}">${escape(this.context.pageContext.web.title)}</p>
            <p class="${styles.description}">${escape(this.context.pageContext.web.serverRelativeUrl)}</p>
            <p class="${styles.description}">${escape(this.context.pageContext.user.displayName)}</p>
            <!-- P018 Culture Info -->
            <ul>
              <li><strong>Current Culture Name</strong>: ${escape(this.context.pageContext.cultureInfo.currentCultureName)}</li>
              <li><strong>Current UI Culture Name</strong>: ${escape(this.context.pageContext.cultureInfo.currentUICultureName)}</li>
              <li><strong>Is Right To Left?</strong>: ${this.context.pageContext.cultureInfo.isRightToLeft}</li>
            </ul>
            <a href="https://aka.ms/spfx" class="${styles.button}">
              <span class="${styles.label}">Learn more</span>
            </a>
          </div>
        </div>
      </div>
      <div id="SPListPlaceHolder"></div>
    </div>`;
  this._getAndRenderLists();
}



  // protected get dataVersion(): Version {
  //   return Version.parse('1.0');
  // }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 label: strings.DescriptionFieldLabel
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }




// P21 Working with TextBoxes

protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages:[
      {
        groups: [
          {
            groupName: "Product Details",
              groupFields: [

                PropertyPaneTextField('productname', {
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter product name", "description": " Name of property field"
                }),

                PropertyPaneTextField('productdescription', {
                  label: "Product Description",
                  multiline: true,
                  resizable:false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter Product Description", "description": "Name property field"
                }),

                PropertyPaneTextField('productcost', {
                  label: "Product cost",
                  multiline: false,
                  resizable:false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter Product Cost", "description": "Number property field"
                }),

                PropertyPaneTextField('quantity', {
                  label: "Product Quantity",
                  multiline: false,
                  resizable:false,
                  deferredValidationTime: 5000,
                  placeholder: "Please enter Product Quantity", "description": "Number property field"
                }),


              ]



          },
          // ////////////////////////////////// P025Working with PropertyPane Toggle
          {
            groupName: "Certification",

              groupFields:
              [
                PropertyPaneToggle('IsCertified',
                  {
                    key:'IsCertified',
                    label: 'Is it Certified',
                    onText: 'ISI Certified',
                    offText: 'Not Certified'

                  }
                )
              ]
            }



        ]
      }




    ]

    }
  }
}


