import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './StockPriceWebPart.module.scss';
import * as strings from 'StockPriceWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

import { getIconClassName } from '@uifabric/styling';

import { PnPClientStorage, dateAdd } from '@pnp/common';
 
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Price : string;
  Symbol :string;
  LastRefreshed :string;
  Volume :string;
  PriceDifference :string;
  IconName :string;
}

export interface IStockPriceWebPartProps {
  description: string;
}

export default class StockPriceWebPart extends BaseClientSideWebPart<IStockPriceWebPartProps> {
 private _getListData(): Promise<ISPLists> {
    try{
      return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('Stock Ticker')/Items?$orderby= Modified desc",SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
          return response.json();
          });
    } catch(e){
      console.log("Error calling loading Stock Price list: " + e.message);
    }
  } 

    private _renderListAsync(): void 
    {
      const pnpStorage = new PnPClientStorage();
      let cacheKeyStockPrice = "SPFX_Stock_Price"
      try
      {
          if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
            let cachedStockPrice: ISPList = pnpStorage.local.get(cacheKeyStockPrice);
            if (!cachedStockPrice) {
              this._getListData()
              .then((response) => {
                this._renderList(response.value[0]);
                // Cache the Stock Price
                pnpStorage.local.put(cacheKeyStockPrice, response.value[0], dateAdd(new Date(), 'minute', 15));  
              });  
              
            } else {
              this._renderList(cachedStockPrice);
            }
            
        }
      } catch(e){
          console.log("Error calling loading Stock Price list: " + e.message);
      }
    }

   //private _renderList(items: ISPList[]): void {
  private _renderList(item: ISPList): void {
    let html: string = '<h2>Unable to load Stock Price information</h2>';
    if(item){
      let iconDir = ``;
      if(item.IconName == "Up"){
        iconDir = `<i data-icon-name="Up" aria-hidden="true" class="ms-Icon root-43"></i>`;
      } else if(item.IconName == "Down"){
        iconDir = `<i data-icon-name="Down" aria-hidden="true" class="ms-Icon root-43"></i>`
      }

      html = 
      `<div class=${styles.stock}>
        <div class=${styles.stockHeaderRow}>
          <div class=${styles.stockHeaderColumn}>${item.Title}</div>
          <div class="${styles.stockHeaderColumn} ${styles.stockSymbol}">${item.Symbol}</div>
        </div>

        <div class=${styles.stockHeaderRow}>
          <div class="${styles.pill}">
            <span class=${styles.stockTrend}>
              ${iconDir}
            </span>
            <span class=${styles.stockValue}>${ item.Price } USD</span>
            <span>${item.PriceDifference}</span>
          </div>
        </div>

        <div class=${styles.stockHeaderRow}>
          <div class=${styles.stockDetails}>
            <div>Last Refreshed</div>
            <div>${item.LastRefreshed}</div>
          </div>
          <div class="${styles.stockDetails} ${styles.stockSymbol}"}>
            <div>Volume:</div>
            <div>${item.Volume}</div>
          </div>
        </div>
      </div>`;
    }
    //});

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `<div id="spListContainer" class="${styles.stockTicker}" />`;
      this._renderListAsync();
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
