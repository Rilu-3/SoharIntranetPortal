import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpOffersDetailsWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import OffersDetails, { IOfferView } from './OfferDetailsHtml';

export interface IWpOffersDetailsWebPartProps {
  description: string;
}

// interfaces for sharepoint list items
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  ID: number;
  Title: string;
  Description: string;
  OfferImage: any;
  Location: string;
  Contact: number;
  OfferPeriod: string;
  Created: any;
  Author: any;
  Status: any;
}

export default class WpOffersDetailsWebPart extends BaseClientSideWebPart<IWpOffersDetailsWebPartProps> {

  public async render(): Promise<void> {
    // Show loader until data is ready
    this.domElement.innerHTML = OffersDetails.wrapperHtml(OffersDetails.loadingHtml);

    // function call to extract query string parameters
    const queryStringParams: any = this.getQueryStringParameters();
    const id: number = parseInt(queryStringParams['OfferDetailID'], 10);

    if (isNaN(id)) {
      this._setContent(OffersDetails.noRecordHtml);
      return;
    }

    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Offers')/items(${id})?$select=Id,Title,Description,Created,Status,OfferImage,OfferPeriod,Location,Contact`;

    try {
      const item: ISPList = await this._getListData(apiUrl);

      // Show only active offers
      if (item.Status !== 'Active') {
        this._setContent(OffersDetails.noRecordHtml);
        return;
      }

      this._setContent(this._buildDetailsHtml(item));
    } catch (error) {
      console.error("Error loading offer", error);
      this._setContent(OffersDetails.noRecordHtml);
    }
  }

  // private method extracts query string parameters from the current URL and returns as an object.
  private getQueryStringParameters(): any {
    const queryStringParams: any = {};
    const queryString = window.location.search;

    if (queryString) {
      const queryParams = queryString.substring(1).split('&');
      // Iterate through each parameter and extract its key-value pair
      queryParams.forEach(param => {
        const [key, value] = param.split('=');
        queryStringParams[key] = value;
      });
    }
    return queryStringParams;
  }

  // Get data from the SharePoint list
  private async _getListData(apiUrl: string): Promise<ISPList> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
    }
    return response.json();
  }

  // Build offer details HTML from list item
  private _buildDetailsHtml(item: ISPList): string {
    const view: IOfferView = {
      title: escape(item.Title || ""),
      description: this._formatRichText(item.Description || ""),
      location: escape(item.Location || ""),
      offerPeriod: escape(item.OfferPeriod || ""),
      contact: item.Contact !== null && item.Contact !== undefined ? escape(String(item.Contact)) : "",
      imageUrl: this._getImageUrl(item, "Offers")
    };
    return OffersDetails.detailsHtml(view);
  }

  // Update the web part content area
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdOfferDetail');
    if (container !== null) {
      container.innerHTML = html;
    }
  }

  // Apply inline styles so enhanced rich text keeps its list formatting
  private _formatRichText(html: string): string {
    if (!html) {
      return "";
    }

    const doc: Document = new DOMParser().parseFromString(html, "text/html");
    const body: HTMLElement = doc.body;

    const applyStyle = (selector: string, styles: { [key: string]: string }): void => {
      const elements: NodeListOf<HTMLElement> = body.querySelectorAll(selector) as NodeListOf<HTMLElement>;
      for (let i = 0; i < elements.length; i++) {
        const el: HTMLElement = elements[i];
        for (const prop in styles) {
          if (styles.hasOwnProperty(prop) && !el.style.getPropertyValue(prop)) {
            el.style.setProperty(prop, styles[prop]);
          }
        }
      }
    };

    // Keep empty paragraphs (blank lines) visible
    const paragraphs: NodeListOf<HTMLElement> = body.querySelectorAll("p") as NodeListOf<HTMLElement>;
    for (let i = 0; i < paragraphs.length; i++) {
      const text: string = (paragraphs[i].textContent || "").replace(/\u200B/g, "").trim();
      if (text === "" && !paragraphs[i].querySelector("img")) {
        paragraphs[i].innerHTML = "&nbsp;";
      }
    }

    applyStyle("p", { "margin": "0 0 12px 0", "line-height": "1.5" });
    applyStyle("strong, b", { "font-weight": "bold" });
    applyStyle("em, i", { "font-style": "italic" });
    applyStyle("u", { "text-decoration": "underline" });
    applyStyle("ul", { "margin": "0 0 12px 0", "padding-left": "20px", "list-style-type": "disc" });
    applyStyle("ol", { "margin": "0 0 12px 0", "padding-left": "20px", "list-style-type": "decimal" });
    applyStyle("li", { "display": "list-item" });

    return body.innerHTML;
  }

  // Resolve image URL from SharePoint Image column JSON
  private _getImageUrl(item: ISPList, listUrlName: string): string {
    if (!item.OfferImage) {
      return "";
    }

    try {
      const img: any = JSON.parse(item.OfferImage);
      const webUrl: string = this.context.pageContext.web.absoluteUrl;

      // Image stored in Site Assets (serverRelativeUrl provided)
      if (img.serverRelativeUrl) {
        const origin: string = img.serverUrl || window.location.origin;
        return origin + encodeURI(img.serverRelativeUrl);
      }

      // Image stored as list item attachment (only fileName provided)
      if (img.fileName) {
        return `${webUrl}/Lists/${listUrlName}/Attachments/${item.ID}/${encodeURIComponent(img.fileName)}`;
      }
    } catch (e) {
      console.error("Unable to parse OfferImage", e);
    }

    return "";
  }

  public async onInit(): Promise<void> {
    await this._loadCSS();
    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }

  // private _loadHome(): void {
  //   const baseUrl: string = this.context.pageContext.web.absoluteUrl;

  //   SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/common.js`)
  //     .then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/home.js`));
  // }

  private _loadCSS(): void {
    const baseUrl: string = this.context.pageContext.web.absoluteUrl;
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/list.css`);
    // // Load all CSS files in parallel
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/variable.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/jquery-ui.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/font-size.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/custom.css`);
    
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/sp-custom.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/home.css`);

    // // Load scripts sequentially with proper dependency handling
    // SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`)
    //   .then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery-ui.js`))
    //   .then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`))
    //   //.then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`))
    //   .then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`))
    //   .then(() => this._loadHome());
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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