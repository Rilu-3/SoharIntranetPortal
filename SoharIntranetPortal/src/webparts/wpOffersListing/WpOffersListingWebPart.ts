import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpOffersListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpOffersListing, { IOfferCardView } from './offersListingHtml';

export interface IWpOffersListingWebPartProps {
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
  Status: any;
  Created: any;
}

export default class WpOffersListingWebPart extends BaseClientSideWebPart<IWpOffersListingWebPartProps> {

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpOffersListing.wrapperHtml(resourceUrl, homeUrl, wpOffersListing.loadingHtml);

    const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('Offers')/items?$select=Id,Title,Description,OfferImage,Status,Created&$filter=Status eq 'Active'&$orderby=Created desc&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);
      this._setContent(this._buildListingHtml(items));
    } catch (error) {
      console.error("Error loading offers", error);
      this._setContent(wpOffersListing.noRecordHtml);
    }
  }

  // Get data from the SharePoint list
  private async _getListData(apiUrl: string): Promise<ISPList[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
    }
    const data: ISPLists = await response.json();
    return data.value || [];
  }

  // Build offer cards HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpOffersListing.noRecordHtml;
    }

    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const defaultImageUrl: string = `${webUrl}/SiteAssets/resources/images/DefaultImages/offerimage.png`;
    let html: string = "";

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const view: IOfferCardView = {
        title: escape(item.Title || ""),
        description: escape(this._getPlainText(item.Description || "")),
        imageUrl: this._getImageUrl(item, "Offers") || defaultImageUrl,
        detailsUrl: `${webUrl}/SitePages/OfferDetails.aspx?OfferDetailID=${item.ID}`
      };
      html += wpOffersListing.cardHtml(view);
    }

    return html;
  }

  // Update the listing container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdOfferList');
    if (container !== null) {
      container.innerHTML = html;
    }
  }

  // Convert enhanced rich text HTML to plain text
  private _getPlainText(html: string): string {
    if (!html) {
      return "";
    }
    const doc: Document = new DOMParser().parseFromString(html, "text/html");
    return (doc.body.textContent || "")
      .replace(/\u200B/g, "")
      .replace(/\s+/g, " ")
      .trim();
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