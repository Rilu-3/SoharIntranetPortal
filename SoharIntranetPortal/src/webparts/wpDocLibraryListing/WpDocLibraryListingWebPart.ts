import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpDocLibraryListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpDocLibraryListing, { IDocLibraryRowView } from './docLibraryListingHtml';

export interface IWpDocLibraryListingWebPartProps {
  description: string;
}

// interfaces for sharepoint list items
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: number;
  Title: string;
  Link?: any;
  Status: string;
  SortOrder?: number;
  Created: string;
}

// Only items with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
const LIST_TITLE: string = 'DocLibraryList';

export default class WpDocLibraryListingWebPart extends BaseClientSideWebPart<IWpDocLibraryListingWebPartProps> {

  public async render(): Promise<void> {
    const webUrl: string = 'https://soharaluminium5.sharepoint.com/sites/DevPortal';
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpDocLibraryListing.wrapperHtml(homeUrl, wpDocLibraryListing.loadingHtml);

    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items` +
      `?$select=Id,Title,Link,Status,SortOrder,Created` +
      `&$filter=Status eq '${PUBLISHED_STATUS}'` +
      `&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);
      this._setContent(this._buildListingHtml(this._sortItems(items)));
    } catch (error) {
      console.error("Error loading document libraries", error);
      this._setContent(wpDocLibraryListing.noRecordHtml);
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

  // SortOrder ascending (empty last); same SortOrder -> latest created first
  private _sortItems(items: ISPList[]): ISPList[] {
    const getOrder = (item: ISPList): number =>
      item.SortOrder !== null && item.SortOrder !== undefined ? item.SortOrder : Number.MAX_VALUE;

    return items.slice().sort((a: ISPList, b: ISPList) => {
      const orderDiff: number = getOrder(a) - getOrder(b);
      if (orderDiff !== 0) {
        return orderDiff;
      }
      return new Date(b.Created).getTime() - new Date(a.Created).getTime();
    });
  }

  // Build document library rows HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpDocLibraryListing.noRecordHtml;
    }

    let html: string = "";

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const view: IDocLibraryRowView = {
        title: escape(item.Title || ""),
        link: this._getLinkUrl(item.Link)
      };
      html += wpDocLibraryListing.rowHtml(view);
    }

    return html;
  }

  // Resolve link from a Hyperlink column ({ Url, Description }) or a text column
  private _getLinkUrl(link: any): string {
    if (!link) {
      return "";
    }

    const url: string = (typeof link === 'object' ? link.Url : String(link)) || "";
    const trimmedUrl: string = url.trim();

    // Allow only http(s) or site-relative links
    if (/^https?:\/\//i.test(trimmedUrl) || trimmedUrl.indexOf('/') === 0) {
      return escape(trimmedUrl);
    }
    return "";
  }

  // Update the listing container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdDocLibraryList');
    if (container !== null) {
      container.innerHTML = html;
    }
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
    const baseUrl: string = "https://soharaluminium5.sharepoint.com/sites/DevPortal";

    // // Load all CSS files in parallel
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/variable.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/jquery-ui.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/font-size.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/custom.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/list.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/sp-custom.css`);
    // SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/home.css`);

    // // Load scripts sequentially with proper dependency handling
    // SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`)
    //   .then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery-ui.js`))
    //   .then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`))
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