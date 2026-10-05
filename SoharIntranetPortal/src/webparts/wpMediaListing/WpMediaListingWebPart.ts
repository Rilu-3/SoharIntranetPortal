import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpMediaListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpMediaListing, { IMediaCardView } from './mediaListingHtml';

export interface IWpMediaListingWebPartProps {
  description: string;
}

// interfaces for sharepoint list items
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: number;
  Title: string;
  Caption: string;
  Status: string;
  SortOrder: number;
  Image: any;
  Year: number;
}

// Only media items with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
const ALL_YEARS: string = 'All';
const LIST_TITLE: string = 'Media_Gallery';

export default class WpMediaListingWebPart extends BaseClientSideWebPart<IWpMediaListingWebPartProps> {

  // All active media loaded from the list (filtered in memory by year)
  private _mediaItems: ISPList[] = [];

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpMediaListing.wrapperHtml(resourceUrl, homeUrl, wpMediaListing.loadingHtml);
    this._bindYearFilter();

    const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items?$select=Id,Title,Caption,Status,SortOrder,Image,Year&$filter=Status eq '${PUBLISHED_STATUS}'&$orderby=SortOrder asc,Id desc&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);

      // Keep only items that have an image
      this._mediaItems = items.filter((item: ISPList) => !!this._getImageUrl(item, LIST_TITLE));

      this._renderYears(this._getDistinctYears(this._mediaItems));
      this._renderMedia(ALL_YEARS);
    } catch (error) {
      console.error("Error loading media gallery", error);
      this._setContent(wpMediaListing.noRecordHtml);
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

  // Distinct years from the items, newest first
  private _getDistinctYears(items: ISPList[]): string[] {
    const years: number[] = [];
    for (let i = 0; i < items.length; i++) {
      const year: number = items[i].Year;
      if (year !== null && year !== undefined && years.indexOf(year) === -1) {
        years.push(year);
      }
    }
    years.sort((a: number, b: number) => b - a);
    return years.map((year: number) => String(year));
  }

  // Populate the year dropdown
  private _renderYears(years: string[]): void {
    const menu: Element | null = this.domElement.querySelector('.sa-filter-dropdown .dropdown-menu');
    if (menu === null) {
      return;
    }

    let html: string = wpMediaListing.yearItemHtml(ALL_YEARS);
    for (let i = 0; i < years.length; i++) {
      html += wpMediaListing.yearItemHtml(escape(years[i]));
    }
    menu.innerHTML = html;
  }

  // Handle year selection (event delegation on the dropdown menu)
  private _bindYearFilter(): void {
    const menu: Element | null = this.domElement.querySelector('.sa-filter-dropdown .dropdown-menu');
    if (menu === null) {
      return;
    }

    menu.addEventListener('click', (event: Event) => {
      const target: HTMLElement = event.target as HTMLElement;
      const year: string | null = target.getAttribute('data-year');
      if (year) {
        this._renderMedia(year);
      }
    });
  }

  // Render media cards for the selected year
  private _renderMedia(year: string): void {
    // Update the dropdown button label
    const label: Element | null = this.domElement.querySelector('#dropdownMenuYear .text-truncate');
    if (label !== null) {
      label.textContent = year;
    }

    const filteredItems: ISPList[] = year === ALL_YEARS
      ? this._mediaItems
      : this._mediaItems.filter((item: ISPList) => String(item.Year) === year);

    this._setContent(this._buildListingHtml(filteredItems));
  }

  // Build media cards HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpMediaListing.noRecordHtml;
    }

    let html: string = "";

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const view: IMediaCardView = {
        title: escape(item.Title || ""),
        caption: escape(this._getPlainText(item.Caption || "")),
        imageUrl: this._getImageUrl(item, LIST_TITLE)
      };
      html += wpMediaListing.cardHtml(view);
    }

    return html;
  }

  // Update the listing container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdMediaList');
    if (container !== null) {
      container.innerHTML = html;
    }
  }

  // Convert rich text HTML to plain text
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
    if (!item.Image) {
      return "";
    }

    try {
      const img: any = JSON.parse(item.Image);
      const webUrl: string = this.context.pageContext.web.absoluteUrl;

      // Image stored in Site Assets (serverRelativeUrl provided)
      if (img.serverRelativeUrl) {
        const origin: string = img.serverUrl || window.location.origin;
        return origin + encodeURI(img.serverRelativeUrl);
      }

      // Image stored as list item attachment (only fileName provided)
      if (img.fileName) {
        return `${webUrl}/Lists/${listUrlName}/Attachments/${item.Id}/${encodeURIComponent(img.fileName)}`;
      }
    } catch (e) {
      console.error("Unable to parse Image", e);
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