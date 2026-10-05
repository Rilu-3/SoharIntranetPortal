import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpBirthdayListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpBirthdayListing, { IBirthdayCardView } from './birthdayListingHtml';

export interface IWpBirthdayListingWebPartProps {
  description: string;
}

// interfaces for sharepoint list items
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: number;
  Title: string;
  EmployeePhoto?: any;
  BirthDate: string;
  Status: string;
  Created: string;
}

// Only birthdays with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
const LIST_TITLE: string = 'Birthday';

export default class WpBirthdayListingWebPart extends BaseClientSideWebPart<IWpBirthdayListingWebPartProps> {

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpBirthdayListing.wrapperHtml(homeUrl, wpBirthdayListing.loadingHtml);

    const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items` +
      `?$select=Id,Title,EmployeePhoto,BirthDate,Status,Created` +
      `&$filter=Status eq '${PUBLISHED_STATUS}'` +
      `&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);
      this._setContent(this._buildListingHtml(this._sortByNearestBirthday(items)));
    } catch (error) {
      console.error("Error loading birthdays", error);
      this._setContent(wpBirthdayListing.noRecordHtml);
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

  // Days from today until the next occurrence of the birthday (0 = today)
  private _daysUntilBirthday(birthDateValue: string): number {
    if (!birthDateValue) {
      return Number.MAX_VALUE;
    }

    const birthDate: Date = new Date(birthDateValue);
    const now: Date = new Date();
    const today: Date = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    let nextBirthday: Date = new Date(today.getFullYear(), birthDate.getMonth(), birthDate.getDate());
    if (nextBirthday.getTime() < today.getTime()) {
      nextBirthday = new Date(today.getFullYear() + 1, birthDate.getMonth(), birthDate.getDate());
    }

    return Math.round((nextBirthday.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  }

  // Nearest birthday first; same day -> latest created first
  private _sortByNearestBirthday(items: ISPList[]): ISPList[] {
    return items.slice().sort((a: ISPList, b: ISPList) => {
      const dayDiff: number = this._daysUntilBirthday(a.BirthDate) - this._daysUntilBirthday(b.BirthDate);
      if (dayDiff !== 0) {
        return dayDiff;
      }
      return new Date(b.Created).getTime() - new Date(a.Created).getTime();
    });
  }

  // Build birthday rows HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpBirthdayListing.noRecordHtml;
    }

    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const defaultImageUrl: string = `${webUrl}/SiteAssets/resources/images/DefaultImages/defaultuser.png`;
    let html: string = "";

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const birthDate: Date | null = item.BirthDate ? new Date(item.BirthDate) : null;

      const view: IBirthdayCardView = {
        title: escape(item.Title || ""),
        imageUrl: this._getImageUrl(item, LIST_TITLE) || defaultImageUrl,
        month: birthDate ? birthDate.toLocaleDateString('en-US', { month: 'short' }) : "",
        day: birthDate ? String(birthDate.getDate()) : ""
      };
      html += wpBirthdayListing.cardHtml(view);
    }

    return html;
  }

  // Update the listing container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdBirthdayList');
    if (container !== null) {
      container.innerHTML = html;
    }
  }

  // Resolve image URL from SharePoint Image column JSON
  private _getImageUrl(item: ISPList, listUrlName: string): string {
    if (!item.EmployeePhoto) {
      return "";
    }

    try {
      const img: any = JSON.parse(item.EmployeePhoto);
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
      console.error("Unable to parse EmployeePhoto", e);
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