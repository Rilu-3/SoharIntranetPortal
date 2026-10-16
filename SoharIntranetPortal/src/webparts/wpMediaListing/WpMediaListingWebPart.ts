
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

// Interfaces for SharePoint list items
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
  GalleryDate: string;
}

// Only media items with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
const ALL_YEARS: string = 'All';
const ALL_MONTHS: string = 'All';
const LIST_TITLE: string = 'Media_Gallery';

const MONTH_NAMES: string[] = [
  'January', 'February', 'March', 'April',
  'May', 'June', 'July', 'August',
  'September', 'October', 'November', 'December'
];

export default class WpMediaListingWebPart extends BaseClientSideWebPart<IWpMediaListingWebPartProps> {

  // All active media loaded from the list
  private _mediaItems: ISPList[] = [];

  // Selected filter values
  private _selectedYear: string = ALL_YEARS;
  private _selectedMonth: string = ALL_MONTHS;

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpMediaListing.wrapperHtml(resourceUrl, homeUrl, wpMediaListing.loadingHtml);
    this._bindYearFilter();
    this._bindGalleryModal();

    const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items?$select=Id,Title,Caption,Status,SortOrder,Image,GalleryDate,Created&$filter=Status eq '${PUBLISHED_STATUS}'&$orderby=SortOrder asc,Created desc&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);

      // Keep only items that have an image
      this._mediaItems = items.filter((item: ISPList) => !!this._getImageUrl(item, LIST_TITLE));

      // Populate Year and Month dropdowns from GalleryDate
      this._renderYears(this._getDistinctYears(this._mediaItems));
      this._renderMonths(this._getDistinctMonths(this._mediaItems));

      // Display all media initially
      this._renderMedia();

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

  // Extract year and month from GalleryDate
  // Supports SharePoint ISO date and MM/DD/YYYY format
  private _getGalleryDateParts(value: string): { year: number; month: number } | null {
    if (!value) {
      return null;
    }

    const dateValue: string = value.trim();

    // SharePoint REST format: 2026-07-11T00:00:00Z
    let match: RegExpExecArray | null = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(dateValue);

    if (match) {
      const year: number = Number(match[1]);
      const month: number = Number(match[2]);

      if (month >= 1 && month <= 12) {
        return { year: year, month: month };
      }
    }

    // Date format: 7/11/2026 (Month/Day/Year)
    match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(dateValue);

    if (match) {
      const month: number = Number(match[1]);
      const year: number = Number(match[3]);

      if (month >= 1 && month <= 12) {
        return { year: year, month: month };
      }
    }

    return null;
  }

  // Distinct years from GalleryDate, newest first
  private _getDistinctYears(items: ISPList[]): string[] {
    const years: number[] = [];

    for (let i = 0; i < items.length; i++) {
      const dateParts = this._getGalleryDateParts(items[i].GalleryDate);

      if (dateParts && years.indexOf(dateParts.year) === -1) {
        years.push(dateParts.year);
      }
    }

    years.sort((a: number, b: number) => b - a);

    return years.map((year: number) => String(year));
  }

  // Distinct months from GalleryDate
  private _getDistinctMonths(items: ISPList[]): string[] {
    const months: number[] = [];

    for (let i = 0; i < items.length; i++) {
      const dateParts = this._getGalleryDateParts(items[i].GalleryDate);

      if (dateParts && months.indexOf(dateParts.month) === -1) {
        months.push(dateParts.month);
      }
    }

    months.sort((a: number, b: number) => a - b);

    return months.map((month: number) => MONTH_NAMES[month - 1]);
  }

  // Populate the year dropdown
  private _renderYears(years: string[]): void {
    const menu: Element | null = this.domElement.querySelector('#dropdownMenuYear + .dropdown-menu');

    if (menu === null) {
      return;
    }

    let html: string = wpMediaListing.yearItemHtml(ALL_YEARS);

    for (let i = 0; i < years.length; i++) {
      html += wpMediaListing.yearItemHtml(escape(years[i]));
    }

    menu.innerHTML = html;
  }

  // Populate the month dropdown
  private _renderMonths(months: string[]): void {
    const menu: Element | null = this.domElement.querySelector('#dropdownMenuMonth + .dropdown-menu');

    if (menu === null) {
      return;
    }

    let html: string = wpMediaListing.monthItemHtml(ALL_MONTHS);

    for (let i = 0; i < months.length; i++) {
      html += wpMediaListing.monthItemHtml(escape(months[i]));
    }

    menu.innerHTML = html;
  }

  // Handle year and month selection
  private _bindYearFilter(): void {
    const yearMenu: Element | null = this.domElement.querySelector('#dropdownMenuYear + .dropdown-menu');
    const monthMenu: Element | null = this.domElement.querySelector('#dropdownMenuMonth + .dropdown-menu');

    if (yearMenu !== null) {
      yearMenu.addEventListener('click', (event: Event) => {
        const target: HTMLElement = event.target as HTMLElement;
        const year: string | null = target.getAttribute('data-year');

        if (year !== null) {
          this._selectedYear = year;
          this._renderMedia();
        }
      });
    }

    if (monthMenu !== null) {
      monthMenu.addEventListener('click', (event: Event) => {
        const target: HTMLElement = event.target as HTMLElement;
        const month: string | null = target.getAttribute('data-month');

        if (month !== null) {
          this._selectedMonth = month;
          this._renderMedia();
        }
      });
    }
  }

  // Handle gallery image click and open selected image in modal

  // Handle gallery image click and open selected image in modal
  private _bindGalleryModal(): void {

    const mediaList: Element | null =
      this.domElement.querySelector('#IdMediaList');

    if (!mediaList) {
      return;
    }

    mediaList.addEventListener('click', (event: Event) => {

      const target: Element = event.target as Element;

      const imageTrigger: Element | null =
        target.closest('.gallery-image-trigger');

      if (!imageTrigger) {
        return;
      }

      event.preventDefault();

      const imageUrl: string | null =
        imageTrigger.getAttribute('data-gallery-image');

      const modalImage: HTMLImageElement | null =
        this.domElement.querySelector('#galleryModalImage');

      const modalElement: HTMLElement | null =
        this.domElement.querySelector('#galleryModal');

      const bootstrap: any = (window as any).bootstrap;

      if (
        imageUrl &&
        modalImage &&
        modalElement &&
        bootstrap &&
        bootstrap.Modal
      ) {

        modalImage.src = imageUrl;

        const modalInstance =
          bootstrap.Modal.getOrCreateInstance(modalElement);

        modalElement.addEventListener(
          'hidden.bs.modal',
          () => {
            (imageTrigger as HTMLElement).focus();
          },
          { once: true }
        );

        modalInstance.show();
      }

    });
  }

  // Render media cards for selected year and month
  private _renderMedia(): void {

    // Update year dropdown button label
    const yearLabel: Element | null = this.domElement.querySelector('#dropdownMenuYear .text-truncate');

    if (yearLabel !== null) {
      yearLabel.textContent = this._selectedYear;
    }

    // Update month dropdown button label
    const monthLabel: Element | null = this.domElement.querySelector('#dropdownMenuMonth .text-truncate');

    if (monthLabel !== null) {
      monthLabel.textContent = this._selectedMonth;
    }

    // Filter media using GalleryDate
    const filteredItems: ISPList[] = this._mediaItems.filter((item: ISPList) => {
      const dateParts = this._getGalleryDateParts(item.GalleryDate);

      if (!dateParts) {
        return this._selectedYear === ALL_YEARS && this._selectedMonth === ALL_MONTHS;
      }

      const yearMatch: boolean =
        this._selectedYear === ALL_YEARS ||
        String(dateParts.year) === this._selectedYear;

      const monthMatch: boolean =
        this._selectedMonth === ALL_MONTHS ||
        MONTH_NAMES[dateParts.month - 1] === this._selectedMonth;

      return yearMatch && monthMatch;
    });

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
        caption: "",
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
//   private _getPlainText(html: string): string {
//   if (!html) {
//     return "";
//   }

//   const doc: Document = new DOMParser().parseFromString(html, "text/html");

//   return (doc.body.textContent || "")
//     .replace(/\u200B/g, "")
//     .replace(/\s+/g, " ")
//     .trim();
// }

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
    SPComponentLoader.loadScript(
  `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
);

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

            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;

            case 'Teams':
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
