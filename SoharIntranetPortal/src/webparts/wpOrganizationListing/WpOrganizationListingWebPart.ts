
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'WpOrganizationListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import wpEventsListing, {
  IEventCardView
} from './eventsListingHtml';

export interface IWpOrganizationListingWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: number;
  Title: string;
  EventDate: string;
  StartTime: string;
  EndTime: string;
  Location: string;
  Status: string;
  Link?: any;
  Description: string;
}

const PUBLISHED_STATUS: string = 'Active';
const LIST_TITLE: string = 'Upcoming Events';

export default class WpOrganizationListingWebPart
  extends BaseClientSideWebPart<IWpOrganizationListingWebPartProps> {

  private _eventItems: ISPList[] = [];

  private _selectedYear: string = 'All';
  private _selectedMonth: string = 'All';

  private readonly _monthNames: string[] = [
    'January', 'February', 'March', 'April',
    'May', 'June', 'July', 'August',
    'September', 'October', 'November', 'December'
  ];

  public async render(): Promise<void> {

    const webUrl: string =
      this.context.pageContext.web.absoluteUrl;

    const homeUrl: string =
      `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    const resourceUrl: string =
      `${webUrl}/SiteAssets/resources`;

    // Reset selected filters when the webpart renders again
    this._selectedYear = 'All';
    this._selectedMonth = 'All';

    // Render page wrapper and loading content
    this.domElement.innerHTML =
      wpEventsListing.wrapperHtml(
        resourceUrl,
        homeUrl,
        wpEventsListing.loadingHtml
      );

    const apiUrl: string =
      `${webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items` +
      `?$select=Id,Title,EventDate,StartTime,EndTime,Location,Status,Link,Description` +
      `&$filter=Status eq '${PUBLISHED_STATUS}'` +
      `&$orderby=EventDate asc&$top=500`;

    try {

      const items: ISPList[] =
        await this._getListData(apiUrl);

      this._eventItems = items;

      // Populate dropdown options
      this._renderYearFilter(items);
      this._renderMonthFilter(items);

      // Bind dropdown selection events
      this._bindEventFilters();

      // Display all events initially
      this._renderFilteredEvents();

    } catch (error) {

      console.error('Error loading events:', error);

      this._setContent(wpEventsListing.noRecordHtml);
    }
  }

  // Get event data from SharePoint
  private async _getListData(
    apiUrl: string
  ): Promise<ISPList[]> {

    const response: SPHttpClientResponse =
      await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

    if (!response.ok) {
      throw new Error(
        `Request failed: ${response.status} ${response.statusText}`
      );
    }

    const data: ISPLists = await response.json();

    return data.value || [];
  }

  // Extract year and month from EventDate
  private _getEventDateParts(
    value: string
  ): { year: number; month: number } | null {

    if (!value) {
      return null;
    }

    // SharePoint date format: YYYY-MM-DDTHH:mm:ssZ
    const isoMatch = value.match(
      /^(\d{4})-(\d{1,2})-(\d{1,2})/
    );

    if (isoMatch) {
      return {
        year: parseInt(isoMatch[1], 10),
        month: parseInt(isoMatch[2], 10)
      };
    }

    const date: Date = new Date(value);

    if (isNaN(date.getTime())) {
      return null;
    }

    return {
      year: date.getFullYear(),
      month: date.getMonth() + 1
    };
  }

  // Populate Year dropdown
  private _renderYearFilter(items: ISPList[]): void {

    const yearButton: HTMLElement | null =
      this.domElement.querySelector('#dropdownMenuYear');

    const yearMenu: HTMLElement | null =
      yearButton
        ? yearButton.nextElementSibling as HTMLElement
        : null;

    if (!yearMenu) {
      console.warn('Year dropdown menu not found.');
      return;
    }

    const years: number[] = [];

    items.forEach((item: ISPList) => {

      const dateParts =
        this._getEventDateParts(item.EventDate);

      if (
        dateParts &&
        years.indexOf(dateParts.year) === -1
      ) {
        years.push(dateParts.year);
      }
    });

    years.sort((a: number, b: number) => b - a);

    yearMenu.innerHTML =
      wpEventsListing.yearItemHtml('All') +
      years.map((year: number) =>
        wpEventsListing.yearItemHtml(year.toString())
      ).join('');
  }

  // Populate Month dropdown
  private _renderMonthFilter(items: ISPList[]): void {

    const monthButton: HTMLElement | null =
      this.domElement.querySelector('#dropdownMenuMonth');

    const monthMenu: HTMLElement | null =
      monthButton
        ? monthButton.nextElementSibling as HTMLElement
        : null;

    if (!monthMenu) {
      console.warn('Month dropdown menu not found.');
      return;
    }

    const months: number[] = [];

    items.forEach((item: ISPList) => {

      const dateParts =
        this._getEventDateParts(item.EventDate);

      if (
        dateParts &&
        months.indexOf(dateParts.month) === -1
      ) {
        months.push(dateParts.month);
      }
    });

    months.sort((a: number, b: number) => a - b);

    monthMenu.innerHTML =
      wpEventsListing.monthItemHtml('All') +
      months.map((month: number) =>
        wpEventsListing.monthItemHtml(
          this._monthNames[month - 1]
        )
      ).join('');
  }

  // Bind Year and Month dropdown click events
  private _bindEventFilters(): void {

    const yearButton: HTMLElement | null =
      this.domElement.querySelector('#dropdownMenuYear');

    const monthButton: HTMLElement | null =
      this.domElement.querySelector('#dropdownMenuMonth');

    const yearMenu: HTMLElement | null =
      yearButton
        ? yearButton.nextElementSibling as HTMLElement
        : null;

    const monthMenu: HTMLElement | null =
      monthButton
        ? monthButton.nextElementSibling as HTMLElement
        : null;

    // Year selection
    if (yearMenu) {

      yearMenu.addEventListener('click', (event: Event) => {

        const target = event.target as HTMLElement;

        const selectedItem =
          target.closest('[data-year]') as HTMLElement | null;

        if (!selectedItem) {
          return;
        }

        event.preventDefault();

        this._selectedYear =
          selectedItem.getAttribute('data-year') || 'All';

        this._renderFilteredEvents();
      });
    }

    // Month selection
    if (monthMenu) {

      monthMenu.addEventListener('click', (event: Event) => {

        const target = event.target as HTMLElement;

        const selectedItem =
          target.closest('[data-month]') as HTMLElement | null;

        if (!selectedItem) {
          return;
        }

        event.preventDefault();

        this._selectedMonth =
          selectedItem.getAttribute('data-month') || 'All';

        this._renderFilteredEvents();
      });
    }
  }

  // Apply selected Year and Month filters
  private _renderFilteredEvents(): void {

    const yearLabel: HTMLElement | null =
      this.domElement.querySelector('#selectedYearLabel');

    const monthLabel: HTMLElement | null =
      this.domElement.querySelector('#selectedMonthLabel');

    if (yearLabel) {
      yearLabel.textContent = this._selectedYear;
    }

    if (monthLabel) {
      monthLabel.textContent = this._selectedMonth;
    }

    const filteredItems: ISPList[] =
      this._eventItems.filter((item: ISPList) => {

        const dateParts =
          this._getEventDateParts(item.EventDate);

        if (!dateParts) {
          return (
            this._selectedYear === 'All' &&
            this._selectedMonth === 'All'
          );
        }

        const yearMatches =
          this._selectedYear === 'All' ||
          dateParts.year.toString() === this._selectedYear;

        const monthMatches =
          this._selectedMonth === 'All' ||
          this._monthNames[dateParts.month - 1] ===
            this._selectedMonth;

        return yearMatches && monthMatches;
      });

    this._setContent(
      this._buildListingHtml(filteredItems)
    );
  }

  // Build event card HTML
  private _buildListingHtml(items: ISPList[]): string {

    if (!items.length) {
      return wpEventsListing.noRecordHtml;
    }

    let html: string = '';

    for (let i = 0; i < items.length; i++) {

      const item: ISPList = items[i];

      const view: IEventCardView = {

        title: escape(item.Title || ''),

        eventDate: item.EventDate
          ? new Date(item.EventDate).toLocaleDateString(
              'en-US',
              {
                month: 'long',
                day: 'numeric',
                year: 'numeric'
              }
            )
          : '',

        time: this._formatTime(
          item.StartTime,
          item.EndTime
        ),

        location: escape(item.Location || ''),

        description: this._formatRichText(
          item.Description || ''
        ),

        link: this._getLinkUrl(item.Link)
      };

      html += wpEventsListing.cardHtml(view);
    }

    return html;
  }

  // Format start and end time
  private _formatTime(
    startTime: string,
    endTime: string
  ): string {

    const start: string =
      escape((startTime || '').trim());

    const end: string =
      escape((endTime || '').trim());

    if (start && end) {
      return `${start} - ${end}`;
    }

    return start || end;
  }

  // Resolve SharePoint hyperlink column
  private _getLinkUrl(link: any): string {

    if (!link) {
      return '';
    }

    const url: string =
      (typeof link === 'object' ? link.Url : String(link)) || '';

    const trimmedUrl: string = url.trim();

    // Allow only HTTP/HTTPS and site-relative links
    if (
      /^https?:\/\//i.test(trimmedUrl) ||
      trimmedUrl.indexOf('/') === 0
    ) {
      return escape(trimmedUrl);
    }

    return '';
  }

  // Update event listing container
  private _setContent(html: string): void {

    const container: HTMLElement | null =
      this.domElement.querySelector('#IdEventList');

    if (container) {
      container.innerHTML = html;
    }
  }

  // Format rich text description
  private _formatRichText(html: string): string {

    if (!html) {
      return '';
    }

    const doc: Document =
      new DOMParser().parseFromString(html, 'text/html');

    const body: HTMLElement = doc.body;

    const visibleText: string =
      (body.textContent || '')
        .replace(/\u200B/g, '')
        .trim();

    if (!visibleText && !body.querySelector('img')) {
      return '';
    }

    const applyStyle = (
      selector: string,
      styles: { [key: string]: string }
    ): void => {

      const elements: NodeListOf<HTMLElement> =
        body.querySelectorAll(selector) as NodeListOf<HTMLElement>;

      for (let i = 0; i < elements.length; i++) {

        const el: HTMLElement = elements[i];

        for (const prop in styles) {

          if (
            Object.prototype.hasOwnProperty.call(styles, prop) &&
            !el.style.getPropertyValue(prop)
          ) {
            el.style.setProperty(prop, styles[prop]);
          }
        }
      }
    };

    // Keep empty paragraphs visible
    const paragraphs: NodeListOf<HTMLElement> =
      body.querySelectorAll('p') as NodeListOf<HTMLElement>;

    for (let i = 0; i < paragraphs.length; i++) {

      const paragraphText: string =
        (paragraphs[i].textContent || '')
          .replace(/\u200B/g, '')
          .trim();

      if (
        paragraphText === '' &&
        !paragraphs[i].querySelector('img')
      ) {
        paragraphs[i].innerHTML = '&nbsp;';
      }
    }

    applyStyle('p', {
      margin: '0 0 12px 0',
      'line-height': '1.5'
    });

    applyStyle('strong, b', {
      'font-weight': 'bold'
    });

    applyStyle('em, i', {
      'font-style': 'italic'
    });

    applyStyle('u', {
      'text-decoration': 'underline'
    });

    applyStyle('ul', {
      margin: '0 0 12px 0',
      'padding-left': '20px',
      'list-style-type': 'disc'
    });

    applyStyle('ol', {
      margin: '0 0 12px 0',
      'padding-left': '20px',
      'list-style-type': 'decimal'
    });

    applyStyle('li', {
      display: 'list-item'
    });

    return body.innerHTML;
  }

  public async onInit(): Promise<void> {

    await this._loadCSS();

    await this._getEnvironmentMessage();
  }

  // Load CSS
  private _loadCSS(): void {

    const baseUrl: string =
      this.context.pageContext.web.absoluteUrl;

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/list.css`
    );
  }

  private _getEnvironmentMessage(): Promise<string> {

    if (!!this.context.sdks.microsoftTeams) {

      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then(context => {

          let environmentMessage: string = '';

          switch (context.app.host.name) {

            case 'Office':
              environmentMessage =
                this.context.isServedFromLocalhost
                  ? strings.AppLocalEnvironmentOffice
                  : strings.AppOfficeEnvironment;
              break;

            case 'Outlook':
              environmentMessage =
                this.context.isServedFromLocalhost
                  ? strings.AppLocalEnvironmentOutlook
                  : strings.AppOutlookEnvironment;
              break;

            case 'Teams':
            case 'TeamsModern':
              environmentMessage =
                this.context.isServedFromLocalhost
                  ? strings.AppLocalEnvironmentTeams
                  : strings.AppTeamsTabEnvironment;
              break;

            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

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
