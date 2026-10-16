
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

// SharePoint list interfaces
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

// Constants
const PUBLISHED_STATUS: string = 'Active';
const LIST_TITLE: string = 'Birthday';
const ALL_YEARS: string = 'All';
const ALL_MONTHS: string = 'All';

const MONTH_NAMES: string[] = [
  'January', 'February', 'March', 'April',
  'May', 'June', 'July', 'August',
  'September', 'October', 'November', 'December'
];

export default class WpBirthdayListingWebPart extends BaseClientSideWebPart<IWpBirthdayListingWebPartProps> {

  private _birthdayItems: ISPList[] = [];

  private _selectedYear: string = ALL_YEARS;
  private _selectedMonth: string = ALL_MONTHS;

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;

    // Display page with loading indicator
    this.domElement.innerHTML = wpBirthdayListing.wrapperHtml(
      resourceUrl,
      homeUrl,
      wpBirthdayListing.loadingHtml
    );

    // SharePoint Birthday list API
    const apiUrl: string =
      `${webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items` +
      `?$select=Id,Title,EmployeePhoto,BirthDate,Status,Created` +
      `&$filter=Status eq '${PUBLISHED_STATUS}'` +
      `&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);

      // Store and sort birthday records
      this._birthdayItems = this._sortByNearestBirthday(items);

      // Populate Year and Month dropdowns
      this._renderYears(this._getDistinctYears(this._birthdayItems));
      this._renderMonths(this._getDistinctMonths(this._birthdayItems));

      // Bind dropdown click events
      this._bindYearFilter();

      // Display all birthdays initially
      this._renderBirthdays();

    } catch (error) {
      console.error('Error loading birthdays', error);
      this._setContent(wpBirthdayListing.noRecordHtml);
    }
  }

  // Get data from SharePoint list
  private async _getListData(apiUrl: string): Promise<ISPList[]> {
    const response: SPHttpClientResponse =
      await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

    if (!response.ok) {
      throw new Error(
        `Request failed with status ${response.status}: ${response.statusText}`
      );
    }

    const data: ISPLists = await response.json();
    return data.value || [];
  }

  // Extract year and month from BirthDate
  // Supports SharePoint ISO date and MM/DD/YYYY formats
  private _getBirthDateParts(
    value: string
  ): { year: number; month: number } | null {

    if (!value) {
      return null;
    }

    const dateValue: string = value.trim();

    // SharePoint ISO format: 1995-09-21T00:00:00Z
    let match: RegExpExecArray | null =
      /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(dateValue);

    if (match) {
      const year: number = Number(match[1]);
      const month: number = Number(match[2]);

      if (month >= 1 && month <= 12) {
        return { year: year, month: month };
      }
    }

    // MM/DD/YYYY format: 9/21/1995
    match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(dateValue);

    if (match) {
      const month: number = Number(match[1]);
      const year: number = Number(match[3]);

      if (month >= 1 && month <= 12) {
        return { year: year, month: month };
      }
    }

    // Fallback for other valid date formats
    const parsedDate: Date = new Date(dateValue);

    if (!isNaN(parsedDate.getTime())) {
      return {
        year: parsedDate.getFullYear(),
        month: parsedDate.getMonth() + 1
      };
    }

    return null;
  }

  // Get distinct birth years, newest first
  private _getDistinctYears(items: ISPList[]): string[] {
    const years: number[] = [];

    for (let i = 0; i < items.length; i++) {
      const dateParts = this._getBirthDateParts(items[i].BirthDate);

      if (dateParts && years.indexOf(dateParts.year) === -1) {
        years.push(dateParts.year);
      }
    }

    years.sort((a: number, b: number) => b - a);

    return years.map((year: number) => String(year));
  }

  // Get distinct birth months, January to December
  private _getDistinctMonths(items: ISPList[]): string[] {
    const months: number[] = [];

    for (let i = 0; i < items.length; i++) {
      const dateParts = this._getBirthDateParts(items[i].BirthDate);

      if (dateParts && months.indexOf(dateParts.month) === -1) {
        months.push(dateParts.month);
      }
    }

    months.sort((a: number, b: number) => a - b);

    return months.map(
      (month: number) => MONTH_NAMES[month - 1]
    );
  }

  // Populate Year dropdown
  private _renderYears(years: string[]): void {
    const menu: Element | null =
      this.domElement.querySelector('#dropdownMenuYear + .dropdown-menu');

    if (menu === null) {
      return;
    }

    let html: string =
      `<li><a class="dropdown-item" href="#" data-year="${ALL_YEARS}">${ALL_YEARS}</a></li>`;

    for (let i = 0; i < years.length; i++) {
      const year: string = escape(years[i]);

      html +=
        `<li><a class="dropdown-item" href="#" data-year="${year}">${year}</a></li>`;
    }

    menu.innerHTML = html;
  }

  // Populate Month dropdown
  private _renderMonths(months: string[]): void {
    const menu: Element | null =
      this.domElement.querySelector('#dropdownMenuMonth + .dropdown-menu');

    if (menu === null) {
      return;
    }

    let html: string =
      `<li><a class="dropdown-item" href="#" data-month="${ALL_MONTHS}">${ALL_MONTHS}</a></li>`;

    for (let i = 0; i < months.length; i++) {
      const month: string = escape(months[i]);

      html +=
        `<li><a class="dropdown-item" href="#" data-month="${month}">${month}</a></li>`;
    }

    menu.innerHTML = html;
  }

  // Handle Year and Month dropdown selections
  private _bindYearFilter(): void {
    const yearMenu: Element | null =
      this.domElement.querySelector('#dropdownMenuYear + .dropdown-menu');

    const monthMenu: Element | null =
      this.domElement.querySelector('#dropdownMenuMonth + .dropdown-menu');

    if (yearMenu !== null) {
      yearMenu.addEventListener('click', (event: Event) => {
        const target = event.target as HTMLElement;

        const selectedItem: HTMLElement | null =
          target.closest('[data-year]') as HTMLElement | null;

        if (selectedItem) {
          event.preventDefault();

          const year: string | null =
            selectedItem.getAttribute('data-year');

          if (year !== null) {
            this._selectedYear = year;
            this._renderBirthdays();
          }
        }
      });
    }

    if (monthMenu !== null) {
      monthMenu.addEventListener('click', (event: Event) => {
        const target = event.target as HTMLElement;

        const selectedItem: HTMLElement | null =
          target.closest('[data-month]') as HTMLElement | null;

        if (selectedItem) {
          event.preventDefault();

          const month: string | null =
            selectedItem.getAttribute('data-month');

          if (month !== null) {
            this._selectedMonth = month;
            this._renderBirthdays();
          }
        }
      });
    }
  }

  // Filter and display birthday cards
  private _renderBirthdays(): void {

    // Update selected Year label
    const yearLabel: Element | null =
      this.domElement.querySelector('#dropdownMenuYear .text-truncate');

    if (yearLabel !== null) {
      yearLabel.textContent = this._selectedYear;
    }

    // Update selected Month label
    const monthLabel: Element | null =
      this.domElement.querySelector('#dropdownMenuMonth .text-truncate');

    if (monthLabel !== null) {
      monthLabel.textContent = this._selectedMonth;
    }

    // Apply Year and Month filters
    const filteredItems: ISPList[] = this._birthdayItems.filter(
      (item: ISPList) => {

        const dateParts = this._getBirthDateParts(item.BirthDate);

        if (!dateParts) {
          return this._selectedYear === ALL_YEARS &&
            this._selectedMonth === ALL_MONTHS;
        }

        const yearMatch: boolean =
          this._selectedYear === ALL_YEARS ||
          String(dateParts.year) === this._selectedYear;

        const monthMatch: boolean =
          this._selectedMonth === ALL_MONTHS ||
          MONTH_NAMES[dateParts.month - 1] === this._selectedMonth;

        return yearMatch && monthMatch;
      }
    );

    this._setContent(this._buildListingHtml(filteredItems));
  }

  // Calculate days until next birthday
  private _daysUntilBirthday(birthDateValue: string): number {
    if (!birthDateValue) {
      return Number.MAX_VALUE;
    }

    const birthDate: Date = new Date(birthDateValue);

    if (isNaN(birthDate.getTime())) {
      return Number.MAX_VALUE;
    }

    const now: Date = new Date();

    const today: Date = new Date(
      now.getFullYear(),
      now.getMonth(),
      now.getDate()
    );

    let nextBirthday: Date = new Date(
      today.getFullYear(),
      birthDate.getMonth(),
      birthDate.getDate()
    );

    if (nextBirthday.getTime() < today.getTime()) {
      nextBirthday = new Date(
        today.getFullYear() + 1,
        birthDate.getMonth(),
        birthDate.getDate()
      );
    }

    return Math.round(
      (nextBirthday.getTime() - today.getTime()) /
      (1000 * 60 * 60 * 24)
    );
  }

  // Sort nearest birthday first
  private _sortByNearestBirthday(items: ISPList[]): ISPList[] {
    return items.slice().sort((a: ISPList, b: ISPList) => {

      const dayDiff: number =
        this._daysUntilBirthday(a.BirthDate) -
        this._daysUntilBirthday(b.BirthDate);

      if (dayDiff !== 0) {
        return dayDiff;
      }

      return new Date(b.Created).getTime() -
        new Date(a.Created).getTime();
    });
  }

  // Build birthday rows HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpBirthdayListing.noRecordHtml;
    }

    const webUrl: string = this.context.pageContext.web.absoluteUrl;

    const defaultImageUrl: string =
      `${webUrl}/SiteAssets/resources/images/DefaultImages/defaultuser.png`;

    let html: string = '';

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];

      const birthDateParts = this._getBirthDateParts(item.BirthDate);

      let month: string = '';
      let day: string = '';

      if (birthDateParts) {
        month = MONTH_NAMES[birthDateParts.month - 1].substring(0, 3);

        const dateMatch: RegExpExecArray | null =
          /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(item.BirthDate);

        if (dateMatch) {
          day = String(Number(dateMatch[3]));
        } else {
          const parsedDate: Date = new Date(item.BirthDate);

          if (!isNaN(parsedDate.getTime())) {
            day = String(parsedDate.getDate());
          }
        }
      }

      const view: IBirthdayCardView = {
        title: escape(item.Title || ''),
        imageUrl: this._getImageUrl(item, LIST_TITLE) || defaultImageUrl,
        month: month,
        day: day
      };

      html += wpBirthdayListing.cardHtml(view);
    }

    return html;
  }

  // Update birthday list content
  private _setContent(html: string): void {
    const container: Element | null =
      this.domElement.querySelector('#IdBirthdayList');

    if (container !== null) {

      // Update the listing content
      container.innerHTML = html;

      // Remove white panel background when no records are available.
      // Add it back when birthday records are displayed.
      const panelCard: HTMLElement | null =
        container.parentElement as HTMLElement | null;

      if (panelCard) {
        const noRecords: boolean =
          html.trim() === wpBirthdayListing.noRecordHtml.trim();

        if (noRecords) {
          panelCard.classList.remove('panel-card');
        } else {
          panelCard.classList.add('panel-card');
        }
      }
    }
  }

  // Resolve image URL from SharePoint Image column JSON
  private _getImageUrl(item: ISPList, listUrlName: string): string {
    if (!item.EmployeePhoto) {
      return '';
    }

    try {
      const img: any = JSON.parse(item.EmployeePhoto);

      const webUrl: string =
        this.context.pageContext.web.absoluteUrl;

      // Image stored in Site Assets
      if (img.serverRelativeUrl) {
        const origin: string =
          img.serverUrl || window.location.origin;

        return origin + encodeURI(img.serverRelativeUrl);
      }

      // Image stored as list item attachment
      if (img.fileName) {
        return `${webUrl}/Lists/${listUrlName}/Attachments/${item.Id}/${encodeURIComponent(img.fileName)}`;
      }

    } catch (e) {
      console.error('Unable to parse EmployeePhoto', e);
    }

    return '';
  }

  public async onInit(): Promise<void> {
    await this._loadCSS();

    return this._getEnvironmentMessage().then(message => {
      // this._environmentMessage = message;
    });
  }

  private _loadCSS(): void {
    const baseUrl: string =
      this.context.pageContext.web.absoluteUrl;

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/list.css`
    );
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {

          let environmentMessage: string = '';

          switch (context.app.host.name) {

            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;

            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;

            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost
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
