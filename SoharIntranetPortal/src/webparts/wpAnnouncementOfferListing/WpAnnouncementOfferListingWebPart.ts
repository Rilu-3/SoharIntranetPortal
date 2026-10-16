import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
 
import * as strings from 'WpAnnouncementListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpAnnouncementOffersListing, { IListingCardView, ListingCategory } from './AnnouncementOfferListingHtml';
 
export interface IWpOffersListingWebPartProps {
  description: string;
}
 
// SharePoint list item shapes
export interface IOfferItem {
  ID: number;
  Title: string;
  Description: string;
  OfferImage: any;
  Status: any;
  Created: any;
}
 
export interface IAnnouncementItem {
  ID: number;
  Title: string;
  ShortDescription: string;
  AnnouncementImage: any;
  Status: any;
  Created: any;
  StartDate: any;
  EndDate: any;
}
 
export default class WpOffersListingWebPart extends BaseClientSideWebPart<IWpOffersListingWebPartProps> {
 
  private _selectedCategory: ListingCategory = 'Offers';
  private _requestId: number = 0; // guards against out-of-order responses when switching quickly
  private _categoryInitialised: boolean = false;
 
  // Preselect the category from the page URL, e.g. ...OffersListing.aspx?category=Announcements
  private _getCategoryFromUrl(): ListingCategory | undefined {
    const value: string = (new URLSearchParams(window.location.search).get('category') || '').toLowerCase();
    if (value === 'announcements' || value === 'announcement') {
      return 'Announcements';
    }
    if (value === 'offers' || value === 'offer') {
      return 'Offers';
    }
    return undefined;
  }
 
  public async render(): Promise<void> {
    if (!this._categoryInitialised) {
      this._categoryInitialised = true;
      this._selectedCategory = this._getCategoryFromUrl() || this._selectedCategory;
    }
 
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;
 
    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpAnnouncementOffersListing.wrapperHtml(
      resourceUrl,
      homeUrl,
      wpAnnouncementOffersListing.loadingHtml,
      this._selectedCategory
    );
 
    this._attachCategoryHandlers();
    await this._loadCategory(this._selectedCategory);
  }
 
  // ---------- Category dropdown ----------
 
  private _attachCategoryHandlers(): void {
    const toggleBtn: HTMLElement | null = this.domElement.querySelector('#dropdownMenuCategory');
    const menu: HTMLElement | null = this.domElement.querySelector('.sa-filter-dropdown .dropdown-menu');
    const items: NodeListOf<Element> = this.domElement.querySelectorAll('.sa-filter-dropdown .dropdown-item[data-category]');
 
    if (!toggleBtn || !menu) {
      return;
    }
 
    // The dropdown is controlled here (not by Bootstrap's data-api), so it works
    // whether or not Bootstrap's JS is loaded on the page and never double-toggles.
    const setOpen = (open: boolean): void => {
      menu.classList.toggle('show', open);
      menu.style.display = open ? 'block' : '';
      toggleBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    };
 
    toggleBtn.addEventListener('click', (e: Event) => {
      e.preventDefault();
      e.stopPropagation();
      setOpen(!menu.classList.contains('show'));
    });
 
    // Close when clicking anywhere else (ignore if this web part was re-rendered)
    const onDocumentClick = (): void => {
      if (!this.domElement.contains(menu)) {
        document.removeEventListener('click', onDocumentClick);
        return;
      }
      setOpen(false);
    };
    document.addEventListener('click', onDocumentClick);
 
    for (let i = 0; i < items.length; i++) {
      items[i].addEventListener('click', (e: Event) => {
        e.stopPropagation();
        const category: string | null = (e.currentTarget as HTMLElement).getAttribute('data-category');
        setOpen(false);
        if (category === 'Offers' || category === 'Announcements') {
          this._onCategoryChange(category);
        }
      });
    }
  }
 
  private async _onCategoryChange(category: ListingCategory): Promise<void> {
    if (category === this._selectedCategory) {
      return;
    }
    this._selectedCategory = category;
 
    const label: Element | null = this.domElement.querySelector('#IdSelectedCategory');
    if (label) {
      label.textContent = category;
    }
    const heading: Element | null = this.domElement.querySelector('#IdListingHeading');
    if (heading) {
      heading.textContent = category.toUpperCase();
    }
 
    this._setContent(wpAnnouncementOffersListing.loadingHtml);
    await this._loadCategory(category);
  }
 
  // ---------- Data loading ----------
 
  private async _loadCategory(category: ListingCategory): Promise<void> {
    const currentRequest: number = ++this._requestId;
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
 
    try {
      let html: string;
      if (category === 'Offers') {
        const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('Offers')/items?$select=Id,Title,Description,OfferImage,Status,Created&$filter=Status eq 'Active'&$orderby=Created desc&$top=500`;
        const items: IOfferItem[] = await this._getListData<IOfferItem>(apiUrl);
        html = this._buildOffersHtml(items);
      } else {
        const filter = `Status eq 'Active' and ${await this._getActiveDateFilter(webUrl)}`;
        const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('Announcements')/items`
          + `?$select=Id,Title,ShortDescription,AnnouncementImage,Created,Status,StartDate,EndDate`
          + `&$filter=(${filter})`
          + `&$orderby=StartDate desc,Created desc`
          + `&$top=500`;
        const items: IAnnouncementItem[] = await this._getListData<IAnnouncementItem>(apiUrl);
        html = this._buildAnnouncementsHtml(items);
      }
 
      // Ignore stale responses if the user switched category meanwhile
      if (currentRequest === this._requestId) {
        this._setContent(html);
      }
    } catch (error) {
      console.error(`Error loading ${category}`, error);
      if (currentRequest === this._requestId) {
        this._setContent(wpAnnouncementOffersListing.noRecordHtml);
      }
    }
  }
 
private async _getSiteOffsetMinutes(webUrl: string): Promise<number> {
  const response: SPHttpClientResponse = await this.context.spHttpClient.get(
    `${webUrl}/_api/web/RegionalSettings/TimeZone`,
    SPHttpClient.configurations.v1
  );
  const json = await response.json();
  return -json.Information.Bias; // Abu Dhabi/Muscat = 240
}
 
private async _getActiveDateFilter(webUrl: string): Promise<string> {
  const offsetMin = await this._getSiteOffsetMinutes(webUrl);
 
  // The viewer's own calendar date, e.g. "2026-10-16"
  const userTimeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const parts = new Date().toLocaleDateString('en-CA', { timeZone: userTimeZone }).split('-');
 
  // That date at 00:00, shifted by the site offset to match how SharePoint stored dates
  const threshold = new Date(
    Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2])) - offsetMin * 60000
  ).toISOString();
 
  return `StartDate le datetime'${threshold}' and EndDate ge datetime'${threshold}'`;
}
  // Get data from a SharePoint list
  private async _getListData<T>(apiUrl: string): Promise<T[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
    }
    const data: { value: T[] } = await response.json();
    return data.value || [];
  }
 
  // ---------- HTML builders ----------
 
  private _buildOffersHtml(items: IOfferItem[]): string {
    if (!items.length) {
      return wpAnnouncementOffersListing.noRecordHtml;
    }
 
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const defaultImageUrl: string = `${webUrl}/SiteAssets/resources/images/DefaultImages/offerimage.png`;
    let html: string = "";
 
    for (let i = 0; i < items.length; i++) {
      const item: IOfferItem = items[i];
      const view: IListingCardView = {
        title: escape(item.Title || ""),
        description: escape(this._getPlainText(item.Description || "")),
        imageUrl: this._getImageUrl(item.OfferImage, item.ID, "Offers") || defaultImageUrl,
        detailsUrl: `${webUrl}/SitePages/OfferDetails.aspx?OfferDetailID=${item.ID}&env=WebViewList`
      };
      html += wpAnnouncementOffersListing.cardHtml(view);
    }
 
    return html;
  }
 
  private _buildAnnouncementsHtml(items: IAnnouncementItem[]): string {
    if (!items.length) {
      return wpAnnouncementOffersListing.noRecordHtml;
    }
 
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const defaultImageUrl: string = `${webUrl}/SiteAssets/resources/images/DefaultImages/announcementimage.png`;
    let html: string = "";
 
    for (let i = 0; i < items.length; i++) {
      const item: IAnnouncementItem = items[i];
      const view: IListingCardView = {
        title: escape(item.Title || ""),
        description: escape(this._getPlainText(item.ShortDescription || "")),
        imageUrl: this._getImageUrl(item.AnnouncementImage, item.ID, "Announcements") || defaultImageUrl,
        detailsUrl: `${webUrl}/SitePages/AnnouncementDetails.aspx?AnnoDetailID=${item.ID}&env=WebViewList`
      };
      html += wpAnnouncementOffersListing.cardHtml(view);
    }
 
    return html;
  }
 
  // Update the listing container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdListingList');
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
  private _getImageUrl(imageColumn: any, itemId: number, listUrlName: string): string {
    if (!imageColumn) {
      return "";
    }
 
    try {
      const img: any = JSON.parse(imageColumn);
      const webUrl: string = this.context.pageContext.web.absoluteUrl;
 
      // Image stored in Site Assets (serverRelativeUrl provided)
      if (img.serverRelativeUrl) {
        const origin: string = img.serverUrl || window.location.origin;
        return origin + encodeURI(img.serverRelativeUrl);
      }
 
      // Image stored as list item attachment (only fileName provided)
      if (img.fileName) {
        return `${webUrl}/Lists/${listUrlName}/Attachments/${itemId}/${encodeURIComponent(img.fileName)}`;
      }
    } catch (e) {
      console.error(`Unable to parse image column for ${listUrlName}`, e);
    }
 
    return "";
  }
 
  // ---------- Lifecycle / setup ----------
 
  public async onInit(): Promise<void> {
    await this._loadCSS();
    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }
 
  private _loadCSS(): void {
    const baseUrl: string = this.context.pageContext.web.absoluteUrl;
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/list.css`);
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
 