import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpNewsListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpNewsListing, { INewsCardView } from './newsListingHtml';

export interface IWpNewsListingWebPartProps {
  description: string;
}

// interfaces for sharepoint list items
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id: number;
  Title: string;
  ShortDescription?: string;
  Category?: string;
  PublishedDate?: string;
  Status?: string;
  NewsImage?: any;
}

// Only news items with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
const ALL_CATEGORIES: string = 'All';

export default class WpNewsListingWebPart extends BaseClientSideWebPart<IWpNewsListingWebPartProps> {

  // All active news loaded from the list (filtered in memory by category)
  private _newsItems: ISPList[] = [];

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpNewsListing.wrapperHtml(resourceUrl, homeUrl, wpNewsListing.loadingHtml);
    this._bindCategoryFilter();

    const listApiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('News')`;
    const itemsApiUrl: string = `${listApiUrl}/items?$select=Id,Title,ShortDescription,Category,PublishedDate,Status,NewsImage&$filter=Status eq '${PUBLISHED_STATUS}'&$orderby=PublishedDate desc&$top=500`;
    const categoryApiUrl: string = `${listApiUrl}/fields/getbyinternalnameortitle('Category')?$select=Choices`;

    try {
      const [items, categories] = await Promise.all([
        this._getNewsItems(itemsApiUrl),
        this._getCategories(categoryApiUrl)
      ]);

      this._newsItems = items;
      this._renderCategories(categories);
      this._renderNews(ALL_CATEGORIES);
    } catch (error) {
      console.error("Error loading news", error);
      this._setContent(wpNewsListing.noRecordHtml);
    }
  }

  // Generic GET request returning JSON
  private async _getJson(apiUrl: string): Promise<any> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
    }
    return response.json();
  }

  // Get active news items
  private async _getNewsItems(apiUrl: string): Promise<ISPList[]> {
    const data: ISPLists = await this._getJson(apiUrl);
    return data.value || [];
  }

  // Get Category choice values (listing still loads if this fails)
  private async _getCategories(apiUrl: string): Promise<string[]> {
    try {
      const data: any = await this._getJson(apiUrl);
      return data.Choices || [];
    } catch (error) {
      console.error("Error loading categories", error);
      return [];
    }
  }

  // Populate the category dropdown
  private _renderCategories(categories: string[]): void {
    const menu: Element | null = this.domElement.querySelector('.sa-filter-dropdown .dropdown-menu');
    if (menu === null) {
      return;
    }

    let html: string = wpNewsListing.categoryItemHtml(ALL_CATEGORIES);
    for (let i = 0; i < categories.length; i++) {
      html += wpNewsListing.categoryItemHtml(escape(categories[i]));
    }
    menu.innerHTML = html;
  }

  // Handle category selection (event delegation on the dropdown menu)
  private _bindCategoryFilter(): void {
    const menu: Element | null = this.domElement.querySelector('.sa-filter-dropdown .dropdown-menu');
    if (menu === null) {
      return;
    }

    menu.addEventListener('click', (event: Event) => {
      const target: HTMLElement = event.target as HTMLElement;
      const category: string | undefined = target.getAttribute('data-category') || undefined;
      if (category) {
        this._renderNews(category);
      }
    });
  }

  // Render news cards for the selected category
  private _renderNews(category: string): void {
    // Update the dropdown button label
    const label: Element | null = this.domElement.querySelector('#dropdownMenuCategory .text-truncate');
    if (label !== null) {
      label.textContent = category;
    }

    const filteredItems: ISPList[] = category === ALL_CATEGORIES
      ? this._newsItems
      : this._newsItems.filter((item: ISPList) => item.Category === category);

    this._setContent(this._buildListingHtml(filteredItems));
  }

  // Build news cards HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpNewsListing.noRecordHtml;
    }

    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const defaultImageUrl: string = `${webUrl}/SiteAssets/resources/images/DefaultImages/newsimage.png`;
    let html: string = "";

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const view: INewsCardView = {
        title: escape(item.Title || ""),
        description: escape(this._getPlainText(item.ShortDescription || "")),
        imageUrl: this._getImageUrl(item, "News") || defaultImageUrl,
        detailsUrl: `${webUrl}/SitePages/NewsDetails.aspx?NewsDetailID=${item.Id}&env=WebViewList`
      };
      html += wpNewsListing.cardHtml(view);
    }

    return html;
  }

  // Update the listing container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdNewsList');
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
    if (!item.NewsImage) {
      return "";
    }

    try {
      const img: any = JSON.parse(item.NewsImage);
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
      console.error("Unable to parse NewsImage", e);
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