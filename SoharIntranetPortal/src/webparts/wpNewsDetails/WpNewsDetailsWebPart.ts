import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpNewsDetailsWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import NewsDetails, { INewsView } from './NewsDetailsHtml';

export interface IWpNewsDetailsWebPartProps {
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
  MainContent?: string;
  PublishedDate?: string;
  Status?: string;
  NewsImage?: any;
}

// Only news items with this status are displayed
const PUBLISHED_STATUS: string = 'Active';

export default class WpNewsDetailsWebPart extends BaseClientSideWebPart<IWpNewsDetailsWebPartProps> {

  public async render(): Promise<void> {
    // Show loader until data is ready
    this.domElement.innerHTML = NewsDetails.wrapperHtml(NewsDetails.loadingHtml);

    // function call to extract query string parameters
    const queryStringParams: any = this.getQueryStringParameters();
    const id: number = parseInt(queryStringParams['NewsDetailID'], 10);

    if (isNaN(id)) {
      this._setContent(NewsDetails.noRecordHtml);
      return;
    }

    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('News')/items(${id})?$select=Id,Title,ShortDescription,Category,MainContent,PublishedDate,Status,NewsImage`;

    try {
      const item: ISPList = await this._getListData(apiUrl);

      // Show only active news
      if (item.Status !== PUBLISHED_STATUS) {
        this._setContent(NewsDetails.noRecordHtml);
        return;
      }

      this._setContent(this._buildDetailsHtml(item));
    } catch (error) {
      console.error("Error loading news", error);
      this._setContent(NewsDetails.noRecordHtml);
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

  // Build news details HTML from list item
  private _buildDetailsHtml(item: ISPList): string {
    const view: INewsView = {
      title: escape(item.Title || ""),
      category: escape(item.Category || ""),
      publishedDate: item.PublishedDate
        ? new Date(item.PublishedDate).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
        : "",
      shortDescription: this._formatShortDescription(item.ShortDescription || ""),
      mainContent: this._formatRichText(item.MainContent || ""),
      imageUrl: this._getImageUrl(item, "News")
    };
    return NewsDetails.detailsHtml(view);
  }

  // Update the web part content area
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#IdNewsDetail');
    if (container !== null) {
      container.innerHTML = html;
    }
  }

  // Short description: plain text with line breaks preserved (falls back to rich text if HTML is returned)
  private _formatShortDescription(value: string): string {
    if (!value) {
      return "";
    }
    if (/<[a-z][\s\S]*>/i.test(value)) {
      return this._formatRichText(value);
    }
    return `<p>${escape(value).replace(/\r?\n/g, "<br>")}</p>`;
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