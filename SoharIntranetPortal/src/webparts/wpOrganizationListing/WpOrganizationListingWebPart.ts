import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WpOrganizationListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpEventsListing, { IEventCardView } from './eventsListingHtml';

export interface IWpOrganizationListingWebPartProps {
  description: string;
}

// interfaces for sharepoint list items
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

// Only events with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
const LIST_TITLE: string = 'Upcoming Events';

export default class WpOrganizationListingWebPart extends BaseClientSideWebPart<IWpOrganizationListingWebPartProps> {

  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;

    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpEventsListing.wrapperHtml(homeUrl, wpEventsListing.loadingHtml);

    const apiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items?$select=Id,Title,EventDate,StartTime,EndTime,Location,Status,Link,Description&$filter=Status eq '${PUBLISHED_STATUS}'&$orderby=EventDate asc&$top=500`;

    try {
      const items: ISPList[] = await this._getListData(apiUrl);
      this._setContent(this._buildListingHtml(items));
    } catch (error) {
      console.error("Error loading events", error);
      this._setContent(wpEventsListing.noRecordHtml);
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

  // Build event cards HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpEventsListing.noRecordHtml;
    }

    let html: string = "";

    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const view: IEventCardView = {
        title: escape(item.Title || ""),
        eventDate: item.EventDate
          ? new Date(item.EventDate).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' })
          : "",
        time: this._formatTime(item.StartTime, item.EndTime),
        location: escape(item.Location || ""),
        description: this._formatRichText(item.Description || ""),
        link: this._getLinkUrl(item.Link)
      };
      html += wpEventsListing.cardHtml(view);
    }

    return html;
  }

  // Combine start and end time (shows whichever exists)
  private _formatTime(startTime: string, endTime: string): string {
    const start: string = escape((startTime || "").trim());
    const end: string = escape((endTime || "").trim());

    if (start && end) {
      return `${start} - ${end}`;
    }
    return start || end;
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
    const container: Element | null = this.domElement.querySelector('#IdEventList');
    if (container !== null) {
      container.innerHTML = html;
    }
  }

  // Apply inline styles so enhanced rich text keeps its list formatting
  private _formatRichText(html: string): string {
    if (!html) {
      return "";
    }

    const doc: Document = new DOMParser().parseFromString(html, "text/html");
    const body: HTMLElement = doc.body;

    // Treat description with no visible text as empty
    const visibleText: string = (body.textContent || "").replace(/\u200B/g, "").trim();
    if (!visibleText && !body.querySelector("img")) {
      return "";
    }

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