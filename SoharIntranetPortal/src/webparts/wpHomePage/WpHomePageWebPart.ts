import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import UabHomePage from './UabHomePage';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import UabAnnouncements from './UabAnnouncements';
import UabOffers from './UabOffers';


/*
 * ============================================================
 * INTERFACES
 * ============================================================
 *
 * These interfaces define the structure of the data returned
 * from the SharePoint Announcements and Offers lists.
 */


/*
 * Represents one item from the Announcements SharePoint list.
 */
export interface IWpHomePageWebPartProps {
  description: string;
}

interface IAnnouncement {
  Id: number;
  Title: string;
  ShortDescription: string;
  Icon: string;
  Created: string;
}


/*
 * Represents one item from the Offers SharePoint list.
 */
interface IOffer {
  Id: number;
  Title: string;
  Description: string;
  OfferImage: string;
  Created: string;
}


/*
 * ============================================================
 * WEB PART CLASS
 * ============================================================
 */
export default class WpHomePageWebPart extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  /*
   * ==========================================================
   * INITIALIZATION
   * ==========================================================
   *
   * onInit() runs when the Web Part is initialized.
   *
   * Here we load all required external CSS and JavaScript
   * files from the SharePoint SiteAssets library.
   */
  public async onInit(): Promise<void> {
    await this.loadCSS();
  }


  /*
   * ==========================================================
   * RENDER HOME PAGE
   * ==========================================================
   *
   * render() is the main method responsible for:
   *
   * 1. Loading the main home page HTML.
   * 2. Loading the Announcement HTML structure.
   * 3. Loading the Offers HTML structure.
   * 4. Setting up the "View All" link.
   * 5. Creating the SharePoint REST API URLs.
   * 6. Loading Announcement data.
   * 7. Loading Offer data.
   */
  public async render(): Promise<void> {
    this.domElement.innerHTML = UabHomePage.allElementsHtml;

    this.domElement.querySelector("#announcement")!.innerHTML =
      UabAnnouncements.allElementsHtml;

    this.domElement.querySelector("#offers")!.innerHTML =
      UabOffers.allElementsHtml;

    this.setupViewAllLink();

    const AnnouncementApiUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Announcements')/items?$select=Id,Title,ShortDescription,Icon,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc`;

    const OfferApiUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Offers')/items?$select=Id,Title,Description,OfferImage,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc`;

    await this._renderAnnouncementsAsync(AnnouncementApiUrl);

    await this._renderOffersAsync(OfferApiUrl);
  }


  /*
   * ==========================================================
   * LOAD CSS AND JAVASCRIPT
   * ==========================================================
   *
   * Loads the external CSS and JavaScript files required
   * by the home page.
   *
   * All files are stored inside:
   *
   * /SiteAssets/resources/
   */
  private async loadCSS(): Promise<void> {

    const baseUrl = this.context.pageContext.web.absoluteUrl;

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/font-size.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`
    );

    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/common.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/home.js`
    );
  }


  /*
   * ==========================================================
   * VIEW ALL LINK
   * ==========================================================
   *
   * The home page has two tabs:
   *
   *     Announcements
   *     Offers
   *
   * Both tabs use the same "View All" link.
   *
   * When the user clicks a tab, the URL of the View All
   * link is changed according to the selected tab.
   */
  private setupViewAllLink(): void {
    const tabs = this.domElement.querySelectorAll('[data-tab-ao]');

    const viewAllLink =
      this.domElement.querySelector('#ao-view-all') as HTMLAnchorElement;

    if (!viewAllLink) {
      console.error('View All link not found');
      return;
    }

    tabs.forEach((tab) => {
      tab.addEventListener('click', () => {

        const selectedTab = tab.getAttribute('data-tab-ao');

        if (selectedTab === 'announcement') {
          viewAllLink.href =
            `${this.context.pageContext.web.absoluteUrl}/SitePages/Announcements.aspx`;
        }

        if (selectedTab === 'offers') {
          viewAllLink.href =
            `${this.context.pageContext.web.absoluteUrl}/SitePages/Offer-List.aspx`;
        }
      });
    });
  }


  /*
   * ==========================================================
   * RENDER ANNOUNCEMENTS
   * ==========================================================
   *
   * Gets Announcement data from SharePoint and converts each
   * SharePoint item into the Announcement HTML template.
   */
  private async _renderAnnouncementsAsync(apiUrl: string): Promise<void> {

    const data: IAnnouncement[] =
      await this._getAnnouncementsData(apiUrl);

    console.log("Announcements data", data);

    let allElementsHtml: string = "";

    try {

      data.forEach((item, index) => {

        let imageUrl = '';

        /*
         * The Icon column contains JSON data.
         * Extract the file name from the JSON and construct
         * the SharePoint attachment URL.
         */
        if (item.Icon) {

          const imageData = JSON.parse(item.Icon);

          // console.log(imageData);

          const fileName = imageData.fileName;

          // Build the image URL using the fileName
          imageUrl =
            `${this.context.pageContext.web.absoluteUrl}/Lists/Announcements/Attachments/${item.Id}/${fileName}`;

          // console.log(imageUrl);
        }

        /*
         * Convert the SharePoint Created date into the
         * required display format.
         */
        let createddate = this.formatDate(item.Created);

        /*
         * Replace the placeholders in the Announcement
         * HTML template with actual SharePoint data.
         */
        let singleElementHtml = UabAnnouncements.singleElementHtml
          .replace("__KEY__ANNOUNCEMENT__ICON__", imageUrl)
          .replace("__KEY__ANNOUNCEMENT__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENT__DESCRIPTION__", item.ShortDescription)
          .replace("__KEY__ANNOUNCEMENT__DATE__", createddate);

        /*
         * Add the generated Announcement HTML to the
         * complete HTML string.
         */
        allElementsHtml += singleElementHtml;
      })

    }
    catch (error) {
      console.error('Error rendering QuickList:', error);
    }

    /*
     * Insert all generated Announcement HTML into the
     * Announcement container.
     */
    this.domElement.querySelector("#announcement-container")!.innerHTML =
      allElementsHtml;
  }


  /*
   * ==========================================================
   * RENDER OFFERS
   * ==========================================================
   *
   * Gets Offer data from SharePoint and converts each
   * SharePoint item into the Offer HTML template.
   */
  private async _renderOffersAsync(apiUrl: string): Promise<void> {

    const data: IOffer[] =
      await this._getOffersData(apiUrl);

    console.log("Offers data", data);

    let allElementsHtml: string = "";

    try {

      data.forEach((item, index) => {

        let imageUrl = '';

        /*
         * The OfferImage column contains JSON data.
         * Extract the file name and construct the
         * SharePoint attachment URL.
         */
        if (item.OfferImage) {

          const imageData = JSON.parse(item.OfferImage);

          // console.log(imageData);

          const fileName = imageData.fileName;

          // Build the image URL using the fileName
          imageUrl =
            `${this.context.pageContext.web.absoluteUrl}/Lists/Offers/Attachments/${item.Id}/${fileName}`;

          // console.log(imageUrl);
        }

        /*
         * Convert the SharePoint Created date into the
         * required display format.
         */
        let createddate = this.formatDate(item.Created);

        /*
         * Replace the placeholders in the Offer HTML template
         * with actual SharePoint data.
         */
        let singleElementHtml = UabOffers.singleElementHtml
          .replace("__KEY__OFFER__ICON__", imageUrl)
          .replace("__KEY__OFFER__TITLE__", item.Title)
          .replace("__KEY__OFFER__DESCRIPTION__", item.Description)
          .replace("__KEY__OFFER__DATE__", createddate);

        /*
         * Add the generated Offer HTML to the complete
         * HTML string.
         */
        allElementsHtml += singleElementHtml;

      })

    }
    catch (error) {
      console.error('Error rendering QuickList:', error);
    }

    /*
     * Insert all generated Offer HTML into the Offer container.
     */
    this.domElement.querySelector("#offer-container")!.innerHTML =
      allElementsHtml;
  }


  /*
   * ==========================================================
   * GET ANNOUNCEMENT DATA
   * ==========================================================
   *
   * Sends a GET request to the SharePoint REST API and
   * returns the Announcement list items.
   */
  private async _getAnnouncementsData(
    apiUrl: string
  ): Promise<IAnnouncement[]> {

    try {

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

      if (response.ok) {

        const data = await response.json();

        return data.value;

      } else {

        console.error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );

        throw new Error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );
      }

    }
    catch (error) {

      console.log("error occured", error);

      throw error;
    }
  }


  /*
   * ==========================================================
   * GET OFFER DATA
   * ==========================================================
   *
   * Sends a GET request to the SharePoint REST API and
   * returns the Offer list items.
   */
  private async _getOffersData(
    apiUrl: string
  ): Promise<IOffer[]> {

    try {

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

      if (response.ok) {

        const data = await response.json();

        return data.value;

      } else {

        console.error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );

        throw new Error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );
      }

    }
    catch (error) {

      console.log("error occured", error);

      throw error;
    }
  }


  /*
   * ==========================================================
   * FORMAT DATE
   * ==========================================================
   *
   * Converts the SharePoint Created date into:
   *
   *     Sep 21, 2026
   *
   * If the date is empty or invalid, the original value
   * is returned where appropriate.
   */
  private formatDate(
    dateValue: string
  ): string {

    if (!dateValue) {
      return '';
    }

    const date =
      new Date(dateValue);

    if (
      isNaN(
        date.getTime()
      )
    ) {

      return dateValue;
    }

    return date.toLocaleDateString(
      'en-US',
      {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
      }
    );
  }
}