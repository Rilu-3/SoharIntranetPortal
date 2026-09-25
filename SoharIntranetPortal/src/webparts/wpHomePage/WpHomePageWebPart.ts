import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';



import AnnouncementOffer from './AnnouncementOffer';


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
     const baseUrl = this.context.pageContext.web.absoluteUrl;

  const workbenchContent = document.getElementById('workbenchPageContent');

  if (workbenchContent) {
    workbenchContent.style.maxWidth = 'none';
  }

  this.domElement.innerHTML = AnnouncementOffer.allElementsHtml;

  const arrowIconUrl =
    `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;

  this.setupViewAllLink(arrowIconUrl);

  const announcementApiUrl =
    `${baseUrl}/_api/web/lists/GetByTitle('Announcements')/items` +
    `?$select=Id,Title,ShortDescription,Icon,Created,Status` +
    `&$filter=Status eq 'Active'` +
    `&$orderby=Created desc` +
    `&$top=3`;

  const offerApiUrl =
    `${baseUrl}/_api/web/lists/GetByTitle('Offers')/items` +
    `?$select=Id,Title,Description,OfferImage,Created,Status` +
    `&$filter=Status eq 'Active'` +
    `&$orderby=Created desc` +
    `&$top=3`;

  await this._renderAnnouncementsAsync(announcementApiUrl);
  await this._renderOffersAsync(offerApiUrl);
  }



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

    
private setupViewAllLink(arrowIconUrl: string): void {

  const baseUrl = this.context.pageContext.web.absoluteUrl;

  const arrowImage =
    this.domElement.querySelector('#ao-view-all-arrow') as HTMLImageElement;
    if (!arrowImage) {
    console.error('arrow image element not found');
    return;
  }

  // Set arrow image
  arrowImage.src = arrowIconUrl;
  const viewAllLink =
    this.domElement.querySelector('#ao-view-all') as HTMLAnchorElement;

  if (!viewAllLink) {
    console.error('View All elements not found');
    return;
  }
  
  viewAllLink.href =
    `${baseUrl}/SitePages/Announcement.aspx`;

  const tabs =
    this.domElement.querySelectorAll('[data-tab-ao]');



  // Set default URL
  // Change URL when tab changes
  tabs.forEach((tab) => {

    tab.addEventListener('click', () => {

      const selectedTab = tab.getAttribute('data-tab-ao');

      if (selectedTab === 'announcement') {
        viewAllLink.href =
          `${baseUrl}/SitePages/Announcement.aspx`;
      }
      else if (selectedTab === 'offers') {
        viewAllLink.href =
          `${baseUrl}/SitePages/Offer-List.aspx`;
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
 try {
    const data: IAnnouncement[] =
      await this._getAnnouncementsData(apiUrl);

    console.log("Announcements data", data);

    let allElementsHtml: string = "";

   

      data.forEach((item) => {

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
        let singleElementHtml = AnnouncementOffer.singleElementHtml
          .replace("__KEY__ANNOUNCEMENTOFFER__ICON__", imageUrl)
          .replace("__KEY__ANNOUNCEMENTOFFER__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENTOFFER__DESCRIPTION__", item.ShortDescription)
          .replace("__KEY__ANNOUNCEMENTOFFER__DATE__", createddate);

        /*
         * Add the generated Announcement HTML to the
         * complete HTML string.
         */
        allElementsHtml += singleElementHtml;
      })
        /*
     * Insert all generated Announcement HTML into the
     * Announcement container.
     */
this.domElement.querySelector("#announcement-container")!.innerHTML =allElementsHtml;
    }
    catch (error) {
      console.error('Error rendering QuickList:', error);
    }

  
    
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
try{
    const data: IOffer[] =
      await this._getOffersData(apiUrl);

    console.log("Offers data", data);

    let allElementsHtml: string = "";



      data.forEach((item) => {

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
        let singleElementHtml = AnnouncementOffer.singleElementHtml
          .replace("__KEY__ANNOUNCEMENTOFFER__ICON__", imageUrl)
          .replace("__KEY__ANNOUNCEMENTOFFER__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENTOFFER__DESCRIPTION__", item.Description)
          .replace("__KEY__ANNOUNCEMENTOFFER__DATE__", createddate);

        /*
         * Add the generated Offer HTML to the complete
         * HTML string.
         */
        allElementsHtml += singleElementHtml;

      })
          /*
     * Insert all generated Offer HTML into the Offer container.
     */
   this.domElement.querySelector("#offer-container")!.innerHTML =
      allElementsHtml;
    }
    catch (error) {
      console.error('Error rendering QuickList:', error);
    }


 
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