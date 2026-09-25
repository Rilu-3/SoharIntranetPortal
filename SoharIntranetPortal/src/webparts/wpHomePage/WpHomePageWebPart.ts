import { Version } from '@microsoft/sp-core-library';
import Wrapper from './Wrapper';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

// import {
//   BaseClientSideWebPart
// } from '@microsoft/sp-webpart-base';
import {
  BaseClientSideWebPart,
  WebPartContext
} from '@microsoft/sp-webpart-base';
import {
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClientV3
} from '@microsoft/sp-http';

import {
  SPComponentLoader
} from '@microsoft/sp-loader';

import { escape } from '@microsoft/sp-lodash-subset';

import UpcomingEventsTemplate from './UpcomingEvents';
import { BannerTemplate } from './BannerTemplate';
import MediaGalleryTemplate from './MediaGalleryTemplate';

import AnnouncementOffer from './AnnouncementOffer';
import QuickLinks from './QuickLinks';
import Birthday from './Birthday';
import SocialMedia from './SocialMedia';
import NewsCentre from './NewsCentre';



export interface IWpHomePageWebPartProps {
  description: string;
}


interface IBannerItem {
  Id: number;
  Title: string;
  Description: string;
  Status: string;
  SortOrder: number;
  Image: any;
}


interface IMediaGalleryItem {
  Id: number;
  Title: string;
  Caption: string;
  Status: string;
  SortOrder: number;
  Image: any;
}


// Interface used for both Outlook My Events
// and SharePoint Organizational Events
export interface IEventItem {
  Id: number;

  Title: string;

  EventDate: string;

  StartTime: string;

  EndTime: string;

  Location: string;

  Status: string;
}
interface IAnnouncement {
  Id: number;
  Title: string;
  ShortDescription: string;
  Icon: string;
  Created: string;
}



interface IOffer {
  Id: number;
  Title: string;
  Description: string;
  OfferImage: string;
  Created: string;
}
export interface IQuickLinksList {
  Id: number;
  Title: string;
  Icon: string;
  URL: {
    Url: string;
  };
  Status: string;
  SortOrder: number;
}


export interface IBirthdayList {
  Id: number;
  Title: string;
  EmployeePhoto?: any;
  BirthDate: string;
  Status: string;
}

interface INewsItem {
  Id: number;
  Title: string;
  NewsIcon?: any;
  ShortDescription?: string;
  Category?: string;
  MainContent?: string;
  PublishedDate?: string;
  Status?: string;
}

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


    private newsCentreItems: INewsItem[] = [];


  // Stores Organizational Events
  // fetched from the SharePoint Upcoming Events list
  private events: IEventItem[] = [];


  // Stores My Events
  // fetched from the logged-in user's Outlook calendar
  private myEvents: IEventItem[] = [];

   private socialMediaObserver: MutationObserver | null = null;

    private hideSocialMediaTutorialLinks(): void {
  const tutorialLinks =
    document.querySelectorAll('.tutorial_link');

  tutorialLinks.forEach((link) => {
    (link as HTMLElement).style.display = 'none';
  });
}

private setupSocialMediaTutorialLinkObserver(): void {

  // Hide links that already exist
  this.hideSocialMediaTutorialLinks();

  // Watch for SociableKIT to add links dynamically
  this.socialMediaObserver = new MutationObserver(() => {
    this.hideSocialMediaTutorialLinks();
  });

  this.socialMediaObserver.observe(document.body, {
    childList: true,
    subtree: true
  });
}


  // ==================== RENDER ====================

  public async render(): Promise<void> {

    

    const workbenchContent =
      document.getElementById('workbenchPageContent');

    if (workbenchContent) {
      workbenchContent.style.maxWidth = 'none';
    }


    const baseUrl =
      this.context.pageContext.web.absoluteUrl;

       


    /*
     * ==========================================================
     * 2. SET STATIC ELEMENTS
     * ==========================================================
     */

    const newsarrowIconUrl =
      `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;

  


    /*
     * ==========================================================
     * 3. CREATE NEWS API URL
     * ==========================================================
     */

    const newsApiUrl =
      `${baseUrl}/_api/web/lists/getbytitle('News')/items` +
      `?$select=Id,Title,NewsIcon,ShortDescription,Category,MainContent,PublishedDate,Status` +
      `&$filter=Status eq 'Active'` +
      `&$orderby=PublishedDate desc`;


    /*
     * ==========================================================
     * 4. GET NEWS DATA AND RENDER IT
     * ==========================================================
     */

  

    /*
     * ==========================================================
     * 5. ATTACH TAB FUNCTIONALITY
     * ==========================================================
     */

   


    // Get the current date
    const today = new Date();


    // Load both event sources at the same time
    await Promise.all([

      // Load Organizational Events from SharePoint
      this.loadEvents(
        today.getFullYear(),
        today.getMonth()
      ),

      // Load My Events from Outlook Calendar
      this.loadMyEvents(
        today.getFullYear(),
        today.getMonth()
      )

    ]);

    const arrowIconUrl =
    `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    /*
     * Create the complete page HTML first.
     *
     * This is important because Banner,
     * Upcoming Events and Media Gallery
     * must have their own containers.
     */
    
    this.domElement.innerHTML =Wrapper.wrapperHtml;
    this.domElement.querySelector('#banner-container')!.innerHTML = BannerTemplate.bannerHtml;
    this.domElement.querySelector('#announcement-offer-container')!.innerHTML = AnnouncementOffer.allElementsHtml;
    this.domElement.querySelector('#quick-links-container')!.innerHTML = QuickLinks.allElementsHtml;
    this.domElement.querySelector('#news-container')!.innerHTML = NewsCentre.allElementsHtml;
    this.domElement.querySelector('#upcoming-events-container')!.innerHTML = UpcomingEventsTemplate.allElementsHtml;
    this.domElement.querySelector('#social-media-container')!.innerHTML = SocialMedia.allElementsHtml;
     this.domElement.querySelector('#birthday-container')!.innerHTML = Birthday.allElementsHtml;
     this.domElement.querySelector('#media-gallery-container')!.innerHTML = MediaGalleryTemplate.allElementsHtml;
    
      // BannerTemplate.bannerHtml +
      // UpcomingEventsTemplate.allElementsHtml +
      // NewsCentre.allElementsHtml+
      // MediaGalleryTemplate.allElementsHtml +
      // MediaGalleryTemplate.galleryModalHtml +   
      // AnnouncementOffer.allElementsHtml +
      // QuickLinks.allElementsHtml +
      // Birthday.allElementsHtml+
      // SocialMedia.allElementsHtml;

      const birthdayViewAll =
      this.domElement.querySelector(
        '.birthday-view-all'
      ) as HTMLAnchorElement;

      if (birthdayViewAll) {
        birthdayViewAll.href =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/SitePages/Upcoming-Birthdys.aspx`;
      }

    const birthdayArrow =
      this.domElement.querySelector(
        '.birthday-view-all-arrow'
      ) as HTMLImageElement;

    if (birthdayArrow) {
      birthdayArrow.src =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    }

      this.newsCentreSetupViewAll(
      newsarrowIconUrl
    );

    


    // Create Upcoming Events inside its existing containers
    this.renderUpcomingEvents();


    // Initialize My Events / Organizational Events tabs
    this.initializeUpcomingEvents();

    
    // Initialize both calendars
    this.initializeCalendar();




 


    this.setupViewAllLink(arrowIconUrl);

    const AnnouncementApiUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Announcements')/items?$select=Id,Title,ShortDescription,Icon,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc&$top=3`;

    const OfferApiUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Offers')/items?$select=Id,Title,Description,OfferImage,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc&$top=3`;

    await this._renderAnnouncementsAsync(AnnouncementApiUrl);

    await this._renderOffersAsync(OfferApiUrl);


    // Load Banner items from SharePoint
     await this._getBannerItems();


    // Load Active Media Gallery items from SharePoint
     await this._getMediaGalleryItems();

    
    await this.initializeQuickLinks();


    await this.renderBirthdays();

    await this.initializeSocialMedia();

    await this._renderNewsAsync(
      newsApiUrl
    );
   this.newsCentreAttachTabEvents();

    // Load home.js only after the HTML is available
    this.loadHomeJS();

  }


  private newsCentreSetupViewAll(
    arrowIconUrl: string
  ): void {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    /*
     * ----------------------------------------------------------
     * Find View All link by ID.
     * ----------------------------------------------------------
     */

    let viewAllLink =
      this.domElement.querySelector(
        '#news-view-all'
      ) as HTMLAnchorElement | null;


    /*
     * ----------------------------------------------------------
     * Fallback:
     * Your older template may still contain the placeholder
     * directly in the href.
     * ----------------------------------------------------------
     */

    if (!viewAllLink) {

      viewAllLink =
        this.domElement.querySelector(
          'a[href="__KEY_URL_VIEW_ALL__"]'
        ) as HTMLAnchorElement | null;
    }


    /*
     * ----------------------------------------------------------
     * Set View All URL.
     * ----------------------------------------------------------
     */

    if (viewAllLink) {

      viewAllLink.href =
        `${baseUrl}/SitePages/News-List.aspx`;

    }
    else {

      console.warn(
        'News View All link not found.'
      );
    }


    /*
     * ----------------------------------------------------------
     * Find arrow image by ID.
     * ----------------------------------------------------------
     */

    let arrowImage =
      this.domElement.querySelector(
        '#news-view-all-arrow'
      ) as HTMLImageElement | null;


    /*
     * ----------------------------------------------------------
     * Fallback:
     * Older template may still contain the placeholder
     * directly in the image src.
     * ----------------------------------------------------------
     */

    if (!arrowImage) {

      arrowImage =
        this.domElement.querySelector(
          'img[src="__KEY_URL_ARROW__"]'
        ) as HTMLImageElement | null;
    }


    /*
     * ----------------------------------------------------------
     * Set arrow image.
     * ----------------------------------------------------------
     */

    if (arrowImage) {

      arrowImage.src =
        arrowIconUrl;

    }
    else {

      console.warn(
        'News View All arrow image not found.'
      );
    }
  }


  /*
   * ============================================================
   * RENDER NEWS
   * ============================================================
   *
   * Gets News data from SharePoint.
   *
   * After getting the data:
   *
   * 1. Store it in newsCentreItems.
   * 2. Render All News.
   * 3. Render Announcements.
   */

  private async _renderNewsAsync(
    apiUrl: string
  ): Promise<void> {

    try {

      /*
       * ========================================================
       * GET SHAREPOINT DATA
       * ========================================================
       */

      const data: INewsItem[] =
        await this._getNewsData(
          apiUrl
        );


      /*
       * Store the data.
       *
       * This is required later for tab filtering.
       */

      this.newsCentreItems =
        data;


      console.log(
        'News items loaded:',
        this.newsCentreItems
      );


      /*
       * ========================================================
       * DEFAULT TAB — ALL
       * ========================================================
       */

      this.newsCentreRenderCategory(
        'All'
      );


      /*
       * ========================================================
       * ANNOUNCEMENTS TAB
       * ========================================================
       */

      this.newsCentreRenderCategory(
        'Announcements'
      );

    }
    catch (error) {

      console.error(
        'Error rendering News:',
        error
      );
    }
  }


  /*
   * ============================================================
   * RENDER CATEGORY
   * ============================================================
   *
   * THIS IS THE MAIN TEMPLATE RENDERING METHOD.
   *
   * The structure is intentionally the same as your
   * AnnouncementOffer rendering pattern:
   *
   * let singleElementHtml =
   *     NewsCentre.singleElementHtml
   *       .replace(...)
   *       .replace(...);
   *
   * allElementsHtml += singleElementHtml;
   *
   * container.innerHTML = allElementsHtml;
   */

  

  private newsCentreRenderCategory(
    category: string
  ): void {
 
 
    /*
     * ==========================================================
     * 1. DETERMINE THE PANEL
     * ==========================================================
     */
 
    let panel: HTMLElement | null = null;
 
 
    /*
     * All
     */
 
    if (category === 'All') {
 
      panel =
        this.domElement.querySelector(
          '#news-panel-all'
        ) as HTMLElement | null;
    }
 
 
    /*
     * Announcements
     */
 
    else if (
      category === 'Announcements'
    ) {
 
      panel =
        this.domElement.querySelector(
          '#news-panel-announcements'
        ) as HTMLElement | null;
    }
 
 
    /*
     * Events / News / Circulars
     *
     * These reuse the All panel.
     */
 
    else if (
      category === 'Events' ||
      category === 'News' ||
      category === 'Circulars'
    ) {
 
      panel =
        this.domElement.querySelector(
          '#news-panel-all'
        ) as HTMLElement | null;
    }
 
 
    /*
     * If the required panel cannot be found,
     * stop safely.
     */
 
    if (!panel) {
 
      console.error(
        `News panel not found for category: ${category}`
      );
 
      return;
    }
 
 
    /*
     * ==========================================================
     * 2. FIND NEWS CONTAINER
     * ==========================================================
     *
     * This is the container into which the generated
     * single News item HTML will be inserted.
     */
 
    const container =
      panel.querySelector(
        '.panel-card-news'
      ) as HTMLElement | null;
 
 
    if (!container) {
 
      console.error(
        `News container not found for category: ${category}`
      );
 
      return;
    }
 
 
    /*
     * ==========================================================
     * 3. FILTER ITEMS
     * ==========================================================
     */
 
    let items: INewsItem[];
 
 
    if (category === 'All') {
 
      /*
       * All News items.
       */
 
      items =
        this.newsCentreItems;
    }
    else {
 
      /*
       * Only the selected category.
       */
 
      items =
        this.newsCentreItems.filter(
          (item) =>
            this.newsCentreNormalizeValue(
              item.Category
            ) ===
            this.newsCentreNormalizeValue(
              category
            )
        );
    }
 
 
    /*
     * ==========================================================
     * 4. START COMPLETE HTML STRING
     * ==========================================================
     */
 
    let allElementsHtml: string = "";
 
 
    /*
     * ==========================================================
     * 5. LOOP THROUGH SHAREPOINT ITEMS
     * ==========================================================
     */
 
    items.forEach(
      (item) => {
 
 
        /*
         * ======================================================
         * IMAGE URL
         * ======================================================
         *
         * Use the helper method here.
         *
         * This means newsCentreGetImageUrl() is no longer
         * an unused method.
         */
 
        const imageUrl =
          this.newsCentreGetImageUrl(
            item
          );
 
 
        /*
         * ======================================================
         * DATE
         * ======================================================
         */
 
        const createddate =
          this.newsCentreFormatDate(
            item.PublishedDate
          );
 
 
        /*
         * ======================================================
         * DETAILS URL
         * ======================================================
         */
 
        const detailsUrl =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/Lists/News/DispForm.aspx?ID=${item.Id}`;
 
 
        /*
         * ======================================================
         * TAKE SINGLE ELEMENT TEMPLATE
         * ======================================================
         *
         * THIS is the important part.
         */
 
        let singleElementHtml =
          NewsCentre.singleElementHtml;
 
 
        /*
         * ======================================================
         * REPLACE PLACEHOLDERS
         * ======================================================
         */
 
        singleElementHtml =
          singleElementHtml
 
            /*
             * News icon
             */
 
            .replace(
              /__KEY_URL_IMGICON__/g,
              imageUrl
            )
 
            /*
             * News title
             */
 
            .replace(
              /__KEY_DATA_TITLE__/g,
              this.newsCentreEscapeHtml(
                item.Title || ''
              )
            )
 
            /*
             * Short description
             */
 
            .replace(
              /__KEY_DATA_DESCRIPTION__/g,
              this.newsCentreEscapeHtml(
                item.ShortDescription || ''
              )
            )
 
            /*
             * Category
             */
 
            .replace(
              /__KEY_DATA_CATEGORY__/g,
              this.newsCentreEscapeHtml(
                item.Category || ''
              )
            )
 
            /*
             * Published date
             */
 
            .replace(
              /__KEY_DATA_DATE__/g,
              createddate
            )
 
            /*
             * News details link
             */
 
            .replace(
              /__KEY_URL_LINK__/g,
              detailsUrl
            )
 
            /*
             * Arrow image
             */
 
            .replace(
              /__KEY_URL_ARROW__/g,
              this.newsCentreGetArrowImageUrl()
            );
 
 
        /*
         * ======================================================
         * ADD GENERATED ITEM TO COMPLETE HTML
         * ======================================================
         */
 
        allElementsHtml +=
          singleElementHtml;
 
      }
    );
 
 
    /*
     * ==========================================================
     * 6. NO DATA
     * ==========================================================
     */
 
    if (!items.length) {
 
      allElementsHtml =
        NewsCentre.noElementHtml;
    }
 
 
    /*
     * ==========================================================
     * 7. INSERT GENERATED HTML INTO CONTAINER
     * ==========================================================
     *
     * Same basic pattern as AnnouncementOffer.
     */
 
    container.innerHTML =
      allElementsHtml;
  }


  /*
   * ============================================================
   * GET NEWS DATA
   * ============================================================
   *
   * Sends GET request to the SharePoint REST API.
   */

  private async _getNewsData(
    apiUrl: string
  ): Promise<INewsItem[]> {

    try {

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept':
                'application/json;odata=nometadata'
            }
          }
        );


      /*
       * Check HTTP response.
       */

      if (!response.ok) {

        throw new Error(
          `News API failed: ${response.status} ${response.statusText}`
        );
      }


      /*
       * Convert response to JSON.
       */

      const data =
        await response.json();


      /*
       * Return SharePoint items.
       */

      return data.value || [];

    }
    catch (error) {

      console.error(
        'Error loading News list:',
        error
      );

      throw error;
    }
  }


  /*
   * ============================================================
   * NEWS IMAGE URL
   * ============================================================
   *
   * NewsIcon contains JSON information about the uploaded
   * attachment.
   */

  private newsCentreGetImageUrl(
    item: INewsItem
  ): string {

    /*
     * No NewsIcon.
     */

    if (!item.NewsIcon) {

      return '';
    }


    try {

      /*
       * NewsIcon may already be an object or may be
       * returned from SharePoint as a JSON string.
       */

      const imageData =
        typeof item.NewsIcon === 'string'
          ? JSON.parse(item.NewsIcon)
          : item.NewsIcon;


      /*
       * Get file name.
       */

      const fileName =
        imageData?.fileName ||
        imageData?.FileName ||
        '';


      /*
       * No file name.
       */

      if (!fileName) {

        return '';
      }


      /*
       * Build SharePoint attachment URL.
       */

      const webUrl =
        this.context.pageContext.web.absoluteUrl;


      return (
        `${webUrl}/Lists/News/Attachments/` +
        `${item.Id}/${fileName}`
      );

    }
    catch (error) {

      console.error(
        'Error parsing NewsIcon:',
        error
      );

      return '';
    }
  }


  /*
   * ============================================================
   * ARROW IMAGE URL
   * ============================================================
   */

  private newsCentreGetArrowImageUrl(): string {

    const webUrl =
      this.context.pageContext.web.absoluteUrl;


    return (
      `${webUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`
    );
  }


  /*
   * ============================================================
   * TAB FUNCTIONALITY
   * ============================================================
   */

  private newsCentreAttachTabEvents(): void {


    /*
     * Find all News tabs.
     */

    const tabs =
      this.domElement.querySelectorAll(
        '#news-tabs .tab-title-pill'
      );


    /*
     * Find All panel.
     */

    const allPanel =
      this.domElement.querySelector(
        '#news-panel-all'
      ) as HTMLElement | null;


    /*
     * Find Announcements panel.
     */

    const announcementPanel =
      this.domElement.querySelector(
        '#news-panel-announcements'
      ) as HTMLElement | null;


    /*
     * No tabs found.
     */

    if (!tabs.length) {

      console.warn(
        'News tabs not found.'
      );

      return;
    }


    /*
     * ==========================================================
     * ADD CLICK EVENT
     * ==========================================================
     */

    tabs.forEach(
      (tab) => {

        tab.addEventListener(
          'click',
          () => {


            /*
             * Get the target panel ID.
             */

            const targetId =
              tab.getAttribute(
                'data-tab-news-id'
              );


            if (!targetId) {

              return;
            }


            /*
             * ==================================================
             * REMOVE ACTIVE CLASS
             * ==================================================
             */

            tabs.forEach(
              (otherTab) => {

                otherTab.classList.remove(
                  'tab-title-pill-active'
                );
              }
            );


            /*
             * Add active class to clicked tab.
             */

            tab.classList.add(
              'tab-title-pill-active'
            );


            /*
             * ==================================================
             * ALL TAB
             * ==================================================
             */

            if (
              targetId ===
              'news-panel-all'
            ) {

              if (allPanel) {

                allPanel.style.display =
                  'block';
              }


              if (announcementPanel) {

                announcementPanel.style.display =
                  'none';
              }


              /*
               * Generate All News HTML.
               */

              this.newsCentreRenderCategory(
                'All'
              );


              return;
            }


            /*
             * ==================================================
             * ANNOUNCEMENTS TAB
             * ==================================================
             */

            if (
              targetId ===
              'news-panel-announcements'
            ) {

              if (allPanel) {

                allPanel.style.display =
                  'none';
              }


              if (announcementPanel) {

                announcementPanel.style.display =
                  'block';
              }


              /*
               * Generate Announcements HTML.
               */

              this.newsCentreRenderCategory(
                'Announcements'
              );


              return;
            }


            /*
             * ==================================================
             * EVENTS / NEWS / CIRCULARS
             * ==================================================
             *
             * Reuse the All panel.
             */

            if (allPanel) {

              allPanel.style.display =
                'block';
            }


            if (announcementPanel) {

              announcementPanel.style.display =
                'none';
            }


            /*
             * Determine selected category.
             */

            let category = '';


            if (
              targetId ===
              'news-panel-events'
            ) {

              category =
                'Events';
            }


            if (
              targetId ===
              'news-panel-news'
            ) {

              category =
                'News';
            }


            if (
              targetId ===
              'news-panel-circulars'
            ) {

              category =
                'Circulars';
            }


            /*
             * Generate HTML using
             * NewsCentre.singleElementHtml.
             */

            this.newsCentreRenderCategory(
              category
            );
          }
        );
      }
    );


    /*
     * ==========================================================
     * DEFAULT TAB STATE
     * ==========================================================
     */

    if (allPanel) {

      allPanel.style.display =
        'block';
    }


    if (announcementPanel) {

      announcementPanel.style.display =
        'none';
    }
  }


  /*
   * ============================================================
   * DATE FORMAT
   * ============================================================
   *
   * Example:
   *
   * Sep 25, 2026
   */

  private newsCentreFormatDate(
    value?: string
  ): string {

    if (!value) {

      return '';
    }


    const date =
      new Date(value);


    if (
      isNaN(
        date.getTime()
      )
    ) {

      return '';
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


  /*
   * ============================================================
   * NORMALIZE VALUE
   * ============================================================
   *
   * Used when comparing categories.
   */

  private newsCentreNormalizeValue(
    value?: string
  ): string {

    return (
      value || ''
    )
      .trim()
      .toLowerCase();
  }


  /*
   * ============================================================
   * ESCAPE HTML
   * ============================================================
   */

  private newsCentreEscapeHtml(
    value: string
  ): string {

    return value
      .replace(
        /&/g,
        '&amp;'
      )
      .replace(
        /</g,
        '&lt;'
      )
      .replace(
        />/g,
        '&gt;'
      )
      .replace(
        /"/g,
        '&quot;'
      )
      .replace(
        /'/g,
        '&#039;'
      );
  }


  // ==================== LOAD CSS ====================

  private loadCSS(): void {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/font-size.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );


    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`
    );

  }


  // ==================== LOAD JS ====================

  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );


    await this.loadBootstrap();


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/common.js`
    );

  }


  // ==================== LOAD BOOTSTRAP ====================

  private async loadBootstrap(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    if (
      typeof (window as any).bootstrap !== 'undefined'
    ) {

      return;

    }


    const bootstrapModule =
      await SPComponentLoader.loadScript<any>(
        `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
      );


    if (bootstrapModule) {

      (window as any).bootstrap =
        bootstrapModule;

    }


    console.log(
      'Bootstrap loaded:',
      typeof (window as any).bootstrap
    );

  }


  // ==================== LOAD HOME JS ====================

  private async loadHomeJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    try {

      await SPComponentLoader.loadScript(
        `${baseUrl}/SiteAssets/resources/js/home.js`
      );


      console.log(
        'home.js loaded'
      );

    } catch (error) {

      console.error(
        'Error loading home.js:',
        error
      );

    }

  }


  // ==================== LOAD EVENTS ====================

  private async loadEvents(
    year: number,
    month: number
  ): Promise<void> {

    // Get the current SharePoint site URL
    const siteUrl =
      this.context.pageContext.web.absoluteUrl;


    /*
     * First day of selected month
     */
    const startDate =
      new Date(
        year,
        month,
        1
      );


    /*
     * First day of next month
     */
    const endDate =
      new Date(
        year,
        month + 1,
        1
      );


    // Convert the dates into ISO format
    // for the SharePoint REST API
    const startDateString =
      startDate.toISOString();


    const endDateString =
      endDate.toISOString();


    // SharePoint REST API URL
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Upcoming Events')/items` +
      `?$select=Id,Title,EventDate,StartTime,EndTime,Location,Status` +
      `&$filter=EventDate ge datetime'${startDateString}' and EventDate lt datetime'${endDateString}'` +
      `&$orderby=EventDate asc`;


    try {

      // Send GET request to the SharePoint REST API
      const response:
        SPHttpClientResponse =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept:
                'application/json;odata=nometadata'
            }
          }
        );


      // Check whether the API request was successful
      if (!response.ok) {

        console.error(
          'Upcoming Events list error:',
          response.status,
          response.statusText
        );


        // Clear events if the request fails
        this.events = [];


        return;

      }


      // Convert API response into JSON
      const data =
        await response.json();


      // Store SharePoint events
      this.events =
        data.value || [];


      // Display the fetched events in the console
      console.log(
        'Organizational Events:',
        this.events
      );


    } catch (error) {

      // Handle SharePoint API errors
      console.error(
        'Error loading Upcoming Events:',
        error
      );


      // Clear events when an error occurs
      this.events = [];

    }

  }


  // ==================== LOAD MY EVENTS ====================

  private async loadMyEvents(
    year: number,
    month: number
  ): Promise<void> {

    try {

      // Create Microsoft Graph client
      const client:
        MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient('3');


      // First day of selected month
      const startDate =
        new Date(
          year,
          month,
          1,
          0,
          0,
          0
        );


      // First day of next month
      const endDate =
        new Date(
          year,
          month + 1,
          1,
          0,
          0,
          0
        );


      // Call Microsoft Graph Calendar API
      const response =
        await client
          .api('/me/calendar/calendarView')

          .query({
            startDateTime:
              startDate.toISOString(),

            endDateTime:
              endDate.toISOString()
          })

          .select(
            'id,subject,start,end,location'
          )

          .orderby(
            'start/dateTime'
          )

          .get();


      // Display Outlook events in the browser console
      console.log(
        'Outlook Calendar Events:',
        response.value
      );


      // Convert Outlook events into the
      // common IEventItem format
      this.myEvents =
        (response.value || []).map(
          (
            event: any,
            index: number
          ): IEventItem => {

            return {

              Id:
                index + 1,

              Title:
                event.subject || '',

              EventDate:
                event.start?.dateTime || '',

              StartTime:
                event.start?.dateTime || '',

              EndTime:
                event.end?.dateTime || '',

              Location:
                event.location?.displayName || '',

              Status:
                'Active'

            };

          }
        );


    } catch (error) {

      // Handle Microsoft Graph errors
      console.error(
        'Error loading Outlook Calendar events:',
        error
      );


      // Clear Outlook events if an error occurs
      this.myEvents = [];

    }

  }


  // ==================== RENDER UPCOMING EVENTS ====================

  private renderUpcomingEvents(): void {

    // Get the SharePoint site URL
    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    // Image used for the event arrow
    const rightArrow =
      `${baseUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


    // Image used for the short arrow
    const arrowRightShort =
      `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;


    // Get future My Events
    const myEvents =
      this.getMyEvents();


    // Get future Organizational Events
    const organizationalEvents =
      this.getOrganizationalEvents();


    // Generate HTML for My Events
    const myEventsHtml =
      this.renderEventElements(
        myEvents,
        rightArrow
      );


    // Generate HTML for Organizational Events
    const organizationalEventsHtml =
      this.renderEventElements(
        organizationalEvents,
        rightArrow
      );


    /*
     * Insert My Events into the existing
     * Upcoming Events container.
     */
    const myEventsList =
      this.domElement.querySelector(
        '#events-list-my'
      );


    if (myEventsList) {

      myEventsList.innerHTML =
        myEventsHtml;

    }


    /*
     * Insert Organizational Events into
     * the existing Upcoming Events container.
     */
    const organizationalEventsList =
      this.domElement.querySelector(
        '#events-list-org'
      );


    if (organizationalEventsList) {

      organizationalEventsList.innerHTML =
        organizationalEventsHtml;

    }


    /*
     * Replace arrow placeholder without
     * replacing the complete page HTML.
     */
    const eventCalendarViews =
      this.domElement.querySelectorAll(
        '.event-calendar-view'
      );


    eventCalendarViews.forEach(
      (view: Element) => {

        view.innerHTML =
          view.innerHTML.replace(
            /__KEY_ARROW_RIGHT_SHORT__/g,
            arrowRightShort
          );

      }
    );

  }


  // ==================== GET MY EVENTS ====================

  private getMyEvents(): IEventItem[] {

    // Get today's date
    const today = new Date();


    // Remove the current time
    today.setHours(
      0,
      0,
      0,
      0
    );


    // Filter Outlook events
    return this.myEvents.filter(
      (event: IEventItem) => {

        // Convert event date into JavaScript Date
        const eventDate =
          new Date(event.EventDate);


        // Remove the time from the event date
        eventDate.setHours(
          0,
          0,
          0,
          0
        );


        // Keep today's and future events
        return eventDate >= today;

      }
    );

  }


  // ==================== GET ORGANIZATIONAL EVENTS ====================

  private getOrganizationalEvents(): IEventItem[] {

    // Get today's date
    const today = new Date();


    // Remove the current time
    today.setHours(
      0,
      0,
      0,
      0
    );


    // Filter SharePoint events
    return this.events.filter(
      (event: IEventItem) => {

        // Ignore events that are not Active
        if (event.Status !== 'Active') {
          return false;
        }


        // Convert event date into JavaScript Date
        const eventDate =
          new Date(event.EventDate);


        // Remove the time from the event date
        eventDate.setHours(
          0,
          0,
          0,
          0
        );


        // Keep today's and future events
        return eventDate >= today;

      }
    );

  }


  // ==================== RENDER EVENT ELEMENTS ====================

  private renderEventElements(
    events: IEventItem[],
    rightArrow: string
  ): string {

    /*
     * No records.
     */
    if (!events.length) {

      return UpcomingEventsTemplate.noRecord;

    }


    /*
     * Show maximum 2 events.
     */
    return events
      .slice(0, 2)
      .map(
        (event: IEventItem) => {

          // Format the event date
          const date =
            this.formatDate(
              event.EventDate
            );


          // Format the event time
          const time =
            this.formatTime(
              event.StartTime,
              event.EndTime
            );


          // Get the individual event HTML template
          let html =
            UpcomingEventsTemplate.singleElementHtml;


          // Replace event month
          html =
            html.replace(
              '__KEY_EVENT_MONTH__',
              escape(date.month)
            );


          // Replace event day
          html =
            html.replace(
              '__KEY_EVENT_DAY__',
              escape(date.day)
            );


          // Replace event title
          html =
            html.replace(
              '__KEY_EVENT_TITLE__',
              escape(event.Title || '')
            );


          // Replace event time
          html =
            html.replace(
              '__KEY_EVENT_TIME__',
              escape(time)
            );


          // Replace event location
          html =
            html.replace(
              '__KEY_EVENT_LOCATION__',
              escape(event.Location || '')
            );


          // Replace event arrow image
          html =
            html.replace(
              '__KEY_EVENT_ARROW__',
              rightArrow
            );


          return html;

        }
      )
      .join('');

  }


  // ==================== FORMAT DATE ====================

  private formatDate(
    eventDate: string
  ): {
    month: string;
    day: string;
  } {

    // Convert string into Date object
    const date =
      new Date(eventDate);


    // Check whether the date is valid
    if (isNaN(date.getTime())) {

      return {
        month: '',
        day: ''
      };

    }


    return {

      // Get short month name
      month:
        date
          .toLocaleString(
            'en-US',
            {
              month: 'short'
            }
          )
          .toUpperCase(),


      // Get day number
      day:
        date
          .getDate()
          .toString()

    };

  }


  // ==================== FORMAT TIME ====================

  private formatTime(
    startTime: string,
    endTime: string
  ): string {

    // Convert start time
    const start =
      this.parseSharePointTime(
        startTime
      );


    // Convert end time
    const end =
      this.parseSharePointTime(
        endTime
      );


    // If both times are empty
    if (!start && !end) {

      return '';

    }


    // If only start time exists
    if (!end) {

      return start;

    }


    // Display start and end time together
    return `${start} – ${end}`;

  }


  // ==================== PARSE TIME ====================

  private parseSharePointTime(
    timeValue: string
  ): string {

    // Return empty value if no time is provided
    if (!timeValue) {

      return '';

    }


    // Check whether the value is an ISO date/time
    if (timeValue.indexOf('T') !== -1) {

      // Convert ISO value into Date object
      const date =
        new Date(timeValue);


      // Check whether the date is valid
      if (!isNaN(date.getTime())) {

        // Convert into 12-hour time
        return date.toLocaleTimeString(
          'en-US',
          {
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
          }
        );

      }

    }


    /*
     * Handle HH:mm or HH:mm:ss.
     */
    const match =
      timeValue.match(
        /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
      );


    if (match) {

      // Extract hour
      const hours =
        parseInt(
          match[1],
          10
        );


      // Extract minutes
      const minutes =
        parseInt(
          match[2],
          10
        );


      // Validate hour and minute values
      if (
        hours >= 0 &&
        hours <= 23 &&
        minutes >= 0 &&
        minutes <= 59
      ) {

        // Decide AM or PM
        const period =
          hours >= 12
            ? 'PM'
            : 'AM';


        // Convert 24-hour hour into 12-hour hour
        const displayHour =
          hours % 12 === 0
            ? 12
            : hours % 12;


        // Return formatted time
        return (
          `${('0' + displayHour).slice(-2)}:` +
          `${('0' + minutes).slice(-2)} ` +
          `${period}`
        );

      }

    }


    // Return original value if it cannot be parsed
    return timeValue;

  }


  // ==================== INITIALIZE TABS ====================

  private initializeUpcomingEvents(): void {

    // Find all event tabs
    const tabs =
      this.domElement.querySelectorAll(
        '.events-tabs-list .etab'
      );


    // Find all event panels
    const panels =
      this.domElement.querySelectorAll(
        '.event-calendar-view'
      );


    // Add click event to each tab
    tabs.forEach(
      (tab: Element) => {

        tab.addEventListener(
          'click',
          () => {

            // Get the panel ID connected to the clicked tab
            const targetId =
              tab.getAttribute(
                'data-tab-event-id'
              );


            /*
             * Remove active class
             * from all tabs.
             */
            tabs.forEach(
              (item: Element) => {

                item.classList.remove(
                  'etab-active'
                );

              }
            );


            panels.forEach(
              (panel: Element) => {

                (
                  panel as HTMLElement
                ).style.display = 'none';

              }
            );


            /*
             * Activate selected tab.
             */
            tab.classList.add(
              'etab-active'
            );


            /*
             * Show selected panel.
             */
            if (targetId) {

              // Find the selected panel
              const selectedPanel =
                this.domElement.querySelector(
                  `#${targetId}`
                );


              if (selectedPanel) {

                // Display the selected panel
                (
                  selectedPanel as HTMLElement
                ).style.display = 'block';

              }

            }

          }
        );

      }
    );

  }


  // ==================== INITIALIZE CALENDAR ====================

  private initializeCalendar(): void {

    // Get jQuery from the global window object
    const $ =
      (window as any).jQuery;


    // Check whether jQuery UI Datepicker is available
    if (
      !$ ||
      !$.fn ||
      !$.fn.datepicker
    ) {

      console.warn(
        'jQuery UI Datepicker is not available.'
      );


      return;

    }


    // Store current web part instance
    // so it can be accessed inside callback functions
    const self = this;


    /*
     * My Events
     */
    $('#events-calendar-my').datepicker({

      // Date display format
      dateFormat: 'dd M yy',


      // Runs when the user changes the calendar month
      onChangeMonthYear:
        async function (
          year: number,
          month: number
        ): Promise<void> {

          // Reload Outlook events for selected month
          await self.loadMyEvents(
            year,
            month - 1
          );


          // Get event arrow image
          const rightArrow =
            `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


          // Rebuild My Events HTML
          const myEventsHtml =
            self.renderEventElements(
              self.getMyEvents(),
              rightArrow
            );


          // Find My Events list container
          const myEventsList =
            self.domElement.querySelector(
              '#events-list-my'
            );


          // Replace old events with new events
          if (myEventsList) {

            myEventsList.innerHTML =
              myEventsHtml;

          }

        }

    });


    /*
     * Organizational Events
     */
    $('#events-calendar-org').datepicker({

      // Date display format
      dateFormat: 'dd M yy',


      // Runs when the user changes the calendar month
      onChangeMonthYear:
        async function (
          year: number,
          month: number
        ): Promise<void> {

          // Reload SharePoint events for selected month
          await self.loadEvents(
            year,
            month - 1
          );


          // Get event arrow image
          const rightArrow =
            `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


          // Rebuild Organizational Events HTML
          const organizationalEventsHtml =
            self.renderEventElements(
              self.getOrganizationalEvents(),
              rightArrow
            );


          // Find Organizational Events list container
          const organizationalEventsList =
            self.domElement.querySelector(
              '#events-list-org'
            );


          // Replace old events with new events
          if (organizationalEventsList) {

            organizationalEventsList.innerHTML =
              organizationalEventsHtml;

          }

        }

    });

  }


  // ==================== GET BANNER ITEMS ====================

  private async _getBannerItems(): Promise<void> {

    try {

      const siteUrl =
        this.context.pageContext.web.absoluteUrl;


      const url =
        `${siteUrl}/_api/web/lists/getbytitle('Banner')/items?$select=Id,Title,Description,Status,SortOrder,Image&$orderby=SortOrder asc`;


      const response =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept':
                'application/json;odata=nometadata'
            }
          }
        );


      if (!response.ok) {

        throw new Error(
          `Banner list request failed: ${response.status}`
        );

      }


      const data =
        await response.json();


      const bannerItems: IBannerItem[] =
        data.value
          .filter(
            (item: IBannerItem) =>
              item.Status === 'Active'
          )
          .sort(
            (a: IBannerItem, b: IBannerItem) =>
              a.SortOrder - b.SortOrder
          );


      console.log(
        'Banner Items:',
        bannerItems
      );


      this._renderBanner(
        bannerItems
      );


    } catch (error) {

      console.error(
        'Error loading Banner list:',
        error
      );


      const divBanner =
        this.domElement.querySelector(
          '#divBanner'
        );


      if (divBanner !== null) {

        divBanner.innerHTML =
          BannerTemplate.noRecord;

      }

    }

  }


  // ==================== RENDER BANNER ====================

  private _renderBanner(
    bannerItems: IBannerItem[]
  ): void {

    let allElementsHtml: string = '';


    bannerItems.forEach(
      (item: IBannerItem) => {

        let imageData: any = {};


        if (item.Image) {

          imageData =
            typeof item.Image === 'string'
              ? JSON.parse(item.Image)
              : item.Image;

        }


        const fileName =
          imageData.fileName || '';


        const imageUrl =
          `${this.context.pageContext.web.absoluteUrl}/Lists/Banner/Attachments/${item.Id}/${fileName}`;


        const singleElementHtml =
          BannerTemplate.singleElementHtml
            .replace(
              '__KEY_BANNER_IMAGE__',
              imageUrl
            )
            .replace(
              /__KEY_BANNER_TITLE__/g,
              item.Title || ''
            )
            .replace(
              '__KEY_BANNER_DESCRIPTION__',
              item.Description || ''
            );


        allElementsHtml +=
          singleElementHtml;

      }
    );


    if (allElementsHtml === '') {

      allElementsHtml =
        BannerTemplate.noRecord;

    }


    const divBanner =
      this.domElement.querySelector(
        '#divBanner'
      );


    if (divBanner !== null) {

      divBanner.innerHTML =
        allElementsHtml;

    }

  }


  // ==================== GET MEDIA GALLERY ITEMS ====================

  private async _getMediaGalleryItems(): Promise<void> {

    try {

      const siteUrl =
        this.context.pageContext.web.absoluteUrl;


      const url =
        `${siteUrl}/_api/web/lists/getbytitle('Media_Gallery')/items?$select=Id,Title,Caption,Status,SortOrder,Image&$orderby=SortOrder asc`;


      const response =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept':
                'application/json;odata=nometadata'
            }
          }
        );


      if (!response.ok) {

        throw new Error(
          `Media Gallery list request failed: ${response.status}`
        );

      }


      const data =
        await response.json();


      const galleryItems: IMediaGalleryItem[] =
        data.value
          .filter(
            (item: IMediaGalleryItem) =>
              item.Status === 'Active'
          )
          .sort(
            (a: IMediaGalleryItem, b: IMediaGalleryItem) =>
              a.SortOrder - b.SortOrder
          );


      console.log(
        'Media Gallery Items:',
        galleryItems
      );


      this._renderMediaGallery(
        galleryItems
      );


    } catch (error) {

      console.error(
        'Error loading Media Gallery list:',
        error
      );


      const galleryWrapper =
        this.domElement.querySelector(
          '.gallery-swiper .swiper-wrapper'
        );


      if (galleryWrapper !== null) {

        galleryWrapper.innerHTML =
          MediaGalleryTemplate.noRecord;

      }

    }

  }


  // ==================== RENDER MEDIA GALLERY ====================

  private _renderMediaGallery(
    galleryItems: IMediaGalleryItem[]
  ): void {

    let allElementsHtml: string = '';


    galleryItems.forEach(
      (item: IMediaGalleryItem) => {

        let imageData: any = {};


        if (item.Image) {

          imageData =
            typeof item.Image === 'string'
              ? JSON.parse(item.Image)
              : item.Image;

        }


        const fileName =
          imageData.fileName || '';


        const imageUrl =
          `${this.context.pageContext.web.absoluteUrl}/Lists/Media_Gallery/Attachments/${item.Id}/${fileName}`;


        console.log(
          'Gallery Image URL:',
          imageUrl
        );


        const singleElementHtml =
          MediaGalleryTemplate.singleElementHtml
            .replace(
              '__KEY_GALLERY_IMAGE__',
              imageUrl
            )
            .replace(
              '__KEY_GALLERY_CAPTION__',
              item.Caption || item.Title || ''
            );


        allElementsHtml +=
          singleElementHtml;

      }
    );


    if (allElementsHtml === '') {

      allElementsHtml =
        MediaGalleryTemplate.noRecord;

    }


    const galleryWrapper =
      this.domElement.querySelector(
        '.gallery-swiper .swiper-wrapper'
      );


    if (galleryWrapper !== null) {

      galleryWrapper.innerHTML =
        allElementsHtml;


      /*
       * Fill Modal Gallery
       */
      const modalWrapper =
        this.domElement.querySelector(
          '.gallery-modal-swiper .swiper-wrapper'
        );


      if (modalWrapper !== null) {

        let modalElementsHtml: string = '';


        galleryItems.forEach(
          (item: IMediaGalleryItem) => {

            let imageData: any = {};


            if (item.Image) {

              imageData =
                typeof item.Image === 'string'
                  ? JSON.parse(item.Image)
                  : item.Image;

            }


            const fileName =
              imageData.fileName || '';


            const imageUrl =
              `${this.context.pageContext.web.absoluteUrl}/Lists/Media_Gallery/Attachments/${item.Id}/${fileName}`;


            modalElementsHtml += `
              <div class="swiper-slide gallery-swiper-slide">
                <img
                  src="${imageUrl}"
                  alt="${item.Caption || item.Title || ''}"
                />
              </div>
            `;

          }
        );


        modalWrapper.innerHTML =
          modalElementsHtml;

      }


      const galleryElement =
        this.domElement.querySelector(
          '.gallery-swiper'
        ) as HTMLElement & {
          swiper?: any;
        };


      if (
        galleryElement &&
        galleryElement.swiper
      ) {

        galleryElement.swiper.update();

      }

    }

  }
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
        let createddate = this.formatDates(item.Created);
 
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
        let createddate = this.formatDates(item.Created);
 
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
  private formatDates(
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

// =========================================================
  // QUICK LINKS
  // =========================================================


  // =========================================================
  // Initialize Quick Links
  // =========================================================

  private async initializeQuickLinks(): Promise<void> {

    try {

      const quickLinks =
        await this.getQuickLinks(
          this.context
        );


      const favouriteIds =
        await this.getUserFavouriteIds(
          this.context
        );


      this.renderQuickLinks(
        this.context,
        this.domElement,
        quickLinks,
        favouriteIds
      );


      this.renderFavouriteOptions(
        this.domElement,
        quickLinks,
        favouriteIds
      );


      this.setupTabs(
        this.domElement
      );


      this.setupAddFavouriteButton(
        this.context,
        this.domElement
      );


      this.setupFavouriteModal(
        this.context,
        this.domElement
      );


      this.updateFavouriteTabVisibility(
        this.domElement,
        favouriteIds
      );

    } catch (error) {

      console.error(
        'Error initializing Quick Links:',
        error
      );


      const container =
        this.domElement.querySelector(
          '.quick-links-grid'
        );


      if (container) {

        container.innerHTML = `
          <div class="p-3">
            Unable to load Quick Links.
          </div>
        `;

      }

    }

  }


  // =========================================================
  // Get Quick Links
  // =========================================================

  private async getQuickLinks(
    context: WebPartContext
  ): Promise<IQuickLinksList[]> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const apiUrl =
      `${webUrl}` +
      `/_api/web/lists/GetByTitle('Quick_Links')/items` +
      `?$select=Id,Title,Icon,URL,Status,SortOrder` +
      `&$orderby=SortOrder asc`;


    const response =
      await context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {

      console.error(
        `Quick Links request failed: ${response.status}`
      );

      return [];

    }


    const data =
      await response.json();


    return data.value;

  }


  // =========================================================
  // Get Current User
  // =========================================================

  private async getCurrentUser(
    context: WebPartContext
  ): Promise<any> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const response =
      await context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {

      throw new Error(
        `Unable to get current user: ${response.status}`
      );

    }


    return await response.json();

  }


  // =========================================================
  // Get User Favourite IDs
  // =========================================================

  private async getUserFavouriteIds(
    context: WebPartContext
  ): Promise<number[]> {

    try {

      const webUrl =
        context.pageContext.web.absoluteUrl;


      const currentUser =
        await this.getCurrentUser(
          context
        );


      const response =
        await context.spHttpClient.get(

          `${webUrl}` +
          `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items` +
          `?$select=Id,Quick_x0020_LinkId` +
          `&$filter=UserId eq ${currentUser.Id}`,

          SPHttpClient.configurations.v1

        );


      if (!response.ok) {

        console.error(
          'Unable to get user favourites:',
          response.status
        );

        return [];

      }


      const data =
        await response.json();


      return data.value.map(
        (item: { Quick_x0020_LinkId: number }) =>
          Number(
            item.Quick_x0020_LinkId
          )
      );

    } catch (error) {

      console.error(
        'Error loading user favourites:',
        error
      );

      return [];

    }

  }


  // =========================================================
  // Get Quick Link Image URL
  // =========================================================

  private getQuickLinkImageUrl(
    context: WebPartContext,
    item: IQuickLinksList
  ): string {

    if (!item.Icon) {

      return '';

    }


    try {

      const imgData =
        JSON.parse(
          item.Icon
        );


      return (
        `${context.pageContext.web.absoluteUrl}` +
        `/Lists/Quick_Links/Attachments/` +
        `${item.Id}/` +
        `${imgData.fileName}`
      );

    } catch (error) {

      console.error(
        'Error parsing Quick Link Icon:',
        error
      );

      return '';

    }

  }


  // =========================================================
  // Render Quick Links
  // =========================================================

  private renderQuickLinks(
    context: WebPartContext,
    domElement: HTMLElement,
    items: IQuickLinksList[],
    favouriteIds: number[]
  ): void {

    const container =
      domElement.querySelector(
        '.quick-links-grid'
      );


    if (!container) {

      console.error(
        'Quick Links container not found.'
      );

      return;

    }


    let allElementsHtml = '';


    const activeItems =
      items.filter(
        (item) =>
          item.Status === 'Active'
      );


    activeItems.forEach(
      (item) => {

        const imageUrl =
          this.getQuickLinkImageUrl(
            context,
            item
          );


        const isFavourite =
          favouriteIds.indexOf(
            item.Id
          ) !== -1;


        const favouriteAttribute =
          isFavourite
            ? 'data-tab-cat="ql-favourite"'
            : '';


        const singleElementHtml =
          QuickLinks.singleElementHtml

            .replace(
              /__KEY_URL_IMGICON__/,
              imageUrl
            )

            .replace(
              /__KEY_DATA_TITLE__/g,
              item.Title || ''
            )

            .replace(
              /__KEY_URL_LINK__/,
              item.URL?.Url || '#'
            )

            // .replace(
            //   /__KEY_URL_TARGET__/,
            //   '_blank'
            // )

            .replace(
              /__KEY_DATA_FAVOURITE__/,
              favouriteAttribute
            );


        allElementsHtml +=
          singleElementHtml;

      }
    );


    container.innerHTML =
      allElementsHtml;

  }


  // =========================================================
  // Render Favourite Options
  // =========================================================

  private renderFavouriteOptions(
    domElement: HTMLElement,
    items: IQuickLinksList[],
    favouriteIds: number[]
  ): void {

    const container =
      domElement.querySelector(
        '.favourite-options'
      );


    if (!container) {

      console.error(
        'Favourite options container not found.'
      );

      return;

    }


    const activeItems =
      items.filter(
        (item) =>
          item.Status === 'Active'
      );


    let optionsHtml = '';


    activeItems.forEach(
      (item) => {

        const isFavourite =
          favouriteIds.indexOf(
            item.Id
          ) !== -1;


        optionsHtml += `

          <label class="favourite-option">

            <input
              type="checkbox"
              value="${item.Id}"
              ${isFavourite ? 'checked' : ''}>

            <span>
              ${item.Title}
            </span>

          </label>

        `;

      }
    );


    container.innerHTML =
      optionsHtml;

  }


  // =========================================================
  // Setup Tabs
  // =========================================================

  private setupTabs(
    domElement: HTMLElement
  ): void {

    const tabs =
      domElement.querySelectorAll(
        '[data-filter-ql]'
      );


    tabs.forEach(
      (tab) => {

        tab.addEventListener(
          'click',
          () => {

            const selectedTab =
              (tab as HTMLElement)
                .getAttribute(
                  'data-filter-ql'
                );


            const allLinks =
              domElement.querySelectorAll(
                '.quick-links-grid .quick-link-box'
              );


            tabs.forEach(
              (item) => {

                item.classList.remove(
                  'panel-title-filter-active'
                );

              }
            );


            tab.classList.add(
              'panel-title-filter-active'
            );


            allLinks.forEach(
              (link) => {

                const favouriteCategory =
                  link.getAttribute(
                    'data-tab-cat'
                  );


                if (
                  selectedTab ===
                  'favourites'
                ) {

                  if (
                    favouriteCategory ===
                    'ql-favourite'
                  ) {

                    (
                      link as HTMLElement
                    ).style.display = '';

                  } else {

                    (
                      link as HTMLElement
                    ).style.display = 'none';

                  }

                } else {

                  (
                    link as HTMLElement
                  ).style.display = '';

                }

              }
            );

          }
        );

      }
    );

  }


  // =========================================================
  // Show / Hide Favourites Tab
  // =========================================================

  private updateFavouriteTabVisibility(
    domElement: HTMLElement,
    favouriteIds: number[]
  ): void {

    const favouriteTab =
      domElement.querySelector(
        '[data-filter-ql="favourites"]'
      ) as HTMLElement;


    if (!favouriteTab) {

      return;

    }


    if (favouriteIds.length === 0) {

      favouriteTab.style.display =
        'none';

    } else {

      favouriteTab.style.display =
        '';

    }

  }


  // =========================================================
  // Setup Add Favourite Button
  // =========================================================

  private setupAddFavouriteButton(
    context: WebPartContext,
    domElement: HTMLElement
  ): void {

    const button =
      domElement.querySelector(
        '#btnAddFavourites'
      );


    if (!button) {

      console.error(
        'Add Favourites button not found.'
      );

      return;

    }


    button.addEventListener(
      'click',
      async () => {

        await this.addFavourites(
          context,
          domElement
        );

      }
    );

  }


  // =========================================================
  // Setup Favourite Modal
  // =========================================================

  private setupFavouriteModal(
    context: WebPartContext,
    domElement: HTMLElement
  ): void {

    const modal =
      domElement.querySelector(
        '#addFavouriteModal'
      );


    if (!modal) {

      console.error(
        'Favourite modal not found.'
      );

      return;

    }


    modal.addEventListener(
      'show.bs.modal',
      async () => {

        const favouriteIds =
          await this.getUserFavouriteIds(
            context
          );


        const checkboxes =
          domElement.querySelectorAll(
            '.favourite-options input[type="checkbox"]'
          );


        checkboxes.forEach(
          (checkbox) => {

            const input =
              checkbox as HTMLInputElement;


            const quickLinkId =
              Number(
                input.value
              );


            input.checked =
              favouriteIds.indexOf(
                quickLinkId
              ) !== -1;

          }
        );

      }
    );

  }


  // =========================================================
  // Add / Remove Favourites
  // =========================================================

  private async addFavourites(
    context: WebPartContext,
    domElement: HTMLElement
  ): Promise<void> {

    try {

      const checkboxes =
        domElement.querySelectorAll(
          '.favourite-options input[type="checkbox"]'
        );


      const selectedIds: number[] = [];


      checkboxes.forEach(
        (checkbox) => {

          const input =
            checkbox as HTMLInputElement;


          if (input.checked) {

            selectedIds.push(
              Number(
                input.value
              )
            );

          }

        }
      );


      const webUrl =
        context.pageContext.web.absoluteUrl;


      const currentUser =
        await this.getCurrentUser(
          context
        );


      // =====================================================
      // Get Existing Favourites
      // =====================================================

      const favouritesResponse =
        await context.spHttpClient.get(

          `${webUrl}` +
          `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items` +
          `?$select=Id,Quick_x0020_LinkId` +
          `&$filter=UserId eq ${currentUser.Id}`,

          SPHttpClient.configurations.v1

        );


      if (!favouritesResponse.ok) {

        console.error(
          'Unable to get existing favourites:',
          favouritesResponse.status
        );

        return;

      }


      const favouritesData =
        await favouritesResponse.json();


      const existingFavourites =
        favouritesData.value;


      const keptIds: number[] = [];


      // =====================================================
      // Remove Unchecked Favourites
      // =====================================================

      for (
        let i = 0;
        i < existingFavourites.length;
        i++
      ) {

        const favourite =
          existingFavourites[i];


        const quickLinkId =
          Number(
            favourite.Quick_x0020_LinkId
          );


        if (
          selectedIds.indexOf(
            quickLinkId
          ) === -1
        ) {

          await this.deleteFavourite(
            context,
            favourite.Id
          );

        } else {

          keptIds.push(
            quickLinkId
          );

        }

      }


      // =====================================================
      // Add Newly Selected Favourites
      // =====================================================

      for (
        let i = 0;
        i < selectedIds.length;
        i++
      ) {

        const quickLinkId =
          selectedIds[i];


        if (
          keptIds.indexOf(
            quickLinkId
          ) === -1
        ) {

          const requestBody = {

            UserId:
              currentUser.Id,

            Quick_x0020_LinkId:
              quickLinkId

          };


          const response =
            await context.spHttpClient.post(

              `${webUrl}` +
              `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items`,

              SPHttpClient.configurations.v1,

              {

                headers: {

                  'Accept':
                    'application/json;odata=nometadata',

                  'Content-Type':
                    'application/json;odata=nometadata'

                },

                body:
                  JSON.stringify(
                    requestBody
                  )

              }

            );


          if (!response.ok) {

            const errorText =
              await response.text();


            console.error(
              'Failed to add favourite:',
              response.status,
              errorText
            );

            return;

          }


          keptIds.push(
            quickLinkId
          );

        }

      }

      const activeTab =
        domElement.querySelector(
          '.panel-title-filter-active'
        ) as HTMLElement;

      const activeTabValue =
        activeTab?.getAttribute(
          'data-filter-ql'
        );



      // =====================================================
      // Refresh Quick Links
      // =====================================================

      const quickLinks =
        await this.getQuickLinks(
          context
        );


      const favouriteIds =
        await this.getUserFavouriteIds(
          context
        );


      this.renderQuickLinks(
        context,
        domElement,
        quickLinks,
        favouriteIds
      );


      this.renderFavouriteOptions(
        domElement,
        quickLinks,
        favouriteIds
      );


      this.updateFavouriteTabVisibility(
            domElement,
            favouriteIds
          );

          if (activeTabValue) {

      const tabToRestore =
        domElement.querySelector(
          `[data-filter-ql="${activeTabValue}"]`
        ) as HTMLElement;

      if (tabToRestore) {
        tabToRestore.click();
      }

    }


      // =====================================================
      // Close Modal
      // =====================================================

      const modal =
        domElement.querySelector(
          '#addFavouriteModal'
        ) as HTMLElement;


      if (modal) {

        modal.classList.remove(
          'show'
        );


        modal.style.display =
          'none';


        modal.setAttribute(
          'aria-hidden',
          'true'
        );


        document.body.classList.remove(
          'modal-open'
        );


        const backdrop =
          document.querySelector(
            '.modal-backdrop'
          );


        if (backdrop) {

          backdrop.remove();

        }

      }

    } catch (error) {

      console.error(
        'Error updating favourites:',
        error
      );

    }

  }


  // =========================================================
  // Delete Favourite
  // =========================================================

  private async deleteFavourite(
    context: WebPartContext,
    favouriteId: number
  ): Promise<void> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const response =
      await context.spHttpClient.post(

        `${webUrl}` +
        `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items(${favouriteId})`,

        SPHttpClient.configurations.v1,

        {

          headers: {

            'Accept':
              'application/json;odata=nometadata',

            'X-HTTP-Method':
              'DELETE',

            'IF-MATCH':
              '*'

          }

        }

      );


    if (!response.ok) {

      const errorText =
        await response.text();


      console.error(
        'Failed to delete favourite:',
        favouriteId,
        response.status,
        errorText
      );


      throw new Error(
        `Failed to delete favourite: ${response.status}`
      );

    }

  }


  // =========================================================
  // BIRTHDAYS
  // =========================================================


  // =========================================================
  // Get Birthdays
  // =========================================================

  private async getBirthdays(
    context: WebPartContext
  ): Promise<IBirthdayList[]> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const endpoint =
      `${webUrl}/_api/web/lists/getbytitle('Birthday')/items` +
      `?$select=Id,Title,EmployeePhoto,BirthDate,Status`;


    try {

      const response =
        await context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept:
                'application/json;odata=nometadata'
            }
          }
        );


      if (!response.ok) {

        console.error(
          'Failed to get birthday data:',
          response.status,
          response.statusText
        );

        return [];

      }


      const data =
        await response.json();


      return data.value;

    } catch (error) {

      console.error(
        'Error getting birthday data:',
        error
      );

      return [];

    }

  }


  // =========================================================
  // Get Employee Photo URL
  // =========================================================

  private getBirthdayImageUrl(
    context: WebPartContext,
    photo: any,
    itemId: number
  ): string {

    if (!photo) {

      return '';

    }


    try {

      if (typeof photo === 'string') {

        photo =
          JSON.parse(
            photo
          );

      }


      if (!photo.fileName) {

        return '';

      }


      return (
        context.pageContext.web.absoluteUrl +
        '/Lists/Birthday/Attachments/' +
        itemId +
        '/' +
        encodeURIComponent(
          photo.fileName
        )
      );

    } catch (error) {

      console.error(
        'Error parsing employee photo:',
        error
      );

      return '';

    }

  }


  // =========================================================
  // Get Upcoming Birthdays
  // =========================================================

  private getUpcomingBirthdays(
    birthdays: IBirthdayList[]
  ): IBirthdayList[] {

    const today =
      new Date();


    today.setHours(
      0,
      0,
      0,
      0
    );


    const upcoming =
      birthdays

        .filter(
          (item) =>
            item.Status === 'Active'
        )

        .map(
          (item) => {

            const birthDate =
              new Date(
                item.BirthDate
              );


            let birthday =
              new Date(
                today.getFullYear(),
                birthDate.getMonth(),
                birthDate.getDate()
              );


            if (
              birthday < today
            ) {

              birthday.setFullYear(
                today.getFullYear() + 1
              );

            }


            const difference =
              Math.floor(
                (
                  birthday.getTime() -
                  today.getTime()
                ) /
                (
                  1000 *
                  60 *
                  60 *
                  24
                )
              );


            return {
              item,
              difference
            };

          }
        )

        .filter(
          (item) =>
            item.difference < 7
        )

        .sort(
          (a, b) =>
            a.difference -
            b.difference
        );


    return upcoming.map(
      (item) =>
        item.item
    );

  }


  // =========================================================
  // Render Birthdays
  // =========================================================

  private async renderBirthdays(): Promise<void> {
    

    const birthdays =
      await this.getBirthdays(
        this.context
      );


    const upcomingBirthdays =
      this.getUpcomingBirthdays(
        birthdays
      );


    const container =
      this.domElement.querySelector(
        '.panel-card-birthdays'
      );


    if (!container) {

      console.error(
        'Birthday container not found.'
      );

      return;

    }


    let html = '';


    upcomingBirthdays.forEach(
      (item) => {

        const birthDate =
          new Date(
            item.BirthDate
          );


        const month =
          birthDate
            .toLocaleString(
              'en-US',
              {
                month: 'short'
              }
            )
            .toUpperCase();


        const day =
          birthDate
            .getDate()
            .toString();


        const imageUrl =
          this.getBirthdayImageUrl(
            this.context,
            item.EmployeePhoto,
            item.Id
          );


        html +=
          Birthday.singleElementHtml

            .replace(
              /__KEY_URL_IMG__/g,
              imageUrl
            )

            .replace(
              /__KEY_TITLE__/g,
              item.Title
            )

            .replace(
              /__KEY_MONTH__/g,
              month
            )

            .replace(
              /__KEY_DAY__/g,
              day
            );

      }
    );


    container.innerHTML =
      html;

  }

  // =====================================================
// HR API - Birthdays
// =====================================================

// private async getHRBirthdays(): Promise<any[]> {

//   const apiUrl = 'HR_API_URL_WILL_BE_PROVIDED';

//   // API implementation will be added
//   // once HR provides endpoint and authentication details.

//   return [];
// }

// =========================================================
// SOCIAL MEDIA
// =========================================================

private async initializeSocialMedia(): Promise<void> {

  // Instagram
  const instagramContainer =
    this.domElement.querySelector(
      '#social-instagram'
    ) as HTMLElement;

  if (instagramContainer) {
    instagramContainer.innerHTML = `
      <div
        class="sk-instagram-feed"
        data-embed-id="25716225">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/instagram-feed/widget.js';
    script.defer = true;

    instagramContainer.appendChild(script);
  }


  // X (Twitter)
  const twitterContainer =
    this.domElement.querySelector(
      '#social-twitter'
    ) as HTMLElement;

  if (twitterContainer) {
    twitterContainer.innerHTML = `
      <div
        class="sk-ww-twitter-feed"
        data-embed-id="25716230">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/twitter-feed/widget.js';
    script.defer = true;

    twitterContainer.appendChild(script);
  }


  // LinkedIn
  const linkedinContainer =
    this.domElement.querySelector(
      '#social-linkedin'
    ) as HTMLElement;

  if (linkedinContainer) {
    linkedinContainer.innerHTML = `
      <div
        class="sk-ww-linkedin-page-post"
        data-embed-id="25716233">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/linkedin-page-posts/widget.js';
    script.defer = true;

    linkedinContainer.appendChild(script);
  }


  // Facebook
  const facebookContainer =
    this.domElement.querySelector(
      '#social-facebook'
    ) as HTMLElement;

  if (facebookContainer) {
    facebookContainer.innerHTML = `
      <div
        class="sk-ww-facebook-page-posts"
        data-embed-id="25716238">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/facebook-page-posts/widget.js';
    script.defer = true;

    facebookContainer.appendChild(script);
  }


  // YouTube
  const youtubeContainer =
    this.domElement.querySelector(
      '#social-youtube'
    ) as HTMLElement;

  if (youtubeContainer) {
    youtubeContainer.innerHTML = `
      <div
        class="sk-ww-youtube-channel-videos"
        data-embed-id="25716248">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/youtube-channel-videos/widget.js';
    script.defer = true;

    youtubeContainer.appendChild(script);
  }

  this.setupSocialMediaTutorialLinkObserver();
}

  // ==================== INITIALIZE WEB PART ====================

  protected async onInit(): Promise<void> {

    this.loadCSS();


    await this.loadJS();


    return super.onInit();

  }


  // ==================== PROPERTY PANE ====================

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {

      pages: [

        {

          header: {
            description: 'Home Page'
          },

          groups: [

            {

              groupName: 'Configuration',

              groupFields: [

                PropertyPaneTextField(
                  'description',
                  {
                    label: 'Description'
                  }
                )

              ]

            }

          ]

        }

      ]

    };

  }


  // ==================== DATA VERSION ====================

  protected get dataVersion(): Version {

    return Version.parse('1.0');

  }

}