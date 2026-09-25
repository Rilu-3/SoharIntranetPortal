import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  SPComponentLoader
} from '@microsoft/sp-loader';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import NewsCentre from './NewsCentre';


/*
 * ============================================================
 * INTERFACE
 * ============================================================
 */

export interface IWpHomePageWebPartProps {
  description: string;
}


/*
 * ============================================================
 * NEWS ITEM INTERFACE
 * ============================================================
 */

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


/*
 * ============================================================
 * WEB PART CLASS
 * ============================================================
 */

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  /*
   * ============================================================
   * NEWS ITEMS
   * ============================================================
   *
   * Stores the News items retrieved from SharePoint.
   *
   * The same data is reused when the user changes tabs.
   */

  private newsCentreItems: INewsItem[] = [];


  /*
   * ============================================================
   * INITIALIZATION
   * ============================================================
   *
   * Load required CSS and JavaScript files.
   */

  public async onInit(): Promise<void> {

    await super.onInit();

    await this.loadCSS();

    await this.loadJS();
  }


  /*
   * ============================================================
   * RENDER
   * ============================================================
   *
   * Main Web Part flow:
   *
   * 1. Set SharePoint workbench width.
   * 2. Add NewsCentre main HTML template.
   * 3. Set static elements.
   * 4. Create SharePoint REST API URL.
   * 5. Load News data.
   * 6. Generate News HTML.
   * 7. Attach tab events.
   */

  public async render(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    /*
     * ----------------------------------------------------------
     * Remove SharePoint workbench maximum width.
     * ----------------------------------------------------------
     */

    const workbenchContent =
      document.getElementById(
        'workbenchPageContent'
      );

    if (workbenchContent) {

      workbenchContent.style.maxWidth =
        'none';
    }


    /*
     * ==========================================================
     * 1. ADD MAIN NEWS CENTRE TEMPLATE
     * ==========================================================
     *
     * This is the same pattern as your AnnouncementOffer code.
     */

    this.domElement.innerHTML =
      NewsCentre.allElementsHtml;


    /*
     * ==========================================================
     * 2. SET STATIC ELEMENTS
     * ==========================================================
     */

    const arrowIconUrl =
      `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;

    this.newsCentreSetupViewAll(
      arrowIconUrl
    );


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

    await this._renderNewsAsync(
      newsApiUrl
    );


    /*
     * ==========================================================
     * 5. ATTACH TAB FUNCTIONALITY
     * ==========================================================
     */

    this.newsCentreAttachTabEvents();
  }


  /*
   * ============================================================
   * VIEW ALL / ARROW
   * ============================================================
   *
   * Sets the View All URL and arrow image after the
   * NewsCentre template has already been inserted.
   */

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


  /*
   * ============================================================
   * LOAD CSS
   * ============================================================
   */

  private async loadCSS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    const cssFiles = [
      'bootstrap.min.css',
      'custom.css',
      'font-size.css',
      'home.css',
      'jquery-ui.css',
      'sp-custom.css',
      'swiper-bundle.min.css',
      'variable.css'
    ];


    for (const file of cssFiles) {

      try {

        await SPComponentLoader.loadCss(
          `${baseUrl}/SiteAssets/resources/css/${file}`
        );

        console.log(
          `CSS loaded: ${file}`
        );

      }
      catch (error) {

        console.error(
          `CSS failed: ${file}`,
          error
        );
      }
    }
  }


  /*
   * ============================================================
   * LOAD JAVASCRIPT
   * ============================================================
   */

  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    const jsBaseUrl =
      `${baseUrl}/SiteAssets/resources/js`;


    const scripts = [
      'jquery-3.6.0.js',
      'jquery-ui.js',
      'jquery.marquee.min.js',
      'swiper-bundle.min.js',
      'bootstrap.bundle.min.js',
      'common.js',
      'home.js'
    ];


    for (const script of scripts) {

      try {

        await SPComponentLoader.loadScript(
          `${jsBaseUrl}/${script}`
        );

        console.log(
          `JS loaded: ${script}`
        );

      }
      catch (error) {

        console.error(
          `JS failed: ${script}`,
          error
        );
      }
    }
  }
}