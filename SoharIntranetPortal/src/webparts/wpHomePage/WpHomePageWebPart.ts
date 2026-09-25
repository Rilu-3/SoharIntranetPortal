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
      this.newsCentreRenderCategory('Events');

      this.newsCentreRenderCategory('News');

      this.newsCentreRenderCategory('Circulars');

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
   * ----------------------------------------------------------
   * All
   * ----------------------------------------------------------
   */

  if (category === 'All') {

    panel =
      this.domElement.querySelector(
        '#news-panel-all'
      ) as HTMLElement | null;

  }


  /*
   * ----------------------------------------------------------
   * Announcements
   * ----------------------------------------------------------
   */

  else if (category === 'Announcements') {

    panel =
      this.domElement.querySelector(
        '#news-panel-announcements'
      ) as HTMLElement | null;

  }


  /*
   * ----------------------------------------------------------
   * Events
   * ----------------------------------------------------------
   */

  else if (category === 'Events') {

    panel =
      this.domElement.querySelector(
        '#news-panel-events'
      ) as HTMLElement | null;

  }


  /*
   * ----------------------------------------------------------
   * News
   * ----------------------------------------------------------
   */

  else if (category === 'News') {

    panel =
      this.domElement.querySelector(
        '#news-panel-news'
      ) as HTMLElement | null;

  }


  /*
   * ----------------------------------------------------------
   * Circulars
   * ----------------------------------------------------------
   */

  else if (category === 'Circulars') {

    panel =
      this.domElement.querySelector(
        '#news-panel-circulars'
      ) as HTMLElement | null;

  }


  /*
   * ----------------------------------------------------------
   * Invalid category
   * ----------------------------------------------------------
   */

  else {

    console.error(
      `Unknown News category: ${category}`
    );

    return;

  }


  /*
   * ==========================================================
   * 2. CHECK WHETHER PANEL EXISTS
   * ==========================================================
   */

  if (!panel) {

    console.error(
      `News panel not found for category: ${category}`
    );

    return;

  }


  /*
   * ==========================================================
   * 3. FIND NEWS CONTAINER
   * ==========================================================
   *
   * Every category panel contains:
   *
   * <div class="panel-card-news">
   *
   * This is where the generated News item HTML
   * will be inserted.
   *
   */

  const container =
    panel.querySelector(
      '.panel-card-news'
    ) as HTMLElement | null;


  /*
   * Container not found
   */

  if (!container) {

    console.error(
      `News container not found for category: ${category}`
    );

    return;

  }


  /*
   * ==========================================================
   * 4. FILTER NEWS ITEMS
   * ==========================================================
   */

  let items: INewsItem[];


  /*
   * ----------------------------------------------------------
   * All category
   * ----------------------------------------------------------
   *
   * Show every News item.
   *
   */

  if (category === 'All') {

    items =
      this.newsCentreItems;

  }


  /*
   * ----------------------------------------------------------
   * Specific category
   * ----------------------------------------------------------
   *
   * Only show items whose Category matches
   * the selected tab.
   *
   */

  else {

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
   * 5. START COMPLETE HTML STRING
   * ==========================================================
   *
   * This will contain all News items for
   * the selected category.
   *
   */

  let allElementsHtml: string = '';


  /*
   * ==========================================================
   * 6. LOOP THROUGH NEWS ITEMS
   * ==========================================================
   */

  items.forEach(
    (item) => {


      /*
       * ======================================================
       * IMAGE URL
       * ======================================================
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
       */

      let singleElementHtml =
        NewsCentre.singleElementHtml;


      /*
       * ======================================================
       * REPLACE NEWS ICON
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_URL_IMGICON__/g,
          imageUrl
        );


      /*
       * ======================================================
       * REPLACE TITLE
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_DATA_TITLE__/g,
          this.newsCentreEscapeHtml(
            item.Title || ''
          )
        );


      /*
       * ======================================================
       * REPLACE DESCRIPTION
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_DATA_DESCRIPTION__/g,
          this.newsCentreEscapeHtml(
            item.ShortDescription || ''
          )
        );


      /*
       * ======================================================
       * REPLACE CATEGORY
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_DATA_CATEGORY__/g,
          this.newsCentreEscapeHtml(
            item.Category || ''
          )
        );


      /*
       * ======================================================
       * REPLACE DATE
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_DATA_DATE__/g,
          createddate
        );


      /*
       * ======================================================
       * REPLACE DETAILS LINK
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_URL_LINK__/g,
          detailsUrl
        );


      /*
       * ======================================================
       * REPLACE ARROW IMAGE
       * ======================================================
       */

      singleElementHtml =
        singleElementHtml.replace(
          /__KEY_URL_ARROW__/g,
          this.newsCentreGetArrowImageUrl()
        );


      /*
       * ======================================================
       * ADD ITEM TO COMPLETE HTML
       * ======================================================
       */

      allElementsHtml +=
        singleElementHtml;

    }
  );


  /*
   * ==========================================================
   * 7. NO DATA
   * ==========================================================
   *
   * If the selected category has no items,
   * show the "No records found." message.
   *
   */

  if (!items.length) {

    allElementsHtml =
      NewsCentre.noElementHtml;

  }


  /*
   * ==========================================================
   * 8. INSERT HTML INTO CATEGORY CONTAINER
   * ==========================================================
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

  /*
 * ============================================================
 * TAB FUNCTIONALITY
 * ============================================================
 */

private newsCentreAttachTabEvents(): void {

  /*
   * ==========================================================
   * 1. FIND ALL TABS
   * ==========================================================
   */

  const tabs =
    this.domElement.querySelectorAll(
      '#news-tabs .tab-title-pill'
    );


  /*
   * ==========================================================
   * 2. FIND ALL NEWS PANELS
   * ==========================================================
   */

  const panels =
    this.domElement.querySelectorAll(
      '#news-tabs .news-panel-tab-view'
    );


  /*
   * ==========================================================
   * 3. CHECK WHETHER TABS EXIST
   * ==========================================================
   */

  if (!tabs.length) {

    console.warn(
      'News tabs not found.'
    );

    return;
  }


  /*
   * ==========================================================
   * 4. ADD CLICK EVENT TO EACH TAB
   * ==========================================================
   */

  tabs.forEach(
    (tab) => {

      tab.addEventListener(
        'click',
        () => {

          /*
           * ==================================================
           * GET TARGET PANEL ID
           * ==================================================
           *
           * Example:
           *
           * data-tab-news-id="news-panel-events"
           *
           */

          const targetId =
            tab.getAttribute(
              'data-tab-news-id'
            );


          /*
           * No target ID
           */

          if (!targetId) {

            return;
          }


          /*
           * ==================================================
           * REMOVE ACTIVE CLASS FROM ALL TABS
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
           * ==================================================
           * ADD ACTIVE CLASS TO CLICKED TAB
           * ==================================================
           */

          tab.classList.add(
            'tab-title-pill-active'
          );


          /*
           * ==================================================
           * HIDE ALL PANELS
           * ==================================================
           */

          panels.forEach(
            (panel) => {

              (panel as HTMLElement).style.display =
                'none';

            }
          );


          /*
           * ==================================================
           * SHOW SELECTED PANEL
           * ==================================================
           */

          const targetPanel =
            this.domElement.querySelector(
              `#${targetId}`
            ) as HTMLElement | null;


          if (targetPanel) {

            targetPanel.style.display =
              'block';

          }
          else {

            console.warn(
              `News panel not found: ${targetId}`
            );

            return;
          }


          /*
           * ==================================================
           * DETERMINE CATEGORY
           * ==================================================
           */

          let category = '';


          if (
            targetId ===
            'news-panel-all'
          ) {

            category = 'All';

          }
          else if (
            targetId ===
            'news-panel-announcements'
          ) {

            category = 'Announcements';

          }
          else if (
            targetId ===
            'news-panel-events'
          ) {

            category = 'Events';

          }
          else if (
            targetId ===
            'news-panel-news'
          ) {

            category = 'News';

          }
          else if (
            targetId ===
            'news-panel-circulars'
          ) {

            category = 'Circulars';

          }


          /*
           * ==================================================
           * RENDER SELECTED CATEGORY
           * ==================================================
           */

          if (category) {

            this.newsCentreRenderCategory(
              category
            );

          }

        }
      );

    }
  );


  /*
   * ==========================================================
   * DEFAULT TAB STATE
   * ==========================================================
   *
   * All tab should be active when the page loads.
   *
   */

  tabs.forEach(
    (tab) => {

      tab.classList.remove(
        'tab-title-pill-active'
      );

    }
  );


  const defaultTab =
    this.domElement.querySelector(
      '[data-tab-news-id="news-panel-all"]'
    ) as HTMLElement | null;


  if (defaultTab) {

    defaultTab.classList.add(
      'tab-title-pill-active'
    );

  }


  /*
   * ==========================================================
   * DEFAULT PANEL STATE
   * ==========================================================
   *
   * Hide every panel first.
   *
   */

  panels.forEach(
    (panel) => {

      (panel as HTMLElement).style.display =
        'none';

    }
  );


  /*
   * ==========================================================
   * SHOW ALL PANEL
   * ==========================================================
   */

  const allPanel =
    this.domElement.querySelector(
      '#news-panel-all'
    ) as HTMLElement | null;


  if (allPanel) {

    allPanel.style.display =
      'block';

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