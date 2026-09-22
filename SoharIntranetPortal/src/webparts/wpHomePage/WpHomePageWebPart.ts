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


export interface IWpHomePageWebPartProps {
  description: string;
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


  private newsCentre!: NewsCentre;

  private newsItems: INewsItem[] = [];


  public async onInit(): Promise<void> {

    await super.onInit();

    await this.loadCSS();

    await this.loadJS();
  }



  


  public async render(): Promise<void> {




    const workbenchContent = document.getElementById('workbenchPageContent');
    if (workbenchContent) {
      workbenchContent.style.maxWidth = 'none';
    }
 


    

    /*
     * ==========================================================
     * CREATE NEWS CENTRE TEMPLATE
     * ==========================================================
     */

    this.newsCentre =
      new NewsCentre();


    /*
     * ==========================================================
     * LOAD SHAREPOINT NEWS
     * ==========================================================
     */

    await this.loadNews();


    /*
     * ==========================================================
     * BUILD DYNAMIC HTML
     * ==========================================================
     */

    const allItemsHtml =
      this.buildCategoryHtml('All');


    const announcementItemsHtml =
      this.buildCategoryHtml('Announcements');


    let html =
      this.newsCentre.allElementsHtml;


    html =
      html.replace(
        /__KEY_ALL_ITEMS__/g,
        allItemsHtml
      );


    html =
      html.replace(
        /__KEY_ANNOUNCEMENT_ITEMS__/g,
        announcementItemsHtml
      );


    html =
      html.replace(
        /__KEY_URL_ARROW__/g,
        this.getArrowImageUrl()
      );


    html =
      html.replace(
        /__KEY_URL_VIEW_ALL__/g,
        `${this.context.pageContext.web.absoluteUrl}/SitePages/News-List.aspx`
      );


    /*
     * ==========================================================
     * INSERT INTO WEBPART
     * ==========================================================
     */

    this.domElement.innerHTML =
      html;


    /*
     * ==========================================================
     * TAB FUNCTIONALITY
     * ==========================================================
     */

    this.attachTabEvents();
  }


  /*
   * ============================================================
   * LOAD NEWS FROM SHAREPOINT
   * ============================================================
   */

  private async loadNews(): Promise<void> {

    const webUrl =
      this.context.pageContext.web.absoluteUrl;


    const apiUrl =
      `${webUrl}/_api/web/lists/getbytitle('News')/items` +
      `?$select=Id,Title,NewsIcon,ShortDescription,Category,MainContent,PublishedDate,Status` +
      `&$filter=Status eq 'Active'` +
      `&$orderby=PublishedDate desc`;


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


      if (!response.ok) {

        throw new Error(
          `News API failed: ${response.status}`
        );
      }


      const data =
        await response.json();


      this.newsItems =
        data.value || [];


      console.log(
        'News items loaded:',
        this.newsItems
      );

    } catch (error) {

      console.error(
        'Error loading News list:',
        error
      );


      this.newsItems = [];
    }
  }


  /*
   * ============================================================
   * BUILD CATEGORY HTML
   * ============================================================
   */

  private buildCategoryHtml(
    category: string
  ): string {


    let items: INewsItem[];


    if (category === 'All') {

      items =
        this.newsItems;

    } else {

      items =
        this.newsItems.filter(
          item =>
            this.normalizeValue(
              item.Category
            ) ===
            this.normalizeValue(
              category
            )
        );
    }


    if (!items.length) {

      return this.newsCentre.noElementHtml;
    }


    return items
      .map(
        item =>
          this.buildSingleNewsItem(item)
      )
      .join('');
  }


  /*
   * ============================================================
   * BUILD SINGLE NEWS ITEM
   * ============================================================
   */

  private buildSingleNewsItem(
    item: INewsItem
  ): string {


    const title =
      this.escapeHtml(
        item.Title || ''
      );


    const description =
      this.escapeHtml(
        item.ShortDescription || ''
      );


    const category =
      this.escapeHtml(
        item.Category || ''
      );


    const date =
      this.formatDate(
        item.PublishedDate
      );


    const imageUrl =
      this.getNewsImageUrl(item);


    const detailsUrl =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/Lists/News/DispForm.aspx?ID=${item.Id}`;


    return this.newsCentre.singleElementHtml

      .replace(
        /__KEY_URL_IMGICON__/g,
        imageUrl
      )

      .replace(
        /__KEY_DATA_TITLE__/g,
        title
      )

      .replace(
        /__KEY_DATA_DESCRIPTION__/g,
        description
      )

      .replace(
        /__KEY_DATA_CATEGORY__/g,
        category
      )

      .replace(
        /__KEY_DATA_DATE__/g,
        date
      )

      .replace(
        /__KEY_URL_LINK__/g,
        detailsUrl
      )

      .replace(
        /__KEY_URL_ARROW__/g,
        this.getArrowImageUrl()
      );
  }


  /*
   * ============================================================
   * NEWS IMAGE
   * ============================================================
   */

  private getNewsImageUrl(
    item: INewsItem
  ): string {


    if (!item.NewsIcon) {

      return '';
    }


    try {

      const imageData =
        typeof item.NewsIcon === 'string'
          ? JSON.parse(item.NewsIcon)
          : item.NewsIcon;


      const fileName =
        imageData?.fileName ||
        imageData?.FileName ||
        '';


      if (!fileName) {

        return '';
      }


      return (
        `${this.context.pageContext.web.absoluteUrl}` +
        `/Lists/News/Attachments/` +
        `${item.Id}/${fileName}`
      );

    } catch (error) {

      console.error(
        'Error parsing NewsIcon:',
        error
      );


      return '';
    }
  }


  /*
   * ============================================================
   * ARROW IMAGE
   * ============================================================
   */

  private getArrowImageUrl(): string {

    return (
      `${this.context.pageContext.web.absoluteUrl}` +
      `/SiteAssets/resources/images/icons/arrow-right-short.svg`
    );
  }


  /*
   * ============================================================
   * TAB FUNCTIONALITY
   * ============================================================
   *
   * The ORIGINAL tab markup is not modified.
   *
   * The WebPart owns the behaviour.
   * ============================================================
   */

  private attachTabEvents(): void {


    const tabs =
      this.domElement.querySelectorAll(
        '#news-tabs .tab-title-pill'
      );


    const allPanel =
      this.domElement.querySelector(
        '#news-panel-all'
      ) as HTMLElement | null;


    const announcementPanel =
      this.domElement.querySelector(
        '#news-panel-announcements'
      ) as HTMLElement | null;


    if (!tabs.length) {

      return;
    }


    tabs.forEach(
      tab => {

        tab.addEventListener(
          'click',
          () => {

            const targetId =
              tab.getAttribute(
                'data-tab-news-id'
              );


            if (!targetId) {

              return;
            }


            /*
             * Remove active class from every tab.
             */

            tabs.forEach(
              otherTab => {

                otherTab.classList.remove(
                  'tab-title-pill-active'
                );
              }
            );


            /*
             * Activate clicked tab.
             */

            tab.classList.add(
              'tab-title-pill-active'
            );


            /*
             * Original HTML only provides actual
             * panel markup for All and Announcements.
             *
             * We therefore keep that original HTML
             * untouched and dynamically change the
             * existing content for the remaining
             * category tabs.
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


              this.renderIntoExistingPanel(
                allPanel,
                this.buildCategoryHtml('All')
              );


              return;
            }


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


              this.renderIntoExistingPanel(
                announcementPanel,
                this.buildCategoryHtml(
                  'Announcements'
                )
              );


              return;
            }


            /*
             * Events / News / Circulars.
             *
             * The original supplied HTML has no
             * corresponding panel elements.
             *
             * We therefore reuse the existing panel
             * structure rather than changing the
             * original HTML template.
             */

            if (allPanel) {

              allPanel.style.display =
                'block';
            }


            if (announcementPanel) {

              announcementPanel.style.display =
                'none';
            }


            let category =
              '';


            if (
              targetId ===
              'news-panel-events'
            ) {

              category = 'Events';
            }


            if (
              targetId ===
              'news-panel-news'
            ) {

              category = 'News';
            }


            if (
              targetId ===
              'news-panel-circulars'
            ) {

              category = 'Circulars';
            }


            this.renderIntoExistingPanel(
              allPanel,
              this.buildCategoryHtml(category)
            );
          }
        );
      }
    );


    /*
     * Initial state.
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
   * RENDER INTO EXISTING ORIGINAL PANEL
   * ============================================================
   */

  private renderIntoExistingPanel(
    panel: HTMLElement | null,
    html: string
  ): void {


    if (!panel) {

      return;
    }


    const container =
      panel.querySelector(
        '.panel-card-news'
      ) as HTMLElement | null;


    if (!container) {

      return;
    }


    container.innerHTML =
      html;
  }


  /*
   * ============================================================
   * DATE FORMAT
   * ============================================================
   */

  private formatDate(
    value?: string
  ): string {


    if (!value) {

      return '';
    }


    const date =
      new Date(value);


    if (isNaN(date.getTime())) {

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
   */

  private normalizeValue(
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

  private escapeHtml(
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
   * CSS
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

        SPComponentLoader.loadCss(
          `${baseUrl}/SiteAssets/resources/css/${file}`
        );


        console.log(
          `CSS loaded: ${file}`
        );

      } catch (error) {

        console.error(
          `CSS failed: ${file}`,
          error
        );
      }
    }
  }


  /*
   * ============================================================
   * JAVASCRIPT
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

      } catch (error) {

        console.error(
          `JS failed: ${script}`,
          error
        );
      }
    }
  }
}