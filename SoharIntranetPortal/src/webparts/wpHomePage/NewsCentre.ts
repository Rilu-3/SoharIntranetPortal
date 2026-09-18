import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import { WebPartContext } from '@microsoft/sp-webpart-base';

interface INewsItem {
  Id: number;
  Title: string;
  NewsIcon: any;
  ShortDescription: string;
  Category: string;
  MainContent: string;
  PublishedDate: string;
  Status: string;
}

export default class NewsCentre {

  private context: WebPartContext;

  private newsItems: INewsItem[] = [];

  /*
   * Root element belonging only to News Centre.
   *
   * This prevents News Centre from interfering
   * with other sections on the final homepage.
   */
  private rootElement!: HTMLElement;


  public static singleElementHtml: string = `

    <div class="news-item flex-column flex-sm-row">

      <div class="d-flex flex-grow-1 align-items-center gap-3">

        <div class="news-icon">

          <img
            src="__KEY_URL_IMGICON__"
            alt="__KEY_DATA_TITLE__"
            class="news-item-image"
          />

        </div>

        <div class="flex-grow-1">

          <p class="news-title">
            __KEY_DATA_TITLE__
          </p>

          <p class="news-desc">
            __KEY_DATA_DESCRIPTION__
          </p>

          <p class="news-meta mt-1">

            <span class="tag">
              __KEY_DATA_CATEGORY__
            </span>

            &nbsp;•&nbsp;

            __KEY_DATA_DATE__

          </p>

        </div>

      </div>

      <a
        href="__KEY_URL_LINK__"
        class="link-arrow text-color-link"
      >

        <span>
          Read More
        </span>

        <img
          src="__KEY_URL_ARROW__"
          alt=""
        />

      </a>

    </div>

  `;


  public static allElementsHtml: string = `

    <div class="panel-card px-2 py-4 d-flex flex-column">

      <div class="panel-header px-2 w-100 float-start mb-4">

        <h2 class="panel-title">
          News Centre
        </h2>

        <a
          href="#"
          class="link-arrow text-color-link"
        >

          <span class="text-sm xxl-text-base font-bold">
            View All News
          </span>

          <img
            src="__KEY_URL_ARROW__"
            alt=""
          />

        </a>

      </div>

      <div
        id="news-tabs"
        class="w-100 float-start d-flex flex-column flex-grow-1 overflow-hidden"
      >

        <ul class="news-tabs-list px-2">

          <li>

            <div
              data-news-category="All"
              class="tab-title-pill tab-title-pill-active"
            >
              All
            </div>

          </li>

          <li>

            <div
              data-news-category="Announcements"
              class="tab-title-pill"
            >
              Announcements
            </div>

          </li>

          <li>

            <div
              data-news-category="Events"
              class="tab-title-pill"
            >
              Events
            </div>

          </li>

          <li>

            <div
              data-news-category="News"
              class="tab-title-pill"
            >
              News
            </div>

          </li>

          <li>

            <div
              data-news-category="Circulars"
              class="tab-title-pill"
            >
              Circulars
            </div>

          </li>

        </ul>

        <div
          id="news-items-container"
          class="news-tabs-content w-100 float-start"
        ></div>

      </div>

    </div>

  `;


  public static noRecord: string = `

    <div class="w-100 float-start text-center py-4">

      <p class="m-0 strive-text-secondary">
        No News Available
      </p>

    </div>

  `;


  constructor(
    context: WebPartContext
  ) {

    this.context = context;

  }


  /*
   * =====================================================
   * RENDER NEWS CENTRE
   * =====================================================
   *
   * The WebPart gives us the root container.
   *
   * NewsCentre renders everything inside that container.
   */
  public render(
    rootElement: HTMLElement
  ): void {

    this.rootElement =
      rootElement;

    /*
     * Arrow icon.
     */
    const arrowUrl =
      `${this.context.pageContext.web.absoluteUrl}` +
      `/SiteAssets/resources/images/icons/arrow-right-short.svg`;

    /*
     * Render News Centre HTML.
     */
    this.rootElement.innerHTML =
      NewsCentre.allElementsHtml
        .replace(
          /__KEY_URL_ARROW__/g,
          arrowUrl
        );

    /*
     * Load News from SharePoint.
     */
    this.loadNews();

  }


  /*
   * =====================================================
   * LOAD NEWS
   * =====================================================
   */
  private async loadNews(): Promise<void> {

    try {

      const webUrl =
        this.context.pageContext.web.absoluteUrl;

      const apiUrl =
        `${webUrl}/_api/web/lists/getbytitle('News')/items` +
        `?$select=Id,Title,NewsIcon,ShortDescription,Category,MainContent,PublishedDate,Status` +
        `&$filter=Status eq 'Active'` +
        `&$orderby=PublishedDate desc`;

      console.log(
        '📰 News API:',
        apiUrl
      );

      const response:
        SPHttpClientResponse =
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
          `News API failed: ${response.status} ${response.statusText}`
        );

      }

      const data =
        await response.json();

      this.newsItems =
        data.value || [];

      console.log(
        '📰 News items:',
        this.newsItems
      );

      /*
       * Initially display all News.
       */
      this.renderNewsItems(
        'All'
      );

      /*
       * Attach category click events.
       */
      this.attachTabEvents();

    } catch (error) {

      console.error(
        '❌ Error loading News:',
        error
      );

      const container =
        this.getNewsItemsContainer();

      if (container) {

        container.innerHTML =
          NewsCentre.noRecord;

      }

    }

  }


  /*
   * =====================================================
   * RENDER NEWS ITEMS
   * =====================================================
   *
   * IMPORTANT:
   *
   * Existing category filtering is preserved.
   *
   * Events → only Events
   * News → only News
   * Announcements → only Announcements
   * Circulars → only Circulars
   * All → all items
   */
  private renderNewsItems(
    category: string
  ): void {

    const container =
      this.getNewsItemsContainer();

    if (!container) {

      console.error(
        '❌ #news-items-container not found.'
      );

      return;

    }

    let filteredItems =
      this.newsItems;

    /*
     * Category filtering.
     */
    if (
      category.toLowerCase() !== 'all'
    ) {

      filteredItems =
        this.newsItems.filter(
          (item: INewsItem) => {

            return (
              item.Category &&
              item.Category
                .toString()
                .trim()
                .toLowerCase() ===
              category
                .toString()
                .trim()
                .toLowerCase()
            );

          }
        );

    }


    /*
     * No records.
     */
    if (
      filteredItems.length === 0
    ) {

      container.innerHTML =
        NewsCentre.noRecord;

      return;

    }


    let allElementsHtml = '';


    filteredItems.forEach(
      (item: INewsItem) => {

        /*
         * =====================================
         * NEWS IMAGE FROM LIST ATTACHMENT
         * =====================================
         *
         * This is the working logic.
         *
         * Image comes from:
         *
         * /Lists/News/Attachments/{ID}/{fileName}
         */

        let imageUrl = '';

        if (item.NewsIcon) {

          try {

            const imgData =
              typeof item.NewsIcon === 'string'
                ? JSON.parse(item.NewsIcon)
                : item.NewsIcon;

            const fileName =
              imgData.fileName || '';

            if (fileName) {

              imageUrl =
                `${this.context.pageContext.web.absoluteUrl}` +
                `/Lists/News/Attachments/` +
                `${item.Id}/${fileName}`;

            }

          } catch (error) {

            console.error(
              '❌ Error parsing NewsIcon:',
              item.NewsIcon,
              error
            );

          }

        }


        console.log(
          'News:',
          item.Title
        );

        console.log(
          'Image:',
          item.NewsIcon
        );

        console.log(
          'Image URL:',
          imageUrl
        );


        /*
         * Date.
         */
        const date =
          this.formatDate(
            item.PublishedDate
          );


        /*
         * Details page.
         */
        const detailsUrl =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/Lists/News/DispForm.aspx?ID=${item.Id}`;


        /*
         * Arrow icon.
         */
        const arrowUrl =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/SiteAssets/resources/images/icons/arrow-right-short.svg`;


        /*
         * Create News HTML.
         */
        const singleElementHtml =
          NewsCentre.singleElementHtml

            .replace(
              /__KEY_URL_IMGICON__/g,
              imageUrl
            )

            .replace(
              /__KEY_DATA_TITLE__/g,
              this.escapeHtml(
                item.Title || ''
              )
            )

            .replace(
              /__KEY_DATA_DESCRIPTION__/g,
              this.escapeHtml(
                item.ShortDescription || ''
              )
            )

            .replace(
              /__KEY_DATA_CATEGORY__/g,
              this.escapeHtml(
                item.Category || ''
              )
            )

            .replace(
              /__KEY_DATA_DATE__/g,
              this.escapeHtml(
                date
              )
            )

            .replace(
              /__KEY_URL_LINK__/g,
              detailsUrl
            )

            .replace(
              /__KEY_URL_ARROW__/g,
              arrowUrl
            );


        allElementsHtml +=
          singleElementHtml;

      }
    );


    /*
     * Put all News items into container.
     */
    container.innerHTML =
      allElementsHtml;

  }


  /*
   * =====================================================
   * GET NEWS ITEMS CONTAINER
   * =====================================================
   *
   * IMPORTANT:
   *
   * Search only inside News Centre.
   *
   * This prevents conflicts with other sections
   * when the complete homepage is merged.
   */
  private getNewsItemsContainer():
    HTMLElement | null {

    if (!this.rootElement) {

      return null;

    }

    return this.rootElement.querySelector(
      '#news-items-container'
    ) as HTMLElement | null;

  }


  /*
   * =====================================================
   * TAB EVENTS
   * =====================================================
   *
   * Existing functionality preserved.
   */
  private attachTabEvents(): void {

    if (!this.rootElement) {

      return;

    }

    const tabs =
      this.rootElement.querySelectorAll(
        '[data-news-category]'
      );


    tabs.forEach(
      (tab: Element) => {

        tab.addEventListener(
          'click',
          () => {

            const category =
              tab.getAttribute(
                'data-news-category'
              );

            if (!category) {

              return;

            }


            /*
             * Remove active class.
             */
            tabs.forEach(
              (item: Element) => {

                item.classList.remove(
                  'tab-title-pill-active'
                );

              }
            );


            /*
             * Set active tab.
             */
            tab.classList.add(
              'tab-title-pill-active'
            );


            /*
             * Filter News.
             */
            this.renderNewsItems(
              category
            );

          }
        );

      }
    );

  }


  /*
   * =====================================================
   * FORMAT DATE
   * =====================================================
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


  /*
   * =====================================================
   * ESCAPE HTML
   * =====================================================
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

}