import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface INewsItem {
  Id: number;
  Title: string;
  NewsIcon?: any;
  ShortDescription?: string;
  Category?: string;
  MainContent?: string;
  PublishedDate?: string;
  Status?: string;
}

export default class NewsCentre {

  private context: WebPartContext;

  private newsItems: INewsItem[] = [];

  /*
   * HTML for ONE news item.
   * This follows the original news-item structure.
   */
  private singleElementHtml: string = `
    <div class="news-item flex-column flex-sm-row">

      <div class="d-flex flex-grow-1 align-items-center gap-3">

        <div class="news-icon">
          <img
            src="__KEY_URL_IMGICON__"
            alt="__KEY_DATA_TITLE__" />
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
        class="link-arrow text-color-link align-self-end align-self-sm-center flex-shrink-0">

        <span class="text-sm xxl-text-base font-bold">
          Read More
        </span>

        <img
          src="__KEY_URL_ARROW__"
          alt="" />

      </a>

    </div>
  `;

  /*
   * COMPLETE News Centre HTML.
   *
   * The structure follows the original HTML.
   * No news-content-* IDs are added.
   */
  private allElementsHtml: string = `
    <div class="col-12 col-lg-6">

      <div class="panel-card px-2 py-4 d-flex flex-column">

        <div class="panel-header px-2 w-100 float-start mb-4">

          <h2 class="panel-title">
            News Centre
          </h2>

          <a
            href="#"
            class="link-arrow text-color-link">

            <span class="text-sm xxl-text-base font-bold">
              View All News
            </span>

            <img
              src="__KEY_URL_ARROW__"
              alt="" />

          </a>

        </div>

        <div
          id="news-tabs"
          class="w-100 float-start d-flex flex-column flex-grow-1 overflow-hidden">

          <ul class="news-tabs-list px-2">

            <li>
              <div
                data-tab-news-id="news-panel-all"
                class="tab-title-pill tab-title-pill-active">
                All
              </div>
            </li>

            <li>
              <div
                data-tab-news-id="news-panel-announcements"
                class="tab-title-pill">
                Announcements
              </div>
            </li>

            <li>
              <div
                data-tab-news-id="news-panel-events"
                class="tab-title-pill">
                Events
              </div>
            </li>

            <li>
              <div
                data-tab-news-id="news-panel-news"
                class="tab-title-pill">
                News
              </div>
            </li>

            <li>
              <div
                data-tab-news-id="news-panel-circulars"
                class="tab-title-pill">
                Circulars
              </div>
            </li>

          </ul>

          <div
            id="news-panel-all"
            class="w-100 float-start news-panel-tab-view flex-grow-1"
            style="display: block;">

            <div
              class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
            </div>

          </div>

          <div
            id="news-panel-announcements"
            class="w-100 float-start news-panel-tab-view flex-grow-1">

            <div
              class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
            </div>

          </div>

          <div
            id="news-panel-events"
            class="w-100 float-start news-panel-tab-view flex-grow-1">

            <div
              class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
            </div>

          </div>

          <div
            id="news-panel-news"
            class="w-100 float-start news-panel-tab-view flex-grow-1">

            <div
              class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
            </div>

          </div>

          <div
            id="news-panel-circulars"
            class="w-100 float-start news-panel-tab-view flex-grow-1">

            <div
              class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
            </div>

          </div>

        </div>

      </div>

    </div>
  `;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Main render method
   */
  public async render(rootElement: HTMLElement): Promise<void> {

    /*
     * Insert the complete original News Centre structure.
     */
    rootElement.innerHTML =
      this.allElementsHtml.replace(
        /__KEY_URL_ARROW__/g,
        this.getArrowImageUrl()
      );

    /*
     * Load SharePoint News list.
     */
    await this.loadNews();

    /*
     * Populate every category.
     */
    this.renderCategory(
      'All',
      'news-panel-all'
    );

    this.renderCategory(
      'Announcements',
      'news-panel-announcements'
    );

    this.renderCategory(
      'Events',
      'news-panel-events'
    );

    this.renderCategory(
      'News',
      'news-panel-news'
    );

    this.renderCategory(
      'Circulars',
      'news-panel-circulars'
    );

    /*
     * Attach tab click events.
     */
    this.attachTabEvents();
  }

  /**
   * Load News items from SharePoint.
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

      const response =
        await fetch(
          apiUrl,
          {
            method: 'GET',

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
        '✅ News items loaded:',
        this.newsItems
      );

    } catch (error) {

      console.error(
        '❌ Error loading News list:',
        error
      );

      this.newsItems = [];
    }
  }

  /**
   * Render a category into the existing
   * .panel-card-news element.
   */
  private renderCategory(
    category: string,
    panelId: string
  ): void {

    const panel =
      document.getElementById(panelId);

    if (!panel) {

      console.warn(
        `⚠️ Panel not found: ${panelId}`
      );

      return;
    }

    /*
     * We use the ORIGINAL .panel-card-news
     * from the HTML.
     *
     * No additional ID is required.
     */
    const container =
      panel.querySelector(
        '.panel-card-news'
      ) as HTMLElement | null;

    if (!container) {

      console.warn(
        `⚠️ .panel-card-news not found inside ${panelId}`
      );

      return;
    }

    /*
     * Filter SharePoint items.
     */
    let filteredItems: INewsItem[];

    if (category === 'All') {

      filteredItems =
        this.newsItems;

    } else {

      filteredItems =
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

    /*
     * No records.
     */
    if (!filteredItems.length) {

      container.innerHTML =
        this.getNoRecordHtml();

      return;
    }

    /*
     * Generate all news items using
     * singleElementHtml.
     */
    const html =
      filteredItems
        .map(
          item =>
            this.createSingleElementHtml(item)
        )
        .join('');

    container.innerHTML =
      html;
  }

  /**
   * Generate ONE news item from
   * singleElementHtml.
   */
  private createSingleElementHtml(
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

    return this.singleElementHtml

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

  /**
   * News image comes from the
   * News list item's attachment.
   */
  private getNewsImageUrl(
    item: INewsItem
  ): string {

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

    return imageUrl;
  }

  /**
   * Existing arrow icon.
   */
  private getArrowImageUrl(): string {

    return (
      `${this.context.pageContext.web.absoluteUrl}` +
      `/SiteAssets/resources/images/icons/arrow-right-short.svg`
    );
  }

  /**
   * Original tab functionality.
   *
   * Uses:
   * data-tab-news-id="news-panel-*"
   */
  private attachTabEvents(): void {

    const tabs =
      document.querySelectorAll(
        '#news-tabs .tab-title-pill'
      );

    const panels =
      document.querySelectorAll(
        '#news-tabs .news-panel-tab-view'
      );

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
             * Hide every panel.
             */
            panels.forEach(
              panel => {

                (
                  panel as HTMLElement
                ).style.display =
                  'none';
              }
            );

            /*
             * Remove active class.
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
             * Show selected panel.
             */
            const targetPanel =
              document.getElementById(
                targetId
              );

            if (targetPanel) {

              (
                targetPanel as HTMLElement
              ).style.display =
                'block';
            }
          }
        );
      }
    );
  }

  /**
   * No records message.
   */
  private getNoRecordHtml(): string {

    return `
      <div class="news-item">

        <div class="flex-grow-1">

          <p class="news-desc">
            No records found.
          </p>

        </div>

      </div>
    `;
  }

  /**
   * Format PublishedDate.
   *
   * Example:
   * May 20, 2024
   */
  private formatDate(
    dateValue?: string
  ): string {

    if (!dateValue) {
      return '';
    }

    const date =
      new Date(dateValue);

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

  /**
   * Normalize category values.
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

  /**
   * Prevent HTML injection from
   * SharePoint text values.
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