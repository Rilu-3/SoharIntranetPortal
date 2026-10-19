export interface INewsCardView {
  id: number;
  type: string;
  title: string;
  tag: string;          // used as Category
  date: string;         // display date (already formatted)
  published: string;    // raw ISO published date, e.g. "2025-03-14T09:30:00Z"
  src: string;
  fileTypeLabel: string;
  iconUrl: string;
  contentHtml: string;
}
 
export default class wpNewsListing {
 
  private static readonly months: string[] = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
 
  // ---------- Filter helpers ----------
 
  public static filterItemHtml(value: string, label: string): string {
    return `<li><a class="dropdown-item" href="#" data-interception="off" data-value="${value}">${label}</a></li>`;
  }
 
  public static categoryItemHtml(categories: string[]): string {
    return wpNewsListing.filterItemHtml("All", "All") +
      categories.map(c => wpNewsListing.filterItemHtml(c, c)).join("");
  }
 
  public static yearItemHtml(years: number[]): string {
    return wpNewsListing.filterItemHtml("All", "All") +
      years.map(y => wpNewsListing.filterItemHtml(String(y), String(y))).join("");
  }
 
  public static monthItemHtml(): string {
    return wpNewsListing.filterItemHtml("All", "All") +
      wpNewsListing.months.map((m, i) => wpNewsListing.filterItemHtml(String(i + 1), m)).join("");
  }
 
  private static filterDropdownHtml(resourceUrl: string, key: string, label: string, itemsHtml: string): string {
    return `
                  <div class="dropdown sa-filter-dropdown" data-filter="${key}">
                    <button
                      class="sa-btn sa-filter-dropdown-btn dropdown-toggle d-flex align-items-center gap-1 text-color-neutral text-sm font-normal"
                      type="button" id="dropdownMenu${label}" data-bs-toggle="dropdown" aria-expanded="false">
                      <div class="d-flex align-items-center gap-1 flex-shrink-0">
                        <img src="${resourceUrl}/images/icons/filter.png" alt="Filter" />
                        <span>${label}:</span>
                      </div>
                      <div class="flex-grow-1 overflow-hidden">
                        <p class="text-truncate text-start filter-selected-text">All</p>
                      </div>
                    </button>
                    <ul class="dropdown-menu text-sm font-normal" id="${key}FilterList" aria-labelledby="dropdownMenu${label}">
                      ${itemsHtml}
                    </ul>
                  </div>`;
  }
 
  // ---------- Page wrapper ----------
 
  public static wrapperHtml(resourceUrl: string, homeUrl: string, content: string): string {
    return `
      <div class="main-wrapper w-100 float-start min-h-screen-wrapper">
        <div class="container container-sa px-3 px-lg-4 py-3 mx-auto">
          <div class="w-100 float-start announcement-page news-list-page">
 
            <div class="w-100 float-start pt-3 pb-5">
              <div class="d-flex justify-content-between w-100 float-start flex-wrap gap-2">
                <div class="float-start d-flex flex-column gap-2">
                  <div class="breadcrumb-container d-flex">
                    <ol class="breadcrumb">
                      <li class="breadcrumb-item">
                        <a href="${homeUrl}" target="_self" data-interception="off">
                          <svg width="12" height="13" viewBox="0 0 12 13" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path
                              d="M5.03333 0.353023C5.30384 0.125038 5.64623 0 6 0C6.35377 0 6.69616 0.125038 6.96667 0.353023L11.4667 4.14769C11.8047 4.43302 12 4.85236 12 5.29436V11.501C12 11.8104 11.8771 12.1072 11.6583 12.326C11.4395 12.5448 11.1428 12.6677 10.8333 12.6677H8.5C8.19058 12.6677 7.89383 12.5448 7.67504 12.326C7.45625 12.1072 7.33333 11.8104 7.33333 11.501V8.16302C7.33333 8.11882 7.31577 8.07643 7.28452 8.04517C7.25326 8.01392 7.21087 7.99636 7.16667 7.99636H4.83333C4.78913 7.99636 4.74674 8.01392 4.71548 8.04517C4.68423 8.07643 4.66667 8.11882 4.66667 8.16302V11.501C4.66667 11.8104 4.54375 12.1072 4.32496 12.326C4.10617 12.5448 3.80942 12.6677 3.5 12.6677H1.16667C1.01346 12.6677 0.861749 12.6375 0.720203 12.5789C0.578656 12.5203 0.450044 12.4343 0.341709 12.326C0.233374 12.2176 0.147438 12.089 0.0888073 11.9475C0.0301768 11.8059 0 11.6542 0 11.501V5.29436C0 4.85236 0.195333 4.43302 0.533333 4.14769L5.03333 0.353023ZM6.322 1.11769C6.23186 1.04181 6.11782 1.0002 6 1.0002C5.88218 1.0002 5.76814 1.04181 5.678 1.11769L1.178 4.91169C1.12231 4.95857 1.07753 5.01704 1.04679 5.08302C1.01604 5.149 1.00007 5.2209 1 5.29369V11.5004C1 11.5924 1.07467 11.667 1.16667 11.667H3.5C3.5442 11.667 3.5866 11.6495 3.61785 11.6182C3.64911 11.587 3.66667 11.5446 3.66667 11.5004V8.16236C3.66667 7.51769 4.18933 6.99569 4.83333 6.99569H7.16667C7.81067 6.99569 8.33333 7.51769 8.33333 8.16236V11.5004C8.33333 11.5924 8.408 11.667 8.5 11.667H10.8333C10.8775 11.667 10.9199 11.6495 10.9512 11.6182C10.9824 11.587 11 11.5446 11 11.5004V5.29369C10.9999 5.2209 10.984 5.149 10.9532 5.08302C10.9225 5.01704 10.8777 4.95857 10.822 4.91169L6.322 1.11769Z"
                              fill="var(--color-secondary)" />
                          </svg>
                          <span>Home</span>
                        </a>
                      </li>
                      <li class="breadcrumb-item" aria-current="page">Library</li>
                    </ol>
                  </div>
                  <p class="font-bold font-bold text-lg md-text-xl text-color-primary">NEWS</p>
                </div>
 
                <!-- Filters: Category, Year, Month -->
                <div class="d-flex align-items-center gap-2 flex-shrink-0 filter-area-wrapper">
                  ${wpNewsListing.filterDropdownHtml(resourceUrl, "category", "Category", wpNewsListing.categoryItemHtml([]))}
                  ${wpNewsListing.filterDropdownHtml(resourceUrl, "year", "Year", wpNewsListing.yearItemHtml([]))}
                  ${wpNewsListing.filterDropdownHtml(resourceUrl, "month", "Month", wpNewsListing.monthItemHtml())}
                </div>
 
              </div>
              <div class="row gy-4">
 
                <!-- News sidebar -->
                <div class="col-12 col-lg-4 col-xl-3">
                  <aside class="news-sidebar-card" id="newsSidebar" aria-label="News list">
                    <div class="news-sidebar-header">
                      <p class="news-sidebar-title">Latest news</p>
                      <span class="news-count" id="newsCount"></span>
                    </div>
                    <div class="w-100 d-flex flex-column overflow-auto panel-card-news-list custom-scroll-view"
                      id="newsList" role="tablist" aria-orientation="vertical" aria-label="News items">
                      ${content}
                    </div>
                  </aside>
                </div>
 
                <!-- Selected news content -->
                <div class="col-12 col-lg-8 col-xl-9">
                  <section class="news-detail-card" id="newsDetail" role="tabpanel" tabindex="-1">
                    <div class="news-detail-header">
                      <div class="flex-grow-1 min-w-0">
                        <button type="button" class="news-back-btn d-lg-none">
                          <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                            <path d="M10 3L5 8l5 5" stroke="currentColor" stroke-width="2" stroke-linecap="round"
                              stroke-linejoin="round" />
                          </svg>
                          <span>Back to list</span>
                        </button>
                        <h2 class="news-detail-title" id="newsDetailTitle" aria-live="polite"></h2>
                        <p class="news-meta mt-1">
                          <span class="tag" id="newsDetailTag"></span><span id="newsDetailSep">&nbsp;•&nbsp;</span><span
                            id="newsDetailDate"></span>
                        </p>
                      </div>
                      <div class="news-detail-actions">
                        <a class="news-action-btn d-none" id="newsOpenLink" href="#" target="_blank" rel="noopener">
                          <svg viewBox="0 0 16 16" fill="none" aria-hidden="true">
                            <path d="M9 2h5v5M14 2L7 9M12 9.5V13a1 1 0 0 1-1 1H3a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h3.5"
                              stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" />
                          </svg>
                          <span>Open</span>
                        </a>
                        <!--
                        <a class="news-action-btn news-action-btn--primary d-none" id="newsDownloadLink" href="#" download>
                          <svg viewBox="0 0 16 16" fill="none" aria-hidden="true">
                            <path d="M8 2v8M4.5 6.5L8 10l3.5-3.5M2.5 13.5h11" stroke="currentColor" stroke-width="1.6"
                              stroke-linecap="round" stroke-linejoin="round" />
                          </svg>
                          <span>Download</span>
                        </a>
                        -->
                      </div>
                    </div>
                    <p class="news-viewer-hint d-none" id="newsViewerHint">
                      Preview not showing on your device? Select <strong>Open</strong> to view the document in a new tab.
                    </p>
                    <div class="news-detail-stage">
                      <div class="news-detail-body" id="newsDetailBody"></div>
                      <div class="news-detail-loader d-none" id="newsDetailLoader">
                        <span class="news-spinner" aria-hidden="true"></span>
                        <span class="visually-hidden">Loading content</span>
                      </div>
                    </div>
                  </section>
                </div>
 
              </div>
            </div>
 
          </div>
        </div>
      </div>`;
  }
 
  // Loading state (Bootstrap spinner) - shown inside the sidebar list
  public static loadingHtml: string = `
    <div class="d-flex justify-content-center py-5">
      <div class="spinner-border text-secondary" role="status">
        <span class="visually-hidden">Loading...</span>
      </div>
    </div>`;
 
  // No records - shown inside the sidebar list
  public static noRecordHtml: string = `
    <div class="p-3">
      <p class="text-color-desc text-sm">No records found</p>
    </div>`;
 
  // Single sidebar news item (values must already be escaped)
  public static cardHtml(data: INewsCardView): string {
    const srcAttr: string = data.src ? ` data-src="${data.src}"` : "";
    const articleTemplate: string = data.type === 'article'
      ? `
                    <!-- Article body rendered in the main content area when this item is selected -->
                    <template class="news-item-content">${data.contentHtml}</template>`
      : "";
 
    return `
                  <div class="news-list-item" id="news-${data.id}" role="tab" aria-controls="newsDetail"
                    data-type="${data.type}" data-title="${data.title}"
                    data-tag="${data.tag}" data-date="${data.date}"
                    data-published="${data.published}"${srcAttr}>
                    <div class="d-flex flex-grow-1 align-items-center gap-3 min-w-0">
                      <div class="news-icon">
                        <img src="${data.iconUrl}" alt="">
                      </div>
                      <div class="flex-grow-1 min-w-0">
                        <p class="news-title line-clamp-2">${data.title}</p>
                        <span class="news-file-type text-sm xxl-text-base font-bold">${data.fileTypeLabel}</span>
                        <p class="news-meta mt-1"><span class="tag">${data.tag}</span> &nbsp;•&nbsp; ${data.date}</p>
                      </div>
                    </div>${articleTemplate}
                  </div>`;
  }
}