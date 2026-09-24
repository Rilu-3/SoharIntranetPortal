export default class NewsCentre {

  /*
   * ============================================================
   * SINGLE NEWS ITEM
   * ============================================================
  
   */
  public readonly singleElementHtml: string = `
    <div class="news-item flex-column flex-sm-row">
      <div class="d-flex flex-grow-1 align-items-center gap-3">
        <div class="news-icon">
          <img src="__KEY_URL_IMGICON__" />
        </div>
        <div class="flex-grow-1">
          <p class="news-title">__KEY_DATA_TITLE__</p>
          <p class="news-desc">__KEY_DATA_DESCRIPTION__</p>
          <p class="news-meta mt-1"><span class="tag">__KEY_DATA_CATEGORY__</span> &nbsp;•&nbsp; __KEY_DATA_DATE__</p>
        </div>
      </div>
      <a href="__KEY_URL_LINK__ "  target="_blank" data-interception="off"
        class="link-arrow text-color-link align-self-end align-self-sm-center flex-shrink-0"><span
          class="text-sm xxl-text-base font-bold">Read More</span><img
          src="__KEY_URL_ARROW__" /></a>
    </div>
  `;


  /*
   * ============================================================
   * COMPLETE NEWS CENTRE HTML
   * ============================================================
  
   */
  public readonly allElementsHtml: string = `
<div class="col-12 col-lg-6">
  <div class="panel-card px-2 py-4 d-flex flex-column">
    <div class="panel-header px-2 w-100 float-start mb-4">
      <h2 class="panel-title">News Centre</h2>
      <a href="__KEY_URL_VIEW_ALL__"  target="_blank" data-interception="off" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
          News</span><img src="__KEY_URL_ARROW__" /></a>
    </div>
    <div id="news-tabs" class="w-100 float-start d-flex flex-column flex-grow-1 overflow-hidden">
      <ul class="news-tabs-list px-2">
        <li>
          <div data-tab-news-id="news-panel-all" class="tab-title-pill tab-title-pill-active">All</div>
        </li>
        <li>
          <div data-tab-news-id="news-panel-announcements" class="tab-title-pill">Announcements</div>
        </li>
        <li>
          <div data-tab-news-id="news-panel-events" class="tab-title-pill">Events</div>
        </li>
        <li>
          <div data-tab-news-id="news-panel-news" class="tab-title-pill">News</div>
        </li>
        <li>
          <div data-tab-news-id="news-panel-circulars" class="tab-title-pill">Circulars</div>
        </li>
      </ul>

      <div id="news-panel-all" class="w-100 float-start news-panel-tab-view flex-grow-1"
        style="display: block;">
        <div class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
          __KEY_ALL_ITEMS__
        </div>
      </div>

      <div id="news-panel-announcements" class="w-100 float-start news-panel-tab-view flex-grow-1">
        <div class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
          __KEY_ANNOUNCEMENT_ITEMS__
        </div>
      </div>

    </div>
  </div>
</div>
  `;


  /*
   * ============================================================
   * NO ELEMENT HTML
   * ============================================================
 
   */
  public readonly noElementHtml: string = `
    <div class="news-item">
      <div class="flex-grow-1">
        <p class="news-desc">No records found.</p>
      </div>
    </div>
  `;


  /*
   * ============================================================
   * CONSTRUCTOR
   * ============================================================
  
   */

}