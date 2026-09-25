export default class QuickLinks {

  // =========================================================
  // Single Quick Link HTML
  // =========================================================

  public static singleElementHtml: string = `
    <a href="__KEY_URL_LINK__"
       target="_blank" data-interception="off"
      class="quick-link-box"
      __KEY_DATA_FAVOURITE__>

      <img
        src="__KEY_URL_IMGICON__"
        alt="__KEY_DATA_TITLE__">

      <span>__KEY_DATA_TITLE__</span>

    </a>
  `;


  // =========================================================
  // Full Quick Links Section
  // =========================================================

  public static allElementsHtml: string = `

    
      <div class="panel-card px-2 py-4">

        <div class="panel-header px-2 w-100 float-start pb-3 d-flex align-items-center">

          <div class="d-flex align-items-center flex-wrap gap-2 gap-sm-3">

            <h2
              data-filter-ql="quick-links"
              class="panel-title panel-title-filter panel-title-filter-active">
              Quick Links
            </h2>

            <h2
              data-filter-ql="favourites"
              class="panel-title panel-title-filter">
              Favourites
            </h2>

          </div>

          <button
            type="button"
            class="btn-appearance-none btn-add-favourite"
            title="Add to Favourites"
            data-bs-toggle="modal"
            data-bs-target="#addFavouriteModal">

            <svg xmlns="http://www.w3.org/2000/svg"
              width="20"
              height="20"
              fill="#000"
              viewBox="0 0 16 16">

              <path
                fill-rule="evenodd"
                d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2" />

            </svg>

          </button>

        </div>


        <div
          class="w-100 float-start p-2 overflow-auto panel-card-quick-links quick-links-grid custom-scroll-view">

          <!-- Dynamic Quick Links will be inserted here -->

        </div>

      </div>
   


   
  `;

}