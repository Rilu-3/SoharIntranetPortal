export default class QuickLinks {

  // =========================================================
  // Single Quick Link HTML
  // =========================================================

  public static singleElementHtml: string = `
    <a href="__KEY_URL_LINK__"
      target="__KEY_URL_TARGET__"
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

    <div class="col-12 col-lg-6">
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
    </div>


    <!-- Add Favourite Modal -->

    <div
      class="modal fade favourite-modal"
      id="addFavouriteModal"
      tabindex="-1"
      aria-labelledby="addFavouriteModalLabel"
      aria-hidden="true">

      <div
        class="modal-dialog modal-dialog-centered modal-dialog-scrollable custom-scroll-view">

        <div class="modal-content">

          <div class="modal-header">

            <h5
              class="modal-title"
              id="addFavouriteModalLabel">
              Add to Favourites
            </h5>

            <svg
              class="cursor-pointer flex-shrink-0"
              data-bs-dismiss="modal"
              aria-label="Close"
              width="15"
              height="15"
              viewBox="0 0 15 15"
              fill="none"
              xmlns="http://www.w3.org/2000/svg">

              <path
                d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016"
                stroke="#005187"
                stroke-width="1.33333"
                stroke-linecap="round"
                stroke-linejoin="round">
              </path>

            </svg>

          </div>


          <div class="modal-body">

            <p class="mb-3 modal-body-fav-title">
              Select the quick links you want to add to your favourites.
            </p>

            <div class="favourite-options">

              <!-- Dynamic Favourite Options -->

            </div>

          </div>


          <div class="modal-footer">

            <button
              type="button"
              class="btn-appearance-none btn-brand btn-brand-secondary px-3 py-1 text-sm"
              data-bs-dismiss="modal">
              Cancel
            </button>

            <button
              type="button"
              class="btn-appearance-none btn-brand btn-brand-primary px-3 py-1 text-sm"
              id="btnAddFavourites">
              Add to Favourites
            </button>

          </div>

        </div>

      </div>

    </div>

  `;

}