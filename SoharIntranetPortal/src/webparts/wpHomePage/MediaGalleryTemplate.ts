export default class MediaGalleryTemplate {

  public static singleElementHtml: string = `
    <div class="swiper-slide">
      <div class="gallery-item">
        <img src="__KEY_GALLERY_IMAGE__"
             alt="__KEY_GALLERY_CAPTION__">
      </div>
    </div>
  `;


  public static allElementsHtml: string = `
    <div class="col-12">
      <div class="panel-card px-3 py-4">
        <div class="panel-header w-100 float-start mb-4">
          <h2 class="panel-title">Media Gallery</h2>
          <a href="https://soharaluminium5.sharepoint.com/sites/DevPortal/SitePages/Media-Gallery.aspx" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
            </span><img src="/sites/DevPortal/SiteAssets/resources/images/icons/arrow-right-short.svg" /></a>
        </div>

        <div class="w-100 d-flex flex-column float-start panel-card-gallery">
          <div class="gallery-swiper swiper">
            <div class="swiper-wrapper">
            </div>

            <div class="swiper-button-next swiper-btn-common gallery-slide-next"></div>
            <div class="swiper-button-prev swiper-btn-common gallery-slide-prev"></div>
          </div>

          <div class="gallery-swiper-pagination"></div>
        </div>
      </div>
    </div>
  `;


  public static galleryModalHtml: string = `
    <div class="modal fade" id="galleryModal" tabindex="-1" aria-labelledby="galleryModal" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable gallery-modal">
        <div class="modal-content py-2 px-2">

          <div class="gallery-modal-header d-flex justify-content-end me-2 mt-2">
            <svg class="cursor-pointer flex-shrink-0"
              data-bs-dismiss="modal"
              aria-label="Close"
              width="15"
              height="15"
              viewBox="0 0 15 15"
              fill="none"
              xmlns="http://www.w3.org/2000/svg">
              <path d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016"
                stroke="#005187"
                stroke-width="1.33333"
                stroke-linecap="round"
                stroke-linejoin="round" />
            </svg>
          </div>

          <div class="modal-body gallery-modal-content">

            <div class="w-100 float-start position-relative gallery-swiper-wrapper">

              <div class="w-100 float-start gallery-modal-swiper swiper">
                <div class="swiper-wrapper">
                </div>
              </div>

              <div class="swiper-button-next swiper-btn-common gallery-swiper-modal-next"></div>
              <div class="swiper-button-prev swiper-btn-common gallery-swiper-modal-prev"></div>

            </div>

          </div>

        </div>
      </div>
    </div>
  `;


  public static noRecord: string = `
    <div class="swiper-slide">
      <div class="gallery-item">
        <div class="d-flex justify-content-center align-items-center p-4">
          <h3 class="m-0">No Media Gallery Available</h3>
        </div>
      </div>
    </div>
  `;

}