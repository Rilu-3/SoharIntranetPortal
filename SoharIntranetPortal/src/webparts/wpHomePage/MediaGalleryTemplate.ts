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
          <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
            </span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
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