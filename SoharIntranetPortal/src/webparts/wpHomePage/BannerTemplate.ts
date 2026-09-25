export class BannerTemplate {

  public static singleElementHtml: string = `

    <div class="swiper-slide">

      <div class="w-100 float-start banner-slide-grid">

        <img src="__KEY_BANNER_IMAGE__" alt="__KEY_BANNER_TITLE__" />

        <div class="banner-slide-info">

          <div class="w-100 text-white d-flex flex-column gap-2 gap-xl-3 px-3 px-lg-4">

            <h4 class="font-bold fhd-text-size-48-combo text-capitalize lh-3 mb-0 line-clamp-3">
              __KEY_BANNER_TITLE__
            </h4>

            <p class="font-light text-base lg-text-lg lh-sm line-clamp-3">
              __KEY_BANNER_DESCRIPTION__
            </p>

          </div>

        </div>

      </div>

    </div>

  `;


  public static bannerHtml: string = `

 

          <div class="col-12">

            <div class="w-100 float-start swiper banner-swiper">

              <div class="swiper-wrapper" id="divBanner"></div>

              <div class="swiper-pagination"></div>

              <div class="swiper-button-next banner-slide-nav-btn banner-slide-next me-3 me-lg-4"></div>

              <div class="swiper-button-prev banner-slide-nav-btn banner-slide-prev ms-3 ms-lg-4"></div>

            </div>

          </div>



  `;


  public static noRecord: string = `

    <div class="swiper-slide">

      <div class="d-flex justify-content-center align-items-center p-4">

        <h3 class="m-0">No Banner Available</h3>

      </div>

    </div>

  `;

}