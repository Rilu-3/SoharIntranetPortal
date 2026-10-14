import { Version } from '@microsoft/sp-core-library';

import {

  type IPropertyPaneConfiguration

} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SPHttpClient } from '@microsoft/sp-http';

import { SPComponentLoader } from '@microsoft/sp-loader';

import { BannerTemplate } from './BannerTemplate';

import MediaGalleryTemplate from './MediaGalleryTemplate';

export interface IWpHomePageWebPartProps {

  description: string;

}

interface ISharePointImage {
  serverRelativeUrl?: string;
  serverUrl?: string;
  fileName?: string;
  [key: string]: unknown;
}

interface IBannerItem {

  Id: number;

  Title: string;

  Description: string;

  Status: string;

  SortOrder: number;

  Image?: string | ISharePointImage;

}

interface IMediaGalleryItem {

  Id: number;

  Title: string;

  Caption: string;

  Status: string;

  SortOrder: number;

  Image?: string | ISharePointImage;

}

export default class WpHomePageWebPart

  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {

  // ==================== RENDER ====================



  public render(): void {

    const workbenchContent = document.getElementById('workbenchPageContent');

    if (workbenchContent) {

      workbenchContent.style.maxWidth = 'none';

    }

    // Render the Banner, Media Gallery and Gallery Modal HTML



    this.domElement.innerHTML =

      BannerTemplate.bannerHtml +

      MediaGalleryTemplate.allElementsHtml +

      MediaGalleryTemplate.galleryModalHtml;



    // Load Banner and Media Gallery data first. Initialize both Swipers
    // only after their corresponding HTML slides have been rendered.
    void Promise.all([
      this._getBannerItems(),
      this._getMediaGalleryItems()
    ]).then(() => {
      this.initBannerSwiper();
      this.initMediaGallery();
    }).catch((error: unknown) => {
      console.error('Error initializing Banner / Media Gallery:', error);
    });

  }

  // ==================== SHAREPOINT IMAGE FIELD HELPER ====================

  private getImageData(image?: string | ISharePointImage): ISharePointImage {
    if (!image) {
      return {};
    }

    if (typeof image === 'string') {
      try {
        return JSON.parse(image) as ISharePointImage;
      } catch (error) {
        console.warn('Unable to parse SharePoint image field:', error);
        return {};
      }
    }

    return image;
  }

  // ==================== BANNER SWIPER INITIALIZATION ====================

  private initBannerSwiper(): void {
    const Swiper = (window as any).Swiper;
    if (!Swiper) { console.warn('Swiper is not available.'); return; }

    new Swiper(".banner-swiper", {
      pagination: { el: ".swiper-pagination", clickable: true },
      navigation: { nextEl: ".banner-slide-next", prevEl: ".banner-slide-prev" },
      speed: 1000,autoplay: {delay: 3000,disableOnInteraction: false}
    });
  }

  // ==================== MEDIA GALLERY SWIPER + MODAL INITIALIZATION ====================

  private initMediaGallery(): void {
    const $ = (window as any).soharJQuery || (window as any).jQuery;
    const Swiper = (window as any).soharSwiper || (window as any).Swiper;

    if (!$ || !Swiper) {
      console.warn('jQuery or Swiper is not available.');
      return;
    }

    const galleryElement = this.domElement.querySelector('.gallery-swiper') as HTMLElement;

    if (!galleryElement) {
      console.warn('Gallery Swiper element not found.');
      return;
    }

    const gallerySwiper = new Swiper(galleryElement, {
      slidesPerView: 1,
      spaceBetween: 12,
      loop: false,
      speed: 500,
      observer: true,
      observeParents: true,
      pagination: { el: '.gallery-swiper-pagination', clickable: true },
      navigation: { nextEl: '.gallery-slide-next', prevEl: '.gallery-slide-prev' },
      breakpoints: {
        400: { slidesPerView: 2, spaceBetween: 12 },
        576: { slidesPerView: 2, spaceBetween: 14 },
        768: { slidesPerView: 3, spaceBetween: 16 },
        992: { slidesPerView: 4, spaceBetween: 16 }
      }
    });

    // Recalculate once everything (images, social widgets, loader) has finished loading
    if (document.readyState === 'complete') {
      setTimeout(() => gallerySwiper.update(), 300);
    } else {
      window.addEventListener('load', () => gallerySwiper.update());
    }

    // Bootstrap is only needed for the popup, so it does not block the slider
    const bootstrap = (window as any).bootstrap;
    if (!bootstrap) {
      console.warn('Bootstrap is not available.');
      return;
    }

    const modalElement = this.domElement.querySelector('#galleryModal') as HTMLElement;

    if (!modalElement) {
      return;
    }

    const modal = bootstrap.Modal.getOrCreateInstance(modalElement);

    let swiperGalleryModal: any = null;
    let gallery_slider_index = 0;

    function initGalleryModalSwiper(): void {
      // Destroy any previous instance before creating a new one
      if (swiperGalleryModal) {
        swiperGalleryModal.destroy(true, true);
        swiperGalleryModal = null;
      }

      swiperGalleryModal = new Swiper('.gallery-modal-swiper', {
        spaceBetween: 10,
        slidesPerView: 1,
        autoplay: false,
        speed: 1400,
        autoHeight: true,
        initialSlide: gallery_slider_index,
        navigation: { nextEl: '.gallery-swiper-modal-next', prevEl: '.gallery-swiper-modal-prev' }
      });

      $('.gallery-modal-content').css('opacity', 1);
    }

    // Set the index BEFORE the modal opens
    $(document).on('click', '.gallery-item img, .gallery-item video', function (this: HTMLElement): void {
      $('.gallery-modal-content').css('opacity', 0);
      gallery_slider_index = $(this).closest('.swiper-slide').index();
      modal.show();
    });

    // shown.bs.modal fires AFTER the transition — Swiper can measure dimensions safely here
    modalElement.addEventListener('shown.bs.modal', () => {
      initGalleryModalSwiper();
    });

    // hidden.bs.modal fires AFTER the modal is fully hidden
    modalElement.addEventListener('hidden.bs.modal', () => {
      if (swiperGalleryModal) {
        swiperGalleryModal.destroy(true, true);
        swiperGalleryModal = null;
      }
    });
  }

  // ==================== GET BANNER ITEMS ====================

  private async _getBannerItems(): Promise<void> {

    try {

      const siteUrl =

        this.context.pageContext.web.absoluteUrl;

      const url =

        `${siteUrl}/_api/web/lists/getbytitle('Banner')/items?$select=Id,Title,Description,Status,SortOrder,Image&$orderby=SortOrder asc`;

      const response =

        await this.context.spHttpClient.get(

          url,

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

          `Banner list request failed: ${response.status}`

        );

      }

      const data =

        await response.json();

      const bannerItems: IBannerItem[] =

        data.value

          .filter(

            (item: IBannerItem) =>

              item.Status === 'Active'

          )

          .sort(

            (a: IBannerItem, b: IBannerItem) =>

              a.SortOrder - b.SortOrder

          );

      console.log('Banner Items:', bannerItems);

      this._renderBanner(

        bannerItems

      );

    } catch (error) {

      console.error(

        'Error loading Banner list:',

        error

      );

      const divBanner =

        this.domElement.querySelector(

          '#divBanner'

        );

      if (divBanner !== null) {

        divBanner.innerHTML =

          BannerTemplate.noRecord;

      }

    }

  }

  // ==================== RENDER BANNER ====================

  private _renderBanner(

    bannerItems: IBannerItem[]

  ): void {

    let allElementsHtml: string = '';

    bannerItems.forEach(

      (item: IBannerItem) => {

        const imageData = this.getImageData(item.Image);

        const defaultImageUrl =

          '/sites/DevPortal/SiteAssets/resources/images/bannerDefault/banner-1.png';



        const imageUrl =

          imageData?.serverRelativeUrl || defaultImageUrl;



        const singleElementHtml =

          BannerTemplate.singleElementHtml

            .replace(

              '__KEY_BANNER_IMAGE__',

              imageUrl

            )

            .replace(

              /__KEY_BANNER_TITLE__/g,

              item.Title || ''

            )

            .replace(

              '__KEY_BANNER_DESCRIPTION__',

              item.Description || ''

            );

        allElementsHtml +=

          singleElementHtml;

      }

    );

    if (allElementsHtml === '') {

      allElementsHtml =

        BannerTemplate.noRecord;

    }

    const divBanner =

      this.domElement.querySelector(

        '#divBanner'

      );

    if (divBanner !== null) {

      divBanner.innerHTML =

        allElementsHtml;

    }

  }

  // ==================== GET MEDIA GALLERY ITEMS ====================

  private async _getMediaGalleryItems(): Promise<void> {

    try {

      const siteUrl =

        this.context.pageContext.web.absoluteUrl;

      const url =

        `${siteUrl}/_api/web/lists/getbytitle('Media_Gallery')/items?$select=Id,Title,Caption,Status,SortOrder,Image&$orderby=SortOrder asc`;

      const response =

        await this.context.spHttpClient.get(

          url,

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

          `Media Gallery list request failed: ${response.status}`

        );

      }

      const data =

        await response.json();

      const galleryItems: IMediaGalleryItem[] =

        data.value

          .filter(

            (item: IMediaGalleryItem) =>

              item.Status === 'Active'

          )

          .sort(

            (a: IMediaGalleryItem, b: IMediaGalleryItem) =>

              a.SortOrder - b.SortOrder

          );

      console.log(

        'Media Gallery Items:',

        galleryItems

      );

      this._renderMediaGallery(

        galleryItems

      );

    } catch (error) {

      console.error(

        'Error loading Media Gallery list:',

        error

      );

      const galleryWrapper =

        this.domElement.querySelector(

          '.gallery-swiper .swiper-wrapper'

        );

      if (galleryWrapper !== null) {

        galleryWrapper.innerHTML =

          MediaGalleryTemplate.noRecord;

      }

    }

  }

  // ==================== RENDER MEDIA GALLERY ====================

  private _renderMediaGallery(

    galleryItems: IMediaGalleryItem[]

  ): void {

    let allElementsHtml: string = '';

    galleryItems.forEach(

      (item: IMediaGalleryItem) => {

        const imageData = this.getImageData(item.Image);

        const imageUrl =

          imageData?.serverRelativeUrl || '';



        console.log(

          'Gallery Image URL:',

          imageUrl

        );

        const singleElementHtml =

          MediaGalleryTemplate.singleElementHtml

            .replace(

              '__KEY_GALLERY_IMAGE__',

              imageUrl

            )

            .replace(

              '__KEY_GALLERY_CAPTION__',

              item.Caption || item.Title || ''

            );

        allElementsHtml +=

          singleElementHtml;

      }

    );

    if (allElementsHtml === '') {

      allElementsHtml =

        MediaGalleryTemplate.noRecord;

    }

    const galleryWrapper =

      this.domElement.querySelector(

        '.gallery-swiper .swiper-wrapper'

      );

    if (galleryWrapper !== null) {

      galleryWrapper.innerHTML =

        allElementsHtml;

      /*

       * Fill Modal Gallery

       */

      const modalWrapper =

        this.domElement.querySelector(

          '.gallery-modal-swiper .swiper-wrapper'

        );

      if (modalWrapper !== null) {

        let modalElementsHtml: string = '';

        galleryItems.forEach(

          (item: IMediaGalleryItem) => {

            const imageData = this.getImageData(item.Image);

            const imageUrl =

              imageData?.serverRelativeUrl || '';



            modalElementsHtml += `

              <div class="swiper-slide gallery-swiper-slide">

                <img

                  src="${imageUrl}"

                  alt="${item.Caption || item.Title || ''}"

                />

              </div>

            `;

          }

        );

        modalWrapper.innerHTML =

          modalElementsHtml;

      }

      const galleryElement =

        this.domElement.querySelector(

          '.gallery-swiper'

        ) as HTMLElement & {

          swiper?: any;

        };

      if (

        galleryElement &&

        galleryElement.swiper

      ) {

        galleryElement.swiper.update();

      }

    }

  }

  // ==================== LOAD BOOTSTRAP ====================



  private async loadBootstrap(): Promise<void> {



    const baseUrl = this.context.pageContext.web.absoluteUrl;

    const win: any = window;



    const bootstrap = await SPComponentLoader.loadScript<any>(

      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`

    );



    if (bootstrap?.Modal) {

      win.bootstrap = bootstrap;



      const original = bootstrap.Modal.getOrCreateInstance;



      if (typeof original === 'function') {

        bootstrap.Modal.getOrCreateInstance = function (

          element: Element,

          config?: any

        ) {

          return original.call(bootstrap.Modal, element, config);

        };

      }

    }



    console.log('Bootstrap loaded:', typeof win.bootstrap);

  }

  // ==================== LOAD HOME JS ====================

  private async loadHomeJS(): Promise<void> {

    const baseUrl =

      this.context.pageContext.web.absoluteUrl;

    try {

      await SPComponentLoader.loadScript(

        `${baseUrl}/SiteAssets/resources/js/home.js`

      );

      console.log(

        'home.js loaded'

      );

    } catch (error) {

      console.error(

        'Error loading home.js:',

        error

      );

    }

  }

  protected get dataVersion(): Version {

    return Version.parse('1.0');

  }

  // ==================== LOAD CSS ====================

  private loadCSS(): void {

    const baseUrl =

      this.context.pageContext.web.absoluteUrl;

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/variable.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/font-size.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/custom.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/home.css`

    );

    SPComponentLoader.loadCss(

      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`

    );

  }

  // ==================== LOAD JS ====================

  private async loadJS(): Promise<void> {

    const baseUrl =

      this.context.pageContext.web.absoluteUrl;

    await SPComponentLoader.loadScript(

      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`

    );

    await this.loadBootstrap();

    await SPComponentLoader.loadScript(

      `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`

    );

    await SPComponentLoader.loadScript(

      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`

    );

    await SPComponentLoader.loadScript(

      `${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`

    );

    await SPComponentLoader.loadScript(

      `${baseUrl}/SiteAssets/resources/js/common.js`

    );

  }

  // ==================== INITIALIZE WEB PART ====================

  protected async onInit(): Promise<void> {

    this.loadCSS();
    this.loadHomeJS();

    await this.loadJS();

    return super.onInit();

  }

  protected getPropertyPaneConfiguration():

    IPropertyPaneConfiguration {

    return {

      pages: []

    };

  }

}