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


interface IBannerItem {
  Id: number;
  Title: string;
  Description: string;
  Status: string;
  SortOrder: number;
  Image: any;
}


interface IMediaGalleryItem {
  Id: number;
  Title: string;
  Caption: string;
  Status: string;
  SortOrder: number;
  Image: any;
}


export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  public render(): void {

    this.domElement.innerHTML =
      BannerTemplate.bannerHtml +
      MediaGalleryTemplate.allElementsHtml;

    this._getBannerItems();

    this._getMediaGalleryItems();

  }


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

      console.log(
        'Banner Items:',
        bannerItems
      );

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


  private _renderBanner(
    bannerItems: IBannerItem[]
  ): void {

    let allElementsHtml: string = '';

    bannerItems.forEach(
      (item: IBannerItem) => {

        let imageData: any = {};

        if (item.Image) {
          imageData =
            typeof item.Image === 'string'
              ? JSON.parse(item.Image)
              : item.Image;
        }

        const fileName =
          imageData.fileName || '';

        const imageUrl =
          `${this.context.pageContext.web.absoluteUrl}/Lists/Banner/Attachments/${item.Id}/${fileName}`;

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


  private _renderMediaGallery(
    galleryItems: IMediaGalleryItem[]
  ): void {

    let allElementsHtml: string = '';

    galleryItems.forEach(
      (item: IMediaGalleryItem) => {

        let imageData: any = {};

        if (item.Image) {
          imageData =
            typeof item.Image === 'string'
              ? JSON.parse(item.Image)
              : item.Image;
        }

        const fileName =
          imageData.fileName || '';

        const imageUrl =
          `${this.context.pageContext.web.absoluteUrl}/Lists/Media_Gallery/Attachments/${item.Id}/${fileName}`;

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


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }


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


  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
    );

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

    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/home.js`
    );

  }


  protected async onInit(): Promise<void> {

    this.loadCSS();

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