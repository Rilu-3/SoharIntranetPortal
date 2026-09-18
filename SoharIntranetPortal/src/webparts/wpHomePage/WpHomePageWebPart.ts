import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import NewsCentre from './NewsCentre';

export interface IWpHomePageWebPartProps {
  description: string;
}

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {

  private newsCentre!: NewsCentre;

  public async onInit(): Promise<void> {

    await super.onInit();

    await this.loadCSS();
    await this.loadJS();
  }

  public render(): void {

    /*
     * =====================================================
     * NEWS CENTRE ROOT
     * =====================================================
     *
     * This WebPart owns the News Centre section.
     *
     * No dependency on UabHomePage.ts.
     */
    this.domElement.innerHTML = `
      <div id="news-centre-root">
      </div>
    `;

    /*
     * =====================================================
     * CREATE NEWS CENTRE
     * =====================================================
     */
    this.newsCentre =
      new NewsCentre(
        this.context
      );

    /*
     * =====================================================
     * RENDER NEWS CENTRE
     * =====================================================
     */
    const newsRoot =
      this.domElement.querySelector(
        '#news-centre-root'
      ) as HTMLElement | null;

    if (!newsRoot) {

      console.error(
        '❌ #news-centre-root not found.'
      );

      return;
    }

    this.newsCentre.render(
      newsRoot
    );
  }

  private async loadCSS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;

    const cssFiles = [
      'bootstrap.min.css',
      'custom.css',
      'font-size.css',
      'home.css',
      'jquery-ui.css',
      'sp-custom.css',
      'swiper-bundle.min.css',
      'variable.css'
    ];

    for (const file of cssFiles) {

      try {

        SPComponentLoader.loadCss(
          `${baseUrl}/SiteAssets/resources/css/${file}`
        );

        console.log(
          `✅ CSS loaded: ${file}`
        );

      } catch (error) {

        console.error(
          `❌ CSS failed: ${file}`,
          error
        );
      }
    }
  }

  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;

    const jsBaseUrl =
      `${baseUrl}/SiteAssets/resources/js`;

    const scripts = [
      'jquery-3.6.0.js',
      'jquery-ui.js',
      'jquery.marquee.min.js',
      'swiper-bundle.min.js',
      'bootstrap.bundle.min.js',
      'common.js',
      'home.js'
    ];

    for (const script of scripts) {

      try {

        await SPComponentLoader.loadScript(
          `${jsBaseUrl}/${script}`
        );

        console.log(
          `✅ JS loaded: ${script}`
        );

      } catch (error) {

        console.error(
          `❌ JS failed: ${script}`,
          error
        );
      }
    }
  }
}