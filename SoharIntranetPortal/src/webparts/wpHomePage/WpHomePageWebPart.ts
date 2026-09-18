import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import UabHomePage from './UabHomePage';
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
     * Render your existing homepage.
     */
    this.domElement.innerHTML =
      UabHomePage.allElementsHtml;

    /*
     * NewsCentre works with the #news-tabs
     * already present in UabHomePage.
     */
    this.newsCentre =
      new NewsCentre(
        this.context
      );

    this.newsCentre.render();
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