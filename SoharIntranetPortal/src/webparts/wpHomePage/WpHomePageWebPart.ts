import { Version } from '@microsoft/sp-core-library';

import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'WpHomePageWebPartStrings';

import QuickLinks from './QuickLinks';

import './QuickLinks.module.scss';


export interface IWpHomePageWebPartProps {
  description: string;
}


export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  // =========================================================
  // Render
  // =========================================================

  public async render(): Promise<void> {

    await this.loadCSS();


    this.domElement.innerHTML =
      QuickLinks.allElementsHtml;


    await this.loadJS();


    await QuickLinks.initialize(
      this.context,
      this.domElement
    );

  }


  // =========================================================
  // Load CSS
  // =========================================================

  private async loadCSS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );


    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );


    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );


    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`
    );

  }


  // =========================================================
  // Load JS
  // =========================================================

  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/common.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/home.js`
    );

  }


  // =========================================================
  // Data Version
  // =========================================================

  protected get dataVersion(): Version {

    return Version.parse('1.0');

  }


  // =========================================================
  // Property Pane
  // =========================================================

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {

      pages: [

        {

          header: {

            description:
              strings.PropertyPaneDescription

          },

          groups: [

            {

              groupName:
                strings.BasicGroupName,

              groupFields: [

                PropertyPaneTextField(
                  'description',
                  {

                    label:
                      strings.DescriptionFieldLabel

                  }
                )

              ]

            }

          ]

        }

      ]

    };

  }

}