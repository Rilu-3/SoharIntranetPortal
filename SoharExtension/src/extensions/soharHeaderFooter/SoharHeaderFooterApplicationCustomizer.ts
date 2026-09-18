import { override } from '@microsoft/decorators';

import { Log } from '@microsoft/sp-core-library';

import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { SPComponentLoader } from '@microsoft/sp-loader';

import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'SoharHeaderFooterApplicationCustomizerStrings';

import Header from './Header';
import Footer from './Footer';

const LOG_SOURCE: string =
  'SoharHeaderFooterApplicationCustomizer';

export interface ISoharHeaderFooterApplicationCustomizerProperties {
  testMessage: string;
}

export default class SoharHeaderFooterApplicationCustomizer
  extends BaseApplicationCustomizer<ISoharHeaderFooterApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public async onInit(): Promise<void> {

    Log.info(
      LOG_SOURCE,
      `Initialized ${strings.Title}`
    );

    this._loadCSS();

    this._loadJS();

    await this._renderHeader();

    this._renderFooter();

    return Promise.resolve();
  }

  private _loadCSS(): void {

    const baseUrl: string =
      this.context.pageContext.web.absoluteUrl;

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/font-size.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );
  }

  private _loadJS(): void {

    const baseUrl: string =
      this.context.pageContext.web.absoluteUrl;

    SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
    );
  }

  private async _getUserDesignation(
    userEmail: string
  ): Promise<string> {

    const siteUrl: string =
      this.context.pageContext.web.absoluteUrl;

    const url =
      `${siteUrl}/_api/web/lists/getbytitle('PRT Master Users')/items` +
      `?$select=Designation` +
      `&$filter=User_x0020_Email eq '${userEmail}'`;

    try {

      const response =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata'
            }
          }
        );

      if (!response.ok) {

        throw new Error(
          `Failed to load designation: ${response.status} ${response.statusText}`
        );
      }

      const data =
        await response.json();

      console.log(
        'User designation data:',
        data.value
      );

      if (
        data.value &&
        data.value.length > 0
      ) {

        return data.value[0].Designation || '';
      }

      return '';

    } catch (error) {

      console.error(
        'Designation loading error:',
        error
      );

      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(String(error))
      );

      return '';
    }
  }

  private async _getDepartments(): Promise<string> {

    const siteUrl: string =
      this.context.pageContext.web.absoluteUrl;

    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Departments')/items?$select=Title,Link`;

    try {

      const response =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json;odata=nometadata'
            }
          }
        );

      if (!response.ok) {

        throw new Error(
          `Failed to load Departments: ${response.status} ${response.statusText}`
        );
      }

      const data =
        await response.json();

      console.log(
        'Departments data:',
        data.value
      );

      const departmentItems =
        data.value.map(
          (department: any) => {

            const link =
              department.Link?.Url || '#';

            return `
              <li>
                <a
                  class="dropdown-item text-sm"
                  href="${link}">
                  ${department.Title}
                </a>
              </li>
            `;
          }
        ).join('');

      return departmentItems;

    } catch (error) {

      console.error(
        'Department list error:',
        error
      );

      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(String(error))
      );

      return '';
    }
  }

  private async _renderHeader(): Promise<void> {

    if (!this._topPlaceholder) {

      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          {
            onDispose: () => {

              Log.info(
                LOG_SOURCE,
                'Top placeholder disposed'
              );

            }
          }
        );
    }

    if (!this._topPlaceholder) {

      Log.error(
        LOG_SOURCE,
        new Error('Top placeholder was not created')
      );

      return;
    }

    const siteUrl: string =
      this.context.pageContext.web.absoluteUrl;

    const userName: string =
      this.context.pageContext.user.displayName;

    const userEmail: string =
      this.context.pageContext.user.email;

    const profilePhoto: string =
      `${siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(userEmail)}`;

    const designation: string =
      await this._getUserDesignation(userEmail);

    const departmentItems: string =
      await this._getDepartments();

    const header: Header =
      new Header(
        siteUrl,
        userName,
        designation,
        profilePhoto,
        departmentItems
      );

    this._topPlaceholder.domElement.innerHTML =
      header.render();
  }

  private _renderFooter(): void {

    if (!this._bottomPlaceholder) {

      this._bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          {
            onDispose: () => {

              Log.info(
                LOG_SOURCE,
                'Bottom placeholder disposed'
              );

            }
          }
        );
    }

    if (!this._bottomPlaceholder) {

      Log.error(
        LOG_SOURCE,
        new Error('Bottom placeholder was not created')
      );

      return;
    }

    const siteUrl: string =
      this.context.pageContext.web.absoluteUrl;

    const footer: Footer =
      new Footer(siteUrl);

    this._bottomPlaceholder.domElement.innerHTML =
      footer.render();
  }
}