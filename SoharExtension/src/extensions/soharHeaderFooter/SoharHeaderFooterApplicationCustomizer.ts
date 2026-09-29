import { override } from '@microsoft/decorators';

import { Log } from '@microsoft/sp-core-library';

import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,

} from '@microsoft/sp-application-base';

import {
  MSGraphClientV3
} from '@microsoft/sp-http';

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
    this._loadCSS();

    Log.info(
      LOG_SOURCE,
      `Initialized ${strings.Title}`
    );


    // this._loadJS();
     await this.loadBootstrap();
    await this._renderHeader();

    await this._renderFooter();

    return Promise.resolve();
    
    
  }


  private _loadHome(): void {
    const baseUrl: string =
      this.context.pageContext.web.absoluteUrl;

    SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/common.js`
    )
      .then(() =>
        SPComponentLoader.loadScript(
          `${baseUrl}/SiteAssets/resources/js/home.js`
        )
      );
  }
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

}
  private _loadCSS(): void {
    const baseUrl: string =
      this.context.pageContext.web.absoluteUrl;

    // Load all CSS files in parallel
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/font-size.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`
    );

    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );

    // Load scripts sequentially with proper dependency handling
    SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    )
      .then(() =>
        SPComponentLoader.loadScript(
          `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
        )
      )
      .then(() =>
        SPComponentLoader.loadScript(
          `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
        )
      )
      //.then(() => SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`))
      .then(() =>
        SPComponentLoader.loadScript(
          `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`
        )
      )
      .then(() => this._loadHome());
  }



  // ============================================================
  // USER DESIGNATION FROM MICROSOFT ENTRA ID
  // ============================================================

  private async _getUserDesignation(): Promise<string> {

    try {

      const client: MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient('3');

      const user =
        await client
          .api('/me')
          .select('jobTitle')
          .get();


      return user.jobTitle || '';

    } catch (error) {

      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(String(error))
      );

      return '';
    }
  }


  // ============================================================
  // DEPARTMENTS
  // ============================================================

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


      const departmentItems =
        data.value.map(
          (department: any) => {

            const link =
              department.Link?.Url || '#';


            return `
              <li>
                <a
              class="dropdown-item text-sm"
              href="${link}"
              target="_blank"
              data-interception="off">
              ${department.Title}
            </a>
              </li>
            `;
          }
        ).join('');


      return departmentItems;

    } catch (error) {

      Log.error(
        LOG_SOURCE,
        error instanceof Error
          ? error
          : new Error(String(error))
      );

      return '';
    }
  }


  // ============================================================
  // RENDER HEADER
  // ============================================================

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
      await this._getUserDesignation();


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

    this._setupSearchFunctionality();

  }


  private _setupSearchFunctionality(): void {

    if (
      !this._topPlaceholder ||
      !this._topPlaceholder.domElement
    ) {
      return;
    }


    const inputMainSearchBox =
      this._topPlaceholder.domElement.querySelector(
        '#inputGlobalSearchBox'
      ) as HTMLInputElement;


    if (inputMainSearchBox) {

      inputMainSearchBox.addEventListener(
        'keydown',
        (event: KeyboardEvent) => {

          if (event.key === 'Enter') {

            const searchKey =
              inputMainSearchBox.value;


            if (searchKey) {

              window.open(
                `/sites/DevPortal/_layouts/15/search.aspx/siteall?q=${encodeURIComponent(searchKey)}`,
                '_blank'
              );
            }


            event.preventDefault();
          }
        }
      );
    }
  }


  // ============================================================
  // GET FOOTER ITEMS FROM SHAREPOINT
  // ============================================================

  private async _getFooterItems(): Promise<any[]> {

    const siteUrl =
      this.context.pageContext.web.absoluteUrl;


    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Footer')/items` +
      `?$select=Id,ContactDetails,Category,Status,Icon` +
      `&$orderby=Id asc`;


    try {

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
          `Failed to load Footer: ${response.status}`
        );
      }


      const data =
        await response.json();


      return data.value || [];


    } catch (error) {

      return [];
    }
  }


  // ============================================================
  // RENDER FOOTER ITEMS
  // ============================================================

  private _renderFooterItems(items: any[]): string {

    return items.map((item: any) => {

      const contactDetails =
        item.ContactDetails || '';


      let imageUrl = '';


      if (item.Icon) {

        const imgData =
          JSON.parse(item.Icon);


        imageUrl =
          `${this.context.pageContext.web.absoluteUrl}/Lists/Footer/Attachments/${item.Id}/${imgData.fileName}`;

      }


      return Footer.itemTemplate
        .replace(
          '__ICON_URL__',
          imageUrl
        )
        .replace(
          '__CONTACT_DETAILS__',
          contactDetails
        );

    }).join('');
  }


  // ============================================================
  // RENDER FOOTER
  // ============================================================

  private async _renderFooter(): Promise<void> {

    // Create Bottom Placeholder
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


    // Check Bottom Placeholder
    if (!this._bottomPlaceholder) {

      Log.error(
        LOG_SOURCE,
        new Error('Bottom placeholder was not created')
      );

      return;
    }


    // Get Footer Items from SharePoint
    const footerItems =
      await this._getFooterItems();


    // Get only Active items
    const activeItems =
      footerItems.filter(
        (item: any) =>
          String(item.Status || '')
            .trim()
            .toLowerCase() === 'active'
      );


    // Important Contacts
    const importantContacts =
      activeItems.filter(
        (item: any) =>
          String(item.Category || '')
            .trim()
            .toLowerCase() ===
          'important contacts'
      );


    // Medical Services
    const medicalServices =
      activeItems.filter(
        (item: any) =>
          String(item.Category || '')
            .trim()
            .toLowerCase() ===
          'medical services'
      );


    // Render Important Contacts Items
    const importantContactsHtml =
      this._renderFooterItems(
        importantContacts
      );


    // Render Medical Services Items
    const medicalServicesHtml =
      this._renderFooterItems(
        medicalServices
      );


    // Render Footer Template
    this._bottomPlaceholder.domElement.innerHTML =
      Footer.html
        .replace(
          '__IMPORTANT_CONTACTS__',
          importantContactsHtml
        )
        .replace(
          '__MEDICAL_SERVICES__',
          medicalServicesHtml
        )
        .replace(
          '__CURRENT_YEAR__',
          new Date().getFullYear().toString()
        );


    // Move the existing footer placeholder into SharePoint Canvas
    const spCanvasElement =
      document.querySelector('.SPCanvas div');


    if (spCanvasElement) {

      // Remove duplicate footer placeholders already inside Canvas
      const existingFooters =
        spCanvasElement.querySelectorAll(
          '#bottomPlaceholder'
        );


      existingFooters.forEach(
        (footer: Element) => {

          if (
            footer !==
            this._bottomPlaceholder!.domElement
          ) {
            footer.remove();
          }

        }
      );


      // Move the actual Bottom Placeholder
      if (
        this._bottomPlaceholder.domElement.parentElement !==
        spCanvasElement
      ) {

        spCanvasElement.appendChild(
          this._bottomPlaceholder.domElement
        );

      }
    }
  }
}