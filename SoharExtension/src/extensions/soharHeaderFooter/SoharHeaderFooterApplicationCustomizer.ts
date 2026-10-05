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

  private loader: HTMLElement | null = null;


  @override

  public async onInit(): Promise<void> {
  await this._loadCSS();
  this._addPageLoader();

  Log.info(
    LOG_SOURCE,
    `Initialized ${strings.Title}`
  );

  await this.loadBootstrap();
  await this._renderHeader();

  this._loadHome();
  await this._renderFooter();
  this._hidePageLoader();
  return Promise.resolve();
}
// ============================================================
// PAGE LOADER
// ============================================================

private _addPageLoader(): void {

  this.loader = document.createElement('div');
  this.loader.className = 'sa-loader';
  this.loader.id = 'saLoader';

  this.loader.innerHTML = `
    <div class="sa-loader-inner">
      <div class="sa-ring sa-ring--secondary"></div>
      <div class="sa-ring sa-ring--primary"></div>
      <img src="/sites/DevPortal/SiteAssets/resources/images/logo-mob.png" alt=""/>
    </div>
  `;

  document.body.appendChild(this.loader);

  const currentUrl = window.location.href.toLowerCase();

  if (
    currentUrl.indexOf('/sites/devportal/sitepages/home.aspx') !== -1
  ) {
    window.addEventListener('load', () => {
      if (this.loader) {
        this.loader.style.display = 'none';
      }
    });
  }
}
private _hidePageLoader(): void {

  setTimeout(() => {

    if (this.loader) {
      this.loader.style.display = 'none';
    }

  }, 4000);
}

  private async _loadHome(): Promise<void> {

  const baseUrl: string =
    'https://soharaluminium5.sharepoint.com/sites/DevPortal';

  await SPComponentLoader.loadScript(
    `${baseUrl}/SiteAssets/resources/js/common.js`
  );

  await SPComponentLoader.loadScript(
    `${baseUrl}/SiteAssets/resources/js/home.js`
  );
}
private async loadBootstrap(): Promise<void> {
 
   const baseUrl: string =
  'https://soharaluminium5.sharepoint.com/sites/DevPortal';
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
 
private async _loadCSS(): Promise<void> {
  const baseUrl: string =
    'https://soharaluminium5.sharepoint.com/sites/DevPortal';

  // Load all CSS files
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

  // Load jQuery 3.6.0
  await SPComponentLoader.loadScript(
    `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
  );

  // Save the original jQuery instance
  (window as any).soharJQuery = (window as any).jQuery;

  // Load jQuery UI
  await SPComponentLoader.loadScript(
    `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
  );

  // Load Swiper
  await SPComponentLoader.loadScript(
    `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`
  );
}

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

private async _getDepartments(): Promise<string[]> {

  const siteUrl: string =
    'https://soharaluminium5.sharepoint.com/sites/DevPortal';

  const url =
    `${siteUrl}/_api/web/lists/getbytitle('Departments')/items?$select=Title,Link,Status,SortOrder`;

  try {

    const response = await this.context.spHttpClient.get(
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

    const data = await response.json();

    const activeDepartments: any[] = data.value
      .filter((department: any) => department.Status === 'Active')
      .sort((a: any, b: any) => Number(a.SortOrder) - Number(b.SortOrder));
const departmentItems: string[] = [];
const itemsPerColumn: number = 4;

for (let i = 0; i < activeDepartments.length; i += itemsPerColumn) {

  const chunk = activeDepartments.slice(
    i,
    i + itemsPerColumn
  );

  departmentItems.push(
    chunk.map((department: any) => {

      const link = department.Link?.Url || '#';

      return `
        <li>
          <a href="${link}" target="_blank" data-interception="off">
            <span>${department.Title}</span>
          </a>
        </li>
      `;

    }).join('')
  );
}

return departmentItems;

  } catch (error) {

    Log.error(
      LOG_SOURCE,
      error instanceof Error ? error : new Error(String(error))
    );

    return ['', '', ''];
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
  'https://soharaluminium5.sharepoint.com/sites/DevPortal';


    const userName: string =
      this.context.pageContext.user.displayName;


    const userEmail: string =
      this.context.pageContext.user.email;


    const profilePhoto: string =
      `${siteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(userEmail)}`;


    const designation: string =
      await this._getUserDesignation();


const departmentItems: string[] =
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
    this._setupDepartmentMegaMenu();

  }


private _setupDepartmentMegaMenu(): void {

  if (!this._topPlaceholder) {
    return;
  }

  const departmentMenu =
    this._topPlaceholder.domElement.querySelector(
      '.has-mega'
    ) as HTMLElement;

  if (!departmentMenu) {
    return;
  }

  const departmentLink =
    departmentMenu.querySelector(
      '.nav-link'
    ) as HTMLElement;

  if (!departmentLink) {
    return;
  }

  departmentLink.addEventListener(
    'click',
    (event: Event) => {

      event.preventDefault();
      event.stopPropagation();

      const isOpen =
        departmentMenu.classList.contains('show');

      departmentMenu.classList.toggle(
        'show',
        !isOpen
      );

      departmentLink.setAttribute(
        'aria-expanded',
        (!isOpen).toString()
      );
    }
  );

  document.addEventListener(
    'click',
    (event: Event) => {

      if (
        !departmentMenu.contains(
          event.target as Node
        )
      ) {

        departmentMenu.classList.remove('show');

        departmentLink.setAttribute(
          'aria-expanded',
          'false'
        );
      }
    }
  );
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

  const searchButton =
    this._topPlaceholder.domElement.querySelector(
      '.search-icon-btn'
    ) as HTMLButtonElement;

  if (!inputMainSearchBox) {
    return;
  }

  const performSearch = (): void => {

    const searchKey: string =
      inputMainSearchBox.value.trim();

   if (searchKey) {
  const siteUrl: string = this.context.pageContext.web.absoluteUrl;

  window.open(
    `${siteUrl}/_layouts/15/search.aspx/siteall?q=${encodeURIComponent(searchKey)}`,
    '_blank'
  );
}
  };

  inputMainSearchBox.addEventListener(
    'keydown',
    (event: KeyboardEvent) => {

      if (event.key === 'Enter') {
        event.preventDefault();
        performSearch();
      }
    }
  );

  if (searchButton) {
    searchButton.addEventListener(
      'click',
      () => {
        performSearch();
      }
    );
  }
}
  // ============================================================
  // GET FOOTER ITEMS FROM SHAREPOINT
  // ============================================================

  private async _getFooterItems(): Promise<any[]> {

    const siteUrl :
      string =
  'https://soharaluminium5.sharepoint.com/sites/DevPortal';


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

  const baseUrl: string =
    'https://soharaluminium5.sharepoint.com/sites/DevPortal';

  return items.map((item: any) => {

    const contactDetails =
      item.ContactDetails || '';

    let imageUrl = '';

    if (item.Icon) {

      const imgData =
        JSON.parse(item.Icon);

      imageUrl =
        `${baseUrl}/Lists/Footer/Attachments/${item.Id}/${imgData.fileName}`;
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


  // Create Footer HTML
  const footerHTML =
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


  // Check if the current page is the SharePoint Search page
  const isSearchPage =
    window.location.pathname
      .toLowerCase()
      .indexOf('/_layouts/15/search.aspx') !== -1;


  // Search Page Footer
  if (isSearchPage) {

    // Find the empty Bottom placeholder inside the Search layout
    const candidates =
      document.querySelectorAll(
        'div[data-sp-placeholder="Bottom"].sp-placeholder-bottom'
      );


    let searchPlaceholder: HTMLElement | null = null;


    for (
      let i = 0;
      i < candidates.length;
      i++
    ) {

      const element =
        candidates[i] as HTMLElement;


      if (
        element.id !== 'spBottomPlaceholder' &&
        element.children.length === 0
      ) {

        searchPlaceholder =
          element;

        break;
      }
    }


    // Render footer inside the Search page placeholder
    if (searchPlaceholder) {

      searchPlaceholder.innerHTML =
        footerHTML;


      // Remove the outer SPFx Bottom placeholder
      // to prevent duplicate footer
      const spfxBottom =
        document.getElementById(
          'spBottomPlaceholder'
        );


      if (spfxBottom) {
        spfxBottom.remove();
      }


      // Clear the cached placeholder reference
      // because the DOM element has been removed
      if (this._bottomPlaceholder) {
        this._bottomPlaceholder = undefined;
      }


      return;
    }
  }


  // Default path for all other SharePoint pages


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
  if (
    !this._bottomPlaceholder ||
    !this._bottomPlaceholder.domElement
  ) {

    Log.error(
      LOG_SOURCE,
      new Error(
        'Bottom placeholder was not created'
      )
    );

    return;
  }


  // Render Footer
  this._bottomPlaceholder.domElement.innerHTML =
    footerHTML;


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