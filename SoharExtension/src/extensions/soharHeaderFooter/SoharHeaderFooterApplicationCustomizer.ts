import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from '@microsoft/sp-application-base';
import { MSGraphClientV3, SPHttpClient } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'SoharHeaderFooterApplicationCustomizerStrings';
import Header from './Header';
import Footer from './Footer';

const LOG_SOURCE: string = 'SoharHeaderFooterApplicationCustomizer';
const SITE_URL: string = 'https://soharaluminium5.sharepoint.com/sites/DevPortal';
const RESOURCES_URL: string = `${SITE_URL}/SiteAssets/resources`;
const LOADER_DURATION_MS: number = 3000;

export interface ISoharHeaderFooterApplicationCustomizerProperties {
  testMessage: string;
}

export default class SoharHeaderFooterApplicationCustomizer extends BaseApplicationCustomizer<ISoharHeaderFooterApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _loader: HTMLElement | null = null;

  // Cached data so header/footer can be restored after page navigation
  private _designation: string = '';
  private _isAdmin: boolean = false;
private _departmentItems: any[] = [];
  private _controlledDocumentItems: any[] = [];
  private _documentCenterItems: any[] = [];
  private _listingPageItems: any[] = [];
  private _footerHtml: string = '';
  private _isInitialised: boolean = false;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
        if (this._isSearchPage()) {
      document.documentElement.classList.add('sa-search-page');
    }
    // 1. Inject stylesheets first so the loader is styled
    this._loadStyles();

    // 2. Show loader on the home page only, removed after exactly 3 seconds
    this._addPageLoader();
    window.setTimeout(() => this._removeLoader(), LOADER_DURATION_MS);

    // 3. Restore header/footer after SharePoint page navigation
    this.context.application.navigatedEvent.add(this, this._onNavigated);

    // 4. Render everything without blocking SharePoint
    this._initialise().catch((error: unknown) => {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
    });

    return Promise.resolve();
  }

  @override
  protected onDispose(): void {
    this.context.application.navigatedEvent.remove(this, this._onNavigated);
    document.removeEventListener('click', this._onDocumentClick);
    this._removeLoader();
  }

  // ============================================================
  // INITIALISATION (parallel loading, correct render order)
  // ============================================================




private async _initialise(): Promise<void> {
  // Load scripts and fetch all header data in parallel
  const scriptsPromise: Promise<void> = this._loadScripts();

const headerDataPromise: Promise<[string, any[], any[], any[], any[], boolean]> =
  Promise.all([
    this._getUserDesignation(),
    this._getDepartments(),
    this._getControlledDocuments(),
    this._getDocumentCenterItems(),
    this._getListingPageItems(),
    this._isAdminUser()
  ]);

  const footerItemsPromise: Promise<any[]> = this._getFooterItems();

  // Receive all five results in the same order as Promise.all
const [
  designation,
  departmentItems,
  controlledDocumentItems,
  documentCenterItems,
  listingPageItems,
  isAdmin
] = await headerDataPromise;

  // Cache the fetched data
  this._designation = designation;
  this._departmentItems = departmentItems;
  this._controlledDocumentItems = controlledDocumentItems;
  
  this._documentCenterItems = documentCenterItems.map((item: any) => ({
  ...item,
  IconUrl: this._getIconUrl(item)
}));
this._listingPageItems = listingPageItems.map((item: any) => ({
  ...item,
  IconUrl: this._getIconUrl(item)
}));

this._isAdmin = isAdmin;



 // Render header after its data is ready
  this._renderHeader();

  // Render footer after its data is ready
  this._footerHtml = this._buildFooterHtml(await footerItemsPromise);
  await this._waitForElement(this._getFooterTargetSelector());
  this._renderFooter();

  // Load page scripts after header and footer HTML exist
  await scriptsPromise;
  await this._loadPageScripts();

  this._isInitialised = true;
}



  private _onNavigated(): void {
    if (!this._isInitialised) {
      return;
    }

    // Re-render header if its placeholder was removed
    if (!this._topPlaceholder || !document.body.contains(this._topPlaceholder.domElement)) {
      this._topPlaceholder = undefined;
      this._renderHeader();
    }

    // Canvas is re-rendered on navigation, so re-place the footer
    this._waitForElement(this._getFooterTargetSelector()).then(() => this._renderFooter()).catch(() => { /* ignore */ });
  }

  // ============================================================
  // RESOURCES
  // ============================================================

  private _loadStyles(): void {
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/variable.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/bootstrap.min.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/jquery-ui.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/swiper-bundle.min.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/font-size.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/custom.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/sp-custom.css`);
    SPComponentLoader.loadCss(`${RESOURCES_URL}/css/home.css`);
  }

  private async _loadScripts(): Promise<void> {
    const w: any = window;

    // Load scripts sequentially: jQuery -> jQuery UI -> Bootstrap -> Swiper
    await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/jquery-3.6.0.js`);

    // Save the original jQuery instance
    w.soharJQuery = w.jQuery;

    await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/jquery-ui.js`);

    // Another script (e.g. SociableKit) may have replaced window.jQuery while jQuery UI was loading.
    // Keep whichever jQuery instance jQuery UI actually attached the datepicker to.
    if (!(w.soharJQuery && w.soharJQuery.fn && w.soharJQuery.fn.datepicker) && w.jQuery && w.jQuery.fn && w.jQuery.fn.datepicker) {
      w.soharJQuery = w.jQuery;
    }

    await this._loadBootstrap();
    await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/swiper-bundle.min.js`);

    // Save the Swiper instance as well, so later scripts cannot replace it
    w.soharSwiper = w.Swiper;
    // await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/jquery.marquee.min.js`);
  }

  private async _loadBootstrap(): Promise<void> {
    const win: any = window;
    const bootstrap = await SPComponentLoader.loadScript<any>(`${RESOURCES_URL}/js/bootstrap.bundle.min.js`);

    if (bootstrap?.Modal) {
      win.bootstrap = bootstrap;

      const original = bootstrap.Modal.getOrCreateInstance;

      if (typeof original === 'function') {
        bootstrap.Modal.getOrCreateInstance = function (element: Element, config?: any) {
          return original.call(bootstrap.Modal, element, config);
        };
      }
    }
  }

  // private async _loadPageScripts(): Promise<void> {
  //   // 1. home.js loads first
  //   try {
  //     await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/home.js`);
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(`Failed to load home.js: ${String(error)}`));
  //   }

  //   // 2. common.js loads after; a failure here does not affect home.js
  //   try {
  //     await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/common.js`);
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(`Failed to load common.js: ${String(error)}`));
  //   }
  // }

  private async _loadPageScripts(): Promise<void> {
    await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/common.js`);
    await SPComponentLoader.loadScript(`${RESOURCES_URL}/js/home.js`);
  }

  // ============================================================
  // PAGE LOADER
  // ============================================================

  private _addPageLoader(): void {
    // Display loader only on the DevPortal home page
    const path: string = window.location.pathname.toLowerCase().replace(/\/$/, '');
    const isHomePage: boolean = path === '/sites/devportal/sitepages/home.aspx' || path === '/sites/devportal';

    if (!isHomePage) {
      return;
    }

    // Prevent duplicate loader
    if (document.getElementById('saLoader')) {
      return;
    }

    this._loader = document.createElement('div');
    this._loader.className = 'sa-loader';
    this._loader.id = 'saLoader';
    this._loader.innerHTML = `
      <div class="sa-loader-inner">
        <div class="sa-ring sa-ring--secondary"></div>
        <div class="sa-ring sa-ring--primary"></div>
        <img src="${RESOURCES_URL}/images/logo-mob.png" alt="" />
      </div>
    `;

    document.body.appendChild(this._loader);
  }

  private _removeLoader(): void {
    if (this._loader && document.body.contains(this._loader)) {
      document.body.removeChild(this._loader);
    }
    this._loader = null;
  }

  private _waitForElement(selector: string, timeoutMs: number = 10000): Promise<Element | null> {
    return new Promise<Element | null>((resolve) => {
      const existing: Element | null = document.querySelector(selector);

      if (existing) {
        resolve(existing);
        return;
      }

      const observer: MutationObserver = new MutationObserver(() => {
        const element: Element | null = document.querySelector(selector);

        if (element) {
          observer.disconnect();
          resolve(element);
        }
      });

      observer.observe(document.body, { childList: true, subtree: true });

      window.setTimeout(() => {
        observer.disconnect();
        resolve(null);
      }, timeoutMs);
    });
  }

  // ============================================================
  // HEADER DATA
  // ============================================================

  private async _getUserDesignation(): Promise<string> {
    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const user = await client.api('/me').select('jobTitle').get();

      return user.jobTitle || '';
    } catch (error) {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
      return '';
    }
  }

private async _getDepartments(): Promise<any[]> {
 
const url: string =
  `${SITE_URL}/_api/web/lists/getbytitle('Departments')/items` +
  `?$select=Title,Link,Status,SortOrder` +
  `&$filter=Status eq 'Active'` +
  `&$orderby=Title asc`;


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

    return data.value || [];

  } catch (error) {
    Log.error(
      LOG_SOURCE,
      error instanceof Error
        ? error
        : new Error(String(error))
    );

    return [];
  }
}
private async _getControlledDocuments(): Promise<any[]> {


const url: string =
  `${SITE_URL}/_api/web/lists/getbytitle('Controlled Documents')/items` +
  `?$select=Id,Title,Link,Status,SortOrder` +
  `&$filter=Status eq 'Active'` +
  `&$orderby=Title asc`;


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
        `Failed to load Controlled Documents: ${response.status} ${response.statusText}`
      );
    }

    const data = await response.json();

    return data.value || [];

  } catch (error) {

    Log.error(
      LOG_SOURCE,
      error instanceof Error
        ? error
        : new Error(String(error))
    );

    return [];
  }
}


private async _getDocumentCenterItems(): Promise<any[]> {
  const url =
    `${SITE_URL}/_api/web/lists/getbytitle('Document Center')/items` +
    `?$select=Id,Title,Link,Icon,Status,SortOrder` +
    `&$filter=Status eq 'Active'` +
    `&$orderby=SortOrder asc`;

  const response = await this.context.spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );

  if (!response.ok) {
    throw new Error('Failed to fetch Document Center items');
  }

  const data = await response.json();

  return data.value || [];
}

private async _getListingPageItems(): Promise<any[]> {
  const url =
    `${SITE_URL}/_api/web/lists/getbytitle('ListingPage')/items` +
    `?$select=Title,Link,Icon,Status,SortOrder` +
    `&$filter=Status eq 'Active'` +
    `&$orderby=SortOrder asc`;

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
        `Failed to load Listing Page: ${response.status} ${response.statusText}`
      );
    }

    const data = await response.json();
    return data.value || [];

  } catch (error) {
    Log.error(
      LOG_SOURCE,
      error instanceof Error ? error : new Error(String(error))
    );

    return [];
  }
}

private _getIconUrl(item: any): string {
  if (!item || !item.Icon) {
    return '';
  }

  try {
    const icon = typeof item.Icon === 'string'
      ? JSON.parse(item.Icon)
      : item.Icon;

    if (icon.serverRelativeUrl) {
      const serverUrl = icon.serverUrl || window.location.origin;
      const imagePath = icon.serverRelativeUrl.startsWith('/')
        ? icon.serverRelativeUrl
        : `/${icon.serverRelativeUrl}`;

      return `${serverUrl.replace(/\/$/, '')}${imagePath}`;
    }

    if (icon.Url) {
      return icon.Url;
    }

    return '';
  } catch {
    return '';
  }
}

private async _isAdminUser(): Promise<boolean> {
  try {
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/groups?$select=Title`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return false;
    }

    const result = await response.json();


    const adminGroups = [
  'Communications Admin',
  'Portal Admin (IT)'
];

    return result.value.some((group: { Title: string }) =>
      adminGroups.some(
        (adminGroup: string) =>
          group.Title.trim().toLowerCase() === adminGroup.toLowerCase()
      )
    );
  } catch {
    return false;
  }
}



  // ============================================================
  // RENDER HEADER
  // ============================================================

  private _renderHeader(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, { onDispose: () => { Log.info(LOG_SOURCE, 'Top placeholder disposed'); } });
    }

    if (!this._topPlaceholder) {
      Log.error(LOG_SOURCE, new Error('Top placeholder was not created'));
      return;
    }

    const userName: string = escape(this.context.pageContext.user.displayName || '');
    const userEmail: string = this.context.pageContext.user.email || '';
    const profilePhoto: string = `${SITE_URL}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(userEmail)}`;

    const header: Header = new Header(
  SITE_URL,
  userName,
  escape(this._designation),
  profilePhoto,
  this._departmentItems,
  this._controlledDocumentItems,
  this._documentCenterItems,
  this._listingPageItems,
  this._isAdmin
  
);

    this._topPlaceholder.domElement.innerHTML = header.render();

    this._setupSearchFunctionality();
    this._setupDepartmentMegaMenu();
    this._setActiveNavLink();
  }


private _setupDepartmentMegaMenu(): void {
  if (!this._topPlaceholder) {
    return;
  }

  const headerElement: HTMLElement = this._topPlaceholder.domElement;

  // Handle department submenu click
  headerElement.addEventListener('click', (event: MouseEvent): void => {
    const target = event.target as HTMLElement;
    const toggle = target.closest('.dept-submenu-toggle') as HTMLAnchorElement | null;
    if (!toggle) {
      return;
    }

const item = toggle.closest('.dept-submenu') as HTMLElement | null;

// Plain link (Listing page, Shared Workspace): navigate
if (!item || !item.querySelector('.dept-submenu-list')) {
 const href: string = toggle.getAttribute('href') || '';
 if (href && href !== '#') {
 event.preventDefault();
 window.location.href = href;
 }
 return;
}

// Real submenu (Departments, Controlled Documents): toggle
event.preventDefault();
event.stopPropagation();

const menu = toggle.closest('.dept-dropdown') as HTMLElement | null;
if (!menu) {
 return;
}

    const isOpen: boolean = item.classList.contains('open');
    const openItems = menu.querySelectorAll('.dept-submenu.open');

    for (let i = 0; i < openItems.length; i++) {
      const openItem = openItems[i] as HTMLElement;
      openItem.classList.remove('open');

      const openToggle = openItem.querySelector('.dept-submenu-toggle') as HTMLElement | null;

      if (openToggle) {
        openToggle.setAttribute('aria-expanded', 'false');
      }
    }

    if (!isOpen) {
      item.classList.add('open');
      toggle.setAttribute('aria-expanded', 'true');
    }
  }, true);

  // Handle department and controlled document search
  headerElement.addEventListener('input', (event: Event): void => {
    const target = event.target as HTMLInputElement;

    if (!target.classList.contains('search-dept-dropdown')) {
      return;
    }

    const searchText: string = target.value.trim().toLowerCase();
    const submenu = target.closest('.dept-submenu-list') as HTMLElement | null;

    if (!submenu) {
      return;
    }

    const listItems = submenu.querySelectorAll('ul > li');

    for (let i = 0; i < listItems.length; i++) {
      const listItem = listItems[i] as HTMLElement;
      const link = listItem.querySelector('a');

      const itemText: string = link && link.textContent
        ? link.textContent.trim().toLowerCase()
        : '';

      listItem.style.display = itemText.indexOf(searchText) !== -1 ? '' : 'none';
    }
  });
}

  private _onDocumentClick = (event: Event): void => {
    const departmentMenu: HTMLElement | null = this._getDepartmentMenu();

    if (departmentMenu && !departmentMenu.contains(event.target as Node)) {
      departmentMenu.classList.remove('show');

      const departmentLink: HTMLElement | null = departmentMenu.querySelector('.nav-link') as HTMLElement;
      if (departmentLink) {
        departmentLink.setAttribute('aria-expanded', 'false');
      }
    }
  };

  private _getDepartmentMenu(): HTMLElement | null {
    return this._topPlaceholder ? this._topPlaceholder.domElement.querySelector('#navbarScrollingDropdown') as HTMLElement : null;
  }

    private _setActiveNavLink(): void {
    if (!this._topPlaceholder) {
      return;
    }

    // Only change the active link when the current site is NOT DevPortal
    const currentWebUrl: string = this.context.pageContext.web.serverRelativeUrl.toLowerCase().replace(/\/$/, '');

    if (currentWebUrl === '/sites/devportal') {
      return;
    }

    const homeLink: HTMLElement | null = this._topPlaceholder.domElement.querySelector('.nav-link.active-nav-link');
    const departmentLink: HTMLElement | null = this._topPlaceholder.domElement.querySelector('#navbarScrollingDropdown');

    if (homeLink) {
      homeLink.classList.remove('active-nav-link');
    }

    if (departmentLink) {
      departmentLink.classList.add('active-nav-link');
    }
  }

  private _setupSearchFunctionality(): void {
    if (!this._topPlaceholder) {
      return;
    }

    const inputMainSearchBox = this._topPlaceholder.domElement.querySelector('#inputGlobalSearchBox') as HTMLInputElement;
    const searchButton = this._topPlaceholder.domElement.querySelector('.search-icon-btn') as HTMLButtonElement;

    if (!inputMainSearchBox) {
      return;
    }

    const performSearch = (): void => {
      const searchKey: string = inputMainSearchBox.value.trim();

      if (searchKey) {
        window.open(`${this.context.pageContext.web.absoluteUrl}/_layouts/15/search.aspx/siteall?q=${encodeURIComponent(searchKey)}`, '_self');
      }
    };

    inputMainSearchBox.addEventListener('keydown', (event: KeyboardEvent) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        performSearch();
      }
    });

    if (searchButton) {
      searchButton.addEventListener('click', (event: Event) => {
        event.preventDefault();
        performSearch();
      });
    }
  }

  // ============================================================
  // FOOTER DATA
  // ============================================================

  private async _getFooterItems(): Promise<any[]> {
const url: string = `${SITE_URL}/_api/web/lists/getbytitle('FooterContacts')/items?$select=Id,ContactDetails,Category,Status,Icon,Created&$orderby=Created desc`;
    try {
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });

      if (!response.ok) {
        throw new Error(`Failed to load Footer: ${response.status}`);
      }

      const data = await response.json();
      return data.value || [];
    } catch (error) {
      Log.error(LOG_SOURCE, error instanceof Error ? error : new Error(String(error)));
      return [];
    }
  }

  private _getFooterIconUrl(item: any): string {
    if (!item.Icon) {
      return '';
    }

    try {
      const img: any = JSON.parse(item.Icon);

      if (img.serverRelativeUrl) {
        return (img.serverUrl || window.location.origin) + encodeURI(img.serverRelativeUrl);
      }

      if (img.fileName) {
        return `${SITE_URL}/Lists/Footer/Attachments/${item.Id}/${encodeURIComponent(img.fileName)}`;
      }
    } catch (error) {
      Log.warn(LOG_SOURCE, `Invalid footer icon for item ${item.Id}`);
    }

    return '';
  }

 
private _renderFooterItems(items: any[]): string {
  return items.map((item: any) => {
    const contactDetails: string = item.ContactDetails || '';
    const imageUrl: string = this._getFooterIconUrl(item);

    return Footer.itemTemplate
      .replace('__ICON_URL__', () => imageUrl)
      .replace('__CONTACT_DETAILS__', () => contactDetails);
  }).join('');
}

private _buildFooterHtml(footerItems: any[]): string {
  const normalise = (value: any): string =>
    String(value || '').trim().toLowerCase();

  const activeItems: any[] = footerItems.filter(
    (item: any) => normalise(item.Status) === 'active'
  );

  const importantContacts: any[] = activeItems.filter(
    (item: any) => normalise(item.Category) === 'important contacts'
  );

  const medicalServices: any[] = activeItems.filter(
    (item: any) => normalise(item.Category) === 'medical services'
  );

  return Footer.html
    .replace(
      '__IMPORTANT_CONTACTS__',
      () => this._renderFooterItems(importantContacts)
    )
    .replace(
      '__MEDICAL_SERVICES__',
      () => this._renderFooterItems(medicalServices)
    )
    .replace(
      '__CURRENT_YEAR__',
      new Date().getFullYear().toString()
    );
}


 
  private _isSearchPage(): boolean {
    return window.location.pathname.toLowerCase().indexOf('/_layouts/15/search.aspx') !== -1;
  }

  // Search page has no .SPCanvas; it has its own Bottom placeholder
  private _getFooterTargetSelector(): string {
    return this._isSearchPage()
      ? 'div[data-sp-placeholder="Bottom"].sp-placeholder-bottom:not(#spBottomPlaceholder)'
      : '.SPCanvas';
  }

  private _renderFooter(): void {
    if (!this._footerHtml) {
      return;
    }

    if (this._isSearchPage()) {
      // 1. Find the bottom placeholder INSIDE the search layout
      //    (empty, or already holding our footer after a re-render)
      const candidates: NodeListOf<Element> = document.querySelectorAll('div[data-sp-placeholder="Bottom"].sp-placeholder-bottom');
      let searchPlaceholder: HTMLElement | null = null;

      for (let i = 0; i < candidates.length; i++) {
        const element: HTMLElement = candidates[i] as HTMLElement;

        if (element.id !== 'spBottomPlaceholder' && (element.children.length === 0 || element.querySelector('.pre-footer-info'))) {
          searchPlaceholder = element;
          break;
        }
      }

      if (searchPlaceholder) {
        searchPlaceholder.innerHTML = this._footerHtml;

        // 2. Remove the outer SPFx bottom placeholder so there are not TWO footers
        const spfxBottom: HTMLElement | null = document.getElementById('spBottomPlaceholder');
        if (spfxBottom) {
          spfxBottom.remove();
        }

        // Clear the cached reference so it is not reused after removal
        this._bottomPlaceholder = undefined;
        return;
      }
    }

    // Default path for all other pages
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, { onDispose: () => { Log.info(LOG_SOURCE, 'Bottom placeholder disposed'); } });
    }

    if (this._bottomPlaceholder && this._bottomPlaceholder.domElement) {
      const footerElement: HTMLElement = this._bottomPlaceholder.domElement;
      footerElement.innerHTML = this._footerHtml;

      const spCanvasElement: Element | null = document.querySelector('.SPCanvas div');

      if (spCanvasElement) {
        const existingFooter: Element | null = spCanvasElement.querySelector('#bottomPlaceholder');

        if (existingFooter && existingFooter !== footerElement) {
          existingFooter.remove();
        }

        spCanvasElement.appendChild(footerElement);
      }
    } else {
      Log.error(LOG_SOURCE, new Error('Could not load bottom placeholder'));
    }
  }
}