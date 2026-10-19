import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
 
import * as strings from 'WpNewsListingWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import wpNewsListing, { INewsCardView } from './newsListingHtml';
 
export interface IWpNewsListingWebPartProps {
  description: string;
}
 
// interfaces for sharepoint list items
export interface ISPLists {
  value: ISPList[];
}
 
export interface ISPList {
  Id: number;
  Title: string;
  ShortDescription?: string;
  Category?: string;
  PublishedDate?: string;
  Status?: string;
  NewsImage?: any;
  MainContent?: string;
}
 
// document link found inside the MainContent column
export interface IDocumentLink {
  type: string;  // image | file (empty when there is no document link)
  src: string;   // url shown in the preview panel
  label: string; // text shown under the title: PDF File | Image | Document
}
 
// Only news items with this status are displayed
const PUBLISHED_STATUS: string = 'Active';
 
export default class WpNewsListingWebPart extends BaseClientSideWebPart<IWpNewsListingWebPartProps> {
 
  // All active news loaded from the list
  private _newsItems: ISPList[] = [];
  private _categoryChoices: string[] = [];
 
  // Current filter selection (values come from data-value of the dropdown items)
  private _filters: { category: string; year: string; month: string } = {
    category: 'All',
    year: 'All',
    month: 'All'
  };
 
  public async render(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const resourceUrl: string = `${webUrl}/SiteAssets/resources`;
    const homeUrl: string = `${webUrl}/SitePages/Home.aspx?env=WebViewList`;
 
    // Reset filters on every render
    this._filters = { category: 'All', year: 'All', month: 'All' };
 
    // Show page header with loader until data is ready
    this.domElement.innerHTML = wpNewsListing.wrapperHtml(resourceUrl, homeUrl, wpNewsListing.loadingHtml);
 
    const listApiUrl: string = `${webUrl}/_api/web/lists/GetByTitle('News')`;
    const itemsApiUrl: string = `${listApiUrl}/items?$select=Id,Title,Category,PublishedDate,Status,MainContent&$filter=Status eq '${PUBLISHED_STATUS}'&$orderby=PublishedDate desc&$top=500`;
 
    try {
            const results = await Promise.all([
        this._getNewsItems(itemsApiUrl),
        this._getCategoryChoices()
      ]);
      this._newsItems = results[0];
      this._categoryChoices = results[1];
      this._renderNews();
      this._initNewsFilters();
    } catch (error) {
      console.error("Error loading news", error);
      this._setContent(wpNewsListing.noRecordHtml);
    }
 
    // Items are in the DOM now, so news-list.js can bind to them
    await this._loadNewsViewer();
  }
 
  // Generic GET request returning JSON
  private async _getJson(apiUrl: string): Promise<any> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
    }
    return response.json();
  }
 
  // Get active news items
  private async _getNewsItems(apiUrl: string): Promise<ISPList[]> {
    const data: ISPLists = await this._getJson(apiUrl);
    return data.value || [];
  }
 
    // Read the Choice values directly from the Category column definition
  private async _getCategoryChoices(): Promise<string[]> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const url: string = `${webUrl}/_api/web/lists/GetByTitle('News')/fields/getbyinternalnameortitle('Category')?$select=Choices`;
 
    try {
      const data: any = await this._getJson(url);
      // nometadata returns an array, verbose returns { results: [...] }
      const choices: string[] = Array.isArray(data.Choices) ? data.Choices : (data.Choices && data.Choices.results) || [];
      return choices;
    } catch (error) {
      console.error("Error loading category choices", error);
      return [];
    }
  }
 
  // Render news items in the sidebar list
  private _renderNews(): void {
    this._setContent(this._buildListingHtml(this._newsItems));
  }
 
  // Build news sidebar items HTML
  private _buildListingHtml(items: ISPList[]): string {
    if (!items.length) {
      return wpNewsListing.noRecordHtml;
    }
 
    const webUrl: string = this.context.pageContext.web.absoluteUrl;
    const iconUrl: string = `${webUrl}/SiteAssets/resources/images/icons/quick-links/quick-link-1.png`;
    let html: string = "";
 
    for (let i = 0; i < items.length; i++) {
      const item: ISPList = items[i];
      const content: string = item.MainContent || "";
 
      // A document link inside MainContent is shown as a preview, otherwise the text is shown
      const doc: IDocumentLink = this._getDocumentLink(content);
      const type: string = doc.type || 'article';
 
      const view: INewsCardView = {
        id: item.Id,
        type: type,
        title: escape(item.Title || ""),
        tag: escape(item.Category || ""),
        date: escape(this._formatDate(item.PublishedDate)),
        published: escape(item.PublishedDate || ""),   // raw ISO date used by the filters
        src: escape(doc.src),
        fileTypeLabel: doc.label || 'Article',
        iconUrl: iconUrl,
        contentHtml: type === 'article' ? content : ""
      };
      html += wpNewsListing.cardHtml(view);
    }
 
    return html;
  }
 
  // Update the news list container
  private _setContent(html: string): void {
    const container: Element | null = this.domElement.querySelector('#newsList');
    if (container !== null) {
      container.innerHTML = html;
    }
  }
 
  // ---------- Filters (Category / Year / Month on published date) ----------
 
  private _initNewsFilters(): void {
    const items: HTMLElement[] = Array.prototype.slice.call(
      this.domElement.querySelectorAll('.news-list-item')
    );
 
    // Category list from the loaded data
    const categories: string[] = [];
    // Year list from the published date
    const years: number[] = [];
 
    for (let i = 0; i < items.length; i++) {
      const tag: string = items[i].getAttribute('data-tag') || "";
      if (tag && categories.indexOf(tag) === -1) {
        categories.push(tag);
      }
 
      const year: number = new Date(items[i].getAttribute('data-published') || "").getFullYear();
      if (!isNaN(year) && years.indexOf(year) === -1) {
        years.push(year);
      }
    }
 
       // Prefer the column's Choice values (keeps the list order); fall back to categories found in the data
    const categoryOptions: string[] = this._categoryChoices.length
      ? this._categoryChoices.map((c: string) => escape(c))
      : categories.sort();
    years.sort((a: number, b: number) => b - a);
 
    const categoryList: Element | null = this.domElement.querySelector('#categoryFilterList');
    const yearList: Element | null = this.domElement.querySelector('#yearFilterList');
    if (categoryList) { categoryList.innerHTML = wpNewsListing.categoryItemHtml(categoryOptions); }
    if (yearList) { yearList.innerHTML = wpNewsListing.yearItemHtml(years); }
 
    // Dropdown item click (event delegation on the filter area)
    const filterArea: Element | null = this.domElement.querySelector('.filter-area-wrapper');
    if (filterArea) {
      filterArea.addEventListener('click', (e: Event) => {
        const target: HTMLElement = e.target as HTMLElement;
        const link: HTMLElement | null = target.closest('.dropdown-item') as HTMLElement | null;
        if (!link) {
          return;
        }
        e.preventDefault();
 
        const dropdown: HTMLElement | null = link.closest('.sa-filter-dropdown') as HTMLElement | null;
        if (!dropdown) {
          return;
        }
 
        const key: string = dropdown.getAttribute('data-filter') || "";
        const value: string = link.getAttribute('data-value') || 'All';
 
        if (key === 'category' || key === 'year' || key === 'month') {
          this._filters[key] = value;
        }
 
        const selectedText: Element | null = dropdown.querySelector('.filter-selected-text');
        if (selectedText) {
          selectedText.textContent = link.textContent;
        }
 
        this._applyNewsFilters(items);
      });
    }
 
    // Sets the initial count
    this._applyNewsFilters(items);
     this._bindFilteredKeyNav();
  }
 
  private _applyNewsFilters(items: HTMLElement[]): void {
    let visible: number = 0;
    let firstVisible: HTMLElement | null = null;
    let selectedStillVisible: boolean = false;
 
    for (let i = 0; i < items.length; i++) {
      const item: HTMLElement = items[i];
      const published: Date = new Date(item.getAttribute('data-published') || "");
      const hasDate: boolean = !isNaN(published.getTime());
 
      const okCategory: boolean = this._filters.category === 'All' || item.getAttribute('data-tag') === this._filters.category;
      const okYear: boolean = this._filters.year === 'All' || (hasDate && String(published.getFullYear()) === this._filters.year);
      const okMonth: boolean = this._filters.month === 'All' || (hasDate && String(published.getMonth() + 1) === this._filters.month);
 
      const show: boolean = okCategory && okYear && okMonth;
      item.classList.toggle('d-none', !show);
 
      if (show) {
        visible++;
        if (!firstVisible) {
          firstVisible = item;
        }
        if (item.classList.contains('is-active')) {
          selectedStillVisible = true;
        }
      }
    }
 
    const count: Element | null = this.domElement.querySelector('#newsCount');
    if (count) {
       count.textContent = visible + (visible === 1 ? ' item' : ' items');
    }
 
    // "No records found" message
    const list: Element | null = this.domElement.querySelector('#newsList');
    if (list) {
      const old: Element | null = list.querySelector('.news-no-result');
      if (old) {
        old.parentNode!.removeChild(old);
      }
          if (visible === 0 && items.length > 0) {
        list.insertAdjacentHTML(
          'beforeend',
          `<div class="p-3 news-no-result"><p class="text-color-desc text-sm">No records found</p></div>`
        );
      }
    }
 
      // Only after news-list.js has loaded, otherwise its click handlers don't exist yet
    if (this._viewerLoaded && items.length > 0) {
      if (visible === 0) {
        this._clearDetailPanel();
      } else if (!selectedStillVisible && firstVisible) {
        this._selectItem(firstVisible);
      }
    }
  }
 
    // Click an item through the script's own handler. On mobile that handler scrolls
  // the page down to the detail panel, so the scroll position is restored afterwards.
  private _selectItem(item: HTMLElement): void {
    const y: number = window.pageYOffset;
    item.click();
    window.scrollTo(0, y);
  }
 
  // Empty state for the right panel (uses the element ids from news-list.js)
  private _clearDetailPanel(): void {
    const root: HTMLElement = this.domElement;
 
    const items: NodeListOf<Element> = root.querySelectorAll('.news-list-item');
    for (let i = 0; i < items.length; i++) {
      items[i].classList.remove('is-active');
      items[i].setAttribute('aria-selected', 'false');
      items[i].setAttribute('tabindex', '-1');
    }
 
    const detail: Element | null = root.querySelector('#newsDetail');
    if (detail) { detail.removeAttribute('aria-labelledby'); }
 
    const setText = (selector: string, text: string): void => {
      const el: Element | null = root.querySelector(selector);
      if (el) { el.textContent = text; }
    };
    setText('#newsDetailTitle', 'No news found');
    setText('#newsDetailTag', '');
    setText('#newsDetailDate', '');
 
    const sep: HTMLElement | null = root.querySelector('#newsDetailSep');
    if (sep) { sep.style.display = 'none'; }
 
    ['#newsOpenLink', '#newsDownloadLink', '#newsViewerHint', '#newsDetailLoader'].forEach((selector: string) => {
      const el: Element | null = root.querySelector(selector);
      if (el) { el.classList.add('d-none'); }
    });
 
    const body: Element | null = root.querySelector('#newsDetailBody');
    if (body) {
      body.className = 'news-detail-body';
      body.innerHTML = '<p class="news-detail-empty">No news matches the selected filters.</p>';
    }
  }
 
  // The script's arrow-key handler loops over ALL items, including filtered-out ones.
  // This capture listener runs first and moves only between the visible items.
  private _bindFilteredKeyNav(): void {
    const list: Element | null = this.domElement.querySelector('#newsList');
    if (!list) {
      return;
    }
 
    list.addEventListener('keydown', (e: Event) => {
      const key: string = (e as KeyboardEvent).key;
      const current: HTMLElement | null = (e.target as HTMLElement).closest('.news-list-item') as HTMLElement | null;
      if (!current || ['ArrowDown', 'ArrowRight', 'ArrowUp', 'ArrowLeft', 'Home', 'End'].indexOf(key) === -1) {
        return;
      }
 
      e.stopPropagation(); // keeps the script's own handler from running
      e.preventDefault();
 
      const visible: HTMLElement[] = Array.prototype.slice.call(list.querySelectorAll('.news-list-item:not(.d-none)'));
      if (!visible.length) {
        return;
      }
 
      const index: number = visible.indexOf(current);
      let next: HTMLElement;
      if (key === 'Home') {
        next = visible[0];
      } else if (key === 'End') {
        next = visible[visible.length - 1];
      } else if (key === 'ArrowDown' || key === 'ArrowRight') {
        next = visible[(index + 1) % visible.length];
      } else {
        next = visible[(index - 1 + visible.length) % visible.length];
      }
 
      this._selectItem(next);
      try { next.focus({ preventScroll: true }); } catch (err) { next.focus(); }
    }, true); // true = capture phase
  }
 
  // ---------- Document link helpers ----------
 
  // Find the first link inside the MainContent HTML (same idea as the home page: every http/https link is previewed)
  private _getDocumentLink(html: string): IDocumentLink {
    const result: IDocumentLink = { type: "", src: "", label: "" };
    if (!html) {
      return result;
    }
 
    const doc: Document = new DOMParser().parseFromString(html, "text/html");
    const links: NodeListOf<HTMLAnchorElement> = doc.querySelectorAll('a[href]');
 
    for (let i = 0; i < links.length; i++) {
      let url: URL;
      try {
        url = new URL(links[i].getAttribute('href') || "", this.context.pageContext.web.absoluteUrl);
      } catch (e) {
        continue;
      }
 
      // Only http and https links are allowed
      if (url.protocol !== 'https:' && url.protocol !== 'http:') {
        continue;
      }
 
      const fileUrl: string = this._getFileUrl(url);
      const match: RegExpExecArray | null = /\.([a-z0-9]+)$/i.exec(new URL(fileUrl).pathname);
      const ext: string = match ? match[1].toLowerCase() : "";
 
      if (/^(png|jpe?g|gif|webp|svg)$/.test(ext)) {
        return { type: 'image', src: fileUrl, label: 'Image' };
      }
 
      // Sharing links look like /:b:/s/Site/... where b = PDF
      const isPdf: boolean = ext === 'pdf' || /^\/:b:\//i.test(url.pathname);
      return { type: 'file', src: this._getPreviewUrl(fileUrl), label: isPdf ? 'PDF File' : 'Document' };
    }
 
    return result;
  }
 
  // Library view links (AllItems.aspx?...&id=/path/file.pdf) keep the real path in "id"
  private _getFileUrl(url: URL): string {
    const filePath: string | null = url.searchParams.get('id');
    return filePath ? window.location.origin + encodeURI(filePath) : url.href;
  }
 
  // Office files need the embed view to render inside an iframe
  private _getPreviewUrl(fileUrl: string): string {
    const ext: string = (fileUrl.split('?')[0].split('.').pop() || '').toLowerCase();
    const separator: string = fileUrl.indexOf('?') > -1 ? '&' : '?';
 
    if (['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx'].indexOf(ext) > -1) {
      return fileUrl + separator + 'action=embedview';
    }
    return ext === 'pdf' ? fileUrl + separator + 'web=1' : fileUrl;
  }
 
  // Format date like "May 20, 2024"
  private _formatDate(value?: string): string {
    if (!value) {
      return "";
    }
    const date: Date = new Date(value);
    if (isNaN(date.getTime())) {
      return "";
    }
    return date.toLocaleDateString('en-US', { month: 'long', day: '2-digit', year: 'numeric' });
  }
 
  public async onInit(): Promise<void> {
    await this._loadCSS();
  }
 
  private _viewerLoaded: boolean = false;
 
  // Load jQuery and news-list.js (selection, viewer, loader, back button)
  private async _loadNewsViewer(): Promise<void> {
    const baseUrl: string = this.context.pageContext.web.absoluteUrl;
    try {
      await SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`, { globalExportsName: 'jQuery' });
      await SPComponentLoader.loadScript(`${baseUrl}/SiteAssets/resources/js/news-list.js`);
      this._viewerLoaded = true;
    } catch (error) {
      console.error("Error loading news viewer script", error);
    }
  }
 
  private _loadCSS(): void {
    const baseUrl: string = this.context.pageContext.web.absoluteUrl;
    // Load CSS files in the same order as the HTML page
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/news-list.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/list.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/jquery-ui.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/variable.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/font-size.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/custom.css`);
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/sp-custom.css`);
  }
 
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
 
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}