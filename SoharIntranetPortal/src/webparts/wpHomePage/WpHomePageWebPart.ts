import { Version } from '@microsoft/sp-core-library';
import Wrapper from './Wrapper';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse, MSGraphClientV3 } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import UpcomingEventsTemplate from './UpcomingEvents';
import { BannerTemplate } from './BannerTemplate';
import MediaGalleryTemplate from './MediaGalleryTemplate';
import AnnouncementOffer from './AnnouncementOffer';
import QuickLinks from './QuickLinks';
import Birthday from './Birthday';
import SocialMedia from './SocialMedia';
import NewsCentre from './NewsCentre';

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
  Image: string | ISharePointImage | undefined;
}

interface IMediaGalleryItem {
  Id: number;
  Title: string;
  Caption: string;
  Status: string;
  SortOrder: number;
  Image: string | ISharePointImage | undefined;
}

// Interface used for both Outlook My Events and SharePoint Organizational Events
export interface IEventItem {
  Id: number;
  Title: string;
  EventDate: string;
  StartTime: string;
  EndTime: string;
  Location: string;
  Status: string;
  Link?: string;
  TeamsUrl?: string;
}

interface IOutlookEvent {
  id?: string;
  subject?: string;
  start?: {
    dateTime?: string;
    timeZone?: string;
  };
  end?: {
    dateTime?: string;
    timeZone?: string;
  };
  location?: {
    displayName?: string;
  };
  onlineMeeting?: {
    joinUrl?: string;
  };
}

interface IAnnouncement {
  Id: number;
  Title: string;
  ShortDescription: string;
  Icon: string;
  Created: string;
}

interface IOffer {
  Id: number;
  Title: string;
  Description: string;
  Created: string;
}

export interface IQuickLinksList {
  Id: number;
  Title: string;
  Icon: string;
  URL: { Url: string };
  Status: string;
  SortOrder: number;
}

export interface IBirthdayList {
  Id: number;
  Title: string;
  EmployeePhoto?: string | ISharePointImage | undefined;
  BirthDate: string;
  Status: string;
}

interface INewsItem {
  Id: number;
  Title: string;
  ShortDescription?: string;
  Category?: string;
  MainContent?: string;
  PublishedDate?: string;
  Status?: string;
}

export default class WpHomePageWebPart extends BaseClientSideWebPart<IWpHomePageWebPartProps> {

  private newsCentreItems: INewsItem[] = [];

  // Organizational Events (SharePoint) and My Events (Outlook)
  private events: IEventItem[] = [];
  private myEvents: IEventItem[] = [];
  private organizationalEventDates: string[] = [];
  private myEventDates: string[] = [];
  private organizationalEventTitles: { [key: string]: string } = {};
  private myEventTitles: { [key: string]: string } = {};
  private socialMediaObserver: MutationObserver | null = null;
  private currentUserId: number | null = null;
  private quickLinksCache: IQuickLinksList[] = [];
  private favouritesCache: { Id: number; LinkId: number }[] = [];

  // ==================== WAIT FOR LIBRARIES (loaded by extension) ====================

  private waitForLibraries(timeout: number = 10000): Promise<void> {
    return new Promise<void>((resolve) => {
      const start = Date.now();
      const check = (): void => {
        const w = window as any;
         if (
        (w.soharJQuery &&
          w.soharJQuery.fn &&
          w.soharJQuery.fn.datepicker &&
          w.Swiper &&
          w.bootstrap) ||
        Date.now() - start > timeout
      ) {
        resolve();
        return;
      }
        setTimeout(check, 50);
      };
      check();
    });
  }



  // ==================== HOME JS (moved from home.js) ====================

  // ---------- BANNER ----------
  private initBannerSwiper(): void {
    const Swiper = (window as any).Swiper;
    if (!Swiper) { console.warn('Swiper is not available.'); return; }

    new Swiper(".banner-swiper", {
      pagination: { el: ".swiper-pagination", clickable: true },
      navigation: { nextEl: ".banner-slide-next", prevEl: ".banner-slide-prev" },
    });
  }

  // ---------- ANNOUNCEMENT / OFFERS TABS ----------
  private initAnnouncementOfferTabs(): void {
    const $ = (window as any).jQuery;
    if (!$) { console.warn('jQuery is not available.'); return; }

    $('.panel-title-tab').on('click', function (this: HTMLElement) {
      const targetId = $(this).data('tab-ao');

      // Update active tab
      $('.panel-title-tab').removeClass('panel-title-tab-active');
      $(this).addClass('panel-title-tab-active');

      // Hide all panels and fade in selected panel
      $('.ao-tab-view').stop(true, true).hide();
      $('#' + targetId).stop(true, true).fadeIn(200);
    });
  }

  // ---------- SOCIAL MEDIA ----------
  private initSocialMediaSwipers(): void {
    const $ = (window as any).jQuery;
    const Swiper = (window as any).Swiper;
    if (!$ || !Swiper) { console.warn('jQuery or Swiper is not available.'); return; }

    const socialSwipers: { [key: string]: any } = {};
    document.querySelectorAll('.social-swiper').forEach(function (el) {
      const panel = el.closest('[id^="social-panel-"]');
      const swiper = new Swiper(el, {
        slidesPerView: 'auto',
        spaceBetween: 12,
        freeMode: { enabled: true, momentum: true },
        grabCursor: true,
        mousewheel: { forceToAxis: true },
        navigation: {
          prevEl: panel ? panel.querySelector('.social-swiper-prev') : null,
          nextEl: panel ? panel.querySelector('.social-swiper-next') : null,
        },
      });
      if (panel) socialSwipers[panel.id] = swiper;
    });

    $('.sm-tab').on('click', function (this: HTMLElement) {
      const targetId = $(this).data('tab-social-id');

      // Update active tab
      $('.sm-tab').removeClass('sm-tab-active');
      $(this).addClass('sm-tab-active');

      // Hide all panels and fade in selected panel
      $('.social-media-view').stop(true, true).hide();
      $('#' + targetId).stop(true, true).fadeIn(200);
    });
  }

  // ---------- MEDIA GALLERY ----------
  private initMediaGallery(): void {
    const $ = (window as any).jQuery;
    // const $ = (window as Window & { soharJQuery?: any }).soharJQuery;
    const Swiper = (window as any).Swiper;
    const bootstrap = (window as any).bootstrap;
    if (!$ || !Swiper || !bootstrap) { console.warn('jQuery, Swiper or Bootstrap is not available.'); return; }

    new Swiper('.gallery-swiper', {
      slidesPerView: 1,
      spaceBetween: 12,
      loop: false,
      speed: 500,
      pagination: { el: ".gallery-swiper-pagination", clickable: true },
      navigation: { nextEl: ".gallery-slide-next", prevEl: ".gallery-slide-prev" },
      breakpoints: {
        400: { slidesPerView: 2, spaceBetween: 12 },
        576: { slidesPerView: 2, spaceBetween: 14 },
        768: { slidesPerView: 3, spaceBetween: 16 },
        992: { slidesPerView: 4, spaceBetween: 16 },
      },
    });

    const modalElement = document.getElementById("galleryModal");
    if (!modalElement) { return; }
    const modal = new bootstrap.Modal.getOrCreateInstance(modalElement);

    let swiperGalleryModal: any = null;
    let gallery_slider_index = 0;

    function initGalleryModalSwiper(): void {
      // Destroy any previous instance before creating a new one
      if (swiperGalleryModal) {
        swiperGalleryModal.destroy(true, true);
        swiperGalleryModal = null;
      }

      swiperGalleryModal = new Swiper(".gallery-modal-swiper", {
        spaceBetween: 10,
        slidesPerView: 1,
        autoplay: false,
        speed: 1400,
        autoHeight: true,
        initialSlide: gallery_slider_index,
        navigation: { nextEl: ".gallery-swiper-modal-next", prevEl: ".gallery-swiper-modal-prev" },
      });

      $(".gallery-modal-content").css("opacity", 1);
    }

    // Set the index BEFORE the modal opens
    $(document).on("click", ".gallery-item img, .gallery-item video", function (this: HTMLElement) {
      $(".gallery-modal-content").css("opacity", 0);
      gallery_slider_index = $(this).closest(".swiper-slide").index();
      modal.show();
    });

    // shown.bs.modal fires AFTER the transition — Swiper can measure dimensions safely here
    modalElement.addEventListener("shown.bs.modal", () => {
      initGalleryModalSwiper();
    });

    // hidden.bs.modal fires AFTER modal is fully hidden
    modalElement.addEventListener("hidden.bs.modal", () => {
      if (swiperGalleryModal) {
        swiperGalleryModal.destroy(true, true);
        swiperGalleryModal = null;
      }
    });
  }

  // ==================== SOCIAL MEDIA TUTORIAL LINKS ====================

  private hideSocialMediaTutorialLinks(root: Document | ShadowRoot = document): void {
    root.querySelectorAll<HTMLElement>('a.tutorial_link').forEach((link) => {
      link.style.setProperty('display', 'none', 'important');
    });

    root.querySelectorAll<HTMLElement>('*').forEach((element) => {
      if (element.shadowRoot) {
        this.hideSocialMediaTutorialLinks(element.shadowRoot);
      }
    });
  }

  private setupSocialMediaTutorialLinkObserver(): void {
    this.socialMediaObserver?.disconnect();
    this.hideSocialMediaTutorialLinks();
    this.socialMediaObserver = new MutationObserver(() => { this.hideSocialMediaTutorialLinks(); });
    this.socialMediaObserver.observe(document.body, { childList: true, subtree: true });
  }

  // ==================== RENDER ====================

  public async render(): Promise<void> {
    const workbenchContent = document.getElementById('workbenchPageContent');
    if (workbenchContent) {
      workbenchContent.style.maxWidth = 'none';
    }

    const baseUrl = this.context.pageContext.web.absoluteUrl;
    const newsarrowIconUrl = `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    const arrowIconUrl = `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;

    const newsApiUrl = `${baseUrl}/_api/web/lists/getbytitle('News')/items?$select=Id,Title,ShortDescription,Category,MainContent,PublishedDate,Status&$filter=Status eq 'Active'&$orderby=PublishedDate desc`;
    const AnnouncementApiUrl = `${baseUrl}/_api/web/lists/GetByTitle('Announcements')/items?$select=Id,Title,ShortDescription,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc&$top=3`;
    const OfferApiUrl = `${baseUrl}/_api/web/lists/GetByTitle('Offers')/items?$select=Id,Title,Description,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc&$top=3`;

    /*
     * 1. Render the page layout immediately (before any API call),
     *    so every section is visible straight away.
     */
    this.domElement.innerHTML = Wrapper.wrapperHtml;
        this.domElement.querySelector('#quick-links-container')!.innerHTML = QuickLinks.allElementsHtml;
    this.domElement.querySelector('#banner-container')!.innerHTML = BannerTemplate.bannerHtml;
    this.domElement.querySelector('#announcement-offer-container')!.innerHTML = AnnouncementOffer.allElementsHtml;
    this.domElement.querySelector('#news-container')!.innerHTML = NewsCentre.allElementsHtml;
    this.domElement.querySelector('#upcoming-events-container')!.innerHTML = UpcomingEventsTemplate.allElementsHtml;
    this.domElement.querySelector('#social-media-container')!.innerHTML = SocialMedia.allElementsHtml;
    this.domElement.querySelector('#birthday-container')!.innerHTML = Birthday.allElementsHtml;
    this.domElement.querySelector('#media-gallery-container')!.innerHTML = MediaGalleryTemplate.allElementsHtml;

    // 2. Static links
    const birthdayViewAll = this.domElement.querySelector('.birthday-view-all') as HTMLAnchorElement;
    if (birthdayViewAll) {
      birthdayViewAll.href = `${baseUrl}/SitePages/Upcoming-Birthdys.aspx`;
    }

    const birthdayArrow = this.domElement.querySelector('.birthday-view-all-arrow') as HTMLImageElement;
    if (birthdayArrow) {
      birthdayArrow.src = `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    }

    this.newsCentreSetupViewAll(newsarrowIconUrl);
    this.setupViewAllLink(arrowIconUrl);

    /*
     * 3. Load every section in parallel.
     *    Each section initialises its own Swiper / tabs as soon as
     *    its own data is ready, so the banner no longer waits for
     *    the other web part sections.
     */
    const libsReady = this.waitForLibraries();
    const today = new Date();

    await Promise.all([

      // Banner (first priority) + Announcement / Offers tabs
      Promise.all([this._getBannerItems(), libsReady]).then(() => {
        this.initBannerSwiper();
        this.initAnnouncementOfferTabs();
      }),

      // Announcements and Offers
      this._renderAnnouncementsAsync(AnnouncementApiUrl),
      this._renderOffersAsync(OfferApiUrl),

      // Media Gallery
      Promise.all([this._getMediaGalleryItems(), libsReady]).then(() => {
        this.initMediaGallery();
      }),

      // Quick Links
      this.initializeQuickLinks(),

      // Birthdays
      this.renderBirthdays(),

      // Social Media
      Promise.all([this.initializeSocialMedia(), libsReady]).then(() => {
        this.initSocialMediaSwipers();
      }),

      // News Centre
      this._renderNewsAsync(newsApiUrl).then(() => {
        this.newsCentreAttachTabEvents();
      }),

      // Upcoming Events
      Promise.all([
        this.loadEvents(today.getFullYear(), today.getMonth(), today.getDate()),
        this.loadMyEvents(today.getFullYear(), today.getMonth(), today.getDate()),
        this.loadOrganizationalEventDates(today.getFullYear(), today.getMonth()),
        this.loadMyEventDates(today.getFullYear(), today.getMonth()),
        libsReady
      ]).then(() => {
        this.renderUpcomingEvents();
        this.initializeUpcomingEvents();
        this.initializeCalendar();

      })
    ]);
  }

  // ==================== NEWS CENTRE ====================

  private newsCentreSetupViewAll(arrowIconUrl: string): void {
    const baseUrl = this.context.pageContext.web.absoluteUrl;

    // Find View All link by ID (fallback: older template placeholder)
    let viewAllLink = this.domElement.querySelector('#news-view-all') as HTMLAnchorElement | null;
    if (!viewAllLink) {
      viewAllLink = this.domElement.querySelector('a[href="__KEY_URL_VIEW_ALL__"]') as HTMLAnchorElement | null;
    }

    if (viewAllLink) {
      viewAllLink.href = `${baseUrl}/SitePages/News-List.aspx`;
    } else {
      console.warn('News View All link not found.');
    }

    // Find arrow image by ID (fallback: older template placeholder)
    let arrowImage = this.domElement.querySelector('#news-view-all-arrow') as HTMLImageElement | null;
    if (!arrowImage) {
      arrowImage = this.domElement.querySelector('img[src="__KEY_URL_ARROW__"]') as HTMLImageElement | null;
    }

    if (arrowImage) {
      arrowImage.src = arrowIconUrl;
    } else {
      console.warn('News View All arrow image not found.');
    }
  }

  private async _renderNewsAsync(apiUrl: string): Promise<void> {
    try {
      const data: INewsItem[] = await this._getNewsData(apiUrl);

      // Store the data (required later for tab filtering)
      this.newsCentreItems = data;

      this.newsCentreRenderCategory('All');
      this.newsCentreRenderCategory('Announcements');
      this.newsCentreRenderCategory('Events');
      this.newsCentreRenderCategory('News');
      this.newsCentreRenderCategory('Circulars');
    } catch (error) {
      console.error('Error rendering News:', error);
    }
  }

  private newsCentreRenderCategory(category: string): void {

    // 1. Determine the panel
    let panel: HTMLElement | null = null;

    if (category === 'All') {
      panel = this.domElement.querySelector('#news-panel-all') as HTMLElement | null;
    } else if (category === 'Announcements') {
      panel = this.domElement.querySelector('#news-panel-announcements') as HTMLElement | null;
    } else if (category === 'Events') {
      panel = this.domElement.querySelector('#news-panel-events') as HTMLElement | null;
    } else if (category === 'News') {
      panel = this.domElement.querySelector('#news-panel-news') as HTMLElement | null;
    } else if (category === 'Circulars') {
      panel = this.domElement.querySelector('#news-panel-circulars') as HTMLElement | null;
    } else {
      console.error(`Unknown News category: ${category}`);
      return;
    }

    // 2. Check whether panel exists
    if (!panel) {
      console.error(`News panel not found for category: ${category}`);
      return;
    }

    // 3. Find news container
    const container = panel.querySelector('.panel-card-news') as HTMLElement | null;
    if (!container) {
      console.error(`News container not found for category: ${category}`);
      return;
    }

    // 4. Filter news items
    let items: INewsItem[];
    if (category === 'All') {
      items = this.newsCentreItems;
    } else {
      items = this.newsCentreItems.filter((item) => this.newsCentreNormalizeValue(item.Category) === this.newsCentreNormalizeValue(category));
    }

    // 5. Build HTML
    let allElementsHtml: string = '';
    const newsIconImg = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/DefaultImages/news-icon.png`;

    items.forEach((item) => {
      const createddate = this.newsCentreFormatDate(item.PublishedDate);
      const detailsUrl = `${this.context.pageContext.web.absoluteUrl}/Lists/News/DispForm.aspx?ID=${item.Id}`;

      let singleElementHtml = NewsCentre.singleElementHtml;
      singleElementHtml = singleElementHtml.replace(/__KEY_URL_IMGICON__/g, newsIconImg);
      singleElementHtml = singleElementHtml.replace(/__KEY_DATA_TITLE__/g, this.newsCentreEscapeHtml(item.Title || ''));
      singleElementHtml = singleElementHtml.replace(/__KEY_DATA_DESCRIPTION__/g, this.newsCentreEscapeHtml(item.ShortDescription || ''));
      singleElementHtml = singleElementHtml.replace(/__KEY_DATA_CATEGORY__/g, this.newsCentreEscapeHtml(item.Category || ''));
      singleElementHtml = singleElementHtml.replace(/__KEY_DATA_DATE__/g, createddate);
      singleElementHtml = singleElementHtml.replace(/__KEY_URL_LINK__/g, detailsUrl);
      singleElementHtml = singleElementHtml.replace(/__KEY_URL_ARROW__/g, this.newsCentreGetArrowImageUrl());

      allElementsHtml += singleElementHtml;
    });

    // 6. No data
    if (!items.length) {
      allElementsHtml = NewsCentre.noElementHtml;
    }

    // 7. Insert HTML
    container.innerHTML = allElementsHtml;
  }

  private async _getNewsData(apiUrl: string): Promise<INewsItem[]> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });

      if (!response.ok) {
        throw new Error(`News API failed: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      return data.value || [];
    } catch (error) {
      console.error('Error loading News list:', error);
      throw error;
    }
  }

  private newsCentreGetArrowImageUrl(): string {
    const webUrl = this.context.pageContext.web.absoluteUrl;
    return `${webUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;
  }

  private newsCentreAttachTabEvents(): void {
    const tabs = this.domElement.querySelectorAll('#news-tabs .tab-title-pill');
    const panels = this.domElement.querySelectorAll('#news-tabs .news-panel-tab-view');

    if (!tabs.length) {
      console.warn('News tabs not found.');
      return;
    }

    tabs.forEach((tab) => {
      tab.addEventListener('click', () => {
        const targetId = tab.getAttribute('data-tab-news-id');
        if (!targetId) {
          return;
        }

        tabs.forEach((otherTab) => { otherTab.classList.remove('tab-title-pill-active'); });
        tab.classList.add('tab-title-pill-active');

        panels.forEach((panel) => { (panel as HTMLElement).style.display = 'none'; });

        const targetPanel = this.domElement.querySelector(`#${targetId}`) as HTMLElement | null;
        if (targetPanel) {
          targetPanel.style.display = 'block';
        } else {
          console.warn(`News panel not found: ${targetId}`);
          return;
        }

        let category = '';
        if (targetId === 'news-panel-all') {
          category = 'All';
        } else if (targetId === 'news-panel-announcements') {
          category = 'Announcements';
        } else if (targetId === 'news-panel-events') {
          category = 'Events';
        } else if (targetId === 'news-panel-news') {
          category = 'News';
        } else if (targetId === 'news-panel-circulars') {
          category = 'Circulars';
        }

        if (category) {
          this.newsCentreRenderCategory(category);
        }
      });
    });

    // Default tab state: All tab active
    tabs.forEach((tab) => { tab.classList.remove('tab-title-pill-active'); });

    const defaultTab = this.domElement.querySelector('[data-tab-news-id="news-panel-all"]') as HTMLElement | null;
    if (defaultTab) {
      defaultTab.classList.add('tab-title-pill-active');
    }

    // Default panel state: show All panel only
    panels.forEach((panel) => { (panel as HTMLElement).style.display = 'none'; });

    const allPanel = this.domElement.querySelector('#news-panel-all') as HTMLElement | null;
    if (allPanel) {
      allPanel.style.display = 'block';
    }
  }

  private newsCentreFormatDate(value?: string): string {
    if (!value) {
      return '';
    }

    const date = new Date(value);
    if (isNaN(date.getTime())) {
      return '';
    }

    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  private newsCentreNormalizeValue(value?: string): string {
    return (value || '').trim().toLowerCase();
  }

  private newsCentreEscapeHtml(value: string): string {
    return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }

  // ==================== LOAD EVENTS ====================

  private async loadEvents(year: number, month: number, day: number): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const startDate = new Date(year, month, day, 0, 0, 0);
    const endDate = new Date(year, month, day + 1, 0, 0, 0);

    const url = `${siteUrl}/_api/web/lists/getbytitle('Upcoming Events')/items?$select=Id,Title,EventDate,StartTime,EndTime,Location,Status,Link&$filter=EventDate ge datetime'${startDate.toISOString()}' and EventDate lt datetime'${endDate.toISOString()}'&$orderby=EventDate asc`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });

      if (!response.ok) {
        console.error('Upcoming Events list error:', response.status, response.statusText);
        this.events = [];
        return;
      }

      const data = await response.json();
      this.events = data.value || [];
    } catch (error) {
      console.error('Error loading Upcoming Events:', error);
      this.events = [];
    }
  }

  private async loadOrganizationalEventDates(year: number, month: number): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const startDate = new Date(year, month, 1, 0, 0, 0);
    const endDate = new Date(year, month + 1, 1, 0, 0, 0);

    const url = `${siteUrl}/_api/web/lists/getbytitle('Upcoming Events')/items?$select=EventDate,Status,Title&$filter=EventDate ge datetime'${startDate.toISOString()}' and EventDate lt datetime'${endDate.toISOString()}' and Status eq 'Active'`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });

      if (!response.ok) {
        this.organizationalEventDates = [];
        this.organizationalEventTitles = {};
        return;
      }

      const data = await response.json();
      this.organizationalEventDates = [];
      this.organizationalEventTitles = {};

      (data.value || []).forEach((item: { EventDate: string; Status: string; Title: string }) => {
        const dateKey = this.getDateKey(item.EventDate);
        if (!dateKey) {
          return;
        }

        if (this.organizationalEventDates.indexOf(dateKey) === -1) {
          this.organizationalEventDates.push(dateKey);
        }

        const title = item.Title || '';
        if (title) {
          if (this.organizationalEventTitles[dateKey]) {
            this.organizationalEventTitles[dateKey] += `, ${title}`;
          } else {
            this.organizationalEventTitles[dateKey] = title;
          }
        }
      });
    } catch (error) {
      console.error('Error loading organizational event dates:', error);
      this.organizationalEventDates = [];
      this.organizationalEventTitles = {};
    }
  }

 private async loadMyEventDates(
  year: number,
  month: number
): Promise<void> {
  try {
    const client: MSGraphClientV3 =
      await this.context.msGraphClientFactory
        .getClient('3');

    const startDate =
      new Date(
        year,
        month,
        1,
        0,
        0,
        0
      );

    const endDate =
      new Date(
        year,
        month + 1,
        1,
        0,
        0,
        0
      );

    const response =
      await client
        .api('/me/calendar/calendarView')
        .header(
          'Prefer',
          'outlook.timezone="India Standard Time"'
        )
        .query({
          startDateTime:
            startDate.toISOString(),
          endDateTime:
            endDate.toISOString()
        })
        .select('start,subject')
        .orderby('start/dateTime')
        .get();

    this.myEventDates = [];
    this.myEventTitles = {};

    (response.value || []).forEach(
      (event: IOutlookEvent) => {

        if (!event.start?.dateTime) {
          return;
        }

        const dateKey =
          this.getDateKey(
            event.start.dateTime
          );

        if (!dateKey) {
          return;
        }

        if (
          this.myEventDates.indexOf(
            dateKey
          ) === -1
        ) {
          this.myEventDates.push(
            dateKey
          );
        }

        const title =
          event.subject || '';

        if (title) {
          if (
            this.myEventTitles[
              dateKey
            ]
          ) {
            this.myEventTitles[
              dateKey
            ] += `, ${title}`;
          } else {
            this.myEventTitles[
              dateKey
            ] = title;
          }
        }
      }
    );
  } catch (error) {
    console.error(
      'Error loading Outlook event dates:',
      error
    );

    this.myEventDates = [];
    this.myEventTitles = {};
  }
}

  private getDateKey(dateValue: string): string {
    const date = new Date(dateValue);
    if (isNaN(date.getTime())) {
      return '';
    }
    return date.getFullYear() + '-' + date.getMonth() + '-' + date.getDate();
  }

  // ==================== LOAD MY EVENTS ====================

private async loadMyEvents(
  year: number,
  month: number,
  day: number
): Promise<void> {
  try {
    const client: MSGraphClientV3 =
      await this.context.msGraphClientFactory
        .getClient('3');

    const startDate =
      new Date(
        year,
        month,
        day,
        0,
        0,
        0
      );

    const endDate =
      new Date(
        year,
        month,
        day + 1,
        0,
        0,
        0
      );

    const response =
      await client
        .api('/me/calendar/calendarView')
        .header(
          'Prefer',
          'outlook.timezone="India Standard Time"'
        )
        .query({
          startDateTime:
            startDate.toISOString(),
          endDateTime:
            endDate.toISOString()
        })
        .select(
          'id,subject,start,end,location,onlineMeeting'
        )
        .orderby(
          'start/dateTime'
        )
        .get();

    this.myEvents =
      (response.value || []).map(
        (
          event: IOutlookEvent,
          index: number
        ): IEventItem => {

           console.log('EVENT START:', event.start?.dateTime);
    console.log('EVENT START TIMEZONE:', event.start?.timeZone);
    console.log('EVENT END:', event.end?.dateTime);
    console.log('EVENT END TIMEZONE:', event.end?.timeZone);



          return {
            Id:
              index + 1,

            Title:
              event.subject || '',

            EventDate:
              event.start?.dateTime || '',

            StartTime:
              event.start?.dateTime || '',

            EndTime:
              event.end?.dateTime || '',

            Location:
              event.location?.displayName || '',

            Status:
              'Active',

            TeamsUrl:
              event.onlineMeeting?.joinUrl || ''
          };
        }
      );

  } catch (error) {
    console.error(
      'Error loading Outlook Calendar events:',
      error
    );

    this.myEvents = [];
  }
}

  // ==================== RENDER UPCOMING EVENTS ====================

  private renderUpcomingEvents(): void {
    const baseUrl = this.context.pageContext.web.absoluteUrl;
    const rightArrow = `${baseUrl}/SiteAssets/resources/images/icons/right-arrow.png`;
    const arrowRightShort = `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    const UpcomingListingPage=`${this.context.pageContext.web.absoluteUrl}/SitePages/Upcoming-Events.aspx`;

    const myEvents = this.getMyEvents();
    const organizationalEvents = this.getOrganizationalEvents();

    let html = UpcomingEventsTemplate.allElementsHtml;
    html = html.replace(/__KEY_ARROW_RIGHT_SHORT__/g, arrowRightShort);
    html=html.replace(/__KEY_URL_UPCOMING__/g,UpcomingListingPage);

    const myEventsHtml = this.renderEventElements(myEvents, rightArrow, true);
    const organizationalEventsHtml = this.renderEventElements(organizationalEvents, rightArrow, false);

    html = html.replace('id="events-list-my">', `id="events-list-my">${myEventsHtml}`);
    html = html.replace('id="events-list-org">', `id="events-list-org">${organizationalEventsHtml}`);

    this.domElement.querySelector('#upcoming-events-container')!.innerHTML = html;
  }

  private getMyEvents(): IEventItem[] {
    return this.myEvents;
  }

  private getOrganizationalEvents(): IEventItem[] {
    return this.events.filter((event: IEventItem) => event.Status === 'Active');
  }

  private renderEventElements(events: IEventItem[], rightArrow: string, isMyEvent: boolean): string {
    if (!events.length) {
      return UpcomingEventsTemplate.noRecord;
    }

    return events.map((event: IEventItem) => {
      const date = this.formatDate(event.EventDate);
      const time = this.formatTime(event.StartTime, event.EndTime);

      let html = UpcomingEventsTemplate.singleElementHtml;
      html = html.replace('__KEY_EVENT_MONTH__', escape(date.month));
      html = html.replace('__KEY_EVENT_DAY__', escape(date.day));
      html = html.replace('__KEY_EVENT_TITLE__', escape(event.Title || ''));
      html = html.replace('__KEY_EVENT_TIME__', escape(time));
      html = html.replace('__KEY_EVENT_LOCATION__', escape(event.Location || ''));

      let arrowHtml = '';
      if (isMyEvent) {
        const teamsUrl = event.TeamsUrl || 'https://teams.microsoft.com/';
        arrowHtml = `<a href="${escape(teamsUrl)}" target="_blank" data-interception="off" rel="noopener noreferrer"><img src="${rightArrow}" /></a>`;
      } else if (event.Link) {
        arrowHtml = `<a href="${escape(event.Link)}" target="_blank" data-interception="off" rel="noopener noreferrer"><img src="${rightArrow}" /></a>`;
      }

      html = html.replace('__KEY_EVENT_ARROW__', arrowHtml);
      return html;
    }).join('');
  }

  private formatDate(eventDate: string): { month: string; day: string } {
    const date = new Date(eventDate);
    if (isNaN(date.getTime())) {
      return { month: '', day: '' };
    }

    return {
      month: date.toLocaleString('en-US', { month: 'short' }).toUpperCase(),
      day: date.getDate().toString()
    };
  }

  private formatTime(startTime: string, endTime: string): string {
    const start = this.parseSharePointTime(startTime);
    const end = this.parseSharePointTime(endTime);

    if (!start && !end) {
      return '';
    }
    if (!end) {
      return start;
    }
    return `${start} – ${end}`;
  }

private parseSharePointTime(timeValue: string): string {
  if (!timeValue) {
    return '';
  }

  if (timeValue.indexOf('T') !== -1) {
    const timePart = timeValue.split('T')[1];

    if (timePart) {
      const match = timePart.match(/^(\d{1,2}):(\d{2})/);

      if (match) {
        const hours = parseInt(match[1], 10);
        const minutes = parseInt(match[2], 10);

        if (
          hours >= 0 &&
          hours <= 23 &&
          minutes >= 0 &&
          minutes <= 59
        ) {
          const period = hours >= 12 ? 'PM' : 'AM';
          const displayHour =
            hours % 12 === 0 ? 12 : hours % 12;

          return `${('0' + displayHour).slice(-2)}:${('0' + minutes).slice(-2)} ${period}`;
        }
      }
    }
  }

  // Handle SharePoint time values such as 13:00
  const match = timeValue.match(
    /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
  );

  if (match) {
    const hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);

    if (
      hours >= 0 &&
      hours <= 23 &&
      minutes >= 0 &&
      minutes <= 59
    ) {
      const period = hours >= 12 ? 'PM' : 'AM';
      const displayHour =
        hours % 12 === 0 ? 12 : hours % 12;

      return `${('0' + displayHour).slice(-2)}:${('0' + minutes).slice(-2)} ${period}`;
    }
  }

  return timeValue;
}

  // ==================== INITIALIZE EVENT TABS ====================

  private initializeUpcomingEvents(): void {
    const tabs = this.domElement.querySelectorAll('.events-tabs-list .etab');
    const panels = this.domElement.querySelectorAll('.event-calendar-view');

    tabs.forEach((tab: Element) => {
      tab.addEventListener('click', () => {
        const targetId = tab.getAttribute('data-tab-event-id');

        tabs.forEach((item: Element) => { item.classList.remove('etab-active'); });
        panels.forEach((panel: Element) => { (panel as HTMLElement).style.display = 'none'; });
        tab.classList.add('etab-active');

        if (targetId) {
          const selectedPanel = this.domElement.querySelector(`#${targetId}`);
          if (selectedPanel) {
            (selectedPanel as HTMLElement).style.display = 'block';
          }
        }
      });
    });
  }

  // ==================== INITIALIZE CALENDAR ====================

  private initializeCalendar(): void {
    // const $ = (window as Window & { jQuery?: any }).jQuery;
    const $ = (window as Window & { soharJQuery?: any }).soharJQuery;

    if (!$ || !$.fn || !$.fn.datepicker) {
      console.warn('jQuery UI Datepicker is not available.');
      return;
    }

    // MY EVENTS CALENDAR
    $('#events-calendar-my').datepicker({
      dateFormat: 'dd M yy',

      beforeShowDay: (date: Date): [boolean, string, string] => {
        const dateKey = this.getDateKey(date.toISOString());
        const hasEvent = this.myEventDates.indexOf(dateKey) !== -1;
        return hasEvent ? [true, 'has-event', this.myEventTitles[dateKey] || 'Special Event'] : [true, '', ''];
      },

      onSelect: async (dateText: string): Promise<void> => {
        const selectedDate = this.parseCalendarDate(dateText);
        if (!selectedDate) {
          return;
        }

        await this.loadMyEvents(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate());

        const rightArrow = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;
        const myEventsList = this.domElement.querySelector('#events-list-my');
        if (myEventsList) {
          myEventsList.innerHTML = this.renderEventElements(this.getMyEvents(), rightArrow, true);
        }
      },

      onChangeMonthYear: async (year: number, month: number): Promise<void> => {
        await this.loadMyEventDates(year, month - 1);
        $('#events-calendar-my').datepicker('refresh');
      }
    });

    $('#events-calendar-my').find('.ui-datepicker-current-day').removeClass('ui-datepicker-current-day').find('.ui-state-active').removeClass('ui-state-active');

    // ORGANIZATIONAL EVENTS CALENDAR
    $('#events-calendar-org').datepicker({
      dateFormat: 'dd M yy',

      beforeShowDay: (date: Date): [boolean, string, string] => {
        const dateKey = this.getDateKey(date.toISOString());
        const hasEvent = this.organizationalEventDates.indexOf(dateKey) !== -1;
        return hasEvent ? [true, 'has-event', this.organizationalEventTitles[dateKey] || 'Special Event'] : [true, '', ''];
      },

      onSelect: async (dateText: string): Promise<void> => {
        const selectedDate = this.parseCalendarDate(dateText);
        if (!selectedDate) {
          return;
        }

        await this.loadEvents(selectedDate.getFullYear(), selectedDate.getMonth(), selectedDate.getDate());

        const rightArrow = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;
        const organizationalEventsList = this.domElement.querySelector('#events-list-org');
        if (organizationalEventsList) {
          organizationalEventsList.innerHTML = this.renderEventElements(this.getOrganizationalEvents(), rightArrow, false);
        }
      },

      onChangeMonthYear: async (year: number, month: number): Promise<void> => {
        await this.loadOrganizationalEventDates(year, month - 1);
        $('#events-calendar-org').datepicker('refresh');
      }
    });

    $('#events-calendar-org').find('.ui-datepicker-current-day').removeClass('ui-datepicker-current-day').find('.ui-state-active').removeClass('ui-state-active');
  }
  private parseCalendarDate(dateText: string): Date | null {
    const parts = dateText.split(' ');
    if (parts.length !== 3) {
      return null;
    }

    const day = parseInt(parts[0], 10);
    const year = parseInt(parts[2], 10);
    const months: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = months.indexOf(parts[1]);

    if (isNaN(day) || isNaN(year) || month === -1) {
      return null;
    }

    return new Date(year, month, day);
  }

  // ==================== BANNER ====================

  private async _getBannerItems(): Promise<void> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const url = `${siteUrl}/_api/web/lists/getbytitle('Banner')/items?$select=Id,Title,Description,Status,SortOrder,Image&$orderby=SortOrder asc`;

      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });

      if (!response.ok) {
        throw new Error(`Banner list request failed: ${response.status}`);
      }

      const data = await response.json();
      const bannerItems: IBannerItem[] = data.value.filter((item: IBannerItem) => item.Status === 'Active').sort((a: IBannerItem, b: IBannerItem) => a.SortOrder - b.SortOrder);

      this._renderBanner(bannerItems);
    } catch (error) {
      console.error('Error loading Banner list:', error);

      const divBanner = this.domElement.querySelector('#divBanner');
      if (divBanner !== null) {
        divBanner.innerHTML = BannerTemplate.noRecord;
      }
    }
  }

  private _renderBanner(bannerItems: IBannerItem[]): void {
    let allElementsHtml: string = '';

    bannerItems.forEach((item: IBannerItem) => {
      let imageData: any = {};
      if (item.Image) {
        imageData = typeof item.Image === 'string' ? JSON.parse(item.Image) : item.Image;
      }

      const defaultImageUrl = '/sites/DevPortal/SiteAssets/resources/images/bannerDefault/banner-1.png';
      const imageUrl = imageData?.serverRelativeUrl || defaultImageUrl;

      const singleElementHtml = BannerTemplate.singleElementHtml
        .replace('__KEY_BANNER_IMAGE__', imageUrl)
        .replace(/__KEY_BANNER_TITLE__/g, item.Title || '')
        .replace('__KEY_BANNER_DESCRIPTION__', item.Description || '');

      allElementsHtml += singleElementHtml;
    });

    if (allElementsHtml === '') {
      allElementsHtml = BannerTemplate.noRecord;
    }

    const divBanner = this.domElement.querySelector('#divBanner');
    if (divBanner !== null) {
      divBanner.innerHTML = allElementsHtml;
    }
  }

  // ==================== MEDIA GALLERY ====================

  private async _getMediaGalleryItems(): Promise<void> {
    try {
      const siteUrl = this.context.pageContext.web.absoluteUrl;
      const url = `${siteUrl}/_api/web/lists/getbytitle('Media_Gallery')/items?$select=Id,Title,Caption,Status,SortOrder,Image&$orderby=SortOrder asc`;

      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });

      if (!response.ok) {
        throw new Error(`Media Gallery list request failed: ${response.status}`);
      }

      const data = await response.json();
      const galleryItems: IMediaGalleryItem[] = data.value.filter((item: IMediaGalleryItem) => item.Status === 'Active').sort((a: IMediaGalleryItem, b: IMediaGalleryItem) => a.SortOrder - b.SortOrder);

      this._renderMediaGallery(galleryItems);
    } catch (error) {
      console.error('Error loading Media Gallery list:', error);

      const galleryWrapper = this.domElement.querySelector('.gallery-swiper .swiper-wrapper');
      if (galleryWrapper !== null) {
        galleryWrapper.innerHTML = MediaGalleryTemplate.noRecord;
      }
    }
  }

  private _renderMediaGallery(galleryItems: IMediaGalleryItem[]): void {
    let allElementsHtml: string = '';

    galleryItems.forEach((item: IMediaGalleryItem) => {
      let imageData: any = {};
      if (item.Image) {
        imageData = typeof item.Image === 'string' ? JSON.parse(item.Image) : item.Image;
      }

      const imageUrl = imageData.serverRelativeUrl || '';

      const singleElementHtml = MediaGalleryTemplate.singleElementHtml
        .replace('__KEY_GALLERY_IMAGE__', imageUrl)
        .replace('__KEY_GALLERY_CAPTION__', item.Caption || item.Title || '');

      allElementsHtml += singleElementHtml;
    });

    if (allElementsHtml === '') {
      allElementsHtml = MediaGalleryTemplate.noRecord;
    }

    const galleryWrapper = this.domElement.querySelector('.gallery-swiper .swiper-wrapper');

    if (galleryWrapper !== null) {
      galleryWrapper.innerHTML = allElementsHtml;

      // Fill Modal Gallery
      const modalWrapper = this.domElement.querySelector('.gallery-modal-swiper .swiper-wrapper');

      if (modalWrapper !== null) {
        let modalElementsHtml: string = '';

        galleryItems.forEach((item: IMediaGalleryItem) => {
          let imageData: any = {};
          if (item.Image) {
            imageData = typeof item.Image === 'string' ? JSON.parse(item.Image) : item.Image;
          }

          const imageUrl = imageData.serverRelativeUrl || '';

          modalElementsHtml += `
            <div class="swiper-slide gallery-swiper-slide">
              <img src="${imageUrl}" alt="${item.Caption || item.Title || ''}" />
            </div>
          `;
        });

        modalWrapper.innerHTML = modalElementsHtml;
      }

      const galleryElement = this.domElement.querySelector('.gallery-swiper') as HTMLElement & { swiper?: any };
      if (galleryElement && galleryElement.swiper) {
        galleryElement.swiper.update();
      }
    }
  }

  // ==================== ANNOUNCEMENT / OFFERS ====================

  private setupViewAllLink(arrowIconUrl: string): void {
    const baseUrl = this.context.pageContext.web.absoluteUrl;

    const arrowImage = this.domElement.querySelector('#ao-view-all-arrow') as HTMLImageElement;
    if (!arrowImage) {
      console.error('arrow image element not found');
      return;
    }

    // Set arrow image
    arrowImage.src = arrowIconUrl;

    const viewAllLink = this.domElement.querySelector('#ao-view-all') as HTMLAnchorElement;
    if (!viewAllLink) {
      console.error('View All elements not found');
      return;
    }

    // Set default URL
    viewAllLink.href = `${baseUrl}/SitePages/Announcement.aspx`;

    // Change URL when tab changes
    const tabs = this.domElement.querySelectorAll('[data-tab-ao]');
    tabs.forEach((tab) => {
      tab.addEventListener('click', () => {
        const selectedTab = tab.getAttribute('data-tab-ao');

        if (selectedTab === 'announcement') {
          viewAllLink.href = `${baseUrl}/SitePages/Announcement.aspx`;
        } else if (selectedTab === 'offers') {
          viewAllLink.href = `${baseUrl}/SitePages/Offer-List.aspx`;
        }
      });
    });
  }

  private async _renderAnnouncementsAsync(apiUrl: string): Promise<void> {
    try {
      const baseUrl = this.context.pageContext.web.absoluteUrl;
      const imageUrl: string = `${baseUrl}/SiteAssets/resources/images/DefaultImages/announcement-icon.png`;
      const data: IAnnouncement[] = await this._getAnnouncementsData(apiUrl);
      let allElementsHtml: string = "";

      if (!data || data.length === 0) {
        this.domElement.querySelector('#announcement-container')!.innerHTML = AnnouncementOffer.noElementHtml;
        return;
      }

      data.forEach((item) => {
        // const imageData = typeof item.Icon === 'string' ? JSON.parse(item.Icon) : item.Icon;

        // // SharePoint stores the real location of the image here
        // if (imageData?.serverRelativeUrl) {
        //   const serverUrl = imageData.serverUrl || new URL(this.context.pageContext.web.absoluteUrl).origin;
        //   imageUrl = `${serverUrl}${imageData.serverRelativeUrl}`;
        // }

        const createddate = this.formatDates(item.Created);

        const singleElementHtml = AnnouncementOffer.singleElementHtml
          .replace("__KEY__ANNOUNCEMENTOFFER__ICON__", imageUrl)
          .replace("__KEY__ANNOUNCEMENTOFFER__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENTOFFER__DESCRIPTION__", item.ShortDescription)
          .replace("__KEY__ANNOUNCEMENTOFFER__DATE__", createddate);

        allElementsHtml += singleElementHtml;
      });

      this.domElement.querySelector("#announcement-container")!.innerHTML = allElementsHtml;
    } catch (error) {
      console.error('Error rendering QuickList:', error);
    }
  }

  private async _renderOffersAsync(apiUrl: string): Promise<void> {
    try {
      const imageUrl = `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/DefaultImages/offer-icon.png`;
      const data: IOffer[] = await this._getOffersData(apiUrl);
      let allElementsHtml: string = "";

      if (!data || data.length === 0) {
        this.domElement.querySelector('#offer-container')!.innerHTML = AnnouncementOffer.noElementHtml;
        return;
      }

      data.forEach((item) => {
        const createddate = this.formatDates(item.Created);

        const singleElementHtml = AnnouncementOffer.singleElementHtml
          .replace("__KEY__ANNOUNCEMENTOFFER__ICON__", imageUrl)
          .replace("__KEY__ANNOUNCEMENTOFFER__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENTOFFER__DESCRIPTION__", item.Description)
          .replace("__KEY__ANNOUNCEMENTOFFER__DATE__", createddate);

        allElementsHtml += singleElementHtml;
      });

      this.domElement.querySelector("#offer-container")!.innerHTML = allElementsHtml;
    } catch (error) {
      console.error('Error rendering QuickList:', error);
    }
  }

  private async _getAnnouncementsData(apiUrl: string): Promise<IAnnouncement[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

    if (response.ok) {
      const data = await response.json();
      return data.value;
    }

    console.error(`Request failed with status ${response.status}: ${response.statusText}`);
    throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
  }

  private async _getOffersData(apiUrl: string): Promise<IOffer[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

    if (response.ok) {
      const data = await response.json();
      return data.value;
    }

    console.error(`Request failed with status ${response.status}: ${response.statusText}`);
    throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
  }

  // Converts the SharePoint Created date into: Sep 21, 2026
  private formatDates(dateValue: string): string {
    if (!dateValue) {
      return '';
    }

    const date = new Date(dateValue);
    if (isNaN(date.getTime())) {
      return dateValue;
    }

    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }

  // ==================== QUICK LINKS ====================

  private async initializeQuickLinks(): Promise<void> {
    try {
      const [quickLinks] = await Promise.all([
        this.getQuickLinks(this.context),
        this.loadFavourites(this.context)
      ]);

      this.quickLinksCache = quickLinks;
      const favouriteIds = this.getFavouriteIds();

      this.renderQuickLinks(this.context, this.domElement, quickLinks, favouriteIds);
      this.renderFavouriteOptions(this.domElement, quickLinks, favouriteIds);
      this.setupTabs(this.domElement);
      this.setupAddFavouriteButton(this.context, this.domElement);
      this.setupFavouriteModal(this.context, this.domElement);
      this.updateFavouriteTabVisibility(this.domElement, favouriteIds);
    } catch (error) {
      console.error('Error initializing Quick Links:', error);
      const container = this.domElement.querySelector('.quick-links-grid');
      if (container) {
        container.innerHTML = `<div class="p-3">Unable to load Quick Links.</div>`;
      }
    }
  }


  private async getQuickLinks(
    context: WebPartContext
  ): Promise<IQuickLinksList[]> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const apiUrl =
      `${webUrl}` +
      `/_api/web/lists/GetByTitle('Quick_Links')/items` +
      `?$select=Id,Title,Icon,URL,Status,SortOrder` +
      `&$orderby=SortOrder asc`;


    const response =
      await context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {

      console.error(
        `Quick Links request failed: ${response.status}`
      );

      return [];

    }


    const data =
      await response.json();


    return data.value;

  }


  private async getCurrentUser(
    context: WebPartContext
  ): Promise<any> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const response =
      await context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {

      throw new Error(
        `Unable to get current user: ${response.status}`
      );

    }


    return await response.json();

  }
  private async getCurrentUserId(context: WebPartContext): Promise<number> {
    if (this.currentUserId !== null) {
      return this.currentUserId;
    }
    const user = await this.getCurrentUser(context);
    this.currentUserId = user.Id;
    return user.Id;
  }

  // Fetches from the server once and fills the cache
  private async loadFavourites(context: WebPartContext): Promise<void> {
    try {
      const webUrl = context.pageContext.web.absoluteUrl;
      const userId = await this.getCurrentUserId(context);

      const response = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items` +
        `?$select=Id,Quick_x0020_LinkId&$filter=UserId eq ${userId}`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Unable to get user favourites:', response.status);
        this.favouritesCache = [];
        return;
      }

      const data = await response.json();
      this.favouritesCache = data.value.map(
        (item: { Id: number; Quick_x0020_LinkId: number }) => ({
          Id: item.Id,
          LinkId: Number(item.Quick_x0020_LinkId)
        })
      );
    } catch (error) {
      console.error('Error loading user favourites:', error);
      this.favouritesCache = [];
    }
  }

  private getFavouriteIds(): number[] {
    return this.favouritesCache.map((f) => f.LinkId);
  }


  private getQuickLinkImageUrl(
    item: IQuickLinksList
  ): string {

    if (!item.Icon) {
      return '';
    }

    try {

      const imgData =
        typeof item.Icon === 'string'
          ? JSON.parse(item.Icon)
          : item.Icon;

      return imgData.serverRelativeUrl || '';

    } catch (error) {

      console.error(
        'Error parsing Quick Link Icon:',
        error
      );

      return '';

    }

  }


  private renderQuickLinks(context: WebPartContext, domElement: HTMLElement, items: IQuickLinksList[], favouriteIds: number[]): void {
    const container = domElement.querySelector('.quick-links-grid');
    if (!container) {
      console.error('Quick Links container not found.');
      return;
    }

    let allElementsHtml = '';
    const activeItems = items.filter((item) => item.Status === 'Active');

    activeItems.forEach((item) => {
      const imageUrl = this.getQuickLinkImageUrl(item);
      const isFavourite = favouriteIds.indexOf(item.Id) !== -1;
      const favouriteAttribute = isFavourite ? 'data-tab-cat="ql-favourite"' : '';

      const singleElementHtml = QuickLinks.singleElementHtml
        .replace(/__KEY_URL_IMGICON__/, imageUrl)
        .replace(/__KEY_DATA_TITLE__/g, item.Title || '')
        .replace(/__KEY_URL_LINK__/, item.URL?.Url || '#')
        // .replace(/__KEY_URL_TARGET__/, '_blank')
        .replace(/__KEY_DATA_FAVOURITE__/, favouriteAttribute);

      allElementsHtml += singleElementHtml;
    });

    container.innerHTML = allElementsHtml;
  }

  private renderFavouriteOptions(domElement: HTMLElement, items: IQuickLinksList[], favouriteIds: number[]): void {
    const container = domElement.querySelector('.favourite-options');
    if (!container) {
      console.error('Favourite options container not found.');
      return;
    }

    const activeItems = items.filter((item) => item.Status === 'Active');
    let optionsHtml = '';

    activeItems.forEach((item) => {
      const isFavourite = favouriteIds.indexOf(item.Id) !== -1;
      optionsHtml += `
        <label class="favourite-option">
          <input type="checkbox" value="${item.Id}" ${isFavourite ? 'checked' : ''}>
          <span>${item.Title}</span>
        </label>
      `;
    });

    container.innerHTML = optionsHtml;
  }
  private async createFavourite(
    context: WebPartContext,
    userId: number,
    quickLinkId: number
  ): Promise<number> {
    const webUrl = context.pageContext.web.absoluteUrl;

    const response = await context.spHttpClient.post(
      `${webUrl}/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({
          UserId: userId,
          Quick_x0020_LinkId: quickLinkId
        })
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to add favourite: ${response.status}`);
    }

    const created = await response.json();
    return created.Id;
  }

  private setupTabs(domElement: HTMLElement): void {
    const tabs = domElement.querySelectorAll('[data-filter-ql]');

    tabs.forEach((tab) => {
      tab.addEventListener('click', () => {
        const selectedTab = (tab as HTMLElement).getAttribute('data-filter-ql');
        const allLinks = domElement.querySelectorAll('.quick-links-grid .quick-link-box');

        tabs.forEach((item) => { item.classList.remove('panel-title-filter-active'); });
        tab.classList.add('panel-title-filter-active');

        allLinks.forEach((link) => {
          const favouriteCategory = link.getAttribute('data-tab-cat');

          if (selectedTab === 'favourites') {
            (link as HTMLElement).style.display = favouriteCategory === 'ql-favourite' ? '' : 'none';
          } else {
            (link as HTMLElement).style.display = '';
          }
        });
      });
    });
  }

  private updateFavouriteTabVisibility(domElement: HTMLElement, favouriteIds: number[]): void {
    const favouriteTab = domElement.querySelector('[data-filter-ql="favourites"]') as HTMLElement;
    if (!favouriteTab) {
      return;
    }
    favouriteTab.style.display = favouriteIds.length === 0 ? 'none' : '';
  }

  private setupAddFavouriteButton(context: WebPartContext, domElement: HTMLElement): void {
    const button = domElement.querySelector('#btnAddFavourites');
    if (!button) {
      console.error('Add Favourites button not found.');
      return;
    }

    button.addEventListener('click', async () => {
      await this.addFavourites(context, domElement);
    });
  }

  private setupFavouriteModal(context: WebPartContext, domElement: HTMLElement): void {
    const modal = domElement.querySelector('#addFavouriteModal');
    if (!modal) {
      console.error('Favourite modal not found.');
      return;
    }

    modal.addEventListener('show.bs.modal', async () => {
      const favouriteIds = await this.getFavouriteIds();
      const checkboxes = domElement.querySelectorAll('.favourite-options input[type="checkbox"]');

      checkboxes.forEach((checkbox) => {
        const input = checkbox as HTMLInputElement;
        const quickLinkId = Number(input.value);
        input.checked = favouriteIds.indexOf(quickLinkId) !== -1;
      });
    });
  }

  private async addFavourites(
    context: WebPartContext,
    domElement: HTMLElement
  ): Promise<void> {

    // 1. Read selection
    const selectedIds: number[] = [];
    domElement
      .querySelectorAll('.favourite-options input[type="checkbox"]')
      .forEach((checkbox) => {
        const input = checkbox as HTMLInputElement;
        if (input.checked) {
          selectedIds.push(Number(input.value));
        }
      });

    // 2. Close modal immediately
    const modal = domElement.querySelector('#addFavouriteModal') as HTMLElement;
    if (modal) {
      const bootstrap = (window as any).bootstrap;
      if (bootstrap?.Modal) {
        bootstrap.Modal.getOrCreateInstance(modal).hide();
      } else {
        modal.classList.remove('show');
        modal.style.display = 'none';
        modal.setAttribute('aria-hidden', 'true');
        document.body.classList.remove('modal-open');
        document.querySelectorAll('.modal-backdrop').forEach((b) => b.remove());
      }
    }

    // 3. Work out what changed using the cache (no network)
    const toDelete = this.favouritesCache.filter(
      (f) => selectedIds.indexOf(f.LinkId) === -1
    );
    const existingLinkIds = this.favouritesCache.map((f) => f.LinkId);
    const toAdd = selectedIds.filter(
      (id) => existingLinkIds.indexOf(id) === -1
    );

    if (toDelete.length === 0 && toAdd.length === 0) {
      return;
    }

    // 4. Update the UI right away (optimistic)
    const activeTabValue = (
      domElement.querySelector('.panel-title-filter-active') as HTMLElement
    )?.getAttribute('data-filter-ql');

    this.refreshFavouritesUI(domElement, selectedIds, activeTabValue);

    // 5. Sync with SharePoint in parallel
    try {
      const userId = await this.getCurrentUserId(context);

      const deletePromises = toDelete.map((f) =>
        this.deleteFavourite(context, f.Id)
      );

      const addPromises = toAdd.map(async (linkId) => ({
        Id: await this.createFavourite(context, userId, linkId),
        LinkId: linkId
      }));

      const [, created] = await Promise.all([
        Promise.all(deletePromises),
        Promise.all(addPromises)
      ]);

      // 6. Update cache with the new server IDs
      const deletedIds = toDelete.map((f) => f.Id);
      this.favouritesCache = this.favouritesCache
        .filter((f) => deletedIds.indexOf(f.Id) === -1)
        .concat(created);

    } catch (error) {
      console.error('Error updating favourites:', error);

      // Roll back to the real server state
      await this.loadFavourites(context);
      this.refreshFavouritesUI(
        domElement,
        this.getFavouriteIds(),
        activeTabValue
      );
    }
  }

  private refreshFavouritesUI(
    domElement: HTMLElement,
    favouriteIds: number[],
    activeTabValue?: string | null
  ): void {
    this.renderQuickLinks(this.context, domElement, this.quickLinksCache, favouriteIds);
    this.renderFavouriteOptions(domElement, this.quickLinksCache, favouriteIds);
    this.updateFavouriteTabVisibility(domElement, favouriteIds);

    if (favouriteIds.length === 0) {
      (domElement.querySelector('[data-filter-ql="quick-links"]') as HTMLElement)?.click();
    } else if (activeTabValue) {
      (domElement.querySelector(`[data-filter-ql="${activeTabValue}"]`) as HTMLElement)?.click();
    }
  }

  private async deleteFavourite(context: WebPartContext, favouriteId: number): Promise<void> {
    const webUrl = context.pageContext.web.absoluteUrl;

    const response = await context.spHttpClient.post(`${webUrl}/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items(${favouriteId})`, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata', 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' }
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('Failed to delete favourite:', favouriteId, response.status, errorText);
      throw new Error(`Failed to delete favourite: ${response.status}`);
    }
  }

  // ==================== BIRTHDAYS ====================

  private async getBirthdays(context: WebPartContext): Promise<IBirthdayList[]> {
    const webUrl = context.pageContext.web.absoluteUrl;
    const endpoint = `${webUrl}/_api/web/lists/getbytitle('Birthday')/items?$select=Id,Title,EmployeePhoto,BirthDate,Status`;

    try {
      const response = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });

      if (!response.ok) {
        console.error('Failed to get birthday data:', response.status, response.statusText);
        return [];
      }

      const data = await response.json();
      return data.value;
    } catch (error) {
      console.error('Error getting birthday data:', error);
      return [];
    }
  }

  private getBirthdayImageUrl(photo: any): string {
    if (!photo) {
      return '';
    }

    try {
      if (typeof photo === 'string') {
        photo = JSON.parse(photo);
      }
      return photo.serverRelativeUrl || '';
    } catch (error) {
      console.error('Error parsing employee photo:', error);
      return '';
    }
  }

  private getUpcomingBirthdays(birthdays: IBirthdayList[]): IBirthdayList[] {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const upcoming = birthdays
      .filter((item) => item.Status === 'Active')
      .map((item) => {
        const birthDate = new Date(item.BirthDate);
        const birthday = new Date(today.getFullYear(), birthDate.getMonth(), birthDate.getDate());

        if (birthday < today) {
          birthday.setFullYear(today.getFullYear() + 1);
        }

        const difference = Math.floor((birthday.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        return { item, difference };
      })
      .filter((item) => item.difference < 7)
      .sort((a, b) => a.difference - b.difference);

    return upcoming.map((item) => item.item);
  }

  private async renderBirthdays(): Promise<void> {
    const birthdays = await this.getBirthdays(this.context);
    const upcomingBirthdays = this.getUpcomingBirthdays(birthdays);

    const container = this.domElement.querySelector('.panel-card-birthdays');
    if (!container) {
      console.error('Birthday container not found.');
      return;
    }

    let html = '';

    upcomingBirthdays.forEach((item) => {
      const birthDate = new Date(item.BirthDate);
      const month = birthDate.toLocaleString('en-US', { month: 'short' }).toUpperCase();
      const day = birthDate.getDate().toString();
      const imageUrl = this.getBirthdayImageUrl(item.EmployeePhoto);

      html += Birthday.singleElementHtml
        .replace(/__KEY_URL_IMG__/g, imageUrl)
        .replace(/__KEY_TITLE__/g, item.Title)
        .replace(/__KEY_MONTH__/g, month)
        .replace(/__KEY_DAY__/g, day);
    });

    container.innerHTML = html;
  }

  // HR API - Birthdays
  // private async getHRBirthdays(): Promise<any[]> {
  //   const apiUrl = 'HR_API_URL_WILL_BE_PROVIDED';
  //   // API implementation will be added once HR provides endpoint and authentication details.
  //   return [];
  // }

  // ==================== SOCIAL MEDIA ====================

  private async initializeSocialMedia(): Promise<void> {

    // Instagram
    const instagramContainer = this.domElement.querySelector('#social-instagram') as HTMLElement;
    if (instagramContainer) {
      instagramContainer.innerHTML = `<div class="sk-instagram-feed" data-embed-id="25716225"></div>`;
      const script = document.createElement('script');
      script.src = 'https://widgets.sociablekit.com/instagram-feed/widget.js';
      script.defer = true;
      instagramContainer.appendChild(script);
    }

    // X (Twitter)
    const twitterContainer = this.domElement.querySelector('#social-twitter') as HTMLElement;
    if (twitterContainer) {
      twitterContainer.innerHTML = `<div class="sk-ww-twitter-feed" data-embed-id="25716230"></div>`;
      const script = document.createElement('script');
      script.src = 'https://widgets.sociablekit.com/twitter-feed/widget.js';
      script.defer = true;
      twitterContainer.appendChild(script);
    }

    // LinkedIn
    const linkedinContainer = this.domElement.querySelector('#social-linkedin') as HTMLElement;
    if (linkedinContainer) {
      linkedinContainer.innerHTML = `<div class="sk-ww-linkedin-page-post" data-embed-id="25716233"></div>`;
      const script = document.createElement('script');
      script.src = 'https://widgets.sociablekit.com/linkedin-page-posts/widget.js';
      script.defer = true;
      linkedinContainer.appendChild(script);
    }

    // Facebook
    const facebookContainer = this.domElement.querySelector('#social-facebook') as HTMLElement;
    if (facebookContainer) {
      facebookContainer.innerHTML = `<div class="sk-ww-facebook-page-posts" data-embed-id="25716238"></div>`;
      const script = document.createElement('script');
      script.src = 'https://widgets.sociablekit.com/facebook-page-posts/widget.js';
      script.defer = true;
      facebookContainer.appendChild(script);
    }

    // YouTube
    const youtubeContainer = this.domElement.querySelector('#social-youtube') as HTMLElement;
    if (youtubeContainer) {
      youtubeContainer.innerHTML = `<div class="sk-ww-youtube-channel-videos" data-embed-id="25716248"></div>`;
      const script = document.createElement('script');
      script.src = 'https://widgets.sociablekit.com/youtube-channel-videos/widget.js';
      script.defer = true;
      youtubeContainer.appendChild(script);
    }

    this.setupSocialMediaTutorialLinkObserver();
  }

  // ==================== INITIALIZE WEB PART ====================

  protected async onInit(): Promise<void> {
    return super.onInit();
  }

  // ==================== PROPERTY PANE ====================

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Home Page' },
          groups: [
            {
              groupName: 'Configuration',
              groupFields: [
                PropertyPaneTextField('description', { label: 'Description' })
              ]
            }
          ]
        }
      ]
    };
  }

  // ==================== DATA VERSION ====================

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}