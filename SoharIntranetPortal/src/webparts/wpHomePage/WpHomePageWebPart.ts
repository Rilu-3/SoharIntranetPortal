import { Version } from '@microsoft/sp-core-library';

import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import {
  BaseClientSideWebPart
} from '@microsoft/sp-webpart-base';

import {
  SPHttpClient,
  SPHttpClientResponse,
  MSGraphClientV3
} from '@microsoft/sp-http';

import {
  SPComponentLoader
} from '@microsoft/sp-loader';

import { escape } from '@microsoft/sp-lodash-subset';

import UpcomingEventsTemplate from './UpcomingEvents';
import { BannerTemplate } from './BannerTemplate';
import MediaGalleryTemplate from './MediaGalleryTemplate';
import UabAnnouncements from './UabAnnouncements';
import UabOffers from './UabOffers';

import UabOfferAnnouncementWrapper from './UabOfferAnnouncementWrapper';


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


// Interface used for both Outlook My Events
// and SharePoint Organizational Events
export interface IEventItem {
  Id: number;

  Title: string;

  EventDate: string;

  StartTime: string;

  EndTime: string;

  Location: string;

  Status: string;
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
  OfferImage: string;
  Created: string;
}

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  // Stores Organizational Events
  // fetched from the SharePoint Upcoming Events list
  private events: IEventItem[] = [];


  // Stores My Events
  // fetched from the logged-in user's Outlook calendar
  private myEvents: IEventItem[] = [];


  // ==================== RENDER ====================

  public async render(): Promise<void> {

    const workbenchContent =
      document.getElementById('workbenchPageContent');

    if (workbenchContent) {
      workbenchContent.style.maxWidth = 'none';
    }


    // Get the current date
    const today = new Date();


    // Load both event sources at the same time
    await Promise.all([

      // Load Organizational Events from SharePoint
      this.loadEvents(
        today.getFullYear(),
        today.getMonth()
      ),

      // Load My Events from Outlook Calendar
      this.loadMyEvents(
        today.getFullYear(),
        today.getMonth()
      )

    ]);

    const arrowIconUrl =
    `${this.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    /*
     * Create the complete page HTML first.
     *
     * This is important because Banner,
     * Upcoming Events and Media Gallery
     * must have their own containers.
     */
    this.domElement.innerHTML =
      BannerTemplate.bannerHtml +
      UpcomingEventsTemplate.allElementsHtml +
      MediaGalleryTemplate.allElementsHtml +
      MediaGalleryTemplate.galleryModalHtml+   UabOfferAnnouncementWrapper.allElementsHtml.replace(
      "__KEY__ARROW__RIGHT__ICON__",
      arrowIconUrl
    );

      ;


    // Create Upcoming Events inside its existing containers
    this.renderUpcomingEvents();


    // Initialize My Events / Organizational Events tabs
    this.initializeUpcomingEvents();


    // Initialize both calendars
    this.initializeCalendar();




 
    this.domElement.querySelector("#announcement")!.innerHTML =
      UabAnnouncements.allElementsHtml;

    this.domElement.querySelector("#offers")!.innerHTML =
      UabOffers.allElementsHtml;

    this.setupViewAllLink();

    const AnnouncementApiUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Announcements')/items?$select=Id,Title,ShortDescription,Icon,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc&$top=3`;

    const OfferApiUrl =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Offers')/items?$select=Id,Title,Description,OfferImage,Created,Status&$filter=Status eq 'Active'&$orderby=Created desc&$top=3`;

    await this._renderAnnouncementsAsync(AnnouncementApiUrl);

    await this._renderOffersAsync(OfferApiUrl);


    // Load Banner items from SharePoint
    await this._getBannerItems();


    // Load Active Media Gallery items from SharePoint
    await this._getMediaGalleryItems();


    // Load home.js only after the HTML is available
    await this.loadHomeJS();

  }


  // ==================== LOAD CSS ====================

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


  // ==================== LOAD JS ====================

  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );


    await this.loadBootstrap();


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

  }


  // ==================== LOAD BOOTSTRAP ====================

  private async loadBootstrap(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    if (
      typeof (window as any).bootstrap !== 'undefined'
    ) {

      return;

    }


    const bootstrapModule =
      await SPComponentLoader.loadScript<any>(
        `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
      );


    if (bootstrapModule) {

      (window as any).bootstrap =
        bootstrapModule;

    }


    console.log(
      'Bootstrap loaded:',
      typeof (window as any).bootstrap
    );

  }


  // ==================== LOAD HOME JS ====================

  private async loadHomeJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    try {

      await SPComponentLoader.loadScript(
        `${baseUrl}/SiteAssets/resources/js/home.js`
      );


      console.log(
        'home.js loaded'
      );

    } catch (error) {

      console.error(
        'Error loading home.js:',
        error
      );

    }

  }


  // ==================== LOAD EVENTS ====================

  private async loadEvents(
    year: number,
    month: number
  ): Promise<void> {

    // Get the current SharePoint site URL
    const siteUrl =
      this.context.pageContext.web.absoluteUrl;


    /*
     * First day of selected month
     */
    const startDate =
      new Date(
        year,
        month,
        1
      );


    /*
     * First day of next month
     */
    const endDate =
      new Date(
        year,
        month + 1,
        1
      );


    // Convert the dates into ISO format
    // for the SharePoint REST API
    const startDateString =
      startDate.toISOString();


    const endDateString =
      endDate.toISOString();


    // SharePoint REST API URL
    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Upcoming Events')/items` +
      `?$select=Id,Title,EventDate,StartTime,EndTime,Location,Status` +
      `&$filter=EventDate ge datetime'${startDateString}' and EventDate lt datetime'${endDateString}'` +
      `&$orderby=EventDate asc`;


    try {

      // Send GET request to the SharePoint REST API
      const response:
        SPHttpClientResponse =
        await this.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept:
                'application/json;odata=nometadata'
            }
          }
        );


      // Check whether the API request was successful
      if (!response.ok) {

        console.error(
          'Upcoming Events list error:',
          response.status,
          response.statusText
        );


        // Clear events if the request fails
        this.events = [];


        return;

      }


      // Convert API response into JSON
      const data =
        await response.json();


      // Store SharePoint events
      this.events =
        data.value || [];


      // Display the fetched events in the console
      console.log(
        'Organizational Events:',
        this.events
      );


    } catch (error) {

      // Handle SharePoint API errors
      console.error(
        'Error loading Upcoming Events:',
        error
      );


      // Clear events when an error occurs
      this.events = [];

    }

  }


  // ==================== LOAD MY EVENTS ====================

  private async loadMyEvents(
    year: number,
    month: number
  ): Promise<void> {

    try {

      // Create Microsoft Graph client
      const client:
        MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient('3');


      // First day of selected month
      const startDate =
        new Date(
          year,
          month,
          1,
          0,
          0,
          0
        );


      // First day of next month
      const endDate =
        new Date(
          year,
          month + 1,
          1,
          0,
          0,
          0
        );


      // Call Microsoft Graph Calendar API
      const response =
        await client
          .api('/me/calendar/calendarView')

          .query({
            startDateTime:
              startDate.toISOString(),

            endDateTime:
              endDate.toISOString()
          })

          .select(
            'id,subject,start,end,location'
          )

          .orderby(
            'start/dateTime'
          )

          .get();


      // Display Outlook events in the browser console
      console.log(
        'Outlook Calendar Events:',
        response.value
      );


      // Convert Outlook events into the
      // common IEventItem format
      this.myEvents =
        (response.value || []).map(
          (
            event: any,
            index: number
          ): IEventItem => {

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
                'Active'

            };

          }
        );


    } catch (error) {

      // Handle Microsoft Graph errors
      console.error(
        'Error loading Outlook Calendar events:',
        error
      );


      // Clear Outlook events if an error occurs
      this.myEvents = [];

    }

  }


  // ==================== RENDER UPCOMING EVENTS ====================

  private renderUpcomingEvents(): void {

    // Get the SharePoint site URL
    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    // Image used for the event arrow
    const rightArrow =
      `${baseUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


    // Image used for the short arrow
    const arrowRightShort =
      `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;


    // Get future My Events
    const myEvents =
      this.getMyEvents();


    // Get future Organizational Events
    const organizationalEvents =
      this.getOrganizationalEvents();


    // Generate HTML for My Events
    const myEventsHtml =
      this.renderEventElements(
        myEvents,
        rightArrow
      );


    // Generate HTML for Organizational Events
    const organizationalEventsHtml =
      this.renderEventElements(
        organizationalEvents,
        rightArrow
      );


    /*
     * Insert My Events into the existing
     * Upcoming Events container.
     */
    const myEventsList =
      this.domElement.querySelector(
        '#events-list-my'
      );


    if (myEventsList) {

      myEventsList.innerHTML =
        myEventsHtml;

    }


    /*
     * Insert Organizational Events into
     * the existing Upcoming Events container.
     */
    const organizationalEventsList =
      this.domElement.querySelector(
        '#events-list-org'
      );


    if (organizationalEventsList) {

      organizationalEventsList.innerHTML =
        organizationalEventsHtml;

    }


    /*
     * Replace arrow placeholder without
     * replacing the complete page HTML.
     */
    const eventCalendarViews =
      this.domElement.querySelectorAll(
        '.event-calendar-view'
      );


    eventCalendarViews.forEach(
      (view: Element) => {

        view.innerHTML =
          view.innerHTML.replace(
            /__KEY_ARROW_RIGHT_SHORT__/g,
            arrowRightShort
          );

      }
    );

  }


  // ==================== GET MY EVENTS ====================

  private getMyEvents(): IEventItem[] {

    // Get today's date
    const today = new Date();


    // Remove the current time
    today.setHours(
      0,
      0,
      0,
      0
    );


    // Filter Outlook events
    return this.myEvents.filter(
      (event: IEventItem) => {

        // Convert event date into JavaScript Date
        const eventDate =
          new Date(event.EventDate);


        // Remove the time from the event date
        eventDate.setHours(
          0,
          0,
          0,
          0
        );


        // Keep today's and future events
        return eventDate >= today;

      }
    );

  }


  // ==================== GET ORGANIZATIONAL EVENTS ====================

  private getOrganizationalEvents(): IEventItem[] {

    // Get today's date
    const today = new Date();


    // Remove the current time
    today.setHours(
      0,
      0,
      0,
      0
    );


    // Filter SharePoint events
    return this.events.filter(
      (event: IEventItem) => {

        // Ignore events that are not Active
        if (event.Status !== 'Active') {
          return false;
        }


        // Convert event date into JavaScript Date
        const eventDate =
          new Date(event.EventDate);


        // Remove the time from the event date
        eventDate.setHours(
          0,
          0,
          0,
          0
        );


        // Keep today's and future events
        return eventDate >= today;

      }
    );

  }


  // ==================== RENDER EVENT ELEMENTS ====================

  private renderEventElements(
    events: IEventItem[],
    rightArrow: string
  ): string {

    /*
     * No records.
     */
    if (!events.length) {

      return UpcomingEventsTemplate.noRecord;

    }


    /*
     * Show maximum 2 events.
     */
    return events
      .slice(0, 2)
      .map(
        (event: IEventItem) => {

          // Format the event date
          const date =
            this.formatDate(
              event.EventDate
            );


          // Format the event time
          const time =
            this.formatTime(
              event.StartTime,
              event.EndTime
            );


          // Get the individual event HTML template
          let html =
            UpcomingEventsTemplate.singleElementHtml;


          // Replace event month
          html =
            html.replace(
              '__KEY_EVENT_MONTH__',
              escape(date.month)
            );


          // Replace event day
          html =
            html.replace(
              '__KEY_EVENT_DAY__',
              escape(date.day)
            );


          // Replace event title
          html =
            html.replace(
              '__KEY_EVENT_TITLE__',
              escape(event.Title || '')
            );


          // Replace event time
          html =
            html.replace(
              '__KEY_EVENT_TIME__',
              escape(time)
            );


          // Replace event location
          html =
            html.replace(
              '__KEY_EVENT_LOCATION__',
              escape(event.Location || '')
            );


          // Replace event arrow image
          html =
            html.replace(
              '__KEY_EVENT_ARROW__',
              rightArrow
            );


          return html;

        }
      )
      .join('');

  }


  // ==================== FORMAT DATE ====================

  private formatDate(
    eventDate: string
  ): {
    month: string;
    day: string;
  } {

    // Convert string into Date object
    const date =
      new Date(eventDate);


    // Check whether the date is valid
    if (isNaN(date.getTime())) {

      return {
        month: '',
        day: ''
      };

    }


    return {

      // Get short month name
      month:
        date
          .toLocaleString(
            'en-US',
            {
              month: 'short'
            }
          )
          .toUpperCase(),


      // Get day number
      day:
        date
          .getDate()
          .toString()

    };

  }


  // ==================== FORMAT TIME ====================

  private formatTime(
    startTime: string,
    endTime: string
  ): string {

    // Convert start time
    const start =
      this.parseSharePointTime(
        startTime
      );


    // Convert end time
    const end =
      this.parseSharePointTime(
        endTime
      );


    // If both times are empty
    if (!start && !end) {

      return '';

    }


    // If only start time exists
    if (!end) {

      return start;

    }


    // Display start and end time together
    return `${start} – ${end}`;

  }


  // ==================== PARSE TIME ====================

  private parseSharePointTime(
    timeValue: string
  ): string {

    // Return empty value if no time is provided
    if (!timeValue) {

      return '';

    }


    // Check whether the value is an ISO date/time
    if (timeValue.indexOf('T') !== -1) {

      // Convert ISO value into Date object
      const date =
        new Date(timeValue);


      // Check whether the date is valid
      if (!isNaN(date.getTime())) {

        // Convert into 12-hour time
        return date.toLocaleTimeString(
          'en-US',
          {
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
          }
        );

      }

    }


    /*
     * Handle HH:mm or HH:mm:ss.
     */
    const match =
      timeValue.match(
        /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
      );


    if (match) {

      // Extract hour
      const hours =
        parseInt(
          match[1],
          10
        );


      // Extract minutes
      const minutes =
        parseInt(
          match[2],
          10
        );


      // Validate hour and minute values
      if (
        hours >= 0 &&
        hours <= 23 &&
        minutes >= 0 &&
        minutes <= 59
      ) {

        // Decide AM or PM
        const period =
          hours >= 12
            ? 'PM'
            : 'AM';


        // Convert 24-hour hour into 12-hour hour
        const displayHour =
          hours % 12 === 0
            ? 12
            : hours % 12;


        // Return formatted time
        return (
          `${('0' + displayHour).slice(-2)}:` +
          `${('0' + minutes).slice(-2)} ` +
          `${period}`
        );

      }

    }


    // Return original value if it cannot be parsed
    return timeValue;

  }


  // ==================== INITIALIZE TABS ====================

  private initializeUpcomingEvents(): void {

    // Find all event tabs
    const tabs =
      this.domElement.querySelectorAll(
        '.events-tabs-list .etab'
      );


    // Find all event panels
    const panels =
      this.domElement.querySelectorAll(
        '.event-calendar-view'
      );


    // Add click event to each tab
    tabs.forEach(
      (tab: Element) => {

        tab.addEventListener(
          'click',
          () => {

            // Get the panel ID connected to the clicked tab
            const targetId =
              tab.getAttribute(
                'data-tab-event-id'
              );


            /*
             * Remove active class
             * from all tabs.
             */
            tabs.forEach(
              (item: Element) => {

                item.classList.remove(
                  'etab-active'
                );

              }
            );


            panels.forEach(
              (panel: Element) => {

                (
                  panel as HTMLElement
                ).style.display = 'none';

              }
            );


            /*
             * Activate selected tab.
             */
            tab.classList.add(
              'etab-active'
            );


            /*
             * Show selected panel.
             */
            if (targetId) {

              // Find the selected panel
              const selectedPanel =
                this.domElement.querySelector(
                  `#${targetId}`
                );


              if (selectedPanel) {

                // Display the selected panel
                (
                  selectedPanel as HTMLElement
                ).style.display = 'block';

              }

            }

          }
        );

      }
    );

  }


  // ==================== INITIALIZE CALENDAR ====================

  private initializeCalendar(): void {

    // Get jQuery from the global window object
    const $ =
      (window as any).jQuery;


    // Check whether jQuery UI Datepicker is available
    if (
      !$ ||
      !$.fn ||
      !$.fn.datepicker
    ) {

      console.warn(
        'jQuery UI Datepicker is not available.'
      );


      return;

    }


    // Store current web part instance
    // so it can be accessed inside callback functions
    const self = this;


    /*
     * My Events
     */
    $('#events-calendar-my').datepicker({

      // Date display format
      dateFormat: 'dd M yy',


      // Runs when the user changes the calendar month
      onChangeMonthYear:
        async function (
          year: number,
          month: number
        ): Promise<void> {

          // Reload Outlook events for selected month
          await self.loadMyEvents(
            year,
            month - 1
          );


          // Get event arrow image
          const rightArrow =
            `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


          // Rebuild My Events HTML
          const myEventsHtml =
            self.renderEventElements(
              self.getMyEvents(),
              rightArrow
            );


          // Find My Events list container
          const myEventsList =
            self.domElement.querySelector(
              '#events-list-my'
            );


          // Replace old events with new events
          if (myEventsList) {

            myEventsList.innerHTML =
              myEventsHtml;

          }

        }

    });


    /*
     * Organizational Events
     */
    $('#events-calendar-org').datepicker({

      // Date display format
      dateFormat: 'dd M yy',


      // Runs when the user changes the calendar month
      onChangeMonthYear:
        async function (
          year: number,
          month: number
        ): Promise<void> {

          // Reload SharePoint events for selected month
          await self.loadEvents(
            year,
            month - 1
          );


          // Get event arrow image
          const rightArrow =
            `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


          // Rebuild Organizational Events HTML
          const organizationalEventsHtml =
            self.renderEventElements(
              self.getOrganizationalEvents(),
              rightArrow
            );


          // Find Organizational Events list container
          const organizationalEventsList =
            self.domElement.querySelector(
              '#events-list-org'
            );


          // Replace old events with new events
          if (organizationalEventsList) {

            organizationalEventsList.innerHTML =
              organizationalEventsHtml;

          }

        }

    });

  }


  // ==================== GET BANNER ITEMS ====================

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


  // ==================== RENDER BANNER ====================

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


  // ==================== GET MEDIA GALLERY ITEMS ====================

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


  // ==================== RENDER MEDIA GALLERY ====================

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


      /*
       * Fill Modal Gallery
       */
      const modalWrapper =
        this.domElement.querySelector(
          '.gallery-modal-swiper .swiper-wrapper'
        );


      if (modalWrapper !== null) {

        let modalElementsHtml: string = '';


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


            modalElementsHtml += `
              <div class="swiper-slide gallery-swiper-slide">
                <img
                  src="${imageUrl}"
                  alt="${item.Caption || item.Title || ''}"
                />
              </div>
            `;

          }
        );


        modalWrapper.innerHTML =
          modalElementsHtml;

      }


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
  private setupViewAllLink(): void {
  const tabs = this.domElement.querySelectorAll('[data-tab-ao]');

  const viewAllLink =
    this.domElement.querySelector('#ao-view-all') as HTMLAnchorElement;

  if (!viewAllLink) {
    console.error('View All link not found');
    return;
  }

  // Set the initial URL for the default Announcements tab
  viewAllLink.href =
    `${this.context.pageContext.web.absoluteUrl}/SitePages/Announcement.aspx`;

  tabs.forEach((tab) => {
    tab.addEventListener('click', () => {

      const selectedTab = tab.getAttribute('data-tab-ao');

      if (selectedTab === 'announcement') {
        viewAllLink.href =
          `${this.context.pageContext.web.absoluteUrl}/SitePages/Announcement.aspx`;
      }

      if (selectedTab === 'offers') {
        viewAllLink.href =
          `${this.context.pageContext.web.absoluteUrl}/SitePages/Offer-List.aspx`;
      }
    });
  });
}
  /*
   * ==========================================================
   * RENDER ANNOUNCEMENTS
   * ==========================================================
   *
   * Gets Announcement data from SharePoint and converts each
   * SharePoint item into the Announcement HTML template.
   */
  private async _renderAnnouncementsAsync(apiUrl: string): Promise<void> {

    const data: IAnnouncement[] =
      await this._getAnnouncementsData(apiUrl);

    console.log("Announcements data", data);

    let allElementsHtml: string = "";

    try {

      data.forEach((item, index) => {

        let imageUrl = '';

        /*
         * The Icon column contains JSON data.
         * Extract the file name from the JSON and construct
         * the SharePoint attachment URL.
         */
        if (item.Icon) {

          const imageData = JSON.parse(item.Icon);

          // console.log(imageData);

          const fileName = imageData.fileName;

          // Build the image URL using the fileName
          imageUrl =
            `${this.context.pageContext.web.absoluteUrl}/Lists/Announcements/Attachments/${item.Id}/${fileName}`;

          // console.log(imageUrl);
        }

        /*
         * Convert the SharePoint Created date into the
         * required display format.
         */
        let createddate = this.formatDates(item.Created);

        /*
         * Replace the placeholders in the Announcement
         * HTML template with actual SharePoint data.
         */
        let singleElementHtml = UabAnnouncements.singleElementHtml
          .replace("__KEY__ANNOUNCEMENT__ICON__", imageUrl)
          .replace("__KEY__ANNOUNCEMENT__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENT__DESCRIPTION__", item.ShortDescription)
          .replace("__KEY__ANNOUNCEMENT__DATE__", createddate);

        /*
         * Add the generated Announcement HTML to the
         * complete HTML string.
         */
        allElementsHtml += singleElementHtml;
      })

    }
    catch (error) {
      console.error('Error rendering QuickList:', error);
    }

    /*
     * Insert all generated Announcement HTML into the
     * Announcement container.
     */
    this.domElement.querySelector("#announcement-container")!.innerHTML =
      allElementsHtml;
  }


  /*
   * ==========================================================
   * RENDER OFFERS
   * ==========================================================
   *
   * Gets Offer data from SharePoint and converts each
   * SharePoint item into the Offer HTML template.
   */
  private async _renderOffersAsync(apiUrl: string): Promise<void> {

    const data: IOffer[] =
      await this._getOffersData(apiUrl);

    console.log("Offers data", data);

    let allElementsHtml: string = "";

    try {

      data.forEach((item, index) => {

        let imageUrl = '';

        /*
         * The OfferImage column contains JSON data.
         * Extract the file name and construct the
         * SharePoint attachment URL.
         */
        if (item.OfferImage) {

          const imageData = JSON.parse(item.OfferImage);

          // console.log(imageData);

          const fileName = imageData.fileName;

          // Build the image URL using the fileName
          imageUrl =
            `${this.context.pageContext.web.absoluteUrl}/Lists/Offers/Attachments/${item.Id}/${fileName}`;

          // console.log(imageUrl);
        }

        /*
         * Convert the SharePoint Created date into the
         * required display format.
         */
        let createddate = this.formatDates(item.Created);

        /*
         * Replace the placeholders in the Offer HTML template
         * with actual SharePoint data.
         */
        let singleElementHtml = UabOffers.singleElementHtml
          .replace("__KEY__OFFER__ICON__", imageUrl)
          .replace("__KEY__OFFER__TITLE__", item.Title)
          .replace("__KEY__OFFER__DESCRIPTION__", item.Description)
          .replace("__KEY__OFFER__DATE__", createddate);

        /*
         * Add the generated Offer HTML to the complete
         * HTML string.
         */
        allElementsHtml += singleElementHtml;

      })

    }
    catch (error) {
      console.error('Error rendering QuickList:', error);
    }

    /*
     * Insert all generated Offer HTML into the Offer container.
     */
    this.domElement.querySelector("#offer-container")!.innerHTML =
      allElementsHtml;
  }


  /*
   * ==========================================================
   * GET ANNOUNCEMENT DATA
   * ==========================================================
   *
   * Sends a GET request to the SharePoint REST API and
   * returns the Announcement list items.
   */
  private async _getAnnouncementsData(
    apiUrl: string
  ): Promise<IAnnouncement[]> {

    try {

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

      if (response.ok) {

        const data = await response.json();

        return data.value;

      } else {

        console.error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );

        throw new Error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );
      }

    }
    catch (error) {

      console.log("error occured", error);

      throw error;
    }
  }


  /*
   * ==========================================================
   * GET OFFER DATA
   * ==========================================================
   *
   * Sends a GET request to the SharePoint REST API and
   * returns the Offer list items.
   */
  private async _getOffersData(
    apiUrl: string
  ): Promise<IOffer[]> {

    try {

      const response: SPHttpClientResponse =
        await this.context.spHttpClient.get(
          apiUrl,
          SPHttpClient.configurations.v1
        );

      if (response.ok) {

        const data = await response.json();

        return data.value;

      } else {

        console.error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );

        throw new Error(
          `Request failed with status ${response.status}: ${response.statusText}`
        );
      }

    }
    catch (error) {

      console.log("error occured", error);

      throw error;
    }
  }


  /*
   * ==========================================================
   * FORMAT DATE
   * ==========================================================
   *
   * Converts the SharePoint Created date into:
   *
   *     Sep 21, 2026
   *
   * If the date is empty or invalid, the original value
   * is returned where appropriate.
   */
  private formatDates(
    dateValue: string
  ): string {

    if (!dateValue) {
      return '';
    }

    const date =
      new Date(dateValue);

    if (
      isNaN(
        date.getTime()
      )
    ) {

      return dateValue;
    }

    return date.toLocaleDateString(
      'en-US',
      {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
      }
    );
  }



  // ==================== INITIALIZE WEB PART ====================

  protected async onInit(): Promise<void> {

    this.loadCSS();


    await this.loadJS();


    return super.onInit();

  }


  // ==================== PROPERTY PANE ====================

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {

      pages: [

        {

          header: {
            description: 'Home Page'
          },

          groups: [

            {

              groupName: 'Configuration',

              groupFields: [

                PropertyPaneTextField(
                  'description',
                  {
                    label: 'Description'
                  }
                )

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