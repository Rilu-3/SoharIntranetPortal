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
export interface IWpHomePageWebPartProps {
  description: string;
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

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  // Stores Organizational Events
  // fetched from the SharePoint Upcoming Events list
  private events: IEventItem[] = [];
  // Stores My Events
  // fetched from the logged-in user's Outlook calendar
  private myEvents: IEventItem[] = [];


  // Main render method of the web part
  public async render(): Promise<void> {

    // Load required CSS and JavaScript resources
    await this.loadCSS();

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

    // Create the Upcoming Events HTML
    this.renderUpcomingEvents();

    // Initialize My Events / Organizational Events tabs
    this.initializeUpcomingEvents();

    // Initialize both calendars
    this.initializeCalendar();
  }


  // Loads CSS files and JavaScript files
  // required by the Upcoming Events UI
  private async loadCSS(): Promise<void> {

    // Get the current SharePoint site URL
    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    // Bootstrap CSS
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );


    // Custom CSS
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );


    // Font size CSS
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/font-size.css`
    );


    // Home page CSS
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );


    // jQuery UI CSS
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`
    );


    // CSS variables
    SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );


    // Load jQuery
    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );


    // Load jQuery UI
    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );
  }


  // Loads Organizational Events
  // from the SharePoint "Upcoming Events" list
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
    //
    // $select -> gets only the required columns
    //
    // $filter -> gets events only from the selected month
    //
    // $orderby -> sorts events by EventDate
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


  // Loads the logged-in user's
  // Outlook Calendar events using Microsoft Graph
 private async loadMyEvents(
  year: number,
  month: number
): Promise<void> {

  try {

    // Create Microsoft Graph client
    //
    // This client is used to communicate with
    // Microsoft Graph APIs using the current
    // SharePoint user's authentication
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
    //
    // /me means the currently logged-in user
    //
    // calendarView returns calendar events
    // between the start and end dates
    const response =
      await client
        .api('/me/calendar/calendarView')

        // Define the date range for Outlook events
        .query({
          startDateTime:
            startDate.toISOString(),

          endDateTime:
            endDate.toISOString()
        })

        // Request only the required Outlook fields
        .select(
          'id,subject,start,end,location'
        )

        // Sort Outlook events by start date/time
        .orderby(
          'start/dateTime'
        )

        // Execute the Microsoft Graph request
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

            // Create a local ID for the event
            Id:
              index + 1,

            // Outlook event subject becomes the title
            Title:
              event.subject || '',

            // Outlook start date/time
            EventDate:
              event.start?.dateTime || '',

            // Outlook start date/time
            StartTime:
              event.start?.dateTime || '',

            // Outlook end date/time
            EndTime:
              event.end?.dateTime || '',

            // Outlook location
            Location:
              event.location?.displayName || '',

            // Mark Outlook events as Active
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


  // Creates the complete Upcoming Events UI
  // using the imported template
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


    // Get the main HTML template
    let html =
      UpcomingEventsTemplate.allElementsHtml;


    // Replace the arrow placeholder
    html =
      html.replace(
        /__KEY_ARROW_RIGHT_SHORT__/g,
        arrowRightShort
      );


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


    // Insert My Events into the template
    html =
      html.replace(
        'id="events-list-my">',
        `id="events-list-my">${myEventsHtml}`
      );


    // Insert Organizational Events into the template
    html =
      html.replace(
        'id="events-list-org">',
        `id="events-list-org">${organizationalEventsHtml}`
      );

    // Display the final HTML in the web part
    this.domElement.innerHTML = html;
  }


  // Returns only future My Events
  // from the Outlook calendar
private getMyEvents(): IEventItem[] {

  // Get today's date
  const today = new Date();

  // Remove the current time
  // so comparison is based only on the date
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


// Returns only future active Organizational Events
// from the SharePoint list
private getOrganizationalEvents(): IEventItem[] {

  // Get today's date
  const today = new Date();

  // Remove the current time
  // so comparison is based only on the date
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


  // Creates the HTML for the event cards
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


  // Converts the event date into
  // month and day for display
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
      // Example: SEP
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
      // Example: 30
      day:
        date
          .getDate()
          .toString()
    };
  }


  // Converts start and end time
  // into a display format
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


  // Converts time values into 12-hour AM/PM format
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


  // Initializes the My Events and Organizational Events tabs
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


  // Initializes the jQuery UI calendars
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


  protected get dataVersion(): Version {

    return Version.parse('1.0');
  }
}