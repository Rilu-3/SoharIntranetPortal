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
  };
  end?: {
    dateTime?: string;
  };
  location?: {
    displayName?: string;
  };
  onlineMeeting?: {
    joinUrl?: string;
  };
}

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {

  private events: IEventItem[] = [];

  private myEvents: IEventItem[] = [];


  public async render(): Promise<void> {

    const workbenchContent =
      document.getElementById(
        'workbenchPageContent'
      );

    if (workbenchContent) {

      workbenchContent.style.maxWidth =
        'none';
    }

    await this.loadCSS();
await this.loadHomeJS();
    const today =
      new Date();

    await Promise.all([

      this.loadEvents(
        today.getFullYear(),
        today.getMonth()
      ),

      this.loadMyEvents(
        today.getFullYear(),
        today.getMonth()
      )

    ]);

    this.renderUpcomingEvents();

    this.initializeUpcomingEvents();

    this.initializeCalendar();

  }


  private async loadCSS(): Promise<void> {

    const baseUrl =
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
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );
  }
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

  private async loadEvents(
    year: number,
    month: number
  ): Promise<void> {

    const siteUrl =
      this.context.pageContext.web.absoluteUrl;


    const startDate =
      new Date(
        year,
        month,
        1
      );


    const endDate =
      new Date(
        year,
        month + 1,
        1
      );


    const startDateString =
      startDate.toISOString();


    const endDateString =
      endDate.toISOString();


    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Upcoming Events')/items` +
      `?$select=Id,Title,EventDate,StartTime,EndTime,Location,Status,Link` +
      `&$filter=EventDate ge datetime'${startDateString}' and EventDate lt datetime'${endDateString}'` +
      `&$orderby=EventDate asc`;


    try {

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


      if (!response.ok) {

        console.error(
          'Upcoming Events list error:',
          response.status,
          response.statusText
        );

        this.events = [];

        return;
      }


      const data =
        await response.json();


      this.events =
        data.value || [];


      console.log(
        'Organizational Events:',
        this.events
      );

    } catch (error) {

      console.error(
        'Error loading Upcoming Events:',
        error
      );

      this.events = [];
    }
  }


  private async loadMyEvents(
    year: number,
    month: number
  ): Promise<void> {

    try {

      const client:
        MSGraphClientV3 =
        await this.context.msGraphClientFactory.getClient(
          '3'
        );


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
          .api(
            '/me/calendar/calendarView'
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


      console.log(
        'Outlook Calendar Events:',
        response.value
      );


      this.myEvents =
        (response.value || []).map(
          (
            event: IOutlookEvent,
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


  private renderUpcomingEvents(): void {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    const rightArrow =
      `${baseUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


    const arrowRightShort =
      `${baseUrl}/SiteAssets/resources/images/icons/arrow-right-short.svg`;


    const myEvents =
      this.getMyEvents();


    const organizationalEvents =
      this.getOrganizationalEvents();


    let html =
      UpcomingEventsTemplate.allElementsHtml;


    html =
      html.replace(
        /__KEY_ARROW_RIGHT_SHORT__/g,
        arrowRightShort
      );


    const myEventsHtml =
      this.renderEventElements(
        myEvents,
        rightArrow,
        true
      );


    const organizationalEventsHtml =
      this.renderEventElements(
        organizationalEvents,
        rightArrow,
        false
      );


    html =
      html.replace(
        'id="events-list-my">',
        `id="events-list-my">${myEventsHtml}`
      );


    html =
      html.replace(
        'id="events-list-org">',
        `id="events-list-org">${organizationalEventsHtml}`
      );


    this.domElement.innerHTML =
      html;
  }


  private getMyEvents(): IEventItem[] {

    const today =
      new Date();


    today.setHours(
      0,
      0,
      0,
      0
    );


    return this.myEvents.filter(
      (event: IEventItem) => {

        const eventDate =
          new Date(
            event.EventDate
          );


        eventDate.setHours(
          0,
          0,
          0,
          0
        );


        return eventDate >= today;
      }
    );
  }


  private getOrganizationalEvents(): IEventItem[] {

    const today =
      new Date();


    today.setHours(
      0,
      0,
      0,
      0
    );


    return this.events.filter(
      (event: IEventItem) => {

        if (
          event.Status !==
          'Active'
        ) {

          return false;
        }


        const eventDate =
          new Date(
            event.EventDate
          );


        eventDate.setHours(
          0,
          0,
          0,
          0
        );


        return eventDate >= today;
      }
    );
  }


  private renderEventElements(
    events: IEventItem[],
    rightArrow: string,
    isMyEvent: boolean
  ): string {

    if (!events.length) {

      return UpcomingEventsTemplate.noRecord;
    }


    const displayEvents = events;
    return displayEvents
      .map(
        (event: IEventItem) => {

          const date =
            this.formatDate(
              event.EventDate
            );


          const time =
            this.formatTime(
              event.StartTime,
              event.EndTime
            );


          let html =
            UpcomingEventsTemplate.singleElementHtml;


          html =
            html.replace(
              '__KEY_EVENT_MONTH__',
              escape(date.month)
            );


          html =
            html.replace(
              '__KEY_EVENT_DAY__',
              escape(date.day)
            );


          html =
            html.replace(
              '__KEY_EVENT_TITLE__',
              escape(event.Title || '')
            );


          html =
            html.replace(
              '__KEY_EVENT_TIME__',
              escape(time)
            );


          html =
            html.replace(
              '__KEY_EVENT_LOCATION__',
              escape(event.Location || '')
            );

let arrowHtml = '';

if (isMyEvent) {

  if (event.TeamsUrl) {

    arrowHtml =
      `<a href="${escape(event.TeamsUrl)}" target="_blank" data-interception="off" rel="noopener noreferrer">
        <img src="${rightArrow}" />
      </a>`;
  }

} else if (event.Link) {

  arrowHtml =
    `<a href="${escape(event.Link)}" target="_blank" data-interception="off" rel="noopener noreferrer">
      <img src="${rightArrow}" />
    </a>`;
}
          html =
            html.replace(
              '__KEY_EVENT_ARROW__',
              arrowHtml
            );


          return html;
        }
      )
      .join('');
  }


  private formatDate(
    eventDate: string
  ): {
    month: string;
    day: string;
  } {

    const date =
      new Date(
        eventDate
      );


    if (
      isNaN(
        date.getTime()
      )
    ) {

      return {
        month: '',
        day: ''
      };
    }


    return {

      month:
        date
          .toLocaleString(
            'en-US',
            {
              month: 'short'
            }
          )
          .toUpperCase(),

      day:
        date
          .getDate()
          .toString()
    };
  }


  private formatTime(
    startTime: string,
    endTime: string
  ): string {

    const start =
      this.parseSharePointTime(
        startTime
      );


    const end =
      this.parseSharePointTime(
        endTime
      );


    if (
      !start &&
      !end
    ) {

      return '';
    }


    if (!end) {

      return start;
    }


    return `${start} – ${end}`;
  }


  private parseSharePointTime(
    timeValue: string
  ): string {

    if (!timeValue) {

      return '';
    }


    if (
      timeValue.indexOf('T') !== -1
    ) {

      const date =
        new Date(
          timeValue
        );


      if (
        !isNaN(
          date.getTime()
        )
      ) {

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


    const match =
      timeValue.match(
        /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
      );


    if (match) {

      const hours =
        parseInt(
          match[1],
          10
        );


      const minutes =
        parseInt(
          match[2],
          10
        );


      if (
        hours >= 0 &&
        hours <= 23 &&
        minutes >= 0 &&
        minutes <= 59
      ) {

        const period =
          hours >= 12
            ? 'PM'
            : 'AM';


        const displayHour =
          hours % 12 === 0
            ? 12
            : hours % 12;


        return (
          `${('0' + displayHour).slice(-2)}:` +
          `${('0' + minutes).slice(-2)} ` +
          `${period}`
        );
      }
    }


    return timeValue;
  }


  private initializeUpcomingEvents(): void {

    const tabs =
      this.domElement.querySelectorAll(
        '.events-tabs-list .etab'
      );


    const panels =
      this.domElement.querySelectorAll(
        '.event-calendar-view'
      );
    tabs.forEach(
      (tab: Element) => {

        tab.addEventListener(
          'click',
          () => {

            const targetId =
              tab.getAttribute(
                'data-tab-event-id'
              );


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
                ).style.display =
                  'none';
              }
            );
            tab.classList.add(
              'etab-active'
            );
            if (targetId) {
              const selectedPanel =
                this.domElement.querySelector(
                  `#${targetId}`
                );
              if (selectedPanel) {
                (
                  selectedPanel as HTMLElement
                ).style.display =
                  'block';
              }
            }
          }
        );
      }
    );
  }
private initializeCalendar(): void {
  const $ =
    (window as any).jQuery;
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
  const self =
    this;
  $('#events-calendar-my').datepicker({
    dateFormat:
      'dd M yy',
    onSelect:
      function (
        dateText: string
      ): void {
        const selectedDate =
          new Date(
            dateText
          );
        selectedDate.setHours(
          0,
          0,
          0,
          0
        );
        const selectedEvents =
          self.getMyEvents().filter(
            (event: IEventItem) => {
              const eventDate =
                new Date(
                  event.EventDate
                );
              eventDate.setHours(
                0,
                0,
                0,
                0
              );
              return (
                eventDate.getTime() ===
                selectedDate.getTime()
              );
            }
          );
        const rightArrow =
          `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;
        const myEventsHtml =
          self.renderEventElements(
            selectedEvents,
            rightArrow,
            true
          );
        const myEventsList =
          self.domElement.querySelector(
            '#events-list-my'
          );
        if (myEventsList) {
          myEventsList.innerHTML =
            myEventsHtml;
        }
      },
    onChangeMonthYear:
      async function (
        year: number,
        month: number
      ): Promise<void> {

        await self.loadMyEvents(
          year,
          month - 1
        );
        const rightArrow =
          `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;
        const myEventsHtml =
          self.renderEventElements(
            self.getMyEvents(),
            rightArrow,
            true
          );
        const myEventsList =
          self.domElement.querySelector(
            '#events-list-my'
          );
        if (myEventsList) {

          myEventsList.innerHTML =
            myEventsHtml;
        }
      }
  });
  $('#events-calendar-org').datepicker({
    dateFormat:
      'dd M yy',
    onSelect:
      function (
        dateText: string
      ): void {

        const selectedDate =
          new Date(
            dateText
          );

        selectedDate.setHours(
          0,
          0,
          0,
          0
        );
        const selectedEvents =
          self.getOrganizationalEvents().filter(
            (event: IEventItem) => {

              const eventDate =
                new Date(
                  event.EventDate
                );

              eventDate.setHours(
                0,
                0,
                0,
                0
              );


              return (
                eventDate.getTime() ===
                selectedDate.getTime()
              );
            }
          );
        const rightArrow =
          `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;
        const organizationalEventsHtml =
          self.renderEventElements(
            selectedEvents,
            rightArrow,
            false
          );


        const organizationalEventsList =
          self.domElement.querySelector(
            '#events-list-org'
          );


        if (organizationalEventsList) {

          organizationalEventsList.innerHTML =
            organizationalEventsHtml;
        }
      },


    onChangeMonthYear:
      async function (
        year: number,
        month: number
      ): Promise<void> {

        await self.loadEvents(
          year,
          month - 1
        );


        const rightArrow =
          `${self.context.pageContext.web.absoluteUrl}/SiteAssets/resources/images/icons/right-arrow.png`;


        const organizationalEventsHtml =
          self.renderEventElements(
            self.getOrganizationalEvents(),
            rightArrow,
            false
          );


        const organizationalEventsList =
          self.domElement.querySelector(
            '#events-list-org'
          );


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
            description:
              'Home Page'
          },

          groups: [

            {

              groupName:
                'Configuration',

              groupFields: [

                PropertyPaneTextField(
                  'description',
                  {
                    label:
                      'Description'
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

    return Version.parse(
      '1.0'
    );
  }
}