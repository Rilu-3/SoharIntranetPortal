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
  SPHttpClientResponse
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

  Category: string;

  Status: string;

}


export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  private events: IEventItem[] = [];


  public async render(): Promise<void> {

    this.loadCSS();

    await this.loadEvents();

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

    // jQuery
    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );

    // jQuery UI
    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );
  }
  private async loadEvents(): Promise<void> {

    const siteUrl =
      this.context.pageContext.web.absoluteUrl;


    const url =
      `${siteUrl}/_api/web/lists/getbytitle('Upcoming Events')/items` +
      `?$select=Id,Title,EventDate,StartTime,EndTime,Location,Category,Status` +
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
        'Upcoming Events:',
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
      rightArrow
    );

  const organizationalEventsHtml =
    this.renderEventElements(
      organizationalEvents,
      rightArrow
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


  // Put your HTML directly into the webpart
  this.domElement.innerHTML = html;
}

  private getMyEvents(): IEventItem[] {

    return this.events.filter(
      (event: IEventItem) => {

        return (
          event.Category === 'My Event' &&
          event.Status === 'Active'
        );

      }
    );
  }


  private getOrganizationalEvents(): IEventItem[] {

    return this.events.filter(
      (event: IEventItem) => {

        return (
          event.Category === 'Organizational Event' &&
          event.Status === 'Active'
        );

      }
    );
  }


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


  private formatDate(
    eventDate: string
  ): {
    month: string;
    day: string;
  } {

    const date =
      new Date(eventDate);


    if (isNaN(date.getTime())) {

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


    if (!start && !end) {

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


    /*
     * SharePoint sometimes returns
     * a complete date/time value.
     */
    if (timeValue.indexOf('T') !== -1) {

      const date =
        new Date(timeValue);


      if (!isNaN(date.getTime())) {

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


            /*
             * Hide all panels.
             */
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

              const selectedPanel =
                this.domElement.querySelector(
                  `#${targetId}`
                );


              if (selectedPanel) {

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


  private initializeCalendar(): void {

    const $ =
      (window as any).jQuery;


    /*
     * jQuery UI may not have loaded yet.
     * Do not stop the whole webpart.
     */
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


    $('#events-calendar-my').datepicker({

      dateFormat: 'dd M yy'

    });


    $('#events-calendar-org').datepicker({

      dateFormat: 'dd M yy'

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