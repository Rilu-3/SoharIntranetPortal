export default class UpcomingEvents {
 
  public static singleElementHtml: string = `
 
    <div class="event-list-item align-items-center">
      <div class="event-date-badge">
        <span class="em">__KEY_EVENT_MONTH__</span>
        <span class="ed">__KEY_EVENT_DAY__</span>
      </div>
 
      <div class="flex-grow-1">
        <p class="event-title">__KEY_EVENT_TITLE__</p>
        <p class="event-meta">
          __KEY_EVENT_TIME__<br>
          __KEY_EVENT_LOCATION__
        </p>
      </div>
 
      __KEY_EVENT_ARROW__
    </div>
 
  `;
 
 
  public static allElementsHtml: string = `
 
  
      <div class="panel-card px-2 py-4">
 
        <div class="panel-header px-2 w-100 float-start">
          <div class="d-flex align-items-center gap-2">
                <img class="panel-title-icon" src="/sites/DevPortal/SiteAssets/resources/images/icons/section-titles/events.svg" />
                <h2 class="panel-title">UPCOMING EVENTS</h2>
              </div>
        </div>
 
        <div class="w-100 float-start pt-4 panel-card-calendar">
 
          <div class="w-100 float-start px-2">
 
            <ul class="events-tabs-list w-100 float-start">
 
              <li>
                <div
                  data-tab-event-id="events-panel-my"
                  class="etab etab-active">
                  My Events
                </div>
              </li>
 
              <li>
                <div
                  data-tab-event-id="events-panel-org"
                  class="etab">
                  Organizational Events
                </div>
              </li>
 
            </ul>
 
          </div>
 
 
          <div
            id="events-panel-my"
            class="w-100 float-start event-calendar-view"
            style="display: block;">
 
            <div class="calendar-wrap w-100 float-start px-2">
 
              <div
                id="events-calendar-my"
                class="w-100 float-start event-calendar">
              </div>
 
            </div>
 
            <div
              class="w-100 float-start d-flex flex-column custom-scroll-view event-list-scroll px-2"
              id="events-list-my">
            </div>
 
            <div
              class="w-100 float-start d-flex justify-content-start px-2">
 
              <a
                href="https://outlook.office.com/calendar/"
                target="_self"
                data-interception="off"
                rel="noopener noreferrer"
                class="link-arrow text-color-link">
 
                <span class="text-sm xxl-text-base font-bold">
                  View All Events
                </span>
 
                <img src="__KEY_ARROW_RIGHT_SHORT__" />
 
              </a>
 
            </div>
 
          </div>
 
 
          <div
            id="events-panel-org"
            class="w-100 float-start event-calendar-view">
 
            <div class="calendar-wrap w-100 float-start px-2">
 
              <div
                id="events-calendar-org"
                class="w-100 float-start event-calendar">
              </div>
 
            </div>
 
            <div
              class="w-100 float-start d-flex flex-column custom-scroll-view event-list-scroll px-2"
              id="events-list-org">
            </div>
 
            <div
              class="w-100 float-start d-flex justify-content-start px-2">
 
              <a
                href="__KEY_URL_UPCOMING__"
                target="_self"
                data-interception="off"
                rel="noopener noreferrer"
                class="link-arrow text-color-link">
 
                <span class="text-sm xxl-text-base font-bold">
                  View All Events
                </span>
 
                <img src="__KEY_ARROW_RIGHT_SHORT__" />
 
              </a>
 
            </div>
 
          </div>
 
        </div>
 
      </div>
  
 
  `;
 
 
  public static noRecord: string = `
 
    <div class="event-list-item align-items-center">
 
      <div class="flex-grow-1">
 
        <p class="event-title">
          No events for this day
        </p>
 
      </div>
 
    </div>
 
  `;
 
}