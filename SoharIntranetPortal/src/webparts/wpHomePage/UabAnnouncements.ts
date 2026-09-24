export default class UabAnnouncements{
    public static allElementsHtml:string=`
    <div
              id="announcement-container"  class="w-100 d-flex flex-column float-start px-2 overflow-auto panel-card-announcement  custom-scroll-view">
          
              </div>`;
    public static singleElementHtml:string=`                <a class="announcement-item">
                  <div class="announcement-icon bg-1">
                    <img src="__KEY__ANNOUNCEMENT__ICON__" />
                  </div>
                  <div class="flex-grow-1">
                    <p class="announcement-title mb-1">__KEY__ANNOUNCEMENT__TITLE__</p>
                    <p class="announcement-desc">__KEY__ANNOUNCEMENT__DESCRIPTION__</p>
                  </div>
                  <span class="announcement-date">__KEY__ANNOUNCEMENT__DATE__</span>
                </a>`;
}