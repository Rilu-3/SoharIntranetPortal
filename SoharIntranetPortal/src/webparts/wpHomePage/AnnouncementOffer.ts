export default class AnnouncementOffer {
    public static allElementsHtml:string=`
         
 
          <div class="panel-card px-2 py-4">
 
            <div class="panel-header px-2 w-100 float-start mb-4 d-flex">
 
              <div class="d-flex align-items-center flex-wrap gap-2 gap-sm-3">
 
                <h2 data-tab-ao="announcement" class="panel-title panel-title-tab panel-title-tab-active">Announcements
 
                </h2>
 
                <h2 data-tab-ao="offers" class="panel-title panel-title-tab">Offers</h2>
 
              </div>
 
              <a href="#" id="ao-view-all" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
 
                </span><img id="ao-view-all-arrow" src="__KEY__ARROW__RIGHT__ICON__" /></a>
 
            </div>
 
             <div id="announcement" class="w-100 float-start ao-tab-view" style="display: block;">
 
              <div
                id="announcement-container"  class="w-100 d-flex flex-column float-start px-2 overflow-auto panel-card-announcement  custom-scroll-view">
           
              </div>
 
            </div>
 
            <div id="offers" class="w-100 float-start ao-tab-view">
              <div
                id="offer-container"
                class="w-100 d-flex flex-column float-start px-2 overflow-auto panel-card-announcement custom-scroll-view">
             
               
             
              </div>
             
 
            </div>
 
          </div>
 
       
    `;
 
    public static singleElementHtml:string=`  <div class="announcement-item">
                  <div class="announcement-icon bg-1">
                    <img src="__KEY__ANNOUNCEMENTOFFER__ICON__" />
                  </div>
                  <div class="flex-grow-1">
                    <p class="announcement-title mb-1">__KEY__ANNOUNCEMENTOFFER__TITLE__</p>
                    <p class="announcement-desc">__KEY__ANNOUNCEMENTOFFER__DESCRIPTION__</p>
                  </div>
                  <span class="announcement-date">__KEY__ANNOUNCEMENTOFFER__DATE__</span>
                </div>`;
 
}