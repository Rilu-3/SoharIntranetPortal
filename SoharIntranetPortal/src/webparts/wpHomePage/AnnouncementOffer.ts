export default class AnnouncementOffer {
    public static allElementsHtml:string=`
         
 
          <div class="panel-card px-2 py-4">
 
            <div class="panel-header px-2 w-100 float-start mb-4 d-flex">
 
              <div class="d-flex align-items-center flex-wrap gap-2 gap-sm-3">
 
               <div data-tab-ao="announcement" class="d-flex align-items-center gap-2 panel-title-tab panel-title-tab-active">
                    <img class="panel-title-icon" src="/sites/DevPortal/SiteAssets/resources/images/icons/section-titles/announcements.svg" />
                    <h2 class="panel-title">Announcements</h2>
                  </div>
 
                <div data-tab-ao="offers" class="d-flex align-items-center gap-2 panel-title-tab">
                    <img class="panel-title-icon" src="/sites/DevPortal/SiteAssets/resources/images/icons/section-titles/offers.png" />
                    <h2 class="panel-title">Offers</h2>
                  </div>

                <div data-tab-ao="new-joiners" class="d-flex align-items-center gap-2 panel-title-tab">
                  <img class="panel-title-icon" src="/sites/DevPortal/SiteAssets/resources/images/icons/section-titles/new-employee.png" />
                  <h2 class="panel-title">New Joiners</h2>
                </div>
 
              </div>
 
              <a href="#"  target="_self" data-interception="off" id="ao-view-all" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
 
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

            <div id="new-joiners" class="w-100 float-start ao-tab-view">
            <div id="new-joiners-container"
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
                    <p class="announcement-title mb-1">        <a href="__KEY__ANNOUNCEMENTOFFER__LINK__" target="_self" data-interception="off"
           class="text-decoration-none announcement-link" style="color:inherit;">__KEY__ANNOUNCEMENTOFFER__TITLE__</a></p>
                    <p class="announcement-desc line-clamp-2">        <a href="__KEY__ANNOUNCEMENTOFFER__LINK__" target="_self" data-interception="off"
           class="text-decoration-none announcement-link" style="color:inherit;">__KEY__ANNOUNCEMENTOFFER__DESCRIPTION__</a>
</p>
                  </div>
                  <span class="announcement-date">__KEY__ANNOUNCEMENTOFFER__DATE__</span>
                </div>`;
 
    public static noElementHtml: string = `
    <div class="news-item">
      <div class="flex-grow-1">
        <p class="news-desc">No records found.</p>
      </div>
    </div>`
 public static singleNewJoinerHtml: string = `
<div class="new-joiners-item">
<img src="__KEY__NEWJOINER__IMAGE__" class="new-joiners-avatar" alt="__KEY__NEWJOINER__NAME__">
<div class="d-flex flex-column flex-grow-1 overflow-hidden">
<p class="new-joiners-name">__KEY__NEWJOINER__NAME__</p>
<p class="new-joiners-desc">__KEY__NEWJOINER__DESCRIPTION__</p>
</div>
<div class="new-joiners-date d-flex">
<span>__KEY__NEWJOINER__MONTH__</span>
<span>__KEY__NEWJOINER__DAY__</span>
</div>
</div>`;
 
public static noNewJoinerHtml: string = `
<div class="news-item">
<div class="flex-grow-1">
<p class="news-desc">No new joiners found.</p>
</div>
</div>`;
 
}
 
 