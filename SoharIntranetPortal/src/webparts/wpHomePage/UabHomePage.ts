export default class UabHomePage{
    public static allElementsHtml:string=`
    <!DOCTYPE html>
<html lang="en">



<body>
  <!-- Widget start here -->
  <nav class="navbar navbar-expand-lg navbar-light fixed-top sa-header py-0 flex-column">

    <div class="w-100 float-start sa-header-top py-2">
      <div class="container container-sa px-3 px-lg-4 d-flex align-items-center gap-3">
        <a class="navbar-brand" href="#">
          <img class="logo-desktop" src="./resources/images/logo.png" />
          <img class="logo-mob" src="./resources/images/logo-mob.png" />
        </a>
        <div class="header-search position-relative mx-lg-auto d-none d-lg-block">
          <input type="text" class="form-control" placeholder="Search people, documents, pages...">
          <button class="search-icon-btn" type="button" aria-label="Search">
            <i class="bi bi-search"></i>
          </button>
        </div>
        <div
          class="ms-lg-0 ms-auto nav-user-info d-flex align-items-center justify-content-between justify-content-lg-start gap-3 gap-lg-4">
          <div class="btn btn-secondary d-flex align-items-center p-0 bg-transparent border-0 gap-2">
            <div class="nav-avatar-wrapper d-flex align-items-center gap-2">
              <img class="nav-avatar" src="./resources/images/avatar.png" />
              <div class="max-w-200 d-flex flex-column text-start d-none d-sm-flex">
                <p class="text-sm xxl-text-base font-bold text-color-title lh-base text-truncate">
                  Lubna Helal
                </p>
                <p class="text-xs font-normal text-color-desc lh-sm text-truncate">
                  Associate Analyst
                </p>
              </div>
            </div>
          </div>
        </div>
        <button class="navbar-toggler" type="button" data-bs-toggle="collapse" data-bs-target="#navbarSupportedContent"
          aria-controls="navbarSupportedContent" aria-expanded="false" aria-label="Toggle navigation">
          <span class="navbar-toggler-icon"></span>
        </button>
      </div>
    </div>
    <div class="w-100 float-start sa-header-bottom">
      <div class="collapse navbar-collapse gap-lg-2" id="navbarSupportedContent">
        <div class="container container-sa px-3 px-lg-4 d-flex gap-3 py-lg-2">
          <ul class="navbar-nav my-2 my-lg-0 navbar-nav-scroll gap-lg-3" style="--bs-scroll-height: 400px;">
            <li class="nav-item d-lg-flex">
              <a class="nav-link active-nav-link d-flex align-items-center gap-2" aria-current="page" href="#">
                <img class="nav-menu-icon" src="./resources/images/icons/home.png" alt="" />
                <span>Home</span></a>
            </li>
            <li class="nav-item d-lg-flex dropdown">
              <a class="nav-link dropdown-toggle d-flex align-items-center gap-2" href="#" id="navbarScrollingDropdown"
                role="button" data-bs-toggle="dropdown" aria-expanded="false">
                <img class="nav-menu-icon" src="./resources/images/icons/department.png" alt="" />
                <span>Departments</span>
                <img class="dropdown-arrow-nav" src="./resources/images/icons/dropdown-arrow.png" />
              </a>
              <ul class="dropdown-menu" aria-labelledby="navbarScrollingDropdown">
                <li><a class="dropdown-item text-sm" href="#">Department 1</a></li>
                <li><a class="dropdown-item text-sm" href="#">Department 2</a></li>
                <li><a class="dropdown-item text-sm" href="#">Department 3</a></li>
              </ul>
            </li>
          </ul>
        </div>
      </div>
    </div>
  </nav>
  <!-- Widget end here -->

  <div class="main-wrapper w-100 float-start min-h-screen-wrapper">
    <div class="container container-sa px-3 px-lg-4 py-3 mx-auto">
      <div class="row gy-3 gx-3 clearfix">
        <div class="col-12">
          <div class="w-100 float-start swiper banner-swiper">
            <div class="swiper-wrapper">
              <div class="swiper-slide">
                <div class="w-100 float-start banner-slide-grid">
                  <img src="./resources/images/banner/banner-1.png" alt="" /><!--1500 × 450 px → 3.33 : 1-->
                  <div class="banner-slide-info">
                    <div class="w-100 text-white d-flex flex-column gap-2 gap-xl-3 px-3 px-lg-4">
                      <h4 class="font-bold fhd-text-size-48-combo text-capitalize lh-3 mb-0 line-clamp-3">Shaping
                        Aluminium.<br />Building Tomorrow.</h4>
                      <p class="font-light text-base lg-text-lg lh-sm line-clamp-3">Our people, our performance and our
                        purpose
                        drive sustainable value for Oman and beyond.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
              <div class="swiper-slide">
                <div class="w-100 float-start banner-slide-grid">
                  <img src="./resources/images/banner/banner-1.png" alt="" /><!--1500 × 450 px → 3.33 : 1-->
                  <div class="banner-slide-info">
                    <div class="w-100 text-white d-flex flex-column gap-2 gap-xl-3 px-3 px-lg-4">
                      <h4 class="font-bold fhd-text-size-40-combo text-capitalize lh-3 mb-0 line-clamp-3">Shaping
                        Aluminium.<br />Building Tomorrow.</h4>
                      <p class="font-light text-base lh-sm line-clamp-3">Our people, our performance and our purpose
                        drive sustainable value for Oman and beyond.
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </div>
            <div class="swiper-pagination"></div>
            <div class="swiper-button-next banner-slide-nav-btn banner-slide-next me-3 me-lg-4"></div>
            <div class="swiper-button-prev banner-slide-nav-btn banner-slide-prev ms-3 ms-lg-4"></div>
          </div>
        </div>
        <div class="col-12 col-lg-6">
          <div class="panel-card px-2 py-4">
            <div class="panel-header px-2 w-100 float-start mb-4 d-flex">
              <div class="d-flex align-items-center flex-wrap gap-2 gap-sm-3">
                <h2 data-tab-ao="announcement" class="panel-title panel-title-tab panel-title-tab-active">Announcements
                </h2>
                <h2 data-tab-ao="offers" class="panel-title panel-title-tab">Offers</h2>
              </div>
              <a href="#" id="ao-view-all" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
                </span><img src="__KEY__ARROW__RIGHT__ICON__" /></a>
            </div>
            <div id="announcement" class="w-100 float-start ao-tab-view" style="display: block;">
            
            </div>
            <div id="offers" class="w-100 float-start ao-tab-view">
             
            </div>
          </div>
        </div>

        <div class="col-12 col-lg-6">
          <div class="panel-card px-2 py-4">

            <div class="panel-header px-2 w-100 float-start pb-3 d-flex align-items-center">
              <div class="d-flex align-items-center flex-wrap gap-2 gap-sm-3">
                <h2 data-filter-ql="quick-links" class="panel-title panel-title-filter panel-title-filter-active">Quick
                  Links
                </h2>
                <h2 data-filter-ql="favourites" class="panel-title panel-title-filter">Favourites</h2>
              </div>
              <button type="button" class="btn-appearance-none btn-add-favourite" title="Add to Favourites"
                data-bs-toggle="modal" data-bs-target="#addFavouriteModal">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" fill="#000" viewBox="0 0 16 16">
                  <path fill-rule="evenodd"
                    d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2" />
                </svg>
              </button>
            </div>

            <div class="w-100 float-start p-2 overflow-auto panel-card-quick-links quick-links-grid custom-scroll-view">
              <a href="#" class="quick-link-box" data-tab-cat="ql-favourite">
                <img src="./resources/images/icons/quick-links/quick-link-1.png" alt=""><span>Payslip</span>
              </a>
              <a href="#" class="quick-link-box">
                <img src="./resources/images/icons/quick-links/quick-link-2.png" alt=""><span>Leave Application</span>
              </a>
              <a href="#" class="quick-link-box" data-tab-cat="ql-favourite">
                <img src="./resources/images/icons/quick-links/quick-link-3.png" alt=""><span>Travel Request</span>
              </a>
              <a href="#" class="quick-link-box">
                <img src="./resources/images/icons/quick-links/quick-link-4.png" alt=""><span>Expense Claim</span>
              </a>
              <a href="#" class="quick-link-box" data-tab-cat="ql-favourite">
                <img src="./resources/images/icons/quick-links/quick-link-5.png" alt=""><span>Training Portal</span>
              </a>
              <a href="#" class="quick-link-box">
                <img src="./resources/images/icons/quick-links/quick-link-6.png" alt=""><span>IT Help Desk</span>
              </a>
              <a href="#" class="quick-link-box" data-tab-cat="ql-favourite">
                <img src="./resources/images/icons/quick-links/quick-link-7.png" alt=""><span>Safety Portal</span>
              </a>
              <a href="#" class="quick-link-box">
                <img src="./resources/images/icons/quick-links/quick-link-8.png" alt=""><span>Idea Box</span>
              </a>
            </div>
          </div>
        </div>
        <div class="col-12 col-lg-6">
          <div class="panel-card px-2 py-4 d-flex flex-column">
            <div class="panel-header px-2 w-100 float-start mb-4">
              <h2 class="panel-title">News Centre</h2>
              <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
                  News</span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
            </div>
            <div id="news-tabs" class="w-100 float-start d-flex flex-column flex-grow-1 overflow-hidden">
              <ul class="news-tabs-list px-2">
                <li>
                  <div data-tab-news-id="news-panel-all" class="tab-title-pill tab-title-pill-active">All</div>
                </li>
                <li>
                  <div data-tab-news-id="news-panel-announcements" class="tab-title-pill">Announcements</div>
                </li>
                <li>
                  <div data-tab-news-id="news-panel-events" class="tab-title-pill">Events</div>
                </li>
                <li>
                  <div data-tab-news-id="news-panel-news" class="tab-title-pill">News</div>
                </li>
                <li>
                  <div data-tab-news-id="news-panel-circulars" class="tab-title-pill">Circulars</div>
                </li>
              </ul>

              <div id="news-panel-all" class="w-100 float-start news-panel-tab-view flex-grow-1"
                style="display: block;">
                <div class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
                  <div class="news-item flex-column flex-sm-row">
                    <div class="d-flex flex-grow-1 align-items-center gap-3">
                      <div class="news-icon">
                        <img src="./resources/images/icons/quick-links/quick-link-1.png" />
                      </div>
                      <div class="flex-grow-1">
                        <p class="news-title">News Title Goes Here</p>
                        <p class="news-desc">Sub-title or short description of the news item goes here.</p>
                        <p class="news-meta mt-1"><span class="tag">News</span> &nbsp;•&nbsp; May 20, 2024</p>
                      </div>
                    </div>
                    <a href="#"
                      class="link-arrow text-color-link align-self-end align-self-sm-center flex-shrink-0"><span
                        class="text-sm xxl-text-base font-bold">Read More</span><img
                        src="./resources/images/icons/arrow-right-short.svg" /></a>
                  </div>
                  <div class="news-item flex-column flex-sm-row">
                    <div class="d-flex flex-grow-1 align-items-center gap-3">
                      <div class="news-icon">
                        <img src="./resources/images/icons/quick-links/quick-link-1.png" />
                      </div>
                      <div class="flex-grow-1">
                        <p class="news-title">Important Update</p>
                        <p class="news-desc">Sub-title or short description of the news item goes here.</p>
                        <p class="news-meta mt-1"><span class="tag">Circulars</span> &nbsp;•&nbsp; May 18, 2024</p>
                      </div>
                    </div>
                    <a href="#"
                      class="link-arrow text-color-link align-self-end align-self-sm-center flex-shrink-0"><span
                        class="text-sm xxl-text-base font-bold">Read More</span><img
                        src="./resources/images/icons/arrow-right-short.svg" /></a>
                  </div>
                  <div class="news-item flex-column flex-sm-row">
                    <div class="d-flex flex-grow-1 align-items-center gap-3">
                      <div class="news-icon">
                        <img src="./resources/images/icons/quick-links/quick-link-1.png" />
                      </div>
                      <div class="flex-grow-1">
                        <p class="news-title">Company Newsletter – May 2024</p>
                        <p class="news-desc">Highlights from our latest achievements, initiatives and upcoming plans.
                        </p>
                        <p class="news-meta mt-1"><span class="tag">Announcements</span> &nbsp;•&nbsp; May 17, 2024</p>
                      </div>
                    </div>
                    <a href="#"
                      class="link-arrow text-color-link align-self-end align-self-sm-center flex-shrink-0"><span
                        class="text-sm xxl-text-base font-bold">Read More</span><img
                        src="./resources/images/icons/arrow-right-short.svg" /></a>
                  </div>
                  <div class="news-item flex-column flex-sm-row">
                    <div class="d-flex flex-grow-1 align-items-center gap-3">
                      <div class="news-icon">
                        <img src="./resources/images/icons/quick-links/quick-link-1.png" />
                      </div>
                      <div class="flex-grow-1">
                        <p class="news-title">News Title Goes Here</p>
                        <p class="news-desc">Sub-title or short description of the news item goes here.</p>
                        <p class="news-meta mt-1"><span class="tag">News</span> &nbsp;•&nbsp; May 20, 2024</p>
                      </div>
                    </div>
                    <a href="#"
                      class="link-arrow text-color-link align-self-end align-self-sm-center flex-shrink-0"><span
                        class="text-sm xxl-text-base font-bold">Read More</span><img
                        src="./resources/images/icons/arrow-right-short.svg" /></a>
                  </div>
                </div>
              </div>

              <div id="news-panel-announcements" class="w-100 float-start news-panel-tab-view flex-grow-1">
                <div class="w-100 d-flex flex-column float-start p-2 overflow-auto panel-card-news custom-scroll-view">
                  <div class="news-item">
                    <div class="news-icon">
                      <img src="./resources/images/icons/quick-links/quick-link-1.png" />
                    </div>
                    <div class="flex-grow-1">
                      <p class="news-title">News Title Goes Here</p>
                      <p class="news-desc">Sub-title or short description of the news item goes here.</p>
                      <p class="news-meta mt-1"><span class="tag">News</span> &nbsp;•&nbsp; May 20, 2024</p>
                    </div>
                    <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">Read
                        More</span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
                  </div>
                </div>
              </div>

            </div>
          </div>
        </div>
        <div class="col-12 col-lg-6">
          <div class="panel-card px-2 py-4">
            <div class="panel-header px-2 w-100 float-start">
              <h2 class="panel-title">Upcoming Events</h2>
            </div>
            <div class="w-100 float-start px-2 pt-4 panel-card-calendar">
              <ul class="events-tabs-list w-100 float-start">
                <li>
                  <div data-tab-event-id="events-panel-my" class="etab etab-active">My Events</div>
                </li>
                <li>
                  <div data-tab-event-id="events-panel-org" class="etab">Organizational Events</div>
                </li>
              </ul>
              <div id="events-panel-my" class="w-100 float-start event-calendar-view" style="display: block;">
                <div class="calendar-wrap w-100 float-start">
                  <div id="events-calendar-my" class="w-100 float-start event-calendar"></div>
                </div>
                <div class="w-100 float-start d-flex flex-column">
                  <div class="event-list-item align-items-center">
                    <div class="event-date-badge"><span class="em">MAY</span><span class="ed">21</span></div>
                    <div class="flex-grow-1">
                      <p class="event-title">Leadership Training Program</p>
                      <p class="event-meta">09:00 AM – 01:00 PM<br>Training Room 1</p>
                    </div>
                    <img src="./resources/images/icons/right-arrow.png" />
                  </div>
                  <div class="event-list-item align-items-center">
                    <div class="event-date-badge"><span class="em">MAY</span><span class="ed">21</span></div>
                    <div class="flex-grow-1">
                      <p class="event-title">Leadership Training Program</p>
                      <p class="event-meta">09:00 AM – 01:00 PM<br>Training Room 1</p>
                    </div>
                    <img src="./resources/images/icons/right-arrow.png" />
                  </div>
                </div>
                <div class="w-100 float-start d-flex justify-content-start">
                  <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
                      Events</span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
                </div>
              </div>
              <div id="events-panel-org" class="w-100 float-start event-calendar-view">
                <div class="calendar-wrap w-100 float-start">
                  <div id="events-calendar-org" class="w-100 float-start event-calendar"></div>
                </div>
                <div class="w-100 float-start d-flex flex-column">
                  <div class="event-list-item align-items-center">
                    <div class="event-date-badge"><span class="em">MAY</span><span class="ed">21</span></div>
                    <div class="flex-grow-1">
                      <p class="event-title">Leadership Training Program</p>
                      <p class="event-meta">09:00 AM – 01:00 PM<br>Training Room 1</p>
                    </div>
                    <img src="./resources/images/icons/right-arrow.png" />
                  </div>
                  <div class="event-list-item align-items-center">
                    <div class="event-date-badge"><span class="em">MAY</span><span class="ed">21</span></div>
                    <div class="flex-grow-1">
                      <p class="event-title">Leadership Training Program</p>
                      <p class="event-meta">09:00 AM – 01:00 PM<br>Training Room 1</p>
                    </div>
                    <img src="./resources/images/icons/right-arrow.png" />
                  </div>
                </div>
                <div class="w-100 float-start d-flex justify-content-start">
                  <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
                      Events</span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="col-12 col-lg-8">
          <div class="panel-card px-2 py-4">
            <div class="panel-header px-2 w-100 float-start">
              <h2 class="panel-title">Social Media</h2>
            </div>
            <div class="w-100 float-start px-2 pt-4 panel-card-calendar">
              <ul class="social-tabs-list w-100 float-start gap-3 gap-lg-4">
                <li>
                  <div data-tab-social-id="social-panel-instagram" class="sm-tab sm-tab-active gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                      class="bi bi-instagram" viewBox="0 0 16 16">
                      <path
                        d="M8 0C5.829 0 5.556.01 4.703.048 3.85.088 3.269.222 2.76.42a3.9 3.9 0 0 0-1.417.923A3.9 3.9 0 0 0 .42 2.76C.222 3.268.087 3.85.048 4.7.01 5.555 0 5.827 0 8.001c0 2.172.01 2.444.048 3.297.04.852.174 1.433.372 1.942.205.526.478.972.923 1.417.444.445.89.719 1.416.923.51.198 1.09.333 1.942.372C5.555 15.99 5.827 16 8 16s2.444-.01 3.298-.048c.851-.04 1.434-.174 1.943-.372a3.9 3.9 0 0 0 1.416-.923c.445-.445.718-.891.923-1.417.197-.509.332-1.09.372-1.942C15.99 10.445 16 10.173 16 8s-.01-2.445-.048-3.299c-.04-.851-.175-1.433-.372-1.941a3.9 3.9 0 0 0-.923-1.417A3.9 3.9 0 0 0 13.24.42c-.51-.198-1.092-.333-1.943-.372C10.443.01 10.172 0 7.998 0zm-.717 1.442h.718c2.136 0 2.389.007 3.232.046.78.035 1.204.166 1.486.275.373.145.64.319.92.599s.453.546.598.92c.11.281.24.705.275 1.485.039.843.047 1.096.047 3.231s-.008 2.389-.047 3.232c-.035.78-.166 1.203-.275 1.485a2.5 2.5 0 0 1-.599.919c-.28.28-.546.453-.92.598-.28.11-.704.24-1.485.276-.843.038-1.096.047-3.232.047s-2.39-.009-3.233-.047c-.78-.036-1.203-.166-1.485-.276a2.5 2.5 0 0 1-.92-.598 2.5 2.5 0 0 1-.6-.92c-.109-.281-.24-.705-.275-1.485-.038-.843-.046-1.096-.046-3.233s.008-2.388.046-3.231c.036-.78.166-1.204.276-1.486.145-.373.319-.64.599-.92s.546-.453.92-.598c.282-.11.705-.24 1.485-.276.738-.034 1.024-.044 2.515-.045zm4.988 1.328a.96.96 0 1 0 0 1.92.96.96 0 0 0 0-1.92m-4.27 1.122a4.109 4.109 0 1 0 0 8.217 4.109 4.109 0 0 0 0-8.217m0 1.441a2.667 2.667 0 1 1 0 5.334 2.667 2.667 0 0 1 0-5.334" />
                    </svg>
                    <span>Instagram</span>
                  </div>
                </li>
                <li>
                  <div data-tab-social-id="social-panel-twitter" class="sm-tab gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                      class="bi bi-twitter-x" viewBox="0 0 16 16">
                      <path
                        d="M12.6.75h2.454l-5.36 6.142L16 15.25h-4.937l-3.867-5.07-4.425 5.07H.316l5.733-6.57L0 .75h5.063l3.495 4.633L12.601.75Zm-.86 13.028h1.36L4.323 2.145H2.865z" />
                    </svg>
                    <span>X (Twitter)</span>
                  </div>
                </li>
                <li>
                  <div data-tab-social-id="social-panel-linkedin" class="sm-tab gap-2">
                    <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                      class="bi bi-linkedin" viewBox="0 0 16 16">
                      <path
                        d="M0 1.146C0 .513.526 0 1.175 0h13.65C15.474 0 16 .513 16 1.146v13.708c0 .633-.526 1.146-1.175 1.146H1.175C.526 16 0 15.487 0 14.854zm4.943 12.248V6.169H2.542v7.225zm-1.2-8.212c.837 0 1.358-.554 1.358-1.248-.015-.709-.52-1.248-1.342-1.248S2.4 3.226 2.4 3.934c0 .694.521 1.248 1.327 1.248zm4.908 8.212V9.359c0-.216.016-.432.08-.586.173-.431.568-.878 1.232-.878.869 0 1.216.662 1.216 1.634v3.865h2.401V9.25c0-2.22-1.184-3.252-2.764-3.252-1.274 0-1.845.7-2.165 1.193v.025h-.016l.016-.025V6.169h-2.4c.03.678 0 7.225 0 7.225z" />
                    </svg>
                    <span>LinkedIn</span>
                  </div>
                </li>
              </ul>
              <div id="social-panel-instagram" class="w-100 float-start social-media-view" style="display: block;">
                <div class="social-swiper swiper">
                  <div class="swiper-wrapper">

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-3.jfif" alt="Aluminium coils">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Celebrating excellence and teamwork across Sohar Aluminium.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">1d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-4.jfif" alt="Team of workers">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Empowering our people, building a stronger future.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">1d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-2.jfif" alt="Molten metal">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Sustainability is at the core of everything we do.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">2d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img
                          src="https://images.unsplash.com/photo-1504328345606-18bbc8c9d7d1?q=80&w=600&auto=format&fit=crop"
                          alt="Training session">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Investing in tomorrow's aluminium leaders through training.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">2h</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-3.jfif" alt="Aluminium coils">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Celebrating excellence and teamwork across Sohar Aluminium.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">1d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                  </div>
                  <div class="swiper-button-next swiper-btn-common social-swiper-next"></div>
                  <div class="swiper-button-prev swiper-btn-common social-swiper-prev"></div>
                </div>
              </div>
              <div id="social-panel-twitter" class="w-100 float-start social-media-view">
                <div class="social-swiper swiper">
                  <div class="swiper-wrapper">

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-3.jfif" alt="Aluminium coils">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Celebrating excellence and teamwork across Sohar Aluminium.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">1d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>
                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-4.jfif" alt="Team of workers">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Empowering our people, building a stronger future.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">1d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                  </div>
                  <div class="swiper-button-next swiper-btn-common social-swiper-next"></div>
                  <div class="swiper-button-prev swiper-btn-common social-swiper-prev"></div>
                </div>
              </div>
              <div id="social-panel-linkedin" class="w-100 float-start social-media-view">
                <div class="social-swiper swiper">
                  <div class="swiper-wrapper">

                    <div class="swiper-slide">
                      <div class="social-card d-flex flex-column">
                        <img src="./resources/images/gallery/gallery-3.jfif" alt="Aluminium coils">
                        <div class="px-3 py-12">
                          <p class="sc-account">sohar.aluminium</p>
                          <p class="sc-text">Celebrating excellence and teamwork across Sohar Aluminium.</p>
                          <div class="sc-footer">
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-heart" viewBox="0 0 16 16">
                              <path
                                d="m8 2.748-.717-.737C5.6.281 2.514.878 1.4 3.053c-.523 1.023-.641 2.5.314 4.385.92 1.815 2.834 3.989 6.286 6.357 3.452-2.368 5.365-4.542 6.286-6.357.955-1.886.838-3.362.314-4.385C13.486.878 10.4.28 8.717 2.01zM8 15C-7.333 4.868 3.279-3.04 7.824 1.143q.09.083.176.171a3 3 0 0 1 .176-.17C12.72-3.042 23.333 4.867 8 15" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-chat" viewBox="0 0 16 16">
                              <path
                                d="M2.678 11.894a1 1 0 0 1 .287.801 11 11 0 0 1-.398 2c1.395-.323 2.247-.697 2.634-.893a1 1 0 0 1 .71-.074A8 8 0 0 0 8 14c3.996 0 7-2.807 7-6s-3.004-6-7-6-7 2.808-7 6c0 1.468.617 2.83 1.678 3.894m-.493 3.905a22 22 0 0 1-.713.129c-.2.032-.352-.176-.273-.362a10 10 0 0 0 .244-.637l.003-.01c.248-.72.45-1.548.524-2.319C.743 11.37 0 9.76 0 8c0-3.866 3.582-7 8-7s8 3.134 8 7-3.582 7-8 7a9 9 0 0 1-2.347-.306c-.52.263-1.639.742-3.468 1.105" />
                            </svg>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-send" viewBox="0 0 16 16">
                              <path
                                d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11ZM6.636 10.07l2.761 4.338L14.13 2.576zm6.787-8.201L1.591 6.602l4.339 2.76z" />
                            </svg>
                            <span class="sc-time">1d</span>
                            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                              class="bi bi-bookmark" viewBox="0 0 16 16">
                              <path
                                d="M2 2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v13.5a.5.5 0 0 1-.777.416L8 13.101l-5.223 2.815A.5.5 0 0 1 2 15.5zm2-1a1 1 0 0 0-1 1v12.566l4.723-2.482a.5.5 0 0 1 .554 0L13 14.566V2a1 1 0 0 0-1-1z" />
                            </svg>
                          </div>
                        </div>
                      </div>
                    </div>

                  </div>
                  <div class="swiper-button-next swiper-btn-common social-swiper-next"></div>
                  <div class="swiper-button-prev swiper-btn-common social-swiper-prev"></div>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="col-12 col-lg-4">
          <div class="panel-card px-2 py-4">
            <div class="panel-header px-2 w-100 float-start mb-4">
              <h2 class="panel-title">Upcoming Birthdays</h2>
              <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
                </span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
            </div>
            <div
              class="w-100 d-flex flex-column float-start px-2 overflow-auto panel-card-birthdays custom-scroll-view">
              <div class="birthday-item">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                  class="bi bi-stars birthday-deco" viewBox="0 0 16 16">
                  <path
                    d="M7.657 6.247c.11-.33.576-.33.686 0l.645 1.937a2.89 2.89 0 0 0 1.829 1.828l1.936.645c.33.11.33.576 0 .686l-1.937.645a2.89 2.89 0 0 0-1.828 1.829l-.645 1.936a.361.361 0 0 1-.686 0l-.645-1.937a2.89 2.89 0 0 0-1.828-1.828l-1.937-.645a.361.361 0 0 1 0-.686l1.937-.645a2.89 2.89 0 0 0 1.828-1.828zM3.794 1.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387A1.73 1.73 0 0 0 4.593 5.69l-.387 1.162a.217.217 0 0 1-.412 0L3.407 5.69A1.73 1.73 0 0 0 2.31 4.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387A1.73 1.73 0 0 0 3.407 2.31zM10.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.16 1.16 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.16 1.16 0 0 0-.732-.732L9.1 2.137a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732z" />
                </svg>
                <img src="https://randomuser.me/api/portraits/women/65.jpg" class="birthday-avatar" alt="Emily Johnson">
                <p class="birthday-name">Emily Johnson</p>
                <div class="birthday-date d-flex">
                  <span>MAY</span>
                  <span>21</span>
                </div>
              </div>
              <div class="birthday-item">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                  class="bi bi-stars birthday-deco" viewBox="0 0 16 16">
                  <path
                    d="M7.657 6.247c.11-.33.576-.33.686 0l.645 1.937a2.89 2.89 0 0 0 1.829 1.828l1.936.645c.33.11.33.576 0 .686l-1.937.645a2.89 2.89 0 0 0-1.828 1.829l-.645 1.936a.361.361 0 0 1-.686 0l-.645-1.937a2.89 2.89 0 0 0-1.828-1.828l-1.937-.645a.361.361 0 0 1 0-.686l1.937-.645a2.89 2.89 0 0 0 1.828-1.828zM3.794 1.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387A1.73 1.73 0 0 0 4.593 5.69l-.387 1.162a.217.217 0 0 1-.412 0L3.407 5.69A1.73 1.73 0 0 0 2.31 4.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387A1.73 1.73 0 0 0 3.407 2.31zM10.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.16 1.16 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.16 1.16 0 0 0-.732-.732L9.1 2.137a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732z" />
                </svg>
                <img src="https://randomuser.me/api/portraits/men/45.jpg" class="birthday-avatar" alt="Michael Brown">
                <p class="birthday-name">Michael Brown</p>
                <div class="birthday-date d-flex">
                  <span>MAY</span>
                  <span>22</span>
                </div>
              </div>
              <div class="birthday-item">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                  class="bi bi-stars birthday-deco" viewBox="0 0 16 16">
                  <path
                    d="M7.657 6.247c.11-.33.576-.33.686 0l.645 1.937a2.89 2.89 0 0 0 1.829 1.828l1.936.645c.33.11.33.576 0 .686l-1.937.645a2.89 2.89 0 0 0-1.828 1.829l-.645 1.936a.361.361 0 0 1-.686 0l-.645-1.937a2.89 2.89 0 0 0-1.828-1.828l-1.937-.645a.361.361 0 0 1 0-.686l1.937-.645a2.89 2.89 0 0 0 1.828-1.828zM3.794 1.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387A1.73 1.73 0 0 0 4.593 5.69l-.387 1.162a.217.217 0 0 1-.412 0L3.407 5.69A1.73 1.73 0 0 0 2.31 4.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387A1.73 1.73 0 0 0 3.407 2.31zM10.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.16 1.16 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.16 1.16 0 0 0-.732-.732L9.1 2.137a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732z" />
                </svg>
                <img src="https://randomuser.me/api/portraits/women/68.jpg" class="birthday-avatar"
                  alt="Sarah Williams">
                <p class="birthday-name">Sarah Williams</p>
                <div class="birthday-date d-flex">
                  <span>MAY</span>
                  <span>24</span>
                </div>
              </div>
              <div class="birthday-item">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                  class="bi bi-stars birthday-deco" viewBox="0 0 16 16">
                  <path
                    d="M7.657 6.247c.11-.33.576-.33.686 0l.645 1.937a2.89 2.89 0 0 0 1.829 1.828l1.936.645c.33.11.33.576 0 .686l-1.937.645a2.89 2.89 0 0 0-1.828 1.829l-.645 1.936a.361.361 0 0 1-.686 0l-.645-1.937a2.89 2.89 0 0 0-1.828-1.828l-1.937-.645a.361.361 0 0 1 0-.686l1.937-.645a2.89 2.89 0 0 0 1.828-1.828zM3.794 1.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387A1.73 1.73 0 0 0 4.593 5.69l-.387 1.162a.217.217 0 0 1-.412 0L3.407 5.69A1.73 1.73 0 0 0 2.31 4.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387A1.73 1.73 0 0 0 3.407 2.31zM10.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.16 1.16 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.16 1.16 0 0 0-.732-.732L9.1 2.137a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732z" />
                </svg>
                <img src="https://randomuser.me/api/portraits/men/22.jpg" class="birthday-avatar" alt="David Lee">
                <p class="birthday-name">David Lee</p>
                <div class="birthday-date d-flex">
                  <span>MAY</span>
                  <span>21</span>
                </div>
              </div>
              <div class="birthday-item">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor"
                  class="bi bi-stars birthday-deco" viewBox="0 0 16 16">
                  <path
                    d="M7.657 6.247c.11-.33.576-.33.686 0l.645 1.937a2.89 2.89 0 0 0 1.829 1.828l1.936.645c.33.11.33.576 0 .686l-1.937.645a2.89 2.89 0 0 0-1.828 1.829l-.645 1.936a.361.361 0 0 1-.686 0l-.645-1.937a2.89 2.89 0 0 0-1.828-1.828l-1.937-.645a.361.361 0 0 1 0-.686l1.937-.645a2.89 2.89 0 0 0 1.828-1.828zM3.794 1.148a.217.217 0 0 1 .412 0l.387 1.162c.173.518.579.924 1.097 1.097l1.162.387a.217.217 0 0 1 0 .412l-1.162.387A1.73 1.73 0 0 0 4.593 5.69l-.387 1.162a.217.217 0 0 1-.412 0L3.407 5.69A1.73 1.73 0 0 0 2.31 4.593l-1.162-.387a.217.217 0 0 1 0-.412l1.162-.387A1.73 1.73 0 0 0 3.407 2.31zM10.863.099a.145.145 0 0 1 .274 0l.258.774c.115.346.386.617.732.732l.774.258a.145.145 0 0 1 0 .274l-.774.258a1.16 1.16 0 0 0-.732.732l-.258.774a.145.145 0 0 1-.274 0l-.258-.774a1.16 1.16 0 0 0-.732-.732L9.1 2.137a.145.145 0 0 1 0-.274l.774-.258c.346-.115.617-.386.732-.732z" />
                </svg>
                <img src="https://randomuser.me/api/portraits/women/33.jpg" class="birthday-avatar" alt="Priya Sharma">
                <p class="birthday-name">Priya Sharma</p>
                <div class="birthday-date d-flex">
                  <span>MAY</span>
                  <span>28</span>
                </div>
              </div>
            </div>
          </div>
        </div>
        <div class="col-12">
          <div class="panel-card px-3 py-4">
            <div class="panel-header w-100 float-start mb-4">
              <h2 class="panel-title">Media Gallery</h2>
              <a href="#" class="link-arrow text-color-link"><span class="text-sm xxl-text-base font-bold">View All
                </span><img src="./resources/images/icons/arrow-right-short.svg" /></a>
            </div>
            <div class="w-100 d-flex flex-column float-start panel-card-gallery">
              <div class="gallery-swiper swiper">
                <div class="swiper-wrapper">
                  <div class="swiper-slide">
                    <div class="gallery-item">
                      <img src="./resources/images/gallery/gallery-1.jfif" alt="Plant at dusk">
                    </div>
                  </div>
                  <div class="swiper-slide">
                    <div class="gallery-item">
                      <img src="./resources/images/gallery/gallery-2.jfif" alt="Molten metal pour">
                    </div>
                  </div>
                  <div class="swiper-slide">
                    <div class="gallery-item">
                      <img src="./resources/images/gallery/gallery-3.jfif" alt="Aluminium sheets">
                    </div>
                  </div>
                  <div class="swiper-slide">
                    <div class="gallery-item">
                      <img src="./resources/images/gallery/gallery-1.jfif" alt="Plant at dusk">
                    </div>
                  </div>
                  <div class="swiper-slide">
                    <div class="gallery-item">
                      <img src="./resources/images/gallery/gallery-2.jfif" alt="Molten metal pour">
                    </div>
                  </div>

                </div>
                <div class="swiper-button-next swiper-btn-common gallery-slide-next"></div>
                <div class="swiper-button-prev swiper-btn-common gallery-slide-prev"></div>
              </div>
              <div class="gallery-swiper-pagination"></div>
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <section class="pre-footer-info">
    <div class="container container-sa px-3 px-lg-4">
      <div class="row gy-4">

        <div class="col-md-6 pre-footer-col pe-md-5">
          <div class="pre-footer-heading">
            <!-- <span class="pre-footer-heading-icon">
              <img src="./resources/images/icons/footer/ip-contacts.png" />
            </span> -->
            <h3 class="pre-footer-title">Important contacts</h3>
          </div>

          <ul class="pre-footer-list">
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/emergency-call.png" />
              </span>
              <span class="pre-footer-item-text">For Emergencies call <strong>(+968)26863333</strong> or Ext. 3333
                internally</span>
            </li>
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/manager.png" />
              </span>
              <span class="pre-footer-item-text">Sohar Aluminium Duty Manager can be reached at
                <strong>(+968)99855907</strong></span>
            </li>
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/email.png" />
              </span>
              <span class="pre-footer-item-text">Should you wish to give feedback or report anything concerning Sohar
                Aluminium please contact on <a href="mailto:hotline@sohar-aluminium.com"
                  class="pre-footer-link">hotline@sohar-aluminium.com</a> or <strong>(+968)26863317</strong></span>
            </li>
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/support.png" />
              </span>
              <span class="pre-footer-item-text">IT Support contact ( Automation: <strong>99100320</strong> , IT
                Support: <strong>92887255</strong> , MES: <strong>92802422</strong> ).</span>
            </li>
          </ul>
        </div>

        <div class="col-md-6 pre-footer-col pre-footer-col-divider ps-md-5">
          <div class="pre-footer-heading">
            <!-- <span class="pre-footer-heading-icon">
              <img src="./resources/images/icons/footer/medical.png" />
            </span> -->
            <h3 class="pre-footer-title">Medical services</h3>
          </div>

          <ul class="pre-footer-list">
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/medical-emergency.png" />
              </span>
              <span class="pre-footer-item-text">SA Medical Emergency: <strong>(+968)26863222</strong> or Ext. 3222
                internally</span>
            </li>
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/medical-front-desk.png" />
              </span>
              <span class="pre-footer-item-text">SA Medical Front Desk: <strong>(+968)26863267</strong> or Ext. 3267
                internally</span>
            </li>
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/medical-helpline.png" />
              </span>
              <span class="pre-footer-item-text">SA Medical helpline (Open 24 hours /7 days a week):
                <strong>92808667</strong></span>
            </li>
            <li>
              <span class="pre-footer-item-icon">
                <img src="./resources/images/icons/footer/nursing-admin.png" />
              </span>
              <span class="pre-footer-item-text">Nursing Admin Desk: Ext. 3281, Dr Hani Al Shuraiqi: Ext. 3282</span>
            </li>
          </ul>
        </div>

      </div>
    </div>
  </section>
  <!-- Widget start here -->
  <footer class="w-100 d-flex align-items-center">
    <div class="container container-sa d-flex align-items-center px-3 px-lg-4 justify-content-center">
      <p class="copyright-text text-white text-center text-xs font-normal m-0 text-capitalize">© 2025 Sohar Aluminium.
        All Rights Reserved.
      </p>
    </div>
  </footer>
  <!-- Widget end here -->


  <div class="modal fade" id="galleryModal" tabindex="-1" aria-labelledby="galleryModal" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable gallery-modal">
      <div class="modal-content py-2 px-2">
        <div class="gallery-modal-header d-flex justify-content-end me-2 mt-2">
          <svg class="cursor-pointer flex-shrink-0" data-bs-dismiss="modal" aria-label="Close" width="15" height="15"
            viewBox="0 0 15 15" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016" stroke="#005187"
              stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round" />
          </svg>
        </div>
        <div class="modal-body gallery-modal-content">
          <!-- <img id="galleryModalImg" src="./resources/images/gallery/gallery-1-lg.png" alt="" /> -->
          <!-- <video controls>
            <source src="" type="video/mp4">
            Your browser does not support the video tag.
          </video> -->
          <div class="w-100 float-start position-relative gallery-swiper-wrapper">
            <div class="w-100 float-start gallery-modal-swiper swiper">
              <div class="swiper-wrapper">
                <!-- <div class="swiper-slide gallery-swiper-slide">
                  <video controls>
                    <source src="./resources/images/gallery/istockphoto-1253263447-640_adpp_is 1.mp4" type="video/mp4">
                    Your browser does not support the video tag.
                  </video>
                </div> -->
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-1.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-2.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-3.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-1.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-2.jfif" alt="" />
                </div>
              </div>
            </div>
            <div class="swiper-button-next swiper-btn-common gallery-swiper-modal-next"></div>
            <div class="swiper-button-prev swiper-btn-common gallery-swiper-modal-prev"></div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="modal fade favourite-modal" id="addFavouriteModal" tabindex="-1" aria-labelledby="addFavouriteModalLabel"
    aria-hidden="true">

    <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable custom-scroll-view">

      <div class="modal-content">

        <div class="modal-header">
          <h5 class="modal-title" id="addFavouriteModalLabel">
            Add to Favourites
          </h5>

          <svg class="cursor-pointer flex-shrink-0" data-bs-dismiss="modal" aria-label="Close" width="15" height="15"
            viewBox="0 0 15 15" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016" stroke="#005187"
              stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"></path>
          </svg>
        </div>

        <div class="modal-body">

          <p class="mb-3 modal-body-fav-title">
            Select the quick links you want to add to your favourites.
          </p>

          <div class="favourite-options">

            <label class="favourite-option">
              <input type="checkbox" value="Payslip" data-favourite="ql-favourite">
              <span>Payslip</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Leave Application" data-favourite="ql-favourite">
              <span>Leave Application</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Travel Request" data-favourite="ql-favourite">
              <span>Travel Request</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Expense Claim" data-favourite="ql-favourite">
              <span>Expense Claim</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Training Portal" data-favourite="ql-favourite">
              <span>Training Portal</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="IT Help Desk" data-favourite="ql-favourite">
              <span>IT Help Desk</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Safety Portal" data-favourite="ql-favourite">
              <span>Safety Portal</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Idea Box" data-favourite="ql-favourite">
              <span>Idea Box</span>
            </label>

          </div>

        </div>

        <div class="modal-footer">
          <button type="button" class="btn-appearance-none btn-brand btn-brand-secondary px-3 py-1 text-sm"
            data-bs-dismiss="modal">
            Cancel
          </button>

          <button type="button" class="btn-appearance-none btn-brand btn-brand-primary px-3 py-1 text-sm" id="btnAddFavourites">
            Add to Favourites
          </button>
        </div>

      </div>

    </div>
  </div>


</body>

</html>
    `;

}