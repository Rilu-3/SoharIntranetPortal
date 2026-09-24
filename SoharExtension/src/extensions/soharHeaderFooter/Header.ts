export default class Header {

  constructor(
    private siteUrl: string,
    private userName: string,
    private designation: string,
    private profilePhoto: string,
    private departmentItems: string
  ) {}

  public render(): string {

    return `
      <!-- Widget start here -->

      <nav class="navbar navbar-expand-lg navbar-light fixed-top sa-header py-0 flex-column">

        <div class="w-100 float-start sa-header-top py-2">

          <div class="container container-sa px-3 px-lg-4 d-flex align-items-center gap-3">

            <a class="navbar-brand" href="#">

              <img
                class="logo-desktop"
                src="${this.siteUrl}/SiteAssets/resources/images/logo.png"
              />

              <img
                class="logo-mob"
                src="${this.siteUrl}/SiteAssets/resources/images/logo-mob.png"
              />

            </a>

            <div class="header-search position-relative mx-lg-auto d-none d-lg-block">

              <input
                type="text"
                class="form-control"
                placeholder="Search people, documents, pages..."
              >

              <button
                class="search-icon-btn"
                type="button"
                aria-label="Search">

                <i class="bi bi-search"></i>

              </button>

            </div>

            <div
              class="ms-lg-0 ms-auto nav-user-info d-flex align-items-center justify-content-between justify-content-lg-start gap-3 gap-lg-4">

              <div
                class="btn btn-secondary d-flex align-items-center p-0 bg-transparent border-0 gap-2">

                <div class="nav-avatar-wrapper d-flex align-items-center gap-2">

                  <img
                    class="nav-avatar"
                    src="${this.profilePhoto}"
                  />

                  <div
                    class="max-w-200 d-flex flex-column text-start d-none d-sm-flex">

                    <p
                      class="text-sm xxl-text-base font-bold text-color-title lh-base text-truncate">
                      ${this.userName}
                    </p>

                    <p
                      class="text-xs font-normal text-color-desc lh-sm text-truncate">
                      ${this.designation}
                    </p>

                  </div>

                </div>

              </div>

            </div>

            <button
              class="navbar-toggler"
              type="button"
              data-bs-toggle="collapse"
              data-bs-target="#navbarSupportedContent"
              aria-controls="navbarSupportedContent"
              aria-expanded="false"
              aria-label="Toggle navigation">

              <span class="navbar-toggler-icon"></span>

            </button>

          </div>

        </div>

        <div class="w-100 float-start sa-header-bottom">

          <div
            class="collapse navbar-collapse gap-lg-2"
            id="navbarSupportedContent">

            <div
              class="container container-sa px-3 px-lg-4 d-flex gap-3 py-lg-2">

              <ul
                class="navbar-nav my-2 my-lg-0 navbar-nav-scroll gap-lg-3"
                style="--bs-scroll-height: 400px;">

                <li class="nav-item d-lg-flex">

                  <a
                    class="nav-link active-nav-link d-flex align-items-center gap-2"
                    aria-current="page"
                    href="#">

                    <img
                      class="nav-menu-icon"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/home.png"
                      alt=""
                    />

                    <span>Home</span>

                  </a>

                </li>

                <li class="nav-item d-lg-flex dropdown">

                  <a 
                      class="nav-link dropdown-toggle d-flex align-items-center gap-2"
  href="/sites/DevPortal/SitePages/Home.aspx?env=WebViewList"
  target="_blank"
  data-interception="off"
  id="navbarScrollingDropdown"
  role="button"
  aria-expanded="false"
  data-bs-toggle="dropdown" > 
                      
                    

                    <img
                      class="nav-menu-icon"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/department.png"
                      alt=""
                    />

                    <span>Departments</span>

                    <img
                      class="dropdown-arrow-nav"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/dropdown-arrow.png"
                      alt=""
                      
                    />

                  </a>

                  <ul
                    class="dropdown-menu"
                    aria-labelledby="navbarScrollingDropdown">

                    ${this.departmentItems}

                  </ul>

                </li>

              </ul>

            </div>

          </div>

        </div>

      </nav>
    `;
  }
}