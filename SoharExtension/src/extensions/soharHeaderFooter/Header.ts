export default class Header {

  constructor(
    private siteUrl: string,
    private userName: string,
    private designation: string,
    private profilePhoto: string,
    private departmentItems: string[]
  ) { }

  public render(): string {

    const homeUrl: string = `${this.siteUrl}/SitePages/Home.aspx?env=WebViewList`;

    return `
      <!-- Widget start here -->

      <nav class="navbar navbar-expand-lg navbar-light fixed-top sa-header py-0 flex-column">

        <div class="w-100 float-start sa-header-top py-2">

          <div class="container container-sa px-3 px-lg-4 d-flex align-items-center gap-3">

            <a class="navbar-brand" href="${homeUrl}">

              <img
                class="logo-desktop"
                src="${this.siteUrl}/SiteAssets/resources/images/logo.png"
                alt="Sohar Aluminium"
              />

              <img
                class="logo-mob"
                src="${this.siteUrl}/SiteAssets/resources/images/logo-mob.png"
                alt="Sohar Aluminium"
              />

            </a>

            <div class="header-search position-relative mx-lg-auto d-none d-lg-block">

              <input
                type="text"
                class="form-control"
                id="inputGlobalSearchBox"
                placeholder="Search..."
                autocomplete="off"
              >

              <button
                class="search-icon-btn"
                type="button"
                aria-label="Search">

                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16" aria-hidden="true">
                  <path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001q.044.06.098.115l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85a1 1 0 0 0-.115-.1zM12 6.5a5.5 5.5 0 1 1-11 0 5.5 5.5 0 0 1 11 0"/>
                </svg>

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
                    alt=""
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
                    href="${homeUrl}">

                    <img
                      class="nav-menu-icon"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/home.png"
                      alt=""
                    />

                    <span>Home</span>

                  </a>

                </li>

                <li class="nav-item d-lg-flex has-mega">

                  <a
                    class="nav-link d-flex align-items-center gap-2"
                    href="#"
                    aria-haspopup="true"
                    aria-expanded="false">

                    <img
                      class="nav-menu-icon"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/department.png"
                      alt=""
                    />

                    <span>Departments</span>

                  </a>

                  <div
                    class="mega-panel"
                    role="menu"
                    aria-label="Department menu">

                    <div class="mega-panel-inner">

                      <div class="mega-panel-content custom-scroll-view">

                        ${this.departmentItems.map((items: string) => `
                          <div>
                            <ul class="mega-links">
                              ${items}
                            </ul>
                          </div>
                        `).join('')}

                      </div>

                    </div>

                  </div>

                </li>

              </ul>

            </div>

          </div>

        </div>

      </nav>
    `;
  }
}