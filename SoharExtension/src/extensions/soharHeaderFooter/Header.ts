
export default class Header {

  constructor(
    private siteUrl: string,
    private userName: string,
    private designation: string,
    private profilePhoto: string,
    private departmentItems: any[],
    private controlledDocumentItems: any[]
  ) { }

  public render(): string {

    return `
      <!-- Widget start here -->

      <nav class="navbar navbar-expand-lg navbar-light fixed-top sa-header py-0 flex-column">

        <div class="w-100 float-start sa-header-top py-2">

          <div class="container container-sa px-3 px-lg-4 d-flex align-items-center gap-3">

            <!-- Logo -->

            <a
              class="navbar-brand"
              href="/sites/DevPortal/SitePages/Home.aspx?env=WebViewList"
              target="_self"
              data-interception="off">

              <img
                class="logo-desktop"
                src="${this.siteUrl}/SiteAssets/resources/images/logo.png"
                alt="Logo"
              />

              <img
                class="logo-mob"
                src="${this.siteUrl}/SiteAssets/resources/images/logo-mob.png"
                alt="Logo"
              />

            </a>

            <!-- Global Search -->

            <div class="header-search position-relative mx-lg-auto d-none d-lg-block">

              <input
                type="text"
                class="form-control"
                id="inputGlobalSearchBox"
                placeholder="Search..."
              />

              <button
                class="search-icon-btn"
                type="button"
                aria-label="Search">

                <i class="bi bi-search"></i>

              </button>

            </div>

            <!-- User Information -->

            <div
              class="ms-lg-0 ms-auto nav-user-info d-flex align-items-center justify-content-between justify-content-lg-start gap-3 gap-lg-4">

              <div
                class="btn btn-secondary d-flex align-items-center p-0 bg-transparent border-0 gap-2">

                <div class="nav-avatar-wrapper d-flex align-items-center gap-2">

                  <img
                    class="nav-avatar"
                    src="${this.profilePhoto}"
                    alt="Profile"
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

            <!-- Mobile Menu Toggle -->

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

        <!-- Navigation Menu -->

        <div class="w-100 float-start sa-header-bottom">

          <div
            class="collapse navbar-collapse gap-lg-2"
            id="navbarSupportedContent">

            <div
              class="container container-sa px-3 px-lg-4 d-flex gap-3 py-lg-2">

              <ul
                class="navbar-nav my-2 my-lg-0 navbar-nav-scroll gap-lg-3"
                style="--bs-scroll-height: 400px;">

                <!-- Home -->

                <li class="nav-item d-lg-flex">

                  <a
                    class="nav-link active-nav-link d-flex align-items-center gap-2"
                    aria-current="page"
                    href="/sites/DevPortal/SitePages/Home.aspx?env=WebViewList"
                    target="_self"
                    data-interception="off">

                    <img
                      class="nav-menu-icon"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/home.png"
                      alt=""
                    />

                    <span>Home</span>

                  </a>

                </li>

                <!-- Document Centre -->

                <li class="nav-item d-lg-flex dropdown">

                  <a
                    data-bs-auto-close="outside"
                    class="nav-link dropdown-toggle d-flex align-items-center gap-2"
                    href="#"
                    id="navbarScrollingDropdown"
                    role="button"
                    data-bs-toggle="dropdown"
                    aria-expanded="false">

                    <img
                      class="nav-menu-icon"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/folder.png"
                      alt=""
                    />

                    <span>Document Center</span>

                    <img
                      class="dropdown-arrow-nav"
                      src="${this.siteUrl}/SiteAssets/resources/images/icons/dropdown-arrow.png"
                      alt=""
                    />

                  </a>

                  <div
                    class="dropdown-menu"
                    aria-labelledby="navbarScrollingDropdown">

                    <ul class="custom-scroll-view dept-dropdown">

                      <!-- Departments Submenu -->

                      <li class="dept-submenu">

                        <a
                          class="dropdown-item text-sm d-flex justify-content-between align-items-center dept-submenu-toggle gap-2"
                          href="#"
                          role="button"
                          aria-expanded="false">

                          <div class="d-flex align-items-center gap-2 flex-grow-1 overflow-hidden min-w-0">

                            <img
                              class="sub-menu-icon sub-menu-icon-default"
                              src="${this.siteUrl}/SiteAssets/resources/images/icons/department-link-icon.png"
                              alt=""
                            />

                            <div class="doc-centre-title">Departments</div>

                          </div>

                          <img
                            class="submenu-arrow"
                            src="${this.siteUrl}/SiteAssets/resources/images/icons/dropdown-arrow-neutral.png"
                            alt=""
                          />

                        </a>

                        <div class="dept-submenu-list">

                          <input
                            type="text"
                            class="form-control search-dept-dropdown"
                            placeholder="Search departments..."
                          />

                          <ul>

                            ${this.departmentItems.map((department: any) => {

                              const link = department.Link?.Url || '#';

                              return `
                                <li>
                                  <a
                                    class="dropdown-item text-sm"
                                    href="${link}"
                                    target="_self"
                                    data-interception="off">

                                    <span>${department.Title}</span>

                                  </a>
                                </li>
                              `;

                            }).join('')}

                          </ul>

                        </div>

                      </li>

                      <!-- Controlled Documents Submenu -->

                      <li class="dept-submenu">

                        <a
                          class="dropdown-item text-sm d-flex justify-content-between align-items-center dept-submenu-toggle gap-2"
                          href="#"
                          role="button"
                          aria-expanded="false">

                          <div class="d-flex align-items-center gap-2 flex-grow-1 overflow-hidden min-w-0">

                            <img
                              class="sub-menu-icon sub-menu-icon-default"
                              src="${this.siteUrl}/SiteAssets/resources/images/icons/controlled-doc.png"
                              alt=""
                            />

                            <div class="doc-centre-title">Controlled Documents</div>

                          </div>

                          <img
                            class="submenu-arrow"
                            src="${this.siteUrl}/SiteAssets/resources/images/icons/dropdown-arrow-neutral.png"
                            alt=""
                          />

                        </a>

                        <div class="dept-submenu-list">

                          <input
                            type="text"
                            class="form-control search-dept-dropdown"
                            placeholder="Search documents..."
                          />

                          <ul>

                            ${this.controlledDocumentItems.map((document: any) => {

                              const link = document.Link?.Url || '#';

                              return `
                                <li>
                                  <a
                                    class="dropdown-item text-sm"
                                    href="${link}"
                                    target="_self"
                                    data-interception="off">

                                    <span>${document.Title}</span>

                                  </a>
                                </li>
                              `;

                            }).join('')}

                          </ul>

                        </div>

                      </li>

                    </ul>

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

