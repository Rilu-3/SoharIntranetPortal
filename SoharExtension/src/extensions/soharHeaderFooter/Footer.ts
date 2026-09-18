export default class Footer {

  private siteUrl: string;

  constructor(siteUrl: string) {
    this.siteUrl = siteUrl;
  }

  public render(): string {
    return `
      <section class="pre-footer-info">
        <div class="container container-sa px-3 px-lg-4">
          <div class="row gy-4">

            <div class="col-md-6 pre-footer-col pe-md-5">
              <div class="pre-footer-heading">

                <!-- <span class="pre-footer-heading-icon">
                  <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/ip-contacts.png" />
                </span> -->

                <h3 class="pre-footer-title">Important contacts</h3>
              </div>

              <ul class="pre-footer-list">

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/emergency-call.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    For Emergencies call
                    <strong>(+968)26863333</strong>
                    or Ext. 3333 internally
                  </span>
                </li>

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/manager.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    Sohar Aluminium Duty Manager can be reached at
                    <strong>(+968)99855907</strong>
                  </span>
                </li>

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/email.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    Should you wish to give feedback or report anything concerning
                    Sohar Aluminium please contact on
                    <a
                      href="mailto:hotline@sohar-aluminium.com"
                      class="pre-footer-link"
                    >
                      hotline@sohar-aluminium.com
                    </a>
                    or <strong>(+968)26863317</strong>
                  </span>
                </li>

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/support.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    IT Support contact (
                    Automation: <strong>99100320</strong>,
                    IT Support: <strong>92887255</strong>,
                    MES: <strong>92802422</strong>
                    ).
                  </span>
                </li>

              </ul>
            </div>

            <div class="col-md-6 pre-footer-col pre-footer-col-divider ps-md-5">

              <div class="pre-footer-heading">

                <!-- <span class="pre-footer-heading-icon">
                  <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/medical.png" />
                </span> -->

                <h3 class="pre-footer-title">Medical services</h3>
              </div>

              <ul class="pre-footer-list">

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/medical-emergency.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    SA Medical Emergency:
                    <strong>(+968)26863222</strong>
                    or Ext. 3222 internally
                  </span>
                </li>

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/medical-front-desk.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    SA Medical Front Desk:
                    <strong>(+968)26863267</strong>
                    or Ext. 3267 internally
                  </span>
                </li>

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/medical-helpline.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    SA Medical helpline
                    (Open 24 hours /7 days a week):
                    <strong>92808667</strong>
                  </span>
                </li>

                <li>
                  <span class="pre-footer-item-icon">
                    <img src="${this.siteUrl}/SiteAssets/resources/images/icons/footer/nursing-admin.png" />
                  </span>

                  <span class="pre-footer-item-text">
                    Nursing Admin Desk: Ext. 3281,
                    Dr Hani Al Shuraiqi: Ext. 3282
                  </span>
                </li>

              </ul>
            </div>

          </div>
        </div>
      </section>

      <!-- Widget start here -->

      <footer class="w-100 d-flex align-items-center">
        <div class="container container-sa d-flex align-items-center px-3 px-lg-4 justify-content-center">

          <p class="copyright-text text-white text-center text-xs font-normal m-0 text-capitalize">
            © 2025 Sohar Aluminium. All Rights Reserved.
          </p>

        </div>
      </footer>
    `;
  }
}
