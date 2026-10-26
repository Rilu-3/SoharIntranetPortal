
export default class Footer {

  public static html: string = `

    <section class="pre-footer-info contact-strip" aria-label="Key contacts">

      <div class="container container-sa px-3 px-lg-4">

        <div class="contact-strip-inner">

          <div class="contact-strip-items">

            <div class="contact-strip-item">

              <span class="pre-footer-item-icon">
<img src="https://soharaluminium5.sharepoint.com/sites/DevPortal/SiteAssets/resources/images/icons/footer/emergency-call.png" alt="" />              </span>

              <div class="contact-strip-text">
                <p class="contact-strip-label">Emergencies</p>
                <p class="contact-strip-value">
                  <a class="contact-tel" href="tel:+96826863333">
                    <strong>(+968)26863333</strong>
                  </a>
                  or Ext. 3333 internally
                </p>
              </div>

            </div>

            <div class="contact-strip-item">

              <span class="pre-footer-item-icon">
                <img src="https://soharaluminium5.sharepoint.com/sites/DevPortal/SiteAssets/resources/images/icons/footer/email.png" alt="" />
              </span>

              <div class="contact-strip-text">
                <p class="contact-strip-label">Feedback &amp; reporting</p>
                <p class="contact-strip-value">
                  <a href="mailto:hotline@sohar-aluminium.com" class="pre-footer-link">
                    hotline@sohar-aluminium.com
                  </a>
                  or
                  <a class="contact-tel" href="tel:+96826863317">
                    <strong>(+968)26863317</strong>
                  </a>
                </p>
              </div>

            </div>

          </div>

          <button
            type="button"
            class="contact-strip-btn"
            data-bs-toggle="offcanvas"
            data-bs-target="#allContactsPanel"
            aria-controls="allContactsPanel">
            All contacts
          </button>

        </div>

      </div>

    </section>

    <footer class="w-100 d-flex align-items-center">

      <div class="container container-sa d-flex align-items-center px-3 px-lg-4 justify-content-center">

        <p class="copyright-text text-white text-center text-xs font-normal m-0 text-capitalize">
          © __CURRENT_YEAR__ Sohar Aluminium. All Rights Reserved.
        </p>

      </div>

    </footer>
<div
  class="offcanvas offcanvas-end contacts-offcanvas"
  tabindex="-1"
  id="allContactsPanel"
  aria-labelledby="allContactsTitle">

  <div class="offcanvas-header">
    <h3 class="offcanvas-title" id="allContactsTitle">
      All contacts
    </h3>

    <button
      type="button"
      class="btn-close btn-close-white"
      data-bs-dismiss="offcanvas"
      aria-label="Close">
    </button>
  </div>

  <div class="offcanvas-body custom-scroll-view">
    <ul class="nav contacts-tabs" role="tablist">

      <li class="nav-item" role="presentation">
        <button
          class="nav-link active"
          id="important-contacts-tab"
          data-bs-toggle="tab"
          data-bs-target="#importantContacts"
          type="button"
          role="tab"
          aria-controls="importantContacts"
          aria-selected="true">
          Important Contacts
        </button>
      </li>

      <li class="nav-item" role="presentation">
        <button
          class="nav-link"
          id="medical-services-tab"
          data-bs-toggle="tab"
          data-bs-target="#medicalServices"
          type="button"
          role="tab"
          aria-controls="medicalServices"
          aria-selected="false">
          Medical Services
        </button>
      </li>

    </ul>

    <div class="tab-content">
      <div
        class="tab-pane fade show active"
        id="importantContacts"
        role="tabpanel"
        aria-labelledby="important-contacts-tab"
        tabindex="0">

        <ul class="pre-footer-list">
          __IMPORTANT_CONTACTS__
        </ul>
      </div>

      <div
        class="tab-pane fade"
        id="medicalServices"
        role="tabpanel"
        aria-labelledby="medical-services-tab"
        tabindex="0">

        <ul class="pre-footer-list">
          __MEDICAL_SERVICES__
        </ul>
      </div>
    </div>
  </div>
</div>
  `;

  public static itemTemplate: string = `
    <li>

      <span class="pre-footer-item-icon">
        <img src="__ICON_URL__" alt="" />
      </span>

      <span class="pre-footer-item-text">
        __CONTACT_DETAILS__
      </span>

    </li>
  `;
}
