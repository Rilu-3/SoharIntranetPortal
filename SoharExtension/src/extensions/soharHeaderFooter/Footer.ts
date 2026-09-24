export default class Footer {

  public static html: string = `

    <section class="pre-footer-info">

      <div class="container container-sa px-3 px-lg-4">

        <div class="row gy-4">

          <div class="col-md-6 pre-footer-col pe-md-5">

            <div class="pre-footer-heading">

              <h3 class="pre-footer-title">
                Important contacts
              </h3>

            </div>

            <ul class="pre-footer-list">
              __IMPORTANT_CONTACTS__
            </ul>

          </div>

          <div class="col-md-6 pre-footer-col pre-footer-col-divider ps-md-5">

            <div class="pre-footer-heading">

              <h3 class="pre-footer-title">
                Medical services
              </h3>

            </div>

            <ul class="pre-footer-list">
              __MEDICAL_SERVICES__
            </ul>

          </div>

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