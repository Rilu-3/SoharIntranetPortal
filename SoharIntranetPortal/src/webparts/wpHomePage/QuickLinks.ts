import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';


export interface IQuickLinksList {

  Id: number;

  Title: string;

  Icon: string;

  URL: {
    Url: string;
  };

  Status: string;

  SortOrder: number;

}


export default class QuickLinks {


  // =========================================================
  // Single Quick Link HTML
  // =========================================================

  public static singleElementHtml: string = `
    <a href="__KEY_URL_LINK__"
      target="__KEY_URL_TARGET__"
      class="quick-link-box"
      __KEY_DATA_FAVOURITE__>

      <img src="__KEY_URL_IMGICON__" alt="__KEY_DATA_TITLE__">

      <span>__KEY_DATA_TITLE__</span>

    </a>
  `;


  // =========================================================
  // Full Quick Links Section
  // =========================================================

  public static allElementsHtml: string = `
    <div class="col-12 col-lg-6">
      <div class="panel-card px-2 py-4">

        <div class="panel-header px-2 w-100 float-start pb-3 d-flex align-items-center">

          <div class="d-flex align-items-center flex-wrap gap-2 gap-sm-3">

            <h2
              data-filter-ql="quick-links"
              class="panel-title panel-title-filter panel-title-filter-active">
              Quick Links
            </h2>

            <h2
              data-filter-ql="favourites"
              class="panel-title panel-title-filter">
              Favourites
            </h2>

          </div>

          <button
            type="button"
            class="btn-appearance-none btn-add-favourite"
            title="Add to Favourites"
            data-bs-toggle="modal"
            data-bs-target="#addFavouriteModal">

            <svg xmlns="http://www.w3.org/2000/svg"
              width="20" height="20"
              fill="#000"
              viewBox="0 0 16 16">

              <path fill-rule="evenodd"
                d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2" />

            </svg>

          </button>

        </div>


        <div class="w-100 float-start p-2 overflow-auto panel-card-quick-links quick-links-grid custom-scroll-view">

          <!-- Dynamic Quick Links will be inserted here -->

        </div>

      </div>
    </div>


    <!-- Add Favourite Modal -->

    <div class="modal fade favourite-modal"
      id="addFavouriteModal"
      tabindex="-1"
      aria-labelledby="addFavouriteModalLabel"
      aria-hidden="true">

      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable custom-scroll-view">

        <div class="modal-content">

          <div class="modal-header">

            <h5 class="modal-title"
              id="addFavouriteModalLabel">
              Add to Favourites
            </h5>

            <svg
              class="cursor-pointer flex-shrink-0"
              data-bs-dismiss="modal"
              aria-label="Close"
              width="15"
              height="15"
              viewBox="0 0 15 15"
              fill="none"
              xmlns="http://www.w3.org/2000/svg">

              <path
                d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016"
                stroke="#005187"
                stroke-width="1.33333"
                stroke-linecap="round"
                stroke-linejoin="round">
              </path>

            </svg>

          </div>


          <div class="modal-body">

            <p class="mb-3 modal-body-fav-title">
              Select the quick links you want to add to your favourites.
            </p>

            <div class="favourite-options">

              <!-- Dynamic Favourite Options -->

            </div>

          </div>


          <div class="modal-footer">

            <button
              type="button"
              class="btn-appearance-none btn-brand btn-brand-secondary px-3 py-1 text-sm"
              data-bs-dismiss="modal">
              Cancel
            </button>

            <button
              type="button"
              class="btn-appearance-none btn-brand btn-brand-primary px-3 py-1 text-sm"
              id="btnAddFavourites">
              Add to Favourites
            </button>

          </div>

        </div>

      </div>

    </div>
  `;


  // =========================================================
  // Initialise Quick Links
  // =========================================================

  public static async initialize(
    context: WebPartContext,
    domElement: HTMLElement
  ): Promise<void> {

    try {

      const quickLinks =
        await this.getQuickLinks(
          context
        );


      const favouriteIds =
        await this.getUserFavouriteIds(
          context
        );


      this.renderQuickLinks(
        context,
        domElement,
        quickLinks,
        favouriteIds
      );


      this.renderFavouriteOptions(
        domElement,
        quickLinks,
        favouriteIds
      );


      this.setupTabs(
        domElement
      );


      this.setupAddFavouriteButton(
        context,
        domElement
      );

      this.setupFavouriteModal(
      context,
      domElement
    );


      this.updateFavouriteTabVisibility(
        domElement,
        favouriteIds
      );


    } catch (error) {

      console.error(
        'Error initializing Quick Links:',
        error
      );


      const container =
        domElement.querySelector(
          '.quick-links-grid'
        );


      if (container) {

        container.innerHTML = `
          <div class="p-3">
            Unable to load Quick Links.
          </div>
        `;

      }

    }

  }


  // =========================================================
  // Get Quick Links from SharePoint
  // =========================================================

  private static async getQuickLinks(
    context: WebPartContext
  ): Promise<IQuickLinksList[]> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const apiUrl =
      `${webUrl}` +
      `/_api/web/lists/GetByTitle('Quick_Links')/items` +
      `?$select=Id,Title,Icon,URL,Status,SortOrder` +
      `&$orderby=SortOrder asc`;


    const response =
      await context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {

      console.error(
        `Quick Links request failed: ${response.status}`
      );

      return [];

    }


    const data =
      await response.json();

    return data.value;

  }


  // =========================================================
  // Get Current User
  // =========================================================

  private static async getCurrentUser(
    context: WebPartContext
  ): Promise<any> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const response =
      await context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );


    if (!response.ok) {

      throw new Error(
        `Unable to get current user: ${response.status}`
      );

    }


    return await response.json();

  }


  // =========================================================
  // Get User Favourite IDs
  // =========================================================

  private static async getUserFavouriteIds(
    context: WebPartContext
  ): Promise<number[]> {

    try {

      const webUrl =
        context.pageContext.web.absoluteUrl;


      const currentUser =
        await this.getCurrentUser(
          context
        );


      const response =
        await context.spHttpClient.get(

          `${webUrl}` +
          `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items` +
          `?$select=Id,Quick_x0020_LinkId` +
          `&$filter=UserId eq ${currentUser.Id}`,

          SPHttpClient.configurations.v1

        );


      if (!response.ok) {

        console.error(
          'Unable to get user favourites:',
          response.status
        );

        return [];

      }


      const data =
        await response.json();


      return data.value.map(
        (item: { Quick_x0020_LinkId: number }) =>
          Number(
            item.Quick_x0020_LinkId
          )
      );


    } catch (error) {

      console.error(
        'Error loading user favourites:',
        error
      );

      return [];

    }

  }


  // =========================================================
  // Get Quick Link Image URL
  // =========================================================

  private static getImageUrl(
    context: WebPartContext,
    item: IQuickLinksList
  ): string {

    if (!item.Icon) {

      return '';

    }


    try {

      const imgData =
        JSON.parse(item.Icon);


      return (
        `${context.pageContext.web.absoluteUrl}` +
        `/Lists/Quick_Links/Attachments/` +
        `${item.Id}/` +
        `${imgData.fileName}`
      );


    } catch (error) {

      console.error(
        'Error parsing Quick Link Icon:',
        error
      );

      return '';

    }

  }


  // =========================================================
  // Render Quick Links
  // =========================================================

  private static renderQuickLinks(
    context: WebPartContext,
    domElement: HTMLElement,
    items: IQuickLinksList[],
    favouriteIds: number[]
  ): void {

    const container =
      domElement.querySelector(
        '.quick-links-grid'
      );


    if (!container) {

      console.error(
        'Quick Links container not found.'
      );

      return;

    }


    let allElementsHtml = '';


    const activeItems =
      items.filter(
        (item) =>
          item.Status === 'Active'
      );


    activeItems.forEach(
      (item) => {

        const imageUrl =
          this.getImageUrl(
            context,
            item
          );


        const isFavourite =
          favouriteIds.indexOf(
            item.Id
          ) !== -1;


        const favouriteAttribute =
          isFavourite
            ? 'data-tab-cat="ql-favourite"'
            : '';


        const singleElementHtml =
          this.singleElementHtml

            .replace(
              /__KEY_URL_IMGICON__/,
              imageUrl
            )

            .replace(
              /__KEY_DATA_TITLE__/g,
              item.Title || ''
            )

            .replace(
              /__KEY_URL_LINK__/,
              item.URL?.Url || '#'
            )

            .replace(
              /__KEY_URL_TARGET__/,
              '_blank'
            )

            .replace(
              /__KEY_DATA_FAVOURITE__/,
              favouriteAttribute
            );


        allElementsHtml +=
          singleElementHtml;

      }
    );


    container.innerHTML =
      allElementsHtml;

  }


  // =========================================================
  // Render Favourite Options in Modal
  // =========================================================

  private static renderFavouriteOptions(
    domElement: HTMLElement,
    items: IQuickLinksList[],
    favouriteIds: number[]
  ): void {

    const container =
      domElement.querySelector(
        '.favourite-options'
      );


    if (!container) {

      console.error(
        'Favourite options container not found.'
      );

      return;

    }


    const activeItems =
      items.filter(
        (item) =>
          item.Status === 'Active'
      );


    let optionsHtml = '';


    activeItems.forEach(
      (item) => {

        const isFavourite =
          favouriteIds.indexOf(
            item.Id
          ) !== -1;


        optionsHtml += `
          <label class="favourite-option">

            <input
              type="checkbox"
              value="${item.Id}"
              ${isFavourite ? 'checked' : ''}>

            <span>
              ${item.Title}
            </span>

          </label>
        `;

      }
    );


    container.innerHTML =
      optionsHtml;

  }


  // =========================================================
  // Setup Tabs
  // =========================================================

  private static setupTabs(
    domElement: HTMLElement
  ): void {

    const tabs =
      domElement.querySelectorAll(
        '[data-filter-ql]'
      );


    tabs.forEach(
      (tab) => {

        tab.addEventListener(
          'click',
          () => {

            const selectedTab =
              (tab as HTMLElement)
                .getAttribute(
                  'data-filter-ql'
                );


            const allLinks =
              domElement.querySelectorAll(
                '.quick-links-grid .quick-link-box'
              );


            tabs.forEach(
              (item) => {

                item.classList.remove(
                  'panel-title-filter-active'
                );

              }
            );


            tab.classList.add(
              'panel-title-filter-active'
            );


            allLinks.forEach(
              (link) => {

                const favouriteCategory =
                  link.getAttribute(
                    'data-tab-cat'
                  );


                if (
                  selectedTab ===
                  'favourites'
                ) {

                  if (
                    favouriteCategory ===
                    'ql-favourite'
                  ) {

                    (link as HTMLElement)
                      .style.display = '';

                  } else {

                    (link as HTMLElement)
                      .style.display = 'none';

                  }

                } else {

                  (link as HTMLElement)
                    .style.display = '';

                }

              }
            );

          }
        );

      }
    );

  }


  // =========================================================
  // Show / Hide Favourites Tab
  // =========================================================

  private static updateFavouriteTabVisibility(
    domElement: HTMLElement,
    favouriteIds: number[]
  ): void {

    const favouriteTab =
      domElement.querySelector(
        '[data-filter-ql="favourites"]'
      ) as HTMLElement;


    if (!favouriteTab) {

      return;

    }


    if (favouriteIds.length === 0) {

      favouriteTab.style.display =
        'none';

    } else {

      favouriteTab.style.display =
        '';

    }

  }


  // =========================================================
  // Setup Add Favourite Button
  // =========================================================

  private static setupAddFavouriteButton(
    context: WebPartContext,
    domElement: HTMLElement
  ): void {

    const button =
      domElement.querySelector(
        '#btnAddFavourites'
      );


    if (!button) {

      console.error(
        'Add Favourites button not found.'
      );

      return;

    }


    button.addEventListener(
      'click',
      async () => {

        await this.addFavourites(
          context,
          domElement
        );

      }
    );

  }
   // =========================================================
  // Setup Favourite Modal
  // =========================================================

  private static setupFavouriteModal(
    context: WebPartContext,
    domElement: HTMLElement
  ): void {

    const modal =
      domElement.querySelector(
        '#addFavouriteModal'
      );

    if (!modal) {

      console.error(
        'Favourite modal not found.'
      );

      return;

    }


    modal.addEventListener(
      'show.bs.modal',
      async () => {

        const favouriteIds =
          await this.getUserFavouriteIds(
            context
          );


        const checkboxes =
          domElement.querySelectorAll(
            '.favourite-options input[type="checkbox"]'
          );


        checkboxes.forEach(
          (checkbox) => {

            const input =
              checkbox as HTMLInputElement;

            const quickLinkId =
              Number(input.value);

            input.checked =
              favouriteIds.indexOf(
                quickLinkId
              ) !== -1;

          }
        );

      }
    );

  }
  // =========================================================
  // Add / Remove Favourites
  // =========================================================

  private static async addFavourites(
    context: WebPartContext,
    domElement: HTMLElement
  ): Promise<void> {

    try {

      const checkboxes =
        domElement.querySelectorAll(
          '.favourite-options input[type="checkbox"]'
        );


      const selectedIds: number[] = [];


      checkboxes.forEach(
        (checkbox) => {

          const input =
            checkbox as HTMLInputElement;


          if (input.checked) {

            selectedIds.push(
              Number(input.value)
            );

          }

        }
      );


      const webUrl =
        context.pageContext.web.absoluteUrl;


      const currentUser =
        await this.getCurrentUser(
          context
        );


      // =====================================================
      // Get existing favourites
      // =====================================================

      const favouritesResponse =
        await context.spHttpClient.get(

          `${webUrl}` +
          `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items` +
          `?$select=Id,Quick_x0020_LinkId` +
          `&$filter=UserId eq ${currentUser.Id}`,

          SPHttpClient.configurations.v1

        );


      if (!favouritesResponse.ok) {

        console.error(
          'Unable to get existing favourites:',
          favouritesResponse.status
        );

        return;

      }


      const favouritesData =
        await favouritesResponse.json();


      const existingFavourites =
        favouritesData.value;


      const keptIds: number[] = [];


      // =====================================================
      // Remove unchecked favourites
      // =====================================================

      for (
        let i = 0;
        i < existingFavourites.length;
        i++
      ) {

        const favourite =
          existingFavourites[i];


        const quickLinkId =
          Number(
            favourite.Quick_x0020_LinkId
          );


        // -----------------------------------------------
        // Remove if unchecked
        // -----------------------------------------------

        if (
          selectedIds.indexOf(
            quickLinkId
          ) === -1
        ) {

          await this.deleteFavourite(
            context,
            favourite.Id
          );

        }


        // -----------------------------------------------
        // Keep existing favourite
        // -----------------------------------------------

        else {

          keptIds.push(
            quickLinkId
          );

        }

      }


      // =====================================================
      // Add newly selected favourites
      // =====================================================

      for (
        let i = 0;
        i < selectedIds.length;
        i++
      ) {

        const quickLinkId =
          selectedIds[i];


        if (
          keptIds.indexOf(
            quickLinkId
          ) === -1
        ) {

          const requestBody = {

            UserId:
              currentUser.Id,

            Quick_x0020_LinkId:
              quickLinkId

          };


          const response =
            await context.spHttpClient.post(

              `${webUrl}` +
              `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items`,

              SPHttpClient.configurations.v1,

              {

                headers: {

                  'Accept':
                    'application/json;odata=nometadata',

                  'Content-Type':
                    'application/json;odata=nometadata'

                },

                body:
                  JSON.stringify(
                    requestBody
                  )

              }

            );


          if (!response.ok) {

            const errorText =
              await response.text();


            console.error(
              'Failed to add favourite:',
              response.status,
              errorText
            );

            return;

          }


          keptIds.push(
            quickLinkId
          );

        }

      }


      // =====================================================
      // Refresh the widget
      // =====================================================

      const quickLinks =
        await this.getQuickLinks(
          context
        );


      const favouriteIds =
        await this.getUserFavouriteIds(
          context
        );


      this.renderQuickLinks(
        context,
        domElement,
        quickLinks,
        favouriteIds
      );


      this.renderFavouriteOptions(
        domElement,
        quickLinks,
        favouriteIds
      );


      this.updateFavouriteTabVisibility(
        domElement,
        favouriteIds
      );

      // =====================================================
      // Close Add Favourite Modal
      // =====================================================

      const modal =
        domElement.querySelector(
          '#addFavouriteModal'
        ) as HTMLElement;


      if (modal) {

        modal.classList.remove(
          'show'
        );

        modal.style.display =
          'none';

        modal.setAttribute(
          'aria-hidden',
          'true'
        );

        document.body.classList.remove(
          'modal-open'
        );


        const backdrop =
          document.querySelector(
            '.modal-backdrop'
          );

        if (backdrop) {

          backdrop.remove();

        }

      }


    } catch (error) {

      console.error(
        'Error updating favourites:',
        error
      );

    }

  }


  // =========================================================
  // Delete Favourite
  // =========================================================

  private static async deleteFavourite(
    context: WebPartContext,
    favouriteId: number
  ): Promise<void> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const response =
      await context.spHttpClient.post(

        `${webUrl}` +
        `/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items(${favouriteId})`,

        SPHttpClient.configurations.v1,

        {

          headers: {

            'Accept':
              'application/json;odata=nometadata',

            'X-HTTP-Method':
              'DELETE',

            'IF-MATCH':
              '*'

          }

        }

      );


    if (!response.ok) {

      const errorText =
        await response.text();


      console.error(
        'Failed to delete favourite:',
        favouriteId,
        response.status,
        errorText
      );


      throw new Error(
        `Failed to delete favourite: ${response.status}`
      );

    }

  }

}