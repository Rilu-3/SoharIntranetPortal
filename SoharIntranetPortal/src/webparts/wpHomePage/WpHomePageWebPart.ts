import { Version } from '@microsoft/sp-core-library';

import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import {
  BaseClientSideWebPart,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import { SPComponentLoader } from '@microsoft/sp-loader';

import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'WpHomePageWebPartStrings';

import QuickLinks from './QuickLinks';

import Birthday from './Birthday';




// =========================================================
// Interfaces
// =========================================================

export interface IWpHomePageWebPartProps {
  description: string;
}


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


export interface IBirthdayList {
  Id: number;
  Title: string;
  EmployeePhoto?: any;
  BirthDate: string;
  Status: string;
}


// =========================================================
// Web Part
// =========================================================

export default class WpHomePageWebPart
  extends BaseClientSideWebPart<IWpHomePageWebPartProps> {


  // =========================================================
  // Render
  // =========================================================

  public async render(): Promise<void> {

    await this.loadCSS();

    const workbenchContent = document.getElementById('workbenchPageContent');
    if (workbenchContent) {
      workbenchContent.style.maxWidth = 'none';
    }
 

    this.domElement.innerHTML =
      QuickLinks.allElementsHtml +
      Birthday.allElementsHtml;

    const birthdayViewAll =
      this.domElement.querySelector(
        '.birthday-view-all'
      ) as HTMLAnchorElement;

      if (birthdayViewAll) {
        birthdayViewAll.href =
          `${this.context.pageContext.web.absoluteUrl}` +
          `/SitePages/Upcoming-Birthdys.aspx`;
      }

    const birthdayArrow =
      this.domElement.querySelector(
        '.birthday-view-all-arrow'
      ) as HTMLImageElement;

    if (birthdayArrow) {
      birthdayArrow.src =
        `${this.context.pageContext.web.absoluteUrl}` +
        `/SiteAssets/resources/images/icons/arrow-right-short.svg`;
    }
    
    await this.loadJS();


    await this.initializeQuickLinks();


    await this.renderBirthdays();

  }


  // =========================================================
  // QUICK LINKS
  // =========================================================


  // =========================================================
  // Initialize Quick Links
  // =========================================================

  private async initializeQuickLinks(): Promise<void> {

    try {

      const quickLinks =
        await this.getQuickLinks(
          this.context
        );


      const favouriteIds =
        await this.getUserFavouriteIds(
          this.context
        );


      this.renderQuickLinks(
        this.context,
        this.domElement,
        quickLinks,
        favouriteIds
      );


      this.renderFavouriteOptions(
        this.domElement,
        quickLinks,
        favouriteIds
      );


      this.setupTabs(
        this.domElement
      );


      this.setupAddFavouriteButton(
        this.context,
        this.domElement
      );


      this.setupFavouriteModal(
        this.context,
        this.domElement
      );


      this.updateFavouriteTabVisibility(
        this.domElement,
        favouriteIds
      );

    } catch (error) {

      console.error(
        'Error initializing Quick Links:',
        error
      );


      const container =
        this.domElement.querySelector(
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
  // Get Quick Links
  // =========================================================

  private async getQuickLinks(
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

  private async getCurrentUser(
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

  private async getUserFavouriteIds(
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

  private getQuickLinkImageUrl(
    context: WebPartContext,
    item: IQuickLinksList
  ): string {

    if (!item.Icon) {

      return '';

    }


    try {

      const imgData =
        JSON.parse(
          item.Icon
        );


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

  private renderQuickLinks(
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
          this.getQuickLinkImageUrl(
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
          QuickLinks.singleElementHtml

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
  // Render Favourite Options
  // =========================================================

  private renderFavouriteOptions(
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

  private setupTabs(
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

                    (
                      link as HTMLElement
                    ).style.display = '';

                  } else {

                    (
                      link as HTMLElement
                    ).style.display = 'none';

                  }

                } else {

                  (
                    link as HTMLElement
                  ).style.display = '';

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

  private updateFavouriteTabVisibility(
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

  private setupAddFavouriteButton(
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

  private setupFavouriteModal(
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
              Number(
                input.value
              );


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

  private async addFavourites(
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
              Number(
                input.value
              )
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
      // Get Existing Favourites
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
      // Remove Unchecked Favourites
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


        if (
          selectedIds.indexOf(
            quickLinkId
          ) === -1
        ) {

          await this.deleteFavourite(
            context,
            favourite.Id
          );

        } else {

          keptIds.push(
            quickLinkId
          );

        }

      }


      // =====================================================
      // Add Newly Selected Favourites
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

      const activeTab =
        domElement.querySelector(
          '.panel-title-filter-active'
        ) as HTMLElement;

      const activeTabValue =
        activeTab?.getAttribute(
          'data-filter-ql'
        );



      // =====================================================
      // Refresh Quick Links
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

          if (activeTabValue) {

      const tabToRestore =
        domElement.querySelector(
          `[data-filter-ql="${activeTabValue}"]`
        ) as HTMLElement;

      if (tabToRestore) {
        tabToRestore.click();
      }

    }


      // =====================================================
      // Close Modal
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

  private async deleteFavourite(
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


  // =========================================================
  // BIRTHDAYS
  // =========================================================


  // =========================================================
  // Get Birthdays
  // =========================================================

  private async getBirthdays(
    context: WebPartContext
  ): Promise<IBirthdayList[]> {

    const webUrl =
      context.pageContext.web.absoluteUrl;


    const endpoint =
      `${webUrl}/_api/web/lists/getbytitle('Birthday')/items` +
      `?$select=Id,Title,EmployeePhoto,BirthDate,Status`;


    try {

      const response =
        await context.spHttpClient.get(
          endpoint,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept:
                'application/json;odata=nometadata'
            }
          }
        );


      if (!response.ok) {

        console.error(
          'Failed to get birthday data:',
          response.status,
          response.statusText
        );

        return [];

      }


      const data =
        await response.json();


      return data.value;

    } catch (error) {

      console.error(
        'Error getting birthday data:',
        error
      );

      return [];

    }

  }


  // =========================================================
  // Get Employee Photo URL
  // =========================================================

  private getBirthdayImageUrl(
    context: WebPartContext,
    photo: any,
    itemId: number
  ): string {

    if (!photo) {

      return '';

    }


    try {

      if (typeof photo === 'string') {

        photo =
          JSON.parse(
            photo
          );

      }


      if (!photo.fileName) {

        return '';

      }


      return (
        context.pageContext.web.absoluteUrl +
        '/Lists/Birthday/Attachments/' +
        itemId +
        '/' +
        encodeURIComponent(
          photo.fileName
        )
      );

    } catch (error) {

      console.error(
        'Error parsing employee photo:',
        error
      );

      return '';

    }

  }


  // =========================================================
  // Get Upcoming Birthdays
  // =========================================================

  private getUpcomingBirthdays(
    birthdays: IBirthdayList[]
  ): IBirthdayList[] {

    const today =
      new Date();


    today.setHours(
      0,
      0,
      0,
      0
    );


    const upcoming =
      birthdays

        .filter(
          (item) =>
            item.Status === 'Active'
        )

        .map(
          (item) => {

            const birthDate =
              new Date(
                item.BirthDate
              );


            let birthday =
              new Date(
                today.getFullYear(),
                birthDate.getMonth(),
                birthDate.getDate()
              );


            if (
              birthday < today
            ) {

              birthday.setFullYear(
                today.getFullYear() + 1
              );

            }


            const difference =
              Math.floor(
                (
                  birthday.getTime() -
                  today.getTime()
                ) /
                (
                  1000 *
                  60 *
                  60 *
                  24
                )
              );


            return {
              item,
              difference
            };

          }
        )

        .filter(
          (item) =>
            item.difference < 7
        )

        .sort(
          (a, b) =>
            a.difference -
            b.difference
        );


    return upcoming.map(
      (item) =>
        item.item
    );

  }


  // =========================================================
  // Render Birthdays
  // =========================================================

  private async renderBirthdays(): Promise<void> {
    

    const birthdays =
      await this.getBirthdays(
        this.context
      );


    const upcomingBirthdays =
      this.getUpcomingBirthdays(
        birthdays
      );


    const container =
      this.domElement.querySelector(
        '.panel-card-birthdays'
      );


    if (!container) {

      console.error(
        'Birthday container not found.'
      );

      return;

    }


    let html = '';


    upcomingBirthdays.forEach(
      (item) => {

        const birthDate =
          new Date(
            item.BirthDate
          );


        const month =
          birthDate
            .toLocaleString(
              'en-US',
              {
                month: 'short'
              }
            )
            .toUpperCase();


        const day =
          birthDate
            .getDate()
            .toString();


        const imageUrl =
          this.getBirthdayImageUrl(
            this.context,
            item.EmployeePhoto,
            item.Id
          );


        html +=
          Birthday.singleElementHtml

            .replace(
              /__KEY_URL_IMG__/g,
              imageUrl
            )

            .replace(
              /__KEY_TITLE__/g,
              item.Title
            )

            .replace(
              /__KEY_MONTH__/g,
              month
            )

            .replace(
              /__KEY_DAY__/g,
              day
            );

      }
    );


    container.innerHTML =
      html;

  }


private async loadCSS(): Promise<void> {
  const baseUrl =
    this.context.pageContext.web.absoluteUrl;

  const cssFiles = [
    'bootstrap.min.css',
    'custom.css',
    'font-size.css',
    'home.css',
    'jquery-ui.css',
    'sp-custom.css',
    'swiper-bundle.min.css',
    'variable.css'
  ];

  for (const file of cssFiles) {
    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/${file}`
    );
  }
}


  // =========================================================
  // LOAD JS
  // =========================================================

  private async loadJS(): Promise<void> {

    const baseUrl =
      this.context.pageContext.web.absoluteUrl;


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/common.js`
    );


    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/home.js`
    );

  }


  // =========================================================
  // DATA VERSION
  // =========================================================

  protected get dataVersion(): Version {

    return Version.parse(
      '1.0'
    );

  }


  // =========================================================
  // PROPERTY PANE
  // =========================================================

  protected getPropertyPaneConfiguration():
    IPropertyPaneConfiguration {

    return {

      pages: [

        {

          header: {

            description:
              strings.PropertyPaneDescription

          },

          groups: [

            {

              groupName:
                strings.BasicGroupName,

              groupFields: [

                PropertyPaneTextField(
                  'description',
                  {

                    label:
                      strings.DescriptionFieldLabel

                  }
                )

              ]

            }

          ]

        }

      ]

    };

  }

}