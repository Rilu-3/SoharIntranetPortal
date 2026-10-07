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

import SocialMedia from './SocialMedia';






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
// SociableKIT Branding Link Handler
// =========================================================

private socialMediaObserver: MutationObserver | null = null;
private currentUserId: number | null = null;
private quickLinksCache: IQuickLinksList[] = [];
private favouritesCache: { Id: number; LinkId: number }[] = [];

private hideSocialMediaTutorialLinks(
  root: Document | ShadowRoot = document
): void {
  root.querySelectorAll<HTMLElement>('a.tutorial_link').forEach((link) => {
    link.style.setProperty('display', 'none', 'important');
  });

  root.querySelectorAll<HTMLElement>('*').forEach((element) => {
    if (element.shadowRoot) {
      this.hideSocialMediaTutorialLinks(element.shadowRoot);
    }
  });
}

private setupSocialMediaTutorialLinkObserver(): void {
  this.socialMediaObserver?.disconnect();

  this.hideSocialMediaTutorialLinks();

  this.socialMediaObserver = new MutationObserver(() => {
    this.hideSocialMediaTutorialLinks();
  });

  this.socialMediaObserver.observe(document.body, {
    childList: true,
    subtree: true
  });
}
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
      Birthday.allElementsHtml +
      SocialMedia.allElementsHtml;

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

    await this.initializeSocialMedia();

  }


  // =========================================================
  // QUICK LINKS
  // =========================================================


  // =========================================================
  // Initialize Quick Links
  // =========================================================

private async initializeQuickLinks(): Promise<void> {
  try {
    const [quickLinks] = await Promise.all([
      this.getQuickLinks(this.context),
      this.loadFavourites(this.context)
    ]);

    this.quickLinksCache = quickLinks;
    const favouriteIds = this.getFavouriteIds();

    this.renderQuickLinks(this.context, this.domElement, quickLinks, favouriteIds);
    this.renderFavouriteOptions(this.domElement, quickLinks, favouriteIds);
    this.setupTabs(this.domElement);
    this.setupAddFavouriteButton(this.context, this.domElement);
    this.setupFavouriteModal(this.context, this.domElement);
    this.updateFavouriteTabVisibility(this.domElement, favouriteIds);
  } catch (error) {
    console.error('Error initializing Quick Links:', error);
    const container = this.domElement.querySelector('.quick-links-grid');
    if (container) {
      container.innerHTML = `<div class="p-3">Unable to load Quick Links.</div>`;
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

private async getCurrentUserId(context: WebPartContext): Promise<number> {
  if (this.currentUserId !== null) {
    return this.currentUserId;
  }
  const user = await this.getCurrentUser(context);
  this.currentUserId = user.Id;
  return user.Id;
}

// Fetches from the server once and fills the cache
private async loadFavourites(context: WebPartContext): Promise<void> {
  try {
    const webUrl = context.pageContext.web.absoluteUrl;
    const userId = await this.getCurrentUserId(context);

    const response = await context.spHttpClient.get(
      `${webUrl}/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items` +
      `?$select=Id,Quick_x0020_LinkId&$filter=UserId eq ${userId}`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      console.error('Unable to get user favourites:', response.status);
      this.favouritesCache = [];
      return;
    }

    const data = await response.json();
    this.favouritesCache = data.value.map(
      (item: { Id: number; Quick_x0020_LinkId: number }) => ({
        Id: item.Id,
        LinkId: Number(item.Quick_x0020_LinkId)
      })
    );
  } catch (error) {
    console.error('Error loading user favourites:', error);
    this.favouritesCache = [];
  }
}

private getFavouriteIds(): number[] {
  return this.favouritesCache.map((f) => f.LinkId);
}

// Returns the new list item's Id
private async createFavourite(
  context: WebPartContext,
  userId: number,
  quickLinkId: number
): Promise<number> {
  const webUrl = context.pageContext.web.absoluteUrl;

  const response = await context.spHttpClient.post(
    `${webUrl}/_api/web/lists/GetByTitle('Quick_Link_Favourites')/items`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata'
      },
      body: JSON.stringify({
        UserId: userId,
        Quick_x0020_LinkId: quickLinkId
      })
    }
  );

  if (!response.ok) {
    throw new Error(`Failed to add favourite: ${response.status}`);
  }

  const created = await response.json();
  return created.Id;
}


  // =========================================================
  // Get Quick Link Image URL
  // =========================================================

 private getQuickLinkImageUrl(
  item: IQuickLinksList
): string {

  if (!item.Icon) {
    return '';
  }

  try {

    const imgData =
      typeof item.Icon === 'string'
        ? JSON.parse(item.Icon)
        : item.Icon;

    return imgData.serverRelativeUrl || '';

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
          this.getQuickLinkImageUrl(item);


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
 
  const favouriteTab = domElement.querySelector(
    '[data-filter-ql="favourites"]'
  ) as HTMLElement;
 
  if (!favouriteTab) {
    return;
  }
 
  const hasFavourites = favouriteIds.length > 0;
 
  favouriteTab.classList.toggle('d-flex', hasFavourites);
  favouriteTab.classList.toggle('d-none', !hasFavourites);
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
  const modal = domElement.querySelector('#addFavouriteModal');
  if (!modal) {
    console.error('Favourite modal not found.');
    return;
  }

  modal.addEventListener('show.bs.modal', () => {
    const favouriteIds = this.getFavouriteIds();

    domElement
      .querySelectorAll('.favourite-options input[type="checkbox"]')
      .forEach((checkbox) => {
        const input = checkbox as HTMLInputElement;
        input.checked = favouriteIds.indexOf(Number(input.value)) !== -1;
      });
  });
}


  // =========================================================
  // Add / Remove Favourites
  // =========================================================

private async addFavourites(
  context: WebPartContext,
  domElement: HTMLElement
): Promise<void> {

  // 1. Read selection
  const selectedIds: number[] = [];
  domElement
    .querySelectorAll('.favourite-options input[type="checkbox"]')
    .forEach((checkbox) => {
      const input = checkbox as HTMLInputElement;
      if (input.checked) {
        selectedIds.push(Number(input.value));
      }
    });

  // 2. Close modal immediately
  const modal = domElement.querySelector('#addFavouriteModal') as HTMLElement;
  if (modal) {
    const bootstrap = (window as any).bootstrap;
    if (bootstrap?.Modal) {
      bootstrap.Modal.getOrCreateInstance(modal).hide();
    } else {
      modal.classList.remove('show');
      modal.style.display = 'none';
      modal.setAttribute('aria-hidden', 'true');
      document.body.classList.remove('modal-open');
      document.querySelectorAll('.modal-backdrop').forEach((b) => b.remove());
    }
  }

  // 3. Work out what changed using the cache (no network)
  const toDelete = this.favouritesCache.filter(
    (f) => selectedIds.indexOf(f.LinkId) === -1
  );
  const existingLinkIds = this.favouritesCache.map((f) => f.LinkId);
  const toAdd = selectedIds.filter(
    (id) => existingLinkIds.indexOf(id) === -1
  );

  if (toDelete.length === 0 && toAdd.length === 0) {
    return;
  }

  // 4. Update the UI right away (optimistic)
  const activeTabValue = (
    domElement.querySelector('.panel-title-filter-active') as HTMLElement
  )?.getAttribute('data-filter-ql');

  this.refreshFavouritesUI(domElement, selectedIds, activeTabValue);

  // 5. Sync with SharePoint in parallel
  try {
    const userId = await this.getCurrentUserId(context);

    const deletePromises = toDelete.map((f) =>
      this.deleteFavourite(context, f.Id)
    );

    const addPromises = toAdd.map(async (linkId) => ({
      Id: await this.createFavourite(context, userId, linkId),
      LinkId: linkId
    }));

    const [, created] = await Promise.all([
      Promise.all(deletePromises),
      Promise.all(addPromises)
    ]);

    // 6. Update cache with the new server IDs
    const deletedIds = toDelete.map((f) => f.Id);
    this.favouritesCache = this.favouritesCache
      .filter((f) => deletedIds.indexOf(f.Id) === -1)
      .concat(created);

  } catch (error) {
    console.error('Error updating favourites:', error);

    // Roll back to the real server state
    await this.loadFavourites(context);
    this.refreshFavouritesUI(
      domElement,
      this.getFavouriteIds(),
      activeTabValue
    );
  }
}

private refreshFavouritesUI(
  domElement: HTMLElement,
  favouriteIds: number[],
  activeTabValue?: string | null
): void {
  this.renderQuickLinks(this.context, domElement, this.quickLinksCache, favouriteIds);
  this.renderFavouriteOptions(domElement, this.quickLinksCache, favouriteIds);
  this.updateFavouriteTabVisibility(domElement, favouriteIds);

  if (favouriteIds.length === 0) {
    (domElement.querySelector('[data-filter-ql="quick-links"]') as HTMLElement)?.click();
  } else if (activeTabValue) {
    (domElement.querySelector(`[data-filter-ql="${activeTabValue}"]`) as HTMLElement)?.click();
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

private getBirthdayImageUrl(photo: any): string {
  if (!photo) {
    return '';
  }

  try {
    if (typeof photo === 'string') {
      photo = JSON.parse(photo);
    }

    return photo.serverRelativeUrl || '';

  } catch (error) {
    console.error('Error parsing employee photo:', error);
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


        const imageUrl = this.getBirthdayImageUrl(
          item.EmployeePhoto
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

  // =====================================================
// HR API - Birthdays
// =====================================================

// private async getHRBirthdays(): Promise<any[]> {

//   const apiUrl = 'HR_API_URL_WILL_BE_PROVIDED';

//   // API implementation will be added
//   // once HR provides endpoint and authentication details.

//   return [];
// }

// =========================================================
// SOCIAL MEDIA
// =========================================================

private async initializeSocialMedia(): Promise<void> {

  // Instagram
  const instagramContainer =
    this.domElement.querySelector(
      '#social-instagram'
    ) as HTMLElement;

  if (instagramContainer) {
    instagramContainer.innerHTML = `
      <div
        class="sk-instagram-feed"
        data-embed-id="25716225">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/instagram-feed/widget.js';
    script.defer = true;

    instagramContainer.appendChild(script);
  }


  // X (Twitter)
  const twitterContainer =
    this.domElement.querySelector(
      '#social-twitter'
    ) as HTMLElement;

  if (twitterContainer) {
    twitterContainer.innerHTML = `
      <div
        class="sk-ww-twitter-feed"
        data-embed-id="25716230">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/twitter-feed/widget.js';
    script.defer = true;

    twitterContainer.appendChild(script);
  }


  // LinkedIn
  const linkedinContainer =
    this.domElement.querySelector(
      '#social-linkedin'
    ) as HTMLElement;

  if (linkedinContainer) {
    linkedinContainer.innerHTML = `
      <div
        class="sk-ww-linkedin-page-post"
        data-embed-id="25716233">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/linkedin-page-posts/widget.js';
    script.defer = true;

    linkedinContainer.appendChild(script);
  }


  // Facebook
  const facebookContainer =
    this.domElement.querySelector(
      '#social-facebook'
    ) as HTMLElement;

  if (facebookContainer) {
    facebookContainer.innerHTML = `
      <div
        class="sk-ww-facebook-page-posts"
        data-embed-id="25716238">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/facebook-page-posts/widget.js';
    script.defer = true;

    facebookContainer.appendChild(script);
  }


  // YouTube
  const youtubeContainer =
    this.domElement.querySelector(
      '#social-youtube'
    ) as HTMLElement;

  if (youtubeContainer) {
    youtubeContainer.innerHTML = `
      <div
        class="sk-ww-youtube-channel-videos"
        data-embed-id="25716248">
      </div>
    `;

    const script = document.createElement('script');
    script.src =
      'https://widgets.sociablekit.com/youtube-channel-videos/widget.js';
    script.defer = true;

    youtubeContainer.appendChild(script);
  }

  this.setupSocialMediaTutorialLinkObserver();
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