export interface IMediaCardView {
  title: string;
  caption: string;
  imageUrl: string;
}

export default class wpMediaListing {

  // Page wrapper (header + year filter + listing container)
  public static wrapperHtml(resourceUrl: string, homeUrl: string, content: string): string {
    return `
      <div class="main-wrapper w-100 float-start min-h-screen-wrapper">
        <div class="container container-sa px-3 px-lg-4 py-3 mx-auto">
          <div class="w-100 float-start announcement-page">

            <div class="w-100 float-start pt-3 pb-5">
              <div class="d-flex justify-content-between w-100 float-start flex-wrap gap-2">
                <div class="float-start d-flex flex-column gap-2">
                  <div class="breadcrumb-container d-flex">
                    <ol class="breadcrumb">
                      <li class="breadcrumb-item">
                        <a href="${homeUrl}" target="_self" data-interception="off">
                          <svg width="12" height="13" viewBox="0 0 12 13" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <path
                              d="M5.03333 0.353023C5.30384 0.125038 5.64623 0 6 0C6.35377 0 6.69616 0.125038 6.96667 0.353023L11.4667 4.14769C11.8047 4.43302 12 4.85236 12 5.29436V11.501C12 11.8104 11.8771 12.1072 11.6583 12.326C11.4395 12.5448 11.1428 12.6677 10.8333 12.6677H8.5C8.19058 12.6677 7.89383 12.5448 7.67504 12.326C7.45625 12.1072 7.33333 11.8104 7.33333 11.501V8.16302C7.33333 8.11882 7.31577 8.07643 7.28452 8.04517C7.25326 8.01392 7.21087 7.99636 7.16667 7.99636H4.83333C4.78913 7.99636 4.74674 8.01392 4.71548 8.04517C4.68423 8.07643 4.66667 8.11882 4.66667 8.16302V11.501C4.66667 11.8104 4.54375 12.1072 4.32496 12.326C4.10617 12.5448 3.80942 12.6677 3.5 12.6677H1.16667C1.01346 12.6677 0.861749 12.6375 0.720203 12.5789C0.578656 12.5203 0.450044 12.4343 0.341709 12.326C0.233374 12.2176 0.147438 12.089 0.0888073 11.9475C0.0301768 11.8059 0 11.6542 0 11.501V5.29436C0 4.85236 0.195333 4.43302 0.533333 4.14769L5.03333 0.353023ZM6.322 1.11769C6.23186 1.04181 6.11782 1.0002 6 1.0002C5.88218 1.0002 5.76814 1.04181 5.678 1.11769L1.178 4.91169C1.12231 4.95857 1.07753 5.01704 1.04679 5.08302C1.01604 5.149 1.00007 5.2209 1 5.29369V11.5004C1 11.5924 1.07467 11.667 1.16667 11.667H3.5C3.5442 11.667 3.5866 11.6495 3.61785 11.6182C3.64911 11.587 3.66667 11.5446 3.66667 11.5004V8.16236C3.66667 7.51769 4.18933 6.99569 4.83333 6.99569H7.16667C7.81067 6.99569 8.33333 7.51769 8.33333 8.16236V11.5004C8.33333 11.5924 8.408 11.667 8.5 11.667H10.8333C10.8775 11.667 10.9199 11.6495 10.9512 11.6182C10.9824 11.587 11 11.5446 11 11.5004V5.29369C10.9999 5.2209 10.984 5.149 10.9532 5.08302C10.9225 5.01704 10.8777 4.95857 10.822 4.91169L6.322 1.11769Z"
                              fill="var(--color-secondary)" />
                          </svg>
                          <span>Home</span>
                        </a>
                      </li>
                      <li class="breadcrumb-item" aria-current="page">Listing</li>
                    </ol>
                  </div>
                  <p class="font-bold font-bold text-lg md-text-xl text-color-primary">MEDIA GALLERY</p>
                </div>
                <div class="d-flex align-items-center gap-2 flex-shrink-0 filter-area-wrapper">
                  <div class="dropdown sa-filter-dropdown">
                    <button
                      class="sa-btn sa-filter-dropdown-btn dropdown-toggle d-flex align-items-center gap-1 text-color-neutral text-sm font-normal"
                      type="button" id="dropdownMenuYear" data-bs-toggle="dropdown" aria-expanded="false">
                      <div class="d-flex align-items-center gap-1 flex-shrink-0">
                        <img src="${resourceUrl}/images/icons/filter.png" alt="Filter" />
                        <span>Year:</span>
                      </div>
                      <div class="flex-grow-1 overflow-hidden">
                        <p class="text-truncate text-start">All</p>
                      </div>
                    </button>
                    <ul class="dropdown-menu text-sm font-normal" aria-labelledby="dropdownMenuYear">
                      ${wpMediaListing.yearItemHtml("All")}
                    </ul>
                  </div>
                </div>
              </div>
              <div class="row gy-4" id="IdMediaList">
                ${content}
              </div>
            </div>

          </div>
        </div>
      </div>`;
  }

  // Single year dropdown item
  public static yearItemHtml(year: string): string {
    return `<li><span class="dropdown-item" data-year="${year}">${year}</span></li>`;
  }

  // Loading state (Bootstrap spinner)
  public static loadingHtml: string = `
    <div class="col-12 d-flex justify-content-center py-5">
      <div class="spinner-border text-secondary" role="status">
        <span class="visually-hidden">Loading...</span>
      </div>
    </div>`;

  // No records
  public static noRecordHtml: string = `
    <div class="col-12">
      <p class="text-color-desc text-sm">No records found</p>
    </div>`;

  // Single media card
  public static cardHtml(data: IMediaCardView): string {
    return `
      <div class="col-12 col-sm-6 col-md-4 col-xl-3">
        <a href="${data.imageUrl}" target="_self" rel="noopener noreferrer" class="w-100 float-start d-flex flex-column sa-list-col" data-interception="off">
          <div class="sa-list-media w-100 float-start">
            <img src="${data.imageUrl}" alt="${data.title}" />
          </div>
          <div class="w-100 d-flex flex-column p-3 gap-2 sa-list-content">
            <h5 class="line-clamp-3">${data.title}</h5>
            <p class="line-clamp-3">${data.caption}</p>
          </div>
        </a>
      </div>`;
  }
}