export default class Wrapper {

    public static wrapperHtml: string = `

<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1" />
  <title>Sohar Aluminum - Home</title>
  
</head>

<body>
  

  <div class="main-wrapper w-100 float-start min-h-screen-wrapper">
    <div class="container container-sa px-3 px-lg-4 py-3 mx-auto">
      <div class="row gy-3 gx-3 clearfix">
        
      
      <div class="col-12" id="banner-container">

        </div>



        <div class="col-12 col-lg-6" id="announcement-offer-container">
       
        </div>

        <div class="col-12 col-lg-6" id="quick-links-container">

        </div>





        <div id="news-container" class="col-12 col-lg-6">
          
        </div>



        <div class="col-12 col-lg-6" id="upcoming-events-container">
          
        </div>





        <div class="col-12 col-lg-8" id="social-media-container">
         
        </div>





        <div class="col-12 col-lg-4" id="birthday-container">
          
        </div>


        <div class="col-12" id="media-gallery-container">

        </div>


      </div>
    </div>
  </div>

 
  <div class="modal fade" id="galleryModal" tabindex="-1" aria-labelledby="galleryModal" aria-hidden="true">
    <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable gallery-modal">
      <div class="modal-content py-2 px-2">
        <div class="gallery-modal-header d-flex justify-content-end me-2 mt-2">
          <svg class="cursor-pointer flex-shrink-0" data-bs-dismiss="modal" aria-label="Close" width="15" height="15"
            viewBox="0 0 15 15" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016" stroke="#005187"
              stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round" />
          </svg>
        </div>
        <div class="modal-body gallery-modal-content">
          <!-- <img id="galleryModalImg" src="./resources/images/gallery/gallery-1-lg.png" alt="" /> -->
          <!-- <video controls>
            <source src="" type="video/mp4">
            Your browser does not support the video tag.
          </video> -->
          <div class="w-100 float-start position-relative gallery-swiper-wrapper">
            <div class="w-100 float-start gallery-modal-swiper swiper">
              <div class="swiper-wrapper">
                <!-- <div class="swiper-slide gallery-swiper-slide">
                  <video controls>
                    <source src="./resources/images/gallery/istockphoto-1253263447-640_adpp_is 1.mp4" type="video/mp4">
                    Your browser does not support the video tag.
                  </video>
                </div> -->
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-1.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-2.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-3.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-1.jfif" alt="" />
                </div>
                <div class="swiper-slide gallery-swiper-slide">
                  <img src="./resources/images/gallery/gallery-2.jfif" alt="" />
                </div>
              </div>
            </div>
            <div class="swiper-button-next swiper-btn-common gallery-swiper-modal-next"></div>
            <div class="swiper-button-prev swiper-btn-common gallery-swiper-modal-prev"></div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="modal fade favourite-modal" id="addFavouriteModal" tabindex="-1" aria-labelledby="addFavouriteModalLabel"
    aria-hidden="true">

    <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable custom-scroll-view">

      <div class="modal-content">

        <div class="modal-header">
          <h5 class="modal-title" id="addFavouriteModalLabel">
            Add to Favourites
          </h5>

          <svg class="cursor-pointer flex-shrink-0" data-bs-dismiss="modal" aria-label="Close" width="15" height="15"
            viewBox="0 0 15 15" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M0.666992 0.666016L14.0003 13.9993M0.666992 13.9993L14.0003 0.666016" stroke="#005187"
              stroke-width="1.33333" stroke-linecap="round" stroke-linejoin="round"></path>
          </svg>
        </div>

        <div class="modal-body">

          <p class="mb-3 modal-body-fav-title">
            Select the quick links you want to add to your favourites.
          </p>

          <div class="favourite-options">

            <label class="favourite-option">
              <input type="checkbox" value="Payslip" data-favourite="ql-favourite">
              <span>Payslip</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Leave Application" data-favourite="ql-favourite">
              <span>Leave Application</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Travel Request" data-favourite="ql-favourite">
              <span>Travel Request</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Expense Claim" data-favourite="ql-favourite">
              <span>Expense Claim</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Training Portal" data-favourite="ql-favourite">
              <span>Training Portal</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="IT Help Desk" data-favourite="ql-favourite">
              <span>IT Help Desk</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Safety Portal" data-favourite="ql-favourite">
              <span>Safety Portal</span>
            </label>

            <label class="favourite-option">
              <input type="checkbox" value="Idea Box" data-favourite="ql-favourite">
              <span>Idea Box</span>
            </label>

          </div>

        </div>

        <div class="modal-footer">
          <button type="button" class="btn-appearance-none btn-brand btn-brand-secondary px-3 py-1 text-sm"
            data-bs-dismiss="modal">
            Cancel
          </button>

          <button type="button" class="btn-appearance-none btn-brand btn-brand-primary px-3 py-1 text-sm" id="btnAddFavourites">
            Add to Favourites
          </button>
        </div>

      </div>

    </div>
  </div>

  
</body>

</html>
 `}