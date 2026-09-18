
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import UabHomePage from './UabHomePage';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import UabAnnouncements from './UabAnnouncements';

export interface IWpHomePageWebPartProps {
  description: string;
}
interface IAnnouncement {
Id: number;
Title: string;
ShortDescription: string;
Icon: string;
CreatedDate: string;
}
export default class WpHomePageWebPart extends BaseClientSideWebPart<IWpHomePageWebPartProps> {
  public async onInit(): Promise<void> {
    await this.loadCSS();
  }
  public async render(): Promise<void> {
    this.domElement.innerHTML = UabHomePage.allElementsHtml;
        const AnnouncementApiUrl=`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Announcements')/items?$select=Id,Title,ShortDescription,Icon,CreatedDate,Status&$filter=Status eq 'Active'`;
        await this._renderAnnouncementsAsync(AnnouncementApiUrl);
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }
  private async loadCSS(): Promise<void> {

    const baseUrl = this.context.pageContext.web.absoluteUrl;
  
    await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/bootstrap.min.css`
    );

     await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/custom.css`
    );

     await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/font-size.css`
    );

     await  SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/home.css`
    );
 
     await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/jquery-ui.css`
    );

     await  SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/sp-custom.css`
    );

     await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/swiper-bundle.min.css`
    );

     await SPComponentLoader.loadCss(
      `${baseUrl}/SiteAssets/resources/css/variable.css`
    );
   console.log(`${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`);
   
   console.log(`${baseUrl}/SiteAssets/resources/js/swiper-bundle.min.js`);
   
   console.log(`${baseUrl}/SiteAssets/resources/js/home.js`);
   
   console.log(`${baseUrl}/SiteAssets/resources/js/common.js`);
   
   console.log(`${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`);
   
   console.log(`${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`);
      await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-3.6.0.js`
    );
                    await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/bootstrap.bundle.min.js`
    );
          await SPComponentLoader.loadScript(
      `${baseUrl}/SiteAssets/resources/js/jquery-ui.js`
    );
    //       await SPComponentLoader.loadScript(
    //   `${baseUrl}/SiteAssets/resources/js/jquery.marquee.min.js`
    // );
    
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
  private async _renderAnnouncementsAsync(apiUrl: string): Promise<void> {
      const data:IAnnouncement[]=await this._getAnnouncementsData(apiUrl) ;
        let allElementsHtml: string = "";
  try{
data.forEach((item,index)=>{
    let imageUrl = '';
  if (item.Icon) {
  const imageData = JSON.parse(item.Icon); // Assuming the Icon field contains JSON data
  // console.log(imageData);
  const fileName = imageData.fileName;//extracting file name from background image

  // Build the image URL using the fileName
   imageUrl = `${this.context.pageContext.web.absoluteUrl}/Lists/Announcements/Attachments/${item.Id}/${fileName}`;
    // console.log(imageUrl);
   
  }
  let createddate=this.formatDate(item.CreatedDate);
          let singleElementHtml = UabAnnouncements.singleElementHtml
          .replace("__KEY__ANNOUNCEMENT__ICON__",imageUrl)
          .replace("__KEY__ANNOUNCEMENT__TITLE__", item.Title)
          .replace("__KEY__ANNOUNCEMENT__DESCRIPTION__", item.ShortDescription).replace("__KEY__ANNOUNCEMENT__DATE__", createddate);
         allElementsHtml += singleElementHtml;
})
  }
catch(error){
      console.error('Error rendering QuickList:', error); 
}
this.domElement.querySelector("#announcement")!.innerHTML=allElementsHtml;

}
    private async _getAnnouncementsData(apiUrl: string): Promise<IAnnouncement[]> { 
    try{
      const response:SPHttpClientResponse=await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
      
        if(response.ok){
           const data=await response.json();
           return data.value;
         
          
        } else {
          console.error(`Request failed with status ${response.status}: ${response.statusText}`);
          throw new Error(`Request failed with status ${response.status}: ${response.statusText}`);
        }
      
    } catch(error){
      console.log("error occured",error);
      throw error;
    }
  }
    private formatDate(
    dateValue: string
  ): string {
 
    if (!dateValue) {
      return '';
    }
 
    const date =
      new Date(dateValue);
 
    if (
      isNaN(
        date.getTime()
      )
    ) {
 
      return dateValue;
    }
 
    return date.toLocaleDateString(
      'en-US',
      {
        month: 'short',
        day: 'numeric',
        year: 'numeric'
      }
    );
  }
}