import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
 
import adminPageHtml, {
  IAdminListView
} from './adminPageHtml';
 
interface ISPList {
  Title: string;
  Hidden: boolean;
  BaseTemplate: number;
  RootFolder: {
    ServerRelativeUrl: string;
  };
}
 
interface ISPListsResponse {
  value: ISPList[];
}
 
export default class AdminPageWebPart
  extends BaseClientSideWebPart<Record<string, never>> {
 
  public async onInit(): Promise<void> {
    const baseUrl: string = this.context.pageContext.web.absoluteUrl;
    SPComponentLoader.loadCss(`${baseUrl}/SiteAssets/resources/css/list.css`);
  }
 
  public async render(): Promise<void> {
    this.domElement.innerHTML = adminPageHtml.wrapperHtml(
      adminPageHtml.loadingHtml
    );
 
    try {
      const lists = await this.getBackendLists();
 
      const content = lists.length
        ? lists
            .map(list => this.toListView(list))
            .map(item => adminPageHtml.singleListHtml(item))
            .join('')
        : adminPageHtml.noRecordHtml;
 
      this.domElement.innerHTML = adminPageHtml.wrapperHtml(content);
 
    } catch (error) {
      console.error('Error loading SharePoint lists:', error);
 
      this.domElement.innerHTML = adminPageHtml.wrapperHtml(
        adminPageHtml.errorHtml
      );
    } finally {
     
      const listCss = document.querySelector('link[href*="list.css"]');
      if (listCss) {
        document.head.appendChild(listCss);
      }
    }
  }
 
  private async getBackendLists(): Promise<ISPList[]> {
    const webUrl = this.context.pageContext.web.absoluteUrl;
 
    const endpoint =
    `${webUrl}/_api/web/lists` +
    `?$select=Id,Title,Hidden,BaseTemplate,` +
    `RootFolder/ServerRelativeUrl` +
    `&$expand=RootFolder` +
    `&$filter=Hidden eq false&$orderby=Title`;
 
    const response = await this.context.spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
 
    if (!response.ok) {
      throw new Error(
        `SharePoint API error: ${response.status}`
      );
    }
 
    const data: ISPListsResponse = await response.json();
 
    // Exclude document libraries; retain ordinary SharePoint lists.
return data.value.filter(
  list =>
    list.BaseTemplate !== 101 &&
    list.Title !== 'Site Pages'
);
  }
 
 
private toListView(list: ISPList): IAdminListView {
  const webUrl = this.context.pageContext.web.absoluteUrl;
 
  const webPath =
    this.context.pageContext.web.serverRelativeUrl
      .replace(/\/$/, '');
 
  const listPath = list.RootFolder.ServerRelativeUrl;
 
  const relativePath =
    listPath.indexOf(webPath + '/') === 0
      ? listPath.substring(webPath.length)
      : listPath;
 
 
  const listUrl =
    `${webUrl.replace(/\/$/, '')}${relativePath}`;
 
    return {
      title: list.Title,
      url: listUrl
    };
}
 
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}