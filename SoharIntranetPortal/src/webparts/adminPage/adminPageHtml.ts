export interface IAdminListView {
  title: string;
  url: string;
}
 
export default class adminPageHtml {
 
  private static escape(value: string): string {
    return (value || '').replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    } as { [key: string]: string })[c]);
  }
 
  public static wrapperHtml(content: string): string {
    return `
      <div class="container-fluid px-0 w-100">
        <div class="panel-card p-3 w-100">
          <div class="panel-header mb-3">
            <h2 class="panel-title">Lists</h2>
          </div>
          <div class="event-meta fw-bold pb-2 w-100">Name</div>
          <div class="w-100">
            ${content}
          </div>
        </div>
      </div>`;
  }
 
  public static loadingHtml: string = `<div class="event-meta">Loading lists...</div>`;
 
  public static noRecordHtml: string = `<div class="event-meta">No SharePoint lists found.</div>`;
 
  public static errorHtml: string = `<div class="event-meta">Unable to load SharePoint lists. Please try again.</div>`;
 
public static singleListHtml(data: IAdminListView): string {
  return `
    <div class="event-list-item align-items-center w-100">
      <span aria-hidden="true">📁</span>
      <a href="${this.escape(data.url)}"
         class="event-title text-decoration-none mb-0 position-relative z-3 pe-auto"
         target="_self">
        ${this.escape(data.title)}
      </a>
    </div>`;
}
}