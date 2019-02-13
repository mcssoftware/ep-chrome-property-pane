// TODO: Refactor or eliminate
export interface IEpModern {
  siteName: string;
  urls: IUrlConfiguration;
}
export interface IUrlConfiguration {
  calendarCenterUrl: string;
  emcUrl?: string;
  newsUrl: string;
  styleLibraryUrl?: string;
  webAbsoluteUrl?: string;
}


export function initGlobalVars(): void {
  const myUrls: IUrlConfiguration = {
    calendarCenterUrl: "https://cwsoft.sharepoint.com/sites/ad/calendars",
    emcUrl: "https://cwsoft.sharepoint.com/sites/ad/ep",
    newsUrl: "https://cwsoft.sharepoint.com/sites/ad/news",
    styleLibraryUrl: "https://multicare.sharepoint.com/Style%20Library",
    webAbsoluteUrl: "https://multicare.sharepoint.com"
  };
  const myVars: IEpModern = {
    siteName: "MHSnet",
    urls: myUrls
  };
  window.Epmodern = myVars;
}
