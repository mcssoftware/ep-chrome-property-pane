// Legacy epconfiguration.js properties (default)
export interface EP {
  bundles: IEpBundles;
  features: IEpFeatures;
  isInEditMode: boolean;
  isTeams: boolean;
  taxonomyGroup: string;
  urls:IEpUrls;
  webpackManifest: Object;
}
export interface IEpUrls {
  calendarCenter: string;
  dependenciesUrl: string;
  emcUrl: string;
  emcAssets: string;
  emcVendorUrl: string;
  epNewsUrl: string;
  iiabUrl: string;
  newsCenterUrl: string;
  siteRoot: string;
  styleLibraryUrl: string;
  tenantHostUrl: string;
  webServerAbsoluteUrl: string;
}
export interface IEpFeatures {
  Discussions: boolean;
  Menu: boolean;
  Logo: boolean;
  Personalization: boolean;
  Pins: boolean;
  RSVP: boolean;
  Salt: boolean;
}
export interface IEpDiscussions {
  sharedAuthMode: boolean;
  sharedAuthRoleId: number;
}
export interface IEpBundles {
  ep: string;
  runtime: string;
}
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
/**
 * About: Inits EPModern Namespace if window.ElevatePoint does not already exist.
 * @param {string} absoluteUrl Input is the webabsoluteURL. From a webpart, the easiest way to accomplish this
 *  from the context of a webpart class is 'this.context.pageContext.site.absoluteUrl'.
 */
export function initNamespace(siteUrl:string):void {
  try {
    switch(!window.ElevatePoint || window.ElevatePoint === undefined || window.ElevatePoint === null) {
      case false: {
        break;
      }
      case true: {
        window.ElevatePoint = <EP> {
          bundles: <IEpBundles> {
            ep: null,
            runtime: null
          },
          features: <IEpFeatures> {
            Discussions: true,
            Menu: true,
            Logo: true,
            Personalization: true,
            Pins: true,
            RSVP: false,
            Salt: false
          },
          isInEditMode: null,
          isTeams: null,
          taxonomyGroup: 'ElevatePoint',
          urls: <IEpUrls> {
            calendarCenter: `${siteUrl}/calendars`,
            emcUrl: `${siteUrl}/sites/ep`,
            iiabUrl: `${siteUrl}`,
            newsCenterUrl: `${siteUrl}/news`,
            siteRoot: `${siteUrl}`,
            styleLibraryUrl: `${siteUrl}/style library`
          },
          webpackManifest: null
        };
        window.EpModern = <IEpModern> {
          siteName: 'MHSnet',
          urls: <IUrlConfiguration> {
            calendarCenterUrl: `${siteUrl}/calendars`,
            emcUrl: 'https://multicare.sharepoint.com/sites/ep',
            newsUrl: `${siteUrl}/news`,
            webAbsoluteUrl: siteUrl
          }
        };
        break;
      }
      default: break;
    }
  } catch (e) {
    console.log(e);
  }
}
