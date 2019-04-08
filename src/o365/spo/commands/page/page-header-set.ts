import auth from '../../SpoAuth';
import config from '../../../../config';
import * as request from 'request-promise-native';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';
import { Auth } from '../../../../Auth';
import { PageHeader, CustomPageHeader, CustomPageHeaderServerProcessedContent, CustomPageHeaderProperties } from './PageHeader';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  altText?: string;
  authors?: string;
  imageUrl?: string;
  kicker?: string;
  layout?: string;
  pageName: string;
  showKicker?: boolean;
  showPublishDate?: boolean;
  textAlignment?: string;
  translateX?: number;
  translateY?: number;
  type?: string;
  webUrl: string;
}

class SpoPageHeaderSetCommand extends SpoCommand {
  public get name(): string {
    return `${commands.PAGE_HEADER_SET}`;
  }

  public get description(): string {
    return 'Sets modern page header';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.altText = typeof args.options.altText !== 'undefined';
    telemetryProps.authors = typeof args.options.authors !== 'undefined';
    telemetryProps.imageUrl = typeof args.options.imageUrl !== 'undefined';
    telemetryProps.kicker = typeof args.options.kicker !== 'undefined';
    telemetryProps.layout = args.options.layout;
    telemetryProps.showKicker = args.options.showKicker;
    telemetryProps.showPublishDate = args.options.showPublishDate;
    telemetryProps.textAlignment = args.options.textAlignment;
    telemetryProps.translateX = typeof args.options.translateX !== 'undefined';
    telemetryProps.translateY = typeof args.options.translateY !== 'undefined';
    telemetryProps.type = args.options.type;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    const noPageHeader: PageHeader = {
      "id": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "instanceId": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "title": "Title Region",
      "description": "Title Region Description",
      "serverProcessedContent": {
        "htmlStrings": {},
        "searchablePlainTexts": {},
        "imageSources": {},
        "links": {}
      },
      "dataVersion": "1.4",
      "properties": {
        "title": "",
        "imageSourceType": 4,
        "layoutType": "NoImage",
        "textAlignment": "Left",
        "showKicker": false,
        "showPublishDate": false,
        "kicker": ""
      }
    };
    const defaultPageHeader: PageHeader = {
      "id": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "instanceId": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "title": "Title Region",
      "description": "Title Region Description",
      "serverProcessedContent": {
        "htmlStrings": {},
        "searchablePlainTexts": {},
        "imageSources": {},
        "links": {}
      },
      "dataVersion": "1.4",
      "properties": {
        "title": "",
        "imageSourceType": 4,
        "layoutType": "FullWidthImage",
        "textAlignment": "Left",
        "showKicker": false,
        "showPublishDate": false,
        "kicker": ""
      }
    };
    const customPageHeader: CustomPageHeader = {
      "id": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "instanceId": "cbe7b0a9-3504-44dd-a3a3-0e5cacd07788",
      "title": "Title Region",
      "description": "Title Region Description",
      "serverProcessedContent": {
        "htmlStrings": {},
        "searchablePlainTexts": {},
        "imageSources": {
          "imageSource": ""
        },
        "links": {},
        "customMetadata": {
          "imageSource": {
            "siteId": "",
            "webId": "",
            "listId": "",
            "uniqueId": ""
          }
        }
      },
      "dataVersion": "1.4",
      "properties": {
        "title": "",
        "imageSourceType": 2,
        "layoutType": "FullWidthImage",
        "textAlignment": "Left",
        "showKicker": false,
        "showPublishDate": false,
        "kicker": "",
        "authors": [],
        "altText": "",
        "webId": "",
        "siteId": "",
        "listId": "",
        "uniqueId": "",
        "translateX": 0,
        "translateY": 0
      }
    };
    let header: PageHeader | CustomPageHeader = defaultPageHeader;
    let siteAccessToken: string = '';
    let pageFullName: string = args.options.pageName.toLowerCase();
    if (pageFullName.indexOf('.aspx') < 0) {
      pageFullName += '.aspx';
    }
    let title: string;

    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}`);
        }

        siteAccessToken = accessToken;

        if (this.verbose) {
          cmd.log(`Retrieving information about the page...`);
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')?$select=IsPageCheckedOutToCurrentUser,Title`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: { IsPageCheckedOutToCurrentUser: boolean, Title: string; }): Promise<void> | request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        title = res.Title;

        if (res.IsPageCheckedOutToCurrentUser) {
          return Promise.resolve();
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')/checkoutpage`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((): Promise<any[] | void> => {
        switch (args.options.type) {
          case 'None':
            header = noPageHeader;
            break;
          case 'Default':
            header = defaultPageHeader;
            break;
          case 'Custom':
            header = customPageHeader;
            break;
          default:
            header = defaultPageHeader;
        }

        header.properties.title = title;
        header.properties.textAlignment = args.options.textAlignment as any || 'Left';
        header.properties.showKicker = args.options.showKicker || false;
        header.properties.kicker = args.options.kicker || '';
        header.properties.showPublishDate = args.options.showPublishDate || false;

        if (args.options.type !== 'None') {
          header.properties.layoutType = args.options.layout as any || 'FullWidthImage';
        }

        if (args.options.type === 'Custom') {
          header.serverProcessedContent.imageSources = {
            imageSource: args.options.imageUrl || ''
          };
          const properties: CustomPageHeaderProperties = header.properties as CustomPageHeaderProperties;
          properties.altText = args.options.altText || '';
          properties.authors = args.options.authors ? args.options.authors.split(',').map(a => a.trim()) : [];
          properties.translateX = args.options.translateX || 0;
          properties.translateY = args.options.translateY || 0;
          header.properties = properties;

          if (!args.options.imageUrl) {
            (header.serverProcessedContent as CustomPageHeaderServerProcessedContent).customMetadata = {
              imageSource: {
                siteId: '',
                webId: '',
                listId: '',
                uniqueId: ''
              }
            };
            properties.listId = '';
            properties.siteId = '';
            properties.uniqueId = '';
            properties.webId = '';
            header.properties = properties;

            return Promise.resolve();
          }

          return Promise.all([
            this.getSiteId(args.options.webUrl, siteAccessToken, this.verbose, this.debug, cmd),
            this.getWebId(args.options.webUrl, siteAccessToken, this.verbose, this.debug, cmd),
            this.getImageInfo(args.options.webUrl, args.options.imageUrl as string, siteAccessToken, this.verbose, this.debug, cmd),
          ]);
        }
        else {
          return Promise.resolve();
        }
      })
      .then((res: void | any[]): request.RequestPromise => {
        if (res) {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          (header.serverProcessedContent as CustomPageHeaderServerProcessedContent).customMetadata = {
            imageSource: {
              siteId: res[0].Id,
              webId: res[1].Id,
              listId: res[2].ListId,
              uniqueId: res[2].UniqueId
            }
          };
          const properties: CustomPageHeaderProperties = header.properties as CustomPageHeaderProperties;
          properties.listId = res[2].ListId;
          properties.siteId = res[0].Id;
          properties.uniqueId = res[2].UniqueId;
          properties.webId = res[1].Id;
          header.properties = properties;
        }

        const requestOptions: any = {
          url: `${args.options.webUrl}/_api/sitepages/pages/GetByUrl('sitepages/${encodeURIComponent(pageFullName)}')/savepage`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata',
            'content-type': 'application/json;odata=nometadata'
          }),
          body: {
            LayoutWebpartsContent: JSON.stringify([header])
          },
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.post(requestOptions);
      })
      .then((res: any): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getSiteId(siteUrl: string, accessToken: string, verbose: boolean, debug: boolean, cmd: CommandInstance): request.RequestPromise {
    if (verbose) {
      cmd.log(`Retrieving information about the site collection...`);
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/site?$select=Id`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${accessToken}`,
        accept: 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (debug) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

    return request.get(requestOptions);
  }

  private getWebId(siteUrl: string, accessToken: string, verbose: boolean, debug: boolean, cmd: CommandInstance): request.RequestPromise {
    if (verbose) {
      cmd.log(`Retrieving information about the site...`);
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/web?$select=Id`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${accessToken}`,
        accept: 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (debug) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

    return request.get(requestOptions);
  }

  private getImageInfo(siteUrl: string, imageUrl: string, accessToken: string, verbose: boolean, debug: boolean, cmd: CommandInstance): request.RequestPromise {
    if (verbose) {
      cmd.log(`Retrieving information about the header image...`);
    }

    const requestOptions: any = {
      url: `${siteUrl}/_api/web/getfilebyserverrelativeurl('${encodeURIComponent(imageUrl)}')?$select=ListId,UniqueId`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${accessToken}`,
        accept: 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (debug) {
      cmd.log('Executing web request...');
      cmd.log(requestOptions);
      cmd.log('');
    }

    return request.get(requestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --pageName <pageName>',
        description: 'Name of the page to set the header for'
      },
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the page to update is located'
      },
      {
        option: '-t, --type [type]',
        description: 'Type of header, allowed values None|Default|Custom. Default Default',
        autocomplete: ['None', 'Default', 'Custom']
      },
      {
        option: '--imageUrl [imageUrl]',
        description: 'Server-relative URL of the image to use in the header. Image must be stored in the same site collection as the page'
      },
      {
        option: '--altText [altText]',
        description: 'Header image alt text'
      },
      {
        option: '-x, --translateX [translateX]',
        description: 'X focal point of the header image'
      },
      {
        option: '-y, --translateY [translateY]',
        description: 'Y focal point of the header image'
      },
      {
        option: '--layout [layout]',
        description: 'Layout to use in the header. Allowed values FullWidthImage|NoImage. Default FullWidthImage',
        autocomplete: ['FullWidthImage', 'NoImage']
      },
      {
        option: '--textAlignment [textAlignment]',
        description: 'How to align text in the header. Allowed values Center|Left. Default Left',
        autocomplete: ['Left', 'Center']
      },
      {
        option: '--showKicker',
        description: 'Set, to show the kicker'
      },
      {
        option: '--showPublishDate',
        description: 'Set, to show the publishing date'
      },
      {
        option: '--kicker [kicker]',
        description: 'Text to show in the kicker, when showKicker is set'
      },
      {
        option: '--authors [authors]',
        description: 'Comma-separated list of page authors to show in the header'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.pageName) {
        return 'Required parameter pageName missing';
      }

      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      if (args.options.type &&
        args.options.type !== 'None' &&
        args.options.type !== 'Default' &&
        args.options.type !== 'Custom') {
        return `${args.options.type} is not a valid type value. Allowed values None|Default|Custom`;
      }

      if (args.options.translateX && isNaN(args.options.translateX)) {
        return `${args.options.translateX} is not a valid number`;
      }

      if (args.options.translateY && isNaN(args.options.translateY)) {
        return `${args.options.translateY} is not a valid number`;
      }

      if (args.options.layout &&
        args.options.layout !== 'FullWidthImage' &&
        args.options.layout !== 'NoImage') {
        return `${args.options.layout} is not a valid layout value. Allowed values FullWidthImage|NoImage`;
      }

      if (args.options.textAlignment &&
        args.options.textAlignment !== 'Left' &&
        args.options.textAlignment !== 'Center') {
        return `${args.options.textAlignment} is not a valid textAlignment value. Allowed values Left|Center`;
      }

      return SpoCommand.isValidSharePointUrl(args.options.webUrl);
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site using 
    the ${chalk.blue(commands.LOGIN)} command.
        
  Remarks:

    To set modern page header, you have to first log in to a SharePoint site
    using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    If the specified ${chalk.grey('name')} doesn't refer to an existing modern page, you will get
    a ${chalk.grey('File doesn\'t exists')} error.

    The ${chalk.blue('showKicker')}, ${chalk.blue('kicker')} and ${chalk.blue('authors')} options are based on preview
    functionality that isn't available on all tenants yet.

  Examples:
  
    Reset the page header to default
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx

    Use the specified image focused on the given coordinates in the page header
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --type Custom --imageUrl /sites/team-a/SiteAssets/hero.jpg --altText 'Sunset over the ocean' --translateX 42.3837520042758 --translateY 56.4285714285714

    Center the page title in the header and show the publishing date
      ${chalk.grey(config.delimiter)} ${this.name} --webUrl https://contoso.sharepoint.com/sites/team-a --pageName home.aspx --textAlignment Center --showPublishDate
`);
  }
}

module.exports = new SpoPageHeaderSetCommand();