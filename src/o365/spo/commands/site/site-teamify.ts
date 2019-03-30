import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  url: string;
}

class SpoSiteGetCommand extends SpoCommand {
  public get name(): string {
    return commands.SITE_TEAMIFY;
  }

  public get description(): string {
    return 'Connects existing team site to Microsoft Teams';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.url);
    let siteAccessToken: string = '';

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Retrieving information about site collection ${args.options.url}...`);
        }

        const requestOptions: any = {
          url: `${args.options.url}/_api/groupsitemanager/EnsureTeamForGroup`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
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
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(res);

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [{
      option: '-u, --url <url>',
      description: 'The URL of the site to connect to Microsoft Teams'
    }];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.url) {
        return 'Required parameter url missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.url);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To get information about a site collection, you have to first log in to
    a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    This command connects a modern team site with a Microsoft Teams team.

    You can not connect a communication site No GroupId for this site. If you try to do so
    the error message will be "No GroupId for this site.".
   
  Examples:
  
    Return information about the ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
    site collection.
      ${chalk.grey(config.delimiter)} ${commands.SITE_GET} -u https://contoso.sharepoint.com/sites/project-x
`);
  }
}

module.exports = new SpoSiteGetCommand();