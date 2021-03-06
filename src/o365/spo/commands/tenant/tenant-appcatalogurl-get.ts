import auth from '../../SpoAuth';
import * as request from 'request-promise-native';
import config from '../../../../config';
import commands from '../../commands';
import Utils from '../../../../Utils';
import SpoCommand from '../../SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOGURL_GET;
  }

  public get description(): string {
    return 'Gets the URL of the tenant app catalog';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(accessToken);
          cmd.log('');
        }

        const requestOptions: any = {
          url: `${auth.site.url}/_api/SP_TenantSettings_Current`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            accept: 'application/json;odata=nometadata'
          })
        };
    
        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }
    
        return request.get(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const json = JSON.parse(res);

        if (json.CorporateCatalogUrl) {
          cmd.log(json.CorporateCatalogUrl);
        }
        else {
          if (this.verbose) {
            cmd.log("Tenant app catalog is not configured.");
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online
    tenant admin site, using the ${chalk.blue(commands.LOGIN)} command.

  Examples:
  
    Get the URL of the tenant app catalog
      ${chalk.grey(config.delimiter)} ${commands.TENANT_APPCATALOGURL_GET}
  ` );
  }
}

module.exports = new SpoTenantAppCatalogUrlGetCommand();