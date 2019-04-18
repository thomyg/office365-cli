import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../GraphCommand';
import Utils from '../../../../Utils';
import { Channel } from './Channel';
import { TeamsApp } from './TeamsApp';
import * as request from 'request-promise-native';
import { GraphResponse } from '../../GraphResponse';
//import { Team } from './Team';



const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  filePath: string;
}

interface TeamsExportDefinition {
  id: string;
  internalId: string;
  isArchived: boolean | undefined;
  webUrl: string;
  messagingSettings: any;
  memberSettings: any;
  guestSettings: any;
  funSettings: any;
  channels: any | undefined,
  installedApps: any | undefined,
}

class GraphTeamsExportCommand extends GraphCommand {

  protected channels: Channel[];
  protected apps: TeamsApp[];
  protected teamsExportDefinition: TeamsExportDefinition;

  /* istanbul ignore next */
  constructor() {
    super();
    this.channels = [];
    this.apps = [];
    this.teamsExportDefinition = {} as TeamsExportDefinition;
  }

  public get name(): string {
    return `${commands.TEAMS_EXPORT}`;
  }

  public get description(): string {
    return 'Lists channels in the specified Microsoft Teams team';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const endpoint: string = `${auth.service.resource}/v1.0/teams/${args.options.teamId}/channels`;

    this
      .getAllChannels(endpoint, cmd, true)
      .then((): void => {
        if (args.options.output === 'json') {
          cmd.log(this.channels);
        }
        else {
          cmd.log(this.channels.map(m => {
            return {
              id: m.id,
              displayName: m.displayName
            }
          }));
        }
      })      
      .then(async (res: any): Promise<string> => {

        await this.getAllTabs(args.options.teamId, cmd);

        return JSON.stringify(this.channels);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Channels');
          cmd.log(JSON.stringify(this.channels));
          cmd.log('');
        }
      })
      .then(async (res:any): Promise<void> => {

        await this.getAllAppsFromTeam(args.options.teamId, cmd, true);

      })
      .then(async (res:any): Promise<void> => {

        await this.getTeamBaseDefinition(args.options.teamId, cmd);

      })
      .then((): void => {

        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }
        this.createExportFile(args.options.filePath);
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }
 
  protected createExportFile(filePath: string){
    var fs = require("fs");
    this.teamsExportDefinition.channels = this.channels;
    this.teamsExportDefinition.installedApps = this.apps;

    var fileContent = JSON.stringify(this.teamsExportDefinition); 

    fs.writeFile(filePath, fileContent, (err:string) => {
        if (err) {
            console.error(err);
            return;
        };
        console.log("File has been created");
    });
  }

  protected async getAllTabs(teamId: string, cmd: CommandInstance): Promise<void> {
    for (var i = 0; i < this.channels.length; i++) {
      cmd.log(i);
      await this.getAllTabsForChannel(teamId, this.channels[i].id, cmd, false);
    }
  }

  protected async getTeamBaseDefinition(teamId: string, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(teamId)}`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
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
        .then((res: TeamsExportDefinition): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          this.teamsExportDefinition = res;
          
          resolve(); 
          
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  protected async getAllChannels(url: string, cmd: CommandInstance, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: url,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
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
        .then((res: GraphResponse<Channel>): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          if (firstRun) {
            this.channels = [];
          }

          this.channels = this.channels.concat(res.value);

          if (res['@odata.nextLink']) {
            this
              .getAllChannels(res['@odata.nextLink'] as string, cmd, false)
              .then((): void => {
                resolve();
              }, (err: any): void => {
                reject(err);
              });
          }
          else {
            resolve();
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }

  protected getAllTabsForChannel(teamId: string, channelId: string, cmd: CommandInstance, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(teamId)}/channels/${encodeURIComponent(channelId)}/tabs`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
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
        .then((res: { value: string; }): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }
          for (var i = 0; i < this.channels.length; i++) {
            if (this.channels[i].id === channelId) {
              this.channels[i].tabs = res.value;
            }
          }
          resolve();
        });
    });
  }

  protected getAllAppsFromTeam(teamId: string, cmd: CommandInstance, firstRun: boolean): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      let endpoint: string = '';
      endpoint = `${auth.service.resource}/v1.0/teams/${encodeURIComponent(teamId)}/installedApps?$expand=teamsAppDefinition`;
      //endpoint += `&$filter=teamsApp/distributionMethod eq 'organization'`;
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: endpoint,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              accept: 'application/json;odata.metadata=none'
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
        .then((res: GraphResponse<TeamsApp>): void => {
          if (this.debug) {
            cmd.log('Response:');
            cmd.log(res);
            cmd.log('');
          }

          if (firstRun) {
            this.apps = [];
          }

          this.apps = this.apps.concat(res.value);

          if (res['@odata.nextLink']) {
            this
              .getAllChannels(res['@odata.nextLink'] as string, cmd, false)
              .then((): void => {
                resolve();
              }, (err: any): void => {
                reject(err);
              });
          }
          else {
            resolve();
          }
        }, (err: any): void => {
          reject(err);
        });
    });
  }


  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team to list the channels of'
      },
      {
        option: '-p, --filePath <filePath>',
        description: 'The path to the file where the exported template should be stored'
      },
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!Utils.isValidGuid(args.options.teamId as string)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!args.options.filePath) {
        return 'Required parameter filePath missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command.
          
  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.

    To list the channels in a Microsoft Teams team, you have to first log in to
    the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:
  
    List the channels in a specified Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId 00000000-0000-0000-0000-000000000000
`   );
  }
}

module.exports = new GraphTeamsExportCommand();