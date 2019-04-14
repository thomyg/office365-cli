import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import { GraphTeamsBaseCommand } from './teams-base';


const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  channelId?: string;
  channelName?: string;
  confirm?: boolean;
}

class GraphTeamsChannelRemoveCommand extends GraphTeamsBaseCommand {
  public get name(): string {
    return `${commands.TEAMS_CHANNEL_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified channel from the specified Microsoft Teams team';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = (!(!args.options.confirm)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const providedChannelId: string = (typeof args.options.channelId !== 'undefined') ? args.options.channelId : "" as string
    const removeChannelById: () => void = (): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): request.RequestPromise => {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(providedChannelId)}`,
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

          return request.delete(requestOptions);
        })


        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    const removeChannelByName: () => void = (): void => {
      auth
        .ensureAccessToken(auth.service.resource, cmd, this.debug)
        .then((): Promise<Object> => {
          let channelName: string = "";
          if (args.options.channelName) {
            channelName = args.options.channelName;
          }
          return this.getChannelIdByChannelName(channelName, cmd);
        })
        .then((res: any): Promise<void> | request.RequestPromise => {
          const requestOptions: any = {
            url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(args.options.teamId)}/channels/${encodeURIComponent(res)}`,
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${auth.service.accessToken}`,
              'accept': 'application/json;odata.metadata=none'
            }),
          };

          if (this.debug) {
            cmd.log('Executing web request...');
            cmd.log(requestOptions);
            cmd.log('');
          }

          return request.delete(requestOptions);
        })


        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      if (args.options.channelId) {
        //removeTeamByChannelId();
      }
      if (args.options.channelName) {
        removeChannelByName();
      }
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the channel ${args.options.channelId}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          if (args.options.channelId) {
            if (3 > 5) {
              removeChannelById();
            }
          }
          if (args.options.channelName) {
            this.teamId = args.options.teamId;
            removeChannelByName();
          }
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the Teams team to remove'
      },
      {
        option: '-c, --channelId [channelId]',
        description: 'The ID of the Teams channel to remove'
      },
      {
        option: '-n, --channelName [channelName]',
        description: 'The name of the Teams channel to remove'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the specified team'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.teamId) {
        return 'Required parameter teamId missing';
      }

      if (!args.options.channelId && !args.options.channelName) {
        return 'Required parameters channelId or channelName missing';
      }

      if (args.options.channelId && args.options.channelName) {
        return 'Specify channelId or channelName but not both';
      }

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if(args.options.channelId){
        if (!this.isValidChannelId(args.options.channelId)) {
          return `${args.options.channelName} is not a valid Teams channelId`;
        }
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

    To remove the specified Microsoft Teams team, you have to first
    log in to the Microsoft Graph using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

    When deleted, Office 365 groups are moved to a temporary container and can be restored within 30 days. 
    After that time, they are permanently deleted. 
    To learn more, see https://docs.microsoft.com/en-us/graph/api/resources/directory?view=graph-rest-1.0. 
    This applies only to Office 365 groups.

  Examples:
  
    Removes the specified team 
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000'

    Removes the specified team without confirmation
      ${chalk.grey(config.delimiter)} ${this.name} --teamId '00000000-0000-0000-0000-000000000000' --confirm 
  `);
  }
}

module.exports = new GraphTeamsChannelRemoveCommand();