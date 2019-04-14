import auth from '../../GraphAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import GraphCommand from '../../GraphCommand';
import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import * as fs from 'fs';
import * as path from 'path';


const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  teamId: string;
  imagePath: string;
}

class GraphTeamsPictureSetCommand extends GraphCommand {
  public get name(): string {
    return `${commands.TEAMS_PICTURE_SET}`;
  }
  public get description(): string {
    return 'Updates the picture of the specified Microsoft Teams team with the given image';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const fullPath: string = path.resolve(args.options.imagePath);
        if (this.verbose) {
          cmd.log(`Setting group logo ${fullPath}...`);
        }

        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/groups/${args.options.teamId}/photo/$value`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${auth.service.accessToken}`,
            'content-type': this.getImageContentType(fullPath)
          }),
          body: fs.readFileSync(fullPath)
        };

        // return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
        //   this.setGroupLogo(requestOptions, GraphO365GroupAddCommand.numRepeat, resolve, reject, cmd);
        // });

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.put(requestOptions);
      })      
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (err: any) => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private getImageContentType(imagePath: string): string {
    let extension: string = imagePath.substr(imagePath.lastIndexOf('.')).toLowerCase();

    switch (extension) {
      case '.png':
        return 'image/png';
      case '.gif':
        return 'image/gif';
      default:
        return 'image/jpeg';
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --teamId <teamId>',
        description: 'The ID of the team'
      },
      {
        option: '--imagePath <imagePath>',
        description: 'The path to the new Team image'
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

      if (!Utils.isValidGuid(args.options.teamId)) {
        return `${args.options.teamId} is not a valid GUID`;
      }

      if (!args.options.imagePath) {
        return 'Required parameter imagePath missing';
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

    To update properties of a specified channel in the given Microsoft Teams
    team, you have to first log in to the Microsoft Graph
    using the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN}`)}.

  Examples:

    Set new description and display name for the specified channel in the given
    Microsoft Teams team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects --description "Channel for new projects"

    Set new display name for the specified channel in the given Microsoft Teams
    team
      ${chalk.grey(config.delimiter)} ${this.name} --teamId "00000000-0000-0000-0000-000000000000" --channelName Reviews --newChannelName Projects
`);
  }
}

module.exports = new GraphTeamsPictureSetCommand();