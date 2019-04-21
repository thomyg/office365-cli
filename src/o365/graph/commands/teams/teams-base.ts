

import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import auth from '../../GraphAuth';
import GraphCommand from '../../GraphCommand';



export abstract class GraphTeamsBaseCommand extends GraphCommand {

  /**
   * Gets or sets the Teams teamId.
   */
  protected teamId: string;


  constructor() {
    super();
    this.teamId = '';
  }

  /**
   * Gets the channelId by providing the channelName    
   * @param channelName the string representing a Teams channel name
   */
  protected getChannelIdByChannelName(channelName: string, cmd: CommandInstance): Promise<string> {
    const requestOptions: any = {
      url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/?$filter=displayName eq '${encodeURIComponent(channelName)}'`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${auth.service.accessToken}`,
        accept: 'application/json;odata.metadata=none'
      }),
      json: true
    };

    if (this.debug) {
      cmd.log('Request:');
      cmd.log(JSON.stringify(requestOptions));
      cmd.log('');
    }

    return new Promise<string>((resolve: any, reject: any): void => {
      request.get(requestOptions).then((res: any) => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(JSON.stringify(res));
          cmd.log('');
        }
        if(!res.value === undefined){
        return resolve(res.value[0].id);
        }
        else {
          return reject(`Cannot proceed. No channel with channelName:'${encodeURIComponent(channelName)}' present`);
        }

        reject('Cannot proceed. Error during getting channelId'); // this is not supposed to happen
      }, (err: any): void => { reject(err); })
    });

  }


  /**
   * Detects if the string in question is a valid Teams ChannelId
   * by the following RegEx ^19:\[0-9a-zA-Z]+@thread.skype$/i
   * @param channelId the string representing a Teams ChannelId
   */
  protected isValidChannelId(channelId: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^19:[0-9a-zA-Z]+@thread.skype$/i);

    return guidRegEx.test(channelId);
  }

}