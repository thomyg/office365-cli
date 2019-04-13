

import Utils from '../../../../Utils';
import * as request from 'request-promise-native';
import auth from '../../GraphAuth';
import GraphCommand from '../../GraphCommand';
import {Channel} from './Channel';


export abstract class GraphTeamsBaseCommand extends GraphCommand {

  /**
   * Gets or sets the Teams teamId.
   */
  protected teamId: string;

  /* istanbul ignore next */
  constructor() {
    super();
    this.teamId = '';
  }


  /**
   * Detects if the string in question is a valid Teams ChannelId
   * by the following RegEx ^19:\[0-9a-zA-Z]+@thread.skype$/i
   * @param channelName the string representing a Teams channel name
   */
  protected getChannelIdByChannelName(channelName: string, cmd: CommandInstance): string {

    let channelToDelete : Channel;
    let rv : string = "";

    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((): request.RequestPromise => {
        const requestOptions: any = {
          url: `${auth.service.resource}/v1.0/teams/${encodeURIComponent(this.teamId)}/channels/?$filter=displayName eq '${encodeURIComponent(channelName)}' `,
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
      .then((res: any): void => {
        if (this.debug) {
          cmd.log('Response:')
          cmd.log(res);
          cmd.log('');
        }

        channelToDelete = res.value[0];
        rv = channelToDelete.id;
      })

    
    return rv;

  }


  /**
   * Detects if the string in question is a valid Teams ChannelId
   * by the following RegEx ^19:\[0-9a-zA-Z]+@thread.skype$/i
   * @param channelId the string representing a Teams ChannelId
   */
  public static isValidTeamsChannelId(channelId: string): boolean {
    const guidRegEx: RegExp = new RegExp(/^19:[0-9a-zA-Z]+@thread.skype$/i);

    return guidRegEx.test(channelId);
  }

}