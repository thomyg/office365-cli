import commands from '../../commands';
import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../GraphAuth';
const command: Command = require('./o365group-user-add');
import * as assert from 'assert';
import * as request from 'request-promise-native';
import Utils from '../../../../Utils';
import { Service } from '../../../../Auth';

describe(commands.O365GROUP_USER_ADD, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let trackEvent: any;
  let telemetry: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
    trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
      telemetry = t;
    });
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      log: (msg: string) => {
        log.push(msg);
      }
    };
    auth.service = new Service();
    telemetry = null;
    (command as any).items = [];
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      appInsights.trackEvent,
      auth.ensureAccessToken,
      auth.restoreAuth
    ]);
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.O365GROUP_USER_ADD), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('defines alias', () => {
    const alias = command.alias();
    assert.notEqual(typeof alias, 'undefined');
  });

  it('defines correct alias', () => {
    const alias = command.alias();
    assert.equal((alias && alias.indexOf(commands.TEAMS_USER_ADD) > -1), true);
  });

  it('calls telemetry', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert(trackEvent.called);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('logs correct telemetry event', (done) => {
    cmdInstance.action = command.action();
    cmdInstance.action({ options: {} }, () => {
      try {
        assert.equal(telemetry.name, commands.O365GROUP_USER_ADD);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the groupId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the teamId is not a valid guid.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        teamId: 'not-c49b-4fd4-8223-28f0ac3a6402'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the groupId is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        role: 'Member'
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation when both groupId and teamId are specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation if the userName is not provided.', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        role: 'Member',
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('fails validation when invalid role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Invalid',
      }
    });
    assert.notEqual(actual, true);
    done();
  });

  it('passes validation when valid groupId, userName and no role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('passes validation when valid groupId, userName and Owner role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Owner'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('passes validation when valid groupId, userName and Member role specified', (done) => {
    const actual = (command.validate() as CommandValidate)({
      options: {
        groupId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
        userName: 'anne.matthews@contoso.onmicrosoft.com',
        role: 'Member'
      }
    });
    assert.equal(actual, true);
    done();
  });

  it('correctly retrieves user and add member to specified Office 365 group', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and add member to specified Office 365 group (debug)', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and add owner to specified Office 365 group', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", role: "Owner" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and add owner to specified Microsoft Teams team (debug)', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/owners/$ref` &&
        JSON.stringify(opts.body) === `{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/00000000-0000-0000-0000-000000000001"}`) {
        addMemberRequestIssued = true;
      }
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com", role: "Owner" } }, () => {
      try {
        assert(addMemberRequestIssued);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly skips adding member or owner when user is not found', (done) => {
    let addMemberRequestIssued = false;

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews.not.found%40contoso.onmicrosoft.com/id`) {
        return Promise.reject({
          "message": "Resource 'anne.matthews.not.found%40contoso.onmicrosoft.com' does not exist or one of its queried reference-property objects are not present."
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref`) {
        addMemberRequestIssued = true;
      }
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, groupId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews.not.found@contoso.onmicrosoft.com" } }, () => {
      try {
        assert(addMemberRequestIssued === false);
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly handles error when user cannot be retrieved', (done) => {
    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/doesnotexist.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'Resource \'doesnotexist.matthews@contoso.onmicrosoft.com\' does not exist or one of its queried reference-property objects are not present.' } } } });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: "doesnotexist.matthews@contoso.onmicrosoft.com" } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Resource \'doesnotexist.matthews@contoso.onmicrosoft.com\' does not exist or one of its queried reference-property objects are not present.')));
        done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('correctly retrieves user and handle error adding member to specified Office 365 group', (done) => {

    sinon.stub(request, 'get').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/users/anne.matthews%40contoso.onmicrosoft.com/id`) {
        return Promise.resolve({
          "value": "00000000-0000-0000-0000-000000000001"
        });
      }

      return Promise.reject('Invalid request');
    });

    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000/members/$ref`) {
        return Promise.reject({ error: { 'odata.error': { message: { value: 'Invalid object identifier' } } } });
      }

      return Promise.reject('Invalid request');
    });

    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000", userName: "anne.matthews@contoso.onmicrosoft.com" } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Invalid object identifier'))); done();
      }
      catch (e) {

        done(e);
      }
    });
  });

  it('aborts when not logged in to Microsoft Graph', (done) => {
    auth.service = new Service();
    auth.service.connected = false;
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to the Microsoft Graph first')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsOption = true;
      }
    });
    assert(containsOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.O365GROUP_USER_ADD));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });

  it('correctly handles lack of valid access token', (done) => {
    Utils.restore(auth.ensureAccessToken);
    sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
    auth.service = new Service();
    auth.service.connected = true;
    auth.service.resource = 'https://graph.microsoft.com';
    cmdInstance.action = command.action();
    cmdInstance.action({ options: { debug: true } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});