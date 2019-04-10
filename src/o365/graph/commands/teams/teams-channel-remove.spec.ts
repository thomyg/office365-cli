// import commands from '../../commands';
// import Command, { CommandOption, CommandError, CommandValidate } from '../../../../Command';
// import * as sinon from 'sinon';
// import appInsights from '../../../../appInsights';
// import auth from '../../GraphAuth';
// const command: Command = require('./teams-remove');
// import * as assert from 'assert';
// import * as request from 'request-promise-native';
// import Utils from '../../../../Utils';
// import { Service } from '../../../../Auth';
// import * as fs from 'fs';

// describe(commands.TEAMS_REMOVE, () => {
//   let vorpal: Vorpal;
//   let log: string[];
//   let cmdInstance: any;
//   let trackEvent: any;
//   let telemetry: any;
//   let promptOptions: any;

//   before(() => {
//     sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
//     sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.resolve('ABC'); });
//     trackEvent = sinon.stub(appInsights, 'trackEvent').callsFake((t) => {
//       telemetry = t;
//     });
//     sinon.stub(fs, 'readFileSync').callsFake(() => 'abc');
//   });

//   beforeEach(() => {
//     vorpal = require('../../../../vorpal-init');
//     log = [];
//     cmdInstance = {
//       log: (msg: string) => {
//         log.push(msg);
//       },
//       prompt: (options: any, cb: (result: { continue: boolean }) => void) => {
//         promptOptions = options;
//         cb({ continue: false });
//       }
//     };
//     auth.service = new Service();
//     telemetry = null;
//     promptOptions = undefined;
//   });

//   afterEach(() => {
//     Utils.restore([
//       vorpal.find,
//       request.get,
//       request.delete,
//       global.setTimeout
//     ]);
//   });

//   after(() => {
//     Utils.restore([
//       appInsights.trackEvent,
//       auth.ensureAccessToken,
//       auth.restoreAuth,
//       fs.readFileSync
//     ]);
//   });

//   it('has correct name', () => {
//     assert.equal(command.name.startsWith(commands.TEAMS_REMOVE), true);
//   });

//   it('has a description', () => {
//     assert.notEqual(command.description, null);
//   });

//   it('calls telemetry', (done) => {
//     cmdInstance.action = command.action();
//     cmdInstance.action({ options: {} }, () => {
//       try {
//         assert(trackEvent.called);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('logs correct telemetry event', (done) => {
//     cmdInstance.action = command.action();
//     cmdInstance.action({ options: {} }, () => {
//       try {
//         assert.equal(telemetry.name, commands.TEAMS_REMOVE);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('fails validation if the teamId is not a valid guid.', (done) => {
//     const actual = (command.validate() as CommandValidate)({
//       options: {
//         teamId: '61703ac8a-c49b-4fd4-8223-28f0ac3a6402'
//       }
//     });
//     assert.notEqual(actual, true);
//     done();
//   });

//   it('fails validation if the teamId is not provided.', (done) => {
//     const actual = (command.validate() as CommandValidate)({
//       options: {
        
//       }
//     });
//     assert.notEqual(actual, true);
//     done();
//   });

//   it('passes validation when valid teamId is specified', (done) => {
//     const actual = (command.validate() as CommandValidate)({
//       options: {
//         teamId: '6703ac8a-c49b-4fd4-8223-28f0ac3a6402',
//       }
//     });
//     assert.equal(actual, true);
//     done();
//   });

//   it('prompts before removing the specified team when confirm option not passed', (done) => {
//     auth.service = new Service('https://graph.microsoft.com');
//     auth.service.connected = true;

//     cmdInstance.action = command.action();
//     cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000"} }, () => {
//       let promptIssued = false;

//       if (promptOptions && promptOptions.type === 'confirm') {
//         promptIssued = true;
//       }

//       try {
//         assert(promptIssued);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('prompts before removing the specified team when confirm option not passed (debug)', (done) => {
//     auth.service = new Service('https://graph.microsoft.com');
//     auth.service.connected = true;

//     cmdInstance.action = command.action();
//     cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
//       let promptIssued = false;

//       if (promptOptions && promptOptions.type === 'confirm') {
//         promptIssued = true;
//       }

//       try {
//         assert(promptIssued);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('aborts removing the specified team when confirm option not passed and prompt not confirmed', (done) => {
//     const postSpy = sinon.spy(request, 'delete');
//     auth.service = new Service('https://graph.microsoft.com');
//     auth.service.connected = true;

//     cmdInstance.action = command.action();
//     cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
//       cb({ continue: false });
//     };
//     cmdInstance.action({ options: { debug: false, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
//       try {
//         assert(postSpy.notCalled);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('aborts removing the specified team when confirm option not passed and prompt not confirmed (debug)', (done) => {
//     const postSpy = sinon.spy(request, 'delete');
//     auth.service = new Service('https://graph.microsoft.com');
//     auth.service.connected = true;

//     cmdInstance.action = command.action();
//     cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
//       cb({ continue: false });
//     };
//     cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
//       try {
//         assert(postSpy.notCalled);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });


//   it('removes the  specified team when prompt confirmed (debug)', (done) => {
//     let teamsDeleteCallIssued = false;
   
//     sinon.stub(request, 'delete').callsFake((opts) => {
//       if (opts.url === `https://graph.microsoft.com/v1.0/groups/00000000-0000-0000-0000-000000000000`) {
//         teamsDeleteCallIssued = true;
//       }
//     });

//     auth.service = new Service('https://graph.microsoft.com');
//     auth.service.connected = true;

//     cmdInstance.action = command.action();
//     cmdInstance.prompt = (options: any, cb: (result: { continue: boolean }) => void) => {
//       cb({ continue: true });
//     };
//     cmdInstance.action({ options: { debug: true, teamId: "00000000-0000-0000-0000-000000000000" } }, () => {
//       try {
//         assert(teamsDeleteCallIssued);
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('aborts when not logged in to Microsoft Graph', (done) => {
//     auth.service = new Service();
//     auth.service.connected = false;
//     cmdInstance.action = command.action();
//     cmdInstance.action({ options: { debug: true } }, (err?: any) => {
//       try {
//         assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Log in to the Microsoft Graph first')));
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });

//   it('supports debug mode', () => {
//     const options = (command.options() as CommandOption[]);
//     let containsOption = false;
//     options.forEach(o => {
//       if (o.option === '--debug') {
//         containsOption = true;
//       }
//     });
//     assert(containsOption);
//   });

//   it('has help referring to the right command', () => {
//     const cmd: any = {
//       log: (msg: string) => { },
//       prompt: () => { },
//       helpInformation: () => { }
//     };
//     const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
//     cmd.help = command.help();
//     cmd.help({}, () => { });
//     assert(find.calledWith(commands.TEAMS_REMOVE));
//   });

//   it('has help with examples', () => {
//     const _log: string[] = [];
//     const cmd: any = {
//       log: (msg: string) => {
//         _log.push(msg);
//       },
//       prompt: () => { },
//       helpInformation: () => { }
//     };
//     sinon.stub(vorpal, 'find').callsFake(() => cmd);
//     cmd.help = command.help();
//     cmd.help({}, () => { });
//     let containsExamples: boolean = false;
//     _log.forEach(l => {
//       if (l && l.indexOf('Examples:') > -1) {
//         containsExamples = true;
//       }
//     });
//     Utils.restore(vorpal.find);
//     assert(containsExamples);
//   });

//   it('correctly handles lack of valid access token', (done) => {
//     Utils.restore(auth.ensureAccessToken);
//     sinon.stub(auth, 'ensureAccessToken').callsFake(() => { return Promise.reject(new Error('Error getting access token')); });
//     auth.service = new Service();
//     auth.service.connected = true;
//     auth.service.resource = 'https://graph.microsoft.com';
//     cmdInstance.action = command.action();
//     cmdInstance.action({ options: { debug: true, confirm: true } }, (err?: any) => {
//       try {
//         assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('Error getting access token')));
//         done();
//       }
//       catch (e) {
//         done(e);
//       }
//     });
//   });
// });