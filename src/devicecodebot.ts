import * as builder from 'botbuilder';

var adal = require('adal-node');
var AuthenticationContext = adal.AuthenticationContext;

const MicrosoftGraph = require("@microsoft/microsoft-graph-client");

export class DeviceCodeBot {
    public readonly Connector: builder.ChatConnector;
    private readonly universalBot: builder.UniversalBot;
    private readonly cache: any;

    constructor(connector: builder.ChatConnector) {
        this.Connector = connector;
        this.cache = new adal.MemoryCache()

        this.universalBot = new builder.UniversalBot(this.Connector);
        this.universalBot.dialog('/', this.defaultDialog);
        this.universalBot.dialog('/signin', this.signInDialog)
    }

    defaultDialog(session: builder.Session, args?: builder.IActionRouteData, next?: Function) {
        if (!session.userData.User) {
            session.send('You need to login to continue');
            session.beginDialog('/signin');
        }
        else {
            if (next) {
                session.endDialog();
                next();
            }
        }
    }
    public signInDialog<T>(session: builder.Session, args?: T): void {
        var context = new AuthenticationContext('https://login.microsoftonline.com/common', null, this.cache);

        context.acquireUserCode('https://graph.microsoft.com', process.env.MICROSOFT_APP_ID, '', (err: any, response: adal.IUserCodeResponse) => {
            if (err) {
                session.send(`I'm having trouble logging in. Please contact your administrators with this message: '${err.stack}`);
            } else {
                var dialog = new builder.SigninCard(session);
                dialog.text(response.message);
                dialog.button('Click here', response.verificationUrl);
                var msg = new builder.Message();
                msg.addAttachment(dialog);
                session.send(msg);
                try {
                    context.acquireTokenWithDeviceCode('https://graph.microsoft.com', process.env.MICROSOFT_APP_ID, response, (err: any, tokenResponse: adal.IDeviceCodeTokenResponse) => {
                        if (err) {
                            session.send(DeviceCodeBot.createErrorMessage(err));
                            session.beginDialog('/signin')

                        } else {
                            session.userData.accessToken = tokenResponse.accessToken;
                            session.send(`Hello ${tokenResponse.givenName} ${tokenResponse.familyName}`);
                            session.sendTyping();
                            const graphClient = MicrosoftGraph.Client.init({
                                authProvider: (done: any) => {
                                    done(null, session.userData.accessToken);
                                }
                            });
                            graphClient.
                                api('/me/jobTitle').
                                version('beta').
                                get((err: any, res: any) => {
                                    if (err) {
                                        session.send(DeviceCodeBot.createErrorMessage(err));
                                    } else {
                                        session.endDialog(`Oh, so you're a ${res.value}`);
                                    }
                                });
                        }
                    });
                } catch (err) {
                    session.send(DeviceCodeBot.createErrorMessage(err));
                }
            }
        });
    }

    public static createErrorMessage(err: any): string {
        return `I have some troubles accessing my inner thoughts. Please contact your administrators with this message: '${err}'`;
    }
}