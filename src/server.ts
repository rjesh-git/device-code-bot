import * as Express from 'express';
import * as https from 'https';
import * as http from 'http';
import * as path from 'path';
import * as morgan from 'morgan';
import * as builder from 'botbuilder';
import { DeviceCodeBot } from './devicecodebot';

require('dotenv').config();

process.on('uncaughtException', (err: any) => {
    console.log('Caught exception: ' + err);
});

let express = Express();
let port = process.env.port || process.env.PORT || 3007;

express.use(morgan('tiny'));

let bot = new DeviceCodeBot(new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
}));

express.post('/api/messages', bot.Connector.listen());

express.set('port', port);
http.createServer(express).listen(port, (err: any) => {
    if (err) {
        return console.error(err);
    }
    console.log(`Server running on ${port}`);

})
