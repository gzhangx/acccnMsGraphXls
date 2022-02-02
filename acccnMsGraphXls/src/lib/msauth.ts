import { Axios } from "axios";
//import * as  promise from 'bluebird';
import * as fs from "fs";
import { get } from 'lodash';

import creds from '../../credentials.json';

export interface IAuthOpt {
    tenantId: string;
    client_id: string;
    promptUser: (msg: string|object, info:object) => void;
    saveToken: (token: object) => Promise<void>;
    loadToken: () => Promise<IAuthCreds>;
    scope?: string;
    pollTime?: number;
}

interface IAuthCreds {
    refresh_token: string;
}
interface ICodeWaitInfo {
    device_code: string;
    message: object;
}

export interface ITokenInfo {
    access_token: string;
    expires_on: number;
}

async function delay(ms: number) {
    return new Promise(resolve => {
            
        setTimeout(() => {
            resolve();
        }, ms);
    });
}

export function getAuth(opt: IAuthOpt) {
    const tenantId = opt.tenantId;
    const client_id = opt.client_id;
    if (!tenantId) throw {
        message: 'tenantId required'
    }
    if (!client_id) throw {
        message: 'client_id required'
    }

    const promptUser = opt.promptUser || console.log;
    const saveToken = opt.saveToken;
    const loadToken = opt.loadToken;

    const resource = 'https://graph.microsoft.com';
    const scope = opt.scope || 'Mail.Read openid profile User.Read email Files.ReadWrite.All Files.ReadWrite Files.Read Files.Read.All Files.Read.Selected Files.ReadWrite.AppFolder Files.ReadWrite.Selected';
    const queryCodeurl = `https://login.microsoftonline.com/${tenantId}/oauth2/token`;
    const request = new Axios({});

    function getFormData(obj: {[id:string]:any}): string {        
        const keys = Object.keys(obj);
        const data = keys.map(key => {
            return `${key}=${encodeURIComponent(obj[key])}`;
        }).join('&')
        return data;
    }
    async function doPost(url: string, data: { [id: string]: any }): Promise<any> {
        const dataStr = getFormData(data);
        return await request.post(url, dataStr, {
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        }).then(r => {
            return JSON.parse(r.data);
        });
    }
    async function getRefreshToken() {        
        const codeWaitInfo = await doPost(`https://login.microsoftonline.com/${tenantId}/oauth2/devicecode`, {
            resource,
            scope,
            client_id,
        }) as ICodeWaitInfo;

        //const user_code = codeWaitInfo.user_code; // presented to the user
        const deviceCode = codeWaitInfo.device_code; // internal code to identify the user
        //const url = codeWaitInfo.verification_url; // URL the user needs to visit & paste in the code
        const message = codeWaitInfo.message; //send user code to url
        await promptUser(message, codeWaitInfo);
        while (true) {
            const rr = await doPost(queryCodeurl, {
                grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                resource: 'https://graph.microsoft.com',
                scope,
                code: deviceCode,
                client_id
            });
            if (rr.error === 'authorization_pending') {
                //await promise.Promise.delay(opt.pollTime || 1000);
                await delay(opt.pollTime || 1000);
                continue;
            }
            ///console.log(rr);
            //const { access_token, refresh_token } = rr;
            //fs.writeFileSync('credentials.json', JSON.stringify(rr, null, 2));
            await saveToken(rr);
            return rr;
        }
    }

    async function getAccessToken(): Promise<ITokenInfo> {
        const credentials = await loadToken();;
        const { refresh_token } = credentials;
        const form = {
            scope,
            resource,
            refresh_token,
            grant_type: 'refresh_token',
            client_id
        };
        const ac = await doPost(queryCodeurl, form) as ITokenInfo;

        return ac;
    }

    //getAuth({tenantId, client_id, promptUser, saveToken, loadToken})
    return {
        getRefreshToken,
        getAccessToken,
    }
}


export interface ITenantClientId{
    tenantId: string;
    client_id: string;
    credentialsPath?: string;
}
export function getDefaultAuth(opt: ITenantClientId) {
    const { tenantId, client_id } = opt;
    const cpath = opt.credentialsPath || 'msgraph';
    return getAuth({
        tenantId,
        client_id,
        promptUser: msg => console.log(msg),
        saveToken: async res => fs.writeFileSync('credentials.json', JSON.stringify({ [cpath]:res }, null, 2)),
        loadToken: () => get(creds, cpath), //get(JSON.parse(fs.readFileSync('credentials.json').toString()),cpath),
    });
}