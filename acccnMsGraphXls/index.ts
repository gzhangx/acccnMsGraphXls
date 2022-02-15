import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getMsDir } from './src/lib/msdir';
import * as store from './src/store';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name]));
    const action = getPrm('action');

    function checkFileName() {
        const fname = getPrm('name');
        if (!fname) {
            context.res = {
                body: 'No filename',
            };        
            return null;
        }
        return fname.replace(/[^a-z0-9-_\/ \.]/gi, '');
    }
    //await store.getAllDataNoCache();
    let responseMessage = null;

    function logger(msg: string) {
        context.log(msg);
    }
    async function getMsDirOpt() {
        const ops = await getMsDir(store.getDefaultMsGraphConfig(), logger);
        return ops;
    }

    type IDataWithError = {
        error?: string;
        length?: number;
    }

    function returnError(msg) {
        context.log(msg);
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: msg
        };
    }
    function getErrorHndl(inf: string) {
        return (err): IDataWithError => {
            responseMessage = {
                error: `${inf} ${err.message}`
            }
            logger(responseMessage.error);
            return responseMessage;
        }
    }
    if (action === "saveGuest") {
        const name = getPrm('name');
        const email = getPrm('email');
        const picture = getPrm('picture') || '';
        context.log(`saveGuest for ${name}:${email}`);
        if (!name || !email) {
            return returnError('Must have name or email');
        } else {
            responseMessage = `user ${name} Saved`;
            await store.addAndSave([name, email, picture], logger).catch(getErrorHndl(`user save error for ${name}:${email}`));
        }
    } else if (action === 'loadData') {
        responseMessage = await store.loadData(!!getPrm('force'), logger).catch(getErrorHndl('loadData Error'));
    } else if (action === 'loadImage') {
        context.res.setHeader("Content-Type", "image/png")
        const fname = checkFileName();
        if (!fname) {            
            return returnError('bad file name, return')
        }
        const ops = await getMsDirOpt();
        const ary = await ops.getFileByPath(fname).then(r => {
            return {
                array: r,
            } as IDataWithError;
        }).catch(getErrorHndl(`unable to load image ${fname}`)) as IDataWithError;
        if (ary.error) {
            responseMessage = ary.error;
        } else {
            context.log(`image size ${ary.length}`)
            context.res = {
                headers: {
                    "Content-Type": "image/png"
                },
                isRaw: true,
                // status: 200, /* Defaults to 200 */
                body: ary, //new Uint8Array(buffer)
            };
            return;
        }        
    } else if (action === 'saveImage') {
        const fname = checkFileName();
        if (!fname) return returnError('No filename for saveImage');
        let dataStr = getPrm('data') as string;
        const sub = dataStr.indexOf('base64,');
        if (sub > 0) {
            dataStr = dataStr.substring(sub + 7).trim();
        }
        const buf = Buffer.from(dataStr, 'base64');
        const ops = await getMsDirOpt();
        try {
            const res = await ops.createFile(fname, buf);
            context.res = {
                body: {
                    id: res.id,
                    file: res.file,
                    size: res.size,
                }
            };
        } catch (err) {
            getErrorHndl(`saveImage createFile error for ${fname} ${buf.length}`)(err);
        }
        return;
    } else if (action === 'createDir') {
        const fname = checkFileName();
        if (!fname) return returnError('createDir: null file name');
        const path = getPrm('path') || '';
        const ops = await getMsDirOpt();
        try {
            const res = await ops.createDir(path, fname);
            context.res = {
                body: {
                    id: res.id,
                    file: res.name,
                    size: res.size,
                }
            };
        } catch (err) {
            getErrorHndl(`createDir error for ${fname} ${path}`)(err);
        }
    }
    

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;