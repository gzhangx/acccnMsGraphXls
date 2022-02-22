import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getMsDir, IMsGraphOps } from './src/lib/msdir';
import * as store from './src/store';

const driveId = 'b!hXChu0dhsUaKN7pqt1bD3_OeafGaVT1FohEO2dBMjAY5XO0eLYVxS7CH5lgurhQd';
const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name]));
    const action = getPrm('action');

    context.log(`action=${action}`);
    function checkFileName() {
        const fname = getPrm('name');
        if (!fname || !fname.trim()) {
            context.res = {
                body: 'No filename',
            };        
            return null;
        }
        return fname;
    }
    //await store.getAllDataNoCache();
    let responseMessage = null;
    
    const msDirPrm: IMsGraphOps = {
        logger: msg=>context.log(msg),
        driveId,
    };
    async function getMsDirOpt() {
        const ops = await getMsDir(store.getDefaultMsGraphConfig(), msDirPrm);
        return ops;
    }

    type IDataWithError = {
        error?: string;
        length?: number;
    }

    function returnError(error) {
        context.log(error);
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: {
                error: error
            }
        };
    }
    function getErrorHndl(inf: string) {
        return (err): IDataWithError => {
            responseMessage = {
                error: `${inf} ${err.message}`
            }
            context.log(responseMessage.error);
            return responseMessage;
        }
    }
    if (action === "saveGuest") {
        const name = getPrm('name');
        const email = getPrm('email') || '';
        const picture = getPrm('picture') || '';
        context.log(`saveGuest for ${name}:${email}`);
        if (!name) {
            return returnError('Must have name or email');
        } else {
            responseMessage = `user ${name} Saved`;
            await store.addAndSave([name, email, picture], msDirPrm).catch(getErrorHndl(`user save error for ${name}:${email}`));
        }
    } else if (action === 'loadData') {
        responseMessage = await store.loadData(msDirPrm).catch(getErrorHndl('loadData Error'));
    } else if (action === 'loadImage') {
        context.res.setHeader("Content-Type", "image/png")
        const fname = checkFileName();
        if (!fname) {            
            return returnError('bad file name, return')
        }
        const ops = await getMsDirOpt();
        const ary = await ops.getFileByPath(fname).catch(getErrorHndl(`unable to load image ${fname}`)) as IDataWithError;
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
    } 
    

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;