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
    if (action === "saveGuest") {
        const name = getPrm('name');
        const email = getPrm('email');
        const picture = getPrm('picture') || '';
        if (!name || !email) {
            responseMessage = 'Must have name or email'
        } else {
            responseMessage = `user ${name} Saved`;
            await store.addAndSave([name, email, picture], logger);
        }
    } else if (action === 'loadData') {
        responseMessage = await store.loadData(!!getPrm('force'),logger);
    } else if (action === 'loadImage') {
        context.res.setHeader("Content-Type", "image/png")
        const fname = checkFileName();
        if (!fname) {
            context.log('bad file name, return')
            return;
        }
        const ops = await getMsDirOpt();
        const ary = await ops.getFileByPath(fname).catch(err => {
            //console.log(err);
            context.log(`Load image error ${err.message}`);
            context.log(err.response?.data?.toString());
            context.log(`file is ${fname}, return empty since it has error`);
            return [];
        });
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
    } else if (action === 'saveImage') {
        const fname = checkFileName();
        if (!fname) return;
        let dataStr = getPrm('data') as string;
        const sub = dataStr.indexOf('base64,');
        if (sub > 0) {
            dataStr = dataStr.substring(sub + 7).trim();
        }
        const buf = Buffer.from(dataStr, 'base64');
        const ops = await getMsDirOpt();
        const res = await ops.createFile(fname, buf);
        context.res = {
            body: {
                id: res.id,
                file: res.file,
                size: res.size,
            }
        };
        return;
    } else if (action === 'createDir') {
        const fname = checkFileName();
        if (!fname) return;
        const path = getPrm('path');
        const ops = await getMsDirOpt();
        const res = await ops.createDir(path, fname);
        context.res = {
            body: {
                id: res.id,
                file: res.name,
                size: res.size,
            }
        };
    }
    

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;