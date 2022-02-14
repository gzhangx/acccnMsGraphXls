import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { getMsDir } from './src/lib/msdir';
import * as store from './src/store';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name]));
    const action = getPrm('action');

    //await store.getAllDataNoCache();
    let responseMessage = null;
    if (action === "saveGuest") {
        const name = getPrm('name');
        const email = getPrm('email');
        const picture = getPrm('picture') || '';
        if (!name || !email) {
            responseMessage = 'Must have name or email'
        } else {
            responseMessage = `user ${name} Saved`;
            await store.addAndSave([name, email, picture]);
        }
    } else if (action === 'loadData') {
        responseMessage = await store.loadData(!!getPrm('force'));
    } else if (action === 'loadImage') {
        context.res.setHeader("Content-Type", "image/png")
        let fname = getPrm('name');
        if (!fname) {
            context.res = {
                body: 'No filename',
            };        
            return;
        }
        fname = fname.replace(/[^a-z0-9-_ ]/gi, '');
        const ops = await getMsDir();
        const ary = await ops.getFileByPath(fname);
        return context.res.raw(ary);
    }
    

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;