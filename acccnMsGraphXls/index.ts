import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import * as store from './src/store';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name]));
    const action = getPrm('action');

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
        responseMessage = await store.loadData();
    }
    

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;