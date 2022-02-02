import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import * as creds from './credentials.json';
import { getMsExcel } from './src/lib/msExcell';
import { store } from './src/store';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name]));
    const name = getPrm('name');
    const action = getPrm('action');

    
    const ops = await store.createMsOps();
    ops.


    const responseMessage = name
        ? "Hello, " + name + ". This HTTP triggered function executed successfully."
        : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;