import { IMsGraphCreds, ILogger } from "./lib/msauth";
import { IMsExcelOps, getMsExcel, IReadSheetValues } from "./lib/msExcell";
//import * as creds from '../credentials.json';
import moment from 'moment';

let msExcelOps: IMsExcelOps = null;
let curSheetData: string[][] = null;

// gzuser.guestSheetId
export function getDefaultMsGraphConfig(): IMsGraphCreds {
    return {
        client_id: process.env['msgp1.CLIENT_ID'],
        refresh_token: process.env['msgp1.refresh_token'],
        tenantId: process.env['msgp1.tenantId'],
        userId: process.env['msgp1.userID'],
    }
}


async function createMsOps(logger: ILogger) {
    if (!msExcelOps) {
        msExcelOps = await getMsExcel({
            itemId: process.env['gzuser.guestSheetId'],
            tenantClientInfo: getDefaultMsGraphConfig(),
        }, logger);

        const today = getToday();
        logger(`creating today ${today}`);
        await msExcelOps.createSheet(today).catch(err => {
            logger(`error creating today ${today}: ${err.message}`);
            logger(err);
        })
    }
    return msExcelOps;
}

function getToday(): string {
    const today = moment().format('YYYY-MM-DD');
    return today;
}
export async function getAllDataNoCache(logger:(msg: string) => void) {
    const today = getToday();
    const ops = await createMsOps(logger);
    const MAX = 10;
    curSheetData = [];
    for (let from = 0; ; from += MAX) {
        const cur = await (await ops.readRange(today, `A${from + 1}`, `C${from + MAX}`)).values;
        const amt = cur.reduce((acc, v) => {
            if (!v[0]) return acc;
            acc++;
            curSheetData.push(v);
            return acc;
        }, 0);
        if (!amt) break;
    }
}

export async function loadData(force:boolean, logger: (msg:string)=>void): Promise<string[][]> {
    if (!curSheetData || force) {
        await getAllDataNoCache(logger);
    }
    return curSheetData;
}

export async function saveData(logger: ILogger): Promise<IReadSheetValues> {
    const ops = await createMsOps(logger);
    const today = getToday();
    return await ops.updateRange(today, `A1`, `C${curSheetData.length}`, curSheetData);
}

export async function addAndSave(ary: string[], logger: ILogger): Promise<any> {
    let curSheetData = await loadData(false, logger);
    curSheetData.push(ary);
    return await saveData(logger);
}
