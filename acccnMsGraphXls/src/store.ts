import { IMsGraphCreds } from "./lib/msauth";
import { IMsExcelOps, getMsExcel } from "./lib/msExcell";
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


async function createMsOps() {
    if (!msExcelOps) {
        msExcelOps = await getMsExcel({
            itemId: process.env['gzuser.guestSheetId'],
            tenantClientInfo: getDefaultMsGraphConfig(),
        });

        await msExcelOps.createSheet(getToday());
    }
    return msExcelOps;
}

function getToday(): string {
    const today = moment().format('YYYY-MM-DD');
    return today;
}
export async function getAllDataNoCache() {
    const today = getToday();
    const ops = await createMsOps();
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

export async function loadData(force:boolean): Promise<string[][]> {
    if (!curSheetData || force) {
        await getAllDataNoCache();
    }
    return curSheetData;
}

export async function saveData(): Promise<void> {
    const ops = await createMsOps();
    const today = getToday();
    await ops.updateRange(today, `A1`, `C${curSheetData.length }`, curSheetData).catch(err => {
        console.log(err);
        throw err;
    })
}

export async function addAndSave(ary: string[]): Promise<void> {
    let curSheetData = await loadData(false);
    curSheetData.push(ary);
    await saveData();
}
