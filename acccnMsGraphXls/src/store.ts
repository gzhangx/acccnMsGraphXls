import { IMsGraphCreds, ILogger } from "./lib/msauth";
import { IMsExcelOps, getMsExcel, IReadSheetValues } from "./lib/msExcell";
import { IMsGraphOps } from './lib/msdir';

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
        //userId: process.env['msgp1.userID'],
    }
}


async function createMsOps(prm: IMsGraphOps) {
    if (!msExcelOps) {
        msExcelOps = await getMsExcel({
            fileName:'新人资料/新人资料表汇总new.xlsx',
            tenantClientInfo: getDefaultMsGraphConfig(),
        }, prm);

        const today = getToday();
        prm.logger(`creating today ${today}`);
        await msExcelOps.createSheet(today).catch(err => {
            prm.logger(`error creating today ${today}: ${err.message}`);
            prm.logger(err);
        })
    }
    return msExcelOps;
}

function getToday(): string {
    const today = moment().format('YYYY-MM-DD');
    return today;
}
export async function getAllDataNoCache(prm: IMsGraphOps) {
    const today = getToday();
    const ops = await createMsOps(prm);
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

export async function loadData(prm: IMsGraphOps): Promise<string[][]> {
        await getAllDataNoCache(prm);
    return curSheetData;
}

export async function saveData(prm:IMsGraphOps): Promise<IReadSheetValues> {
    const ops = await createMsOps(prm);
    const today = getToday();
    return await ops.updateRange(today, `A1`, `C${curSheetData.length}`, curSheetData);
}

export async function addAndSave(ary: string[], prm: IMsGraphOps): Promise<any> {
    const curSheetData = await loadData(prm);
    let found = false;
    curSheetData.forEach(s => {
        if (s[0] === ary[0]) {
            found = true;
            for (let i = 0; i < s.length; i++) {
                s[i] = ary[i];
            }
        } 
    });
    if (!found) curSheetData.push(ary);
    return await saveData(prm);
}
