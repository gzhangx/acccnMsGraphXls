import { getDefaultAuth, ITenantClientId, ITokenInfo } from "./msauth";
import Axios from 'axios';

export interface IMsGraphExcelItemOpt {
    tenantClientInfo: ITenantClientId;
    userId: string;
    itemId: string;
    tokenInfo?: ITokenInfo;
    sheetInfo?: IWorkSheetInfo;
}

interface IWorkSheetInfo {
    '@odata.context': string;
    value:
    {
        '@odata.id': string;
        id: string;
        name: string;
        position: number;
        visibility: string; //'Visible'
    }[];
    
}

interface IReadSheetValues {
    '@odata.context': string; //https://graph.microsoft.com/v1.0/$metadata#workbookRange
    '@odata.type': string; //'#microsoft.graph.workbookRange',
    '@odata.id': string;
    address: string; //'SheetName!A1:C12'
    addressLocal: string;
    columnCount: number;
    cellCount: number;
    columnHidden: boolean;
    rowHidden: boolean;
    numberFormat: string[][];
    columnIndex: number;
    text: string[][];
    formulas: string[][];
    formulasLocal: string[][];
    hidden: boolean;
    rowCount: number;
    rowIndex: number;
    valueTypes: string[][];
    values: string[][];
}

export interface IMsExcelOps {
    getWorkSheets: () => Promise<IWorkSheetInfo>;
    createSheet: (name: string) => Promise<any>;
    readRange: (name: string, from: string, to: string) => Promise<IReadSheetValues>;
    getRangeFormat: (name: string, from: string, to: string) => Promise<IReadSheetValues>;
    updateRange: (name: string, from: string, to: string, values: string[][]) => Promise<IReadSheetValues>;
}

export async function getMsExcel(opt: IMsGraphExcelItemOpt): Promise<IMsExcelOps> {
    const now = new Date().getTime();
    async function getToken(): Promise<ITokenInfo> {
        if (!opt.tokenInfo || opt.tokenInfo.expires_on < now / 1000) {
            const { getAccessToken } = getDefaultAuth(opt.tenantClientInfo);
            console.log('getting new token');
            const tok = await getAccessToken();
            opt.tokenInfo = tok;
        }
        return opt.tokenInfo;
    }

    async function getHeaders() {
        const tok = await getToken();
        return {
            headers: {
                "Authorization": `Bearer ${tok.access_token}`
            }
        };
    }

    function parseResp(r: { data: any }) {
        return r.data;
    }
    async function doGet(url: string) {
        return await Axios.get(getUrl(url), await getHeaders())
            .then(parseResp);
    }

    async function doPost(postFix: string, data: object) {
        return Axios.post(getUrl(postFix), data, await getHeaders()).then(parseResp);
    }

    async function doPatch(postFix: string, data: object) {
        return Axios.patch(getUrl(postFix), data, await getHeaders())
            .then(parseResp);
    }

    const getUrl = (postFix: string) => `https://graph.microsoft.com/v1.0/users('${opt.userId}')/drive/items('${opt.itemId}')/workbook/worksheets${postFix}`;
    async function getWorkSheets(): Promise<IWorkSheetInfo> {
        return await doGet('');
    }

    async function createSheet(name: string): Promise<any> {
        if (!opt.sheetInfo) {
            opt.sheetInfo = await getWorkSheets();
        }
        const found = (opt.sheetInfo.value.find(v => v.name === name));
        if (found) return found;
        return await doPost('', {
            name
        });
    }

    async function readRange(name: string, from: string, to: string): Promise<IReadSheetValues> {
        return doGet((`/${name}/range(address='${from}:${to}')`));
    }

    async function getRangeFormat(name: string, from: string, to: string): Promise<IReadSheetValues> {
        return doGet((`/${name}/range(address='${from}:${to}')/format`));
    }

    async function updateRange(name: string, from: string, to: string, values: string[][]): Promise<IReadSheetValues> {
        return doPatch((`/${name}/range(address='${from}:${to}')`), {
            values,
        });
    }

    return {
        getWorkSheets,
        createSheet,
        readRange,
        getRangeFormat,
        updateRange,
    }

}