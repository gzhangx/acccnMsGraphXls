import { getDefaultMsGraphConn } from "./msauth";


export interface IMsDirOps {
    doGet: (itemId: string, action: string) => Promise<any>;
    doPost: (itemId: string, action: string, data: object) => Promise<any>;
    doSearch: (itemId: string, name: string) => Promise<IFileSearchResult>;
    createFile: (path: string, data: Buffer) => Promise<any>;
}

export async function gtMsDir(): Promise<IMsDirOps> {
    const ops = await getDefaultMsGraphConn();
    const getPostFix = (itemId: string, action: string) => `/drive/items('${itemId}')/${action}`
    async function doGet(itemId: string, action: string) : Promise<any> {
        return ops.doGet(getPostFix(itemId, action));
    }

    async function doPost(itemId: string, action: string, data: object) {
        return ops.doPost(getPostFix(itemId, action), data);
    }    

    //const getDriveUrl = () => `https://graph.microsoft.com/v1.0/users('${opt.userId}')/drive`
    //const getUrl = (itemId: string, action: string) => `${getDriveUrl()}/items('${itemId}')/${action}`;
   
    async function createFile(path: string, data: Buffer) {
        return ops.doPut(`drive/root:/${path}:/content`, data);
    }
    return {
        doGet,
        doPost,
        doSearch: (itemId: string, name: string) => doSearch(itemId, name, doGet),
        createFile,
    }

}


export interface IFileSearchResult {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#Collection(driveItem)";
    value: {
        "@odata.type": "#microsoft.graph.driveItem";
        createdDateTime: string;
        id: string;
        lastModifiedDateTime: string;
        name: string;
        webUrl: string; //https://acccnusa-my.sharepoint.com/personal/gangzhang_acccn_org/Documents/NewUserImages,
        size: number;
        createdBy: {
            user: {
                email: string;
                displayName: string;
            };
        };
        lastModifiedBy: {
            user: {
                email: string;
                displayName: string;
            };
        };
        parentReference: {
            driveId: string;
            driveType: string; //"business",
            id: string;
        };
        fileSystemInfo: {
            createdDateTime: string;
            lastModifiedDateTime: string;
        };
        folder?: {
            childCount: number;
        };
        searchResult: object;
    }[];
}
async function doSearch(itemId: string, name: string, doGet: (itemId: string, action: string) => Promise<any>)
    : Promise<IFileSearchResult>{
    const res = await doGet(itemId, `search(q='${name}')`);
    return res as IFileSearchResult;
}

