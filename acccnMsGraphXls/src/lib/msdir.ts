import { getDefaultMsGraphConn } from "./msauth";


export interface IMsDirOps {
    doGet: (itemId: string, action: string) => Promise<any>;
    doPost: (itemId: string, action: string, data: object) => Promise<any>;
    doSearch: (itemId: string, name: string) => Promise<IFileSearchResult>;
    createFile: (path: string, data: Buffer) => Promise<any>;
    getFile: (itemId: string) => Promise<any>;
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
   
    async function createFile(path: string, data: Buffer): Promise<IFileCreateResponse> {
        return ops.doPut(`drive/root:/${path}:/content`, data);
    }

    async function getFile(itemId: string): Promise<any> {
        //01XX2KYFI2ZEYM7DGTM5FZGNFFNPF6DARZ
        return ops.doGet(getPostFix(itemId, 'content'));
    }

    return {
        doGet,
        doPost,
        doSearch: (itemId: string, name: string) => doSearch(itemId, name, doGet),
        createFile,
        getFile,
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



interface IFileCreateResponse {
    '@odata.context': string;
    '@microsoft.graph.downloadUrl': string;
    createdDateTime: string;
    eTag: string;
    id: string;
    lastModifiedDateTime: string;
    name: string;
    webUrl: string;
    cTag: string;
    size: number,
    createdBy: {
        application: {
            id: string;
            displayName: string;
        };
        user: {
            email: string;
            id: string;
            displayName: string;
        };
    };
    lastModifiedBy: {
        application: {
            id: string;
            displayName: string;
        };
        user: {
            email: string;
            id: string;
            displayName: string;
        };
    };
    parentReference: {
        driveId: string;
        driveType: string;
        id: string;
        path: string;
    },
    file: {
        mimeType: string; //'text/plain',
        hashes: { quickXorHash: string; };
    };
    fileSystemInfo: {
        createdDateTime: string;
        lastModifiedDateTime: string;
    }
}