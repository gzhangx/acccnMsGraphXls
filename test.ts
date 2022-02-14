import { gtMsDir } from './acccnMsGraphXls/src/lib/msdir';

import creds from './acccnMsGraphXls/credentials.json'
async function test() {
    const dir = await gtMsDir({
        tenantClientInfo: {
            tenantId: creds.gzuser.tenantId,
            client_id: creds.gzuser.client_id,
        },
        userId: creds.gzuser.userId,
    });
    dir.doSearch(creds.dirInfo.NewGuestImageDir, `new`).then(r => {
        //console.log(JSON.stringify(r,null,2));
        console.log(`len=${r.value.length} 0.size=${r.value[0].size} 0.size=${r.value[0].name} 0.folder.childcount=${r.value[0].folder?.childCount}` )
    })

}

test();