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
    dir.doGet(creds.dirInfo.NewGuestImageDir, `search(q='new')`).then(r => {
        console.log(r);
    })

}

test();