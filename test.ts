import { gtMsDir } from './acccnMsGraphXls/src/lib/msdir';

import creds from './acccnMsGraphXls/credentials.json'
async function test() {
    const dir = await gtMsDir();
    await dir.doSearch(creds.dirInfo.NewGuestImageDir, `new`).then(r => {
        //console.log(JSON.stringify(r,null,2));
        console.log(`len=${r.value.length} 0.size=${r.value[0].size} 0.size=${r.value[0].name} 0.folder.childcount=${r.value[0].folder?.childCount}` )
    })


    await dir.createFile('NewUserImages/test.txt', Buffer.from('testtest1234')).then(res => {
        console.log(res); 
    });

}

test();