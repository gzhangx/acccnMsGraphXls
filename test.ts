import { gtMsDir } from './acccnMsGraphXls/src/lib/msdir';

import creds from './acccnMsGraphXls/credentials.json'
const fs = require('fs');
async function test() {
    const dir = await gtMsDir();
    await dir.doSearch(creds.dirInfo.NewGuestImageDir, `new`).then(r => {
        //console.log(JSON.stringify(r,null,2));
        console.log(`len=${r.value.length} 0.size=${r.value[0].size} 0.size=${r.value[0].name} 0.folder.childcount=${r.value[0].folder?.childCount}` )
    })


    const buf = fs.readFileSync('d:/temp/IMG_0158.JPG');
    await dir.createFile('NewUserImages/IMG_0158.jpg', buf).then(res => {
        console.log(res); 
    });

    await dir.getFileById('01XX2KYFI2ZEYM7DGTM5FZGNFFNPF6DARZ').then(res => {
        console.log(res instanceof Buffer);
        console.log(res.length);
        fs.writeFileSync('d:\\temp\\testtest_byId.jpg', res);
    })

    await dir.getFileByPath('NewUserImages/IMG_0158.jpg').then(res => {
        console.log(res instanceof Buffer);
        console.log(res.length);
        fs.writeFileSync('d:\\temp\\testtest_bypath.jpg', res);
    })

}

test();