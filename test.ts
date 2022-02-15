import { getMsDir } from './acccnMsGraphXls/src/lib/msdir';
import { getDefaultMsGraphConfig } from './acccnMsGraphXls/src/store';
import { getMsExcel } from './acccnMsGraphXls/src/lib/msExcell';

//import creds from './acccnMsGraphXls/credentials.json'
const fs = require('fs');
async function test1() {
    const dir = await getMsDir(getDefaultMsGraphConfig(), msg => {
        console.log(msg);
    });
    //await dir.doSearch(creds.dirInfo.NewGuestImageDir, `new`).then(r => {
    //    //console.log(JSON.stringify(r,null,2));
    //    console.log(`len=${r.value.length} 0.size=${r.value[0].size} 0.size=${r.value[0].name} 0.folder.childcount=${r.value[0].folder?.childCount}` )
    //})


    //const buf = fs.readFileSync('d:/temp/IMG_0158.JPG');
    //await dir.createFile('NewUserImages/IMG_0158.jpg', buf).then(res => {
    //    console.log(res); 
    //});

    //await dir.getFileById('01XX2KYFI2ZEYM7DGTM5FZGNFFNPF6DARZ').then(res => {
    //    console.log(res instanceof Buffer);
    //    console.log(res.length);
    //    fs.writeFileSync('d:\\temp\\testtest_byId.jpg', res);
    //})

    console.log('getting img');
    await dir.getFileByPath('NewUserImages/IMG_0158.jpg').then(res => {
        console.log(res instanceof Buffer);
        console.log(res.length);
        fs.writeFileSync('d:\\temp\\testtest_bypath.jpg', res);
    }).catch(err => {
        console.log(err.message)
        console.log(err.response)
    });

    await dir.createDir("NewUserImages", "dir111").then(res => {
        console.log(res);
    }).catch(err => {
        console.log(err.message)
        console.log(err.response)
    })

}


async function testExcell() {

    const ops  = await getMsExcel({
        itemId: '01XX2KYFM4CINDUVRDIJGICH2EHDH5G3EY',
        tenantClientInfo: getDefaultMsGraphConfig(),
    }, msg=>console.log(msg));
    await ops.createSheet('2022-01=tet');
}
testExcell();