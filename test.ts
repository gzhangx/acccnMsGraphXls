import { getMsDir, IMsGraphOps } from './acccnMsGraphXls/src/lib/msdir';
import { getDefaultMsGraphConfig } from './acccnMsGraphXls/src/store';
import { getMsExcel } from './acccnMsGraphXls/src/lib/msExcell';
import * as auth from './acccnMsGraphXls/src/lib/msauth';
import * as store from './acccnMsGraphXls/src/store';

//import creds from './acccnMsGraphXls/credentials.json'
const fs = require('fs');

const driveId = 'b!hXChu0dhsUaKN7pqt1bD3_OeafGaVT1FohEO2dBMjAY5XO0eLYVxS7CH5lgurhQd';
const msDirPrm: IMsGraphOps = {
    logger: msg => console.log(msg),
    driveId,
};
function getDefaultDirOpt(): IMsGraphOps {
    return {
        logger: msg => console.log(msg),
        driveId: '',
        sharedUrl: 'https://acccnusa.sharepoint.com/:f:/r/sites/newcomer/Shared%20Documents/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99?csf=1&web=1&e=8k5NUF',
    }
}

async function getDriveIdOfUrl(url) {    
    const aut = await auth.getDefaultMsGraphConn(getDefaultMsGraphConfig());
    await aut.getSharedItemInfo(url).then(res => {
        console.log(`Drive id=${res.parentReference.driveId}`);
        console.log(`id=${res.id} ${res.name}`);

    });
}

getDriveIdOfUrl('https://acccnusa-my.sharepoint.com/:x:/r/personal/gangzhang_acccn_org/_layouts/15/Doc.aspx?sourcedoc=%7B3A1A129C-2356-4C42-811F-4438CFD36C98%7D&file=AcccnNewGuests.xlsx&action=default&mobileredirect=true')
    .then(async () => {
        await testExcellOld();
    }).catch(err => {
        console.log(err);
})

async function testPathFile() {

    //const aut = await auth.getDefaultMsGraphConn(getDefaultMsGraphConfig());
    //await aut.getSharedItemInfo('https://acccnusa.sharepoint.com/:f:/r/sites/newcomer/Shared%20Documents/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99?csf=1&web=1&e=8k5NUF').then(res => {
    //    console.log(res.parentReference.driveId);
    //    console.log(`id=${res.id}`);
    //})
    const dir = await getMsDir(getDefaultMsGraphConfig(), getDefaultDirOpt());
    
    await dir.createDir("新人资料","tempNewUserImages");
    await dir.createFile("新人资料/tempNewUserImages/testdir/test.text", Buffer.from('testtest'));


}



//testPathFile().catch(err => {console.log(err);})

async function test1() {
    const dir = await getMsDir(getDefaultMsGraphConfig(), getDefaultDirOpt());
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


async function testExcellOld() {

    const ops  = await getMsExcel({
        fileName: '新人资料/新人资料表汇总new.xlsx',
        //fileName:'AcccnNewGuests.xlsx',
        tenantClientInfo: getDefaultMsGraphConfig(),
    }, {
        ...msDirPrm,
        //driveId: 'b!38QulSRtuEmvnv_ky3EDLMmYEAkrtktPhNRbouRgk9FGZ0JsesssSZRVEVKm90fq'
    });
    await ops.createSheet('2022-01=tet');
}



async function testExcell() {

    const testerr = {
        message: 'msg1',
        response: {
            data: {
                message: 'datamsg',
                error: {
                    message:'errmsg'
                }
            }
        }
    }

    
    const rr = auth.axiosErrorProcessing(testerr);
    if (rr) {
        return console.log(rr);
    }
    
    await store.loadData(msDirPrm);
    await store.addAndSave(['test1', 'test2', 'test3'], msDirPrm).catch(err => {
        console.log(err.message);
        console.log(Object.keys(err));
        console.log(err.isAxiosError);
        console.log(err.response?.data);
    })
}
//testExcell();