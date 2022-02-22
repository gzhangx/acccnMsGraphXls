import { getMsDir } from './acccnMsGraphXls/src/lib/msdir';
import { getDefaultMsGraphConfig } from './acccnMsGraphXls/src/store';
import { getMsExcel } from './acccnMsGraphXls/src/lib/msExcell';
import * as auth from './acccnMsGraphXls/src/lib/msauth';
import * as store from './acccnMsGraphXls/src/store';

//import creds from './acccnMsGraphXls/credentials.json'
const fs = require('fs');


async function testPathFile() {
    const dir = await getMsDir(getDefaultMsGraphConfig(), msg => {
        console.log(msg);
    });
    
    //await dir.createFile("NewUserImages/test.text", Buffer.from('testtest'));

    const sharingUrl = 'https://acccnusa.sharepoint.com/:f:/r/sites/newcomer/Shared%20Documents/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99?csf=1&web=1&e=aGy7vS';
    //see https://docs.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0&irgwc=1&OCID=AID2200057_aff_7593_1243925&tduid=(ir__ksd0kmgl9ckf6nyskg6fwnqce32xt3umkhw9f9gn00)(7593)(1243925)(je6NUbpObpQ-XTpQa0NuXTfWX1VU38TMYg)()&irclickid=_ksd0kmgl9ckf6nyskg6fwnqce32xt3umkhw9f9gn00&tabs=http#encoding-sharing-urls&ranMID=24542&ranEAID=je6NUbpObpQ&ranSiteID=je6NUbpObpQ-XTpQa0NuXTfWX1VU38TMYg&epi=je6NUbpObpQ-XTpQa0NuXTfWX1VU38TMYg   
    const base64Value = Buffer.from(sharingUrl).toString('base64');
    console.log(base64Value);
    console.log(Buffer.from(base64Value,'base64').toString())
    //string encodedUrl = "u!" + base64Value .TrimEnd('=').Replace('/', '_').Replace('+', '-');
    const encodedUrl = base64Value.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
    const resUrl = encodeSharedUrl(sharingUrl);
    console.log(resUrl);


}

function encodeSharedUrl(sharingUrl: string) : string {    
    //see https://docs.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0&irgwc=1&OCID=AID2200057_aff_7593_1243925&tduid=(ir__ksd0kmgl9ckf6nyskg6fwnqce32xt3umkhw9f9gn00)(7593)(1243925)(je6NUbpObpQ-XTpQa0NuXTfWX1VU38TMYg)()&irclickid=_ksd0kmgl9ckf6nyskg6fwnqce32xt3umkhw9f9gn00&tabs=http#encoding-sharing-urls&ranMID=24542&ranEAID=je6NUbpObpQ&ranSiteID=je6NUbpObpQ-XTpQa0NuXTfWX1VU38TMYg&epi=je6NUbpObpQ-XTpQa0NuXTfWX1VU38TMYg   
    const base64Value = Buffer.from(sharingUrl).toString('base64');    
    //string encodedUrl = "u!" + base64Value .TrimEnd('=').Replace('/', '_').Replace('+', '-');
    const encodedUrl = base64Value.replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
    const resUrl = `u!${encodedUrl}`;
    return resUrl;
}

testPathFile().catch(err => {
    console.log(err);
 })

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


async function testExcellOld() {

    const ops  = await getMsExcel({
        itemId: '01XX2KYFM4CINDUVRDIJGICH2EHDH5G3EY',
        tenantClientInfo: getDefaultMsGraphConfig(),
    }, msg=>console.log(msg));
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
    const logger = (msg: string) => console.log(msg);
    await store.loadData(true, logger);
    await store.addAndSave(['test1', 'test2', 'test3'], logger).catch(err => {
        console.log(err.message);
        console.log(Object.keys(err));
        console.log(err.isAxiosError);
        console.log(err.response?.data);
    })
}
//testExcell();