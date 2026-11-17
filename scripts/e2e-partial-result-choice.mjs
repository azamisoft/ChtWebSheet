// Deterministic UI acceptance: actual CWS host, no inference/provider requests.
import {chromium} from '@playwright/test';
import {readFile,mkdir,writeFile} from 'node:fs/promises';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import {createRequire} from 'node:module';
const require=createRequire(new URL('../cws-agent/package.json',import.meta.url));
const {rolldown}=await import(require.resolve('rolldown'));
const appUrl=process.env.CWS_PARTIAL_URL || 'http://127.0.0.1:8774/';
const appHost=new URL(appUrl).hostname;
const out=process.env.CWS_PARTIAL_EVIDENCE || '/tmp/cws-partial-result-e2e';await mkdir(out,{recursive:true});
const workbook=new ExcelJS.Workbook();const sheet=workbook.addWorksheet('Inventory');sheet.getCell('A1').value='変更前';sheet.getCell('A2').value='領収書 合計';sheet.getCell('B2').value=1234;await workbook.xlsx.writeFile(`${out}/baseline.xlsx`);
const instrumentation=`
  window.partialTest={
    setup:()=>{Object.assign(aiState(),{user:{id:99117,displayName:'テスト利用者',isDeveloper:false},open:true,tab:'chat',activeSessionId:'99117',sessionsLoaded:true,historyLoaded:true,aiStatus:'ready',panelWidth:370,panelHeight:620});renderCwsAiPanel();},
    create:async()=>{
      const baseline=runtime.operations.captureTaskWorkbook();
      const record={key:cwsAiRecoveryKey(),operationId:'partial-'+Date.now(),sessionId:cwsAiActiveSessionId(),baseline,requestMessage:'領収書一覧を作成',requestHash:'fixture',workbookKey:'fixture'};
      const ops=[{type:'setCell',sheetName:'Inventory',cell:'A1',value:'変更後'}];
      const result=runtime.operations.previewOps(ops);
      const message={id:record.operationId,operationId:record.operationId,role:'assistant',text:'作業は未完了です。',ops,previewId:result.previewId,previewVisible:true,taskStatus:'incomplete',pending:false,operationFinalized:true,opValidation:{valid:false,message:'税額の再確認が必要です。'},visualIssues:[{sheetName:'領収書',reason:'重複候補が1件あります。'}]};
      aiState().messages.push(message);
      await cwsAiOfferPartialResult(message,record);renderCwsAiPanel();return record.operationId;
    },
    debug:()=>({session:cwsAiActiveSessionId(),recordSession:cwsAiRecoveryCache.record?.sessionId,key:cwsAiRecoveryKey(),recordKey:cwsAiRecoveryCache.record?.key,compare:cwsAiPartialState.comparison ? {shown:cwsAiPartialState.comparison.shown,released:cwsAiPartialState.comparison.released,ending:!!cwsAiPartialState.comparison.ending}:null,error:aiState().statusMessage,decisionError:cwsAiRecoveryCache.record?.userDecision?.error}),
    read:()=>({value:runtime.operations.captureTaskWorkbook().sheets[0].cells['1:1']?.raw,record:cwsAiRecoveryCache.record?.userDecision,status:aiState().messages.at(-1)?.taskStatus,done:aiState().messages.at(-1)?.activityDone}),
    reloadDecision:async()=>{aiState().messages=[];cwsAiRecoveryCache={key:'',record:null};renderCwsAiPanel();},
    restoreBaseline:async()=>{await runtime.operations.restoreTaskWorkbook(cwsAiRecoveryCache.record.baseline);},
    recover:async()=>{await cwsAiRecoveryDb('readonly',store=>store.get(cwsAiRecoveryKey()));renderCwsAiPanel();},
  };
`;
const bundle=await rolldown({input:'cws-agent/src/agent.js',plugins:[{name:'acceptance-hooks',transform(code,id){if(id.endsWith('/src/modules/history-and-task-state.js'))return {code:code.replace(/async function refreshCwsAiUser\(([^)]*)\) \{/, 'async function refreshCwsAiUser($1) { return;'),map:null};if(!id.endsWith('/src/agent-core.js'))return;return {code:"import {cwsAiOfferPartialResult} from './modules/partial-results.js';\nimport {cwsAiActiveSessionId} from './modules/panel-and-sessions.js';\n"+code.replace('registerCwsAgent();',instrumentation+'\nregisterCwsAgent();'),map:null};}}]});
const {output}=await bundle.generate({format:'iife'});const injected=output[0].code;await bundle.close();
const browser=await chromium.launch({headless:true});const context=await browser.newContext({viewport:{width:1280,height:820}});const page=await context.newPage();const errors=[];let inference=0;page.setDefaultTimeout(8000);
page.on('pageerror',e=>errors.push(e.message));
await page.route('**/*',async route=>{
 const u=new URL(route.request().url());
 if(u.pathname.endsWith('/cws-agent/agent.js'))return route.fulfill({contentType:'text/javascript',body:injected});
 if(u.pathname.endsWith('/cws-agent/agent.css'))return route.fulfill({contentType:'text/css',body:await readFile('cws-agent/src/agent.css','utf8')});
 if(u.pathname.includes('/api/')){if(/run|route|plan|review/.test(u.pathname))inference++;return route.fulfill({contentType:'application/json',body:JSON.stringify({ok:true,user:{id:99117,displayName:'テスト利用者',isDeveloper:false},sessions:[],messages:[],quota:null})});}
 if(u.hostname!==appHost)return route.abort();
 return route.continue();
});
try{
 await page.goto(appUrl);
 await page.setInputFiles('#fileInput',`${out}/baseline.xlsx`);
 await page.waitForFunction(()=>/読み込みました|読み込み済み|已加载/.test(document.querySelector('#statusText')?.textContent || ''));
 await page.waitForFunction(()=>!window.ChtWebSheetAgentRuntime?.state?.persistenceLoading);
 await page.waitForFunction(()=>window.ChtWebSheetAgentRuntime?.state?.model && window.partialTest);
 await page.evaluate(()=>window.partialTest.setup());
 await page.evaluate(()=>window.partialTest.create());
 const choice=page.locator('[data-cws-ai-partial-choice="kept"]');await choice.waitFor();
 assert.equal((await page.evaluate(()=>window.partialTest.read())).value,'変更後');
 const compare=page.locator('[data-cws-ai-compare]');await compare.hover();await page.mouse.down();
 await page.waitForFunction(()=>window.partialTest.read().value==='変更前');await page.screenshot({path:`${out}/before.png`});
 await page.mouse.up();await page.waitForFunction(()=>window.partialTest.read().value==='変更後');await page.screenshot({path:`${out}/after.png`});
 await compare.focus();await page.keyboard.down('Space');await page.waitForFunction(()=>window.partialTest.read().value==='変更前');await page.keyboard.up('Space');await page.waitForFunction(()=>window.partialTest.read().value==='変更後');
 await page.reload();await page.waitForFunction(()=>window.ChtWebSheetAgentRuntime?.state?.model && window.partialTest);await page.evaluate(()=>window.partialTest.setup());
 await choice.waitFor();await page.waitForFunction(()=>window.partialTest.read().record?.status==='pending');assert.equal((await page.evaluate(()=>window.partialTest.read())).value,'変更後');assert.equal(await page.locator('[data-cws-ai-compare]').count(),1);
 await choice.click();await page.waitForFunction(()=>window.partialTest.read().record?.status==='kept');assert.equal((await page.evaluate(()=>window.partialTest.read())).value,'変更後');
 await page.screenshot({path:`${out}/kept.png`});
 // Full reload restores real workbook persistence and the IndexedDB decision.
 await page.reload();await page.waitForFunction(()=>window.ChtWebSheetAgentRuntime?.state?.model && window.partialTest);await page.evaluate(()=>window.partialTest.setup());
 await page.waitForFunction(()=>window.partialTest.read().record?.status==='kept');assert.equal((await page.evaluate(()=>window.partialTest.read())).value,'変更後');
 await page.evaluate(()=>window.partialTest.restoreBaseline());await page.evaluate(()=>window.partialTest.create());
 await page.locator('[data-cws-ai-partial-choice="cancelled"]').click();await page.waitForFunction(()=>window.partialTest.read().record?.status==='cancelled');
 const cancelled=await page.evaluate(()=>window.partialTest.read());assert.equal(cancelled.value,'変更前');assert.equal(cancelled.done,false);assert.equal(cancelled.status,'incomplete');
 await page.screenshot({path:`${out}/cancelled.png`});
 await page.reload();await page.waitForFunction(()=>window.ChtWebSheetAgentRuntime?.state?.model && window.partialTest);await page.evaluate(()=>window.partialTest.setup());await page.waitForFunction(()=>window.partialTest.read().record?.status==='cancelled');assert.equal((await page.evaluate(()=>window.partialTest.read())).value,'変更前');
 assert.equal(inference,0);assert.deepEqual(errors,[]);
 const result={ok:true,checks:['native worksheet before/after','pointer press/release','keyboard press/release','decision reload','keep persisted across page reload','cancel exact baseline persisted across reload','nondeveloper failure details','incomplete guard'],inferenceRequests:inference,errors};await writeFile(`${out}/result.json`,JSON.stringify(result,null,2));console.log(JSON.stringify(result));
}catch(error){await page.screenshot({path:`${out}/failure.png`}).catch(()=>{});console.error('UI state',await page.evaluate(()=>({debug:window.partialTest?.debug(),value:window.partialTest?.read().value,status:window.partialTest?.read().record?.status,ui:document.querySelector('#cwsAiPanelRoot')?.innerText.slice(-1600)})));throw error;}finally{await browser.close();}
