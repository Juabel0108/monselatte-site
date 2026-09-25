/* Local regression for B1/B2/B3/B7. All submissions and analytics blocked.
 * NODE_PATH=<Playwright runtime> node tests/audit-b-regression.test.js
 */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');
const base = process.env.WIZARD_TEST_URL || 'http://127.0.0.1:8905/';
const out = process.env.AUDIT_SCREENSHOTS || '/tmp/monselatte-b-fixes/screenshots';
const viewports = [[320,568],[375,812],[390,844],[768,1024],[1024,768],[1280,800],[1440,900],[1920,1080]];
const settle = page => page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
(async()=>{
  fs.mkdirSync(out,{recursive:true});
  const browser=await chromium.launch({channel:'chrome',headless:true});
  const report=[]; const errors=[];
  try {
    for(const dpr of [1,2]) for(const [width,height] of viewports) {
      const context=await browser.newContext({viewport:{width,height},deviceScaleFactor:dpr});
      await context.route('**/*',route=>{
        const url=new URL(route.request().url());
        if(url.hostname==='127.0.0.1'||url.hostname.includes('fonts.g')) return route.continue();
        if(url.hostname.includes('googletagmanager')) return route.fulfill({body:'',contentType:'application/javascript'});
        return route.abort();
      });
      await context.addInitScript(()=>{
        window.fetch=()=>{throw Error('Real submission forbidden');};
        window.__heroShifts=[];
        new PerformanceObserver(list=>{
          for(const e of list.getEntries()) if(e.sources?.some(s=>s.node?.closest?.('.monselatte-hero'))) window.__heroShifts.push(e.value);
        }).observe({type:'layout-shift',buffered:true});
      });
      const page=await context.newPage();
      page.on('pageerror',e=>errors.push(e.message));
      page.on('console',m=>{if(m.type()==='error') errors.push(m.text());});
      await page.goto(base);
      await page.evaluate(()=>document.fonts.ready);
      await page.locator('.monselatte-hero img').evaluate(img=>img.decode());
      await settle(page);
      const hero=await page.locator('.monselatte-hero img').evaluate(img=>({
        src:img.currentSrc,width:img.clientWidth,height:img.clientHeight,
        priority:img.fetchPriority,loading:img.loading,
        shifts:window.__heroShifts,
        requested:performance.getEntriesByType('resource').filter(r=>/galeria-4|optimized\/sobre.webp/.test(r.name)).map(r=>r.name)
      }));
      assert.match(hero.src,/galeria-4-(480|768|1024|1440|1536|1920)\.webp$/);
      assert.equal(hero.priority,'high');assert.equal(hero.loading,'eager');
      assert(Math.abs(hero.width/hero.height-1.5)<.02);
      assert.equal(hero.requested.length,1,'Preload and picture must reuse one hero request');
      assert(!hero.requested.some(r=>r.includes('/sobre.webp')));
      const filename=path.basename(hero.src); const bytes=fs.statSync(path.join('assets/img/optimized',filename)).size;
      report.push({width,height,dpr,filename,bytes,rendered:hero.width,heroShifts:hero.shifts});
      if(dpr===2){await context.close();continue;}
      await page.screenshot({path:path.join(out,`hero-${width}.png`)});
      const schema=await page.evaluate(()=>Array.from(document.querySelectorAll('script[type="application/ld+json"]'),s=>JSON.parse(s.textContent)));
      assert.deepEqual(schema.map(s=>s['@type']),['LocalBusiness','Organization','WebSite','FAQPage','Service']);
      assert(!/OfferCatalog|InStock|SearchAction|CafeOrCoffeeShop|mocha|Básico/.test(JSON.stringify(schema)));
      const normalize=t=>t.replace(/\s+/g,' ').trim();
      const questions=await page.locator('.faq-q > span').allTextContents();
      const answers=await page.locator('.faq-a-inner').allTextContents();
      schema.find(s=>s['@type']==='FAQPage').mainEntity.forEach((q,i)=>{
        assert.equal(normalize(q.name),normalize(questions[i]));
        assert.equal(normalize(q.acceptedAnswer.text),normalize(answers[i]));
      });
      assert.equal(schema.find(s=>s['@type']==='FAQPage').mainEntity.length,10);
      const cdp=await context.newCDPSession(page);
      const ax=()=>cdp.send('Accessibility.getFullAXTree');
      let tree=await ax();
      assert(!tree.nodes.some(n=>!n.ignored && n.name?.value?.includes('Recomendamos reservar con')));
      const q=page.locator('#faq-question-1'),panel=page.locator('#faq-answer-1');
      await q.press('Enter');await page.waitForTimeout(350);
      assert.equal(await q.getAttribute('aria-expanded'),'true');
      assert.equal(await q.getAttribute('aria-controls'),'faq-answer-1');
      tree=await ax();
      assert(tree.nodes.some(n=>!n.ignored && n.name?.value?.includes('Recomendamos reservar con')));
      // An open answer must fit after rotation/width changes and content reflow.
      for(const next of [{width:320,height:568},{width:1440,height:900},{width,height}]) {
        await page.setViewportSize(next);await page.waitForTimeout(350);
        assert(await panel.evaluate(e=>e.clientHeight>=e.querySelector('.faq-a-inner').scrollHeight-1));
      }
      await panel.locator('.faq-a-inner').evaluate(e=>{e.dataset.original=e.innerHTML;e.append(document.createTextNode(' Texto adicional de prueba.'.repeat(20)));});
      await page.waitForTimeout(350);
      assert(await panel.evaluate(e=>e.clientHeight>=e.querySelector('.faq-a-inner').scrollHeight-1));
      await panel.locator('.faq-a-inner').evaluate(e=>{e.innerHTML=e.dataset.original;});
      await q.press('Space');await page.waitForTimeout(350);
      assert.equal(await panel.getAttribute('aria-hidden'),'true');
      assert(await panel.evaluate(e=>e.inert && e.clientHeight===0));
      tree=await ax();
      assert(!tree.nodes.some(n=>!n.ignored && n.name?.value?.includes('Recomendamos reservar con')));
      await q.press('ArrowDown');assert(await page.locator('#faq-question-2').evaluate(e=>e===document.activeElement));
      await page.emulateMedia({reducedMotion:'reduce'});await q.click();
      assert.equal(await panel.evaluate(e=>getComputedStyle(e).transitionDuration),'0s');
      await q.click();await page.emulateMedia({reducedMotion:'no-preference'});
      // Check the floating target through the page and each wizard step.
      const overlaps=()=>page.evaluate(()=>{
        const wa=document.querySelector('.whatsapp-float'),r=wa.getBoundingClientRect();
        const header=document.querySelector('.site-header'), headerBottom=header.getBoundingClientRect().bottom;
        return Array.from(document.querySelectorAll('a,button,input,select,textarea,label,[tabindex]'))
          .filter(e=>e!==wa && !wa.contains(e) && !e.closest('[inert],[aria-hidden="true"]') && e.getClientRects().length)
          .filter(e=>{const b=e.getBoundingClientRect();return b.width && b.height && b.left<r.right && b.right>r.left && (header.contains(e)?b.top:Math.max(b.top,headerBottom))<r.bottom && b.bottom>r.top;})
          .map(e=>({tag:e.tagName,id:e.id,cls:e.className,text:e.textContent.trim().slice(0,60),rect:e.getBoundingClientRect().toJSON(),wa:r.toJSON()}));
      });
      const docHeight=await page.evaluate(()=>document.documentElement.scrollHeight);
      for(let y=0;y<docHeight;y+=Math.floor(height/3)){
        await page.evaluate(y=>scrollTo({top:y,behavior:'instant'}),y);await settle(page);
        assert.deepEqual(await overlaps(),[],`WhatsApp overlap ${width} at ${y}`);
      }
      for(let step=1;step<=5;step++){
        await page.evaluate(n=>window.__wizardGoTo(n),step);await page.waitForTimeout(400);
        for(const target of ['#reserva','#wizardNav']){
          if(!await page.locator(target).count())continue;
          await page.locator(target).evaluate(e=>scrollTo({top:e.getBoundingClientRect().top+scrollY-innerHeight/2,behavior:'instant'}));
          await settle(page);assert.deepEqual(await overlaps(),[],`Wizard ${step}, ${width}`);
        }
      }
      const wa=page.locator('.whatsapp-float');
      assert.equal(await wa.getAttribute('href'),'https://wa.me/17876108953');
      const waBox=await wa.boundingBox();assert(waBox.width>=44&&waBox.height>=44);
      await wa.evaluate(e=>{e.addEventListener('click',event=>event.preventDefault(),{once:true});e.click();});
      assert(await page.evaluate(()=>dataLayer.some(e=>e[0]==='event'&&e[1]==='contact'&&e[2].method==='whatsapp'&&e[2].location==='floating')));
      assert(!await page.evaluate(()=>document.documentElement.scrollWidth>innerWidth));
      await page.locator('#faq-question-1').click();await page.waitForTimeout(350);
      await page.locator('#faq').evaluate(e=>scrollTo({top:e.getBoundingClientRect().top+scrollY-100,behavior:'instant'}));
      await settle(page);await page.screenshot({path:path.join(out,`faq-${width}.png`)});
      console.log(`${width}x${height}: schema, FAQ resize/AX/keyboard/content/reduced motion, WhatsApp scroll/wizard/tracking, hero PASS`);
      await context.close();
    }
    assert.deepEqual(errors,[]);
    fs.writeFileSync(path.join(out,'report.json'),JSON.stringify(report,null,2));
    console.log('All 8 viewports, DPR 1/2, no real submissions/analytics, zero console errors PASS');
  } finally {await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
