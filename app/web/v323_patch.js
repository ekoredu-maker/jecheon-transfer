/* v3.2.3 - standard input workflow */
(function(){
  'use strict';
  var VERSION='v3.2.3-InputWorkflow';
  try{ Store.version=VERSION; }catch(e){}

  var INPUTS={
    schools:{id:'file-schools',label:'학교기본정보'},
    quota:{id:'file-quota',label:'정원·현원'},
    intra:{id:'file-intra',label:'관내전보·전입명부'},
    out:{id:'file-out',label:'관외전출·퇴직명부'}
  };
  var IN_TYPES=['복직','복귀','관내전보','타시군전입','타시도전입','비정기전입','신규임용'];
  var OUT_TYPES=['타시군전출','타시도전출','정년퇴직','명예퇴직','면직','휴직','파견','승진·전직','국립전출','기타'];

  function cfg(){
    Store.state.config=Store.state.config||{};
    Store.state.config.inputStatus323=Store.state.config.inputStatus323||{};
    return Store.state.config;
  }
  function text(v){ return String(v==null?'':v).trim(); }
  function norm323(v){ return text(v).replace(/\s+/g,''); }
  function canon(v){ return typeof canonicalSchool315==='function'?canonicalSchool315(v):norm323(v); }
  function hasHeader(target,name){
    var hs=(Store.state.config&&Store.state.config.lastImport319&&Store.state.config.lastImport319[target]&&Store.state.config.lastImport319[target].headers)||[];
    return hs.some(function(h){return norm323(h)===norm323(name);});
  }
  function inType(v){
    var s=norm323(v);
    if(/복직/.test(s))return '복직';
    if(/복귀/.test(s))return '복귀';
    if(/관내/.test(s))return '관내전보';
    if(/타시도|타시·도|시도간|시·도간|교류입|시도교류/.test(s))return '타시도전입';
    if(/타시군|타시·군|시군간|시·군간|청간/.test(s))return '타시군전입';
    if(/비정기/.test(s))return '비정기전입';
    if(/신규/.test(s))return '신규임용';
    if(/관내|희망전보|학교만기|급지만기|지역만기|구역만기|만기/.test(s))return '관내전보';
    return text(v);
  }
  function outType(v){
    var s=norm323(v);
    if(/타시군/.test(s))return '타시군전출';
    if(/타시도|시도간|시도교류|교류전출/.test(s))return '타시도전출';
    if(/명예퇴직|명퇴/.test(s))return '명예퇴직';
    if(/정년퇴직|정퇴/.test(s))return '정년퇴직';
    if(/면직/.test(s))return '면직';
    if(/휴직/.test(s))return '휴직';
    if(/파견/.test(s))return '파견';
    if(/승진|전직/.test(s))return '승진·전직';
    if(/국립/.test(s))return '국립전출';
    if(/관외전출|전출/.test(s))return '기타';
    if(/기타/.test(s))return '기타';
    return text(v);
  }

  var baseMapIntra=mapIntra;
  mapIntra=function(rows){
    var src=(rows||[]).map(function(r){
      var o=Object.assign({},r);
      if(text(r['배치유형'])&&!text(r['발령사유']))o['발령사유']=r['배치유형'];
      if(text(r['복귀·복직대상교'])&&!text(r['복직대상학교']))o['복직대상학교']=r['복귀·복직대상교'];
      return o;
    });
    var out=baseMapIntra(src);
    out.forEach(function(t){
      t.배치유형=inType(t.발령사유||t.배치유형);
      if(t.배치유형)t.발령사유=t.배치유형;
    });
    return out;
  };

  var baseMapOut=mapOut;
  mapOut=function(rows){
    var src=(rows||[]).map(function(r){
      var o=Object.assign({},r);
      if(text(r['전출유형'])&&!text(r['구분']))o['구분']=r['전출유형'];
      if(text(r['세부사유'])&&!text(r['사유']))o['사유']=r['세부사유'];
      return o;
    });
    var out=baseMapOut(src);
    out.forEach(function(o){o.전출유형=outType([o.전출유형,o.구분,o.처리,o.사유].filter(Boolean).join(' '));});
    return out;
  };

  function audit(target){
    var s=Store.state||{}, errors=[], warnings=[], meta=cfg().inputStatus323[target]||{};
    if(target==='schools'){
      var a=s.schools||[], seen={};
      if(!a.length)errors.push('학교기본정보가 없습니다.');
      a.forEach(function(x,i){
        var n=norm323(x.학교명);
        if(seen[n])errors.push('학교명 중복: '+x.학교명);
        seen[n]=1;
        if(!text(x.학교급))errors.push((x.학교명||(i+1)+'행')+' 학교급 누락');
        if(!text(x.급지)&&!/센터|지원청/.test(text(x.학교급)+text(x.학교명)))errors.push((x.학교명||(i+1)+'행')+' 급지 누락');
      });
    }
    if(target==='quota'){
      var q=s.quota||[], master={};
      (s.schools||[]).forEach(function(x){master[norm323(canon(x.학교명))]=1;});
      if(!q.length)errors.push('정원·현원 자료가 없습니다.');
      var qseen={};
      q.forEach(function(x){
        var k=norm323(canon(x.학교명))+'|'+norm323(x.교원구분);
        if(qseen[k])errors.push('정원 키 중복: '+x.학교명+' / '+x.교원구분);
        qseen[k]=1;
        if(Object.keys(master).length&&!master[norm323(canon(x.학교명))])errors.push('학교마스터 미등록: '+x.학교명);
        if(x.현원228===null||x.현원228===undefined||x.현원228==='')errors.push(x.학교명+'/'+x.교원구분+': 2.28 실제현원 누락');
      });
      if(!hasHeader('quota','2.28현원'))warnings.push('표준서식의 [2.28현원] 열 사용 권장');
    }
    if(target==='intra'){
      var ia=s.intra||[];
      if(!ia.length)errors.push('관내전보·전입명부가 없습니다.');
      ia.forEach(function(t,i){
        var row=t.sourceRowNo||i+2, typ=inType(t.배치유형||t.발령사유);
        if(!text(t.성명))errors.push(row+'행 성명 누락');
        if(!text(t.생년월일))errors.push((t.성명||row+'행')+' 생년월일 누락');
        if(!text(t.교원구분))errors.push((t.성명||row+'행')+' 교원구분 누락');
        if(IN_TYPES.indexOf(typ)<0)errors.push((t.성명||row+'행')+' 배치유형 확인 필요: '+(typ||'미입력'));
        if(['복직','복귀','관내전보'].indexOf(typ)>=0&&!text(t.현임교))errors.push((t.성명||row+'행')+' 현임교 누락');
        if(['관내전보','타시군전입','타시도전입','비정기전입'].indexOf(typ)>=0&&!text(t.지망1))warnings.push(t.성명+': 1지망 미입력');
        if(Number(t.전보순위)!==i+1)warnings.push(t.성명+': 명부 위→아래 순위와 내부순위 불일치');
      });
      if(!hasHeader('intra','배치유형'))warnings.push('호환서식 읽기: 다음부터 [배치유형] 열 사용 권장');
    }
    if(target==='out'){
      if(cfg().noOut323===true)return {ready:true,errors:[],warnings:[],loaded:true,noOut:true};
      var oa=s.out||[];
      if(!meta.loaded&&!oa.length)errors.push('관외전출명부를 올리거나 [관외전출 없음]을 선택하세요.');
      oa.forEach(function(o,i){
        var row=o.sourceRowNo||i+2, typ=outType([o.전출유형,o.구분,o.처리,o.사유].filter(Boolean).join(' '));
        if(!text(o.성명))errors.push(row+'행 성명 누락');
        if(!text(o.생년월일))warnings.push((o.성명||row+'행')+' 생년월일 누락');
        if(!text(o.교원구분))warnings.push((o.성명||row+'행')+' 교원구분 누락');
        if(!text(o.현임교))errors.push((o.성명||row+'행')+' 현임교 누락');
        if(OUT_TYPES.indexOf(typ)<0)errors.push((o.성명||row+'행')+' 전출유형 확인 필요: '+(typ||'미입력'));
      });
      if(meta.loaded&&!hasHeader('out','전출유형'))warnings.push('호환서식 읽기: 다음부터 [전출유형] 열 사용 권장');
    }
    return {ready:errors.length===0,errors:Array.from(new Set(errors)),warnings:Array.from(new Set(warnings)),loaded:!!meta.loaded||((target==='out')&&(Store.state.out||[]).length>0)};
  }
  window.inputAudit323=audit;

  function status(){
    var c=cfg(), setup=!!(text(c.appointDate)&&text(c.personnelYear));
    var A={schools:audit('schools'),quota:audit('quota'),intra:audit('intra'),out:audit('out')};
    var inputs=setup&&A.schools.ready&&A.quota.ready&&A.intra.ready&&A.out.ready;
    var vacancy=inputs&&c.vacConfirmed===true;
    var active=(Store.state.intra||[]).filter(function(t){return t.상태!=='관외확정';});
    var unresolved=active.filter(function(t){return !((t.상태==='배치'&&t.배치교)||(t.상태==='잔류'&&t.현임교));}).length;
    var placement=vacancy&&unresolved===0, final=false;
    try{final=placement&&typeof validateAllStable==='function'&&(validateAllStable(true).errors||[]).length===0;}catch(e){}
    return {setup:setup,A:A,inputs:inputs,vacancy:vacancy,placement:placement,final:final,unresolved:unresolved};
  }
  window.workflowStatus323=status;

  function renderCards(){
    Object.keys(INPUTS).forEach(function(target){
      var x=INPUTS[target], el=document.getElementById(x.id);
      if(!el)return;
      var card=el.closest('.card-base'); if(!card)return;
      var box=card.querySelector('.input-status-323');
      if(!box){box=document.createElement('div');box.className='input-status-323';card.appendChild(box);}
      var a=audit(target), meta=cfg().inputStatus323[target]||{};
      box.className='input-status-323 '+(a.ready?(a.warnings.length?'warn':'ok'):(a.loaded?'err':''));
      var state=a.ready?'✓ 준비완료':(a.loaded?'✕ 확인필요':'○ 미등록');
      var file=meta.noOut?'관외전출 없음':(meta.file||'');
      var detail=a.errors[0]||a.warnings[0]||String(meta.rows||0)+'건 반영';
      box.innerHTML='<b>'+state+'</b>'+(file?' · '+escapeHtml(file):'')+'<div class="mt-0.5">'+escapeHtml(detail)+'</div>';
    });
  }

  function refreshWorkflow323(){
    var st=status(), holder=document.getElementById('workflow-steps-322');
    if(holder){
      var steps=[['발령설정',st.setup],['기준자료',st.inputs],['결원확정',st.vacancy],['배치완료',st.placement],['최종검증',st.final],['전의안',st.final],['발령통지',st.final]];
      var p=steps.findIndex(function(x){return !x[1];}); if(p<0)p=steps.length;
      holder.innerHTML=steps.map(function(x,i){
        var cls=x[1]?'done':(i===p?'current':'locked');
        return '<div class="wf-step '+cls+'"><span class="wf-num">'+(x[1]?'✓':i+1)+'</span><span class="truncate">'+x[0]+'</span></div>';
      }).join('');
    }
    renderCards();
    var rs=document.getElementById('rule-status-322');
    if(rs){
      var bad=Object.keys(st.A).filter(function(k){return !st.A[k].ready;}).map(function(k){return INPUTS[k].label;});
      rs.innerHTML=bad.length?'다음 확인 필요: <b class="text-rose-600">'+bad.join(' · ')+'</b>':'입력 4종 검증 완료 · 결원 확정 단계로 진행할 수 있습니다.';
    }
  }
  window.refreshWorkflow323=refreshWorkflow323;
  window.refreshWorkflow322=refreshWorkflow323;

  function template(target){
    if(typeof XLSX==='undefined')return toast('엑셀 라이브러리를 먼저 불러와 주세요.','error');
    var wb=XLSX.utils.book_new();
    function add(name,rows){XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),name);}
    if(target==='schools'){
      add('학교기본정보',[['교번','학교명','학교급','급지','학교코드','전의안표시명','정렬순서','사용여부']]);
      add('전보금칙설정',[
        ['규칙코드','규칙명','사용여부','처리수준','적용조건','설명'],
        ['SCHOOL_MATURITY_CURRENT','학교만기 현임교 배치','Y','금지','학교만기=Y AND 배치교=현임교','현임교 재배치 금지'],
        ['REGION_MATURITY_SAME_ZONE','지역만기 동일급지 배치','Y','금지','지역만기=Y AND 현임교급지=배치교급지','다른 급지로 이동'],
        ['REGION_MATURITY_ZONE_MISSING','지역만기 급지 미확인','Y','검토','지역만기=Y AND 급지 누락','급지 확인 후 배치'],
        ['FAMILY_SAME_SCHOOL','가족교원 상피','Y','검토','배치교=가족교원근무교 또는 개인금칙학교','기관 기준에 따라 금지로 변경 가능']
      ]);
      add('개인금칙',[['성명','생년월일','교원구분','금칙학교','금칙유형','처리수준','사용여부','비고']]);
      add('학교별칭',[['입력학교명','정식학교명','사용여부','비고']]);
    }else if(target==='quota'){
      add('정원·현원',[['학교명','교원구분','2.28정원','2.28현원','3.1정원','2.28학급','2.28전담','2.28기타전담','2.28수석','3.1학급','3.1전담','3.1기타전담','3.1과원','3.1수석','기간제','별도정원']]);
    }else if(target==='intra'){
      add('관내전보·전입명부',[['성명','생년월일','교원구분','배치유형','현임교','복직대상학교','1지망','2지망','학교만기','지역만기','가족교원근무교','별도정원','비고']]);
      add('작성안내',[['배치유형 허용값']].concat(IN_TYPES.map(function(x){return [x];})).concat([[],['순위 원칙'],['데이터 행의 위→아래 순서가 관내전보 순위입니다. 별도 전보순위 열은 필요하지 않습니다.']]));
    }else if(target==='out'){
      add('관외전출·퇴직명부',[['성명','생년월일','교원구분','현임교','전출유형','세부사유','별도정원','비고']]);
      add('작성안내',[['전출유형 허용값']].concat(OUT_TYPES.map(function(x){return [x];})));
    }
    XLSX.writeFile(wb,INPUTS[target].label+'_표준서식_v3.2.3.xlsx');
  }

  function decorate(){
    var style=document.createElement('style');
    style.textContent='.input-status-323{margin-top:.55rem;padding:.5rem .6rem;border-radius:.65rem;font-size:10px;line-height:1.45;border:1px solid #e2e8f0;background:#f8fafc;color:#64748b}.input-status-323.ok{border-color:#bbf7d0;background:#f0fdf4;color:#166534}.input-status-323.warn{border-color:#fde68a;background:#fffbeb;color:#92400e}.input-status-323.err{border-color:#fecaca;background:#fff1f2;color:#9f1239}.template-btn-323{margin-top:.45rem;width:100%;padding:.42rem .55rem;border:1px solid #e2e8f0;border-radius:.55rem;background:white;color:#475569;font-size:10px;font-weight:800}.template-btn-323:hover{background:#f8fafc;border-color:#c7d2fe;color:#4338ca}.wf-step.locked{opacity:.52;background:#f8fafc}';
    document.head.appendChild(style);
    Object.keys(INPUTS).forEach(function(target){
      var x=INPUTS[target], el=document.getElementById(x.id); if(!el)return;
      var card=el.closest('.card-base'); if(!card||card.querySelector('[data-template323="'+target+'"]'))return;
      var b=document.createElement('button');b.type='button';b.setAttribute('data-template323',target);b.className='template-btn-323';
      b.innerHTML='<i class="fa-solid fa-file-arrow-down mr-1"></i>표준서식 다운로드';
      b.onclick=async function(){try{await ensureXLSXReady();template(target);}catch(e){toast('표준서식 생성 실패','error');}};
      var drop=el.closest('label'); if(drop)drop.insertAdjacentElement('afterend',b);
      if(target==='out'){
        var n=document.createElement('button');n.type='button';n.id='btn-no-out-323';n.className='template-btn-323';
        n.innerHTML='<i class="fa-solid fa-circle-check mr-1"></i>관외전출 없음';
        n.onclick=function(){
          var c=cfg();c.noOut323=true;c.inputStatus323.out={loaded:true,noOut:true,file:'관외전출 없음',rows:0,at:new Date().toISOString()};
          Store.state.out=[];try{Store.save();}catch(e){}refreshWorkflow323();toast('관외전출·퇴직 없음으로 설정했습니다.');
        };
        b.insertAdjacentElement('afterend',n);
      }
    });
  }

  async function import323(file,target){
    var c=cfg();if(target==='out')c.noOut323=false;
    var parsed=await window.importFile319(file,target);
    var meta=(Store.state.config&&Store.state.config.lastImport319&&Store.state.config.lastImport319[target])||{};
    cfg().inputStatus323[target]={loaded:true,file:file.name,rows:Array.isArray(parsed)?parsed.length:0,at:new Date().toISOString(),headers:meta.headers||[]};
    try{Store.save();}catch(e){}
    refreshWorkflow323();
    var a=audit(target);
    if(!a.ready)toast(INPUTS[target].label+': '+(a.errors[0]||'확인 필요'),'error');
    else if(a.warnings.length)toast(INPUTS[target].label+': '+a.warnings[0],'warn');
    return parsed;
  }
  window.importFile323=import323;

  function bind(){
    Object.keys(INPUTS).forEach(function(target){
      var el=document.getElementById(INPUTS[target].id);if(!el)return;
      el.onchange=async function(e){
        var file=e.target.files&&e.target.files[0];if(!file)return;
        try{await import323(file,target);}
        catch(err){console.error(err);log(file.name+': '+(err.message||err),'error');toast('파일 읽기 실패: '+(err.message||err),'error');}
        finally{e.target.value='';refreshWorkflow323();}
      };
    });
  }

  function guard(){
    var vac=document.getElementById('btn-confirm-vac');
    if(vac&&!vac.dataset.guard323){
      var oldVac=vac.onclick;vac.dataset.guard323='1';
      vac.onclick=function(e){var st=status();if(!st.inputs){refreshWorkflow323();return toast('발령설정과 입력 4종 검증을 먼저 완료하세요.','error');}if(oldVac)return oldVac.call(this,e);};
    }
    var auto=document.getElementById('btn-autoplace');
    if(auto&&!auto.dataset.guard323){
      var oldAuto=auto.onclick;auto.dataset.guard323='1';
      auto.onclick=function(e){if(cfg().vacConfirmed!==true)return toast('결원 확정을 먼저 완료하세요.','error');if(oldAuto)return oldAuto.call(this,e);};
    }
  }

  var baseRefresh=refreshAll;
  refreshAll=function(){var r=baseRefresh.apply(this,arguments);try{refreshWorkflow323();}catch(e){}return r;};

  window.addEventListener('DOMContentLoaded',function(){
    try{
      decorate();bind();guard();refreshWorkflow323();
      var title=document.querySelector('header span.text-indigo-600');if(title)title.textContent='v3.2.3 (Input Workflow)';
      log('v3.2.3 적용: 입력 4종 표준서식 · 업로드 준비상태 · 단계 잠금 · 명부순서 순위 원칙 고정','info');
    }catch(e){console.error('v3.2.3 init',e);}
  });
})();