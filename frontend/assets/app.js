// ── CONFIG ──
const ENGSAFE_CFG = window.ENGSAFE_CONFIG || {};
// URL e chave publishable vêm do config.js, preenchido no build a partir das variáveis de ambiente.
const SB_URL = String(ENGSAFE_CFG.supabaseUrl || '').replace(/\/+$/, '');
const SB_KEY = String(ENGSAFE_CFG.supabaseKey || '');
if (!/^https:\/\/[a-z0-9]+\.supabase\.co$/.test(SB_URL) || !SB_KEY.startsWith('sb_publishable_')) {
  const aviso = 'Configuração do Supabase ausente. Rode "npm run build" com SUPABASE_URL e SUPABASE_PUBLISHABLE_KEY definidas e abra a pasta dist.';
  console.error(aviso);
  document.addEventListener('DOMContentLoaded', () => { document.body.insertAdjacentHTML('afterbegin', `<div role="alert" style="background:#b00020;color:#fff;padding:12px;font:14px sans-serif">${aviso}</div>`); });
}
// Envio manual pelo Outlook Web. A conta remetente é definida pela sessão aberta no Outlook.
const OUTLOOK_COMPOSE_URL = 'https://outlook.office.com/mail/deeplink/compose';
const OUTLOOK_SENDER_REMINDER = 'engenharia.ser@ambientalserra.com.br';
const AUTH_SESSION_KEY = 'engsafe.auth.session.v2'; // v2: banco novo, descarta sessões do projeto antigo

// ── STATE ──
let DATA = {ocorrencias:[], atividades:[], contratadas:[], autoridades:[], historicoEmails:[], usuarios:[]};
let charts = {};
let AUTH_SESSION=null;
let AUTH_USER=null;
let AUTH_PROFILE=null;
let PENDING_EMAIL_LOG=null;
const ACTION_VALIDATION_STATE=new Map();

// ── UTILS ──
const el = id => document.getElementById(id);
const fmt = d => d ? new Date(d+'T12:00').toLocaleDateString('pt-BR') : '—';
const today = () => new Date().toISOString().slice(0,10);
const canManage = () => ['colaborador','admin'].includes(AUTH_PROFILE?.perfil)&&AUTH_PROFILE?.status_acesso==='aprovado'&&AUTH_PROFILE?.ativo!==false;
const isAdmin = () => AUTH_PROFILE?.perfil==='admin'&&AUTH_PROFILE?.status_acesso==='aprovado'&&AUTH_PROFILE?.ativo!==false;
const canManageAccess = () => isAdmin()&&AUTH_PROFILE?.gerencia_acessos===true;
function requireManage(){if(canManage())return true;toast('Seu perfil permite somente leitura e download de PDF.','warn');return false;}
function requireAccessManager(){if(canManageAccess())return true;toast('A gestão de acessos é exclusiva da administradora responsável.','warn');return false;}

function authHeaders(contentType='application/json',prefer){
  const h={apikey:SB_KEY,Authorization:`Bearer ${AUTH_SESSION?.access_token||SB_KEY}`};
  if(contentType)h['Content-Type']=contentType;if(prefer)h.Prefer=prefer;return h;
}
async function storageSignedUrl(bucket,path){
  const r=await fetch(`${SB_URL}/storage/v1/object/sign/${bucket}/${path}`,{method:'POST',headers:authHeaders(),body:JSON.stringify({expiresIn:900})});
  const data=await r.json().catch(()=>({}));if(!r.ok)throw new Error(data.message||'Falha ao autorizar anexo');
  const signed=data.signedURL||data.signedUrl;return signed?.startsWith('http')?signed:`${SB_URL}/storage/v1${signed}`;
}
function saveSession(session){AUTH_SESSION=session||null;if(session)localStorage.setItem(AUTH_SESSION_KEY,JSON.stringify(session));else localStorage.removeItem(AUTH_SESSION_KEY);}
function showAuthMessage(msg,type='err'){const box=el('auth-message');box.textContent=msg;box.className=`auth-msg show ${type}`;}
function showAuthTab(tab){el('login-form').style.display=tab==='login'?'block':'none';el('signup-form').style.display=tab==='signup'?'block':'none';el('tab-login').classList.toggle('active',tab==='login');el('tab-signup').classList.toggle('active',tab==='signup');el('auth-message').className='auth-msg';}
function showPending(profile){el('auth-forms').style.display='none';el('auth-pending').style.display='block';const rejected=profile?.status_acesso==='rejeitado',inactive=profile?.ativo===false;el('pending-title').textContent=rejected?'Acesso não aprovado':inactive?'Acesso bloqueado':'Aguardando aprovação';el('pending-text').textContent=rejected?'Sua solicitação foi recusada. Procure a administradora do EngSafe.':inactive?'Seu acesso está inativo. Procure a administradora do EngSafe.':'Sua solicitação foi recebida e aguarda liberação da administradora.';}

async function authRequest(path,options={}){
  const res=await fetch(`${SB_URL}/auth/v1${path}`,{...options,headers:{apikey:SB_KEY,'Content-Type':'application/json',...(options.headers||{})}});const data=await res.json().catch(()=>({}));if(!res.ok)throw new Error(data.msg||data.error_description||data.message||data.error||`Falha de autenticação (${res.status})`);return data;
}
async function loginEngsafe(event){event.preventDefault();const btn=el('login-btn');btn.disabled=true;btn.textContent='Entrando...';try{const data=await authRequest('/token?grant_type=password',{method:'POST',body:JSON.stringify({email:el('login-email').value.trim().toLowerCase(),password:el('login-password').value})});saveSession(data);await iniciarSessao();}catch(e){showAuthMessage(e.message);}finally{btn.disabled=false;btn.textContent='Entrar';}}
async function solicitarAcesso(event){event.preventDefault();const btn=el('signup-btn');btn.disabled=true;btn.textContent='Enviando...';try{const data=await authRequest('/signup',{method:'POST',body:JSON.stringify({email:el('signup-email').value.trim().toLowerCase(),password:el('signup-password').value,data:{nome:el('signup-name').value.trim()}})});if(data.access_token){saveSession(data);await iniciarSessao();}else{showAuthMessage('Solicitação enviada. Confirme o e-mail corporativo e depois aguarde a aprovação da administradora.','ok');el('signup-form').reset();}}catch(e){showAuthMessage(e.message);}finally{btn.disabled=false;btn.textContent='Enviar solicitação';}}
async function recuperarSenha(){const email=(el('login-email').value||'').trim().toLowerCase();if(!email){showAuthMessage('Informe seu e-mail corporativo no campo acima.');return;}try{await authRequest('/recover',{method:'POST',body:JSON.stringify({email,redirect_to:location.href.split('#')[0]})});showAuthMessage('Se o e-mail estiver cadastrado, você receberá as instruções de recuperação.','ok');}catch(e){showAuthMessage(e.message);}}
async function logoutEngsafe(){try{if(AUTH_SESSION?.access_token)await authRequest('/logout',{method:'POST',headers:{Authorization:`Bearer ${AUTH_SESSION.access_token}`}});}catch(e){}saveSession(null);AUTH_USER=null;AUTH_PROFILE=null;location.reload();}
async function refreshSession(){if(!AUTH_SESSION?.refresh_token)return false;try{const data=await authRequest('/token?grant_type=refresh_token',{method:'POST',body:JSON.stringify({refresh_token:AUTH_SESSION.refresh_token})});saveSession(data);return true;}catch(e){return false;}}
function sessionFromHash(){const p=new URLSearchParams(location.hash.slice(1));if(!p.get('access_token'))return null;const s={access_token:p.get('access_token'),refresh_token:p.get('refresh_token'),expires_in:Number(p.get('expires_in')||3600),token_type:'bearer'};history.replaceState(null,'',location.pathname+location.search);return s;}
async function getAuthenticatedUser(){let res=await fetch(`${SB_URL}/auth/v1/user`,{headers:authHeaders(null)});if(res.status===401&&await refreshSession())res=await fetch(`${SB_URL}/auth/v1/user`,{headers:authHeaders(null)});if(!res.ok)throw new Error('Sessão expirada. Entre novamente.');return res.json();}
async function iniciarSessao(){try{AUTH_USER=await getAuthenticatedUser();const rows=await sbGet('usuarios',`id=eq.${encodeURIComponent(AUTH_USER.id)}&select=*`);AUTH_PROFILE=rows[0];if(!AUTH_PROFILE){showPending({status_acesso:'pendente'});return;}if(AUTH_PROFILE.status_acesso!=='aprovado'||AUTH_PROFILE.ativo===false){showPending(AUTH_PROFILE);return;}aplicarPerfil();document.body.classList.remove('auth-locked');await loadAll();}catch(e){saveSession(null);showAuthMessage(e.message);}}
function aplicarPerfil(){document.body.classList.remove('access-leitor','access-colaborador','access-admin','access-manager','no-access-manager');document.body.classList.add(`access-${AUTH_PROFILE.perfil}`,canManageAccess()?'access-manager':'no-access-manager');const nome=AUTH_PROFILE.nome||AUTH_USER?.email||'Usuário';el('user-name').textContent=nome;el('sidebar-user-name').textContent=nome;el('user-role').textContent={leitor:'Leitor',colaborador:'Colaborador',admin:'Administrador'}[AUTH_PROFILE.perfil]||AUTH_PROFILE.perfil;el('user-avatar').textContent=nome.trim().charAt(0).toUpperCase();}
async function bootAuth(){const hashSession=sessionFromHash();if(hashSession)saveSession(hashSession);else{try{AUTH_SESSION=JSON.parse(localStorage.getItem(AUTH_SESSION_KEY)||'null');}catch(e){AUTH_SESSION=null;}}if(AUTH_SESSION)await iniciarSessao();}

function norm(s){ return (s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));}
function normEmp(s){
  const n=norm(s);
  if(n.includes('macro')) return 'Macro Ambiental';
  if(n.includes('a3')) return 'A3 Engenharia';
  if(n.includes('ae eng')) return 'AE Engenharia';
  if(n.includes('ambiental vila')) return 'Ambiental Vila Velha';
  if(n.includes('ambiental cariacica')) return 'Ambiental Cariacica';
  if(n.includes('caj')) return 'CAJ';
  return s.trim();
}
function normTec(s){
  const n=norm(s);
  if(n.includes('robervan cord')) return 'Robervan Cordeiro';
  if(n.includes('robervan')) return 'Robervan Cordeiro';
  if(n.includes('jairo')) return 'Jairo Filho';
  return s.trim();
}

function critOrder(c){return c==='alta'?0:c==='media'?1:2;}

function isoLocal(date){
  const y=date.getFullYear(),m=String(date.getMonth()+1).padStart(2,'0'),d=String(date.getDate()).padStart(2,'0');
  return `${y}-${m}-${d}`;
}
function periodoDatas(tipo){
  const agora=new Date(),fim=new Date(agora.getFullYear(),agora.getMonth(),agora.getDate());
  let ini=new Date(fim);
  if(tipo==='7dias') ini.setDate(ini.getDate()-6);
  else if(tipo==='30dias') ini.setDate(ini.getDate()-29);
  else if(tipo==='mes') ini=new Date(agora.getFullYear(),agora.getMonth(),1);
  else if(tipo==='mes_anterior'){
    ini=new Date(agora.getFullYear(),agora.getMonth()-1,1);
    return {ini:isoLocal(ini),fim:isoLocal(new Date(agora.getFullYear(),agora.getMonth(),0))};
  }
  return {ini:isoLocal(ini),fim:isoLocal(fim)};
}
function onPeriodoOcorrenciasChange(){
  const tipo=el('fPeriodo')?.value||'';
  const custom=el('occ-period-custom');
  custom?.classList.toggle('show',tipo==='personalizado');
  if(tipo&&tipo!=='personalizado'){
    const datas=periodoDatas(tipo);
    el('fDtI').value=datas.ini;el('fDtF').value=datas.fim;
  }else if(!tipo){el('fDtI').value='';el('fDtF').value='';}
  renderOcorrencias();
}

function badge(type,val){
  const map={alta:'badge-alta',media:'badge-media',baixa:'badge-baixa',aberta:'badge-aberta',concluida:'badge-concluida',fechada:'badge-fechada'};
  const label={alta:'Alta',media:'Média',baixa:'Baixa',aberta:'Aberta',concluida:'Concluída',fechada:'Fechada'};
  return `<span class="badge ${map[val]||''}">${label[val]||val}</span>`;
}

function isVencida(r){
  return r.prazo && r.status!=='concluida' && r.status!=='fechada' && new Date(r.prazo)<new Date();
}
function getContratadaOcorrencia(r){
  return DATA.contratadas.find(c=>c.id===r.contratada_id)||DATA.contratadas.find(c=>norm(c.nome_fantasia)===norm(r.empresa)||norm(c.razao_social)===norm(r.empresa));
}
function podeEnviarContratada(r){return r.status==='aberta'&&(isVencida(r)||r.criticidade==='alta');}

// ── TOAST ──
function toast(msg, type='ok'){
  const icons={ok:'✓',err:'✕',warn:'⚠'};
  const t=document.createElement('div');
  t.className=`toast ${type}`;
  t.innerHTML=`<span>${icons[type]}</span><span>${msg}</span>`;
  el('toasts').appendChild(t);
  setTimeout(()=>t.remove(), 3500);
}

// ── SUPABASE ──
async function sbGet(table, params=''){
  const r=await fetch(`${SB_URL}/rest/v1/${table}?${params}&limit=500`,{headers:authHeaders(null)});
  if(!r.ok) throw new Error(`${r.status}: ${await r.text()}`);
  return r.json();
}
async function sbInsert(table, row){
  if(!requireManage())throw new Error('Perfil sem permissão de escrita.');
  const r=await fetch(`${SB_URL}/rest/v1/${table}`,{method:'POST',headers:authHeaders('application/json','return=representation'),body:JSON.stringify(row)});
  if(!r.ok) throw new Error(`${r.status}: ${await r.text()}`);
  return r.json();
}
async function sbUpdate(table, id, row){
  if(!requireManage())throw new Error('Perfil sem permissão de escrita.');
  const r=await fetch(`${SB_URL}/rest/v1/${table}?id=eq.${id}`,{method:'PATCH',headers:authHeaders('application/json','return=minimal'),body:JSON.stringify(row)});
  if(!r.ok) throw new Error(`${r.status}: ${await r.text()}`);
}
async function sbDelete(table, id){
  if(!requireManage())throw new Error('Perfil sem permissão de exclusão.');
  const r=await fetch(`${SB_URL}/rest/v1/${table}?id=eq.${id}`,{method:'DELETE',headers:authHeaders('application/json','return=minimal')});
  if(!r.ok) throw new Error(`${r.status}: ${await r.text()}`);
}

// ── LOAD ──
async function loadAll(){
  try{
    const [oc, at, co, au, he] = await Promise.all([
      sbGet('ocorrencias','select=*&order=criado_em.desc'),
      sbGet('atividades_sst','select=*&order=criado_em.desc').catch(()=>[]),
      sbGet('contratadas','select=*&order=nome_fantasia.asc').catch(()=>[]),
      sbGet('autoridades_contratada','select=*&order=nome.asc').catch(()=>[]),
      sbGet('historico_emails','select=*&order=data_envio.desc').catch(()=>[])
    ]);
    DATA.ocorrencias = oc.map(r=>({...r, empresa:normEmp(r.empresa), tecnico:normTec(r.tecnico)}));
    DATA.atividades  = at;
    DATA.contratadas = co;
    DATA.autoridades = au;
    DATA.historicoEmails = he;
    el('occ-error').style.display='none';
  }catch(e){
    el('occ-error').style.display='block';
    el('occ-error').textContent='Falha ao carregar: '+e.message;
  }
  populateFilters();
  renderCurrentPage();
}

function populateFilters(){
  const empresas=[...new Set([...DATA.ocorrencias.map(r=>r.empresa),...DATA.contratadas.filter(r=>r.ativo).map(r=>r.nome_fantasia)].filter(Boolean))].sort();
  const tecnicos=[...new Set(DATA.ocorrencias.map(r=>r.tecnico).filter(Boolean))].sort();

  // occ empresa filter
  const fe=el('fEmpresa');
  const cur=fe.value;
  fe.innerHTML='<option value="">Todas empresas</option>'+empresas.map(e=>`<option ${cur===e?'selected':''}>${e}</option>`).join('');

  // relatório selects
  ['rEmp'].forEach(id=>{
    const s=el(id); if(!s) return;
    const v=s.value;
    s.innerHTML='<option value="">Todas</option>'+empresas.map(e=>`<option ${v===e?'selected':''}>${e}</option>`).join('');
  });
  const rt=el('rTec'); if(rt){
    const v=rt.value;
    rt.innerHTML='<option value="">Todos</option>'+tecnicos.map(t=>`<option ${v===t?'selected':''}>${t}</option>`).join('');
  }

  // datalists
  const empDL=el('empList');
  if(empDL) empDL.innerHTML=empresas.map(e=>`<option value="${e}">`).join('');
  const tecDL=el('tecList');
  if(tecDL) tecDL.innerHTML=tecnicos.map(t=>`<option value="${t}">`).join('');
  const fiscais=[...new Set(DATA.ocorrencias.map(r=>r.fiscal).filter(Boolean))].sort();
  const fDL=el('fiscList');if(fDL) fDL.innerHTML=fiscais.map(f=>`<option value="${f}">`).join('');
  const encs=[...new Set(DATA.ocorrencias.map(r=>r.encarregado).filter(Boolean))].sort();
  const eDL=el('encarList');if(eDL) eDL.innerHTML=encs.map(e=>`<option value="${e}">`).join('');
}

// ── NAVIGATION ──
let currentPage = 'dashboard';
function goPage(p){
  if(['nova','novaAtiv'].includes(p)&&!requireManage()) return;
  if(p==='acessos'&&!requireAccessManager()) return;
  document.querySelectorAll('.page').forEach(d=>d.classList.remove('active'));
  document.querySelectorAll('.sb-item').forEach(a=>a.classList.remove('active'));
  el('page-'+p)?.classList.add('active');
  document.querySelector(`[data-page="${p}"]`)?.classList.add('active');
  currentPage = p;
  renderCurrentPage();
  if(p==='nova') resetNovaForm();
  if(p==='novaAtiv') resetAtivForm();
  if(p==='acessos') loadAcessos();
}

document.querySelectorAll('.sb-item[data-page]').forEach(a=>{
  a.addEventListener('click',()=>goPage(a.dataset.page));
});

function renderCurrentPage(){
  if(currentPage==='dashboard') renderDashboard();
  if(currentPage==='ocorrencias') renderOcorrencias();
  if(currentPage==='atividades') renderAtividades();
  if(currentPage==='relatorios') renderRelatorio();
  if(currentPage==='empresas') renderEmpresas();
  if(currentPage==='contratadas') renderContratadas();
  if(currentPage==='acessos') renderAcessos();
}

async function loadAcessos(){
  if(!requireAccessManager()) return;
  try{
    DATA.usuarios=await sbGet('usuarios','select=*&order=created_at.desc');
    renderAcessos();
  }catch(e){toast('Erro ao carregar acessos: '+e.message,'err');}
}

function renderAcessos(){
  const list=el('access-list');if(!list||!canManageAccess()) return;
  const q=norm(el('access-search')?.value||''),status=el('access-status')?.value||'';
  const rows=DATA.usuarios.filter(u=>(!q||norm(`${u.nome||''} ${u.email||''}`).includes(q))&&(!status||u.status_acesso===status));
  el('access-summary').textContent=`${rows.length} solicitação${rows.length!==1?'ões':'ão'} exibida${rows.length!==1?'s':''}`;
  list.innerHTML=rows.length?rows.map(u=>{
    const self=u.id===AUTH_USER?.id;
    return `<div class="access-card">
      <div class="access-person"><div class="access-avatar">${escapeHtml((u.nome||u.email||'?').charAt(0).toUpperCase())}</div><div><b>${escapeHtml(u.nome||'Sem nome')}</b><span>${escapeHtml(u.email||'')}</span></div></div>
      <div class="access-fields">
        <select id="perfil-${u.id}" ${self?'disabled':''}><option value="leitor" ${u.perfil==='leitor'?'selected':''}>Leitor</option><option value="colaborador" ${u.perfil==='colaborador'?'selected':''}>Colaborador</option><option value="admin" ${u.perfil==='admin'?'selected':''}>Administrador</option></select>
        <select id="status-${u.id}" ${self?'disabled':''}><option value="pendente" ${u.status_acesso==='pendente'?'selected':''}>Pendente</option><option value="aprovado" ${u.status_acesso==='aprovado'?'selected':''}>Aprovado</option><option value="rejeitado" ${u.status_acesso==='rejeitado'?'selected':''}>Rejeitado</option></select>
        <label class="access-active"><input id="ativo-${u.id}" type="checkbox" ${u.ativo!==false?'checked':''} ${self?'disabled':''}> Ativo</label>
        ${!self&&u.status_acesso==='pendente'?`<button class="btn-sec" style="color:#b91c1c;border-color:#fecaca" onclick="adminAtualizarUsuario('${u.id}','rejeitado')">Rejeitar</button><button class="btn-primary" onclick="adminAtualizarUsuario('${u.id}','aprovado')">Aprovar</button>`:`<button class="btn-primary" ${self?'disabled title="Seu próprio acesso não pode ser alterado aqui"':''} onclick="adminAtualizarUsuario('${u.id}')">Salvar</button>`}
      </div>
    </div>`;
  }).join(''):'<div class="empty-state">Nenhuma solicitação encontrada.</div>';
}

async function adminAtualizarUsuario(id,statusOverride){
  if(!requireAccessManager()) return;
  const perfil=el(`perfil-${id}`).value,status=statusOverride||el(`status-${id}`).value,ativo=el(`ativo-${id}`).checked;
  try{
    const r=await fetch(`${SB_URL}/rest/v1/rpc/admin_atualizar_usuario`,{method:'POST',headers:authHeaders(),body:JSON.stringify({p_usuario_id:id,p_status:status,p_perfil:perfil,p_ativo:ativo})});
    const data=await r.json().catch(()=>null);if(!r.ok)throw new Error(data?.message||data?.error||`Falha (${r.status})`);
    toast('Acesso atualizado com sucesso.');await loadAcessos();
  }catch(e){toast('Erro ao atualizar acesso: '+e.message,'err');}
}

// ── DASHBOARD ──
function renderDashboard(){
  const oc=DATA.ocorrencias;
  const total=oc.length;
  const abertas=oc.filter(r=>r.status==='aberta').length;
  const concluidas=oc.filter(r=>r.status==='concluida').length;
  const criticas=oc.filter(r=>r.criticidade==='alta').length;
  const vencidas=oc.filter(r=>isVencida(r)).length;
  const taxa=total?Math.round((concluidas+oc.filter(r=>r.status==='fechada').length)/total*100):0;

  const fechadas=oc.filter(r=>r.status==='fechada').length;
  const pAb=total?Math.round(abertas/total*100):0;
  const pCo=total?Math.round(concluidas/total*100):0;
  const pCr=total?Math.round(criticas/total*100):0;
  el('dash-kpis').innerHTML=`
    ${kpiCard('Total de ocorrências',total,'','',true,new Date().toLocaleDateString('pt-BR',{month:'long',year:'numeric'}))}
    ${kpiCard('Em aberto',abertas,'','','',pAb+'% do total',pAb,abertas>3?'danger':abertas>1?'warn':'')}
    ${kpiCard('Concluídas',concluidas,'','','',pCo+'% de resolução',pCo)}
    ${kpiCard('Criticidade alta',criticas,'','','',pCr+'% do total',pCr)}
    ${kpiCard('Vencidas',vencidas,'','','',vencidas?'Atenção imediata':'Sem atrasos',vencidas?Math.min(100,vencidas*20):0,vencidas?'warn':'')}
    ${kpiCard('Taxa de resolução',taxa+'%','','','',''+concluidas+fechadas+' de '+total+' encerradas',taxa,taxa>=80?'success':taxa>=50?'warn':'danger')}
  `;

  const byEmp = groupBy(oc,'empresa');
  const empS=Object.entries(byEmp).sort((a,b)=>a[1].length-b[1].length);
  mkChart('chartEmpresa','bar',empS.map(e=>e[0]),empS.map(e=>e[1].length),null,'horizontal');

  // Criticidade
  const byCrit={'Alta':criticas,'Média':oc.filter(r=>r.criticidade==='media').length,'Baixa':oc.filter(r=>r.criticidade==='baixa').length};
  mkChart('chartCrit','doughnut',Object.keys(byCrit),Object.values(byCrit),['#c0392b','#c8841a','#00a89b']);

  // Mensal
  const byMes={};
  oc.forEach(r=>{
    if(!r.data_abertura) return;
    const m=r.data_abertura.slice(0,7);
    byMes[m]=(byMes[m]||0)+1;
  });
  const meses=Object.keys(byMes).sort();
  mkChart('chartMensal','line',meses.map(m=>{const[y,mo]=m.split('-');return `${['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][+mo-1]}/${y.slice(2)}`}),meses.map(m=>byMes[m]),'#00c6b8');

  // Status
  const bySt={'Aberta':abertas,'Concluída':concluidas,'Fechada':oc.filter(r=>r.status==='fechada').length};
  mkChart('chartStatus','doughnut',Object.keys(bySt),Object.values(bySt),['#1a4fa0','#00a89b','#8baec8']);

  // Recentes
  const rec=oc.filter(r=>r.status==='aberta').slice(0,5);
  el('dash-recentes').innerHTML = rec.length
    ? `<table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead><tr>${['Título','Empresa','Criticidade','Prazo'].map(h=>`<th style="text-align:left;padding:8px 10px;border-bottom:1px solid var(--line);font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase">${h}</th>`).join('')}</tr></thead>
        <tbody>${rec.map(r=>`<tr style="cursor:pointer" onclick="verOcorrencia('${r.id}')">
          <td style="padding:10px;border-bottom:1px solid var(--line)">${r.titulo.slice(0,50)}${r.titulo.length>50?'…':''}</td>
          <td style="padding:10px;border-bottom:1px solid var(--line)">${r.empresa||'—'}</td>
          <td style="padding:10px;border-bottom:1px solid var(--line)">${badge('crit',r.criticidade)}</td>
          <td style="padding:10px;border-bottom:1px solid var(--line);${isVencida(r)?'color:#ef4444;font-weight:700':''}">${fmt(r.prazo)}</td>
        </tr>`).join('')}</tbody>
      </table>`
    : `<div style="text-align:center;padding:30px;color:var(--muted);font-size:13.5px">Nenhuma ocorrência aberta 🎉</div>`;
}

function kpiCard(label,val,_ic,_col,hl=false,hint='',pct=-1,variant=''){
  const cls=hl?'hl':(variant||'');
  const bar=pct>=0?`<div class="kpi-bar"><div class="kpi-bar-fill" style="width:${Math.min(100,Math.round(pct))}%"></div></div>`:'';
  if(hl) return `<div class="kpi-card hl"><div class="kpi-label">${label}</div><div class="kpi-val">${val}</div>${hint?`<div class="kpi-hint">${hint}</div>`:''}</div>`;
  return `<div class="kpi-card ${cls}"><div class="kpi-label">${label}</div><div class="kpi-val">${val}</div>${hint?`<div class="kpi-hint">${hint}</div>`:''} ${bar}</div>`;
}

function groupBy(arr,key){
  const r={};arr.forEach(a=>{const k=a[key]||'Desconhecido';r[k]=(r[k]||[]);r[k].push(a);});return r;
}
const EXEC=['#1a4fa0','#00c6b8','#0d3d7a','#00a89b','#1e56c4','#2ec4b6','#154090','#17b8ab'];
function colorArr(i){return EXEC[i%EXEC.length];}
function mkChart(id,type,labels,data,colors,dir){
  if(charts[id]) charts[id].destroy();
  const cv=el(id);if(!cv) return;
  const ctx=cv.getContext('2d');
  const isH=dir==='horizontal',isD=type==='doughnut',isL=type==='line',isB=type==='bar';
  const BF={family:"'Inter',system-ui,sans-serif",size:11,weight:'500'};
  let bg,bc,bw=0;
  if(isD){bg=['#1a4fa0','#00c6b8','#0d3d7a','#00a89b','#1e56c4','#2ec4b6'];bc=bg.map(()=>'rgba(255,255,255,.4)');bw=0.5;}
  else if(isL){const g=ctx.createLinearGradient(0,0,0,230);g.addColorStop(0,'rgba(26,79,160,.16)');g.addColorStop(1,'rgba(0,198,184,.02)');bg=g;bc='#1a4fa0';bw=2.5;}
  else{bg=data.map((_,i)=>{const g=ctx.createLinearGradient(isH?200:0,0,0,isH?0:200);const p=[['#1a4fa0','#2563eb'],['#00a89b','#00c6b8'],['#0d3d7a','#1a6fa0'],['#065f46','#00a89b'],['#1e3a8a','#3b82f6'],['#6d28d9','#8b5cf6']];const[c1,c2]=p[i%p.length];g.addColorStop(0,c1);g.addColorStop(1,c2);return g;});bc=data.map((_,i)=>EXEC[i%EXEC.length]);bw=0;}
  charts[id]=new Chart(ctx,{type,data:{labels,datasets:[{data,backgroundColor:bg,borderColor:bc,borderWidth:bw,borderRadius:isB?5:0,borderSkipped:false,tension:.42,fill:isL,pointRadius:isL?0:0,pointHoverRadius:isL?5:0,pointBackgroundColor:'#fff',pointBorderColor:'#1a4fa0',pointBorderWidth:2,hoverOffset:isD?8:0,cutout:isD?'66%':undefined}]},
  options:{maintainAspectRatio:false,indexAxis:isH?'y':'x',animation:{duration:500,easing:'easeInOutQuart'},
  plugins:{
    legend:{display:isD,position:'bottom',labels:{font:{...BF,size:12},padding:16,color:'rgba(255,255,255,.8)',usePointStyle:true,pointStyle:'circle',boxWidth:10}},
    tooltip:{backgroundColor:'#0d2137',titleColor:'rgba(139,174,200,.9)',bodyColor:'#fff',borderColor:'rgba(0,198,184,.25)',borderWidth:1,padding:11,cornerRadius:7,callbacks:{label:c=>{const tot=Array.isArray(data)?data.reduce((a,b)=>a+b,0):c.raw;return ` ${c.raw} (${tot?Math.round(c.raw/tot*100):0}%)`;},title:t=>t[0].label}}
  },
  scales:isD?{}:{
    x:{grid:{color:isH?'rgba(226,232,240,.5)':'transparent',lineWidth:.5},border:{display:false},ticks:{font:{...BF},color:'#8baec8',padding:6,maxTicksLimit:8}},
    y:{grid:{color:isH?'transparent':'rgba(226,232,240,.5)',lineWidth:.5},border:{display:false},ticks:{font:{...BF},color:'#8baec8',padding:6,callback:v=>Number.isInteger(v)?v:null}}
  }}});
}

// ── OCORRÊNCIAS ──
function filteredOc(){
  const q=norm(el('occSearch')?.value||'');
  const fc=el('fCrit')?.value||'';
  const fs=el('fStatus')?.value||'';
  const fe=el('fEmpresa')?.value||'';
  const di=el('fDtI')?.value||'',df=el('fDtF')?.value||'';
  const fo=el('fOrdem')?.value||'recente';
  let arr=DATA.ocorrencias.filter(r=>{
    if(q&&!norm(r.titulo+' '+r.empresa+' '+r.tecnico+' '+r.frente_servico).includes(q)) return false;
    if(fc&&r.criticidade!==fc) return false;
    if(fs&&r.status!==fs) return false;
    if(fe&&r.empresa!==fe) return false;
    const data=(r.data_abertura||r.criado_em||'').slice(0,10);
    if(di&&(!data||data<di)) return false;
    if(df&&(!data||data>df)) return false;
    return true;
  });
  arr.sort((a,b)=>{
    if(fo==='crit') return critOrder(a.criticidade)-critOrder(b.criticidade);
    if(fo==='antiga') return new Date(a.criado_em)-new Date(b.criado_em);
    return new Date(b.criado_em)-new Date(a.criado_em);
  });
  return arr;
}

function renderOcorrencias(){
  const arr=filteredOc();
  el('occ-tbody').innerHTML=arr.length
    ? arr.map(r=>`<tr style="cursor:pointer">
        <td onclick="verOcorrencia('${r.id}')" style="max-width:200px"><b>${r.titulo.slice(0,45)}${r.titulo.length>45?'…':''}</b></td>
        <td>${r.empresa||'—'}</td>
        <td>${r.tecnico||'—'}</td>
        <td>${r.encarregado||'—'}</td>
        <td>${r.fiscal||'—'}</td>
        <td>${badge('crit',r.criticidade)}</td>
        <td>${badge('st',r.status)}</td>
        <td style="${isVencida(r)?'color:#fca5a5;font-weight:700':''}">${fmt(r.prazo)}</td>
        <td style="white-space:nowrap">
          <button class="btn-sec" style="padding:5px 10px;font-size:12px" onclick="verOcorrencia('${r.id}')">Ver</button>
          ${canManage()&&podeEnviarContratada(r)?`<button class="btn-sec" style="padding:5px 9px;margin-left:4px;color:var(--teal);border-color:var(--teal)" onclick="abrirNotificacao('${r.id}')" title="Enviar para Contratada">✉</button>`:''}
          ${canManage()?`<button class="btn-danger" style="margin-left:4px" onclick="confirmarExcluir('${r.id}','ocorrencia')">🗑</button>`:''}
        </td>
      </tr>`).join('')
    : '<tr><td colspan="9" style="text-align:center;padding:40px;color:rgba(255,255,255,.5)">Nenhuma ocorrência encontrada.</td></tr>';
  const _oc=el('occ-count');if(_oc){_oc.textContent=`${arr.length} ocorrência${arr.length!==1?'s':''} encontrada${arr.length!==1?'s':''}`;_oc.style.color='rgba(255,255,255,.55)';}
}

function globalSearchFn(v){
  const q=norm(v);
  if(!q){el('occSearch')&&(el('occSearch').value='');renderOcorrencias();return;}
  if(el('occSearch')) el('occSearch').value=v;
  goPage('ocorrencias');
}

// ── VER OCORRÊNCIA ──
async function verOcorrencia(id){
  const r=DATA.ocorrencias.find(x=>x.id===id);
  if(!r) return;
  const cobrancas=DATA.historicoEmails.filter(h=>h.ocorrencia_id===id);
  ACTION_VALIDATION_STATE.set(id,{existingPhotos:0,existingCorrectivePhotos:0,loading:true,saved:false});
  el('modal-title').textContent=r.titulo;
  el('modal-body').innerHTML=`
    <div style="display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap">
      ${badge('crit',r.criticidade)} ${badge('st',r.status)}
      ${isVencida(r)?'<span class="badge" style="background:#fef2f2;color:#dc2626">⏰ Vencida</span>':''}
    </div>
    <div class="detail-row"><div class="detail-key">Empresa</div><div class="detail-val">${r.empresa||'—'}</div></div>
    <div class="detail-row"><div class="detail-key">Frente de Serviço</div><div class="detail-val">${r.frente_servico||'—'}</div></div>
    <div class="detail-row"><div class="detail-key">Técnico</div><div class="detail-val">${r.tecnico||'—'}</div></div>
    <div class="detail-row"><div class="detail-key">Encarregado</div><div class="detail-val">${r.encarregado||'—'}</div></div>
    <div class="detail-row"><div class="detail-key">Fiscal de Campo</div><div class="detail-val">${r.fiscal||'—'}</div></div>
    <div class="detail-row"><div class="detail-key">Abertura</div><div class="detail-val">${fmt(r.data_abertura)}</div></div>
    <div class="detail-row"><div class="detail-key">Prazo</div><div class="detail-val" style="${isVencida(r)?'color:#ef4444;font-weight:700':''}">${fmt(r.prazo)}</div></div>
    ${r.descricao?`<div style="margin-top:14px"><div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">Descrição</div><div style="font-size:13.5px;line-height:1.6;white-space:pre-wrap;background:#f8fafc;padding:12px;border-radius:8px">${r.descricao}</div></div>`:''}
    ${r.acao_corretiva?`<div style="margin-top:12px"><div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">Ação Corretiva</div><div style="font-size:13.5px;line-height:1.6;white-space:pre-wrap;background:#f0fdf4;padding:12px;border-radius:8px">${r.acao_corretiva}</div></div>`:''}
    <div class="cobranca-list">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px">
        <div><div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em">Cobranças enviadas</div><div style="font-size:11px;color:var(--muted);margin-top:3px">Histórico de notificações desta ocorrência.</div></div>
        ${canManage()&&podeEnviarContratada(r)?`<button class="btn-sec" style="color:var(--teal);border-color:var(--teal);font-size:11px" onclick="abrirNotificacao('${r.id}')">✉ Nova cobrança</button>`:''}
      </div>
      ${cobrancas.length?cobrancas.map(h=>`<div class="cobranca-item"><span>${new Date(h.data_envio).toLocaleString('pt-BR')}</span><span title="${escapeHtml((h.destinatarios||[]).join(', '))}">${escapeHtml((h.destinatarios||[]).join(', '))}</span><b style="color:${['enviado','enviado_manual'].includes(h.status_envio)?'#15803d':'#dc2626'}">${h.status_envio==='enviado_manual'?'Enviado manualmente':h.status_envio==='enviado'?'Enviado':'Falha'}</b></div>`).join(''):'<div style="font-size:12px;color:var(--muted);padding:14px 0">Nenhuma cobrança enviada.</div>'}
    </div>
    ${canManage()&&r.status==='aberta'?`
    <div class="validation-card">
      <div class="validation-eyebrow">Validação para conclusão</div>
      <div class="validation-title">A ação corretiva foi realizada?</div>
      <div class="validation-help">Confirme a execução da ação corretiva antes de concluir a ocorrência.</div>
      <div class="validation-options">
        <label class="validation-option">
          <input type="radio" name="acao-realizada-${r.id}" value="sim" onchange="responderAcaoCorretiva('${r.id}',true)">
          Sim
        </label>
        <label class="validation-option">
          <input type="radio" name="acao-realizada-${r.id}" value="nao" onchange="responderAcaoCorretiva('${r.id}',false)">
          Não
        </label>
      </div>
      <div class="validation-feedback" id="validation-feedback-${r.id}">A ocorrência permanecerá com o status Aberta.</div>
      <div id="validation-action-${r.id}" style="display:none;margin-top:14px;padding-top:14px;border-top:1px solid #cbd5e1">
        <div class="validation-title">Registre a ação corretiva</div>
        <div class="validation-help">Preencha a descrição e anexe a evidência fotográfica antes de concluir.</div>

        <div style="margin-top:12px">
          <label for="acao-descricao-${r.id}" style="display:block;margin-bottom:6px;font-size:12px;font-weight:700;color:#334155">
            Descreva a ação adotada <span style="color:#dc2626">*</span>
          </label>
          <textarea
            id="acao-descricao-${r.id}"
            rows="4"
            placeholder="Informe detalhadamente a ação corretiva executada..."
            oninput="invalidarAcaoCorretiva('${r.id}')"
            style="width:100%;resize:vertical;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:#0f172a;font:inherit"
          >${escapeHtml(r.acao_corretiva||'')}</textarea>
        </div>

        <div style="margin-top:12px">
          <div style="font-size:12px;font-weight:700;color:#334155;margin-bottom:6px">
            Evidência fotográfica <span style="color:#dc2626">*</span>
          </div>
          <button type="button" class="btn-sec" onclick="el('acao-fotos-${r.id}').click()" style="font-size:12px;padding:8px 12px">
            📷 Adicionar foto
          </button>
          <input
            type="file"
            id="acao-fotos-${r.id}"
            accept="image/*"
            multiple
            style="display:none"
            onchange="validarFotosAcao('${r.id}',this)"
          >
          <div id="acao-contador-${r.id}" style="margin-top:7px;font-size:11px;color:#64748b">
            Verificando evidências cadastradas...
          </div>
          <div id="acao-preview-${r.id}" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(82px,1fr));gap:8px;margin-top:9px"></div>
        </div>

        <button
          type="button"
          id="salvar-acao-${r.id}"
          class="btn-primary"
          onclick="salvarAcaoCorretiva('${r.id}')"
          style="margin-top:13px;font-size:12px;padding:9px 14px"
        >
          Salvar ação corretiva
        </button>
      </div>
      <div class="validation-step2" id="validation-step2-${r.id}">
        <div class="validation-title">Deseja concluir esta ocorrência?</div>
        <div class="validation-help">Ao confirmar, o status será alterado automaticamente para Concluída.</div>
        <div class="validation-options">
          <label class="validation-option" id="btn-concluir-${r.id}">
            <input type="radio" name="concluir-ocorrencia-${r.id}" value="sim" onchange="concluirOcorrencia('${r.id}')">
            Sim
          </label>
          <label class="validation-option">
            <input type="radio" name="concluir-ocorrencia-${r.id}" value="nao" onchange="naoConcluirOcorrencia('${r.id}')">
            Não
          </label>
        </div>
      </div>
    </div>`:''}
    ${canManage()&&r.status!=='aberta'?`<div style="margin-top:16px">
      <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px">Alterar Status</div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        ${r.status!=='aberta'?`<button onclick="mudarStatus('${r.id}','aberta')" class="btn-sec" style="font-size:12px;padding:6px 12px;color:#123f9d;border-color:#123f9d">↺ Reabrir ocorrência</button>`:''}
        ${r.status!=='fechada'?`<button onclick="mudarStatus('${r.id}','fechada')" class="btn-sec" style="font-size:12px;padding:6px 12px">Fechar ocorrência</button>`:'<span class="badge" style="background:#f1f5f9;color:#475569">Ocorrência fechada</span>'}
      </div>
    </div>`:''}
    <div style="margin-top:18px;border-top:1px solid var(--line);padding-top:14px">
      <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:10px">Evidências</div>
      <div id="modal-anexos" style="font-size:13px;color:var(--muted)">Carregando...</div>
    </div>`;
  el('modal-foot').innerHTML=`
    ${canManage()?`<button class="btn-danger" onclick="confirmarExcluir('${r.id}','ocorrencia');closeModal()">🗑 Excluir</button>`:''}
    ${canManage()&&podeEnviarContratada(r)?`<button class="btn-primary" onclick="abrirNotificacao('${r.id}')">✉ Enviar para Contratada</button>`:''}
    <button class="btn-sec" onclick="imprimirOcorrencia('${r.id}')" style="gap:5px">🖨 Imprimir</button>
    <button class="btn-sec" onclick="closeModal()">Fechar</button>`;
  el('overlay').classList.add('show');

  // Busca anexos do bucket evidencias
  try{
    const res=await fetch(`${SB_URL}/storage/v1/object/list/evidencias`,{
      method:'POST',
      headers:authHeaders(),
      body:JSON.stringify({prefix:r.id+'/',limit:50,sortBy:{column:'created_at',order:'desc'}})
    });
    const files=res.ok?await res.json():[];
    const existingPhotos=(files||[]).filter(f=>/\.(jpg|jpeg|png|gif|webp|heic)$/i.test(f.name||'')).length;
    const existingCorrectivePhotos=(files||[]).filter(f=>/_acao_corretiva_/i.test(f.name||'')&&/\.(jpg|jpeg|png|gif|webp|heic)$/i.test(f.name||'')).length;
    const actionState=ACTION_VALIDATION_STATE.get(r.id)||{};
    ACTION_VALIDATION_STATE.set(r.id,{...actionState,existingPhotos,existingCorrectivePhotos,loading:false});
    atualizarContadorAcao(r.id);
    const div=document.getElementById('modal-anexos');
    if(!div) return;
    if(!files||!files.length){div.textContent='Nenhuma evidência anexada.';return;}
    const anexos=await Promise.all(files.map(async f=>({f,url:await storageSignedUrl('evidencias',`${r.id}/${f.name}`)})));
    div.innerHTML=`<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:10px">
      ${anexos.map(({f,url})=>{
        const isImg=/\.(jpg|jpeg|png|gif|webp|heic)$/i.test(f.name);
        const isPdf=/\.pdf$/i.test(f.name);
        if(isImg) return `<a href="${url}" target="_blank" style="display:block;border-radius:9px;overflow:hidden;border:2px solid var(--line);aspect-ratio:1">
          <img src="${url}" alt="${f.name}" style="width:100%;height:100%;object-fit:cover" onerror="this.parentElement.innerHTML='<div style=padding:12px;font-size:11px;color:var(--muted)>'+decodeURIComponent('${encodeURIComponent(f.name)}')+'</div>'">
        </a>`;
        if(isPdf) return `<a href="${url}" target="_blank" style="display:flex;flex-direction:column;align-items:center;justify-content:center;border-radius:9px;border:2px solid var(--line);padding:14px 8px;text-decoration:none;gap:6px;background:#f8fafc">
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#e34948" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
          <span style="font-size:10px;color:#64748b;text-align:center;word-break:break-all;line-height:1.3">${f.name.slice(0,30)}</span>
        </a>`;
        return `<a href="${url}" target="_blank" style="display:flex;flex-direction:column;align-items:center;justify-content:center;border-radius:9px;border:2px solid var(--line);padding:14px 8px;text-decoration:none;gap:6px;background:#f8fafc">
          <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><polyline points="13 2 13 9 20 9"/></svg>
          <span style="font-size:10px;color:#64748b;text-align:center;word-break:break-all;line-height:1.3">${f.name.slice(0,30)}</span>
        </a>`;
      }).join('')}
    </div>`;
  }catch(e){
    const actionState=ACTION_VALIDATION_STATE.get(r.id)||{};
    ACTION_VALIDATION_STATE.set(r.id,{...actionState,loading:false});
    atualizarContadorAcao(r.id);
    const div=document.getElementById('modal-anexos');
    if(div) div.textContent='Erro ao carregar anexos.';
  }
}

async function mudarStatus(id,status){
  if(!requireManage()) return;
  const ocorrencia=DATA.ocorrencias.find(x=>x.id===id);
  if(status==='concluida'){
    toast('A conclusão exige a validação da ação corretiva.','warn');
    return;
  }
  if(status==='fechada'&&ocorrencia?.status==='aberta'){
    toast('Registre a ação corretiva e conclua a ocorrência antes de fechá-la.','warn');
    return;
  }
  try{
    await sbUpdate('ocorrencias',id,{status});
    const i=DATA.ocorrencias.findIndex(x=>x.id===id);
    if(i>-1) DATA.ocorrencias[i].status=status;
    toast('Status atualizado!');
    closeModal();
    renderCurrentPage();
  }catch(e){toast('Erro: '+e.message,'err');}
}

function responderAcaoCorretiva(id,realizada){
  const action=el(`validation-action-${id}`);
  const step2=el(`validation-step2-${id}`);
  const feedback=el(`validation-feedback-${id}`);
  if(!action||!step2||!feedback) return;
  action.style.display=realizada?'block':'none';
  step2.classList.remove('show');
  feedback.classList.toggle('show',!realizada);
  if(realizada){
    feedback.classList.remove('show');
    atualizarContadorAcao(id);
  }else{
    feedback.textContent='A ocorrência permanecerá com o status Aberta.';
    document.querySelectorAll(`input[name="concluir-ocorrencia-${id}"]`).forEach(r=>r.checked=false);
    const state=ACTION_VALIDATION_STATE.get(id)||{};
    ACTION_VALIDATION_STATE.set(id,{...state,saved:false});
  }
}

function invalidarAcaoCorretiva(id){
  const state=ACTION_VALIDATION_STATE.get(id)||{};
  ACTION_VALIDATION_STATE.set(id,{...state,saved:false});
  el(`validation-step2-${id}`)?.classList.remove('show');
  document.querySelectorAll(`input[name="concluir-ocorrencia-${id}"]`).forEach(r=>r.checked=false);
}

function atualizarContadorAcao(id){
  const state=ACTION_VALIDATION_STATE.get(id)||{};
  const input=el(`acao-fotos-${id}`);
  const selected=Array.from(input?.files||[]).length;
  const existing=Number(state.existingPhotos||0);
  const total=Math.min(4,existing+selected);
  const remaining=Math.max(0,4-existing);
  const counter=el(`acao-contador-${id}`);
  if(!counter) return;
  if(state.loading){
    counter.textContent='Verificando evidências cadastradas...';
    return;
  }
  counter.textContent=`Evidências: ${total}/4 fotos cadastradas ou selecionadas. ${remaining?`Você pode adicionar até ${remaining} foto${remaining!==1?'s':''}.`:'Limite máximo de 4 evidências atingido.'}`;
}

function validarFotosAcao(id,input){
  invalidarAcaoCorretiva(id);
  const state=ACTION_VALIDATION_STATE.get(id)||{};
  const existing=Number(state.existingPhotos||0);
  const available=Math.max(0,4-existing);
  const chosen=Array.from(input.files||[]);
  const images=chosen.filter(file=>file.type.startsWith('image/'));

  if(!available){
    input.value='';
    toast('Limite máximo de 4 evidências atingido.','warn');
    atualizarContadorAcao(id);
    return;
  }
  if(images.length!==chosen.length) toast('Selecione somente arquivos de imagem.','warn');
  if(images.length>available) toast(`Você pode adicionar somente ${available} foto${available!==1?'s':''}.`,'warn');

  const accepted=images.slice(0,available);
  const transfer=new DataTransfer();
  accepted.forEach(file=>transfer.items.add(file));
  input.files=transfer.files;

  const preview=el(`acao-preview-${id}`);
  if(preview){
    preview.innerHTML=accepted.map(file=>`
      <div style="position:relative;aspect-ratio:1;border:1px solid #cbd5e1;border-radius:8px;overflow:hidden;background:#f8fafc">
        <img src="${URL.createObjectURL(file)}" alt="${escapeHtml(file.name)}" style="width:100%;height:100%;object-fit:cover">
      </div>
    `).join('');
  }
  atualizarContadorAcao(id);
}

async function uploadFotosAcao(files,recordId){
  for(const file of files){
    const safe=file.name.replace(/[^a-zA-Z0-9._-]/g,'_');
    const unique=globalThis.crypto?.randomUUID?.()||`${Date.now()}_${Math.random().toString(36).slice(2)}`;
    const path=`${recordId}/${unique}_acao_corretiva_${safe}`;
    const res=await fetch(`${SB_URL}/storage/v1/object/evidencias/${path}`,{
      method:'POST',
      headers:{...authHeaders(null),'Content-Type':file.type,'x-upsert':'false'},
      body:file
    });
    if(!res.ok){
      const message=await res.text().catch(()=> '');
      throw new Error(message||'Não foi possível salvar a evidência fotográfica.');
    }
  }
}

async function salvarAcaoCorretiva(id){
  if(!requireManage()) return;
  const description=el(`acao-descricao-${id}`)?.value.trim()||'';
  const input=el(`acao-fotos-${id}`);
  const files=Array.from(input?.files||[]);
  const state=ACTION_VALIDATION_STATE.get(id)||{};
  const existing=Number(state.existingPhotos||0);
  const existingCorrective=Number(state.existingCorrectivePhotos||0);

  if(!description){
    toast('Descreva a ação adotada antes de continuar.','warn');
    el(`acao-descricao-${id}`)?.focus();
    return;
  }
  if(state.loading){
    toast('Aguarde a verificação das evidências cadastradas.','warn');
    return;
  }
  if(existingCorrective+files.length<1){
    toast('Adicione pelo menos uma evidência fotográfica.','warn');
    return;
  }
  if(existing+files.length>4){
    toast('Limite máximo de 4 evidências atingido.','warn');
    return;
  }

  const button=el(`salvar-acao-${id}`);
  if(button){button.disabled=true;button.textContent='Salvando...';}

  try{
    await sbUpdate('ocorrencias',id,{acao_corretiva:description});
    if(files.length) await uploadFotosAcao(files,id);

    const index=DATA.ocorrencias.findIndex(x=>x.id===id);
    if(index>-1) DATA.ocorrencias[index].acao_corretiva=description;

    ACTION_VALIDATION_STATE.set(id,{
      ...state,
      existingPhotos:existing+files.length,
      existingCorrectivePhotos:existingCorrective+files.length,
      loading:false,
      saved:true
    });
    if(input) input.value='';
    const preview=el(`acao-preview-${id}`);
    if(preview) preview.innerHTML='';
    atualizarContadorAcao(id);

    const feedback=el(`validation-feedback-${id}`);
    if(feedback){
      feedback.textContent='Ação corretiva e evidência salvas. Agora você pode concluir a ocorrência.';
      feedback.classList.add('show');
    }
    el(`validation-step2-${id}`)?.classList.add('show');
    toast('Ação corretiva salva com sucesso!');
  }catch(e){
    toast('Erro ao salvar a ação corretiva: '+e.message,'err');
  }finally{
    if(button){button.disabled=false;button.textContent='Salvar ação corretiva';}
  }
}

function naoConcluirOcorrencia(id){
  const feedback=el(`validation-feedback-${id}`);
  if(feedback){
    feedback.textContent='A ação corretiva foi confirmada, mas a ocorrência permanecerá com o status Aberta.';
    feedback.classList.add('show');
  }
  toast('Ocorrência mantida como Aberta.','warn');
}

async function concluirOcorrencia(id){
  if(!requireManage()) return;
  const confirmacao=document.querySelector(`input[name="acao-realizada-${id}"][value="sim"]:checked`);
  if(!confirmacao){
    toast('Confirme primeiro que a ação corretiva foi realizada.','warn');
    return;
  }
  const state=ACTION_VALIDATION_STATE.get(id)||{};
  if(!state.saved){
    toast('Salve a descrição e a evidência fotográfica antes de concluir.','warn');
    return;
  }
  const btn=el(`btn-concluir-${id}`);
  if(btn){
    btn.style.pointerEvents='none';
    btn.style.opacity='.65';
  }
  try{
    await sbUpdate('ocorrencias',id,{status:'concluida'});
    const i=DATA.ocorrencias.findIndex(x=>x.id===id);
    if(i>-1) DATA.ocorrencias[i].status='concluida';
    toast('Ocorrência concluída com sucesso!');
    closeModal();
    renderCurrentPage();
  }catch(e){
    if(btn){
      btn.style.pointerEvents='';
      btn.style.opacity='';
    }
    toast('Erro ao concluir: '+e.message,'err');
  }
}

// ── NOVA OCORRÊNCIA ──
function resetNovaForm(){
  ['n_emp','n_tit','n_fren','n_tec','n_fisc','n_encar','n_desc','n_acao'].forEach(id=>{const e=el(id);if(e)e.value='';});
  const oi=el('occ-files');if(oi)oi.value='';const og=el('occ-grid');if(og)og.innerHTML='';
  el('n_crit').value='media';
  el('n_st').value='aberta';
  el('n_data').value=today();
  el('n_prazo').value=today();
}

async function salvarOcorrencia(){
  if(!requireManage()) return;
  const emp=el('n_emp').value.trim();
  const tit=el('n_tit').value.trim();
  if(!tit){toast('Informe o título da ocorrência','warn');return;}
  const contratada=DATA.contratadas.find(c=>norm(c.nome_fantasia)===norm(emp)||norm(c.razao_social)===norm(emp));
  const row={
    empresa:normEmp(emp)||emp,titulo:tit,
    contratada_id:contratada?.id||null,
    frente_servico:el('n_fren').value.trim()||null,
    tecnico:normTec(el('n_tec').value.trim())||null,
    fiscal:el('n_fisc').value.trim()||null,
    encarregado:el('n_encar')?.value.trim()||null,
    criticidade:el('n_crit').value,
    status:el('n_st').value,
    data_abertura:el('n_data').value||null,
    prazo:el('n_prazo').value||null,
    descricao:el('n_desc').value.trim()||null,
    acao_corretiva:el('n_acao').value.trim()||null,
  };
  try{
    const saved=await sbInsert('ocorrencias',row);
    const rec=Array.isArray(saved)?saved[0]:saved;
    DATA.ocorrencias.unshift({...row,...rec});
    populateFilters();
    const files=Array.from(el('occ-files')?.files||[]);
    if(files.length){
      toast(`Salvando ${files.length} evidência${files.length>1?'s':''}...`,'ok');
      await uploadFiles(files,'evidencias',rec.id,'occ-prog','occ-prog-fill');
    }
    toast('Ocorrência registrada com sucesso! ✓');
    goPage('ocorrencias');
  }catch(e){toast('Erro ao salvar: '+e.message,'err');console.error(e);}
}

// ── ATIVIDADES ──
function renderAtividades(){
  const q=norm(el('ativSearch')?.value||'');
  const ft=el('fTipo')?.value||'';
  let arr=DATA.atividades.filter(r=>{
    if(q&&!norm((r.titulo||'')+' '+(r.empresa||'')+' '+(r.responsavel||'')).includes(q)) return false;
    if(ft&&r.tipo!==ft) return false;
    return true;
  });
  el('ativ-tbody').innerHTML=arr.length
    ? arr.map(r=>`<tr>
        <td><b>${(r.titulo||'').slice(0,55)}</b></td>
        <td><span class="badge" style="background:#eff6ff;color:#3b82f6">${r.tipo||'—'}</span></td>
        <td>${r.empresa||'—'}</td>
        <td>${r.responsavel||'—'}</td>
        <td>${fmt(r.data_atividade)}</td>
        <td>${canManage()?`<button class="btn-danger" onclick="confirmarExcluir('${r.id}','atividade')">🗑</button>`:''}</td>
      </tr>`).join('')
    : '<tr><td colspan="6" style="text-align:center;padding:40px;color:var(--muted)">Nenhuma atividade registrada.</td></tr>';
}

function resetAtivForm(){
  ['a_tit','a_emp','a_fren','a_resp','a_desc'].forEach(id=>{const e=el(id);if(e)e.value='';});
  const ai=el('ativ-files');if(ai)ai.value='';const ag=el('ativ-grid');if(ag)ag.innerHTML='';
  el('a_tipo').value='DDS';
  el('a_data').value=today();
}

async function salvarAtividade(){
  if(!requireManage()) return;
  const tit=el('a_tit').value.trim();
  if(!tit){toast('Informe o título da atividade','warn');return;}
  const row={
    titulo:tit,tipo:el('a_tipo').value,
    empresa:el('a_emp').value.trim()||null,
    frente_servico:el('a_fren').value.trim()||null,
    responsavel:el('a_resp').value.trim()||null,
    data_atividade:el('a_data').value||null,
    descricao:el('a_desc').value.trim()||null,
  };
  try{
    const saved=await sbInsert('atividades_sst',row);
    const rec=Array.isArray(saved)?saved[0]:saved;
    DATA.atividades.unshift({...row,...rec});
    const files=Array.from(el('ativ-files')?.files||[]);
    if(files.length){
      toast(`Salvando ${files.length} evidência${files.length>1?'s':''}...`,'ok');
      await uploadFiles(files,'atividades',rec.id,'ativ-prog','ativ-prog-fill');
    }
    toast('Atividade registrada! ✓');
    goPage('atividades');
  }catch(e){toast('Erro ao salvar: '+e.message,'err');console.error(e);}
}

// ── RELATÓRIO ──
function filteredRel(){
  const di=el('rDtI')?.value,df=el('rDtF')?.value;
  const emp=el('rEmp')?.value,tec=el('rTec')?.value;
  const crit=el('rCrit')?.value,st=el('rSt')?.value;
  return DATA.ocorrencias.filter(r=>{
    if(di&&r.data_abertura&&r.data_abertura<di) return false;
    if(df&&r.data_abertura&&r.data_abertura>df) return false;
    if(emp&&r.empresa!==emp) return false;
    if(tec&&r.tecnico!==tec) return false;
    if(crit&&r.criticidade!==crit) return false;
    if(st&&r.status!==st) return false;
    return true;
  });
}

function renderRelatorio(){
  const arr=filteredRel();
  const abertas=arr.filter(r=>r.status==='aberta').length;
  const concluidas=arr.filter(r=>r.status==='concluida').length;
  const criticas=arr.filter(r=>r.criticidade==='alta').length;
  const vencidas=arr.filter(r=>isVencida(r)).length;
  el('rel-kpis').innerHTML=`
    ${kpiCard('Total',arr.length,'📋','#3b82f6',true)}
    ${kpiCard('Abertas',abertas,'⚠️','#f59e0b')}
    ${kpiCard('Concluídas',concluidas,'✅','#22c55e')}
    ${kpiCard('Críticas',criticas,'🔥','#ef4444')}
    ${kpiCard('Vencidas',vencidas,'⏰','#8b5cf6')}
  `;

  const byEmp=groupBy(arr,'empresa');
  mkChart('rChartEmp','bar',Object.keys(byEmp),Object.values(byEmp).map(v=>v.length),Object.values(byEmp).map((_,i)=>colorArr(i)),'horizontal');

  const byCrit={'Alta':criticas,'Média':arr.filter(r=>r.criticidade==='media').length,'Baixa':arr.filter(r=>r.criticidade==='baixa').length};
  mkChart('rChartCrit','doughnut',Object.keys(byCrit),Object.values(byCrit),['#ef4444','#f59e0b','#22c55e']);

  el('rel-tbody').innerHTML=arr.length
    ? arr.map(r=>`<tr>
        <td style="max-width:240px">${r.titulo.slice(0,50)}${r.titulo.length>50?'…':''}</td>
        <td>${r.empresa||'—'}</td><td>${r.tecnico||'—'}</td>
        <td>${badge('crit',r.criticidade)}</td><td>${badge('st',r.status)}</td>
        <td>${fmt(r.data_abertura)}</td><td>${fmt(r.prazo)}</td>
      </tr>`).join('')
    : '<tr><td colspan="7" style="text-align:center;padding:40px;color:var(--muted)">Sem dados para os filtros selecionados.</td></tr>';
}

function limparFiltrosRel(){
  ['rDtI','rDtF'].forEach(id=>{const e=el(id);if(e)e.value='';});
  ['rEmp','rTec','rCrit','rSt'].forEach(id=>{const e=el(id);if(e)e.value='';});
  renderRelatorio();
}

// ── PDF E ENVIO DO RELATÓRIO ──
function relatorioContexto(origem='relatorios'){
  const arr=origem==='ocorrencias'?filteredOc():filteredRel();
  const filtros=[];
  if(origem==='ocorrencias'){
    const periodo=el('fPeriodo')?.selectedOptions?.[0]?.textContent;
    if(el('fPeriodo')?.value) filtros.push(`Período: ${periodo}`);
    if(el('fDtI')?.value||el('fDtF')?.value) filtros.push(`Datas: ${fmt(el('fDtI').value)} a ${fmt(el('fDtF').value)}`);
    if(el('fEmpresa')?.value) filtros.push(`Empresa: ${el('fEmpresa').value}`);
    if(el('fCrit')?.value) filtros.push(`Criticidade: ${el('fCrit').selectedOptions[0].textContent}`);
    if(el('fStatus')?.value) filtros.push(`Status: ${el('fStatus').selectedOptions[0].textContent}`);
    if(el('occSearch')?.value) filtros.push(`Busca: ${el('occSearch').value}`);
  }else{
    if(el('rDtI')?.value||el('rDtF')?.value) filtros.push(`Datas: ${fmt(el('rDtI').value)} a ${fmt(el('rDtF').value)}`);
    [['rEmp','Empresa'],['rTec','Técnico'],['rCrit','Criticidade'],['rSt','Status']].forEach(([id,label])=>{if(el(id)?.value) filtros.push(`${label}: ${el(id).selectedOptions[0].textContent}`);});
  }
  return {arr,filtros:filtros.length?filtros:['Sem filtros — todo o período']};
}

const PDF_TEMPLATE_DATA=window.ENGSAFE_PDF_TEMPLATE;

function gerarRelatorioPDF(origem='relatorios'){
  if(!window.jspdf?.jsPDF) throw new Error('Biblioteca de PDF não carregada. Verifique sua conexão.');
  const {jsPDF}=window.jspdf;
  const {arr,filtros}=relatorioContexto(origem);
  const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  const total=arr.length,abertas=arr.filter(r=>r.status==='aberta').length,concluidas=arr.filter(r=>r.status==='concluida').length;
  const fechadas=arr.filter(r=>r.status==='fechada').length,vencidas=arr.filter(isVencida).length,criticas=arr.filter(r=>r.criticidade==='alta').length;
  const responsavel=AUTH_PROFILE?.nome||AUTH_PROFILE?.email||'Equipe SST';
  const cards=[['TOTAL',total],['ABERTAS',abertas],['CONCLUÍDAS',concluidas],['FECHADAS',fechadas],['VENCIDAS',vencidas],['ALTA CRITICIDADE',criticas]];
  function desenharPagina(numero){
    doc.addImage(PDF_TEMPLATE_DATA,'PNG',0,0,297,167);
    doc.setFont('helvetica','bold');doc.setTextColor(48,56,143);doc.setFontSize(16);doc.text('Relatório Executivo SST',283,36,{align:'right'});
    doc.setFont('helvetica','normal');doc.setTextColor(71,85,105);doc.setFontSize(7.5);doc.text(`Emissão: ${new Date().toLocaleString('pt-BR')}  |  Responsável: ${responsavel}`,283,42,{align:'right'});
    doc.setDrawColor(68,192,197);doc.setLineWidth(.8);doc.line(14,48,283,48);
    if(numero===1){
      doc.setFontSize(7.5);doc.setTextColor(71,85,105);doc.text(doc.splitTextToSize(`Filtros aplicados: ${filtros.join(' | ')}`,269),14,54);
      cards.forEach(([label,val],i)=>{
        const x=14+i*45;
        doc.setFillColor(i%2?239:245,i%2?250:247,i%2?250:255);doc.setDrawColor(i%2?68:48,i%2?192:56,i%2?197:143);doc.setLineWidth(.45);doc.roundedRect(x,61,40,18,2.5,2.5,'FD');
        doc.setTextColor(71,85,105);doc.setFontSize(6.3);doc.setFont('helvetica','bold');doc.text(label,x+4,67);
        doc.setTextColor(48,56,143);doc.setFontSize(14);doc.text(String(val),x+4,75);
      });
      doc.setTextColor(48,56,143);doc.setFontSize(9.5);doc.text('Lista detalhada de ocorrências',14,88);
    }
  }
  doc.autoTable({startY:92,margin:{top:56,right:14,bottom:14,left:14},head:[['Ocorrência','Empresa','Técnico','Criticidade','Status','Abertura','Prazo']],body:arr.map(r=>[r.titulo||'—',r.empresa||'—',r.tecnico||'—',{alta:'Alta',media:'Média',baixa:'Baixa'}[r.criticidade]||r.criticidade||'—',{aberta:'Aberta',concluida:'Concluída',fechada:'Fechada'}[r.status]||r.status||'—',fmt(r.data_abertura),fmt(r.prazo)]),theme:'grid',styles:{fontSize:7,cellPadding:2.2,textColor:[30,41,59],lineColor:[218,223,240],lineWidth:.2},headStyles:{fillColor:[48,56,143],textColor:255,fontStyle:'bold'},alternateRowStyles:{fillColor:[246,249,252]},columnStyles:{0:{cellWidth:62},1:{cellWidth:47},2:{cellWidth:39}},willDrawPage:data=>desenharPagina(data.pageNumber),didDrawPage:()=>{doc.setFillColor(48,56,143);doc.rect(0,204,297,6,'F');doc.setFontSize(7);doc.setTextColor(255,255,255);doc.text('EngSafe - Gestão de Saúde e Segurança do Trabalho',14,208);doc.text(`Página ${doc.internal.getNumberOfPages()}`,283,208,{align:'right'});}});
  const nome=`Relatorio_SST_${isoLocal(new Date()).replaceAll('-','')}.pdf`;
  return {doc,nome,arr,filtros};
}

function baixarRelatorioPDF(origem){
  try{const {doc,nome}=gerarRelatorioPDF(origem);doc.save(nome);toast('PDF gerado com os filtros aplicados! ✓');}
  catch(e){toast(e.message,'err');}
}

function abrirEnvioRelatorio(origem){
  if(!requireManage()) return;
  const {arr}=relatorioContexto(origem);
  if(!arr.length){toast('Não há ocorrências nos filtros atuais para enviar.','warn');return;}
  el('modal-title').textContent='Enviar relatório por e-mail';
  el('modal-body').innerHTML=`
    <div class="field"><label>E-mail(s) destinatário(s) *</label><input id="rel-email" type="text" placeholder="nome@empresa.com.br" autocomplete="email"><div style="font-size:11px;color:#64748b;margin-top:5px">Para mais de um destinatário, separe os e-mails por vírgula.</div></div>
    <div class="field"><label>Assunto</label><input id="rel-assunto" value="Relatório Executivo SST"></div>
    <div class="field"><label>Mensagem</label><textarea id="rel-mensagem" rows="4">Prezados,\n\nSegue em anexo o Relatório Executivo SST com os filtros selecionados.\n\nAtenciosamente,\nEquipe SST</textarea></div>
    <div style="padding:12px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:9px;color:#1e40af;font-size:12px;line-height:1.5">O PDF será baixado e o Outlook abrirá com destinatários, assunto e mensagem preenchidos. Anexe o arquivo baixado e confirme que o remetente é <b>${OUTLOOK_SENDER_REMINDER}</b> antes de enviar.</div>`;
  el('modal-foot').innerHTML=`<button class="btn-sec" onclick="closeModal()">Cancelar</button><button class="btn-primary" id="btn-enviar-relatorio" onclick="enviarRelatorioEmail('${origem}')">Baixar PDF e abrir Outlook</button>`;
  el('overlay').classList.add('show');setTimeout(()=>el('rel-email')?.focus(),100);
}

function emailValido(email){return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);}

function abrirOutlook(destinatarios,assunto,mensagem){
  const params=new URLSearchParams({
    to:destinatarios.join(';'),
    subject:assunto,
    body:mensagem
  });
  const link=document.createElement('a');
  link.href=`${OUTLOOK_COMPOSE_URL}?${params.toString()}`;
  link.target='_blank';
  link.rel='noopener noreferrer';
  document.body.appendChild(link);link.click();link.remove();
}

async function enviarRelatorioEmail(origem){
  if(!requireManage()) return;
  const destinatarios=[...new Set((el('rel-email')?.value||'').split(/[,;\s]+/).map(v=>v.trim().toLowerCase()).filter(Boolean))];
  if(!destinatarios.length||destinatarios.some(e=>!emailValido(e))){toast('Informe um ou mais e-mails válidos.','warn');return;}
  if(destinatarios.length>20){toast('Informe no máximo 20 destinatários.','warn');return;}
  try{
    const {doc,nome}=gerarRelatorioPDF(origem);
    doc.save(nome);
    abrirOutlook(destinatarios,(el('rel-assunto')?.value||'Relatório Executivo SST').trim(),(el('rel-mensagem')?.value||'').trim());
    closeModal();toast('PDF baixado. Anexe-o no Outlook e revise o remetente antes de enviar.','ok');
  }catch(e){toast('Erro ao preparar o e-mail: '+e.message,'err');}
}

// ── EXCLUIR ──
function confirmarExcluir(id,tipo){
  if(!requireManage()) return;
  el('modal-title').textContent='⚠ Confirmar exclusão';
  el('modal-body').innerHTML=`
    <p style="font-size:14px;margin-bottom:12px">Tem certeza que deseja excluir esta ${tipo==='ocorrencia'?'ocorrência':'atividade'}?</p>
    <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:9px;padding:12px 14px;font-size:13px;color:#dc2626">Esta ação não pode ser desfeita.</div>`;
  el('modal-foot').innerHTML=`
    <button class="btn-sec" onclick="closeModal()">Cancelar</button>
    <button class="btn-primary" style="background:#ef4444;border-color:#ef4444" onclick="executarExcluir('${id}','${tipo}')">Excluir definitivamente</button>`;
  el('overlay').classList.add('show');
}

async function executarExcluir(id,tipo){
  if(!requireManage()) return;
  const table=tipo==='ocorrencia'?'ocorrencias':'atividades_sst';
  try{
    await sbDelete(table,id);
    if(tipo==='ocorrencia') DATA.ocorrencias=DATA.ocorrencias.filter(r=>r.id!==id);
    else DATA.atividades=DATA.atividades.filter(r=>r.id!==id);
    toast('Registro excluído','ok');
    closeModal();
    renderCurrentPage();
  }catch(e){toast('Erro ao excluir: '+e.message,'err');}
}

// ── MODAL ──
function closeModal(){el('overlay').classList.remove('show');}

// ── NOTIFICAÇÃO PARA CONTRATADA ──
function abrirNotificacao(id){
  if(!requireManage()) return;
  const r=DATA.ocorrencias.find(x=>x.id===id);
  if(!r) return;
  const contratada=getContratadaOcorrencia(r);
  if(!contratada){toast('Vincule esta ocorrência a uma contratada cadastrada antes de enviar.','warn');return;}
  const autoridades=DATA.autoridades.filter(a=>a.contratada_id===contratada.id&&a.ativo!==false&&a.email);
  el('modal-title').textContent='Enviar para Contratada';
  el('modal-body').innerHTML=`
    <div class="notify-summary">
      <div><small>Empresa</small><b>${escapeHtml(contratada.nome_fantasia)}</b></div>
      <div><small>Ocorrência</small><b>${escapeHtml(r.titulo)}</b></div>
      <div><small>Prazo</small><b>${fmt(r.prazo)}</b></div>
      <div><small>Criticidade</small><b>${{alta:'Alta',media:'Média',baixa:'Baixa'}[r.criticidade]||r.criticidade}</b></div>
    </div>
    <div style="font-size:11px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-bottom:6px">Destinatários cadastrados</div>
    <div id="notify-destinatarios">${autoridades.length?autoridades.map(a=>`<label class="email-check"><input type="checkbox" name="notify-email" value="${escapeHtml(a.email)}" checked><div><b>${escapeHtml(a.nome)} · ${escapeHtml(a.cargo)}</b><span>${escapeHtml(a.email)}</span></div></label>`).join(''):'<div style="font-size:12px;color:var(--muted);padding:8px 0">Nenhuma autoridade com e-mail cadastrada.</div>'}</div>
    <div class="field" style="margin-top:14px"><label>Outros e-mails</label><input id="notify-email-extra" type="text" placeholder="nome@empresa.com.br; outro@empresa.com.br"><div style="font-size:11px;color:#64748b;margin-top:5px">Opcional. Separe vários e-mails por vírgula ou ponto e vírgula.</div></div>
    <div style="margin-top:14px;padding:12px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:9px;color:#1e40af;font-size:11.5px;line-height:1.5">O Outlook será aberto com o e-mail preenchido. Confirme que o remetente é <b>${OUTLOOK_SENDER_REMINDER}</b>, revise e clique em <b>Enviar</b>. O EngSafe não envia automaticamente.</div>`;
  el('modal-foot').innerHTML=`<button class="btn-sec" onclick="closeModal()">Cancelar</button><button class="btn-primary" id="btn-enviar-notificacao" onclick="enviarNotificacao('${r.id}')">Abrir no Outlook</button>`;
  el('overlay').classList.add('show');
}

async function enviarNotificacao(id){
  if(!requireManage()) return;
  const r=DATA.ocorrencias.find(x=>x.id===id);if(!r)return;
  const contratada=getContratadaOcorrencia(r);if(!contratada)return;
  const cadastrados=[...document.querySelectorAll('input[name="notify-email"]:checked')].map(i=>i.value.trim().toLowerCase());
  const extras=(el('notify-email-extra')?.value||'').split(/[,;\s]+/).map(v=>v.trim().toLowerCase()).filter(Boolean);
  const destinatarios=[...new Set([...cadastrados,...extras])];
  if(!destinatarios.length){toast('Selecione pelo menos um destinatário.','warn');return;}
  if(destinatarios.some(e=>!emailValido(e))){toast('Revise os endereços de e-mail informados.','warn');return;}
  if(destinatarios.length>20){toast('Selecione no máximo 20 destinatários.','warn');return;}
  const assunto='Notificação de Ocorrência SST Pendente';
  const criticidade={alta:'Alta',media:'Média',baixa:'Baixa'}[r.criticidade]||r.criticidade||'—';
  const mensagem=`Prezados,\n\nFoi identificada uma ocorrência SST em aberto junto à contratada:\n\nEmpresa: ${contratada.nome_fantasia||contratada.razao_social||r.empresa||'—'}\n\nOcorrência:\n${r.titulo||'—'}\n\nDescrição:\n${r.descricao||'—'}\n\nCriticidade:\n${criticidade}\n\nPrazo:\n${fmt(r.prazo)}\n\nSolicitamos tratativa imediata.\n\nAtenciosamente,\n\nEquipe SST`;
  abrirOutlook(destinatarios,assunto,mensagem);
  PENDING_EMAIL_LOG={id,contratada_id:contratada.id,destinatarios,assunto};
  el('modal-foot').innerHTML=`<button class="btn-sec" onclick="closeModal()">Fechar</button><button class="btn-primary" onclick="registrarCobrancaManual()">Já enviei — registrar cobrança</button>`;
  toast(`Outlook aberto. Confirme o remetente ${OUTLOOK_SENDER_REMINDER}.`,'ok');
}

async function registrarCobrancaManual(){
  if(!requireManage())return;
  try{
    const p=PENDING_EMAIL_LOG;if(!p)throw new Error('Nenhuma cobrança aguardando confirmação.');
    const [registro]=await sbInsert('historico_emails',{ocorrencia_id:p.id,contratada_id:p.contratada_id,tipo_envio:'notificacao',destinatarios:p.destinatarios,assunto:p.assunto,status_envio:'enviado_manual',enviado_por:String(AUTH_PROFILE?.email||AUTH_USER?.email||OUTLOOK_SENDER_REMINDER).toLowerCase(),data_envio:new Date().toISOString()});
    if(registro)DATA.historicoEmails.unshift(registro);
    PENDING_EMAIL_LOG=null;closeModal();toast('Cobrança manual registrada no histórico.','ok');renderCurrentPage();
  }catch(e){toast('O Outlook foi aberto, mas o histórico não pôde ser registrado: '+e.message,'err');}
}

// ── UPLOAD & DRAG-DROP ──
function fmtSz(b){return b>1048576?(b/1048576).toFixed(1)+' MB':(b/1024).toFixed(0)+' KB';}
function dzMerge(srcInput, targetId, gridId){
  const target = el(targetId);
  if(!target) return;
  if(srcInput === target){
    // direct input change — just preview
    dzPreview(targetId, gridId);
    return;
  }
  if(!srcInput.files.length) return;
  const existing = Array.from(target.files || []);
  const added    = Array.from(srcInput.files);
  const dt = new DataTransfer();
  [...existing, ...added].forEach(f => dt.items.add(f));
  target.files = dt.files;
  dzPreview(targetId, gridId);
  srcInput.value = '';
}

function dzOver(e,id){e.preventDefault();el(id)?.classList.add('drag');}
function dzLeave(id){el(id)?.classList.remove('drag');}
function dzDrop(e,inputId,zoneId,gridId){
  e.preventDefault();dzLeave(zoneId);
  const inp=el(inputId);if(!inp) return;
  const existing=Array.from(inp.files||[]);
  const added=Array.from(e.dataTransfer.files);
  const dt=new DataTransfer();
  [...existing,...added].forEach(f=>dt.items.add(f));
  inp.files=dt.files;
  dzPreview(inputId,gridId);
}
function dzPreview(inputId,gridId){
  const inp=el(inputId);const grid=el(gridId);
  if(!inp||!grid) return;
  const files=Array.from(inp.files);
  // Update count badge
  const countEl=el(inputId+'-count');
  if(countEl){countEl.textContent=files.length?files.length+' arquivo'+(files.length!==1?'s':''):'';countEl.style.display=files.length?'inline':'none';}
  if(!files.length){grid.innerHTML='';return;}
  grid.innerHTML=`
    <div style="grid-column:1/-1;display:flex;align-items:center;justify-content:space-between;margin-bottom:2px">
      <span style="font-size:12px;font-weight:600;color:rgba(255,255,255,.6)">${files.length} arquivo${files.length!==1?'s':''} selecionado${files.length!==1?'s':''}</span>
      <button onclick="rmAllFiles('${inputId}','${gridId}')" style="font-size:11px;color:rgba(255,255,255,.4);background:none;border:none;cursor:pointer;padding:2px 6px;border-radius:4px">Limpar tudo</button>
    </div>
    ${files.map((f,i)=>{
      const isImg=f.type.startsWith('image/');
      const isPdf=f.type==='application/pdf';
      const preview=isImg
        ?`<img src="${URL.createObjectURL(f)}" alt="${f.name}" style="width:100%;height:100%;object-fit:cover">`
        :`<div class="attach-pdf"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="${isPdf?'#ef4444':'#94a3b8'}" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg><div class="attach-name" title="${f.name}">${f.name.slice(0,18)}${f.name.length>18?'…':''}</div><div style="font-size:10px;color:#94a3b8;margin-top:2px">${fmtSz(f.size)}</div></div>`;
      return `<div class="attach-thumb" id="att-${inputId}-${i}">
        ${preview}
        <button class="rm-btn" onclick="rmFile('${inputId}','${gridId}',${i})" title="Remover">✕</button>
      </div>`;
    }).join('')}`;
}

function rmAllFiles(inputId,gridId){
  const inp=el(inputId);if(!inp)return;
  inp.value='';
  dzPreview(inputId,gridId);
}
function rmFile(inputId,gridId,idx){
  const inp=el(inputId);if(!inp) return;
  const dt=new DataTransfer();
  Array.from(inp.files).forEach((f,i)=>{if(i!==idx)dt.items.add(f);});
  inp.files=dt.files;
  dzPreview(inputId,gridId);
}

async function uploadFiles(files,bucket,recordId,progId,fillId){
  if(!requireManage()) return;
  const prog=el(progId),fill=el(fillId);
  if(prog) prog.style.display='block';
  let done=0;
  for(const file of files){
    try{
      const safe=file.name.replace(/[^a-zA-Z0-9._-]/g,'_');
      const path=`${recordId}/${Date.now()}_${safe}`;
      const res=await fetch(`${SB_URL}/storage/v1/object/${bucket}/${path}`,{
        method:'POST',
        headers:{...authHeaders(null),'Content-Type':file.type,'x-upsert':'true'},
        body:file
      });
      if(!res.ok){const t=await res.text();console.warn('Upload fail:',t);}
    }catch(e){console.warn('Upload err:',e);}
    done++;
    if(fill) fill.style.width=Math.round(done/files.length*100)+'%';
  }
  if(prog) setTimeout(()=>{prog.style.display='none';if(fill)fill.style.width='0%';},800);
}

// ── CONTRATADAS ──
function renderContratadas(){
  const body=el('contratada-tbody');
  if(!body) return;
  const q=norm(el('contratadaSearch')?.value||'');
  const st=el('contratadaStatus')?.value||'';
  const rows=DATA.contratadas.filter(c=>{
    if(q&&!norm(`${c.razao_social||''} ${c.nome_fantasia||''} ${c.cnpj||''} ${c.responsavel||''}`).includes(q)) return false;
    if(st&&String(c.ativo)!==st) return false;
    return true;
  });
  body.innerHTML=rows.length?rows.map(c=>{
    const autoridades=DATA.autoridades.filter(a=>a.contratada_id===c.id&&a.ativo!==false).length;
    return `<tr>
      <td><b>${escapeHtml(c.nome_fantasia)}</b><div style="font-size:11px;color:rgba(255,255,255,.5);margin-top:3px">${escapeHtml(c.razao_social)}</div></td>
      <td>${escapeHtml(c.cnpj||'—')}</td>
      <td>${escapeHtml(c.responsavel||'—')}<div style="font-size:11px;color:rgba(255,255,255,.5);margin-top:3px">${escapeHtml(c.cargo||'')}</div></td>
      <td>${escapeHtml(c.email||'—')}<div style="font-size:11px;color:rgba(255,255,255,.5);margin-top:3px">${escapeHtml(c.telefone||'')}</div></td>
      <td>${autoridades}</td>
      <td><span class="status-dot ${c.ativo===false?'off':''}"></span>${c.ativo===false?'Inativa':'Ativa'}</td>
      <td style="white-space:nowrap">${canManage()?`<button class="btn-sec" style="padding:5px 10px" onclick="editarContratada('${c.id}')">Editar</button> <button class="btn-danger" style="padding:5px 9px" onclick="confirmarExcluirContratada('${c.id}')">🗑</button>`:'—'}</td>
    </tr>`;
  }).join(''):'<tr><td colspan="7" style="text-align:center;padding:40px;color:rgba(255,255,255,.5)">Nenhuma contratada cadastrada.</td></tr>';
  el('contratada-count').textContent=`${rows.length} contratada${rows.length!==1?'s':''} encontrada${rows.length!==1?'s':''}`;
}

function novaContratada(){if(requireManage())abrirEditorContratada(null);}
function editarContratada(id){if(requireManage())abrirEditorContratada(DATA.contratadas.find(c=>c.id===id));}

function abrirEditorContratada(c){
  const autoridades=c?DATA.autoridades.filter(a=>a.contratada_id===c.id):[];
  el('modal-title').textContent=c?'Editar Contratada':'Nova Contratada';
  el('modal-body').innerHTML=`
    <input type="hidden" id="c_id" value="${escapeHtml(c?.id||'')}">
    <div class="grid2f">
      <div class="field"><label>Razão Social *</label><input id="c_razao" value="${escapeHtml(c?.razao_social||'')}"></div>
      <div class="field"><label>Nome Fantasia *</label><input id="c_fantasia" value="${escapeHtml(c?.nome_fantasia||'')}"></div>
    </div>
    <div class="grid2f">
      <div class="field"><label>CNPJ *</label><input id="c_cnpj" maxlength="18" placeholder="00.000.000/0000-00" value="${escapeHtml(c?.cnpj||'')}"></div>
      <div class="field"><label>Status</label><select id="c_ativo"><option value="true" ${c?.ativo!==false?'selected':''}>Ativa</option><option value="false" ${c?.ativo===false?'selected':''}>Inativa</option></select></div>
    </div>
    <div class="grid2f">
      <div class="field"><label>Responsável</label><input id="c_responsavel" value="${escapeHtml(c?.responsavel||'')}"></div>
      <div class="field"><label>Cargo</label><input id="c_cargo" value="${escapeHtml(c?.cargo||'')}"></div>
    </div>
    <div class="grid2f">
      <div class="field"><label>E-mail</label><input id="c_email" type="email" value="${escapeHtml(c?.email||'')}"></div>
      <div class="field"><label>Telefone</label><input id="c_telefone" value="${escapeHtml(c?.telefone||'')}"></div>
    </div>
    <div class="autoridades-box">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:10px"><div><b style="font-size:13px;color:var(--navy)">Autoridades da Contratada</b><div style="font-size:11.5px;color:var(--muted);margin-top:3px">Diretor, gerente, coordenador, engenheiro ou responsável SST.</div></div><button class="btn-sec" onclick="addAutoridadeEditor()">+ Autoridade</button></div>
      <div id="autoridades-editor">${autoridades.map(a=>autoridadeEditorHtml(a)).join('')}</div>
    </div>`;
  el('modal-foot').innerHTML=`<button class="btn-sec" onclick="closeModal()">Cancelar</button><button class="btn-primary" id="btn-salvar-contratada" onclick="salvarContratada()">Salvar Contratada</button>`;
  el('overlay').classList.add('show');
}

function autoridadeEditorHtml(a={}){
  return `<div class="autoridade-item" data-id="${escapeHtml(a.id||'')}">
    <div class="field"><label>Nome *</label><input data-a="nome" value="${escapeHtml(a.nome||'')}"></div>
    <div class="field"><label>Cargo *</label><select data-a="cargo">${['Diretor','Gerente','Coordenador','Engenheiro','Responsável SST'].map(x=>`<option ${a.cargo===x?'selected':''}>${x}</option>`).join('')}</select></div>
    <div class="field"><label>E-mail *</label><input data-a="email" type="email" value="${escapeHtml(a.email||'')}"></div>
    <div class="field"><label>Telefone</label><input data-a="telefone" value="${escapeHtml(a.telefone||'')}"></div>
    <button class="btn-danger" onclick="this.closest('.autoridade-item').remove()">Remover</button>
  </div>`;
}

function addAutoridadeEditor(){el('autoridades-editor').insertAdjacentHTML('beforeend',autoridadeEditorHtml());}

async function salvarContratada(){
  if(!requireManage()) return;
  const id=el('c_id').value;
  const row={razao_social:el('c_razao').value.trim(),nome_fantasia:el('c_fantasia').value.trim(),cnpj:el('c_cnpj').value.trim(),responsavel:el('c_responsavel').value.trim()||null,cargo:el('c_cargo').value.trim()||null,email:el('c_email').value.trim()||null,telefone:el('c_telefone').value.trim()||null,ativo:el('c_ativo').value==='true'};
  if(!row.razao_social||!row.nome_fantasia||!row.cnpj){toast('Preencha Razão Social, Nome Fantasia e CNPJ','warn');return;}
  const autoridades=[...document.querySelectorAll('#autoridades-editor .autoridade-item')].map(d=>({nome:d.querySelector('[data-a="nome"]').value.trim(),cargo:d.querySelector('[data-a="cargo"]').value,email:d.querySelector('[data-a="email"]').value.trim(),telefone:d.querySelector('[data-a="telefone"]').value.trim()||null,ativo:true})).filter(a=>a.nome||a.email);
  if(autoridades.some(a=>!a.nome||!a.email)){toast('Preencha nome e e-mail de todas as autoridades','warn');return;}
  const btn=el('btn-salvar-contratada');btn.disabled=true;btn.textContent='Salvando...';
  try{
    let contratadaId=id;
    if(id){await sbUpdate('contratadas',id,row);Object.assign(DATA.contratadas.find(c=>c.id===id),row);}
    else{const saved=await sbInsert('contratadas',row);const rec=Array.isArray(saved)?saved[0]:saved;contratadaId=rec.id;DATA.contratadas.push({...row,...rec});}
    const antigas=DATA.autoridades.filter(a=>a.contratada_id===contratadaId);
    for(const a of antigas) await sbDelete('autoridades_contratada',a.id);
    DATA.autoridades=DATA.autoridades.filter(a=>a.contratada_id!==contratadaId);
    for(const a of autoridades){const saved=await sbInsert('autoridades_contratada',{...a,contratada_id:contratadaId});const rec=Array.isArray(saved)?saved[0]:saved;DATA.autoridades.push({...a,...rec,contratada_id:contratadaId});}
    populateFilters();closeModal();renderContratadas();toast('Contratada salva com sucesso! ✓');
  }catch(e){toast('Erro ao salvar contratada: '+e.message,'err');btn.disabled=false;btn.textContent='Salvar Contratada';}
}

function confirmarExcluirContratada(id){
  if(!requireManage()) return;
  const vinculada=DATA.ocorrencias.some(o=>o.contratada_id===id);
  if(vinculada){toast('A contratada possui ocorrências vinculadas. Inative o cadastro em vez de excluir.','warn');return;}
  el('modal-title').textContent='⚠ Excluir Contratada';
  el('modal-body').innerHTML='<p style="font-size:14px">Deseja excluir esta contratada e todas as autoridades cadastradas?</p>';
  el('modal-foot').innerHTML=`<button class="btn-sec" onclick="closeModal()">Cancelar</button><button class="btn-primary" style="background:#ef4444" onclick="excluirContratada('${id}')">Excluir definitivamente</button>`;
  el('overlay').classList.add('show');
}

async function excluirContratada(id){
  if(!requireManage()) return;
  try{await sbDelete('contratadas',id);DATA.contratadas=DATA.contratadas.filter(c=>c.id!==id);DATA.autoridades=DATA.autoridades.filter(a=>a.contratada_id!==id);closeModal();populateFilters();renderContratadas();toast('Contratada excluída','ok');}catch(e){toast('Erro ao excluir: '+e.message,'err');}
}

// ── EMPRESAS ──
function renderEmpresas(){
  const byEmp={};
  DATA.ocorrencias.forEach(r=>{const k=r.empresa||'Não informada';if(!byEmp[k])byEmp[k]=[];byEmp[k].push(r);});
  const sorted=Object.entries(byEmp).sort((a,b)=>b[1].length-a[1].length);
  const COLORS_EMP=['#1a4fa0','#00a89b','#1e56c4','#0d3d7a','#2ec4b6','#154090'];
  el('emp-cards').innerHTML=sorted.map(([emp,rows],i)=>{
    const alta=rows.filter(r=>r.criticidade==='alta').length;
    const abertas=rows.filter(r=>r.status==='aberta').length;
    const concl=rows.filter(r=>r.status==='concluida'||r.status==='fechada').length;
    const taxa=rows.length?Math.round(concl/rows.length*100):0;
    const c=COLORS_EMP[i%COLORS_EMP.length];
    return `<div onclick="openEmpDetail('${emp.replace(/'/g,"\'")}',${i})" style="background:#1e3a5f;border:1px solid rgba(255,255,255,.1);border-left:4px solid ${c};border-radius:12px;padding:20px;cursor:pointer;transition:transform .15s,box-shadow .15s" onmouseover="this.style.transform='translateY(-2px)';this.style.boxShadow='0 6px 20px rgba(13,33,55,.3)'" onmouseout="this.style.transform='';this.style.boxShadow=''">
      <div style="font-size:15px;font-weight:700;color:#fff;margin-bottom:12px">${emp}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:14px">
        <div style="text-align:center;padding:8px;background:rgba(255,255,255,.06);border-radius:8px">
          <div style="font-size:20px;font-weight:800;color:#fff">${rows.length}</div>
          <div style="font-size:10px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.05em">Total</div>
        </div>
        <div style="text-align:center;padding:8px;background:rgba(255,255,255,.06);border-radius:8px">
          <div style="font-size:20px;font-weight:800;color:${alta>0?'#fca5a5':'#fff'}">${alta}</div>
          <div style="font-size:10px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.05em">Críticas</div>
        </div>
        <div style="text-align:center;padding:8px;background:rgba(255,255,255,.06);border-radius:8px">
          <div style="font-size:20px;font-weight:800;color:${abertas>0?'#fcd34d':'#5eead4'}">${abertas}</div>
          <div style="font-size:10px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:.05em">Abertas</div>
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:center;font-size:12px;color:rgba(255,255,255,.6);margin-bottom:6px">
        <span>Taxa de resolução</span><span style="font-weight:700;color:#fff">${taxa}%</span>
      </div>
      <div style="height:4px;background:rgba(255,255,255,.12);border-radius:3px;overflow:hidden">
        <div style="height:100%;width:${taxa}%;background:${c};border-radius:3px"></div>
      </div>
    </div>`;
  }).join('');
  el('emp-detail').style.display='none';
  el('emp-view-cards').style.display='block';
}

function openEmpDetail(emp){
  const rows=DATA.ocorrencias.filter(r=>r.empresa===emp);
  el('emp-detail-name').textContent=emp+' — '+rows.length+' ocorrência'+(rows.length!==1?'s':'');
  el('emp-tbody').innerHTML=rows.length
    ?rows.map(r=>`<tr style="cursor:pointer" onclick="verOcorrencia('${r.id}')">
        <td><b>${r.titulo.slice(0,45)}${r.titulo.length>45?'…':''}</b></td>
        <td>${r.tecnico||'—'}</td>
        <td>${r.encarregado||'—'}</td>
        <td>${badge('crit',r.criticidade)}</td>
        <td>${badge('st',r.status)}</td>
        <td style="${isVencida(r)?'color:#fca5a5;font-weight:700':''}">${fmt(r.prazo)}</td>
        <td><button class="btn-sec" style="padding:4px 10px;font-size:12px" onclick="event.stopPropagation();verOcorrencia('${r.id}')">Ver</button></td>
      </tr>`).join('')
    :'<tr><td colspan="7" style="text-align:center;padding:30px;color:rgba(255,255,255,.5)">Nenhuma ocorrência.</td></tr>';
  el('emp-view-cards').style.display='none';
  el('emp-detail').style.display='block';
}

function closeEmpDetail(){
  el('emp-detail').style.display='none';
  el('emp-view-cards').style.display='block';
}

// ── PDF DE OCORRÊNCIA ──
function imprimirOcorrencia(id){
  const r=DATA.ocorrencias.find(x=>x.id===id);
  if(!r) return;
  const crit={alta:'Alta',media:'Média',baixa:'Baixa'};
  const st={aberta:'Aberta',concluida:'Concluída',fechada:'Fechada'};
  const win=window.open('','_blank','width=800,height=900');
  win.document.write(`<!DOCTYPE html><html lang="pt-BR"><head>
  <meta charset="UTF-8"><title>Ocorrência SST — ${r.titulo}</title>
  <style>
    *{box-sizing:border-box;margin:0;padding:0}
    body{font-family:'Segoe UI',system-ui,sans-serif;color:#0f172a;background:#fff;padding:40px}
    .header{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:18px;border-bottom:3px solid #1a4fa0;margin-bottom:24px}
    .header-left h1{font-size:13px;font-weight:700;color:#1a4fa0;letter-spacing:.1em;text-transform:uppercase;margin-bottom:4px}
    .header-left h2{font-size:20px;font-weight:800;color:#0d2137}
    .header-right{text-align:right;font-size:11px;color:#64748b}
    .badges{display:flex;gap:8px;margin-bottom:20px}
    .badge{padding:4px 12px;border-radius:20px;font-size:12px;font-weight:700}
    .b-alta{background:#fecaca;color:#7f1d1d}
    .b-media{background:#fde68a;color:#78350f}
    .b-baixa{background:#bbf7d0;color:#14532d}
    .b-aberta{background:#bfdbfe;color:#1e3a8a}
    .b-concluida{background:#99f6e4;color:#134e4a}
    .b-fechada{background:#e2e8f0;color:#1e293b}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:0;margin-bottom:20px;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden}
    .row{display:contents}
    .key{background:#f8fafc;padding:10px 14px;font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.06em;border-bottom:1px solid #e2e8f0}
    .val{padding:10px 14px;font-size:13px;color:#0f172a;border-bottom:1px solid #e2e8f0}
    .section{margin-bottom:18px}
    .section-title{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}
    .box{background:#f8fafc;border-radius:8px;padding:14px;font-size:13.5px;line-height:1.7;color:#0f172a;white-space:pre-wrap}
    .box-green{background:#f0fdf4}
    .assin{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:50px;padding-top:24px;border-top:1px solid #e2e8f0}
    .assin-field{text-align:center}
    .assin-line{border-top:1px solid #0d2137;margin-bottom:8px;margin-top:50px}
    .assin-label{font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.06em}
    .assin-name{font-size:13px;color:#0d2137;font-weight:700;margin-top:3px}
    .footer{margin-top:30px;padding-top:14px;border-top:1px solid #e2e8f0;font-size:10px;color:#94a3b8;display:flex;justify-content:space-between}
    @media print{body{padding:28px}button{display:none}}
  </style>
  </head><body>
  <div class="header">
    <div class="header-left">
      <h1>Registro de Ocorrência SST — EngSafe</h1>
      <h2>${r.titulo}</h2>
    </div>
    <div class="header-right">
      <div>ID: ${r.id?.slice(0,8)||'—'}</div>
      <div>Emitido em: ${new Date().toLocaleDateString('pt-BR')}</div>
    </div>
  </div>

  <div class="badges">
    <span class="badge b-${r.criticidade}">${crit[r.criticidade]||r.criticidade}</span>
    <span class="badge b-${r.status}">${st[r.status]||r.status}</span>
    ${isVencida(r)?'<span class="badge" style="background:#fecaca;color:#7f1d1d">⏰ Vencida</span>':''}
  </div>

  <div class="grid">
    <div class="key">Empresa</div><div class="val">${r.empresa||'—'}</div>
    <div class="key">Frente de Serviço</div><div class="val">${r.frente_servico||'—'}</div>
    <div class="key">Técnico SST</div><div class="val">${r.tecnico||'—'}</div>
    <div class="key">Encarregado</div><div class="val">${r.encarregado||'—'}</div>
    <div class="key">Fiscal de Campo</div><div class="val">${r.fiscal||'—'}</div>
    <div class="key">Data de Abertura</div><div class="val">${fmt(r.data_abertura)}</div>
    <div class="key">Prazo</div><div class="val" style="${isVencida(r)?'color:#dc2626;font-weight:700':''}">${fmt(r.prazo)}</div>
    <div class="key">Status</div><div class="val">${st[r.status]||r.status}</div>
  </div>

  ${r.descricao?`<div class="section"><div class="section-title">Descrição da Ocorrência</div><div class="box">${r.descricao}</div></div>`:''}
  ${r.acao_corretiva?`<div class="section"><div class="section-title">Ação Corretiva</div><div class="box box-green">${r.acao_corretiva}</div></div>`:''}

  <div class="assin">
    <div class="assin-field">
      <div class="assin-line"></div>
      <div class="assin-label">Técnico de Segurança do Trabalho</div>
      <div class="assin-name">${r.tecnico||'_______________________________'}</div>
    </div>
    <div class="assin-field">
      <div class="assin-line"></div>
      <div class="assin-label">Encarregado da Obra</div>
      <div class="assin-name">${r.encarregado||'_______________________________'}</div>
    </div>
  </div>

  <div class="footer">
    <span>EngSafe — Sistema de Gestão SST</span>
    <span>Gerado em ${new Date().toLocaleString('pt-BR')}</span>
  </div>

  <script>window.onload=()=>window.print()<\/script>

<!-- CÂMERA MODAL -->
<div class="cam-overlay" id="camOverlay">
  <div style="width:100%;max-width:480px;padding:0 16px">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <span style="color:#fff;font-size:14px;font-weight:600">Câmera</span>
      <span id="camCount" style="color:#00c6b8;font-size:13px"></span>
    </div>
    <video id="camVideo" class="cam-video" autoplay playsinline muted></video>
    <div class="cam-bar">
      <button class="cam-switch" id="camSwitch" onclick="camToggleFacing()" title="Alternar câmera">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 4v6h6"/><path d="M23 20v-6h-6"/><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"/></svg>
      </button>
      <button class="cam-shoot" onclick="camCapture()" title="Tirar foto">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#0d2137" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>
      </button>
      <button class="cam-close" onclick="camClose()" title="Fechar">✕</button>
    </div>
    <canvas id="camCanvas" style="display:none"></canvas>
    <div class="cam-preview-strip" id="camStrip"></div>
    <button onclick="camConfirm()" style="margin-top:16px;width:100%;padding:12px;border-radius:10px;background:linear-gradient(135deg,#1a4fa0,#00c6b8);color:#fff;font-size:14px;font-weight:700;border:none;cursor:pointer;display:none" id="camConfirmBtn">
      ✓ Usar fotos (<span id="camConfirmCount">0</span>)
    </button>
  </div>
</div>

</body></html>`);
  win.document.close();
}

// ── IMPRIMIR OCORRÊNCIA (PDF) ──
function imprimirOcorrencia(id){
  const r=DATA.ocorrencias.find(x=>x.id===id);
  if(!r) return;
  const critica={'alta':'Alta','media':'Média','baixa':'Baixa'}[r.criticidade]||r.criticidade;
  const statusL={'aberta':'Aberta','concluida':'Concluída','fechada':'Fechada'}[r.status]||r.status;
  const w=window.open('','_blank','width=794,height=1123');
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
<title>Ocorrência SST — ${r.titulo}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',system-ui,sans-serif;color:#0f172a;background:#fff;padding:40px 48px;font-size:13px}
  .header{display:flex;justify-content:space-between;align-items:flex-start;border-bottom:3px solid #0d2137;padding-bottom:16px;margin-bottom:24px}
  .logo-area h1{font-size:22px;font-weight:800;color:#0d2137;letter-spacing:-.3px}
  .logo-area p{font-size:11px;color:#64748b;margin-top:2px;letter-spacing:.05em;text-transform:uppercase}
  .doc-info{text-align:right;font-size:11px;color:#64748b}
  .doc-info strong{display:block;font-size:13px;color:#0d2137;font-weight:700}
  .title-block{background:linear-gradient(135deg,#0d2137,#1a4fa0);border-radius:10px;padding:18px 22px;margin-bottom:22px;color:#fff}
  .title-block h2{font-size:18px;font-weight:700;margin-bottom:8px}
  .title-block .tags{display:flex;gap:8px;flex-wrap:wrap}
  .tag{padding:3px 12px;border-radius:20px;font-size:11px;font-weight:700;background:rgba(255,255,255,.18);color:#fff}
  .section{margin-bottom:20px}
  .section-title{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#64748b;border-bottom:1px solid #e2e8f0;padding-bottom:6px;margin-bottom:12px}
  .info-grid{display:grid;grid-template-columns:1fr 1fr;gap:0}
  .info-row{padding:8px 0;border-bottom:1px solid #f1f5f9;display:flex;gap:12px}
  .info-row:last-child{border-bottom:none}
  .info-key{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.05em;min-width:130px;flex-shrink:0}
  .info-val{color:#0f172a;font-size:13px}
  .text-block{background:#f8fafc;border-radius:8px;padding:14px;font-size:13px;line-height:1.7;white-space:pre-wrap;color:#0f172a;border-left:3px solid #1a4fa0}
  .text-block.green{border-left-color:#00a89b;background:#f0fdf9}
  .sign-section{margin-top:40px;border-top:2px solid #e2e8f0;padding-top:24px}
  .sign-grid{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:16px}
  .sign-box{text-align:center}
  .sign-line{border-top:1.5px solid #0d2137;margin:60px 16px 8px}
  .sign-label{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.07em}
  .sign-name{font-size:12px;color:#0f172a;margin-top:3px;font-weight:600}
  .footer{margin-top:30px;text-align:center;font-size:10px;color:#94a3b8;border-top:1px solid #e2e8f0;padding-top:12px}
  @media print{body{padding:20px 28px}}
</style></head><body>
<div class="header">
  <div class="logo-area">
    <h1>EngSafe</h1>
    <p>Sistema de Gestão SST</p>
  </div>
  <div class="doc-info">
    <strong>Registro de Ocorrência</strong>
    Gerado em ${new Date().toLocaleDateString('pt-BR',{day:'2-digit',month:'long',year:'numeric'})}<br>
    ID: ${r.id?.slice(0,8)||'—'}
  </div>
</div>

<div class="title-block">
  <h2>${r.titulo}</h2>
  <div class="tags">
    <span class="tag">Criticidade: ${critica}</span>
    <span class="tag">Status: ${statusL}</span>
    ${r.data_abertura?`<span class="tag">Abertura: ${new Date(r.data_abertura+'T12:00').toLocaleDateString('pt-BR')}</span>`:''}
    ${r.prazo?`<span class="tag">Prazo: ${new Date(r.prazo+'T12:00').toLocaleDateString('pt-BR')}</span>`:''}
  </div>
</div>

<div class="section">
  <div class="section-title">Informações da ocorrência</div>
  <div class="info-row"><div class="info-key">Empresa</div><div class="info-val">${r.empresa||'—'}</div></div>
  <div class="info-row"><div class="info-key">Frente de Serviço</div><div class="info-val">${r.frente_servico||'—'}</div></div>
  <div class="info-row"><div class="info-key">Técnico de Segurança</div><div class="info-val">${r.tecnico||'—'}</div></div>
  <div class="info-row"><div class="info-key">Encarregado da Obra</div><div class="info-val">${r.encarregado||'—'}</div></div>
  <div class="info-row"><div class="info-key">Fiscal de Campo</div><div class="info-val">${r.fiscal||'—'}</div></div>
</div>

${r.descricao?`<div class="section">
  <div class="section-title">Descrição da ocorrência</div>
  <div class="text-block">${r.descricao}</div>
</div>`:''}

${r.acao_corretiva?`<div class="section">
  <div class="section-title">Ação corretiva</div>
  <div class="text-block green">${r.acao_corretiva}</div>
</div>`:''}

<div class="sign-section">
  <div class="section-title">Assinaturas</div>
  <div class="sign-grid">
    <div class="sign-box">
      <div class="sign-line"></div>
      <div class="sign-label">Técnico de Segurança</div>
      <div class="sign-name">${r.tecnico||'_____________________'}</div>
    </div>
    <div class="sign-box">
      <div class="sign-line"></div>
      <div class="sign-label">Encarregado da Obra</div>
      <div class="sign-name">${r.encarregado||'_____________________'}</div>
    </div>
  </div>
</div>

<div class="footer">EngSafe — Gestão SST · Documento gerado automaticamente · ${new Date().toLocaleString('pt-BR')}</div>

<!-- CÂMERA MODAL -->
<div class="cam-overlay" id="camOverlay">
  <div style="width:100%;max-width:480px;padding:0 16px">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <span style="color:#fff;font-size:14px;font-weight:600">Câmera</span>
      <span id="camCount" style="color:#00c6b8;font-size:13px"></span>
    </div>
    <video id="camVideo" class="cam-video" autoplay playsinline muted></video>
    <div class="cam-bar">
      <button class="cam-switch" id="camSwitch" onclick="camToggleFacing()" title="Alternar câmera">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 4v6h6"/><path d="M23 20v-6h-6"/><path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4l-4.64 4.36A9 9 0 0 1 3.51 15"/></svg>
      </button>
      <button class="cam-shoot" onclick="camCapture()" title="Tirar foto">
        <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#0d2137" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>
      </button>
      <button class="cam-close" onclick="camClose()" title="Fechar">✕</button>
    </div>
    <canvas id="camCanvas" style="display:none"></canvas>
    <div class="cam-preview-strip" id="camStrip"></div>
    <button onclick="camConfirm()" style="margin-top:16px;width:100%;padding:12px;border-radius:10px;background:linear-gradient(135deg,#1a4fa0,#00c6b8);color:#fff;font-size:14px;font-weight:700;border:none;cursor:pointer;display:none" id="camConfirmBtn">
      ✓ Usar fotos (<span id="camConfirmCount">0</span>)
    </button>
  </div>
</div>

</body></html>`);
  w.document.close();
  w.focus();
  setTimeout(()=>w.print(),400);
}

// ── CÂMERA getUserMedia ──
let _camStream=null, _camFacing='environment', _camTarget=null, _camGrid=null, _camBlobs=[];

async function camOpen(targetId, gridId){
  _camTarget=targetId; _camGrid=gridId; _camBlobs=[];
  el('camStrip').innerHTML=''; el('camConfirmCount').textContent='0';
  el('camConfirmBtn').style.display='none';
  el('camOverlay').classList.add('show');
  await camStart();
}

async function camStart(){
  if(_camStream){_camStream.getTracks().forEach(t=>t.stop());_camStream=null;}
  try{
    _camStream=await navigator.mediaDevices.getUserMedia({
      video:{facingMode:_camFacing,width:{ideal:1920},height:{ideal:1080}},audio:false
    });
    el('camVideo').srcObject=_camStream;
  }catch(e){
    // fallback: sem facing mode específico
    try{
      _camStream=await navigator.mediaDevices.getUserMedia({video:true,audio:false});
      el('camVideo').srcObject=_camStream;
    }catch(e2){
      alert('Câmera não disponível: '+e2.message);
      camClose();
    }
  }
}

function camToggleFacing(){
  _camFacing=(_camFacing==='environment'?'user':'environment');
  camStart();
}

function camCapture(){
  const v=el('camVideo'), c=el('camCanvas');
  c.width=v.videoWidth; c.height=v.videoHeight;
  c.getContext('2d').drawImage(v,0,0);
  c.toBlob(blob=>{
    if(!blob) return;
    _camBlobs.push(blob);
    // preview thumbnail
    const img=document.createElement('img');
    img.className='cam-thumb'; img.src=URL.createObjectURL(blob);
    el('camStrip').appendChild(img);
    el('camConfirmCount').textContent=_camBlobs.length;
    el('camCount').textContent=_camBlobs.length+' foto'+(_camBlobs.length!==1?'s':'');
    el('camConfirmBtn').style.display='block';
    // flash effect
    const f=document.createElement('div');
    f.style.cssText='position:fixed;inset:0;background:#fff;opacity:.6;z-index:400;pointer-events:none';
    document.body.appendChild(f);
    setTimeout(()=>f.remove(),120);
  },'image/jpeg',0.92);
}

function camConfirm(){
  if(!_camBlobs.length) return;
  const input=el(_camTarget);
  if(!input){camClose();return;}
  const existing=Array.from(input.files||[]);
  const newFiles=_camBlobs.map((b,i)=>new File([b],`foto_${Date.now()}_${i+1}.jpg`,{type:'image/jpeg'}));
  const dt=new DataTransfer();
  [...existing,...newFiles].forEach(f=>dt.items.add(f));
  input.files=dt.files;
  toast(`${newFiles.length} foto${newFiles.length!==1?'s':''} adicionada${newFiles.length!==1?'s':''}!`,'ok');
  dzPreview(_camTarget,_camGrid);
  camClose();
}

function camClose(){
  if(_camStream){_camStream.getTracks().forEach(t=>t.stop());_camStream=null;}
  el('camVideo').srcObject=null;
  el('camOverlay').classList.remove('show');
  _camBlobs=[];
}

// ── MOBILE NAV ──
function isMobile(){return window.innerWidth<=768;}

function openMobMenu(){
  el('sidebar-aside').classList.add('mob-open');
  el('mobOverlay').classList.add('show');
}
function closeMobMenu(){
  el('sidebar-aside').classList.remove('mob-open');
  el('mobOverlay').classList.remove('show');
}

function syncBottomNav(page){
  document.querySelectorAll('.bottom-nav a[data-page]').forEach(a=>{
    a.classList.toggle('active', a.dataset.page===page);
  });
}

document.querySelectorAll('.bottom-nav a[data-page]').forEach(a=>{
  a.addEventListener('click',()=>{
    closeMobMenu();
    goPage(a.dataset.page);
  });
});

// Show/hide hamburger on resize
function handleResize(){
  const mob=isMobile();
  const btn=el('mobMenuBtn');
  if(btn) btn.style.display=mob?'block':'none';
}
window.addEventListener('resize',handleResize);
handleResize();

// Patch goPage to sync bottom nav
const _origGoPage=goPage;
goPage=function(p){
  _origGoPage(p);
  syncBottomNav(p);
  if(isMobile()) closeMobMenu();
};

// ── BOOT ──
el('topDate').textContent='Ref: '+new Date().toLocaleDateString('pt-BR');
el('sidebar-brand-logo').src=document.querySelector('.auth-logo-white').src;
resetNovaForm();
resetAtivForm();
bootAuth();
