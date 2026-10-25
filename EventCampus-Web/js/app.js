// js/app.js — Painel Web EventCampus (Professor e Coordenação) integrado à API

// ── Estado global ────────────────────────────────────────────
let usuarioAtual = null;
let paginaAtual = null;
let eventoSelecionadoId = null;
let cacheEventos = [];          // eventos do professor (ou todos, para admin)
let timerQR = null;             // relógio da tela de QR
let pollQR = null;              // atualização automática de presentes
let qrEstado = null;            // último retorno de /qr

// ════════════════════════════════════════════════════════════
// UTILITÁRIOS
// ════════════════════════════════════════════════════════════
function esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/** "2026-10-21" → "21/10/2026" (sem passar por Date, para não errar o fuso) */
function fmtData(iso) {
  if (!iso) return '—';
  const [a, m, d] = String(iso).slice(0, 10).split('-');
  return d && m && a ? `${d}/${m}/${a}` : String(iso);
}
/** "19:00:00" → "19:00" */
function fmtHora(h) { return h ? String(h).slice(0, 5) : ''; }
/** "2026-10-21T19:10:00" → "21/10/2026 19:10" */
function fmtDataHora(iso) {
  if (!iso) return '—';
  const s = String(iso);
  return fmtData(s.slice(0, 10)) + (s.length > 10 ? ' ' + s.slice(11, 16) : '');
}
function horarioEvento(e) {
  const i = fmtHora(e.horarioInicio), f = fmtHora(e.horarioFim);
  return i && f ? `${i}–${f}` : (i || '—');
}
function iniciais(nome) {
  return String(nome || '?').replace(/^(Prof\.?|Profa\.?)\s+/i, '').split(/\s+/).filter(Boolean)
    .slice(0, 2).map(p => p[0].toUpperCase()).join('');
}
function hojeISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
const pct = (a, b) => (b > 0 ? Math.round((a / b) * 100) : 0);

const STATUS_EVENTO = {
  RASCUNHO:  { label: 'Rascunho',  badge: 'badge-gray'  },
  ATIVO:     { label: 'Ativo',     badge: 'badge-green' },
  CONCLUIDO: { label: 'Concluído', badge: 'badge-blue'  },
  ENCERRADO: { label: 'Encerrado', badge: 'badge-red'   },
};
const TIPOS = { PALESTRA: 'Palestra', WORKSHOP: 'Workshop', CURSO: 'Curso', SEMINARIO: 'Seminário' };
const badgeStatus = s => `<span class="badge ${(STATUS_EVENTO[s] || {}).badge || 'badge-gray'}">${(STATUS_EVENTO[s] || {}).label || esc(s)}</span>`;

function toast(msg, tipo = 'ok') {
  const el = document.createElement('div');
  el.className = 'toast ' + tipo;
  el.innerHTML = icone(tipo === 'erro' ? 'circle-alert' : 'circle-check', 18) + '<span>' + esc(msg) + '</span>';
  document.getElementById('toasts').appendChild(el);
  setTimeout(() => el.remove(), tipo === 'erro' ? 6000 : 3500);
}
function erro(e) { console.error(e); toast(e.message || 'Algo deu errado', 'erro'); }

const carregando = (txt = 'Carregando...') => `<div class="loading-box"><div class="spinner"></div>${esc(txt)}</div>`;
const vazio = (ic, txt) => `<div class="empty"><div class="empty-icon">${icone(ic, 40)}</div><p>${esc(txt)}</p></div>`;

function setPagina(id, html) {
  const el = document.getElementById('page-' + id);
  if (el) el.innerHTML = html;
}

/** Desabilita o botão enquanto a promessa roda. */
async function comLoading(btn, fn) {
  if (btn) btn.classList.add('loading');
  try { return await fn(); }
  finally { if (btn) btn.classList.remove('loading'); }
}

// ── Exportações (sem bibliotecas) ────────────────────────────
/** Gera um .csv que o Excel abre direto (separador ; e acentos ok). */
function baixarCSV(nomeArquivo, cabecalho, linhas) {
  const cel = v => {
    const s = String(v ?? '');
    return /[;"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const csv = [cabecalho, ...linhas].map(l => l.map(cel).join(';')).join('\r\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = nomeArquivo.replace(/[\\/:*?"<>|]+/g, '_') + '.csv';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/** Abre uma janela com o HTML e chama a impressão (o usuário escolhe "Salvar como PDF"). */
function imprimirHTML(titulo, corpo, { paisagem = false } = {}) {
  const w = window.open('', '_blank');
  if (!w) { toast('Permita pop-ups para gerar o PDF', 'erro'); return; }
  w.document.write(`<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8"><title>${esc(titulo)}</title>
  <style>
    @page { size: A4 ${paisagem ? 'landscape' : 'portrait'}; margin: ${paisagem ? '0' : '18mm'}; }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Segoe UI', Arial, sans-serif; color: #2E2A45; font-size: 12px;
           -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    h1 { font-size: 20px; color: #5A4FB0; margin-bottom: 4px; }
    h2 { font-size: 14px; margin: 18px 0 8px; color: #2E2A45; }
    .sub { color: #8580A3; margin-bottom: 14px; }
    table { width: 100%; border-collapse: collapse; margin-top: 6px; }
    th, td { border: 1px solid #ddd; padding: 6px 8px; text-align: left; }
    th { background: #EFEDFC; font-size: 11px; text-transform: uppercase; }
    .stats { display: flex; gap: 10px; margin: 12px 0; }
    .stat { flex: 1; border: 1px solid #ddd; border-radius: 8px; padding: 10px; text-align: center; }
    .stat b { display: block; font-size: 18px; color: #5A4FB0; }
    .rodape { margin-top: 24px; color: #8580A3; font-size: 10px; }
  </style></head><body>${corpo}
  <script>window.onload = () => { setTimeout(() => window.print(), 300); };<\/script></body></html>`);
  w.document.close();
}

// ════════════════════════════════════════════════════════════
// MODAL GENÉRICO
// ════════════════════════════════════════════════════════════
function abrirModal(titulo, corpoHTML, { grande = false } = {}) {
  document.getElementById('modal-titulo').textContent = titulo;
  document.getElementById('modal-corpo').innerHTML = corpoHTML;
  document.getElementById('modal-box').classList.toggle('modal-lg', grande);
  document.getElementById('modal').classList.add('open');
  const primeiro = document.querySelector('#modal-corpo input, #modal-corpo textarea, #modal-corpo select');
  if (primeiro) setTimeout(() => primeiro.focus(), 50);
}
function fecharModal() { document.getElementById('modal').classList.remove('open'); }
function erroModal(msg) {
  const el = document.getElementById('modal-erro');
  if (el) { el.textContent = msg; el.classList.add('show'); } else toast(msg, 'erro');
}

/** Confirmação simples. Retorna Promise<boolean>. */
function confirmar(titulo, mensagem, textoOk = 'Confirmar', perigo = false) {
  return new Promise(resolve => {
    abrirModal(titulo, `
      <p class="text-sm" style="margin-bottom:20px;">${mensagem}</p>
      <div class="modal-actions">
        <button class="btn btn-outline" id="conf-nao">Cancelar</button>
        <button class="btn ${perigo ? 'btn-danger' : 'btn-primary'}" id="conf-sim">${esc(textoOk)}</button>
      </div>`);
    document.getElementById('conf-nao').onclick = () => { fecharModal(); resolve(false); };
    document.getElementById('conf-sim').onclick = () => { fecharModal(); resolve(true); };
  });
}

document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { fecharModal(); fecharQRTelaCheia(); fecharDropdowns(); }
});
document.addEventListener('click', e => { if (!e.target.closest('.dropdown')) fecharDropdowns(); });
function fecharDropdowns() { document.querySelectorAll('.dropdown.open').forEach(d => d.classList.remove('open')); }
function toggleDropdown(btn) {
  const dd = btn.closest('.dropdown');
  const aberto = dd.classList.contains('open');
  fecharDropdowns();
  if (!aberto) dd.classList.add('open');
}

// ════════════════════════════════════════════════════════════
// LOGIN / SESSÃO
// ════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('ano-atual').textContent = new Date().getFullYear();
  aplicarIcones();
  if (window.innerWidth <= 768) document.getElementById('menu-btn').style.display = 'flex';

  // Sessão salva? Confere com a API antes de entrar.
  if (Sessao.token) {
    try {
      const me = await API.me();
      usuarioAtual = { ...Sessao.usuario, ...me };
      if (usuarioAtual.perfil === 'ALUNO') throw new Error('aluno');
      entrarNoApp();
      return;
    } catch { Sessao.limpar(); }
  }
  mostrarLogin();
});

function mostrarLogin() {
  pararQR();
  document.getElementById('login-screen').style.display = 'flex';
  document.getElementById('app').style.display = 'none';
  setTimeout(() => document.getElementById('login-email').focus(), 50);
}

function preencherLogin(email) {
  document.getElementById('login-email').value = email;
  document.getElementById('login-senha').value = '123456';
}

async function fazerLogin() {
  const email = document.getElementById('login-email').value.trim();
  const senha = document.getElementById('login-senha').value;
  const erroEl = document.getElementById('login-erro');
  erroEl.classList.remove('show');

  await comLoading(document.getElementById('btn-login'), async () => {
    try {
      const dados = await API.login(email, senha);
      if (dados.perfil === 'ALUNO') {
        erroEl.textContent = 'Alunos acessam pelo aplicativo EventCampus no celular.';
        erroEl.classList.add('show');
        return;
      }
      Sessao.salvar(dados);
      usuarioAtual = dados;
      document.getElementById('login-senha').value = '';
      entrarNoApp();
    } catch (e) {
      erroEl.textContent = e.message;
      erroEl.classList.add('show');
    }
  });
}

function fazerLogout() {
  Sessao.limpar();
  usuarioAtual = null;
  eventoSelecionadoId = null;
  cacheEventos = [];
  mostrarLogin();
}

function sessaoExpirada() {
  if (!usuarioAtual) return;
  usuarioAtual = null;
  toast('Sua sessão expirou. Entre novamente.', 'erro');
  mostrarLogin();
}

const ehAdmin = () => usuarioAtual && usuarioAtual.perfil === 'ADMIN';

function entrarNoApp() {
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('app').style.display = 'flex';
  renderSidebar();
  navegarPara(ehAdmin() ? 'dashboard' : 'eventos');
}

function toggleSidebar() { document.getElementById('sidebar').classList.toggle('open'); }

// ════════════════════════════════════════════════════════════
// SIDEBAR / NAVEGAÇÃO
// ════════════════════════════════════════════════════════════
function renderSidebar() {
  const u = usuarioAtual;
  document.getElementById('sidebar-iniciais').textContent = iniciais(u.nome);
  document.getElementById('topbar-avatar').textContent = iniciais(u.nome);
  document.getElementById('sidebar-nome').textContent = u.nome;
  document.getElementById('sidebar-email').textContent = u.email;
  const badge = document.getElementById('sidebar-badge');
  badge.textContent = ehAdmin() ? 'Admin' : 'Professor';
  badge.className = 'sidebar-badge ' + (ehAdmin() ? 'badge-admin' : 'badge-prof');

  const item = (pag, ic, txt, extra = '') =>
    `<div class="nav-item" onclick="navegarPara('${pag}')" data-page="${pag}"><span class="nav-icon">${icone(ic, 18)}</span> ${txt}${extra}</div>`;

  const eventos = `
    <div class="nav-section">Eventos</div>
    ${ehAdmin() ? item('eventos-admin', 'calendar', 'Todos os Eventos') : item('eventos', 'calendar', 'Meus Eventos')}
    ${item('qrcode', 'qr-code', 'QR Code (Presença)')}
    ${item('inscritos', 'users', 'Inscritos')}
    ${item('relatorio', 'chart-column', 'Relatórios')}`;

  document.getElementById('sidebar-nav').innerHTML = ehAdmin() ? `
    <div class="nav-section">Visão Geral</div>
    ${item('dashboard', 'layout-dashboard', 'Dashboard')}
    <div class="nav-section">Acadêmico</div>
    ${item('cert-externos', 'file-text', 'Certificados Externos', '<span class="nav-badge" id="badge-pendentes" style="display:none;"></span>')}
    ${item('horas', 'clock', 'Horas Complementares')}
    ${item('cert-final', 'award', 'Certificado Final')}
    ${eventos}
    <div class="nav-section">Sistema</div>
    ${item('usuarios', 'user-cog', 'Usuários')}
  ` : eventos;

  if (ehAdmin()) atualizarBadgePendentes();
}

async function atualizarBadgePendentes() {
  try {
    const lista = await API.certificados('PENDENTE');
    const el = document.getElementById('badge-pendentes');
    if (!el) return;
    el.textContent = lista.length;
    el.style.display = lista.length ? '' : 'none';
  } catch { /* silencioso */ }
}

const TITULOS = {
  'dashboard': 'Dashboard Geral',
  'eventos': 'Meus Eventos',
  'eventos-admin': 'Todos os Eventos',
  'qrcode': 'QR Code de Presença',
  'inscritos': 'Inscritos do Evento',
  'relatorio': 'Relatório do Evento',
  'cert-externos': 'Certificados Externos',
  'horas': 'Horas Complementares',
  'cert-final': 'Certificado Final',
  'usuarios': 'Usuários',
};

function navegarPara(pagina) {
  if (paginaAtual === 'qrcode' && pagina !== 'qrcode') pararQR();
  paginaAtual = pagina;
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + pagina)?.classList.add('active');
  document.querySelector(`[data-page="${pagina}"]`)?.classList.add('active');
  document.getElementById('topbar-title').textContent = TITULOS[pagina] || 'EventCampus';
  document.getElementById('sidebar').classList.remove('open');

  const renders = {
    'dashboard': renderDashboard,
    'eventos': renderEventos,
    'eventos-admin': renderEventos,
    'qrcode': renderQRCode,
    'inscritos': renderInscritos,
    'relatorio': renderRelatorio,
    'cert-externos': renderCertExternos,
    'horas': renderHoras,
    'cert-final': renderCertFinal,
    'usuarios': renderUsuarios,
  };
  renders[pagina]?.();
}

function recarregarPagina() { if (paginaAtual) navegarPara(paginaAtual); }

function abrirEvento(id, pagina) {
  eventoSelecionadoId = id;
  navegarPara(pagina);
}

// ════════════════════════════════════════════════════════════
// EVENTOS (lista, criar, editar, status)
// ════════════════════════════════════════════════════════════
async function carregarEventos() {
  cacheEventos = await API.meusEventos();   // professor: os dele | admin: todos
  return cacheEventos;
}

async function renderEventos() {
  const pag = ehAdmin() ? 'eventos-admin' : 'eventos';
  setPagina(pag, carregando('Carregando eventos...'));
  try {
    const eventos = await carregarEventos();
    const filtro = renderEventos.filtro || 'TODOS';
    const lista = filtro === 'TODOS' ? eventos : eventos.filter(e => e.status === filtro);
    const contagem = s => eventos.filter(e => e.status === s).length;

    const linhas = lista.map(e => `
      <tr>
        <td>
          <p class="fw-600">${esc(e.titulo)}</p>
          <p class="text-mute">${TIPOS[e.tipo] || esc(e.tipo)} · ${esc(e.local || 'Local a definir')}${ehAdmin() && e.professor ? ' · ' + esc(e.professor) : ''}</p>
        </td>
        <td>${fmtData(e.data)}<br><span class="text-mute">${horarioEvento(e)}</span></td>
        <td>${e.inscritos} / ${e.vagas}</td>
        <td>${e.presentes}</td>
        <td>${e.cargaHoraria}h</td>
        <td>${badgeStatus(e.status)} ${e.qrAtivo ? '<span class="badge badge-amber"><span class="dot"></span>QR ativo</span>' : ''}</td>
        <td>
          <div class="row" style="flex-wrap:nowrap;">
            <button class="btn btn-outline btn-sm" onclick="abrirEvento(${e.id},'qrcode')">QR</button>
            <button class="btn btn-outline btn-sm" onclick="abrirEvento(${e.id},'inscritos')">Inscritos</button>
            <button class="btn btn-outline btn-sm" onclick="abrirEvento(${e.id},'relatorio')">Relatório</button>
            <div class="dropdown">
              <button class="btn btn-ghost btn-sm" onclick="toggleDropdown(this)" title="Mais ações">${icone('ellipsis', 18)}</button>
              <div class="dropdown-menu">
                <button onclick="abrirFormEvento(${e.id})">${icone('pencil', 14)} Editar</button>
                ${e.status !== 'ATIVO' ? `<button onclick="mudarStatus(${e.id},'ATIVO')">${icone('send', 14)} Publicar (Ativo)</button>` : ''}
                ${e.status !== 'CONCLUIDO' ? `<button onclick="mudarStatus(${e.id},'CONCLUIDO')">${icone('flag', 14)} Marcar como concluído</button>` : ''}
                ${e.status !== 'RASCUNHO' ? `<button onclick="mudarStatus(${e.id},'RASCUNHO')">${icone('undo-2', 14)} Voltar para rascunho</button>` : ''}
                ${e.status !== 'ENCERRADO' ? `<button class="danger" onclick="mudarStatus(${e.id},'ENCERRADO')">${icone('ban', 14)} Encerrar/cancelar</button>` : ''}
                <button class="danger" onclick="excluirEvento(${e.id})">${icone('trash-2', 14)} Excluir</button>
              </div>
            </div>
          </div>
        </td>
      </tr>`).join('');

    const aba = (s, txt, n) =>
      `<button class="tab ${filtro === s ? 'active' : ''}" onclick="renderEventos.filtro='${s}';renderEventos()">${txt}<span class="count">${n}</span></button>`;

    setPagina(pag, `
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:20px;font-weight:600;">${ehAdmin() ? 'Todos os Eventos' : 'Meus Eventos'}</h2>
          <p class="text-mute mt-4">Crie, publique e acompanhe a presença dos eventos</p>
        </div>
        <button class="btn btn-primary" onclick="abrirFormEvento()">${icone('plus')} Novo Evento</button>
      </div>
      <div class="card" style="padding:0;">
        <div class="tabs">
          ${aba('TODOS', 'Todos', eventos.length)}
          ${aba('ATIVO', 'Ativos', contagem('ATIVO'))}
          ${aba('RASCUNHO', 'Rascunhos', contagem('RASCUNHO'))}
          ${aba('CONCLUIDO', 'Concluídos', contagem('CONCLUIDO'))}
          ${aba('ENCERRADO', 'Encerrados', contagem('ENCERRADO'))}
        </div>
        <div class="table-wrap" style="overflow:visible;">
          ${lista.length ? `<table>
            <thead><tr><th>Evento</th><th>Data</th><th>Inscritos/Vagas</th><th>Presentes</th><th>Carga</th><th>Status</th><th>Ações</th></tr></thead>
            <tbody>${linhas}</tbody>
          </table>` : vazio('inbox', eventos.length ? 'Nenhum evento com este status.' : 'Você ainda não criou eventos. Clique em "+ Novo Evento".')}
        </div>
      </div>
      <p class="text-mute row" style="gap:6px;">${icone('info', 14)}<span>Eventos em <b>Rascunho</b> não aparecem para os alunos. Publique para abrir as inscrições.</span></p>
    `);
  } catch (e) {
    setPagina(pag, vazio('triangle-alert', e.message));
  }
}

function abrirFormEvento(id) {
  fecharDropdowns();
  const e = id ? cacheEventos.find(x => x.id === id) : null;
  const v = (campo, padrao = '') => esc(e ? (e[campo] ?? padrao) : padrao);
  abrirModal(e ? 'Editar Evento' : 'Novo Evento', `
    <div class="form-error" id="modal-erro"></div>
    <div class="form-group">
      <label class="form-label">Título do evento *</label>
      <input class="form-input" id="ev-titulo" placeholder="Ex: Palestra de Inovação" value="${v('titulo')}" />
    </div>
    <div class="form-group">
      <label class="form-label">Descrição</label>
      <textarea class="form-input" id="ev-descricao" placeholder="Sobre o que é o evento?">${v('descricao')}</textarea>
    </div>
    <div class="grid-2">
      <div class="form-group">
        <label class="form-label">Tipo</label>
        <select class="form-select" id="ev-tipo">
          ${Object.entries(TIPOS).map(([k, t]) => `<option value="${k}" ${e && e.tipo === k ? 'selected' : ''}>${t}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Carga horária (h) *</label>
        <input class="form-input" id="ev-carga" type="number" min="1" value="${v('cargaHoraria', 4)}" />
      </div>
    </div>
    <div class="grid-2">
      <div class="form-group">
        <label class="form-label">Data *</label>
        <input class="form-input" id="ev-data" type="date" value="${e ? esc(e.data) : hojeISO()}" />
      </div>
      <div class="form-group">
        <label class="form-label">Vagas *</label>
        <input class="form-input" id="ev-vagas" type="number" min="1" value="${v('vagas', 50)}" />
      </div>
    </div>
    <div class="grid-2">
      <div class="form-group">
        <label class="form-label">Início</label>
        <input class="form-input" id="ev-inicio" type="time" value="${e ? fmtHora(e.horarioInicio) : '19:00'}" />
      </div>
      <div class="form-group">
        <label class="form-label">Término</label>
        <input class="form-input" id="ev-fim" type="time" value="${e ? fmtHora(e.horarioFim) : '21:00'}" />
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Local</label>
      <input class="form-input" id="ev-local" placeholder="Ex: Auditório Central" value="${v('local')}" />
    </div>
    ${e ? '' : `<label class="check mb-16"><input type="checkbox" id="ev-publicar" checked /> Publicar agora (alunos já podem se inscrever)</label>`}
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="fecharModal()">Cancelar</button>
      <button class="btn btn-primary" id="ev-salvar" onclick="salvarEvento(${e ? e.id : 'null'}, this)">Salvar evento</button>
    </div>`, { grande: true });
}

async function salvarEvento(id, btn) {
  const val = s => document.getElementById(s).value.trim();
  const dados = {
    titulo: val('ev-titulo'),
    descricao: val('ev-descricao'),
    tipo: val('ev-tipo'),
    cargaHoraria: parseInt(val('ev-carga'), 10),
    data: val('ev-data'),
    vagas: parseInt(val('ev-vagas'), 10),
    horarioInicio: val('ev-inicio') || null,
    horarioFim: val('ev-fim') || null,
    local: val('ev-local'),
  };
  if (!dados.titulo) return erroModal('Informe o título do evento.');
  if (!dados.data) return erroModal('Informe a data do evento.');
  if (!(dados.cargaHoraria > 0)) return erroModal('A carga horária deve ser maior que zero.');
  if (!(dados.vagas > 0)) return erroModal('O número de vagas deve ser maior que zero.');
  if (dados.horarioInicio && dados.horarioFim && dados.horarioFim < dados.horarioInicio)
    return erroModal('O término deve ser depois do início.');
  if (!id && document.getElementById('ev-publicar')?.checked) dados.status = 'ATIVO';

  await comLoading(btn, async () => {
    try {
      const salvo = id ? await API.editarEvento(id, dados) : await API.criarEvento(dados);
      fecharModal();
      toast(id ? 'Evento atualizado!' : `Evento "${salvo.titulo}" criado!`);
      eventoSelecionadoId = salvo.id;
      recarregarPagina();
    } catch (e) { erroModal(e.message); }
  });
}

async function mudarStatus(id, status) {
  fecharDropdowns();
  const textos = {
    ATIVO: 'Publicar o evento? Os alunos poderão vê-lo e se inscrever.',
    CONCLUIDO: 'Marcar como concluído? A presença por QR será fechada.',
    RASCUNHO: 'Voltar para rascunho? O evento some do app dos alunos.',
    ENCERRADO: 'Encerrar/cancelar o evento? Ele deixa de aceitar inscrições e presença.',
  };
  if (!(await confirmar('Alterar status', textos[status], 'Sim, alterar', status === 'ENCERRADO'))) return;
  try {
    await API.statusEvento(id, status);
    toast('Status alterado para ' + STATUS_EVENTO[status].label);
    recarregarPagina();
  } catch (e) { erro(e); }
}

async function excluirEvento(id) {
  fecharDropdowns();
  if (!(await confirmar('Excluir evento', 'Esta ação não pode ser desfeita. Continuar?', 'Excluir', true))) return;
  try {
    await API.excluirEvento(id);
    if (eventoSelecionadoId === id) eventoSelecionadoId = null;
    toast('Evento excluído');
    recarregarPagina();
  } catch (e) { erro(e); }
}

/** Seletor de evento usado nas telas de QR, Inscritos e Relatório. */
async function seletorEvento(pagina, filtro = () => true) {
  if (!cacheEventos.length) await carregarEventos();
  const lista = cacheEventos.filter(filtro);
  if (!lista.length) return { html: '', evento: null, lista };
  if (!eventoSelecionadoId || !lista.some(e => e.id === eventoSelecionadoId)) {
    const ativo = lista.find(e => e.status === 'ATIVO');
    eventoSelecionadoId = (ativo || lista[0]).id;
  }
  const evento = lista.find(e => e.id === eventoSelecionadoId);
  const html = `
    <div class="event-picker no-print">
      <label class="form-label" style="margin:0;">Evento:</label>
      <select class="form-select" onchange="abrirEvento(Number(this.value),'${pagina}')">
        ${lista.map(e => `<option value="${e.id}" ${e.id === eventoSelecionadoId ? 'selected' : ''}>
          ${esc(e.titulo)} — ${fmtData(e.data)} (${(STATUS_EVENTO[e.status] || {}).label || e.status})</option>`).join('')}
      </select>
    </div>`;
  return { html, evento, lista };
}

// ════════════════════════════════════════════════════════════
// QR CODE DE PRESENÇA
// ════════════════════════════════════════════════════════════
function svgQR(texto, tamanhoCelula = 8) {
  const qr = qrcode(0, 'M');
  qr.addData(texto);
  qr.make();
  return qr.createSvgTag({ cellSize: tamanhoCelula, margin: 2, scalable: true });
}

function pararQR() {
  clearInterval(timerQR); timerQR = null;
  clearInterval(pollQR); pollQR = null;
}

async function renderQRCode() {
  pararQR();
  setPagina('qrcode', carregando());
  try {
    const { html: picker, evento } = await seletorEvento('qrcode', e => e.status !== 'ENCERRADO');
    if (!evento) {
      setPagina('qrcode', vazio('qr-code', 'Nenhum evento disponível. Crie e publique um evento primeiro.'));
      return;
    }
    qrEstado = await API.qr(evento.id);
    const q = qrEstado;
    const podeAtivar = evento.status === 'ATIVO' || evento.status === 'RASCUNHO';
    const temLocal = q.latitude != null && q.longitude != null;

    setPagina('qrcode', `
      ${picker}
      <div class="grid-2" style="align-items:start;gap:24px;">
        <div>
          <div class="card">
            <p class="card-title">Evento selecionado</p>
            <p class="fw-600" style="font-size:15px;margin-bottom:4px;">${esc(evento.titulo)}</p>
            <p class="text-mute">${fmtData(evento.data)} · ${horarioEvento(evento)} · ${esc(evento.local || '')}</p>
            <p class="mt-8">${badgeStatus(evento.status)} <span class="badge badge-blue">${evento.cargaHoraria}h</span></p>
            <div class="row mt-16" style="gap:24px;">
              <div style="text-align:center;"><p style="font-size:22px;font-weight:700;color:var(--mint-dark);" id="qr-presentes">${q.presentes}</p><p class="text-mute">Presentes</p></div>
              <div style="text-align:center;"><p style="font-size:22px;font-weight:700;">${q.inscritos}</p><p class="text-mute">Inscritos</p></div>
              <div style="text-align:center;"><p style="font-size:22px;font-weight:700;color:var(--peach);" id="qr-pendentes">${q.inscritos - q.presentes}</p><p class="text-mute">Pendentes</p></div>
            </div>
          </div>

          <div class="card">
            <p class="card-title">Configurar presença</p>
            <div class="grid-2">
              <div class="form-group">
                <label class="form-label">Raio permitido (metros)</label>
                <input class="form-input" id="qr-raio" type="number" min="10" value="${q.raioMetros || 100}" />
              </div>
              <div class="form-group">
                <label class="form-label">Aberto por (minutos)</label>
                <input class="form-input" id="qr-minutos" type="number" min="0" placeholder="vazio = até encerrar" value="30" />
              </div>
            </div>
            <label class="check mb-8"><input type="checkbox" id="qr-usar-local" checked /> Usar a localização deste computador como local do evento</label>
            <p class="form-hint mb-16">
              ${temLocal ? `${icone('map-pin', 12)} Local salvo: ${Number(q.latitude).toFixed(5)}, ${Number(q.longitude).toFixed(5)}` : `${icone('map-pin', 12)} Sem local salvo: sem localização, o app aceita presença de qualquer lugar.`}
              <br>O aluno só confirma presença se estiver dentro do raio (cálculo de Haversine).
            </p>
            ${podeAtivar ? `
              <button class="btn btn-primary" style="width:100%;justify-content:center;" onclick="ativarPresenca(this)">${icone('play', 14)} ${q.qrAtivo ? 'Reativar / atualizar' : 'Ativar presença'}</button>
              <button class="btn btn-danger mt-8" style="width:100%;justify-content:center;" onclick="encerrarPresenca(this)" ${q.qrAtivo ? '' : 'disabled'}>${icone('square', 14)} Encerrar presença</button>
            ` : `<p class="text-mute">Este evento está ${STATUS_EVENTO[evento.status].label.toLowerCase()}: a presença por QR não pode ser aberta.</p>`}
          </div>
        </div>

        <div class="card" style="text-align:center;">
          <p class="card-title" style="text-align:left;">QR Code ao vivo</p>
          <div class="qr-display">
            <div>${q.qrAtivo ? '<span class="badge badge-green"><span class="dot"></span>Presença ativa</span>' : '<span class="badge badge-gray">Presença fechada</span>'}</div>
            <div class="qr-svg ${q.qrAtivo ? '' : 'qr-inativo'}" style="width:260px;">${svgQR(q.qrCodigo)}</div>
            <div class="qr-timer" id="qr-timer">${q.qrAtivo ? '--:--' : '00:00'}</div>
            <p class="qr-label" id="qr-timer-label">${q.qrAtivo ? (q.qrExpiraEm ? 'tempo restante' : 'aberto até você encerrar') : 'Ative a presença para os alunos escanearem pelo app'}</p>
            <div class="row" style="justify-content:center;flex-wrap:wrap;">
              <button class="btn btn-outline btn-sm" onclick="regenerarQR(this)" title="Gera um código novo; prints/fotos do antigo param de funcionar">${icone('refresh-cw', 14)} Regenerar código</button>
              <button class="btn btn-primary btn-sm" onclick="abrirQRTelaCheia()" ${q.qrAtivo ? '' : 'disabled'}>${icone('maximize', 14)} Projetar em tela cheia</button>
            </div>
          </div>
          <div class="mt-16">
            <button class="btn btn-outline" style="width:100%;justify-content:center;" onclick="navegarPara('inscritos')">${icone('users')} Ver lista de presença</button>
          </div>
        </div>
      </div>
    `);

    if (q.qrAtivo) iniciarRelogioQR();
  } catch (e) {
    setPagina('qrcode', vazio('triangle-alert', e.message));
  }
}

function iniciarRelogioQR() {
  pararQR();
  const inicio = Date.now();
  // fim calculado a partir dos segundos informados pela API (não depende do fuso do navegador)
  const fim = qrEstado && qrEstado.segundosRestantes != null ? Date.now() + qrEstado.segundosRestantes * 1000 : null;
  const tick = () => {
    const el = document.getElementById('qr-timer');
    if (!el) return;
    let seg;
    if (fim) {
      seg = Math.max(0, Math.round((fim - Date.now()) / 1000));
      if (seg === 0) { pararQR(); toast('Tempo de presença encerrado'); renderQRCode(); return; }
    } else {
      seg = Math.round((Date.now() - inicio) / 1000);
    }
    el.textContent = `${String(Math.floor(seg / 60)).padStart(2, '0')}:${String(seg % 60).padStart(2, '0')}`;
  };
  tick();
  timerQR = setInterval(tick, 1000);

  // Atualiza o número de presentes a cada 5 s enquanto o QR está ativo
  pollQR = setInterval(async () => {
    if (paginaAtual !== 'qrcode' || !eventoSelecionadoId) return;
    try {
      const q = await API.qr(eventoSelecionadoId);
      const antes = qrEstado ? qrEstado.presentes : q.presentes;
      qrEstado = q;
      const p = document.getElementById('qr-presentes');
      if (p) p.textContent = q.presentes;
      const pe = document.getElementById('qr-pendentes');
      if (pe) pe.textContent = q.inscritos - q.presentes;
      const fs = document.getElementById('qr-fs-presentes');
      if (fs) fs.textContent = `${q.presentes} presentes`;
      if (q.presentes > antes) toast(`+${q.presentes - antes} presença(s) confirmada(s)`);
      if (!q.qrAtivo) renderQRCode();
    } catch { /* ignora falhas pontuais */ }
  }, 5000);
}

function obterLocalizacao() {
  return new Promise(resolve => {
    if (!navigator.geolocation) return resolve(null);
    navigator.geolocation.getCurrentPosition(
      p => resolve({ latitude: p.coords.latitude, longitude: p.coords.longitude }),
      () => resolve(null),
      { enableHighAccuracy: true, timeout: 10000 }
    );
  });
}

async function ativarPresenca(btn) {
  const cfg = {
    raioMetros: parseInt(document.getElementById('qr-raio').value, 10) || 100,
    minutos: parseInt(document.getElementById('qr-minutos').value, 10) || null,
  };
  await comLoading(btn, async () => {
    if (document.getElementById('qr-usar-local').checked) {
      const loc = await obterLocalizacao();
      if (loc) Object.assign(cfg, loc);
      else toast('Não foi possível obter a localização do navegador. Mantendo o local salvo (se houver).', 'erro');
    }
    try {
      await API.ativarQR(eventoSelecionadoId, cfg);
      toast('Presença ativada! Os alunos já podem escanear.');
      await carregarEventos();
      renderQRCode();
    } catch (e) { erro(e); }
  });
}

async function encerrarPresenca(btn) {
  await comLoading(btn, async () => {
    try {
      await API.encerrarQR(eventoSelecionadoId);
      toast('Presença encerrada.');
      await carregarEventos();
      renderQRCode();
    } catch (e) { erro(e); }
  });
}

async function regenerarQR(btn) {
  if (!(await confirmar('Regenerar QR Code', 'O código atual deixará de funcionar (fotos ou prints compartilhados ficam inválidos). Continuar?', 'Regenerar'))) return;
  try {
    await API.regenerarQR(eventoSelecionadoId);
    toast('Novo QR Code gerado');
    renderQRCode();
  } catch (e) { erro(e); }
}

function abrirQRTelaCheia() {
  if (!qrEstado) return;
  const ev = cacheEventos.find(e => e.id === eventoSelecionadoId);
  document.getElementById('qr-fs-titulo').textContent = ev ? ev.titulo : 'EventCampus';
  document.getElementById('qr-fs-codigo').innerHTML = svgQR(qrEstado.qrCodigo, 12);
  document.getElementById('qr-fs-presentes').textContent = `${qrEstado.presentes} presentes`;
  document.getElementById('qr-fullscreen').classList.add('open');
}
function fecharQRTelaCheia() { document.getElementById('qr-fullscreen').classList.remove('open'); }

// ════════════════════════════════════════════════════════════
// INSCRITOS
// ════════════════════════════════════════════════════════════
let cacheInscritos = [];

async function renderInscritos() {
  setPagina('inscritos', carregando());
  try {
    const { html: picker, evento } = await seletorEvento('inscritos');
    if (!evento) { setPagina('inscritos', vazio('users', 'Nenhum evento cadastrado ainda.')); return; }
    cacheInscritos = await API.inscritos(evento.id);
    const presentes = cacheInscritos.filter(i => i.presencaConfirmada).length;

    setPagina('inscritos', `
      ${picker}
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:18px;font-weight:600;">${esc(evento.titulo)}</h2>
          <p class="text-mute mt-4">${fmtData(evento.data)} · ${esc(evento.local || '')} · ${evento.cargaHoraria}h</p>
        </div>
        <div class="row">
          <span class="badge badge-green">${presentes} presentes</span>
          <span class="badge badge-amber">${cacheInscritos.length - presentes} pendentes</span>
          <button class="btn btn-outline btn-sm" onclick="exportarInscritos()" ${cacheInscritos.length ? '' : 'disabled'}>${icone('download', 14)} Exportar Excel</button>
        </div>
      </div>
      <div class="card" style="padding:0;margin-bottom:16px;">
        <div style="padding:14px 20px;border-bottom:1px solid var(--border);">
          <input class="form-input search-input" placeholder="Buscar aluno, matrícula ou curso..." style="max-width:340px;" oninput="filtrarTabela('tabela-inscritos', this.value)">
        </div>
        <div class="table-wrap" id="tabela-inscritos">
          ${cacheInscritos.length ? `<table>
            <thead><tr><th>Aluno</th><th>Matrícula</th><th>Curso</th><th>Status</th><th>Confirmação</th><th>Ações</th></tr></thead>
            <tbody>${cacheInscritos.map(i => `
              <tr>
                <td><p class="fw-600">${esc(i.nome)}</p><p class="text-mute">${esc(i.email)}</p></td>
                <td>${esc(i.matricula)}</td>
                <td>${esc(i.curso || '—')}</td>
                <td><span class="badge ${i.presencaConfirmada ? 'badge-green' : 'badge-amber'}">${i.presencaConfirmada ? 'Presente' : 'Pendente'}</span></td>
                <td class="text-mute">${i.presencaConfirmada
                  ? `${fmtDataHora(i.confirmadaEm)}<br>${i.formaConfirmacao === 'MANUAL' ? icone('pen-line', 12) + ' manual' : icone('smartphone', 12) + ' QR Code'}${i.distanciaMetros != null ? ` · ${Math.round(i.distanciaMetros)} m` : ''}`
                  : '—'}</td>
                <td>
                  ${i.presencaConfirmada
                    ? `<button class="btn btn-danger btn-sm" onclick="presencaManual(${i.id}, false)">${icone('x', 14)} Remover</button>`
                    : `<button class="btn btn-success btn-sm" onclick="presencaManual(${i.id}, true)">${icone('check', 14)} Confirmar</button>`}
                </td>
              </tr>`).join('')}
            </tbody>
          </table>` : vazio('inbox', 'Nenhum aluno inscrito neste evento ainda.')}
        </div>
      </div>
    `);
  } catch (e) {
    setPagina('inscritos', vazio('triangle-alert', e.message));
  }
}

async function presencaManual(id, confirmarPresenca) {
  try {
    await API.presencaManual(id, confirmarPresenca);
    toast(confirmarPresenca ? 'Presença confirmada manualmente' : 'Presença removida');
    await carregarEventos();
    renderInscritos();
  } catch (e) { erro(e); }
}

function filtrarTabela(idWrap, busca) {
  const b = busca.toLowerCase();
  document.querySelectorAll(`#${idWrap} tbody tr`).forEach(tr => {
    tr.style.display = tr.textContent.toLowerCase().includes(b) ? '' : 'none';
  });
}

function exportarInscritos() {
  const ev = cacheEventos.find(e => e.id === eventoSelecionadoId) || {};
  baixarCSV(`Inscritos - ${ev.titulo || 'evento'}`,
    ['Nome', 'E-mail', 'Matrícula', 'Curso', 'Presença', 'Confirmada em', 'Forma', 'Distância (m)'],
    cacheInscritos.map(i => [i.nome, i.email, i.matricula, i.curso, i.presencaConfirmada ? 'Presente' : 'Ausente',
      i.presencaConfirmada ? fmtDataHora(i.confirmadaEm) : '', i.formaConfirmacao || '',
      i.distanciaMetros != null ? Math.round(i.distanciaMetros) : '']));
}

// ════════════════════════════════════════════════════════════
// RELATÓRIO DO EVENTO (RF-ADM-003)
// ════════════════════════════════════════════════════════════
let relatorioAtual = null;
const CORES = ['var(--lav)', 'var(--mint)', 'var(--sun)', 'var(--peach)', 'var(--lav-dark)', 'var(--red)'];

async function renderRelatorio() {
  setPagina('relatorio', carregando('Gerando relatório...'));
  try {
    const { html: picker, evento } = await seletorEvento('relatorio');
    if (!evento) { setPagina('relatorio', vazio('chart-column', 'Nenhum evento cadastrado ainda.')); return; }
    const r = relatorioAtual = await API.relatorioEvento(evento.id);
    const e = r.evento;

    const barras = r.porCurso.length ? r.porCurso.map((c, i) => `
      <div style="margin-bottom:12px;">
        <div class="row-between mb-4">
          <span class="text-sm">${esc(c.curso)}</span>
          <span class="text-sm fw-600">${c.presentes}/${c.inscritos} · ${c.percentual}%</span>
        </div>
        <div class="progress-bar"><div class="progress-fill" style="width:${c.percentual}%;background:${CORES[i % CORES.length]};"></div></div>
      </div>`).join('') : '<p class="text-mute">Sem inscritos ainda.</p>';

    setPagina('relatorio', `
      ${picker}
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:18px;font-weight:600;">${esc(e.titulo)}</h2>
          <p class="text-mute mt-4">${fmtData(e.data)} · ${horarioEvento(e)} · ${esc(e.local || '')} ${e.professor ? '· ' + esc(e.professor) : ''}</p>
        </div>
        <div class="row">
          <button class="btn btn-outline btn-sm" onclick="exportarRelatorioPDF()">${icone('file-text', 14)} Exportar PDF</button>
          <button class="btn btn-outline btn-sm" onclick="exportarRelatorioExcel()">${icone('file-spreadsheet', 14)} Exportar Excel</button>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat-card"><div class="stat-icon" style="background:var(--lav-light);color:var(--lav-dark);">${icone('users', 22)}</div><div class="stat-info"><p>${r.inscritos}</p><span>Inscritos de ${e.vagas} vagas</span></div></div>
        <div class="stat-card"><div class="stat-icon" style="background:var(--mint-light);color:var(--mint-dark);">${icone('circle-check', 22)}</div><div class="stat-info"><p>${r.presentes}</p><span>Presentes</span></div></div>
        <div class="stat-card"><div class="stat-icon" style="background:var(--sun-light);color:var(--sun-dark);">${icone('percent', 22)}</div><div class="stat-info"><p>${r.taxaPresenca}%</p><span>Taxa de presença</span></div></div>
        <div class="stat-card"><div class="stat-icon" style="background:var(--peach-light);color:var(--peach);">${icone('clock', 22)}</div><div class="stat-info"><p>${r.horasGeradas}h</p><span>Horas geradas (${e.cargaHoraria}h cada)</span></div></div>
      </div>

      <div class="grid-2">
        <div class="card">
          <p class="card-title">Inscritos por curso</p>
          ${barras}
        </div>
        <div class="card">
          <p class="card-title">Resumo</p>
          <table><tbody>
            ${[
              ['Tipo', TIPOS[e.tipo] || e.tipo],
              ['Status', (STATUS_EVENTO[e.status] || {}).label || e.status],
              ['Data', fmtData(e.data)],
              ['Horário', horarioEvento(e)],
              ['Local', e.local || '—'],
              ['Vagas', e.vagas],
              ['Ausentes', r.ausentes],
              ['Confirmados via QR Code', r.confirmadosViaQR],
              ['Confirmados manualmente', r.confirmadosManual],
            ].map(([k, v]) => `<tr><td class="text-mute">${k}</td><td class="fw-600">${esc(v)}</td></tr>`).join('')}
          </tbody></table>
        </div>
      </div>

      <div class="card" style="padding:0;">
        <p class="card-title" style="padding:16px 20px 0;">Lista de presença</p>
        <div class="table-wrap">
          ${r.alunos.length ? `<table>
            <thead><tr><th>Aluno</th><th>Matrícula</th><th>Curso</th><th>Presença</th><th>Confirmada em</th></tr></thead>
            <tbody>${r.alunos.map(a => `<tr>
              <td class="fw-600">${esc(a.nome)}</td><td>${esc(a.matricula)}</td><td>${esc(a.curso || '—')}</td>
              <td><span class="badge ${a.presente ? 'badge-green' : 'badge-red'}">${a.presente ? 'Presente' : 'Ausente'}</span></td>
              <td class="text-mute">${a.presente ? fmtDataHora(a.confirmadaEm) : '—'}</td></tr>`).join('')}
            </tbody></table>` : vazio('inbox', 'Nenhum inscrito.')}
        </div>
      </div>
    `);
  } catch (e) {
    setPagina('relatorio', vazio('triangle-alert', e.message));
  }
}

function exportarRelatorioExcel() {
  const r = relatorioAtual; if (!r) return;
  const e = r.evento;
  baixarCSV(`Relatorio - ${e.titulo}`, ['Relatório de presença — ' + e.titulo], [
    ['Data', fmtData(e.data)], ['Horário', horarioEvento(e)], ['Local', e.local || ''],
    ['Carga horária', e.cargaHoraria + 'h'], ['Inscritos', r.inscritos], ['Presentes', r.presentes],
    ['Ausentes', r.ausentes], ['Taxa de presença', r.taxaPresenca + '%'], ['Horas geradas', r.horasGeradas],
    [],
    ['Curso', 'Inscritos', 'Presentes', '% dos inscritos'],
    ...r.porCurso.map(c => [c.curso, c.inscritos, c.presentes, c.percentual + '%']),
    [],
    ['Aluno', 'Matrícula', 'Curso', 'Presença', 'Confirmada em', 'Forma'],
    ...r.alunos.map(a => [a.nome, a.matricula, a.curso, a.presente ? 'Presente' : 'Ausente',
      a.presente ? fmtDataHora(a.confirmadaEm) : '', a.forma || '']),
  ]);
}

function exportarRelatorioPDF() {
  const r = relatorioAtual; if (!r) return;
  const e = r.evento;
  imprimirHTML(`Relatório - ${e.titulo}`, `
    <h1>Relatório de Presença</h1>
    <p class="sub">EventCampus · UNIFAN — gerado em ${new Date().toLocaleString('pt-BR')}</p>
    <h2>${esc(e.titulo)}</h2>
    <p>${TIPOS[e.tipo] || ''} · ${fmtData(e.data)} · ${horarioEvento(e)} · ${esc(e.local || '')} ${e.professor ? '· Responsável: ' + esc(e.professor) : ''}</p>
    <div class="stats">
      <div class="stat"><b>${r.inscritos}</b>Inscritos</div>
      <div class="stat"><b>${r.presentes}</b>Presentes</div>
      <div class="stat"><b>${r.taxaPresenca}%</b>Taxa de presença</div>
      <div class="stat"><b>${r.horasGeradas}h</b>Horas geradas</div>
    </div>
    <h2>Por curso</h2>
    <table><thead><tr><th>Curso</th><th>Inscritos</th><th>Presentes</th><th>%</th></tr></thead><tbody>
      ${r.porCurso.map(c => `<tr><td>${esc(c.curso)}</td><td>${c.inscritos}</td><td>${c.presentes}</td><td>${c.percentual}%</td></tr>`).join('')}
    </tbody></table>
    <h2>Lista de presença</h2>
    <table><thead><tr><th>#</th><th>Aluno</th><th>Matrícula</th><th>Curso</th><th>Presença</th><th>Confirmada em</th></tr></thead><tbody>
      ${r.alunos.map((a, i) => `<tr><td>${i + 1}</td><td>${esc(a.nome)}</td><td>${esc(a.matricula)}</td><td>${esc(a.curso || '')}</td>
        <td>${a.presente ? 'Presente' : 'Ausente'}</td><td>${a.presente ? fmtDataHora(a.confirmadaEm) : '—'}</td></tr>`).join('')}
    </tbody></table>
    <p class="rodape">Documento gerado automaticamente pelo sistema EventCampus.</p>
  `);
}

// ════════════════════════════════════════════════════════════
// ADMIN — DASHBOARD
// ════════════════════════════════════════════════════════════
async function renderDashboard() {
  setPagina('dashboard', carregando());
  try {
    const s = await API.dashboard();
    const maxMes = Math.max(1, ...s.eventosPorMes.map(m => m.total));
    const barras = s.eventosPorMes.map(m => `
      <div style="display:flex;flex-direction:column;align-items:center;gap:6px;flex:1;">
        <span style="font-size:11px;color:var(--ink-mute);">${m.total}</span>
        <div style="width:100%;background:var(--lav-light);border-radius:6px;overflow:hidden;height:80px;display:flex;align-items:flex-end;">
          <div style="width:100%;height:${Math.round((m.total / maxMes) * 100)}%;background:linear-gradient(180deg,var(--lav),var(--lav-dark));border-radius:6px;"></div>
        </div>
        <span style="font-size:11px;color:var(--ink-mute);">${esc(m.mes)}</span>
      </div>`).join('');

    const totalCerts = s.certsAprovados + s.certsPendentes + s.certsReprovados;
    const dist = [
      { label: 'Aprovados', val: s.certsAprovados, cor: 'var(--mint)' },
      { label: 'Em análise', val: s.certsPendentes, cor: 'var(--sun)' },
      { label: 'Reprovados', val: s.certsReprovados, cor: 'var(--peach)' },
    ];

    setPagina('dashboard', `
      <div class="stat-grid">
        <div class="stat-card"><div class="stat-icon" style="background:var(--lav-light);color:var(--lav-dark);">${icone('graduation-cap', 22)}</div><div class="stat-info"><p>${s.alunosAtivos.toLocaleString('pt-BR')}</p><span>Alunos cadastrados</span></div></div>
        <div class="stat-card"><div class="stat-icon" style="background:var(--mint-light);color:var(--mint-dark);">${icone('calendar', 22)}</div><div class="stat-info"><p>${s.eventosAtivos}</p><span>Eventos ativos (${s.eventosTotal} no total)</span></div></div>
        <div class="stat-card"><div class="stat-icon" style="background:var(--sun-light);color:var(--sun-dark);">${icone('award', 22)}</div><div class="stat-info"><p>${s.certsFinaisEmitidos}</p><span>Certificados finais emitidos</span></div></div>
        <div class="stat-card"><div class="stat-icon" style="background:var(--peach-light);color:var(--peach);">${icone('clock', 22)}</div><div class="stat-info"><p>${s.horasRegistradas.toLocaleString('pt-BR')}h</p><span>Horas registradas</span></div></div>
      </div>

      <div class="grid-2">
        <div class="card">
          <p class="card-title">Eventos por mês (últimos 6 meses)</p>
          <div style="display:flex;align-items:flex-end;gap:8px;height:120px;">${barras}</div>
        </div>
        <div class="card">
          <p class="card-title">Certificados externos</p>
          ${dist.map(c => `
            <div style="margin-bottom:12px;">
              <div class="row-between mb-4"><span class="text-sm">${c.label}</span><span class="text-sm fw-600">${c.val} (${pct(c.val, totalCerts)}%)</span></div>
              <div class="progress-bar"><div class="progress-fill" style="width:${pct(c.val, totalCerts)}%;background:${c.cor};"></div></div>
            </div>`).join('')}
          <div class="mt-16" style="text-align:center;">
            ${s.certsPendentes
              ? `<button class="badge badge-red" style="font-size:13px;padding:6px 16px;cursor:pointer;" onclick="navegarPara('cert-externos')">${s.certsPendentes} pendente(s) — analisar</button>`
              : '<span class="badge badge-green" style="font-size:13px;padding:6px 16px;">Nenhum certificado pendente</span>'}
          </div>
        </div>
      </div>

      <div class="grid-3">
        <div class="card"><p class="text-mute">Professores</p><p style="font-size:22px;font-weight:700;">${s.professores}</p></div>
        <div class="card"><p class="text-mute">Inscrições realizadas</p><p style="font-size:22px;font-weight:700;">${s.inscricoesTotal}</p></div>
        <div class="card"><p class="text-mute">Alunos que atingiram a meta</p><p style="font-size:22px;font-weight:700;">${s.alunosConcluidos}</p></div>
      </div>
    `);
    atualizarBadgePendentes();
  } catch (e) {
    setPagina('dashboard', vazio('triangle-alert', e.message));
  }
}

// ════════════════════════════════════════════════════════════
// ADMIN — CERTIFICADOS EXTERNOS
// ════════════════════════════════════════════════════════════
let abaCert = 'PENDENTE';
let cacheCerts = [];

async function renderCertExternos() {
  setPagina('cert-externos', carregando());
  try {
    const todos = await API.certificados();
    const cont = s => todos.filter(c => c.status === s).length;
    cacheCerts = abaCert === 'TODOS' ? todos : todos.filter(c => c.status === abaCert);
    const aba = (s, txt, n) => `<button class="tab ${abaCert === s ? 'active' : ''}" onclick="abaCert='${s}';renderCertExternos()">${txt}<span class="count">${n}</span></button>`;
    const badge = s => s === 'APROVADO' ? '<span class="badge badge-green">Aprovado</span>'
      : s === 'REPROVADO' ? '<span class="badge badge-red">Reprovado</span>' : '<span class="badge badge-amber">Pendente</span>';

    setPagina('cert-externos', `
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:20px;font-weight:600;">Certificados Externos</h2>
          <p class="text-mute mt-4">Analise os certificados enviados pelos alunos. Os aprovados somam nas horas complementares.</p>
        </div>
        ${cont('PENDENTE') ? `<span class="badge badge-red" style="font-size:13px;padding:6px 16px;">${cont('PENDENTE')} pendente(s)</span>` : ''}
      </div>
      <div class="card" style="padding:0;">
        <div class="tabs">
          ${aba('PENDENTE', 'Pendentes', cont('PENDENTE'))}
          ${aba('APROVADO', 'Aprovados', cont('APROVADO'))}
          ${aba('REPROVADO', 'Reprovados', cont('REPROVADO'))}
          ${aba('TODOS', 'Todos', todos.length)}
        </div>
        <div style="padding:14px 20px;border-bottom:1px solid var(--border);">
          <input class="form-input search-input" placeholder="Buscar aluno ou curso..." style="max-width:320px;" oninput="filtrarTabela('cert-tabela', this.value)">
        </div>
        <div class="table-wrap" id="cert-tabela">
          ${cacheCerts.length ? `<table>
            <thead><tr><th>Aluno</th><th>Curso / Evento</th><th>Carga</th><th>Enviado em</th><th>Status</th><th>Ações</th></tr></thead>
            <tbody>${cacheCerts.map(c => `
              <tr>
                <td><p class="fw-600">${esc(c.aluno)}</p><p class="text-mute">${esc(c.matricula || '')} ${c.curso ? '· ' + esc(c.curso) : ''}</p></td>
                <td><p>${esc(c.tituloCurso)}</p><p class="text-mute">${esc(c.instituicao || '')}</p>
                  ${c.status === 'REPROVADO' && c.motivoReprovacao ? `<p class="text-mute" style="color:var(--red);">Motivo: ${esc(c.motivoReprovacao)}</p>` : ''}</td>
                <td>${c.cargaHoraria}h</td>
                <td>${fmtDataHora(c.enviadoEm)}</td>
                <td>${badge(c.status)}${c.analisadoPor ? `<p class="text-mute">por ${esc(c.analisadoPor)}</p>` : ''}</td>
                <td>
                  <div class="row" style="flex-wrap:nowrap;">
                    ${c.temArquivo ? `<button class="btn btn-ghost btn-sm" onclick="verAnexo(${c.id})" title="Ver arquivo enviado">${icone('paperclip', 14)} Ver</button>` : '<span class="text-mute" title="Sem anexo">—</span>'}
                    ${c.status !== 'APROVADO' ? `<button class="btn btn-primary btn-sm" onclick="modalAprovar(${c.id})">${icone('check', 14)} Aprovar</button>` : ''}
                    ${c.status !== 'REPROVADO' ? `<button class="btn btn-danger btn-sm" onclick="modalReprovar(${c.id})">${icone('x', 14)} Reprovar</button>` : ''}
                  </div>
                </td>
              </tr>`).join('')}
            </tbody></table>` : vazio('inbox', abaCert === 'PENDENTE' ? 'Nenhum certificado aguardando análise.' : 'Nenhum certificado aqui.')}
        </div>
      </div>
    `);
    atualizarBadgePendentes();
  } catch (e) {
    setPagina('cert-externos', vazio('triangle-alert', e.message));
  }
}

function modalAprovar(id) {
  const c = cacheCerts.find(x => x.id === id);
  abrirModal('Aprovar certificado', `
    <div class="form-error" id="modal-erro"></div>
    <p class="text-sm mb-16"><b>${esc(c.aluno)}</b> — ${esc(c.tituloCurso)} ${c.instituicao ? '(' + esc(c.instituicao) + ')' : ''}</p>
    <div class="form-group">
      <label class="form-label">Carga horária aceita (h)</label>
      <input class="form-input" id="aprov-carga" type="number" min="1" value="${c.cargaHoraria}" />
      <p class="form-hint">Você pode ajustar se o regulamento limitar as horas deste tipo de atividade.</p>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="fecharModal()">Cancelar</button>
      <button class="btn btn-primary" onclick="aprovarCert(${id}, this)">${icone('check', 14)} Aprovar</button>
    </div>`);
}

async function aprovarCert(id, btn) {
  const carga = parseInt(document.getElementById('aprov-carga').value, 10);
  if (!(carga > 0)) return erroModal('Informe uma carga horária válida.');
  await comLoading(btn, async () => {
    try {
      await API.aprovarCert(id, carga);
      fecharModal();
      toast('Certificado aprovado');
      renderCertExternos();
    } catch (e) { erroModal(e.message); }
  });
}

function modalReprovar(id) {
  const c = cacheCerts.find(x => x.id === id);
  abrirModal('Reprovar certificado', `
    <div class="form-error" id="modal-erro"></div>
    <p class="text-sm mb-16"><b>${esc(c.aluno)}</b> — ${esc(c.tituloCurso)}</p>
    <div class="form-group">
      <label class="form-label">Motivo (o aluno verá no app) *</label>
      <textarea class="form-input" id="reprov-motivo" placeholder="Ex: Certificado ilegível, envie novamente."></textarea>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="fecharModal()">Cancelar</button>
      <button class="btn btn-danger" onclick="reprovarCert(${id}, this)">${icone('x', 14)} Reprovar</button>
    </div>`);
}

async function reprovarCert(id, btn) {
  const motivo = document.getElementById('reprov-motivo').value.trim();
  if (!motivo) return erroModal('Informe o motivo da reprovação.');
  await comLoading(btn, async () => {
    try {
      await API.reprovarCert(id, motivo);
      fecharModal();
      toast('Certificado reprovado');
      renderCertExternos();
    } catch (e) { erroModal(e.message); }
  });
}

async function verAnexo(id) {
  try {
    const a = await API.arquivoCert(id);
    const bytes = Uint8Array.from(atob(a.base64), ch => ch.charCodeAt(0));
    const url = URL.createObjectURL(new Blob([bytes], { type: a.tipo }));
    const w = window.open(url, '_blank');
    if (!w) { // pop-up bloqueado → baixa
      const link = document.createElement('a');
      link.href = url; link.download = a.nome; link.click();
    }
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  } catch (e) { erro(e); }
}

// ════════════════════════════════════════════════════════════
// ADMIN — HORAS COMPLEMENTARES
// ════════════════════════════════════════════════════════════
let cacheHoras = [];
let filtroHoras = 'TODOS';

async function renderHoras() {
  setPagina('horas', carregando());
  try {
    cacheHoras = await API.horasAlunos();
    const lista = filtroHoras === 'CONCLUIDO' ? cacheHoras.filter(a => a.concluido)
      : filtroHoras === 'ANDAMENTO' ? cacheHoras.filter(a => !a.concluido) : cacheHoras;
    const aba = (s, txt, n) => `<button class="tab ${filtroHoras === s ? 'active' : ''}" onclick="filtroHoras='${s}';renderHoras()">${txt}<span class="count">${n}</span></button>`;

    setPagina('horas', `
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:20px;font-weight:600;">Horas Complementares</h2>
          <p class="text-mute mt-4">Horas = eventos com presença confirmada + certificados externos aprovados</p>
        </div>
        <div class="row">
          <input class="form-input search-input" placeholder="Buscar aluno..." style="max-width:240px;" oninput="filtrarTabela('tabela-horas', this.value)">
          <button class="btn btn-outline btn-sm" onclick="exportarHoras()">${icone('download', 14)} Excel</button>
        </div>
      </div>
      <div class="card" style="padding:0;">
        <div class="tabs">
          ${aba('TODOS', 'Todos', cacheHoras.length)}
          ${aba('ANDAMENTO', 'Em andamento', cacheHoras.filter(a => !a.concluido).length)}
          ${aba('CONCLUIDO', 'Meta atingida', cacheHoras.filter(a => a.concluido).length)}
        </div>
        <div class="table-wrap" id="tabela-horas">
          ${lista.length ? `<table>
            <thead><tr><th>Aluno</th><th>Curso</th><th>Eventos</th><th>Externas</th><th>Progresso</th><th>Situação</th><th>Ações</th></tr></thead>
            <tbody>${lista.map(a => `
              <tr>
                <td><p class="fw-600">${esc(a.nome)}</p><p class="text-mute">${esc(a.matricula)}</p></td>
                <td>${esc(a.curso || '—')}</td>
                <td>${a.horasEventos}h</td>
                <td>${a.horasExternas}h</td>
                <td>
                  <div style="min-width:140px;">
                    <div class="row-between mb-4"><span class="text-sm">${a.horasCompletas}h / ${a.horasTotal}h</span><span class="text-sm fw-600">${a.percentual}%</span></div>
                    <div class="progress-bar"><div class="progress-fill" style="width:${a.percentual}%;${a.concluido ? 'background:var(--mint);' : ''}"></div></div>
                  </div>
                </td>
                <td><span class="badge ${a.concluido ? 'badge-green' : 'badge-amber'}">${a.concluido ? 'Concluído' : `Faltam ${a.horasRestantes}h`}</span></td>
                <td>${a.certFinalEmitido
                  ? `<button class="btn btn-outline btn-sm" onclick="abrirCertFinal(${a.id})">${icone('award', 14)} Ver certificado</button>`
                  : a.concluido
                    ? `<button class="btn btn-primary btn-sm" onclick="abrirCertFinal(${a.id})">${icone('award', 14)} Emitir certificado</button>`
                    : '<span class="text-mute text-sm">Aguardando meta</span>'}</td>
              </tr>`).join('')}
            </tbody></table>` : vazio('graduation-cap', 'Nenhum aluno nesta lista.')}
        </div>
      </div>
    `);
  } catch (e) {
    setPagina('horas', vazio('triangle-alert', e.message));
  }
}

function exportarHoras() {
  baixarCSV('Horas complementares', ['Aluno', 'Matrícula', 'Curso', 'Horas em eventos', 'Horas externas', 'Total', 'Meta', '%', 'Situação', 'Certificado final'],
    cacheHoras.map(a => [a.nome, a.matricula, a.curso, a.horasEventos, a.horasExternas, a.horasCompletas, a.horasTotal,
      a.percentual + '%', a.concluido ? 'Concluído' : 'Em andamento', a.certFinalEmitido ? 'Emitido' : '']));
}

function abrirCertFinal(alunoId) {
  renderCertFinal.alunoSelecionado = alunoId;
  navegarPara('cert-final');
}

// ════════════════════════════════════════════════════════════
// ADMIN — CERTIFICADO FINAL
// ════════════════════════════════════════════════════════════
let certFinalAtual = null;

async function renderCertFinal() {
  setPagina('cert-final', carregando());
  try {
    if (!cacheHoras.length) cacheHoras = await API.horasAlunos();
    const aptos = cacheHoras.filter(a => a.concluido);

    setPagina('cert-final', `
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:20px;font-weight:600;">Certificado Final de Horas Complementares</h2>
          <p class="text-mute mt-4">Emitido para alunos que atingiram a meta de horas do curso</p>
        </div>
      </div>
      <div class="grid-2" style="align-items:start;">
        <div class="card">
          <p class="card-title">Alunos aptos (${aptos.length})</p>
          ${aptos.length ? aptos.map(a => `
            <div class="row-between" style="padding:10px 0;border-bottom:1px solid var(--border);">
              <div>
                <p class="fw-600 text-sm">${esc(a.nome)}</p>
                <p class="text-mute" style="font-size:11px;">${esc(a.curso || '')} · ${a.horasCompletas}h de ${a.horasTotal}h</p>
              </div>
              ${a.certFinalEmitido
                ? `<button class="btn btn-outline btn-sm" onclick="visualizarCertFinal(${a.id})">${icone('eye', 14)} Ver</button>`
                : `<button class="btn btn-primary btn-sm" onclick="emitirCertFinal(${a.id}, this)">${icone('award', 14)} Emitir</button>`}
            </div>`).join('') : vazio('hourglass', 'Nenhum aluno atingiu a meta ainda.')}
        </div>
        <div class="card" style="background:var(--lav-light);text-align:center;">
          <p class="card-title" style="text-align:left;">Pré-visualização</p>
          <div id="cert-visualizacao" style="padding:12px;">
            <p class="text-mute">Selecione um aluno para emitir ou visualizar o certificado</p>
          </div>
        </div>
      </div>
    `);

    const sel = renderCertFinal.alunoSelecionado;
    renderCertFinal.alunoSelecionado = null;
    if (sel) {
      const a = cacheHoras.find(x => x.id === sel);
      if (a && a.certFinalEmitido) visualizarCertFinal(sel);
      else if (a && a.concluido) emitirCertFinal(sel);
    }
  } catch (e) {
    setPagina('cert-final', vazio('triangle-alert', e.message));
  }
}

async function emitirCertFinal(alunoId, btn) {
  const a = cacheHoras.find(x => x.id === alunoId);
  if (!(await confirmar('Emitir certificado final',
    `Emitir o certificado de conclusão para <b>${esc(a ? a.nome : 'o aluno')}</b> com ${a ? a.horasCompletas : ''}h? Ele só pode ser emitido uma vez.`, 'Emitir'))) return;
  await comLoading(btn, async () => {
    try {
      certFinalAtual = await API.emitirFinal(alunoId);
      toast('Certificado final emitido!');
      cacheHoras = await API.horasAlunos();
      await renderCertFinal();
      mostrarCertFinal(certFinalAtual);
    } catch (e) { erro(e); }
  });
}

async function visualizarCertFinal(alunoId) {
  try {
    certFinalAtual = await API.verFinal(alunoId);
    mostrarCertFinal(certFinalAtual);
  } catch (e) { erro(e); }
}

function htmlCertificado(c, impressao = false) {
  const data = c.emitidoEm ? new Date(c.emitidoEm).toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' }) : '';
  const s = impressao ? 1.6 : 1;   // escala para o PDF
  return `
    <div style="border:${2 * s}px solid #8B7FD9;border-radius:${16 * s}px;padding:${32 * s}px;background:white;${impressao ? 'width:270mm;height:186mm;margin:12mm auto;display:flex;flex-direction:column;justify-content:center;' : 'max-width:400px;margin:0 auto;'}text-align:center;font-family:'Segoe UI',Arial,sans-serif;">
      <div style="background:#8B7FD9;height:${6 * s}px;border-radius:3px;margin-bottom:${24 * s}px;"></div>
      <p style="font-size:${11 * s}px;text-transform:uppercase;letter-spacing:.1em;color:#8580A3;margin-bottom:${8 * s}px;">Centro Universitário Alfredo Nasser · UNIFAN</p>
      <p style="font-size:${15 * s}px;font-weight:700;color:#5A4FB0;margin-bottom:${16 * s}px;">CERTIFICADO DE CONCLUSÃO DE HORAS COMPLEMENTARES</p>
      <p style="font-size:${11 * s}px;color:#8580A3;margin-bottom:${4 * s}px;">Certificamos que</p>
      <p style="font-size:${20 * s}px;font-weight:700;color:#2E2A45;margin-bottom:${4 * s}px;">${esc(c.aluno)}</p>
      <p style="font-size:${11 * s}px;color:#8580A3;margin-bottom:${16 * s}px;">${esc(c.curso || '')}${c.matricula ? ' · Matrícula ' + esc(c.matricula) : ''}</p>
      <p style="font-size:${12 * s}px;color:#2E2A45;line-height:1.6;margin-bottom:${20 * s}px;">
        concluiu a carga horária total de <strong>${c.horasTotais} horas</strong> de atividades complementares
        (${c.horasEventos}h em eventos acadêmicos e ${c.horasExternas}h em certificados externos),
        conforme exigido pela coordenação acadêmica.
      </p>
      <p style="font-size:${11 * s}px;color:#2E2A45;margin-bottom:${4 * s}px;">${esc(c.emitidoPor || 'Coordenação Acadêmica')}</p>
      <p style="font-size:${10 * s}px;color:#B3AFCB;">Emitido em ${data} · Código de validação: <b>${esc(c.codigoValidacao || '')}</b></p>
      <div style="background:#8B7FD9;height:${4 * s}px;border-radius:3px;margin-top:${20 * s}px;"></div>
    </div>`;
}

function mostrarCertFinal(c) {
  const el = document.getElementById('cert-visualizacao');
  if (!el) return;
  el.innerHTML = `
    ${htmlCertificado(c)}
    <div class="row mt-16" style="justify-content:center;gap:10px;">
      <button class="btn btn-primary btn-sm" onclick="baixarCertificadoPDF()">${icone('download', 14)} Baixar PDF</button>
    </div>`;
}

function baixarCertificadoPDF() {
  if (!certFinalAtual) return;
  imprimirHTML(`Certificado - ${certFinalAtual.aluno}`, htmlCertificado(certFinalAtual, true), { paisagem: true });
}

// ════════════════════════════════════════════════════════════
// ADMIN — USUÁRIOS
// ════════════════════════════════════════════════════════════
async function renderUsuarios() {
  setPagina('usuarios', carregando());
  try {
    const lista = await API.usuarios();
    const perfil = p => p === 'ADMIN' ? '<span class="badge badge-amber">Coordenação</span>'
      : p === 'PROFESSOR' ? '<span class="badge badge-blue">Professor</span>' : '<span class="badge badge-gray">Aluno</span>';
    setPagina('usuarios', `
      <div class="row-between mb-24">
        <div>
          <h2 style="font-size:20px;font-weight:600;">Usuários</h2>
          <p class="text-mute mt-4">Professores e coordenadores são cadastrados aqui. Alunos se cadastram pelo app.</p>
        </div>
        <button class="btn btn-primary" onclick="abrirFormUsuario()">${icone('plus')} Novo usuário</button>
      </div>
      <div class="card" style="padding:0;">
        <div style="padding:14px 20px;border-bottom:1px solid var(--border);">
          <input class="form-input search-input" placeholder="Buscar..." style="max-width:300px;" oninput="filtrarTabela('tabela-usuarios', this.value)">
        </div>
        <div class="table-wrap" id="tabela-usuarios">
          <table>
            <thead><tr><th>Nome</th><th>E-mail</th><th>Matrícula</th><th>Curso</th><th>Perfil</th></tr></thead>
            <tbody>${lista.map(u => `<tr>
              <td class="fw-600">${esc(u.nome)}</td><td>${esc(u.email)}</td><td>${esc(u.matricula || '—')}</td>
              <td>${esc(u.curso || '—')}</td><td>${perfil(u.perfil)}</td></tr>`).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `);
  } catch (e) {
    setPagina('usuarios', vazio('triangle-alert', e.message));
  }
}

function abrirFormUsuario() {
  abrirModal('Novo usuário', `
    <div class="form-error" id="modal-erro"></div>
    <div class="form-group"><label class="form-label">Nome completo *</label><input class="form-input" id="us-nome" /></div>
    <div class="form-group"><label class="form-label">E-mail *</label><input class="form-input" id="us-email" type="email" placeholder="nome@uni.edu.br" /></div>
    <div class="grid-2">
      <div class="form-group"><label class="form-label">Senha inicial *</label><input class="form-input" id="us-senha" type="text" value="123456" /></div>
      <div class="form-group"><label class="form-label">Perfil</label>
        <select class="form-select" id="us-perfil">
          <option value="PROFESSOR">Professor</option>
          <option value="ADMIN">Coordenação (admin)</option>
          <option value="ALUNO">Aluno</option>
        </select>
      </div>
    </div>
    <div class="grid-2">
      <div class="form-group"><label class="form-label">Matrícula / registro</label><input class="form-input" id="us-matricula" /></div>
      <div class="form-group"><label class="form-label">Curso</label><input class="form-input" id="us-curso" placeholder="Ex: Engenharia de Software" /></div>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="fecharModal()">Cancelar</button>
      <button class="btn btn-primary" onclick="salvarUsuario(this)">Cadastrar</button>
    </div>`);
}

async function salvarUsuario(btn) {
  const v = id => document.getElementById(id).value.trim();
  const dados = { nome: v('us-nome'), email: v('us-email'), senha: v('us-senha'), perfil: v('us-perfil'),
    matricula: v('us-matricula'), curso: v('us-curso') };
  if (!dados.nome || !dados.email || !dados.senha) return erroModal('Preencha nome, e-mail e senha.');
  if (dados.senha.length < 6) return erroModal('A senha deve ter pelo menos 6 caracteres.');
  await comLoading(btn, async () => {
    try {
      await API.criarUsuario(dados);
      fecharModal();
      toast('Usuário cadastrado!');
      renderUsuarios();
    } catch (e) { erroModal(e.message); }
  });
}