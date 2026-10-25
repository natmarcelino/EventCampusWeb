// js/api.js — comunicação com a API EventCampus (JWT no header Authorization)

const Sessao = {
  get token() { try { return localStorage.getItem('ec_token'); } catch { return Sessao._token || null; } },
  get usuario() {
    try { return JSON.parse(localStorage.getItem('ec_usuario') || 'null'); }
    catch { return Sessao._usuario || null; }
  },
  salvar(dados) {
    Sessao._token = dados.token; Sessao._usuario = dados;
    try {
      localStorage.setItem('ec_token', dados.token);
      localStorage.setItem('ec_usuario', JSON.stringify(dados));
    } catch {}
  },
  limpar() {
    Sessao._token = null; Sessao._usuario = null;
    try { localStorage.removeItem('ec_token'); localStorage.removeItem('ec_usuario'); } catch {}
  },
};

class ApiError extends Error {
  constructor(msg, status) { super(msg); this.status = status; }
}

async function api(metodo, rota, corpo) {
  const headers = { 'Content-Type': 'application/json' };
  const token = Sessao.token;
  if (token) headers.Authorization = 'Bearer ' + token;

  let resp;
  try {
    resp = await fetch(API_URL + rota, {
      method: metodo,
      headers,
      body: corpo !== undefined ? JSON.stringify(corpo) : undefined,
    });
  } catch (e) {
    throw new ApiError('Não foi possível conectar à API em ' + API_URL + '. Se ela estiver online, pode estar acordando: aguarde uns 40 segundos e tente de novo.', 0);
  }

  const texto = await resp.text();
  let dados = null;
  if (texto) { try { dados = JSON.parse(texto); } catch { dados = { erro: texto.slice(0, 200) }; } }

  if (resp.status === 401 && rota !== '/api/auth/login') {
    Sessao.limpar();
    if (typeof sessaoExpirada === 'function') sessaoExpirada();
    throw new ApiError((dados && dados.erro) || 'Sessão expirada. Entre novamente.', 401);
  }
  if (!resp.ok) {
    throw new ApiError((dados && (dados.erro || dados.message)) || ('Erro ' + resp.status), resp.status);
  }
  return dados;
}

const API = {
  // auth
  login: (email, senha) => api('POST', '/api/auth/login', { email, senha }),
  me:    ()             => api('GET', '/api/auth/me'),

  // eventos
  meusEventos:   ()            => api('GET', '/api/eventos/meus'),
  evento:        (id)          => api('GET', `/api/eventos/${id}`),
  criarEvento:   (dados)       => api('POST', '/api/eventos', dados),
  editarEvento:  (id, dados)   => api('PUT', `/api/eventos/${id}`, dados),
  statusEvento:  (id, status)  => api('PATCH', `/api/eventos/${id}/status`, { status }),
  excluirEvento: (id)          => api('DELETE', `/api/eventos/${id}`),
  qr:            (id)          => api('GET', `/api/eventos/${id}/qr`),
  ativarQR:      (id, cfg)     => api('PATCH', `/api/eventos/${id}/ativar-qr`, cfg || {}),
  encerrarQR:    (id)          => api('PATCH', `/api/eventos/${id}/encerrar-qr`),
  regenerarQR:   (id)          => api('PATCH', `/api/eventos/${id}/regenerar-qr`),

  // inscrições / presença
  inscritos:       (eventoId)        => api('GET', `/api/inscricoes/evento/${eventoId}`),
  presencaManual:  (inscId, confirmar) => api('PATCH', `/api/inscricoes/${inscId}/presenca`, { confirmar }),

  // relatórios
  relatorioEvento: (id) => api('GET', `/api/relatorios/evento/${id}`),

  // certificados (coordenação)
  certificados:  (status)       => api('GET', '/api/certificados' + (status ? `?status=${status}` : '')),
  aprovarCert:   (id, carga)    => api('PATCH', `/api/certificados/${id}/aprovar`, carga ? { cargaHoraria: carga } : {}),
  reprovarCert:  (id, motivo)   => api('PATCH', `/api/certificados/${id}/reprovar`, { motivo }),
  arquivoCert:   (id)           => api('GET', `/api/certificados/${id}/arquivo`),
  horasAlunos:   ()             => api('GET', '/api/certificados/horas-alunos'),
  emitirFinal:   (alunoId)      => api('POST', `/api/certificados/final/${alunoId}`),
  verFinal:      (alunoId)      => api('GET', `/api/certificados/final/${alunoId}`),

  // admin
  dashboard:     ()        => api('GET', '/api/admin/dashboard'),
  todosEventos:  ()        => api('GET', '/api/admin/eventos'),
  usuarios:      (perfil)  => api('GET', '/api/admin/usuarios' + (perfil ? `?perfil=${perfil}` : '')),
  criarUsuario:  (dados)   => api('POST', '/api/admin/usuarios', dados),
};