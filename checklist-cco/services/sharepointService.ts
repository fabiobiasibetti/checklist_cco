
// @google/genai guidelines: Use direct process.env.API_KEY, no UI for keys, use correct model names.
// Correct models: 'gemini-3-flash-preview', 'gemini-3-pro-preview', 'gemini-2.5-flash-image', etc.

import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord, RouteDeparture, RouteOperationMapping, RouteConfig } from '../types';

export interface DailyWarning {
  id: string;
  operacao: string; // Título
  celula: string;   // Email do responsável
  rota: string;
  descricao: string;
  dataOcorrencia: string; // ISO Date
  visualizado: boolean;
}

const SITE_PATH = "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, { mapping: Record<string, string>, readOnly: Set<string>, internalNames: Set<string> }> = {};

async function graphFetch(endpoint: string, token: string, options: RequestInit = {}) {
  const separator = endpoint.includes('?') ? '&' : '?';
  const url = endpoint.startsWith('https://') 
    ? endpoint 
    : `https://graph.microsoft.com/v1.0${endpoint}${options.method === 'GET' || !options.method ? `${separator}t=${Date.now()}` : ''}`;
    
  const headers: Record<string, string> = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
    'Prefer': 'HonorNonIndexedQueriesWarningMayFailOverLargeLists'
  };

  const res = await fetch(url, { ...options, headers: { ...headers, ...options.headers } });

  if (!res.ok) {
    let errDetail = "";
    try { 
        const err = await res.json(); 
        errDetail = err.error?.message || JSON.stringify(err); 
    } catch(e) { 
        errDetail = await res.text(); 
    }
    console.error(`[SHAREPOINT_API_FAILURE] URL: ${url} STATUS: ${res.status} ERROR: ${errDetail}`);
    throw new Error(errDetail);
  }
  return res.status === 204 ? null : res.json();
}

async function getResolvedSiteId(token: string): Promise<string> {
  if (cachedSiteId) return cachedSiteId;
  const siteData = await graphFetch(`/sites/${SITE_PATH}`, token);
  cachedSiteId = siteData.id;
  return siteData.id;
}

async function findListByIdOrName(siteId: string, listName: string, token: string): Promise<any> {
  try { return await graphFetch(`/sites/${siteId}/lists/${listName}`, token); } 
  catch (e) {
    const data = await graphFetch(`/sites/${siteId}/lists`, token);
    const found = data.value.find((l: any) => 
      l.name?.toLowerCase() === listName.toLowerCase() || 
      l.displayName?.toLowerCase() === listName.toLowerCase()
    );
    if (found) return found;
  }
  throw new Error(`Lista '${listName}' não encontrada.`);
}

function normalizeString(str: string): string {
  if (!str) return "";
  return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "").trim();
}

async function getListColumnMapping(siteId: string, listId: string, token: string) {
  const cacheKey = `${siteId}_${listId}`;
  if (columnMappingCache[cacheKey]) return columnMappingCache[cacheKey];
  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  const readOnly = new Set<string>();
  const internalNames = new Set<string>();
  
  columns.value.forEach((col: any) => {
    const internalName = col.name;
    mapping[normalizeString(col.name)] = internalName;
    mapping[normalizeString(col.displayName)] = internalName;
    internalNames.add(internalName);
    if (col.readOnly || internalName.startsWith('_') || ['ID', 'Author', 'Created'].includes(internalName)) {
        if (internalName !== 'Title') readOnly.add(internalName);
    }
  });
  
  columnMappingCache[cacheKey] = { mapping, readOnly, internalNames };
  return columnMappingCache[cacheKey];
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  const normalized = normalizeString(target);
  if (normalized === 'titulo' || normalized === 'rota') {
      if (mapping['title']) return 'Title';
  }
  return mapping[normalized] || target;
}

export const SharePointService = {
  // ... (métodos existentes mantidos)

  async createBulkDepartures(token: string, departures: RouteDeparture[]): Promise<{ success: string[], failed: string[] }> {
      const results = { success: [] as string[], failed: [] as string[] };
      for (const departure of departures) {
          try {
              const newId = await this.updateDeparture(token, departure);
              results.success.push(newId);
          } catch (e: any) {
              console.error(`Erro ao criar rota ${departure.rota}:`, e);
              results.failed.push(departure.rota);
          }
      }
      return results;
  },

  async getAllListsMetadata(token: string): Promise<any[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const listsToExplore = ['Tarefas_Checklist', 'Operacoes_Checklist', 'Status_Checklist', 'Historico_checklist_web', 'Usuarios_cco', 'CONFIG_SAIDA_DE_ROTAS', 'Dados_Saida_de_rotas', 'Rotas_Operacao_Checklist', 'avisos_diarios_checklist'];
      const results = await Promise.all(listsToExplore.map(async (listName) => {
        try {
          const list = await findListByIdOrName(siteId, listName, token);
          const columnsResponse = await graphFetch(`/sites/${siteId}/lists/${list.id}/columns`, token);
          return { list: { id: list.id, displayName: list.displayName, webUrl: list.webUrl }, columns: columnsResponse.value || [], error: null };
        } catch (e: any) {
          return { list: { displayName: listName, id: listName, webUrl: '#' }, columns: [], error: e.message };
        }
      }));
      return results;
    } catch (e) { return []; }
  },

  async getTasks(token: string): Promise<SPTask[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => ({
          id: String(item.fields.id || item.id),
          Title: item.fields.Title || "Sem Título",
          Descricao: item.fields[resolveFieldName(mapping, 'Descricao')] || "",
          Categoria: item.fields[resolveFieldName(mapping, 'Categoria')] || "Geral",
          Horario: item.fields[resolveFieldName(mapping, 'Horario')] || "--:--",
          Ativa: item.fields[resolveFieldName(mapping, 'Ativa')] !== false,
          Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 999
        })).sort((a: any, b: any) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const emailField = mapping['responsavel'] || 'Responsavel';
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || [])
          .map((item: any) => ({
            id: String(item.fields.id || item.id),
            Title: item.fields.Title || "OP",
            Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 0,
            Email: (item.fields[emailField] || "").toString().trim()
          }))
          .filter((op: SPOperation) => op.Email.toLowerCase() === userEmail.toLowerCase().trim())
          .sort((a: SPOperation, b: SPOperation) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async getTeamMembers(token: string): Promise<string[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => item.fields.Title).filter(Boolean).sort();
    } catch (e) { return ['Logística 1', 'Logística 2', 'Supervisor']; }
  },

  async getRegisteredUsers(token: string, _userEmail?: string): Promise<string[]> { return this.getTeamMembers(token); },

  async ensureMatrix(token: string, tasks: SPTask[], ops: SPOperation[]): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);
    const today = new Date().toISOString().split('T')[0];
    const colData = resolveFieldName(mapping, 'DataReferencia');
    const filter = `fields/${colData} ge '${today}T00:00:00Z' and fields/${colData} le '${today}T23:59:59Z'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=999`, token);
    const existingKeys = new Set((existing.value || []).map((i: any) => i.fields.Title));

    for (const task of tasks) {
      if (!task.Ativa) continue;
      for (const op of ops) {
        const uniqueKey = `${today.replace(/-/g, '')}_${task.id}_${op.Title}`;
        if (!existingKeys.has(uniqueKey)) {
          const rawFields: any = { Title: uniqueKey, ChaveUnica: uniqueKey, DataReferencia: today + 'T12:00:00Z', TarefaID: task.id, OperacaoSigla: op.Title, Status: 'PR', Usuario: 'Sistema' };
          const fields: any = {};
          Object.keys(rawFields).forEach(key => {
            const int = resolveFieldName(mapping, key);
            if (internalNames.has(int) && (!readOnly.has(int) || int === 'Title')) fields[int] = rawFields[key];
          });
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) }).catch(() => null);
        }
      }
    }
  },

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const colData = resolveFieldName(mapping, 'DataReferencia');
        const filter = `fields/${colData} ge '${date}T00:00:00Z' and fields/${colData} le '${date}T23:59:59Z'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}&$top=999`, token);
        return (data.value || []).map((item: any) => ({
          id: item.id, DataReferencia: item.fields[colData], TarefaID: String(item.fields[resolveFieldName(mapping, 'TarefaID')] || ""), OperacaoSigla: item.fields[resolveFieldName(mapping, 'OperacaoSigla')], Status: item.fields[resolveFieldName(mapping, 'Status')], Usuario: item.fields[resolveFieldName(mapping, 'Usuario')], Title: item.fields.Title
        }));
    } catch (e) { return []; }
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const { mapping, readOnly, internalNames } = await getListColumnMapping(siteId, list.id, token);
    const filter = `fields/Title eq '${status.Title}'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    const fields: any = {};
    if (!existing.value?.length) {
        const raw = { Title: status.Title, ChaveUnica: status.Title, DataReferencia: new Date(status.DataReferencia).toISOString(), TarefaID: status.TarefaID, OperacaoSigla: status.OperacaoSigla, Status: status.Status, Usuario: status.Usuario };
        Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (internalNames.has(int)) fields[int] = (raw as any)[k]; });
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
    } else {
        const raw = { Status: status.Status, Usuario: status.Usuario };
        Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (internalNames.has(int) && !readOnly.has(int)) fields[int] = (raw as any)[k]; });
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${existing.value[0].id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);
    const celulaInternalName = mapping['celula'] || 'celula';
    const raw = { Title: record.resetBy || 'Reset', Data: new Date(record.timestamp).toISOString(), DadosJSON: JSON.stringify(record.tasks) };
    const fields: any = {};
    Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (internalNames.has(int)) fields[int] = (raw as any)[k]; });
    fields[celulaInternalName] = record.email;
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);
      const celulaField = mapping['celula'] || 'celula';
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
      return (data.value || []).map((item: any) => ({ id: item.id, timestamp: item.fields.Data, resetBy: item.fields.Title, email: (item.fields[celulaField] || "").toString().trim(), tasks: JSON.parse(item.fields.DadosJSON || '[]') })).filter((record: HistoryRecord) => record.email?.toLowerCase() === userEmail.toLowerCase().trim()).sort((a: any, b: any) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
    } catch (e) { return []; }
  },

  // FIX: Explicitly return RouteConfig[] instead of any[] to satisfy TypeScript in the component.
  async getRouteConfigs(token: string, userEmail: string): Promise<RouteConfig[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'CONFIG_SAIDA_DE_ROTAS', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any): RouteConfig => { 
          const f = item.fields; 
          return { 
            operacao: String(f[resolveFieldName(mapping, 'OPERACAO')] || ""), 
            email: String(f[resolveFieldName(mapping, 'EMAIL')] || "").toString().toLowerCase().trim(), 
            tolerancia: String(f[resolveFieldName(mapping, 'TOLERANCIA')] || "00:00:00") 
          }; 
        }).filter(c => c.email === userEmail.toLowerCase().trim());
    } catch (e) { return []; }
  },

  async getRouteOperationMappings(token: string): Promise<RouteOperationMapping[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Rotas_Operacao_Checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => ({ id: item.id, Title: item.fields.Title, OPERACAO: item.fields[resolveFieldName(mapping, 'OPERACAO')] }));
    } catch (e) { return []; }
  },

  async addRouteOperationMapping(token: string, routeName: string, operation: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Rotas_Operacao_Checklist', token);
    const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);
    const fields: any = { Title: routeName, [resolveFieldName(mapping, 'OPERACAO')]: operation };
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
  },

  async getDepartures(token: string): Promise<RouteDeparture[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
      const { mapping } = await getListColumnMapping(siteId, list.id, token);
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
      return (data.value || []).map((item: any) => {
        const f = item.fields;
        return {
          id: String(item.id), semana: f[resolveFieldName(mapping, 'Semana')] || "", rota: f.Title || "", data: f[resolveFieldName(mapping, 'DataOperacao')] ? f[resolveFieldName(mapping, 'DataOperacao')].split('T')[0] : "", inicio: f[resolveFieldName(mapping, 'HorarioInicio')] || "00:00:00", motorista: f[resolveFieldName(mapping, 'Motorista')] || "", placa: f[resolveFieldName(mapping, 'Placa')] || "", saida: f[resolveFieldName(mapping, 'HorarioSaida')] || "00:00:00", motivo: f[resolveFieldName(mapping, 'MotivoAtraso')] || "", observacao: f[resolveFieldName(mapping, 'Observacao')] || "", statusGeral: f[resolveFieldName(mapping, 'StatusGeral')] || "OK", aviso: f[resolveFieldName(mapping, 'Aviso')] || "NÃO", operacao: f[resolveFieldName(mapping, 'Operacao')] || "", statusOp: f[resolveFieldName(mapping, 'StatusOp')] || "OK", tempo: f[resolveFieldName(mapping, 'TempoGap')] || "OK", createdAt: f.Created || new Date().toISOString()
        };
      });
    } catch (e) { return []; }
  },

  async getArchivedDepartures(token: string, operation: string, startDate: string, endDate: string): Promise<RouteDeparture[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
      const { mapping } = await getListColumnMapping(siteId, historyListId, token);
      
      const colData = resolveFieldName(mapping, 'DataOperacao');
      const colOp = resolveFieldName(mapping, 'Operacao');
      
      let filter = `fields/${colData} ge '${startDate}T00:00:00Z' and fields/${colData} le '${endDate}T23:59:59Z'`;
      if (operation) {
          filter += ` and fields/${colOp} eq '${operation}'`;
      }

      console.log(`[ARCHIVE_QUERY] Filter: ${filter}`);
      const data = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items?expand=fields&$filter=${filter}&$top=999`, token);
      
      return (data.value || []).map((item: any) => {
        const f = item.fields;
        return {
          id: String(item.id), semana: f[resolveFieldName(mapping, 'Semana')] || "", rota: f.Title || "", data: f[colData] ? f[colData].split('T')[0] : "", inicio: f[resolveFieldName(mapping, 'HorarioInicio')] || "00:00:00", motorista: f[resolveFieldName(mapping, 'Motorista')] || "", placa: f[resolveFieldName(mapping, 'Placa')] || "", saida: f[resolveFieldName(mapping, 'HorarioSaida')] || "00:00:00", motivo: f[resolveFieldName(mapping, 'MotivoAtraso')] || "", observacao: f[resolveFieldName(mapping, 'Observacao')] || "", statusGeral: f[resolveFieldName(mapping, 'StatusGeral')] || "OK", aviso: f[resolveFieldName(mapping, 'Aviso')] || "NÃO", operacao: f[colOp] || "", statusOp: f[resolveFieldName(mapping, 'StatusOp')] || "OK", tempo: f[resolveFieldName(mapping, 'TempoGap')] || "OK", createdAt: f.Created || new Date().toISOString()
        };
      });
    } catch (e: any) {
        console.error("[ARCHIVE_FETCH_ERROR]", e.message);
        return [];
    }
  },

  async updateDeparture(token: string, departure: RouteDeparture): Promise<string> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    const { mapping, internalNames, readOnly } = await getListColumnMapping(siteId, list.id, token);
    const raw: any = { Title: departure.rota, Semana: departure.semana, DataOperacao: departure.data ? new Date(departure.data + 'T12:00:00Z').toISOString() : null, HorarioInicio: departure.inicio, Motorista: departure.motorista, Placa: departure.placa, HorarioSaida: departure.saida, MotivoAtraso: departure.motivo, Observacao: departure.observacao, StatusGeral: departure.statusGeral, Aviso: departure.aviso, Operacao: departure.operacao, StatusOp: departure.statusOp, TempoGap: departure.tempo };
    const fields: any = {};
    Object.keys(raw).forEach(k => { const int = resolveFieldName(mapping, k); if (int === 'Title' || (internalNames.has(int) && !readOnly.has(int))) { fields[int] = raw[k]; } });
    const isUpdate = departure.id && departure.id !== "" && departure.id !== "0" && !isNaN(Number(departure.id));
    if (isUpdate) {
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${departure.id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
      return departure.id;
    } else {
      const res = await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { method: 'POST', body: JSON.stringify({ fields }) });
      return String(res.id);
    }
  },

  async deleteDeparture(token: string, id: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${id}`, token, { method: 'DELETE' });
  },

  async moveDeparturesToHistory(token: string, items: RouteDeparture[]): Promise<{ success: number, failed: number, lastError?: string }> {
    console.log(`[ARCHIVE_START] Iniciando migração de ${items.length} itens.`);
    const siteId = await getResolvedSiteId(token);
    const sourceList = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
    const { mapping: histMapping, internalNames: histInternals } = await getListColumnMapping(siteId, historyListId, token);
    
    let successCount = 0;
    let failedCount = 0;
    let lastErrorMessage = "";

    for (const item of items) {
        try {
            const raw: any = { Title: item.rota, Semana: item.semana, DataOperacao: item.data ? new Date(item.data + 'T12:00:00Z').toISOString() : null, HorarioInicio: item.inicio, Motorista: item.motorista, Placa: item.placa, HorarioSaida: item.saida, MotivoAtraso: item.motivo, Observacao: item.observacao, StatusGeral: item.statusGeral, Aviso: item.aviso, Operacao: item.operacao, StatusOp: item.statusOp, TempoGap: item.tempo };
            const histFields: any = {};
            Object.keys(raw).forEach(k => { const int = resolveFieldName(histMapping, k); if (histInternals.has(int)) histFields[int] = raw[k]; });
            const postRes = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, { method: 'POST', body: JSON.stringify({ fields: histFields }) });
            if (postRes && postRes.id) {
                await graphFetch(`/sites/${siteId}/lists/${sourceList.id}/items/${item.id}`, token, { method: 'DELETE' });
                successCount++;
            } else { failedCount++; lastErrorMessage = "Sem confirmação ID."; }
        } catch (err: any) { failedCount++; lastErrorMessage = err.message; }
    }
    return { success: successCount, failed: failedCount, lastError: lastErrorMessage };
  },

  // MÉTODOS PARA AVISOS DIÁRIOS
  async addDailyWarning(token: string, warning: Omit<DailyWarning, 'id' | 'visualizado'>): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'avisos_diarios_checklist', token);
    const { mapping, internalNames } = await getListColumnMapping(siteId, list.id, token);
    
    // Mapeamento manual de segurança e conversão para strings (coluna Texto no SharePoint)
    // FIX: Usando warning.dataOcorrencia + T12:00:00Z para evitar erro de fuso
    const raw: any = {
        Title: warning.operacao || 'SEM OPERACAO',
        celula: warning.celula,
        rota: warning.rota,
        descricao: warning.descricao,
        data_referencia: new Date(warning.dataOcorrencia + 'T12:00:00Z').toISOString(),
        visualizado: "false" 
    };

    const fields: any = {};
    Object.keys(raw).forEach(k => {
        const int = resolveFieldName(mapping, k);
        if (internalNames.has(int)) {
            fields[int] = raw[k];
        } else if (internalNames.has(k)) {
            fields[k] = raw[k];
        }
    });

    if (!fields['Title']) fields['Title'] = raw.Title;

    console.log('[DEBUG WARNING] Final Payload Fields:', JSON.stringify(fields, null, 2));

    try {
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, { 
            method: 'POST', 
            body: JSON.stringify({ fields }) 
        });
    } catch (error: any) {
        console.error('[DEBUG ERROR] Critical failure saving warning:', error.message || error);
        throw error;
    }
  },

  async getDailyWarnings(token: string, userEmail: string): Promise<DailyWarning[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'avisos_diarios_checklist', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        
        const celulaCol = resolveFieldName(mapping, 'celula');
        const visualizadoCol = resolveFieldName(mapping, 'visualizado');
        const rotaCol = resolveFieldName(mapping, 'rota');
        const descCol = resolveFieldName(mapping, 'descricao');
        const dataCol = resolveFieldName(mapping, 'data_referencia');

        // FIX: Mantemos o filtro apenas em visualizado e celula para garantir que avisos não lidos apareçam sempre
        const filter = `fields/${celulaCol} eq '${userEmail.trim()}' and fields/${visualizadoCol} eq 'false'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        
        return (data.value || []).map((item: any) => {
            const f = item.fields;
            return {
                id: String(item.id),
                operacao: f.Title || "",
                celula: f[celulaCol] || "",
                rota: f[rotaCol] || "",
                descricao: f[descCol] || "",
                dataOcorrencia: f[dataCol] || "",
                visualizado: f[visualizadoCol] === 'true'
            };
        });
    } catch (e) {
        console.error("Erro ao carregar avisos:", e);
        return [];
    }
  },

  async markWarningAsViewed(token: string, id: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'avisos_diarios_checklist', token);
    const { mapping } = await getListColumnMapping(siteId, list.id, token);
    const visualizadoCol = resolveFieldName(mapping, 'visualizado');
    
    const fields: any = { [visualizadoCol]: "true" }; 
    await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${id}/fields`, token, { method: 'PATCH', body: JSON.stringify(fields) });
  }
};
