
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord, RouteDeparture, RouteOperationMapping } from '../types';

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
    // Log crítico para debug em produção
    console.error(`[GRAPH_ERROR] Falha na chamada API!
      Endpoint: ${endpoint}
      Status: ${res.status}
      Detalhe do SharePoint: ${errDetail}
    `);
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
  throw new Error(`Lista '${listName}' não encontrada no site.`);
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
  if (normalized === 'titulo' || normalized === 'rota') return 'Title';
  return mapping[normalized] || target;
}

export const SharePointService = {
  async getAllListsMetadata(token: string): Promise<any[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const listsToExplore = ['Tarefas_Checklist', 'Operacoes_Checklist', 'Status_Checklist', 'Historico_checklist_web', 'Usuarios_cco', 'CONFIG_SAIDA_DE_ROTAS', 'Dados_Saida_de_rotas', 'Rotas_Operacao_Checklist'];
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

  async getRouteConfigs(token: string, userEmail: string): Promise<any[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'CONFIG_SAIDA_DE_ROTAS', token);
        const { mapping } = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || []).map((item: any) => { const f = item.fields; return { operacao: f[resolveFieldName(mapping, 'OPERACAO')] || "", email: (f[resolveFieldName(mapping, 'EMAIL')] || "").toString().toLowerCase().trim(), tolerancia: f[resolveFieldName(mapping, 'TOLERANCIA')] || "00:00:00" }; }).filter(c => c.email === userEmail.toLowerCase().trim());
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

  async moveDeparturesToHistory(token: string, items: RouteDeparture[]): Promise<{ success: number, failed: number }> {
    console.log(`[SP_SERVICE] Iniciando arquivamento atômico de ${items.length} itens.`);
    const siteId = await getResolvedSiteId(token);
    const sourceList = await findListByIdOrName(siteId, 'Dados_Saida_de_rotas', token);
    
    // GUID LIMPO: Sem encoded chars
    const historyListId = "856bf9d5-6081-4360-bcad-e771cbabfda8";
    
    const { mapping: histMapping, internalNames: histInternals } = await getListColumnMapping(siteId, historyListId, token);
    
    let successCount = 0;
    let failedCount = 0;

    for (const item of items) {
        try {
            console.log(`[SP_SERVICE] Preparando item: ${item.rota}`);
            
            const raw: any = {
                Title: item.rota,
                Semana: item.semana,
                DataOperacao: item.data ? new Date(item.data + 'T12:00:00Z').toISOString() : null,
                HorarioInicio: item.inicio,
                Motorista: item.motorista,
                Placa: item.placa,
                HorarioSaida: item.saida,
                MotivoAtraso: item.motivo,
                Observacao: item.observacao,
                StatusGeral: item.statusGeral,
                Aviso: item.aviso,
                Operacao: item.operacao,
                StatusOp: item.statusOp,
                TempoGap: item.tempo
            };

            const histFields: any = {};
            Object.keys(raw).forEach(k => {
                const int = resolveFieldName(histMapping, k);
                if (histInternals.has(int)) histFields[int] = raw[k];
            });

            console.log('[DEBUG PAYLOAD]', JSON.stringify(histFields));

            const postRes = await graphFetch(`/sites/${siteId}/lists/${historyListId}/items`, token, {
                method: 'POST',
                body: JSON.stringify({ fields: histFields })
            });

            if (postRes && postRes.id) {
                console.log(`[SP_SERVICE] POST concluído. Deletando original ID ${item.id}...`);
                await graphFetch(`/sites/${siteId}/lists/${sourceList.id}/items/${item.id}`, token, {
                    method: 'DELETE'
                });
                successCount++;
            } else {
                failedCount++;
            }
        } catch (err: any) {
            console.error(`[SP_SERVICE] Erro no item ${item.rota}:`, err.message);
            failedCount++;
        }
    }

    console.log(`[SP_SERVICE] Arquivamento Finalizado. S: ${successCount}, F: ${failedCount}`);
    return { success: successCount, failed: failedCount };
  }
};
