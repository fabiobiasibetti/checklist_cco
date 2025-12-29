
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord } from '../types';

const SITE_PATH = "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, Record<string, string>> = {};

async function graphFetch(endpoint: string, token: string, options: RequestInit = {}) {
  const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  
  const res = await fetch(url, {
    ...options,
    headers: {
      ...options.headers,
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Prefer': 'HonorNonIndexedQueriesWarningMayFailOverLargeLists'
    }
  });

  if (!res.ok) {
    let errDetail = "";
    try {
      const err = await res.json();
      errDetail = err.error?.message || JSON.stringify(err);
    } catch(e) {
      errDetail = await res.text();
    }
    
    const errorMsg = `Erro API SharePoint [${res.status}]: ${errDetail}`;
    console.error(errorMsg);
    
    if (res.status === 403) throw new Error("ACESSO NEGADO: Sua conta não tem permissão para ler ou editar esta lista no SharePoint.");
    if (res.status === 404) throw new Error(`NÃO ENCONTRADO: O recurso no caminho ${endpoint} não existe.`);
    throw new Error(errorMsg);
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
  try {
    // Tenta busca direta pelo nome/id
    return await graphFetch(`/sites/${siteId}/lists/${listName}`, token);
  } catch (e) {
    // Fallback: lista todas as listas e procura por displayName
    const data = await graphFetch(`/sites/${siteId}/lists`, token);
    const found = data.value.find((l: any) => 
      l.name?.toLowerCase() === listName.toLowerCase() || 
      l.displayName?.toLowerCase() === listName.toLowerCase()
    );
    if (found) return found;
  }
  throw new Error(`CRÍTICO: A lista '${listName}' não foi encontrada no site CCO. Verifique se o nome está correto no SharePoint.`);
}

function normalizeString(str: string): string {
  if (!str) return "";
  return str.toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") 
    .replace(/[^a-z0-9]/g, "")       
    .trim();
}

async function getListColumnMapping(siteId: string, listId: string, token: string): Promise<Record<string, string>> {
  const cacheKey = `${siteId}_${listId}`;
  if (columnMappingCache[cacheKey]) return columnMappingCache[cacheKey];

  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  
  columns.value.forEach((col: any) => {
    mapping[normalizeString(col.name)] = col.name;
    mapping[normalizeString(col.displayName)] = col.name;
  });

  columnMappingCache[cacheKey] = mapping;
  return mapping;
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  const normalizedTarget = normalizeString(target);
  if (mapping[normalizedTarget]) return mapping[normalizedTarget];
  return target;
}

export const SharePointService = {
  async validateConnection(token: string): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    // Verifica as 3 listas essenciais
    await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
    await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
    await findListByIdOrName(siteId, 'Status_Checklist', token);
  },

  async getTasks(token: string): Promise<SPTask[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
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
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    const colEmail = resolveFieldName(mapping, 'Email');
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
    
    const filtered = (data.value || [])
      .filter((item: any) => (item.fields[colEmail] || "").toLowerCase().trim() === userEmail.toLowerCase().trim());
    
    if (filtered.length === 0) {
        console.warn(`Nenhuma operação vinculada ao e-mail ${userEmail} na lista Operacoes_Checklist.`);
    }

    return filtered.map((item: any) => ({
        id: String(item.fields.id || item.id),
        Title: item.fields.Title || "OP",
        Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 0,
        Email: item.fields[colEmail] || ""
    })).sort((a: any, b: any) => a.Ordem - b.Ordem);
  },

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    const colData = resolveFieldName(mapping, 'DataReferencia');
    
    const filter = `fields/${colData} eq '${date}'`;
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    return (data.value || []).map((item: any) => ({
      id: item.id,
      DataReferencia: item.fields[colData],
      TarefaID: String(item.fields[resolveFieldName(mapping, 'TarefaID')]),
      OperacaoSigla: item.fields[resolveFieldName(mapping, 'OperacaoSigla')],
      Status: item.fields[resolveFieldName(mapping, 'Status')],
      Usuario: item.fields[resolveFieldName(mapping, 'Usuario')],
      Title: item.fields.Title
    }));
  },

  async updateStatus(token: string, status: SPStatus): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const fields = {
      Title: status.Title,
      [resolveFieldName(mapping, 'DataReferencia')]: status.DataReferencia,
      [resolveFieldName(mapping, 'TarefaID')]: status.TarefaID,
      [resolveFieldName(mapping, 'OperacaoSigla')]: status.OperacaoSigla,
      [resolveFieldName(mapping, 'Status')]: status.Status,
      [resolveFieldName(mapping, 'Usuario')]: status.Usuario
    };

    const filter = `fields/Title eq '${status.Title}'`;
    const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    if (existing?.value && existing.value.length > 0) {
      const itemId = existing.value[0].id;
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
        method: 'PATCH',
        body: JSON.stringify(fields)
      });
    } else {
      await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
        method: 'POST',
        body: JSON.stringify({ fields })
      });
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const listName = 'Historico_checklist_web';
    const list = await findListByIdOrName(siteId, listName, token);

    const fields: any = {
      Title: record.resetBy, 
      Data: record.timestamp,
      DadosJSON: JSON.stringify(record.tasks),
      Celula: record.email
    };

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
      method: 'POST',
      body: JSON.stringify({ fields })
    });
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
    const filter = `fields/Celula eq '${userEmail}'`;
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    return (data.value || []).map((item: any) => ({
      id: item.fields.id || item.id,
      timestamp: item.fields.Data,
      resetBy: item.fields.Title, 
      email: item.fields.Celula,
      tasks: JSON.parse(item.fields.DadosJSON || '[]')
    })).sort((a: any, b: any) => new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime());
  },

  async getRegisteredUsers(token: string, email: string): Promise<string[]> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const colEmail = resolveFieldName(mapping, 'Email');
    const colNome = resolveFieldName(mapping, 'Nome');
    
    const filter = `fields/${colEmail} eq '${email}'`;
    const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
    
    return (data.value || []).map((item: any) => item.fields[colNome] || "").filter(Boolean);
  }
};
