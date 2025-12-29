
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord, SPListInfo, SPColumnInfo } from '../types';

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
    let errorCode = "";
    try {
      const err = await res.json();
      errorCode = err.error?.code || "";
      errDetail = err.error?.message || JSON.stringify(err);
    } catch(e) {
      errDetail = await res.text();
    }
    
    console.error(`Graph API Error [${res.status}]:`, errDetail);
    
    if (res.status === 403) {
        throw new Error(`Acesso Negado (403): O App não tem permissão para escrever nesta lista ou o usuário não tem permissão de edição no SharePoint. Verifique os escopos Sites.ReadWrite.All.`);
    }
    if (res.status === 404) {
        throw new Error(`Não Encontrado (404): Verifique se a URL do site ou o nome da lista estão corretos.`);
    }
    
    throw new Error(`${errorCode}: ${errDetail}`);
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
    return await graphFetch(`/sites/${siteId}/lists/${listName}`, token);
  } catch (e) {
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
  async getAllLists(token: string): Promise<SPListInfo[]> {
    const siteId = await getResolvedSiteId(token);
    const data = await graphFetch(`/sites/${siteId}/lists`, token);
    return data.value.map((l: any) => ({
        id: l.id,
        displayName: l.displayName,
        name: l.name,
        webUrl: l.webUrl
    }));
  },

  async getListColumns(token: string, listId: string): Promise<SPColumnInfo[]> {
    const siteId = await getResolvedSiteId(token);
    const data = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
    return data.value.map((c: any) => ({
        name: c.name,
        displayName: c.displayName,
        description: c.description || "",
        type: c.text ? 'Text' : c.dateTime ? 'DateTime' : c.number ? 'Number' : c.choice ? 'Choice' : 'Other'
    }));
  },

  async getTasks(token: string): Promise<SPTask[]> {
    try {
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
    } catch (e) { return []; }
  },

  async getOperations(token: string, userEmail: string): Promise<SPOperation[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Operacoes_Checklist', token);
        const mapping = await getListColumnMapping(siteId, list.id, token);
        const colEmail = resolveFieldName(mapping, 'Email');
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
        return (data.value || [])
          .filter((item: any) => (item.fields[colEmail] || "").toLowerCase().trim() === userEmail.toLowerCase().trim())
          .map((item: any) => ({
            id: String(item.fields.id || item.id),
            Title: item.fields.Title || "OP",
            Ordem: Number(item.fields[resolveFieldName(mapping, 'Ordem')]) || 0,
            Email: item.fields[colEmail] || ""
          })).sort((a: any, b: any) => a.Ordem - b.Ordem);
    } catch (e) { return []; }
  },

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
        const filter = `fields/DataReferencia eq '${date}'`;
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        return (data.value || []).map((item: any) => ({
          id: item.id,
          DataReferencia: item.fields.DataReferencia,
          TarefaID: String(item.fields.TarefaID),
          OperacaoSigla: item.fields.OperacaoSigla,
          Status: item.fields.Status,
          Usuario: item.fields.Usuario,
          Title: item.fields.Title
        }));
    } catch (e) { return []; }
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
    try {
        const filter = `fields/Title eq '${status.Title}'`;
        const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        if (existing?.value?.length > 0) {
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${existing.value[0].id}/fields`, token, {
            method: 'PATCH',
            body: JSON.stringify(fields)
          });
        } else {
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
            method: 'POST',
            body: JSON.stringify({ fields })
          });
        }
    } catch (e) {
        throw e; // Repassa para o TaskManager tratar
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const listName = 'Historico_checklist_web';
    const list = await findListByIdOrName(siteId, listName, token);
    const mapping = await getListColumnMapping(siteId, list.id, token);

    const fields: any = {
      Title: record.resetBy, 
      [resolveFieldName(mapping, 'Data')]: record.timestamp,
      [resolveFieldName(mapping, 'DadosJSON')]: JSON.stringify(record.tasks),
      [resolveFieldName(mapping, 'Celula')]: record.email
    };

    await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
        method: 'POST',
        body: JSON.stringify({ fields })
    });
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
      const mapping = await getListColumnMapping(siteId, list.id, token);
      const colEmail = resolveFieldName(mapping, 'Celula');
      const colData = resolveFieldName(mapping, 'Data');
      const colJson = resolveFieldName(mapping, 'DadosJSON');
      
      const filter = `fields/${colEmail} eq '${userEmail}'`;
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
      
      return (data.value || []).map((item: any) => ({
        id: item.fields.id || item.id,
        timestamp: item.fields[colData],
        resetBy: item.fields.Title, 
        email: item.fields[colEmail],
        tasks: JSON.parse(item.fields[colJson] || '[]')
      })).sort((a: any, b: any) => new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime());
    } catch (e) { return []; }
  },

  async getRegisteredUsers(token: string, email: string): Promise<string[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
      const mapping = await getListColumnMapping(siteId, list.id, token);
      
      const colEmail = resolveFieldName(mapping, 'Email');
      const colNome = resolveFieldName(mapping, 'Nome');
      
      const filter = `fields/${colEmail} eq '${email}'`;
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
      
      return (data.value || []).map((item: any) => item.fields[colNome] || "").filter(Boolean);
    } catch (e) {
      return [];
    }
  }
};
