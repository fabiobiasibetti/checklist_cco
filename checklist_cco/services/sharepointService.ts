
import { SPTask, SPOperation, SPStatus, Task, OperationStatus, HistoryRecord } from '../types';

const SITE_PATH = "vialacteoscombr.sharepoint.com:/sites/CCO";
let cachedSiteId: string | null = null;
const columnMappingCache: Record<string, Record<string, string>> = {};

export const getLocalDateString = () => {
  const date = new Date();
  const offset = date.getTimezoneOffset();
  const adjustedDate = new Date(date.getTime() - (offset * 60 * 1000));
  return adjustedDate.toISOString().split('T')[0];
};

async function graphFetch(endpoint: string, token: string, options: RequestInit = {}) {
  const url = endpoint.startsWith('https://') ? endpoint : `https://graph.microsoft.com/v1.0${endpoint}`;
  
  const res = await fetch(url, {
    ...options,
    headers: {
      ...options.headers,
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
      'Prefer': 'HonorNonIndexedQueriesWarningMayFailRandomly, HonorNonIndexedQueriesWarningMayFailOverLargeLists'
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
    console.error(`Graph API Error [${res.status}] at ${endpoint}:`, errDetail);
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
  throw new Error(`Lista '${listName}' não encontrada.`);
}

function normalizeString(str: string): string {
  if (!str) return "";
  return str.toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") 
    .replace(/[^a-z0-9]/g, "")       
    .trim();
}

const SYSTEM_FIELDS_TO_EXCLUDE = new Set([
  'id', 'linktitle', 'linktitlenomenu', 'modified', 'created', 'author', 'editor', 
  'attachments', 'contenttype', 'version', 'complianceassetid'
]);

async function getListColumnMapping(siteId: string, listId: string, token: string): Promise<Record<string, string>> {
  const cacheKey = `${siteId}_${listId}`;
  if (columnMappingCache[cacheKey]) return columnMappingCache[cacheKey];

  const columns = await graphFetch(`/sites/${siteId}/lists/${listId}/columns`, token);
  const mapping: Record<string, string> = {};
  
  columns.value.forEach((col: any) => {
    const internalName = col.name;
    const displayName = col.displayName;
    const normInternal = normalizeString(internalName);
    const normDisplay = normalizeString(displayName);
    if (SYSTEM_FIELDS_TO_EXCLUDE.has(normInternal)) return;
    if (!SYSTEM_FIELDS_TO_EXCLUDE.has(normDisplay)) {
      mapping[normDisplay] = internalName;
    }
    mapping[normInternal] = internalName;
  });

  columnMappingCache[cacheKey] = mapping;
  return mapping;
}

function resolveFieldName(mapping: Record<string, string>, target: string): string {
  const normalizedTarget = normalizeString(target);
  return mapping[normalizedTarget] || target;
}

export const SharePointService = {
  async getTasks(token: string): Promise<SPTask[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Tarefas_Checklist', token);
        const mapping = await getListColumnMapping(siteId, list.id, token);
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=999`, token);
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
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=999`, token);
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
        // Using $top 5000 and manual filter if the site filter fails due to indexing
        const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=5000`, token);
        return (data.value || [])
          .filter((item: any) => item.fields.DataReferencia === date)
          .map((item: any) => ({
            id: item.id,
            DataReferencia: item.fields.DataReferencia,
            TarefaID: String(item.fields.TarefaID),
            OperacaoSigla: item.fields.OperacaoSigla,
            Status: item.fields.Status,
            Usuario: item.fields.Usuario,
            Title: item.fields.Title
          }));
    } catch (e) { 
        console.error("Error fetching status by date:", e);
        return []; 
    }
  },

  async updateStatus(token: string, status: SPStatus, existingId?: string): Promise<string> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
    const mapping = await getListColumnMapping(siteId, list.id, token);
    
    const chaveUnicaField = resolveFieldName(mapping, 'ChaveUnica');
    const fields: any = {
      Title: status.Title,
      [chaveUnicaField]: status.Title,
      [resolveFieldName(mapping, 'DataReferencia')]: status.DataReferencia,
      [resolveFieldName(mapping, 'TarefaID')]: status.TarefaID,
      [resolveFieldName(mapping, 'OperacaoSigla')]: status.OperacaoSigla,
      [resolveFieldName(mapping, 'Status')]: status.Status,
      [resolveFieldName(mapping, 'Usuario')]: status.Usuario
    };

    if (existingId) {
        // Direct PATCH if we have the ID, no filter needed
        const patchFields = { ...fields };
        delete patchFields.Title;
        delete patchFields[chaveUnicaField];
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${existingId}/fields`, token, {
            method: 'PATCH',
            body: JSON.stringify(patchFields)
        });
        return existingId;
    } else {
        // Try to find by ChaveUnica or Title first to avoid duplicate if indexing allows, 
        // but if it fails, we catch and POST.
        try {
            const res = await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
                method: 'POST',
                body: JSON.stringify({ fields })
            });
            return res.id;
        } catch (e: any) {
            // If POST fails because it exists (and unique constraint active), we'd need the ID.
            // But usually we just throw here.
            throw e;
        }
    }
  },

  async saveHistory(token: string, record: HistoryRecord): Promise<void> {
    const siteId = await getResolvedSiteId(token);
    const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
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
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Historico_checklist_web', token);
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$top=500`, token);
      return (data.value || [])
        .filter((item: any) => (item.fields.Celula || "").toLowerCase() === userEmail.toLowerCase())
        .map((item: any) => ({
          id: item.id,
          timestamp: item.fields.Data,
          resetBy: item.fields.Title, 
          email: item.fields.Celula,
          tasks: JSON.parse(item.fields.DadosJSON || '[]')
        })).sort((a: any, b: any) => new Date(b.timestamp || 0).getTime() - new Date(a.timestamp || 0).getTime());
    } catch (e) { return []; }
  },

  async getRegisteredUsers(token: string, email: string): Promise<string[]> {
    try {
      const siteId = await getResolvedSiteId(token);
      const list = await findListByIdOrName(siteId, 'Usuarios_cco', token);
      const mapping = await getListColumnMapping(siteId, list.id, token);
      const colEmail = resolveFieldName(mapping, 'Email');
      const data = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields`, token);
      return (data.value || [])
        .filter((item: any) => (item.fields[colEmail] || "").toLowerCase() === email.toLowerCase())
        .map((item: any) => item.fields[resolveFieldName(mapping, 'Nome')] || "")
        .filter(Boolean);
    } catch (e) { return []; }
  }
};
