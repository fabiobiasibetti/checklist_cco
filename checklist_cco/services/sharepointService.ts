
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
    console.error(`Graph API Error [${res.status}] em ${url}:`, errDetail);
    if (res.status === 403) throw new Error("Acesso Negado: Verifique as permissões de EDIÇÃO na lista.");
    if (res.status === 400) throw new Error(`Requisição Inválida (400): ${errDetail}. Verifique se as colunas existem na lista.`);
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
    } catch (e) { 
        console.error("Erro ao buscar tarefas:", e);
        return []; 
    }
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
    } catch (e) { 
        console.error("Erro ao buscar operações:", e);
        return []; 
    }
  },

  async getStatusByDate(token: string, date: string): Promise<SPStatus[]> {
    try {
        const siteId = await getResolvedSiteId(token);
        const list = await findListByIdOrName(siteId, 'Status_Checklist', token);
        const mapping = await getListColumnMapping(siteId, list.id, token);
        const colData = resolveFieldName(mapping, 'DataReferencia');
        
        // Filtro por data. Nota: Algumas listas SP podem exigir que a coluna esteja indexada para filtros.
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
    } catch (e) { 
        console.error("Erro ao buscar status por data:", e);
        return []; 
    }
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
        // Busca se já existe um item com o mesmo Title (chave única lógica)
        const filter = `fields/Title eq '${status.Title}'`;
        const existing = await graphFetch(`/sites/${siteId}/lists/${list.id}/items?expand=fields&$filter=${filter}`, token);
        
        if (existing?.value && existing.value.length > 0) {
          // Se houver mais de um (duplicata acidental), atualizamos o primeiro e poderíamos deletar os outros se necessário.
          const itemId = existing.value[0].id;
          console.log(`Atualizando item existente ${itemId} para ${status.Status}`);
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items/${itemId}/fields`, token, {
            method: 'PATCH',
            body: JSON.stringify(fields)
          });
        } else {
          // Criar novo
          console.log(`Criando novo registro de status: ${status.Title}`);
          await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
            method: 'POST',
            body: JSON.stringify({ fields })
          });
        }
    } catch (e: any) {
        console.error("Falha ao atualizar/criar status no SharePoint:", e);
        // Fallback: Tenta criar se a busca falhou (ex: erro de filtro)
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

    try {
        await graphFetch(`/sites/${siteId}/lists/${list.id}/items`, token, {
          method: 'POST',
          body: JSON.stringify({ fields })
        });
    } catch (error: any) {
        console.error("Erro ao salvar histórico:", error);
        throw new Error(`Erro ao gravar na lista ${listName}: ${error.message}`);
    }
  },

  async getHistory(token: string, userEmail: string): Promise<HistoryRecord[]> {
    try {
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
    } catch (e) { 
        console.error("Erro ao buscar histórico:", e);
        return []; 
    }
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
      console.error("Erro ao buscar usuários cadastrados:", e);
      return [];
    }
  }
};
