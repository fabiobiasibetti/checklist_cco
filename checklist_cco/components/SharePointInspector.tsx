
import React, { useState, useEffect } from 'react';
import { SharePointService } from '../services/sharepointService';
import { SPListInfo, SPColumnInfo, User } from '../types';
import { Database, Search, Table, Info, RefreshCw, AlertCircle, ExternalLink, ChevronRight, ChevronDown } from 'lucide-react';

interface SharePointInspectorProps {
    currentUser: User;
}

const SharePointInspector: React.FC<SharePointInspectorProps> = ({ currentUser }) => {
    const [lists, setLists] = useState<SPListInfo[]>([]);
    const [selectedList, setSelectedList] = useState<SPListInfo | null>(null);
    const [columns, setColumns] = useState<SPColumnInfo[]>([]);
    const [isLoading, setIsLoading] = useState(true);
    const [isLoadingColumns, setIsLoadingColumns] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const loadLists = async () => {
        setIsLoading(true);
        setError(null);
        try {
            const data = await SharePointService.getAllLists(currentUser.accessToken!);
            setLists(data.sort((a, b) => a.displayName.localeCompare(b.displayName)));
        } catch (err: any) {
            setError(err.message);
        } finally {
            setIsLoading(false);
        }
    };

    const loadColumns = async (list: SPListInfo) => {
        setSelectedList(list);
        setIsLoadingColumns(true);
        try {
            const cols = await SharePointService.getListColumns(currentUser.accessToken!, list.id);
            setColumns(cols);
        } catch (err: any) {
            alert(`Erro ao carregar colunas: ${err.message}`);
        } finally {
            setIsLoadingColumns(false);
        }
    };

    useEffect(() => {
        if (currentUser.accessToken) loadLists();
    }, [currentUser]);

    return (
        <div className="flex flex-col h-full bg-slate-50 dark:bg-slate-950 rounded-2xl border dark:border-slate-800 shadow-sm overflow-hidden animate-fade-in">
            <div className="p-6 border-b dark:border-slate-800 bg-white dark:bg-slate-900 flex justify-between items-center">
                <div>
                    <h2 className="text-xl font-black text-slate-800 dark:text-white flex items-center gap-2">
                        <Database className="text-blue-600" />
                        SharePoint Inspector
                    </h2>
                    <p className="text-xs text-slate-500 font-bold uppercase tracking-widest mt-1">Diagnóstico de Estrutura e Permissões</p>
                </div>
                <button 
                    onClick={loadLists} 
                    className="p-2 bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 rounded-xl hover:bg-slate-200 transition-all flex items-center gap-2 text-xs font-bold"
                >
                    <RefreshCw size={16} className={isLoading ? 'animate-spin' : ''} />
                    Recarregar Site
                </button>
            </div>

            {error && (
                <div className="m-6 p-4 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-xl text-red-600 dark:text-red-400 flex items-start gap-3">
                    <AlertCircle className="shrink-0 mt-0.5" />
                    <div className="text-sm">
                        <p className="font-bold mb-1">Erro Crítico de Acesso</p>
                        <p className="opacity-80">{error}</p>
                        <p className="mt-3 text-[10px] uppercase font-black">Sugestão: Verifique se o App Registration no Azure AD possui a permissão 'Sites.ReadWrite.All' e se o usuário atual tem acesso ao site CCO.</p>
                    </div>
                </div>
            )}

            <div className="flex-1 flex overflow-hidden">
                {/* List of Lists */}
                <div className="w-1/3 border-r dark:border-slate-800 overflow-y-auto p-4 space-y-2 bg-white/50 dark:bg-slate-900/50">
                    <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4 px-2">Listas Encontradas</h3>
                    {isLoading ? (
                        <div className="flex flex-col items-center py-10 gap-3">
                            <RefreshCw className="animate-spin text-blue-500" />
                            <span className="text-[10px] text-slate-400 font-bold uppercase">Mapeando site...</span>
                        </div>
                    ) : lists.map(list => (
                        <button 
                            key={list.id}
                            onClick={() => loadColumns(list)}
                            className={`w-full text-left p-3 rounded-xl transition-all border flex flex-col gap-1 group
                                ${selectedList?.id === list.id 
                                    ? 'bg-blue-600 border-blue-500 text-white shadow-lg shadow-blue-500/20' 
                                    : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:border-blue-400 dark:hover:border-blue-800'
                                }
                            `}
                        >
                            <div className="flex items-center justify-between">
                                <span className="font-bold text-sm truncate">{list.displayName}</span>
                                <ChevronRight size={14} className={selectedList?.id === list.id ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'} />
                            </div>
                            <span className={`text-[9px] font-mono opacity-60 truncate ${selectedList?.id === list.id ? 'text-blue-100' : 'text-slate-500'}`}>
                                ID: {list.id}
                            </span>
                        </button>
                    ))}
                </div>

                {/* Column Detail */}
                <div className="flex-1 overflow-y-auto p-6 bg-white dark:bg-slate-900">
                    {!selectedList ? (
                        <div className="h-full flex flex-col items-center justify-center text-slate-400 opacity-30">
                            <Table size={64} className="mb-4" />
                            <p className="font-black uppercase text-sm">Selecione uma lista para inspecionar</p>
                        </div>
                    ) : (
                        <div className="space-y-6">
                            <div className="flex justify-between items-start">
                                <div>
                                    <h3 className="text-2xl font-black text-slate-800 dark:text-white">{selectedList.displayName}</h3>
                                    <a href={selectedList.webUrl} target="_blank" className="text-[10px] text-blue-500 font-bold flex items-center gap-1 hover:underline mt-1">
                                        ABRIR NO SHAREPOINT <ExternalLink size={10} />
                                    </a>
                                </div>
                                <div className="px-3 py-1 bg-slate-100 dark:bg-slate-800 rounded-lg text-[10px] font-mono text-slate-500">
                                    Graph Name: {selectedList.name}
                                </div>
                            </div>

                            <div className="bg-blue-50 dark:bg-blue-900/20 border border-blue-100 dark:border-blue-800/50 p-4 rounded-2xl flex items-start gap-3">
                                <Info size={18} className="text-blue-500 shrink-0 mt-0.5" />
                                <p className="text-xs text-blue-700 dark:text-blue-300 leading-relaxed font-medium">
                                    <b>Nota:</b> Use o <b>Internal Name</b> (campo 'Nome Técnico') para realizar filtros e atualizações via API. Se o nome contiver caracteres especiais ou espaços no SharePoint, o Graph API costuma converter para algo como <code>_x0020_</code>.
                                </p>
                            </div>

                            <div className="overflow-hidden border dark:border-slate-800 rounded-xl">
                                <table className="w-full text-left text-xs">
                                    <thead className="bg-slate-50 dark:bg-slate-800/50 text-slate-500 dark:text-slate-400 font-black uppercase tracking-widest text-[9px]">
                                        <tr>
                                            <th className="p-3 border-b dark:border-slate-800">Exibição</th>
                                            <th className="p-3 border-b dark:border-slate-800">Nome Técnico (API)</th>
                                            <th className="p-3 border-b dark:border-slate-800">Tipo</th>
                                        </tr>
                                    </thead>
                                    <tbody className="divide-y dark:divide-slate-800">
                                        {isLoadingColumns ? (
                                            <tr>
                                                <td colSpan={3} className="p-10 text-center">
                                                    <RefreshCw className="animate-spin inline-block mr-2" /> Carregando colunas...
                                                </td>
                                            </tr>
                                        ) : columns.map(col => (
                                            <tr key={col.name} className="hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors">
                                                <td className="p-3 font-bold text-slate-700 dark:text-slate-200">{col.displayName}</td>
                                                <td className="p-3 font-mono text-blue-600 dark:text-blue-400 bg-blue-50/30 dark:bg-blue-900/10">{col.name}</td>
                                                <td className="p-3">
                                                    <span className="px-2 py-0.5 bg-slate-100 dark:bg-slate-800 rounded text-[9px] font-bold text-slate-500">
                                                        {col.type}
                                                    </span>
                                                </td>
                                            </tr>
                                        ))}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    )}
                </div>
            </div>
        </div>
    );
};

export default SharePointInspector;
