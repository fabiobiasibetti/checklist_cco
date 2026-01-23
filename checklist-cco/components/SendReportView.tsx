
import React, { useState, useEffect, useMemo } from 'react';
import { SharePointService } from '../services/sharepointService';
import { RouteDeparture, Task, User, RouteConfig } from '../types';
import { 
  TowerControl, Send, RefreshCw, Loader2, 
  CheckCircle2, AlertCircle, Clock, Filter,
  ChevronDown
} from 'lucide-react';

interface SummaryItem {
  id: string;
  operacao: string;
  timestamp: string;
  relativeTime: string;
  status: string;
  statusColor: string;
  details?: string;
}

const SendReportView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [isLoading, setIsLoading] = useState(true);
  const [departures, setDepartures] = useState<RouteDeparture[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [lastSync, setLastSync] = useState(new Date());

  const fetchAllData = async () => {
    setIsLoading(true);
    const token = (window as any).__access_token;
    if (!token) return;

    try {
      // Buscamos dados para o resumo
      const [depData, configs] = await Promise.all([
        SharePointService.getDepartures(token),
        SharePointService.getRouteConfigs(token, currentUser.email)
      ]);
      
      setDepartures(depData);
      setUserConfigs(configs);
      setLastSync(new Date());
    } catch (e) {
      console.error("Erro ao carregar resumo:", e);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchAllData();
    const interval = setInterval(fetchAllData, 60000); // Atualiza a cada minuto
    return () => clearInterval(interval);
  }, []);

  const getRelativeTime = (dateStr: string) => {
    if (!dateStr) return "há -- horas";
    const date = new Date(dateStr);
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffHours = Math.floor(diffMs / (1000 * 60 * 60));
    return `há ${diffHours} horas`;
  };

  // Lógica para processar a lista de SAÍDAS
  const departuresSummary = useMemo(() => {
    const ops = Array.from(new Set(userConfigs.map(c => c.operacao)));
    return ops.map(op => {
      const opRoutes = departures.filter(d => d.operacao === op);
      const lastRoute = opRoutes.length > 0 ? opRoutes[opRoutes.length - 1] : null;
      
      let status = "PREVISTO";
      let color = "bg-slate-300 text-slate-600";

      if (opRoutes.some(r => r.statusOp === 'Atrasada')) {
        status = "ATUALIZAR";
        color = "bg-blue-500 text-white";
      } else if (opRoutes.length > 0 && opRoutes.every(r => r.statusOp === 'OK')) {
        status = "OK";
        color = "bg-emerald-500 text-white";
      }

      return {
        id: op,
        operacao: op,
        timestamp: lastRoute?.createdAt || new Date().toISOString(),
        relativeTime: getRelativeTime(lastRoute?.createdAt || ""),
        status: status,
        statusColor: color
      };
    });
  }, [departures, userConfigs]);

  // Lógica para processar a lista de NÃO COLETAS (Simulada com base no status das plantas)
  const nonCollectionsSummary = useMemo(() => {
    const ops = Array.from(new Set(userConfigs.map(c => c.operacao)));
    return ops.map(op => {
      // Simulamos a contagem de não coletas baseado em rotas NOK ou dados pendentes
      const opRoutes = departures.filter(d => d.operacao === op);
      const nokCount = opRoutes.filter(r => r.statusGeral === 'NOK').length;
      
      return {
        id: op,
        operacao: op,
        timestamp: new Date().toISOString(),
        relativeTime: getRelativeTime(new Date().toISOString()),
        status: nokCount > 0 ? `${nokCount} NÃO COLETAS` : "TODOS COLETADOS",
        statusColor: nokCount > 0 ? "bg-red-500 text-white" : "bg-emerald-500 text-white"
      };
    });
  }, [departures, userConfigs]);

  if (isLoading && departures.length === 0) {
    return (
      <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4">
        <Loader2 size={48} className="animate-spin" />
        <p className="font-black text-xs uppercase tracking-widest">Sincronizando Resumo...</p>
      </div>
    );
  }

  return (
    <div className="h-full flex flex-col bg-[#F8FAFC] dark:bg-slate-950 p-4 font-sans overflow-hidden">
      {/* Header Estilo Print */}
      <div className="flex justify-between items-center mb-6 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-[#075985] text-white rounded-xl shadow-lg">
            <TowerControl size={24} />
          </div>
          <div>
            <h1 className="text-xl font-black text-[#075985] dark:text-sky-400 uppercase tracking-tight">
              Envio de Saídas e Não Coletas
            </h1>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
              <Clock size={12} /> Última atualização: {lastSync.toLocaleTimeString()}
            </p>
          </div>
        </div>
        <button onClick={fetchAllData} className="p-2 text-slate-400 hover:text-primary-600 transition-colors">
          <RefreshCw size={20} className={isLoading ? 'animate-spin' : ''} />
        </button>
      </div>

      {/* Grid Principal */}
      <div className="flex-1 grid grid-cols-1 md:grid-cols-2 gap-6 min-h-0">
        
        {/* Coluna Saídas */}
        <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col overflow-hidden">
          <div className="p-4 border-b dark:border-slate-800 flex justify-between items-center bg-slate-50/50 dark:bg-slate-800/50">
            <div className="flex items-center gap-2">
              <Filter size={16} className="text-slate-400" />
              <h2 className="font-black text-[#075985] dark:text-sky-400 uppercase tracking-widest text-sm">Saídas</h2>
            </div>
          </div>
          
          <div className="flex-1 overflow-y-auto p-4 space-y-3 scrollbar-thin">
            {departuresSummary.map(item => (
              <div key={item.id} className="flex justify-between items-center p-3 rounded-xl border border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-all">
                <div>
                  <h3 className="font-black text-slate-700 dark:text-slate-200 text-sm">{item.operacao}</h3>
                  <p className="text-[10px] text-slate-400 font-medium">{new Date(item.timestamp).toLocaleString()}</p>
                </div>
                <div className="text-right">
                  <p className="text-[10px] font-bold text-slate-500 mb-1">{item.relativeTime}</p>
                  <span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-tighter ${item.statusColor}`}>
                    {item.status}
                  </span>
                </div>
              </div>
            ))}
          </div>

          <div className="p-4 bg-slate-50 dark:bg-slate-800/80 border-t dark:border-slate-800 space-y-3">
             <div className="flex items-center gap-2">
                <select className="flex-1 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-2 text-xs font-bold outline-none appearance-none cursor-pointer">
                  <option>SELECIONAR OPERAÇÃO...</option>
                  {userConfigs.map(c => <option key={c.operacao}>{c.operacao}</option>)}
                </select>
                <button className="bg-[#075985] text-white px-6 py-2 rounded-lg font-black uppercase text-[10px] flex items-center gap-2 hover:bg-sky-800 transition-all">
                  <Send size={14} /> Enviar
                </button>
             </div>
             <div className="flex items-center justify-between">
                <label className="flex items-center gap-2 cursor-pointer group">
                  <span className="text-[10px] font-black text-slate-500 uppercase">Atualização?</span>
                  <div className="relative w-8 h-4 bg-slate-200 dark:bg-slate-700 rounded-full transition-colors group-hover:bg-slate-300">
                    <div className="absolute left-0.5 top-0.5 w-3 h-3 bg-white rounded-full shadow-sm"></div>
                  </div>
                </label>
             </div>
          </div>
        </div>

        {/* Coluna Não Coletas */}
        <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-sm flex flex-col overflow-hidden">
          <div className="p-4 border-b dark:border-slate-800 flex justify-center items-center bg-slate-50/50 dark:bg-slate-800/50">
            <h2 className="font-black text-[#075985] dark:text-sky-400 uppercase tracking-widest text-sm">Não Coletas</h2>
          </div>

          <div className="flex-1 overflow-y-auto p-4 space-y-3 scrollbar-thin">
            {nonCollectionsSummary.map(item => (
              <div key={item.id} className="flex justify-between items-center p-3 rounded-xl border border-slate-100 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800/50 transition-all">
                <div>
                  <h3 className="font-black text-slate-700 dark:text-slate-200 text-sm">{item.operacao}</h3>
                  <p className="text-[10px] text-slate-400 font-medium">{new Date(item.timestamp).toLocaleString()}</p>
                </div>
                <div className="text-right">
                  <p className="text-[10px] font-bold text-slate-500 mb-1">{item.relativeTime}</p>
                  <span className={`px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-tighter ${item.statusColor}`}>
                    {item.status}
                  </span>
                </div>
              </div>
            ))}
          </div>

          <div className="p-4 bg-slate-50 dark:bg-slate-800/80 border-t dark:border-slate-800 flex items-center gap-2">
            <select className="flex-1 bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 rounded-lg p-2 text-xs font-bold outline-none appearance-none cursor-pointer">
              <option>SELECIONAR OPERAÇÃO...</option>
              {userConfigs.map(c => <option key={c.operacao}>{c.operacao}</option>)}
            </select>
            <button className="bg-[#075985] text-white px-6 py-2 rounded-lg font-black uppercase text-[10px] flex items-center gap-2 hover:bg-sky-800 transition-all">
              <Send size={14} /> Enviar
            </button>
          </div>
        </div>

      </div>

      {/* Footer Status Bar */}
      <div className="mt-6 bg-white dark:bg-slate-900 p-2 rounded-xl border border-slate-200 dark:border-slate-800 flex items-center justify-center gap-4">
        <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Status Resumo</span>
        <span className="px-4 py-1 bg-red-500 text-white rounded-full text-[10px] font-black uppercase">Não Enviado</span>
      </div>
    </div>
  );
};

export default SendReportView;
