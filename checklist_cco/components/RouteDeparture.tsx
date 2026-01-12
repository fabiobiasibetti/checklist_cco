
import React, { useState, useEffect, useMemo } from 'react';
import { RouteDeparture, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDepartures } from '../services/geminiService';
import { Plus, Trash2, Save, Clock, Maximize2, Minimize2, X, Upload, Sparkles, Loader2, RefreshCw, AlertTriangle, ShieldCheck } from 'lucide-react';

interface RouteConfig {
    operacao: string;
    email: string;
    tolerancia: string;
}

const RouteDepartureView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importText, setImportText] = useState('');
  const [currentTime, setCurrentTime] = useState(new Date());
  
  const getAccessToken = () => (window as any).__access_token;

  // Form State
  const [formData, setFormData] = useState<Partial<RouteDeparture>>({
    rota: '',
    data: new Date().toISOString().split('T')[0],
    inicio: '00:00:00',
    saida: '00:00:00',
    motorista: '',
    placa: '',
    operacao: '',
    motivo: '',
    observacao: '',
    statusGeral: 'OK',
    aviso: 'NÃO',
  });

  const loadData = async () => {
    const token = getAccessToken();
    if (!token || !currentUser) return;
    setIsLoading(true);
    try {
      const configs = await SharePointService.getRouteConfigs(token, currentUser.email);
      setUserConfigs(configs);

      const spData = await SharePointService.getDepartures(token);
      
      // Filtrar rotas baseadas nas operações que o usuário tem permissão na lista de config
      const allowedOps = new Set(configs.map(c => c.operacao.toUpperCase().trim()));
      const filtered = spData.filter(r => allowedOps.has(r.operacao.toUpperCase().trim()));
      
      setRoutes(filtered);
    } catch (e) {
      console.error("Falha ao carregar rotas sincronizadas:", e);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadData();
    const timer = setInterval(() => setCurrentTime(new Date()), 30000); // Atualiza a cada 30s
    return () => clearInterval(timer);
  }, [currentUser]);

  const timeToSeconds = (timeStr: string): number => {
    if (!timeStr || !timeStr.includes(':')) return 0;
    const parts = timeStr.split(':').map(Number);
    const h = parts[0] || 0;
    const m = parts[1] || 0;
    const s = parts[2] || 0;
    return (h * 3600) + (m * 60) + s;
  };

  const secondsToTime = (totalSeconds: number): string => {
    const isNegative = totalSeconds < 0;
    const absSeconds = Math.abs(totalSeconds);
    const h = Math.floor(absSeconds / 3600);
    const m = Math.floor((absSeconds % 3600) / 60);
    const s = absSeconds % 60;
    const formatted = [h, m, s].map(v => v.toString().padStart(2, '0')).join(':');
    return isNegative ? `-${formatted}` : formatted;
  };

  const calculateGap = (inicio: string, saida: string, toleranceStr: string = "00:00:00"): { gap: string, status: string, isOutOfTolerance: boolean } => {
    if (!inicio || !saida || inicio === '00:00:00' || saida === '00:00:00') return { gap: 'OK', status: 'OK', isOutOfTolerance: false };
    
    const startSec = timeToSeconds(inicio);
    const endSec = timeToSeconds(saida);
    const diff = endSec - startSec;
    const toleranceSec = timeToSeconds(toleranceStr);
    
    if (diff === 0) return { gap: 'OK', status: 'OK', isOutOfTolerance: false };
    
    const gapFormatted = secondsToTime(diff);
    const isOutOfTolerance = Math.abs(diff) > toleranceSec;
    
    let status = 'OK';
    if (isOutOfTolerance) {
        status = diff > 0 ? 'Atrasado' : 'Adiantado';
    }
    
    return { gap: gapFormatted, status, isOutOfTolerance };
  };

  const calculateWeekString = (dateStr: string) => {
    if (!dateStr) return '';
    const date = new Date(dateStr + 'T12:00:00');
    const monthNames = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
    const month = monthNames[date.getMonth()];
    const day = date.getDate();
    const firstDayOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
    const firstDayWeekday = firstDayOfMonth.getDay();
    const adjustedFirstDayWeekday = firstDayWeekday === 0 ? 6 : firstDayWeekday - 1;
    const weekNum = Math.ceil((day + adjustedFirstDayWeekday) / 7);
    return `${month} S${weekNum}`;
  };

  const openModal = () => {
    setFormData({
      rota: '',
      data: new Date().toISOString().split('T')[0],
      inicio: '00:00:00',
      saida: '00:00:00',
      motorista: '',
      placa: '',
      operacao: userConfigs.length > 0 ? userConfigs[0].operacao : '',
      motivo: '',
      observacao: '',
      statusGeral: 'OK',
      aviso: 'NÃO',
    });
    setIsModalOpen(true);
  };

  const closeModal = () => setIsModalOpen(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = getAccessToken();
    if (!token) return;

    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === formData.operacao?.toUpperCase().trim());
    const { gap, status, isOutOfTolerance } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia);

    if (isOutOfTolerance && (!formData.motivo || !formData.observacao)) {
        alert("Atenção: Para rotas fora da tolerância, o Motivo e a Observação são obrigatórios.");
        return;
    }

    setIsSyncing(true);
    try {
        const week = calculateWeekString(formData.data || '');
        const newRoute: RouteDeparture = {
            ...formData as RouteDeparture,
            id: '', 
            semana: week,
            tempo: gap,
            statusOp: status,
            createdAt: new Date().toISOString()
        };

        const newId = await SharePointService.updateDeparture(token, newRoute);
        setRoutes(prev => [...prev, { ...newRoute, id: newId }]);
        closeModal();
    } catch (err: any) {
        alert(`ERRO AO SALVAR: ${err.message}`);
    } finally {
        setIsSyncing(false);
    }
  };

  const removeRow = async (id: string) => {
    const token = getAccessToken();
    if (!token) return;
    if (confirm('Excluir permanentemente da lista do SharePoint?')) {
      setIsSyncing(true);
      try {
        await SharePointService.deleteDeparture(token, id);
        setRoutes(routes.filter(r => r.id !== id));
      } catch (err: any) {
          alert(`Erro ao excluir: ${err.message}`);
      } finally {
          setIsSyncing(false);
      }
    }
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    const token = getAccessToken();
    if (!token) return;

    const route = routes.find(r => r.id === id);
    if (!route) return;

    let updatedRoute = { ...route, [field]: value };
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === updatedRoute.operacao.toUpperCase().trim());
    
    if (field === 'inicio' || field === 'saida') {
        const { gap, status } = calculateGap(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia);
        updatedRoute.tempo = gap;
        updatedRoute.statusOp = status;
    }

    if (field === 'data') {
        updatedRoute.semana = calculateWeekString(value);
    }

    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));

    setIsSyncing(true);
    try {
      await SharePointService.updateDeparture(token, updatedRoute);
    } catch (err: any) {
      console.error("Erro na sincronização:", err.message);
    } finally {
      setIsSyncing(false);
    }
  };

  const getRowStyle = (route: RouteDeparture) => {
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === route.operacao.toUpperCase().trim());
    const toleranceSec = timeToSeconds(config?.tolerancia || "00:00:00");
    const { isOutOfTolerance } = calculateGap(route.inicio, route.saida, config?.tolerancia);

    // Laranja: Fora da tolerância com Saída Real
    if (route.saida !== '00:00:00' && isOutOfTolerance) {
        return 'bg-orange-500 text-white font-bold border-orange-600';
    }

    // Amarelo: Início já passou + tolerância e NÃO tem saída real ainda
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) {
        return 'bg-yellow-400 text-slate-900 border-yellow-500';
    }

    return 'bg-white dark:bg-slate-900 border-gray-100 dark:border-slate-800';
  };

  const getStatusBadgeStyle = (val: string, isWhiteText: boolean) => {
    if (val === 'Atrasado' || val === 'Adiantado') {
        return isWhiteText ? 'bg-white/20 text-white font-black' : 'bg-red-500 text-white font-black';
    }
    return 'bg-green-500 text-white font-black';
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-xs uppercase tracking-widest text-center">
            SINCRONIZANDO CONFIGURAÇÕES...<br/>
            <span className="text-[10px] font-normal italic">Filtrando por {currentUser.email}</span>
        </p>
    </div>
  );

  return (
    <div className={`flex flex-col h-full animate-fade-in ${isFullscreen ? 'fixed inset-0 z-[60] bg-white dark:bg-slate-950 p-4' : ''}`}>
      <div className="flex justify-between items-center mb-4 shrink-0">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-600 text-white rounded-2xl shadow-lg">
            <Clock size={24} />
          </div>
          <div>
            <h2 className="text-xl font-bold text-gray-800 dark:text-white flex items-center gap-2">
              Saída de Rotas
              {isSyncing && <Loader2 size={16} className="animate-spin text-blue-500 ml-2"/>}
            </h2>
            <div className="flex items-center gap-2">
                <p className="text-[10px] text-gray-500 dark:text-gray-400 font-bold uppercase tracking-tighter">
                    Filtrado por: {currentUser.email}
                </p>
                <div className="w-1 h-1 rounded-full bg-slate-300"></div>
                <p className="text-[10px] text-blue-600 font-black uppercase tracking-tighter">
                    {currentTime.toLocaleTimeString('pt-BR')}
                </p>
            </div>
          </div>
        </div>
        <div className="flex gap-2">
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-blue-500 transition-colors" title="Sincronizar">
              <RefreshCw size={18} />
          </button>
          <button 
            onClick={() => setIsImportModalOpen(true)}
            className="flex items-center gap-2 px-3 py-2 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 rounded-lg hover:bg-emerald-100 border border-emerald-100 dark:border-emerald-800 shadow-sm"
          >
            <Upload size={18} />
            <span className="text-xs font-bold hidden sm:inline">Importar Excel</span>
          </button>
          <button 
            onClick={() => setIsFullscreen(!isFullscreen)}
            className={`flex items-center gap-2 px-3 py-2 rounded-lg transition-all border shadow-sm ${isFullscreen ? 'bg-blue-600 text-white border-blue-500' : 'bg-gray-100 dark:bg-slate-800 text-gray-600 dark:text-gray-300 border-gray-200 dark:border-slate-700'}`}
          >
            {isFullscreen ? <Minimize2 size={18} /> : <Maximize2 size={18} />}
          </button>
          <button 
            onClick={openModal}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg transition-all shadow-md active:scale-95"
          >
            <Plus size={18} />
            Nova Rota
          </button>
        </div>
      </div>

      <div className="flex-1 overflow-auto bg-white dark:bg-slate-900 rounded-xl border border-gray-200 dark:border-slate-800 shadow-sm relative scrollbar-thin">
        <table className={`w-full border-collapse text-[10px] ${isFullscreen ? 'min-w-full' : 'min-w-[1500px]'}`}>
          <thead className="sticky top-0 z-20 bg-blue-900 dark:bg-blue-950 text-white uppercase font-black tracking-widest text-center shadow-lg h-10">
            <tr>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-20">Semana</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Rota</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-28">Data</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Início</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 min-w-[150px]">Motorista</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Placa</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Saída</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-32">Motivo</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 min-w-[200px]">Obs</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-16">Geral</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-16">Av</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-40">Operação</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Status Op</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Tempo</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-10 sticky right-0 bg-blue-900 dark:bg-blue-950">#</th>
            </tr>
          </thead>
          <tbody>
            {routes.length === 0 && (
              <tr>
                <td colSpan={15} className="p-16 text-center text-slate-400 font-black uppercase italic text-xs tracking-widest">
                   {userConfigs.length === 0 
                     ? "Seu e-mail não possui permissão em nenhuma operação (CONFIG_SAIDA_DE_ROTAS)" 
                     : "Nenhuma rota cadastrada para suas operações autorizadas."}
                </td>
              </tr>
            )}
            {routes.map((route) => {
              const rowStyle = getRowStyle(route);
              const isWhiteText = rowStyle.includes('text-white');
              
              return (
                <tr key={route.id} className={`hover:brightness-95 transition-all border-b h-11 ${rowStyle}`}>
                  <td className="p-0 border-r border-gray-200/20 text-center font-black">{route.semana}</td>
                  <td className="p-0 border-r border-gray-200/20">
                    <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none font-black text-center" />
                  </td>
                  <td className="p-0 border-r border-gray-200/20">
                    <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center" />
                  </td>
                  <td className="p-0 border-r border-gray-200/20">
                    <input type="text" value={route.inicio} onChange={(e) => updateCell(route.id, 'inicio', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" placeholder="00:00:00" />
                  </td>
                  <td className="p-0 border-r border-gray-200/20 px-2 uppercase truncate font-bold">{route.motorista}</td>
                  <td className="p-0 border-r border-gray-200/20 text-center font-black uppercase">{route.placa}</td>
                  <td className="p-0 border-r border-gray-200/20">
                    <input type="text" value={route.saida} onChange={(e) => updateCell(route.id, 'saida', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" placeholder="00:00:00" />
                  </td>
                  <td className="p-0 border-r border-gray-200/20">
                    <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none cursor-pointer text-center font-bold">
                      <option value="">Nenhum</option>
                      <option value="Manutenção">Manutenção</option>
                      <option value="Mão de obra">Mão de obra</option>
                      <option value="Atraso coleta">Atraso coleta</option>
                      <option value="Atraso carregamento">Atraso carregamento</option>
                      <option value="Logística">Logística</option>
                      <option value="Outros">Outros</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-gray-200/20">
                    <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none font-bold italic" />
                  </td>
                  <td className="p-0 border-r border-gray-200/20 text-center">
                     <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none text-center font-black">
                      <option value="OK">OK</option>
                      <option value="NOK">NOK</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-gray-200/20 text-center">
                    <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none text-center font-black">
                      <option value="SIM">SIM</option>
                      <option value="NÃO">NÃO</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-gray-200/20 text-center font-black uppercase tracking-tight">{route.operacao}</td>
                  <td className="p-0 border-r border-gray-200/20 text-center">
                    <div className={`w-full h-full flex items-center justify-center font-black ${getStatusBadgeStyle(route.statusOp, isWhiteText)}`}>
                      {route.statusOp}
                    </div>
                  </td>
                  <td className="p-0 border-r border-gray-200/20 text-center font-mono font-black">{route.tempo}</td>
                  <td className="p-1 sticky right-0 bg-inherit text-center">
                    <button onClick={() => removeRow(route.id)} className={`${isWhiteText ? 'text-white hover:text-red-200' : 'text-slate-300 hover:text-red-500'} transition-colors p-1`} title="Excluir">
                      <Trash2 size={14} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* MODALS */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
             <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border dark:border-slate-700">
                <div className="bg-emerald-800 dark:bg-slate-800 text-white p-4 flex justify-between items-center">
                    <h3 className="font-bold flex items-center gap-2"><Upload size={20} /> Importar Excel</h3>
                    <button onClick={() => setIsImportModalOpen(false)}><X size={24} /></button>
                </div>
                <div className="p-6">
                    <textarea value={importText} onChange={e => setImportText(e.target.value)} className="w-full h-64 p-4 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-xs font-mono dark:text-white mb-4" placeholder="Cole as linhas do Excel aqui..."/>
                    <div className="flex gap-3">
                        <button onClick={() => setIsImportModalOpen(false)} className="flex-1 py-3 bg-gray-200 dark:bg-slate-700 font-bold rounded-xl">Cancelar</button>
                        <button onClick={() => {}} disabled className="flex-[2] py-3 bg-emerald-600 text-white font-bold rounded-xl flex items-center justify-center gap-2 opacity-50 cursor-not-allowed">
                             Em breve: Importação com Gemini
                        </button>
                    </div>
                </div>
             </div>
        </div>
      )}

      {isModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden border dark:border-slate-700">
            <div className="bg-blue-600 dark:bg-slate-800 text-white p-6 flex justify-between items-center">
                <h3 className="font-black uppercase tracking-widest text-sm flex items-center gap-3"><Plus size={20} /> Registrar Rota</h3>
                <button onClick={closeModal} className="hover:bg-white/10 p-1 rounded-full"><X size={24} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-8 space-y-6 bg-gray-50 dark:bg-slate-900">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-bold shadow-sm"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Rota</label>
                        <input type="text" required placeholder="Número" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-black text-blue-600 shadow-sm"/>
                    </div>
                </div>
                <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Operação</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-black shadow-sm outline-none">
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Motorista</label>
                        <input type="text" required placeholder="Nome" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm shadow-sm"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-black shadow-sm"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Início Previsto</label>
                        <input type="text" required placeholder="00:00:00" value={formData.inicio} onChange={e => setFormData({...formData, inicio: e.target.value})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-mono shadow-sm"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Saída Real</label>
                        <input type="text" placeholder="00:00:00" value={formData.saida} onChange={e => setFormData({...formData, saida: e.target.value})} className="w-full p-3 border dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-mono shadow-sm"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-slate-900 hover:bg-slate-800 text-white font-black uppercase tracking-widest text-xs rounded-2xl flex items-center justify-center gap-3 shadow-xl transition-all active:scale-95 border-b-4 border-slate-700">
                    {isSyncing ? <Loader2 size={18} className="animate-spin" /> : <Save size={18} />} GRAVAR NO SHAREPOINT
                </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
