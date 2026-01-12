
import React, { useState, useEffect } from 'react';
import { RouteDeparture, User } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDepartures, parseRouteDeparturesManual } from '../services/geminiService';
// Added ShieldCheck to the imports
import { Plus, Trash2, Save, Clock, Maximize2, Minimize2, X, Upload, Sparkles, Loader2, RefreshCw, AlertTriangle, CheckCircle2, ShieldCheck, FileSpreadsheet } from 'lucide-react';

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
      const allowedOps = new Set(configs.map(c => c.operacao.toUpperCase().trim()));
      const filtered = spData.filter(r => allowedOps.has(r.operacao.toUpperCase().trim()));
      setRoutes(filtered);
    } catch (e) {
      console.error(e);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadData();
    const timer = setInterval(() => setCurrentTime(new Date()), 10000);
    return () => clearInterval(timer);
  }, [currentUser]);

  const timeToSeconds = (timeStr: string): number => {
    if (!timeStr || !timeStr.includes(':')) return 0;
    const parts = timeStr.split(':').map(Number);
    return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0);
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
    const gapFormatted = secondsToTime(diff);
    const isOutOfTolerance = Math.abs(diff) > toleranceSec;
    const status = isOutOfTolerance ? (diff > 0 ? 'Atrasado' : 'Adiantado') : 'OK';
    return { gap: gapFormatted, status, isOutOfTolerance };
  };

  const calculateWeekString = (dateStr: string) => {
    if (!dateStr || dateStr === '') return '';
    try {
        const date = new Date(dateStr + 'T12:00:00');
        if (isNaN(date.getTime())) return '';
        const monthNames = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"];
        const month = monthNames[date.getMonth()];
        const day = date.getDate();
        const weekNum = Math.ceil(day / 7);
        return `${month} S${weekNum}`;
    } catch(e) { return ''; }
  };

  const handleImport = async (useAI: boolean = false) => {
    const token = getAccessToken();
    if (!importText.trim() || !token) return;
    setIsProcessingImport(true);
    try {
        let parsed: Partial<RouteDeparture>[] = [];
        
        if (useAI) {
            parsed = await parseRouteDepartures(importText);
        } else {
            parsed = parseRouteDeparturesManual(importText);
        }

        if (parsed.length === 0) {
            throw new Error("Nenhum dado válido identificado. Verifique se copiou as colunas ROTA e DATA corretamente.");
        }

        const importPromises = parsed.map(p => {
            // Atribuição automática da operação
            let op = p.operacao?.toUpperCase().trim();
            if (!op || op === '') {
                op = userConfigs.length > 0 ? userConfigs[0].operacao.toUpperCase().trim() : '';
            }

            const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === op);
            const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia);
            
            const r: RouteDeparture = {
                rota: p.rota || '',
                data: p.data || '', // Já vem em YYYY-MM-DD do parser robusto
                inicio: p.inicio || '00:00:00',
                motorista: p.motorista || '',
                placa: p.placa || '',
                saida: p.saida || '00:00:00',
                motivo: p.motivo || '',
                observacao: p.observacao || '',
                operacao: op,
                id: '',
                semana: calculateWeekString(p.data || ''),
                statusGeral: 'OK',
                aviso: 'NÃO',
                statusOp: status,
                tempo: gap,
                createdAt: new Date().toISOString()
            } as RouteDeparture;

            return SharePointService.updateDeparture(token, r);
        });

        await Promise.all(importPromises);
        await loadData();
        setIsImportModalOpen(false);
        setImportText('');
        alert(`Sucesso! ${parsed.length} rotas importadas.`);
    } catch (error: any) {
        alert(`Erro na importação: ${error.message}`);
    } finally {
        setIsProcessingImport(false);
    }
  };

  const removeRow = async (id: string) => {
    const token = getAccessToken();
    if (!token) return;
    if (confirm('Deseja excluir permanentemente?')) {
      setIsSyncing(true);
      try {
        await SharePointService.deleteDeparture(token, id);
        setRoutes(routes.filter(r => r.id !== id));
      } catch (err: any) {
          alert(err.message);
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

    if (field === 'data') updatedRoute.semana = calculateWeekString(value);

    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));
    setIsSyncing(true);
    try {
      await SharePointService.updateDeparture(token, updatedRoute);
    } catch (err: any) {
      console.error(err);
    } finally {
      setIsSyncing(false);
    }
  };

  const getRowStyle = (route: RouteDeparture) => {
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === route.operacao.toUpperCase().trim());
    const toleranceSec = timeToSeconds(config?.tolerancia || "00:00:00");
    const { isOutOfTolerance } = calculateGap(route.inicio, route.saida, config?.tolerancia);

    if (route.saida !== '00:00:00' && isOutOfTolerance) {
        return 'bg-orange-600 text-white font-bold border-orange-700 shadow-inner';
    }

    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) {
        return 'bg-yellow-400 text-slate-900 font-bold border-yellow-500';
    }

    return 'bg-white dark:bg-slate-900 border-gray-100 dark:border-slate-800';
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = getAccessToken();
    if (!token) return;

    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === formData.operacao?.toUpperCase().trim());
    const { status, isOutOfTolerance } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia);

    if (isOutOfTolerance && (!formData.motivo || !formData.observacao)) {
        alert("Atenção: Motivo e Observação são OBRIGATÓRIOS para rotas fora da tolerância.");
        return;
    }

    setIsSyncing(true);
    try {
        const week = calculateWeekString(formData.data || '');
        const { gap } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia);
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
        setIsModalOpen(false);
    } catch (err: any) {
        alert(err.message);
    } finally {
        setIsSyncing(false);
    }
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-xs uppercase tracking-widest text-center">Conectando ao SharePoint CCO...</p>
    </div>
  );

  return (
    <div className={`flex flex-col h-full animate-fade-in ${isFullscreen ? 'fixed inset-0 z-[60] bg-white dark:bg-slate-950 p-4' : ''}`}>
      <div className="flex justify-between items-center mb-4 shrink-0">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-700 text-white rounded-2xl shadow-lg ring-4 ring-blue-700/10">
            <Clock size={24} />
          </div>
          <div>
            <h2 className="text-xl font-black text-slate-800 dark:text-white flex items-center gap-2 uppercase tracking-tight">
              Saída de Rotas
              {isSyncing && <Loader2 size={16} className="animate-spin text-blue-500 ml-2"/>}
            </h2>
            <div className="flex items-center gap-2">
                <ShieldCheck size={12} className="text-emerald-500"/>
                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-tighter">Sincronizado via {currentUser.email}</p>
            </div>
          </div>
        </div>
        <div className="flex gap-2">
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-blue-500 transition-colors" title="Atualizar">
              <RefreshCw size={18} />
          </button>
          <button 
            onClick={() => setIsImportModalOpen(true)}
            className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-xl hover:bg-emerald-700 font-bold shadow-md transition-all active:scale-95 border-b-4 border-emerald-800"
          >
            <Upload size={18} />
            <span className="text-xs uppercase tracking-widest">Importar Excel</span>
          </button>
          <button 
            onClick={() => setIsModalOpen(true)}
            className="flex items-center gap-2 px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl transition-all shadow-md active:scale-95 border-b-4 border-blue-800"
          >
            <Plus size={18} />
            <span className="text-xs uppercase tracking-widest">Nova Rota</span>
          </button>
        </div>
      </div>

      <div className="flex-1 overflow-auto bg-slate-100 dark:bg-slate-900 rounded-2xl border dark:border-slate-800 shadow-xl relative scrollbar-thin">
        <table className={`w-full border-collapse text-[10px] ${isFullscreen ? 'min-w-full' : 'min-w-[1600px]'}`}>
          <thead className="sticky top-0 z-20 bg-blue-900 dark:bg-slate-950 text-white uppercase font-black tracking-widest text-center shadow-lg h-12">
            <tr>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-20">Semana</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Rota</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-28">Data</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Início</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 min-w-[180px]">Motorista</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Placa</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Saída</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-36">Motivo</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 min-w-[250px]">Observação</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-16">Geral</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-16">Av</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-40">Operação</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Status</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-24">Tempo</th>
              <th className="p-2 border border-blue-800 dark:border-blue-900 w-10 sticky right-0 bg-blue-900 dark:bg-slate-950">#</th>
            </tr>
          </thead>
          <tbody>
            {routes.map((route) => {
              const rowStyle = getRowStyle(route);
              const isWhiteText = rowStyle.includes('text-white');
              
              return (
                <tr key={route.id} className={`hover:brightness-95 transition-all border-b h-12 ${rowStyle}`}>
                  <td className="p-0 border-r border-slate-200/10 text-center font-black">{route.semana}</td>
                  <td className="p-0 border-r border-slate-200/10">
                    <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className="w-full h-full p-2 bg-transparent outline-none font-black text-center" />
                  </td>
                  <td className="p-0 border-r border-slate-200/10">
                    <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className="w-full h-full p-2 bg-transparent outline-none text-center" />
                  </td>
                  <td className="p-0 border-r border-slate-200/10">
                    <input type="text" value={route.inicio} onChange={(e) => updateCell(route.id, 'inicio', e.target.value)} className="w-full h-full p-2 bg-transparent outline-none text-center font-mono font-bold" />
                  </td>
                  <td className="p-0 border-r border-slate-200/10 px-3 uppercase truncate font-bold">{route.motorista}</td>
                  <td className="p-0 border-r border-slate-200/10 text-center font-black uppercase tracking-widest">{route.placa}</td>
                  <td className="p-0 border-r border-slate-200/10">
                    <input type="text" value={route.saida} onChange={(e) => updateCell(route.id, 'saida', e.target.value)} className="w-full h-full p-2 bg-transparent outline-none text-center font-mono font-bold" />
                  </td>
                  <td className="p-0 border-r border-slate-200/10">
                    <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className={`w-full h-full p-1 bg-transparent outline-none cursor-pointer font-black text-center ${isWhiteText ? 'text-white' : ''}`}>
                      <option value="" className="text-slate-900">Selecione...</option>
                      <option value="Manutenção" className="text-slate-900">Manutenção</option>
                      <option value="Mão de obra" className="text-slate-900">Mão de obra</option>
                      <option value="Atraso coleta" className="text-slate-900">Atraso coleta</option>
                      <option value="Atraso carregamento" className="text-slate-900">Atraso carregamento</option>
                      <option value="Fábrica" className="text-slate-900">Fábrica</option>
                      <option value="Infraestrutura" className="text-slate-900">Infraestrutura</option>
                      <option value="Logística" className="text-slate-900">Logística</option>
                      <option value="Outros" className="text-slate-900">Outros</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-200/10">
                    <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className="w-full h-full p-2 bg-transparent outline-none font-bold italic" placeholder="Descreva o motivo..." />
                  </td>
                  <td className="p-0 border-r border-slate-200/10 text-center">
                     <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none font-black text-center">
                      <option value="OK">OK</option>
                      <option value="NOK">NOK</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-200/10 text-center">
                    <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none font-black text-center">
                      <option value="SIM">SIM</option>
                      <option value="NÃO">NÃO</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-200/10 text-center font-black uppercase tracking-tight">{route.operacao}</td>
                  <td className="p-0 border-r border-slate-200/10 text-center">
                    <span className={`px-2 py-1 rounded font-black text-[9px] ${route.statusOp === 'OK' ? 'bg-emerald-500 text-white' : 'bg-white/20 text-white'}`}>
                        {route.statusOp}
                    </span>
                  </td>
                  <td className="p-0 border-r border-slate-200/10 text-center font-mono font-black">{route.tempo}</td>
                  <td className="p-1 sticky right-0 bg-inherit text-center">
                    <button onClick={() => removeRow(route.id)} className={`${isWhiteText ? 'text-white hover:text-red-200' : 'text-slate-400 hover:text-red-600'} transition-colors p-1`}>
                      <Trash2 size={14} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-md z-[100] flex items-center justify-center p-4">
             <div className="bg-white dark:bg-slate-900 rounded-[2rem] shadow-2xl w-full max-w-2xl overflow-hidden border dark:border-slate-700 animate-in zoom-in duration-200">
                <div className="bg-emerald-600 dark:bg-slate-800 text-white p-6 flex justify-between items-center border-b border-emerald-700 dark:border-slate-700">
                    <div className="flex items-center gap-3">
                        <Upload size={24} />
                        <h3 className="font-black uppercase tracking-widest text-sm">Importar Planilha CCO</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/20 p-1 rounded-full"><X size={24} /></button>
                </div>
                <div className="p-8">
                    <div className="mb-4 text-xs font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                        <FileSpreadsheet size={14} className="text-blue-500"/> Copie as linhas do Excel e cole abaixo
                    </div>
                    <textarea 
                        value={importText} 
                        onChange={e => setImportText(e.target.value)} 
                        className="w-full h-64 p-5 border-2 dark:border-slate-700 rounded-2xl bg-slate-50 dark:bg-slate-800 text-xs font-mono dark:text-white mb-6 focus:ring-4 focus:ring-blue-500/20 outline-none transition-all" 
                        placeholder="Cole aqui os dados copiados do Excel..."
                    />
                    <div className="flex flex-col gap-3">
                        <button onClick={() => handleImport(false)} disabled={isProcessingImport || !importText.trim()} className="w-full py-4 bg-slate-900 text-white font-black uppercase tracking-widest text-xs rounded-2xl shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-slate-800 disabled:opacity-50 active:scale-95 border-b-4 border-slate-950">
                            {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : <><FileSpreadsheet size={20} /> Importação Direta (Recomendado)</>}
                        </button>
                        <p className="text-[9px] text-center text-slate-400 font-bold uppercase">Utiliza algoritmo robusto para identificar as colunas automaticamente</p>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="w-full mt-4 py-3 text-slate-400 font-black uppercase text-[10px] tracking-widest hover:text-slate-600">Cancelar</button>
                </div>
             </div>
        </div>
      )}

      {/* NEW ROUTE MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-black/70 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden border dark:border-slate-700">
            <div className="bg-blue-600 dark:bg-slate-800 text-white p-6 flex justify-between items-center border-b border-blue-700 dark:border-slate-700">
                <h3 className="font-black uppercase tracking-widest text-sm flex items-center gap-3"><Plus size={20} /> Registrar Rota Manual</h3>
                <button onClick={() => setIsModalOpen(false)} className="hover:bg-white/20 p-1 rounded-full"><X size={24} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-8 space-y-5 bg-slate-50 dark:bg-slate-900">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-bold shadow-sm focus:ring-4 focus:ring-blue-500/20 outline-none"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Rota</label>
                        <input type="text" required placeholder="Número" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-black text-blue-600 shadow-sm focus:ring-4 focus:ring-blue-500/20 outline-none"/>
                    </div>
                </div>
                <div className="space-y-1">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Operação Autorizada</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-black shadow-sm outline-none focus:ring-4 focus:ring-blue-500/20">
                        <option value="">Selecione...</option>
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Motorista</label>
                        <input type="text" required placeholder="Nome" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm shadow-sm focus:ring-4 focus:ring-blue-500/20 outline-none"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-black shadow-sm focus:ring-4 focus:ring-blue-500/20 outline-none"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Início Previsto</label>
                        <input type="text" required placeholder="00:00:00" value={formData.inicio} onChange={e => setFormData({...formData, inicio: e.target.value})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-mono shadow-sm focus:ring-4 focus:ring-blue-500/20 outline-none"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Saída Real</label>
                        <input type="text" placeholder="00:00:00" value={formData.saida} onChange={e => setFormData({...formData, saida: e.target.value})} className="w-full p-3 border-2 dark:border-slate-700 rounded-2xl bg-white dark:bg-slate-800 text-sm font-mono shadow-sm focus:ring-4 focus:ring-blue-500/20 outline-none"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-slate-900 hover:bg-slate-800 text-white font-black uppercase tracking-widest text-xs rounded-2xl flex items-center justify-center gap-3 shadow-xl transition-all active:scale-95 border-b-4 border-slate-700 mt-4">
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
