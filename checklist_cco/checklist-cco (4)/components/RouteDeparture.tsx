
import React, { useState, useEffect } from 'react';
import { RouteDeparture } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDepartures } from '../services/geminiService';
import { Plus, Trash2, Download, Save, AlertTriangle, CheckCircle2, Clock, Maximize2, Minimize2, X, FileText, User, Calendar, MapPin, Upload, Sparkles, Loader2, RefreshCw } from 'lucide-react';

const RouteDepartureView: React.FC = () => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importText, setImportText] = useState('');
  
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
    if (!token) return;
    setIsLoading(true);
    try {
      const spData = await SharePointService.getDepartures(token);
      setRoutes(spData);
    } catch (e) {
      console.error(e);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadData();
  }, []);

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

  const calculateGap = (inicio: string, saida: string): { gap: string, status: string } => {
    if (!inicio || !saida || inicio === '00:00:00' || saida === '00:00:00') return { gap: 'OK', status: 'OK' };
    
    const startSec = timeToSeconds(inicio);
    const endSec = timeToSeconds(saida);
    const diff = endSec - startSec;
    
    if (diff === 0) return { gap: 'OK', status: 'OK' };
    
    const gapFormatted = secondsToTime(diff);
    const status = diff > 0 ? 'Atrasado' : 'OK';
    
    return { gap: gapFormatted, status };
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
      operacao: '',
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

    setIsSyncing(true);
    try {
        const week = calculateWeekString(formData.data || '');
        const { gap, status } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00');
        
        const newRoute: RouteDeparture = {
            ...formData as RouteDeparture,
            id: '', 
            semana: week,
            tempo: gap,
            statusOp: status,
            createdAt: new Date().toISOString()
        };

        const newId = await SharePointService.updateDeparture(token, newRoute);
        const finalRoute = { ...newRoute, id: newId };
        setRoutes(prev => [...prev, finalRoute]);
        closeModal();
    } catch (err: any) {
        console.error("Erro ao salvar:", err);
        alert(`ERRO AO SALVAR NO SHAREPOINT: ${err.message}`);
    } finally {
        setIsSyncing(false);
    }
  };

  const handleImport = async () => {
    const token = getAccessToken();
    if (!importText.trim() || !token) return;
    setIsProcessingImport(true);
    try {
        const parsed = await parseRouteDepartures(importText);
        const importPromises = parsed.map(p => {
            const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00');
            const r: RouteDeparture = {
                ...p,
                id: '',
                semana: p.semana || calculateWeekString(p.data || ''),
                saida: p.saida || '00:00:00',
                motivo: p.motivo || '',
                statusGeral: p.statusGeral || 'OK',
                aviso: p.aviso || 'NÃO',
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
        alert(`Rotas importadas e sincronizadas com sucesso!`);
    } catch (error: any) {
        console.error(error);
        alert(`Erro ao processar importação: ${error.message}`);
    } finally {
        setIsProcessingImport(false);
    }
  };

  const removeRow = async (id: string) => {
    const token = getAccessToken();
    if (!token) return;
    if (confirm('Deseja excluir esta rota permanentemente do SharePoint?')) {
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
    
    if (field === 'inicio' || field === 'saida') {
        const { gap, status } = calculateGap(updatedRoute.inicio, updatedRoute.saida);
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
    if (route.statusOp === 'Atrasado') {
      return 'bg-amber-50 dark:bg-amber-900/10 border-amber-200 dark:border-amber-800/50';
    }
    return 'bg-white dark:bg-slate-900 border-gray-100 dark:border-slate-800';
  };

  const getStatusBadgeStyle = (val: string) => {
    if (val === 'Atrasado') return 'bg-red-500 text-white font-bold';
    if (val === 'OK') return 'bg-green-500 text-white font-bold';
    return 'bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300';
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-xs uppercase tracking-widest">Sincronizando Banco de Saídas...</p>
    </div>
  );

  return (
    <div className={`flex flex-col h-full animate-fade-in ${isFullscreen ? 'fixed inset-0 z-[60] bg-white dark:bg-slate-950 p-4' : ''}`}>
      <div className="flex justify-between items-center mb-4 shrink-0">
        <div>
          <h2 className="text-xl font-bold text-gray-800 dark:text-white flex items-center gap-2">
            <Clock className="text-blue-600" />
            Saída de Rotas
            {isSyncing && <Loader2 size={16} className="animate-spin text-blue-500 ml-2"/>}
          </h2>
          <p className="text-xs text-gray-500 dark:text-gray-400">Controle de horários e motivos de atraso</p>
        </div>
        <div className="flex gap-2">
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-blue-500 transition-colors" title="Atualizar dados">
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
          <thead className="sticky top-0 z-20 bg-blue-900 dark:bg-blue-950 text-white uppercase font-bold tracking-wider text-center">
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
                <td colSpan={15} className="p-12 text-center text-slate-400 font-bold uppercase italic">
                   Nenhuma rota sincronizada do SharePoint.
                </td>
              </tr>
            )}
            {routes.map((route) => (
              <tr key={route.id} className={`hover:bg-blue-50/50 dark:hover:bg-blue-900/10 border-b ${getRowStyle(route)}`}>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center font-bold bg-slate-50 dark:bg-slate-800/50">{route.semana}</td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none font-bold text-center" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.inicio} onChange={(e) => updateCell(route.id, 'inicio', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" placeholder="00:00:00" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none uppercase truncate" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-bold uppercase" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.saida} onChange={(e) => updateCell(route.id, 'saida', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none text-center font-mono" placeholder="00:00:00" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none cursor-pointer">
                    <option value="">Nenhum</option>
                    <option value="Manutenção">Manutenção</option>
                    <option value="Mão de obra">Mão de obra</option>
                    <option value="Atraso coleta">Atraso coleta</option>
                    <option value="Atraso carregamento">Atraso carregamento</option>
                    <option value="Outros">Outros</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center">
                   <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none text-center font-bold">
                    <option value="OK">OK</option>
                    <option value="NOK">NOK</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center">
                  <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full p-1 bg-transparent outline-none text-center font-bold">
                    <option value="SIM">SIM</option>
                    <option value="NÃO">NÃO</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800">
                  <input type="text" value={route.operacao} onChange={(e) => updateCell(route.id, 'operacao', e.target.value)} className="w-full h-full p-1.5 bg-transparent outline-none uppercase font-bold" />
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center">
                  <select value={route.statusOp} onChange={(e) => updateCell(route.id, 'statusOp', e.target.value)} className={`w-full h-full p-1 bg-transparent outline-none text-center font-bold ${getStatusBadgeStyle(route.statusOp)}`}>
                    <option value="OK">OK</option>
                    <option value="Atrasado">Atrasado</option>
                  </select>
                </td>
                <td className="p-0 border border-gray-200 dark:border-slate-800 text-center font-mono font-bold bg-slate-50/50 dark:bg-slate-800/30">
                   {route.tempo}
                </td>
                <td className="p-1 border border-gray-200 dark:border-slate-800 sticky right-0 bg-white dark:bg-slate-900 text-center">
                  <button onClick={() => removeRow(route.id)} className="text-gray-400 hover:text-red-500 transition-colors p-1">
                    <Trash2 size={14} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
             <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-2xl overflow-hidden border dark:border-slate-700">
                <div className="bg-emerald-800 dark:bg-slate-800 text-white p-4 flex justify-between items-center">
                    <div className="flex items-center gap-2">
                        <Upload size={20} />
                        <h3 className="font-bold">Importar do Excel (SharePoint)</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)}><X size={24} /></button>
                </div>
                <div className="p-6">
                    <textarea value={importText} onChange={e => setImportText(e.target.value)} className="w-full h-64 p-4 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-xs font-mono dark:text-white mb-4" placeholder="Cole as linhas do Excel aqui..."/>
                    <div className="flex gap-3">
                        <button onClick={() => setIsImportModalOpen(false)} className="flex-1 py-3 bg-gray-200 dark:bg-slate-700 font-bold rounded-xl">Cancelar</button>
                        <button onClick={handleImport} disabled={isProcessingImport} className="flex-[2] py-3 bg-emerald-600 text-white font-bold rounded-xl flex items-center justify-center gap-2">
                            {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : <><Sparkles size={20} /> Processar e Sincronizar</>}
                        </button>
                    </div>
                </div>
             </div>
        </div>
      )}

      {/* ADD ROUTE MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden border dark:border-slate-700">
            <div className="bg-blue-600 dark:bg-slate-800 text-white p-4 flex justify-between items-center">
                <h3 className="font-bold flex items-center gap-2"><Plus size={20} /> Nova Saída</h3>
                <button onClick={closeModal}><X size={24} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-6 space-y-4 bg-gray-50 dark:bg-slate-900">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Rota</label>
                        <input type="text" required placeholder="Ex: 24133D" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-bold"/>
                    </div>
                </div>
                <div className="space-y-1">
                    <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Operação</label>
                    <input type="text" required placeholder="CLIENTE / UNIDADE" value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value.toUpperCase()})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-bold"/>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Motorista</label>
                        <input type="text" required placeholder="NOME" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-bold"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Início Previsto</label>
                        <input type="text" required placeholder="00:00:00" value={formData.inicio} onChange={e => setFormData({...formData, inicio: e.target.value})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-mono"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">Saída Real</label>
                        <input type="text" placeholder="00:00:00" value={formData.saida} onChange={e => setFormData({...formData, saida: e.target.value})} className="w-full p-2 border dark:border-slate-700 rounded-lg bg-white dark:bg-slate-800 text-sm font-mono"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-3 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl flex items-center justify-center gap-2 shadow-lg transition-all active:scale-95">
                    {isSyncing ? <Loader2 size={20} className="animate-spin" /> : <Save size={20} />} Confirmar e Gravar no SharePoint
                </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
