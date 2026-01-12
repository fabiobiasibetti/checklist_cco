
import React, { useState, useEffect, useRef } from 'react';
import { RouteDeparture, User, RouteOperationMapping } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDepartures, parseRouteDeparturesManual } from '../services/geminiService';
import { 
  Plus, Trash2, Save, Clock, X, Upload, 
  Loader2, RefreshCw, ShieldCheck, FileSpreadsheet,
  AlertTriangle, Link, CheckCircle2, ChevronRight, ChevronDown
} from 'lucide-react';

interface RouteConfig {
    operacao: string;
    email: string;
    tolerancia: string;
}

const RouteDepartureView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [routeMappings, setRouteMappings] = useState<RouteOperationMapping[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [isLinkModalOpen, setIsLinkModalOpen] = useState(false);
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importText, setImportText] = useState('');
  const [currentTime, setCurrentTime] = useState(new Date());

  // Pending items from import that need operation linking
  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  // Resizing state
  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 80,
    rota: 120,
    data: 130,
    inicio: 100,
    motorista: 250,
    placa: 110,
    saida: 100,
    motivo: 140,
    observacao: 300,
    geral: 80,
    aviso: 80,
    operacao: 150,
    status: 90,
    tempo: 90,
  });

  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

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
      const [configs, mappings, spData] = await Promise.all([
        SharePointService.getRouteConfigs(token, currentUser.email),
        SharePointService.getRouteOperationMappings(token),
        SharePointService.getDepartures(token)
      ]);
      
      setUserConfigs(configs);
      setRouteMappings(mappings);

      const allowedOps = new Set(configs.map(c => c.operacao.toUpperCase().trim()));
      
      // Auto-fix existing routes without operations if they exist in mappings
      const fixedData = spData.map(route => {
        if (!route.operacao || route.operacao === "") {
            const match = mappings.find(m => m.Title === route.rota);
            if (match && allowedOps.has(match.OPERACAO.toUpperCase().trim())) {
                return { ...route, operacao: match.OPERACAO.toUpperCase().trim() };
            }
        }
        return route;
      });

      const filtered = fixedData.filter(r => allowedOps.has(r.operacao.toUpperCase().trim()) || !r.operacao);
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
    
    const handleMouseMove = (e: MouseEvent) => {
      if (resizingRef.current) {
        const { col, startX, startWidth } = resizingRef.current;
        const newWidth = Math.max(50, startWidth + (e.clientX - startX));
        setColWidths(prev => ({ ...prev, [col]: newWidth }));
      }
    };

    const handleMouseUp = () => {
      resizingRef.current = null;
      document.body.style.cursor = 'default';
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    
    return () => {
      clearInterval(timer);
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [currentUser]);

  const startResize = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    resizingRef.current = {
      col,
      startX: e.clientX,
      startWidth: colWidths[col]
    };
    document.body.style.cursor = 'col-resize';
  };

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

  const handleImport = async () => {
    if (!importText.trim()) return;
    setIsProcessingImport(true);
    try {
        const parsed = parseRouteDeparturesManual(importText);
        if (parsed.length === 0) throw new Error("Nenhum dado válido identificado.");

        const token = getAccessToken();
        const allowedOps = new Set(userConfigs.map(c => c.operacao.toUpperCase().trim()));
        
        const itemsToSave: Partial<RouteDeparture>[] = [];
        const itemsToLink: Partial<RouteDeparture>[] = [];

        for (const item of parsed) {
            const routeName = item.rota?.trim() || "";
            const mapping = routeMappings.find(m => m.Title.trim() === routeName);
            
            if (mapping && allowedOps.has(mapping.OPERACAO.toUpperCase().trim())) {
                itemsToSave.push({ ...item, operacao: mapping.OPERACAO.toUpperCase().trim() });
            } else {
                itemsToLink.push({ ...item, operacao: "" });
            }
        }

        // Save automatic mappings
        if (itemsToSave.length > 0) {
            await Promise.all(itemsToSave.map(async (p) => {
                const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === p.operacao!.toUpperCase().trim());
                const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia);
                
                const r: RouteDeparture = {
                    ...p,
                    id: '',
                    semana: calculateWeekString(p.data || ''),
                    statusGeral: 'OK',
                    aviso: 'NÃO',
                    statusOp: status,
                    tempo: gap,
                    createdAt: new Date().toISOString()
                } as RouteDeparture;

                return SharePointService.updateDeparture(token!, r);
            }));
        }

        if (itemsToLink.length > 0) {
            setPendingItems(itemsToLink);
            setIsLinkModalOpen(true);
        } else if (itemsToSave.length > 0) {
            alert(`Sucesso! ${itemsToSave.length} rotas importadas automaticamente.`);
        }

        await loadData();
        setIsImportModalOpen(false);
        setImportText('');
    } catch (error: any) {
        alert(`Erro na importação: ${error.message}`);
    } finally {
        setIsProcessingImport(false);
    }
  };

  const handleLinkPending = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = getAccessToken();
    if (!token) return;

    setIsProcessingImport(true);
    try {
        const ready = pendingItems.filter(p => p.operacao && p.operacao !== "");
        
        await Promise.all(ready.map(async (p) => {
            // Check if mapping already exists to avoid duplicates
            const exists = routeMappings.some(m => m.Title === p.rota && m.OPERACAO === p.operacao);
            if (!exists) {
                await SharePointService.addRouteOperationMapping(token, p.rota!, p.operacao!);
            }
            
            const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === p.operacao!.toUpperCase().trim());
            const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia);
            
            const r: RouteDeparture = {
                ...p,
                id: '',
                semana: calculateWeekString(p.data || ''),
                statusGeral: 'OK',
                aviso: 'NÃO',
                statusOp: status,
                tempo: gap,
                createdAt: new Date().toISOString()
            } as RouteDeparture;

            return SharePointService.updateDeparture(token, r);
        }));

        await loadData();
        setIsLinkModalOpen(false);
        setPendingItems([]);
        alert(`Sucesso! ${ready.length} novas vinculações e rotas salvas.`);
    } catch (err: any) {
        alert("Erro ao salvar: " + err.message);
    } finally {
        setIsProcessingImport(false);
    }
  };

  const getRowStyle = (route: RouteDeparture) => {
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === route.operacao.toUpperCase().trim());
    const { isOutOfTolerance } = calculateGap(route.inicio, route.saida, config?.tolerancia);

    if (route.saida !== '00:00:00' && isOutOfTolerance) {
        return 'bg-orange-500 text-white font-bold';
    }

    const toleranceSec = timeToSeconds(config?.tolerancia || "00:00:00");
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) {
        return 'bg-yellow-300 text-slate-900 font-bold';
    }

    return 'bg-white dark:bg-slate-900 text-slate-700 dark:text-slate-200';
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
        <p className="font-bold animate-pulse text-xs uppercase tracking-widest text-center">Sincronizando Dados CCO...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-slate-50 dark:bg-slate-950 p-4 overflow-hidden">
      {/* HEADER SECTION */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 shrink-0 gap-4">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-600 text-white rounded-2xl shadow-xl">
            <Clock size={24} />
          </div>
          <div>
            <h2 className="text-2xl font-black text-slate-800 dark:text-white uppercase tracking-tight flex items-center gap-3">
              Saída de Rotas
              {isSyncing && <Loader2 size={18} className="animate-spin text-blue-500"/>}
            </h2>
            <div className="flex items-center gap-2">
                <ShieldCheck size={14} className="text-emerald-500"/>
                <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">CCO Logística: {currentUser.name}</p>
            </div>
          </div>
        </div>
        <div className="flex gap-2 w-full sm:w-auto">
          <button onClick={loadData} className="p-2.5 text-slate-400 hover:text-blue-600 hover:bg-white dark:hover:bg-slate-900 rounded-xl transition-all shadow-sm border border-slate-200 dark:border-slate-800">
              <RefreshCw size={20} />
          </button>
          <button 
            onClick={() => setIsImportModalOpen(true)}
            className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-5 py-2.5 bg-emerald-500 text-white rounded-xl hover:bg-emerald-600 font-black shadow-lg transition-all active:scale-95 border-b-4 border-emerald-700 uppercase text-xs tracking-widest"
          >
            <Upload size={18} /> Importar Excel
          </button>
          <button 
            onClick={() => setIsModalOpen(true)}
            className="flex-1 sm:flex-none flex items-center justify-center gap-2 px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl transition-all shadow-lg active:scale-95 border-b-4 border-blue-800 uppercase text-xs tracking-widest"
          >
            <Plus size={18} /> Nova Rota
          </button>
        </div>
      </div>

      {/* DATA GRID - EXCEL STYLE */}
      <div className="flex-1 overflow-auto bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 shadow-2xl relative scrollbar-thin">
        <table className="border-collapse table-fixed w-full min-w-max bg-white dark:bg-slate-900">
          <thead className="sticky top-0 z-20 bg-slate-50 dark:bg-slate-950">
            <tr className="border-b border-slate-200 dark:border-slate-800">
              {[
                { id: 'semana', label: 'SEMANA' },
                { id: 'rota', label: 'ROTA' },
                { id: 'data', label: 'DATA' },
                { id: 'inicio', label: 'INÍCIO' },
                { id: 'motorista', label: 'MOTORISTA' },
                { id: 'placa', label: 'PLACA' },
                { id: 'saida', label: 'SAÍDA' },
                { id: 'motivo', label: 'MOTIVO' },
                { id: 'observacao', label: 'OBSERVAÇÃO' },
                { id: 'geral', label: 'GERAL' },
                { id: 'aviso', label: 'AV' },
                { id: 'operacao', label: 'OPERAÇÃO' },
                { id: 'status', label: 'STATUS' },
                { id: 'tempo', label: 'TEMPO' }
              ].map(col => (
                <th 
                  key={col.id} 
                  style={{ width: colWidths[col.id] }}
                  className="relative p-3 text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-widest text-left select-none border-r border-slate-200 dark:border-slate-800 group"
                >
                  <div className="truncate pr-4">{col.label}</div>
                  <div 
                    onMouseDown={(e) => startResize(e, col.id)}
                    className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-blue-500 transition-colors z-10 flex items-center justify-center"
                  >
                     <div className="w-[1px] h-4 bg-slate-200 dark:bg-slate-700 group-hover:bg-blue-300" />
                  </div>
                </th>
              ))}
              <th className="p-3 w-12 sticky right-0 bg-slate-50 dark:bg-slate-950 z-30"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
            {routes.map((route) => {
              const rowStyle = getRowStyle(route);
              const isColored = rowStyle.includes('bg-orange') || rowStyle.includes('bg-yellow');
              
              const baseCellClass = "p-0 border-r border-slate-100 dark:border-slate-800/50";
              const inputClass = `w-full h-full bg-transparent outline-none border-none p-2.5 text-center transition-all focus:bg-blue-500/10 ${isColored ? 'text-inherit placeholder-white/60' : 'text-slate-700 dark:text-slate-300 placeholder-slate-300'}`;

              return (
                <tr key={route.id} className={`${rowStyle} group hover:brightness-95 transition-all h-12 text-[11px] font-medium`}>
                  <td className={`${baseCellClass} text-center font-bold bg-slate-50/30 dark:bg-slate-900/30`}>{route.semana}</td>
                  <td className={baseCellClass}>
                    <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={inputClass} />
                  </td>
                  <td className={baseCellClass}>
                    <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={inputClass} />
                  </td>
                  <td className={baseCellClass}>
                    <input type="text" value={route.inicio} onChange={(e) => updateCell(route.id, 'inicio', e.target.value)} className={`${inputClass} font-mono`} />
                  </td>
                  <td className={baseCellClass}>
                    <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value.toUpperCase())} className={`${inputClass} text-left uppercase font-bold px-3`} />
                  </td>
                  <td className={baseCellClass}>
                    <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value.toUpperCase())} className={`${inputClass} font-bold tracking-widest`} />
                  </td>
                  <td className={baseCellClass}>
                    <input type="text" value={route.saida} onChange={(e) => updateCell(route.id, 'saida', e.target.value)} className={`${inputClass} font-mono`} />
                  </td>
                  <td className={baseCellClass}>
                    <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full h-full bg-transparent outline-none p-1 text-center cursor-pointer appearance-none">
                      <option value="">--</option>
                      <option value="Manutenção">Manutenção</option>
                      <option value="Mão de obra">Mão de obra</option>
                      <option value="Atraso coleta">Atraso coleta</option>
                      <option value="Atraso carregamento">Atraso carregamento</option>
                      <option value="Fábrica">Fábrica</option>
                      <option value="Infraestrutura">Infraestrutura</option>
                      <option value="Logística">Logística</option>
                      <option value="Outros">Outros</option>
                    </select>
                  </td>
                  <td className={baseCellClass}>
                    <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className={`${inputClass} text-left italic px-3`} placeholder="..." />
                  </td>
                  <td className={baseCellClass}>
                    <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full bg-transparent outline-none p-1 text-center font-bold appearance-none">
                      <option value="OK">OK</option>
                      <option value="NOK">NOK</option>
                    </select>
                  </td>
                  <td className={baseCellClass}>
                    <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full bg-transparent outline-none p-1 text-center font-bold appearance-none">
                      <option value="SIM">SIM</option>
                      <option value="NÃO">NÃO</option>
                    </select>
                  </td>
                  <td className={`${baseCellClass} text-center font-black uppercase text-[9px] px-2`}>
                    {route.operacao || <span className="text-red-500 animate-pulse underline">VINCULAR</span>}
                  </td>
                  <td className={`${baseCellClass} text-center px-2`}>
                    <span className={`px-2 py-0.5 rounded text-[9px] font-black ${route.statusOp === 'OK' ? 'bg-emerald-500 text-white' : 'bg-white/30 text-inherit'}`}>
                      {route.statusOp}
                    </span>
                  </td>
                  <td className={`${baseCellClass} text-center font-mono font-bold`}>{route.tempo}</td>
                  <td className="p-2 sticky right-0 bg-inherit text-center shadow-[-4px_0_8px_rgba(0,0,0,0.05)]">
                    <button onClick={() => removeRow(route.id)} className="text-slate-300 hover:text-red-500 p-1 rounded-lg transition-colors">
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* MODAL: LINK ROUTES TO OPERATIONS */}
      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[110] flex items-center justify-center p-4">
            <div className="bg-white dark:bg-slate-900 rounded-[2rem] shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 dark:border-slate-800 animate-in zoom-in duration-300">
                <div className="bg-amber-600 p-6 flex justify-between items-center text-white shadow-lg">
                    <div className="flex items-center gap-3">
                        <Link size={24} />
                        <h3 className="font-black uppercase tracking-widest text-sm">Vincular Operações de Rota</h3>
                    </div>
                </div>
                <div className="p-8">
                    <div className="flex items-center gap-4 p-5 bg-amber-50 dark:bg-amber-900/20 border border-amber-100 dark:border-amber-800 rounded-2xl mb-6 text-amber-700 dark:text-amber-400">
                        <AlertTriangle size={28} className="shrink-0" />
                        <p className="text-xs font-bold leading-relaxed">
                            Atenção: Algumas rotas importadas não foram reconhecidas. 
                            Selecione a operação correspondente para salvar e automatizar futuras importações.
                        </p>
                    </div>

                    <div className="max-h-[300px] overflow-y-auto space-y-3 mb-8 pr-2 scrollbar-thin">
                        {pendingItems.map((item, idx) => (
                            <div key={idx} className="flex items-center gap-4 p-4 bg-slate-50 dark:bg-slate-800/50 rounded-2xl border border-slate-100 dark:border-slate-700 hover:border-amber-200 transition-colors">
                                <div className="flex-1">
                                    <span className="text-[9px] text-slate-400 font-black uppercase tracking-widest block mb-1">Rota Identificada</span>
                                    <div className="font-black text-blue-600 dark:text-blue-400 text-sm truncate">{item.rota}</div>
                                </div>
                                <div className="w-[60%]">
                                    <span className="text-[9px] text-slate-400 font-black uppercase tracking-widest block mb-1">Operação Associada</span>
                                    <select 
                                        value={item.operacao} 
                                        onChange={(e) => {
                                            const newPending = [...pendingItems];
                                            newPending[idx].operacao = e.target.value;
                                            setPendingItems(newPending);
                                        }}
                                        className="w-full p-2.5 bg-white dark:bg-slate-950 border border-slate-200 dark:border-slate-700 rounded-xl text-xs font-black outline-none focus:ring-4 focus:ring-blue-500/20 shadow-sm transition-all"
                                    >
                                        <option value="">Selecione...</option>
                                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                                    </select>
                                </div>
                            </div>
                        ))}
                    </div>

                    <div className="flex gap-4">
                        <button 
                            onClick={() => { setPendingItems([]); setIsLinkModalOpen(false); }} 
                            className="flex-1 py-4 text-slate-500 font-black uppercase tracking-widest text-[10px] hover:bg-slate-100 rounded-2xl transition-all"
                        >
                            Ignorar e Sair
                        </button>
                        <button 
                            onClick={handleLinkPending} 
                            disabled={isProcessingImport || !pendingItems.some(p => p.operacao)} 
                            className="flex-[2] py-4 bg-blue-600 text-white font-black uppercase tracking-widest text-xs rounded-2xl shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-blue-700 active:scale-95 disabled:opacity-50 border-b-4 border-blue-800"
                        >
                            {isProcessingImport ? <Loader2 size={18} className="animate-spin" /> : <><CheckCircle2 size={18} /> Salvar Vínculos</>}
                        </button>
                    </div>
                </div>
            </div>
        </div>
      )}

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[100] flex items-center justify-center p-4">
             <div className="bg-white dark:bg-slate-900 rounded-[2rem] shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 dark:border-slate-800 animate-in zoom-in duration-200">
                <div className="bg-emerald-600 p-6 flex justify-between items-center text-white shadow-lg">
                    <div className="flex items-center gap-3">
                        <Upload size={24} />
                        <h3 className="font-black uppercase tracking-widest text-sm">Importar Planilha CCO</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/20 p-2 rounded-xl transition-colors"><X size={24} /></button>
                </div>
                <div className="p-8">
                    <div className="mb-4 text-xs font-bold text-slate-500 uppercase tracking-widest flex items-center gap-2">
                        <FileSpreadsheet size={16} className="text-blue-500"/> Copie as linhas do Excel e cole abaixo
                    </div>
                    <textarea 
                        value={importText} 
                        onChange={e => setImportText(e.target.value)} 
                        className="w-full h-64 p-5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-xs font-mono mb-6 focus:ring-4 focus:ring-blue-500/20 outline-none transition-all dark:text-white shadow-inner" 
                        placeholder="Cole aqui os dados copiados do Excel..."
                    />
                    <div className="flex flex-col gap-3">
                        <button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-4 bg-slate-900 text-white font-black uppercase tracking-widest text-xs rounded-2xl shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-slate-800 disabled:opacity-50 active:scale-95 border-b-4 border-slate-950">
                            {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : <><FileSpreadsheet size={20} /> Iniciar Importação Inteligente</>}
                        </button>
                        <p className="text-[9px] text-center text-slate-400 font-bold uppercase tracking-widest">A vinculação automática ocorrerá após a importação utilizando a lista 'Rotas_Operacao_Checklist'</p>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="w-full mt-4 py-3 text-slate-400 font-black uppercase text-[10px] tracking-widest hover:text-slate-600 transition-colors">Cancelar</button>
                </div>
             </div>
        </div>
      )}

      {/* NEW ROUTE MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden border border-slate-200 dark:border-slate-800">
            <div className="bg-blue-600 text-white p-6 flex justify-between items-center shadow-lg">
                <h3 className="font-black uppercase tracking-widest text-sm flex items-center gap-3"><Plus size={20} /> Registrar Rota Manual</h3>
                <button onClick={() => setIsModalOpen(false)} className="hover:bg-white/20 p-2 rounded-xl transition-colors"><X size={24} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-8 space-y-6">
                <div className="grid grid-cols-2 gap-5">
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Data Operação</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-bold focus:ring-4 focus:ring-blue-500/20 outline-none dark:text-white transition-all"/>
                    </div>
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Nº Rota</label>
                        <input type="text" required placeholder="Ex: 24133" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-black text-blue-600 focus:ring-4 focus:ring-blue-500/20 outline-none shadow-sm transition-all"/>
                    </div>
                </div>
                <div className="space-y-1.5">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Operação CCO</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-black focus:ring-4 focus:ring-blue-500/20 outline-none dark:text-white transition-all shadow-sm">
                        <option value="">Selecione...</option>
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-5">
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Motorista</label>
                        <input type="text" required placeholder="Nome Completo" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-bold focus:ring-4 focus:ring-blue-500/20 outline-none dark:text-white transition-all shadow-sm"/>
                    </div>
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Placa</label>
                        <input type="text" required placeholder="ABC-1234" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-black focus:ring-4 focus:ring-blue-500/20 outline-none dark:text-white transition-all shadow-sm"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-5">
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Início Previsto</label>
                        <input type="text" required placeholder="00:00:00" value={formData.inicio} onChange={e => setFormData({...formData, inicio: e.target.value})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-mono focus:ring-4 focus:ring-blue-500/20 outline-none dark:text-white transition-all shadow-sm"/>
                    </div>
                    <div className="space-y-1.5">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Saída Realizada</label>
                        <input type="text" placeholder="00:00:00" value={formData.saida} onChange={e => setFormData({...formData, saida: e.target.value})} className="w-full p-3.5 border-2 border-slate-100 dark:border-slate-800 rounded-2xl bg-slate-50 dark:bg-slate-950 text-sm font-mono focus:ring-4 focus:ring-blue-500/20 outline-none dark:text-white transition-all shadow-sm"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-5 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-widest text-xs rounded-2xl flex items-center justify-center gap-3 shadow-xl transition-all active:scale-95 border-b-4 border-blue-800 mt-4">
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
