import React, { useState, useEffect, useRef } from 'react';
import { RouteDeparture, User, RouteOperationMapping } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDepartures, parseRouteDeparturesManual } from '../services/geminiService';
import { 
  Plus, Trash2, Save, Clock, X, Upload, 
  Loader2, RefreshCw, ShieldCheck, FileSpreadsheet,
  AlertTriangle, Link, CheckCircle2, ChevronRight, ChevronDown, Sparkles
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

  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 80,
    rota: 120,
    data: 140,
    inicio: 100,
    motorista: 260,
    placa: 110,
    saida: 100,
    motivo: 140,
    observacao: 300,
    geral: 70,
    aviso: 70,
    operacao: 150,
    status: 90,
    tempo: 90,
  });

  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  const getAccessToken = () => (window as any).__access_token;

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
      
      const fixedData = spData.map(route => {
        if (!route.operacao || route.operacao === "") {
            const match = mappings.find(m => m.Title === route.rota);
            if (match && allowedOps.has(match.OPERACAO.toUpperCase().trim())) {
                return { ...route, operacao: match.OPERACAO.toUpperCase().trim() };
            }
        }
        return route;
      });

      // Se houver rotas sem operação, força o popup
      const itemsToLink = fixedData.filter(r => !r.operacao || r.operacao === "");
      if (itemsToLink.length > 0) {
          setPendingItems(itemsToLink);
          setIsLinkModalOpen(true);
      }

      setRoutes(fixedData);
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
        const newWidth = Math.max(40, startWidth + (e.clientX - startX));
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

  const formatTimeInput = (value: string): string => {
    let clean = value.replace(/[^0-9:]/g, '');
    if (!clean) return '00:00:00';
    
    const parts = clean.split(':');
    let h = parts[0] || '00';
    let m = parts[1] || '00';
    let s = parts[2] || '00';

    if (parts.length === 1 && h.length > 0) {
        m = '00';
        s = '00';
    } else if (parts.length === 2) {
        s = '00';
    }

    h = h.padStart(2, '0').substring(0, 2);
    m = m.padStart(2, '0').substring(0, 2);
    s = s.padStart(2, '0').substring(0, 2);

    return `${h}:${m}:${s}`;
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
        return `${monthNames[date.getMonth()]} S${Math.ceil(date.getDate() / 7)}`;
    } catch(e) { return ''; }
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    const token = getAccessToken();
    if (!token) return;

    const route = routes.find(r => r.id === id);
    if (!route) return;

    let finalValue = value;
    if (field === 'inicio' || field === 'saida') {
        finalValue = formatTimeInput(value);
    }

    let updatedRoute = { ...route, [field]: finalValue };
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

        if (itemsToSave.length > 0) {
            await Promise.all(itemsToSave.map(async (p) => {
                const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === p.operacao!.toUpperCase().trim());
                const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia);
                const r: RouteDeparture = {
                    ...p, id: '', semana: calculateWeekString(p.data || ''), statusGeral: 'OK', aviso: 'NÃO',
                    statusOp: status, tempo: gap, createdAt: new Date().toISOString()
                } as RouteDeparture;
                return SharePointService.updateDeparture(token!, r);
            }));
        }

        if (itemsToLink.length > 0) {
            setPendingItems(itemsToLink);
            setIsLinkModalOpen(true);
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

    if (pendingItems.some(p => !p.operacao || p.operacao === "")) {
        alert("Por favor, selecione a operação para TODAS as rotas antes de salvar.");
        return;
    }

    setIsProcessingImport(true);
    try {
        await Promise.all(pendingItems.map(async (p) => {
            const exists = routeMappings.some(m => m.Title === p.rota && m.OPERACAO === p.operacao);
            if (!exists) {
                await SharePointService.addRouteOperationMapping(token, p.rota!, p.operacao!);
            }
            const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === p.operacao!.toUpperCase().trim());
            const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia);
            const r: RouteDeparture = {
                ...p, id: p.id || '', semana: calculateWeekString(p.data || ''), statusGeral: 'OK', aviso: 'NÃO',
                statusOp: status, tempo: gap, createdAt: new Date().toISOString()
            } as RouteDeparture;
            return SharePointService.updateDeparture(token, r);
        }));

        await loadData();
        setIsLinkModalOpen(false);
        setPendingItems([]);
    } catch (err: any) {
        alert("Erro ao salvar: " + err.message);
    } finally {
        setIsProcessingImport(false);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = getAccessToken();
    if (!token) return;

    setIsSyncing(true);
    try {
        const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === formData.operacao?.toUpperCase().trim());
        const { gap, status } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia);
        
        const newRoute: RouteDeparture = {
            ...formData, id: '', semana: calculateWeekString(formData.data || ''), statusOp: status, tempo: gap,
            statusGeral: formData.statusGeral || 'OK', aviso: formData.aviso || 'NÃO', createdAt: new Date().toISOString()
        } as RouteDeparture;

        const newId = await SharePointService.updateDeparture(token, newRoute);
        setRoutes(prev => [{ ...newRoute, id: newId }, ...prev]);
        setIsModalOpen(false);
        setFormData({
            rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '00:00:00',
            motorista: '', placa: '', operacao: '', motivo: '', observacao: '', statusGeral: 'OK', aviso: 'NÃO',
        });
    } catch (err: any) {
        alert("Erro ao salvar: " + err.message);
    } finally {
        setIsSyncing(false);
    }
  };

  const removeRow = async (id: string) => {
    const token = getAccessToken();
    if (!token) return;
    if (confirm('Excluir permanentemente do SharePoint?')) {
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

  const getRowClass = (route: RouteDeparture) => {
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === route.operacao.toUpperCase().trim());
    const { isOutOfTolerance } = calculateGap(route.inicio, route.saida, config?.tolerancia);
    
    if (route.saida !== '00:00:00' && isOutOfTolerance) return 'bg-orange-600 text-white font-black shadow-inner';
    
    const toleranceSec = timeToSeconds(config?.tolerancia || "00:00:00");
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) return 'bg-yellow-400 text-slate-900 font-black shadow-inner animate-pulse-subtle';
    
    return 'bg-white text-slate-700';
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-[10px] uppercase tracking-widest">Sincronizando CCO...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-slate-950 p-4 overflow-hidden select-none">
      {/* HEADER */}
      <div className="flex justify-between items-center mb-6 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-600 text-white rounded-2xl shadow-2xl">
            <Clock size={24} />
          </div>
          <div>
            <h2 className="text-2xl font-black text-white uppercase tracking-tighter flex items-center gap-3">
              Saída de Rotas
              {isSyncing && <Loader2 size={18} className="animate-spin text-blue-500"/>}
            </h2>
            <div className="flex items-center gap-2">
                <ShieldCheck size={14} className="text-emerald-500"/>
                <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Operação Ativa: {currentUser.name}</p>
            </div>
          </div>
        </div>
        <div className="flex gap-3">
          <button onClick={loadData} className="p-2.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-xl transition-all border border-slate-800">
              <RefreshCw size={20} />
          </button>
          <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-5 py-2.5 bg-emerald-600 text-white rounded-xl hover:bg-emerald-700 font-black shadow-lg border-b-4 border-emerald-900 uppercase text-[10px] tracking-widest transition-all active:scale-95">
            <Upload size={18} /> Importar Excel
          </button>
          <button onClick={() => setIsModalOpen(true)} className="flex items-center gap-2 px-5 py-2.5 bg-blue-600 text-white rounded-xl hover:bg-blue-700 font-black shadow-lg border-b-4 border-blue-900 uppercase text-[10px] tracking-widest transition-all active:scale-95">
            <Plus size={18} /> Nova Rota
          </button>
        </div>
      </div>

      {/* GRADE BRANCA */}
      <div className="flex-1 overflow-auto bg-white rounded-2xl border border-slate-800 shadow-2xl relative scrollbar-thin">
        <table className="border-collapse table-fixed w-full min-w-max">
          <thead className="sticky top-0 z-20 bg-blue-600 text-white">
            <tr className="border-none h-12">
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
                  className="relative p-2 text-[10px] font-black uppercase tracking-widest text-center select-none group border-r border-blue-500/30"
                >
                  {col.label}
                  <div 
                    onMouseDown={(e) => startResize(e, col.id)}
                    className="absolute right-0 top-2 bottom-2 w-1.5 cursor-col-resize hover:bg-white transition-colors z-10 flex items-center justify-center"
                  >
                     <div className="w-[1px] h-full bg-blue-400 group-hover:bg-white" />
                  </div>
                </th>
              ))}
              <th className="p-3 w-12 sticky right-0 bg-blue-600"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {routes.map((route) => {
              const rowClass = getRowClass(route);
              const isAlert = rowClass.includes('bg-orange') || rowClass.includes('bg-yellow');
              
              const inputClass = `w-full h-full bg-transparent outline-none border-none p-2.5 text-center transition-all focus:bg-blue-500/10 ${isAlert ? 'placeholder-white/50 text-inherit' : 'text-slate-700 placeholder-slate-300'}`;

              return (
                <tr key={route.id} className={`${rowClass} group hover:brightness-95 transition-all h-11 text-[11px] font-medium`}>
                  <td className="p-0 border-r border-slate-100 text-center font-black bg-slate-50/20">{route.semana}</td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={inputClass} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${inputClass} text-[10px]`} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.inicio} onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} defaultValue={route.inicio} className={`${inputClass} font-mono`} placeholder="00:00" />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value.toUpperCase())} className={`${inputClass} text-left uppercase font-black px-4`} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value.toUpperCase())} className={`${inputClass} font-black tracking-widest`} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    {/* Fixed typo below: changed inputStyle to inputClass */}
                    <input type="text" value={route.saida} onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} defaultValue={route.saida} className={`${inputClass} font-mono`} placeholder="00:00" />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full h-full bg-transparent outline-none p-1 text-center cursor-pointer appearance-none font-bold">
                      <option value="">Selecione...</option>
                      {['Manutenção', 'Mão de obra', 'Atraso coleta', 'Atraso carregamento', 'Fábrica', 'Infraestrutura', 'Logística', 'Outros'].map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className={`${inputClass} text-left italic px-4`} placeholder="..." />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full bg-transparent outline-none p-1 text-center font-black appearance-none">
                      <option value="OK">OK</option>
                      <option value="NOK">NOK</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full bg-transparent outline-none p-1 text-center font-black appearance-none">
                      <option value="SIM">SIM</option>
                      <option value="NÃO">NÃO</option>
                    </select>
                  </td>
                  <td className="p-2 border-r border-slate-100 text-center font-black uppercase text-[9px] truncate">
                      {route.operacao || <span className="text-red-600 animate-pulse underline">VINCULAR</span>}
                  </td>
                  <td className="p-2 border-r border-slate-100 text-center">
                    <span className={`px-2 py-0.5 rounded text-[9px] font-black ${route.statusOp === 'OK' ? 'bg-emerald-600 text-white' : 'bg-red-600 text-white'}`}>
                      {route.statusOp}
                    </span>
                  </td>
                  <td className="p-2 border-r border-slate-100 text-center font-mono font-black">{route.tempo}</td>
                  <td className="p-2 sticky right-0 bg-white group-hover:bg-slate-50 transition-colors shadow-[-4px_0_8px_rgba(0,0,0,0.05)]">
                    <button onClick={() => removeRow(route.id)} className="text-slate-300 hover:text-red-600 p-1.5 transition-colors">
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* VINCULATION MODAL (MODO BLACK) */}
      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-xl z-[110] flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-[3rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-300">
                <div className="bg-blue-600 p-8 flex justify-between items-center text-white shadow-2xl">
                    <div className="flex items-center gap-4">
                        <Link size={32} className="bg-white/20 p-2 rounded-2xl" />
                        <div>
                            <h3 className="font-black uppercase tracking-widest text-lg">Mapeamento de Rotas</h3>
                            <p className="text-blue-200 text-xs font-bold mt-1 tracking-tight">Vincule as novas rotas detectadas</p>
                        </div>
                    </div>
                </div>
                
                <div className="p-10">
                    <div className="flex items-center gap-5 p-6 bg-slate-950 border border-blue-500/20 rounded-[2rem] mb-8 text-blue-400">
                        <AlertTriangle size={32} className="shrink-0 animate-pulse" />
                        <p className="text-xs font-black leading-relaxed uppercase tracking-widest">
                            Atenção: Rotas sem operação não podem ser processadas. Vincule agora para automatizar importações futuras.
                        </p>
                    </div>

                    <div className="max-h-[340px] overflow-y-auto space-y-4 mb-10 pr-4 scrollbar-thin">
                        {pendingItems.map((item, idx) => (
                            <div key={idx} className="flex items-center gap-6 p-6 bg-slate-950 border border-slate-800 rounded-[2.5rem] hover:border-blue-500/50 transition-all group">
                                <div className="flex-1">
                                    <span className="text-[9px] text-slate-500 font-black uppercase tracking-[0.2em] block mb-2">Rota Identificada</span>
                                    <div className="font-black text-white text-xl tracking-tighter truncate group-hover:text-blue-400 transition-colors">{item.rota}</div>
                                </div>
                                <div className="w-[55%] relative">
                                    <span className="text-[9px] text-slate-500 font-black uppercase tracking-[0.2em] block mb-2">Vincular Operação</span>
                                    <select 
                                        value={item.operacao} 
                                        onChange={(e) => {
                                            const newPending = [...pendingItems];
                                            newPending[idx].operacao = e.target.value;
                                            setPendingItems(newPending);
                                        }}
                                        className="w-full p-4 bg-slate-900 border-2 border-slate-800 rounded-2xl text-xs font-black text-white outline-none focus:ring-4 focus:ring-blue-500/20 focus:border-blue-500 transition-all appearance-none cursor-pointer"
                                    >
                                        <option value="">Selecione...</option>
                                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                                    </select>
                                    <div className="absolute right-4 bottom-4 pointer-events-none text-slate-600"><ChevronDown size={18}/></div>
                                </div>
                            </div>
                        ))}
                    </div>

                    <button 
                        onClick={handleLinkPending} 
                        disabled={isProcessingImport || pendingItems.some(p => !p.operacao)} 
                        className="w-full py-6 bg-blue-600 text-white font-black uppercase tracking-[0.3em] text-sm rounded-[2rem] shadow-[0_20px_40px_rgba(37,99,235,0.2)] flex items-center justify-center gap-4 transition-all hover:bg-blue-700 active:scale-95 disabled:opacity-50 border-b-8 border-blue-900"
                    >
                        {isProcessingImport ? <Loader2 size={24} className="animate-spin" /> : <><CheckCircle2 size={24} /> Concluir e Salvar Tudo</>}
                    </button>
                    
                    <p className="text-center mt-6 text-[9px] text-slate-600 font-black uppercase tracking-[0.4em]">Design CCO Digital V3.0</p>
                </div>
            </div>
        </div>
      )}

      {/* MODAL IMPORTAÇÃO */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
             <div className="bg-slate-900 border border-slate-800 rounded-[3rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200">
                <div className="bg-emerald-600 p-8 flex justify-between items-center text-white">
                    <div className="flex items-center gap-4">
                        <Upload size={24} className="bg-white/20 p-2 rounded-xl" />
                        <h3 className="font-black uppercase tracking-widest text-sm">Importar Planilha CCO</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-2 rounded-2xl transition-all"><X size={24} /></button>
                </div>
                <div className="p-10">
                    <textarea 
                        value={importText} 
                        onChange={e => setImportText(e.target.value)} 
                        className="w-full h-72 p-6 border-2 border-slate-800 rounded-[2rem] bg-slate-950 text-xs font-mono mb-8 focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 outline-none transition-all text-white shadow-inner" 
                        placeholder="Cole aqui os dados da sua planilha Excel..."
                    />
                    <button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-5 bg-emerald-600 text-white font-black uppercase tracking-widest text-xs rounded-[2rem] shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-emerald-700 disabled:opacity-50 border-b-8 border-emerald-900">
                        {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : <><Sparkles size={20} /> Processar Inteligente</>}
                    </button>
                </div>
             </div>
        </div>
      )}

      {/* MODAL NOVA ROTA */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-[3rem] shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in">
            <div className="bg-blue-600 text-white p-8 flex justify-between items-center shadow-lg">
                <h3 className="font-black uppercase tracking-widest text-sm flex items-center gap-3"><Plus size={24} /> Registro Manual</h3>
                <button onClick={() => setIsModalOpen(false)} className="hover:bg-white/10 p-2 rounded-2xl transition-all"><X size={24} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-10 space-y-6">
                <div className="grid grid-cols-2 gap-6">
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-white text-sm font-bold focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all"/>
                    </div>
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Nº Rota</label>
                        <input type="text" required placeholder="Ex: 24133" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-sm font-black text-blue-400 focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 outline-none transition-all"/>
                    </div>
                </div>
                <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Operação</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-sm font-black text-white focus:ring-4 focus:ring-blue-500/10 outline-none appearance-none cursor-pointer">
                        <option value="">Selecione...</option>
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-6">
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Motorista</label>
                        <input type="text" required placeholder="Nome Completo" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-white text-sm font-bold focus:ring-4 focus:ring-blue-500/10 outline-none transition-all"/>
                    </div>
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-white text-sm font-black focus:ring-4 focus:ring-blue-500/10 outline-none transition-all"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-6">
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Início</label>
                        <input type="text" required placeholder="00:00" onBlur={(e) => setFormData({...formData, inicio: formatTimeInput(e.target.value)})} defaultValue={formData.inicio} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-white text-sm font-mono focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-1">Saída</label>
                        <input type="text" placeholder="00:00" onBlur={(e) => setFormData({...formData, saida: formatTimeInput(e.target.value)})} defaultValue={formData.saida} className="w-full p-4 border-2 border-slate-800 rounded-2xl bg-slate-950 text-white text-sm font-mono focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-6 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-[0.2em] text-xs rounded-[2rem] flex items-center justify-center gap-3 shadow-xl transition-all active:scale-95 border-b-8 border-blue-900 mt-4">
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