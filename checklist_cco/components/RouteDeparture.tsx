
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

  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  // Excel-like Resizing state
  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 80,
    rota: 110,
    data: 140,
    inicio: 110,
    motorista: 260,
    placa: 110,
    saida: 110,
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

  // Time formatting logic: "15" -> "15:00:00", "15:30" -> "15:30:00"
  const formatTimeInput = (value: string): string => {
    let clean = value.replace(/[^0-9:]/g, '');
    if (!clean) return '00:00:00';
    
    const parts = clean.split(':');
    let h = parts[0] || '00';
    let m = parts[1] || '00';
    let s = parts[2] || '00';

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
                ...p, id: '', semana: calculateWeekString(p.data || ''), statusGeral: 'OK', aviso: 'NÃO',
                statusOp: status, tempo: gap, createdAt: new Date().toISOString()
            } as RouteDeparture;
            return SharePointService.updateDeparture(token, r);
        }));

        await loadData();
        setIsLinkModalOpen(false);
        setPendingItems([]);
        alert("Vínculos criados e rotas salvas com sucesso.");
    } catch (err: any) {
        alert("Erro ao salvar: " + err.message);
    } finally {
        setIsProcessingImport(false);
    }
  };

  // Fixed missing handleSubmit function for manual route registration form
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = getAccessToken();
    if (!token) return;

    setIsSyncing(true);
    try {
        const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === formData.operacao?.toUpperCase().trim());
        const { gap, status } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia);
        
        const newRoute: RouteDeparture = {
            ...formData,
            id: '', // Empty ID indicates a new entry
            semana: calculateWeekString(formData.data || ''),
            statusOp: status,
            tempo: gap,
            statusGeral: formData.statusGeral || 'OK',
            aviso: formData.aviso || 'NÃO',
            createdAt: new Date().toISOString()
        } as RouteDeparture;

        const newId = await SharePointService.updateDeparture(token, newRoute);
        
        // Optimistically update the local list with the new entry and its returned ID
        setRoutes(prev => [{ ...newRoute, id: newId }, ...prev]);
        setIsModalOpen(false);
        
        // Reset form data for next use
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
        alert("Rota registrada com sucesso no SharePoint!");
    } catch (err: any) {
        console.error("Manual submission failed:", err);
        alert("Erro ao salvar a rota: " + err.message);
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

  const getRowColor = (route: RouteDeparture) => {
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === route.operacao.toUpperCase().trim());
    const { isOutOfTolerance } = calculateGap(route.inicio, route.saida, config?.tolerancia);
    if (route.saida !== '00:00:00' && isOutOfTolerance) return 'bg-orange-50 border-orange-100 text-orange-900';
    const toleranceSec = timeToSeconds(config?.tolerancia || "00:00:00");
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) return 'bg-yellow-50 border-yellow-100 text-yellow-900';
    return 'bg-white text-slate-700';
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-[10px] uppercase tracking-[0.2em]">Sincronizando Banco de Dados...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-slate-50 p-4 overflow-hidden select-none">
      {/* HEADER */}
      <div className="flex justify-between items-center mb-6 shrink-0">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-blue-600 text-white rounded-2xl shadow-lg ring-4 ring-blue-500/10">
            <Clock size={24} />
          </div>
          <div>
            <h2 className="text-2xl font-black text-slate-800 uppercase tracking-tight flex items-center gap-3">
              Saída de Rotas
              {isSyncing && <Loader2 size={18} className="animate-spin text-blue-500"/>}
            </h2>
            <div className="flex items-center gap-2">
                <ShieldCheck size={14} className="text-emerald-500"/>
                <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Sincronizado: {currentUser.name}</p>
            </div>
          </div>
        </div>
        <div className="flex gap-3">
          <button onClick={loadData} className="p-2.5 text-slate-400 hover:text-blue-600 hover:bg-white rounded-xl transition-all border border-transparent hover:border-slate-200">
              <RefreshCw size={20} />
          </button>
          <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-5 py-2.5 bg-emerald-500 text-white rounded-xl hover:bg-emerald-600 font-black shadow-lg transition-all active:scale-95 border-b-4 border-emerald-700 uppercase text-xs tracking-widest">
            <Upload size={18} /> Importar Excel
          </button>
          <button onClick={() => setIsModalOpen(true)} className="flex items-center gap-2 px-5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl transition-all shadow-lg active:scale-95 border-b-4 border-blue-800 uppercase text-xs tracking-widest">
            <Plus size={18} /> Nova Rota
          </button>
        </div>
      </div>

      {/* EXCEL GRID */}
      <div className="flex-1 overflow-auto bg-white rounded-2xl border border-slate-200 shadow-xl relative scrollbar-thin">
        <table className="border-collapse table-fixed w-full min-w-max bg-white">
          <thead className="sticky top-0 z-20 bg-slate-50">
            <tr className="border-b border-slate-200">
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
                  className="relative p-3 text-[10px] font-black text-slate-500 uppercase tracking-widest text-center select-none border-r border-slate-200/50 group"
                >
                  {col.label}
                  <div 
                    onMouseDown={(e) => startResize(e, col.id)}
                    className="absolute right-0 top-0 bottom-0 w-1.5 cursor-col-resize hover:bg-blue-500 transition-colors z-10 flex items-center justify-center"
                  >
                     <div className="w-[1px] h-4 bg-slate-200 group-hover:bg-blue-300" />
                  </div>
                </th>
              ))}
              <th className="p-3 w-12 sticky right-0 bg-slate-50"></th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {routes.map((route) => {
              const rowColor = getRowColor(route);
              const isAlerted = rowColor.includes('bg-orange') || rowColor.includes('bg-yellow');
              
              const inputStyle = `w-full h-full bg-transparent outline-none border-none p-3 text-center transition-all focus:bg-blue-500/5 ${isAlerted ? 'font-black' : 'font-medium'}`;

              return (
                <tr key={route.id} className={`${rowColor} group hover:bg-blue-50/20 transition-all h-11 text-[11px]`}>
                  <td className="p-0 border-r border-slate-100 text-center font-bold bg-slate-50/20">{route.semana}</td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={inputStyle} />
                  </td>
                  <td className="p-0 border-r border-slate-100 text-center">
                    <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${inputStyle} text-[10px]`} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.inicio} onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} defaultValue={route.inicio} className={`${inputStyle} font-mono`} placeholder="00:00" />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value.toUpperCase())} className={`${inputStyle} text-left uppercase px-4`} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value.toUpperCase())} className={`${inputStyle} font-black uppercase tracking-widest`} />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.saida} onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} defaultValue={route.saida} className={`${inputStyle} font-mono`} placeholder="00:00" />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full h-full bg-transparent outline-none px-2 text-center cursor-pointer appearance-none">
                      <option value="">Selecione...</option>
                      {['Manutenção', 'Mão de obra', 'Atraso coleta', 'Atraso carregamento', 'Fábrica', 'Infraestrutura', 'Logística', 'Outros'].map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className={`${inputStyle} text-left italic px-4`} placeholder="Observações..." />
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full bg-transparent outline-none font-bold text-center appearance-none">
                      <option value="OK">OK</option>
                      <option value="NOK">NOK</option>
                    </select>
                  </td>
                  <td className="p-0 border-r border-slate-100">
                    <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className="w-full h-full bg-transparent outline-none font-bold text-center appearance-none">
                      <option value="SIM">SIM</option>
                      <option value="NÃO">NÃO</option>
                    </select>
                  </td>
                  <td className="p-2 border-r border-slate-100 text-center font-black uppercase text-[9px] truncate">{route.operacao}</td>
                  <td className="p-2 border-r border-slate-100 text-center">
                    <span className={`px-2 py-0.5 rounded text-[9px] font-black ${route.statusOp === 'OK' ? 'bg-emerald-500 text-white' : 'bg-orange-500 text-white'}`}>
                      {route.statusOp}
                    </span>
                  </td>
                  <td className="p-2 border-r border-slate-100 text-center font-mono font-bold">{route.tempo}</td>
                  <td className="p-2 sticky right-0 bg-white group-hover:bg-slate-50 transition-colors shadow-[-4px_0_8px_rgba(0,0,0,0.02)]">
                    <button onClick={() => removeRow(route.id)} className="text-slate-300 hover:text-red-500 p-1.5 transition-colors">
                      <Trash2 size={16} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* VINCULATION MODAL (CLEAN DESIGN) */}
      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[110] flex items-center justify-center p-4">
            <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 animate-in zoom-in duration-300">
                <div className="bg-amber-600 p-8 flex justify-between items-center text-white shadow-xl relative overflow-hidden">
                    <div className="absolute top-0 right-0 p-8 opacity-10">
                        <Link size={120} />
                    </div>
                    <div className="flex items-center gap-4 relative z-10">
                        <div className="p-3 bg-white/20 rounded-2xl backdrop-blur-md">
                            <Link size={28} />
                        </div>
                        <div>
                            <h3 className="font-black uppercase tracking-widest text-lg">Mapeamento de Rotas</h3>
                            <p className="text-amber-100 text-xs font-bold mt-1">Vincule as novas rotas às operações CCO</p>
                        </div>
                    </div>
                </div>
                
                <div className="p-10">
                    <div className="flex items-center gap-4 p-5 bg-amber-50 border border-amber-100 rounded-3xl mb-8 text-amber-800">
                        <AlertTriangle size={28} className="shrink-0" />
                        <p className="text-xs font-black leading-relaxed uppercase tracking-tight">
                            Atenção: Estas rotas não foram reconhecidas. Vincule cada uma a uma operação válida para salvar no SharePoint e automatizar importações futuras.
                        </p>
                    </div>

                    <div className="max-h-[340px] overflow-y-auto space-y-4 mb-10 pr-4 scrollbar-thin">
                        {pendingItems.map((item, idx) => (
                            <div key={idx} className="flex items-center gap-6 p-6 bg-slate-50 rounded-[2rem] border border-slate-100 hover:border-blue-200 transition-all shadow-sm">
                                <div className="flex-1">
                                    <span className="text-[9px] text-slate-400 font-black uppercase tracking-widest block mb-2">Rota Identificada</span>
                                    <div className="font-black text-blue-600 text-lg truncate tracking-tight">{item.rota}</div>
                                </div>
                                <div className="w-[55%]">
                                    <span className="text-[9px] text-slate-400 font-black uppercase tracking-widest block mb-2">Operação de Destino</span>
                                    <select 
                                        value={item.operacao} 
                                        onChange={(e) => {
                                            const newPending = [...pendingItems];
                                            newPending[idx].operacao = e.target.value;
                                            setPendingItems(newPending);
                                        }}
                                        className="w-full p-4 bg-white border-2 border-slate-100 rounded-2xl text-xs font-black outline-none focus:ring-4 focus:ring-blue-500/10 focus:border-blue-400 shadow-sm transition-all appearance-none cursor-pointer"
                                    >
                                        <option value="">Selecione...</option>
                                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                                    </select>
                                </div>
                            </div>
                        ))}
                    </div>

                    <button 
                        onClick={handleLinkPending} 
                        disabled={isProcessingImport || pendingItems.some(p => !p.operacao)} 
                        className="w-full py-5 bg-blue-600 text-white font-black uppercase tracking-widest text-sm rounded-[2rem] shadow-xl flex items-center justify-center gap-4 transition-all hover:bg-blue-700 active:scale-[0.98] disabled:opacity-50 border-b-8 border-blue-800"
                    >
                        {isProcessingImport ? <Loader2 size={24} className="animate-spin" /> : <><CheckCircle2 size={24} /> Concluir e Salvar Tudo</>}
                    </button>
                    
                    <p className="text-center mt-6 text-[10px] text-slate-400 font-black uppercase tracking-widest">Atenção: Não é permitido ignorar este passo para manter a integridade do banco de dados.</p>
                </div>
            </div>
        </div>
      )}

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[100] flex items-center justify-center p-4">
             <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-200 animate-in zoom-in duration-300">
                <div className="bg-emerald-600 p-8 flex justify-between items-center text-white shadow-lg">
                    <div className="flex items-center gap-4">
                        <Upload size={24} />
                        <h3 className="font-black uppercase tracking-widest text-sm">Importar Planilha Excel</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/20 p-2 rounded-2xl transition-all"><X size={24} /></button>
                </div>
                <div className="p-10">
                    <div className="mb-6 text-[10px] font-black text-slate-400 uppercase tracking-widest flex items-center gap-2">
                        <FileSpreadsheet size={16} className="text-blue-500"/> Copie do Excel (ROTA | DATA | INÍCIO | MOTORISTA | PLACA | SAÍDA)
                    </div>
                    <textarea 
                        value={importText} 
                        onChange={e => setImportText(e.target.value)} 
                        className="w-full h-72 p-6 border-2 border-slate-100 rounded-[2rem] bg-slate-50 text-xs font-mono mb-8 focus:ring-4 focus:ring-blue-500/10 outline-none transition-all shadow-inner" 
                        placeholder="Cole aqui os dados..."
                    />
                    <button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-5 bg-slate-900 text-white font-black uppercase tracking-widest text-xs rounded-[2rem] shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-slate-800 disabled:opacity-50 border-b-8 border-slate-950">
                        {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : <><SparklesIcon /> Processar Dados da Planilha</>}
                    </button>
                </div>
             </div>
        </div>
      )}

      {/* MANUAL ENTRY MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden border border-slate-200 animate-in zoom-in">
            <div className="bg-blue-600 text-white p-8 flex justify-between items-center shadow-lg">
                <h3 className="font-black uppercase tracking-widest text-sm flex items-center gap-3"><Plus size={20} /> Registrar Rota Manual</h3>
                <button onClick={() => setIsModalOpen(false)} className="hover:bg-white/20 p-2 rounded-2xl transition-all"><X size={24} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-10 space-y-6">
                <div className="grid grid-cols-2 gap-6">
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-bold focus:ring-4 focus:ring-blue-500/10 outline-none transition-all"/>
                    </div>
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Rota</label>
                        <input type="text" required placeholder="Nº Rota" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-black text-blue-600 focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                </div>
                <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Operação CCO</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-black focus:ring-4 focus:ring-blue-500/10 outline-none appearance-none cursor-pointer">
                        <option value="">Selecione...</option>
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-6">
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Motorista</label>
                        <input type="text" required placeholder="Nome" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-bold focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-black focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-6">
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Horário Previsto</label>
                        <input type="text" required placeholder="00:00:00" onBlur={(e) => setFormData({...formData, inicio: formatTimeInput(e.target.value)})} defaultValue={formData.inicio} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-mono focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                    <div className="space-y-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest ml-1">Horário Saída</label>
                        <input type="text" placeholder="00:00:00" onBlur={(e) => setFormData({...formData, saida: formatTimeInput(e.target.value)})} defaultValue={formData.saida} className="w-full p-4 border-2 border-slate-100 rounded-2xl bg-slate-50 text-sm font-mono focus:ring-4 focus:ring-blue-500/10 outline-none"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-5 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-widest text-xs rounded-[2rem] flex items-center justify-center gap-3 shadow-xl transition-all active:scale-95 border-b-8 border-blue-800 mt-4">
                    {isSyncing ? <Loader2 size={18} className="animate-spin" /> : <Save size={18} />} GRAVAR NO SHAREPOINT
                </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

const SparklesIcon = () => (
    <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
        <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"/>
    </svg>
);

export default RouteDepartureView;
