
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteDeparture, User, RouteOperationMapping } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDeparturesManual } from '../services/geminiService';
import { 
  Plus, Trash2, Save, Clock, X, Upload, 
  Loader2, RefreshCw, ShieldCheck, FileSpreadsheet,
  AlertTriangle, Link, CheckCircle2, ChevronDown, 
  Filter, Search, Check, CheckSquare, Square,
  Sparkles
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
  
  // Zoom State (Ctrl + Scroll)
  const [zoomLevel, setZoomLevel] = useState(1);

  // Filter States
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});

  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 80,
    rota: 120,
    data: 130,
    inicio: 100,
    motorista: 240,
    placa: 100,
    saida: 100,
    motivo: 140,
    observacao: 280,
    geral: 70,
    aviso: 70,
    operacao: 150,
    status: 90,
    tempo: 90,
  });

  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);
  const filterRef = useRef<HTMLDivElement>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);

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
    };

    const handleClickOutside = (e: MouseEvent) => {
        if (filterRef.current && !filterRef.current.contains(e.target as Node)) {
            setActiveFilterCol(null);
        }
    };

    // Control + Scroll Zoom Handler
    const handleWheel = (e: WheelEvent) => {
        if (e.ctrlKey) {
            e.preventDefault();
            const delta = e.deltaY > 0 ? -0.05 : 0.05;
            setZoomLevel(prev => Math.min(Math.max(prev + delta, 0.5), 1.5));
        }
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('mousedown', handleClickOutside);
    
    const container = tableContainerRef.current;
    if (container) {
        container.addEventListener('wheel', handleWheel, { passive: false });
    }
    
    return () => {
      clearInterval(timer);
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      window.removeEventListener('mousedown', handleClickOutside);
      if (container) container.removeEventListener('wheel', handleWheel);
    };
  }, [currentUser]);

  const startResize = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    resizingRef.current = {
      col,
      startX: e.clientX,
      startWidth: colWidths[col]
    };
  };

  const formatTimeInput = (value: string): string => {
    let clean = value.replace(/[^0-9:]/g, '');
    if (!clean) return '00:00:00';
    
    const parts = clean.split(':');
    let h = parts[0] || '00';
    let m = parts[1] || '00';
    let s = parts[2] || '00';

    if (parts.length === 1) {
        if (h.length === 1) h = '0' + h;
        m = '00'; s = '00';
    } else if (parts.length === 2) {
        if (h.length === 1) h = '0' + h;
        if (m.length === 1) m = '0' + m;
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

  // Fixed potential unknown types from Object.entries by casting the result to specific [string, T][] tuples
  const filteredRoutes = useMemo(() => {
    return routes.filter(r => {
        return (Object.entries(colFilters) as [string, string][]).every(([col, val]) => {
            if (!val) return true;
            const field = r[col as keyof RouteDeparture]?.toString().toLowerCase() || "";
            return field.includes(val.toLowerCase());
        }) && (Object.entries(selectedFilters) as [string, string[]][]).every(([col, vals]) => {
            if (!vals || vals.length === 0) return true;
            const field = r[col as keyof RouteDeparture]?.toString() || "";
            return vals.includes(field);
        });
    });
  }, [routes, colFilters, selectedFilters]);

  const uniqueValuesForCol = (col: string) => {
    const vals = Array.from(new Set(routes.map(r => r[col as keyof RouteDeparture]?.toString() || "")));
    return vals.sort();
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
    
    if (route.saida !== '00:00:00' && isOutOfTolerance) return 'bg-[#FF4500] text-white font-black hover:brightness-110';
    
    const toleranceSec = timeToSeconds(config?.tolerancia || "00:00:00");
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) return 'bg-[#FFD700] text-slate-900 font-black hover:brightness-105';
    
    return 'bg-white text-slate-800 hover:bg-slate-50';
  };

  const FilterDropdown = ({ col }: { col: string }) => {
    const values = uniqueValuesForCol(col);
    const selected = (selectedFilters[col] as string[]) || [];

    const toggleValue = (val: string) => {
        const current = (selectedFilters[col] as string[]) || [];
        if (current.includes(val)) {
            setSelectedFilters({ ...selectedFilters, [col]: current.filter(v => v !== val) });
        } else {
            setSelectedFilters({ ...selectedFilters, [col]: [...current, val] });
        }
    };

    return (
        <div ref={filterRef} className="absolute top-10 left-0 z-50 bg-white dark:bg-slate-900 border border-slate-400 shadow-2xl rounded-xl w-64 p-3 text-slate-800 animate-in fade-in slide-in-from-top-2">
            <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 dark:bg-slate-950 rounded-lg border border-slate-300">
                <Search size={14} className="text-slate-400" />
                <input 
                    type="text" 
                    placeholder="Pesquisar..." 
                    value={colFilters[col] || ""}
                    onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })}
                    className="w-full bg-transparent outline-none text-[10px] font-bold dark:text-white"
                />
            </div>
            <div className="max-h-40 overflow-y-auto space-y-1 mb-3 scrollbar-thin border-t border-b border-slate-100 py-2">
                {values.map(v => (
                    <div 
                        key={v} 
                        onClick={() => toggleValue(v)}
                        className="flex items-center gap-2 p-1.5 hover:bg-blue-50 dark:hover:bg-slate-800 rounded-md cursor-pointer transition-colors"
                    >
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} className="text-slate-300" />}
                        <span className="text-[9px] font-black uppercase truncate dark:text-slate-300">{v || "(vazio)"}</span>
                    </div>
                ))}
            </div>
            <div className="flex gap-2">
                <button onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); }} className="flex-1 py-1.5 text-[9px] font-black uppercase text-red-500 border border-red-100 hover:bg-red-50 rounded-md">Limpar</button>
                <button onClick={() => setActiveFilterCol(null)} className="flex-1 py-1.5 bg-blue-600 text-white text-[9px] font-black uppercase rounded-md shadow-md">Aplicar</button>
            </div>
        </div>
    );
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4 bg-slate-950">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-[10px] uppercase tracking-[0.3em]">Carregando Sistema CCO...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-slate-950 p-4 overflow-hidden select-none">
      {/* HEADER */}
      <div className="flex justify-between items-center mb-4 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-2.5 bg-blue-600 text-white rounded-xl shadow-2xl ring-4 ring-blue-600/10">
            <Clock size={20} />
          </div>
          <div>
            <h2 className="text-xl font-black text-white uppercase tracking-tighter flex items-center gap-3">
              Saída de Rotas
              {isSyncing && <Loader2 size={16} className="animate-spin text-blue-500"/>}
            </h2>
            <div className="flex items-center gap-2">
                <ShieldCheck size={12} className="text-emerald-500"/>
                <p className="text-[9px] text-slate-400 font-black uppercase tracking-widest">Sincronizado: {currentUser.name}</p>
            </div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <div className="hidden lg:flex items-center gap-2 px-3 py-1 bg-slate-900 border border-slate-800 rounded-lg text-[9px] text-slate-500 font-bold uppercase mr-4">
              Zoom: {Math.round(zoomLevel * 100)}% <span className="opacity-50">(Ctrl+Scroll)</span>
          </div>
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-800">
              <RefreshCw size={18} />
          </button>
          <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 font-black shadow-lg border-b-2 border-emerald-900 uppercase text-[9px] tracking-widest transition-all active:scale-95">
            <Upload size={16} /> Importar
          </button>
          <button onClick={() => setIsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-black shadow-lg border-b-2 border-blue-900 uppercase text-[9px] tracking-widest transition-all active:scale-95">
            <Plus size={16} /> Nova Rota
          </button>
        </div>
      </div>

      {/* EXCEL STONE GRID */}
      <div 
        ref={tableContainerRef}
        className="flex-1 overflow-auto bg-white rounded-xl border-2 border-slate-400 shadow-2xl relative scrollbar-thin overflow-x-auto"
      >
        <div style={{ transform: `scale(${zoomLevel})`, transformOrigin: 'top left', width: `${100 / zoomLevel}%` }}>
            <table className="border-collapse table-fixed w-full min-w-max">
              <thead className="sticky top-0 z-40 bg-blue-600 text-white">
                <tr className="border-none h-10">
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
                  ].map(col => {
                    // Use double negation and optional chaining with nullish coalescing to ensure a safe boolean result for ternary ops
                    const hasFilter = !!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0;
                    
                    return (
                      <th 
                        key={col.id} 
                        style={{ width: colWidths[col.id] }}
                        className="relative p-1 border-r border-blue-400 text-[9px] font-black uppercase tracking-widest text-left select-none group"
                      >
                        <div className="flex items-center justify-between px-1.5 h-full">
                          <span className="truncate">{col.label}</span>
                          <button 
                              onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }}
                              className={`p-1 rounded hover:bg-white/20 transition-all ${hasFilter ? 'text-yellow-300' : 'text-white/40'}`}
                          >
                              <Filter size={10} fill={hasFilter ? 'currentColor' : 'none'} />
                          </button>
                        </div>
                        {activeFilterCol === col.id && <FilterDropdown col={col.id} />}
                        <div 
                          onMouseDown={(e) => startResize(e, col.id)}
                          className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10"
                        />
                      </th>
                    );
                  })}
                  <th className="p-2 w-10 sticky right-0 bg-blue-600 border-l border-blue-400"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300">
                {filteredRoutes.map((route) => {
                  const rowClass = getRowClass(route);
                  const isAlert = rowClass.includes('bg-[#FF');
                  
                  // Unified styling for text
                  const inputClass = `w-full h-full bg-transparent outline-none border-none px-2 py-2 text-center text-[10px] font-black uppercase tracking-tight transition-all focus:bg-blue-500/10 ${isAlert ? 'text-white' : 'text-slate-800'}`;

                  return (
                    <tr key={route.id} className={`${rowClass} group transition-all h-9 border-b border-slate-300`}>
                      <td className="p-0 border-r border-slate-300 text-center font-black bg-slate-100/50 text-[10px]">{route.semana}</td>
                      <td className="p-0 border-r border-slate-300">
                        <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={inputClass} />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${inputClass} text-[9px]`} />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <input 
                            type="text" 
                            defaultValue={route.inicio} 
                            onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} 
                            className={`${inputClass} font-mono`} 
                            placeholder="00:00" 
                        />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value.toUpperCase())} className={`${inputClass} text-left px-3`} />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value.toUpperCase())} className={`${inputClass} tracking-widest`} />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <input 
                            type="text" 
                            defaultValue={route.saida} 
                            onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} 
                            className={`${inputClass} font-mono`} 
                            placeholder="00:00" 
                        />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className={`${inputClass} cursor-pointer appearance-none`}>
                          <option value="" className="text-slate-900">Selecione...</option>
                          {['Manutenção', 'Mão de obra', 'Atraso coleta', 'Atraso carregamento', 'Fábrica', 'Infraestrutura', 'Logística', 'Outros'].map(m => (
                            <option key={m} value={m} className="text-slate-900">{m}</option>
                          ))}
                        </select>
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className={`${inputClass} text-left italic px-3`} placeholder="..." />
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className={`${inputClass} appearance-none`}>
                          <option value="OK" className="text-slate-900">OK</option>
                          <option value="NOK" className="text-slate-900">NOK</option>
                        </select>
                      </td>
                      <td className="p-0 border-r border-slate-300">
                        <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className={`${inputClass} appearance-none`}>
                          <option value="SIM" className="text-slate-900">SIM</option>
                          <option value="NÃO" className="text-slate-900">NÃO</option>
                        </select>
                      </td>
                      <td className="p-1 border-r border-slate-300 text-center font-black uppercase text-[8px] truncate">
                          {route.operacao || <span className="text-red-600 animate-pulse underline">VINCULAR</span>}
                      </td>
                      <td className="p-1 border-r border-slate-300 text-center">
                        <span className={`px-2 py-0.5 rounded text-[8px] font-black border ${route.statusOp === 'OK' ? 'bg-emerald-600 border-emerald-700 text-white' : 'bg-red-600 border-red-700 text-white'}`}>
                          {route.statusOp}
                        </span>
                      </td>
                      <td className="p-1 border-r border-slate-300 text-center font-mono font-black text-[9px]">{route.tempo}</td>
                      <td className="p-1 sticky right-0 bg-white group-hover:bg-slate-100 transition-colors shadow-[-4px_0_8px_rgba(0,0,0,0.05)] text-center">
                        <button onClick={() => removeRow(route.id)} className="text-slate-300 hover:text-red-600 p-1 transition-colors">
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
        </div>
      </div>

      {/* COMPACT VINCULATION MODAL */}
      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-slate-950/95 backdrop-blur-2xl z-[110] flex items-center justify-center p-4">
            <div className="bg-slate-900 border border-slate-800 rounded-[2.5rem] shadow-[0_0_100px_rgba(37,99,235,0.2)] w-full max-w-xl max-h-[70vh] overflow-hidden animate-in zoom-in duration-300 flex flex-col">
                <div className="bg-blue-600 p-5 flex justify-between items-center text-white shrink-0 shadow-lg">
                    <div className="flex items-center gap-3">
                        <Link size={20} className="bg-white/20 p-1.5 rounded-lg" />
                        <div>
                            <h3 className="font-black uppercase tracking-widest text-xs">Vínculo de Operação</h3>
                            <p className="text-blue-200 text-[9px] font-bold tracking-tight">Novos registros pendentes</p>
                        </div>
                    </div>
                </div>
                
                <div className="p-6 flex-1 overflow-y-auto space-y-3 scrollbar-thin">
                    <div className="flex items-center gap-3 p-3 bg-blue-500/10 border border-blue-500/20 rounded-xl text-blue-400">
                        <AlertTriangle size={20} className="shrink-0 animate-pulse" />
                        <p className="text-[9px] font-black uppercase tracking-widest">Rotas sem operação detectadas. Vincule agora.</p>
                    </div>

                    {pendingItems.map((item, idx) => (
                        <div key={idx} className="flex items-center gap-3 p-4 bg-slate-950 border border-slate-800 rounded-2xl group hover:border-blue-500/30 transition-all">
                            <div className="flex-1">
                                <span className="text-[8px] text-slate-500 font-black uppercase tracking-widest block mb-1">Identificador Rota</span>
                                <div className="font-black text-white text-base tracking-tighter truncate group-hover:text-blue-400">{item.rota}</div>
                            </div>
                            <div className="w-[45%] relative">
                                <select 
                                    value={item.operacao} 
                                    onChange={(e) => {
                                        const newPending = [...pendingItems];
                                        newPending[idx].operacao = e.target.value;
                                        setPendingItems(newPending);
                                    }}
                                    className="w-full p-2.5 bg-slate-900 border-2 border-slate-800 rounded-lg text-[10px] font-black text-white outline-none focus:border-blue-500 appearance-none cursor-pointer"
                                >
                                    <option value="">Selecione...</option>
                                    {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                                </select>
                            </div>
                        </div>
                    ))}
                </div>

                <div className="p-6 bg-slate-900 border-t border-slate-800 shrink-0">
                    <button 
                        onClick={handleLinkPending} 
                        disabled={isProcessingImport || pendingItems.some(p => !p.operacao)} 
                        className="w-full py-4 bg-blue-600 text-white font-black uppercase tracking-[0.3em] text-[10px] rounded-xl shadow-2xl transition-all hover:bg-blue-700 active:scale-95 disabled:opacity-50 border-b-4 border-blue-900"
                    >
                        {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : "Confirmar Vínculos"}
                    </button>
                </div>
            </div>
        </div>
      )}

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
             <div className="bg-slate-900 border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200">
                <div className="bg-emerald-600 p-6 flex justify-between items-center text-white">
                    <div className="flex items-center gap-3">
                        <Upload size={20} className="bg-white/20 p-1.5 rounded-lg" />
                        <h3 className="font-black uppercase tracking-widest text-xs">Importar CCO Excel</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button>
                </div>
                <div className="p-8">
                    <textarea 
                        value={importText} 
                        onChange={e => setImportText(e.target.value)} 
                        className="w-full h-64 p-5 border-2 border-slate-800 rounded-2xl bg-slate-950 text-[10px] font-mono mb-6 focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 outline-none transition-all text-white shadow-inner scrollbar-thin" 
                        placeholder="Cole os dados aqui (Rota | Data | Início...)"
                    />
                    <button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-4 bg-emerald-600 text-white font-black uppercase tracking-widest text-[10px] rounded-xl shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-emerald-700 disabled:opacity-50 border-b-4 border-emerald-900">
                        {isProcessingImport ? <Loader2 size={18} className="animate-spin" /> : <Sparkles size={16} />} <span>Processar Dados</span>
                    </button>
                </div>
             </div>
        </div>
      )}

      {/* MANUAL ENTRY MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in">
            <div className="bg-blue-600 text-white p-6 flex justify-between items-center shadow-lg">
                <h3 className="font-black uppercase tracking-widest text-xs flex items-center gap-3"><Plus size={20} /> Registro Manual</h3>
                <button onClick={() => setIsModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-8 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-white text-[11px] font-bold outline-none focus:border-blue-500 transition-all"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Rota</label>
                        <input type="text" required placeholder="Nº Rota" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-[11px] font-black text-blue-400 outline-none focus:border-blue-500 transition-all"/>
                    </div>
                </div>
                <div className="space-y-1">
                    <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Operação</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-[11px] font-black text-white outline-none appearance-none cursor-pointer focus:border-blue-500">
                        <option value="">Selecione...</option>
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Motorista</label>
                        <input type="text" required placeholder="Nome" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-white text-[11px] font-bold outline-none focus:border-blue-500 transition-all"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-white text-[11px] font-black outline-none focus:border-blue-500 transition-all"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Início</label>
                        <input type="text" required placeholder="00:00" onBlur={(e) => setFormData({...formData, inicio: formatTimeInput(e.target.value)})} defaultValue={formData.inicio} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-white text-[11px] font-mono outline-none focus:border-blue-500"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Saída</label>
                        <input type="text" placeholder="00:00" onBlur={(e) => setFormData({...formData, saida: formatTimeInput(e.target.value)})} defaultValue={formData.saida} className="w-full p-3 border border-slate-800 rounded-xl bg-slate-950 text-white text-[11px] font-mono outline-none focus:border-blue-500"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-[0.2em] text-[10px] rounded-xl flex items-center justify-center gap-2 shadow-xl transition-all active:scale-95 border-b-4 border-blue-900 mt-4">
                    {isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />} SALVAR NO SHAREPOINT
                </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
