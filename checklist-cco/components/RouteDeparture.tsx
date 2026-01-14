
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteDeparture, User, RouteOperationMapping } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDeparturesManual } from '../services/geminiService';
import { 
  Plus, Trash2, Save, Clock, X, Upload, 
  Loader2, RefreshCw, ShieldCheck,
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
  
  // Zoom State
  const [zoomLevel, setZoomLevel] = useState(0.9); // Começa um pouco menor para caber mais dados

  // Filter States
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});

  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    semana: 80,
    rota: 120,
    data: 120,
    inicio: 95,
    motorista: 230,
    placa: 100,
    saida: 95,
    motivo: 150,
    observacao: 280,
    geral: 70,
    aviso: 70,
    operacao: 140,
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
    if (!token) return;
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

    const handleWheel = (e: WheelEvent) => {
        if (e.ctrlKey) {
            e.preventDefault();
            const delta = e.deltaY > 0 ? -0.05 : 0.05;
            setZoomLevel(prev => Math.min(Math.max(prev + delta, 0.4), 1.3));
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
    let h = (parts[0] || '00').padStart(2, '0').substring(0, 2);
    let m = (parts[1] || '00').padStart(2, '0').substring(0, 2);
    let s = (parts[2] || '00').padStart(2, '0').substring(0, 2);
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
    
    if (field === 'inicio' || field === 'saida' || field === 'operacao') {
        const { gap, status } = calculateGap(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00");
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
                const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia || "00:00:00");
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
        alert("Selecione a operação para todas as rotas.");
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
            const { gap, status } = calculateGap(p.inicio || '00:00:00', p.saida || '00:00:00', config?.tolerancia || "00:00:00");
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
        const { gap, status } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia || "00:00:00");
        
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

  const getAlertStyles = (route: RouteDeparture) => {
    const config = userConfigs.find(c => c.operacao.toUpperCase().trim() === route.operacao.toUpperCase().trim());
    // Fix: Explicitly handle potential undefined/unknown by providing a fallback string and casting.
    const tolerance = (config?.tolerancia as string) || "00:00:00";
    const { isOutOfTolerance } = calculateGap(route.inicio, route.saida, tolerance);
    
    // Crítico (Vermelho) se estiver fora da tolerância e com saída registrada
    if (route.saida !== '00:00:00' && isOutOfTolerance) {
        return "border-l-4 border-[#F75A68] bg-[#F75A68]/10";
    }
    
    // Alerta (Laranja) se ainda não saiu e passou do horário planejado + tolerância
    // Fix: Use the same tolerance string constant or variable to ensure type safety.
    const toleranceSec = timeToSeconds(tolerance);
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(route.inicio);
    
    if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) {
        return "border-l-4 border-[#FF9000] bg-[#FF9000]/10";
    }
    
    return "border-l-4 border-transparent";
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
        <div ref={filterRef} className="absolute top-10 left-0 z-50 bg-[#1e1e24] border border-slate-700 shadow-2xl rounded-xl w-64 p-3 text-slate-200 animate-in fade-in slide-in-from-top-2">
            <div className="flex items-center gap-2 mb-3 p-2 bg-[#121214] rounded-lg border border-slate-800">
                <Search size={14} className="text-slate-500" />
                <input 
                    type="text" 
                    placeholder="Pesquisar..." 
                    value={colFilters[col] || ""}
                    onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })}
                    className="w-full bg-transparent outline-none text-[10px] font-bold text-white"
                />
            </div>
            <div className="max-h-40 overflow-y-auto space-y-1 mb-3 scrollbar-thin border-t border-b border-slate-800 py-2">
                {values.map(v => (
                    <div 
                        key={v} 
                        onClick={() => toggleValue(v)}
                        className="flex items-center gap-2 p-1.5 hover:bg-slate-800 rounded-md cursor-pointer transition-colors"
                    >
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-500" /> : <Square size={14} className="text-slate-600" />}
                        <span className="text-[9px] font-black uppercase truncate text-slate-300">{v || "(vazio)"}</span>
                    </div>
                ))}
            </div>
            <div className="flex gap-2">
                <button onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); }} className="flex-1 py-1.5 text-[9px] font-black uppercase text-red-400 border border-red-900/50 hover:bg-red-900/20 rounded-md">Limpar</button>
                <button onClick={() => setActiveFilterCol(null)} className="flex-1 py-1.5 bg-blue-600 text-white text-[9px] font-black uppercase rounded-md shadow-md">Aplicar</button>
            </div>
        </div>
    );
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4 bg-[#121214]">
        <Loader2 size={40} className="animate-spin" />
        <p className="font-bold animate-pulse text-[10px] uppercase tracking-[0.3em]">Sincronizando CCO Digital...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-[#020617] p-4 overflow-hidden select-none">
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
                <p className="text-[9px] text-slate-400 font-black uppercase tracking-widest">CCO Logística: {currentUser.name}</p>
            </div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <div className="hidden lg:flex items-center gap-2 px-3 py-1 bg-[#121214] border border-slate-800 rounded-lg text-[9px] text-slate-500 font-bold uppercase mr-4">
              Zoom: {Math.round(zoomLevel * 100)}% <span className="opacity-50 text-[8px]">(Ctrl + Scroll)</span>
          </div>
          <button onClick={loadData} className="p-2 text-slate-500 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-800">
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

      {/* MODERN DARK DATA GRID */}
      <div 
        ref={tableContainerRef}
        className="flex-1 overflow-auto bg-[#121214] rounded-xl border-2 border-[#1e1e24] shadow-2xl relative scrollbar-thin overflow-x-auto"
      >
        <div style={{ transform: `scale(${zoomLevel})`, transformOrigin: 'top left', width: `${100 / zoomLevel}%` }}>
            <table className="border-collapse table-fixed w-full min-w-max">
              <thead className="sticky top-0 z-40 bg-blue-700 text-white shadow-lg">
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
                    const hasFilter = !!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0;
                    return (
                      <th 
                        key={col.id} 
                        style={{ width: colWidths[col.id] }}
                        className="relative p-1 border-r border-blue-600/50 text-[9px] font-black uppercase tracking-widest text-left select-none"
                      >
                        <div className="flex items-center justify-between px-1.5 h-full">
                          <span className="truncate">{col.label}</span>
                          <button 
                              onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }}
                              className={`p-1 rounded hover:bg-white/20 transition-all ${hasFilter ? 'text-yellow-400' : 'text-white/40'}`}
                          >
                              <Filter size={10} fill={hasFilter ? 'currentColor' : 'none'} />
                          </button>
                        </div>
                        {activeFilterCol === col.id && <FilterDropdown col={col.id} />}
                        <div onMouseDown={(e) => startResize(e, col.id)} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                      </th>
                    );
                  })}
                  <th className="p-2 w-10 sticky right-0 bg-blue-700 border-l border-blue-600/50"></th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/30">
                {filteredRoutes.map((route, idx) => {
                  const alertClasses = getAlertStyles(route);
                  const isEven = idx % 2 === 0;
                  
                  const baseInputClass = "w-full h-full bg-transparent outline-none border-none px-2 py-2 text-center text-[10px] font-medium transition-all focus:bg-blue-500/5 text-[#E1E1E6]";
                  const monoInputClass = "w-full h-full bg-transparent outline-none border-none px-2 py-2 text-center text-[10px] font-mono transition-all focus:bg-blue-500/5 text-[#E1E1E6]";
                  const textLeftClass = "w-full h-full bg-transparent outline-none border-none px-3 py-2 text-left text-[10px] font-semibold transition-all focus:bg-blue-500/5 text-[#E1E1E6]";

                  return (
                    <tr key={route.id} className={`${isEven ? 'bg-[#121214]' : 'bg-[#18181b]'} ${alertClasses} group transition-all h-9 hover:bg-blue-900/10`}>
                      <td className="p-0 border-r border-[#1e1e24] text-center font-bold text-[#8D8D99] text-[9px]">{route.semana}</td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={`${textLeftClass} font-bold`} />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${monoInputClass} text-[9px]`} />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input 
                            type="text" 
                            value={route.inicio} 
                            onChange={(e) => {
                                const val = e.target.value;
                                setRoutes(prev => prev.map(r => r.id === route.id ? { ...r, inicio: val } : r));
                            }}
                            onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} 
                            className={monoInputClass} 
                            placeholder="00:00:00" 
                        />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value.toUpperCase())} className={textLeftClass} />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value.toUpperCase())} className={`${monoInputClass} tracking-widest font-bold`} />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input 
                            type="text" 
                            value={route.saida} 
                            onChange={(e) => {
                                const val = e.target.value;
                                setRoutes(prev => prev.map(r => r.id === route.id ? { ...r, saida: val } : r));
                            }}
                            onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} 
                            className={monoInputClass} 
                            placeholder="00:00:00" 
                        />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <div className="flex items-center justify-center h-full px-2">
                            <select 
                                value={route.motivo} 
                                onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} 
                                className="w-full bg-[#1e1e24] border border-slate-800 rounded px-1 py-0.5 text-[9px] font-black uppercase text-slate-300 outline-none cursor-pointer hover:border-slate-600 appearance-none text-center"
                            >
                                <option value="">SELECIONE...</option>
                                {['Manutenção', 'Mão de obra', 'Atraso coleta', 'Atraso carregamento', 'Fábrica', 'Infraestrutura', 'Logística', 'Outros'].map(m => (
                                    <option key={m} value={m}>{m.toUpperCase()}</option>
                                ))}
                            </select>
                        </div>
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <input type="text" value={route.observacao} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} className={`${textLeftClass} italic text-slate-500 font-normal truncate`} placeholder="..." />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className={`${baseInputClass} font-black appearance-none`}>
                          <option value="OK">OK</option>
                          <option value="NOK">NOK</option>
                        </select>
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className={`${baseInputClass} font-black appearance-none`}>
                          <option value="SIM">SIM</option>
                          <option value="NÃO">NÃO</option>
                        </select>
                      </td>
                      <td className="p-1 border-r border-[#1e1e24] text-center font-black uppercase text-[8px] truncate text-[#8D8D99]">
                          {route.operacao || <span className="text-[#F75A68] animate-pulse underline">VINCULAR</span>}
                      </td>
                      <td className="p-1 border-r border-[#1e1e24] text-center">
                        <span className={`px-2 py-0.5 rounded-[4px] text-[8px] font-black border ${route.statusOp === 'OK' ? 'bg-emerald-900/30 border-emerald-800 text-emerald-400' : 'bg-red-900/30 border-red-800 text-red-400'}`}>
                          {route.statusOp}
                        </span>
                      </td>
                      <td className="p-1 border-r border-[#1e1e24] text-center font-mono font-bold text-[9px] text-[#8D8D99]">{route.tempo}</td>
                      <td className="p-1 sticky right-0 bg-[#121214] group-hover:bg-[#18181b] transition-colors shadow-[-4px_0_12px_rgba(0,0,0,0.5)] text-center">
                        <button onClick={() => removeRow(route.id)} className="text-slate-600 hover:text-red-400 p-1.5 transition-colors">
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

      {/* COMPACT DARK VINCULATION MODAL */}
      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-xl z-[110] flex items-center justify-center p-4">
            <div className="bg-[#121214] border border-slate-800 rounded-3xl shadow-[0_0_100px_rgba(37,99,235,0.1)] w-full max-w-lg max-h-[80vh] overflow-hidden animate-in zoom-in duration-300 flex flex-col">
                <div className="bg-blue-700 p-6 flex justify-between items-center text-white shrink-0">
                    <div className="flex items-center gap-3">
                        <Link size={24} className="bg-white/20 p-2 rounded-xl" />
                        <div>
                            <h3 className="font-black uppercase tracking-widest text-xs">Ajuste de Operação</h3>
                            <p className="text-blue-100 text-[10px] font-bold tracking-tight">Rotas pendentes de vinculação</p>
                        </div>
                    </div>
                </div>
                
                <div className="p-6 flex-1 overflow-y-auto space-y-3 scrollbar-thin">
                    <div className="flex items-center gap-3 p-4 bg-red-950/20 border border-red-900/30 rounded-2xl text-red-400">
                        <AlertTriangle size={20} className="shrink-0" />
                        <p className="text-[10px] font-black uppercase tracking-widest leading-relaxed">As rotas abaixo não possuem operação vinculada e não serão processadas se não forem corrigidas.</p>
                    </div>

                    {pendingItems.map((item, idx) => (
                        <div key={idx} className="flex items-center gap-4 p-4 bg-[#18181b] border border-slate-800 rounded-2xl group">
                            <div className="flex-1">
                                <span className="text-[8px] text-slate-500 font-black uppercase tracking-widest block mb-1">Rota Identificada</span>
                                <div className="font-black text-white text-base tracking-tighter truncate group-hover:text-blue-500">{item.rota}</div>
                            </div>
                            <div className="w-[45%]">
                                <select 
                                    value={item.operacao} 
                                    onChange={(e) => {
                                        const newPending = [...pendingItems];
                                        newPending[idx].operacao = e.target.value;
                                        setPendingItems(newPending);
                                    }}
                                    className="w-full p-2.5 bg-[#121214] border border-slate-700 rounded-lg text-[10px] font-black text-white outline-none focus:border-blue-600 appearance-none cursor-pointer"
                                >
                                    <option value="">Selecione...</option>
                                    {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                                </select>
                            </div>
                        </div>
                    ))}
                </div>

                <div className="p-6 bg-[#18181b] border-t border-slate-800 shrink-0">
                    <button 
                        onClick={handleLinkPending} 
                        disabled={isProcessingImport || pendingItems.some(p => !p.operacao)} 
                        className="w-full py-4 bg-blue-600 text-white font-black uppercase tracking-[0.3em] text-[10px] rounded-xl shadow-2xl transition-all hover:bg-blue-700 active:scale-95 disabled:opacity-50 border-b-4 border-blue-900"
                    >
                        {isProcessingImport ? <Loader2 size={20} className="animate-spin" /> : "Gravar Vínculos"}
                    </button>
                </div>
            </div>
        </div>
      )}

      {/* IMPORT MODAL */}
      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
             <div className="bg-[#121214] border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200">
                <div className="bg-emerald-600 p-6 flex justify-between items-center text-white">
                    <div className="flex items-center gap-3">
                        <Upload size={20} className="bg-white/20 p-1.5 rounded-lg" />
                        <h3 className="font-black uppercase tracking-widest text-xs">Importação Rápida Excel</h3>
                    </div>
                    <button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button>
                </div>
                <div className="p-8">
                    <textarea 
                        value={importText} 
                        onChange={e => setImportText(e.target.value)} 
                        className="w-full h-64 p-5 border-2 border-slate-800 rounded-2xl bg-[#020617] text-[10px] font-mono mb-6 focus:ring-2 focus:ring-emerald-500 outline-none transition-all text-white shadow-inner scrollbar-thin" 
                        placeholder="Cole os dados aqui..."
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
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
          <div className="bg-[#121214] border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in">
            <div className="bg-blue-700 text-white p-6 flex justify-between items-center shadow-lg">
                <h3 className="font-black uppercase tracking-widest text-xs flex items-center gap-3"><Plus size={20} /> Novo Registro</h3>
                <button onClick={() => setIsModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button>
            </div>
            <form onSubmit={handleSubmit} className="p-8 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Data</label>
                        <input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-bold outline-none focus:border-blue-600 transition-all"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Rota</label>
                        <input type="text" required placeholder="Ex: 24001D" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-[11px] font-black text-blue-400 outline-none focus:border-blue-600 transition-all"/>
                    </div>
                </div>
                <div className="space-y-1">
                    <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Operação</label>
                    <select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-[11px] font-black text-white outline-none appearance-none cursor-pointer focus:border-blue-600">
                        <option value="">Selecione...</option>
                        {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                    </select>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Motorista</label>
                        <input type="text" required placeholder="Nome Completo" value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-bold outline-none focus:border-blue-600 transition-all"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Placa</label>
                        <input type="text" required placeholder="XXX-0000" value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-black outline-none focus:border-blue-600 transition-all"/>
                    </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Início</label>
                        <input type="text" required placeholder="00:00:00" onBlur={(e) => setFormData({...formData, inicio: formatTimeInput(e.target.value)})} defaultValue={formData.inicio} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-mono outline-none focus:border-blue-600"/>
                    </div>
                    <div className="space-y-1">
                        <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest ml-1">Saída</label>
                        <input type="text" placeholder="00:00:00" onBlur={(e) => setFormData({...formData, saida: formatTimeInput(e.target.value)})} defaultValue={formData.saida} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-mono outline-none focus:border-blue-600"/>
                    </div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-[0.2em] text-[10px] rounded-xl flex items-center justify-center gap-2 shadow-xl transition-all active:scale-95 border-b-4 border-blue-900 mt-4">
                    {isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />} SALVAR REGISTRO
                </button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
