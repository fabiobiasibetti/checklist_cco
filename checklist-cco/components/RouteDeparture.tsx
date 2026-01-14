
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteDeparture, User, RouteOperationMapping } from '../types';
import { SharePointService } from '../services/sharepointService';
import { parseRouteDeparturesManual } from '../services/geminiService';
import { 
  Plus, Trash2, Save, Clock, X, Upload, 
  Loader2, RefreshCw, ShieldCheck,
  AlertTriangle, Link, CheckCircle2, ChevronDown, 
  Filter, Search, Check, CheckSquare, Square,
  BarChart3, PieChart as PieChartIcon, TrendingUp,
  Activity, EyeOff, ChevronRight, AlignLeft, Type as TypeIcon
} from 'lucide-react';
import { PieChart, Pie, Cell, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, Legend } from 'recharts';

interface RouteConfig {
    operacao: string;
    email: string;
    tolerancia: string;
}

const MOTIVOS = [
  'Fábrica',
  'Logística',
  'Mão de obra',
  'Manutenção',
  'Divergência de Roteirização',
  'Solicitado pelo Cliente',
  'Infraestrutura'
];

const OBSERVATION_TEMPLATES: Record<string, string[]> = {
  'Fábrica': ["Atraso na descarga | Entrada **:**h - Saída **:**h"],
  'Logística': [
    "Atraso no lavador | Chegada da rota anterior às **:**h - Entrada na fábrica às **:**h",
    "Motorista adiantou a rota devido à desvios",
    "Atraso na rota anterior (nome da rota)",
    "Atraso na rota anterior | Chegada no lavador **:**h - Entrada na fábrica às **:**h",
    "Falta de material de coleta para realizar a rota"
  ],
  'Mão de obra': [
    "Atraso do motorista",
    "Adiantamento do motorista",
    "A rota iniciou atrasada devido à interjornada do motorista | Atrasou na rota anterior devido à",
    "Troca do motorista previsto devido à saúde"
  ],
  'Manutenção': [
    "Precisou realizar a troca de pneus | Início **:**h - Término **:**h",
    "Troca de mola | Início **:**h - Término **:**h",
    "Manutenção na parte elétrica | Início **:**h - Término **:**h",
    "Manutenção nos freios | Início **:**h - Término **:**h",
    "Manutenção na bomba de carregamento de leite | Início **:**h - Término **:**h"
  ],
  'Divergência de Roteirização': [
    "Horário de saída da rota não atende os produtores",
    "Horário de saída da rota precisa ser alterado devido à entrada de produtores"
  ],
  'Solicitado pelo Cliente': [
    "Rota saiu adiantada para realizar socorro",
    "Cliente solicitou para a rota sair adiantada"
  ],
  'Infraestrutura': []
};

const FilterDropdown = ({ col, routes, colFilters, setColFilters, selectedFilters, setSelectedFilters, onClose, innerRef }: any) => {
    const values: string[] = Array.from(new Set(routes.map((r: any) => String(r[col] || "")))).sort() as string[];
    const selected = (selectedFilters[col] as string[]) || [];
    
    const toggleValue = (val: string) => {
        const next = selected.includes(val) ? selected.filter(v => v !== val) : [...selected, val];
        setSelectedFilters({ ...selectedFilters, [col]: next });
    };

    return (
        <div ref={innerRef} className="absolute top-10 left-0 z-[100] bg-white border border-slate-200 shadow-xl rounded-xl w-64 p-3 text-slate-700 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 rounded-lg border border-slate-200">
                <Search size={14} className="text-slate-400" />
                <input 
                    type="text" 
                    placeholder="Filtrar..." 
                    autoFocus
                    value={colFilters[col] || ""} 
                    onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })} 
                    className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800"
                />
            </div>
            <div className="max-h-56 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-100 py-2">
                {values.filter(v => v.toLowerCase().includes((colFilters[col] || "").toLowerCase())).map(v => (
                    <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-all">
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} className="text-slate-300" />}
                        <span className="text-[10px] font-bold uppercase truncate text-slate-600">{v || "(VAZIO)"}</span>
                    </div>
                ))}
            </div>
            <button 
                onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); onClose(); }} 
                className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 hover:bg-red-100 rounded-lg border border-red-100 transition-colors"
            >
                Limpar Filtro
            </button>
        </div>
    );
};

const RouteDepartureView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [routeMappings, setRouteMappings] = useState<RouteOperationMapping[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [isLinkModalOpen, setIsLinkModalOpen] = useState(false);
  const [isStatsModalOpen, setIsStatsModalOpen] = useState(false);
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importText, setImportText] = useState('');
  const [currentTime, setCurrentTime] = useState(new Date());
  
  const [zoomLevel, setZoomLevel] = useState(0.9);
  const [activeObsId, setActiveObsId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [isTextWrapEnabled, setIsTextWrapEnabled] = useState(false);

  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});

  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    select: 35, rota: 140, data: 125, inicio: 95, motorista: 230, placa: 100, saida: 95, motivo: 170, observacao: 400, geral: 70, operacao: 140, status: 90, tempo: 90
  });

  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);
  const filterRef = useRef<HTMLDivElement>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const obsDropdownRef = useRef<HTMLDivElement>(null);

  const getAccessToken = (): string => (window as any).__access_token || '';

  const [formData, setFormData] = useState<Partial<RouteDeparture>>({
    rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '00:00:00', motorista: '', placa: '', operacao: '', motivo: '', observacao: '', statusGeral: 'OK', aviso: 'NÃO'
  });

  const loadData = async () => {
    const token = getAccessToken();
    if (!token) return;
    setIsLoading(true);
    try {
      const [configs, mappings, spData] = (await Promise.all([
        SharePointService.getRouteConfigs(token, currentUser.email),
        SharePointService.getRouteOperationMappings(token),
        SharePointService.getDepartures(token)
      ])) as [RouteConfig[], RouteOperationMapping[], RouteDeparture[]];
      
      setUserConfigs(configs || []);
      setRouteMappings(mappings || []);
      const allowedOps = new Set(configs.map(c => (c.operacao || "").toUpperCase().trim()));
      const fixedData = spData.map(route => {
        if (!route.operacao || route.operacao === "") {
            const match = mappings.find(m => m.Title === route.rota);
            if (match && allowedOps.has((match.OPERACAO || "").toUpperCase().trim())) {
                return { ...route, operacao: match.OPERACAO.toUpperCase().trim() };
            }
        }
        return route;
      });
      const itemsToLink = fixedData.filter(r => !r.operacao || r.operacao === "");
      if (itemsToLink.length > 0) { setPendingItems(itemsToLink); setIsLinkModalOpen(true); }
      setRoutes(fixedData);
    } catch (e: any) { console.error(e); } finally { setIsLoading(false); }
  };

  const removeSelectedRows = async () => {
    if (selectedIds.size === 0) return;
    const token = getAccessToken();
    if (!token) return;
    if (confirm(`Deseja excluir os ${selectedIds.size} registros selecionados?`)) {
        setIsSyncing(true);
        try {
            await Promise.all(Array.from(selectedIds).map((id: string) => SharePointService.deleteDeparture(token, id)));
            setRoutes(prev => prev.filter(r => !selectedIds.has(r.id)));
            setSelectedIds(new Set());
        } catch (err: any) { alert("Erro ao excluir: " + err.message); } finally { setIsSyncing(false); }
    }
  };

  useEffect(() => { loadData(); const timer = setInterval(() => setCurrentTime(new Date()), 10000); return () => clearInterval(timer); }, [currentUser]);

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (resizingRef.current) {
        const { col, startX, startWidth } = resizingRef.current;
        const newWidth = Math.max(10, startWidth + (e.clientX - startX));
        setColWidths(prev => ({ ...prev, [col]: newWidth }));
      }
    };
    const handleMouseUp = () => { resizingRef.current = null; };
    const handleClickOutside = (e: MouseEvent) => {
        if (filterRef.current && !filterRef.current.contains(e.target as Node)) { setActiveFilterCol(null); }
        if (obsDropdownRef.current && !obsDropdownRef.current.contains(e.target as Node)) { setActiveObsId(null); }
    };
    const handleKeyDown = (e: KeyboardEvent) => {
        if (e.key === 'Delete' && selectedIds.size > 0) {
            const target = e.target as HTMLElement;
            if (target.tagName !== 'INPUT' && target.tagName !== 'TEXTAREA' && target.tagName !== 'SELECT') { removeSelectedRows(); }
        }
    };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('mousedown', handleClickOutside);
    window.addEventListener('keydown', handleKeyDown);
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      window.removeEventListener('mousedown', handleClickOutside);
      window.removeEventListener('keydown', handleKeyDown);
    };
  }, [selectedIds]);

  const startResize = (e: React.MouseEvent, col: string) => { e.preventDefault(); resizingRef.current = { col, startX: e.clientX, startWidth: colWidths[col] }; };

  const formatTimeInput = (value: string): string => {
    let clean = (value || "").replace(/[^0-9:]/g, '');
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
    const sInicio = inicio || '00:00:00';
    const sSaida = saida || '00:00:00';
    if (sInicio === '00:00:00' || sSaida === '00:00:00') return { gap: 'OK', status: 'OK', isOutOfTolerance: false };
    const startSec = timeToSeconds(sInicio);
    const endSec = timeToSeconds(sSaida);
    const diff = endSec - startSec;
    const toleranceSec = timeToSeconds(toleranceStr || "00:00:00");
    const gapFormatted = secondsToTime(diff);
    const isOutOfTolerance = Math.abs(diff) > toleranceSec;
    const status = isOutOfTolerance ? (diff > 0 ? 'Atrasado' : 'Adiantado') : 'OK';
    return { gap: gapFormatted, status, isOutOfTolerance };
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    const token = getAccessToken();
    if (!token) return;
    const route = routes.find(r => r.id === id);
    if (!route) return;
    let finalValue = value;
    if (field === 'inicio' || field === 'saida') finalValue = formatTimeInput(value);
    let updatedRoute = { ...route, [field]: finalValue };
    const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (updatedRoute.operacao || "").toUpperCase().trim());
    if (field === 'inicio' || field === 'saida' || field === 'operacao') {
        const { gap, status } = calculateGap(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00");
        updatedRoute.tempo = gap;
        updatedRoute.statusOp = status;
    }
    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));
    setIsSyncing(true);
    try { await SharePointService.updateDeparture(token, updatedRoute); } 
    catch (err: any) { console.error(err); } 
    finally { setIsSyncing(false); }
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

  const dashboardStats = useMemo(() => {
    const total = filteredRoutes.length;
    if (total === 0) return null;
    const okCount = filteredRoutes.filter(r => r.statusOp === 'OK').length;
    const delayedCount = filteredRoutes.filter(r => r.statusOp === 'Atrasado').length;
    const earlyCount = filteredRoutes.filter(r => r.statusOp === 'Adiantado').length;
    const reasonCounts: Record<string, number> = {};
    filteredRoutes.forEach(r => { if (r.statusOp !== 'OK') { const reason = r.motivo || 'NÃO INFORMADO'; reasonCounts[reason] = (reasonCounts[reason] || 0) + 1; } });
    const statusPie = [ { name: 'OK', value: okCount, color: '#10b981' }, { name: 'Atrasado', value: delayedCount, color: '#f97316' }, { name: 'Adiantado', value: earlyCount, color: '#3b82f6' } ];
    const reasonData = Object.entries(reasonCounts).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    return { total, okCount, delayedCount, earlyCount, statusPie, reasonData };
  }, [filteredRoutes]);

  const handleImport = async () => {
    if (!importText.trim()) return;
    setIsProcessingImport(true);
    try {
        const parsed = parseRouteDeparturesManual(importText);
        if (parsed.length === 0) throw new Error("Nenhum dado válido identificado.");
        const token = getAccessToken();
        for (const item of parsed) {
            const rotaStr = (item.rota || "").trim();
            const mapping = routeMappings.find(m => (m.Title || "").trim() === rotaStr);
            const op = (mapping?.OPERACAO || "").toUpperCase().trim();
            const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === op);
            const { gap, status } = calculateGap(item.inicio || '00:00:00', item.saida || '00:00:00', config?.tolerancia || "00:00:00");
            await SharePointService.updateDeparture(token, { ...item, id: '', statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture);
        }
        await loadData();
        setIsImportModalOpen(false);
    } catch (e: any) { alert(e.message); } finally { setIsProcessingImport(false); }
  };

  const getAlertStyles = (route: RouteDeparture) => {
    const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (route.operacao || "").toUpperCase().trim());
    const tolerance = String(config?.tolerancia || "00:00:00");
    const inicio = String(route.inicio || "00:00:00");
    const saida = String(route.saida || "00:00:00");
    const { isOutOfTolerance, status } = calculateGap(inicio, saida, tolerance);
    
    // Design Minimalista: Cores Mais Vibrantes para Visibilidade (Realce)
    if (saida !== '00:00:00' && status === 'Atrasado') return "border-l-[6px] border-orange-600 bg-orange-100/80 shadow-sm";
    if (status === 'Adiantado') return "border-l-[6px] border-blue-600 bg-blue-100/80 shadow-sm";
    
    const toleranceSec = timeToSeconds(tolerance);
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(inicio);
    if (saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) return "border-l-[6px] border-yellow-500 bg-yellow-100/80 shadow-sm";
    
    return "border-l-4 border-transparent";
  };

  const handleLinkPending = async () => {
    const token = getAccessToken();
    if (!token) return;
    setIsSyncing(true);
    try {
        const promises = pendingItems.map(async (item: any) => {
            if (item.operacao && item.rota) {
                const existingMapping = routeMappings.find(m => m.Title === (item.rota as string));
                if (!existingMapping) await SharePointService.addRouteOperationMapping(token, item.rota as string, item.operacao as string);
                const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (item.operacao as string).toUpperCase().trim());
                const { gap, status } = calculateGap(String(item.inicio || '00:00:00'), String(item.saida || '00:00:00'), String(config?.tolerancia || "00:00:00"));
                return SharePointService.updateDeparture(token, { ...item, statusOp: status, tempo: gap } as RouteDeparture);
            }
            return Promise.resolve();
        });
        await Promise.all(promises);
        await loadData();
        setIsLinkModalOpen(false);
        setPendingItems([]);
    } catch (e: any) { alert("Erro ao gravar vínculos: " + e.message); } finally { setIsSyncing(false); }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    const token = getAccessToken();
    if (!token) return;
    setIsSyncing(true);
    try {
        const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (formData.operacao || "").toUpperCase().trim());
        const { gap, status } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia || "00:00:00");
        const newRoute: RouteDeparture = { ...formData, id: '', statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
        await SharePointService.updateDeparture(token, newRoute);
        await loadData();
        setIsModalOpen(false);
        setFormData({ rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '00:00:00', motorista: '', placa: '', operacao: '', motivo: '', observacao: '', statusGeral: 'OK', aviso: 'NÃO' });
    } catch (err: any) { alert(err.message); } finally { setIsSyncing(false); }
  };

  const toggleSelection = (id: string) => {
    const newSelected = new Set(selectedIds);
    if (newSelected.has(id)) newSelected.delete(id);
    else newSelected.add(id);
    setSelectedIds(newSelected);
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4 bg-[#020617]">
        <Loader2 size={48} className="animate-spin" />
        <p className="font-bold animate-pulse text-[10px] uppercase tracking-[0.3em] text-slate-400">Gestão de Rotas CCO...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-[#020617] p-4 overflow-hidden select-none">
      <div className="flex justify-between items-center mb-6 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-primary-600 text-white rounded-2xl shadow-lg shadow-primary-600/20"><Clock size={20} /></div>
          <div>
            <h2 className="text-xl font-black text-white uppercase tracking-tight flex items-center gap-3">Saída de Rotas{isSyncing && <Loader2 size={16} className="animate-spin text-primary-500"/>}</h2>
            <div className="flex items-center gap-2"><ShieldCheck size={12} className="text-emerald-500"/><p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest">CCO Logística: {currentUser.name}</p></div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <button 
            onClick={() => setIsTextWrapEnabled(!isTextWrapEnabled)} 
            className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] tracking-wide transition-all shadow-sm ${isTextWrapEnabled ? 'bg-primary-600 text-white border-primary-600' : 'bg-slate-800 text-slate-300 border-slate-700 hover:bg-slate-700'}`}
          >
            <AlignLeft size={16} /> Quebra de Linha
          </button>
          <button onClick={() => setIsStatsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm"><BarChart3 size={16} /> Indicadores</button>
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-700 bg-slate-900"><RefreshCw size={18} /></button>
          <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 font-bold border border-emerald-700 uppercase text-[10px] tracking-wide shadow-sm transition-all"><Upload size={16} /> Importar</button>
          <button onClick={() => setIsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-primary-600 text-white rounded-lg hover:bg-primary-700 font-bold border border-primary-700 uppercase text-[10px] tracking-wide shadow-md transition-all"><Plus size={16} /> Nova Rota</button>
        </div>
      </div>

      <div ref={tableContainerRef} className="flex-1 overflow-auto bg-white rounded-2xl border border-slate-700/50 shadow-2xl relative scrollbar-thin overflow-x-auto">
        <div style={{ transform: `scale(${zoomLevel})`, transformOrigin: 'top left', width: `${100 / zoomLevel}%` }}>
            <table className="border-collapse table-fixed w-full min-w-max h-px">
              <thead className="sticky top-0 z-50 bg-[#1e293b] text-white shadow-md">
                <tr className="h-12">
                  {[
                    { id: 'select', label: '' },
                    { id: 'rota', label: 'ROTA' },
                    { id: 'data', label: 'DATA' },
                    { id: 'inicio', label: 'INÍCIO' },
                    { id: 'motorista', label: 'MOTORISTA' },
                    { id: 'placa', label: 'PLACA' },
                    { id: 'saida', label: 'SAÍDA' },
                    { id: 'motivo', label: 'MOTIVO' },
                    { id: 'observacao', label: 'OBSERVAÇÃO' },
                    { id: 'geral', label: 'GERAL' },
                    { id: 'operacao', label: 'OPERAÇÃO' },
                    { id: 'status', label: 'STATUS' },
                    { id: 'tempo', label: 'TEMPO' }
                  ].map(col => {
                    if (col.id === 'select') return <th key={col.id} style={{ width: colWidths.select }} className="bg-slate-900/50"></th>;
                    const hasFilter = !!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0;
                    return (
                      <th key={col.id} style={{ width: colWidths[col.id] }} className="relative p-1 border-r border-slate-700/50 text-[10px] font-black uppercase tracking-wider text-left select-none group">
                        <div className="flex items-center justify-between px-2 h-full">
                          <span className="flex items-center gap-1.5">{col.label}</span>
                          <button onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }} className={`p-1 rounded transition-all ${hasFilter ? 'text-yellow-400' : 'text-white/40 hover:text-white/60'}`}><Filter size={11} fill={hasFilter ? 'currentColor' : 'none'} /></button>
                        </div>
                        {activeFilterCol === col.id && (
                            <FilterDropdown 
                                col={col.id} 
                                routes={routes} 
                                colFilters={colFilters} 
                                setColFilters={setColFilters} 
                                selectedFilters={selectedFilters} 
                                setSelectedFilters={setSelectedFilters}
                                onClose={() => setActiveFilterCol(null)}
                                innerRef={filterRef}
                            />
                        )}
                        <div onMouseDown={(e) => startResize(e, col.id)} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredRoutes.map((route, idx) => {
                  const alertClasses = getAlertStyles(route);
                  const isSelected = selectedIds.has(route.id);
                  const rowBg = isSelected ? 'bg-primary-100/50' : 'bg-white hover:bg-slate-50';
                  const textClass = "w-full h-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold text-slate-800 uppercase transition-all placeholder-slate-300";
                  
                  const displayStatus = route.statusOp;
                  const showDetails = displayStatus !== 'OK';

                  return (
                    <tr key={route.id} className={`${rowBg} ${alertClasses} group transition-all h-auto`}>
                      <td 
                        className={`p-0 border-r border-slate-100 cursor-pointer transition-colors w-[35px] ${isSelected ? 'bg-primary-500' : 'hover:bg-slate-200'}`} 
                        onClick={() => toggleSelection(route.id)}
                      ></td>
                      <td className="p-0 border-r border-slate-100"><input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={`${textClass} font-black text-primary-600`} /></td>
                      <td className="p-0 border-r border-slate-100"><input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${textClass} text-center text-slate-600`} /></td>
                      <td className="p-0 border-r border-slate-100"><input type="text" value={route.inicio} onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} className={`${textClass} font-mono text-center`} /></td>
                      <td className="p-0 border-r border-slate-100"><input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value)} className={`${textClass}`} /></td>
                      <td className="p-0 border-r border-slate-100"><input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value)} className={`${textClass} font-mono text-center`} /></td>
                      <td className="p-0 border-r border-slate-100"><input type="text" value={route.saida} onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} className={`${textClass} font-mono text-center`} /></td>
                      <td className="p-0 border-r border-slate-100">
                        {showDetails ? (
                          <div className="flex items-center justify-center h-full px-2">
                              <select 
                                value={route.motivo} 
                                onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} 
                                className="w-full bg-slate-100 border-none rounded-lg px-2 py-1 text-[10px] font-bold text-slate-700 outline-none appearance-none text-center shadow-sm"
                              >
                                  <option value="">Selecione...</option>
                                  {MOTIVOS.map(m => (<option key={m} value={m}>{m}</option>))}
                              </select>
                          </div>
                        ) : null}
                      </td>
                      <td className="p-0 border-r border-slate-100 relative group/obs align-top h-full min-h-[44px]">
                        {showDetails ? (
                          <div className="flex items-start w-full h-full relative p-0 min-h-[44px]">
                            <textarea 
                                value={route.observacao || ""}
                                onChange={(e) => updateCell(route.id, 'observacao', e.target.value)}
                                onFocus={() => setActiveObsId(route.id)}
                                placeholder="Descreva..."
                                className={`w-full h-full min-h-[44px] bg-transparent outline-none border-none px-3 py-2 text-[11px] font-normal text-slate-800 placeholder-slate-400 resize-none overflow-hidden ${isTextWrapEnabled ? 'whitespace-normal break-words leading-relaxed' : 'truncate pr-8'}`}
                                style={{ height: isTextWrapEnabled ? 'auto' : '44px' }}
                                onInput={(e) => {
                                    if (isTextWrapEnabled) {
                                        const el = e.target as HTMLTextAreaElement;
                                        el.style.height = 'auto';
                                        el.style.height = el.scrollHeight + 'px';
                                    }
                                }}
                                ref={(el) => {
                                    if (el && isTextWrapEnabled) {
                                        el.style.height = 'auto';
                                        el.style.height = el.scrollHeight + 'px';
                                    }
                                }}
                            />
                            {!isTextWrapEnabled && (
                                <button onClick={(e) => { e.stopPropagation(); setActiveObsId(activeObsId === route.id ? null : route.id); }} className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 text-slate-400 hover:text-primary-600 transition-colors opacity-40 group-hover/obs:opacity-100"><ChevronDown size={12} /></button>
                            )}
                          </div>
                        ) : null}
                        {activeObsId === route.id && (
                          <div ref={obsDropdownRef} className="absolute top-full left-0 w-full z-[110] bg-white border border-slate-200 rounded-xl shadow-2xl overflow-hidden animate-in fade-in slide-in-from-top-1">
                            <div className="p-2 border-b border-slate-100 flex items-center justify-between"><span className="text-[8px] font-black uppercase text-slate-400 tracking-widest">Modelos: {route.motivo || 'Geral'}</span><X size={10} className="text-slate-400 cursor-pointer" onClick={() => setActiveObsId(null)} /></div>
                            <div className="max-h-48 overflow-y-auto scrollbar-thin">
                              {(route.motivo ? (OBSERVATION_TEMPLATES[route.motivo] || []) : Object.values(OBSERVATION_TEMPLATES).flat()).filter(t => t.toLowerCase().includes((route.observacao || "").toLowerCase())).map((template, tIdx) => (
                                  <div key={tIdx} onClick={() => { updateCell(route.id, 'observacao', template); setActiveObsId(null); }} className="p-2 text-[10px] text-slate-600 hover:bg-primary-50 hover:text-primary-700 cursor-pointer transition-all border-b border-slate-50 last:border-0 flex items-center gap-2"><ChevronRight size={10} className="shrink-0" />{template}</div>
                                ))}
                            </div>
                          </div>
                        )}
                      </td>
                      <td className="p-0 border-r border-slate-100 align-middle"><select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full bg-transparent border-none text-[10px] font-bold text-center appearance-none text-slate-700"><option value="OK">OK</option><option value="NOK">NOK</option></select></td>
                      <td className="p-1 border-r border-slate-100 text-center font-bold uppercase text-[9px] text-slate-500 align-middle">{route.operacao || "---"}</td>
                      <td className="p-1 border-r border-slate-100 text-center align-middle">
                        <span className={`px-2 py-1 rounded-full text-[8px] font-black border ${displayStatus === 'OK' ? 'bg-emerald-50 border-emerald-100 text-emerald-600' : 'bg-red-50 border-red-100 text-red-600'}`}>{displayStatus}</span>
                      </td>
                      <td className="p-1 border-r border-slate-100 text-center font-mono font-bold text-[10px] text-slate-600 align-middle">{route.tempo}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
        </div>
      </div>

      {isStatsModalOpen && dashboardStats && (
        <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-md z-[200] flex items-center justify-center p-6">
            <div className="bg-white border border-slate-200 rounded-[2rem] shadow-2xl w-full max-w-5xl max-h-[90vh] overflow-hidden flex flex-col animate-in zoom-in duration-300">
                <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white">
                    <div className="flex items-center gap-4"><div className="p-2.5 bg-white/10 rounded-xl"><TrendingUp size={24} /></div><div><h3 className="font-black uppercase tracking-widest text-base leading-none">Dashboard de Performance</h3><p className="text-slate-400 text-[10px] font-bold uppercase mt-1">Resumo ({dashboardStats.total} rotas filtradas)</p></div></div>
                    <button onClick={() => setIsStatsModalOpen(false)} className="hover:bg-white/10 p-2 rounded-xl transition-all"><X size={28} /></button>
                </div>
                <div className="p-8 flex-1 overflow-y-auto space-y-8 scrollbar-thin bg-slate-50">
                    <div className="grid grid-cols-4 gap-6">
                        {[
                            { label: 'Total Filtrado', value: dashboardStats.total, icon: Activity, color: 'text-slate-700 bg-white shadow-sm' },
                            { label: 'No Horário', value: `${Math.round((dashboardStats.okCount / (dashboardStats.total || 1)) * 100)}%`, icon: CheckCircle2, color: 'text-emerald-600 bg-emerald-50' },
                            { label: 'Atrasadas', value: `${Math.round((dashboardStats.delayedCount / (dashboardStats.total || 1)) * 100)}%`, icon: AlertTriangle, color: 'text-orange-600 bg-orange-50' },
                            { label: 'Adiantadas', value: `${Math.round((dashboardStats.earlyCount / (dashboardStats.total || 1)) * 100)}%`, icon: TrendingUp, color: 'text-blue-600 bg-blue-50' }
                        ].map((stat, idx) => (
                            <div key={idx} className={`p-6 rounded-2xl border border-slate-200 flex flex-col gap-2 ${stat.color}`}>
                                <stat.icon size={20} /><span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mt-2">{stat.label}</span>
                                <div className="text-3xl font-black tracking-tighter">{stat.value}</div>
                            </div>
                        ))}
                    </div>
                    <div className="grid grid-cols-2 gap-8">
                        <div className="p-6 rounded-3xl bg-white border border-slate-200 shadow-sm h-[400px] flex flex-col">
                            <h4 className="text-slate-700 font-black uppercase text-xs tracking-widest mb-6 flex items-center gap-2"><PieChartIcon size={16} className="text-primary-500" /> Distribuição de Status</h4>
                            <div className="flex-1"><ResponsiveContainer width="100%" height="100%"><PieChart><Pie data={dashboardStats.statusPie} innerRadius={80} outerRadius={110} paddingAngle={5} dataKey="value">{dashboardStats.statusPie.map((entry, index) => (<Cell key={`cell-${index}`} fill={entry.color} />))}</Pie><Tooltip contentStyle={{ borderRadius: '12px', border: 'none', shadow: '0 10px 15px -3px rgba(0,0,0,0.1)' }} /><Legend verticalAlign="bottom" height={36} formatter={(value) => <span className="text-slate-500 font-bold uppercase text-[10px]">{value}</span>} /></PieChart></ResponsiveContainer></div>
                        </div>
                        <div className="p-6 rounded-3xl bg-white border border-slate-200 shadow-sm h-[400px] flex flex-col">
                            <h4 className="text-slate-700 font-black uppercase text-xs tracking-widest mb-6 flex items-center gap-2"><BarChart3 size={16} className="text-yellow-500" /> Motivos de Desvio</h4>
                            <div className="flex-1"><ResponsiveContainer width="100%" height="100%"><BarChart data={dashboardStats.reasonData} layout="vertical"><XAxis type="number" hide /><YAxis dataKey="name" type="category" width={120} tick={{ fill: '#64748b', fontSize: 10, fontWeight: 'bold' }} axisLine={false} tickLine={false} /><Tooltip contentStyle={{ borderRadius: '12px', border: 'none' }} cursor={{ fill: 'rgba(0,0,0,0.02)' }} /><Bar dataKey="value" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={20} /></BarChart></ResponsiveContainer></div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
      )}

      {isImportModalOpen && (
        <div className="fixed inset-0 bg-slate-950/60 backdrop-blur-md z-[200] flex items-center justify-center p-4">
             <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200">
                <div className="bg-emerald-500 p-6 flex justify-between items-center text-white"><div className="flex items-center gap-3"><Upload size={20} className="bg-white/20 p-1.5 rounded-lg" /><h3 className="font-black uppercase tracking-widest text-xs">Importar Dados Excel</h3></div><button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button></div>
                <div className="p-8">
                    <textarea value={importText} onChange={e => setImportText(e.target.value)} className="w-full h-64 p-5 border-2 border-slate-100 rounded-2xl bg-slate-50 text-[11px] font-mono mb-6 focus:ring-2 focus:ring-emerald-500 outline-none text-slate-800 shadow-inner scrollbar-thin" placeholder="Cole aqui..." />
                    <button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-4 bg-emerald-500 text-white font-black uppercase tracking-widest text-[11px] rounded-xl shadow-lg flex items-center justify-center gap-3 transition-all hover:bg-emerald-600 disabled:opacity-50">{isProcessingImport ? <Loader2 size={18} className="animate-spin" /> : <span>Processar Importação</span>}</button>
                </div>
             </div>
        </div>
      )}

      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-950/60 backdrop-blur-md z-[200] flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in">
            <div className="bg-primary-600 text-white p-6 flex justify-between items-center"><h3 className="font-black uppercase tracking-widest text-xs flex items-center gap-3"><Plus size={20} /> Novo Registro</h3><button onClick={() => setIsModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button></div>
            <form onSubmit={handleSubmit} className="p-8 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Data</label><input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border border-slate-100 rounded-xl bg-slate-50 text-slate-800 text-[11px] font-bold outline-none focus:border-primary-600 transition-all"/></div>
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Rota</label><input type="text" required value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value})} className="w-full p-3 border border-slate-100 rounded-xl bg-slate-50 text-[11px] font-black text-primary-600 outline-none focus:border-primary-600 transition-all"/></div>
                </div>
                <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Operação</label><select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border border-slate-100 rounded-xl bg-slate-50 text-[11px] font-black text-slate-700 outline-none focus:border-primary-600"><option value="">Selecione...</option>{userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}</select></div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Motorista</label><input type="text" required value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value})} className="w-full p-3 border border-slate-100 rounded-xl bg-slate-50 text-slate-800 text-[11px] font-bold outline-none focus:border-primary-600 transition-all"/></div>
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Placa</label><input type="text" required value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value})} className="w-full p-3 border border-slate-100 rounded-xl bg-slate-50 text-slate-800 text-[11px] font-black outline-none focus:border-primary-600 transition-all"/></div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-primary-600 hover:bg-primary-700 text-white font-black uppercase tracking-widest text-[11px] rounded-xl flex items-center justify-center gap-2 shadow-xl transition-all mt-4">{isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />} SALVAR NO SHAREPOINT</button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
