
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
  Activity, EyeOff, ChevronRight
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
  'Fábrica': [
    "Atraso na descarga | Entrada **:**h - Saída **:**h"
  ],
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
  
  const [isAvisoVisible, setIsAvisoVisible] = useState(false);
  const [contextMenu, setContextMenu] = useState<{ x: number, y: number, id: string } | null>(null);
  const [zoomLevel, setZoomLevel] = useState(0.9);
  const [activeObsId, setActiveObsId] = useState<string | null>(null);

  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});

  const [pendingItems, setPendingItems] = useState<Partial<RouteDeparture>[]>([]);

  const [colWidths, setColWidths] = useState<Record<string, number>>({
    rota: 140,
    data: 125,
    inicio: 95,
    motorista: 230,
    placa: 100,
    saida: 95,
    motivo: 170,
    observacao: 320,
    geral: 70,
    aviso: 70,
    operacao: 140,
    status: 90,
    tempo: 90,
  });

  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);
  const filterRef = useRef<HTMLDivElement>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const obsDropdownRef = useRef<HTMLDivElement>(null);

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

  const clearAllFilters = () => {
    setColFilters({});
    setSelectedFilters({});
    setActiveFilterCol(null);
  };

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
      if (itemsToLink.length > 0) {
          setPendingItems(itemsToLink);
          setIsLinkModalOpen(true);
      }

      setRoutes(fixedData);
    } catch (e: any) {
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

    const handleMouseUp = () => { resizingRef.current = null; };

    const handleClickOutside = (e: MouseEvent) => {
        if (filterRef.current && !filterRef.current.contains(e.target as Node)) {
            setActiveFilterCol(null);
        }
        if (contextMenu) setContextMenu(null);
        if (obsDropdownRef.current && !obsDropdownRef.current.contains(e.target as Node)) {
          setActiveObsId(null);
        }
    };

    const handleWheel = (e: WheelEvent) => {
        if (e.ctrlKey) {
            e.preventDefault();
            const delta = e.deltaY > 0 ? -0.05 : 0.05;
            setZoomLevel(prev => Math.min(Math.max(prev + delta, 0.4), 1.3));
        }
    };

    const handleKeyDown = (e: KeyboardEvent) => {
        if (e.ctrlKey && e.shiftKey && e.key.toLowerCase() === 'l') {
            e.preventDefault();
            clearAllFilters();
        }
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('mousedown', handleClickOutside);
    window.addEventListener('keydown', handleKeyDown);
    
    const container = tableContainerRef.current;
    if (container) {
        container.addEventListener('wheel', handleWheel, { passive: false });
    }
    
    return () => {
      clearInterval(timer);
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      window.removeEventListener('mousedown', handleClickOutside);
      window.removeEventListener('keydown', handleKeyDown);
      if (container) container.removeEventListener('wheel', handleWheel);
    };
  }, [currentUser, contextMenu]);

  const handleContextMenu = (e: React.MouseEvent, id: string) => {
    e.preventDefault();
    setContextMenu({ x: e.clientX, y: e.clientY, id });
  };

  const startResize = (e: React.MouseEvent, col: string) => {
    e.preventDefault();
    resizingRef.current = { col, startX: e.clientX, startWidth: colWidths[col] };
  };

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
    const statusPie = [ { name: 'OK', value: okCount, color: '#10b981' }, { name: 'Atrasado', value: delayedCount, color: '#f75a68' }, { name: 'Adiantado', value: earlyCount, color: '#3b82f6' } ];
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
            
            const inicioStr = item.inicio || '00:00:00';
            const saidaStr = item.saida || '00:00:00';
            const toleranceStr = config?.tolerancia || "00:00:00";
            const { gap, status } = calculateGap(inicioStr, saidaStr, toleranceStr);
            
            await SharePointService.updateDeparture(token!, { ...item, id: '', statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture);
        }
        await loadData();
        setIsImportModalOpen(false);
    } catch (e: any) { alert(e.message); } finally { setIsProcessingImport(false); }
  };

  const removeRow = async (id: string) => {
    const token = getAccessToken();
    if (!token) return;
    if (confirm('Excluir registro permanentemente?')) {
      setIsSyncing(true);
      try {
        await SharePointService.deleteDeparture(token, id);
        setRoutes(routes.filter(r => r.id !== id));
      } catch (err: any) { alert(err.message); } finally { setIsSyncing(false); }
    }
  };

  const getAlertStyles = (route: RouteDeparture) => {
    const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (route.operacao || "").toUpperCase().trim());
    const tolerance = String(config?.tolerancia || "00:00:00");
    const inicio = String(route.inicio || "00:00:00");
    const saida = String(route.saida || "00:00:00");
    const { isOutOfTolerance } = calculateGap(inicio, saida, tolerance);
    if (saida !== '00:00:00' && isOutOfTolerance) return "border-l-4 border-[#F75A68] bg-[#F75A68]/10";
    const toleranceSec = timeToSeconds(tolerance);
    const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
    const scheduledStartSec = timeToSeconds(inicio);
    if (saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) return "border-l-4 border-[#FF9000] bg-[#FF9000]/10";
    return "border-l-4 border-transparent";
  };

  const handleLinkPending = async () => {
    const token = getAccessToken();
    if (!token) return;
    setIsSyncing(true);
    try {
        const promises = pendingItems.map(async (item) => {
            if (item.operacao && item.rota) {
                const existingMapping = routeMappings.find(m => m.Title === (item.rota as string));
                if (!existingMapping) await SharePointService.addRouteOperationMapping(token, item.rota as string, item.operacao as string);
                const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (item.operacao as string).toUpperCase().trim());
                // Explicitly cast arguments to string to avoid "unknown" type assignment errors
                const { gap, status } = calculateGap(
                  String(item.inicio || '00:00:00'), 
                  String(item.saida || '00:00:00'), 
                  String(config?.tolerancia || "00:00:00")
                );
                // Ensure token is cast to string as required by SharePointService
                return SharePointService.updateDeparture(String(token), { ...item, statusOp: status, tempo: gap } as RouteDeparture);
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

  const FilterDropdown = ({ col }: { col: string }) => {
    const values = Array.from(new Set(routes.map(r => String(r[col as keyof RouteDeparture] || "")))).sort();
    const selected = (selectedFilters[col] as string[]) || [];
    const toggleValue = (val: string) => {
        const next = selected.includes(val) ? selected.filter(v => v !== val) : [...selected, val];
        setSelectedFilters({ ...selectedFilters, [col]: next });
    };

    return (
        <div ref={filterRef} className="absolute top-10 left-0 z-[100] bg-[#1e1e24] border border-slate-700 shadow-[0_10px_40px_rgba(0,0,0,0.5)] rounded-2xl w-64 p-3 text-slate-200 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-center gap-2 mb-3 p-2 bg-[#121214] rounded-lg border border-slate-800">
                <Search size={14} className="text-slate-500" />
                <input type="text" placeholder="Filtrar..." value={colFilters[col] || ""} onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })} className="w-full bg-transparent outline-none text-[10px] font-bold text-white"/>
            </div>
            <div className="max-h-56 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-800 py-2">
                {values.map(v => (
                    <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2 hover:bg-slate-800 rounded-lg cursor-pointer transition-all">
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-500" /> : <Square size={14} className="text-slate-600" />}
                        <span className="text-[10px] font-bold uppercase truncate text-slate-100">{v || "(VAZIO)"}</span>
                    </div>
                ))}
            </div>
            <button onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); }} className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-400 bg-red-900/10 hover:bg-red-900/20 rounded-lg border border-red-900/30 transition-colors">Limpar Filtro</button>
        </div>
    );
  };

  if (isLoading) return (
    <div className="h-full flex flex-col items-center justify-center text-blue-600 gap-4 bg-[#020617]">
        <Loader2 size={48} className="animate-spin" />
        <p className="font-bold animate-pulse text-[10px] uppercase tracking-[0.3em]">Gestão de Rotas CCO...</p>
    </div>
  );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-[#020617] p-4 overflow-hidden select-none">
      <div className="flex justify-between items-center mb-4 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-2.5 bg-blue-600 text-white rounded-xl shadow-2xl ring-4 ring-blue-600/10"><Clock size={20} /></div>
          <div>
            <h2 className="text-xl font-black text-white uppercase tracking-tighter flex items-center gap-3">Saída de Rotas{isSyncing && <Loader2 size={16} className="animate-spin text-blue-500"/>}</h2>
            <div className="flex items-center gap-2"><ShieldCheck size={12} className="text-emerald-500"/><p className="text-[9px] text-slate-400 font-black uppercase tracking-widest">CCO Logística: {currentUser.name}</p></div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <button onClick={() => setIsStatsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-200 rounded-lg hover:bg-slate-700 font-black border border-slate-700 uppercase text-[9px] tracking-widest transition-all mr-2 shadow-lg"><BarChart3 size={16} /> Indicadores</button>
          <div className="hidden lg:flex items-center gap-2 px-3 py-1 bg-[#121214] border border-slate-800 rounded-lg text-[9px] text-slate-300 font-bold uppercase mr-4">Zoom: {Math.round(zoomLevel * 100)}%</div>
          <button onClick={loadData} className="p-2 text-slate-500 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-800"><RefreshCw size={18} /></button>
          <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 font-black shadow-lg border-b-2 border-emerald-900 uppercase text-[9px] tracking-widest transition-all active:scale-95"><Upload size={16} /> Importar</button>
          <button onClick={() => setIsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 font-black shadow-lg border-b-2 border-blue-900 uppercase text-[9px] tracking-widest transition-all active:scale-95"><Plus size={16} /> Nova Rota</button>
        </div>
      </div>

      {contextMenu && (
        <div className="fixed z-[300] bg-[#1e1e24] border border-slate-700 rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-100" style={{ top: contextMenu.y, left: contextMenu.x }}>
          <button onClick={() => removeRow(contextMenu.id)} className="w-full px-4 py-3 flex items-center gap-3 text-red-400 hover:bg-red-500/10 transition-colors font-black uppercase text-[10px]">
            <Trash2 size={16} /> Excluir Registro
          </button>
        </div>
      )}

      <div ref={tableContainerRef} className="flex-1 overflow-auto bg-[#121214] rounded-xl border border-[#1e1e24] shadow-2xl relative scrollbar-thin overflow-x-auto">
        <div style={{ transform: `scale(${zoomLevel})`, transformOrigin: 'top left', width: `${100 / zoomLevel}%` }}>
            <table className="border-collapse table-fixed w-full min-w-max">
              <thead className="sticky top-0 z-50 bg-blue-700 text-white shadow-lg">
                <tr className="h-10">
                  {[
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
                    const isAviso = col.id === 'aviso';
                    const hasFilter = !!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0;
                    if (isAviso && !isAvisoVisible) {
                        return (
                          <th key={col.id} className="w-8 border-r border-blue-600/50 bg-blue-800 flex items-center justify-center cursor-pointer" onClick={() => setIsAvisoVisible(true)} title="Expandir Aviso">
                            <ChevronDown size={12} className="-rotate-90 text-white/50" />
                          </th>
                        );
                    }
                    return (
                      <th key={col.id} style={{ width: colWidths[col.id] }} className="relative p-1 border-r border-blue-600/50 text-[9px] font-black uppercase tracking-widest text-left select-none group">
                        <div className="flex items-center justify-between px-1.5 h-full">
                          <span className="flex items-center gap-1.5">
                            {col.label}
                            {isAviso && <button onClick={(e) => {e.stopPropagation(); setIsAvisoVisible(false);}} className="p-0.5 hover:bg-white/20 rounded"><EyeOff size={10}/></button>}
                          </span>
                          <button onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }} className={`p-1 rounded transition-all ${hasFilter ? 'text-yellow-400 bg-white/10' : 'text-white/40 hover:bg-white/10'}`}><Filter size={10} fill={hasFilter ? 'currentColor' : 'none'} /></button>
                        </div>
                        {activeFilterCol === col.id && <FilterDropdown col={col.id} />}
                        <div onMouseDown={(e) => startResize(e, col.id)} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/30">
                {filteredRoutes.map((route, idx) => {
                  const alertClasses = getAlertStyles(route);
                  const isEven = idx % 2 === 0;
                  const rowBg = isEven ? 'bg-[#121214]' : 'bg-[#18181b]';
                  const textClass = "w-full h-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-bold text-white uppercase tracking-tight transition-all focus:bg-blue-500/10 placeholder-slate-400";

                  const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (route.operacao || "").toUpperCase().trim());
                  const tolerance = String(config?.tolerancia || "00:00:00");
                  const toleranceSec = timeToSeconds(tolerance);
                  const nowSec = (currentTime.getHours() * 3600) + (currentTime.getMinutes() * 60) + currentTime.getSeconds();
                  const scheduledStartSec = timeToSeconds(String(route.inicio || "00:00:00"));
                  let displayStatus = route.statusOp;
                  if (route.saida === '00:00:00' && nowSec > (scheduledStartSec + toleranceSec)) {
                      displayStatus = 'Atrasado';
                  }

                  return (
                    <tr key={route.id} className={`${rowBg} ${alertClasses} group transition-all h-9 hover:bg-blue-900/10`}>
                      <td className="p-0 border-r border-[#1e1e24]" onContextMenu={(e) => handleContextMenu(e, route.id)}>
                        <input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={`${textClass} text-left font-black text-blue-400 cursor-help`} title="Clique direito p/ opções" />
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]"><input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${textClass} font-mono text-center text-white`} /></td>
                      <td className="p-0 border-r border-[#1e1e24]"><input type="text" value={route.inicio} onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} className={`${textClass} font-mono text-center text-white`} /></td>
                      <td className="p-0 border-r border-[#1e1e24]"><input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value.toUpperCase())} className={`${textClass} text-left text-white`} /></td>
                      <td className="p-0 border-r border-[#1e1e24]"><input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value.toUpperCase())} className={`${textClass} font-mono tracking-widest text-center text-white`} /></td>
                      <td className="p-0 border-r border-[#1e1e24]"><input type="text" value={route.saida} onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} className={`${textClass} font-mono text-center text-white`} /></td>
                      <td className="p-0 border-r border-[#1e1e24]">
                        <div className="flex items-center justify-center h-full px-2">
                            <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full bg-[#1e1e24] border border-slate-700 rounded px-1 py-0.5 text-[10px] font-black uppercase text-white outline-none appearance-none text-center focus:border-blue-500">
                                <option value="">SELECIONE...</option>
                                {MOTIVOS.map(m => (<option key={m} value={m}>{m.toUpperCase()}</option>))}
                            </select>
                        </div>
                      </td>
                      <td className="p-0 border-r border-[#1e1e24] relative group/obs">
                        <div className="flex items-center w-full h-full relative">
                          <input 
                            type="text" 
                            value={route.observacao} 
                            onFocus={() => setActiveObsId(route.id)}
                            onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} 
                            className={`${textClass} italic font-normal text-slate-100 text-left truncate pr-8`} 
                            placeholder="Descreva..." 
                          />
                          <button 
                            onClick={(e) => { e.stopPropagation(); setActiveObsId(activeObsId === route.id ? null : route.id); }}
                            className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 text-slate-500 hover:text-blue-400 transition-colors opacity-30 group-hover/obs:opacity-100"
                          >
                            <ChevronDown size={12} />
                          </button>
                        </div>
                        {activeObsId === route.id && (
                          <div 
                            ref={obsDropdownRef}
                            className="absolute top-full left-0 w-full z-[110] bg-[#1e1e24] border border-slate-700 rounded-xl shadow-[0_15px_40px_rgba(0,0,0,0.6)] overflow-hidden animate-in fade-in slide-in-from-top-1"
                          >
                            <div className="p-2 border-b border-slate-700 flex items-center justify-between">
                              <span className="text-[8px] font-black uppercase text-slate-500 tracking-widest">Modelos Sugeridos: {route.motivo || 'Geral'}</span>
                              <X size={10} className="text-slate-500 cursor-pointer" onClick={() => setActiveObsId(null)} />
                            </div>
                            <div className="max-h-48 overflow-y-auto scrollbar-thin">
                              {(route.motivo ? (OBSERVATION_TEMPLATES[route.motivo] || []) : Object.values(OBSERVATION_TEMPLATES).flat())
                                .filter(t => t.toLowerCase().includes((route.observacao || "").toLowerCase()))
                                .map((template, tIdx) => (
                                  <div 
                                    key={tIdx} 
                                    onClick={() => { updateCell(route.id, 'observacao', template); setActiveObsId(null); }}
                                    className="p-2 text-[10px] text-slate-300 hover:bg-blue-600 hover:text-white cursor-pointer transition-all border-b border-slate-800 last:border-0 flex items-center gap-2"
                                  >
                                    <ChevronRight size={10} className="shrink-0" />
                                    <span className="truncate">{template}</span>
                                  </div>
                                ))}
                              {route.motivo && (!OBSERVATION_TEMPLATES[route.motivo] || OBSERVATION_TEMPLATES[route.motivo].length === 0) && (
                                <div className="p-4 text-center text-[10px] text-slate-500 italic">Nenhum template para este motivo.</div>
                              )}
                            </div>
                          </div>
                        )}
                      </td>
                      <td className="p-0 border-r border-[#1e1e24]"><select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className={`${textClass} text-center appearance-none text-white`}><option value="OK">OK</option><option value="NOK">NOK</option></select></td>
                      {isAvisoVisible ? (
                        <td className="p-0 border-r border-[#1e1e24]"><select value={route.aviso} onChange={(e) => updateCell(route.id, 'aviso', e.target.value)} className={`${textClass} text-center appearance-none text-white`}><option value="SIM">SIM</option><option value="NÃO">NÃO</option></select></td>
                      ) : (
                        <td className="w-8 border-r border-[#1e1e24] bg-black/10"></td>
                      )}
                      <td className="p-1 border-r border-[#1e1e24] text-center font-black uppercase text-[9px] text-slate-300">{route.operacao || "---"}</td>
                      <td className="p-1 border-r border-[#1e1e24] text-center">
                        <span className={`px-2 py-0.5 rounded-[4px] text-[8px] font-black border ${displayStatus === 'OK' ? 'bg-emerald-900/30 border-emerald-800 text-emerald-400' : 'bg-red-900/30 border-red-800 text-red-400'}`}>{displayStatus}</span>
                      </td>
                      <td className="p-1 border-r border-[#1e1e24] text-center font-mono font-bold text-[10px] text-white">{route.tempo}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
        </div>
      </div>

      {isStatsModalOpen && dashboardStats && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-xl z-[200] flex items-center justify-center p-6">
            <div className="bg-[#121214] border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-5xl max-h-[90vh] overflow-hidden flex flex-col animate-in zoom-in duration-300">
                <div className="bg-blue-700 p-6 flex justify-between items-center text-white">
                    <div className="flex items-center gap-4">
                        <div className="p-2.5 bg-white/20 rounded-xl"><TrendingUp size={24} /></div>
                        <div><h3 className="font-black uppercase tracking-widest text-base">Dashboard de Performance</h3><p className="text-blue-100 text-xs font-bold">Resumo ({dashboardStats.total} rotas filtradas)</p></div>
                    </div>
                    <button onClick={() => setIsStatsModalOpen(false)} className="hover:bg-white/10 p-2 rounded-xl transition-all"><X size={28} /></button>
                </div>
                <div className="p-8 flex-1 overflow-y-auto space-y-8 scrollbar-thin">
                    <div className="grid grid-cols-4 gap-6">
                        {[
                            { label: 'Total Filtrado', value: dashboardStats.total, icon: Activity, color: 'text-white bg-slate-800' },
                            { label: 'No Horário', value: `${Math.round((dashboardStats.okCount / (dashboardStats.total || 1)) * 100)}%`, icon: CheckCircle2, color: 'text-emerald-400 bg-emerald-900/20' },
                            { label: 'Atrasadas', value: `${Math.round((dashboardStats.delayedCount / (dashboardStats.total || 1)) * 100)}%`, icon: AlertTriangle, color: 'text-red-400 bg-red-900/20' },
                            { label: 'Adiantadas', value: `${Math.round((dashboardStats.earlyCount / (dashboardStats.total || 1)) * 100)}%`, icon: TrendingUp, color: 'text-blue-400 bg-blue-900/20' }
                        ].map((stat, idx) => (
                            <div key={idx} className="p-6 rounded-[2rem] bg-[#18181b] border border-slate-800 flex flex-col gap-2 shadow-lg">
                                <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${stat.color}`}><stat.icon size={20} /></div>
                                <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest mt-2">{stat.label}</span>
                                <div className="text-3xl font-black text-white tracking-tighter">{stat.value}</div>
                            </div>
                        ))}
                    </div>
                    <div className="grid grid-cols-2 gap-8">
                        <div className="p-6 rounded-[2rem] bg-[#18181b] border border-slate-800 shadow-xl h-[400px] flex flex-col">
                            <h4 className="text-slate-200 font-black uppercase text-xs tracking-widest mb-6 flex items-center gap-2"><PieChartIcon size={16} className="text-blue-500" /> Distribuição de Status</h4>
                            <div className="flex-1">
                                <ResponsiveContainer width="100%" height="100%">
                                    <PieChart>
                                        <Pie data={dashboardStats.statusPie} innerRadius={80} outerRadius={110} paddingAngle={5} dataKey="value">{dashboardStats.statusPie.map((entry, index) => (<Cell key={`cell-${index}`} fill={entry.color} />))}</Pie>
                                        <Tooltip contentStyle={{ backgroundColor: '#121214', border: 'none', borderRadius: '10px', fontSize: '12px', fontWeight: 'bold' }} />
                                        <Legend verticalAlign="bottom" height={36} formatter={(value) => <span className="text-slate-300 font-bold uppercase text-[10px]">{value}</span>} />
                                    </PieChart>
                                </ResponsiveContainer>
                            </div>
                        </div>
                        <div className="p-6 rounded-[2rem] bg-[#18181b] border border-slate-800 shadow-xl h-[400px] flex flex-col">
                            <h4 className="text-slate-200 font-black uppercase text-xs tracking-widest mb-6 flex items-center gap-2"><BarChart3 size={16} className="text-yellow-500" /> Motivos de Desvio</h4>
                            <div className="flex-1">
                                <ResponsiveContainer width="100%" height="100%">
                                    <BarChart data={dashboardStats.reasonData} layout="vertical">
                                        <XAxis type="number" hide />
                                        <YAxis dataKey="name" type="category" width={120} tick={{ fill: '#8D8D99', fontSize: 10, fontWeight: 'bold' }} axisLine={false} tickLine={false} />
                                        <Tooltip contentStyle={{ backgroundColor: '#121214', border: 'none', borderRadius: '10px', fontSize: '12px' }} cursor={{ fill: 'rgba(255,255,255,0.05)' }} />
                                        <Bar dataKey="value" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={20} />
                                    </BarChart>
                                </ResponsiveContainer>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
      )}

      {isLinkModalOpen && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-xl z-[250] flex items-center justify-center p-4">
            <div className="bg-[#121214] border border-slate-800 rounded-3xl shadow-2xl w-full max-w-lg flex flex-col overflow-hidden animate-in zoom-in duration-300">
                <div className="bg-blue-700 p-6 flex justify-between items-center text-white"><div className="flex items-center gap-3"><Link size={24} className="bg-white/20 p-2 rounded-xl" /><h3 className="font-black uppercase tracking-widest text-xs">Vínculo de Operação</h3></div></div>
                <div className="p-6 overflow-y-auto max-h-[60vh] space-y-3">
                    {pendingItems.map((item, idx) => (
                        <div key={idx} className="p-4 bg-[#18181b] border border-slate-800 rounded-2xl flex items-center gap-4">
                            <div className="flex-1 min-w-0"><span className="text-[8px] text-slate-500 font-black uppercase">Rota</span><div className="font-black text-white truncate">{item.rota}</div></div>
                            <select value={item.operacao} onChange={(e) => { const newItems = [...pendingItems]; newItems[idx].operacao = e.target.value; setPendingItems(newItems); }} className="p-2 bg-[#121214] border border-slate-700 rounded-lg text-xs font-bold text-white outline-none focus:border-blue-600"><option value="">---</option>{userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}</select>
                        </div>
                    ))}
                </div>
                <div className="p-6 bg-[#18181b] border-t border-slate-800">
                    <button onClick={handleLinkPending} disabled={isSyncing} className="w-full py-4 bg-blue-600 text-white font-black uppercase text-[11px] rounded-xl shadow-xl transition-all hover:bg-blue-700 disabled:opacity-50 flex items-center justify-center gap-2">
                        {isSyncing ? <Loader2 size={16} className="animate-spin" /> : "Gravar Registros"}
                    </button>
                </div>
            </div>
        </div>
      )}

      {isImportModalOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
             <div className="bg-[#121214] border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in duration-200">
                <div className="bg-emerald-600 p-6 flex justify-between items-center text-white"><div className="flex items-center gap-3"><Upload size={20} className="bg-white/20 p-1.5 rounded-lg" /><h3 className="font-black uppercase tracking-widest text-xs">Importar Dados Excel</h3></div><button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button></div>
                <div className="p-8">
                    <textarea value={importText} onChange={e => setImportText(e.target.value)} className="w-full h-64 p-5 border-2 border-slate-800 rounded-2xl bg-[#020617] text-[11px] font-mono mb-6 focus:ring-2 focus:ring-emerald-500 outline-none text-white shadow-inner scrollbar-thin" placeholder="Cole aqui..." />
                    <button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-4 bg-emerald-600 text-white font-black uppercase tracking-widest text-[11px] rounded-xl shadow-xl flex items-center justify-center gap-3 transition-all hover:bg-emerald-700 disabled:opacity-50">{isProcessingImport ? <Loader2 size={18} className="animate-spin" /> : <span>Processar Importação</span>}</button>
                </div>
             </div>
        </div>
      )}

      {isModalOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
          <div className="bg-[#121214] border border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in">
            <div className="bg-blue-700 text-white p-6 flex justify-between items-center"><h3 className="font-black uppercase tracking-widest text-xs flex items-center gap-3"><Plus size={20} /> Novo Registro</h3><button onClick={() => setIsModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg transition-all"><X size={20} /></button></div>
            <form onSubmit={handleSubmit} className="p-8 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Data</label><input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-bold outline-none focus:border-blue-600 transition-all"/></div>
                    <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Rota</label><input type="text" required placeholder="Ex: 24001D" value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-[11px] font-black text-blue-400 outline-none focus:border-blue-600 transition-all"/></div>
                </div>
                <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Operação</label><select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-[11px] font-black text-white outline-none focus:border-blue-600"><option value="">Selecione...</option>{userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}</select></div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Motorista</label><input type="text" required value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-bold outline-none focus:border-blue-600 transition-all"/></div>
                    <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Placa</label><input type="text" required value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value.toUpperCase()})} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-black outline-none focus:border-blue-600 transition-all"/></div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Início</label><input type="text" onBlur={(e) => setFormData({...formData, inicio: formatTimeInput(e.target.value)})} defaultValue={formData.inicio} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-mono outline-none focus:border-blue-600"/></div>
                    <div className="space-y-1"><label className="text-[9px] font-black text-slate-500 uppercase">Saída</label><input type="text" onBlur={(e) => setFormData({...formData, saida: formatTimeInput(e.target.value)})} defaultValue={formData.saida} className="w-full p-3 border border-slate-800 rounded-xl bg-[#020617] text-white text-[11px] font-mono outline-none focus:border-blue-600"/></div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-blue-600 hover:bg-blue-700 text-white font-black uppercase text-[11px] rounded-xl flex items-center justify-center gap-2 shadow-xl transition-all border-b-4 border-blue-900 mt-4">{isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />} SALVAR NO SHAREPOINT</button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
