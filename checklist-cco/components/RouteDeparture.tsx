
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
  Activity, EyeOff, ChevronRight, AlignLeft, Type as TypeIcon,
  Archive
} from 'lucide-react';
import { PieChart, Pie, Cell, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, Legend } from 'recharts';

const Sparkles = ({ size = 20, className = "" }) => (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" className={className}>
        <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z"/>
        <path d="M5 3v4"/><path d="M19 17v4"/><path d="M3 5h4"/><path d="M17 19h4"/>
    </svg>
);

interface RouteConfig {
    operacao: string;
    email: string;
    tolerancia: string;
}

const MOTIVOS = [
  'Fábrica', 'Logística', 'Mão de obra', 'Manutenção', 'Divergência de Roteirização', 'Solicitado pelo Cliente', 'Infraestrutura'
];

const OBSERVATION_TEMPLATES: Record<string, string[]> = {
  'Fábrica': ["Atraso na descarga | Entrada **:**h - Saída **:**h"],
  'Logística': ["Atraso no lavador | Chegada da rota anterior às **:**h - Entrada na fábrica às **:**h", "Motorista adiantou a rota devido à desvios", "Atraso na rota anterior (nome da rota)", "Atraso na rota anterior | Chegada no lavador **:**h - Entrada na fábrica às **:**h", "Falta de material de coleta para realizar a rota"],
  'Mão de obra': ["Atraso do motorista", "Adiantamento do motorista", "A rota iniciou atrasada devido à interjornada do motorista | Atrasou na rota anterior devido à", "Troca do motorista previsto devido à saúde"],
  'Manutenção': ["Precisou realizar a troca de pneus | Início **:**h - Término **:**h", "Troca de mola | Início **:**h - Término **:**h", "Manutenção na parte elétrica | Início **:**h - Término **:**h", "Manutenção nos freios | Início **:**h - Término **:**h", "Manutenção na bomba de carregamento de leite | Início **:**h - Término **:**h"],
  'Divergência de Roteirização': ["Horário de saída da rota não atende os produtores", "Horário de saída da rota precisa ser alterado devido à entrada de produtores"],
  'Solicitado pelo Cliente': ["Rota saiu adiantada para realizar socorro", "Cliente solicitou para a rota sair adiantada"],
  'Infraestrutura': []
};

const FilterDropdown = ({ col, routes, colFilters, setColFilters, selectedFilters, setSelectedFilters, onClose, innerRef }: any) => {
    const values: string[] = Array.from(new Set(routes.map((r: any) => String(r[col] || "")))).sort() as string[];
    const selected = (selectedFilters[col] as string[]) || [];
    const toggleValue = (val: string) => { const next = selected.includes(val) ? selected.filter(v => v !== val) : [...selected, val]; setSelectedFilters({ ...selectedFilters, [col]: next }); };
    return (
        <div ref={innerRef} className="absolute top-10 left-0 z-[100] bg-white border border-slate-200 shadow-xl rounded-xl w-64 p-3 text-slate-700 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 rounded-lg border border-slate-200">
                <Search size={14} className="text-slate-400" />
                <input type="text" placeholder="Filtrar..." autoFocus value={colFilters[col] || ""} onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })} className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800" />
            </div>
            <div className="max-h-56 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-100 py-2">
                {values.filter(v => v.toLowerCase().includes((colFilters[col] || "").toLowerCase())).map(v => (
                    <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2 hover:bg-slate-50 rounded-lg cursor-pointer transition-all">
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} className="text-slate-300" />}
                        <span className="text-[10px] font-bold uppercase truncate text-slate-600">{v || "(VAZIO)"}</span>
                    </div>
                ))}
            </div>
            <button onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); onClose(); }} className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 hover:bg-red-100 rounded-lg border border-red-100 transition-colors"> Limpar Filtro </button>
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
  const [isStatsModalOpen, setIsStatsModalOpen] = useState(false);
  const [isProcessingImport, setIsProcessingImport] = useState(false);
  const [importText, setImportText] = useState('');
  const [currentTime, setCurrentTime] = useState(new Date());
  
  const [zoomLevel] = useState(0.9);
  const [activeObsId, setActiveObsId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [isTextWrapEnabled, setIsTextWrapEnabled] = useState(false);
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});
  const [colWidths, setColWidths] = useState<Record<string, number>>({ select: 35, rota: 140, data: 125, inicio: 95, motorista: 230, placa: 100, saida: 95, motivo: 170, observacao: 400, geral: 70, operacao: 140, status: 90, tempo: 90 });

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
      setRoutes(fixedData);
    } catch (e: any) { console.error(e); } finally { setIsLoading(false); }
  };

  const handleArchiveFiltered = async () => {
    const visibleRoutes = filteredRoutes;
    if (visibleRoutes.length === 0) { alert("Nenhum item visível para arquivar."); return; }
    if (confirm(`Atenção: Você está prestes a mover ${visibleRoutes.length} rotas para o Histórico Definitivo. Elas sairão desta grade imediatamente. Confirma?`)) {
        const token = getAccessToken();
        setIsSyncing(true);
        try {
            const result = await SharePointService.moveDeparturesToHistory(token, visibleRoutes);
            if (result.failed > 0) {
              alert(`Arquivamento concluído com Falhas Parciais!\n\n- Sucesso: ${result.success}\n- Falhas: ${result.failed}\n\nMOTIVO TÉCNICO: ${result.lastError}\n\nVerifique o console (F12) para o payload exato.`);
            } else {
              alert(`Sucesso absoluto! ${result.success} rotas movidas para o histórico.`);
            }
            await loadData();
        } catch (err: any) { 
            alert("ERRO CRÍTICO NO ARQUIVAMENTO:\n" + err.message); 
        } finally { 
            setIsSyncing(false); 
        }
    }
  };

  useEffect(() => { loadData(); const timer = setInterval(() => setCurrentTime(new Date()), 10000); return () => clearInterval(timer); }, [currentUser]);

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => { if (resizingRef.current) { const { col, startX, startWidth } = resizingRef.current; const newWidth = Math.max(10, startWidth + (e.clientX - startWidth)); setColWidths(prev => ({ ...prev, [col]: newWidth })); } };
    const handleMouseUp = () => { resizingRef.current = null; };
    const handleClickOutside = (e: MouseEvent) => { if (filterRef.current && !filterRef.current.contains(e.target as Node)) { setActiveFilterCol(null); } if (obsDropdownRef.current && !obsDropdownRef.current.contains(e.target as Node)) { setActiveObsId(null); } };
    const handleKeyDown = (e: KeyboardEvent) => {
        const target = e.target as HTMLElement;
        if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.tagName === 'SELECT') return;

        if (e.ctrlKey && e.shiftKey && e.key.toUpperCase() === 'L') { e.preventDefault(); setColFilters({}); setSelectedFilters({}); setSelectedIds(new Set()); }
        if (e.key === 'Delete' && selectedIds.size > 0) {
            if (confirm(`Excluir permanentemente os ${selectedIds.size} itens selecionados?`)) {
                const token = getAccessToken();
                setIsSyncing(true);
                Promise.all(Array.from(selectedIds).map((id: string) => SharePointService.deleteDeparture(token, id))).then(() => {
                   alert("Exclusão concluída.");
                   loadData();
                }).finally(() => setIsSyncing(false));
            }
        }
    };
    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    window.addEventListener('mousedown', handleClickOutside);
    window.addEventListener('keydown', handleKeyDown);
    return () => { window.removeEventListener('mousemove', handleMouseMove); window.removeEventListener('mouseup', handleMouseUp); window.removeEventListener('mousedown', handleClickOutside); window.removeEventListener('keydown', handleKeyDown); };
  }, [selectedIds, loadData]);

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

  const timeToSeconds = (timeStr: string): number => { if (!timeStr || !timeStr.includes(':')) return 0; const parts = timeStr.split(':').map(Number); return (parts[0] || 0) * 3600 + (parts[1] || 0) * 60 + (parts[2] || 0); };
  const secondsToTime = (totalSeconds: number): string => { const isNegative = totalSeconds < 0; const absSeconds = Math.abs(totalSeconds); const h = Math.floor(absSeconds / 3600); const m = Math.floor((absSeconds % 3600) / 60); const s = absSeconds % 60; const formatted = [h, m, s].map(v => v.toString().padStart(2, '0')).join(':'); return isNegative ? `-${formatted}` : formatted; };

  const calculateGap = (inicio: string, saida: string, toleranceStr: string = "00:00:00"): { gap: string, status: string, isOutOfTolerance: boolean } => {
    const sInicio = inicio || '00:00:00'; const sSaida = saida || '00:00:00'; if (sInicio === '00:00:00' || sSaida === '00:00:00') return { gap: 'OK', status: 'OK', isOutOfTolerance: false };
    const startSec = timeToSeconds(sInicio); const endSec = timeToSeconds(sSaida); const diff = endSec - startSec; const toleranceSec = timeToSeconds(toleranceStr || "00:00:00");
    const gapFormatted = secondsToTime(diff); const isOutOfTolerance = Math.abs(diff) > toleranceSec; const status = isOutOfTolerance ? (diff > 0 ? 'Atrasado' : 'Adiantado') : 'OK';
    return { gap: gapFormatted, status, isOutOfTolerance };
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    const token = getAccessToken();
    const route = routes.find(r => r.id === id);
    if (!route) return;
    let finalValue = value;
    if (field === 'inicio' || field === 'saida') finalValue = formatTimeInput(value);
    let updatedRoute = { ...route, [field]: finalValue };
    const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (updatedRoute.operacao || "").toUpperCase().trim());
    if (field === 'inicio' || field === 'saida' || field === 'operacao') {
        const { gap, status } = calculateGap(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00");
        updatedRoute.tempo = gap; updatedRoute.statusOp = status;
    }
    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));
    setIsSyncing(true);
    try { await SharePointService.updateDeparture(token, updatedRoute); } catch (err: any) { console.error(err); } finally { setIsSyncing(false); }
  };

  const filteredRoutes = useMemo(() => {
    return routes.filter(r => {
        return (Object.entries(colFilters) as [string, string][]).every(([col, val]) => {
            if (!val) return true; const field = r[col as keyof RouteDeparture]?.toString().toLowerCase() || ""; return field.includes(val.toLowerCase());
        }) && (Object.entries(selectedFilters) as [string, string[]][]).every(([col, vals]) => {
            if (!vals || vals.length === 0) return true; const field = r[col as keyof RouteDeparture]?.toString() || ""; return vals.includes(field);
        });
    });
  }, [routes, colFilters, selectedFilters]);

  const dashboardStats = useMemo(() => {
    const total = filteredRoutes.length; if (total === 0) return null;
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
    if (!importText.trim()) return; setIsProcessingImport(true);
    try {
        const parsed = parseRouteDeparturesManual(importText); if (parsed.length === 0) throw new Error("Sem dados válidos.");
        const token = getAccessToken();
        for (const item of parsed) {
            const rotaStr = (item.rota || "").trim();
            const mapping = routeMappings.find(m => (m.Title || "").trim() === rotaStr);
            const op = (mapping?.OPERACAO || "").toUpperCase().trim();
            const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === op);
            const { gap, status } = calculateGap(item.inicio || '00:00:00', item.saida || '00:00:00', config?.tolerancia || "00:00:00");
            await SharePointService.updateDeparture(token, { ...item, id: '', statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture);
        }
        await loadData(); setIsImportModalOpen(false);
    } catch (e: any) { alert(e.message); } finally { setIsProcessingImport(false); }
  };

  const getAlertStyles = (route: RouteDeparture) => {
    const isDelayed = route.statusOp === 'Atrasado';
    const isEarly = route.statusOp === 'Adiantado';
    // Estilo de linha ATRASADA muito evidente
    if (route.saida !== '00:00:00' && isDelayed) return "bg-orange-200 border-l-[8px] border-orange-600 shadow-inner";
    if (isEarly) return "bg-blue-100 border-l-[8px] border-blue-600";
    return "border-l-4 border-transparent";
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault(); const token = getAccessToken(); setIsSyncing(true);
    try {
        const config = userConfigs.find(c => (c.operacao || "").toUpperCase().trim() === (formData.operacao || "").toUpperCase().trim());
        const { gap, status } = calculateGap(formData.inicio || '00:00:00', formData.saida || '00:00:00', config?.tolerancia || "00:00:00");
        const newRoute: RouteDeparture = { ...formData, id: '', statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
        await SharePointService.updateDeparture(token, newRoute); await loadData(); setIsModalOpen(false);
        setFormData({ rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '00:00:00', motorista: '', placa: '', operacao: '', motivo: '', observacao: '', statusGeral: 'OK', aviso: 'NÃO' });
    } catch (err: any) { alert(err.message); } finally { setIsSyncing(false); }
  };

  const toggleSelection = (id: string) => { const newSelected = new Set(selectedIds); if (newSelected.has(id)) newSelected.delete(id); else newSelected.add(id); setSelectedIds(newSelected); };

  if (isLoading) return ( <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4 bg-[#020617]"><Loader2 size={48} className="animate-spin" /><p className="font-bold text-[10px] uppercase tracking-[0.3em] text-slate-400">CCO Logística...</p></div> );

  return (
    <div className="flex flex-col h-full animate-fade-in bg-[#020617] p-4 overflow-hidden select-none font-sans">
      <div className="flex justify-between items-center mb-6 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-primary-600 text-white rounded-2xl shadow-lg shadow-primary-600/20"><Clock size={20} /></div>
          <div>
            <h2 className="text-xl font-black text-white uppercase tracking-tight flex items-center gap-3">Controle de Saídas {isSyncing && <Loader2 size={16} className="animate-spin text-primary-500"/>}</h2>
            <div className="flex items-center gap-2"><ShieldCheck size={12} className="text-emerald-500"/><p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest">Responsável: {currentUser.name}</p></div>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <button onClick={() => setIsTextWrapEnabled(!isTextWrapEnabled)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] tracking-wide shadow-sm transition-all ${isTextWrapEnabled ? 'bg-primary-600 text-white border-primary-600' : 'bg-slate-800 text-slate-300 border-slate-700 hover:bg-slate-700'}`}><AlignLeft size={16} /> Quebra</button>
          <button onClick={() => setIsStatsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm"><BarChart3 size={16} /> Indicadores</button>
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-700 bg-slate-900"><RefreshCw size={18} /></button>
          <button onClick={handleArchiveFiltered} disabled={isSyncing || filteredRoutes.length === 0} className="flex items-center gap-2 px-4 py-2 bg-slate-900 text-slate-300 rounded-lg hover:bg-slate-800 font-bold border border-slate-700 uppercase text-[10px] tracking-wide shadow-sm transition-all disabled:opacity-30"><Archive size={16} /> Arquivar</button>
          <button onClick={() => setIsImportModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 font-bold border border-emerald-700 uppercase text-[10px] tracking-wide shadow-sm transition-all"><Upload size={16} /> Importar</button>
          <button onClick={() => setIsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-primary-600 text-white rounded-lg hover:bg-primary-700 font-bold border border-primary-700 uppercase text-[10px] tracking-wide shadow-md transition-all"><Plus size={16} /> Nova Rota</button>
        </div>
      </div>

      <div ref={tableContainerRef} className="flex-1 overflow-auto bg-white rounded-2xl border border-slate-700/50 shadow-2xl relative scrollbar-thin overflow-x-auto">
        <div style={{ transform: `scale(${zoomLevel})`, transformOrigin: 'top left', width: `${100 / zoomLevel}%` }}>
            <table className="border-collapse table-fixed w-full min-w-max h-px">
              <thead className="sticky top-0 z-50 bg-[#1e293b] text-white shadow-md">
                <tr className="h-12">
                  {[ { id: 'select', label: '' }, { id: 'rota', label: 'ROTA' }, { id: 'data', label: 'DATA' }, { id: 'inicio', label: 'INÍCIO' }, { id: 'motorista', label: 'MOTORISTA' }, { id: 'placa', label: 'PLACA' }, { id: 'saida', label: 'SAÍDA' }, { id: 'motivo', label: 'MOTIVO' }, { id: 'observacao', label: 'OBSERVAÇÃO' }, { id: 'geral', label: 'GERAL' }, { id: 'operacao', label: 'OPERAÇÃO' }, { id: 'status', label: 'STATUS' }, { id: 'tempo', label: 'TEMPO' } ].map(col => {
                    if (col.id === 'select') return <th key={col.id} style={{ width: colWidths.select }} className="bg-slate-900/50 border border-slate-700/50"></th>;
                    const hasFilter = !!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0;
                    return (
                      <th key={col.id} style={{ width: colWidths[col.id] }} className="relative p-1 border border-slate-700/50 text-[10px] font-black uppercase tracking-wider text-left select-none group">
                        <div className="flex items-center justify-between px-2 h-full"><span>{col.label}</span><button onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }} className={`p-1 rounded transition-all ${hasFilter ? 'text-yellow-400' : 'text-white/40 hover:text-white/60'}`}><Filter size={11} /></button></div>
                        {activeFilterCol === col.id && <FilterDropdown col={col.id} routes={routes} colFilters={colFilters} setColFilters={setColFilters} selectedFilters={selectedFilters} setSelectedFilters={setSelectedFilters} onClose={() => setActiveFilterCol(null)} innerRef={filterRef} />}
                        <div onMouseDown={(e) => startResize(e, col.id)} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300">
                {filteredRoutes.map((route) => {
                  const alertClasses = getAlertStyles(route);
                  const isSelected = selectedIds.has(route.id);
                  const rowBg = isSelected ? 'bg-primary-100/90' : alertClasses || 'bg-white hover:bg-slate-50';
                  const inputClass = "w-full h-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold text-slate-800 uppercase transition-all placeholder-slate-400 focus:bg-white/80 focus:ring-1 focus:ring-primary-500/20";
                  const cellClass = "p-0 border border-slate-300 transition-all overflow-hidden";
                  const isDelayed = route.statusOp === 'Atrasado';

                  return (
                    <tr key={route.id} className={`${rowBg} group transition-all h-auto`}>
                      <td className={`${cellClass} cursor-pointer transition-colors w-[35px] ${isSelected ? 'bg-primary-600' : 'hover:bg-slate-200'}`} onClick={() => toggleSelection(route.id)}></td>
                      <td className={cellClass}><input type="text" value={route.rota} onChange={(e) => updateCell(route.id, 'rota', e.target.value)} className={`${inputClass} font-black text-primary-700`} /></td>
                      <td className={cellClass}><input type="date" value={route.data} onChange={(e) => updateCell(route.id, 'data', e.target.value)} className={`${inputClass} text-center text-slate-600`} /></td>
                      <td className={cellClass}><input type="text" value={route.inicio} onBlur={(e) => updateCell(route.id, 'inicio', e.target.value)} className={`${inputClass} font-mono text-center`} /></td>
                      <td className={cellClass}><input type="text" value={route.motorista} onChange={(e) => updateCell(route.id, 'motorista', e.target.value)} className={`${inputClass}`} /></td>
                      <td className={cellClass}><input type="text" value={route.placa} onChange={(e) => updateCell(route.id, 'placa', e.target.value)} className={`${inputClass} font-mono text-center`} /></td>
                      <td className={cellClass}><input type="text" value={route.saida} onBlur={(e) => updateCell(route.id, 'saida', e.target.value)} className={`${inputClass} font-mono text-center`} /></td>
                      <td className={cellClass}>
                        <div className="flex items-center justify-center h-full px-1">
                            <select value={route.motivo} onChange={(e) => updateCell(route.id, 'motivo', e.target.value)} className="w-full bg-white/40 border border-slate-200 rounded-md px-2 py-1 text-[10px] font-bold text-slate-700 outline-none shadow-sm appearance-none">
                                <option value="">---</option>{MOTIVOS.map(m => (<option key={m} value={m}>{m}</option>))}
                            </select>
                        </div>
                      </td>
                      <td className={`${cellClass} relative group/obs align-top h-full min-h-[44px]`}>
                        <div className="flex items-start w-full h-full relative p-0 min-h-[44px]">
                          <textarea value={route.observacao || ""} onChange={(e) => updateCell(route.id, 'observacao', e.target.value)} onFocus={() => setActiveObsId(route.id)} placeholder="..." className={`w-full h-full min-h-[44px] bg-transparent outline-none border-none px-3 py-2 text-[11px] font-normal text-slate-800 placeholder-slate-500 resize-none overflow-hidden ${isTextWrapEnabled ? 'whitespace-normal break-words leading-relaxed' : 'truncate pr-8'}`} style={{ height: isTextWrapEnabled ? 'auto' : '44px' }} onInput={(e) => { if (isTextWrapEnabled) { const el = e.target as HTMLTextAreaElement; el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px'; } }} />
                          {!isTextWrapEnabled && <button onClick={(e) => { e.stopPropagation(); setActiveObsId(activeObsId === route.id ? null : route.id); }} className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 text-slate-500 hover:text-primary-700 transition-colors opacity-60 group-hover/obs:opacity-100"><ChevronDown size={14} /></button>}
                        </div>
                        {activeObsId === route.id && (
                          <div ref={obsDropdownRef} className="absolute top-full left-0 w-full z-[110] bg-white border border-slate-300 rounded-xl shadow-2xl overflow-hidden animate-in fade-in slide-in-from-top-1">
                            <div className="p-2 border-b border-slate-100 flex justify-between items-center bg-slate-50"><span className="text-[9px] font-black uppercase text-slate-500">Auto-Completar</span><X size={12} className="text-slate-400 cursor-pointer" onClick={() => setActiveObsId(null)} /></div>
                            <div className="max-h-48 overflow-y-auto scrollbar-thin">{(route.motivo ? (OBSERVATION_TEMPLATES[route.motivo] || []) : []).map((template, tIdx) => ( <div key={tIdx} onClick={() => { updateCell(route.id, 'observacao', template); setActiveObsId(null); }} className="p-3 text-[10px] text-slate-700 hover:bg-primary-100 cursor-pointer border-b border-slate-100 flex items-center gap-2"><ChevronRight size={12} className="shrink-0 text-primary-500" />{template}</div> ))}</div>
                          </div>
                        )}
                      </td>
                      <td className={cellClass}><select value={route.statusGeral} onChange={(e) => updateCell(route.id, 'statusGeral', e.target.value)} className="w-full h-full bg-transparent border-none text-[10px] font-bold text-center appearance-none text-slate-800"><option value="OK">OK</option><option value="NOK">NOK</option></select></td>
                      <td className={`${cellClass} bg-slate-50/50`}>
                          <select value={route.operacao} onChange={(e) => updateCell(route.id, 'operacao', e.target.value)} className="w-full h-full bg-transparent border-none text-[9px] font-black text-center text-slate-600 uppercase">
                              <option value="">OP...</option>
                              {userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}
                          </select>
                      </td>
                      <td className={`${cellClass} text-center`}><span className={`px-2 py-0.5 rounded-full text-[8px] font-black border ${route.statusOp === 'OK' ? 'bg-emerald-100 border-emerald-400 text-emerald-800' : 'bg-red-100 border-red-400 text-red-800'}`}>{route.statusOp}</span></td>
                      <td className={`${cellClass} text-center font-mono font-bold text-[10px] text-slate-700`}>{route.tempo}</td>
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
                <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white"><div className="flex items-center gap-4"><div className="p-2.5 bg-white/10 rounded-xl"><TrendingUp size={24} /></div><h3 className="font-black uppercase tracking-widest text-base">Dashboard Operacional</h3></div><button onClick={() => setIsStatsModalOpen(false)} className="hover:bg-white/10 p-2 rounded-xl transition-all"><X size={28} /></button></div>
                <div className="p-8 flex-1 overflow-y-auto space-y-8 bg-slate-50">
                    <div className="grid grid-cols-4 gap-6">{[{ label: 'Total', value: dashboardStats.total, icon: Activity, color: 'text-slate-700 bg-white' }, { label: 'OK', value: `${Math.round((dashboardStats.okCount / dashboardStats.total) * 100)}%`, icon: CheckCircle2, color: 'text-emerald-600 bg-emerald-50' }, { label: 'Atrasos', value: `${Math.round((dashboardStats.delayedCount / dashboardStats.total) * 100)}%`, icon: AlertTriangle, color: 'text-orange-600 bg-orange-50' }].map((stat, idx) => ( <div key={idx} className={`p-6 rounded-2xl border border-slate-200 flex flex-col gap-2 ${stat.color}`}><stat.icon size={20} /><span className="text-[10px] font-black uppercase text-slate-400 mt-2">{stat.label}</span><div className="text-3xl font-black">{stat.value}</div></div> ))}</div>
                </div>
            </div>
        </div>
      )}

      {isImportModalOpen && (
        <div className="fixed inset-0 bg-slate-950/60 backdrop-blur-md z-[200] flex items-center justify-center p-4">
             <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-2xl w-full max-w-2xl overflow-hidden animate-in zoom-in">
                <div className="bg-emerald-600 p-6 flex justify-between items-center text-white font-black uppercase tracking-widest text-xs"><div className="flex items-center gap-3"><Upload size={20} /> Importar Dados</div><button onClick={() => setIsImportModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg"><X size={20} /></button></div>
                <div className="p-8"><textarea value={importText} onChange={e => setImportText(e.target.value)} className="w-full h-64 p-5 border-2 border-slate-100 rounded-2xl bg-slate-50 text-[11px] font-mono mb-6 outline-none shadow-inner" placeholder="Cole aqui..." /><button onClick={handleImport} disabled={isProcessingImport || !importText.trim()} className="w-full py-4 bg-emerald-600 text-white font-black uppercase text-[11px] rounded-xl transition-all hover:bg-emerald-700 shadow-lg disabled:opacity-50">{isProcessingImport ? <Loader2 size={18} className="animate-spin" /> : 'Processar Agora'}</button></div>
             </div>
        </div>
      )}

      {isModalOpen && (
        <div className="fixed inset-0 bg-slate-950/60 backdrop-blur-md z-[200] flex items-center justify-center p-4">
          <div className="bg-white border border-slate-200 rounded-[2.5rem] shadow-2xl w-full max-w-lg overflow-hidden animate-in zoom-in">
            <div className="bg-primary-600 text-white p-6 flex justify-between items-center font-black uppercase text-xs"><div className="flex items-center gap-3"><Plus size={20} /> Nova Rota</div><button onClick={() => setIsModalOpen(false)} className="hover:bg-white/10 p-1.5 rounded-lg"><X size={20} /></button></div>
            <form onSubmit={handleSubmit} className="p-8 space-y-4">
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase">Data</label><input type="date" required value={formData.data} onChange={e => setFormData({...formData, data: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-800 text-[11px] font-bold outline-none"/></div>
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase">Rota</label><input type="text" required value={formData.rota} onChange={e => setFormData({...formData, rota: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 text-[11px] font-black text-primary-600 outline-none"/></div>
                </div>
                <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase">Operação</label><select required value={formData.operacao} onChange={e => setFormData({...formData, operacao: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 text-[11px] font-black text-slate-700 outline-none"><option value="">Selecione...</option>{userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}</select></div>
                <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase">Motorista</label><input type="text" required value={formData.motorista} onChange={e => setFormData({...formData, motorista: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-800 text-[11px] font-bold outline-none"/></div>
                    <div className="space-y-1"><label className="text-[10px] font-black text-slate-400 uppercase">Placa</label><input type="text" required value={formData.placa} onChange={e => setFormData({...formData, placa: e.target.value})} className="w-full p-3 border border-slate-200 rounded-xl bg-slate-50 text-[11px] font-black outline-none"/></div>
                </div>
                <button type="submit" disabled={isSyncing} className="w-full py-4 bg-primary-600 hover:bg-primary-700 text-white font-black uppercase text-[11px] rounded-xl flex items-center justify-center gap-2 shadow-xl transition-all mt-4">{isSyncing ? <Loader2 size={16} className="animate-spin" /> : <Save size={16} />} SALVAR NO SHAREPOINT</button>
            </form>
          </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
