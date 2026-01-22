
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { RouteDeparture, User, RouteOperationMapping, RouteConfig } from '../types';
import { SharePointService } from '../services/sharepointService';
import { 
  Clock, X, Loader2, RefreshCw, ShieldCheck,
  AlertTriangle, CheckCircle2, ChevronDown, 
  Filter, Search, CheckSquare, Square,
  BarChart3, TrendingUp,
  Activity, ChevronRight, AlignLeft,
  Archive, Database, Save, Link as LinkIcon,
  Layers, Trash2, Settings2
} from 'lucide-react';

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
        <div ref={innerRef} className="absolute top-10 left-0 z-[100] bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xl rounded-xl w-64 p-3 text-slate-700 dark:text-slate-300 animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-center gap-2 mb-3 p-2 bg-slate-50 dark:bg-slate-900 rounded-lg border border-slate-200 dark:border-slate-700">
                <Search size={14} className="text-slate-400" />
                <input type="text" placeholder="Filtrar..." autoFocus value={colFilters[col] || ""} onChange={e => setColFilters({ ...colFilters, [col]: e.target.value })} className="w-full bg-transparent outline-none text-[10px] font-bold text-slate-800 dark:text-white" />
            </div>
            <div className="max-h-56 overflow-y-auto space-y-1 scrollbar-thin border-t border-slate-100 dark:border-slate-700 py-2">
                {values.filter(v => v.toLowerCase().includes((colFilters[col] || "").toLowerCase())).map(v => (
                    <div key={v} onClick={() => toggleValue(v)} className="flex items-center gap-2 p-2 hover:bg-slate-50 dark:hover:bg-slate-700 rounded-lg cursor-pointer transition-all">
                        {selected.includes(v) ? <CheckSquare size={14} className="text-blue-600" /> : <Square size={14} className="text-slate-300" />}
                        <span className="text-[10px] font-bold uppercase truncate">{v || "(VAZIO)"}</span>
                    </div>
                ))}
            </div>
            <button onClick={() => { setColFilters({ ...colFilters, [col]: "" }); setSelectedFilters({ ...selectedFilters, [col]: [] }); onClose(); }} className="w-full mt-2 py-2 text-[10px] font-black uppercase text-red-600 bg-red-50 dark:bg-red-900/30 hover:bg-red-100 rounded-lg border border-red-100 dark:border-red-900/50 transition-colors"> Limpar Filtro </button>
        </div>
    );
};

const RouteDepartureView: React.FC<{ currentUser: User }> = ({ currentUser }) => {
  const [routes, setRoutes] = useState<RouteDeparture[]>([]);
  const [userConfigs, setUserConfigs] = useState<RouteConfig[]>([]);
  const [routeMappings, setRouteMappings] = useState<RouteOperationMapping[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [zoomLevel] = useState(0.9);
  const [currentTime, setCurrentTime] = useState(new Date());

  // Bulk state para criação de rotas
  const [bulkStatus, setBulkStatus] = useState<{ active: boolean, current: number, total: number } | null>(null);
  const [pendingBulkRoutes, setPendingBulkRoutes] = useState<string[]>([]);
  const [isBulkMappingModalOpen, setIsBulkMappingModalOpen] = useState(false);

  // Ghost Row State - Saída inicia VAZIA e status Programada
  const [ghostRow, setGhostRow] = useState<Partial<RouteDeparture>>({
    id: 'ghost', rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '', motorista: '', placa: '', statusGeral: 'OK', aviso: 'NÃO', operacao: '', statusOp: 'Programada', tempo: '', semana: ''
  });

  const [isStatsModalOpen, setIsStatsModalOpen] = useState(false);
  const [isHistoryModalOpen, setIsHistoryModalOpen] = useState(false);
  const [isMappingModalOpen, setIsMappingModalOpen] = useState(false);
  const [pendingMappingRoute, setPendingMappingRoute] = useState<string | null>(null);
  
  const [histStart, setHistStart] = useState(new Date().toISOString().split('T')[0]);
  const [histEnd, setHistEnd] = useState(new Date().toISOString().split('T')[0]);
  const [archivedResults, setArchivedResults] = useState<RouteDeparture[]>([]);
  const [isSearchingArchive, setIsSearchingArchive] = useState(false);

  const [activeObsId, setActiveObsId] = useState<string | null>(null);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [isTextWrapEnabled, setIsTextWrapEnabled] = useState(false);
  const [activeFilterCol, setActiveFilterCol] = useState<string | null>(null);
  const [colFilters, setColFilters] = useState<Record<string, string>>({});
  const [selectedFilters, setSelectedFilters] = useState<Record<string, string[]>>({});
  const [colWidths, setColWidths] = useState<Record<string, number>>({ rota: 140, data: 125, inicio: 95, motorista: 230, placa: 100, saida: 95, motivo: 170, observacao: 400, geral: 70, operacao: 140, status: 90, tempo: 90 });

  const filterRef = useRef<HTMLDivElement>(null);
  const obsDropdownRef = useRef<HTMLDivElement>(null);
  const resizingRef = useRef<{ col: string; startX: number; startWidth: number } | null>(null);

  const getAccessToken = (): string => (window as any).__access_token || '';

  // Atualiza o relógio interno para cálculos de atraso em tempo real
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 30000);
    return () => clearInterval(timer);
  }, []);

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

  const calculateStatusWithTolerance = (inicio: string, saida: string, toleranceStr: string = "00:00:00", routeDate: string): { status: string, gap: string } => {
    if (!inicio || inicio === '00:00:00') return { status: 'Pendente', gap: '' };
    if (!routeDate) return { status: 'Pendente', gap: '' };

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const [y, m, d] = routeDate.split('-').map(Number);
    const rDate = new Date(y, m - 1, d);
    rDate.setHours(0, 0, 0, 0);

    // CASO 0: DATA FUTURA
    if (rDate > today) return { status: 'Programada', gap: '' };

    const toleranceSec = timeToSeconds(toleranceStr);
    const startSec = timeToSeconds(inicio);

    // CASO 1: NÃO SAIU AINDA (Saída Vazia ou 00:00:00)
    if (!saida || saida === '00:00:00' || saida === '') {
      // Se a data já passou, está atrasada independente do horário
      if (rDate < today) return { status: 'Atrasada', gap: '' };

      // Se for hoje, checa o horário atual contra o início + tolerância
      const nowSec = currentTime.getHours() * 3600 + currentTime.getMinutes() * 60 + currentTime.getSeconds();
      if (nowSec > (startSec + toleranceSec)) {
        return { status: 'Atrasada', gap: '' };
      }
      return { status: 'Pendente', gap: '' };
    }

    // CASO 2: SAIU (Saída preenchida)
    const endSec = timeToSeconds(saida);
    const diff = endSec - startSec;
    const gapFormatted = secondsToTime(diff);

    if (diff < -toleranceSec) return { status: 'Adiantada', gap: gapFormatted };
    if (diff > toleranceSec) return { status: 'Atrasada', gap: gapFormatted };

    return { status: 'OK', gap: gapFormatted };
  };

  const formatTimeInput = (value: string): string => {
    let clean = (value || "").replace(/[^0-9:]/g, '');
    if (!clean) return '';
    const parts = clean.split(':');
    let h = (parts[0] || '00').padStart(2, '0').substring(0, 2);
    let m = (parts[1] || '00').padStart(2, '0').substring(0, 2);
    let s = (parts[2] || '00').padStart(2, '0').substring(0, 2);
    return `${h}:${m}:${s}`;
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
      setRoutes(spData || []);
    } catch (e) { console.error(e); } finally { setIsLoading(false); }
  };

  useEffect(() => { loadData(); }, [currentUser]);

  const handleDeleteRoute = async (id: string) => {
    if (!confirm('Deseja excluir permanentemente esta rota do SharePoint?')) return;
    const token = getAccessToken();
    setIsSyncing(true);
    try {
      await SharePointService.deleteDeparture(token, id);
      setRoutes(prev => prev.filter(r => r.id !== id));
      setSelectedIds(prev => { const next = new Set(prev); next.delete(id); return next; });
    } catch (e) { alert("Erro ao excluir item."); }
    finally { setIsSyncing(false); }
  };

  const handleDeleteSelected = async () => {
    if (selectedIds.size === 0) return;
    if (!confirm(`Deseja excluir as ${selectedIds.size} rotas selecionadas do SharePoint?`)) return;
    const token = getAccessToken();
    setIsSyncing(true);
    let success = 0;
    for (const id of Array.from(selectedIds)) {
        try { await SharePointService.deleteDeparture(token, id); success++; } catch (e) {}
    }
    setRoutes(prev => prev.filter(r => !selectedIds.has(r.id!)));
    setSelectedIds(new Set());
    setIsSyncing(false);
    alert(`${success} rotas excluídas.`);
  };

  const toggleSelection = (id: string) => {
    setSelectedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const handleBulkCreateSave = async (operacao: string) => {
    const token = getAccessToken();
    const total = pendingBulkRoutes.length;
    setIsBulkMappingModalOpen(false);
    setBulkStatus({ active: true, current: 0, total });
    const newRoutes: RouteDeparture[] = [];
    const config = userConfigs.find(c => c.operacao === operacao);
    for (let i = 0; i < total; i++) {
        const rotaName = pendingBulkRoutes[i];
        setBulkStatus((prev: { active: boolean, current: number, total: number } | null) => prev ? { ...prev, current: i + 1 } : null);
        const { status, gap } = calculateStatusWithTolerance(ghostRow.inicio || '00:00:00', ghostRow.saida || '', config?.tolerancia || "00:00:00", ghostRow.data || "");
        const payload: RouteDeparture = { ...ghostRow, id: '', rota: rotaName, operacao: operacao, statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
        try { const newId = await SharePointService.updateDeparture(token, payload); newRoutes.push({ ...payload, id: newId }); } catch (e) {}
    }
    setRoutes(prev => [...prev, ...newRoutes]);
    setBulkStatus(null);
    setPendingBulkRoutes([]);
    setGhostRow({ id: 'ghost', rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '', motorista: '', placa: '', statusGeral: 'OK', aviso: 'NÃO', operacao: '', statusOp: 'Programada', tempo: '' });
  };

  const handleMultilinePaste = async (field: keyof RouteDeparture, startRowIndex: number, value: string) => {
    const lines = value.split(/[\n\r]/).map(l => l.trim()).filter(Boolean);
    if (lines.length <= 1) return;
    if (!confirm(`Distribuir ${lines.length} linhas verticalmente?`)) return;
    const token = getAccessToken();
    setIsSyncing(true);
    const targetRoutes = (filteredRoutes as RouteDeparture[]).slice(startRowIndex, startRowIndex + lines.length);
    for (let i = 0; i < targetRoutes.length; i++) {
        const route = targetRoutes[i];
        let finalValue = lines[i];
        if (field === 'inicio' || field === 'saida') finalValue = formatTimeInput(finalValue);
        const updatedRoute: RouteDeparture = { ...route, [field]: finalValue };
        const config = userConfigs.find(c => c.operacao === updatedRoute.operacao);
        const { status, gap } = calculateStatusWithTolerance(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00", updatedRoute.data);
        updatedRoute.statusOp = status;
        updatedRoute.tempo = gap;
        try { await SharePointService.updateDeparture(token, updatedRoute); setRoutes(prev => prev.map(r => r.id === route.id ? updatedRoute : r)); } catch (err) {}
    }
    setIsSyncing(false);
  };

  const updateCell = async (id: string, field: keyof RouteDeparture, value: string) => {
    if (id === 'ghost') {
        if (field === 'rota' && (value.includes('\n') || value.includes(';'))) {
            const lines = value.split(/[\n;]/).map(l => l.trim()).filter(Boolean);
            if (lines.length > 1) { setPendingBulkRoutes(lines); setIsBulkMappingModalOpen(true); return; }
        }
        const updatedGhost = { ...ghostRow, [field]: value };
        if (field === 'rota' && value !== "") {
            const mapping = routeMappings.find(m => m.Title === value);
            if (mapping) updatedGhost.operacao = mapping.OPERACAO;
            else { setPendingMappingRoute(value); setIsMappingModalOpen(true); }
        }
        if (field !== 'rota' && updatedGhost.rota) {
            setIsSyncing(true);
            try {
                const config = userConfigs.find(c => c.operacao === updatedGhost.operacao);
                const { status, gap } = calculateStatusWithTolerance(updatedGhost.inicio || '00:00:00', updatedGhost.saida || '', config?.tolerancia || "00:00:00", updatedGhost.data || "");
                const payload = { ...updatedGhost, statusOp: status, tempo: gap, createdAt: new Date().toISOString() } as RouteDeparture;
                const newId = await SharePointService.updateDeparture(getAccessToken(), payload);
                setRoutes(prev => [...prev, { ...payload, id: newId }]);
                setGhostRow({ id: 'ghost', rota: '', data: new Date().toISOString().split('T')[0], inicio: '00:00:00', saida: '', motorista: '', placa: '', statusGeral: 'OK', aviso: 'NÃO', operacao: '', statusOp: 'Programada', tempo: '' });
            } catch (e) {} finally { setIsSyncing(false); }
        } else { setGhostRow(updatedGhost); }
        return;
    }

    const route = routes.find(r => r.id === id);
    if (!route) return;
    let finalValue = value;
    if (field === 'inicio' || field === 'saida') finalValue = formatTimeInput(value);
    let updatedRoute = { ...route, [field]: finalValue };
    const config = userConfigs.find(c => c.operacao === updatedRoute.operacao);
    const { status, gap } = calculateStatusWithTolerance(updatedRoute.inicio, updatedRoute.saida, config?.tolerancia || "00:00:00", updatedRoute.data);
    updatedRoute.statusOp = status;
    updatedRoute.tempo = gap;
    if (status !== 'Atrasada' && status !== 'Adiantada') { updatedRoute.motivo = ""; updatedRoute.observacao = ""; }
    setRoutes(prev => prev.map(r => r.id === id ? updatedRoute : r));
    setIsSyncing(true);
    try { await SharePointService.updateDeparture(getAccessToken(), updatedRoute); } catch (e) {} finally { setIsSyncing(false); }
  };

  const getRowStyle = (route: RouteDeparture | Partial<RouteDeparture>) => {
    if (route.id === 'ghost') return "bg-slate-50 dark:bg-slate-900 italic text-slate-400";
    const status = route.statusOp;
    
    // ⚪ PROGRAMADA (Data Futura)
    if (status === 'Programada') {
        return "bg-slate-100 dark:bg-slate-800 border-l-4 border-slate-400 text-slate-500 dark:text-slate-400";
    }

    // ✅ OK - Verde Soft
    if (status === 'OK') return "bg-emerald-50 dark:bg-emerald-900/10 border-l-4 border-emerald-600";
    
    // ⏰ ATRASADA SEM SAÍDA (Amarelo de Alerta)
    if (status === 'Atrasada' && (!route.saida || route.saida === '00:00:00' || route.saida === '')) {
      return "bg-yellow-300 dark:bg-yellow-500/30 text-slate-900 dark:text-yellow-100 font-bold border-l-[12px] border-yellow-600 shadow-lg";
    }
    
    // 🟠 ATRASADA COM SAÍDA OU ADIANTADA (Laranja Operacional)
    if (status === 'Atrasada' || status === 'Adiantada') {
      return "bg-orange-500 dark:bg-orange-600/30 text-white font-bold border-l-[12px] border-orange-700 shadow-lg";
    }
    
    return "bg-white dark:bg-slate-900 border-l-4 border-transparent";
  };

  const filteredRoutes = useMemo(() => {
    return routes.filter(r => {
        return (Object.entries(colFilters) as [string, string][]).every(([col, val]) => {
            if (!val) return true; return r[col as keyof RouteDeparture]?.toString().toLowerCase().includes(val.toLowerCase());
        }) && (Object.entries(selectedFilters) as [string, string[]][]).every(([col, vals]) => {
            if (!vals || vals.length === 0) return true; return vals.includes(r[col as keyof RouteDeparture]?.toString() || "");
        });
    });
  }, [routes, colFilters, selectedFilters]);

  const dashboardStats = useMemo(() => {
    const total = filteredRoutes.length; if (total === 0) return null;
    const okCount = filteredRoutes.filter(r => r.statusOp === 'OK').length;
    const delayedCount = filteredRoutes.filter(r => r.statusOp === 'Atrasada').length;
    return { total, okCount, delayedCount };
  }, [filteredRoutes]);

  const handleSearchArchive = async () => {
    setIsSearchingArchive(true);
    try {
        const results = await SharePointService.getArchivedDepartures(getAccessToken(), '', histStart, histEnd);
        const myOps = new Set(userConfigs.map(c => c.operacao));
        setArchivedResults(results.filter(r => myOps.has(r.operacao)));
    } catch (err) { alert("Erro na busca."); } finally { setIsSearchingArchive(false); }
  };

  const tableColumns = [
    { id: 'rota', label: 'ROTA' }, { id: 'data', label: 'DATA' }, { id: 'inicio', label: 'INÍCIO' },
    { id: 'motorista', label: 'MOTORISTA' }, { id: 'placa', label: 'PLACA' }, { id: 'saida', label: 'SAÍDA' },
    { id: 'motivo', label: 'MOTIVO' }, { id: 'observacao', label: 'OBSERVAÇÃO' }, { id: 'geral', label: 'GERAL' },
    { id: 'operacao', label: 'OPERAÇÃO' }, { id: 'status', label: 'STATUS' }, { id: 'tempo', label: 'TEMPO' }
  ];

  if (isLoading) return <div className="h-full flex flex-col items-center justify-center text-primary-500 gap-4"><Loader2 size={48} className="animate-spin" /><p className="font-bold text-[10px] uppercase tracking-widest">Carregando Grid...</p></div>;

  return (
    <div className="flex flex-col h-full bg-[#020617] p-4 overflow-hidden select-none font-sans animate-fade-in relative">
      
      {bulkStatus?.active && (
          <div className="fixed inset-0 z-[500] bg-slate-950/60 backdrop-blur-sm flex items-center justify-center animate-in fade-in duration-300">
              <div className="bg-white dark:bg-slate-900 p-8 rounded-[2.5rem] border border-primary-500 shadow-2xl flex flex-col items-center gap-6 max-w-sm w-full">
                  <div className="relative"><Loader2 size={64} className="text-primary-600 animate-spin" /><Layers size={24} className="absolute inset-0 m-auto text-primary-400" /></div>
                  <div className="text-center"><h3 className="text-lg font-black uppercase text-slate-800 dark:text-white">Criando Lote</h3><p className="text-xs text-slate-400 font-bold uppercase mt-1 tracking-widest">{bulkStatus.current} de {bulkStatus.total}</p></div>
                  <div className="w-full bg-slate-200 dark:bg-slate-800 h-2 rounded-full overflow-hidden"><div className="h-full bg-primary-600 transition-all duration-300" style={{ width: `${(bulkStatus.current / bulkStatus.total) * 100}%` }}></div></div>
              </div>
          </div>
      )}

      <div className="flex justify-between items-center mb-6 shrink-0 px-2">
        <div className="flex items-center gap-4">
          <div className="p-3 bg-primary-600 text-white rounded-2xl shadow-lg"><Clock size={20} /></div>
          <div>
            <h2 className="text-xl font-black text-white uppercase tracking-tight flex items-center gap-3">Controle de Saídas {isSyncing && <Loader2 size={16} className="animate-spin text-primary-500"/>}</h2>
            <p className="text-[9px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2"><ShieldCheck size={12} className="text-emerald-500"/> Operador: {currentUser.name}</p>
          </div>
        </div>
        <div className="flex gap-2 items-center">
          <button onClick={() => setIsTextWrapEnabled(!isTextWrapEnabled)} className={`flex items-center gap-2 px-4 py-2 rounded-lg font-bold border uppercase text-[10px] transition-all ${isTextWrapEnabled ? 'bg-primary-600 text-white border-primary-600' : 'bg-slate-800 text-slate-300 border-slate-700'}`}><AlignLeft size={16} /> Quebra</button>
          <button onClick={() => setIsHistoryModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm"><Database size={16} /> Histórico</button>
          <button onClick={() => setIsStatsModalOpen(true)} className="flex items-center gap-2 px-4 py-2 bg-slate-800 text-slate-300 rounded-lg hover:bg-slate-700 font-bold border border-slate-700 uppercase text-[10px] tracking-wide transition-all shadow-sm"><BarChart3 size={16} /> Dashboard</button>
          <button onClick={loadData} className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg transition-all border border-slate-700 bg-slate-900"><RefreshCw size={18} /></button>
          <button onClick={() => { if(confirm(`Arquivar ${filteredRoutes.length} itens?`)) SharePointService.moveDeparturesToHistory(getAccessToken(), filteredRoutes as RouteDeparture[]).then(loadData); }} disabled={isSyncing || filteredRoutes.length === 0} className="flex items-center gap-2 px-4 py-2 bg-slate-900 text-slate-300 rounded-lg hover:bg-slate-800 font-bold border border-slate-700 uppercase text-[10px] shadow-sm disabled:opacity-30 transition-all"><Archive size={16} /> Arquivar Grade</button>
        </div>
      </div>

      <div className="flex-1 overflow-auto bg-white dark:bg-slate-900 rounded-2xl border border-slate-700/50 shadow-2xl relative scrollbar-thin">
        <div style={{ transform: `scale(${zoomLevel})`, transformOrigin: 'top left', width: `${100 / zoomLevel}%` }}>
            <table className="border-collapse table-fixed w-full min-w-max">
              <thead className="sticky top-0 z-50 bg-[#1e293b] text-white shadow-md">
                <tr className="h-12">
                  {tableColumns.map(col => (
                    <th key={col.id} style={{ width: colWidths[col.id] }} className="relative p-1 border border-slate-700/50 text-[10px] font-black uppercase tracking-wider text-left group">
                      <div className="flex items-center justify-between px-2 h-full"><span>{col.label}</span><button onClick={(e) => { e.stopPropagation(); setActiveFilterCol(activeFilterCol === col.id ? null : col.id); }} className={`p-1 rounded ${!!colFilters[col.id] || (selectedFilters[col.id]?.length ?? 0) > 0 ? 'text-yellow-400' : 'text-white/40'}`}><Filter size={11} /></button></div>
                      {activeFilterCol === col.id && <FilterDropdown col={col.id} routes={routes} colFilters={colFilters} setColFilters={setColFilters} selectedFilters={selectedFilters} setSelectedFilters={setSelectedFilters} onClose={() => setActiveFilterCol(null)} />}
                      <div onMouseDown={(e) => { e.preventDefault(); resizingRef.current = { col: col.id, startX: e.clientX, startWidth: colWidths[col.id] }; }} className="absolute right-0 top-0 bottom-0 w-1 cursor-col-resize z-10" />
                    </th>
                  ))}
                  <th style={{ width: 60 }} className="relative p-1 border border-slate-700/50 text-[10px] font-black uppercase text-center bg-slate-900/50">
                      {selectedIds.size > 0 ? (
                          <button onClick={handleDeleteSelected} className="p-1 text-red-500 hover:text-red-400 transition-colors" title="Deletar Selecionados"><Trash2 size={16} /></button>
                      ) : <Settings2 size={14} className="mx-auto opacity-40" />}
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-300 dark:divide-slate-700">
                {[...filteredRoutes, ghostRow].map((route, rowIndex) => {
                  const rowStyle = getRowStyle(route);
                  const isGhost = route.id === 'ghost';
                  const isSelected = selectedIds.has(route.id!);
                  const isDelayed = route.statusOp === 'Atrasada' || route.statusOp === 'Adiantada';
                  const isDelayedFilled = isDelayed && (route.saida !== '' && route.saida !== '00:00:00');
                  
                  const inputClass = `w-full h-full bg-transparent outline-none border-none px-3 py-2 text-[11px] font-semibold uppercase transition-all ${isDelayedFilled ? 'text-white placeholder-white/50' : 'text-slate-800 dark:text-slate-200 placeholder-slate-400'}`;

                  return (
                    <tr key={route.id} className={`${isSelected ? 'bg-primary-600/20' : rowStyle} group transition-all h-auto`}>
                      <td className="p-0 border border-slate-300 dark:border-slate-700">
                          {isGhost ? (
                              <textarea rows={1} value={route.rota} placeholder="Digite p/ criar..." onChange={(e) => updateCell(route.id!, 'rota', e.target.value)} onInput={(e) => { const el = e.target as HTMLTextAreaElement; el.style.height = 'auto'; el.style.height = (el.scrollHeight) + 'px'; }} className={`${inputClass} font-black resize-none overflow-hidden min-h-[38px]`} />
                          ) : (
                              <input type="text" value={route.rota} onChange={(e) => updateCell(route.id!, 'rota', e.target.value)} className={`${inputClass} font-black`} />
                          )}
                      </td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><input type="date" value={route.data} onChange={(e) => updateCell(route.id!, 'data', e.target.value)} className={`${inputClass} text-center`} /></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><input type="text" value={route.inicio} onPaste={(e) => { const val = e.clipboardData.getData('text'); if (val.includes('\n')) { e.preventDefault(); handleMultilinePaste('inicio', rowIndex, val); } }} onBlur={(e) => updateCell(route.id!, 'inicio', e.target.value)} className={`${inputClass} font-mono text-center`} /></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><input type="text" value={route.motorista} onPaste={(e) => { const val = e.clipboardData.getData('text'); if (val.includes('\n')) { e.preventDefault(); handleMultilinePaste('motorista', rowIndex, val); } }} onChange={(e) => updateCell(route.id!, 'motorista', e.target.value)} className={`${inputClass}`} /></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><input type="text" value={route.placa} onPaste={(e) => { const val = e.clipboardData.getData('text'); if (val.includes('\n')) { e.preventDefault(); handleMultilinePaste('placa', rowIndex, val); } }} onChange={(e) => updateCell(route.id!, 'placa', e.target.value)} className={`${inputClass} font-mono text-center`} /></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><input type="text" value={route.saida} placeholder="--:--:--" onPaste={(e) => { const val = e.clipboardData.getData('text'); if (val.includes('\n')) { e.preventDefault(); handleMultilinePaste('saida', rowIndex, val); } }} onBlur={(e) => updateCell(route.id!, 'saida', e.target.value)} className={`${inputClass} font-mono text-center`} /></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700">
                        {isDelayed && (
                          <select value={route.motivo} onChange={(e) => updateCell(route.id!, 'motivo', e.target.value)} className="w-full bg-white/20 dark:bg-slate-800/20 border-none px-2 py-1 text-[10px] font-bold text-inherit outline-none appearance-none cursor-pointer">
                              <option value="" className="text-slate-800">---</option>{MOTIVOS.map(m => (<option key={m} value={m} className="text-slate-800">{m}</option>))}
                          </select>
                        )}
                      </td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700 relative align-top">
                        {isDelayed && (
                          <div className="flex items-start w-full h-full relative p-0 min-h-[44px]">
                            <textarea value={route.observacao || ""} onPaste={(e) => { const val = e.clipboardData.getData('text'); if (val.includes('\n')) { e.preventDefault(); handleMultilinePaste('observacao', rowIndex, val); } }} onChange={(e) => updateCell(route.id!, 'observacao', e.target.value)} onFocus={() => setActiveObsId(route.id!)} placeholder="..." className={`w-full h-full min-h-[44px] bg-transparent outline-none border-none px-3 py-2 text-[11px] font-normal resize-none overflow-hidden ${isTextWrapEnabled ? 'whitespace-normal' : 'truncate pr-8'}`} onInput={(e) => { if (isTextWrapEnabled) { const el = e.target as HTMLTextAreaElement; el.style.height = 'auto'; el.style.height = el.scrollHeight + 'px'; } }} />
                            {!isTextWrapEnabled && <button onClick={(e) => { e.stopPropagation(); setActiveObsId(activeObsId === route.id ? null : route.id!); }} className="absolute right-2 top-1/2 -translate-y-1/2 p-0.5 opacity-60"><ChevronDown size={14} /></button>}
                            {activeObsId === route.id && (
                              <div ref={obsDropdownRef} className="absolute top-full left-0 w-full z-[110] bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-xl shadow-2xl overflow-hidden animate-in fade-in slide-in-from-top-1">
                                <div className="max-h-48 overflow-y-auto scrollbar-thin">{(route.motivo ? (OBSERVATION_TEMPLATES[route.motivo] || []) : []).map((template, tIdx) => ( <div key={tIdx} onClick={() => { updateCell(route.id!, 'observacao', template); setActiveObsId(null); }} className="p-3 text-[10px] text-slate-700 dark:text-slate-300 hover:bg-primary-100 dark:hover:bg-slate-700 cursor-pointer border-b border-slate-100 dark:border-slate-700 flex items-center gap-2"><ChevronRight size={12} className="shrink-0 text-primary-500" />{template}</div> ))}</div>
                              </div>
                            )}
                          </div>
                        )}
                      </td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><select value={route.statusGeral} onChange={(e) => updateCell(route.id!, 'statusGeral', e.target.value)} className="w-full h-full bg-transparent border-none text-[10px] font-bold text-center appearance-none"><option value="OK">OK</option><option value="NOK">NOK</option></select></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700"><select value={route.operacao} onChange={(e) => updateCell(route.id!, 'operacao', e.target.value)} className="w-full h-full bg-transparent border-none text-[9px] font-black text-center uppercase"><option value="">OP...</option>{userConfigs.map(c => <option key={c.operacao} value={c.operacao}>{c.operacao}</option>)}</select></td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700 text-center">
                        <span className={`px-2 py-0.5 rounded-full text-[8px] font-black border ${route.statusOp === 'OK' ? 'bg-emerald-100 border-emerald-400 text-emerald-800' : route.statusOp === 'Atrasada' ? 'bg-yellow-100 border-yellow-400 text-yellow-800' : route.statusOp === 'Programada' ? 'bg-slate-200 border-slate-400 text-slate-600' : 'bg-red-100 border-red-400 text-red-800'}`}>{route.statusOp}</span>
                      </td>
                      <td className="p-0 border border-slate-300 dark:border-slate-700 text-center font-mono font-bold text-[10px]">{route.tempo}</td>
                      
                      <td className="p-0 border border-slate-300 dark:border-slate-700 flex items-center justify-center gap-1 h-12">
                          {!isGhost && (
                              <>
                                <button onClick={() => toggleSelection(route.id!)} className={`p-1.5 rounded-lg transition-colors ${isSelected ? 'text-primary-500 bg-primary-500/10' : 'text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800'}`}>
                                    {isSelected ? <CheckSquare size={16}/> : <Square size={16}/>}
                                </button>
                                <button onClick={() => handleDeleteRoute(route.id!)} className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-red-500/10 rounded-lg transition-colors">
                                    <Trash2 size={16} />
                                </button>
                              </>
                          )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
        </div>
      </div>

      {isBulkMappingModalOpen && (
          <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4">
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-primary-500 shadow-2xl animate-in zoom-in">
                  <div className="flex items-center gap-3 text-primary-500 mb-6 font-black uppercase text-xs"><Layers size={24} /> Atribuir Planta para Lote</div>
                  <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">Você colou <span className="text-primary-500 font-black">{pendingBulkRoutes.length} rotas</span>. Escolha a operação:</p>
                  <div className="grid grid-cols-2 gap-3 max-h-64 overflow-y-auto pr-2 scrollbar-thin">
                      {userConfigs.map(c => ( <button key={c.operacao} onClick={() => handleBulkCreateSave(c.operacao)} className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-2xl hover:bg-primary-600 hover:text-white transition-all font-black text-xs uppercase">{c.operacao}</button> ))}
                  </div>
                  <button onClick={() => setIsBulkMappingModalOpen(false)} className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400">Cancelar</button>
              </div>
          </div>
      )}

      {isMappingModalOpen && (
          <div className="fixed inset-0 bg-slate-950/90 backdrop-blur-md z-[300] flex items-center justify-center p-4">
              <div className="bg-white dark:bg-slate-900 rounded-[2.5rem] p-8 w-full max-w-md border border-primary-500 animate-in zoom-in">
                  <div className="flex items-center gap-3 text-primary-500 mb-6 font-black uppercase text-xs"><LinkIcon size={24} /> Vínculo Necessário</div>
                  <p className="text-sm text-slate-500 dark:text-slate-400 mb-6">A rota <span className="text-primary-500 font-black">{pendingMappingRoute}</span> não possui planta vinculada:</p>
                  <div className="grid grid-cols-2 gap-3">
                      {userConfigs.map(c => ( <button key={c.operacao} onClick={() => { SharePointService.addRouteOperationMapping(getAccessToken(), pendingMappingRoute!, c.operacao); setGhostRow(prev => ({...prev, operacao: c.operacao})); setIsMappingModalOpen(false); }} className="p-4 bg-slate-50 dark:bg-slate-800 border border-slate-200 rounded-2xl hover:bg-primary-600 hover:text-white transition-all font-black text-xs uppercase">{c.operacao}</button> ))}
                  </div>
                  <button onClick={() => { setIsMappingModalOpen(false); setGhostRow(prev => ({...prev, rota: ''})); }} className="w-full mt-6 py-4 text-[10px] font-black uppercase text-slate-400">Cancelar</button>
              </div>
          </div>
      )}

      {isHistoryModalOpen && (
          <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[200] flex items-center justify-center p-4">
              <div className="bg-white dark:bg-slate-900 border dark:border-slate-800 rounded-[2.5rem] shadow-2xl w-full max-w-6xl max-h-[90vh] overflow-hidden flex flex-col">
                  <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white">
                      <div className="flex items-center gap-4"><Database size={24} /><h3 className="font-black uppercase tracking-widest text-base">Histórico Definitivo</h3></div>
                      <button onClick={() => setIsHistoryModalOpen(false)}><X size={28} /></button>
                  </div>
                  <div className="p-6 bg-slate-50 dark:bg-slate-900 border-b dark:border-slate-800 grid grid-cols-3 gap-4">
                      <input type="date" value={histStart} onChange={e => setHistStart(e.target.value)} className="p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white" />
                      <input type="date" value={histEnd} onChange={e => setHistEnd(e.target.value)} className="p-3 border dark:border-slate-700 rounded-xl bg-white dark:bg-slate-800 text-[11px] font-bold outline-none dark:text-white" />
                      <button onClick={handleSearchArchive} disabled={isSearchingArchive} className="py-3 bg-primary-600 text-white font-black uppercase text-[11px] rounded-xl flex items-center justify-center gap-2 hover:bg-primary-700 shadow-lg"> {isSearchingArchive ? <Loader2 size={16} className="animate-spin" /> : <Search size={16} />} BUSCAR </button>
                  </div>
                  <div className="flex-1 overflow-auto p-4 bg-slate-50 dark:bg-slate-950">
                      {archivedResults.length > 0 ? (
                        <table className="w-full border-collapse text-[10px]">
                            <thead className="sticky top-0 bg-slate-200 dark:bg-slate-800 text-slate-600 font-black uppercase">
                                <tr><th className="p-2 border border-slate-300 dark:border-slate-700 text-left">Rota</th><th className="p-2 border border-slate-300 text-center">Data</th><th className="p-2 border border-slate-300 text-center">Saída</th><th className="p-2 border border-slate-300 text-left">Motivo</th><th className="p-2 border border-slate-300 text-center">OP</th></tr>
                            </thead>
                            <tbody>
                                {archivedResults.map((r, i) => (<tr key={i} className="hover:bg-slate-50 dark:hover:bg-slate-800 border-b border-slate-200 dark:border-slate-800"><td className="p-2 font-bold text-primary-700">{r.rota}</td><td className="p-2 text-center">{r.data}</td><td className="p-2 text-center font-mono">{r.saida}</td><td className="p-2">{r.motivo || "---"}</td><td className="p-2 text-center font-black">{r.operacao}</td></tr>))}
                            </tbody>
                        </table>
                      ) : <div className="h-full flex flex-col items-center justify-center text-slate-400 italic font-bold">Nenhum dado retornado para este período</div>}
                  </div>
              </div>
          </div>
      )}

      {isStatsModalOpen && dashboardStats && (
        <div className="fixed inset-0 bg-slate-950/70 backdrop-blur-md z-[200] flex items-center justify-center p-4">
            <div className="bg-white dark:bg-slate-900 rounded-[2rem] shadow-2xl w-full max-w-4xl overflow-hidden border dark:border-slate-800 animate-in zoom-in">
                <div className="bg-[#1e293b] p-6 flex justify-between items-center text-white"><div className="flex items-center gap-4"><TrendingUp size={24} /><h3 className="font-black uppercase tracking-widest text-base">Dashboard Operacional</h3></div><button onClick={() => setIsStatsModalOpen(false)}><X size={28} /></button></div>
                <div className="p-8 grid grid-cols-3 gap-6 bg-slate-50 dark:bg-slate-950">
                    {[{ label: 'Total', value: dashboardStats.total, icon: Activity, color: 'text-slate-700 bg-white' }, { label: 'OK', value: `${Math.round((dashboardStats.okCount / dashboardStats.total) * 100)}%`, icon: CheckCircle2, color: 'text-emerald-600 bg-emerald-50' }, { label: 'Atrasos', value: `${Math.round((dashboardStats.delayedCount / dashboardStats.total) * 100)}%`, icon: AlertTriangle, color: 'text-orange-600 bg-orange-50' }].map((stat, idx) => ( <div key={idx} className={`p-6 rounded-2xl border dark:border-slate-800 flex flex-col gap-2 ${stat.color}`}><stat.icon size={20} /><span className="text-[10px] font-black uppercase text-slate-400 mt-2">{stat.label}</span><div className="text-3xl font-black">{stat.value}</div></div> ))}
                </div>
            </div>
        </div>
      )}
    </div>
  );
};

export default RouteDepartureView;
