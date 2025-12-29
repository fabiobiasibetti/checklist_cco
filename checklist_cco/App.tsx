
import React, { useState, useEffect, useRef } from 'react';
import { HashRouter as Router, Routes, Route, useNavigate } from 'react-router-dom';
import { CheckSquare, History, Truck, Moon, Sun, LogOut, ChevronLeft, ChevronRight, Loader2, RefreshCw, PauseCircle } from 'lucide-react';
import TaskManager from './components/TaskManager';
import HistoryViewer from './components/HistoryViewer';
import RouteDepartureView from './components/RouteDeparture';
import Login from './components/Login';
import { SharePointService } from './services/sharepointService';
import { Task, User } from './types';
import { setCurrentUser as setStorageUser } from './services/storageService';

const SidebarLink = ({ to, icon: Icon, label, active, collapsed }: any) => (
  <a href={`#${to}`} className={`flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${active ? 'bg-blue-600 text-white' : 'text-slate-500 hover:bg-slate-100'} ${collapsed ? 'justify-center' : ''}`}>
    <Icon size={20} />
    {!collapsed && <span className="font-medium whitespace-nowrap">{label}</span>}
  </a>
);

const AppContent = () => {
  const [currentUser, setUser] = useState<User | null>(null);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [locations, setLocations] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSyncing, setIsSyncing] = useState(false);
  const [isSyncPaused, setIsSyncPaused] = useState(false);
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [collapsed, setCollapsed] = useState(true);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);
  
  const partialSaveAttemptedRef = useRef(false);
  const lastListTimestampRef = useRef<string | null>(null);
  const isSyncBlockedRef = useRef(false);
  const cooldownTimeoutRef = useRef<number | null>(null);
  
  const navigate = useNavigate();

  const loadDataFromSharePoint = async (user: User) => {
    if (!user.accessToken) return;
    (window as any).__access_token = user.accessToken; 
    setIsLoading(true);
    try {
      const spTasks = await SharePointService.getTasks(user.accessToken);
      const spOps = await SharePointService.getOperations(user.accessToken, user.email);
      const today = new Date().toISOString().split('T')[0];
      const spStatus = await SharePointService.getStatusByDate(user.accessToken, today);

      const meta = await SharePointService.getListMetadata(user.accessToken, 'Status_Checklist');
      lastListTimestampRef.current = meta.lastModifiedDateTime;

      const opSiglas = spOps.map(o => o.Title);
      setLocations(opSiglas);

      const matrixTasks: Task[] = spTasks.map(t => {
        const ops: Record<string, any> = {};
        opSiglas.forEach(sigla => {
          const statusMatch = spStatus.find(s => s.TarefaID === t.id && s.OperacaoSigla === sigla);
          ops[sigla] = statusMatch ? statusMatch.Status : 'PR';
        });

        return {
          id: t.id,
          title: t.Title,
          description: t.Descricao,
          category: t.Categoria,
          timeRange: t.Horario,
          operations: ops,
          createdAt: new Date().toISOString(),
          isDaily: true,
          active: t.Ativa
        };
      });

      setTasks(matrixTasks.filter(t => t.active !== false));

      if (!partialSaveAttemptedRef.current) {
        checkAndTriggerPartialSave(user, matrixTasks.filter(t => t.active !== false));
        partialSaveAttemptedRef.current = true;
      }
    } catch (err) {
      console.error("Erro ao carregar SharePoint:", err);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    if (!currentUser?.accessToken || isLoading) return;

    const syncInterval = setInterval(async () => {
      if (isSyncing || isSyncBlockedRef.current) return;
      try {
        const meta = await SharePointService.getListMetadata(currentUser.accessToken!, 'Status_Checklist');
        if (meta.lastModifiedDateTime !== lastListTimestampRef.current) {
          setIsSyncing(true);
          const today = new Date().toISOString().split('T')[0];
          const newSpStatus = await SharePointService.getStatusByDate(currentUser.accessToken!, today);
          setTasks(prevTasks => prevTasks.map(task => {
            const updatedOps = { ...task.operations };
            let hasChanged = false;
            locations.forEach(sigla => {
              const statusMatch = newSpStatus.find(s => s.TarefaID === task.id && s.OperacaoSigla === sigla);
              const newStatus = statusMatch ? statusMatch.Status : 'PR';
              if (updatedOps[sigla] !== newStatus) {
                updatedOps[sigla] = newStatus;
                hasChanged = true;
              }
            });
            return hasChanged ? { ...task, operations: updatedOps } : task;
          }));
          lastListTimestampRef.current = meta.lastModifiedDateTime;
          setIsSyncing(false);
        }
      } catch (e) {
        setIsSyncing(false);
      }
    }, 6000);
    return () => clearInterval(syncInterval);
  }, [currentUser, isLoading, isSyncing, locations]);

  const handleManualSaveComplete = async () => {
    if (!currentUser?.accessToken) return;
    isSyncBlockedRef.current = true;
    setIsSyncPaused(true);
    if (cooldownTimeoutRef.current) window.clearTimeout(cooldownTimeoutRef.current);
    cooldownTimeoutRef.current = window.setTimeout(() => {
        isSyncBlockedRef.current = false;
        setIsSyncPaused(false);
        cooldownTimeoutRef.current = null;
    }, 8000); 
  };

  const checkAndTriggerPartialSave = async (user: User, currentTasks: Task[]) => {
    const now = new Date();
    const hours = now.getHours();
    if (hours >= 10 && hours < 22) {
        try {
            const history = await SharePointService.getHistory(user.accessToken!, user.email);
            const todayStr = now.toISOString().split('T')[0];
            const alreadyHasPartial = history.some(h => h.isPartial && h.timestamp.startsWith(todayStr));
            if (!alreadyHasPartial && currentTasks.length > 0) {
                await SharePointService.saveHistory(user.accessToken!, {
                    id: `partial_${Date.now()}`, timestamp: now.toISOString(), tasks: currentTasks, resetBy: user.name, email: user.email, isPartial: true
                });
            }
        } catch (e) {}
    }
  };

  const handleLogout = () => {
    // Limpa o cache local
    setUser(null);
    setStorageUser(null);
    delete (window as any).__access_token;
    
    // Limpa o cache do MSAL no navegador
    localStorage.clear(); 
    sessionStorage.clear();
    
    navigate('/');
  };

  useEffect(() => {
    if (isDarkMode) document.documentElement.classList.add('dark');
    else document.documentElement.classList.remove('dark');
  }, [isDarkMode]);

  if (!currentUser) return <Login onLogin={(u) => { setUser(u); loadDataFromSharePoint(u); }} />;

  return (
    <div className="flex h-screen bg-slate-50 dark:bg-slate-950 overflow-hidden">
      <aside className={`bg-white dark:bg-slate-900 border-r dark:border-slate-800 transition-all ${collapsed ? 'w-20' : 'w-64'} p-4 flex flex-col`}>
        <div className="mb-10 flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center text-white font-bold">V</div>
          {!collapsed && <h1 className="font-bold dark:text-white text-sm">CCO Digital</h1>}
        </div>
        <nav className="flex-1 space-y-2">
          <SidebarLink to="/" icon={CheckSquare} label="Checklist" active={window.location.hash === '#/'} collapsed={collapsed} />
          <SidebarLink to="/departures" icon={Truck} label="Saídas" active={window.location.hash === '#/departures'} collapsed={collapsed} />
          <SidebarLink to="/history" icon={History} label="Histórico" active={window.location.hash === '#/history'} collapsed={collapsed} />
        </nav>
        
        <div className={`mt-auto mb-4 p-2 rounded-lg flex items-center justify-center gap-2 transition-colors ${isSyncPaused ? 'text-amber-500 bg-amber-50 dark:bg-amber-900/20' : isSyncing ? 'text-blue-500 bg-blue-50 dark:bg-blue-900/20' : 'text-slate-400'}`}>
            {isSyncPaused ? <PauseCircle size={14} /> : <RefreshCw size={14} className={isSyncing ? 'animate-spin' : ''} />}
            {!collapsed && <span className="text-[10px] font-bold uppercase tracking-tighter">{isSyncPaused ? 'Pausa Segura' : isSyncing ? 'Sincronizando' : 'Em Nuvem'}</span>}
        </div>

        <div className="space-y-2 border-t pt-4 dark:border-slate-800">
           <button onClick={() => setIsDarkMode(!isDarkMode)} className="p-2 w-full flex justify-center text-slate-500 hover:bg-slate-100 rounded-lg">
             {isDarkMode ? <Sun size={20}/> : <Moon size={20}/>}
           </button>
           <button onClick={() => setCollapsed(!collapsed)} className="p-2 w-full flex justify-center text-slate-500 hover:bg-slate-100 rounded-lg">
             {collapsed ? <ChevronRight size={20}/> : <ChevronLeft size={20}/>}
           </button>
        </div>
      </aside>
      <main className="flex-1 overflow-hidden p-4">
        {isLoading ? (
          <div className="h-full flex items-center justify-center flex-col gap-4 text-blue-600">
             <Loader2 size={40} className="animate-spin" />
             <p className="font-bold animate-pulse">Sincronizando com SharePoint...</p>
          </div>
        ) : (
          <Routes>
            <Route path="/" element={
              <TaskManager 
                tasks={tasks} 
                setTasks={setTasks} 
                locations={locations} 
                setLocations={setLocations} 
                onUserSwitch={() => loadDataFromSharePoint(currentUser)} 
                collapsedCategories={collapsedCategories} 
                setCollapsedCategories={setCollapsedCategories} 
                currentUser={currentUser}
                onLogout={handleLogout}
                onInteractionStart={() => { isSyncBlockedRef.current = true; setIsSyncPaused(true); }}
                onInteractionEnd={handleManualSaveComplete}
              />
            } />
            <Route path="/departures" element={<RouteDepartureView />} />
            <Route path="/history" element={<HistoryViewer currentUser={currentUser} />} />
          </Routes>
        )}
      </main>
    </div>
  );
};

const App = () => (<Router><AppContent /></Router>);
export default App;
