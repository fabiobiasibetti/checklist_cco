
import React, { useState, useEffect } from 'react';
import { HashRouter as Router, Routes, Route, useNavigate } from 'react-router-dom';
import { CheckSquare, History, Truck, Moon, Sun, LogOut, ChevronLeft, ChevronRight, Loader2, AlertTriangle, RefreshCw } from 'lucide-react';
import TaskManager from './components/TaskManager';
import HistoryViewer from './components/HistoryViewer';
import RouteDepartureView from './components/RouteDeparture';
import Login from './components/Login';
import { SharePointService } from './services/sharepointService';
import { Task, User, SPTask, SPOperation, SPStatus } from './types';
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
  const [syncError, setSyncError] = useState<string | null>(null);
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [collapsed, setCollapsed] = useState(true);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);
  
  const navigate = useNavigate();

  const loadDataFromSharePoint = async (user: User) => {
    if (!user.accessToken) return;
    (window as any).__access_token = user.accessToken; 
    setIsLoading(true);
    setSyncError(null);
    
    try {
      // 1. Validação inicial: Tenta encontrar as listas. Se falhar aqui, já para tudo.
      await SharePointService.validateConnection(user.accessToken);

      // 2. Busca os dados
      const [spTasks, spOps] = await Promise.all([
        SharePointService.getTasks(user.accessToken),
        SharePointService.getOperations(user.accessToken, user.email)
      ]);

      if (spOps.length === 0) {
          throw new Error(`Seu e-mail (${user.email}) não está vinculado a nenhuma operação na lista 'Operacoes_Checklist'.`);
      }

      const today = new Date().toISOString().split('T')[0];
      const spStatus = await SharePointService.getStatusByDate(user.accessToken, today);

      const opSiglas = spOps.map(o => o.Title);
      setLocations(opSiglas);

      const matrixTasks: Task[] = spTasks.map(t => {
        const ops: Record<string, any> = {};
        opSiglas.forEach(sigla => {
          const matchedStatuses = spStatus.filter(s => s.TarefaID === t.id && s.OperacaoSigla === sigla);
          const statusMatch = matchedStatuses.length > 0 ? matchedStatuses[matchedStatuses.length - 1] : null;
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
    } catch (err: any) {
      console.error("Erro crítico de sincronização:", err);
      setSyncError(err.message || "Erro desconhecido ao conectar com SharePoint.");
    } finally {
      setIsLoading(false);
    }
  };

  const handleLogout = () => {
    setUser(null);
    setStorageUser(null);
    delete (window as any).__access_token;
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
        <div className="mt-auto space-y-2 border-t pt-4 dark:border-slate-800">
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
        ) : syncError ? (
          <div className="h-full flex items-center justify-center p-8">
            <div className="bg-white dark:bg-slate-900 p-10 rounded-[2.5rem] shadow-xl border dark:border-slate-800 max-w-xl w-full text-center flex flex-col items-center">
                <div className="w-20 h-20 bg-red-100 dark:bg-red-900/30 text-red-600 dark:text-red-400 rounded-3xl flex items-center justify-center mb-6">
                    <AlertTriangle size={40} />
                </div>
                <h2 className="text-2xl font-black text-slate-800 dark:text-white mb-4 uppercase tracking-tight">Falha na Sincronização</h2>
                <div className="bg-red-50 dark:bg-red-900/10 border border-red-100 dark:border-red-900/50 p-4 rounded-2xl text-red-600 dark:text-red-400 text-sm font-medium mb-8 text-left w-full">
                    {syncError}
                </div>
                <div className="flex flex-col gap-3 w-full">
                    <button 
                        onClick={() => loadDataFromSharePoint(currentUser)}
                        className="w-full py-4 bg-blue-600 hover:bg-blue-700 text-white rounded-2xl font-bold flex items-center justify-center gap-2 transition-all active:scale-95"
                    >
                        <RefreshCw size={20} /> Tentar Novamente
                    </button>
                    <button 
                        onClick={handleLogout}
                        className="w-full py-4 bg-slate-200 dark:bg-slate-800 text-slate-600 dark:text-slate-300 rounded-2xl font-bold hover:bg-slate-300 dark:hover:bg-slate-700 transition-all"
                    >
                        Sair da Conta
                    </button>
                </div>
                <p className="mt-8 text-[10px] text-slate-400 font-bold uppercase tracking-widest leading-relaxed">
                    Certifique-se de que a lista <b>Status_Checklist</b> existe no site CCO e que você possui permissões de edição.
                </p>
            </div>
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
