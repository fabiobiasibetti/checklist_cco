
import React, { useState, useEffect } from 'react';
import { HashRouter as Router, Routes, Route, useNavigate } from 'react-router-dom';
import { CheckSquare, History, Truck, Moon, Sun, LogOut, ChevronLeft, ChevronRight, Loader2 } from 'lucide-react';
import TaskManager from './components/TaskManager';
import HistoryViewer from './components/HistoryViewer';
import RouteDepartureView from './components/RouteDeparture';
import Login from './components/Login';
import { SharePointService, getLocalDateString } from './services/sharepointService';
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
  const [isDarkMode, setIsDarkMode] = useState(true);
  const [collapsed, setCollapsed] = useState(true);
  const [collapsedCategories, setCollapsedCategories] = useState<string[]>([]);
  
  const navigate = useNavigate();

  const loadDataFromSharePoint = async (user: User) => {
    if (!user.accessToken) return;
    (window as any).__access_token = user.accessToken; 
    setIsLoading(true);
    try {
      // 1. Get fundamental metadata
      const [spTasks, spOps] = await Promise.all([
        SharePointService.getTasks(user.accessToken),
        SharePointService.getOperations(user.accessToken, user.email)
      ]);

      const opSiglas = spOps.map(o => o.Title);
      setLocations(opSiglas);

      // 2. Get all persistent status
      const spStatus = await SharePointService.getAllStatus(user.accessToken);
      const statusMap = new Map<string, SPStatus>();
      spStatus.forEach(s => statusMap.set(s.Title, s));

      // 3. Auto-Provisioning: Check if all Task-Op combinations exist
      const provisionPromises: Promise<any>[] = [];
      spTasks.forEach(task => {
        opSiglas.forEach(sigla => {
            const key = `${task.id}_${sigla}`;
            if (!statusMap.has(key)) {
                console.log(`Célula faltante detectada: ${key}. Criando no SharePoint...`);
                provisionPromises.push(
                    SharePointService.ensureCellExists(user.accessToken!, String(task.id), sigla)
                        .then(newStatus => statusMap.set(newStatus.Title, newStatus))
                );
            }
        });
      });

      if (provisionPromises.length > 0) {
          await Promise.all(provisionPromises);
      }

      // 4. Update TaskManager internal ID ref
      if ((window as any).refreshSpIds) {
          (window as any).refreshSpIds(Array.from(statusMap.values()));
      }

      // 5. Build UI state
      const matrixTasks: Task[] = spTasks.map(t => {
        const ops: Record<string, any> = {};
        opSiglas.forEach(sigla => {
          const key = `${t.id}_${sigla}`;
          const statusMatch = statusMap.get(key);
          ops[sigla] = statusMatch ? statusMatch.Status : 'PR';
        });

        return {
          id: String(t.id),
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
    } catch (err) {
      console.error("Erro ao carregar SharePoint:", err);
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
             <p className="font-bold animate-pulse">Sincronizando Estado Atual...</p>
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
