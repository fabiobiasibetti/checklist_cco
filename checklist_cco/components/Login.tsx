
import React, { useState, useEffect } from 'react';
import { User } from '../types';
import { Loader2, AlertCircle, ShieldCheck } from 'lucide-react';
import { setCurrentUser } from '../services/storageService';
import { PublicClientApplication, Configuration } from "@azure/msal-browser";

const msalConfig: Configuration = {
    auth: {
        clientId: "c176306d-f849-4cf4-bfca-22ff214cdaad",
        authority: "https://login.microsoftonline.com/7d9754b3-dcdb-4efe-8bb7-c0e5587b86ed",
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: "localStorage", // Salva o login no cache do navegador
        storeAuthStateInCookie: true,  // Ajuda na persistência em alguns navegadores
    }
};

const msalInstance = new PublicClientApplication(msalConfig);

const MicrosoftIcon = () => (
    <svg width="20" height="20" viewBox="0 0 23 23" xmlns="http://www.w3.org/2000/svg">
        <path fill="#f35325" d="M1 1h10v10H1z"/><path fill="#81bc06" d="M12 1h10v10H12z"/><path fill="#05a6f0" d="M1 12h10v10H1z"/><path fill="#ffba08" d="M12 12h10v10H12z"/>
    </svg>
);

const Login: React.FC<{ onLogin: (user: User) => void }> = ({ onLogin }) => {
  const [isCheckingSession, setIsCheckingSession] = useState(true);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const checkSession = async () => {
        try {
            await msalInstance.initialize();
            
            // Verifica se o usuário acabou de voltar de um redirecionamento de login
            const result = await msalInstance.handleRedirectPromise();
            if (result) {
                completeLogin(result);
                return;
            }

            // Busca contas já existentes no cache do navegador
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                try {
                    // Tenta obter um novo token sem pedir senha (silenciosamente)
                    const silentRequest = {
                        scopes: ["User.Read", "Sites.ReadWrite.All"],
                        account: accounts[0]
                    };
                    const response = await msalInstance.acquireTokenSilent(silentRequest);
                    completeLogin(response);
                } catch (e) {
                    console.log("Sessão expirada, necessário login manual.");
                    setIsCheckingSession(false);
                }
            } else {
                setIsCheckingSession(false);
            }
        } catch (err) {
            console.error("Erro MSAL:", err);
            setIsCheckingSession(false);
        }
    };
    checkSession();
  }, [onLogin]);

  const completeLogin = (response: any) => {
    setCurrentUser(response.account.username);
    onLogin({
        email: response.account.username,
        name: response.account.name || response.account.username,
        accessToken: response.accessToken
    });
  };

  const handleMicrosoftLogin = async () => {
    setIsLoggingIn(true);
    setError(null);
    try {
        const loginRequest = {
            scopes: ["User.Read", "Sites.ReadWrite.All"],
            prompt: "select_account" // Força escolher a conta na primeira vez ou após logout
        };
        const response = await msalInstance.loginPopup(loginRequest);
        if (response && response.account) {
            completeLogin(response);
        }
    } catch (err: any) {
        console.error(err);
        setError("Não foi possível completar o login corporativo.");
    } finally {
        setIsLoggingIn(false);
    }
  };

  if (isCheckingSession) {
    return (
        <div className="min-h-screen bg-slate-100 flex items-center justify-center">
            <div className="flex flex-col items-center gap-4">
                <Loader2 className="animate-spin text-blue-600" size={40} />
                <p className="text-slate-500 font-bold animate-pulse">Restaurando Sessão...</p>
            </div>
        </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 flex items-center justify-center p-6">
      <div className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-[440px] border overflow-hidden">
        <div className="h-2 w-full bg-blue-600"></div>
        <div className="p-10 flex flex-col items-center">
            <div className="mb-8"><img src="https://viagroup.com.br/assets/via_group-22fac685.png" alt="VIA Group" className="max-w-[180px]"/></div>
            <h1 className="text-2xl font-black text-slate-800 mb-2">Checklist CCO</h1>
            <p className="text-slate-500 text-sm mb-8">Gestão de Operações em Tempo Real</p>
            
            {error && <div className="w-full mb-4 p-3 bg-red-50 text-red-600 text-xs rounded-xl flex items-center gap-2"><AlertCircle size={14}/>{error}</div>}
            
            <button 
                onClick={handleMicrosoftLogin}
                disabled={isLoggingIn}
                className="w-full bg-slate-900 text-white py-4 rounded-2xl font-bold flex items-center justify-center gap-3 transition-all hover:bg-slate-800 disabled:opacity-70"
            >
                {isLoggingIn ? <Loader2 className="animate-spin" size={20} /> : <MicrosoftIcon />}
                <span>Entrar com Microsoft</span>
            </button>
            
            <div className="mt-8 text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
                <ShieldCheck size={12} className="text-blue-500" /> Acesso Seguro SharePoint
            </div>
        </div>
      </div>
    </div>
  );
};

export default Login;
