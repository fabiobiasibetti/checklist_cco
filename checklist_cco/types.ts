
export type OperationStatus = 'PR' | 'OK' | 'EA' | 'AR' | 'ATT' | 'AT';

export interface OperationStates {
  [location: string]: OperationStatus;
}

// Added missing TaskPriority enum
export enum TaskPriority {
  LOW = 'Baixa',
  MEDIUM = 'Média',
  HIGH = 'Alta'
}

// Added missing TaskStatus enum
export enum TaskStatus {
  TODO = 'TODO',
  DONE = 'DONE'
}

export interface Task {
  id: string;
  timeRange: string;
  title: string;
  description: string;
  category: string;
  operations: OperationStates;
  createdAt: string;
  isDaily: boolean;
  active?: boolean;
  // Added optional fields to support objects returned by Gemini
  priority?: TaskPriority | string;
  status?: TaskStatus;
  dueDate?: string;
}

// Added missing Customer interface
export interface Customer {
  id: string;
  name: string;
  company: string;
  email: string;
  phone: string;
  status: 'Lead' | 'Ativo' | 'Inativo';
}

export interface SPListInfo {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  columns?: SPColumnInfo[];
}

export interface SPColumnInfo {
  name: string;
  displayName: string;
  description: string;
  type: string;
}

export interface User {
  email: string;
  name: string;
  accessToken?: string;
}

export interface HistoryRecord {
  id: string;
  timestamp: string;
  tasks: Task[];
  resetBy?: string;
  email?: string;
}

export interface RouteDeparture {
  id: string;
  semana: string;
  rota: string;
  data: string;
  inicio: string;
  motorista: string;
  placa: string;
  saida: string;
  motivo: string;
  observacao: string;
  statusGeral: string;
  aviso: string;
  operacao: string;
  statusOp: string;
  tempo: string;
  createdAt: string;
}

export interface SPTask {
  id: string;
  Title: string;
  Descricao: string;
  Categoria: string;
  Horario: string;
  Ativa: boolean;
  Ordem: number;
}

export interface SPOperation {
  id: string;
  Title: string;
  Ordem: number;
  Email: string;
}

export interface SPStatus {
  id?: string;
  DataReferencia: string;
  TarefaID: string;
  OperacaoSigla: string;
  Status: OperationStatus;
  Usuario: string;
  Title: string;
}

export const VALID_USERS = [];
export const INITIAL_LOCATIONS: string[] = [];
export const LOGISTICA_2_LOCATIONS: string[] = ['LAT-CWB', 'LAT-SJP', 'LAT-LDB', 'LAT-MGA'];
