
import { GoogleGenAI, Type } from "@google/genai";
import { Task, TaskPriority, TaskStatus, RouteDeparture } from "../types";

export const parseExcelContentToTasks = async (rawText: string): Promise<Partial<Task>[]> => {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  const prompt = `Analise o texto bruto de uma planilha e extraia as tarefas em JSON. Texto: """${rawText}"""`;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              title: { type: Type.STRING },
              description: { type: Type.STRING },
              priority: { type: Type.STRING, enum: ["Baixa", "Média", "Alta"] },
              category: { type: Type.STRING },
              dueDate: { type: Type.STRING },
            },
            required: ["title", "priority", "category"]
          }
        }
      }
    });

    if (response.text) {
      const parsed = JSON.parse(response.text.trim());
      return parsed.map((item: any) => ({
        ...item,
        status: TaskStatus.TODO,
        createdAt: new Date().toISOString()
      }));
    }
    return [];
  } catch (error) {
    console.error("Error parsing tasks:", error);
    throw error;
  }
};

/**
 * Manual parser for Excel data (Tab-separated)
 * Expected Order: ROTA | DATA | INICIO | MOTORISTA | PLACA | SAIDA | MOTIVO | OBSERVAÇÃO | (OPERAÇÃO)
 */
export const parseRouteDeparturesManual = (rawText: string): Partial<RouteDeparture>[] => {
  const rows = rawText.trim().split(/\r?\n/);
  
  const convertDate = (dateStr: string) => {
    if (!dateStr) return '';
    const parts = dateStr.includes('/') ? dateStr.split('/') : dateStr.split('-');
    if (parts.length === 3) {
      // If DD/MM/YYYY -> YYYY-MM-DD
      if (parts[0].length === 2 && parts[2].length === 4) {
        return `${parts[2]}-${parts[1]}-${parts[0]}`;
      }
      return dateStr;
    }
    return dateStr;
  };

  const formatTime = (timeStr: string) => {
    if (!timeStr || timeStr === '00:00:00' || timeStr.trim() === '') return '00:00:00';
    // Se for apenas HH:mm, adiciona :00
    if (timeStr.length === 5 && timeStr.includes(':')) return `${timeStr}:00`;
    return timeStr;
  };

  return rows.map(row => {
    // Splits by Tab (Standard Excel) or double spaces (fallback)
    const cols = row.split(/\t| {2,}/).map(c => c.trim());
    
    return {
      rota: cols[0] || '',
      data: convertDate(cols[1] || ''),
      inicio: formatTime(cols[2] || '00:00:00'),
      motorista: cols[3] || '',
      placa: cols[4] || '',
      saida: formatTime(cols[5] || '00:00:00'),
      motivo: cols[6] || '',
      observacao: cols[7] || '',
      operacao: cols[8] || '' 
    };
  }).filter(r => r.rota && r.data);
};

export const parseRouteDepartures = async (rawText: string): Promise<Partial<RouteDeparture>[]> => {
  // Check if API KEY is available before calling
  if (!process.env.API_KEY) {
      throw new Error("API Key não configurada no ambiente. Por favor, utilize a Importação Direta.");
  }

  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  
  const prompt = `
    Atue como um especialista em logística. Extraia os dados do texto abaixo (copiado de um Excel) para um JSON estruturado.
    MODELO DE COLUNAS: ROTA, DATA, INICIO, MOTORISTA, PLACA, SAIDA, MOTIVO, OBSERVAÇÃO, OPERAÇÃO.
    
    REGRAS:
    - DATA: Converta para YYYY-MM-DD.
    - HORÁRIOS (INICIO/SAIDA): Garanta formato HH:mm:ss. Se vazio use "00:00:00".
    - OPERAÇÃO: Identifique a sigla (ex: LAT-UNA, LAT-TER).
    
    TEXTO:
    """
    ${rawText}
    """
  `;

  try {
    const response = await ai.models.generateContent({
      model: "gemini-3-flash-preview",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              rota: { type: Type.STRING },
              data: { type: Type.STRING },
              inicio: { type: Type.STRING },
              motorista: { type: Type.STRING },
              placa: { type: Type.STRING },
              saida: { type: Type.STRING },
              motivo: { type: Type.STRING },
              observacao: { type: Type.STRING },
              operacao: { type: Type.STRING }
            },
            required: ["rota", "data", "operacao"]
          }
        }
      }
    });

    if (response.text) {
      return JSON.parse(response.text.trim());
    }
    return [];
  } catch (error) {
    console.error("Error parsing departures with Gemini:", error);
    throw error;
  }
};
