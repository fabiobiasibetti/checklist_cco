
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

export const parseRouteDepartures = async (rawText: string): Promise<Partial<RouteDeparture>[]> => {
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
