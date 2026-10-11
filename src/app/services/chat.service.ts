import { Injectable } from '@angular/core';

/**
 * Cliente del nuevo asistente IA (sv-excel-agent + DeepSeek).
 * Reemplaza al anterior basado en Google Apps Script.
 *
 * El backend (Excel/excel-chat-backend, FastAPI en http://localhost:8001)
 * recibe el Excel actual y la pregunta, ejecuta el agente sobre el archivo
 * y devuelve la respuesta + URL de descarga del Excel modificado.
 */
const CHAT_API_URL = 'http://localhost:8001';

export interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

export interface ChatResult {
  reply: string;
  downloadUrl?: string;
}

@Injectable({ providedIn: 'root' })
export class ChatService {

  async send(file: Blob, fileName: string, history: ChatMessage[], question: string): Promise<ChatResult> {
    const form = new FormData();
    form.append('message', question);
    form.append('file', file, fileName.toLowerCase().endsWith('.xlsx') ? fileName : `${fileName}.xlsx`);
    form.append('history', JSON.stringify(history));

    const response = await fetch(`${CHAT_API_URL}/api/chat`, {
      method: 'POST',
      body: form,
    });

    if (!response.ok) throw new Error(`Chat error ${response.status}`);

    const result = await response.json() as {
      success: boolean; response?: string; output_file?: string; error?: string;
    };
    if (!result.success || !result.response) throw new Error(result.error || 'Sin respuesta');

    return {
      reply: result.response.trim() || 'Sin respuesta del asistente.',
      downloadUrl: result.output_file ? `${CHAT_API_URL}${result.output_file}` : undefined,
    };
  }
}
