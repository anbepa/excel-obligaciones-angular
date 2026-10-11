import { Injectable } from '@angular/core';

/**
 * Cliente del asistente IA (sv-excel-agent + DeepSeek).
 *
 * En desarrollo apunta al backend local (http://localhost:8001).
 * En produccion (Vercel) usa la ruta relativa /api, servida por la
 * funcion Python del mismo repositorio.
 *
 * El backend devuelve el Excel modificado en base64 dentro de la misma
 * respuesta (las funciones serverless no guardan estado entre peticiones).
 */
const CHAT_API_URL =
  window.location.hostname === 'localhost' ? 'http://localhost:8001' : '';

export interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

export interface ChatResult {
  reply: string;
  downloadUrl?: string;
  downloadFileName?: string;
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
      success: boolean; response?: string; file_base64?: string; file_name?: string; error?: string;
    };
    if (!result.success || !result.response) throw new Error(result.error || 'Sin respuesta');

    let downloadUrl: string | undefined;
    if (result.file_base64) {
      const bytes = Uint8Array.from(atob(result.file_base64), c => c.charCodeAt(0));
      const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      downloadUrl = URL.createObjectURL(blob);
    }

    return {
      reply: result.response.trim() || 'Sin respuesta del asistente.',
      downloadUrl,
      downloadFileName: result.file_name || 'modificado.xlsx',
    };
  }
}
