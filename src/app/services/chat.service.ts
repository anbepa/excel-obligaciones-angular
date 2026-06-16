import { Injectable } from '@angular/core';

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwR3Nh9ea7cmmodxrKLS6TMbosaRtEXksYKHfTZXGWOFMT-6jHYD0-WQNM_cD3YHCCZ/exec';

export interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

@Injectable({ providedIn: 'root' })
export class ChatService {

  async send(sheetId: string, history: ChatMessage[], question: string): Promise<string> {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify({ action: 'chat', id: sheetId, history, question }),
    });

    if (!response.ok) throw new Error(`Chat error ${response.status}`);

    const result = await response.json() as { success: boolean; reply?: string; error?: string };
    if (!result.success || !result.reply) throw new Error(result.error || 'Sin respuesta');
    return result.reply?.trim() || 'Sin respuesta del asistente.';
  }
}
