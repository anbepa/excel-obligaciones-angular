import { Injectable } from '@angular/core';

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbywXULtaug-OOBl4HJs0OcVLg0ZeL8e3u7jRZjl5Z_J9OnlTJLDomm3QP-BTJS0Fa4-/exec';

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
