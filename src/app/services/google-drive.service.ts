import { Injectable } from '@angular/core';

const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbywXULtaug-OOBl4HJs0OcVLg0ZeL8e3u7jRZjl5Z_J9OnlTJLDomm3QP-BTJS0Fa4-/exec';

@Injectable({ providedIn: 'root' })
export class GoogleDriveService {

  async uploadAndGetPreviewUrl(buffer: ArrayBuffer, fileName: string, folderId?: string): Promise<string> {
    const base64 = this.arrayBufferToBase64(buffer);

    const payload: Record<string, string> = { fileBase64: base64, fileName };
    if (folderId) payload['folderId'] = folderId;

    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error(`Google Apps Script responded with ${response.status}`);
    }

    const result = await response.json() as { success: boolean; url?: string; error?: string };

    if (!result.success || !result.url) {
      throw new Error(result.error || 'Google Apps Script returned failure');
    }

    return result.url;
  }

  async deleteFile(fileId: string): Promise<void> {
    const url = `${APPS_SCRIPT_URL}?action=delete&id=${encodeURIComponent(fileId)}`;
    const response = await fetch(url, { method: 'GET' });
    if (!response.ok) throw new Error(`Delete failed with ${response.status}`);
    const result = await response.json() as { success: boolean };
    if (!result.success) throw new Error('Delete failed');
  }

  private arrayBufferToBase64(buffer: ArrayBuffer): string {
    const bytes = new Uint8Array(buffer);
    const chunks: string[] = [];
    const CHUNK_SIZE = 0x8000; // 32KB chunks
    for (let i = 0; i < bytes.length; i += CHUNK_SIZE) {
      const chunk = bytes.subarray(i, i + CHUNK_SIZE);
      chunks.push(String.fromCharCode(...chunk));
    }
    return btoa(chunks.join(''));
  }
}
