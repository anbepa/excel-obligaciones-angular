import { Injectable } from '@angular/core';
import { createClient, SupabaseClient } from '@supabase/supabase-js';

const SUPABASE_URL = 'https://acpecmikvvnxwghsirjp.supabase.co';
const SUPABASE_KEY = 'sb_publishable_FrYXsyipxTI_JrKlNpOkbg_lZ2LOwhv';
const BUCKET_NAME = 'excel-previews';

export interface SimulationRecord {
  id: string;
  google_sheet_id: string;
  google_sheet_url: string;
  file_name: string;
  record_count: number;
  created_at: string;
}

export interface PresetPrompt {
  id: string;
  text: string;
  sort_order: number;
}

@Injectable({ providedIn: 'root' })
export class SupabaseStorageService {
  private readonly supabase: SupabaseClient;

  constructor() {
    this.supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
  }

  // ── Google Drive (existing) ──

  async uploadAndGetUrl(buffer: ArrayBuffer, fileName: string): Promise<string> {
    const sanitizedName = fileName.replace(/[^a-zA-Z0-9._-]/g, '_');
    const path = `${Date.now()}-${sanitizedName}`;
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    const { data, error } = await this.supabase.storage
      .from(BUCKET_NAME)
      .upload(path, blob, { upsert: false });

    if (error) throw error;

    const { data: urlData } = this.supabase.storage
      .from(BUCKET_NAME)
      .getPublicUrl(data.path);

    return urlData.publicUrl;
  }

  // ── Simulation history (new) ──

  async saveSimulation(sheetId: string, sheetUrl: string, fileName: string, recordCount: number, folderId: string): Promise<void> {
    const { error } = await this.supabase.from('simulations').insert({
      google_sheet_id: sheetId,
      google_sheet_url: sheetUrl,
      file_name: fileName,
      record_count: recordCount,
      folder_id: folderId
    });
    if (error) console.warn('Supabase history save failed:', error);
  }

  async deleteSimulation(id: string): Promise<void> {
    const { error } = await this.supabase.from('simulations').delete().eq('id', id);
    if (error) throw error;
  }

  async getSimulations(folderId: string, limit = 10): Promise<SimulationRecord[]> {
    const query = this.supabase
      .from('simulations')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(limit);

    // Traer registros con folder_id coincidente O sin folder_id (legacy)
    const { data, error } = await query.or(`folder_id.eq.${folderId},folder_id.is.null`);

    if (error) { console.warn('Supabase history load failed:', error); return []; }
    return (data ?? []) as SimulationRecord[];
  }

  // ── Preset prompts ──

  async getPresets(): Promise<PresetPrompt[]> {
    const { data, error } = await this.supabase
      .from('preset_prompts')
      .select('*')
      .order('sort_order', { ascending: true });

    if (error) { console.warn('Presets load failed:', error); return []; }
    return (data ?? []) as PresetPrompt[];
  }

  async addPreset(text: string): Promise<void> {
    const { data: max } = await this.supabase.from('preset_prompts').select('sort_order').order('sort_order', { ascending: false }).limit(1).single();
    const next = (max?.sort_order ?? 0) + 1;
    const { error } = await this.supabase.from('preset_prompts').insert({ text, sort_order: next });
    if (error) console.warn('Preset add failed:', error);
  }

  async updatePreset(id: string, text: string): Promise<void> {
    const { error } = await this.supabase.from('preset_prompts').update({ text }).eq('id', id);
    if (error) console.warn('Preset update failed:', error);
  }

  async deletePreset(id: string): Promise<void> {
    const { error } = await this.supabase.from('preset_prompts').delete().eq('id', id);
    if (error) console.warn('Preset delete failed:', error);
  }

  // ── User preferences ──

  async getPreferences(): Promise<{ driveFolderId: string; outputFileName: string }> {
    const { data, error } = await this.supabase
      .from('user_preferences')
      .select('drive_folder_id, output_file_name')
      .eq('id', 'default')
      .single();

    if (error || !data) {
      return { driveFolderId: '', outputFileName: '' };
    }
    return {
      driveFolderId: data.drive_folder_id || '',
      outputFileName: data.output_file_name || ''
    };
  }

  async savePreference(key: 'driveFolderId' | 'outputFileName', value: string): Promise<void> {
    const column = key === 'driveFolderId' ? 'drive_folder_id' : 'output_file_name';
    const { error } = await this.supabase
      .from('user_preferences')
      .upsert({ id: 'default', [column]: value }, { onConflict: 'id' });

    if (error) console.warn('Preference save failed:', error);
  }
}
