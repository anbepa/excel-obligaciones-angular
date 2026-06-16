import { Component, HostListener, OnDestroy, OnInit, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { DomSanitizer, SafeResourceUrl } from '@angular/platform-browser';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { ExcelObligationsService } from './services/excel-obligations.service';
import { GoogleDriveService } from './services/google-drive.service';
import { ChatService, ChatMessage } from './services/chat.service';
import { SupabaseStorageService, SimulationRecord, PresetPrompt } from './services/supabase-storage.service';
import { GenerationLog, GenerationOptions, ObligationRecord } from './models/obligation.model';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent implements OnInit, OnDestroy {
  templateFile = signal<File | null>(null);
  obligationsFile = signal<File | null>(null);
  obligations = signal<ObligationRecord[]>([]);
  logs = signal<GenerationLog[]>([]);
  isProcessing = signal(false);
  errorMessage = signal<string | null>(null);
  previewUrl = signal<SafeResourceUrl | null>(null);
  previewDirectUrl = signal<string | null>(null);
  previewLoading = signal(false);
  isDragOver = signal(false);
  progressPhase = signal<'idle' | 'generating' | 'uploading' | 'done'>('idle');
  history = signal<SimulationRecord[]>([]);
  historyLoaded = signal(false);
  driveFolderId = signal('1u51rRx9XiKdzVzBjSg7wYgTOqmuoEjLC');
  sidebarCollapsed = signal(false);
  pasteMode = signal(false);
  pastedCsv = signal('');
  chatOpen = signal(false);
  chatMessages = signal<ChatMessage[]>([]);
  chatLoading = signal(false);
  chatInput = signal('');
  chatExpanded = signal(true);
  presets = signal<PresetPrompt[]>([]);
  presetEditing = signal<string | null>(null);
  presetNewText = signal('');
  presetsPanelOpen = signal(true);
  sqlHelpOpen = signal(false);
  paramsOpen = signal(false);
  historyOpen = signal(false);
  currentSheetId = signal('');
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  private _folderDebounce: any = null;

  sqlByClient = `SELECT to2.invoice_number, to2.invoice_value, to2.created_at,
       to2.created_at AS transfer_date, to2.net_due_date, to2.total_due_date,
       to2.provider_rate, to2.remunerative_rate, to2.prepayment_rate
FROM schsaf.tbl_obligations to2
WHERE to2.client_id IN (
  SELECT c.id FROM schsaf.tbl_client c
  WHERE c.document_number IN ('1999231309')
) AND to2.state IN ('CURRENT','EXPIRED')`;

  sqlByObligation = `SELECT to2.invoice_number, to2.invoice_value, to2.created_at,
       to2.created_at AS transfer_date, to2.net_due_date, to2.total_due_date,
       to2.provider_rate, to2.remunerative_rate, to2.prepayment_rate
FROM schsaf.tbl_obligations to2
WHERE to2.id = 713409
AND to2.state IN ('CURRENT','EXPIRED')`;

  options: GenerationOptions = {
    divideRatesBy100: true,
    keepObligationsSheet: true,
    outputFileName: 'simulaciones_obligaciones.xlsx'
  };

  constructor(
    private readonly excelService: ExcelObligationsService,
    private readonly driveService: GoogleDriveService,
    private readonly chatService: ChatService,
    private readonly supabaseService: SupabaseStorageService,
    private readonly sanitizer: DomSanitizer,
  ) {}

  ngOnInit(): void {
    this.loadPresets();
    this.loadPreferences();
  }

  ngOnDestroy(): void {
    this.unlockBodyScroll();
  }

  get canGenerate(): boolean {
    return Boolean(this.templateFile()) && !this.isProcessing();
  }

  // ── Helpers ──

  formatSize(bytes: number | undefined): string {
    if (bytes === undefined || bytes === null) return '';
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1048576).toFixed(1)} MB`;
  }

  private messageFromError(error: unknown): string {
    if (error instanceof Error) return error.message;
    return String(error);
  }

  private safeOutputName(): string {
    const name = this.options.outputFileName.trim() || 'simulaciones_obligaciones.xlsx';
    return name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`;
  }

  saveDriveFolderId(): void {
    const id = this.driveFolderId();
    localStorage.setItem('driveFolderId', id);
    this.supabaseService.savePreference('driveFolderId', id);
    clearTimeout(this._folderDebounce);
    this._folderDebounce = setTimeout(() => this.loadHistory(), 500);
  }

  saveFileName(): void {
    const name = this.options.outputFileName;
    localStorage.setItem('outputFileName', name);
    this.supabaseService.savePreference('outputFileName', name);
  }

  private loadPreferences(): void {
    this.supabaseService.getPreferences().then(prefs => {
      if (prefs.driveFolderId) {
        this.driveFolderId.set(prefs.driveFolderId);
        localStorage.setItem('driveFolderId', prefs.driveFolderId);
      } else {
        const local = localStorage.getItem('driveFolderId');
        if (local) this.driveFolderId.set(local);
      }
      if (prefs.outputFileName) {
        this.options.outputFileName = prefs.outputFileName;
        localStorage.setItem('outputFileName', prefs.outputFileName);
      } else {
        const local = localStorage.getItem('outputFileName');
        if (local) this.options.outputFileName = local;
      }
      // Cargar historial con el ID que quedó
      this.loadHistory();
    }).catch(() => {
      const localDrive = localStorage.getItem('driveFolderId');
      if (localDrive) this.driveFolderId.set(localDrive);
      const localFile = localStorage.getItem('outputFileName');
      if (localFile) this.options.outputFileName = localFile;
      this.loadHistory();
    });
  }

  // ── Drag & Drop ──

  onDragOver(event: DragEvent): void { event.preventDefault(); this.isDragOver.set(true); }
  onDragLeave(event: DragEvent): void { event.preventDefault(); this.isDragOver.set(false); }

  onDropTemplate(event: DragEvent): void {
    event.preventDefault(); this.isDragOver.set(false);
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setTemplate(file);
  }

  onDropObligations(event: DragEvent): void {
    event.preventDefault(); this.isDragOver.set(false);
    const file = event.dataTransfer?.files?.[0];
    if (file) this.setObligationsFile(file);
  }

  // ── Template ──

  onTemplateSelected(event: Event): void {
    const file = this.getSelectedFile(event);
    if (file) this.setTemplate(file);
  }

  removeTemplate(): void {
    this.templateFile.set(null);
    this.previewUrl.set(null);
    this.previewDirectUrl.set(null);
    this.unlockBodyScroll();
  }

  private setTemplate(file: File): void {
    this.templateFile.set(file);
    this.errorMessage.set(null);
    this.logs.set([]);
    this.previewUrl.set(null);
    this.previewDirectUrl.set(null);
    this.unlockBodyScroll();
  }

  // ── Obligations ──

  async onObligationsSelected(event: Event): Promise<void> {
    const file = this.getSelectedFile(event);
    if (file) await this.setObligationsFile(file);
  }

  removeObligations(): void {
    this.obligationsFile.set(null);
    this.obligations.set([]);
    this.pasteMode.set(false);
    this.pastedCsv.set('');
  }

  openPasteMode(): void {
    this.pasteMode.set(true);
    this.pastedCsv.set('');
    this.obligationsFile.set(null);
    this.obligations.set([]);
  }

  cancelPaste(): void {
    this.pasteMode.set(false);
    this.pastedCsv.set('');
  }

  loadFromClipboard(): void {
    const text = this.pastedCsv().trim();
    if (!text) return;
    try {
      const records = this.excelService.parseObligationsFromCsvText(text);
      if (!records.length) { this.errorMessage.set('No se encontraron obligaciones válidas.'); return; }
      this.obligations.set(records);
      this.obligationsFile.set(new File([text], 'clipboard.csv', { type: 'text/csv' }));
      this.errorMessage.set(null);
      this.pasteMode.set(false);
    } catch (error) { this.errorMessage.set(this.messageFromError(error)); this.obligations.set([]); }
  }

  onPasteObligations(event: ClipboardEvent): void {
    const text = event.clipboardData?.getData('text')?.trim();
    if (!text) return;
    if (!text.includes('\n') || (!text.includes(',') && !text.includes('\t') && !text.includes(';'))) return;
    event.preventDefault();
    try {
      const records = this.excelService.parseObligationsFromCsvText(text);
      if (!records.length) return;
      this.obligations.set(records);
      this.obligationsFile.set(new File([text], 'clipboard.csv', { type: 'text/csv' }));
      this.errorMessage.set(null);
    } catch (error) { this.errorMessage.set(this.messageFromError(error)); this.obligations.set([]); }
  }

  @HostListener('document:paste', ['$event'])
  onDocumentPaste(event: ClipboardEvent): void {
    if (this.obligationsFile() || this.pasteMode() || this.isProcessing()) return;
    const text = event.clipboardData?.getData('text')?.trim();
    if (!text) return;
    if (!text.includes('\n') || (!text.includes(',') && !text.includes('\t') && !text.includes(';'))) return;
    event.preventDefault();
    try {
      const records = this.excelService.parseObligationsFromCsvText(text);
      if (!records.length) return;
      this.obligations.set(records);
      this.obligationsFile.set(new File([text], 'clipboard.csv', { type: 'text/csv' }));
      this.errorMessage.set(null);
    } catch (error) { this.errorMessage.set(this.messageFromError(error)); this.obligations.set([]); }
  }

  private async setObligationsFile(file: File): Promise<void> {
    this.obligationsFile.set(file);
    this.errorMessage.set(null);
    this.logs.set([]);
    this.previewUrl.set(null);
    this.previewDirectUrl.set(null);
    try {
      this.isProcessing.set(true);
      const obligations = await this.excelService.readObligations(file);
      this.obligations.set(obligations);
      this.logs.set([{ type: 'info', message: `${obligations.length} registro${obligations.length !== 1 ? 's' : ''} cargado${obligations.length !== 1 ? 's' : ''} desde ${file.name}` }]);
    } catch (error) {
      this.errorMessage.set(this.messageFromError(error));
      this.obligations.set([]);
    } finally { this.isProcessing.set(false); }
  }

  // ── Generate ──

  async generate(): Promise<void> {
    const template = this.templateFile();
    if (!template) return;
    this.errorMessage.set(null); this.logs.set([]); this.isProcessing.set(true);
    this.progressPhase.set('generating');
    this.previewUrl.set(null); this.previewDirectUrl.set(null); this.previewLoading.set(true);
    this.unlockBodyScroll();

    try {
      const workbook = await this.excelService.readWorkbook(template);
      let obligations = this.obligations();
      if (!this.obligationsFile()) {
        const ws = workbook.getWorksheet('Obligaciones');
        if (!ws) throw new Error('Carga un archivo de obligaciones o usa una plantilla con hoja "Obligaciones".');
        obligations = this.excelService.readObligationsFromWorksheet(ws);
        this.obligations.set(obligations);
      }
      if (!obligations.length) throw new Error('No hay obligaciones válidas para procesar.');

      const { buffer } = await this.excelService.generateWorkbook(workbook, obligations, this.options);
      const fileName = this.safeOutputName();
      saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), fileName);
      this.progressPhase.set('uploading');

      try {
        this.replaceFormulaForGoogleSheets(workbook);
        const driveBuffer = await workbook.xlsx.writeBuffer();
        const editUrl = await this.driveService.uploadAndGetPreviewUrl(driveBuffer as ArrayBuffer, fileName, this.driveFolderId());
        this.previewDirectUrl.set(editUrl);
        this.previewUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(editUrl));
        this.lockBodyScroll();
        const sheetId = editUrl.match(/\/d\/([^/]+)/)?.[1] ?? '';
        this.currentSheetId.set(sheetId);
        this.chatMessages.set([]); this.chatOpen.set(false);
        this.supabaseService.saveSimulation(sheetId, editUrl, fileName, obligations.length, this.driveFolderId()).then(() => { this.loadHistory(); this.historyOpen.set(true); }).catch(() => {});
        this.progressPhase.set('done');
      } catch (uploadErr) {
        console.error('Drive upload failed:', uploadErr);
        const msg = this.messageFromError(uploadErr);
        this.logs.set([{ type: 'warning', message: 'No se pudo subir a Drive: ' + msg }]);
        this.previewDirectUrl.set(null); this.progressPhase.set('done');
      }
    } catch (error) { this.errorMessage.set(this.messageFromError(error)); }
    finally { this.isProcessing.set(false); this.previewLoading.set(false); }
  }

  private replaceFormulaForGoogleSheets(workbook: ExcelJS.Workbook): void {
    workbook.worksheets.forEach(sheet => {
      const cell = sheet.getCell('G8');
      const raw = cell.value;
      if (typeof raw === 'object' && raw !== null && 'formula' in raw) {
        cell.value = { formula: 'ROUND(((B8/I8)^(360/H8))-1;6)' };
      }
    });
  }

  private getSelectedFile(event: Event): File | null {
    const input = event.target as HTMLInputElement;
    return input.files?.[0] ?? null;
  }

  // ── Scroll lock ──

  private lockBodyScroll(): void { document.body.style.overflow = 'hidden'; }
  private unlockBodyScroll(): void { document.body.style.overflow = ''; }

  // ── Sidebar ──

  toggleSidebar(): void { this.sidebarCollapsed.update(v => !v); }

  // ── History ──

  private loadHistory(): void {
    this.supabaseService.getSimulations(this.driveFolderId()).then(data => { this.history.set(data); this.historyLoaded.set(true); });
  }

  restoreSimulation(record: SimulationRecord): void {
    const url = record.google_sheet_url;
    this.previewUrl.set(this.sanitizer.bypassSecurityTrustResourceUrl(url));
    this.previewDirectUrl.set(url);
    this.lockBodyScroll();
    this.currentSheetId.set(record.google_sheet_id);
    this.chatMessages.set([]); this.chatOpen.set(false);
    this.progressPhase.set('idle');
  }

  deleteHistoryItem(event: MouseEvent, record: SimulationRecord): void {
    event.stopPropagation();
    this.driveService.deleteFile(record.google_sheet_id).catch(() => {});
    this.supabaseService.deleteSimulation(record.id).then(() => this.loadHistory()).catch(() => this.loadHistory());
  }

  // ── Chat ──

  toggleChat(): void {
    this.chatOpen.update(v => !v);
    if (this.chatOpen()) { this.chatExpanded.set(true); setTimeout(() => document.getElementById('chat-input')?.focus(), 100); }
  }

  toggleChatSize(): void { this.chatExpanded.update(v => !v); }

  async sendChatMessage(): Promise<void> {
    const question = this.chatInput().trim();
    if (!question || !this.currentSheetId()) return;
    const messages = this.chatMessages();
    this.chatMessages.set([...messages, { role: 'user', content: question }]);
    this.chatInput.set(''); this.chatLoading.set(true);
    try {
      const reply = await this.chatService.send(this.currentSheetId(), messages, question);
      this.chatMessages.set([...this.chatMessages(), { role: 'assistant', content: reply?.trim() || 'Sin respuesta.' }]);
    } catch { this.chatMessages.set([...this.chatMessages(), { role: 'assistant', content: 'Error al comunicarse con el asistente IA.' }]); }
    finally { this.chatLoading.set(false); setTimeout(() => document.getElementById('chat-scroll')?.scrollTo(0, 99999), 50); }
  }

  onChatKeydown(event: KeyboardEvent): void {
    if (event.key === 'Enter' && !event.shiftKey) { event.preventDefault(); this.sendChatMessage(); }
  }

  mdToHtml(text: string): string {
    if (!text) return '';
    let html = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

    // Tablas markdown: formato | pipe | o tabulado con ---
    html = html.replace(/(?:^|\n\n)((?:\|.+\|(?:\n|$))+)/gm, (_match: string, block: string) => renderTable(block, '|'));
    html = html.replace(/(?:^|\n\n)((?:[^\n]+\t[^\n]+(?:\n|$)){2,})/gm, (_match: string, block: string) => {
      // Tab-separado: detectar si tiene separador --- o son solo líneas de datos
      const lines = block.trim().split('\n');
      // Si es formato con separador --- (línea de solo --- y \t)
      const hasSeparator = lines.some(l => /^[\t\s\-:]+$/.test(l) && l.includes('-'));
      if (hasSeparator) return renderTable(block, '\t');
      // Si son solo líneas de datos (sin separador), primera línea = header
      if (lines.length >= 2) return renderTable(block, '\t');
      return _match;
    });
    // Tablas con pipes internos pero sin pipes en bordes: "A | B | C"
    html = html.replace(/(?:^|\n\n)((?:[^\n]+\s*\|\s*[^\n]+(?:\n|$)){2,})/gm, (_match: string, block: string) => {
      const lines = block.trim().split('\n');
      if (lines.length >= 2 && lines.every(l => l.includes('|'))) return renderTable(block, '|');
      return _match;
    });

    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
    html = html.replace(/`(.+?)`/g, '<code>$1</code>');
    html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
    html = html.replace(/((?:<li>.*<\/li>\n?)+)/g, '<ul>$1</ul>');
    html = html.replace(/\n\n/g, '</p><p>');
    html = html.replace(/\n/g, '<br>');
    if (!html.startsWith('<')) html = '<p>' + html + '</p>';
    return html;

    function renderTable(block: string, sep: string): string {
      const lines = block.trim().split('\n');
      if (lines.length < 2) return block;
      const dataLines = lines.filter(l => !/^[\s\-:|]+\n?$/.test(l) && !/^[\t\s\-:]+\n?$/.test(l));
      if (dataLines.length === 0) return block;
      const thead = dataLines[0];
      const tbody = dataLines.slice(1);
      const splitCells = (row: string) => row.split(sep).filter(c => c.trim());
      const headerCells = splitCells(thead).map(c => `<th>${c.trim()}</th>`).join('');
      const bodyRows = tbody.map(row => {
        const cells = splitCells(row).map(c => `<td>${c.trim()}</td>`).join('');
        return `<tr>${cells}</tr>`;
      }).join('');
      return `\n<div class="md-table-wrap"><table><thead><tr>${headerCells}</tr></thead><tbody>${bodyRows}</tbody></table></div>\n`;
    }
  }

  // ── Presets ──

  sendPreset(text: string): void {
    if (!this.currentSheetId()) return;
    if (!this.chatOpen()) this.toggleChat();
    this.chatInput.set(text);
    setTimeout(() => this.sendChatMessage(), 100);
  }

  addPreset(): void {
    const text = this.presetNewText().trim(); if (!text) return;
    this.supabaseService.addPreset(text).then(() => { this.presetNewText.set(''); this.loadPresets(); });
  }

  startEditPreset(preset: PresetPrompt): void { this.presetEditing.set(preset.id); this.presetNewText.set(preset.text); }

  saveEditPreset(id: string): void {
    const text = this.presetNewText().trim(); if (!text) return;
    this.supabaseService.updatePreset(id, text).then(() => { this.presetEditing.set(null); this.presetNewText.set(''); this.loadPresets(); });
  }

  cancelEditPreset(): void { this.presetEditing.set(null); this.presetNewText.set(''); }

  deletePreset(id: string): void { this.supabaseService.deletePreset(id).then(() => this.loadPresets()); }

  copySql(sql: string): void {
    navigator.clipboard.writeText(sql).then(() => {
      const prev = this.errorMessage();
      this.errorMessage.set('SQL copiado al portapapeles.');
      setTimeout(() => this.errorMessage.set(prev), 1500);
    }).catch(() => {});
  }

  private loadPresets(): void { this.supabaseService.getPresets().then(data => this.presets.set(data)); }
}
