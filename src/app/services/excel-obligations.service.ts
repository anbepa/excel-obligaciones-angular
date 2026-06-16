import { Injectable } from '@angular/core';
import * as ExcelJS from 'exceljs';
import * as Papa from 'papaparse';
import { CELL_MAP, GenerationLog, GenerationOptions, ObligationField, ObligationRecord } from '../models/obligation.model';

const HEADER_ALIASES: Record<ObligationField, string[]> = {
  obligationId: ['obligacion', 'obligación', 'id obligacion', 'id obligación', 'numero obligacion', 'número obligación', 'invoice_number', 'invoice', 'obligation', 'obligation id'],
  capital: ['capital', 'valor capital', 'monto', 'valor', 'invoice_value', 'amount', 'value'],
  disbursementDate: ['fecha de desembolso', 'fecha desembolso', 'desembolso', 'created_at', 'disbursement_date', 'disbursement date', 'fecha desembolso real'],
  transferDate: ['fecha de transferencia', 'fecha transferencia', 'transferencia', 'transfer_date', 'transfer date', 'fecha transferencia real', 'created_at_2'],
  netDueDate: ['fecha vcto neto', 'fecha vencimiento neto', 'vcto neto', 'vencimiento neto', 'net_due_date', 'net due date'],
  totalDueDate: ['fecha vcto total', 'fecha vencimiento total', 'vcto total', 'vencimiento total', 'total_due_date', 'total due date'],
  businessRate: ['tasa negocio', 'tasa/spread proveedor (tasa negocio)', 'tasa spread proveedor', 'tasa/proveedor', 'tasa proveedor', 'provider_rate', 'business_rate', 'business rate', 'spread'],
  remuneratoryRate: ['tasa interes remuneratorio', 'tasa interés remuneratorio', 'tasa ea interés remuneratorio', 'tasa ea interes remuneratorio', 'interés remuneratorio', 'interes remuneratorio', 'remunerative_rate', 'remuneratory_rate', 'remuneratory rate', 'interest rate'],
  dppRate: ['tasa dpp', 'dpp', 'prepayment_rate', 'dpp_rate', 'dpp rate', 'prepayment']
};

@Injectable({ providedIn: 'root' })
export class ExcelObligationsService {
  async readWorkbook(file: File): Promise<ExcelJS.Workbook> {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return workbook;
  }

  async readObligations(file: File): Promise<ObligationRecord[]> {
    const extension = this.getExtension(file.name);
    if (extension === 'csv') {
      return this.readObligationsFromCsv(file);
    }

    const workbook = await this.readWorkbook(file);
    const worksheet = workbook.getWorksheet('Obligaciones') ?? workbook.worksheets[0];
    if (!worksheet) {
      throw new Error('No se encontró una hoja con obligaciones en el archivo cargado.');
    }
    return this.readObligationsFromWorksheet(worksheet);
  }

  readObligationsFromWorksheet(worksheet: ExcelJS.Worksheet): ObligationRecord[] {
    const headerRowNumber = this.findHeaderRow(worksheet);
    if (!headerRowNumber) {
      throw new Error('No se encontraron encabezados válidos para Obligación, Capital y Tasas.');
    }

    const headerMap = this.buildHeaderMap(worksheet.getRow(headerRowNumber));
    const missing = this.getMissingRequiredFields(headerMap);
    if (missing.length) {
      throw new Error(`Faltan columnas obligatorias: ${missing.join(', ')}`);
    }

    const records: ObligationRecord[] = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRowNumber) return;
      if (this.isTargetCellHelperRow(row)) return;

      const record = this.rowToRecord(row, headerMap, rowNumber);
      if (record.obligationId) {
        records.push(record);
      }
    });

    return records;
  }

  parseObligationsFromCsvText(csvText: string): ObligationRecord[] {
    // Track duplicate headers (e.g. two "created_at" columns)
    const seenHeaders = new Map<string, number>();

    const parsed = Papa.parse<Record<string, string>>(csvText, {
      header: true,
      skipEmptyLines: true,
      delimiter: '',
      transformHeader: (header) => {
        const h = header.trim();
        const count = seenHeaders.get(h) || 0;
        seenHeaders.set(h, count + 1);
        // Rename second occurrence: "created_at" → "created_at_2" so both are preserved
        return count > 0 ? `${h}_${count + 1}` : h;
      }
    });

    if (parsed.errors.length) {
      const first = parsed.errors[0];
      throw new Error(`Error leyendo CSV en fila ${first.row ?? 'N/A'}: ${first.message}`);
    }

    const headers = parsed.meta.fields ?? [];
    const map = this.buildObjectHeaderMap(headers);
    const missing = this.getMissingRequiredFields(map);
    if (missing.length) {
      throw new Error(`Faltan columnas obligatorias en el CSV: ${missing.join(', ')}`);
    }

    return parsed.data
      .map((row, index) => this.objectToRecord(row, map, index + 2))
      .filter((record) => Boolean(record.obligationId));
  }

  async generateWorkbook(
    templateWorkbook: ExcelJS.Workbook,
    obligations: ObligationRecord[],
    options: GenerationOptions
  ): Promise<{ buffer: ArrayBuffer; logs: GenerationLog[] }> {
    const logs: GenerationLog[] = [];
    const baseSheet = templateWorkbook.getWorksheet('Base');
    if (!baseSheet) {
      throw new Error('La plantilla debe tener una hoja llamada exactamente "Base".');
    }

    if (!options.keepObligationsSheet) {
      const obligationsSheet = templateWorkbook.getWorksheet('Obligaciones');
      if (obligationsSheet) templateWorkbook.removeWorksheet(obligationsSheet.id);
    }

    const usedNames = new Set(templateWorkbook.worksheets.map((ws) => ws.name));

    obligations.forEach((obligation, index) => {
      const sheetName = this.makeUniqueWorksheetName(obligation.obligationId, usedNames);
      const newSheet = templateWorkbook.addWorksheet(sheetName);
      this.cloneWorksheet(baseSheet, newSheet);
      this.applyObligationToSheet(newSheet, obligation, options);
      usedNames.add(sheetName);
      logs.push({ type: 'success', message: `${index + 1}. Hoja creada: ${sheetName}` });
    });

    templateWorkbook.calcProperties.fullCalcOnLoad = true;
    (templateWorkbook.calcProperties as unknown as { forceFullCalc?: boolean }).forceFullCalc = true;

    const output = await templateWorkbook.xlsx.writeBuffer();
    return { buffer: output as ArrayBuffer, logs };
  }

  private async readObligationsFromCsv(file: File): Promise<ObligationRecord[]> {
    const text = await file.text();
    const parsed = Papa.parse<Record<string, string>>(text, {
      header: true,
      skipEmptyLines: true,
      delimiter: '',
      transformHeader: (header) => header.trim()
    });

    if (parsed.errors.length) {
      const first = parsed.errors[0];
      throw new Error(`Error leyendo CSV en fila ${first.row ?? 'N/A'}: ${first.message}`);
    }

    const headers = parsed.meta.fields ?? [];
    const map = this.buildObjectHeaderMap(headers);
    const missing = this.getMissingRequiredFields(map);
    if (missing.length) {
      throw new Error(`Faltan columnas obligatorias en el CSV: ${missing.join(', ')}`);
    }

    return parsed.data
      .map((row, index) => this.objectToRecord(row, map, index + 2))
      .filter((record) => Boolean(record.obligationId));
  }

  private applyObligationToSheet(sheet: ExcelJS.Worksheet, record: ObligationRecord, options: GenerationOptions): void {
    sheet.getCell(CELL_MAP.obligationId).value = record.obligationId;
    sheet.getCell(CELL_MAP.capital).value = this.normalizeNumberOrText(record.capital);
    this.setCellDate(sheet, CELL_MAP.disbursementDate, record.disbursementDate);
    this.setCellDate(sheet, CELL_MAP.transferDate, record.transferDate);
    this.setCellDate(sheet, CELL_MAP.netDueDate, record.netDueDate);
    this.setCellDate(sheet, CELL_MAP.totalDueDate, record.totalDueDate);
    sheet.getCell(CELL_MAP.businessRate).value = this.normalizeRate(record.businessRate, options.divideRatesBy100);
    sheet.getCell(CELL_MAP.remuneratoryRate).value = this.normalizeRate(record.remuneratoryRate, options.divideRatesBy100);
    sheet.getCell(CELL_MAP.dppRate).value = this.normalizeRate(record.dppRate, options.divideRatesBy100);

    if (options.divideRatesBy100) {
      sheet.getCell(CELL_MAP.businessRate).numFmt = '0.00%';
      sheet.getCell(CELL_MAP.remuneratoryRate).numFmt = '0.00%';
      sheet.getCell(CELL_MAP.dppRate).numFmt = '0.00%';
    }
  }

  private setCellDate(sheet: ExcelJS.Worksheet, cellRef: string, value: unknown): void {
    const cell = sheet.getCell(cellRef);
    const serial = this.toExcelDateSerial(value);
    if (serial !== null) {
      cell.value = serial;
      cell.numFmt = 'dd/mm/yyyy';
    } else {
      cell.value = value !== null && value !== undefined ? String(value) : null;
    }
  }

  private toExcelDateSerial(value: unknown): number | null {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'object' && 'result' in value) {
      return this.toExcelDateSerial((value as { result?: unknown }).result);
    }

    if (typeof value === 'number') {
      if (value > 25569 && value < 100000) return Math.floor(value);
      return this.dateToSerial(new Date(value));
    }

    if (value instanceof Date) return this.dateToSerial(value);

    const text = String(value).trim();
    const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      return this.dateToSerial(new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3])));
    }

    const parsed = this.parseSpanishDate(text);
    return parsed ? this.dateToSerial(parsed) : null;
  }

  private dateToSerial(date: Date): number {
    const MS_PER_DAY = 86400000;
    const EXCEL_EPOCH = 25569;
    // Use UTC to get a pure integer serial — no timezone fraction
    const utcMs = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
    const serial = utcMs / MS_PER_DAY + EXCEL_EPOCH;
    return Math.floor(serial);
  }

  private cloneWorksheet(source: ExcelJS.Worksheet, target: ExcelJS.Worksheet): void {
    target.properties = this.deepClone(source.properties);
    target.pageSetup = this.deepClone(source.pageSetup);
    target.headerFooter = this.deepClone(source.headerFooter);
    target.views = this.deepClone(source.views);
    target.autoFilter = source.autoFilter ? this.deepClone(source.autoFilter) : undefined;

    source.columns.forEach((column, index) => {
      const targetColumn = target.getColumn(index + 1);
      targetColumn.width = column.width;
      if (column.hidden !== undefined) targetColumn.hidden = column.hidden;
      if (column.outlineLevel !== undefined) targetColumn.outlineLevel = column.outlineLevel;
      if (column.style) targetColumn.style = this.deepClone(column.style);
    });

    source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const targetRow = target.getRow(rowNumber);
      targetRow.height = row.height;
      targetRow.hidden = row.hidden;
      targetRow.outlineLevel = row.outlineLevel;
      if (row.model?.style) {
        (targetRow as unknown as { _style?: Partial<ExcelJS.Style> })._style = this.deepClone(row.model.style);
      }

      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const targetCell = targetRow.getCell(colNumber);
        targetCell.value = this.deepClone(cell.value);
        targetCell.style = this.deepClone(cell.style);
        targetCell.numFmt = cell.numFmt;
        targetCell.font = this.deepClone(cell.font);
        targetCell.alignment = this.deepClone(cell.alignment);
        targetCell.border = this.deepClone(cell.border);
        targetCell.fill = this.deepClone(cell.fill);
        targetCell.protection = this.deepClone(cell.protection);
      });
      targetRow.commit();
    });

    const sourceWithMerges = source as unknown as { _merges?: Record<string, { model: { top: number; left: number; bottom: number; right: number } }> };
    Object.values(sourceWithMerges._merges ?? {}).forEach((merge) => {
      const { top, left, bottom, right } = merge.model;
      target.mergeCells(top, left, bottom, right);
    });
  }

  private rowToRecord(row: ExcelJS.Row, headerMap: Map<ObligationField, number>, sourceRow: number): ObligationRecord {
    const get = (field: ObligationField) => row.getCell(headerMap.get(field) ?? 0).value as string | number | Date | null;
    return {
      obligationId: String(get('obligationId') ?? '').trim(),
      capital: get('capital'),
      disbursementDate: get('disbursementDate'),
      transferDate: get('transferDate'),
      netDueDate: get('netDueDate'),
      totalDueDate: get('totalDueDate'),
      businessRate: get('businessRate') as number | string | null,
      remuneratoryRate: get('remuneratoryRate') as number | string | null,
      dppRate: get('dppRate') as number | string | null,
      sourceRow
    };
  }

  private objectToRecord(row: Record<string, string>, headerMap: Map<ObligationField, string>, sourceRow: number): ObligationRecord {
    const get = (field: ObligationField) => row[headerMap.get(field) ?? ''] ?? null;
    return {
      obligationId: String(get('obligationId') ?? '').trim(),
      capital: get('capital'),
      disbursementDate: get('disbursementDate'),
      transferDate: get('transferDate'),
      netDueDate: get('netDueDate'),
      totalDueDate: get('totalDueDate'),
      businessRate: get('businessRate'),
      remuneratoryRate: get('remuneratoryRate'),
      dppRate: get('dppRate'),
      sourceRow
    };
  }

  private findHeaderRow(worksheet: ExcelJS.Worksheet): number | null {
    for (let rowNumber = 1; rowNumber <= Math.min(10, worksheet.rowCount); rowNumber += 1) {
      const map = this.buildHeaderMap(worksheet.getRow(rowNumber));
      if (map.has('obligationId') && map.has('capital')) return rowNumber;
    }
    return null;
  }

  private buildHeaderMap(row: ExcelJS.Row): Map<ObligationField, number> {
    const map = new Map<ObligationField, number>();
    row.eachCell((cell, colNumber) => {
      const normalized = this.normalizeHeader(this.cellToPlainText(cell.value));
      (Object.keys(HEADER_ALIASES) as ObligationField[]).forEach((field) => {
        if (!map.has(field) && HEADER_ALIASES[field].some((alias) => normalized === this.normalizeHeader(alias) || normalized.includes(this.normalizeHeader(alias)))) {
          map.set(field, colNumber);
        }
      });
    });
    return map;
  }

  private buildObjectHeaderMap(headers: string[]): Map<ObligationField, string> {
    const map = new Map<ObligationField, string>();
    headers.forEach((header) => {
      const normalized = this.normalizeHeader(header);
      (Object.keys(HEADER_ALIASES) as ObligationField[]).forEach((field) => {
        if (!map.has(field) && HEADER_ALIASES[field].some((alias) => normalized === this.normalizeHeader(alias) || normalized.includes(this.normalizeHeader(alias)))) {
          map.set(field, header);
        }
      });
    });
    return map;
  }

  private getMissingRequiredFields(map: Map<ObligationField, unknown>): string[] {
    const required: ObligationField[] = [
      'obligationId', 'capital', 'disbursementDate', 'transferDate', 'netDueDate',
      'totalDueDate', 'businessRate', 'remuneratoryRate', 'dppRate'
    ];
    return required.filter((field) => !map.has(field)).map((field) => `${field} → ${CELL_MAP[field]}`);
  }

  private isTargetCellHelperRow(row: ExcelJS.Row): boolean {
    const values = row.values as Array<unknown>;
    const text = values.map((value) => String(value ?? '').trim().toUpperCase()).join('|');
    return ['A8', 'B8', 'B4', 'C8', 'D8', 'E8', 'F8', 'L8', 'AB8'].every((cellRef) => text.includes(cellRef));
  }

  private normalizeRate(value: unknown, divideBy100: boolean): number | null {
    const numberValue = this.toNumber(value);
    if (numberValue === null) return null;
    return divideBy100 ? numberValue / 100 : numberValue;
  }

  private normalizeNumberOrText(value: unknown): number | string | Date | null {
    const numberValue = this.toNumber(value);
    return numberValue ?? this.normalizeDateOrValue(value);
  }

  private normalizeDateOrValue(value: unknown): string | number | Date | null {
    if (value === null || value === undefined || value === '') return null;
    if (value instanceof Date) return value;
    if (typeof value === 'number') return value;
    if (typeof value === 'object' && 'result' in value) {
      return this.normalizeDateOrValue((value as { result?: unknown }).result);
    }

    const text = String(value).trim();
    const date = this.parseSpanishDate(text);
    return date ?? text;
  }

  private parseSpanishDate(text: string): Date | null {
    const match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (!match) return null;
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);
    if (!day || !month || !year) return null;
    return new Date(year, month - 1, day);
  }

  private toNumber(value: unknown): number | null {
    if (value === null || value === undefined || value === '') return null;
    if (typeof value === 'number') return value;
    if (typeof value === 'object' && 'result' in value) {
      return this.toNumber((value as { result?: unknown }).result);
    }

    const raw = String(value).replace(/%/g, '').replace(/\s/g, '').trim();
    if (!raw) return null;

    const lastComma = raw.lastIndexOf(',');
    const lastDot = raw.lastIndexOf('.');
    let normalized = raw;

    if (lastComma >= 0 && lastDot >= 0) {
      // Si existen ambos símbolos, el último se asume decimal y el otro separador de miles.
      normalized = lastComma > lastDot
        ? raw.replace(/\./g, '').replace(',', '.')
        : raw.replace(/,/g, '');
    } else if (lastComma >= 0) {
      // Formato latino: 16,6544 -> 16.6544 o 1.234,56 -> 1234.56.
      normalized = raw.replace(/\./g, '').replace(',', '.');
    } else if (lastDot >= 0) {
      // Formato internacional: 16.6544 queda 16.6544; 1,234.56 queda 1234.56.
      const dotGroups = raw.split('.');
      const looksLikeThousands = dotGroups.length > 2 && dotGroups.slice(1).every((group) => group.length === 3);
      normalized = looksLikeThousands ? raw.replace(/\./g, '') : raw.replace(/,/g, '');
    }

    const numberValue = Number(normalized);
    return Number.isFinite(numberValue) ? numberValue : null;
  }

  private makeUniqueWorksheetName(rawName: string, usedNames: Set<string>): string {
    const cleaned = (rawName || 'Obligacion')
      .replace(/[\\/*?:\[\]]/g, '-')
      .replace(/'/g, '')
      .trim()
      .slice(0, 31) || 'Obligacion';

    if (!usedNames.has(cleaned)) return cleaned;

    let counter = 2;
    let candidate = cleaned;
    do {
      const suffix = `_${counter}`;
      candidate = `${cleaned.slice(0, 31 - suffix.length)}${suffix}`;
      counter += 1;
    } while (usedNames.has(candidate));

    return candidate;
  }

  private normalizeHeader(value: string): string {
    return value
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .trim();
  }

  private cellToPlainText(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') {
      const cellObject = value as { text?: unknown; richText?: Array<{ text: string }>; result?: unknown };
      if ('text' in cellObject) return String(cellObject.text ?? '');
      if (Array.isArray(cellObject.richText)) return cellObject.richText.map((part) => part.text).join('');
      if ('result' in cellObject) return String(cellObject.result ?? '');
    }
    return String(value);
  }

  private getExtension(fileName: string): string {
    return fileName.split('.').pop()?.toLowerCase() ?? '';
  }

  private deepClone<T>(value: T): T {
    if (value === null || value === undefined) return value;
    if (value instanceof Date) return new Date(value.getTime()) as T;
    if (Array.isArray(value)) return value.map((item) => this.deepClone(item)) as T;
    if (typeof value === 'object') {
      return Object.fromEntries(Object.entries(value as Record<string, unknown>).map(([key, item]) => [key, this.deepClone(item)])) as T;
    }
    return value;
  }
}
