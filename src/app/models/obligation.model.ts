export interface ObligationRecord {
  obligationId: string;
  capital: number | string | Date | null;
  disbursementDate: string | Date | number | null;
  transferDate: string | Date | number | null;
  netDueDate: string | Date | number | null;
  totalDueDate: string | Date | number | null;
  businessRate: number | string | null;
  remuneratoryRate: number | string | null;
  dppRate: number | string | null;
  sourceRow: number;
}

export interface GenerationLog {
  type: 'success' | 'warning' | 'error' | 'info';
  message: string;
}

export interface GenerationOptions {
  divideRatesBy100: boolean;
  keepObligationsSheet: boolean;
  outputFileName: string;
}

export const CELL_MAP = {
  obligationId: 'A8',
  capital: 'B8',
  disbursementDate: 'B4',
  transferDate: 'C8',
  netDueDate: 'D8',
  totalDueDate: 'E8',
  businessRate: 'F8',
  remuneratoryRate: 'L8',
  dppRate: 'AB8'
} as const;

export type ObligationField = keyof typeof CELL_MAP;
