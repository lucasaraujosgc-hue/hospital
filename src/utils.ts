import { differenceInMinutes, format, isValid, parse } from 'date-fns';
import * as XLSX from 'xlsx';

export interface ProcessedRecord {
  id: string;
  admissao: Date | null;
  horaClassificacao: Date | null;
  atendimento: Date | null;
  fimAtendimento: Date | null;
  classificacao: string;
  // Store raw strings for display if needed
  rawAdmissao: string;
  rawHoraClassificacao: string;
  rawAtendimento: string;
  rawFimAtendimento: string;
}

export const columnOptions = [
  { id: 'admissao', label: 'Admissão' },
  { id: 'horaClassificacao', label: 'Hora da Classificação' },
  { id: 'atendimento', label: 'Atendimento' },
  { id: 'fimAtendimento', label: 'Fim Atendimento' }
];

export function parseCustomDate(dateVal: any): Date | null {
  if (!dateVal) return null;
  
  // If it's already a JS Date object (Excel parsed it)
  if (dateVal instanceof Date && isValid(dateVal)) {
    return dateVal;
  }

  // If it's a number (Excel date code)
  if (typeof dateVal === 'number') {
    // Excel dates are days since Dec 30, 1899
    // SheetJS often converts these to strings if raw:false, but let's handle if it is a number
    const date = XLSX.SSF.parse_date_code(dateVal);
    if (date) {
      return new Date(date.y, date.m - 1, date.d, date.H, date.M, date.S);
    }
  }

  // If it's a string, try various formats
  const cleanStr = String(dateVal).trim();
  let parsedDate = null;
  
  // Formats shown in screenshot: 8/1/2026 1:46:58 AM or PM
  const formatsToTry = [
    'M/d/yyyy h:mm:ss a',
    'MM/dd/yyyy h:mm:ss a',
    'd/M/yyyy h:mm:ss a',
    'dd/MM/yyyy h:mm:ss a',
    'dd/MM/yyyy HH:mm:ss',
    'dd/MM/yyyy HH:mm',
    'yyyy-MM-dd HH:mm:ss',
    'yyyy-MM-dd HH:mm',
    'MM/dd/yyyy HH:mm',
  ];

  for (const fmt of formatsToTry) {
    parsedDate = parse(cleanStr, fmt, new Date());
    if (isValid(parsedDate)) {
      return parsedDate;
    }
  }
  
  // Last resort, standard new Date
  const d = new Date(cleanStr);
  if (isValid(d)) return d;

  return null;
}

export function processExcelData(arrayBuffer: ArrayBuffer): ProcessedRecord[] {
  const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  
  // Convert to array of objects.
  const rawData = XLSX.utils.sheet_to_json(worksheet, { raw: true, defval: '' }) as any[];
  
  return rawData.map((row, index) => {
    // Handle potential variations in column names from Excel
    const getVal = (keys: string[]) => {
      for (const key of keys) {
        if (row[key] !== undefined && row[key] !== '') return row[key];
      }
      return '';
    };

    const valAdmissao = getVal(['ADMISSÃO', 'ADMISSAO', 'Admissão', 'admissão']);
    const valClass = getVal(['HORA DA CLASSIFICAÇÃO', 'HORA DA CLASSIFICACAO', 'Hora da Classificação']);
    const valAtend = getVal(['ATENDIMENTO', 'Atendimento', 'atendimento']);
    const valFim = getVal(['FIM ATENDIMENTO', 'Fim Atendimento', 'fim atendimento']);
    const classificacao = getVal(['CLASSIFICACAO', 'CLASSIFICAÇÃO', 'Classificacao', 'Classificação', 'Cor']);

    return {
      id: `row-${index}-${Date.now()}`,
      admissao: parseCustomDate(valAdmissao),
      horaClassificacao: parseCustomDate(valClass),
      atendimento: parseCustomDate(valAtend),
      fimAtendimento: parseCustomDate(valFim),
      classificacao: String(classificacao).trim() || 'Sem Classificação',
      rawAdmissao: String(valAdmissao),
      rawHoraClassificacao: String(valClass),
      rawAtendimento: String(valAtend),
      rawFimAtendimento: String(valFim)
    };
  }).filter(r => r.admissao || r.horaClassificacao || r.atendimento || r.fimAtendimento || r.classificacao !== 'Sem Classificação');
}

export function getDurationMinutes(record: ProcessedRecord, startCol: string, endCol: string): number | null {
  const start = record[startCol as keyof ProcessedRecord] as Date | null;
  const end = record[endCol as keyof ProcessedRecord] as Date | null;
  if (start && end) {
    const diff = differenceInMinutes(end, start);
    return diff >= 0 ? diff : null; // Ignore negative times if any
  }
  return null;
}

// --- Aggregation Functions ---

export function getAverageWaitTimeByClassificacao(data: ProcessedRecord[], startCol: string, endCol: string) {
  const map = new Map<string, { total: number, count: number }>();
  
  data.forEach(d => {
    const duration = getDurationMinutes(d, startCol, endCol);
    if (duration === null) return;
    
    const color = d.classificacao;
    if (!map.has(color)) map.set(color, { total: 0, count: 0 });
    
    const entry = map.get(color)!;
    entry.total += duration;
    entry.count += 1;
  });
  
  return Array.from(map.entries()).map(([name, stats]) => ({
    name,
    mediaEspera: Number((stats.total / stats.count).toFixed(1))
  })).sort((a, b) => b.mediaEspera - a.mediaEspera);
}

export function getAppointmentsByClassificacao(data: ProcessedRecord[]) {
  const map = new Map<string, number>();
  data.forEach(d => {
    map.set(d.classificacao, (map.get(d.classificacao) || 0) + 1);
  });
  return Array.from(map.entries()).map(([name, atendimentos]) => ({
    name,
    atendimentos
  })).sort((a, b) => b.atendimentos - a.atendimentos);
}

export function getWaitTimeTimeline(data: ProcessedRecord[], startCol: string, endCol: string) {
  // Aggregate by Date (Day/Month/Year) of the start column
  const map = new Map<string, { total: number, count: number, _date: Date }>();
  
  data.forEach(d => {
    const startDate = d[startCol as keyof ProcessedRecord] as Date | null;
    if (!startDate) return;
    
    const duration = getDurationMinutes(d, startCol, endCol);
    if (duration === null) return;

    const dateKey = format(startDate, 'dd/MM/yyyy');
    
    if (!map.has(dateKey)) {
      const normalizedDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      map.set(dateKey, { total: 0, count: 0, _date: normalizedDate });
    }
    const entry = map.get(dateKey)!;
    entry.total += duration;
    entry.count += 1;
  });

  return Array.from(map.entries()).map(([date, stats]) => ({
    name: date,
    mediaEspera: Number((stats.total / stats.count).toFixed(1)),
    volume: stats.count,
    _sort: stats._date.getTime()
  })).sort((a, b) => a._sort - b._sort);
}

export function getTimelineVolumeByClassificacao(data: ProcessedRecord[], startCol: string) {
  // Aggregate volume by day and classification color
  const map = new Map<string, { _date: Date, [key: string]: any }>();
  
  data.forEach(d => {
    const startDate = d[startCol as keyof ProcessedRecord] as Date | null;
    if (!startDate) return;
    
    const dateKey = format(startDate, 'dd/MM/yyyy');
    if (!map.has(dateKey)) {
      const normalizedDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      map.set(dateKey, { name: dateKey, _date: normalizedDate });
    }
    const entry = map.get(dateKey)!;
    
    const color = d.classificacao;
    entry[color] = (entry[color] || 0) + 1;
  });

  return Array.from(map.values()).sort((a, b) => a._date.getTime() - b._date.getTime());
}

export function getDailyVolumeSummary(data: ProcessedRecord[], startCol: string) {
  const map = new Map<string, { total: number, _date: Date }>();
  
  data.forEach(d => {
    const startDate = d[startCol as keyof ProcessedRecord] as Date | null;
    if (!startDate) return;
    
    const dateKey = format(startDate, 'dd/MM/yyyy');
    if (!map.has(dateKey)) {
      const normalizedDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
      map.set(dateKey, { total: 0, _date: normalizedDate });
    }
    map.get(dateKey)!.total += 1;
  });

  return Array.from(map.entries())
    .map(([date, stats]) => ({ date, total: stats.total, _sort: stats._date.getTime() }))
    .sort((a, b) => a._sort - b._sort);
}

