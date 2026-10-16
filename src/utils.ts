import { differenceInMinutes, format, isValid, parse } from 'date-fns';
import * as XLSX from 'xlsx';

export interface ProcessedRecord {
  id: string;
  admissao: Date | null;
  horaClassificacao: Date | null;
  atendimento: Date | null;
  fimAtendimento: Date | null;
  classificacao: string;
  medico: string;
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
    const date = XLSX.SSF.parse_date_code(dateVal);
    if (date) {
      return new Date(date.y, date.m - 1, date.d, date.H, date.M, date.S);
    }
  }

  // If it's a string, try various formats
  const cleanStr = String(dateVal).trim();
  if (!cleanStr) return null;

  // Formats shown in common spreadsheets and Brazilian/US formats
  const formatsToTry = [
    'M/d/yyyy h:mm:ss a',
    'MM/dd/yyyy h:mm:ss a',
    'd/M/yyyy h:mm:ss a',
    'dd/MM/yyyy h:mm:ss a',
    'M/d/yyyy H:mm:ss',
    'dd/MM/yyyy HH:mm:ss',
    'dd/MM/yyyy HH:mm',
    'dd-MM-yyyy HH:mm:ss',
    'dd-MM-yyyy HH:mm',
    'yyyy-MM-dd HH:mm:ss',
    'yyyy-MM-dd HH:mm',
    'MM/dd/yyyy HH:mm:ss',
    'MM/dd/yyyy HH:mm',
    'dd/MM/yyyy',
    'yyyy-MM-dd'
  ];

  for (const fmt of formatsToTry) {
    const parsedDate = parse(cleanStr, fmt, new Date());
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
  
  // Convert sheet to array of rows (2D array) to safely access Column F (index 5)
  const rows2D = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true, defval: '' }) as any[][];
  
  // Convert to array of objects with headers
  const rawObjects = XLSX.utils.sheet_to_json(worksheet, { raw: true, defval: '' }) as Record<string, any>[];
  
  const results: ProcessedRecord[] = [];

  for (let i = 0; i < rawObjects.length; i++) {
    const row = rawObjects[i];
    // Row 0 in rows2D is header, row i+1 corresponds to rawObjects[i]
    const rowArr = (rows2D && rows2D.length > i + 1) ? rows2D[i + 1] : [];

    // Helper to get value by case-insensitive key search
    const getVal = (keys: string[]) => {
      const lowerKeys = keys.map(k => k.toLowerCase().trim());
      for (const objKey of Object.keys(row)) {
        if (lowerKeys.includes(objKey.toLowerCase().trim())) {
          const val = row[objKey];
          if (val !== undefined && val !== null && String(val).trim() !== '') {
            return val;
          }
        }
      }
      return '';
    };

    // 1. Admissão
    let valAdmissao = getVal(['ADMISSÃO', 'ADMISSAO', 'Admissão', 'Admissao', 'admissão', 'admissao', 'DATA/HORA ADMISSAO', 'DATA ADMISSAO']);
    if (!valAdmissao && rowArr[0] !== undefined) {
      valAdmissao = rowArr[0];
    }

    // 2. Classificação Timestamp (Hora da Classificação)
    // IMPORTANT: User specified that the classification column may not have "hora" in the title
    // (e.g. just "CLASSIFICAÇÃO", "CLASSIFICACAO", "TRIAGEM", etc.)
    let valClass = getVal([
      'HORA DA CLASSIFICAÇÃO', 'HORA DA CLASSIFICACAO', 'Hora da Classificação', 
      'HORA CLASSIFICAÇÃO', 'HORA CLASSIFICACAO', 'HORA DA TRIAGEM', 'HORA TRIAGEM', 
      'TRIAGEM', 'HORA'
    ]);

    // If not found via specific "hora" headers, check generic "Classificação" header:
    // If its value is a date, it is the classification timestamp!
    if (!valClass) {
      const genericClassVal = getVal(['CLASSIFICAÇÃO', 'CLASSIFICACAO', 'Classificação', 'Classificacao']);
      if (genericClassVal && parseCustomDate(genericClassVal) !== null) {
        valClass = genericClassVal;
      } else if (rowArr[1] !== undefined && parseCustomDate(rowArr[1]) !== null) {
        valClass = rowArr[1];
      }
    }

    // 3. Atendimento
    let valAtend = getVal(['ATENDIMENTO', 'Atendimento', 'atendimento', 'HORA ATENDIMENTO', 'HORA DO ATENDIMENTO', 'INICIO ATENDIMENTO']);
    if (!valAtend && rowArr[2] !== undefined && parseCustomDate(rowArr[2]) !== null) {
      valAtend = rowArr[2];
    }

    // 4. Fim Atendimento
    let valFim = getVal(['FIM ATENDIMENTO', 'Fim Atendimento', 'fim atendimento', 'FIM DO ATENDIMENTO', 'TERMINO ATENDIMENTO', 'ALTA']);
    if (!valFim && rowArr[3] !== undefined && parseCustomDate(rowArr[3]) !== null) {
      valFim = rowArr[3];
    }

    // 5. Cor / Classificação de Risco (Manchester Color)
    let classificacao = getVal([
      'COR', 'Cor', 'cor', 'COR / RISCO', 'COR/RISCO', 'CLASSIFICAÇÃO DE RISCO', 
      'CLASSIFICACAO DE RISCO', 'PRIORIDADE', 'COR DA CLASSIFICAÇÃO', 'COR DA CLASSIFICACAO', 'RISCO'
    ]);

    // If still empty, check if row has a text-based "CLASSIFICAÇÃO" column that isn't a date
    if (!classificacao) {
      const genericClassVal = getVal(['CLASSIFICAÇÃO', 'CLASSIFICACAO', 'Classificação', 'Classificacao']);
      if (genericClassVal && parseCustomDate(genericClassVal) === null) {
        classificacao = genericClassVal;
      } else if (rowArr[4] !== undefined && parseCustomDate(rowArr[4]) === null && String(rowArr[4]).trim() !== '') {
        classificacao = rowArr[4];
      }
    }

    // 6. Médico / Profissional (User specified: column "profissional", name "médico", which is in column F / index 5)
    let valMedico = getVal([
      'PROFISSIONAL', 'Profissional', 'profissional', 
      'MÉDICO', 'Médico', 'médico', 'MEDICO', 'Medico', 'medico',
      'MÉDICA', 'Médica', 'médica', 'MEDICA', 'Medica',
      'PROFISSIONAL / MÉDICO', 'PROFISSIONAL/MÉDICO', 'PROFISSIONAL (MÉDICO)',
      'NOME DO PROFISSIONAL', 'NOME PROFISSIONAL', 'NOME DO MÉDICO', 'NOME MEDICO', 
      'DR', 'DRA', 'DOUTR', 'CLÍNICO'
    ]);

    // If not found via header, check Column F (index 5: A=0, B=1, C=2, D=3, E=4, F=5)
    if (!valMedico && rowArr && rowArr.length > 5 && rowArr[5] !== undefined) {
      valMedico = rowArr[5];
    }

    const medicoStr = String(valMedico || '').trim();

    const parsedAdmissao = parseCustomDate(valAdmissao);
    const parsedHoraClassificacao = parseCustomDate(valClass);
    const parsedAtendimento = parseCustomDate(valAtend);
    const parsedFimAtendimento = parseCustomDate(valFim);
    const classStr = String(classificacao || '').trim() || 'Sem Classificação';

    // Filter out completely empty rows
    if (parsedAdmissao || parsedHoraClassificacao || parsedAtendimento || parsedFimAtendimento || classStr !== 'Sem Classificação' || medicoStr) {
      results.push({
        id: `row-${i}-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
        admissao: parsedAdmissao,
        horaClassificacao: parsedHoraClassificacao,
        atendimento: parsedAtendimento,
        fimAtendimento: parsedFimAtendimento,
        classificacao: classStr,
        medico: medicoStr || 'Não informado',
        rawAdmissao: String(valAdmissao || ''),
        rawHoraClassificacao: String(valClass || ''),
        rawAtendimento: String(valAtend || ''),
        rawFimAtendimento: String(valFim || '')
      });
    }
  }
  
  return results;
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

// --- Aggregation Functions: Classificação / Cores ---

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

// --- Aggregation Functions: Médicos / Profissionais ---

export function getAppointmentsByMedico(data: ProcessedRecord[]) {
  const map = new Map<string, number>();
  data.forEach(d => {
    const med = d.medico?.trim() || 'Não informado';
    map.set(med, (map.get(med) || 0) + 1);
  });
  return Array.from(map.entries())
    .map(([name, atendimentos]) => ({
      name,
      atendimentos
    }))
    .sort((a, b) => b.atendimentos - a.atendimentos);
}

export function getAverageWaitTimeByMedico(data: ProcessedRecord[], startCol: string, endCol: string) {
  const map = new Map<string, { total: number, count: number }>();
  
  data.forEach(d => {
    const duration = getDurationMinutes(d, startCol, endCol);
    if (duration === null) return;
    
    const med = d.medico?.trim() || 'Não informado';
    if (!map.has(med)) map.set(med, { total: 0, count: 0 });
    
    const entry = map.get(med)!;
    entry.total += duration;
    entry.count += 1;
  });
  
  return Array.from(map.entries()).map(([name, stats]) => ({
    name,
    mediaEspera: Number((stats.total / stats.count).toFixed(1)),
    count: stats.count
  })).sort((a, b) => b.mediaEspera - a.mediaEspera);
}

export function getDoctorRankingSummary(data: ProcessedRecord[], startCol: string, endCol: string) {
  const appointments = getAppointmentsByMedico(data);
  const waitTimes = getAverageWaitTimeByMedico(data, startCol, endCol);
  const waitTimeMap = new Map(waitTimes.map(w => [w.name, w.mediaEspera]));

  return appointments.map(item => ({
    name: item.name,
    atendimentos: item.atendimentos,
    mediaEspera: waitTimeMap.has(item.name) ? waitTimeMap.get(item.name)! : null
  }));
}

// --- Aggregation Functions: Volume Diário ---

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
