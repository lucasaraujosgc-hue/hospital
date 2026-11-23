import React, { useState, useMemo, useEffect } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  Cell
} from 'recharts';
import { 
  UploadCloud, FileSpreadsheet, Loader2, Filter, Calendar, Trash2, Settings, Download, 
  Edit2, Save, X, Search, Clock, UserCheck, Stethoscope, RotateCcw,
  ArrowUpDown, ArrowUp, ArrowDown, AlertCircle
} from 'lucide-react';
import { startOfDay, endOfDay, parseISO, format } from 'date-fns';
import * as XLSX from 'xlsx';
import { 
  ProcessedRecord, processExcelData, columnOptions, getAverageWaitTimeByClassificacao, 
  getAppointmentsByClassificacao, getDailyVolumeSummary, parseCustomDate,
  getDoctorRankingSummary, getAppointmentsByMedico, isValidDoctor
} from './utils';

// Tailwind colors for common classifications
const colorMap: Record<string, string> = {
  'VERMELHO': '#ef4444',
  'Vermelho': '#ef4444',
  'LARANJA': '#f97316',
  'Laranja': '#f97316',
  'AMARELO': '#eab308',
  'Amarelo': '#eab308',
  'VERDE': '#22c55e',
  'Verde': '#22c55e',
  'AZUL': '#3b82f6',
  'Azul': '#3b82f6',
  'BRANCO': '#94a3b8',
  'Branco': '#94a3b8',
  'Sem Classificação': '#64748b'
};

const DEFAULT_COLORS = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#06b6d4'];

export function App() {
  const [data, setData] = useState<ProcessedRecord[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'data'>('dashboard');

  // Interval column configuration
  const [startCol, setStartCol] = useState<string>('admissao');
  const [endCol, setEndCol] = useState<string>('horaClassificacao');

  // Time & Duration filter state
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [minWaitTime, setMinWaitTime] = useState<string>('');
  const [maxWaitTime, setMaxWaitTime] = useState<string>('');

  // Classification & Doctor Filters
  const [selectedColors, setSelectedColors] = useState<Set<string>>(new Set());
  const [selectedMedicos, setSelectedMedicos] = useState<Set<string>>(new Set());
  const [medicoSearchTerm, setMedicoSearchTerm] = useState<string>('');

  // Raw Data Search Filter
  const [searchTerm, setSearchTerm] = useState<string>('');

  // Missing / Unfilled Columns Filter
  const [missingColsFilter, setMissingColsFilter] = useState<Set<string>>(new Set());
  const [missingColsMode, setMissingColsMode] = useState<'all' | 'any'>('all');

  // Doctor Ranking Table Sorting state
  const [doctorSortField, setDoctorSortField] = useState<'rank' | 'name' | 'atendimentos' | 'mediaEspera'>('rank');
  const [doctorSortOrder, setDoctorSortOrder] = useState<'asc' | 'desc'>('asc');

  // Bulk Delete Modal state
  const [isDeleteModalOpen, setIsDeleteModalOpen] = useState(false);
  const [deleteStart, setDeleteStart] = useState<string>('');
  const [deleteEnd, setDeleteEnd] = useState<string>('');

  // Inline editing state for Raw Data
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<Partial<ProcessedRecord>>({});

  // --- API Functions ---
  const fetchRecords = async () => {
    try {
      setIsLoading(true);
      const res = await fetch('/api/records');
      const json = await res.json();
      
      // Re-hydrate dates
      const hydrated = json.map((r: any) => ({
        ...r,
        admissao: r.admissao ? new Date(r.admissao) : null,
        horaClassificacao: r.horaClassificacao ? new Date(r.horaClassificacao) : null,
        atendimento: r.atendimento ? new Date(r.atendimento) : null,
        fimAtendimento: r.fimAtendimento ? new Date(r.fimAtendimento) : null,
        medico: r.medico || 'Não informado',
      }));
      
      setData(hydrated);
    } catch (error) {
      console.error('Failed to fetch records:', error);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchRecords();
  }, []);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        const processed = processExcelData(arrayBuffer);
        
        // Save to DB
        const res = await fetch('/api/records/bulk', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ records: processed })
        });
        
        if (res.ok) {
          // Append new records to current state
          setData(prev => [...prev, ...processed]);
          setSelectedColors(new Set());
          setSelectedMedicos(new Set());
          setMissingColsFilter(new Set());
          setStartDate('');
          setEndDate('');
          alert(`Importação concluída! ${processed.length} registros adicionados com sucesso.`);
        } else {
          throw new Error('Falha na API ao salvar');
        }
      } catch (error) {
        console.error("Error parsing Excel:", error);
        alert("Erro ao processar o arquivo Excel e salvar no banco de dados.");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // reset input
  };

  const uniqueColors = useMemo(() => Array.from(new Set(data.map(d => d.classificacao).filter(Boolean))).sort(), [data]);
  
  const uniqueMedicos = useMemo(() => {
    const list = Array.from(new Set(data.map(d => d.medico?.trim()).filter(med => isValidDoctor(med)) as string[])).sort();
    return list;
  }, [data]);

  const filteredData = useMemo(() => {
    let sDate: Date | null = null;
    let eDate: Date | null = null;

    if (startDate) sDate = startOfDay(parseISO(startDate));
    if (endDate) eDate = endOfDay(parseISO(endDate));
    
    const searchLower = searchTerm.toLowerCase();
    const minTimeNum = minWaitTime !== '' ? Number(minWaitTime) : null;
    const maxTimeNum = maxWaitTime !== '' ? Number(maxWaitTime) : null;

    return data.filter(d => {
      // Time Filter based on startCol
      const recordDate = d[startCol as keyof ProcessedRecord] as Date | null;
      if (sDate && recordDate && recordDate < sDate) return false;
      if (eDate && recordDate && recordDate > eDate) return false;

      // Color Filter
      if (selectedColors.size > 0 && !selectedColors.has(d.classificacao)) return false;

      // Doctor Filter
      if (selectedMedicos.size > 0) {
        const med = d.medico || 'Não informado';
        if (!selectedMedicos.has(med)) return false;
      }
      
      // Duration Filter
      if (minTimeNum !== null || maxTimeNum !== null) {
        const startT = d[startCol as keyof ProcessedRecord] as Date | null;
        const endT = d[endCol as keyof ProcessedRecord] as Date | null;
        
        if (!startT || !endT) return false;
        
        const diffMinutes = (endT.getTime() - startT.getTime()) / (1000 * 60);
        if (minTimeNum !== null && diffMinutes < minTimeNum) return false;
        if (maxTimeNum !== null && diffMinutes > maxTimeNum) return false;
      }

      // Text Search Filter (applies to Doctor, Color, and Dates)
      if (searchLower) {
        const matchStr = `${d.classificacao} ${d.medico || ''} ${d.rawAdmissao} ${d.rawHoraClassificacao} ${d.rawAtendimento} ${d.rawFimAtendimento}`.toLowerCase();
        if (!matchStr.includes(searchLower)) return false;
      }

      // Missing / Unfilled Columns Filter
      if (missingColsFilter.size > 0) {
        const isEmptyVal = (val?: string | null) => {
          if (!val) return true;
          const trimmed = val.trim().toLowerCase();
          return trimmed === '' || trimmed === '-' || trimmed === '--' || trimmed === 'null' || trimmed === 'undefined';
        };
        const isColEmpty = (colId: string): boolean => {
          switch (colId) {
            case 'admissao':
              return !d.admissao && isEmptyVal(d.rawAdmissao);
            case 'horaClassificacao':
              return !d.horaClassificacao && isEmptyVal(d.rawHoraClassificacao);
            case 'atendimento':
              return !d.atendimento && isEmptyVal(d.rawAtendimento);
            case 'fimAtendimento':
              return !d.fimAtendimento && isEmptyVal(d.rawFimAtendimento);
            case 'classificacao':
              return isEmptyVal(d.classificacao) || d.classificacao === 'Sem Classificação';
            case 'medico':
              return !isValidDoctor(d.medico);
            default:
              return false;
          }
        };

        const selectedCols = Array.from(missingColsFilter) as string[];
        if (missingColsMode === 'all') {
          if (!selectedCols.every(colId => isColEmpty(colId))) return false;
        } else {
          if (!selectedCols.some(colId => isColEmpty(colId))) return false;
        }
      }
      
      return true;
    });
  }, [data, selectedColors, selectedMedicos, startDate, endDate, startCol, endCol, searchTerm, minWaitTime, maxWaitTime, missingColsFilter, missingColsMode]);

  const hasData = data.length > 0;
  
  // Dashboard Aggregations
  const waitTimeData = hasData ? getAverageWaitTimeByClassificacao(filteredData, startCol, endCol) : [];
  const appointmentsData = hasData ? getAppointmentsByClassificacao(filteredData) : [];
  const dailyVolumeData = hasData ? getDailyVolumeSummary(filteredData, startCol) : [];
  const doctorAppointmentsData = hasData ? getAppointmentsByMedico(filteredData) : [];
  
  // Doctor Ranking with Original Position
  const baseDoctorRankingData = useMemo(() => {
    if (!hasData) return [];
    const list = getDoctorRankingSummary(filteredData, startCol, endCol);
    return list.map((item, idx) => ({
      ...item,
      originalRank: idx + 1
    }));
  }, [hasData, filteredData, startCol, endCol]);

  // Sorted Doctor Ranking based on user selection
  const sortedDoctorRankingData = useMemo(() => {
    return [...baseDoctorRankingData].sort((a, b) => {
      let comparison = 0;
      if (doctorSortField === 'rank') {
        comparison = a.originalRank - b.originalRank;
      } else if (doctorSortField === 'name') {
        comparison = a.name.localeCompare(b.name, 'pt-BR', { sensitivity: 'base' });
      } else if (doctorSortField === 'atendimentos') {
        comparison = a.atendimentos - b.atendimentos;
      } else if (doctorSortField === 'mediaEspera') {
        const valA = a.mediaEspera !== null ? a.mediaEspera : (doctorSortOrder === 'asc' ? Infinity : -Infinity);
        const valB = b.mediaEspera !== null ? b.mediaEspera : (doctorSortOrder === 'asc' ? Infinity : -Infinity);
        comparison = valA - valB;
      }
      return doctorSortOrder === 'asc' ? comparison : -comparison;
    });
  }, [baseDoctorRankingData, doctorSortField, doctorSortOrder]);

  const handleDoctorSort = (field: 'rank' | 'name' | 'atendimentos' | 'mediaEspera') => {
    if (doctorSortField === field) {
      setDoctorSortOrder(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setDoctorSortField(field);
      setDoctorSortOrder(field === 'name' || field === 'rank' ? 'asc' : 'desc');
    }
  };

  const handlePositionOrNameSort = () => {
    if (doctorSortField === 'name') {
      if (doctorSortOrder === 'asc') {
        setDoctorSortOrder('desc');
      } else {
        setDoctorSortField('rank');
        setDoctorSortOrder('asc');
      }
    } else if (doctorSortField === 'rank') {
      if (doctorSortOrder === 'asc') {
        setDoctorSortOrder('desc');
      } else {
        setDoctorSortField('name');
        setDoctorSortOrder('asc');
      }
    } else {
      setDoctorSortField('rank');
      setDoctorSortOrder('asc');
    }
  };

  const toggleFilter = (setter: React.Dispatch<React.SetStateAction<Set<string>>>, value: string) => {
    setter(prev => {
      const next = new Set(prev);
      if (next.has(value)) next.delete(value);
      else next.add(value);
      return next;
    });
  };

  const clearAllFilters = () => {
    setSelectedColors(new Set());
    setSelectedMedicos(new Set());
    setMissingColsFilter(new Set());
    setStartDate('');
    setEndDate('');
    setMinWaitTime('');
    setMaxWaitTime('');
    setSearchTerm('');
  };

  const hasActiveFilters = selectedColors.size > 0 || selectedMedicos.size > 0 || missingColsFilter.size > 0 || startDate !== '' || endDate !== '' || minWaitTime !== '' || maxWaitTime !== '' || searchTerm !== '';

  const emptyColumnOptions = [
    { id: 'admissao', label: 'Admissão' },
    { id: 'horaClassificacao', label: 'Classificação (Hora)' },
    { id: 'atendimento', label: 'Atendimento' },
    { id: 'fimAtendimento', label: 'Fim Atendimento' },
    { id: 'classificacao', label: 'Cor' },
    { id: 'medico', label: 'Médico / Profissional' },
  ];

  const getColorHex = (name: string, index: number) => {
    return colorMap[name] || DEFAULT_COLORS[index % DEFAULT_COLORS.length];
  };

  // --- Row Actions ---
  const handleEditClick = (record: ProcessedRecord) => {
    setEditingId(record.id);
    setEditForm({
      rawAdmissao: record.admissao ? format(record.admissao, 'dd/MM/yyyy HH:mm:ss') : record.rawAdmissao,
      rawHoraClassificacao: record.horaClassificacao ? format(record.horaClassificacao, 'dd/MM/yyyy HH:mm:ss') : record.rawHoraClassificacao,
      rawAtendimento: record.atendimento ? format(record.atendimento, 'dd/MM/yyyy HH:mm:ss') : record.rawAtendimento,
      rawFimAtendimento: record.fimAtendimento ? format(record.fimAtendimento, 'dd/MM/yyyy HH:mm:ss') : record.rawFimAtendimento,
      classificacao: record.classificacao,
      medico: record.medico || 'Não informado',
    });
  };

  const handleSaveEdit = async (id: string) => {
    const admissao = parseCustomDate(editForm.rawAdmissao);
    const horaClassificacao = parseCustomDate(editForm.rawHoraClassificacao);
    const atendimento = parseCustomDate(editForm.rawAtendimento);
    const fimAtendimento = parseCustomDate(editForm.rawFimAtendimento);
    
    const updatedRecord = {
      admissao,
      horaClassificacao,
      atendimento,
      fimAtendimento,
      classificacao: editForm.classificacao || 'Sem Classificação',
      medico: editForm.medico?.trim() || 'Não informado',
      rawAdmissao: String(editForm.rawAdmissao || ''),
      rawHoraClassificacao: String(editForm.rawHoraClassificacao || ''),
      rawAtendimento: String(editForm.rawAtendimento || ''),
      rawFimAtendimento: String(editForm.rawFimAtendimento || ''),
    };

    try {
      const res = await fetch(`/api/records/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updatedRecord)
      });

      if (res.ok) {
        setData(prev => prev.map(item => item.id === id ? { ...item, ...updatedRecord } as ProcessedRecord : item));
        setEditingId(null);
      } else {
        throw new Error('Falha ao atualizar');
      }
    } catch (error) {
      console.error(error);
      alert('Erro ao salvar edição no banco de dados.');
    }
  };

  const handleDeleteRow = async (id: string) => {
    if (confirm("Tem certeza que deseja excluir esta linha?")) {
      try {
        const res = await fetch(`/api/records/${id}`, { method: 'DELETE' });
        if (res.ok) {
          setData(prev => prev.filter(d => d.id !== id));
        } else {
          throw new Error('Falha ao excluir');
        }
      } catch (error) {
        console.error(error);
        alert('Erro ao excluir do banco de dados.');
      }
    }
  };

  const handleBulkDelete = async () => {
    const sDate = deleteStart ? startOfDay(parseISO(deleteStart)) : null;
    const eDate = deleteEnd ? endOfDay(parseISO(deleteEnd)) : null;
    
    if (!sDate && !eDate) {
      alert("Por favor, selecione ao menos uma data (Inicial ou Final) para exclusão.");
      return;
    }

    const idsToDelete = data.filter(d => {
      const recordDate = d[startCol as keyof ProcessedRecord] as Date | null;
      if (!recordDate) return false; 
      let isWithin = true;
      if (sDate && recordDate < sDate) isWithin = false;
      if (eDate && recordDate > eDate) isWithin = false;
      return isWithin;
    }).map(d => d.id);

    if (idsToDelete.length === 0) {
      alert("Nenhum registro encontrado no período selecionado.");
      return;
    }

    try {
      const res = await fetch('/api/records/bulk-delete', {
        method: 'DELETE',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids: idsToDelete })
      });

      if (res.ok) {
        setData(prev => prev.filter(d => !idsToDelete.includes(d.id)));
        setIsDeleteModalOpen(false);
        setDeleteStart('');
        setDeleteEnd('');
        alert(`${idsToDelete.length} registros excluídos com sucesso.`);
      } else {
        throw new Error('Falha na exclusão em lote');
      }
    } catch (error) {
      console.error(error);
      alert('Erro ao excluir registros em lote.');
    }
  };

  const handleExport = () => {
    if (filteredData.length === 0) {
      alert("Não há dados para exportar.");
      return;
    }
    
    const exportData = filteredData.map(d => ({
      'Admissão': d.admissao ? format(d.admissao, 'dd/MM/yyyy HH:mm:ss') : d.rawAdmissao,
      'Hora da Classificação': d.horaClassificacao ? format(d.horaClassificacao, 'dd/MM/yyyy HH:mm:ss') : d.rawHoraClassificacao,
      'Atendimento': d.atendimento ? format(d.atendimento, 'dd/MM/yyyy HH:mm:ss') : d.rawAtendimento,
      'Fim Atendimento': d.fimAtendimento ? format(d.fimAtendimento, 'dd/MM/yyyy HH:mm:ss') : d.rawFimAtendimento,
      'Classificação': d.classificacao,
      'Médico / Profissional': d.medico || 'Não informado',
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dados_Brutos");
    XLSX.writeFile(wb, `Painel_Classificacao_${format(new Date(), 'yyyyMMdd_HHmm')}.xlsx`);
  };

  if (isLoading) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center">
        <Loader2 className="w-8 h-8 text-teal-600 animate-spin mb-4" />
        <p className="text-slate-500 font-medium">Carregando dados do banco de dados persistente...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 p-4 sm:p-8 font-sans relative">
      
      {/* Bulk Delete Modal */}
      {isDeleteModalOpen && (
        <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-lg max-w-md w-full p-6 animate-in zoom-in-95 duration-200">
            <h3 className="text-lg font-bold text-slate-800 mb-2">Excluir em Lote</h3>
            <p className="text-sm text-slate-500 mb-4">
              Selecione o período (baseado na coluna "{columnOptions.find(c => c.id === startCol)?.label}") para excluir os registros permanentemente do banco de dados.
            </p>
            <div className="grid grid-cols-2 gap-4 mb-6">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">A partir de</label>
                <input 
                  type="date" 
                  value={deleteStart} 
                  onChange={e => setDeleteStart(e.target.value)} 
                  className="w-full border border-slate-300 rounded-lg p-2 text-sm focus:border-teal-500 focus:ring-1 focus:ring-teal-500" 
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">Até</label>
                <input 
                  type="date" 
                  value={deleteEnd} 
                  onChange={e => setDeleteEnd(e.target.value)} 
                  className="w-full border border-slate-300 rounded-lg p-2 text-sm focus:border-teal-500 focus:ring-1 focus:ring-teal-500" 
                />
              </div>
            </div>
            <div className="flex justify-end gap-3">
              <button 
                onClick={() => setIsDeleteModalOpen(false)}
                className="px-4 py-2 border border-slate-200 text-slate-600 hover:bg-slate-50 rounded-lg text-sm font-medium transition-colors"
              >
                Cancelar
              </button>
              <button 
                onClick={handleBulkDelete}
                className="px-4 py-2 bg-rose-600 text-white hover:bg-rose-700 rounded-lg text-sm font-medium transition-colors"
              >
                Confirmar Exclusão
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
          <div>
            <div className="flex items-center gap-3">
              <h1 className="text-2xl font-bold text-slate-800 tracking-tight">Indicadores do Hospital</h1>
              <span className="text-xs font-semibold px-2.5 py-0.5 rounded-full bg-teal-100 text-teal-800 border border-teal-200">
                SQLite Persistente
              </span>
            </div>
            <p className="text-sm text-slate-500 mt-1">
              Análise de tempo de espera, produtividade por médico e exportação de dados
            </p>
          </div>

          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 px-4 py-2.5 bg-teal-600 text-white hover:bg-teal-700 rounded-xl cursor-pointer font-medium text-sm transition-all shadow-sm shadow-teal-700/20 active:scale-95">
              <UploadCloud className="w-4 h-4" />
              <span>{hasData ? "Adicionar XLSX" : "Carregar Planilha XLSX"}</span>
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                onChange={handleFileUpload} 
                className="hidden" 
                disabled={isProcessing}
              />
            </label>
          </div>
        </header>

        {hasData && (
          <div className="space-y-6">
            {/* Global Filters & Configuration */}
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
              
              {/* Card 1: Configuração de Tempos */}
              <div className="bg-white p-4 border border-slate-200 rounded-xl shadow-sm space-y-3">
                <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                  <Settings className="w-4 h-4 text-slate-400" /> Configuração de Tempos
                </h3>
                <div>
                  <label className="block text-[11px] font-medium text-slate-500 mb-1">Início do Intervalo</label>
                  <select 
                    value={startCol}
                    onChange={e => setStartCol(e.target.value)}
                    className="w-full border border-slate-300 rounded-md p-1.5 text-xs text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  >
                    {columnOptions.map(opt => <option key={opt.id} value={opt.id}>{opt.label}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-slate-500 mb-1">Fim do Intervalo</label>
                  <select 
                    value={endCol}
                    onChange={e => setEndCol(e.target.value)}
                    className="w-full border border-slate-300 rounded-md p-1.5 text-xs text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  >
                    {columnOptions.map(opt => <option key={opt.id} value={opt.id}>{opt.label}</option>)}
                  </select>
                </div>
              </div>

              {/* Card 2: Date Range */}
              <div className="bg-white p-4 border border-slate-200 rounded-xl shadow-sm space-y-3">
                <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-slate-400" /> Período
                </h3>
                <div>
                  <label className="block text-[11px] font-medium text-slate-500 mb-1">Data Inicial</label>
                  <input 
                    type="date" 
                    value={startDate} 
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full border border-slate-300 rounded-md p-1.5 text-xs text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-slate-500 mb-1">Data Final</label>
                  <input 
                    type="date" 
                    value={endDate} 
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full border border-slate-300 rounded-md p-1.5 text-xs text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  />
                </div>
              </div>

              {/* Card 3: Duration Filter */}
              <div className="bg-white p-4 border border-slate-200 rounded-xl shadow-sm space-y-3">
                <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                  <Clock className="w-4 h-4 text-slate-400" /> Duração (minutos)
                </h3>
                <div>
                  <label className="block text-[11px] font-medium text-slate-500 mb-1">Tempo Mínimo (Ex: 240)</label>
                  <input 
                    type="number" 
                    min="0"
                    value={minWaitTime} 
                    onChange={(e) => setMinWaitTime(e.target.value)}
                    placeholder="Mínimo em min"
                    className="w-full border border-slate-300 rounded-md p-1.5 text-xs text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-medium text-slate-500 mb-1">Tempo Máximo</label>
                  <input 
                    type="number" 
                    min="0"
                    value={maxWaitTime} 
                    onChange={(e) => setMaxWaitTime(e.target.value)}
                    placeholder="Máximo em min"
                    className="w-full border border-slate-300 rounded-md p-1.5 text-xs text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  />
                </div>
              </div>

              {/* Card 4: Colors Filter */}
              <div className="bg-white p-4 border border-slate-200 rounded-xl shadow-sm space-y-2 flex flex-col">
                <div className="flex items-center justify-between">
                  <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                    <Filter className="w-4 h-4 text-slate-400" /> Cores
                  </h3>
                  {selectedColors.size > 0 && (
                    <button 
                      onClick={() => setSelectedColors(new Set())}
                      className="text-[11px] text-teal-600 hover:underline"
                    >
                      Limpar
                    </button>
                  )}
                </div>
                <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto flex-1 content-start">
                  {uniqueColors.map(color => (
                    <button
                      key={color}
                      onClick={() => toggleFilter(setSelectedColors, color)}
                      className={`px-2 py-1 text-xs font-medium rounded-full border transition-colors ${
                        selectedColors.has(color) 
                          ? 'bg-teal-50 border-teal-500 text-teal-700 shadow-sm' 
                          : 'bg-white border-slate-200 text-slate-600 hover:border-slate-300'
                      }`}
                    >
                      <span 
                        className="inline-block w-2 h-2 rounded-full mr-1.5" 
                        style={{backgroundColor: getColorHex(color, 0)}} 
                      />
                      {color}
                    </button>
                  ))}
                </div>
              </div>

              {/* Card 5: Doctor Filter */}
              <div className="bg-white p-4 border border-slate-200 rounded-xl shadow-sm space-y-2 flex flex-col">
                <div className="flex items-center justify-between">
                  <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                    <Stethoscope className="w-4 h-4 text-teal-600" /> Médico / Profissional
                  </h3>
                  {selectedMedicos.size > 0 && (
                    <button 
                      onClick={() => setSelectedMedicos(new Set())}
                      className="text-[11px] text-teal-600 hover:underline"
                    >
                      Limpar
                    </button>
                  )}
                </div>
                <div className="relative">
                  <input 
                    type="text"
                    value={medicoSearchTerm}
                    onChange={e => setMedicoSearchTerm(e.target.value)}
                    placeholder="Buscar médico..."
                    className="w-full text-xs p-1 border border-slate-200 rounded text-slate-700 outline-none focus:border-teal-500"
                  />
                </div>
                <div className="flex flex-wrap gap-1.5 max-h-20 overflow-y-auto flex-1 content-start">
                  {uniqueMedicos
                    .filter(med => med.toLowerCase().includes(medicoSearchTerm.toLowerCase()))
                    .map(medico => (
                      <button
                        key={medico}
                        onClick={() => toggleFilter(setSelectedMedicos, medico)}
                        className={`px-2 py-1 text-xs font-medium rounded-full border transition-colors truncate max-w-[150px] ${
                          selectedMedicos.has(medico) 
                            ? 'bg-teal-50 border-teal-500 text-teal-700 shadow-sm font-semibold' 
                            : 'bg-white border-slate-200 text-slate-600 hover:border-slate-300'
                        }`}
                        title={medico}
                      >
                        {medico}
                      </button>
                  ))}
                  {uniqueMedicos.length === 0 && (
                    <span className="text-xs text-slate-400 italic">Nenhum médico cadastrado</span>
                  )}
                </div>
              </div>

            </div>

            {/* Card 6: Missing / Unfilled Columns Filter */}
            <div className="bg-white p-4 border border-slate-200 rounded-xl shadow-sm">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-3">
                <div className="flex items-center gap-2">
                  <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                    <AlertCircle className="w-4 h-4 text-amber-600" /> Colunas Sem Preenchimento
                  </h3>
                  <span className="text-xs text-slate-400 hidden sm:inline">
                    (Selecione uma ou mais colunas vazias para levantar os dados)
                  </span>
                </div>
                <div className="flex items-center gap-3">
                  {missingColsFilter.size > 1 && (
                    <div className="inline-flex items-center bg-slate-100 p-0.5 rounded-lg border border-slate-200 text-[11px]">
                      <button
                        type="button"
                        onClick={() => setMissingColsMode('all')}
                        className={`px-2 py-0.5 rounded-md font-medium transition-colors ${
                          missingColsMode === 'all'
                            ? 'bg-white text-teal-700 shadow-sm'
                            : 'text-slate-500 hover:text-slate-700'
                        }`}
                        title="Exibir registros em que TODAS as colunas selecionadas estão sem preenchimento"
                      >
                        Todas vazias (E)
                      </button>
                      <button
                        type="button"
                        onClick={() => setMissingColsMode('any')}
                        className={`px-2 py-0.5 rounded-md font-medium transition-colors ${
                          missingColsMode === 'any'
                            ? 'bg-white text-teal-700 shadow-sm'
                            : 'text-slate-500 hover:text-slate-700'
                        }`}
                        title="Exibir registros em que QUALQUER UMA das colunas selecionadas está sem preenchimento"
                      >
                        Qualquer uma (OU)
                      </button>
                    </div>
                  )}
                  {missingColsFilter.size > 0 && (
                    <button
                      onClick={() => setMissingColsFilter(new Set())}
                      className="text-[11px] text-teal-600 hover:underline font-medium"
                    >
                      Limpar
                    </button>
                  )}
                </div>
              </div>
              <div className="flex flex-wrap gap-2">
                {emptyColumnOptions.map(col => (
                  <button
                    key={col.id}
                    onClick={() => toggleFilter(setMissingColsFilter, col.id)}
                    className={`px-3 py-1.5 text-xs font-medium rounded-full border transition-colors ${
                      missingColsFilter.has(col.id)
                        ? 'bg-amber-50 border-amber-500 text-amber-800 shadow-sm font-semibold'
                        : 'bg-white border-slate-200 text-slate-600 hover:border-slate-300'
                    }`}
                  >
                    Sem {col.label}
                  </button>
                ))}
              </div>
            </div>

            {/* Active Filters Clear Bar */}
            {hasActiveFilters && (
              <div className="flex items-center justify-between bg-teal-50/70 border border-teal-200 px-4 py-2 rounded-xl text-xs text-teal-800">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-semibold">Filtros Ativos:</span>
                  {selectedColors.size > 0 && (
                    <span className="bg-white px-2 py-0.5 rounded border border-teal-200 font-medium">
                      {selectedColors.size} cor(es)
                    </span>
                  )}
                  {selectedMedicos.size > 0 && (
                    <span className="bg-white px-2 py-0.5 rounded border border-teal-200 font-medium">
                      {selectedMedicos.size} médico(s)
                    </span>
                  )}
                  {missingColsFilter.size > 0 && (
                    <span className="bg-white px-2 py-0.5 rounded border border-amber-300 text-amber-800 font-medium">
                      Sem preenchimento ({missingColsFilter.size > 1 ? (missingColsMode === 'all' ? 'Todas: ' : 'Qualquer: ') : ''}
                      {Array.from(missingColsFilter).map(id => emptyColumnOptions.find(o => o.id === id)?.label || id).join(', ')})
                    </span>
                  )}
                  {(startDate || endDate) && (
                    <span className="bg-white px-2 py-0.5 rounded border border-teal-200 font-medium">
                      Período: {startDate || 'Início'} até {endDate || 'Fim'}
                    </span>
                  )}
                  {(minWaitTime || maxWaitTime) && (
                    <span className="bg-white px-2 py-0.5 rounded border border-teal-200 font-medium">
                      Duração: {minWaitTime ? `≥ ${minWaitTime} min` : ''} {maxWaitTime ? `≤ ${maxWaitTime} min` : ''}
                    </span>
                  )}
                  {searchTerm && (
                    <span className="bg-white px-2 py-0.5 rounded border border-teal-200 font-medium">
                      Busca: "{searchTerm}"
                    </span>
                  )}
                </div>
                <button 
                  onClick={clearAllFilters}
                  className="flex items-center gap-1 font-semibold text-teal-700 hover:text-teal-900 ml-4 hover:underline"
                >
                  <RotateCcw className="w-3 h-3" />
                  <span>Limpar Todos os Filtros</span>
                </button>
              </div>
            )}

            <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
              
              {/* Tabs */}
              <div className="flex border-b border-slate-200 px-2 bg-slate-50/50">
                <button 
                  onClick={() => setActiveTab('dashboard')}
                  className={`px-5 py-3 text-sm font-medium transition-colors border-b-2 ${
                    activeTab === 'dashboard' ? 'border-teal-600 text-teal-700' : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
                  }`}
                >
                  Dashboard
                </button>
                <button 
                  onClick={() => setActiveTab('data')}
                  className={`px-5 py-3 text-sm font-medium transition-colors border-b-2 ${
                    activeTab === 'data' ? 'border-teal-600 text-teal-700' : 'border-transparent text-slate-500 hover:text-slate-700 hover:border-slate-300'
                  }`}
                >
                  Dados Brutos
                </button>
              </div>

              <div className="p-6">
                {activeTab === 'dashboard' ? (
                  <div className="space-y-8 animate-in fade-in duration-300">
                    
                    {/* SECTION 1: Classificação por Cor (Charts) */}
                    <div>
                      <h2 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                        <span>Classificação e Triagem</span>
                      </h2>
                      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                        
                        {/* Chart: Appointments by Color */}
                        <div id="chart-appointments" className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm">
                          <h3 className="text-base font-semibold text-slate-800 mb-4">Total de Atendimentos por Cor</h3>
                          <div className="h-[280px]">
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={appointmentsData} margin={{ top: 10, right: 10, left: 0, bottom: 20 }}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                <XAxis dataKey="name" tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                                <YAxis tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                                <RechartsTooltip 
                                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                                  cursor={{fill: '#f8fafc'}}
                                />
                                <Bar dataKey="atendimentos" radius={[4, 4, 0, 0]}>
                                  {appointmentsData.map((entry, index) => (
                                    <Cell key={`cell-${index}`} fill={getColorHex(entry.name, index)} />
                                  ))}
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        </div>

                        {/* Chart: Wait Time by Color */}
                        <div id="chart-wait-time" className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm">
                          <h3 className="text-base font-semibold text-slate-800 mb-4">
                            Tempo Médio (minutos) por Cor ({columnOptions.find(c => c.id === startCol)?.label} → {columnOptions.find(c => c.id === endCol)?.label})
                          </h3>
                          <div className="h-[280px]">
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={waitTimeData} margin={{ top: 10, right: 10, left: 0, bottom: 20 }}>
                                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                <XAxis dataKey="name" tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                                <YAxis tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                                <RechartsTooltip 
                                  contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                                  cursor={{fill: '#f8fafc'}}
                                  formatter={(value: number) => [`${value} min`, 'Média']}
                                />
                                <Bar dataKey="mediaEspera" radius={[4, 4, 0, 0]}>
                                  {waitTimeData.map((entry, index) => (
                                    <Cell key={`cell-${index}`} fill={getColorHex(entry.name, index)} />
                                  ))}
                                </Bar>
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        </div>

                      </div>
                    </div>

                    {/* SECTION 2: Médicos / Profissionais (Ranking & Volume) */}
                    <div>
                      <div className="flex items-center justify-between mb-4">
                        <h2 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                          <Stethoscope className="w-5 h-5 text-teal-600" />
                          <span>Desempenho e Produtividade por Médico / Profissional</span>
                        </h2>
                        <span className="text-xs text-slate-500 font-medium">
                          Obedecendo intervalo: {columnOptions.find(c => c.id === startCol)?.label} → {columnOptions.find(c => c.id === endCol)?.label}
                        </span>
                      </div>

                      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                        
                        {/* Table: Ranking Resumo de Atendimentos por Médico */}
                        <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
                          <div className="px-5 py-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                            <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-2">
                              <UserCheck className="w-4 h-4 text-teal-600" />
                              Ranking de Atendimentos por Médico
                            </h3>
                            <div className="flex items-center gap-2">
                              <span className="text-xs text-slate-500 font-medium">
                                Total: {sortedDoctorRankingData.length}
                              </span>
                            </div>
                          </div>
                          <div className="overflow-x-auto overflow-y-auto max-h-[350px] flex-1">
                            <table className="w-full text-sm text-left">
                              <thead className="bg-slate-50 text-slate-500 font-medium sticky top-0 z-10 shadow-sm">
                                <tr>
                                  <th 
                                    onClick={handlePositionOrNameSort}
                                    className="px-5 py-3 border-b border-slate-200 bg-slate-50 cursor-pointer select-none hover:bg-slate-100 transition-colors group"
                                    title="Clique para alternar ordenação por Posição (1º, 2º...) ou Nome do Médico (A-Z)"
                                  >
                                    <div className="flex items-center gap-1.5">
                                      <span>Posição / Médico</span>
                                      {doctorSortField === 'name' ? (
                                        <span className="inline-flex items-center text-[10px] font-bold text-teal-700 bg-teal-50 px-1.5 py-0.5 rounded border border-teal-200">
                                          {doctorSortOrder === 'asc' ? 'Nome A→Z' : 'Nome Z→A'}
                                          {doctorSortOrder === 'asc' ? <ArrowUp className="w-2.5 h-2.5 ml-0.5" /> : <ArrowDown className="w-2.5 h-2.5 ml-0.5 text-teal-700" />}
                                        </span>
                                      ) : doctorSortField === 'rank' ? (
                                        <span className="inline-flex items-center text-[10px] font-bold text-teal-700 bg-teal-50 px-1.5 py-0.5 rounded border border-teal-200">
                                          {doctorSortOrder === 'asc' ? '1º→Último' : 'Último→1º'}
                                          {doctorSortOrder === 'asc' ? <ArrowUp className="w-2.5 h-2.5 ml-0.5" /> : <ArrowDown className="w-2.5 h-2.5 ml-0.5 text-teal-700" />}
                                        </span>
                                      ) : (
                                        <ArrowUpDown className="w-3.5 h-3.5 text-slate-400 opacity-40 group-hover:opacity-100 transition-opacity" />
                                      )}
                                    </div>
                                  </th>

                                  <th 
                                    onClick={() => handleDoctorSort('atendimentos')}
                                    className="px-5 py-3 border-b border-slate-200 bg-slate-50 text-right cursor-pointer select-none hover:bg-slate-100 transition-colors group"
                                    title="Clique para ordenar por Quantidade de Atendimentos"
                                  >
                                    <div className="flex items-center justify-end gap-1.5">
                                      <span>Qtd. Atendimentos</span>
                                      {doctorSortField === 'atendimentos' ? (
                                        <span className="inline-flex items-center text-[10px] font-bold text-teal-700 bg-teal-50 px-1.5 py-0.5 rounded border border-teal-200">
                                          {doctorSortOrder === 'desc' ? 'Maior' : 'Menor'}
                                          {doctorSortOrder === 'asc' ? <ArrowUp className="w-2.5 h-2.5 ml-0.5" /> : <ArrowDown className="w-2.5 h-2.5 ml-0.5 text-teal-700" />}
                                        </span>
                                      ) : (
                                        <ArrowUpDown className="w-3.5 h-3.5 text-slate-400 opacity-40 group-hover:opacity-100 transition-opacity" />
                                      )}
                                    </div>
                                  </th>

                                  <th 
                                    onClick={() => handleDoctorSort('mediaEspera')}
                                    className="px-5 py-3 border-b border-slate-200 bg-slate-50 text-right cursor-pointer select-none hover:bg-slate-100 transition-colors group"
                                    title="Clique para ordenar por Tempo Médio"
                                  >
                                    <div className="flex items-center justify-end gap-1.5">
                                      <span>Tempo Médio</span>
                                      {doctorSortField === 'mediaEspera' ? (
                                        <span className="inline-flex items-center text-[10px] font-bold text-teal-700 bg-teal-50 px-1.5 py-0.5 rounded border border-teal-200">
                                          {doctorSortOrder === 'desc' ? 'Maior' : 'Menor'}
                                          {doctorSortOrder === 'asc' ? <ArrowUp className="w-2.5 h-2.5 ml-0.5" /> : <ArrowDown className="w-2.5 h-2.5 ml-0.5 text-teal-700" />}
                                        </span>
                                      ) : (
                                        <ArrowUpDown className="w-3.5 h-3.5 text-slate-400 opacity-40 group-hover:opacity-100 transition-opacity" />
                                      )}
                                    </div>
                                  </th>
                                </tr>
                              </thead>
                              <tbody>
                                {sortedDoctorRankingData.map((item, index) => (
                                  <tr key={index} className="hover:bg-slate-50 transition-colors border-b border-slate-100 last:border-0">
                                    <td className="px-5 py-3 font-medium text-slate-700 flex items-center gap-2.5">
                                      <span className={`inline-flex items-center justify-center w-5 h-5 rounded-full text-[11px] font-bold shrink-0 ${
                                        item.originalRank === 1 ? 'bg-amber-100 text-amber-800 border border-amber-300' :
                                        item.originalRank === 2 ? 'bg-slate-200 text-slate-700 border border-slate-300' :
                                        item.originalRank === 3 ? 'bg-orange-100 text-orange-800 border border-orange-300' :
                                        'bg-slate-100 text-slate-500'
                                      }`}>
                                        {item.originalRank}
                                      </span>
                                      <span className="truncate max-w-[200px]" title={item.name}>{item.name}</span>
                                    </td>
                                    <td className="px-5 py-3 text-right tabular-nums font-semibold text-slate-800">{item.atendimentos}</td>
                                    <td className="px-5 py-3 text-right tabular-nums text-teal-700 font-medium">
                                      {item.mediaEspera !== null ? `${item.mediaEspera} min` : '-'}
                                    </td>
                                  </tr>
                                ))}
                                {sortedDoctorRankingData.length === 0 && (
                                  <tr>
                                    <td colSpan={3} className="px-5 py-8 text-center text-slate-500">
                                      Nenhum atendimento com médico informado no período ou filtro atual.
                                    </td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Chart: Quantidade de Atendimentos por Médico */}
                        <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm flex flex-col">
                          <div className="flex items-center justify-between mb-4">
                            <div>
                              <h3 className="text-base font-semibold text-slate-800">Quantidade de Atendimentos por Médico</h3>
                              <p className="text-xs text-slate-500">Top médicos com maior volume (exclui registros sem médico)</p>
                            </div>
                            <span className="text-xs font-semibold px-2 py-0.5 rounded bg-teal-50 text-teal-700 border border-teal-200">
                              {doctorAppointmentsData.length} médicos
                            </span>
                          </div>

                          <div className="h-[300px] flex-1">
                            {doctorAppointmentsData.length > 0 ? (
                              <ResponsiveContainer width="100%" height="100%">
                                <BarChart data={doctorAppointmentsData.slice(0, 10)} margin={{ top: 10, right: 10, left: 0, bottom: 40 }}>
                                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                                  <XAxis 
                                    dataKey="name" 
                                    tick={{ fontSize: 11, fill: '#64748b' }} 
                                    interval={0}
                                    angle={-25}
                                    textAnchor="end"
                                    axisLine={false} 
                                    tickLine={false} 
                                  />
                                  <YAxis tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                                  <RechartsTooltip 
                                    contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                                    cursor={{fill: '#f8fafc'}}
                                  />
                                  <Bar dataKey="atendimentos" fill="#0f766e" radius={[4, 4, 0, 0]} name="Atendimentos" />
                                </BarChart>
                              </ResponsiveContainer>
                            ) : (
                              <div className="h-full flex flex-col items-center justify-center text-slate-400 text-sm">
                                <Stethoscope className="w-8 h-8 text-slate-300 mb-2" />
                                <p>Nenhum atendimento com médico identificado no período.</p>
                              </div>
                            )}
                          </div>
                        </div>

                      </div>
                    </div>

                    {/* SECTION 3: Resumo por Cor & Volume Diário Geral */}
                    <div>
                      <h2 className="text-lg font-bold text-slate-800 mb-4">Volume Diário e Resumo Geral</h2>
                      <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                        
                        {/* Summary Table: By Color */}
                        <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
                          <div className="px-5 py-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                            <h3 className="text-sm font-semibold text-slate-800">Resumo de Atendimentos por Cor</h3>
                          </div>
                          <div className="overflow-x-auto overflow-y-auto max-h-[350px] flex-1">
                            <table className="w-full text-sm text-left">
                              <thead className="bg-slate-50 text-slate-500 font-medium sticky top-0 z-10 shadow-sm">
                                <tr>
                                  <th className="px-5 py-3 border-b border-slate-200 bg-slate-50">Classificação</th>
                                  <th className="px-5 py-3 border-b border-slate-200 bg-slate-50 text-right">Qtd. Atendimentos</th>
                                  <th className="px-5 py-3 border-b border-slate-200 bg-slate-50 text-right">Tempo Médio ({columnOptions.find(c => c.id === startCol)?.label} → {columnOptions.find(c => c.id === endCol)?.label})</th>
                                </tr>
                              </thead>
                              <tbody>
                                {appointmentsData.map((item, index) => {
                                  const waitTimeItem = waitTimeData.find(w => w.name === item.name);
                                  return (
                                    <tr key={index} className="hover:bg-slate-50 transition-colors border-b border-slate-100 last:border-0">
                                      <td className="px-5 py-3 font-medium text-slate-700 flex items-center gap-2">
                                        <span 
                                          className="inline-block w-3 h-3 rounded-full" 
                                          style={{backgroundColor: getColorHex(item.name, index)}} 
                                        />
                                        {item.name}
                                      </td>
                                      <td className="px-5 py-3 text-right tabular-nums">{item.atendimentos}</td>
                                      <td className="px-5 py-3 text-right tabular-nums">
                                        {waitTimeItem ? `${waitTimeItem.mediaEspera} min` : '-'}
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Summary Table: General Daily Volume */}
                        <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
                          <div className="px-5 py-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                            <h3 className="text-sm font-semibold text-slate-800">Volume Diário Geral</h3>
                          </div>
                          <div className="overflow-x-auto overflow-y-auto max-h-[350px] flex-1">
                            <table className="w-full text-sm text-left">
                              <thead className="bg-slate-50 text-slate-500 font-medium sticky top-0 z-10 shadow-sm">
                                <tr>
                                  <th className="px-5 py-3 border-b border-slate-200 bg-slate-50">Data</th>
                                  <th className="px-5 py-3 border-b border-slate-200 bg-slate-50 text-right">Total de Atendimentos</th>
                                </tr>
                              </thead>
                              <tbody>
                                {dailyVolumeData.map((item, index) => (
                                  <tr key={index} className="hover:bg-slate-50 transition-colors border-b border-slate-100 last:border-0">
                                    <td className="px-5 py-3 font-medium text-slate-700">{item.date}</td>
                                    <td className="px-5 py-3 text-right tabular-nums">{item.total}</td>
                                  </tr>
                                ))}
                                {dailyVolumeData.length > 0 && (
                                  <tr className="bg-slate-50/80 font-semibold text-slate-800 border-t-2 border-slate-200 sticky bottom-0 shadow-sm">
                                    <td className="px-5 py-3">Total Geral</td>
                                    <td className="px-5 py-3 text-right tabular-nums">
                                      {dailyVolumeData.reduce((acc, curr) => acc + curr.total, 0)}
                                    </td>
                                  </tr>
                                )}
                                {dailyVolumeData.length === 0 && (
                                  <tr>
                                    <td colSpan={2} className="px-5 py-8 text-center text-slate-500">
                                      Nenhum dado encontrado para o período.
                                    </td>
                                  </tr>
                                )}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      </div>

                      {/* Chart: Volume de Atendimentos por Dia (Bar Chart posicionado abaixo da tabela de volume diário geral) */}
                      <div id="chart-daily-volume" className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm mt-6">
                        <div className="flex items-center justify-between mb-4">
                          <h3 className="text-base font-semibold text-slate-800">
                            Volume de Atendimentos (Picos por Dia)
                          </h3>
                          <span className="text-xs text-slate-500 font-medium">
                            Gráfico de barras geral sem distinção de cor
                          </span>
                        </div>
                        <div className="h-[320px]">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dailyVolumeData} margin={{ top: 10, right: 20, left: 0, bottom: 20 }}>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                              <XAxis 
                                dataKey="date" 
                                tick={{ fontSize: 12, fill: '#64748b' }} 
                                axisLine={false} 
                                tickLine={false} 
                              />
                              <YAxis tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                              <RechartsTooltip 
                                contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                                cursor={{fill: '#f8fafc'}}
                              />
                              <Bar dataKey="total" name="Total Atendimentos" fill="#0d9488" radius={[4, 4, 0, 0]} />
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>

                    </div>

                  </div>
                ) : (
                  <div className="animate-in fade-in duration-300">
                    {/* Action Bar for Raw Data */}
                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
                      <div className="flex items-center gap-4 w-full sm:w-auto flex-1">
                        <div className="text-sm text-slate-600 whitespace-nowrap">
                          Exibindo <span className="font-semibold text-slate-800">{filteredData.length}</span> registros
                        </div>
                        <div className="relative w-full max-w-sm">
                          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                          <input 
                            type="text" 
                            placeholder="Buscar por médico, cor, data..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                            className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-sm focus:ring-1 focus:ring-teal-500 outline-none"
                          />
                        </div>
                      </div>
                      <div className="flex items-center gap-3 w-full sm:w-auto">
                        <button 
                          onClick={() => setIsDeleteModalOpen(true)}
                          className="flex items-center justify-center gap-2 px-4 py-2 bg-rose-50 border border-rose-200 text-rose-700 hover:bg-rose-100 rounded-lg text-sm font-medium transition-colors w-full sm:w-auto"
                        >
                          <Trash2 className="w-4 h-4" />
                          <span>Excluir em Lote</span>
                        </button>
                        <button 
                          onClick={handleExport}
                          className="flex items-center justify-center gap-2 px-4 py-2 bg-teal-50 border border-teal-200 text-teal-700 hover:bg-teal-100 rounded-lg text-sm font-medium transition-colors w-full sm:w-auto"
                        >
                          <Download className="w-4 h-4" />
                          <span>Exportar XLSX</span>
                        </button>
                      </div>
                    </div>

                    {/* Raw Data Table with Edit/Delete */}
                    <div className="border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                      <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                        <table className="w-full text-sm text-left whitespace-nowrap relative">
                          <thead className="bg-slate-50 text-slate-600 font-semibold sticky top-0 z-10 shadow-sm">
                            <tr>
                              <th className="px-4 py-3 border-b border-slate-200">Admissão</th>
                              <th className="px-4 py-3 border-b border-slate-200">Classificação</th>
                              <th className="px-4 py-3 border-b border-slate-200">Atendimento</th>
                              <th className="px-4 py-3 border-b border-slate-200">Fim Atendimento</th>
                              <th className="px-4 py-3 border-b border-slate-200">Cor</th>
                              <th className="px-4 py-3 border-b border-slate-200">Médico / Profissional</th>
                              <th className="px-4 py-3 border-b border-slate-200 text-right">Ações</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {filteredData.map((row) => (
                              <tr key={row.id} className="hover:bg-slate-50/80 transition-colors">
                                {editingId === row.id ? (
                                  <>
                                    <td className="px-4 py-2">
                                      <input 
                                        type="text" 
                                        value={editForm.rawAdmissao || ''} 
                                        onChange={e => setEditForm({...editForm, rawAdmissao: e.target.value})}
                                        className="w-full text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
                                      />
                                    </td>
                                    <td className="px-4 py-2">
                                      <input 
                                        type="text" 
                                        value={editForm.rawHoraClassificacao || ''} 
                                        onChange={e => setEditForm({...editForm, rawHoraClassificacao: e.target.value})}
                                        className="w-full text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
                                      />
                                    </td>
                                    <td className="px-4 py-2">
                                      <input 
                                        type="text" 
                                        value={editForm.rawAtendimento || ''} 
                                        onChange={e => setEditForm({...editForm, rawAtendimento: e.target.value})}
                                        className="w-full text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
                                      />
                                    </td>
                                    <td className="px-4 py-2">
                                      <input 
                                        type="text" 
                                        value={editForm.rawFimAtendimento || ''} 
                                        onChange={e => setEditForm({...editForm, rawFimAtendimento: e.target.value})}
                                        className="w-full text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
                                      />
                                    </td>
                                    <td className="px-4 py-2">
                                      <input 
                                        type="text" 
                                        value={editForm.classificacao || ''} 
                                        onChange={e => setEditForm({...editForm, classificacao: e.target.value})}
                                        className="w-28 text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
                                      />
                                    </td>
                                    <td className="px-4 py-2">
                                      <input 
                                        type="text" 
                                        value={editForm.medico || ''} 
                                        onChange={e => setEditForm({...editForm, medico: e.target.value})}
                                        placeholder="Nome do Médico"
                                        className="w-36 text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
                                      />
                                    </td>
                                    <td className="px-4 py-2 text-right">
                                      <div className="flex justify-end gap-1">
                                        <button onClick={() => handleSaveEdit(row.id)} className="p-1.5 text-white bg-teal-600 hover:bg-teal-700 rounded transition-colors" title="Salvar">
                                          <Save className="w-4 h-4"/>
                                        </button>
                                        <button onClick={() => setEditingId(null)} className="p-1.5 text-slate-600 bg-slate-200 hover:bg-slate-300 rounded transition-colors" title="Cancelar">
                                          <X className="w-4 h-4"/>
                                        </button>
                                      </div>
                                    </td>
                                  </>
                                ) : (
                                  <>
                                    <td className="px-4 py-3">{row.admissao ? format(row.admissao, 'dd/MM/yyyy HH:mm:ss') : row.rawAdmissao}</td>
                                    <td className="px-4 py-3">{row.horaClassificacao ? format(row.horaClassificacao, 'dd/MM/yyyy HH:mm:ss') : row.rawHoraClassificacao}</td>
                                    <td className="px-4 py-3">{row.atendimento ? format(row.atendimento, 'dd/MM/yyyy HH:mm:ss') : row.rawAtendimento}</td>
                                    <td className="px-4 py-3">{row.fimAtendimento ? format(row.fimAtendimento, 'dd/MM/yyyy HH:mm:ss') : row.rawFimAtendimento}</td>
                                    <td className="px-4 py-3 font-medium flex items-center gap-2">
                                      <span 
                                        className="inline-block w-2 h-2 rounded-full" 
                                        style={{backgroundColor: getColorHex(row.classificacao, 0)}} 
                                      />
                                      {row.classificacao}
                                    </td>
                                    <td className="px-4 py-3 text-slate-700">
                                      <span className="inline-flex items-center gap-1.5">
                                        <Stethoscope className="w-3.5 h-3.5 text-teal-600" />
                                        <span className="font-medium">{row.medico || 'Não informado'}</span>
                                      </span>
                                    </td>
                                    <td className="px-4 py-3 text-right">
                                      <div className="flex justify-end gap-2 opacity-60 hover:opacity-100 transition-opacity">
                                        <button onClick={() => handleEditClick(row)} className="p-1 text-teal-600 hover:bg-teal-50 rounded" title="Editar">
                                          <Edit2 className="w-4 h-4" />
                                        </button>
                                        <button onClick={() => handleDeleteRow(row.id)} className="p-1 text-rose-600 hover:bg-rose-50 rounded" title="Excluir">
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                    </td>
                                  </>
                                )}
                              </tr>
                            ))}
                            {filteredData.length === 0 && (
                              <tr>
                                <td colSpan={7} className="px-4 py-8 text-center text-slate-500">
                                  Nenhum registro encontrado.
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {!hasData && !isProcessing && !isLoading && (
          <div className="bg-white border border-slate-200 border-dashed rounded-2xl p-12 text-center max-w-xl mx-auto shadow-sm">
            <div className="w-16 h-16 bg-teal-50 text-teal-600 rounded-full flex items-center justify-center mx-auto mb-4">
              <FileSpreadsheet className="w-8 h-8" />
            </div>
            <h3 className="text-lg font-bold text-slate-800 mb-1">Nenhum dado importado</h3>
            <p className="text-sm text-slate-500 mb-6">
              Envie sua planilha Excel (.xlsx) contendo as colunas de Admissão, Classificação, Atendimento, Fim do Atendimento, Cor e Médico / Profissional (Coluna F).
            </p>
            <label className="inline-flex items-center gap-2 px-5 py-2.5 bg-teal-600 text-white hover:bg-teal-700 rounded-xl cursor-pointer font-medium text-sm transition-all shadow-sm shadow-teal-700/20">
              <UploadCloud className="w-4 h-4" />
              <span>Selecionar Arquivo XLSX</span>
              <input 
                type="file" 
                accept=".xlsx, .xls" 
                onChange={handleFileUpload} 
                className="hidden" 
              />
            </label>
          </div>
        )}

      </div>
    </div>
  );
}
export default App;
