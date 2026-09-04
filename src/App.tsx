import React, { useState, useMemo, useEffect } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer,
  Cell, LineChart, Line
} from 'recharts';
import { UploadCloud, FileSpreadsheet, Loader2, Filter, Calendar, Trash2, Settings, Download, Edit2, Save, X, Search, Clock } from 'lucide-react';
import { startOfDay, endOfDay, parseISO, format } from 'date-fns';
import * as XLSX from 'xlsx';
import { 
  ProcessedRecord, processExcelData, columnOptions, getAverageWaitTimeByClassificacao, 
  getAppointmentsByClassificacao, getTimelineVolumeByClassificacao, parseCustomDate, getDailyVolumeSummary
} from './utils';

// Tailwind colors for common classifications
const colorMap: Record<string, string> = {
  'Verde': '#22c55e',
  'Amarelo': '#eab308',
  'Laranja': '#f97316',
  'Vermelho': '#ef4444',
  'Azul': '#3b82f6',
  'Sem Classificação': '#94a3b8'
};

const DEFAULT_COLORS = ['#0f766e', '#0369a1', '#be123c', '#b45309', '#6d28d9', '#15803d'];

export default function App() {
  const [data, setData] = useState<ProcessedRecord[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  
  // Tab state
  const [activeTab, setActiveTab] = useState<'dashboard' | 'data'>('dashboard');

  // Settings state
  const [startCol, setStartCol] = useState<string>('admissao');
  const [endCol, setEndCol] = useState<string>('horaClassificacao');
  
  // Filter state
  const [showFilters, setShowFilters] = useState(false);
  const [selectedColors, setSelectedColors] = useState<Set<string>>(new Set());
  
  // Time filter state
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [minWaitTime, setMinWaitTime] = useState<string>('');
  const [maxWaitTime, setMaxWaitTime] = useState<string>('');

  // Raw Data Search Filter
  const [searchTerm, setSearchTerm] = useState<string>('');

  // Row Edit state
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editForm, setEditForm] = useState<Partial<ProcessedRecord>>({});

  // Bulk Delete state
  const [isDeleteModalOpen, setIsDeleteModalOpen] = useState(false);
  const [deleteStart, setDeleteStart] = useState<string>('');
  const [deleteEnd, setDeleteEnd] = useState<string>('');

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
          // Append new records to current state to avoid full refetch delay
          setData(prev => [...prev, ...processed]);
          setSelectedColors(new Set());
          setStartDate('');
          setEndDate('');
          alert(`Importação concluída! ${processed.length} registros adicionados.`);
        } else {
          throw new Error('Falha na API');
        }
      } catch (error) {
        console.error("Error parsing Excel:", error);
        alert("Erro ao processar o arquivo Excel e salvar no banco.");
      } finally {
        setIsProcessing(false);
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // reset input
  };

  const uniqueColors = useMemo(() => Array.from(new Set(data.map(d => d.classificacao).filter(Boolean))).sort(), [data]);
  
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
      
      // Duration Filter
      if (minTimeNum !== null || maxTimeNum !== null) {
        const startT = d[startCol as keyof ProcessedRecord] as Date | null;
        const endT = d[endCol as keyof ProcessedRecord] as Date | null;
        
        if (!startT || !endT) return false;
        
        const diffMinutes = (endT.getTime() - startT.getTime()) / (1000 * 60);
        if (minTimeNum !== null && diffMinutes < minTimeNum) return false;
        if (maxTimeNum !== null && diffMinutes > maxTimeNum) return false;
      }

      // Text Search Filter (only relevant if term exists)
      if (searchLower) {
        const matchStr = `${d.classificacao} ${d.rawAdmissao} ${d.rawHoraClassificacao} ${d.rawAtendimento} ${d.rawFimAtendimento}`.toLowerCase();
        if (!matchStr.includes(searchLower)) return false;
      }
      
      return true;
    });
  }, [data, selectedColors, startDate, endDate, startCol, endCol, searchTerm, minWaitTime, maxWaitTime]);

  const hasData = data.length > 0;
  
  const waitTimeData = hasData ? getAverageWaitTimeByClassificacao(filteredData, startCol, endCol) : [];
  const appointmentsData = hasData ? getAppointmentsByClassificacao(filteredData) : [];
  const volumeTimelineData = hasData ? getTimelineVolumeByClassificacao(filteredData, startCol) : [];
  const dailyVolumeData = hasData ? getDailyVolumeSummary(filteredData, startCol) : [];

  const toggleFilter = (setter: React.Dispatch<React.SetStateAction<Set<string>>>, value: string) => {
    setter(prev => {
      const next = new Set(prev);
      if (next.has(value)) next.delete(value);
      else next.add(value);
      return next;
    });
  };

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

    // Determine which IDs to delete
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
        <p className="text-slate-500 font-medium">Carregando dados do servidor...</p>
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
              Selecione o período (baseado na coluna "{columnOptions.find(c => c.id === startCol)?.label}") para excluir os registros permanentemente.
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
                className="px-4 py-2 text-sm font-medium text-slate-600 hover:bg-slate-100 rounded-lg transition-colors"
              >
                Cancelar
              </button>
              <button 
                onClick={handleBulkDelete} 
                className="px-4 py-2 text-sm font-medium text-white bg-rose-600 hover:bg-rose-700 rounded-lg transition-colors flex items-center gap-2 shadow-sm"
              >
                <Trash2 className="w-4 h-4" /> Excluir
              </button>
            </div>
          </div>
        </div>
      )}

      <div className="max-w-[1400px] mx-auto space-y-6">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row md:items-start justify-between gap-4 bg-white p-6 rounded-xl shadow-sm border border-slate-200">
          <div>
            <h1 className="text-2xl font-bold text-slate-800 tracking-tight">Painel de Classificação de Risco</h1>
            <p className="text-slate-500 text-sm mt-1 max-w-xl">Carregue suas planilhas XLSX. Os dados serão acumulados no banco de dados.</p>
          </div>
          
          <div className="flex flex-wrap items-center gap-3">
            <label className="flex items-center justify-center gap-2 px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg cursor-pointer transition-colors text-sm font-medium shadow-sm">
              {isProcessing ? <Loader2 className="w-4 h-4 animate-spin" /> : <UploadCloud className="w-4 h-4" />}
              <span>{isProcessing ? 'Processando...' : 'Adicionar XLSX'}</span>
              <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} className="hidden" />
            </label>
          </div>
        </header>

        {hasData && (
          <div className="space-y-6">
            {/* Global Filters & Configuration */}
            <div className="grid grid-cols-1 lg:grid-cols-4 gap-4">
              {/* Configuração de Tempos */}
              <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm space-y-4">
                <h3 className="text-sm font-semibold text-slate-700 flex items-center gap-2">
                  <Settings className="w-4 h-4 text-slate-400" /> Configuração de Tempos
                </h3>
                
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">Início do Intervalo</label>
                  <select 
                    value={startCol}
                    onChange={e => setStartCol(e.target.value)}
                    className="w-full border border-slate-300 rounded-md p-2 text-sm text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  >
                    {columnOptions.map(opt => <option key={opt.id} value={opt.id}>{opt.label}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-medium text-slate-500 mb-1">Fim do Intervalo</label>
                  <select 
                    value={endCol}
                    onChange={e => setEndCol(e.target.value)}
                    className="w-full border border-slate-300 rounded-md p-2 text-sm text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                  >
                    {columnOptions.map(opt => <option key={opt.id} value={opt.id}>{opt.label}</option>)}
                  </select>
                </div>
              </div>

              {/* Date Range */}
              <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm">
                <h3 className="text-sm font-semibold text-slate-700 mb-3 flex items-center gap-2">
                  <Calendar className="w-4 h-4 text-slate-400" /> Período
                </h3>
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-500 mb-1">Data Inicial</label>
                    <input 
                      type="date" 
                      value={startDate} 
                      onChange={(e) => setStartDate(e.target.value)}
                      className="w-full border border-slate-300 rounded-md p-2 text-sm text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-500 mb-1">Data Final</label>
                    <input 
                      type="date" 
                      value={endDate} 
                      onChange={(e) => setEndDate(e.target.value)}
                      className="w-full border border-slate-300 rounded-md p-2 text-sm text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                    />
                  </div>
                </div>
              </div>

              {/* Duration Filter */}
              <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm">
                <h3 className="text-sm font-semibold text-slate-700 mb-3 flex items-center gap-2">
                  <Clock className="w-4 h-4 text-slate-400" /> Duração (minutos)
                </h3>
                <div className="space-y-3">
                  <div>
                    <label className="block text-xs font-medium text-slate-500 mb-1">Tempo Mínimo (Ex: 240)</label>
                    <input 
                      type="number" 
                      min="0"
                      value={minWaitTime} 
                      onChange={(e) => setMinWaitTime(e.target.value)}
                      placeholder="Mínimo"
                      className="w-full border border-slate-300 rounded-md p-2 text-sm text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                    />
                  </div>
                  <div>
                    <label className="block text-xs font-medium text-slate-500 mb-1">Tempo Máximo</label>
                    <input 
                      type="number" 
                      min="0"
                      value={maxWaitTime} 
                      onChange={(e) => setMaxWaitTime(e.target.value)}
                      placeholder="Máximo"
                      className="w-full border border-slate-300 rounded-md p-2 text-sm text-slate-700 focus:border-teal-500 focus:ring-1 focus:ring-teal-500"
                    />
                  </div>
                </div>
              </div>

              {/* Colors Filter */}
              <div className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm">
                <h3 className="text-sm font-semibold text-slate-700 mb-3 flex items-center gap-2">
                  <Filter className="w-4 h-4 text-slate-400" /> Filtrar por Cor
                </h3>
                <div className="flex flex-wrap gap-2 max-h-32 overflow-y-auto">
                  {uniqueColors.map(color => (
                    <button
                      key={color}
                      onClick={() => toggleFilter(setSelectedColors, color)}
                      className={`px-3 py-1.5 text-xs font-medium rounded-full border transition-colors ${
                        selectedColors.has(color) 
                          ? 'bg-teal-50 border-teal-500 text-teal-700 shadow-sm' 
                          : 'bg-white border-slate-200 text-slate-600 hover:border-slate-300'
                      }`}
                    >
                      <span 
                        className="inline-block w-2 h-2 rounded-full mr-2" 
                        style={{backgroundColor: getColorHex(color, 0)}} 
                      />
                      {color}
                    </button>
                  ))}
                </div>
              </div>
            </div>

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
                  <div className="space-y-6 animate-in fade-in duration-300">
                    {/* Dashboard Grid */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    
                    {/* Chart: Appointments by Color */}
                    <div id="chart-appointments" className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm">
                      <h3 className="text-base font-semibold text-slate-800 mb-4">Total de Atendimentos por Cor</h3>
                      <div className="h-[300px]">
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
                      <h3 className="text-base font-semibold text-slate-800 mb-4">Tempo Médio (minutos) por Cor</h3>
                      <div className="h-[300px]">
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

                    {/* Chart: Volume Timeline */}
                    <div id="chart-timeline" className="bg-white p-5 border border-slate-200 rounded-xl shadow-sm lg:col-span-2">
                      <h3 className="text-base font-semibold text-slate-800 mb-4">Volume de Atendimentos (Picos por Dia)</h3>
                      <div className="h-[350px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={volumeTimelineData} margin={{ top: 10, right: 20, left: 0, bottom: 20 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                            <XAxis 
                              dataKey="name" 
                              tick={{ fontSize: 12, fill: '#64748b' }} 
                              axisLine={false} 
                              tickLine={false} 
                            />
                            <YAxis tick={{ fontSize: 12, fill: '#64748b' }} axisLine={false} tickLine={false} />
                            <RechartsTooltip 
                              contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                            />
                            <Legend iconType="circle" wrapperStyle={{ fontSize: '12px', paddingTop: '20px' }} />
                            
                            {(selectedColors.size > 0 ? Array.from(selectedColors) : uniqueColors).map((color, idx) => (
                              <Line 
                                key={color}
                                type="monotone" 
                                dataKey={color} 
                                name={color}
                                stroke={getColorHex(color, idx)} 
                                strokeWidth={2}
                                dot={{ r: 4, strokeWidth: 0, fill: getColorHex(color, idx) }}
                                activeDot={{ r: 6 }}
                              />
                            ))}
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                    </div>
                  </div>

                  {/* Summary Tables Grid */}
                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-6 mt-6">
                    {/* Summary Table: By Color */}
                    <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden flex flex-col">
                      <div className="px-5 py-4 border-b border-slate-200 bg-slate-50 flex items-center justify-between">
                        <h3 className="text-sm font-semibold text-slate-800">Resumo de Atendimentos por Cor</h3>
                      </div>
                      <div className="overflow-x-auto flex-1">
                        <table className="w-full text-sm text-left h-full">
                          <thead className="bg-slate-50 text-slate-500 font-medium">
                            <tr>
                              <th className="px-5 py-3 border-b border-slate-200">Classificação</th>
                              <th className="px-5 py-3 border-b border-slate-200 text-right">Qtd. Atendimentos</th>
                              <th className="px-5 py-3 border-b border-slate-200 text-right">Tempo Médio ({columnOptions.find(c => c.id === startCol)?.label} → {columnOptions.find(c => c.id === endCol)?.label})</th>
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
                      <div className="overflow-x-auto overflow-y-auto max-h-[400px] flex-1">
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
                          placeholder="Buscar registros..."
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
                            <th className="px-4 py-3 border-b border-slate-200">Classificação (Hora)</th>
                            <th className="px-4 py-3 border-b border-slate-200">Atendimento</th>
                            <th className="px-4 py-3 border-b border-slate-200">Fim Atendimento</th>
                            <th className="px-4 py-3 border-b border-slate-200">Cor</th>
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
                                      className="w-32 text-xs p-1.5 border border-teal-300 rounded focus:ring-1 focus:ring-teal-500 outline-none"
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
                              <td colSpan={6} className="px-4 py-8 text-center text-slate-500">
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
          <div className="flex flex-col items-center justify-center p-12 bg-white rounded-xl border border-slate-200 border-dashed">
            <div className="w-16 h-16 bg-teal-50 rounded-full flex items-center justify-center mb-4 text-teal-600">
              <FileSpreadsheet className="w-8 h-8" />
            </div>
            <h2 className="text-lg font-medium text-slate-800">Nenhum dado carregado</h2>
            <p className="text-slate-500 text-sm mt-1 max-w-sm text-center">Faça o upload de uma planilha XLSX contendo as colunas de Admissão, Classificação, Atendimento e Cor para visualizar o painel. Os dados serão salvos para as próximas visitas.</p>
          </div>
        )}

      </div>
    </div>
  );
}
