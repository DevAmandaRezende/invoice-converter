import React, { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from './lib/supabase';
import type { User as SupabaseUser } from '@supabase/supabase-js';
import AuthPage from './components/AuthPage';
import {
  Upload,
  FileSpreadsheet,
  ArrowRight,
  Settings,
  CheckCircle2,
  Download,
  Table,
  Sparkles,
  RefreshCw,
  Trash2,
  ChevronRight,
  ChevronDown,
  Plus,
  Minus,
  Check,
  HelpCircle,
  AlertCircle,
  Eye,
  Hash,
  Database,
  Grid,
  History,
  Calendar,
  LogOut,
  User
} from 'lucide-react';
import {
  TARGET_COLUMNS,
  ColumnMappingConfig,
  ExcelRow,
  PreviewData,
  TransformationType,
  SourceType
} from './types';
import {
  sugerirMapeamentoAutomatico,
  processarLinha,
  applyTransformation,
  sanitizeCsvValue,
  SAMPLE_TECH_STORE_ROWS,
  SAMPLE_WEB_ACADEMY_ROWS
} from './utils';

export interface HistoryItem {
  id: string;
  filename: string;
  clientName: string;
  monthName: string;
  timestamp: string;
  filePath: string;
}

export default function App() {
  const [user, setUser] = useState<SupabaseUser | null>(null);
  const [authLoading, setAuthLoading] = useState(true);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setUser(data.session?.user ?? null);
      setAuthLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => {
      setUser(session?.user ?? null);
    });
    return () => listener.subscription.unsubscribe();
  }, []);

  if (authLoading) {
    return (
      <div className="min-h-screen bg-[#0F0E1A] flex items-center justify-center">
        <div className="h-8 w-8 border-2 border-[#6366F1] border-t-transparent rounded-full animate-spin" />
      </div>
    );
  }

  if (!user) return <AuthPage />;

  return <AppContent user={user} />;
}

function AppContent({ user }: { user: SupabaseUser }) {
  // File data state
  const [uploadedFileName, setUploadedFileName] = useState<string | null>(null);
  const [uploadedColumns, setUploadedColumns] = useState<string[]>([]);
  const [rawRows, setRawRows] = useState<ExcelRow[]>([]);
  const [mappingConfigs, setMappingConfigs] = useState<ColumnMappingConfig[]>([]);
  
  // Custom suggestion notification
  const [suggestionCount, setSuggestionCount] = useState<number>(0);
  const [showNotification, setShowNotification] = useState<boolean>(false);

  // Tabs structure state
  const [activePreviewTab, setActivePreviewTab] = useState<'result' | 'source'>('result');

  // Input states for custom fixed values or details
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [dragActive, setDragActive] = useState<boolean>(false);

  // Export modal and history states
  const [isExportModalOpen, setIsExportModalOpen] = useState<boolean>(false);
  const [clientNameInput, setClientNameInput] = useState<string>('');
  const [monthNameInput, setMonthNameInput] = useState<string>(() => {
    const months = [
      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
    ];
    return months[new Date().getMonth()];
  });
  
  const [exportHistory, setExportHistory] = useState<HistoryItem[]>([]);
  const [showUserModal, setShowUserModal] = useState(false);
  const [openHistoryMenu, setOpenHistoryMenu] = useState<string | null>(null);
  const [menuPosition, setMenuPosition] = useState<{ top: number; right: number } | null>(null);
  const profileName: string | null = user.user_metadata?.name ?? null;
  const [editingName, setEditingName] = useState(false);
  const [nameInput, setNameInput] = useState('');
  const [savingName, setSavingName] = useState(false);

  const saveProfileName = async () => {
    if (!nameInput.trim()) return;
    setSavingName(true);
    const { error } = await supabase.auth.updateUser({ data: { name: nameInput.trim() } });
    setSavingName(false);
    if (!error) setEditingName(false);
  };

  useEffect(() => {
    supabase
      .from('export_history')
      .select('*')
      .eq('user_id', user.id)
      .order('created_at', { ascending: false })
      .then(({ data, error }) => {
        if (error) { console.error(error); return; }
        if (data) {
          setExportHistory(data.map((row) => ({
            id: row.id,
            filename: row.filename,
            clientName: row.client_name,
            monthName: row.month_name,
            timestamp: row.timestamp,
            filePath: row.file_path,
          })));
        }
      });
  }, [user.id]);

  // Error/Success statuses
  const [alertMsg, setAlertMsg] = useState<{ type: 'success' | 'info' | 'error', text: string } | null>(null);

  // Toggle advanced mapping view
  const [expandedField, setExpandedField] = useState<string | null>(null);

  // Initialize with empty mapping configurations
  useEffect(() => {
    resetMapping();
  }, []);

  const resetMapping = () => {
    const initial = TARGET_COLUMNS.map((col) => ({
      targetKey: col.key,
      sourceType: 'column' as SourceType,
      sourceColumn: '',
      fixedValue: col.defaultFixed || '',
      concatColumns: [] as string[],
      concatSeparator: ' ',
      transformations: getSuggestedTransformations(col.key),
    }));
    setMappingConfigs(initial);
  };

  const notify = (text: string, type: 'success' | 'info' | 'error' = 'info') => {
    setAlertMsg({ type, text });
    setTimeout(() => {
      setAlertMsg(null);
    }, 5000);
  };

  function getSuggestedTransformations(targetKey: string): TransformationType[] {
    if (targetKey === 'data_do_pagamento') return ['converter_data'];
    if (targetKey === 'endereco_cep') return ['formatar_cep'];
    if (targetKey === 'documento_do_cliente') return [];
    if (targetKey.includes('email')) return ['minusculo', 'remover_espacos'];
    if (targetKey.includes('nome')) return ['remover_espacos'];
    return [];
  }

  // Handle excel drag & drop events
  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      const ext = file.name.split('.').pop()?.toLowerCase();
      if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
        parseFile(file);
      } else {
        notify("Formato inválido. Por favor, envie um arquivo Excel (.xlsx, .xls) ou CSV.", "error");
      }
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      parseFile(e.target.files[0]);
    }
  };

  // Main workbook parser (via SheetJS)
  const parseFile = (file: File) => {
    if (file.size > 20 * 1024 * 1024) {
      notify('Arquivo muito grande. Limite de 20MB.', 'error');
      return;
    }
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const binaryString = evt.target?.result;
        const workbook = XLSX.read(binaryString, { type: 'binary', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Parse raw rows as JSON array
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: "" }) as ExcelRow[];
        
        if (jsonRows.length === 0) {
          notify("A planilha enviada está vazia ou sem dados estruturados.", "error");
          return;
        }

        const columns = Object.keys(jsonRows[0]);
        setRawRows(jsonRows);
        setUploadedColumns(columns);
        setUploadedFileName(file.name);

        // Run smart suggestion algorithms on newly loaded columns
        let autoSuggestedCount = 0;
        const newConfigs = TARGET_COLUMNS.map((target) => {
          const suggestion = sugerirMapeamentoAutomatico(columns, target.key);
          const hasSuggest = suggestion.sourceColumn !== "";
          if (hasSuggest) {
            autoSuggestedCount++;
          }
          
          return {
            targetKey: target.key,
            sourceType: hasSuggest ? ('column' as SourceType) : (target.defaultFixed ? 'fixed' : 'column'),
            sourceColumn: suggestion.sourceColumn,
            fixedValue: suggestion.fixedValue || target.defaultFixed || '',
            concatColumns: [] as string[],
            concatSeparator: ' ',
            transformations: getSuggestedTransformations(target.key),
          };
        });

        setMappingConfigs(newConfigs);
        setSuggestionCount(autoSuggestedCount);
        setShowNotification(true);
        notify(`Planilha carregada e mapeada! ${autoSuggestedCount} correspondências encontradas.`, "success");
      } catch (err) {
        console.error("Erro ao ler arquivo:", err);
        notify("Não foi possível processar o arquivo. Verifique se não está corrompido.", "error");
      }
    };
    reader.readAsBinaryString(file);
  };

  // Load Example Datasets
  const loadPreset = (type: 'hardware' | 'academy') => {
    resetMapping();
    if (type === 'hardware') {
      const cols = Object.keys(SAMPLE_TECH_STORE_ROWS[0]);
      setRawRows(SAMPLE_TECH_STORE_ROWS);
      setUploadedColumns(cols);
      setUploadedFileName('modelo_vendas_techstore.xlsx');

      // Setup optimal presets
      const hardwareConfigs = TARGET_COLUMNS.map((target) => {
        let sourceType: SourceType = 'column';
        let sourceColumn = '';
        let fixedValue = '';
        let concatColumns: string[] = [];
        let transformations: TransformationType[] = getSuggestedTransformations(target.key);

        if (target.key === 'id_da_venda') sourceColumn = 'Código';
        else if (target.key === 'data_do_pagamento') sourceColumn = 'Data_Compra';
        else if (target.key === 'nome_do_produto') sourceColumn = 'Produto_Nome';
        else if (target.key === 'id_do_sku') sourceColumn = 'Ref_SKU';
        else if (target.key === 'nome_do_sku') sourceColumn = 'Ref_SKU';
        else if (target.key === 'valor_unitario') sourceColumn = 'Preço';
        else if (target.key === 'quantidade') sourceColumn = 'Qtd';
        else if (target.key === 'nome_do_cliente') {
          sourceColumn = 'Cliente_Nome';
          transformations = ['remover_espacos', 'maiusculo'];
        }
        else if (target.key === 'email_do_cliente') sourceColumn = 'Cliente_Email';
        else if (target.key === 'telefone_do_cliente') {
          sourceColumn = 'Apenas_Celular';
          transformations = ['remover_especiais'];
        }
        else if (target.key === 'documento_do_cliente') {
          sourceColumn = 'CPF_CNPJ';
          transformations = ['formatar_cpf'];
        }
        else if (target.key === 'endereco_rua') sourceColumn = 'Logradouro';
        else if (target.key === 'endereco_numero') sourceColumn = 'Num';
        else if (target.key === 'endereco_bairro') sourceColumn = 'Bairro';
        else if (target.key === 'endereco_cidade') {
          // showcase fixed value or map column directly
          sourceColumn = 'Cidade_UF'; 
          transformations = ['remover_espacos'];
        }
        else if (target.key === 'endereco_cep') {
          sourceColumn = 'CEP';
          transformations = ['formatar_cep'];
        }
        else if (target.key === 'moeda') {
          sourceType = 'fixed';
          fixedValue = 'BRL';
        }
        else if (target.key === 'tipo_de_cliente') {
          sourceType = 'fixed';
          fixedValue = 'Pessoa Física';
        }
        else if (target.key === 'endereco_pais') {
          sourceType = 'fixed';
          fixedValue = 'Brasil';
        } else {
          sourceType = 'fixed';
          fixedValue = '';
        }

        return {
          targetKey: target.key,
          sourceType,
          sourceColumn,
          fixedValue,
          concatColumns,
          concatSeparator: ' ',
          transformations,
        };
      });

      setMappingConfigs(hardwareConfigs);
      notify("Preset 'TechStore' carregado com sucesso. Examine as correspondências e transformações!", "success");
    } else {
      const cols = Object.keys(SAMPLE_WEB_ACADEMY_ROWS[0]);
      setRawRows(SAMPLE_WEB_ACADEMY_ROWS);
      setUploadedColumns(cols);
      setUploadedFileName('modelo_matriculas_webacademy.xlsx');

      // Setup complex presets (including field concatenations)
      const academyConfigs = TARGET_COLUMNS.map((target) => {
        let sourceType: SourceType = 'column';
        let sourceColumn = '';
        let fixedValue = '';
        let concatColumns: string[] = [];
        let transformations: TransformationType[] = getSuggestedTransformations(target.key);

        if (target.key === 'id_da_venda') sourceColumn = 'id_compra';
        else if (target.key === 'data_do_pagamento') sourceColumn = 'data_pagto';
        else if (target.key === 'nome_do_produto') sourceColumn = 'curso_comprado';
        else if (target.key === 'id_do_sku') sourceColumn = 'cod_sku';
        else if (target.key === 'nome_do_sku') sourceColumn = 'cod_sku';
        else if (target.key === 'valor_unitario') sourceColumn = 'valor';
        else if (target.key === 'quantidade') sourceColumn = 'quantidade_itens';
        else if (target.key === 'nome_do_cliente') {
          // Showcase dynamic column concatenation!
          sourceType = 'concat';
          concatColumns = ['nome_completo', 'sobrenome'];
          transformations = ['remover_espacos', 'maiusculo'];
        }
        else if (target.key === 'email_do_cliente') {
          sourceColumn = 'email_contato';
          transformations = ['minusculo', 'remover_espacos'];
        }
        else if (target.key === 'telefone_do_cliente') {
          sourceColumn = 'telefone_ddd';
          transformations = ['remover_especiais'];
        }
        else if (target.key === 'documento_do_cliente') {
          sourceColumn = 'documento_identificação';
          transformations = ['formatar_cnpj']; // Demonstrates auto CNPJ support too
        }
        else if (target.key === 'endereco_rua') sourceColumn = 'rua_entrega';
        else if (target.key === 'endereco_numero') sourceColumn = 'numero_casa';
        else if (target.key === 'endereco_complemento') sourceColumn = 'complemento_apto';
        else if (target.key === 'endereco_cidade') sourceColumn = 'cidade_cadastro';
        else if (target.key === 'endereco_estado') {
          sourceColumn = 'estado_cadastro';
          transformations = ['maiusculo'];
        }
        else if (target.key === 'endereco_cep') {
          sourceColumn = 'cep_entrega';
          transformations = ['formatar_cep'];
        }
        else if (target.key === 'moeda') {
          sourceType = 'fixed';
          fixedValue = 'BRL';
        }
        else if (target.key === 'tipo_de_cliente') {
          sourceType = 'fixed';
          fixedValue = 'Pessoa Física';
        }
        else if (target.key === 'endereco_pais') {
          sourceType = 'fixed';
          fixedValue = 'Brasil';
        } else {
          sourceType = 'fixed';
          fixedValue = '';
        }

        return {
          targetKey: target.key,
          sourceType,
          sourceColumn,
          fixedValue,
          concatColumns,
          concatSeparator: ' ',
          transformations,
        };
      });

      setMappingConfigs(academyConfigs);
      notify("Preset 'WebAcademy' carregado. Nota: O Nome do Cliente foi concatenado automaticamente!", "success");
    }
  };

  const handleRemoveFile = () => {
    setUploadedFileName(null);
    setRawRows([]);
    setUploadedColumns([]);
    resetMapping();
    setSuggestionCount(0);
    setShowNotification(false);
    notify("Arquivo removido e configurações reiniciadas.", "info");
  };

  // Modify individual column mapping
  const updateMappingField = (targetKey: string, updates: Partial<ColumnMappingConfig>) => {
    setMappingConfigs((prev) =>
      prev.map((cfg) => (cfg.targetKey === targetKey ? { ...cfg, ...updates } : cfg))
    );
  };

  // Toggle single transformation
  const toggleTransformation = (targetKey: string, transType: TransformationType) => {
    setMappingConfigs((prev) =>
      prev.map((cfg) => {
        if (cfg.targetKey !== targetKey) return cfg;
        const index = cfg.transformations.indexOf(transType);
        let updated: TransformationType[];
        if (index > -1) {
          updated = cfg.transformations.filter((t) => t !== transType);
        } else {
          updated = [...cfg.transformations, transType];
        }
        return { ...cfg, transformations: updated };
      })
    );
  };

  const generateCsvBlob = (resultRows: any[]): Blob => {
    const worksheet = XLSX.utils.json_to_sheet(resultRows, {
      header: TARGET_COLUMNS.map((c) => c.key),
    });
    const csvContent = XLSX.utils.sheet_to_csv(worksheet, { FS: ';' });
    return new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
  };

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const redownloadFromStorage = async (item: HistoryItem) => {
    const { data, error } = await supabase.storage.from('exports').createSignedUrl(item.filePath, 60);
    if (error || !data) { notify('Erro ao baixar documento.', 'error'); return; }
    const link = document.createElement('a');
    link.href = data.signedUrl;
    link.setAttribute('download', item.filename);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    notify('Planilha baixada novamente!', 'success');
  };

  // Export spreadsheet triggered by "Gerar Planilha"
  const exportResult = () => {
    if (rawRows.length === 0) {
      notify("Carregue uma planilha primeiro para converter.", "error");
      return;
    }
    // Initialize/reset modal options
    setClientNameInput('');
    setIsExportModalOpen(true);
  };

  const confirmExport = async () => {
    if (!clientNameInput.trim()) {
      notify("Por favor, informe o nome do cliente.", "error");
      return;
    }

    try {
      const resultRows = rawRows.map((row, index) => {
        const transformedRow = processarLinha(row, mappingConfigs, index);
        const orderedRow: any = {};
        TARGET_COLUMNS.forEach((col) => {
          orderedRow[col.key] = sanitizeCsvValue(transformedRow[col.key] || '');
        });
        return orderedRow;
      });

      const cleanClientName = clientNameInput.trim().toUpperCase();
      const filename = `Planilha de Importação- ${monthNameInput} - ${cleanClientName}.csv`;
      const id = crypto.randomUUID();
      const filePath = `${user.id}/${id}.csv`;

      const blob = generateCsvBlob(resultRows);
      downloadBlob(blob, filename);

      const { error: uploadError } = await supabase.storage
        .from('exports')
        .upload(filePath, blob, { contentType: 'text/csv;charset=utf-8' });

      if (uploadError) {
        console.error('Storage upload error:', uploadError);
        notify('Planilha gerada, mas erro ao salvar no histórico.', 'error');
        return;
      }

      const historyItem: HistoryItem = {
        id,
        filename,
        clientName: cleanClientName,
        monthName: monthNameInput,
        timestamp: new Date().toLocaleString('pt-BR'),
        filePath,
      };

      supabase.from('export_history').insert({
        id: historyItem.id,
        user_id: user.id,
        filename: historyItem.filename,
        client_name: historyItem.clientName,
        month_name: historyItem.monthName,
        timestamp: historyItem.timestamp,
        file_path: historyItem.filePath,
      }).then(({ error }) => { if (error) console.error(error); });

      setExportHistory((prev) => [historyItem, ...prev]);
      setIsExportModalOpen(false);
      notify("Planilha gerada e salva no histórico!", "success");
    } catch (err) {
      console.error("Erro ao exportar:", err);
      notify("Ocorreu um erro ao estruturar o arquivo de exportação.", "error");
    }
  };

  return (
    <div className="min-h-screen bg-[#0F0E1A] text-[#F4F5FC] selection:bg-[#6366F1] selection:text-white pb-16">
      
      {/* Visual background atmospheric effects */}
      <div className="absolute top-0 left-0 right-0 h-96 bg-gradient-to-b from-[#252344]/15 to-transparent pointer-events-none" />
      <div className="absolute top-1/4 right-0 w-96 h-96 bg-[#7C3AED]/5 rounded-full filter blur-[120px] pointer-events-none" />
      <div className="absolute bottom-1/4 left-0 w-96 h-96 bg-[#6366F1]/5 rounded-full filter blur-[120px] pointer-events-none" />

      {/* Internal Navigation / Header Area */}
      <header className="relative border-b border-[#2E2C4A] bg-[#0F0E1A]/80 backdrop-blur-md sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col sm:flex-row items-center justify-between gap-4">
          
          {/* Logo / Title */}
          <button
            onClick={handleRemoveFile}
            className="flex items-center gap-3 group cursor-pointer"
            title="Voltar à tela inicial"
          >
            <div className="h-10 w-10 rounded-xl bg-gradient-to-tr from-[#7C3AED] to-[#6366F1] flex items-center justify-center shadow-lg shadow-[#7C3AED]/20 group-hover:shadow-[#7C3AED]/40 transition-shadow">
              <FileSpreadsheet className="h-5.5 w-5.5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-lg font-bold tracking-tight text-[#F4F5FC] group-hover:text-white transition-colors">Conversor de Planilhas Fiscal</h1>
                <span className="text-[10px] bg-[#1A1830] text-[#8B80FF] px-2 py-0.5 rounded-full border border-[#2E2C4A]">Uso Interno</span>
              </div>
            </div>
          </button>

          {/* Action Header / Actions */}
          <div className="flex items-center gap-3">
            {/* User avatar */}
            <div className="relative">
              <button
                onClick={() => setShowUserModal((v) => !v)}
                className="flex items-center gap-2 cursor-pointer group"
                title="Conta"
              >
                {profileName && (
                  <span className="text-xs font-semibold text-[#A1A6C4] group-hover:text-[#F4F5FC] transition-colors hidden sm:block">{profileName}</span>
                )}
                <div className="h-9 w-9 rounded-full bg-gradient-to-tr from-[#6366F1] to-[#7C3AED] flex items-center justify-center shadow-md shadow-[#6366F1]/20 group-hover:shadow-[#6366F1]/40 border-2 border-[#2E2C4A] group-hover:border-[#6366F1] transition-all">
                  <User className="h-4.5 w-4.5 text-white" />
                </div>
              </button>

              {showUserModal && (
                <>
                  <div className="fixed inset-0 z-40" onClick={() => setShowUserModal(false)} />
                  <div className="absolute right-0 top-12 z-50 w-64 bg-[#141229] border border-[#2E2C4A] rounded-2xl shadow-2xl shadow-black/40 p-4 animate-in fade-in zoom-in-95 duration-150">
                    <div className="flex flex-col items-center gap-3 pb-4 border-b border-[#2E2C4A]">
                      <div className="h-14 w-14 rounded-full bg-gradient-to-tr from-[#6366F1] to-[#7C3AED] flex items-center justify-center shadow-lg shadow-[#6366F1]/20">
                        <User className="h-7 w-7 text-white" />
                      </div>
                      <div className="text-center w-full">
                        {editingName ? (
                          <div className="flex items-center gap-1.5 mt-1">
                            <input
                              autoFocus
                              value={nameInput}
                              onChange={(e) => setNameInput(e.target.value)}
                              onKeyDown={(e) => { if (e.key === 'Enter') saveProfileName(); if (e.key === 'Escape') setEditingName(false); }}
                              placeholder="Seu nome"
                              className="flex-1 bg-[#0F0E1A] border border-[#6366F1] rounded-lg px-2 py-1 text-xs text-[#F4F5FC] outline-none w-full"
                            />
                            <button
                              onClick={saveProfileName}
                              disabled={savingName || !nameInput.trim()}
                              className="px-2 py-1 bg-[#6366F1] text-white text-[10px] font-bold rounded-lg disabled:opacity-50 cursor-pointer"
                            >
                              {savingName ? '...' : 'OK'}
                            </button>
                            <button
                              onClick={() => setEditingName(false)}
                              className="px-2 py-1 bg-[#1A1830] text-[#A1A6C4] text-[10px] rounded-lg cursor-pointer border border-[#2E2C4A]"
                            >
                              ✕
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center justify-center gap-1.5">
                            <p className="text-sm font-bold text-[#F4F5FC]">
                              {profileName ?? 'Sem nome'}
                            </p>
                            <button
                              onClick={() => { setNameInput(profileName ?? ''); setEditingName(true); }}
                              className="text-[#A1A6C4] hover:text-[#8B80FF] transition-colors cursor-pointer"
                              title="Editar nome"
                            >
                              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          </div>
                        )}
                        <p className="text-[11px] text-[#A1A6C4] break-all mt-0.5">{user.email}</p>
                      </div>
                    </div>
                    <button
                      onClick={() => { setShowUserModal(false); supabase.auth.signOut(); }}
                      className="mt-3 w-full flex items-center justify-center gap-2 px-4 py-2 bg-[#1A1830] hover:bg-red-500/10 text-[#A1A6C4] hover:text-red-400 border border-[#2E2C4A] hover:border-red-500/40 rounded-xl text-xs font-semibold transition-all cursor-pointer"
                    >
                      <LogOut className="h-3.5 w-3.5" />
                      Sair da Conta
                    </button>
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="relative max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-6">

        {/* Global Notifications Panel */}
        {alertMsg && (
          <div className={`p-4 mb-6 rounded-xl border text-sm flex items-start gap-3 animate-fade-in transition-all duration-300 ${
            alertMsg.type === 'success'
              ? 'bg-[#22C55E]/10 border-[#22C55E]/30 text-[#22C55E]'
              : alertMsg.type === 'error'
              ? 'bg-[#FF3D00]/10 border-[#FF3D00]/30 text-[#FF3D00]'
              : 'bg-[#3B82F6]/10 border-[#3B82F6]/30 text-[#3B82F6]'
          }`}>
            <AlertCircle className="h-5 w-5 shrink-0 mt-0.5" />
            <div>
              <p className="font-semibold">{alertMsg.type === 'success' ? 'Sucesso' : alertMsg.type === 'error' ? 'Erro' : 'Informação'}</p>
              <p className="opacity-90">{alertMsg.text}</p>
            </div>
          </div>
        )}

        {/* INITIAL STATE / EMPTY PLACEHOLDER CONTAINER */}
        {rawRows.length === 0 ? (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-stretch pt-4">
            
            {/* Main Upload Dropzone */}
            <div className="col-span-1 lg:col-span-8">
              <div
                id="dropzone-container"
                onDragEnter={handleDrag}
                onDragOver={handleDrag}
                onDragLeave={handleDrag}
                onDrop={handleDrop}
                className={`relative group bg-[#141229] border-2 border-dashed rounded-3xl p-10 sm:p-16 text-center flex flex-col items-center justify-center min-h-[440px] transition-all duration-300 cursor-pointer ${
                  dragActive
                    ? 'border-[#8B80FF] bg-[#1A1830] scale-[0.99] shadow-xl shadow-indigo-500/5'
                    : 'border-[#2E2C4A] hover:border-[#6366F1]/50 hover:bg-[#1A1830]/40'
                }`}
                onClick={() => fileInputRef.current?.click()}
              >
                
                {/* Visual subtle grid background inside active dropzone */}
                <div className="absolute inset-0 bg-[linear-gradient(to_right,#1A1830_1px,transparent_1px),linear-gradient(to_bottom,#1A1830_1px,transparent_1px)] bg-[size:24px_24px] pointer-events-none opacity-20 rounded-3xl" />

                <input
                  type="file"
                  ref={fileInputRef}
                  onChange={handleFileChange}
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                />

                <div className="h-16 w-16 rounded-2xl bg-[#1A1830] border border-[#2E2C4A] group-hover:border-[#6366F1]/40 flex items-center justify-center transition-all duration-300 mb-6 relative shadow-inner">
                  <Upload className="h-7 w-7 text-[#8B80FF] group-hover:scale-110 transition-transform duration-300" />
                  <div className="absolute -top-1 -right-1 h-3.5 w-3.5 bg-[#6366F1] rounded-full animate-pulse" />
                </div>

                <h3 className="text-xl font-bold text-[#F4F5FC] mb-1">Upload da planilha de vendas</h3>
                <p className="text-[#A1A6C4] text-sm max-w-md mb-8">
                  Arraste seu arquivo <strong className="text-[#F4F5FC]">.xlsx, .xls ou .csv</strong> para esta área ou clique para selecionar do dispositivo.
                </p>


              </div>
            </div>

            {/* Side info, design references, and Demo preloads */}
            <div className="col-span-1 lg:col-span-4 flex flex-col gap-6">
              
              {/* Export History Widget */}
              <div className="bg-[#141229] border border-[#2E2C4A] p-6 rounded-3xl flex-1 flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between mb-3 border-b border-[#2E2C4A]/40 pb-2">
                    <div className="flex items-center gap-2">
                      <History className="h-4 w-4 text-[#8B80FF]" />
                      <h4 className="text-xs font-bold text-[#F4F5FC] uppercase tracking-wider">Histórico de Planilhas</h4>
                    </div>
                    {exportHistory.length > 0 && (
                      <button
                        onClick={() => {
                          if (confirm("Tem certeza que deseja limpar todo o histórico?")) {
                            supabase.from('export_history').delete().eq('user_id', user.id).then(({ error }) => { if (error) console.error(error); });
                            setExportHistory([]);
                          }
                        }}
                        className="text-[10px] text-red-400 hover:text-red-300 transition-colors flex items-center gap-1 cursor-pointer"
                      >
                        <Trash2 className="h-3 w-3" />
                        Limpar
                      </button>
                    )}
                  </div>

                  {exportHistory.length === 0 ? (
                    <div className="flex flex-col items-center justify-center text-center py-6 text-[#A1A6C4] space-y-2">
                      <FileSpreadsheet className="h-8 w-8 text-[#2E2C4A]" />
                      <p className="text-xs leading-relaxed">
                        Nenhuma planilha foi gerada ainda.<br/>Mapeie e exporte seus arquivos para preencher este histórico.
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-3 max-h-[190px] overflow-y-auto pr-1">
                      {exportHistory.map((item) => (
                        <div key={item.id} className="bg-[#0F0E1A]/85 border border-[#2E2C4A] p-3 rounded-xl flex items-center justify-between gap-3 group hover:border-[#6366F1]/50 transition-colors">
                          <div className="min-w-0 flex-1">
                            <div className="text-xs font-bold text-[#F4F5FC] truncate" title={item.filename}>
                              {item.filename}
                            </div>
                            <div className="text-[10px] text-[#A1A6C4] mt-0.5 flex flex-wrap gap-x-2 gap-y-0.5 font-sans">
                              <span>👤 {item.clientName}</span>
                              <span>📅 {item.monthName}</span>
                            </div>
                          </div>
                          <button
                            onClick={(e) => {
                              if (openHistoryMenu === item.id) {
                                setOpenHistoryMenu(null);
                                setMenuPosition(null);
                              } else {
                                const rect = (e.currentTarget as HTMLElement).getBoundingClientRect();
                                setMenuPosition({ top: rect.bottom + 6, right: window.innerWidth - rect.right });
                                setOpenHistoryMenu(item.id);
                              }
                            }}
                            className="p-1.5 bg-[#1A1830] hover:bg-[#252344] text-[#A1A6C4] hover:text-white rounded-lg border border-[#2E2C4A] transition-all cursor-pointer shrink-0"
                            title="Opções"
                          >
                            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><circle cx="5" cy="12" r="2"/><circle cx="12" cy="12" r="2"/><circle cx="19" cy="12" r="2"/></svg>
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

              </div>

            </div>
          </div>
        ) : (
          
          /* ACTIVE CONVERTING EXPERIENCE PANEL */
          <div className="space-y-8 pt-2">
            
            {/* Quick Status Bar */}
            <div className="bg-[#141229] border border-[#2E2C4A] rounded-2.5xl p-4 sm:p-5 flex flex-col md:flex-row items-center justify-between gap-4">
              <div className="flex flex-wrap items-center gap-3 sm:gap-4">
                <div className="h-9 w-9 rounded-lg bg-[#22C55E]/10 border border-[#22C55E]/20 flex items-center justify-center">
                  <FileSpreadsheet className="h-5 w-5 text-[#22C55E]" />
                </div>
                <div>
                  <div className="text-xs text-[#A1A6C4] font-medium">Arquivo em processo:</div>
                  <div className="text-sm font-bold text-[#F4F5FC] flex items-center gap-2">
                    {uploadedFileName}
                    <span className="text-[10px] font-normal px-2 py-0.5 rounded bg-[#1A1830] border border-[#2E2C4A] text-[#8B80FF]">
                      {rawRows.length} linhas importadas
                    </span>
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2 shrink-0">
                <button
                  id="reload-file-btn"
                  onClick={() => fileInputRef.current?.click()}
                  className="bg-[#1A1830] hover:bg-[#252344] border border-[#2E2C4A] px-3.5 py-1.8 text-xs font-semibold text-[#F4F5FC] rounded-lg transition-all duration-150 flex items-center gap-1.5"
                  title="Trocar Arquivo"
                >
                  <RefreshCw className="h-3.5 w-3.5" />
                  Substituir
                </button>
                <button
                  id="reset-mapping-btn"
                  onClick={resetMapping}
                  className="bg-[#1A1830] hover:bg-[#252344] border border-[#2E2C4A] px-3.5 py-1.8 text-xs font-semibold text-rose-400 rounded-lg transition-all duration-150 flex items-center gap-1.5"
                  title="Resetar todos os mapeamentos corporais"
                >
                  Limpar
                </button>
              </div>
            </div>

            {/* Conversion Settings Panel */}
            <div className="grid grid-cols-1 gap-8 items-start">

              {/* MAPPING WORKSPACE PANELS AND FORMS */}
              <div className="bg-[#141229] border border-[#2E2C4A] rounded-3xl p-6 flex flex-col justify-between">
                
                {/* Panel Header */}
                <div>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-5 border-b border-[#2E2C4A] pb-4">
                    <div>
                      <h3 className="text-md font-bold text-[#F4F5FC] flex items-center gap-2">
                        <Settings className="h-4 w-4 text-[#8B80FF]" />
                        Campos de Entrada e Destino
                      </h3>
                      <p className="text-xs text-[#A1A6C4] mt-0.5">Associe as colunas importadas às colunas obrigatórias do faturamento fiscal</p>
                    </div>
                  </div>

                  {/* Column Configuration Scrollable Checklist */}
                  <div className="space-y-4 max-h-[640px] overflow-y-auto pr-1">
                    {TARGET_COLUMNS.map((col) => {
                      const cfg = mappingConfigs.find((c) => c.targetKey === col.key);
                      if (!cfg) return null;

                      const isLocked = col.key === 'tipo_de_cliente';
                      const isExpanded = expandedField === col.key && !isLocked;

                      // Status bullet color calculations for the dashboard look
                      let statusBulletColor = "bg-[#FF3D00]"; // red / default unmapped
                      if (isLocked) {
                        statusBulletColor = "bg-[#3B82F6]"; // blue for locked / autogenerated
                      } else if (cfg.sourceType === 'column' && cfg.sourceColumn) {
                        statusBulletColor = "bg-[#22C55E]"; // green
                      } else if (cfg.sourceType === 'fixed' && cfg.fixedValue !== undefined) {
                        statusBulletColor = "bg-[#FBBF24]"; // amber
                      } else if (cfg.sourceType === 'concat' && cfg.concatColumns.length > 0) {
                        statusBulletColor = "bg-[#7C3AED]"; // purple / concatenated
                      }

                      return (
                        <div
                          key={col.key}
                          id={`field-container-${col.key}`}
                          className={`glass rounded-xl p-4 transition-all duration-200 mapping-row border ${
                            isExpanded ? 'border-[#6366F1] bg-[#1A1830]/40 accent-glow' : 'border-[#2E2C4A]'
                          }`}
                        >
                          {/* Summary Header */}
                          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            
                            {/* Label and Info with interactive status bullet */}
                            <div className="flex items-start gap-3">
                              <div className="flex items-center gap-1.5 shrink-0 mt-1">
                                <span className="h-5 w-5 rounded bg-[#1A1830] border border-[#2E2C4A] flex items-center justify-center text-[9px] text-[#8B80FF] font-mono leading-none">
                                  {TARGET_COLUMNS.findIndex((x) => x.key === col.key) + 1}
                                </span>
                                <div className={`w-2.5 h-2.5 rounded-full ${statusBulletColor} animate-pulse shadow-sm`} title={isLocked ? "Cálculo Inteligente Ativo" : "Status do mapeamento"} />
                              </div>
                              <div>
                                <div className="flex items-center gap-2 flex-wrap">
                                  <h4 className="text-xs font-bold text-[#F4F5FC]">{col.label}</h4>
                                  <code className="text-[9px] text-[#8B80FF] bg-[#1A1830] px-1 rounded font-mono border border-[#2E2C4A]/60 font-medium">col: {col.key}</code>
                                  
                                  {isLocked ? (
                                    <span className="text-[9px] bg-[#3B82F6]/10 text-[#3B82F6] border border-[#3B82F6]/20 px-1.5 py-0.2 rounded-md font-medium">
                                      🔒 Vinculado ao Documento
                                    </span>
                                  ) : (
                                    <>
                                      {cfg.sourceType === 'column' && cfg.sourceColumn && (
                                        <span className="text-[9px] bg-[#22C55E]/10 text-[#22C55E] border border-[#22C55E]/20 px-1.5 py-0.2 rounded-md font-medium">
                                          {cfg.sourceColumn}
                                        </span>
                                      )}
                                      {cfg.sourceType === 'fixed' && (
                                        <span className="text-[9px] bg-amber-500/10 text-amber-400 border border-amber-500/20 px-1.5 py-0.2 rounded-md font-mono">
                                          Fixo: &quot;{cfg.fixedValue || 'vazio'}&quot;
                                        </span>
                                      )}
                                      {cfg.sourceType === 'concat' && (
                                        <span className="text-[9px] bg-[#7C3AED]/15 text-[#8B80FF] border border-[#7C3AED]/30 px-1.5 py-0.2 rounded-md">
                                          União de ({cfg.concatColumns.length} colunas)
                                        </span>
                                      )}
                                    </>
                                  )}
                                </div>
                                <p className="text-[10px] text-[#A1A6C4] mt-0.5">
                                  {isLocked ? "Calculado automaticamente: pessoa_fisica para CPF, pessoa_juridica para CNPJ e internacional para outros/aleatórios." : col.description}
                                </p>
                              </div>
                            </div>

                            {/* Toggle expanded parameters */}
                            {isLocked ? (
                              <div className="bg-[#1A1830] text-blue-400 border border-blue-500/20 px-2.5 py-1.5 rounded-lg text-[10px] font-semibold flex items-center gap-1 select-none">
                                🔒 Auto-Calculado
                              </div>
                            ) : (
                              <button
                                id={`expand-btn-${col.key}`}
                                onClick={() => setExpandedField(isExpanded ? null : col.key)}
                                className="bg-[#1A1830] hover:bg-[#252344] text-[#A1A6C4] hover:text-[#F4F5FC] border border-[#2E2C4A] px-2.5 py-1.5 rounded-lg text-xs flex items-center gap-1 transition-all cursor-pointer shrink-0"
                              >
                                <Settings className="h-3 w-3" />
                                <span className="text-[10px] font-semibold">Configurar</span>
                                {isExpanded ? <ChevronDown className="h-3 w-3" /> : <ChevronRight className="h-3 w-3" />}
                              </button>
                            )}
                          </div>

                          {/* Extended Mapping options (only if isExpanded is true) */}
                          {isExpanded && (
                            <div className="mt-4 pt-4 border-t border-[#2E2C4A] space-y-4 animate-fade-in text-xs">
                              
                              {/* Source mapping buttons */}
                              <div>
                                <label className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider block mb-1.5">Origem da Informação</label>
                                <div className="grid grid-cols-3 gap-2">
                                  <button
                                    type="button"
                                    onClick={() => updateMappingField(col.key, { sourceType: 'column' })}
                                    className={`px-2 py-1.5 text-[11px] font-semibold rounded-lg border transition-all cursor-pointer ${
                                      cfg.sourceType === 'column'
                                        ? 'bg-[#6366F1]/10 border-[#6366F1] text-[#8B80FF]'
                                        : 'bg-[#141229] border-[#2E2C4A] text-[#A1A6C4] hover:bg-[#1A1830]'
                                    }`}
                                  >
                                    Coluna do Cliente
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => updateMappingField(col.key, { sourceType: 'fixed' })}
                                    className={`px-2 py-1.5 text-[11px] font-semibold rounded-lg border transition-all cursor-pointer ${
                                      cfg.sourceType === 'fixed'
                                        ? 'bg-[#6366F1]/10 border-[#6366F1] text-[#8B80FF]'
                                        : 'bg-[#141229] border-[#2E2C4A] text-[#A1A6C4] hover:bg-[#1A1830]'
                                    }`}
                                  >
                                    Valor Constante
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() => updateMappingField(col.key, { sourceType: 'concat' })}
                                    className={`px-2 py-1.5 text-[11px] font-semibold rounded-lg border transition-all cursor-pointer ${
                                      cfg.sourceType === 'concat'
                                        ? 'bg-[#6366F1]/10 border-[#6366F1] text-[#8B80FF]'
                                        : 'bg-[#141229] border-[#2E2C4A] text-[#A1A6C4] hover:bg-[#1A1830]'
                                    }`}
                                  >
                                    Concatenar Colunas
                                  </button>
                                </div>
                              </div>

                              {/* SOURCE Type SPECIFICS */}
                              {cfg.sourceType === 'column' && (
                                <div>
                                  <label className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider block mb-1.5">Escolha a coluna correspondente</label>
                                  <select
                                    id={`select-column-${col.key}`}
                                    value={cfg.sourceColumn}
                                    onChange={(e) => updateMappingField(col.key, { sourceColumn: e.target.value })}
                                    className="w-full bg-[#1A1830] border border-[#2E2C4A] text-xs text-[#F4F5FC] px-3 py-2 rounded-lg focus:outline-none focus:border-[#6366F1]"
                                  >
                                    <option value="">-- Ignorar ou deixar vazio --</option>
                                    {uploadedColumns.map((sourceCol) => (
                                      <option key={sourceCol} value={sourceCol}>
                                        {sourceCol}
                                      </option>
                                    ))}
                                  </select>
                                </div>
                              )}

                              {cfg.sourceType === 'fixed' && (
                                <div>
                                  <label className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider block mb-1.5">Digite o valor fixo a ser preenchido</label>
                                  <div className="relative">
                                    <input
                                      id={`input-fixed-${col.key}`}
                                      type="text"
                                      value={cfg.fixedValue}
                                      onChange={(e) => updateMappingField(col.key, { fixedValue: e.target.value })}
                                      className="w-full bg-[#1A1830] border border-[#2E2C4A] text-xs text-[#F4F5FC] px-3 py-2 rounded-lg focus:outline-none focus:border-[#6366F1] pr-28 placeholder-[#A1A6C4]/30"
                                      placeholder={col.placeholder || "Valor fiscal regular"}
                                    />
                                    {col.defaultFixed && cfg.fixedValue !== col.defaultFixed && (
                                      <button
                                        type="button"
                                        onClick={() => updateMappingField(col.key, { fixedValue: col.defaultFixed })}
                                        className="absolute right-2 top-1.5 px-2 py-1 bg-[#141229] border border-[#2E2C4A] text-[9px] text-[#8B80FF] rounded hover:text-white transition-colors cursor-pointer font-medium"
                                      >
                                        Preencher Padrão ({col.defaultFixed})
                                      </button>
                                    )}
                                  </div>
                                </div>
                              )}

                              {cfg.sourceType === 'concat' && (
                                <div className="space-y-3">
                                  <div>
                                    <label className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider block mb-1.5">Escolha as colunas a serem unidas</label>
                                    <div className="grid grid-cols-2 md:grid-cols-3 gap-2 bg-[#1A1830] p-3 rounded-xl border border-[#2E2C4A] max-h-[145px] overflow-y-auto">
                                      {uploadedColumns.map((sourceCol) => {
                                        const isChecked = cfg.concatColumns.includes(sourceCol);
                                        return (
                                          <label key={sourceCol} className="flex items-center gap-1.5 bg-[#141229] p-2 rounded border border-[#2E2C4A]/50 hover:border-[#6366F1]/50 text-[10px] text-[#F4F5FC] cursor-pointer select-none">
                                            <input
                                              type="checkbox"
                                              checked={isChecked}
                                              onChange={() => {
                                                const updatedConcat = isChecked
                                                  ? cfg.concatColumns.filter((c) => c !== sourceCol)
                                                  : [...cfg.concatColumns, sourceCol];
                                                updateMappingField(col.key, { concatColumns: updatedConcat });
                                              }}
                                              className="rounded border-[#2E2C4A] text-[#6366F1] focus:ring-[#6366F1] bg-[#141229] h-3 w-3"
                                            />
                                            <span className="truncate">{sourceCol}</span>
                                          </label>
                                        );
                                      })}
                                    </div>
                                  </div>

                                  <div>
                                    <label className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider block mb-1">Delimitador de Junção (Separador)</label>
                                    <div className="flex items-center gap-2">
                                      <input
                                        id={`input-separator-${col.key}`}
                                        type="text"
                                        value={cfg.concatSeparator}
                                        onChange={(e) => updateMappingField(col.key, { concatSeparator: e.target.value })}
                                        className="w-20 bg-[#1A1830] border border-[#2E2C4A] text-xs text-[#F4F5FC] px-2.5 py-1.5 rounded-lg focus:outline-none focus:border-[#6366F1] text-center"
                                        placeholder="espaço"
                                      />
                                      <span className="text-[10px] text-[#A1A6C4]">Espaço em branco, hífen (&quot;-&quot;) ou vírgula.</span>
                                    </div>
                                  </div>
                                </div>
                              )}

                              {/* TRANSFORMATIONS PIPELINE ACCORDION */}
                              <div className="space-y-2 pt-2 border-t border-[#2E2C4A]/50">
                                <div className="flex items-center gap-1.5">
                                  <label className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider">Pipeline de Formatação / Limpeza</label>
                                  <span className="text-[8px] bg-[#1A1830] text-[#A1A6C4] px-1.5 py-0.2 rounded font-mono">Executados em ordem</span>
                                </div>
                                <div className="flex flex-wrap gap-1.5">
                                  {[
                                    { type: 'remover_especiais', label: 'Remover Especiais', desc: 'Remove símbolos/pontuações mantendo apenas alfa-numéricos' },
                                    { type: 'remover_espacos', label: 'Remover Espaços Extras', desc: 'Remove espaços duplos e colaterais' },
                                    { type: 'maiusculo', label: 'MAIÚSCULO', desc: 'Converte todas as letras para caixa alta' },
                                    { type: 'minusculo', label: 'minúsculo', desc: 'Converte todas as letras para caixa baixa' },
                                    { type: 'formatar_cpf', label: 'Formatar CPF', desc: 'Ajusta formato numérico para XXX.XXX.XXX-XX' },
                                    { type: 'formatar_cnpj', label: 'Formatar CNPJ', desc: 'Ajusta formato corporativo para XX.XXX.XXX/XXXX-XX' },
                                    { type: 'formatar_cep', label: 'Formatar CEP', desc: 'Garante o formato postal brasileiro XXXXX-XXX' },
                                    { type: 'converter_data', label: 'Corrigir Data', desc: 'Busca datas no texto e padroniza para DD/MM/AAAA' },
                                  ].map((trans) => {
                                    const isActive = cfg.transformations.includes(trans.type as TransformationType);
                                    return (
                                      <button
                                        key={trans.type}
                                        type="button"
                                        onClick={() => toggleTransformation(col.key, trans.type as TransformationType)}
                                        className={`px-2 py-1 rounded-md border text-[10px] font-semibold flex items-center gap-1 transition-all cursor-pointer ${
                                          isActive
                                            ? 'bg-[#6366F1]/10 border-[#6366F1] text-[#8B80FF]'
                                            : 'bg-[#1A1830]/60 border-[#2E2C4A] text-[#A1A6C4] hover:border-[#6366F1]/40 hover:text-white'
                                        }`}
                                        title={trans.desc}
                                      >
                                        {isActive && <Check className="h-2.5 w-2.5 text-[#22C55E]" />}
                                        {trans.label}
                                      </button>
                                    );
                                  })}
                                </div>
                              </div>

                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>



              </div>

            </div>

            {/* REAL-TIME PREVIEW MODULE (Spans full widescreen 12-columns at the bottom for readability) */}
            <div className="w-full bg-[#141229] border border-[#2E2C4A] rounded-3xl p-6 flex flex-col min-h-[500px]">
              
              {/* Top Header Selector */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-5 border-b border-[#2E2C4A]/60 pb-4">
                <div>
                  <h3 className="text-sm font-bold text-[#F4F5FC] flex items-center gap-1.5">
                    <Eye className="h-4 w-4 text-[#8B80FF]" />
                    Visualização em Tempo Real (Lote de Conferência)
                  </h3>
                  <p className="text-[11px] text-[#A1A6C4] mt-0.5">As primeiras 8 linhas comparando origem com resultado do mapeamento</p>
                </div>

                <div className="flex bg-[#0F0E1A] p-0.8 rounded-lg border border-[#2E2C4A] shrink-0">
                  <button
                    onClick={() => setActivePreviewTab('result')}
                    className={`px-3 py-1 text-[11px] font-semibold rounded-md transition-all cursor-pointer ${
                      activePreviewTab === 'result'
                        ? 'bg-[#6366F1] text-white shadow'
                        : 'text-[#A1A6C4] hover:text-[#F4F5FC]'
                    }`}
                  >
                    💎 Resultado Traduzido (22 colunas)
                  </button>
                  <button
                    onClick={() => setActivePreviewTab('source')}
                    className={`px-3 py-1 text-[11px] font-semibold rounded-md transition-all cursor-pointer ${
                      activePreviewTab === 'source'
                        ? 'bg-[#6366F1] text-white shadow'
                        : 'text-[#A1A6C4] hover:text-[#F4F5FC]'
                    }`}
                  >
                    📂 Planilha Original Recebida
                  </button>
                </div>
              </div>

              {/* Dynamic comparison table wrapper with vertical & horizontal scrolling */}
              <div className="flex-1 overflow-auto bg-[#0F0E1A] border border-[#2E2C4A] rounded-2xl relative min-h-[380px]">
                
                {activePreviewTab === 'result' ? (
                  
                  /* RESULT PLANILHA VIEW: Exact output structure */
                  <div className="absolute inset-0 flex flex-col overflow-auto">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead className="bg-[#141229] border-b border-[#2E2C4A] sticky top-0 z-10">
                        <tr>
                          <th className="p-3 font-semibold text-[#A1A6C4] text-[9px] uppercase font-mono tracking-wider text-center border-r border-[#2E2C4A] bg-[#141229] min-w-[50px]">LID</th>
                          {TARGET_COLUMNS.map((col) => (
                            <th key={col.key} className="p-3 font-semibold text-[#F4F5FC] border-r border-[#2E2C4A] font-mono min-w-[170px] bg-[#141229]">
                              <div className="flex flex-col">
                                <span className="text-[11px]">{col.label}</span>
                                <span className="text-[9px] text-[#8B80FF] font-normal leading-tight lowercase font-sans">col: {col.key}</span>
                              </div>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rawRows.slice(0, 8).map((row, rIdx) => {
                          const processed = processarLinha(row, mappingConfigs, rIdx);
                          return (
                            <tr key={rIdx} className="border-b border-[#2E2C4A] hover:bg-[#1A1830]/40 transition-colors">
                              <td className="p-3 font-mono text-center text-[#8B80FF] bg-[#141229]/60 border-r border-[#2E2C4A]">{rIdx + 1}</td>
                              {TARGET_COLUMNS.map((col) => (
                                <td key={col.key} className="p-3 border-r border-[#2E2C4A]/60 max-w-[200px] truncate text-xs font-mono">
                                  {processed[col.key] !== undefined && processed[col.key] !== "" ? (
                                    <span className="text-[#F4F5FC]">{String(processed[col.key])}</span>
                                  ) : (
                                    <span className="text-[#A1A6C4]/30 italic text-[10px]">- vazio -</span>
                                  )}
                                </td>
                              ))}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    
                    {rawRows.length > 8 && (
                      <div className="p-3 bg-[#141229] text-center text-[11px] text-[#A1A6C4] border-t border-[#2E2C4A] sticky bottom-0 mt-auto">
                        Exibindo as primeiras 8 passagens de {rawRows.length} linhas calculadas.
                      </div>
                    )}
                  </div>

                ) : (

                  /* SOURCE PLANILHA VIEW: Raw columns uploaded */
                  <div className="absolute inset-0 flex flex-col overflow-auto">
                    <table className="w-full border-collapse text-left text-xs">
                      <thead className="bg-[#141229] border-b border-[#2E2C4A] sticky top-0 z-10">
                        <tr>
                          <th className="p-3 font-semibold text-[#A1A6C4] text-[9px] uppercase font-mono text-center border-r border-[#2E2C4A] bg-[#141229] min-w-[50px]">LID</th>
                          {uploadedColumns.map((col) => (
                            <th key={col} className="p-3 font-semibold text-[#F4F5FC] border-r border-[#2E2C4A] bg-[#141229] min-w-[140px] font-mono">
                              {col}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rawRows.slice(0, 8).map((row, rIdx) => (
                          <tr key={rIdx} className="border-b border-[#2E2C4A] hover:bg-[#1A1830]/40 transition-colors">
                            <td className="p-3 font-mono text-center text-[#8B80FF] bg-[#141229]/50 border-r border-[#2E2C4A]">{rIdx + 1}</td>
                            {uploadedColumns.map((col) => (
                              <td key={col} className="p-3 border-r border-[#2E2C4A]/60 text-[#A1A6C4] text-xs max-w-[200px] truncate font-mono">
                                {row[col] !== undefined ? String(row[col]) : ''}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    
                    {rawRows.length > 8 && (
                      <div className="p-3 bg-[#141229] text-center text-[11px] text-[#A1A6C4] border-t border-[#2E2C4A] sticky bottom-0 mt-auto">
                        Visualizando as primeiras 8 linhas para referência da disposição recebida.
                      </div>
                    )}
                  </div>

                )}
              </div>



            </div>

          </div>
        )}

      </main>

      {/* Export Confirmation Modal */}
      {isExportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center sm:p-4 bg-black/75 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-[#0F0E1A] w-full sm:max-w-lg rounded-t-3xl sm:rounded-3xl shadow-2xl relative animate-in slide-in-from-bottom-4 sm:zoom-in-95 duration-200 overflow-hidden">

            {/* Top accent bar */}
            <div className="h-1 w-full bg-gradient-to-r from-[#6366F1] to-[#7C3AED]" />

            <div className="p-6">
              {/* Header */}
              <div className="flex items-center justify-between mb-6">
                <div className="flex items-center gap-3">
                  <div className="h-9 w-9 rounded-xl bg-[#6366F1]/15 border border-[#6366F1]/30 flex items-center justify-center">
                    <Download className="h-4 w-4 text-[#8B80FF]" />
                  </div>
                  <div>
                    <h3 className="text-sm font-bold text-[#F4F5FC]">Exportar Planilha</h3>
                    <p className="text-[11px] text-[#A1A6C4]">{rawRows.length} linhas prontas para exportação</p>
                  </div>
                </div>
                <button
                  onClick={() => setIsExportModalOpen(false)}
                  className="h-8 w-8 flex items-center justify-center rounded-lg bg-[#1A1830] border border-[#2E2C4A] text-[#A1A6C4] hover:text-white transition-colors cursor-pointer text-lg leading-none"
                >
                  &times;
                </button>
              </div>

              {/* Form — side by side on wider screens */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
                <div className="bg-[#141229] border border-[#2E2C4A] rounded-2xl p-4">
                  <label className="block text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider mb-2">
                    Nome do Cliente <span className="text-red-400">*</span>
                  </label>
                  <input
                    type="text"
                    value={clientNameInput}
                    onChange={(e) => setClientNameInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter' && clientNameInput.trim()) confirmExport(); }}
                    placeholder="Ex: TECHSTORE"
                    className="w-full text-sm bg-[#0F0E1A] border border-[#2E2C4A] focus:border-[#6366F1] rounded-xl px-3 py-2 outline-none font-semibold text-[#F4F5FC] uppercase placeholder:text-[#A1A6C4]/20 transition-colors"
                    autoFocus
                  />
                </div>

                <div className="bg-[#141229] border border-[#2E2C4A] rounded-2xl p-4">
                  <label className="block text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider mb-2">
                    Mês de Referência
                  </label>
                  <select
                    value={monthNameInput}
                    onChange={(e) => setMonthNameInput(e.target.value)}
                    className="w-full text-sm bg-[#0F0E1A] border border-[#2E2C4A] focus:border-[#6366F1] rounded-xl px-3 py-2 outline-none font-semibold text-[#F4F5FC] cursor-pointer transition-colors"
                  >
                    {[
                      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
                    ].map((month) => (
                      <option key={month} value={month} className="bg-[#141229] font-medium text-white">
                        {month}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* File name preview */}
              <div className={`rounded-2xl border px-4 py-3 mb-6 transition-all duration-200 ${clientNameInput.trim() ? 'bg-[#6366F1]/5 border-[#6366F1]/30' : 'bg-[#141229] border-[#2E2C4A]'}`}>
                <p className="text-[10px] font-bold text-[#A1A6C4] uppercase tracking-wider mb-1">Arquivo gerado</p>
                <code className="text-xs font-mono text-[#8B80FF] break-all">
                  {clientNameInput.trim()
                    ? `Planilha de Importação- ${monthNameInput} - ${clientNameInput.trim().toUpperCase()}.csv`
                    : <span className="text-[#A1A6C4]/40 italic">Preencha o nome do cliente...</span>
                  }
                </code>
              </div>

              {/* Action Buttons */}
              <div className="flex items-center gap-3">
                <button
                  onClick={() => setIsExportModalOpen(false)}
                  className="px-5 py-2.5 bg-[#141229] hover:bg-[#1A1830] text-[#A1A6C4] font-bold text-xs rounded-xl border border-[#2E2C4A] transition-colors cursor-pointer"
                >
                  Cancelar
                </button>
                <button
                  onClick={confirmExport}
                  disabled={!clientNameInput.trim()}
                  className={`flex-1 px-4 py-2.5 text-sm font-bold rounded-xl transition-all flex items-center justify-center gap-2 cursor-pointer ${
                    clientNameInput.trim()
                      ? 'bg-gradient-to-r from-[#6366F1] to-[#7C3AED] text-white shadow-lg shadow-indigo-500/20 hover:shadow-indigo-500/40 hover:-translate-y-0.5'
                      : 'bg-[#1A1830] text-[#A1A6C4]/40 cursor-not-allowed border border-[#2E2C4A]'
                  }`}
                >
                  <Download className="h-4 w-4" />
                  Gerar e Baixar CSV
                </button>
              </div>
            </div>

          </div>
        </div>
      )}

      {/* History item context menu */}
      {openHistoryMenu && menuPosition && (() => {
        const item = exportHistory.find((h) => h.id === openHistoryMenu);
        if (!item) return null;
        return (
          <>
            <div className="fixed inset-0 z-40" onClick={() => { setOpenHistoryMenu(null); setMenuPosition(null); }} />
            <div
              className="fixed z-50 w-44 bg-[#141229] border border-[#2E2C4A] rounded-xl shadow-2xl shadow-black/40 overflow-hidden"
              style={{ top: menuPosition.top, right: menuPosition.right }}
            >
              <button
                onClick={() => { setOpenHistoryMenu(null); setMenuPosition(null); redownloadFromStorage(item); }}
                className="w-full flex items-center gap-2.5 px-3 py-2.5 text-xs text-[#F4F5FC] hover:bg-[#1A1830] transition-colors cursor-pointer"
              >
                <Download className="h-3.5 w-3.5 text-[#8B80FF]" />
                Baixar arquivo
              </button>
              <div className="h-px bg-[#2E2C4A]" />
              <button
                onClick={async () => {
                  setOpenHistoryMenu(null);
                  setMenuPosition(null);
                  await supabase.storage.from('exports').remove([item.filePath]);
                  await supabase.from('export_history').delete().eq('id', item.id);
                  setExportHistory((prev) => prev.filter((h) => h.id !== item.id));
                }}
                className="w-full flex items-center gap-2.5 px-3 py-2.5 text-xs text-red-400 hover:bg-red-500/10 transition-colors cursor-pointer"
              >
                <Trash2 className="h-3.5 w-3.5" />
                Excluir
              </button>
            </div>
          </>
        );
      })()}

      {/* Floating action buttons */}
      {rawRows.length > 0 && (
        <div className="fixed bottom-6 right-6 z-40 flex items-center gap-2">
          <button
            onClick={handleRemoveFile}
            className="flex items-center gap-2 px-4 py-3 bg-[#141229] border border-[#2E2C4A] hover:border-red-500/50 text-[#A1A6C4] hover:text-red-400 text-sm font-bold rounded-2xl shadow-xl shadow-black/20 hover:-translate-y-0.5 transition-all duration-200"
          >
            <Trash2 className="h-4 w-4" />
            Limpar
          </button>
          <button
            onClick={exportResult}
            className="flex items-center gap-2.5 px-5 py-3 bg-gradient-to-r from-[#6366F1] to-[#7C3AED] text-white text-sm font-bold rounded-2xl shadow-xl shadow-indigo-500/30 hover:shadow-indigo-500/50 hover:-translate-y-0.5 transition-all duration-200"
          >
            <Download className="h-4 w-4" />
            Exportar CSV
          </button>
        </div>
      )}

    </div>
  );
}
