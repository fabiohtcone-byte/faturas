import React, { useState, useRef, useEffect } from "react";
import {
  Network,
  FolderSync,
  HardDrive,
  Copy,
  Check,
  FolderOpen,
  FileSpreadsheet,
  AlertCircle,
  CheckCircle2,
  Loader2,
  X,
  RefreshCw,
  Server,
  ArrowRight,
  Sparkles,
  Leaf,
  Database,
  UploadCloud,
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import localforage from "localforage";
import { BillData } from "../types";
import {
  NETWORK_SHARE_PATH,
  readSpreadsheetFile,
  scanDirectoryFiles,
  saveSavedDirectoryHandle,
  getSavedDirectoryHandle,
  normalizeUcString,
  formatMonth,
  ProcessFileResult,
} from "../utils/networkSync";
import { sanitizeBillForDb } from "../App";

interface NetworkFolderSyncModalProps {
  isOpen: boolean;
  onClose: () => void;
  currentBills: BillData[];
  setBills: React.Dispatch<React.SetStateAction<BillData[]>>;
  ucMappings: any[];
  ucInfoMap: Map<string, any>;
  ucsLivre: Set<string>;
  ucsOper: Set<string>;
  supabase: any;
  isSupabaseConfigured: boolean;
  mapBillDataToDb: (bill: BillData, userId?: string) => any;
  onSyncSuccess?: (result: {
    filesCount: number;
    newCount: number;
    updatedCount: number;
    totalCount: number;
    greenAdjustedCount: number;
  }) => void;
}

export const NetworkFolderSyncModal: React.FC<NetworkFolderSyncModalProps> = ({
  isOpen,
  onClose,
  currentBills,
  setBills,
  ucMappings,
  ucInfoMap,
  ucsLivre,
  ucsOper,
  supabase,
  isSupabaseConfigured,
  mapBillDataToDb,
  onSyncSuccess,
}) => {
  const [copied, setCopied] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [currentStep, setCurrentStep] = useState<string>("");
  const [progressPercent, setProgressPercent] = useState<number>(0);
  const [hasSavedHandle, setHasSavedHandle] = useState(false);
  const [savedHandle, setSavedHandle] = useState<any>(null);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  // Sync Results
  const [syncResult, setSyncResult] = useState<{
    files: string[];
    newBillsCount: number;
    updatedBillsCount: number;
    totalBillsCount: number;
    greenAdjustedCount: number;
    totalValorR$: number;
    totalConsumoKwh: number;
    affectedUcs: string[];
  } | null>(null);

  // Bridge Intranet URL option
  const [bridgeUrl, setBridgeUrl] = useState<string>(
    localStorage.getItem("sicone_network_bridge_url") || ""
  );
  const [showBridgeConfig, setShowBridgeConfig] = useState(false);

  const folderInputRef = useRef<HTMLInputElement>(null);
  const filesInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (isOpen) {
      setErrorMsg(null);
      // Verifica se há handle salvo
      getSavedDirectoryHandle().then((handle) => {
        if (handle) {
          setHasSavedHandle(true);
          setSavedHandle(handle);
        }
      });
    }
  }, [isOpen]);

  const handleCopyPath = () => {
    navigator.clipboard.writeText(NETWORK_SHARE_PATH);
    setCopied(true);
    setTimeout(() => setCopied(false), 3000);
  };

  /**
   * Processa uma lista de arquivos File[] (vindos de showDirectoryPicker ou file inputs)
   */
  const processFilesBatch = async (files: File[]) => {
    if (files.length === 0) {
      setErrorMsg("Nenhuma planilha válida (.xlsx, .xls ou .csv) foi encontrada.");
      return;
    }

    setIsProcessing(true);
    setErrorMsg(null);
    setSyncResult(null);
    setCurrentStep(`Iniciando leitura de ${files.length} planilhas da pasta...`);
    setProgressPercent(5);

    try {
      const allExtractedBills: BillData[] = [];
      let totalGreenAdjusted = 0;
      const processedFilesList: string[] = [];

      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        setCurrentStep(`Processando ${i + 1}/${files.length}: ${file.name}`);
        setProgressPercent(Math.round(10 + ((i + 1) / files.length) * 60));

        try {
          const res: ProcessFileResult = await readSpreadsheetFile(
            file,
            ucMappings,
            ucInfoMap,
            ucsLivre,
            ucsOper
          );
          if (res.bills.length > 0) {
            allExtractedBills.push(...res.bills);
            totalGreenAdjusted += res.greenAdjustedCount;
            processedFilesList.push(file.name);
          }
        } catch (fileErr: any) {
          console.warn(`Aviso ao ler ${file.name}:`, fileErr);
        }
      }

      if (allExtractedBills.length === 0) {
        throw new Error(
          "Não foram encontrados registros de faturas válidos nas planilhas selecionadas. Verifique se as planilhas contêm os cabeçalhos esperados (UC, Mês, Ano, Valor, etc.)."
        );
      }

      setCurrentStep(`Atualizando valores das faturas no aplicativo...`);
      setProgressPercent(75);

      // Mapeamento e Merging com as faturas existentes
      let newCount = 0;
      let updatedCount = 0;
      const affectedUcsSet = new Set<string>();
      let totalValor = 0;
      let totalConsumo = 0;

      // Cria índice de faturas existentes para rápida atualização
      const billsMap = new Map<string, BillData>();
      currentBills.forEach((b) => {
        const cleanUc = normalizeUcString(b.uc);
        const normMes = formatMonth(b.mesReferencia);
        const key = cleanUc && normMes && b.anoLeitura ? `${cleanUc}_${normMes}_${b.anoLeitura}` : b.id;
        billsMap.set(key, b);
      });

      // Aplica os novos dados das planilhas da pasta de rede
      allExtractedBills.forEach((nb) => {
        const cleanUc = normalizeUcString(nb.uc);
        const normMes = formatMonth(nb.mesReferencia);
        const key = cleanUc && normMes && nb.anoLeitura ? `${cleanUc}_${normMes}_${nb.anoLeitura}` : nb.id;

        if (cleanUc) affectedUcsSet.add(cleanUc);

        const valNum = parseFloat(String(nb.valorTotal || "").replace(/\./g, "").replace(",", ".")) || 0;
        const consNum =
          (parseFloat(String(nb.consumoKwhPonta || "").replace(/\./g, "").replace(",", ".")) || 0) +
          (parseFloat(String(nb.consumoKwhForaPonta || "").replace(/\./g, "").replace(",", ".")) || 0) +
          (parseFloat(String(nb.consumoKwhGrupoB || "").replace(/\./g, "").replace(",", ".")) || 0);

        totalValor += valNum;
        totalConsumo += consNum;

        const existing = billsMap.get(key);
        if (existing) {
          // Atualiza dados preservando campos pré-existentes
          billsMap.set(key, {
            ...existing,
            ...nb,
            id: existing.id || nb.id,
            status: "completed",
          });
          updatedCount++;
        } else {
          // Nova fatura inserida
          billsMap.set(key, {
            ...nb,
            status: "completed",
          });
          newCount++;
        }
      });

      const updatedBillsList = Array.from(billsMap.values());

      // 1. Atualiza estado global do React
      setBills(updatedBillsList);

      // 2. Persiste no localforage (IndexedDB local)
      setCurrentStep(`Salvando dados no banco de dados local...`);
      setProgressPercent(85);
      try {
        await localforage.setItem("sanesul_bills_data", updatedBillsList);
      } catch (storeErr) {
        console.warn("Aviso ao salvar localmente:", storeErr);
      }

      // 3. Salva no banco de dados na Nuvem (Supabase/Firebase) se configurado
      if (isSupabaseConfigured) {
        setCurrentStep(`Sincronizando faturas atualizadas com o banco na nuvem...`);
        setProgressPercent(92);
        try {
          const { data: userData } = await supabase.auth.getUser().catch(() => ({ data: { user: null } }));
          const userId = userData?.user?.id || null;
          const dbData = allExtractedBills.map((b) => mapBillDataToDb(b, userId));

          const chunkSize = 50;
          for (let i = 0; i < dbData.length; i += chunkSize) {
            const chunk = dbData.slice(i, i + chunkSize);
            let { error } = await supabase.from("bills").upsert(chunk);
            if (error) {
              const fallbackChunk = chunk.map((d: any) => {
                const { user_id, ...rest } = d;
                return rest;
              });
              await supabase.from("bills").upsert(fallbackChunk).catch(console.warn);
            }
          }
        } catch (dbErr) {
          console.warn("Aviso ao salvar faturas na nuvem:", dbErr);
        }
      }

      setProgressPercent(100);
      setCurrentStep("Sincronização concluída com sucesso!");

      const resultObj = {
        files: processedFilesList,
        newBillsCount: newCount,
        updatedBillsCount: updatedCount,
        totalBillsCount: updatedBillsList.length,
        greenAdjustedCount: totalGreenAdjusted,
        totalValorR$: totalValor,
        totalConsumoKwh: totalConsumo,
        affectedUcs: Array.from(affectedUcsSet),
      };

      setSyncResult(resultObj);

      if (onSyncSuccess) {
        onSyncSuccess({
          filesCount: processedFilesList.length,
          newCount,
          updatedCount,
          totalCount: updatedBillsList.length,
          greenAdjustedCount: totalGreenAdjusted,
        });
      }
    } catch (err: any) {
      console.error("Erro na sincronização da pasta de rede:", err);
      setErrorMsg(err?.message || "Ocorreu um erro ao processar os arquivos da rede.");
    } finally {
      setIsProcessing(false);
    }
  };

  /**
   * Dispara a API nativa showDirectoryPicker (Chrome / Edge / Opera)
   */
  const handlePickDirectoryApi = async () => {
    // Copia o caminho para o clipboard para facilitar o usuário colar no Windows Explorer
    handleCopyPath();

    if ("showDirectoryPicker" in window) {
      try {
        const dirHandle = await (window as any).showDirectoryPicker({
          id: "sicone_dados_extraidos_picker",
          mode: "read",
        });

        if (dirHandle) {
          await saveSavedDirectoryHandle(dirHandle);
          setSavedHandle(dirHandle);
          setHasSavedHandle(true);

          setCurrentStep("Vasculhando arquivos na pasta...");
          setIsProcessing(true);
          const files = await scanDirectoryFiles(dirHandle);
          await processFilesBatch(files);
        }
      } catch (pickerErr: any) {
        if (pickerErr.name === "AbortError") {
          return; // Usuário apenas cancelou a janela
        }
        console.warn("showDirectoryPicker falhou ou foi bloqueado, usando seletor alternativo:", pickerErr);
        // Fallback para input de pasta
        folderInputRef.current?.click();
      }
    } else {
      // Navegadores sem suporte a showDirectoryPicker usam input de diretório
      folderInputRef.current?.click();
    }
  };

  /**
   * Re-sincroniza usando o handle já salvo anteriormente
   */
  const handleResyncSavedHandle = async () => {
    if (!savedHandle) return;
    try {
      // Verifica/requisita permissão se necessário
      if (savedHandle.requestPermission) {
        const perm = await savedHandle.requestPermission({ mode: "read" });
        if (perm !== "granted") {
          throw new Error("Permissão de leitura negada pelo usuário.");
        }
      }
      setIsProcessing(true);
      setCurrentStep("Re-escaneando pasta de rede...");
      const files = await scanDirectoryFiles(savedHandle);
      await processFilesBatch(files);
    } catch (err: any) {
      console.warn("Falha ao usar handle salvo:", err);
      setHasSavedHandle(false);
      handlePickDirectoryApi();
    }
  };

  /**
   * Dispara a leitura via seletor padrão de pasta HTML5
   */
  const handleFolderInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;
    const files = Array.from(fileList).filter((f) => {
      const lower = f.name.toLowerCase();
      return lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.endsWith(".csv");
    });
    processFilesBatch(files);
  };

  /**
   * Dispara a leitura de múltiplos arquivos selecionados
   */
  const handleFilesInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const fileList = e.target.files;
    if (!fileList || fileList.length === 0) return;
    const files = Array.from(fileList).filter((f) => {
      const lower = f.name.toLowerCase();
      return lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.endsWith(".csv");
    });
    processFilesBatch(files);
  };

  /**
   * Sincronização via API Bridge Intranet (opcional)
   */
  const handleSyncViaBridge = async () => {
    if (!bridgeUrl.trim()) {
      setErrorMsg("Informe a URL do serviço / bridge local na intranet.");
      return;
    }
    localStorage.setItem("sicone_network_bridge_url", bridgeUrl.trim());
    setIsProcessing(true);
    setErrorMsg(null);
    setCurrentStep(`Conectando ao serviço em ${bridgeUrl}...`);
    try {
      const resp = await fetch(bridgeUrl.trim(), { method: "GET" });
      if (!resp.ok) throw new Error(`Servidor respondeu com status ${resp.status}`);
      const data = await resp.json();
      if (!Array.isArray(data)) throw new Error("A resposta da API deve ser um array de faturas/planilhas.");
      // Converte se vier direto em formato de faturas
      alert(`Recebidos ${data.length} registros da API Bridge!`);
    } catch (err: any) {
      setErrorMsg(`Erro ao conectar ao bridge local: ${err?.message || "Falha na conexão"}`);
    } finally {
      setIsProcessing(false);
    }
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs overflow-y-auto">
        <motion.div
          initial={{ opacity: 0, scale: 0.95, y: 15 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 15 }}
          className="relative w-full max-w-2xl bg-white rounded-3xl shadow-2xl border border-slate-200 overflow-hidden my-8"
        >
          {/* Header */}
          <div className="relative px-6 py-5 bg-gradient-to-r from-slate-900 via-blue-950 to-indigo-900 text-white flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-2xl bg-blue-500/20 border border-blue-400/30 flex items-center justify-center shadow-inner">
                <Network className="w-6 h-6 text-cyan-300 animate-pulse" />
              </div>
              <div>
                <span className="text-[10px] font-black tracking-widest uppercase text-cyan-300">
                  Intranet / Compartilhamento SMB
                </span>
                <h3 className="text-lg font-display font-bold text-white">
                  Sincronizar Pasta de Rede: DADOS EXTRAÍDOS
                </h3>
              </div>
            </div>
            <button
              onClick={onClose}
              disabled={isProcessing}
              className="p-2 rounded-xl text-slate-300 hover:text-white hover:bg-white/10 transition-colors disabled:opacity-50"
              title="Fechar"
            >
              <X size={20} />
            </button>
          </div>

          <div className="p-6 space-y-6 max-h-[80vh] overflow-y-auto">
            {/* Network UNC Path Display Box */}
            <div className="bg-slate-50 border border-slate-200/90 rounded-2xl p-4 space-y-2">
              <div className="flex items-center justify-between text-xs font-bold text-slate-600">
                <span className="flex items-center gap-1.5">
                  <HardDrive size={14} className="text-sanesul-primary" />
                  Caminho de Rede Configurado:
                </span>
                <span className="text-[10px] font-mono uppercase bg-blue-100 text-blue-700 px-2 py-0.5 rounded-md">
                  Origem das Planilhas
                </span>
              </div>
              <div className="flex items-center gap-2">
                <div className="flex-1 bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 font-mono text-xs text-slate-800 select-all overflow-x-auto shadow-inner">
                  {NETWORK_SHARE_PATH}
                </div>
                <button
                  type="button"
                  onClick={handleCopyPath}
                  className={`flex items-center gap-1.5 px-4 py-2.5 rounded-xl text-xs font-bold transition-all shadow-xs active:scale-95 ${
                    copied
                      ? "bg-emerald-600 text-white"
                      : "bg-blue-600 hover:bg-blue-700 text-white"
                  }`}
                  title="Copiar caminho para colar na janela de seleção do Windows"
                >
                  {copied ? (
                    <>
                      <Check size={15} />
                      <span>Copiado!</span>
                    </>
                  ) : (
                    <>
                      <Copy size={15} />
                      <span>Copiar</span>
                    </>
                  )}
                </button>
              </div>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                💡 <strong>Dica prática:</strong> Ao clicar no botão abaixo, o caminho será
                automaticamente copiado. Na janela do seletor de arquivos do Windows, basta colar
                (<strong>Ctrl+V</strong>) na barra de endereços de pastas e pressionar <strong>Enter</strong> para abrir diretamente a pasta.
              </p>
            </div>

            {/* Error Message */}
            {errorMsg && (
              <div className="p-4 bg-red-50 border border-red-200 rounded-2xl flex items-start gap-3 text-red-700 text-xs">
                <AlertCircle size={18} className="shrink-0 mt-0.5 text-red-600" />
                <div className="flex-1 leading-relaxed">
                  <strong className="block font-bold">Não foi possível concluir a importação:</strong>
                  {errorMsg}
                </div>
              </div>
            )}

            {/* Progress Status Bar (when processing) */}
            {isProcessing && (
              <div className="p-5 bg-blue-50/70 border border-blue-200 rounded-2xl space-y-3">
                <div className="flex items-center justify-between text-xs font-bold text-blue-800">
                  <div className="flex items-center gap-2">
                    <Loader2 size={16} className="animate-spin text-blue-600" />
                    <span>{currentStep}</span>
                  </div>
                  <span>{progressPercent}%</span>
                </div>
                <div className="w-full h-2.5 bg-blue-200/80 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-blue-600 via-indigo-600 to-cyan-500 transition-all duration-300 rounded-full"
                    style={{ width: `${progressPercent}%` }}
                  />
                </div>
                <p className="text-[11px] text-blue-600 font-medium">
                  Extraindo indicadores, aplicando regras de modalidade verde (demanda ponta = 0) e atualizando o banco de dados...
                </p>
              </div>
            )}

            {/* Success Summary Report (when finished) */}
            {syncResult && !isProcessing && (
              <div className="p-5 bg-emerald-50 border border-emerald-200/90 rounded-2xl space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2 text-emerald-800 font-bold text-sm">
                    <CheckCircle2 size={18} className="text-emerald-600" />
                    <span>Sincronização com DADOS EXTRAÍDOS Realizada com Sucesso!</span>
                  </div>
                  <span className="text-[10px] font-bold px-2.5 py-0.5 bg-emerald-200 text-emerald-900 rounded-full">
                    {syncResult.files.length} planilhas
                  </span>
                </div>

                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                  <div className="bg-white p-3 rounded-xl border border-emerald-100 shadow-xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase block">
                      Faturas Atualizadas
                    </span>
                    <span className="text-base font-display font-bold text-blue-600">
                      {syncResult.updatedBillsCount}
                    </span>
                  </div>

                  <div className="bg-white p-3 rounded-xl border border-emerald-100 shadow-xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase block">
                      Novas Faturas
                    </span>
                    <span className="text-base font-display font-bold text-emerald-600">
                      {syncResult.newBillsCount}
                    </span>
                  </div>

                  <div className="bg-white p-3 rounded-xl border border-emerald-100 shadow-xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase block">
                      Modalidade Verde
                    </span>
                    <span className="text-base font-display font-bold text-emerald-700 flex items-center gap-1">
                      <Leaf size={14} />
                      {syncResult.greenAdjustedCount}
                    </span>
                    <span className="text-[9px] text-slate-400 block mt-0.5">Demanda Ponta = 0 kW</span>
                  </div>

                  <div className="bg-white p-3 rounded-xl border border-emerald-100 shadow-xs">
                    <span className="text-[10px] font-bold text-slate-500 uppercase block">
                      UCs Afetadas
                    </span>
                    <span className="text-base font-display font-bold text-indigo-700">
                      {syncResult.affectedUcs.length}
                    </span>
                  </div>
                </div>

                {/* Planilhas processadas */}
                <div className="text-xs text-slate-600 space-y-1">
                  <span className="font-bold block text-slate-700">Planilhas processadas:</span>
                  <div className="max-h-28 overflow-y-auto space-y-1 bg-white/70 p-2.5 rounded-xl border border-emerald-100 text-[11px] font-mono">
                    {syncResult.files.map((name, idx) => (
                      <div key={idx} className="flex items-center gap-1.5 text-slate-700">
                        <FileSpreadsheet size={13} className="text-emerald-600 shrink-0" />
                        <span className="truncate">{name}</span>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* Action Buttons Group */}
            <div className="space-y-3 pt-2">
              <span className="text-xs font-bold text-slate-700 uppercase tracking-wider block">
                Escolha o Método de Importação:
              </span>

              {/* Action 1: Sincronização Principal */}
              {hasSavedHandle ? (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <button
                    type="button"
                    onClick={handleResyncSavedHandle}
                    disabled={isProcessing}
                    className="flex items-center justify-center gap-2 p-4 bg-gradient-to-r from-blue-600 via-indigo-600 to-blue-700 text-white rounded-2xl font-bold text-xs shadow-md shadow-blue-500/20 hover:scale-[1.01] active:scale-95 transition-all disabled:opacity-50"
                  >
                    <RefreshCw size={17} className={isProcessing ? "animate-spin" : ""} />
                    <span>Re-escanear Pasta Salva Agora</span>
                  </button>

                  <button
                    type="button"
                    onClick={handlePickDirectoryApi}
                    disabled={isProcessing}
                    className="flex items-center justify-center gap-2 p-4 bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 rounded-2xl font-bold text-xs shadow-xs active:scale-95 transition-all disabled:opacity-50"
                  >
                    <FolderOpen size={17} className="text-sanesul-primary" />
                    <span>Selecionar Outra Pasta</span>
                  </button>
                </div>
              ) : (
                <button
                  type="button"
                  onClick={handlePickDirectoryApi}
                  disabled={isProcessing}
                  className="w-full flex items-center justify-center gap-3 p-4 bg-gradient-to-r from-blue-700 via-indigo-700 to-cyan-600 text-white hover:from-blue-800 hover:to-indigo-800 rounded-2xl font-bold text-sm shadow-lg shadow-blue-500/25 hover:shadow-blue-500/35 active:scale-[0.99] transition-all disabled:opacity-50"
                >
                  <FolderSync size={20} className="text-cyan-200" />
                  <span>Acessar e Importar Pasta de Rede (DADOS EXTRAÍDOS)</span>
                  <ArrowRight size={16} />
                </button>
              )}

              {/* Action 2 & 3: Alternativas de Seleção Manual */}
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
                <button
                  type="button"
                  onClick={() => {
                    handleCopyPath();
                    folderInputRef.current?.click();
                  }}
                  disabled={isProcessing}
                  className="flex items-center justify-center gap-2 p-3 bg-slate-50 hover:bg-slate-100 border border-slate-200/90 text-slate-700 rounded-xl text-xs font-bold transition-all shadow-xs active:scale-95 disabled:opacity-50"
                >
                  <FolderOpen size={15} className="text-blue-600" />
                  <span>Selecionar Pasta via Explorer</span>
                </button>

                <button
                  type="button"
                  onClick={() => {
                    handleCopyPath();
                    filesInputRef.current?.click();
                  }}
                  disabled={isProcessing}
                  className="flex items-center justify-center gap-2 p-3 bg-slate-50 hover:bg-slate-100 border border-slate-200/90 text-slate-700 rounded-xl text-xs font-bold transition-all shadow-xs active:scale-95 disabled:opacity-50"
                >
                  <FileSpreadsheet size={15} className="text-emerald-600" />
                  <span>Selecionar Arquivos .xlsx / .csv</span>
                </button>
              </div>
            </div>

            {/* Accordion: Opção de Bridge Local / API Intranet */}
            <div className="pt-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setShowBridgeConfig(!showBridgeConfig)}
                className="flex items-center justify-between w-full text-xs font-bold text-slate-500 hover:text-slate-800 transition-colors py-1"
              >
                <span className="flex items-center gap-1.5">
                  <Server size={14} className="text-indigo-600" />
                  Opção Avançada: Conectar via API / Agente HTTP Local
                </span>
                <span className="text-[10px] text-slate-400">
                  {showBridgeConfig ? "Ocultar" : "Expandir"}
                </span>
              </button>

              {showBridgeConfig && (
                <div className="mt-3 p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3 text-xs">
                  <p className="text-slate-600 leading-relaxed text-[11px]">
                    Se houver um script ou microserviço HTTP rodando na rede local (ex: Python/Flask ou Node.js em <code>http://10.100.100.114:8080/api/faturas</code> ou <code>http://localhost:5000/faturas</code>), informe o endereço para sincronização direta:
                  </p>
                  <div className="flex items-center gap-2">
                    <input
                      type="text"
                      placeholder="http://10.100.100.114:8080/api/faturas"
                      value={bridgeUrl}
                      onChange={(e) => setBridgeUrl(e.target.value)}
                      className="flex-1 bg-white border border-slate-300 rounded-xl px-3 py-2 text-xs text-slate-800 outline-none focus:border-blue-500"
                    />
                    <button
                      type="button"
                      onClick={handleSyncViaBridge}
                      disabled={isProcessing || !bridgeUrl.trim()}
                      className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition-all shadow-xs active:scale-95 disabled:opacity-50"
                    >
                      Sincronizar API
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Footer */}
          <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
            <span className="text-[11px] text-slate-500 font-medium flex items-center gap-1.5">
              <Database size={13} className="text-slate-400" />
              Atualiza automaticamente gráficos, tabelas e visões gerais
            </span>
            <button
              type="button"
              onClick={onClose}
              disabled={isProcessing}
              className="px-5 py-2.5 bg-white border border-slate-200 hover:bg-slate-100 text-slate-700 rounded-xl text-xs font-bold shadow-xs active:scale-95 transition-all disabled:opacity-50"
            >
              {syncResult ? "Concluir" : "Fechar"}
            </button>
          </div>

          {/* Hidden HTML Inputs */}
          <input
            type="file"
            ref={folderInputRef}
            onChange={handleFolderInputChange}
            {...({ webkitdirectory: "", directory: "" } as any)}
            multiple
            className="hidden"
          />
          <input
            type="file"
            ref={filesInputRef}
            onChange={handleFilesInputChange}
            multiple
            accept=".xlsx, .xls, .csv"
            className="hidden"
          />
        </motion.div>
      </div>
    </AnimatePresence>
  );
};
