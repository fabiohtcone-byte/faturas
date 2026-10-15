import * as XLSX from "xlsx";
import Papa from "papaparse";
import { BillData } from "../types";
import { UC_MIGRATION_MAP } from "../uc_migration";

export const NETWORK_SHARE_PATH = "\\\\10.100.100.114\\Sicone_app\\Faturas\\DADOS EXTRAIDOS";

export const normalizeHeaderKey = (str: any): string =>
  String(str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

export const normalizeUcString = (val: any): string => {
  if (val === undefined || val === null) return "";
  let s = String(val).trim();
  // Se vier em notação científica comum em planilhas Excel (ex: 8.9431905152e+10 ou 8.94319E+10)
  if (/^[0-9]+(\.[0-9]+)?[eE]\+[0-9]+$/.test(s)) {
    try {
      s = BigInt(Math.round(Number(s))).toString();
    } catch {
      s = Number(s).toFixed(0);
    }
  }
  // Remove sufixo decimal gerado pelo Excel para números (ex: "89431905152.0" -> "89431905152")
  s = s.replace(/\.0+$/, "");
  // Remove pontos, traços, barras e espaços
  s = s.replace(/[\.\-\/\s]/g, "");
  // Verifica no mapa de migração
  if (UC_MIGRATION_MAP[s]) {
    s = UC_MIGRATION_MAP[s];
  }
  return s;
};

export const formatMonth = (month: string | number): string => {
  if (!month) return "";
  let normalized = month.toString().toLowerCase().trim();
  if (normalized.includes("/")) {
    normalized = normalized.split("/")[0].trim();
  }

  const monthMap: Record<string, string> = {
    "01": "Janeiro",
    "1": "Janeiro",
    janeiro: "Janeiro",
    jan: "Janeiro",
    "02": "Fevereiro",
    "2": "Fevereiro",
    fevereiro: "Fevereiro",
    fev: "Fevereiro",
    "03": "Março",
    "3": "Março",
    março: "Março",
    marco: "Março",
    mar: "Março",
    "04": "Abril",
    "4": "Abril",
    abril: "Abril",
    abr: "Abril",
    "05": "Maio",
    "5": "Maio",
    maio: "Maio",
    mai: "Maio",
    "06": "Junho",
    "6": "Junho",
    junho: "Junho",
    jun: "Junho",
    "07": "Julho",
    "7": "Julho",
    julho: "Julho",
    jul: "Julho",
    "08": "Agosto",
    "8": "Agosto",
    agosto: "Agosto",
    ago: "Agosto",
    "09": "Setembro",
    "9": "Setembro",
    setembro: "Setembro",
    set: "Setembro",
    "10": "Outubro",
    outubro: "Outubro",
    out: "Outubro",
    "11": "Novembro",
    novembro: "Novembro",
    nov: "Novembro",
    "12": "Dezembro",
    dezembro: "Dezembro",
    dez: "Dezembro",
  };

  return monthMap[normalized] || month.toString();
};

export const parseValue = (val: any): number => {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const str = String(val).trim();
  if (!str) return 0;
  // Trata formato brasileiro 1.234,56 ou internacional 1234.56
  const clean = str.replace(/[R$\s]/g, "");
  if (clean.includes(",") && clean.includes(".")) {
    return parseFloat(clean.replace(/\./g, "").replace(",", ".")) || 0;
  }
  if (clean.includes(",")) {
    return parseFloat(clean.replace(",", ".")) || 0;
  }
  return parseFloat(clean) || 0;
};

export const COLUMN_ALIASES: Record<string, string> = {
  nomedoarquivo: "fileName",
  arquivo: "fileName",
  filename: "fileName",
  nomearquivo: "fileName",
  planilha: "fileName",
  uc: "uc",
  unidadeconsumidora: "uc",
  codigocliente: "uc",
  matricula: "uc",
  numerocliente: "uc",
  instalacao: "uc",
  codinstalacao: "uc",
  gerencia: "gerencia",
  locins: "locin",
  locin: "locin",
  cidade: "cidade",
  municipio: "cidade",
  endereco: "endereco",
  endereço: "endereco",
  enderecocompleto: "endereco",
  endereçocompleto: "endereco",
  logradouro: "endereco",
  rua: "endereco",
  address: "endereco",
  local: "endereco",
  localidade: "endereco",
  end: "endereco",
  enderecoimovel: "endereco",
  endereçodoimovel: "endereco",
  enderecodoimovel: "endereco",
  enderecouc: "endereco",
  endereçouc: "endereco",
  enderecodaunidade: "endereco",
  localizacao: "endereco",
  localização: "endereco",
  bairro: "endereco",
  tipo: "tipo",
  mercado: "mercado",
  concessionaria: "concessionaria",
  distribuidora: "concessionaria",
  mesreferencia: "mesReferencia",
  mesref: "mesReferencia",
  mes: "mesReferencia",
  referencia: "mesReferencia",
  anoleitura: "anoLeitura",
  ano: "anoLeitura",
  vencimento: "dataVencimento",
  datavencimento: "dataVencimento",
  notafiscal: "numeroNotaFiscal",
  numeronotafiscal: "numeroNotaFiscal",
  nf: "numeroNotaFiscal",
  numeronf: "numeroNotaFiscal",
  modalidadetarifaria: "modalidadeTarifaria",
  modalidade: "modalidadeTarifaria",
  tarifa: "modalidadeTarifaria",
  subgrupo: "subgrupo",
  valortotal: "valorTotal",
  valortotalr: "valorTotal",
  valortotalfatura: "valorTotal",
  totalfatura: "valorTotal",
  valorapagar: "valorTotal",
  demandacontratadapontakw: "demandaPontaKW",
  demandapontakw: "demandaPontaKW",
  demandacontratadaponta: "demandaPontaKW",
  demandacontratadaforapontakw: "demandaForaPontaKW",
  demandaforapontakw: "demandaForaPontaKW",
  demandacontratadaforaponta: "demandaForaPontaKW",
  demandatodosperiodoskw: "demandaTodosPeriodosKW",
  demandatodosperiodos: "demandaTodosPeriodosKW",
  demandamedidapontakw: "demandaPotenciaMedidaPonta",
  demandamedidaponta: "demandaPotenciaMedidaPonta",
  valordemandamedidapontar: "valorDemandaPotenciaMedidaPonta",
  demandamedidaforapontakw: "demandaPotenciaMedidaForaPonta",
  demandamedidaforaponta: "demandaPotenciaMedidaForaPonta",
  valordemandamedidaforapontar: "valorDemandaPotenciaMedidaForaPonta",
  consumoemkwhgrupob: "consumoKwhGrupoB",
  consumoemkwhgrupobkwh: "consumoKwhGrupoB",
  consumoemkwhgrupobr: "consumoKwhGrupoB",
  consumokwhgrupob: "consumoKwhGrupoB",
  consumogrupobkwh: "consumoKwhGrupoB",
  consumogrupob: "consumoKwhGrupoB",
  consumogrupobemkwh: "consumoKwhGrupoB",
  consumoemkwhb: "consumoKwhGrupoB",
  consumokwhb: "consumoKwhGrupoB",
  consumobkwh: "consumoKwhGrupoB",
  consumob: "consumoKwhGrupoB",
  consumoativogrupob: "consumoKwhGrupoB",
  consumofaturadogrupob: "consumoKwhGrupoB",
  consumomedidogrupob: "consumoKwhGrupoB",
  consumototalgrupob: "consumoKwhGrupoB",
  consumodogrupob: "consumoKwhGrupoB",
  consumoemkwhdogrupob: "consumoKwhGrupoB",
  consumogrupobr: "consumoKwhGrupoB",
  valorconsumoemkwhgrupob: "valorConsumoKwhGrupoB",
  valorconsumoemkwhgrupobr: "valorConsumoKwhGrupoB",
  valorconsumogrupobr: "valorConsumoKwhGrupoB",
  valorconsumogrupob: "valorConsumoKwhGrupoB",
  valorconsumokwhgrupob: "valorConsumoKwhGrupoB",
  valorconsumokwhgrupobr: "valorConsumoKwhGrupoB",
  valordoconsumogrupob: "valorConsumoKwhGrupoB",
  valordoconsumoemkwhgrupob: "valorConsumoKwhGrupoB",
  valorconsumob: "valorConsumoKwhGrupoB",
  valorconsumobr: "valorConsumoKwhGrupoB",
  valorgrupobr: "valorConsumoKwhGrupoB",
  valorgrupob: "valorConsumoKwhGrupoB",
  consumorgrupob: "valorConsumoKwhGrupoB",
  consumorgrupobr: "valorConsumoKwhGrupoB",
  consumopontakwh: "consumoKwhPonta",
  consumoemkwhponta: "consumoKwhPonta",
  valorconsumopontar: "valorConsumoKwhPonta",
  consumoemkwhintermediaria: "consumoKwhIntermediaria",
  consumointermediariakwh: "consumoKwhIntermediaria",
  consumointermediaria: "consumoKwhIntermediaria",
  consumointermediario: "consumoKwhIntermediaria",
  consumoforapontakwh: "consumoKwhForaPonta",
  consumoemkwhforaponta: "consumoKwhForaPonta",
  valorconsumoforapontar: "valorConsumoKwhForaPonta",
  demandanaoconsumidapontakw: "demandaPotenciaNaoConsumidaPonta",
  valordemandanaoconsumidapontar: "valorDemandaPotenciaNaoConsumidaPonta",
  demandanaoconsumidaforapontakw: "demandaPotenciaNaoConsumidaFPonta",
  valordemandanaoconsumidaforapontar: "valorDemandaPotenciaNaoConsumidaFPonta",
  ultrapassagempontakw: "demandaPotenciaAtivaUltrapPonta",
  valorultrapassagempontar: "valorDemandaPotenciaAtivaUltrapPonta",
  ultrapassagemforapontakw: "demandaPotenciaAtivaUltrapFPonta",
  valorultrapassagemforapontar: "valorDemandaPotenciaAtivaUltrapFPonta",
  reativaexcedentepontakvarh: "energiaReativaExcedPonta",
  valorreativaexcedentepontar: "valorEnergiaReativaExcedPonta",
  reativaexcedenteforapontakvarh: "energiaReativaExcedFPonta",
  valorreativaexcedenteforapontar: "valorEnergiaReativaExcedFPonta",
  gdiouckwh: "energiaAtvInjetadaGDIOUC",
  valorgdioucr: "valorEnergiaAtvInjetadaGDIOUC",
  gdimuckwh: "energiaAtvInjetadaGDIMUC",
  valorgdimucr: "valorEnergiaAtvInjetadaGDIMUC",
  energiainjetadakwh: "energiaInjetadaKwh",
  energiacompensadakwh: "energiaCompensadaKwh",
  saldoacumuladogd: "saldoAcumuladoGD",
  saldoacumulado: "saldoAcumuladoGD",
  saldoacumuladogdkwh: "saldoAcumuladoGD",
  saldoatualgd: "saldoAcumuladoGD",
  saldoatual: "saldoAcumuladoGD",
  cipr: "cip",
  cip: "cip",
  outrosencargosr: "outrosEncargos",
  outrosencargos: "outrosEncargos",
  pisr: "pis",
  pis: "pis",
  cofinsr: "cofins",
  cofins: "cofins",
  icmsr: "icms",
  icms: "icms",
};

export interface ProcessFileResult {
  fileName: string;
  bills: BillData[];
  greenAdjustedCount: number;
}

/**
 * Converte linhas de uma planilha (XLSX, XLS ou CSV) em BillData[]
 */
export const parseSheetRows = (
  rowsData: any[][],
  sourceFileName: string,
  ucMappings: any[] = [],
  ucInfoMap: Map<string, any> = new Map(),
  ucsLivre: Set<string> = new Set(),
  ucsOper: Set<string> = new Set()
): { bills: BillData[]; greenAdjustedCount: number } => {
  if (!rowsData || rowsData.length < 2) {
    return { bills: [], greenAdjustedCount: 0 };
  }

  const rawHeaders = rowsData[0] as string[];
  const rows = rowsData.slice(1);
  let greenAdjustedCount = 0;

  const bills: BillData[] = rows
    .filter((row) => row && row.some((val) => val !== undefined && val !== null && String(val).trim() !== ""))
    .map((row, index) => {
      const billData: any = {
        id: crypto.randomUUID(),
        fileName: sourceFileName ? `${sourceFileName}_Linha_${index + 1}` : `Planilha_Linha_${index + 1}`,
        status: "completed",
        createdAt: Date.now() + index,
        tipo: "normal",
      };

      rawHeaders.forEach((header, colIndex) => {
        const rawHeader = String(header || "").trim();
        const normHeader = normalizeHeaderKey(rawHeader);
        let targetKey = COLUMN_ALIASES[normHeader];

        if (!targetKey) {
          if (normHeader.includes("grupob") || (normHeader.includes("grupo") && normHeader.includes("b"))) {
            if (normHeader.includes("valor") || normHeader.includes("r$") || normHeader.includes("rs")) {
              targetKey = "valorConsumoKwhGrupoB";
            } else if (normHeader.includes("consumo")) {
              targetKey = "consumoKwhGrupoB";
            }
          }
        }

        if (targetKey) {
          const val = row[colIndex];
          billData[targetKey] = val !== undefined && val !== null ? String(val).trim() : "";
        }
      });

      // Normaliza UC
      const cleanUc = normalizeUcString(billData.uc);
      billData.uc = cleanUc;

      // Normaliza Mês e Ano
      if (billData.mesReferencia) {
        billData.mesReferencia = formatMonth(billData.mesReferencia);
      }
      if (billData.anoLeitura) {
        const matchYear = String(billData.anoLeitura).match(/\d{4}/);
        if (matchYear) billData.anoLeitura = matchYear[0];
      }

      // Grupo B
      const valConsumoB = billData.consumoKwhGrupoB || billData.consumoGrupoB || billData.consumoKwh || "";
      if (valConsumoB !== "") {
        const cleanVal = String(valConsumoB).trim();
        billData.consumoKwhGrupoB = cleanVal;
        billData.consumoGrupoB = cleanVal;
        billData.consumoKwh = cleanVal;
      }
      if (billData.valorConsumoKwhGrupoB) {
        billData.valorConsumoKwhGrupoB = String(billData.valorConsumoKwhGrupoB).trim();
      }

      // Mapeamento de Cidade, LOCIN, Gerência e Endereço
      const mapping = ucMappings.find((m) => String(m.uc) === cleanUc);
      if (mapping) {
        billData.gerencia = mapping.gerencia;
        billData.locin = mapping.locin;
        if (!billData.cidade) billData.cidade = mapping.cidade;
        if (!billData.endereco && mapping.endereco) billData.endereco = mapping.endereco;
      }
      if (!billData.endereco) {
        const knownInfo = ucInfoMap.get(cleanUc);
        if (knownInfo?.endereco) billData.endereco = knownInfo.endereco;
      }

      // Modalidade VERDE: Demanda Contratada Ponta é SEMPRE 0 (não há contratação na ponta)
      const modTarifaria = (billData.modalidadeTarifaria || "").toUpperCase();
      if (modTarifaria.includes("VERDE") || cleanUc === "163738105109") {
        billData.demandaPontaKW = "0";
        greenAdjustedCount++;
      }

      // Limpa saldo GD
      if (billData.saldoAcumuladoGD) {
        billData.saldoAcumuladoGD = String(billData.saldoAcumuladoGD).replace(/\./g, "").trim();
      }

      // Mercado Livre vs Cativo
      if (cleanUc && ucsLivre.has(cleanUc)) {
        billData.mercado = "LIVRE";
        billData.tipo = "LIVRE";
        if (!modTarifaria.includes("LIVRE")) {
          billData.modalidadeTarifaria = billData.modalidadeTarifaria
            ? `${billData.modalidadeTarifaria} - LIVRE`
            : "LIVRE";
        }
      } else {
        billData.mercado = "CATIVO";
      }

      if (cleanUc && ucsOper.has(cleanUc)) {
        billData.tipo = "OPER";
      }

      // ID determinístico para atualização transparente
      const normMes = formatMonth(billData.mesReferencia);
      if (cleanUc && normMes && billData.anoLeitura) {
        billData.id = `${cleanUc}_${normMes}_${billData.anoLeitura}`;
      }

      return billData as BillData;
    });

  return { bills, greenAdjustedCount };
};

/**
 * Lê um arquivo individual de planilha (XLSX, XLS ou CSV)
 */
export const readSpreadsheetFile = async (
  file: File,
  ucMappings: any[] = [],
  ucInfoMap: Map<string, any> = new Map(),
  ucsLivre: Set<string> = new Set(),
  ucsOper: Set<string> = new Set()
): Promise<ProcessFileResult> => {
  const fileName = file.name;
  const isCsv = fileName.toLowerCase().endsWith(".csv");

  if (isCsv) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: false,
        skipEmptyLines: true,
        complete: (results) => {
          try {
            const { bills, greenAdjustedCount } = parseSheetRows(
              results.data as any[][],
              fileName,
              ucMappings,
              ucInfoMap,
              ucsLivre,
              ucsOper
            );
            resolve({ fileName, bills, greenAdjustedCount });
          } catch (err) {
            reject(err);
          }
        },
        error: (err) => reject(err),
      });
    });
  }

  // Arquivo XLSX / XLS
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  let allBills: BillData[] = [];
  let totalGreenAdjusted = 0;

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) as any[][];
    if (data.length > 1) {
      const { bills, greenAdjustedCount } = parseSheetRows(
        data,
        `${fileName} [${sheetName}]`,
        ucMappings,
        ucInfoMap,
        ucsLivre,
        ucsOper
      );
      allBills.push(...bills);
      totalGreenAdjusted += greenAdjustedCount;
    }
  }

  return {
    fileName,
    bills: allBills,
    greenAdjustedCount: totalGreenAdjusted,
  };
};

/**
 * Escaneia recursivamente um FileSystemDirectoryHandle e retorna todos os arquivos .xlsx, .xls e .csv
 */
export const scanDirectoryFiles = async (
  dirHandle: any,
  maxFiles = 500
): Promise<File[]> => {
  const files: File[] = [];

  const readDir = async (handle: any) => {
    if (files.length >= maxFiles) return;
    for await (const entry of handle.values()) {
      if (entry.kind === "file") {
        const lower = entry.name.toLowerCase();
        if (
          lower.endsWith(".xlsx") ||
          lower.endsWith(".xls") ||
          lower.endsWith(".csv")
        ) {
          try {
            const f = await entry.getFile();
            files.push(f);
            if (files.length >= maxFiles) break;
          } catch (e) {
            console.warn("Erro ao obter arquivo do entry:", entry.name, e);
          }
        }
      } else if (entry.kind === "directory") {
        // Escaneia subpastas
        await readDir(entry);
      }
    }
  };

  await readDir(dirHandle);
  return files;
};

/**
 * Armazena o handle no IndexedDB para permitir re-sincronização com 1 clique
 */
const DB_NAME = "sicone_network_handles";
const STORE_NAME = "handles";

export const saveSavedDirectoryHandle = async (handle: any) => {
  try {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    return new Promise<void>((resolve, reject) => {
      request.onsuccess = () => {
        const db = request.result;
        const tx = db.transaction(STORE_NAME, "readwrite");
        const store = tx.objectStore(STORE_NAME);
        store.put(handle, "dados_extraidos_handle");
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      };
      request.onerror = () => reject(request.error);
    });
  } catch (err) {
    console.warn("Não foi possível salvar handle no IndexedDB:", err);
  }
};

export const getSavedDirectoryHandle = async (): Promise<any | null> => {
  try {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    return new Promise((resolve) => {
      request.onsuccess = () => {
        const db = request.result;
        const tx = db.transaction(STORE_NAME, "readonly");
        const store = tx.objectStore(STORE_NAME);
        const getReq = store.get("dados_extraidos_handle");
        getReq.onsuccess = () => resolve(getReq.result || null);
        getReq.onerror = () => resolve(null);
      };
      request.onerror = () => resolve(null);
    });
  } catch (err) {
    console.warn("Não foi possível carregar handle do IndexedDB:", err);
    return null;
  }
};
