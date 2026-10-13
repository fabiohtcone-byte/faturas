import React, { useState, useRef, useEffect, useMemo } from "react";
import {
  Bot,
  Send,
  Sparkles,
  AlertTriangle,
  DollarSign,
  Zap,
  Leaf,
  CheckCircle2,
  Copy,
  Check,
  RotateCcw,
  ShieldCheck,
  TrendingDown,
  ArrowRight,
  HelpCircle,
  FileText,
  Search,
  Table as TableIcon,
  X,
  ChevronDown,
  Layers,
  Database,
  MapPin,
} from "lucide-react";
import { GoogleGenAI } from "@google/genai";
import { BillData } from "../types";

export interface FaturaAssistente {
  id: string;
  uc: string;
  gerencia: string;
  locins: string;
  cidade: string;
  endereco: string;
  tipo: string;
  mercado: string;
  concessionaria: string;
  mesReferencia: string;
  anoLeitura: string;
  referenciaFormatada: string;
  vencimento: string;
  notaFiscal: string;
  modalidadeTarifaria: string;
  subgrupo: string;
  valorTotal: number;
  consumoGrupoB: number;
  valorConsumoGrupoB: number;
  consumoIntermediaria: number;
  demandaContratadaPonta: number;
  demandaContratadaForaPonta: number;
  demandaTodosPeriodos: number;
  demandaMedidaPonta: number;
  valorDemandaMedidaPonta: number;
  demandaMedidaForaPonta: number;
  valorDemandaMedidaForaPonta: number;
  consumoPonta: number;
  valorConsumoPonta: number;
  consumoForaPonta: number;
  valorConsumoForaPonta: number;
  demandaNaoConsumidaPonta: number;
  valorDemandaNaoConsumidaPonta: number;
  demandaNaoConsumidaForaPonta: number;
  valorDemandaNaoConsumidaForaPonta: number;
  ultrapassagemPonta: number;
  valorUltrapassagemPonta: number;
  ultrapassagemForaPonta: number;
  valorUltrapassagemForaPonta: number;
  reativaExcedentePonta: number;
  valorReativaExcedentePonta: number;
  reativaExcedenteForaPonta: number;
  valorReativaExcedenteForaPonta: number;
  gdiOucKwh: number;
  valorGdiOuc: number;
  gdiMucKwh: number;
  valorGdiMuc: number;
  energiaInjetadaKwh: number;
  energiaCompensadaKwh: number;
  saldoAcumuladoGD: number;
  cip: number;
  outrosEncargos: number;
  pis: number;
  cofins: number;
  icms: number;
}

interface AssistentePortalProps {
  bills: BillData[];
  dashboardData: any[];
  savingsData?: any[];
  customUsinaSaldoGD: Record<string, number>;
  isUcPPP: (uc: string, ref?: string) => boolean;
  UCS_USINA: Set<string>;
  UCS_PPP: Set<string>;
  formatMonth: (m: string | number) => string;
  parseValue: (v: any) => number;
  setActiveTab: (tab: any) => void;
  setIsActionCenterOpen: (open: boolean) => void;
  getLocin?: (uc: string) => string;
  getGerencia?: (uc: string) => string;
  getCidade?: (uc: string, fallback?: string) => string;
  getEndereco?: (uc: string, fallback?: string) => string;
}

interface Message {
  id: string;
  sender: "user" | "assistant";
  text: string;
  timestamp: string;
  quickAction?: {
    label: string;
    action: () => void;
  };
}

export const AssistentePortal: React.FC<AssistentePortalProps> = ({
  bills,
  dashboardData,
  savingsData = [],
  customUsinaSaldoGD,
  isUcPPP,
  UCS_USINA,
  UCS_PPP,
  formatMonth,
  parseValue,
  setActiveTab,
  setIsActionCenterOpen,
  getLocin,
  getGerencia,
  getCidade,
  getEndereco,
}) => {
  const [messages, setMessages] = useState<Message[]>([
    {
      id: "welcome",
      sender: "assistant",
      text: `Olá! Sou o **Assistente do Portal de Inteligência Energética (COTAA)**.\n\nTodos os **52 campos de cada fatura** foram carregados na minha base de consulta direta:\n- **Identificação & Localização:** UC, **Endereço Completo ou Parcial**, Gerência, LOCINS, Cidade, Tipo, Mercado, Concessionária\n- **Faturamento:** Mês Referência, Ano Leitura, Vencimento, Nota Fiscal, Modalidade Tarifária, Subgrupo, Valor Total (R$)\n- **Consumo:** Grupo B (kWh e R$), Intermediária, Ponta (kWh e R$), Fora Ponta (kWh e R$)\n- **Demanda:** Contratada Ponta/Fora Ponta/Todos os Períodos, Medida Ponta/Fora Ponta (kW e R$), Demanda Não Consumida Ponta/Fora Ponta (kW e R$)\n- **Multas & Penalidades:** Ultrapassagem Ponta/Fora Ponta (kW e R$), Reativa Excedente Ponta/Fora Ponta (kVArh e R$)\n- **Geração Distribuída:** GDI oUC (kWh e R$), GDI mUC (kWh e R$), Energia Injetada, Compensada e Saldo GD\n- **Tributos:** CIP, Outros Encargos, PIS, COFINS e ICMS\n\n> 📍 *Localização de UCs:* Você pode **localizar qualquer UC pelo endereço completo ou parte dele** (ex: *"qual a UC da Rua Maracaju?"*, *"Afonso Pena"* ou *"Rua das Flores 123"*).\n> ℹ️ *Aviso importante:* Atuo exclusivamente em **modo consulta**. Não realizo qualquer alteração cadastral no sistema.`,
      timestamp: new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }),
    },
  ]);

  const [input, setInput] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [isDataViewerOpen, setIsDataViewerOpen] = useState(false);
  const [viewerPage, setViewerPage] = useState(1);
  const [searchFilter, setSearchFilter] = useState("");
  const messagesEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages, isLoading]);

  // 1. Mapeamento e Carregamento estruturado de todos os 52 campos de todas as faturas
  const faturasCarregadas: FaturaAssistente[] = useMemo(() => {
    const sourceBills = bills.length > 0 ? bills : [];

    return sourceBills.map((b, index) => {
      const rawUc = String(b.uc || "").trim();
      const gerencia = b.gerencia && b.gerencia !== "---" ? b.gerencia : (getGerencia ? getGerencia(rawUc) : "---");
      const locins = b.locin && b.locin !== "---" ? b.locin : (getLocin ? getLocin(rawUc) : "---");
      const cidade = b.cidade && b.cidade !== "---" ? b.cidade : (getCidade ? getCidade(rawUc, b.cidade) : "---");
      const endereco = b.endereco && b.endereco !== "---" ? b.endereco : (getEndereco ? getEndereco(rawUc, b.endereco) : (b.endereco || "---"));
      const tipo = b.tipo || (isUcPPP(rawUc) ? "PPP Fotovoltaica" : (UCS_USINA.has(rawUc) ? "Usina Sanesul" : "Normal"));
      const mercado = b.mercado || ((b.modalidadeTarifaria || "").toUpperCase().includes("LIVRE") ? "LIVRE" : "CATIVO");
      const concessionaria = b.concessionaria || (b.fileName?.toLowerCase().includes("elektro") ? "ELEKTRO" : "ENERGISA");
      const mesReferencia = b.mesReferencia || "";
      const anoLeitura = b.anoLeitura || "";
      const referenciaFormatada = `${formatMonth(b.mesReferencia)} / ${b.anoLeitura}`;
      const vencimento = b.dataVencimento || "---";
      const notaFiscal = b.numeroNotaFiscal || "---";
      const modalidadeTarifaria = b.modalidadeTarifaria || "---";
      const subgrupo = b.subgrupo || "---";

      const valorTotal = parseValue(b.valorTotal);
      const consumoGrupoB = parseValue(b.consumoKwhGrupoB || b.consumoGrupoB || b.consumoKwh);
      const valorConsumoGrupoB = parseValue(b.valorConsumoKwhGrupoB);
      const consumoIntermediaria = parseValue(b.consumoKwhIntermediaria);

      const isVerde = (b.modalidadeTarifaria || "").toUpperCase().includes("VERDE");
      const demandaContratadaPonta = isVerde ? 0 : parseValue(b.demandaPontaKW);
      const demandaContratadaForaPonta = parseValue(b.demandaForaPontaKW);
      const demandaTodosPeriodos = parseValue(b.demandaTodosPeriodosKW || (b as any).demandaTodosPeriodos);

      const demandaMedidaPonta = parseValue(b.demandaPotenciaMedidaPonta);
      const valorDemandaMedidaPonta = parseValue(b.valorDemandaPotenciaMedidaPonta);
      const demandaMedidaForaPonta = parseValue(b.demandaPotenciaMedidaForaPonta);
      const valorDemandaMedidaForaPonta = parseValue(b.valorDemandaPotenciaMedidaForaPonta);

      const consumoPonta = parseValue(b.consumoKwhPonta);
      const valorConsumoPonta = parseValue(b.valorConsumoKwhPonta);
      const consumoForaPonta = parseValue(b.consumoKwhForaPonta);
      const valorConsumoForaPonta = parseValue(b.valorConsumoKwhForaPonta);

      const demandaNaoConsumidaPonta = parseValue(b.demandaPotenciaNaoConsumidaPonta);
      const valorDemandaNaoConsumidaPonta = parseValue(b.valorDemandaPotenciaNaoConsumidaPonta);
      const demandaNaoConsumidaForaPonta = parseValue(b.demandaPotenciaNaoConsumidaFPonta);
      const valorDemandaNaoConsumidaForaPonta = parseValue(b.valorDemandaPotenciaNaoConsumidaFPonta);

      const ultrapassagemPonta = parseValue(b.demandaPotenciaAtivaUltrapPonta);
      const valorUltrapassagemPonta = parseValue(b.valorDemandaPotenciaAtivaUltrapPonta);
      const ultrapassagemForaPonta = parseValue(b.demandaPotenciaAtivaUltrapFPonta);
      const valorUltrapassagemForaPonta = parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta);

      const reativaExcedentePonta = parseValue(b.energiaReativaExcedPonta);
      const valorReativaExcedentePonta = parseValue(b.valorEnergiaReativaExcedPonta);
      const reativaExcedenteForaPonta = parseValue(b.energiaReativaExcedFPonta);
      const valorReativaExcedenteForaPonta = parseValue(b.valorEnergiaReativaExcedFPonta);

      const gdiOucKwh = parseValue(b.energiaAtvInjetadaGDIOUC);
      const valorGdiOuc = parseValue(b.valorEnergiaAtvInjetadaGDIOUC);
      const gdiMucKwh = parseValue(b.energiaAtvInjetadaGDIMUC);
      const valorGdiMuc = parseValue(b.valorEnergiaAtvInjetadaGDIMUC);

      const energiaInjetadaKwh = parseValue(b.energiaInjetadaKwh || (gdiOucKwh + gdiMucKwh));
      const energiaCompensadaKwh = parseValue(b.energiaCompensadaKwh);

      const customKey = `${rawUc}_${referenciaFormatada}`;
      const saldoAcumuladoGD = (customUsinaSaldoGD && customUsinaSaldoGD[customKey] !== undefined)
        ? customUsinaSaldoGD[customKey]
        : parseValue(b.saldoAcumuladoGD);

      const cip = parseValue(b.cip);
      const outrosEncargos = parseValue(b.outrosEncargos);
      const pis = parseValue(b.pis);
      const cofins = parseValue(b.cofins);
      const icms = parseValue(b.icms);

      return {
        id: b.id || `${rawUc}_${mesReferencia}_${anoLeitura}_${index}`,
        uc: rawUc,
        gerencia,
        locins,
        cidade,
        endereco,
        tipo,
        mercado,
        concessionaria,
        mesReferencia,
        anoLeitura,
        referenciaFormatada,
        vencimento,
        notaFiscal,
        modalidadeTarifaria,
        subgrupo,
        valorTotal,
        consumoGrupoB,
        valorConsumoGrupoB,
        consumoIntermediaria,
        demandaContratadaPonta,
        demandaContratadaForaPonta,
        demandaTodosPeriodos,
        demandaMedidaPonta,
        valorDemandaMedidaPonta,
        demandaMedidaForaPonta,
        valorDemandaMedidaForaPonta,
        consumoPonta,
        valorConsumoPonta,
        consumoForaPonta,
        valorConsumoForaPonta,
        demandaNaoConsumidaPonta,
        valorDemandaNaoConsumidaPonta,
        demandaNaoConsumidaForaPonta,
        valorDemandaNaoConsumidaForaPonta,
        ultrapassagemPonta,
        valorUltrapassagemPonta,
        ultrapassagemForaPonta,
        valorUltrapassagemForaPonta,
        reativaExcedentePonta,
        valorReativaExcedentePonta,
        reativaExcedenteForaPonta,
        valorReativaExcedenteForaPonta,
        gdiOucKwh,
        valorGdiOuc,
        gdiMucKwh,
        valorGdiMuc,
        energiaInjetadaKwh,
        energiaCompensadaKwh,
        saldoAcumuladoGD,
        cip,
        outrosEncargos,
        pis,
        cofins,
        icms,
      };
    });
  }, [bills, getLocin, getGerencia, getCidade, getEndereco, isUcPPP, UCS_USINA, formatMonth, parseValue, customUsinaSaldoGD]);

  // 2. Índices e Totais calculados de todos os 52 campos em passagem única O(N)
  const calculatedContext = useMemo(() => {
    let totalFaturado = 0;
    let totalPis = 0;
    let totalCofins = 0;
    let totalIcms = 0;
    let totalCip = 0;
    let totalOutrosEncargos = 0;
    let totalConsumoGrupoB = 0;
    let totalValorConsumoGrupoB = 0;
    let totalConsumoIntermediaria = 0;
    let totalConsumoPonta = 0;
    let totalValorConsumoPonta = 0;
    let totalConsumoForaPonta = 0;
    let totalValorConsumoForaPonta = 0;
    let totalDemandaMedidaPonta = 0;
    let totalValorDemandaMedidaPonta = 0;
    let totalDemandaMedidaForaPonta = 0;
    let totalValorDemandaMedidaForaPonta = 0;
    let totalDemandaNaoConsumidaPonta = 0;
    let totalValorDemandaNaoConsumidaPonta = 0;
    let totalDemandaNaoConsumidaForaPonta = 0;
    let totalValorDemandaNaoConsumidaForaPonta = 0;
    let totalUltrapassagemPonta = 0;
    let totalValorUltrapassagemPonta = 0;
    let totalUltrapassagemForaPonta = 0;
    let totalValorUltrapassagemForaPonta = 0;
    let totalReativoPonta = 0;
    let totalValorReativoPonta = 0;
    let totalReativoForaPonta = 0;
    let totalValorReativoForaPonta = 0;
    let totalGdiOucKwh = 0;
    let totalValorGdiOuc = 0;
    let totalGdiMucKwh = 0;
    let totalValorGdiMuc = 0;
    let totalEnergiaInjetada = 0;
    let totalEnergiaCompensada = 0;
    let totalSaldoGD = 0;

    const ucsMap = new Map<string, FaturaAssistente[]>();
    const gerenciasMap = new Map<string, FaturaAssistente[]>();
    const cidadesMap = new Map<string, FaturaAssistente[]>();

    const len = faturasCarregadas.length;
    for (let i = 0; i < len; i++) {
      const f = faturasCarregadas[i];
      totalFaturado += f.valorTotal;
      totalPis += f.pis;
      totalCofins += f.cofins;
      totalIcms += f.icms;
      totalCip += f.cip;
      totalOutrosEncargos += f.outrosEncargos;
      totalConsumoGrupoB += f.consumoGrupoB;
      totalValorConsumoGrupoB += f.valorConsumoGrupoB;
      totalConsumoIntermediaria += f.consumoIntermediaria;
      totalConsumoPonta += f.consumoPonta;
      totalValorConsumoPonta += f.valorConsumoPonta;
      totalConsumoForaPonta += f.consumoForaPonta;
      totalValorConsumoForaPonta += f.valorConsumoForaPonta;
      totalDemandaMedidaPonta += f.demandaMedidaPonta;
      totalValorDemandaMedidaPonta += f.valorDemandaMedidaPonta;
      totalDemandaMedidaForaPonta += f.demandaMedidaForaPonta;
      totalValorDemandaMedidaForaPonta += f.valorDemandaMedidaForaPonta;
      totalDemandaNaoConsumidaPonta += f.demandaNaoConsumidaPonta;
      totalValorDemandaNaoConsumidaPonta += f.valorDemandaNaoConsumidaPonta;
      totalDemandaNaoConsumidaForaPonta += f.demandaNaoConsumidaForaPonta;
      totalValorDemandaNaoConsumidaForaPonta += f.valorDemandaNaoConsumidaForaPonta;
      totalUltrapassagemPonta += f.ultrapassagemPonta;
      totalValorUltrapassagemPonta += f.valorUltrapassagemPonta;
      totalUltrapassagemForaPonta += f.ultrapassagemForaPonta;
      totalValorUltrapassagemForaPonta += f.valorUltrapassagemForaPonta;
      totalReativoPonta += f.reativaExcedentePonta;
      totalValorReativoPonta += f.valorReativaExcedentePonta;
      totalReativoForaPonta += f.reativaExcedenteForaPonta;
      totalValorReativoForaPonta += f.valorReativaExcedenteForaPonta;
      totalGdiOucKwh += f.gdiOucKwh;
      totalValorGdiOuc += f.valorGdiOuc;
      totalGdiMucKwh += f.gdiMucKwh;
      totalValorGdiMuc += f.valorGdiMuc;
      totalEnergiaInjetada += f.energiaInjetadaKwh;
      totalEnergiaCompensada += f.energiaCompensadaKwh;
      totalSaldoGD += f.saldoAcumuladoGD;

      // Group by UC
      let uList = ucsMap.get(f.uc);
      if (!uList) {
        uList = [];
        ucsMap.set(f.uc, uList);
      }
      uList.push(f);

      // Group by Gerência
      const g = f.gerencia || "Outras";
      let gList = gerenciasMap.get(g);
      if (!gList) {
        gList = [];
        gerenciasMap.set(g, gList);
      }
      gList.push(f);

      // Group by Cidade
      const c = f.cidade || "Outras";
      let cList = cidadesMap.get(c);
      if (!cList) {
        cList = [];
        cidadesMap.set(c, cList);
      }
      cList.push(f);
    }

    const totalMultaUltrapassagem = totalValorUltrapassagemPonta + totalValorUltrapassagemForaPonta;
    const totalMultaReativo = totalValorReativoPonta + totalValorReativoForaPonta;
    const totalMultasGeral = totalMultaUltrapassagem + totalMultaReativo;

    // Ranking de UCs com multas
    const finesByUc: Array<{ uc: string; cidade: string; gerencia: string; multas: number; ultrapassagem: number; reativo: number }> = [];
    ucsMap.forEach((fatList, uc) => {
      let ultrap = 0;
      let reat = 0;
      for (let j = 0; j < fatList.length; j++) {
        const item = fatList[j];
        ultrap += item.valorUltrapassagemPonta + item.valorUltrapassagemForaPonta;
        reat += item.valorReativaExcedentePonta + item.valorReativaExcedenteForaPonta;
      }
      const total = ultrap + reat;
      if (total > 0) {
        finesByUc.push({
          uc,
          cidade: fatList[0].cidade,
          gerencia: fatList[0].gerencia,
          multas: total,
          ultrapassagem: ultrap,
          reativo: reat,
        });
      }
    });
    finesByUc.sort((a, b) => b.multas - a.multas);

    // Central de Ações
    const totalAjustesSugeridos = savingsData.length;
    let totalEconomiaEstimada = 0;
    for (let k = 0; k < savingsData.length; k++) {
      const s = savingsData[k];
      totalEconomiaEstimada += (s.economiaTotal || s.economia || 0);
    }

    return {
      totalFaturas: len,
      totalFaturado,
      totalPis,
      totalCofins,
      totalIcms,
      totalCip,
      totalOutrosEncargos,
      totalConsumoGrupoB,
      totalValorConsumoGrupoB,
      totalConsumoIntermediaria,
      totalConsumoPonta,
      totalValorConsumoPonta,
      totalConsumoForaPonta,
      totalValorConsumoForaPonta,
      totalDemandaMedidaPonta,
      totalValorDemandaMedidaPonta,
      totalDemandaMedidaForaPonta,
      totalValorDemandaMedidaForaPonta,
      totalDemandaNaoConsumidaPonta,
      totalValorDemandaNaoConsumidaPonta,
      totalDemandaNaoConsumidaForaPonta,
      totalValorDemandaNaoConsumidaForaPonta,
      totalUltrapassagemPonta,
      totalValorUltrapassagemPonta,
      totalUltrapassagemForaPonta,
      totalValorUltrapassagemForaPonta,
      totalMultaUltrapassagem,
      totalReativoPonta,
      totalValorReativoPonta,
      totalReativoForaPonta,
      totalValorReativoForaPonta,
      totalMultaReativo,
      totalMultasGeral,
      totalGdiOucKwh,
      totalValorGdiOuc,
      totalGdiMucKwh,
      totalValorGdiMuc,
      totalEnergiaInjetada,
      totalEnergiaCompensada,
      totalSaldoGD,
      qtdUcs: ucsMap.size,
      qtdGerencias: gerenciasMap.size,
      qtdCidades: cidadesMap.size,
      ucsMap,
      gerenciasMap,
      cidadesMap,
      finesByUc,
      totalAjustesSugeridos,
      totalEconomiaEstimada,
    };
  }, [faturasCarregadas, savingsData]);

  // Lista unificada e deduplicada de UCs com agregação cadastral e endereço
  const uniqueUcsList = useMemo(() => {
    const map = new Map<
      string,
      {
        uc: string;
        gerencia: string;
        locins: string;
        cidade: string;
        endereco: string;
        tipo: string;
        mercado: string;
        concessionaria: string;
        modalidadeTarifaria: string;
        subgrupo: string;
        totalFaturas: number;
        totalFaturado: number;
        totalMultas: number;
        demandaContratadaPonta: number;
        demandaContratadaForaPonta: number;
        demandaTodosPeriodos: number;
        saldoGD: number;
        faturas: FaturaAssistente[];
      }
    >();

    for (let i = 0; i < faturasCarregadas.length; i++) {
      const f = faturasCarregadas[i];
      let item = map.get(f.uc);
      if (!item) {
        item = {
          uc: f.uc,
          gerencia: f.gerencia,
          locins: f.locins,
          cidade: f.cidade,
          endereco: f.endereco && f.endereco !== "---" ? f.endereco : "",
          tipo: f.tipo,
          mercado: f.mercado,
          concessionaria: f.concessionaria,
          modalidadeTarifaria: f.modalidadeTarifaria,
          subgrupo: f.subgrupo,
          totalFaturas: 0,
          totalFaturado: 0,
          totalMultas: 0,
          demandaContratadaPonta: f.demandaContratadaPonta,
          demandaContratadaForaPonta: f.demandaContratadaForaPonta,
          demandaTodosPeriodos: f.demandaTodosPeriodos,
          saldoGD: f.saldoAcumuladoGD,
          faturas: [],
        };
        map.set(f.uc, item);
      } else {
        if ((!item.endereco || item.endereco === "---") && f.endereco && f.endereco !== "---") {
          item.endereco = f.endereco;
        }
        if ((!item.cidade || item.cidade === "---") && f.cidade && f.cidade !== "---") {
          item.cidade = f.cidade;
        }
        if ((!item.gerencia || item.gerencia === "---") && f.gerencia && f.gerencia !== "---") {
          item.gerencia = f.gerencia;
        }
        if ((!item.locins || item.locins === "---") && f.locins && f.locins !== "---") {
          item.locins = f.locins;
        }
      }

      item.totalFaturas += 1;
      item.totalFaturado += f.valorTotal;
      item.totalMultas +=
        f.valorUltrapassagemPonta +
        f.valorUltrapassagemForaPonta +
        f.valorReativaExcedentePonta +
        f.valorReativaExcedenteForaPonta;
      item.faturas.push(f);
    }

    return Array.from(map.values());
  }, [faturasCarregadas]);

  // Normalização de texto para pesquisa de endereços (ignora acentos, pontuações e expande abreviações)
  const normalizeAddressText = (text: string) => {
    return (text || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "") // remove acentos
      .replace(/[,;.:\/\-–_#ºª°\\()]/g, " ") // substitui pontuações
      .replace(/\br\b|\br\.\b/g, "rua")
      .replace(/\bav\b|\bav\.\b/g, "avenida")
      .replace(/\bal\b|\bal\.\b/g, "alameda")
      .replace(/\brod\b|\brod\.\b/g, "rodovia")
      .replace(/\bjd\b|\bjd\.\b/g, "jardim")
      .replace(/\btrav\b|\btrav\.\b/g, "travessa")
      .replace(/\bpr\b|\bpç\b|\bpraca\b/g, "praça")
      .replace(/\bbl\b|\bbl\.\b/g, "bloco")
      .replace(/\bap\b|\bap\.\b|\bapto\b/g, "apartamento")
      .replace(/\bs\/n\b|\bsn\b/g, "sem numero")
      .replace(/\s+/g, " ")
      .trim();
  };

  // Extrai o núcleo do endereço a partir de consultas em linguagem natural
  const extractAddressSearchTerm = (query: string): string => {
    let q = normalizeAddressText(query);

    q = q
      .replace(/^(por favor |me informe |me diga |poderia informar |gostaria de saber |quero saber |como localizar |como encontrar )+/g, "")
      .trim();

    q = q
      .replace(/^(qual( a)? uc (e|que e|fica no|fica na|fica em|esta no|esta na|esta em|tem o|tem|pertence ao|corresponde ao|do|da|de|no|na) )+/g, "")
      .replace(/^(localizar( a)? uc (pelo|por|no|na|com|de|do|da) )+/g, "")
      .replace(/^(encontrar( a)? uc (pelo|por|no|na|com|de|do|da) )+/g, "")
      .replace(/^(buscar( a)? uc (pelo|por|no|na|com|de|do|da) )+/g, "")
      .replace(/^(pesquisar( a)? uc (pelo|por|no|na|com|de|do|da) )+/g, "")
      .replace(/^(onde fica a uc (do|da|de|no|na|com) )+/g, "")
      .replace(/^(localizacao da uc (do|da|de|no|na|com) )+/g, "")
      .replace(/^(endereco da uc |endereco de uc |endereco da |endereco do |endereco )+/g, "")
      .replace(/^(uc no endereco |uc na rua |uc na avenida |uc da rua |uc da avenida |uc do endereco )+/g, "")
      .trim();

    q = q.replace(/[?!.]+$/, "").trim();
    return q;
  };

  // Motor de busca de UCs por endereço completo ou parte dele
  const searchUcsByAddress = (rawQuery: string) => {
    const rawNormalized = normalizeAddressText(rawQuery);
    const cleanedSearch = extractAddressSearchTerm(rawQuery);
    const termToUse = cleanedSearch.length >= 2 ? cleanedSearch : rawNormalized;

    if (!termToUse || termToUse.length < 2) return [];

    const ignoreTokens = new Set([
      "de", "do", "da", "dos", "das", "em", "no", "na", "nos", "nas",
      "e", "com", "qual", "uc", "ucs", "o", "a", "os", "as"
    ]);
    const queryTokens = termToUse
      .split(/\s+/)
      .filter((t) => t.length > 0 && !ignoreTokens.has(t));

    const matches: Array<{
      uc: string;
      endereco: string;
      cidade: string;
      gerencia: string;
      locins: string;
      tipo: string;
      mercado: string;
      concessionaria: string;
      modalidadeTarifaria: string;
      subgrupo: string;
      totalFaturado: number;
      totalMultas: number;
      totalFaturas: number;
      demandaContratadaPonta: number;
      demandaContratadaForaPonta: number;
      demandaTodosPeriodos: number;
      saldoGD: number;
      faturas: FaturaAssistente[];
      score: number;
    }> = [];

    for (let i = 0; i < uniqueUcsList.length; i++) {
      const u = uniqueUcsList[i];
      if (!u.endereco || u.endereco === "---" || u.endereco.trim().length === 0) {
        continue;
      }

      const normEnd = normalizeAddressText(u.endereco);
      const normCidade = normalizeAddressText(u.cidade);
      const fullLocation = `${normEnd} ${normCidade}`;

      let score = 0;

      // 1. Correspondência exata da frase de busca no endereço
      if (normEnd.includes(termToUse)) {
        score = 100 + (termToUse.length / normEnd.length) * 20;
      } else if (fullLocation.includes(termToUse)) {
        score = 90 + (termToUse.length / fullLocation.length) * 20;
      } else if (normEnd.includes(rawNormalized) || (rawNormalized.length > 5 && rawNormalized.includes(normEnd))) {
        score = 85;
      } else if (queryTokens.length > 0) {
        // 2. Correspondência ponderada por tokens (rua, número, nome do logradouro)
        let matchedCount = 0;
        let matchedWeight = 0;

        for (let j = 0; j < queryTokens.length; j++) {
          const token = queryTokens[j];
          if (normEnd.includes(token)) {
            matchedCount++;
            matchedWeight += /^\d+$/.test(token) ? 2 : (token.length > 3 ? 1.5 : 1);
          } else if (normCidade.includes(token)) {
            matchedCount += 0.5;
            matchedWeight += 0.5;
          }
        }

        const matchRatio = matchedCount / queryTokens.length;

        if (matchRatio >= 0.6 || (queryTokens.length >= 2 && matchedCount >= 2)) {
          score = 50 + matchRatio * 40 + Math.min(10, matchedWeight);
        }
      }

      if (score >= 40) {
        matches.push({
          ...u,
          score,
        });
      }
    }

    matches.sort((a, b) => b.score - a.score || b.totalFaturas - a.totalFaturas);
    return matches;
  };

  // Formatação em Reais
  const fmtMoney = (v: number) =>
    `R$ ${v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  // Formatação Numérica
  const fmtNum = (v: number, unit = "") =>
    `${v.toLocaleString("pt-BR", { minimumFractionDigits: 0, maximumFractionDigits: 2 })}${unit ? " " + unit : ""}`;

  // 3. Motor de Busca e Resposta Determinística Local
  const generateLocalResponse = (query: string): { text: string; quickAction?: { label: string; action: () => void } } => {
    const q = query.toLowerCase().trim();

    // 1. Busca por UC específica numérica (se não for busca com intenção de endereço tipo "rua ...")
    const ucDigitsMatch = q.match(/\b\d{6,15}\b/);
    const hasAddressKeywords = /\b(rua|av|avenida|alameda|rodovia|travessa|bairro|endereco|endereço)\b/i.test(q);

    if (ucDigitsMatch && !hasAddressKeywords) {
      const searchedUc = ucDigitsMatch[0];
      const matchingBills = faturasCarregadas.filter(
        (f) => f.uc === searchedUc || f.uc.includes(searchedUc)
      );

      if (matchingBills.length > 0) {
        const u = matchingBills[0];
        const totalFaturadoUc = matchingBills.reduce((acc, f) => acc + f.valorTotal, 0);
        const totalMultasUc = matchingBills.reduce(
          (acc, f) =>
            acc +
            f.valorUltrapassagemPonta +
            f.valorUltrapassagemForaPonta +
            f.valorReativaExcedentePonta +
            f.valorReativaExcedenteForaPonta,
          0
        );

        const ultimasFaturas = matchingBills
          .slice(0, 3)
          .map(
            (f) =>
              `- **Ref ${f.referenciaFormatada}** (NF ${f.notaFiscal}, Venc. ${f.vencimento}): Total ${fmtMoney(f.valorTotal)} | Consumo Ponta ${fmtNum(f.consumoPonta, "kWh")} (${fmtMoney(f.valorConsumoPonta)}) | Fora Ponta ${fmtNum(f.consumoForaPonta, "kWh")} (${fmtMoney(f.valorConsumoForaPonta)}) | Demanda Med. FP ${fmtNum(f.demandaMedidaForaPonta, "kW")} (${fmtMoney(f.valorDemandaMedidaForaPonta)}) | Ultrapassagem ${fmtMoney(f.valorUltrapassagemPonta + f.valorUltrapassagemForaPonta)} | Reativo ${fmtMoney(f.valorReativaExcedentePonta + f.valorReativaExcedenteForaPonta)} | PIS/COFINS/ICMS: ${fmtMoney(f.pis + f.cofins + f.icms)}`
          )
          .join("\n");

        return {
          text: `### 📋 Dados Completos Carregados da UC **${u.uc}**\n\n` +
            `📍 **Identificação Cadastral & Localização:**\n` +
            `- **UC:** ${u.uc}\n` +
            `- **Endereço Completo:** ${u.endereco && u.endereco !== "---" ? u.endereco : "Não informado no cadastro"}\n` +
            `- **Gerência:** ${u.gerencia}\n` +
            `- **LOCINS:** ${u.locins}\n` +
            `- **Cidade:** ${u.cidade}\n` +
            `- **Tipo:** ${u.tipo}\n` +
            `- **Mercado:** ${u.mercado}\n` +
            `- **Concessionária:** ${u.concessionaria}\n` +
            `- **Modalidade Tarifária:** ${u.modalidadeTarifaria}\n` +
            `- **Subgrupo:** ${u.subgrupo}\n\n` +
            `💰 **Valores Acumulados no Período (${matchingBills.length} fatura(s)):**\n` +
            `- **Valor Total Faturado:** ${fmtMoney(totalFaturadoUc)}\n` +
            `- **Total de Multas (Ultrapassagem + Reativo):** ${fmtMoney(totalMultasUc)}\n` +
            `- **CIP Acumulado:** ${fmtMoney(matchingBills.reduce((acc, f) => acc + f.cip, 0))}\n` +
            `- **PIS:** ${fmtMoney(matchingBills.reduce((acc, f) => acc + f.pis, 0))} | **COFINS:** ${fmtMoney(matchingBills.reduce((acc, f) => acc + f.cofins, 0))} | **ICMS:** ${fmtMoney(matchingBills.reduce((acc, f) => acc + f.icms, 0))}\n\n` +
            `⚡ **Contratos de Demanda & Consumo Médio:**\n` +
            `- **Demanda Contratada Ponta:** ${fmtNum(u.demandaContratadaPonta, "kW")} | **Fora Ponta:** ${fmtNum(u.demandaContratadaForaPonta, "kW")} | **Todos Períodos:** ${fmtNum(u.demandaTodosPeriodos, "kW")}\n` +
            `- **Demanda Não Consumida:** ${fmtMoney(matchingBills.reduce((acc, f) => acc + f.valorDemandaNaoConsumidaPonta + f.valorDemandaNaoConsumidaForaPonta, 0))}\n` +
            `- **GDI oUC:** ${fmtNum(matchingBills.reduce((acc, f) => acc + f.gdiOucKwh, 0), "kWh")} (${fmtMoney(matchingBills.reduce((acc, f) => acc + f.valorGdiOuc, 0))})\n` +
            `- **GDI mUC:** ${fmtNum(matchingBills.reduce((acc, f) => acc + f.gdiMucKwh, 0), "kWh")} (${fmtMoney(matchingBills.reduce((acc, f) => acc + f.valorGdiMuc, 0))})\n` +
            `- **Saldo GD:** ${fmtNum(u.saldoAcumuladoGD, "kWh")}\n\n` +
            `📑 **Detalhamento das Últimas Faturas:**\n${ultimasFaturas}`,
          quickAction: {
            label: "Ver Gestão de Faturas",
            action: () => setActiveTab("faturas"),
          },
        };
      }
    }

    // 2. Consulta de estatísticas / listagem de UCs com endereço cadastrado
    if (
      (q.includes("quantas") || q.includes("quais") || q.includes("lista") || q.includes("todas") || q.includes("cobertura")) &&
      (q.includes("endereço") || q.includes("endereco"))
    ) {
      const ucsComEndereco = uniqueUcsList.filter((u) => u.endereco && u.endereco !== "---" && u.endereco.trim().length > 0);
      const taxa = uniqueUcsList.length > 0 ? Math.round((ucsComEndereco.length / uniqueUcsList.length) * 100) : 0;

      const amostraLista = ucsComEndereco
        .slice(0, 8)
        .map((u) => `- **UC ${u.uc}:** ${u.endereco} (${u.cidade} / ${u.gerencia})`)
        .join("\n");

      return {
        text: `### 📋 Cobertura de Endereços Cadastrados nas UCs\n\n` +
          `- **Total de UCs Únicas na Base:** ${uniqueUcsList.length}\n` +
          `- **UCs com Endereço Completo:** ${ucsComEndereco.length} (${taxa}% de cobertura)\n` +
          `- **UCs sem Endereço Registrado:** ${uniqueUcsList.length - ucsComEndereco.length}\n\n` +
          `#### Exemplos de UCs com Endereço Registrado:\n${amostraLista || "Nenhum endereço cadastrado até o momento."}\n\n` +
          `> 💡 *Você pode localizar qualquer uma dessas UCs digitando o nome da rua, número, avenida ou bairro no chat.* Para importar ou atualizar endereços em lote, utilize a importação via planilha Excel ou arquivo TXT na aba **Gestão de Faturas**.`,
        quickAction: {
          label: "Inspecionar Base de Faturas",
          action: () => setIsDataViewerOpen(true),
        },
      };
    }

    // 3. Como pesquisar / Ajuda sobre localização de UC por endereço
    if (
      (q.includes("como") || q.includes("ajuda") || q.includes("tutorial")) &&
      (q.includes("localizar") || q.includes("encontrar") || q.includes("buscar")) &&
      (q.includes("endereço") || q.includes("endereco"))
    ) {
      const ucsComEndereco = uniqueUcsList.filter((u) => u.endereco && u.endereco !== "---" && u.endereco.trim().length > 0);
      const exemplos = ucsComEndereco
        .slice(0, 3)
        .map((u) => `- *"Qual a UC do endereço ${u.endereco}?"* (Retorna UC **${u.uc}**)`)
        .join("\n");

      return {
        text: `### 📍 Como Localizar UCs pelo Endereço no Assistente\n\n` +
          `O Assistente possui motor de busca com inteligência fonética e correspondência por partes de endereço (rua, avenida, bairro, número ou cidade).\n\n` +
          `**Você pode perguntar de várias formas:**\n` +
          `1. **Endereço Completo:** *"Qual a UC do endereço Rua Maracaju, 450?"*\n` +
          `2. **Parte do Endereço:** *"UC da Rua das Flores"*, *"Afonso Pena"*, *"Av. Brasil"*\n` +
          `3. **Apenas o Logradouro ou Bairro:** *"rua cuiaba"*, *"jardim dos estados"*\n` +
          `4. **Número da Residência/Edifício:** *"localizar UC número 500"*\n\n` +
          (exemplos ? `**Exemplos que você pode testar agora:**\n${exemplos}\n\n` : "") +
          `O sistema retornará o número da UC, dados contratuais, cidade, gerência e histórico de faturas.`,
        quickAction: {
          label: "Ver Base (52 Campos)",
          action: () => setIsDataViewerOpen(true),
        },
      };
    }

    // 4. Execução do Motor de Localização de UC por Endereço
    const addressMatches = searchUcsByAddress(query);
    const isAddressExplicit = /endere[cç]o|localiz|onde fica|qual.*uc|encontr.*uc|busca.*uc|rua\b|avenida\b|\bav\b|alameda\b|rodovia\b|travessa\b|bairro\b|logradouro\b/i.test(q);

    if (addressMatches.length > 0) {
      if (addressMatches.length === 1) {
        const m = addressMatches[0];
        const ultimas = m.faturas
          .slice(0, 3)
          .map(
            (f) =>
              `- **Ref ${f.referenciaFormatada}** (NF ${f.notaFiscal}, Venc. ${f.vencimento}): Total ${fmtMoney(f.valorTotal)} | Consumo Ponta ${fmtNum(f.consumoPonta, "kWh")} | Fora Ponta ${fmtNum(f.consumoForaPonta, "kWh")} | Demanda Med. FP ${fmtNum(f.demandaMedidaForaPonta, "kW")}`
          )
          .join("\n");

        return {
          text: `### 🎯 UC Localizada com Sucesso pelo Endereço!\n\n` +
            `🏢 **Unidade Consumidora (UC):** \`${m.uc}\`\n\n` +
            `📍 **Localização & Identificação:**\n` +
            `- **Endereço Completo:** **${m.endereco}**\n` +
            `- **Cidade:** ${m.cidade}\n` +
            `- **Gerência:** ${m.gerencia}\n` +
            `- **LOCINS:** ${m.locins}\n` +
            `- **Tipo:** ${m.tipo} | **Mercado:** ${m.mercado} | **Concessionária:** ${m.concessionaria}\n` +
            `- **Modalidade Tarifária:** ${m.modalidadeTarifaria} | **Subgrupo:** ${m.subgrupo}\n\n` +
            `💰 **Histórico no Sistema (${m.totalFaturas} fatura(s)):**\n` +
            `- **Faturamento Total:** ${fmtMoney(m.totalFaturado)}\n` +
            `- **Total de Multas:** ${fmtMoney(m.totalMultas)}\n` +
            `- **Demanda Contratada Ponta:** ${fmtNum(m.demandaContratadaPonta, "kW")} | **Fora Ponta:** ${fmtNum(m.demandaContratadaForaPonta, "kW")} | **Todos Períodos:** ${fmtNum(m.demandaTodosPeriodos, "kW")}\n` +
            (m.saldoGD > 0 ? `- **Saldo GD Acumulado:** ${fmtNum(m.saldoGD, "kWh")}\n` : "") +
            `\n📑 **Últimas Faturas Registradas:**\n${ultimas || "Nenhuma fatura detalhada disponível."}`,
          quickAction: {
            label: `Ver Faturas da UC ${m.uc}`,
            action: () => setActiveTab("faturas"),
          },
        };
      } else {
        // Múltiplas UCs encontradas para o endereço informado
        const listaUcs = addressMatches
          .slice(0, 10)
          .map(
            (m, idx) =>
              `${idx + 1}. **UC ${m.uc}**\n` +
              `   - 📍 **Endereço:** ${m.endereco}\n` +
              `   - 🏙️ **Localidade:** ${m.cidade} (${m.gerencia}) | **LOCINS:** ${m.locins}\n` +
              `   - ⚡ **Tipo / Mercado:** ${m.tipo} (${m.mercado}) | **Faturamento:** ${fmtMoney(m.totalFaturado)} (${m.totalFaturas} fatura(s))`
          )
          .join("\n\n");

        const searchedTerm = extractAddressSearchTerm(query);

        return {
          text: `### 📍 Foram Localizadas ${addressMatches.length} UCs para o Endereço Informado\n\n` +
            `**Termo pesquisado:** *"${searchedTerm || query}"*\n\n` +
            `${listaUcs}\n\n` +
            (addressMatches.length > 10 ? `> *(Exibindo as 10 UCs mais relevantes de ${addressMatches.length} encontradas)*\n\n` : "") +
            `💡 *Para ver detalhes de uma UC específica, pergunte pelo número dela (ex: "detalhes da UC ${addressMatches[0].uc}").*`,
          quickAction: {
            label: "Inspecionar na Base de Faturas",
            action: () => {
              setSearchFilter(searchedTerm || query);
              setIsDataViewerOpen(true);
            },
          },
        };
      }
    } else if (
      isAddressExplicit &&
      !q.includes("multa") &&
      !q.includes("demanda") &&
      !q.includes("reativ") &&
      !q.includes("usina") &&
      !q.includes("pis") &&
      !q.includes("cofins") &&
      !q.includes("icms") &&
      !q.includes("cip")
    ) {
      const searchedTerm = extractAddressSearchTerm(query);
      const ucsComEndereco = uniqueUcsList.filter((u) => u.endereco && u.endereco !== "---" && u.endereco.trim().length > 0);
      const exemplos = ucsComEndereco
        .slice(0, 3)
        .map((u) => `- *"${u.endereco}"* (UC **${u.uc}**)`)
        .join("\n");

      return {
        text: `### ⚠️ Nenhuma UC Localizada para o Endereço Informado\n\n` +
          `Não foi encontrada nenhuma Unidade Consumidora (UC) correspondente ao termo **"${searchedTerm || query}"** na base de dados carregada.\n\n` +
          `**Dicas para localização de UCs:**\n` +
          `- Tente buscar por **apenas uma parte do endereço**, como o nome principal da rua ou avenida (ex: *"Afonso Pena"* ou *"Maracaju"*).\n` +
          `- Verifique se o endereço foi importado ou cadastrado para essa UC na aba **Gestão de Faturas** (botão *Editar Fatura* ou *Importar Mapeamento UC/Endereço*).\n` +
          `- Atualmente, **${ucsComEndereco.length} UCs** possuem endereço completo cadastrado na base.\n\n` +
          (exemplos ? `**Exemplos de endereços válidos cadastrados no sistema:**\n${exemplos}` : ""),
        quickAction: {
          label: "Ver Faturas para Cadastrar Endereço",
          action: () => setActiveTab("faturas"),
        },
      };
    }

    // 2. Tributos (PIS, COFINS, ICMS, CIP, Encargos)
    if (q.includes("pis") || q.includes("cofins") || q.includes("icms") || q.includes("cip") || q.includes("imposto") || q.includes("tribut")) {
      return {
        text: `### 🏛️ Consolidação de Tributos e Encargos (52 Campos Carregados)\n\n` +
          `- **Total de ICMS:** ${fmtMoney(calculatedContext.totalIcms)}\n` +
          `- **Total de PIS:** ${fmtMoney(calculatedContext.totalPis)}\n` +
          `- **Total de COFINS:** ${fmtMoney(calculatedContext.totalCofins)}\n` +
          `- **Total de PIS/COFINS Somados:** ${fmtMoney(calculatedContext.totalPis + calculatedContext.totalCofins)}\n` +
          `- **Total de CIP (Contrib. Iluminação Pública):** ${fmtMoney(calculatedContext.totalCip)}\n` +
          `- **Outros Encargos:** ${fmtMoney(calculatedContext.totalOutrosEncargos)}\n\n` +
          `*Todos esses valores são totalizados diretamente a partir dos campos fiscais de todas as faturas faturadas.*`,
        quickAction: {
          label: "Ver Relatório Financeiro",
          action: () => setActiveTab("relatorio"),
        },
      };
    }

    // 3. Consumo Grupo B & Intermediária
    if (q.includes("grupo b") || q.includes("intermediária") || q.includes("intermediaria")) {
      return {
        text: `### 💡 Consumo Grupo B & Horário Intermediário\n\n` +
          `- **Consumo Grupo B Total:** ${fmtNum(calculatedContext.totalConsumoGrupoB, "kWh")}\n` +
          `- **Valor Consumo Grupo B Total:** ${fmtMoney(calculatedContext.totalValorConsumoGrupoB)}\n` +
          `- **Consumo Intermediária Total:** ${fmtNum(calculatedContext.totalConsumoIntermediaria, "kWh")}\n` +
          `- **Consumo Ponta Total:** ${fmtNum(calculatedContext.totalConsumoPonta, "kWh")} (${fmtMoney(calculatedContext.totalValorConsumoPonta)})\n` +
          `- **Consumo Fora Ponta Total:** ${fmtNum(calculatedContext.totalConsumoForaPonta, "kWh")} (${fmtMoney(calculatedContext.totalValorConsumoForaPonta)})\n\n` +
          `*Para UCs convencionais do Grupo B, o consumo e valor são faturados em campo específico de energia ativa.*`,
        quickAction: {
          label: "Ver Dashboard Analítico",
          action: () => setActiveTab("dashboard"),
        },
      };
    }

    // 4. Demanda Contratada, Medida e Não Consumida
    if (q.includes("demanda") || q.includes("contratad") || q.includes("subutiliz") || q.includes("não consumida")) {
      return {
        text: `### 📈 Demanda Medida vs Contratada & Subutilização\n\n` +
          `**Demanda Medida:**\n` +
          `- Ponta: ${fmtNum(calculatedContext.totalDemandaMedidaPonta, "kW")} (${fmtMoney(calculatedContext.totalValorDemandaMedidaPonta)})\n` +
          `- Fora Ponta: ${fmtNum(calculatedContext.totalDemandaMedidaForaPonta, "kW")} (${fmtMoney(calculatedContext.totalValorDemandaMedidaForaPonta)})\n\n` +
          `**Demanda Não Consumida (Subutilização):**\n` +
          `- Ponta: ${fmtNum(calculatedContext.totalDemandaNaoConsumidaPonta, "kW")} (${fmtMoney(calculatedContext.totalValorDemandaNaoConsumidaPonta)})\n` +
          `- Fora Ponta: ${fmtNum(calculatedContext.totalDemandaNaoConsumidaForaPonta, "kW")} (${fmtMoney(calculatedContext.totalValorDemandaNaoConsumidaForaPonta)})\n` +
          `- **Custo Total Subutilizado:** ${fmtMoney(calculatedContext.totalValorDemandaNaoConsumidaPonta + calculatedContext.totalValorDemandaNaoConsumidaForaPonta)}\n\n` +
          `**Ultrapassagem de Demanda Contratada:**\n` +
          `- Total Penalidade: ${fmtMoney(calculatedContext.totalMultaUltrapassagem)}\n` +
          `  - Ponta: ${fmtNum(calculatedContext.totalUltrapassagemPonta, "kW")} (${fmtMoney(calculatedContext.totalValorUltrapassagemPonta)})\n` +
          `  - Fora Ponta: ${fmtNum(calculatedContext.totalUltrapassagemForaPonta, "kW")} (${fmtMoney(calculatedContext.totalValorUltrapassagemForaPonta)})`,
        quickAction: {
          label: "Ver Monitoramento de Demanda",
          action: () => setActiveTab("monitoramento"),
        },
      };
    }

    // 5. Multas e Ultrapassagens
    if (q.includes("multa") || q.includes("ultrapassa") || q.includes("penalidade")) {
      const topList = calculatedContext.finesByUc
        .slice(0, 5)
        .map(
          (u, idx) =>
            `${idx + 1}. **UC ${u.uc}** (${u.cidade} / ${u.gerencia}): Total ${fmtMoney(u.multas)} (Ultrapassagem: ${fmtMoney(u.ultrapassagem)}, Reativo: ${fmtMoney(u.reativo)})`
        )
        .join("\n");

      return {
        text: `### 🚨 Diagnóstico de Multas e Penalidades\n\n` +
          `- **Total Geral de Multas Identificadas:** ${fmtMoney(calculatedContext.totalMultasGeral)}\n` +
          `- **Multas de Ultrapassagem:** ${fmtMoney(calculatedContext.totalMultaUltrapassagem)}\n` +
          `  - Ponta: ${fmtNum(calculatedContext.totalUltrapassagemPonta, "kW")} (${fmtMoney(calculatedContext.totalValorUltrapassagemPonta)})\n` +
          `  - Fora Ponta: ${fmtNum(calculatedContext.totalUltrapassagemForaPonta, "kW")} (${fmtMoney(calculatedContext.totalValorUltrapassagemForaPonta)})\n` +
          `- **Multas por Excedente Reativo:** ${fmtMoney(calculatedContext.totalMultaReativo)}\n` +
          `  - Ponta: ${fmtNum(calculatedContext.totalReativoPonta, "kVArh")} (${fmtMoney(calculatedContext.totalValorReativoPonta)})\n` +
          `  - Fora Ponta: ${fmtNum(calculatedContext.totalReativoForaPonta, "kVArh")} (${fmtMoney(calculatedContext.totalValorReativoForaPonta)})\n\n` +
          `#### Principais UCs com Penalidades:\n${topList || "Nenhuma UC com multas registradas."}`,
        quickAction: {
          label: "Visualizar Módulo de Multas",
          action: () => setActiveTab("multas"),
        },
      };
    }

    // 6. Monitoramento Reativo
    if (q.includes("reativ") || q.includes("kvarh") || q.includes("fator de potencia")) {
      return {
        text: `### ⚡ Monitoramento de Energia Reativa Excedente\n\n` +
          `- **Custo Total de Reativo Excedente:** ${fmtMoney(calculatedContext.totalMultaReativo)}\n` +
          `- **Reativa Excedente Ponta:** ${fmtNum(calculatedContext.totalReativoPonta, "kVArh")} (${fmtMoney(calculatedContext.totalValorReativoPonta)})\n` +
          `- **Reativa Excedente Fora Ponta:** ${fmtNum(calculatedContext.totalReativoForaPonta, "kVArh")} (${fmtMoney(calculatedContext.totalValorReativoForaPonta)})\n\n` +
          `*A adequação de bancos de capacitores elimina 100% dessas cobranças de excedente reativo.*`,
        quickAction: {
          label: "Ver Monitoramento Reativo",
          action: () => setActiveTab("monitoramento_reativo"),
        },
      };
    }

    // 7. GDI, Usinas e Saldo GD
    if (q.includes("usina") || q.includes("solar") || q.includes("gdi") || q.includes("injetad") || q.includes("saldo gd") || q.includes("compensad")) {
      return {
        text: `### ☀️ Geração Distribuída (GDI), Usinas & Saldo GD\n\n` +
          `- **GDI oUC:** ${fmtNum(calculatedContext.totalGdiOucKwh, "kWh")} (${fmtMoney(calculatedContext.totalValorGdiOuc)})\n` +
          `- **GDI mUC:** ${fmtNum(calculatedContext.totalGdiMucKwh, "kWh")} (${fmtMoney(calculatedContext.totalValorGdiMuc)})\n` +
          `- **Energia Injetada Total:** ${fmtNum(calculatedContext.totalEnergiaInjetada, "kWh")}\n` +
          `- **Energia Compensada Total:** ${fmtNum(calculatedContext.totalEnergiaCompensada, "kWh")}\n` +
          `- **Saldo Acumulado GD Total:** ${fmtNum(calculatedContext.totalSaldoGD, "kWh")}\n\n` +
          `*Todos os créditos das usinas PPP Fotovoltaica e Usinas Sanesul estão mapeados individualmente por fatura.*`,
        quickAction: {
          label: "Ver Monitoramento Usinas",
          action: () => setActiveTab("monitoramento_usinas"),
        },
      };
    }

    // 8. Resumo Completo de Todos os 52 Campos
    return {
      text: `### 📊 Resumo Executivo da Base de Dados (52 Campos Carregados)\n\n` +
        `- **Total de Faturas Carregadas:** ${calculatedContext.totalFaturas.toLocaleString("pt-BR")}\n` +
        `- **Unidades Consumidoras Únicas (UCs):** ${calculatedContext.qtdUcs}\n` +
        `- **Gerências Atendidas:** ${calculatedContext.qtdGerencias} | **Cidades:** ${calculatedContext.qtdCidades}\n` +
        `- **Faturamento Total:** ${fmtMoney(calculatedContext.totalFaturado)}\n` +
        `- **Tributos (PIS + COFINS + ICMS + CIP):** ${fmtMoney(calculatedContext.totalPis + calculatedContext.totalCofins + calculatedContext.totalIcms + calculatedContext.totalCip)}\n` +
        `- **Multas (Ultrapassagem + Reativo):** ${fmtMoney(calculatedContext.totalMultasGeral)}\n` +
        `- **GDI (oUC + mUC Compensado):** ${fmtMoney(calculatedContext.totalValorGdiOuc + calculatedContext.totalValorGdiMuc)}\n` +
        `- **Saldo GD Acumulado:** ${fmtNum(calculatedContext.totalSaldoGD, "kWh")}\n\n` +
        `Você pode me perguntar sobre uma **UC específica**, sobre **tributos (ICMS/PIS/COFINS/CIP)**, **multas**, **demanda**, **reativo** ou clicar em **"Explorar Faturas Carregadas"** abaixo para inspecionar linha a linha.`,
      quickAction: {
        label: "Ver Dashboard Analítico",
        action: () => setActiveTab("dashboard"),
      },
    };
  };

  const handleSend = async (userText?: string) => {
    const query = (userText || input).trim();
    if (!query || isLoading) return;

    const userMessage: Message = {
      id: crypto.randomUUID(),
      sender: "user",
      text: query,
      timestamp: new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }),
    };

    setMessages((prev) => [...prev, userMessage]);
    setInput("");
    setIsLoading(true);

    const apiKey =
      (typeof process !== "undefined"
        ? process.env.GEMINI_API_KEY || process.env.API_KEY
        : "") ||
      import.meta.env.VITE_GEMINI_API_KEY ||
      "";

    try {
      if (apiKey) {
        const ai = new GoogleGenAI({ apiKey });

        // Identifica se a query faz referência a alguma UC específica por dígitos
        const ucMatch = query.match(/\b\d{6,15}\b/);
        let relevantRecords: FaturaAssistente[] = [];
        if (ucMatch) {
          relevantRecords = faturasCarregadas.filter((f) => f.uc === ucMatch[0] || f.uc.includes(ucMatch[0])).slice(0, 12);
        } else {
          // Busca por correspondência de endereço completo ou parcial
          const addressMatches = searchUcsByAddress(query);
          if (addressMatches.length > 0) {
            const matchedUcsSet = new Set(addressMatches.map((m) => m.uc));
            relevantRecords = faturasCarregadas.filter((f) => matchedUcsSet.has(f.uc)).slice(0, 15);
          } else {
            // Top UCs ou amostragem relevante
            relevantRecords = faturasCarregadas.slice(0, 8);
          }
        }

        const systemPrompt = `Você é o Assistente do Portal de Inteligência Energética da SANESUL (COTAA).
Sua missão é consultar, esclarecer e fornecer análises precisas sobre faturas de energia elétrica, multas, demanda, reativo, usinas fotovoltaicas e a Central de Ações, além de LOCALIZAR UNIDADES CONSUMIDORAS (UCs) PELO ENDEREÇO COMPLETO OU PARTE DELE.

REGRAS RÍGIDAS DE ATUAÇÃO:
1. Você atua ESTRITAMENTE em MODO CONSULTA: você só consulta e explica dados já calculados no sistema.
2. Você NUNCA altera nada no cadastro ou nos parâmetros das faturas/UCs.
3. Você possui todos os seguintes 52 CAMPOS de todas as faturas carregados na base de conhecimento, incluindo o ENDEREÇO COMPLETO de cada UC:
   - UC, Endereço, Gerência, LOCINS, Cidade, Tipo, Mercado, Concessionária, Mês Referência, Ano Leitura, Vencimento, Nota Fiscal, Modalidade Tarifária, Subgrupo, Valor Total (R$), Consumo Grupo B (kWh e R$), Consumo Intermediária, Demanda Contratada Ponta/Fora Ponta/Todos os Períodos (kW), Demanda Medida Ponta/Fora Ponta (kW e R$), Consumo Ponta/Fora Ponta (kWh e R$), Demanda Não Consumida Ponta/Fora Ponta (kW e R$), Ultrapassagem Ponta/Fora Ponta (kW e R$), Reativa Excedente Ponta/Fora Ponta (kVArh e R$), GDI oUC (kWh e R$), GDI mUC (kWh e R$), Energia Injetada (kWh), Energia Compensada (kWh), Saldo Acumulado GD, CIP (R$), Outros Encargos (R$), PIS (R$), COFINS (R$), ICMS (R$).
4. LOCALIZAÇÃO DE UC POR ENDEREÇO (COMPLETO OU PARCIAL):
   - Se o usuário perguntar qual a UC de determinado endereço, rua, avenida, logradouro, bairro ou número, identifique com precisão todas as UCs correspondentes na base carregada.
   - Forneça: Número da UC, Endereço Completo, Cidade, Gerência, LOCINS, Tipo e Histórico de faturamento.
   - Se não encontrar correspondência, oriente o usuário com termos parciais e informe que endereços podem ser atualizados na aba Gestão de Faturas.

TOTAIS CONSOLIDADOS DO SISTEMA:
- Total Faturas: ${calculatedContext.totalFaturas}
- Valor Faturado Total: R$ ${calculatedContext.totalFaturado.toFixed(2)}
- Total ICMS: R$ ${calculatedContext.totalIcms.toFixed(2)} | PIS: R$ ${calculatedContext.totalPis.toFixed(2)} | COFINS: R$ ${calculatedContext.totalCofins.toFixed(2)} | CIP: R$ ${calculatedContext.totalCip.toFixed(2)}
- Total Multas Geral: R$ ${calculatedContext.totalMultasGeral.toFixed(2)}
- Multas Ultrapassagem Demanda: R$ ${calculatedContext.totalMultaUltrapassagem.toFixed(2)} (Ponta: ${calculatedContext.totalUltrapassagemPonta.toFixed(1)} kW / R$ ${calculatedContext.totalValorUltrapassagemPonta.toFixed(2)}, Fora Ponta: ${calculatedContext.totalUltrapassagemForaPonta.toFixed(1)} kW / R$ ${calculatedContext.totalValorUltrapassagemForaPonta.toFixed(2)})
- Multas Reativo Excedente: R$ ${calculatedContext.totalMultaReativo.toFixed(2)} (Ponta: ${calculatedContext.totalReativoPonta.toFixed(1)} kVArh / R$ ${calculatedContext.totalValorReativoPonta.toFixed(2)}, Fora Ponta: ${calculatedContext.totalReativoForaPonta.toFixed(1)} kVArh / R$ ${calculatedContext.totalValorReativoForaPonta.toFixed(2)})
- Demanda Não Consumida: R$ ${(calculatedContext.totalValorDemandaNaoConsumidaPonta + calculatedContext.totalValorDemandaNaoConsumidaForaPonta).toFixed(2)}
- Total GDI Compensado: R$ ${(calculatedContext.totalValorGdiOuc + calculatedContext.totalValorGdiMuc).toFixed(2)} (oUC: ${calculatedContext.totalGdiOucKwh.toFixed(0)} kWh, mUC: ${calculatedContext.totalGdiMucKwh.toFixed(0)} kWh)
- Saldo GD Total: ${calculatedContext.totalSaldoGD.toFixed(0)} kWh

AMOSTRA DE REGISTROS RELEVANTES CARREGADOS COM OS 52 CAMPOS:
${JSON.stringify(relevantRecords)}

Responda em português brasileiro de forma técnica, limpa, direta, com valores formatados em moeda (R$) e grandezas (kWh, kW, kVArh).`;

        // Timeout de 7 segundos para resposta ágil sem congelamento da UI
        const apiPromise = ai.models.generateContent({
          model: "gemini-2.5-flash",
          contents: query,
          config: {
            systemInstruction: systemPrompt,
          },
        });

        const timeoutPromise = new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error("Timeout na consulta à IA")), 7000)
        );

        const response = (await Promise.race([apiPromise, timeoutPromise])) as any;
        const replyText = response?.text || generateLocalResponse(query).text;
        const localMeta = generateLocalResponse(query);

        setMessages((prev) => [
          ...prev,
          {
            id: crypto.randomUUID(),
            sender: "assistant",
            text: replyText,
            timestamp: new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }),
            quickAction: localMeta.quickAction,
          },
        ]);
      } else {
        await new Promise((resolve) => setTimeout(resolve, 350));
        const localRes = generateLocalResponse(query);
        setMessages((prev) => [
          ...prev,
          {
            id: crypto.randomUUID(),
            sender: "assistant",
            text: localRes.text,
            timestamp: new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }),
            quickAction: localRes.quickAction,
          },
        ]);
      }
    } catch (err) {
      console.warn("Gemini query error, using local engine:", err);
      const localRes = generateLocalResponse(query);
      setMessages((prev) => [
        ...prev,
        {
          id: crypto.randomUUID(),
          sender: "assistant",
          text: localRes.text,
          timestamp: new Date().toLocaleTimeString("pt-BR", { hour: "2-digit", minute: "2-digit" }),
          quickAction: localRes.quickAction,
        },
      ]);
    } finally {
      setIsLoading(false);
    }
  };

  const handleCopy = (id: string, text: string) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const filteredBillsForViewer = useMemo(() => {
    if (!isDataViewerOpen) return [];
    if (!searchFilter.trim()) return faturasCarregadas;
    const s = searchFilter.toLowerCase().trim();
    return faturasCarregadas.filter(
      (f) =>
        f.uc.includes(s) ||
        (f.endereco && f.endereco.toLowerCase().includes(s)) ||
        f.cidade.toLowerCase().includes(s) ||
        f.gerencia.toLowerCase().includes(s) ||
        f.locins.toLowerCase().includes(s) ||
        f.notaFiscal.toLowerCase().includes(s) ||
        f.referenciaFormatada.toLowerCase().includes(s)
    );
  }, [faturasCarregadas, searchFilter, isDataViewerOpen]);

  const pageSize = 25;
  const totalViewerPages = Math.max(1, Math.ceil(filteredBillsForViewer.length / pageSize));
  const paginatedBills = useMemo(() => {
    if (!isDataViewerOpen) return [];
    const start = (viewerPage - 1) * pageSize;
    return filteredBillsForViewer.slice(start, start + pageSize);
  }, [filteredBillsForViewer, viewerPage, isDataViewerOpen]);

  const quickPrompts = [
    { label: "📍 Localizar UC por Endereço", query: "Como localizar uma UC pelo endereço completo ou parte dele?" },
    { label: "📋 UCs com Endereço Cadastrado", query: "Quais e quantas UCs possuem endereço cadastrado?" },
    { label: "Resumo dos 52 Campos", query: "Qual o resumo geral dos 52 campos de todas as faturas carregadas?" },
    { label: "Tributos (ICMS/PIS/COFINS/CIP)", query: "Qual o valor total de ICMS, PIS, COFINS, CIP e Outros Encargos?" },
    { label: "Consumo Grupo B & Intermediária", query: "Qual o consumo e valor do Grupo B e o consumo no horário intermediário?" },
    { label: "Demanda Medida vs Contratada", query: "Qual o detalhamento da Demanda Medida, Contratada e Não Consumida?" },
    { label: "Multas de Ultrapassagem & Reativo", query: "Quais são os valores de ultrapassagem de ponta/fora ponta e excedente reativo?" },
    { label: "Usinas, Injeção e Saldo GD", query: "Qual o total de GDI oUC, GDI mUC, energia injetada e saldo GD acumulado?" },
  ];

  return (
    <div className="py-4 space-y-6 max-w-6xl mx-auto">
      {/* Header do Assistente */}
      <div className="bg-white rounded-3xl p-6 md:p-8 border border-slate-200/80 shadow-sm relative overflow-hidden">
        <div className="absolute top-0 right-0 w-96 h-96 bg-gradient-to-bl from-blue-100/60 via-emerald-50/40 to-transparent rounded-full -mr-20 -mt-20 blur-2xl pointer-events-none" />

        <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 relative z-10">
          <div className="flex items-start gap-4">
            <div className="w-14 h-14 rounded-2xl bg-gradient-to-tr from-sanesul-primary to-blue-500 text-white flex items-center justify-center shadow-lg shadow-blue-500/20 flex-shrink-0">
              <Bot size={28} />
            </div>
            <div>
              <div className="flex items-center gap-2 mb-1 flex-wrap">
                <h1 className="text-2xl md:text-3xl font-display font-bold text-slate-800">
                  Assistente do Portal
                </h1>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-blue-100 text-blue-800 border border-blue-200">
                  52 Campos Carregados
                </span>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-indigo-100 text-indigo-800 border border-indigo-200 flex items-center gap-1">
                  <MapPin size={11} />
                  Busca por Endereço
                </span>
                <span className="px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider bg-emerald-100 text-emerald-800 border border-emerald-200">
                  {faturasCarregadas.length} Faturas Indexadas
                </span>
              </div>
              <p className="text-sm text-slate-600 max-w-2xl leading-relaxed">
                Base completa carregada com 52 campos, histórico de faturamento, tributos, multas e localização ágil de UCs por endereço completo ou parcial.
              </p>
            </div>
          </div>

          <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2">
            <button
              onClick={() => setIsDataViewerOpen(true)}
              className="flex items-center justify-center gap-2 px-3.5 py-2 bg-slate-100 hover:bg-slate-200/80 border border-slate-200 rounded-xl text-slate-700 text-xs font-bold transition-all shadow-2xs active:scale-95"
              title="Inspecionar tabela com os 52 campos carregados"
            >
              <Database size={15} className="text-sanesul-primary" />
              <span>Ver Base (52 Campos)</span>
            </button>

            <div className="flex items-center gap-2 px-3.5 py-2 bg-emerald-50 border border-emerald-200 rounded-xl text-emerald-800 text-xs font-bold">
              <ShieldCheck size={16} className="text-emerald-600 flex-shrink-0" />
              <span>Modo Consulta: sem alterações cadastrais.</span>
            </div>
          </div>
        </div>

        {/* Mini KPIs dos 52 Campos */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 mt-6 pt-6 border-t border-slate-100">
          <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              Valor Faturado
            </span>
            <span className="text-sm md:text-base font-bold text-slate-800 block truncate">
              {fmtMoney(calculatedContext.totalFaturado)}
            </span>
          </div>

          <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              Total ICMS
            </span>
            <span className="text-sm md:text-base font-bold text-blue-600 block truncate">
              {fmtMoney(calculatedContext.totalIcms)}
            </span>
          </div>

          <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              PIS + COFINS
            </span>
            <span className="text-sm md:text-base font-bold text-indigo-600 block truncate">
              {fmtMoney(calculatedContext.totalPis + calculatedContext.totalCofins)}
            </span>
          </div>

          <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              Total CIP
            </span>
            <span className="text-sm md:text-base font-bold text-amber-600 block truncate">
              {fmtMoney(calculatedContext.totalCip)}
            </span>
          </div>

          <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              Multas (Ultrap. + Reat.)
            </span>
            <span className="text-sm md:text-base font-bold text-red-600 block truncate">
              {fmtMoney(calculatedContext.totalMultasGeral)}
            </span>
          </div>

          <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400 block mb-0.5">
              Saldo GD Acumulado
            </span>
            <span className="text-sm md:text-base font-bold text-emerald-600 block truncate">
              {fmtNum(calculatedContext.totalSaldoGD, "kWh")}
            </span>
          </div>
        </div>
      </div>

      {/* Caixa de Mensagens */}
      <div className="bg-white rounded-3xl border border-slate-200/80 shadow-sm overflow-hidden flex flex-col h-[580px]">
        {/* Lista de Mensagens */}
        <div className="flex-1 overflow-y-auto p-4 md:p-6 space-y-4">
          {messages.map((m) => {
            const isUser = m.sender === "user";
            return (
              <div
                key={m.id}
                className={`flex gap-3 ${isUser ? "justify-end" : "justify-start"}`}
              >
                {!isUser && (
                  <div className="w-8 h-8 rounded-xl bg-blue-100 text-sanesul-primary flex items-center justify-center flex-shrink-0 mt-1 shadow-xs">
                    <Bot size={17} />
                  </div>
                )}

                <div
                  className={`max-w-[85%] md:max-w-[80%] rounded-2xl p-4 text-sm leading-relaxed shadow-xs relative group ${
                    isUser
                      ? "bg-sanesul-primary text-white rounded-tr-xs"
                      : "bg-slate-50 border border-slate-200/80 text-slate-800 rounded-tl-xs"
                  }`}
                >
                  <div className="whitespace-pre-wrap font-sans space-y-2">
                    {m.text}
                  </div>

                  {m.quickAction && (
                    <div className="mt-3 pt-3 border-t border-slate-200/80 flex items-center justify-between">
                      <span className="text-xs text-slate-500 font-medium">Atalho direto:</span>
                      <button
                        onClick={m.quickAction.action}
                        className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition-all shadow-xs"
                      >
                        <span>{m.quickAction.label}</span>
                        <ArrowRight size={13} />
                      </button>
                    </div>
                  )}

                  <div className="flex items-center justify-between gap-4 mt-2 pt-1 border-t border-black/5 text-[10px] opacity-70">
                    <span>{m.timestamp}</span>
                    <button
                      onClick={() => handleCopy(m.id, m.text)}
                      className="opacity-0 group-hover:opacity-100 transition-opacity hover:text-blue-500 flex items-center gap-1"
                      title="Copiar mensagem"
                    >
                      {copiedId === m.id ? (
                        <>
                          <Check size={11} className="text-emerald-500" />
                          <span className="text-emerald-500">Copiado</span>
                        </>
                      ) : (
                        <>
                          <Copy size={11} />
                          <span>Copiar</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </div>
            );
          })}

          {isLoading && (
            <div className="flex gap-3 justify-start">
              <div className="w-8 h-8 rounded-xl bg-blue-100 text-sanesul-primary flex items-center justify-center flex-shrink-0 mt-1 animate-pulse">
                <Bot size={17} />
              </div>
              <div className="bg-slate-50 border border-slate-200/80 rounded-2xl rounded-tl-xs p-4 text-xs font-bold text-slate-500 flex items-center gap-2">
                <span className="w-2 h-2 rounded-full bg-blue-500 animate-ping" />
                <span>Consultando os 52 campos da base de faturas...</span>
              </div>
            </div>
          )}

          <div ref={messagesEndRef} />
        </div>

        {/* Sugestões Rápidas */}
        <div className="p-3 bg-slate-50/80 border-t border-slate-100 flex items-center gap-2 overflow-x-auto no-scrollbar">
          <span className="text-[11px] font-bold uppercase tracking-wider text-slate-400 pl-2 whitespace-nowrap flex items-center gap-1">
            <Sparkles size={13} className="text-amber-500" />
            Sugestões:
          </span>
          {quickPrompts.map((p, idx) => (
            <button
              key={idx}
              onClick={() => handleSend(p.query)}
              disabled={isLoading}
              className="px-3 py-1.5 bg-white hover:bg-blue-50 hover:border-blue-200 border border-slate-200 text-slate-600 hover:text-blue-700 text-xs font-bold rounded-xl whitespace-nowrap transition-all shadow-2xs active:scale-95 disabled:opacity-50"
            >
              {p.label}
            </button>
          ))}
        </div>

        {/* Input Bar */}
        <div className="p-4 bg-white border-t border-slate-200/80">
          <form
            onSubmit={(e) => {
              e.preventDefault();
              handleSend();
            }}
            className="flex items-center gap-2"
          >
            <input
              type="text"
              value={input}
              onChange={(e) => setInput(e.target.value)}
              placeholder="Localize qualquer UC pelo endereço completo/parcial, número da UC, multas, demanda, tributos..."
              disabled={isLoading}
              className="flex-1 px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:bg-white focus:ring-2 focus:ring-sanesul-primary/30 focus:border-sanesul-primary outline-none transition-all"
            />
            <button
              type="submit"
              disabled={!input.trim() || isLoading}
              className="px-5 py-3 bg-sanesul-primary hover:bg-sanesul-secondary text-white rounded-xl text-sm font-bold flex items-center gap-2 transition-all shadow-md shadow-sanesul-primary/20 disabled:opacity-50 disabled:shadow-none active:scale-95"
            >
              <Send size={16} />
              <span className="hidden sm:inline">Enviar</span>
            </button>
          </form>
        </div>
      </div>

      {/* Modal / Inspecionador da Base Completa dos 52 Campos */}
      {isDataViewerOpen && (
        <div className="fixed inset-0 z-[120] flex items-center justify-center bg-black/60 backdrop-blur-sm p-4">
          <div className="bg-white w-full max-w-6xl max-h-[90vh] rounded-3xl shadow-2xl flex flex-col overflow-hidden animate-in zoom-in duration-200">
            {/* Modal Header */}
            <div className="p-5 border-b border-slate-200 flex items-center justify-between gap-4 bg-slate-50">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-blue-100 text-sanesul-primary flex items-center justify-center">
                  <Database size={20} />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-800">
                    Base de Faturas Carregadas (52 Campos)
                  </h3>
                  <p className="text-xs text-slate-500">
                    {faturasCarregadas.length} faturas indexadas e prontas para consulta pelo Assistente
                  </p>
                </div>
              </div>

              <button
                onClick={() => setIsDataViewerOpen(false)}
                className="p-2 text-slate-400 hover:text-slate-600 rounded-xl hover:bg-slate-200/60 transition-all"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Search Bar */}
            <div className="p-4 border-b border-slate-100 flex items-center gap-3 bg-white">
              <div className="relative flex-1">
                <Search size={16} className="absolute left-3.5 top-3 text-slate-400" />
                <input
                  type="text"
                  value={searchFilter}
                  onChange={(e) => setSearchFilter(e.target.value)}
                  placeholder="Filtrar por UC, Endereço, Cidade, Gerência, LOCINS, NF ou Mês/Ano..."
                  className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium focus:bg-white outline-none focus:ring-2 focus:ring-sanesul-primary/20"
                />
              </div>
              <span className="text-xs font-bold text-slate-500 whitespace-nowrap">
                {filteredBillsForViewer.length} registro(s)
              </span>
            </div>

            {/* Modal Table with horizontal scroll */}
            <div className="flex-1 overflow-auto p-4">
              <table className="w-full border-collapse text-left text-xs whitespace-nowrap">
                <thead className="bg-slate-100/90 text-slate-700 sticky top-0 font-bold uppercase tracking-wider text-[10px] z-10">
                  <tr>
                    <th className="p-2.5 border-b">UC</th>
                    <th className="p-2.5 border-b">Gerência</th>
                    <th className="p-2.5 border-b">LOCINS</th>
                    <th className="p-2.5 border-b">Cidade</th>
                    <th className="p-2.5 border-b">Endereço</th>
                    <th className="p-2.5 border-b">Tipo</th>
                    <th className="p-2.5 border-b">Mercado</th>
                    <th className="p-2.5 border-b">Concessionária</th>
                    <th className="p-2.5 border-b">Mês Ref</th>
                    <th className="p-2.5 border-b">Ano Leitura</th>
                    <th className="p-2.5 border-b">Vencimento</th>
                    <th className="p-2.5 border-b">Nota Fiscal</th>
                    <th className="p-2.5 border-b">Modalidade</th>
                    <th className="p-2.5 border-b">Subgrupo</th>
                    <th className="p-2.5 border-b text-right">Valor Total</th>
                    <th className="p-2.5 border-b text-right">Cons. Grupo B</th>
                    <th className="p-2.5 border-b text-right">Val. Cons. Grupo B</th>
                    <th className="p-2.5 border-b text-right">Cons. Intermed.</th>
                    <th className="p-2.5 border-b text-right">Dem. Cont. P</th>
                    <th className="p-2.5 border-b text-right">Dem. Cont. FP</th>
                    <th className="p-2.5 border-b text-right">Dem. Todos Per.</th>
                    <th className="p-2.5 border-b text-right">Dem. Med. P</th>
                    <th className="p-2.5 border-b text-right">Val. Dem. Med. P</th>
                    <th className="p-2.5 border-b text-right">Dem. Med. FP</th>
                    <th className="p-2.5 border-b text-right">Val. Dem. Med. FP</th>
                    <th className="p-2.5 border-b text-right">Cons. Ponta</th>
                    <th className="p-2.5 border-b text-right">Val. Cons. Ponta</th>
                    <th className="p-2.5 border-b text-right">Cons. Fora Ponta</th>
                    <th className="p-2.5 border-b text-right">Val. Cons. Fora Ponta</th>
                    <th className="p-2.5 border-b text-right">Dem. Não Cons. P</th>
                    <th className="p-2.5 border-b text-right">Val. Não Cons. P</th>
                    <th className="p-2.5 border-b text-right">Dem. Não Cons. FP</th>
                    <th className="p-2.5 border-b text-right">Val. Não Cons. FP</th>
                    <th className="p-2.5 border-b text-right">Ultrap. P</th>
                    <th className="p-2.5 border-b text-right">Val. Ultrap. P</th>
                    <th className="p-2.5 border-b text-right">Ultrap. FP</th>
                    <th className="p-2.5 border-b text-right">Val. Ultrap. FP</th>
                    <th className="p-2.5 border-b text-right">Reat. Exced. P</th>
                    <th className="p-2.5 border-b text-right">Val. Reat. Exced. P</th>
                    <th className="p-2.5 border-b text-right">Reat. Exced. FP</th>
                    <th className="p-2.5 border-b text-right">Val. Reat. Exced. FP</th>
                    <th className="p-2.5 border-b text-right">GDI oUC (kWh)</th>
                    <th className="p-2.5 border-b text-right">Val. GDI oUC</th>
                    <th className="p-2.5 border-b text-right">GDI mUC (kWh)</th>
                    <th className="p-2.5 border-b text-right">Val. GDI mUC</th>
                    <th className="p-2.5 border-b text-right">Energia Injetada</th>
                    <th className="p-2.5 border-b text-right">Energia Compensada</th>
                    <th className="p-2.5 border-b text-right">Saldo GD</th>
                    <th className="p-2.5 border-b text-right">CIP</th>
                    <th className="p-2.5 border-b text-right">Outros Encargos</th>
                    <th className="p-2.5 border-b text-right">PIS</th>
                    <th className="p-2.5 border-b text-right">COFINS</th>
                    <th className="p-2.5 border-b text-right">ICMS</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredBillsForViewer.slice(0, 100).map((f) => (
                    <tr key={f.id} className="hover:bg-slate-50 transition-colors">
                      <td className="p-2.5 font-bold text-sanesul-primary">{f.uc}</td>
                      <td className="p-2.5 text-slate-600">{f.gerencia}</td>
                      <td className="p-2.5 text-slate-600">{f.locins}</td>
                      <td className="p-2.5 text-slate-600">{f.cidade}</td>
                      <td className="p-2.5 text-slate-600 max-w-[200px] truncate" title={f.endereco}>{f.endereco}</td>
                      <td className="p-2.5 text-slate-600">{f.tipo}</td>
                      <td className="p-2.5 text-slate-600">{f.mercado}</td>
                      <td className="p-2.5 text-slate-600">{f.concessionaria}</td>
                      <td className="p-2.5 text-slate-600">{f.mesReferencia}</td>
                      <td className="p-2.5 text-slate-600">{f.anoLeitura}</td>
                      <td className="p-2.5 text-slate-600">{f.vencimento}</td>
                      <td className="p-2.5 text-slate-600">{f.notaFiscal}</td>
                      <td className="p-2.5 text-slate-600">{f.modalidadeTarifaria}</td>
                      <td className="p-2.5 text-slate-600">{f.subgrupo}</td>
                      <td className="p-2.5 text-right font-bold text-slate-800">{fmtMoney(f.valorTotal)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.consumoGrupoB)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorConsumoGrupoB)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.consumoIntermediaria)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaContratadaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaContratadaForaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaTodosPeriodos)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaMedidaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorDemandaMedidaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaMedidaForaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorDemandaMedidaForaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.consumoPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorConsumoPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.consumoForaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorConsumoForaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaNaoConsumidaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorDemandaNaoConsumidaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtNum(f.demandaNaoConsumidaForaPonta)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.valorDemandaNaoConsumidaForaPonta)}</td>
                      <td className="p-2.5 text-right text-red-600">{fmtNum(f.ultrapassagemPonta)}</td>
                      <td className="p-2.5 text-right text-red-600">{fmtMoney(f.valorUltrapassagemPonta)}</td>
                      <td className="p-2.5 text-right text-red-600">{fmtNum(f.ultrapassagemForaPonta)}</td>
                      <td className="p-2.5 text-right text-red-600">{fmtMoney(f.valorUltrapassagemForaPonta)}</td>
                      <td className="p-2.5 text-right text-amber-600">{fmtNum(f.reativaExcedentePonta)}</td>
                      <td className="p-2.5 text-right text-amber-600">{fmtMoney(f.valorReativaExcedentePonta)}</td>
                      <td className="p-2.5 text-right text-amber-600">{fmtNum(f.reativaExcedenteForaPonta)}</td>
                      <td className="p-2.5 text-right text-amber-600">{fmtMoney(f.valorReativaExcedenteForaPonta)}</td>
                      <td className="p-2.5 text-right text-emerald-600">{fmtNum(f.gdiOucKwh)}</td>
                      <td className="p-2.5 text-right text-emerald-600">{fmtMoney(f.valorGdiOuc)}</td>
                      <td className="p-2.5 text-right text-emerald-600">{fmtNum(f.gdiMucKwh)}</td>
                      <td className="p-2.5 text-right text-emerald-600">{fmtMoney(f.valorGdiMuc)}</td>
                      <td className="p-2.5 text-right text-emerald-600">{fmtNum(f.energiaInjetadaKwh)}</td>
                      <td className="p-2.5 text-right text-emerald-600">{fmtNum(f.energiaCompensadaKwh)}</td>
                      <td className="p-2.5 text-right text-amber-600 font-bold">{fmtNum(f.saldoAcumuladoGD)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.cip)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.outrosEncargos)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.pis)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.cofins)}</td>
                      <td className="p-2.5 text-right text-slate-600">{fmtMoney(f.icms)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredBillsForViewer.length === 0 && (
                <div className="py-12 text-center text-slate-400">
                  Nenhuma fatura encontrada para o filtro informado.
                </div>
              )}
            </div>

            {/* Modal Footer */}
            <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between text-xs text-slate-500">
              <span>Exibindo até 100 faturas filtradas para rápida inspeção.</span>
              <button
                onClick={() => setIsDataViewerOpen(false)}
                className="px-4 py-2 bg-sanesul-primary text-white rounded-xl font-bold shadow-xs active:scale-95"
              >
                Fechar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
