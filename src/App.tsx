/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useRef, useMemo, useEffect } from 'react';
import domtoimage from 'dom-to-image-more';
import { jsPDF } from 'jspdf';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, ShadingType, ImageRun, VerticalAlign } from 'docx';
import { saveAs } from 'file-saver';
import localforage from 'localforage';
import { GoogleGenAI, Type, GenerateContentResponse, ThinkingLevel } from "@google/genai";
import * as XLSX from 'xlsx';
import { supabase, isSupabaseConfigured } from './lib/supabase';
import { 
  Upload, 
  FileText, 
  Download, 
  Trash2, 
  CheckCircle2, 
  AlertCircle, 
  Loader2,
  Clock,
  Table as TableIcon,
  Plus,
  LayoutDashboard,
  BarChart3,
  TrendingUp,
  TrendingDown,
  FileSpreadsheet,
  DollarSign,
  Zap,
  RotateCcw,
  CheckSquare,
  ChevronRight,
  Calendar,
  Filter,
  X,
  Printer,
  LogOut,
  Pencil,
  Save,
  ArrowLeft,
  Search,
  ChevronDown,
  Calculator,
  ArrowUp,
  ArrowDown,
  Menu,
  User,
  Home,
  BarChart2,
  GitCompare,
  Activity,
  Battery,
  ZapOff,
  Leaf,
  Key,
  AlertTriangle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  ResponsiveContainer, 
  LineChart, 
  Line,
  AreaChart,
  Area,
  ComposedChart,
  PieChart,
  Pie,
  Cell,
  LabelList
} from 'recharts';

// --- Types ---

interface BillData {
  id: string;
  fileName: string;
  uc: string;
  demandaPontaKW: string;
  demandaForaPontaKW: string;
  demandaPotenciaMedidaPonta: string;
  demandaPotenciaMedidaForaPonta: string;
  anoLeitura: string;
  mesReferencia: string;
  consumoKwhPonta: string;
  consumoKwhForaPonta: string;
  valorConsumoKwhPonta: string;
  valorConsumoKwhForaPonta: string;
  valorTotal: string;
  cidade: string;
  demandaPotenciaNaoConsumidaPonta: string;
  demandaPotenciaNaoConsumidaFPonta: string;
  demandaPotenciaAtivaUltrapPonta: string;
  demandaPotenciaAtivaUltrapFPonta: string;
  energiaReativaExcedPonta: string;
  energiaReativaExcedFPonta: string;
  energiaInjetadaKwh: string;
  energiaCompensadaKwh: string;
  valorDemandaPotenciaMedidaPonta: string;
  valorDemandaPotenciaMedidaForaPonta: string;
  valorDemandaPotenciaNaoConsumidaPonta: string;
  valorDemandaPotenciaNaoConsumidaFPonta: string;
  valorDemandaPotenciaAtivaUltrapPonta: string;
  valorDemandaPotenciaAtivaUltrapFPonta: string;
  valorEnergiaReativaExcedPonta: string;
  valorEnergiaReativaExcedFPonta: string;
  energiaAtvInjetadaGDIOUC: string;
  valorEnergiaAtvInjetadaGDIOUC: string;
  energiaAtvInjetadaGDIMUC: string;
  valorEnergiaAtvInjetadaGDIMUC: string;
  cip: string;
  outrosEncargos: string;
  pis?: string;
  cofins?: string;
  icms?: string;
  concessionaria?: string;
  numeroNotaFiscal?: string;
  modalidadeTarifaria?: string;
  subgrupo?: string;
  tipo?: string;
  mercado?: string;
  dataVencimento?: string;
  status: 'pending' | 'processing' | 'completed' | 'error';
  error?: string;
  file?: File;
  progress?: number;
  abortController?: AbortController;
  createdAt?: number;
}

// --- Constants ---

const UCS_LIVRE_MERCADO_LIVRE = new Set([
  "33836", "74312", "117383", "128413", "138553", "152443", "188105", "206447", "247300", "477410", 
  "503164", "525348", "964549", "1097648", "1313470", "1409760", "1626524", "1717627", "1753181", "1755547", 
  "1850184", "3141598", "3192611", "3209502", "3301031", "3315068", "3321872", "3390621", "3390665", "3408044", 
  "3462151", "9000076", "9000079", "9000210", "9000211", "9000483", "9000941", "9000943", "9001396", "12030", 
  "12031", "30723", "30724", "39038513", "41904974", "41905059", "18256767", "18256830"
]);

const UCS_OPER = new Set([
  "1069", "12010", "12018", "12030", "12031", "30723", "30724", "33836", "34548", "34724", "36245", "36246", "36247", "74245", "74246", "74255", "74256", "74262", "74265", "74304", "74305", "74307", "74310", "74312", "74313", "79418", "83685", "94005", "94006", "94009", "94018", "99484", "101432", "101434", "101435", "101436", "102336", "102438", "102690", "105115", "105116", "105120", "108130", "108131", "108132", "108134", "112153", "112154", "112156", "112157", "112158", "112159", "112160", "112161", "117377", "117383", "117384", "121252", "121253", "121254", "121600", "128387", "128389", "128392", "128393", "128394", "128395", "128397", "128413", "132652", "132666", "132670", "132671", "132887", "138550", "138551", "138553", "138555", "138556", "138557", "138558", "152425", "152427", "152428", "152434", "152438", "152439", "152440", "152443", "152444", "152445", "152446", "152447", "152480", "152481", "158691", "158694", "158697", "158717", "163515", "163517", "163520", "163522", "163524", "171999", "172000", "172001", "172002", "172003", "172004", "172005", "172006", "172007", "176812", "176813", "176817", "176819", "179855", "179857", "179858", "179859", "180129", "188105", "188107", "188116", "188117", "189729", "189730", "189731", "189732", "197711", "197712", "197713", "206446", "206447", "209223", "209225", "209226", "209227", "211933", "211934", "211935", "211936", "211937", "216226", "216229", "220535", "220536", "220539", "220542", "220546", "222648", "222650", "223697", "231302", "231303", "231304", "231305", "231306", "233303", "233304", "233306", "238383", "238385", "238407", "244193", "244194", "244195", "244196", "244199", "244200", "244202", "244204", "244205", "244206", "244209", "247298", "247299", "247300", "249349", "249353", "249354", "249355", "249356", "249545", "252248", "252249", "252250", "252251", "253245", "253246", "256199", "256200", "256201", "258183", "258184", "264010", "264986", "264988", "269110", "269118", "270079", "270080", "272605", "272951", "273988", "273989", "274331", "274332", "276549", "276552", "277100", "277101", "279006", "279007", "280857", "280858", "280860", "281808", "281809", "282012", "283247", "283248", "283352", "283480", "283507", "453683", "453827", "455028", "455300", "456560", "456731", "457351", "457765", "457891", "458050", "458289", "458661", "460570", "460571", "461216", "461759", "462534", "462964", "462965", "463730", "463783", "463908", "464215", "464406", "464453", "464549", "464764", "464765", "465134", "465135", "465971", "465974", "466787", "467063", "467064", "467145", "471659", "477410", "477421", "482891", "502682", "511994", "518898", "525348", "525890", "527978", "530714", "533217", "534928", "905272", "908728", "910387", "924494", "925640", "934044", "934045", "938246", "943615", "947571", "964549", "966936", "968191", "973292", "978192", "978395", "983857", "984681", "988341", "996818", "1000652", "1021334", "1029172", "1030939", "1034959", "1047248", "1047259", "1084731", "1089791", "1095819", "1097648", "1101962", "1113637", "1115969", "1126680", "1126687", "1127638", "1136937", "1141976", "1142030", "1142916", "1144446", "1148016", "1151043", "1155477", "1204522", "1223492", "1257197", "1273099", "1273120", "1273146", "1292715", "1306946", "1309765", "1313470", "1320065", "1352920", "1361474", "1388271", "1409760", "1467369", "1479647", "1479890", "1491784", "1543691", "1548221", "1565771", "1572158", "1600326", "1602335", "1617283", "1617386", "1626524", "1626678", "1650695", "1656911", "1665340", "1666808", "1673014", "1673468", "1677710", "1686836", "1687606", "1687779", "1690088", "1698882", "1698936", "1698960", "1699438", "1701676", "1702111", "1717627", "1721067", "1722992", "1741897", "1743417", "1745575", "1745856", "1745857", "1748386", "1753181", "1755547", "1821234", "1822779", "1846263", "1850184", "1877305", "1879309", "1879837", "1893630", "1899594", "1901647", "1908604", "1908949", "1916616", "1923325", "1924437", "1926536", "1936660", "1975585", "2065093", "2093921", "2140053", "2172369", "2188959", "2199504", "2199895", "2203819", "2213694", "2218803", "2243641", "2250642", "2283427", "2337244", "2339906", "2342909", "2392852", "2398903", "2400975", "2414930", "2420193", "2451844", "2480085", "2498097", "2524079", "2527881", "2532624", "2558334", "2559619", "2563141", "2578479", "2588841", "2589075", "2601678", "2601732", "2611984", "2613087", "2656959", "2657735", "2659601", "2664911", "2700471", "2712151", "2713805", "2754088", "2765031", "2797860", "2858514", "2861580", "2868808", "2878685", "2884288", "2884288", "2916984", "2918742", "2922080", "2934674", "2936673", "2941716", "2941788", "2954787", "2963688", "2965103", "2968124", "2993670", "2995984", "2999073", "3000898", "3001597", "3001613", "3005931", "3005999", "3011291", "3017560", "3024170", "3026610", "3026782", "3035283", "3036982", "3047887", "3048079", "3049042", "3062915", "3065851", "3066798", "3066903", "3066996", "3070666", "3074559", "3080035", "3091306", "3092484", "3094151", "3096891", "3101742", "3102586", "3102869", "3106362", "3118096", "3141335", "3141598", "3145893", "3147125", "3149991", "3154041", "3166697", "3175195", "3176636", "3178147", "3181887", "3188966", "3192195", "3192583", "3192611", "3197354", "3197368", "3202095", "3202323", "3204970", "3206144", "3207043", "3209502", "3216941", "3227756", "3228332", "3234575", "3234876", "3235300", "3244861", "3245811", "3248099", "3275218", "3275502", "3282876", "3296302", "3300719", "3301031", "3301943", "3302837", "3310090", "3313761", "3315068", "3320484", "3321872", "3324168", "3331751", "3337032", "3338046", "3338689", "3341371", "3341373", "3341380", "3343169", "3348432", "3351605", "3355489", "3356368", "3356380", "3357484", "3357491", "3362329", "3362339", "3366554", "3366558", "3366584", "3367248", "3367575", "3367708", "3371198", "3375315", "3378194", "3389001", "3389705", "3390621", "3390659", "3390665", "3390948", "3398011", "3401807", "3408044", "3409248", "3409655", "3412949", "3414263", "3417002", "3418302", "3421139", "3422802", "3426808", "3433559", "3440686", "3443659", "3453805", "3462151", "3481163", "3481691", "3498071", "3498079", "3530681", "3537726", "3538203", "3545512", "3560022", "3584509", "3587946", "3617822", "3633305", "3633473", "3635126", "3640464", "3659120", "3706867", "3710835", "3725144", "3727956", "3730751", "3739981", "3762346", "9000076", "9000079", "9000210", "9000211", "9000469", "9000483", "9000941", "9000943"
]);

const UCS_ADM = new Set([
  "158690", "152441", "216228", "2863310", "163521", "280856", "209397", "172705", "3310057", "3344018", 
  "2233618", "254634", "112581", "3181", "457892", "3302798", "456907", "2765050", "3211", "495695", 
  "34643", "34594", "457766", "3645948", "3680569", "277102", "3058557", "3045978", "2632342", "189790", 
  "3331889", "139161", "3280232", "33857", "3243926", "75226", "163944", "128898", "3330204", "189791", 
  "2716454", "3360238", "153011", "108347", "3342396", "3495962", "3214000", "931227", "3263428", "3341702",
  "457352", "3242", "3324892", "216861", "503164", "9001396"
]);

const EXCEL_COLUMNS = [
  { header: 'UC', key: 'uc' },
  { header: 'Tipo', key: 'tipo' },
  { header: 'Concessionária', key: 'concessionaria' },
  { header: 'Cidade', key: 'cidade' },
  { header: 'Mês Referência', key: 'mesReferencia' },
  { header: 'Ano Leitura', key: 'anoLeitura' },
  { header: 'Vencimento', key: 'dataVencimento' },
  { header: 'Nota Fiscal', key: 'numeroNotaFiscal' },
  { header: 'Modalidade Tarifária', key: 'modalidadeTarifaria' },
  { header: 'Subgrupo', key: 'subgrupo' },
  { header: 'Valor Total (R$)', key: 'valorTotal' },
  { header: 'Demanda Contratada Ponta (kW)', key: 'demandaPontaKW' },
  { header: 'Demanda Contratada Fora Ponta (kW)', key: 'demandaForaPontaKW' },
  { header: 'Demanda Medida Ponta (kW)', key: 'demandaPotenciaMedidaPonta' },
  { header: 'Valor Demanda Medida Ponta (R$)', key: 'valorDemandaPotenciaMedidaPonta' },
  { header: 'Demanda Medida Fora Ponta (kW)', key: 'demandaPotenciaMedidaForaPonta' },
  { header: 'Valor Demanda Medida Fora Ponta (R$)', key: 'valorDemandaPotenciaMedidaForaPonta' },
  { header: 'Consumo Ponta (kWh)', key: 'consumoKwhPonta' },
  { header: 'Valor Consumo Ponta (R$)', key: 'valorConsumoKwhPonta' },
  { header: 'Consumo Fora Ponta (kWh)', key: 'consumoKwhForaPonta' },
  { header: 'Valor Consumo Fora Ponta (R$)', key: 'valorConsumoKwhForaPonta' },
  { header: 'Demanda Não Consumida Ponta (kW)', key: 'demandaPotenciaNaoConsumidaPonta' },
  { header: 'Valor Demanda Não Consumida Ponta (R$)', key: 'valorDemandaPotenciaNaoConsumidaPonta' },
  { header: 'Demanda Não Consumida Fora Ponta (kW)', key: 'demandaPotenciaNaoConsumidaFPonta' },
  { header: 'Valor Demanda Não Consumida Fora Ponta (R$)', key: 'valorDemandaPotenciaNaoConsumidaFPonta' },
  { header: 'Ultrapassagem Ponta (kW)', key: 'demandaPotenciaAtivaUltrapPonta' },
  { header: 'Valor Ultrapassagem Ponta (R$)', key: 'valorDemandaPotenciaAtivaUltrapPonta' },
  { header: 'Ultrapassagem Fora Ponta (kW)', key: 'demandaPotenciaAtivaUltrapFPonta' },
  { header: 'Valor Ultrapassagem Fora Ponta (R$)', key: 'valorDemandaPotenciaAtivaUltrapFPonta' },
  { header: 'Reativa Excedente Ponta (kVArh)', key: 'energiaReativaExcedPonta' },
  { header: 'Valor Reativa Excedente Ponta (R$)', key: 'valorEnergiaReativaExcedPonta' },
  { header: 'Reativa Excedente Fora Ponta (kVArh)', key: 'energiaReativaExcedFPonta' },
  { header: 'Valor Reativa Excedente Fora Ponta (R$)', key: 'valorEnergiaReativaExcedFPonta' },
  { header: 'Energia Injetada (kWh)', key: 'energiaInjetadaKwh' },
  { header: 'Energia Compensada (kWh)', key: 'energiaCompensadaKwh' },
  { header: 'GDI oUC (kWh)', key: 'energiaAtvInjetadaGDIOUC' },
  { header: 'Valor GDI oUC (R$)', key: 'valorEnergiaAtvInjetadaGDIOUC' },
  { header: 'GDI mUC (kWh)', key: 'energiaAtvInjetadaGDIMUC' },
  { header: 'Valor GDI mUC (R$)', key: 'valorEnergiaAtvInjetadaGDIMUC' },
  { header: 'CIP (R$)', key: 'cip' },
  { header: 'Outros Encargos (R$)', key: 'outrosEncargos' },
  { header: 'PIS (R$)', key: 'pis' },
  { header: 'COFINS (R$)', key: 'cofins' },
  { header: 'ICMS (R$)', key: 'icms' }
];

const EXTRACTION_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    uc: { type: Type.STRING, description: "Código da Unidade Consumidora (UC). Para ENERGISA, use o 'CÓDIGO DO CLIENTE' completo ou a 'MATRÍCULA'. Extraia apenas os números e caracteres identificadores (ex: 10/1069-4)." },
    demandaPontaKW: { type: Type.STRING, description: "Demanda contratada na Ponta em kW. Geralmente encontrada na seção 'Grandezas Contratadas' ou 'Dados da Unidade Consumidora'." },
    demandaForaPontaKW: { type: Type.STRING, description: "Demanda contratada Fora Ponta em kW. Geralmente encontrada na seção 'Grandezas Contratadas'." },
    demandaPotenciaMedidaPonta: { type: Type.STRING, description: "Demanda de Potência Medida no horário de Ponta (kW). Procure na tabela de 'Itens da Fatura'." },
    demandaPotenciaMedidaForaPonta: { type: Type.STRING, description: "Demanda de Potência Medida no horário Fora Ponta (kW). Procure na tabela de 'Itens da Fatura'." },
    anoLeitura: { type: Type.STRING, description: "Ano de referência da fatura (ex: 2025). Extraia apenas os 4 dígitos do ano." },
    mesReferencia: { type: Type.STRING, description: "Mês de referência da fatura (ex: Agosto). Extraia apenas o nome do mês, sem o ano." },
    consumoKwhPonta: { type: Type.STRING, description: "Quantidade de consumo em kWh no horário de Ponta. Procure por 'Consumo Ponta' ou 'Consumo Ativo Ponta'." },
    valorConsumoKwhPonta: { type: Type.STRING, description: "Valor total em R$ do consumo no horário de Ponta." },
    consumoKwhForaPonta: { type: Type.STRING, description: "Quantidade de consumo em kWh no horário Fora Ponta. Procure por 'Consumo Fora Ponta' ou 'Consumo Ativo Fora Ponta'." },
    valorConsumoKwhForaPonta: { type: Type.STRING, description: "Valor total em R$ do consumo no horário Fora Ponta." },
    valorTotal: { type: Type.STRING, description: "Valor total da fatura a pagar (R$). Geralmente em destaque." },
    cidade: { type: Type.STRING, description: "Cidade onde se localiza a Unidade Consumidora." },
    demandaPotenciaNaoConsumidaPonta: { type: Type.STRING, description: "Demanda de Potência Não Consumida - Ponta (kW)." },
    demandaPotenciaNaoConsumidaFPonta: { type: Type.STRING, description: "Demanda de Potência Não Consumida - Fora Ponta (kW)." },
    demandaPotenciaAtivaUltrapPonta: { type: Type.STRING, description: "Demanda de Potência Ativa - Ultrapassagem - Ponta (kW)." },
    demandaPotenciaAtivaUltrapFPonta: { type: Type.STRING, description: "Demanda de Potência Ativa - Ultrapassagem - Fora Ponta (kW)." },
    energiaReativaExcedPonta: { type: Type.STRING, description: "Energia Reativa Excedente - Ponta (kVArh)." },
    energiaReativaExcedFPonta: { type: Type.STRING, description: "Energia Reativa Excedente - Fora Ponta (kVArh)." },
    energiaInjetadaKwh: { type: Type.STRING, description: "Energia Injetada em kWh (Geração Distribuída). Procure por 'Energia Injetada' ou 'GD'." },
    energiaCompensadaKwh: { type: Type.STRING, description: "Energia Compensada em kWh (Geração Distribuída). Procure por 'Energia Compensada' ou 'Consumo Reaturado'." },
    valorDemandaPotenciaMedidaPonta: { type: Type.STRING, description: "Valor em R$ da Demanda de Potência Medida - Ponta." },
    valorDemandaPotenciaMedidaForaPonta: { type: Type.STRING, description: "Valor em R$ da Demanda de Potência Medida - Fora Ponta." },
    valorDemandaPotenciaNaoConsumidaPonta: { type: Type.STRING, description: "Valor em R$ da Demanda Potência Não Consumida - Ponta." },
    valorDemandaPotenciaNaoConsumidaFPonta: { type: Type.STRING, description: "Valor em R$ da Demanda Potência Não Consumida - Fora Ponta." },
    valorDemandaPotenciaAtivaUltrapPonta: { type: Type.STRING, description: "Valor em R$ da Demanda Potência Ativa - Ultrapassagem - Ponta." },
    valorDemandaPotenciaAtivaUltrapFPonta: { type: Type.STRING, description: "Valor em R$ da Demanda Potência Ativa - Ultrapassagem - Fora Ponta." },
    valorEnergiaReativaExcedPonta: { type: Type.STRING, description: "Valor em R$ da Energia Reativa Excedente - Ponta." },
    valorEnergiaReativaExcedFPonta: { type: Type.STRING, description: "Valor em R$ da Energia Reativa Excedente - Fora Ponta." },
    energiaAtvInjetadaGDIOUC: { type: Type.STRING, description: "Energia Ativa Injetada GDI oUC (kWh)." },
    valorEnergiaAtvInjetadaGDIOUC: { type: Type.STRING, description: "Valor em R$ da Energia Ativa Injetada GDI oUC." },
    energiaAtvInjetadaGDIMUC: { type: Type.STRING, description: "Energia Ativa Injetada GDI mUC (kWh)." },
    valorEnergiaAtvInjetadaGDIMUC: { type: Type.STRING, description: "Valor em R$ da Energia Ativa Injetada GDI mUC." },
    cip: { type: Type.STRING, description: "Valor em R$ da Contribuição de Iluminação Pública (CIP ou COSIP)." },
    outrosEncargos: { type: Type.STRING, description: "Soma de outros encargos, multas, juros ou adicionais de bandeira tarifária." },
    pis: { type: Type.STRING, description: "Valor em R$ do PIS." },
    cofins: { type: Type.STRING, description: "Valor em R$ do COFINS." },
    icms: { type: Type.STRING, description: "Valor em R$ do ICMS." },
    concessionaria: { type: Type.STRING, description: "Nome da empresa concessionária (ex: ENERGISA, ELEKTRO, CPFL)." },
    numeroNotaFiscal: { type: Type.STRING, description: "Número da Nota Fiscal ou Número da Fatura." },
    dataVencimento: { type: Type.STRING, description: "Data de vencimento da fatura (ex: 15/08/2025)." },
    modalidadeTarifaria: { type: Type.STRING, description: "Modalidade Tarifária (ex: AZUL, VERDE, BRANCA, CONVENCIONAL)." },
    subgrupo: { type: Type.STRING, description: "Subgrupo tarifário (ex: A4, B1, B3)." }
  },
  required: ["uc", "anoLeitura", "mesReferencia"],
};

const AGRUPADORA_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    concessionaria: { type: Type.STRING, description: "Nome da concessionária (ex: ELEKTRO, ENERGISA)" },
    valorTotal: { type: Type.STRING, description: "Valor total da fatura (R$)" },
    mesReferencia: { type: Type.STRING, description: "Mês/Ano de referência (ex: Fevereiro/2026)" },
    vencimento: { type: Type.STRING, description: "Data de vencimento" },
    numeroNotaFiscal: { type: Type.STRING, description: "Número da Nota Fiscal ou Fatura (ex: AGP-01...)" },
    pis: { type: Type.STRING, description: "Valor do PIS (R$)" },
    cofins: { type: Type.STRING, description: "Valor do COFINS (R$)" },
    icms: { type: Type.STRING, description: "Valor do ICMS (R$)" },
    cip: { type: Type.STRING, description: "Valor da CIP (R$)" },
  },
  required: ["valorTotal", "mesReferencia"],
};

interface AgrupadoraData {
  concessionaria: string;
  valorTotal: number;
  mesReferencia: string;
  vencimento: string;
  numeroNotaFiscal: string;
  pis: number;
  cofins: number;
  icms: number;
  cip: number;
  fileName: string;
}

// --- Helper Functions ---

const Logo = ({ className = "h-10", showText = true, isLogin = false }: { className?: string, showText?: boolean, isLogin?: boolean }) => {
  const [error, setError] = useState(false);
  
  if (!error) {
    return (
      <img 
        src="/logo.png" 
        alt="Sanesul Energy" 
        className={`object-contain ${className}`}
        onError={() => setError(true)}
      />
    );
  }

  // Fallback to HTML logo if image is missing
  return (
    <div className={`flex ${isLogin ? 'flex-col items-center' : 'items-center gap-3'} ${className}`}>
      <div className={`${isLogin ? 'w-16 h-16 rounded-2xl mb-4' : className.includes('h-12') ? 'w-12 h-12 rounded-xl' : 'w-10 h-10 rounded-xl'} bg-sanesul-primary flex items-center justify-center shadow-lg shadow-sanesul-primary/20 shrink-0`}>
        <Zap className="text-white" size={isLogin ? 32 : className.includes('h-12') ? 24 : 20} />
      </div>
      {showText && (
        <div className={isLogin ? 'text-center' : ''}>
          <h1 className={`${isLogin ? 'text-3xl' : className.includes('h-12') ? 'text-3xl md:text-4xl' : 'text-2xl'} font-display font-bold tracking-tight text-sanesul-primary leading-none`}>
            Sanesul <span className="text-sanesul-secondary">Energy</span>
          </h1>
          <p className="text-[9px] uppercase tracking-[0.2em] text-sanesul-muted font-bold mt-1">
            {isLogin ? 'Acesso Restrito' : 'Portal de Inteligência Energética'}
          </p>
        </div>
      )}
    </div>
  );
};

const deduplicateBills = (bills: BillData[]) => {
  const seenIds = new Set();
  const uniqueBills = [];
  const seenKeys = new Set();

  // Iterate backwards to keep the latest processed bill for a given UC + Mes + Ano
  for (let i = bills.length - 1; i >= 0; i--) {
    const bill = bills[i];
    if (!bill.id || seenIds.has(bill.id)) continue;
    
    // Deduplicate completed bills based on content
    if (bill.status === 'completed' && bill.uc && bill.mesReferencia && bill.anoLeitura) {
      const key = `${bill.uc}-${bill.mesReferencia}-${bill.anoLeitura}`;
      if (seenKeys.has(key)) {
        continue; // Skip older duplicate
      }
      seenKeys.add(key);
    }
    
    seenIds.add(bill.id);
    uniqueBills.unshift(bill); // Add to front to maintain original order
  }
  
  return uniqueBills;
};

const ensureApiKey = async () => {
  if (window.aistudio && typeof window.aistudio.hasSelectedApiKey === 'function') {
    try {
      const hasKey = await window.aistudio.hasSelectedApiKey();
      if (!hasKey) {
        await window.aistudio.openSelectKey();
      }
    } catch (e) {
      console.warn("Erro ao verificar/abrir diálogo de chave de API:", e);
    }
  }
};

const generateContentWithRetry = async (
  ai: GoogleGenAI,
  params: any,
  retries = 5,
  delay = 5000
): Promise<GenerateContentResponse> => {
  try {
    // Add a timeout of 120 seconds to the API call
    let timeoutId: NodeJS.Timeout;
    const timeoutPromise = new Promise((_, reject) => {
      timeoutId = setTimeout(() => reject(new Error('TIMEOUT_API: A API demorou muito para responder (120s).')), 120000);
    });

    try {
      const response = await Promise.race([
        ai.models.generateContent(params),
        timeoutPromise
      ]) as GenerateContentResponse;
      
      clearTimeout(timeoutId!);
      return response;
    } catch (error) {
      clearTimeout(timeoutId!);
      throw error;
    }
  } catch (error: any) {
    // Extract error details
    let errorStr = '';
    let errorCode = 0;
    let errorStatus = '';
    
    if (typeof error === 'string') {
      errorStr = error;
    } else if (error && typeof error === 'object') {
      // Handle the specific format provided by the user
      const nestedError = error.error || error;
      errorCode = nestedError.code || error.status || 0;
      errorStatus = nestedError.status || '';
      errorStr = nestedError.message || error.message || JSON.stringify(error);
    }

    console.error('generateContentWithRetry Error:', {
      params: JSON.stringify(params),
      errorCode,
      errorStatus,
      errorStr,
      retries
    });

    const isTransientError = 
      errorCode === 429 || 
      errorCode === 500 || 
      errorCode === 502 || 
      errorCode === 503 || 
      errorCode === 504 ||
      errorStatus === 'RESOURCE_EXHAUSTED' ||
      errorStatus === 'INTERNAL' ||
      errorStatus === 'UNAVAILABLE' ||
      errorStr.includes('429') || 
      errorStr.includes('500') ||
      errorStr.includes('502') ||
      errorStr.includes('503') ||
      errorStr.includes('504') ||
      errorStr.includes('RESOURCE_EXHAUSTED') ||
      errorStr.includes('INTERNAL') ||
      errorStr.includes('UNAVAILABLE');

    const isTimeout = errorStr.includes('TIMEOUT_API');
    const isLockError = errorStr.includes('Lock broken by another request');
    const isHardQuota = errorStr.includes('spending cap') || errorStr.includes('monthly limit');
    const isRateLimit = errorCode === 429 || errorStatus === 'RESOURCE_EXHAUSTED' || errorStr.includes('429') || errorStr.includes('RESOURCE_EXHAUSTED');
    const isExpired = errorStr.includes('API key expired') || errorStr.includes('API_KEY_INVALID') || errorStr.includes('expired');
    const isInvalid = errorStr.includes('invalid API key') || errorStr.includes('invalid key') || (errorCode === 401 && errorStr.includes('invalid'));
    const isNotFound = errorStr.includes('Requested entity was not found') || errorStr.includes('API key not found');

    if (isNotFound || isExpired || isInvalid) {
      const msg = isNotFound 
        ? 'Chave de API não encontrada ou inválida. Por favor, use o botão "Trocar Conta" para selecionar uma chave válida.'
        : 'A chave da API expirou ou é inválida. Por favor, use o botão "Trocar Conta" para selecionar uma nova chave.';
      throw new Error(msg);
    }

    if (retries > 0 && (isTransientError || isTimeout || isLockError || isRateLimit) && !isHardQuota) {
      console.warn(`${isTimeout ? 'Timeout' : isLockError ? 'Lock error' : isRateLimit ? 'Rate limit' : 'Transient error (' + errorCode + ')'} hit, retrying in ${delay}ms... (${retries} retries left)`);
      await new Promise(resolve => setTimeout(resolve, delay));
      return generateContentWithRetry(ai, params, retries - 1, delay * 2);
    }
    
    // If it's a quota error or we're out of retries, throw
    if (isHardQuota || isRateLimit || isTransientError) {
      if (isHardQuota || isRateLimit) {
        const msg = isHardQuota 
          ? "O limite de gastos do seu projeto foi atingido. Verifique sua conta do Google Cloud (https://ai.google.dev/gemini-api/docs/billing)."
          : "Cota da API excedida ou limite de taxa atingido. Verifique seu plano e detalhes de faturamento no Google AI Studio (https://ai.google.dev/gemini-api/docs/billing). " + errorStr;
        const quotaError = new Error(msg);
        (quotaError as any).isQuotaError = true;
        throw quotaError;
      }
    }
    
    throw error;
  }
};

const parseValue = (val: string | number) => {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  
  let str = val.toString().trim();
  
  const lastDot = str.lastIndexOf('.');
  const lastComma = str.lastIndexOf(',');
  
  if (lastComma > lastDot) {
    // Brazilian format: 1.234,56
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (lastDot > lastComma && lastComma !== -1) {
    // US format: 1,234.56
    str = str.replace(/,/g, '');
  } else if (lastComma !== -1) {
    // Only comma: 1234,56
    str = str.replace(',', '.');
  }
  // If only dot, it's already in US format: 1234.56
  
  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
};

const formatNumber = (val: number, isCurrency: boolean = false, precision: number = 2) => {
  return val.toLocaleString('pt-BR', {
    minimumFractionDigits: isCurrency ? precision : 0,
    maximumFractionDigits: precision
  });
};

const formatMonth = (month: string | number) => {
  if (!month) return '';
  let normalized = month.toString().toLowerCase().trim();
  
  // Se vier no formato MM/YYYY, pega apenas o mês
  if (normalized.includes('/')) {
    normalized = normalized.split('/')[0];
  }

  const monthMap: Record<string, string> = {
    '01': 'Janeiro', '1': 'Janeiro', 'janeiro': 'Janeiro', 'jan': 'Janeiro',
    '02': 'Fevereiro', '2': 'Fevereiro', 'fevereiro': 'Fevereiro', 'fev': 'Fevereiro',
    '03': 'Março', '3': 'Março', 'março': 'Março', 'marco': 'Março', 'mar': 'Março',
    '04': 'Abril', '4': 'Abril', 'abril': 'Abril', 'abr': 'Abril',
    '05': 'Maio', '5': 'Maio', 'maio': 'Maio', 'mai': 'Maio',
    '06': 'Junho', '6': 'Junho', 'junho': 'Junho', 'jun': 'Junho',
    '07': 'Julho', '7': 'Julho', 'julho': 'Julho', 'jul': 'Julho',
    '08': 'Agosto', '8': 'Agosto', 'agosto': 'Agosto', 'ago': 'Agosto',
    '09': 'Setembro', '9': 'Setembro', 'setembro': 'Setembro', 'set': 'Setembro',
    '10': 'Outubro', 'outubro': 'Outubro', 'out': 'Outubro',
    '11': 'Novembro', 'novembro': 'Novembro', 'nov': 'Novembro',
    '12': 'Dezembro', 'dezembro': 'Dezembro', 'dez': 'Dezembro'
  };
  return monthMap[normalized] || month.toString();
};

const getMonthNumber = (month: string) => {
  if (!month) return 0;
  const monthOrder: Record<string, number> = {
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12,
    '01': 1, '02': 2, '03': 3, '04': 4, '05': 5, '06': 6,
    '07': 7, '08': 8, '09': 9, '10': 10, '11': 11, '12': 12,
    '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9
  };
  return monthOrder[month.toLowerCase().trim()] || 0;
};

const formatReference = (ref: string) => {
  if (!ref) return '';
  if (ref.includes('/')) {
    const parts = ref.split('/');
    if (parts.length === 2) {
      return `${formatMonth(parts[0])}/${parts[1]}`;
    }
  }
  return formatMonth(ref);
};

const mapDbToBillData = (dbBill: any): BillData => {
  let mod = dbBill.modalidade_tarifaria || '';
  let tipo = dbBill.tipo || '';
  if (UCS_OPER.has(String(dbBill.uc))) {
    tipo = 'OPER';
  } else if (UCS_ADM.has(String(dbBill.uc))) {
    tipo = 'ADM';
  } else if (UCS_LIVRE_MERCADO_LIVRE.has(String(dbBill.uc))) {
    tipo = 'LIVRE';
    if (!mod.toUpperCase().includes('LIVRE')) {
      mod = mod ? `${mod} - LIVRE` : 'LIVRE';
    }
  }

  return {
  id: dbBill.id,
  fileName: dbBill.file_name,
  uc: dbBill.uc || '',
  demandaPontaKW: dbBill.demanda_ponta_kw || '',
  demandaForaPontaKW: dbBill.demanda_fora_ponta_kw || '',
  demandaPotenciaMedidaPonta: dbBill.demanda_potencia_medida_ponta || '',
  demandaPotenciaMedidaForaPonta: dbBill.demanda_potencia_medida_fora_ponta || '',
  anoLeitura: dbBill.ano_leitura || '',
  mesReferencia: dbBill.mes_referencia || '',
  consumoKwhPonta: dbBill.consumo_kwh_ponta || '',
  consumoKwhForaPonta: dbBill.consumo_kwh_fora_ponta || '',
  valorConsumoKwhPonta: dbBill.valor_consumo_kwh_ponta || '',
  valorConsumoKwhForaPonta: dbBill.valor_consumo_kwh_fora_ponta || '',
  valorTotal: dbBill.valor_total || '',
  cidade: dbBill.cidade || '',
  demandaPotenciaNaoConsumidaPonta: dbBill.demanda_potencia_nao_consumida_ponta || '',
  demandaPotenciaNaoConsumidaFPonta: dbBill.demanda_potencia_nao_consumida_f_ponta || '',
  demandaPotenciaAtivaUltrapPonta: dbBill.demanda_potencia_ativa_ultrap_ponta || '',
  demandaPotenciaAtivaUltrapFPonta: dbBill.demanda_potencia_ativa_ultrap_f_ponta || '',
  energiaReativaExcedPonta: dbBill.energia_reativa_exced_ponta || '',
  energiaReativaExcedFPonta: dbBill.energia_reativa_exced_f_ponta || '',
  energiaInjetadaKwh: dbBill.energia_injetada_kwh || '',
  energiaCompensadaKwh: dbBill.energia_compensada_kwh || '',
  valorDemandaPotenciaMedidaPonta: dbBill.valor_demanda_potencia_medida_ponta || '',
  valorDemandaPotenciaMedidaForaPonta: dbBill.valor_demanda_potencia_medida_fora_ponta || '',
  valorDemandaPotenciaNaoConsumidaPonta: dbBill.valor_demanda_potencia_nao_consumida_ponta || '',
  valorDemandaPotenciaNaoConsumidaFPonta: dbBill.valor_demanda_potencia_nao_consumida_f_ponta || '',
  valorDemandaPotenciaAtivaUltrapPonta: dbBill.valor_demanda_potencia_ativa_ultrap_ponta || '',
  valorDemandaPotenciaAtivaUltrapFPonta: dbBill.valor_demanda_potencia_ativa_ultrap_f_ponta || '',
  valorEnergiaReativaExcedPonta: dbBill.valor_energia_reativa_exced_ponta || '',
  valorEnergiaReativaExcedFPonta: dbBill.valor_energia_reativa_exced_f_ponta || '',
  energiaAtvInjetadaGDIOUC: dbBill.energia_atv_injetada_gdi_ouc || '',
  valorEnergiaAtvInjetadaGDIOUC: dbBill.valor_energia_atv_injetada_gdi_ouc || '',
  energiaAtvInjetadaGDIMUC: dbBill.energia_atv_injetada_gdi_muc || '',
  valorEnergiaAtvInjetadaGDIMUC: dbBill.valor_energia_atv_injetada_gdi_muc || '',
  cip: dbBill.cip || '',
  outrosEncargos: dbBill.outros_encargos || '',
  pis: dbBill.pis || '',
  cofins: dbBill.cofins || '',
  icms: dbBill.icms || '',
  concessionaria: dbBill.concessionaria || '',
  numeroNotaFiscal: dbBill.numero_nota_fiscal || '',
  modalidadeTarifaria: mod,
  subgrupo: dbBill.subgrupo || '',
  tipo: tipo,
  mercado: UCS_LIVRE_MERCADO_LIVRE.has(String(dbBill.uc)) ? 'LIVRE' : 'CATIVO',
  dataVencimento: dbBill.data_vencimento || '',
  status: dbBill.status as any,
  error: dbBill.error || undefined,
  createdAt: dbBill.created_at ? new Date(dbBill.created_at).getTime() : Date.now()
  };
};

const mapBillDataToDb = (bill: BillData, userId: string) => ({
  file_name: bill.fileName,
  uc: bill.uc,
  demanda_ponta_kw: bill.demandaPontaKW,
  demanda_fora_ponta_kw: bill.demandaForaPontaKW,
  demanda_potencia_medida_ponta: bill.demandaPotenciaMedidaPonta,
  demanda_potencia_medida_fora_ponta: bill.demandaPotenciaMedidaForaPonta,
  ano_leitura: bill.anoLeitura,
  mes_referencia: bill.mesReferencia,
  consumo_kwh_ponta: bill.consumoKwhPonta,
  consumo_kwh_fora_ponta: bill.consumoKwhForaPonta,
  valor_consumo_kwh_ponta: bill.valorConsumoKwhPonta,
  valor_consumo_kwh_fora_ponta: bill.valorConsumoKwhForaPonta,
  valor_total: bill.valorTotal,
  cidade: bill.cidade,
  demanda_potencia_nao_consumida_ponta: bill.demandaPotenciaNaoConsumidaPonta,
  demanda_potencia_nao_consumida_f_ponta: bill.demandaPotenciaNaoConsumidaFPonta,
  demanda_potencia_ativa_ultrap_ponta: bill.demandaPotenciaAtivaUltrapPonta,
  demanda_potencia_ativa_ultrap_f_ponta: bill.demandaPotenciaAtivaUltrapFPonta,
  energia_reativa_exced_ponta: bill.energiaReativaExcedPonta,
  energia_reativa_exced_f_ponta: bill.energiaReativaExcedFPonta,
  energia_injetada_kwh: bill.energiaInjetadaKwh,
  energia_compensada_kwh: bill.energiaCompensadaKwh,
  valor_demanda_potencia_medida_ponta: bill.valorDemandaPotenciaMedidaPonta,
  valor_demanda_potencia_medida_fora_ponta: bill.valorDemandaPotenciaMedidaForaPonta,
  valor_demanda_potencia_nao_consumida_ponta: bill.valorDemandaPotenciaNaoConsumidaPonta,
  valor_demanda_potencia_nao_consumida_f_ponta: bill.valorDemandaPotenciaNaoConsumidaFPonta,
  valor_demanda_potencia_ativa_ultrap_ponta: bill.valorDemandaPotenciaAtivaUltrapPonta,
  valor_demanda_potencia_ativa_ultrap_f_ponta: bill.valorDemandaPotenciaAtivaUltrapFPonta,
  valor_energia_reativa_exced_ponta: bill.valorEnergiaReativaExcedPonta,
  valor_energia_reativa_exced_f_ponta: bill.valorEnergiaReativaExcedFPonta,
  energia_atv_injetada_gdi_ouc: bill.energiaAtvInjetadaGDIOUC,
  valor_energia_atv_injetada_gdi_ouc: bill.valorEnergiaAtvInjetadaGDIOUC,
  energia_atv_injetada_gdi_muc: bill.energiaAtvInjetadaGDIMUC,
  valor_energia_atv_injetada_gdi_muc: bill.valorEnergiaAtvInjetadaGDIMUC,
  cip: bill.cip,
  outros_encargos: bill.outrosEncargos,
  pis: bill.pis || '',
  cofins: bill.cofins || '',
  icms: bill.icms || '',
  concessionaria: bill.concessionaria || '',
  numero_nota_fiscal: bill.numeroNotaFiscal || '',
  modalidade_tarifaria: bill.modalidadeTarifaria || '',
  subgrupo: bill.subgrupo || '',
  tipo: bill.tipo || '',
  mercado: bill.mercado || '',
  data_vencimento: bill.dataVencimento || '',
  status: bill.status,
  error: bill.error || null,
  user_id: userId
});

// --- Components ---

const MetricCard = ({ title, custo, consumo, isReference = false, rightElement, titleColorClass = "text-sanesul-primary" }: { title: React.ReactNode, custo: number, consumo: number, isReference?: boolean, rightElement?: React.ReactNode, titleColorClass?: string }) => {
  const tarifaMedia = consumo > 0 ? (custo / consumo) : 0;
  const tarifaLabel = 'Tarifa Média (R$/kWh)';
  const custoLabel = isReference ? 'CUSTO (R$)' : 'Custo (R$)';
  const consumoLabel = isReference ? 'CONSUMO (kWh)' : 'Consumo (kWh)';

  return (
    <div className="bg-white p-6 rounded-[24px] border border-sanesul-primary/10 shadow-lg hover:shadow-xl transition-all hover:border-sanesul-primary/30">
      <div className="flex justify-between items-center mb-4">
        <h3 className={`text-sm font-display font-bold uppercase tracking-wider ${titleColorClass}`}>{title}</h3>
        {rightElement}
      </div>
      <div className="space-y-3">
        <div className="flex justify-between items-end border-b border-slate-100 pb-2">
          <span className="text-[10px] font-bold text-sanesul-muted uppercase tracking-wider">{custoLabel}</span>
          <span className="text-lg font-bold text-slate-800">R$ {custo.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
        </div>
        <div className="flex justify-between items-end border-b border-slate-100 pb-2">
          <span className="text-[10px] font-bold text-sanesul-muted uppercase tracking-wider">{consumoLabel}</span>
          <span className="text-lg font-bold text-slate-800">{consumo.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}</span>
        </div>
        <div className="flex justify-between items-end pt-1">
          <span className="text-[10px] font-bold text-sanesul-secondary uppercase tracking-wider">{tarifaLabel}</span>
          <span className="text-xl font-bold text-sanesul-secondary">R$ {tarifaMedia.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
        </div>
      </div>
    </div>
  );
};

const UCS_PPP = new Set([
  "12018", "33857", "34548", "34594", "34724", "36245", "36246", "74255", "74256", "75226", "94009", "101432", "101434", "101435", "102438", "108130", "108132", "112154", "112156", "112157", "112158", "112160", "112581", "121252", "121253", "121254", "121600", "128389", "128392", "128394", "128898", "132652", "132666", "132671", "138551", "138555", "138557", "152427", "152434", "152446", "153011", "158690", "163517", "163521", "163944", "172004", "172005", "172705", "179857", "179858", "179859", "189729", "189730", "189731", "189732", "189790", "197712", "206446", "209223", "209225", "209226", "209227", "209397", "211934", "211935", "211936", "211937", "216228", "216229", "220535", "220546", "222648", "222650", "223697", "231303", "231304", "231305", "231306", "233304", "233306", "238383", "244193", "244194", "244199", "244200", "244205", "244206", "244209", "247298", "247299", "249353", "249354", "252248", "252249", "252250", "252251", "253245", "253246", "256200", "258183", "264010", "264986", "264988", "269110", "269118", "270079", "274331", "464406", "1047259", "1084731", "1113637", "1126680", "1151043", "2414930", "2420193", "2558334", "2657735", "2700471", "2716454", "2754088", "2765050", "2797860", "2858514", "2884288", "2954787", "2999073", "3047887", "3058557", "3070666", "3102869", "3141335", "3175195", "3181887", "3188966", "3206144", "3207043", "3214000", "3234876", "3235300", "3248099", "3275502", "3301943", "3302837", "3310090", "3313761", "3331889", "3341371", "3341373", "3341380", "3343169", "3348432", "3366558", "3367575", "3371198", "3375315", "3390948", "3409248", "3412949", "3414263", "3417002", "3418302", "3421139", "3426808", "3481691", "3498079", "272605", "272951", "273988", "273989", "274332", "276549", "277100", "277101", "279007", "280858", "280860", "281808", "281809", "283247", "283248", "283352", "283480", "453683", "453827", "456560", "456731", "456907", "457351", "457765", "457766", "457891", "458050", "458289", "458661", "460570", "460571", "461216", "461759", "462534", "462964", "463730", "463783", "463908", "464549", "464764", "464765", "465134", "465135", "465971", "466787", "467063", "467064", "467145", "482891", "518898", "527978", "533217", "905272", "925640", "938246", "973292", "978395", "984681", "988341", "996818", "1000652", "1034959", "1047248", "1089791", "1126687", "1127638", "1136937", "1142030", "1142916", "1144446", "1148016", "1204522", "1223492", "1273099", "1273146", "1292715", "1309765", "1320065", "1352920", "1361474", "1388271", "1467369", "1479890", "1491784", "1543691", "1548221", "1600326", "1650695", "1656911", "1673468", "1677710", "1686836", "1690088", "1698936", "1698960", "1699438", "1701676", "1702111", "1745575", "1745856", "1748386", "1821234", "1877305", "1879309", "1879837", "1899594", "1901647", "1916616", "1923325", "1924437", "1936660", "1975585", "2065093", "2093921", "2140053", "2188959", "2203819", "2233618", "2283427", "2337244", "2342909", "2392852", "2398903", "2480085", "2524079", "2527881", "2563141", "2613087", "2632342", "3001597", "3001613", "3005931", "3005999", "3011291", "3036982", "3422802", "3443659", "3495962", "2601732", "2601678"
]);

const UCS_USINA = new Set([
  "2400975", "1602335", "279006", "176817", "176812", "102690", "3211"
]);

const VisaoGeralDashboard = ({ data, setCurrentPage, handleLogout, hasApiKey, handleSelectKey }: { data: any[], setCurrentPage: (page: string) => void, handleLogout: () => void, hasApiKey: boolean, handleSelectKey: () => void }) => {
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [hoveredLivreType, setHoveredLivreType] = useState<'azul' | 'verde' | null>(null);

  const availableMonths = Array.from(new Set(data.map(d => d.name))).filter(Boolean).sort((a, b) => {
    const [mA, yA] = String(a).split('/');
    const [mB, yB] = String(b).split('/');
    if (yA !== yB) return Number(yB) - Number(yA);
    return getMonthNumber(mB) - getMonthNumber(mA);
  });

  const filteredData = selectedMonth === 'all' ? data : data.filter(d => d.name === selectedMonth);

  const calc = (filterFn: (d: any) => boolean) => {
    const filtered = filteredData.filter(filterFn);
    const custo = filtered.reduce((acc, curr) => acc + curr.valorTotal, 0);
    const consumo = filtered.reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0);
    return { custo, consumo, tarifa: consumo > 0 ? custo / consumo : 0 };
  };

  const isGrupoA = (d: any) => d.demandaContratadaPonta > 0 || d.demandaContratadaForaPonta > 0;
  const isGrupoB = (d: any) => !isGrupoA(d);
  const isLivre = (d: any) => d.modalidadeTarifaria.includes('LIVRE') || d.tipo === 'LIVRE';
  const isCativo = (d: any) => !isLivre(d);
  
  const isAzul = (d: any) => d.modalidadeTarifaria.includes('AZUL');
  const isVerde = (d: any) => d.modalidadeTarifaria.includes('VERDE');
  const isOutrosGrupoA = (d: any) => isGrupoA(d) && !isAzul(d) && !isVerde(d);
  
  const hasCompensacao = (d: any) => d.solarInjetadaOUC > 0 || d.solarInjetadaMUC > 0;
  
  const isConsumoMinimo = (d: any) => isGrupoB(d) && (d.consumoPonta + d.consumoForaPonta) <= 100 && d.valorTotal < 150;
  const isPPP = (d: any) => isGrupoB(d) && UCS_PPP.has(String(d.uc));
  const isUsina = (d: any) => isGrupoB(d) && UCS_USINA.has(String(d.uc));
  const isOptanteB = (d: any) => isGrupoB(d) && (d.modalidadeTarifaria || '').toUpperCase().includes('OPTANTE');
  const isGeral = (d: any) => isGrupoB(d) && !isConsumoMinimo(d) && !isPPP(d) && !isUsina(d) && !isOptanteB(d);

  const totalGeral = calc(() => true);
  const grupoA = calc(isGrupoA);
  const grupoB = calc(isGrupoB);

  const livre = calc(d => isGrupoA(d) && isLivre(d));
  const livreAzul = calc(d => isGrupoA(d) && isLivre(d) && isAzul(d));
  const livreVerde = calc(d => isGrupoA(d) && isLivre(d) && isVerde(d));

  const cativo = calc(d => isGrupoA(d) && isCativo(d));
  const cativoAzul = calc(d => isGrupoA(d) && isCativo(d) && isAzul(d));
  const cativoVerde = calc(d => isGrupoA(d) && isCativo(d) && isVerde(d));
  const cativoOutras = calc(d => isGrupoA(d) && isCativo(d) && isOutrosGrupoA(d));

  const semCompensacao = calc(d => isGrupoB(d) && !hasCompensacao(d));
  const geral = calc(isGeral);
  const consumosMinimos = calc(isConsumoMinimo);
  const optanteB = calc(isOptanteB);

  const comCompensacao = calc(d => isGrupoB(d) && hasCompensacao(d));
  const ppp = calc(isPPP);
  const usinas = calc(isUsina);

  const totalSolarInjetada = filteredData.reduce((acc, curr) => acc + (curr.solarInjetadaOUC || 0) + (curr.solarInjetadaMUC || 0), 0);
  const emissoesEvitadas = totalSolarInjetada * 0.0426; // Fator médio do SIN

  const getMonthlyData = (sourceData: any[], filterFn?: (d: any) => boolean) => {
    interface GroupedItem {
      name: string;
      month: number;
      year: number;
      consumo: number;
      custo: number;
    }

    const filtered = filterFn ? sourceData.filter(filterFn) : sourceData;

    const grouped = filtered.reduce((acc, curr) => {
      const name = curr.name; // e.g. "Janeiro/2026"
      if (!acc[name]) {
        const [month, year] = name.split('/');
        acc[name] = { 
          name, 
          month: getMonthNumber(month), 
          year: parseInt(year), 
          consumo: 0, 
          custo: 0 
        };
      }
      acc[name].consumo += curr.consumoPonta + curr.consumoForaPonta;
      acc[name].custo += curr.valorTotal;
      return acc;
    }, {} as Record<string, GroupedItem>);

    const sorted = (Object.values(grouped) as GroupedItem[]).sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      return a.month - b.month;
    });

    const ptAbbr: Record<string, string> = {
      'Janeiro': 'Jan', 'Fevereiro': 'Fev', 'Março': 'Mar', 'Abril': 'Abr',
      'Maio': 'Mai', 'Junho': 'Jun', 'Julho': 'Jul', 'Agosto': 'Ago',
      'Setembro': 'Set', 'Outubro': 'Out', 'Novembro': 'Nov', 'Dezembro': 'Dez'
    };

    return sorted.map(item => {
      const [month] = item.name.split('/');
      const fullMonth = formatMonth(month);
      const abbr = ptAbbr[fullMonth] || month.substring(0, 3);
      return {
        name: abbr,
        consumo: item.consumo,
        custo: item.custo
      };
    });
  };

  const monthlyData = useMemo(() => getMonthlyData(data), [data]);

  const chartDomain = useMemo(() => {
    const maxConsumo = Math.max(...monthlyData.map(d => d.consumo), 0);
    const maxCusto = Math.max(...monthlyData.map(d => d.custo), 0);
    const maxVal = Math.max(maxConsumo, maxCusto);
    return [0, Math.ceil(maxVal * 1.1)];
  }, [monthlyData]);

  const monthlyDataAzul = useMemo(() => getMonthlyData(data, d => isGrupoA(d) && isLivre(d) && isAzul(d)), [data]);
  const monthlyDataVerde = useMemo(() => getMonthlyData(data, d => isGrupoA(d) && isLivre(d) && isVerde(d)), [data]);

  const sparklineDataAzul = monthlyDataAzul.map(m => ({ value: m.custo }));
  const sparklineDataVerde = monthlyDataVerde.map(m => ({ value: m.custo }));

  const MetricRow = ({ icon: Icon, label, value, unit, isCurrency }: { icon: any, label: string, value: number, unit?: string, isCurrency?: boolean }) => (
    <div className="flex items-center justify-between border-b border-slate-100 pb-1 mb-2 group/row">
      <div className="flex items-center gap-2">
        <Icon size={14} className="text-slate-400 group-hover/row:text-blue-500 transition-colors" />
        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">{label}</span>
      </div>
      <div className="text-right flex items-baseline gap-1">
        {isCurrency && <span className="text-[10px] text-slate-400 font-bold">R$</span>}
        <span className="text-lg font-black text-slate-900 tracking-tight">
          {formatNumber(value, isCurrency, isCurrency ? 2 : 0)}
        </span>
        {unit && <span className="text-[10px] text-slate-400 font-bold ml-1">{unit}</span>}
      </div>
    </div>
  );

  const SummaryCard = ({ title, data, icon: Icon, color = "blue", className = "" }: { title: string, data: any, icon: any, color?: "blue" | "indigo" | "sky", className?: string }) => {
    const colorStyles = {
      blue: { icon: "text-blue-600", bg: "bg-blue-50", border: "border-blue-100" },
      indigo: { icon: "text-indigo-600", bg: "bg-indigo-50", border: "border-indigo-100" },
      sky: { icon: "text-sky-600", bg: "bg-sky-50", border: "border-sky-100" },
    }[color];

    return (
      <div className={`rounded-[2rem] p-5 relative overflow-hidden bg-white text-slate-900 border border-slate-200 group transition-all duration-500 hover:shadow-xl shadow-sm ${className}`}>
        <div className="absolute top-0 right-0 w-64 h-64 bg-slate-50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-30 group-hover:scale-110 transition-transform duration-700 blur-3xl"></div>
        
        <div className="flex justify-between items-center mb-4 relative z-10">
          <div className="flex items-center gap-4">
            <div className={`p-2.5 ${colorStyles.bg} rounded-xl border ${colorStyles.border} shadow-sm group-hover:bg-white transition-colors duration-300`}>
              <Icon size={20} className={colorStyles.icon} />
            </div>
            <h3 className="text-xs font-black uppercase tracking-[0.2em] text-slate-800">{title}</h3>
          </div>
        </div>
        
        <div className="space-y-1 relative z-10">
          <MetricRow icon={DollarSign} label="Custo Total" value={data.custo} isCurrency />
          <MetricRow icon={Zap} label="Consumo Total" value={data.consumo} unit="kWh" />
          <MetricRow icon={Calculator} label="Tarifa Média" value={data.tarifa} isCurrency />
        </div>
      </div>
    );
  };

  const DetailCard = ({ title, data, color = "blue", icon: Icon }: { title: string, data: any, color?: "blue" | "green" | "slate" | "indigo", icon?: any }) => {
    const colorStyles = {
      blue: { text: "text-blue-600", bg: "bg-blue-50", border: "border-blue-100", iconBg: "bg-blue-100", iconText: "text-blue-600", hover: "hover:border-blue-200 hover:bg-blue-100/50", valueText: "text-slate-900" },
      green: { text: "text-emerald-600", bg: "bg-emerald-50", border: "border-emerald-100", iconBg: "bg-emerald-100", iconText: "text-emerald-600", hover: "hover:border-emerald-200 hover:bg-emerald-100/50", valueText: "text-slate-900" },
      slate: { text: "text-slate-600", bg: "bg-slate-50", border: "border-slate-200", iconBg: "bg-slate-100", iconText: "text-slate-600", hover: "hover:border-slate-300 hover:bg-slate-100/50", valueText: "text-slate-900" },
      indigo: { text: "text-indigo-600", bg: "bg-indigo-50", border: "border-indigo-100", iconBg: "bg-indigo-100", iconText: "text-indigo-600", hover: "hover:border-indigo-200 hover:bg-indigo-100/50", valueText: "text-slate-900" },
    }[color];
    
    return (
      <div className={`rounded-2xl p-3.5 border ${colorStyles.border} ${colorStyles.bg} flex-1 transition-all duration-300 ${colorStyles.hover} group relative overflow-hidden shadow-sm`}>
        <div className="absolute top-0 right-0 w-32 h-32 bg-white/50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-40 group-hover:scale-110 transition-transform duration-500 blur-xl"></div>
        <div className="flex items-center gap-3 mb-3 relative z-10">
          {Icon && (
            <div className={`p-2 rounded-xl ${colorStyles.iconBg} ${colorStyles.iconText} group-hover:scale-110 transition-transform shadow-sm border border-white`}>
              <Icon size={16} />
            </div>
          )}
          <h4 className={`text-xs font-bold uppercase tracking-wider ${colorStyles.text}`}>{title}</h4>
        </div>
        <div className="space-y-2 relative z-10">
          <div className="flex justify-between items-end bg-white/50 p-1.5 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Custo</span>
            <span className={`text-sm font-bold ${colorStyles.valueText}`}>R$ {formatNumber(data.custo, true)}</span>
          </div>
          <div className="flex justify-between items-end bg-white/50 p-1.5 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Consumo</span>
            <span className={`text-sm font-bold ${colorStyles.valueText}`}>{formatNumber(data.consumo, false, 0)} <span className="text-[10px] text-slate-400 font-medium">kWh</span></span>
          </div>
          <div className="flex justify-between items-end pt-2 border-t border-slate-100">
            <span className={`text-[10px] font-bold uppercase tracking-wider ${colorStyles.text}`}>Tarifa Média</span>
            <span className={`text-sm font-bold ${colorStyles.text}`}>R$ {formatNumber(data.tarifa, true)}</span>
          </div>
        </div>
      </div>
    );
  };

  const SparklineCard = ({ title, data, color = "blue", sparklineData, fullMonthlyData, onMouseEnter, onMouseLeave }: { title: string, data: any, color?: "blue" | "green", sparklineData: any[], fullMonthlyData: any[], onMouseEnter?: () => void, onMouseLeave?: () => void }) => {
    const colorHex = color === "blue" ? "#3b82f6" : "#10b981";
    const bgClass = color === "blue" ? "bg-blue-50" : "bg-emerald-50";
    const textClass = color === "blue" ? "text-blue-600" : "text-emerald-600";
    const valueTextClass = "text-slate-900";
    const borderClass = color === "blue" ? "border-blue-100" : "border-emerald-100";
    const hoverBorderClass = color === "blue" ? "hover:border-blue-200 hover:bg-blue-100/50" : "hover:border-emerald-200 hover:bg-emerald-100/50";
    const iconBgClass = color === "blue" ? "bg-blue-100" : "bg-emerald-100";
    
    const maxConsumo = Math.max(...fullMonthlyData.map(d => d.consumo), 0);
    const maxCusto = Math.max(...fullMonthlyData.map(d => d.custo), 0);
    const maxVal = Math.max(maxConsumo, maxCusto);
    const domain = [0, Math.ceil(maxVal * 1.1)];

    return (
      <div 
        onMouseEnter={onMouseEnter}
        onMouseLeave={onMouseLeave}
        className={`rounded-2xl p-4 border ${borderClass} ${bgClass} flex items-center justify-between mt-3 transition-all duration-300 ${hoverBorderClass} group relative overflow-visible shadow-sm`}
      >
        <div className="absolute top-0 right-0 w-24 h-24 bg-white/50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-40 group-hover:scale-110 transition-transform duration-500 blur-xl"></div>
        <div className="flex items-center gap-4 relative z-10">
          <div className={`w-10 h-10 rounded-xl ${iconBgClass} flex items-center justify-center shadow-sm border border-white group-hover:scale-110 transition-transform`}>
            <TrendingUp size={20} className={textClass} />
          </div>
          <div>
            <h4 className={`text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1`}>{title}</h4>
            <p className={`text-lg font-bold ${valueTextClass}`}>R$ {formatNumber(data.custo, true)}</p>
          </div>
        </div>
        <div className="w-24 h-12 opacity-80 relative z-10 group-hover:opacity-10 transition-opacity">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={sparklineData}>
              <Line type="monotone" dataKey="value" stroke={colorHex} strokeWidth={2.5} dot={false} isAnimationActive={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* Hover Chart */}
        <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-4 w-[450px] bg-white rounded-3xl shadow-2xl border border-slate-200 p-6 opacity-0 invisible group-hover:opacity-100 group-hover:visible transition-all duration-300 z-50">
          <div className="flex justify-between items-center mb-6">
            <div className="flex items-center gap-3">
              <div className={`w-2 h-8 rounded-full ${color === 'blue' ? 'bg-blue-500' : 'bg-emerald-500'}`}></div>
              <div>
                <h4 className="text-sm font-bold text-slate-900 uppercase tracking-widest">{title}</h4>
                <p className="text-[10px] text-slate-500 font-medium mt-0.5">Evolução Mensal</p>
              </div>
            </div>
            <div className="flex items-center gap-4 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200">
              <div className="flex items-center gap-2">
                <div className="w-2 h-2 rounded-full bg-[#0ea5e9]"></div>
                <span className="text-[9px] font-bold text-slate-600 uppercase tracking-wider">Consumo</span>
              </div>
              <div className="flex items-center gap-2">
                <div className="w-2 h-2 rounded-full bg-[#6366f1]"></div>
                <span className="text-[9px] font-bold text-slate-600 uppercase tracking-wider">Custo</span>
              </div>
            </div>
          </div>
          <div className="h-[200px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={fullMonthlyData} margin={{ top: 10, right: 10, bottom: 0, left: 0 }}>
                <defs>
                  <linearGradient id={`colorConsumo-${color}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id={`colorCusto-${color}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#6366f1" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} horizontal={true} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }} dy={10} />
                <YAxis 
                  yAxisId="left" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }} 
                  dx={-10} 
                  tickFormatter={(val) => formatNumber(val, false, 0)}
                  domain={domain}
                />
                <YAxis 
                  yAxisId="right" 
                  orientation="right" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }} 
                  dx={10} 
                  tickFormatter={(val) => `R$ ${formatNumber(val, true, 0)}`}
                  domain={domain}
                />
                <Tooltip 
                  contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)', padding: '12px 16px' }}
                  itemStyle={{ fontSize: '12px', fontWeight: 600, padding: '4px 0' }}
                  labelStyle={{ fontSize: '11px', fontWeight: 700, color: '#64748b', marginBottom: '8px', textTransform: 'uppercase', letterSpacing: '0.05em' }}
                  formatter={(value: number, name: string) => [
                    name === 'custo' ? `R$ ${formatNumber(value, true)}` : `${formatNumber(value, false)} kWh`,
                    name === 'custo' ? 'Custo' : 'Consumo'
                  ]}
                />
                <Area yAxisId="left" type="monotone" dataKey="consumo" stroke="#0ea5e9" strokeWidth={3} fillOpacity={1} fill={`url(#colorConsumo-${color})`} activeDot={{ r: 6, strokeWidth: 0, fill: '#0ea5e9' }} />
                <Area yAxisId="right" type="monotone" dataKey="custo" stroke="#6366f1" strokeWidth={3} fillOpacity={1} fill={`url(#colorCusto-${color})`} activeDot={{ r: 6, strokeWidth: 0, fill: '#6366f1' }} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] pb-8 text-slate-600 selection:bg-blue-500/30">
      {/* Header */}
      <header className="flex items-center justify-between px-8 py-4 bg-white border-b border-slate-200">
        <Logo className="h-10" />
        <div className="flex items-center gap-4">
          <button
            onClick={() => setCurrentPage('sistema')}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white hover:bg-blue-500 transition-all rounded-xl text-xs font-bold tracking-wider shadow-md active:scale-95"
          >
            <LayoutDashboard size={16} />
            Acessar Sistema
          </button>
          <button
            onClick={handleSelectKey}
            className={`flex items-center gap-2 px-4 py-2 ${hasApiKey ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-amber-50 text-amber-700 border-amber-200'} border hover:opacity-80 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95`}
            title={hasApiKey ? "Trocar Chave de API" : "Selecionar Chave de API"}
          >
            <Key size={16} />
            {hasApiKey ? "Trocar Conta" : "Configurar API"}
          </button>
          <button
            onClick={handleLogout}
            className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
            title="Sair"
          >
            <LogOut size={16} />
            Sair
          </button>
        </div>
      </header>

      <div className="max-w-[1600px] mx-auto px-6 py-8 space-y-8">
        {/* Title Section */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-white p-6 rounded-3xl shadow-sm border border-slate-200">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-blue-50 rounded-2xl flex items-center justify-center border border-blue-100">
              <Activity className="text-blue-600" size={24} />
            </div>
            <div>
              <h2 className="text-2xl font-bold text-slate-900 tracking-tight">Análise de Insumo de Energia Elétrica</h2>
              <p className="text-sm text-slate-500 font-medium">Visão geral consolidada de custos e consumos por grupos tarifários.</p>
            </div>
          </div>
          <div className="flex items-center gap-3 bg-slate-50 px-5 py-3 rounded-2xl border border-slate-200 shadow-inner">
            <Calendar size={18} className="text-blue-600" />
            <select 
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(e.target.value)}
              className="bg-transparent text-sm font-bold text-slate-700 outline-none cursor-pointer appearance-none pr-4"
            >
              <option value="all">Período Completo</option>
              {availableMonths.map(month => (
                <option key={month} value={month}>{month}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Main Chart */}
        <div className="bg-white rounded-[2.5rem] p-6 shadow-sm border border-slate-200 relative overflow-hidden group hover:shadow-md transition-all duration-500">
          <div className="absolute top-0 right-0 w-[500px] h-[500px] bg-gradient-to-br from-blue-50 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700 blur-3xl"></div>
          
          <div className="flex justify-between items-center mb-6 relative z-10">
            <div className="flex items-center gap-5">
              <div className="w-3 h-10 bg-gradient-to-b from-blue-500 to-blue-700 rounded-full shadow-sm"></div>
              <div>
                <h3 className="text-lg font-bold text-slate-900 uppercase tracking-widest">Evolução de Consumo e Custo</h3>
                <p className="text-xs text-slate-500 font-medium mt-1">Histórico Mensal Consolidado</p>
              </div>
            </div>
            <div className="flex items-center gap-6 bg-slate-50 px-5 py-3 rounded-2xl border border-slate-200 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="w-3 h-3 rounded-full bg-[#0ea5e9]"></div>
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Consumo (kWh)</span>
              </div>
              <div className="flex items-center gap-3">
                <div className="w-3 h-3 rounded-full bg-[#6366f1]"></div>
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Custo (R$)</span>
              </div>
            </div>
          </div>
          <div className="h-[165px] w-full relative z-10">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={monthlyData} margin={{ top: 20, right: 80, bottom: 10, left: 80 }}>
                <defs>
                  <linearGradient id="colorConsumo" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id="colorCusto" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#6366f1" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} horizontal={true} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 11, fontWeight: 600 }} dy={10} padding={{ left: 10, right: 10 }} />
                <YAxis 
                  yAxisId="left" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748b', fontSize: 11, fontWeight: 600 }} 
                  dx={-10} 
                  tickFormatter={(val) => formatNumber(val, false, 0)}
                  domain={chartDomain}
                  label={{ value: 'Consumo (kWh)', angle: -90, position: 'insideLeft', offset: -55, style: { textAnchor: 'middle', fill: '#64748b', fontSize: 10, fontWeight: 700, textTransform: 'uppercase' } }}
                />
                <YAxis 
                  yAxisId="right" 
                  orientation="right" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748b', fontSize: 11, fontWeight: 600 }} 
                  dx={10} 
                  tickFormatter={(val) => formatNumber(val, false, 0)}
                  domain={chartDomain}
                  label={{ value: 'Custo (R$)', angle: 90, position: 'insideRight', offset: -55, style: { textAnchor: 'middle', fill: '#64748b', fontSize: 10, fontWeight: 700, textTransform: 'uppercase' } }}
                />
                <Tooltip 
                  contentStyle={{ borderRadius: '20px', border: '1px solid #e2e8f0', boxShadow: '0 10px 15px -3px rgba(0,0,0,0.1)', padding: '16px 20px', fontWeight: 'bold', backgroundColor: 'rgba(255, 255, 255, 0.95)', backdropFilter: 'blur(12px)' }}
                  formatter={(value: number, name: string) => [
                    name === 'consumo' ? formatNumber(value, false, 0) + ' kWh' : formatNumber(value, true),
                    name === 'consumo' ? 'Consumo' : 'Custo'
                  ]}
                  labelStyle={{ color: '#64748b', marginBottom: '8px', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.1em' }}
                  itemStyle={{ color: '#1e293b' }}
                />
                <Area 
                  yAxisId="left" 
                  type="monotone" 
                  dataKey="consumo" 
                  stroke="#0ea5e9" 
                  strokeWidth={4} 
                  fillOpacity={1} 
                  fill="url(#colorConsumo)" 
                  dot={{ r: 4, fill: '#fff', strokeWidth: 2, stroke: '#0ea5e9' }} 
                  activeDot={{ r: 6, strokeWidth: 0, fill: '#0ea5e9' }} 
                />
                <Area 
                  yAxisId="right" 
                  type="monotone" 
                  dataKey="custo" 
                  stroke="#6366f1" 
                  strokeWidth={4} 
                  fillOpacity={1} 
                  fill="url(#colorCusto)" 
                  dot={{ r: 4, fill: '#fff', strokeWidth: 2, stroke: '#6366f1' }} 
                  activeDot={{ r: 6, strokeWidth: 0, fill: '#6366f1' }} 
                />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Summary Cards */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
          <div className="md:col-span-2">
            <SummaryCard title="TOTAL GERAL" data={totalGeral} icon={DollarSign} color="indigo" />
          </div>
          <SummaryCard title="GRUPO A (MT/AT)" data={grupoA} icon={Zap} color="blue" />
          <SummaryCard title="GRUPO B (BT)" data={grupoB} icon={Battery} color="sky" />
        </div>

        {/* Detalhamento Section */}
        <div className="space-y-12">
          
          {/* Grupo A Section */}
          <div className="space-y-8">
            <div className="flex items-center justify-between px-2 bg-white/50 py-3 rounded-2xl border border-slate-100 backdrop-blur-sm">
              <h2 className="text-xl font-bold text-slate-800 flex items-center gap-4">
                <div className="w-12 h-12 rounded-2xl bg-gradient-to-br from-blue-600 to-blue-800 flex items-center justify-center shadow-lg shadow-blue-500/30 border border-blue-400/20">
                  <Zap size={24} className="text-white" />
                </div>
                Detalhamento Grupo A
              </h2>
              <span className="text-xs font-bold text-blue-600 bg-blue-50 px-3 py-1.5 rounded-lg uppercase tracking-wider border border-blue-100">Alta Tensão</span>
            </div>
            
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* LIVRE Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-blue-100 relative group hover:shadow-xl transition-all duration-500 min-h-[500px] flex flex-col">
                <div className="absolute inset-0 overflow-hidden rounded-[2rem] pointer-events-none">
                  <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-blue-50 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>
                </div>
                
                <AnimatePresence mode="wait">
                  {!hoveredLivreType ? (
                    <motion.div
                      key="content"
                      initial={{ opacity: 0 }}
                      animate={{ opacity: 1 }}
                      exit={{ opacity: 0 }}
                      className="flex-1 flex flex-col"
                    >
                      <div className="flex justify-between items-center mb-8 relative z-10">
                        <div className="flex items-center gap-5">
                          <div className="w-3 h-12 bg-gradient-to-b from-blue-500 to-blue-700 rounded-full shadow-md"></div>
                          <div>
                            <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">MERCADO LIVRE</h3>
                            <p className="text-xs text-slate-500 font-medium mt-1">Ambiente de Contratação Livre</p>
                          </div>
                        </div>
                        <div className="p-3 bg-blue-50 rounded-2xl border border-blue-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                          <Activity size={24} className="text-blue-600" />
                        </div>
                      </div>
                      
                      <div className="space-y-1 relative z-10 mb-8">
                        <MetricRow icon={DollarSign} label="Custo Total" value={livre.custo} isCurrency />
                        <MetricRow icon={Zap} label="Consumo Total" value={livre.consumo} unit="kWh" />
                        <MetricRow icon={Calculator} label="Tarifa Média" value={livre.tarifa} isCurrency />
                      </div>

                      <div className="space-y-6 relative z-10 mt-auto">
                        <div className="grid grid-cols-2 gap-6">
                          <DetailCard title="Faturas Azul" data={livreAzul} color="blue" icon={Zap} />
                          <DetailCard title="Faturas Verde" data={livreVerde} color="green" icon={Zap} />
                        </div>
                        <div className="grid grid-cols-2 gap-6">
                          <SparklineCard 
                            title="Evolução Azul" 
                            data={livreAzul} 
                            color="blue" 
                            sparklineData={sparklineDataAzul} 
                            fullMonthlyData={monthlyDataAzul} 
                            onMouseEnter={() => setHoveredLivreType('azul')}
                          />
                          <SparklineCard 
                            title="Evolução Verde" 
                            data={livreVerde} 
                            color="green" 
                            sparklineData={sparklineDataVerde} 
                            fullMonthlyData={monthlyDataVerde} 
                            onMouseEnter={() => setHoveredLivreType('verde')}
                          />
                        </div>
                      </div>
                    </motion.div>
                  ) : (
                    <motion.div
                      key="chart"
                      initial={{ opacity: 0, scale: 0.95 }}
                      animate={{ opacity: 1, scale: 1 }}
                      exit={{ opacity: 0, scale: 0.95 }}
                      className="absolute inset-0 p-8 flex flex-col z-20 bg-white rounded-[2rem]"
                      onMouseLeave={() => setHoveredLivreType(null)}
                    >
                      <div className="flex justify-between items-center mb-8">
                        <div className="flex items-center gap-4">
                          <div className={`w-3 h-10 rounded-full ${hoveredLivreType === 'azul' ? 'bg-blue-500' : 'bg-emerald-500'} shadow-lg`}></div>
                          <div>
                            <h3 className="text-xl font-black text-slate-900 uppercase tracking-tighter">
                              Evolução {hoveredLivreType === 'azul' ? 'Azul' : 'Verde'}
                            </h3>
                            <p className="text-xs text-slate-500 font-bold uppercase tracking-widest opacity-60">Análise Detalhada de Mercado Livre</p>
                          </div>
                        </div>
                        <button 
                          onClick={() => setHoveredLivreType(null)}
                          className="p-2 hover:bg-slate-100 rounded-full transition-colors text-slate-400"
                        >
                          <X size={20} />
                        </button>
                      </div>

                      <div className="flex-1 w-full min-h-[250px] bg-slate-50/50 rounded-3xl p-4 border border-slate-100">
                        <ResponsiveContainer width="100%" height="100%">
                          <AreaChart 
                            data={hoveredLivreType === 'azul' ? monthlyDataAzul : monthlyDataVerde} 
                            margin={{ top: 20, right: 30, bottom: 20, left: 10 }}
                          >
                            <defs>
                              <linearGradient id="colorConsumoLivre" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.3}/>
                                <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0}/>
                              </linearGradient>
                              <linearGradient id="colorCustoLivre" x1="0" y1="0" x2="0" y2="1">
                                <stop offset="5%" stopColor="#6366f1" stopOpacity={0.3}/>
                                <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                              </linearGradient>
                            </defs>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                            <XAxis 
                              dataKey="name" 
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fill: '#64748b', fontSize: 11, fontWeight: 700 }} 
                              dy={15}
                            />
                            <YAxis 
                              yAxisId="left"
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }}
                              tickFormatter={(val) => formatNumber(val, false, 0)}
                            />
                            <YAxis 
                              yAxisId="right"
                              orientation="right"
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fill: '#64748b', fontSize: 10, fontWeight: 600 }}
                              tickFormatter={(val) => `R$ ${formatNumber(val, true, 0)}`}
                            />
                            <Tooltip 
                              contentStyle={{ 
                                backgroundColor: 'rgba(255, 255, 255, 0.95)', 
                                borderRadius: '16px', 
                                border: 'none', 
                                boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1)' 
                              }}
                              itemStyle={{ fontSize: '11px', fontWeight: 'bold' }}
                            />
                            <Area 
                              yAxisId="left"
                              type="monotone" 
                              dataKey="consumo" 
                              name="Consumo (kWh)"
                              stroke="#0ea5e9" 
                              strokeWidth={3}
                              fillOpacity={1} 
                              fill="url(#colorConsumoLivre)" 
                            />
                            <Area 
                              yAxisId="right"
                              type="monotone" 
                              dataKey="custo" 
                              name="Custo (R$)"
                              stroke="#6366f1" 
                              strokeWidth={3}
                              fillOpacity={1} 
                              fill="url(#colorCustoLivre)" 
                            />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                      
                      <div className="mt-6 grid grid-cols-2 gap-4">
                        <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                          <p className="text-[10px] font-bold text-blue-600 uppercase tracking-widest mb-1">Consumo Médio</p>
                          <p className="text-xl font-black text-blue-900">
                            {formatNumber((hoveredLivreType === 'azul' ? livreAzul : livreVerde).consumo / (hoveredLivreType === 'azul' ? monthlyDataAzul : monthlyDataVerde).length || 0, false, 0)} <span className="text-xs font-bold opacity-60">kWh</span>
                          </p>
                        </div>
                        <div className="bg-indigo-50/50 p-4 rounded-2xl border border-indigo-100/50">
                          <p className="text-[10px] font-bold text-indigo-600 uppercase tracking-widest mb-1">Custo Médio</p>
                          <p className="text-xl font-black text-indigo-900">
                            <span className="text-xs font-bold opacity-60 mr-1">R$</span>
                            {formatNumber((hoveredLivreType === 'azul' ? livreAzul : livreVerde).custo / (hoveredLivreType === 'azul' ? monthlyDataAzul : monthlyDataVerde).length || 0, true, 2)}
                          </p>
                        </div>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>

              {/* Cativo Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative group hover:shadow-xl transition-all duration-500">
                <div className="absolute inset-0 overflow-hidden rounded-[2rem] pointer-events-none">
                  <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-slate-100 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>
                </div>
                
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-slate-400 to-slate-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">Consumidor Cativo</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Ambiente de Contratação Regulada</p>
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <ZapOff size={24} className="text-slate-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={cativo.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={cativo.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={cativo.tarifa} isCurrency />
                </div>

                <div className="grid grid-cols-2 gap-6 relative z-10">
                  <DetailCard title="Faturas Azul" data={cativoAzul} color="blue" icon={Zap} />
                  <DetailCard title="Faturas Verde" data={cativoVerde} color="green" icon={Zap} />
                </div>
              </div>
            </div>
          </div>

          {/* Grupo B Section */}
          <div className="space-y-8">
            <div className="flex items-center justify-between px-2 bg-white/50 py-3 rounded-2xl border border-slate-100 backdrop-blur-sm">
              <h2 className="text-xl font-bold text-slate-800 flex items-center gap-4">
                <div className="w-12 h-12 rounded-2xl bg-gradient-to-br from-sky-500 to-sky-700 flex items-center justify-center shadow-lg shadow-sky-500/30 border border-sky-400/20">
                  <Battery size={24} className="text-white" />
                </div>
                Detalhamento Grupo B
              </h2>
              <span className="text-xs font-bold text-sky-600 bg-sky-50 px-3 py-1.5 rounded-lg uppercase tracking-wider border border-sky-100">Baixa Tensão</span>
            </div>
            
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* Sem Compensação Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-slate-100 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>
                
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-slate-400 to-slate-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">UC's Sem Compensação</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Consumo Padrão da Rede</p>
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Activity size={24} className="text-slate-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={semCompensacao.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={semCompensacao.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={semCompensacao.tarifa} isCurrency />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10">
                  <DetailCard title="Consumo Geral" data={geral} color="slate" icon={Activity} />
                  <DetailCard title="Consumos Mínimos" data={consumosMinimos} color="indigo" icon={ArrowDown} />
                  <DetailCard title="Optante B" data={optanteB} color="blue" icon={Battery} />
                </div>
              </div>

              {/* Com Compensação Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-emerald-100 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>
                
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-emerald-400 to-emerald-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">UC's Com Compensação</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Geração Distribuída e Sustentável</p>
                    </div>
                  </div>
                  <div className="p-3 bg-emerald-50 rounded-2xl border border-emerald-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Zap size={24} className="text-emerald-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={comCompensacao.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={comCompensacao.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={comCompensacao.tarifa} isCurrency />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10 mb-6">
                  <DetailCard title="PPP Fotovoltaica" data={ppp} color="green" icon={Zap} />
                  <DetailCard title="Usinas Sanesul" data={usinas} color="blue" icon={Activity} />
                  
                  {/* Crédito de Carbono Card */}
                  <div className="bg-emerald-50/50 rounded-2xl p-5 border border-emerald-100 relative overflow-hidden group">
                    <div className="flex items-center gap-3 mb-4">
                      <div className="p-2 bg-emerald-100 rounded-lg text-emerald-600">
                        <Leaf size={20} />
                      </div>
                      <h4 className="text-sm font-bold text-emerald-800 uppercase tracking-wider">Crédito de Carbono</h4>
                    </div>
                    <div className="space-y-3">
                      <div className="flex justify-between items-center bg-white/60 p-2 rounded-lg border border-emerald-100/50">
                        <span className="text-[10px] font-bold text-emerald-600/70 uppercase tracking-wider">Mês/Ano</span>
                        <span className="text-xs font-bold text-emerald-900">{selectedMonth === 'all' ? 'Todos os Meses' : selectedMonth}</span>
                      </div>
                      <div className="flex justify-between items-center bg-white/60 p-2 rounded-lg border border-emerald-100/50">
                        <span className="text-[10px] font-bold text-emerald-600/70 uppercase tracking-wider">Geração Solar Abatida</span>
                        <span className="text-xs font-bold text-emerald-900">{formatNumber(totalSolarInjetada, false, 0)} <span className="text-[10px] text-emerald-600/70">kWh</span></span>
                      </div>
                      <div className="flex justify-between items-center bg-white/60 p-2 rounded-lg border border-emerald-100/50">
                        <span className="text-[10px] font-bold text-emerald-600/70 uppercase tracking-wider">Emissões Evitadas</span>
                        <span className="text-xs font-bold text-emerald-900">{formatNumber(emissoesEvitadas, false, 2)} <span className="text-[10px] text-emerald-600/70">KgCO₂</span></span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(() => {
    return localStorage.getItem('sanesul_auth') === 'true';
  });

  const [searchUC, setSearchUC] = useState('');

  React.useEffect(() => {
    if (!isSupabaseConfigured) {
      console.warn('Supabase não configurado. Ignorando verificação de sessão.');
      return;
    }

    supabase.auth.getSession().then(({ data: { session } }) => {
      setIsAuthenticated(!!session);
      if (session) {
        localStorage.setItem('sanesul_auth', 'true');
      } else {
        localStorage.removeItem('sanesul_auth');
      }
    }).catch(async err => {
      console.error('Erro ao buscar sessão do Supabase:', err);
      if (err.message?.includes('Refresh Token Not Found') || err.message?.includes('invalid refresh token')) {
        try {
          await supabase.auth.signOut();
        } catch (signOutErr) {
          console.error('Erro ao realizar signOut:', signOutErr);
        }
        localStorage.removeItem('sanesul_auth');
        setIsAuthenticated(false);
      }
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setIsAuthenticated(!!session);
      if (session) {
        localStorage.setItem('sanesul_auth', 'true');
      } else {
        localStorage.removeItem('sanesul_auth');
      }
    });

    return () => subscription.unsubscribe();
  }, []);
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [loginError, setLoginError] = useState('');

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError('');
    
    console.log('Tentando login com:', loginUsername);
    
    try {
      if (!isSupabaseConfigured) {
        setLoginError('O Supabase não está configurado. Por favor, adicione as chaves VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY nas configurações do AI Studio.');
        return;
      }

      const { data, error } = await supabase.auth.signInWithPassword({
        email: loginUsername,
        password: loginPassword,
      });

      if (error) {
        console.error('Erro no Supabase Auth:', error);
        setLoginError(`Erro de autenticação: ${error.message}`);
      } else if (data.user) {
        console.log('Login bem-sucedido:', data.user);
        setIsAuthenticated(true);
        localStorage.setItem('sanesul_auth', 'true');
      }
    } catch (err: any) {
      console.error('Erro inesperado no handleLogin:', err);
      if (err.message === 'Failed to fetch') {
        setLoginError('Erro de conexão com o Supabase. Verifique se a URL está correta nas configurações.');
      } else {
        setLoginError('Erro inesperado ao tentar logar. Verifique a conexão.');
      }
    }
  };

  const [hasApiKey, setHasApiKey] = useState<boolean>(false);

  useEffect(() => {
    const checkApiKey = async () => {
      const aiStudio = (window as any).aistudio;
      if (aiStudio?.hasSelectedApiKey) {
        const selected = await aiStudio.hasSelectedApiKey();
        setHasApiKey(selected);
      }
    };
    checkApiKey();
  }, []);

  const handleSelectKey = async () => {
    const aiStudio = (window as any).aistudio;
    if (aiStudio?.openSelectKey) {
      await aiStudio.openSelectKey();
      setHasApiKey(true);
    }
  };

  const handleLogout = async () => {
    if (isSupabaseConfigured) {
      await supabase.auth.signOut();
    }
    setIsAuthenticated(false);
    localStorage.removeItem('sanesul_auth');
  };

  const [bills, setBills] = useState<BillData[]>(() => {
    const saved = localStorage.getItem('sanesul_bills');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Deduplicate on load
        const deduplicated = deduplicateBills(Array.isArray(parsed) ? parsed : []);
        
        // One-time fix for Elektro UC from localStorage
        return deduplicated.map(b => {
          const isElektro = (b.concessionaria || '').toUpperCase().includes('ELEKTRO');
          if (isElektro && b.status === 'completed') {
            const fileNameNumbers = b.fileName.replace(/\.[^/.]+$/, "").replace(/\D/g, "");
            if (fileNameNumbers.length >= 5 && b.uc !== fileNameNumbers) {
              return { ...b, uc: fileNameNumbers };
            }
          }
          return b;
        });
      } catch (e) {
        return [];
      }
    }
    return [];
  });

  React.useEffect(() => {
    localforage.getItem<Record<string, File>>('sanesul_bills_files').then(filesMap => {
      if (filesMap) {
        setBills(prev => prev.map(b => {
          if (filesMap[b.id]) {
            return { ...b, file: filesMap[b.id] };
          }
          return b;
        }));
      }
    }).catch(err => {
      console.warn('Failed to load files from localforage:', err);
    });
  }, []);

  React.useEffect(() => {
    const fetchBills = async () => {
      if (!isSupabaseConfigured || !isAuthenticated) return;
      
      try {
        const { data: { user } } = await supabase.auth.getUser();
        if (!user) return;

        let allData: any[] = [];
        let from = 0;
        let to = 999;
        let finished = false;

        while (!finished) {
          const { data, error } = await supabase
            .from('bills')
            .select('*')
            .range(from, to)
            .order('created_at', { ascending: false });

          if (error) {
            console.error('Erro ao buscar faturas do Supabase:', error);
            finished = true;
            return;
          }

          if (data && data.length > 0) {
            allData = [...allData, ...data];
            if (data.length < 1000) {
              finished = true;
            } else {
              from += 1000;
              to += 1000;
            }
          } else {
            finished = true;
          }
        }

        if (allData.length > 0) {
          const mappedBills = allData.map(mapDbToBillData);
          
          // Apply fixes to data from Supabase
          let hasChanges = false;
          const updatedBills = mappedBills.map((b, i) => {
            let updatedBill = { ...b };
            let changed = false;

            const dbBill = allData[i];
            if (UCS_OPER.has(String(b.uc))) {
              if (dbBill.tipo !== 'OPER') {
                changed = true;
              }
            } else if (UCS_LIVRE_MERCADO_LIVRE.has(String(b.uc))) {
              let expectedMod = dbBill.modalidade_tarifaria || '';
              if (!expectedMod.toUpperCase().includes('LIVRE')) {
                expectedMod = expectedMod ? `${expectedMod} - LIVRE` : 'LIVRE';
              }
              if (dbBill.modalidade_tarifaria !== expectedMod || dbBill.tipo !== 'LIVRE') {
                changed = true;
              }
            }

            // Fix Elektro UC
            const isElektro = (b.concessionaria || '').toUpperCase().includes('ELEKTRO');
            if (isElektro && b.status === 'completed') {
              const fileNameNumbers = b.fileName.replace(/\.[^/.]+$/, "").replace(/\D/g, "");
              if (fileNameNumbers.length >= 5 && b.uc !== fileNameNumbers) {
                updatedBill.uc = fileNameNumbers;
                changed = true;
              }
            }

            // Fix 45839/2025 or other Excel dates
            if (updatedBill.mesReferencia === '45839' || updatedBill.mesReferencia === '45839/2025') {
              updatedBill.mesReferencia = 'Julho';
              updatedBill.anoLeitura = '2025';
              changed = true;
            } else if (/^\d{5}$/.test(updatedBill.mesReferencia) && parseInt(updatedBill.mesReferencia) > 40000) {
              const excelDate = parseInt(updatedBill.mesReferencia);
              const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
              const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
              updatedBill.mesReferencia = monthNames[jsDate.getUTCMonth()];
              updatedBill.anoLeitura = jsDate.getUTCFullYear().toString();
              changed = true;
            }

            if (changed) {
              hasChanges = true;
              return updatedBill;
            }
            return b;
          });

          setBills(prev => {
            const pendingBills = prev.filter(b => b.status === 'pending' || b.status === 'processing' || b.status === 'error');
            return deduplicateBills([...pendingBills, ...updatedBills]);
          });

          // If changes were made, update Supabase sequentially to avoid locking
          if (hasChanges) {
            const changedBills = updatedBills.filter((b, i) => 
              b.uc !== mappedBills[i].uc || 
              b.mesReferencia !== mappedBills[i].mesReferencia || 
              b.anoLeitura !== mappedBills[i].anoLeitura ||
              b.modalidadeTarifaria !== (allData[i].modalidade_tarifaria || '') ||
              b.tipo !== (allData[i].tipo || '')
            );
            for (const billToSave of changedBills) {
              try {
                const dbData = mapBillDataToDb(billToSave, user.id);
                await supabase.from('bills').update(dbData).eq('id', billToSave.id);
              } catch (err) {
                console.error('Erro ao atualizar dados no Supabase:', err);
              }
            }
          }
        } else {
          setBills(prev => {
            const pendingBills = prev.filter(b => b.status === 'pending' || b.status === 'processing' || b.status === 'error');
            return deduplicateBills([...pendingBills]);
          });
        }
      } catch (err) {
        console.error('Erro inesperado ao buscar faturas:', err);
      }
    };

    fetchBills();
  }, [isAuthenticated]);

  React.useEffect(() => {
    try {
      const billsToSave = bills.map(b => {
        // We cannot serialize File objects, so we remove it before saving
        const { file, ...rest } = b as any;
        return rest;
      });
      localStorage.setItem('sanesul_bills', JSON.stringify(billsToSave));

      // Save files to localforage
      const filesMap: Record<string, File> = {};
      bills.forEach(b => {
        if ((b as any).file) {
          filesMap[b.id] = (b as any).file;
        }
      });
      localforage.setItem('sanesul_bills_files', filesMap).catch(err => {
        console.warn('Failed to save files to localforage:', err);
      });
    } catch (e) {
      console.warn('LocalStorage limit reached, skipping save:', e);
    }
  }, [bills]);

  const [isProcessing, setIsProcessing] = useState(false);
  const isProcessingRef = useRef(false);

  useEffect(() => {
    isProcessingRef.current = isProcessing;
  }, [isProcessing]);
  const [isDragging, setIsDragging] = useState(false);
  const [currentPage, setCurrentPage] = useState<'visao_geral' | 'sistema'>('visao_geral');
  const [activeTab, setActiveTab] = useState<'faturas' | 'multas' | 'dashboard' | 'analises' | 'monitoramento' | 'monitoramento_reativo' | 'relatorio'>('faturas');
  const [multasMonth, setMultasMonth] = useState<string>('all');
  const [selectedMultaType, setSelectedMultaType] = useState<'ultrapassagem' | 'reativa' | 'subutilizacao' | 'total'>('total');
  const [multasSortDirection, setMultasSortDirection] = useState<'asc' | 'desc'>('desc');
  const [filterReference, setFilterReference] = useState<string>('all');
  const [sortConfig, setSortConfig] = useState<{ key: keyof BillData | 'referencia', direction: 'asc' | 'desc' } | null>(null);
  const [analysisData, setAnalysisData] = useState<any[]>([]);
  const [memoNumber, setMemoNumber] = useState(`001447/${new Date().getFullYear()}/GEDEO/DCO`);
  const [memoNfEnergisa, setMemoNfEnergisa] = useState('');
  const [memoNfElektro, setMemoNfElektro] = useState('');
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);

  const handleDownloadPDF = async () => {
    const element = document.getElementById('memo-content');
    if (!element) return;
    
    setIsGeneratingPDF(true);
    try {
      // Create a clone of the element to modify it for PDF generation
      const clone = element.cloneNode(true) as HTMLElement;
      
      // Update the memo number in the clone immediately
      const memoNumEl = clone.querySelector('.memo-number-text');
      if (memoNumEl) {
        memoNumEl.textContent = `MEMO Nº ${memoNumber}`;
      }

      // Create a temporary container off-screen to render the clone
      const container = document.createElement('div');
      container.style.position = 'absolute';
      container.style.left = '-9999px';
      container.style.top = '0';
      container.style.width = '210mm'; // A4 width
      container.appendChild(clone);
      document.body.appendChild(container);

      // Convert to image using dom-to-image-more (handles modern CSS like oklch better)
      const dataUrl = await domtoimage.toJpeg(clone, {
        quality: 0.98,
        bgcolor: '#ffffff',
        width: clone.offsetWidth,
        height: clone.offsetHeight,
        style: {
          transform: 'scale(1)',
          transformOrigin: 'top left'
        }
      });

      // Calculate dimensions for A4
      const pdf = new jsPDF('p', 'mm', 'a4');
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (clone.offsetHeight * pdfWidth) / clone.offsetWidth;

      pdf.addImage(dataUrl, 'JPEG', 0, 0, pdfWidth, pdfHeight);
      pdf.save(`memorando_faturamento_${new Date().getTime()}.pdf`);

      // Cleanup
      document.body.removeChild(container);
    } catch (error) {
      console.error('Erro ao gerar PDF:', error);
      showAlert('Erro', 'Erro ao gerar o PDF. Tente novamente.');
      setIsGeneratingPDF(false);
    } finally {
      setIsGeneratingPDF(false);
    }
  };

  const handleDownloadDocx = async () => {
    try {
      const doc = new Document({
        styles: {
          default: {
            document: {
              run: {
                font: "Arial",
                size: 24, // 12pt
              },
            },
          },
        },
        sections: [{
          properties: {},
          children: [
            // Header Text (Safe fallback instead of images to prevent corruption)
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "EMPRESA DE SANEAMENTO DE MATO GROSSO DO SUL S.A.", bold: true, size: 24, color: "0070C0" }),
              ]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "DIRETORIA DA PRESIDÊNCIA", bold: true, size: 20, color: "0070C0" }),
              ]
            }),
            new Paragraph({}),
            new Paragraph({}),
            
            // Memo Number
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [new TextRun({ text: `MEMO Nº ${memoNumber || "-"}`, bold: true, size: 24 })],
            }),
            new Paragraph({}),
            
            // Date
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [new TextRun({ text: `Campo Grande, ${new Intl.DateTimeFormat('pt-BR', { dateStyle: 'long' }).format(new Date())}.`, size: 24 })],
            }),
            new Paragraph({}),
            
            // To/From/Subject
            new Paragraph({ children: [new TextRun({ text: "DE: ", bold: true, size: 24 }), new TextRun({ text: "GEDEO - Gerência de Desenvolvimento Operacional", size: 24 })] }),
            new Paragraph({}),
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [new TextRun({ text: "PARA: ", bold: true, size: 24 }), new TextRun({ text: "GEFI - Gerência Financeira e Gestão de Recursos", size: 24 })]
            }),
            new Paragraph({}),
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({ text: "ASSUNTO: ", bold: true, size: 24 }), 
                new TextRun({ 
                  text: `Faturas Agrupadora Operacional Energisa e Agrupadora Elektro — ${selectedRelatorioMonth === 'all' ? 'Consolidado' : selectedRelatorioMonth}${!selectedRelatorioType.includes('all') ? ` (${selectedRelatorioType.join(', ')})` : ''}.`, 
                  size: 24 
                })
              ]
            }),
            new Paragraph({}),
            
            // Body
            new Paragraph({ children: [new TextRun({ text: "        Prezado(a),", size: 24 })] }),
            new Paragraph({}),
            new Paragraph({ 
              alignment: AlignmentType.JUSTIFIED,
              children: [
                new TextRun({ text: "        Seguem anexas para pagamento as faturas de energia elétrica Agrupadora da concessionária Energisa MS, e Agrupadora da concessionária Elektro — todas referentes ao mês de ", size: 24 }),
                new TextRun({ text: selectedRelatorioMonth === 'all' ? 'todos os períodos' : selectedRelatorioMonth, bold: true, color: "0070C0", size: 24 }),
                new TextRun({ text: !selectedRelatorioType.includes('all') ? ` (Tipo: ${selectedRelatorioType.join(', ')})` : '', size: 24 }),
                new TextRun({ text: " e correspondentes às unidades operacionais da SANESUL.", size: 24 })
              ] 
            }),
            new Paragraph({}),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Na Tabela 1 são especificadas as faturas anexas.", size: 24 })] }),
            new Paragraph({}),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tabela 1 - Faturas Anexas", bold: true, size: 24 })] }),
            
            // Table
            new Table({
              width: { size: 10000, type: WidthType.DXA },
              columnWidths: [4000, 2000, 2000, 2000],
              borders: {
                top: { style: BorderStyle.SINGLE, size: 4 },
                bottom: { style: BorderStyle.SINGLE, size: 4 },
                left: { style: BorderStyle.SINGLE, size: 4 },
                right: { style: BorderStyle.SINGLE, size: 4 },
                insideHorizontal: { style: BorderStyle.SINGLE, size: 4 },
                insideVertical: { style: BorderStyle.SINGLE, size: 4 },
              },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "LOCALIDADE", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "VALOR (R$)", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NOTA FISCAL", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REF: MÊS / ANO", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
                  ],
                }),
                // Energisa Main Row
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Agrupadora Energisa Operacional", bold: true, color: "0070C0" })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `R$ ${(memoData?.energisa?.total || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` })] })] }),
                    new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.energisa?.nf || "-" })] })] }),
                    new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.energisa?.mesRef || "-", bold: true, color: "0070C0" })] })] }),
                  ],
                }),
                // Energisa Details
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PIS", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.pis || 0) > 0 ? `R$ ${(memoData?.energisa?.pis || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "COFINS", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.cofins || 0) > 0 ? `R$ ${(memoData?.energisa?.cofins || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ICMS", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.icms || 0) > 0 ? `R$ ${(memoData?.energisa?.icms || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tarifa de Iluminação Pública", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.cip || 0) > 0 ? `R$ ${(memoData?.energisa?.cip || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                // Elektro Main Row
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Agrupadora Elektro", bold: true, color: "ED7D31" })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `R$ ${(memoData?.elektro?.total || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` })] })] }),
                    new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.elektro?.nf || "-" })] })] }),
                    new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.elektro?.mesRef || "-", bold: true, color: "ED7D31" })] })] }),
                  ],
                }),
                // Elektro Details
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PIS", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.pis || 0) > 0 ? `R$ ${(memoData?.elektro?.pis || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "COFINS", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.cofins || 0) > 0 ? `R$ ${(memoData?.elektro?.cofins || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ICMS", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.icms || 0) > 0 ? `R$ ${(memoData?.elektro?.icms || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tarifa de Iluminação Pública", size: 20 })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.cip || 0) > 0 ? `R$ ${(memoData?.elektro?.cip || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
                  ],
                }),
                // Total Row
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "TOTAL ( Agrupadora ENERGISA + ELEKTRO)", bold: true, italics: true })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `R$ ${((memoData?.energisa?.total || 0) + (memoData?.elektro?.total || 0)).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "-------------------" })] })] }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "-------------------" })] })] }),
                  ],
                }),
              ],
            }),
            new Paragraph({}),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Proc. N.º 694/2018    Nota Orçamentária Nº 003/2019", italics: true, size: 20 }),
              ],
            }),
            new Paragraph({
              pageBreakBefore: true,
              alignment: AlignmentType.JUSTIFIED,
              children: [
                new TextRun({ text: "        A planilha contendo a estratificação dos dados apresentados neste memorando está disponível em \\\\srv-fs-01\\DADOS\\DCO\\GEDEO\\OPERACAO_AGUA\\COTAA\\ENERGIA\\FATURAS.", size: 24 }),
              ],
            }),
            new Paragraph({}),
            new Paragraph({}),
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({ text: "Atenciosamente,", size: 24 }),
              ],
            }),
            new Paragraph({}),
            new Paragraph({}),
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({ text: "Fabio Roberto Alves da Silva", bold: true, size: 24 }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [
                new TextRun({ text: "Engenheiro Eletricista/GEDEO/Gerência de Desenvolvimento Operacional", size: 24 }),
              ],
            }),
          ],
        }],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `memorando_faturamento_${new Date().getTime()}.docx`);
    } catch (error) {
      console.error("Erro ao gerar DOCX:", error);
      showAlert('Erro', 'Ocorreu um erro ao gerar o arquivo DOCX. Verifique o console para mais detalhes.');
    }
  };

  const requestSort = (key: keyof BillData | 'referencia') => {
    if (sortConfig && sortConfig.key === key) {
      if (sortConfig.direction === 'asc') {
        setSortConfig({ key, direction: 'desc' });
      } else {
        setSortConfig(null);
      }
    } else {
      setSortConfig({ key, direction: 'asc' });
    }
  };

  const availableReferences = React.useMemo(() => {
    const refs = new Set<string>();
    bills.forEach(b => {
      if (b.mesReferencia && b.anoLeitura) {
        refs.add(`${formatMonth(b.mesReferencia)}/${b.anoLeitura}`);
      }
    });
    return Array.from(refs).sort((a, b) => {
      const [mA, yA] = a.split('/');
      const [mB, yB] = b.split('/');
      if (yA !== yB) return parseInt(yB) - parseInt(yA);
      return getMonthNumber(mB) - getMonthNumber(mA);
    });
  }, [bills]);

  const sortedBills = React.useMemo(() => {
    let filtered = [...bills];
    if (filterReference !== 'all') {
      filtered = filtered.filter(b => `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` === filterReference);
    }

    // UC Search Filter
    if (searchUC.trim() !== '') {
      const search = searchUC.toLowerCase().trim();
      filtered = filtered.filter(b => (b.uc || '').toLowerCase().includes(search));
    }

    let sortableBills = filtered;
    
    sortableBills.sort((a, b) => {
      // Priority mapping for statuses
      const getStatusPriority = (status: string) => {
        if (status === 'processing') return 0;
        if (status === 'error') return 1;
        return 2;
      };

      const aPriority = getStatusPriority(a.status);
      const bPriority = getStatusPriority(b.status);

      if (aPriority !== bPriority) {
        return aPriority - bPriority;
      }

      if (sortConfig !== null) {
        const extractNumericValue = (str: string | number) => {
          const s = String(str);
          const matches = s.match(/\d+/g);
          return matches ? parseInt(matches.join(''), 10) : 0;
        };

        let aValue: any;
        let bValue: any;

        if (sortConfig.key === 'referencia') {
          const monthA = getMonthNumber(a.mesReferencia);
          const monthB = getMonthNumber(b.mesReferencia);
          aValue = parseInt(a.anoLeitura || '0', 10) * 100 + monthA;
          bValue = parseInt(b.anoLeitura || '0', 10) * 100 + monthB;
        } else if (sortConfig.key === 'uc') {
          aValue = extractNumericValue(a.uc || '');
          bValue = extractNumericValue(b.uc || '');
        } else if (sortConfig.key === 'fileName') {
          aValue = extractNumericValue(a.fileName || '');
          bValue = extractNumericValue(b.fileName || '');
        } else if (sortConfig.key === 'concessionaria') {
          aValue = (a.concessionaria || '').toLowerCase();
          bValue = (b.concessionaria || '').toLowerCase();
        } else {
          aValue = a[sortConfig.key as keyof BillData];
          bValue = b[sortConfig.key as keyof BillData];
        }

        if (aValue < bValue) return sortConfig.direction === 'asc' ? -1 : 1;
        if (aValue > bValue) return sortConfig.direction === 'asc' ? 1 : -1;
      }

      // Default sort by createdAt descending (newest first)
      return (b.createdAt || 0) - (a.createdAt || 0);
    });

    return sortableBills;
  }, [bills, sortConfig, filterReference, searchUC]);

  const [analysisResults, setAnalysisResults] = useState<any>(null);
  const [dashboardSubTab, setDashboardSubTab] = useState<'operacionais' | 'financeiro'>('operacionais');
  const [operationalSubTab, setOperationalSubTab] = useState<'consumo' | 'ultrapassagem' | 'subutilizacao' | 'reativa' | 'solar'>('consumo');
  const [financialSubTab, setFinancialSubTab] = useState<'despesas' | 'multa_ultrapassagem' | 'multa_reativa' | 'tarifa_media' | 'energia_solar'>('despesas');
  const [selectedUC, setSelectedUC] = useState<string>('');
  const [selectedMonth, setSelectedMonth] = useState<string>('all');
  const [selectedConcessionaria, setSelectedConcessionaria] = useState<string>('all');
  const [selectedRelatorioMonth, setSelectedRelatorioMonth] = useState<string>('all');
  const [selectedRelatorioType, setSelectedRelatorioType] = useState<string[]>(['all']);
  const [isRelatorioTypeDropdownOpen, setIsRelatorioTypeDropdownOpen] = useState(false);
  const [selectedReactiveMonth, setSelectedReactiveMonth] = useState<string>('all');
  const [reactiveSortField, setReactiveSortField] = useState<string>('totalGeral');
  const [reactiveSortDirection, setReactiveSortDirection] = useState<'asc' | 'desc'>('desc');
  const [dashboardSort, setDashboardSort] = useState<{ key: string, direction: 'asc' | 'desc' }>({ key: 'name', direction: 'desc' });
  const [showMemo, setShowMemo] = useState(false);
  const [showMemoNumberPrompt, setShowMemoNumberPrompt] = useState(false);
  const [tempMemoNumber, setTempMemoNumber] = useState('');
  const [tempMemoNfEnergisa, setTempMemoNfEnergisa] = useState('');
  const [tempMemoNfElektro, setTempMemoNfElektro] = useState('');
  const [uploadProgress, setUploadProgress] = useState<Record<string, { 
    status: string, 
    percent: number, 
    fileName: string, 
    fileSize: number,
    abortController: AbortController | null 
  }>>({});
  const [isBillModalOpen, setIsBillModalOpen] = useState(false);
  const [editingBill, setEditingBill] = useState<Partial<BillData> | null>(null);
  const fileInputEnergisaRef = useRef<HTMLInputElement>(null);
  const fileInputElektroRef = useRef<HTMLInputElement>(null);
  const [agrupadoraFiles, setAgrupadoraFiles] = useState<Record<string, AgrupadoraData>>(() => {
    const saved = localStorage.getItem('sanesul_agrupadora_files');
    if (saved) {
      try {
        return JSON.parse(saved);
      } catch (e) {
        return {};
      }
    }
    return {};
  });

  const [showConfirmModal, setShowConfirmModal] = useState(false);
  const [confirmModalData, setConfirmModalData] = useState<{
    title: string;
    message: string;
    onConfirm: () => void;
    type: 'danger' | 'info';
    isAlert?: boolean;
  } | null>(null);

  const showAlert = (title: string, message: string) => {
    setConfirmModalData({
      title,
      message,
      onConfirm: () => setShowConfirmModal(false),
      type: 'info',
      isAlert: true
    });
    setShowConfirmModal(true);
  };

  const showConfirm = (title: string, message: string, onConfirm: () => void, type: 'danger' | 'info' = 'info') => {
    setConfirmModalData({
      title,
      message,
      onConfirm: () => {
        onConfirm();
        setShowConfirmModal(false);
      },
      type,
      isAlert: false
    });
    setShowConfirmModal(true);
  };

  React.useEffect(() => {
    localStorage.setItem('sanesul_agrupadora_files', JSON.stringify(agrupadoraFiles));
  }, [agrupadoraFiles]);
  const agrupadoraInputRef = useRef<HTMLInputElement>(null);
  const energisaInputRef = useRef<HTMLInputElement>(null);
  const detailedElektroInputRef = useRef<HTMLInputElement>(null);

  const handleAgrupadoraUpload = async (event: React.ChangeEvent<HTMLInputElement>, reportType: 'summary' | 'detailed' = 'summary') => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileId = `${file.name}-${Date.now()}`;
    const abortController = new AbortController();
    const statusPrefix = reportType === 'detailed' ? 'Relatório Detalhado' : 'Fatura Agrupadora';

    setUploadProgress(prev => ({
      ...prev,
      [fileId]: { status: `Lendo ${statusPrefix}...`, percent: 0, fileName: file.name, fileSize: file.size, abortController }
    }));
    
    const apiKey = (typeof process !== 'undefined' ? (process.env.GEMINI_API_KEY || process.env.API_KEY) : '') || import.meta.env.VITE_GEMINI_API_KEY || '';
    const ai = new GoogleGenAI({ apiKey });

    try {
      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve, reject) => {
        const timeout = setTimeout(() => reject(new Error('Erro ao ler arquivo: Tempo limite excedido (60s)')), 60000);
        
        reader.onprogress = (data) => {
          if (data.lengthComputable) {
            const progress = Math.round((data.loaded / data.total) * 30);
            setUploadProgress(prev => ({
              ...prev,
              [fileId]: { ...prev[fileId], status: `Lendo ${statusPrefix}...`, percent: progress }
            }));
          }
        };
        reader.onload = () => {
          clearTimeout(timeout);
          const base64 = (reader.result as string).split(',')[1];
          setUploadProgress(prev => ({
            ...prev,
            [fileId]: { ...prev[fileId], status: 'Processando com IA...', percent: 30 }
          }));
          resolve(base64);
        };
        reader.onerror = () => {
          clearTimeout(timeout);
          reject(new Error('Erro ao ler arquivo'));
        };
        reader.readAsDataURL(file);
      });

      const base64Data = await base64Promise;

      let prompt = "Você é um especialista em faturas agrupadoras de energia elétrica. Sua tarefa é extrair os dados consolidados desta fatura.\n\nINSTRUÇÕES:\n1. CONCESSIONÁRIA: Identifique se é ELEKTRO ou ENERGISA.\n2. VALOR TOTAL: Extraia o valor total a pagar da fatura agrupadora.\n3. REFERÊNCIA: Identifique o mês e ano de referência (ex: Fevereiro/2026).\n4. NOTA FISCAL: Procure pelo número da Nota Fiscal ou Fatura (ex: AGP-01... ou similar).\n5. IMPOSTOS: Extraia os valores de PIS, COFINS, ICMS e CIP. Para faturas da Energisa, os impostos federais (PIS/COFINS) podem estar agrupados como 'Imp. Fed.'.\n\nSe algum valor não for encontrado, retorne 0 ou string vazia.\n\nIMPORTANTE: SEMPRE RESPONDA EM PORTUGUÊS.";

      if (reportType === 'detailed') {
        prompt = "VOCÊ É UM AUDITOR CONTÁBIL ESPECIALISTA EM FATURAS DE ENERGIA. Sua tarefa é analisar TODAS AS PÁGINAS deste relatório detalhado para consolidar o valor da CIP.\n\nINSTRUÇÕES DETALHADAS:\n1. Percorra TODAS as páginas do documento, sem exceção.\n2. Em cada página, localize a tabela de itens faturados.\n3. Procure pelas descrições: 'COBRANCA ILUM PUBLICA', 'CIP', 'ILUMINACAO PUBLICA' ou 'CONTRIBUIÇÃO DE ILUMINAÇÃO PÚBLICA'.\n4. Extraia o valor monetário associado a cada uma dessas linhas.\n5. SOMA TOTAL: Você deve somar TODOS os valores encontrados em todas as páginas para obter o total da CIP do grupo.\n6. RETORNO: Retorne o JSON preenchendo o campo 'cip' com a soma total calculada. Os campos 'valorTotal', 'pis', 'cofins', 'icms' devem ser preenchidos como 0, a menos que você encontre um valor consolidado claro para eles no documento.\n7. Identifique a 'concessionaria' e o 'mesReferencia'.\n\nIMPORTANTE: SEMPRE RESPONDA EM PORTUGUÊS.";
      }

      if (abortController.signal.aborted) throw new Error('Upload cancelado');

      const progressInterval = setInterval(() => {
        setUploadProgress(prev => {
          if (!prev || !prev[fileId] || prev[fileId].percent >= 95) {
            clearInterval(progressInterval);
            return prev;
          }
          return { ...prev, [fileId]: { ...prev[fileId], percent: prev[fileId].percent + 5 } };
        });
      }, 1000);

      let response;
      try {
        response = await generateContentWithRetry(ai, {
          model: "gemini-3.1-flash-lite-preview",
          contents: [
            {
              parts: [
                {
                  inlineData: {
                    mimeType: file.type || 'application/pdf',
                    data: base64Data
                  }
                },
                { text: "Extraia os dados desta fatura seguindo as instruções do sistema." }
              ]
            }
          ],
          config: {
            responseMimeType: "application/json",
            responseSchema: AGRUPADORA_SCHEMA,
            thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
            systemInstruction: prompt
          }
        });
      } finally {
        clearInterval(progressInterval);
      }

        if (abortController.signal.aborted) throw new Error('Upload cancelado');

        setUploadProgress(prev => ({
          ...prev,
          [fileId]: { ...prev[fileId], status: 'Concluído!', percent: 100 }
        }));
        setTimeout(() => setUploadProgress(prev => {
          const next = { ...prev };
          delete next[fileId];
          return next;
        }), 2000);

        let text = response.text || '{}';
        text = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
        const result = JSON.parse(text);
        const concessionariaRaw = (result.concessionaria || 'DESCONHECIDA').toUpperCase();
        const key = concessionariaRaw.includes('ENERGISA') ? 'ENERGISA' : 'ELEKTRO';
        
        if (reportType === 'detailed') {
          const detailedKey = `${key}_DETALHADO`;
          setAgrupadoraFiles(prev => ({
            ...prev,
            [detailedKey]: {
              cip: typeof result.cip === 'string' ? parseValue(result.cip) : result.cip,
              concessionaria: `${concessionariaRaw} (DETALHADO)`,
              mesReferencia: formatMonth(result.mesReferencia || ''),
              valorTotal: 0,
              vencimento: '',
              numeroNotaFiscal: '',
              pis: 0,
              cofins: 0,
              icms: 0,
              fileName: file.name
            }
          }));
        } else {
          const newData: AgrupadoraData = {
            concessionaria: concessionariaRaw,
            valorTotal: typeof result.valorTotal === 'string' ? parseValue(result.valorTotal) : result.valorTotal,
            mesReferencia: formatMonth(result.mesReferencia || ''),
            vencimento: result.vencimento || '',
            numeroNotaFiscal: result.numeroNotaFiscal || '',
            pis: typeof result.pis === 'string' ? parseValue(result.pis) : result.pis,
            cofins: typeof result.cofins === 'string' ? parseValue(result.cofins) : result.cofins,
            icms: typeof result.icms === 'string' ? parseValue(result.icms) : result.icms,
            cip: typeof result.cip === 'string' ? parseValue(result.cip) : result.cip,
            fileName: file.name
          };

          setAgrupadoraFiles(prev => ({
            ...prev,
            [key]: newData
          }));
        }

    } catch (error: any) {
      console.error("Agrupadora extraction error:", {
        fileName: file.name,
        fileType: file.type,
        fileSize: file.size,
        reportType,
        error
      });
      const errorStr = error?.message || (typeof error === 'string' ? error : JSON.stringify(error));
      
      if (error?.isQuotaError || errorStr.includes('quota') || errorStr.includes('429') || errorStr.includes('RESOURCE_EXHAUSTED')) {
        showAlert('Limite Atingido', "Cota da API excedida. Verifique seu plano e detalhes de faturamento no Google AI Studio. Se você já tem um plano pago, aguarde alguns minutos.");
      } else {
        showAlert('Erro', "Erro ao processar fatura agrupadora: " + errorStr);
      }
      setUploadProgress(null);
    } finally {
      setIsProcessing(false);
      if (agrupadoraInputRef.current) agrupadoraInputRef.current.value = '';
      if (energisaInputRef.current) energisaInputRef.current.value = '';
      if (detailedElektroInputRef.current) detailedElektroInputRef.current.value = '';
    }
  };

  const downloadExcelTemplate = () => {
    const ws = XLSX.utils.json_to_sheet([{}], { header: EXCEL_COLUMNS.map(c => c.header) });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Modelo Faturas");
    XLSX.writeFile(wb, "Modelo_Importacao_Faturas.xlsx");
  };

  const handleExcelImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });

        if (data.length < 2) {
          alert('A planilha está vazia ou não possui dados.');
          return;
        }

        const headers = data[0] as string[];
        const rows = data.slice(1) as any[][];

        const newBills: BillData[] = rows.map((row, index) => {
          const billData: any = {
            id: crypto.randomUUID(),
            fileName: `Importado_${file.name}_Linha_${index + 1}`,
            status: 'completed',
            createdAt: Date.now() + index,
            tipo: 'normal'
          };

          headers.forEach((header, colIndex) => {
            const columnDef = EXCEL_COLUMNS.find(c => c.header === header);
            if (columnDef) {
              const value = row[colIndex];
              billData[columnDef.key] = value !== undefined && value !== null ? String(value) : '';
            }
          });

          return billData as BillData;
        }).map(bill => {
          bill.mercado = bill.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(bill.uc)) ? 'LIVRE' : 'CATIVO';
          if (bill.uc && UCS_OPER.has(String(bill.uc))) {
            bill.tipo = 'OPER';
          } else if (bill.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(bill.uc))) {
            let mod = bill.modalidadeTarifaria || '';
            if (!mod.toUpperCase().includes('LIVRE')) {
              bill.modalidadeTarifaria = mod ? `${mod} - LIVRE` : 'LIVRE';
            }
            bill.tipo = 'LIVRE';
          }
          return bill;
        });

        if (isSupabaseConfigured && isAuthenticated) {
          const { data: userData } = await supabase.auth.getUser();
          if (userData.user) {
            const dbData = newBills.map(bill => ({
              id: bill.id,
              user_id: userData.user!.id,
              file_name: bill.fileName,
              uc: bill.uc || '',
              demanda_ponta_kw: bill.demandaPontaKW || '',
              demanda_fora_ponta_kw: bill.demandaForaPontaKW || '',
              demanda_potencia_medida_ponta: bill.demandaPotenciaMedidaPonta || '',
              demanda_potencia_medida_fora_ponta: bill.demandaPotenciaMedidaForaPonta || '',
              ano_leitura: bill.anoLeitura || '',
              mes_referencia: bill.mesReferencia || '',
              consumo_kwh_ponta: bill.consumoKwhPonta || '',
              consumo_kwh_fora_ponta: bill.consumoKwhForaPonta || '',
              valor_consumo_kwh_ponta: bill.valorConsumoKwhPonta || '',
              valor_consumo_kwh_fora_ponta: bill.valorConsumoKwhForaPonta || '',
              valor_total: bill.valorTotal || '',
              cidade: bill.cidade || '',
              demanda_potencia_nao_consumida_ponta: bill.demandaPotenciaNaoConsumidaPonta || '',
              demanda_potencia_nao_consumida_f_ponta: bill.demandaPotenciaNaoConsumidaFPonta || '',
              demanda_potencia_ativa_ultrap_ponta: bill.demandaPotenciaAtivaUltrapPonta || '',
              demanda_potencia_ativa_ultrap_f_ponta: bill.demandaPotenciaAtivaUltrapFPonta || '',
              energia_reativa_exced_ponta: bill.energiaReativaExcedPonta || '',
              energia_reativa_exced_f_ponta: bill.energiaReativaExcedFPonta || '',
              energia_injetada_kwh: bill.energiaInjetadaKwh || '',
              energia_compensada_kwh: bill.energiaCompensadaKwh || '',
              valor_demanda_potencia_medida_ponta: bill.valorDemandaPotenciaMedidaPonta || '',
              valor_demanda_potencia_medida_fora_ponta: bill.valorDemandaPotenciaMedidaForaPonta || '',
              valor_demanda_potencia_nao_consumida_ponta: bill.valorDemandaPotenciaNaoConsumidaPonta || '',
              valor_demanda_potencia_nao_consumida_f_ponta: bill.valorDemandaPotenciaNaoConsumidaFPonta || '',
              valor_demanda_potencia_ativa_ultrap_ponta: bill.valorDemandaPotenciaAtivaUltrapPonta || '',
              valor_demanda_potencia_ativa_ultrap_f_ponta: bill.valorDemandaPotenciaAtivaUltrapFPonta || '',
              valor_energia_reativa_exced_ponta: bill.valorEnergiaReativaExcedPonta || '',
              valor_energia_reativa_exced_f_ponta: bill.valorEnergiaReativaExcedFPonta || '',
              energia_atv_injetada_gdi_ouc: bill.energiaAtvInjetadaGDIOUC || '',
              valor_energia_atv_injetada_gdi_ouc: bill.valorEnergiaAtvInjetadaGDIOUC || '',
              energia_atv_injetada_gdi_muc: bill.energiaAtvInjetadaGDIMUC || '',
              valor_energia_atv_injetada_gdi_muc: bill.valorEnergiaAtvInjetadaGDIMUC || '',
              cip: bill.cip || '',
              outros_encargos: bill.outrosEncargos || '',
              pis: bill.pis || '',
              cofins: bill.cofins || '',
              icms: bill.icms || '',
              concessionaria: bill.concessionaria || '',
              numero_nota_fiscal: bill.numeroNotaFiscal || '',
              modalidade_tarifaria: bill.modalidadeTarifaria || '',
              subgrupo: bill.subgrupo || '',
              tipo: bill.tipo || 'normal',
              mercado: bill.mercado || '',
              data_vencimento: bill.dataVencimento || '',
              status: bill.status,
              created_at: new Date(bill.createdAt || Date.now()).toISOString()
            }));
            
            let { error } = await supabase.from('bills').insert(dbData);
            
            if (error && (error.message.includes('data_vencimento') || error.message.includes('mercado') || error.details?.includes('data_vencimento') || error.details?.includes('mercado') || error.code === 'PGRST204')) {
              console.warn('Coluna data_vencimento ou mercado não encontrada. Inserindo sem elas...');
              const fallbackData = dbData.map((d: any) => {
                const { data_vencimento, mercado, ...rest } = d;
                return rest;
              });
              const fallbackRes = await supabase.from('bills').insert(fallbackData);
              error = fallbackRes.error;
            }
            
            if (error) {
              console.error('Erro ao salvar no Supabase:', error);
              alert('Erro ao salvar os dados importados no banco de dados.');
            }
          }
        }

        setBills(prev => [...prev, ...newBills]);
        alert(`${newBills.length} faturas importadas com sucesso!`);
      } catch (error) {
        console.error('Erro ao importar planilha:', error);
        alert('Erro ao processar o arquivo Excel.');
      }
    };
    reader.readAsBinaryString(file);
  };

  const addFiles = (files: FileList | File[], concessionaria?: string) => {
    const now = Date.now();
    const newBills: BillData[] = (Array.from(files) as File[]).map((file, index) => ({
      id: crypto.randomUUID(),
      createdAt: now + index,
      fileName: file.name,
      concessionaria: concessionaria || '',
      uc: '',
      demandaPontaKW: '',
      demandaForaPontaKW: '',
      demandaPotenciaMedidaPonta: '',
      demandaPotenciaMedidaForaPonta: '',
      anoLeitura: '',
      mesReferencia: '',
      consumoKwhPonta: '',
      valorConsumoKwhPonta: '',
      consumoKwhForaPonta: '',
      valorConsumoKwhForaPonta: '',
      valorTotal: '',
      cidade: '',
      demandaPotenciaNaoConsumidaPonta: '',
      demandaPotenciaNaoConsumidaFPonta: '',
      demandaPotenciaAtivaUltrapPonta: '',
      demandaPotenciaAtivaUltrapFPonta: '',
      energiaReativaExcedPonta: '',
      energiaReativaExcedFPonta: '',
      energiaInjetadaKwh: '',
      energiaCompensadaKwh: '',
      valorDemandaPotenciaMedidaPonta: '',
      valorDemandaPotenciaMedidaForaPonta: '',
      valorDemandaPotenciaNaoConsumidaPonta: '',
      valorDemandaPotenciaNaoConsumidaFPonta: '',
      valorDemandaPotenciaAtivaUltrapPonta: '',
      valorDemandaPotenciaAtivaUltrapFPonta: '',
      valorEnergiaReativaExcedPonta: '',
      valorEnergiaReativaExcedFPonta: '',
      energiaAtvInjetadaGDIOUC: '',
      valorEnergiaAtvInjetadaGDIOUC: '',
      energiaAtvInjetadaGDIMUC: '',
      valorEnergiaAtvInjetadaGDIMUC: '',
      cip: '',
      outrosEncargos: '',
      status: 'pending',
      file: file
    } as any));

    setBills(prev => deduplicateBills([...prev, ...newBills]));
  };

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>, concessionaria?: string) => {
    const files = event.target.files;
    if (!files) return;
    addFiles(files, concessionaria);
    event.target.value = '';
  };

  const runAnalysis = () => {
    const completedBills = bills.filter(b => b.status === 'completed');
    if (completedBills.length === 0) return;

    const parsedData = completedBills.map(b => {
      let dmp = parseValue(b.demandaPotenciaMedidaPonta);
      
      // Correção manual para UC 3401807 em Dezembro/2025 conforme solicitado pelo usuário
      // O usuário informou que houve uma leitura incorreta da demanda medida ponta neste mês.
      const isUC3401807 = String(b.uc) === '3401807';
      const isDec2025 = (b.mesReferencia?.toLowerCase() === 'dezembro' || b.mesReferencia === '12' || b.mesReferencia === '12/2025') && b.anoLeitura === '2025';
      
      if (isUC3401807 && isDec2025) {
        // Se a leitura foi incorreta (provavelmente um pico errôneo), ajustamos para 0 
        // para que não influencie o cálculo da demanda ideal (que usa o máximo do período).
        dmp = 0;
      }

      return {
        fileName: b.fileName,
        mes: b.mesReferencia || 'N/A',
        ano: b.anoLeitura || '',
        uc: b.uc || 'N/A',
        dcp: parseValue(b.demandaPontaKW),
        dmp: dmp,
        dcfp: parseValue(b.demandaForaPontaKW),
        dmfp: parseValue(b.demandaPotenciaMedidaForaPonta),
        modalidade: (b.modalidadeTarifaria || '').toUpperCase(),
        // Valores financeiros para o Gasto Real
        vDmpP: parseValue(b.valorDemandaPotenciaMedidaPonta),
        vDmfpFP: parseValue(b.valorDemandaPotenciaMedidaForaPonta),
        vUltrapP: parseValue(b.valorDemandaPotenciaAtivaUltrapPonta),
        vUltrapFP: parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta),
        vNaoConsP: parseValue(b.valorDemandaPotenciaNaoConsumidaPonta),
        vNaoConsFP: parseValue(b.valorDemandaPotenciaNaoConsumidaFPonta),
        tipo: b.tipo || '',
        mercado: UCS_LIVRE_MERCADO_LIVRE.has(b.uc) ? 'LIVRE' : 'CATIVO'
      };
    }).filter(d => d.dcp > 0 || d.dcfp > 0);

    // 2. Calculate Optimal Fixed Demand (Resolução 1000)
    // The optimal fixed demand is the one that minimizes the total cost over the period.
    // Group by UC to find the optimal demand per UC
    const ucs = Array.from(new Set(parsedData.map(d => d.uc)));
    const optimalDemands: Record<string, { ponta: number, foraPonta: number }> = {};

    ucs.forEach(uc => {
      const ucData = parsedData.filter(d => d.uc === uc);
      // Verifica se a UC tem contrato de ponta (ex: Tarifa Azul)
      // Se dcp for 0 em todos os meses, não sugerimos valor para ponta (ex: Tarifa Verde)
      // NOVO: Se a modalidade for VERDE, forçamos hasPontaContract para false
      const isVerde = ucData.some(d => d.modalidade.includes('VERDE'));
      const hasPontaContract = !isVerde && ucData.some(d => d.dcp > 0);
      
      const maxDmp = Math.max(...ucData.map(d => d.dmp));
      const maxDmfp = Math.max(...ucData.map(d => d.dmfp));
      
      const roundDemand = (val: number) => {
        const minRequired = val / 1.05;
        return Math.ceil(minRequired * 2) / 2;
      };

      optimalDemands[String(uc)] = {
        ponta: hasPontaContract ? roundDemand(maxDmp) : 0,
        foraPonta: Math.max(30, roundDemand(maxDmfp))
      };
    });

    // 3. Second Pass: Calculate costs and savings based on the fixed optimal demand
    const tp = 15.00; 
    const tfp = 10.00;

    const results = parsedData.flatMap(row => {
      const { mes, ano, uc, dcp, dmp, dcfp, dmfp, vDmpP, vDmfpFP, vUltrapP, vUltrapFP, vNaoConsP, vNaoConsFP, modalidade } = row;
      const opt = optimalDemands[String(uc)];
      
      if (!opt) return [];

      // Gasto Real conforme solicitado: Soma dos valores financeiros da fatura
      const currentTotal = vDmpP + vDmfpFP + vUltrapP + vUltrapFP + vNaoConsP + vNaoConsFP;
      
      // Para o cálculo da economia, precisamos dos custos base (Ponta e Fora Ponta)
      // que compõem esse Gasto Real.
      const costPonta = vDmpP + vUltrapP + vNaoConsP;
      const costForaPonta = vDmfpFP + vUltrapFP + vNaoConsFP;

      // Definir tarifas ideais com base na modalidade (Valores solicitados pelo usuário)
      const isAzul = modalidade.includes('AZUL');
      const isVerde = modalidade.includes('VERDE');
      
      const tp_ideal = isAzul ? 91.115690 : 0;
      const tfp_ideal = isAzul ? 45.702760 : (isVerde ? 43.177150 : 10.00);

      let optimizedTotal = 0;
      let economy = 0;

      if (isVerde) {
        // Fórmula específica solicitada pelo usuário para UC VERDE:
        // Economia = Gasto Real - (Valor Demanda de Potência Medida Ponta + (Demanda Ideal fora ponta * 45.702760))
        optimizedTotal = vDmpP + (opt.foraPonta * 45.702760);
        economy = currentTotal - optimizedTotal;
      } else {
        // Optimized Cost (using the FIXED optimal demand and specific rates)
        const optCostPonta = (dcp > 0 && opt.ponta > 0 && tp_ideal > 0)
            ? (dmp > opt.ponta * 1.05 ? (opt.ponta * tp_ideal) + ((dmp - opt.ponta) * tp_ideal * 2) : (Math.max(dmp, opt.ponta) * tp_ideal))
            : 0;
            
        const optCostForaPonta = dmfp > opt.foraPonta * 1.05 
          ? (opt.foraPonta * tfp_ideal) + ((dmfp - opt.foraPonta) * tfp_ideal * 2) 
          : (Math.max(dmfp, opt.foraPonta) * tfp_ideal);
        
        optimizedTotal = optCostPonta + optCostForaPonta;
        economy = currentTotal - optimizedTotal;
      }

      // Ultrapassagem
      const overrunPonta = (dcp > 0 && dmp > dcp * 1.05) ? dmp - dcp : 0;
      const overrunForaPonta = dmfp > dcfp * 1.05 ? dmfp - dcfp : 0;

      // Subutilização
      const subPonta = (dcp > 0 && dmp < dcp) ? dcp - dmp : 0;
      const subForaPonta = dmfp < dcfp ? dcfp - dmfp : 0;

      return [{
        fileName: row.fileName,
        mes,
        ano,
        uc,
        dcp,
        dmp,
        dcfp,
        dmfp,
        optimizedPonta: opt.ponta,
        optimizedForaPonta: opt.foraPonta,
        currentTotal,
        optimizedTotal,
        economy,
        overrunPonta,
        overrunForaPonta,
        subPonta,
        subForaPonta,
        isOverrun: overrunPonta > 0 || overrunForaPonta > 0,
        isSub: subPonta > 0 || subForaPonta > 0,
        tipo: row.tipo,
        mercado: row.mercado
      }];
    });

    // Sort results by month (newest to oldest)
    const monthOrder: Record<string, number> = {
      'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4,
      'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
      'outubro': 10, 'novembro': 11, 'dezembro': 12
    };

    results.sort((a, b) => {
      const monthA = monthOrder[a.mes.toLowerCase()] || 0;
      const monthB = monthOrder[b.mes.toLowerCase()] || 0;
      return monthB - monthA;
    });

    setAnalysisResults(results);
  };

  const runMonitoringAnalysis = () => {
    const completedBills = bills.filter(b => b.status === 'completed');
    if (completedBills.length === 0) return;

    const tp = 15.00; 
    const tfp = 10.00;

    const getYear = (year: string) => parseInt(year || '0', 10);

    // Get all unique UCs
    const ucs = Array.from(new Set(completedBills.map(b => b.uc))).filter(uc => {
      const ucBills = completedBills.filter(b => b.uc === uc);
      return ucBills.some(b => parseValue(b.demandaPontaKW) > 0 || parseValue(b.demandaForaPontaKW) > 0);
    });

    const allUcData = ucs.map(uc => {
      let ucBills = completedBills.filter(b => b.uc === uc);
      const city = ucBills[0]?.cidade || 'N/A';
      
      // Sort bills chronologically (Oldest to Newest) for change detection
      ucBills.sort((a, b) => {
        const yearA = getYear(a.anoLeitura);
        const yearB = getYear(b.anoLeitura);
        if (yearA !== yearB) return yearA - yearB;
        return getMonthNumber(a.mesReferencia) - getMonthNumber(b.mesReferencia);
      });

      // Calculate optimal demand (Ideal) based on all history (Max measured)
      // NOVO: Se a modalidade for VERDE, não consideramos a demanda ponta para o cálculo da demanda ideal
      const isVerde = ucBills.some(b => (b.modalidadeTarifaria || '').toUpperCase().includes('VERDE'));
      
      const maxDmp = isVerde ? 0 : Math.max(...ucBills.map(b => parseValue(b.demandaPotenciaMedidaPonta)));
      const maxDmfp = Math.max(...ucBills.map(b => parseValue(b.demandaPotenciaMedidaForaPonta)));
      
      const roundDemand = (val: number) => {
        const minRequired = val / 1.05;
        return Math.ceil(minRequired * 2) / 2;
      };

      const optPonta = (maxDmp > 0 && !isVerde) ? roundDemand(maxDmp) : 0;
      const optForaPonta = Math.max(30, roundDemand(maxDmfp));

      // 1st Pass: Calculate Current Total for ALL bills first
      const processedBills = ucBills.map(b => {
        const dcp = parseValue(b.demandaPontaKW);
        const dmp = parseValue(b.demandaPotenciaMedidaPonta);
        const dcfp = parseValue(b.demandaForaPontaKW);
        const dmfp = parseValue(b.demandaPotenciaMedidaForaPonta);

        const vDmpP = parseValue(b.valorDemandaPotenciaMedidaPonta);
        const vDmpFp = parseValue(b.valorDemandaPotenciaMedidaForaPonta);
        const vDncP = parseValue(b.valorDemandaPotenciaNaoConsumidaPonta);
        const vDncFp = parseValue(b.valorDemandaPotenciaNaoConsumidaFPonta);
        const vUltrapP = parseValue(b.valorDemandaPotenciaAtivaUltrapPonta);
        const vUltrapFp = parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta);

        const currentTotal = vDmpP + vDmpFp + vDncP + vDncFp + vUltrapP + vUltrapFp;

        return {
          originalBill: b,
          dcp, dmp, dcfp, dmfp,
          currentTotal,
          mes: b.mesReferencia,
          ano: b.anoLeitura
        };
      });

      // 2nd Pass: Detect Changes and Calculate Reference/Economy
      let activeContract: { ponta: number, foraPonta: number } | null = null;
      let currentContractStartIdx = 0;
      let previousContractAverageCost = 0;
      let accumulatedEconomy = 0;

      const monthlyData = processedBills.map((b, index) => {
        // Initialize active contract on first bill
        if (activeContract === null) {
          activeContract = { ponta: b.dcp, foraPonta: b.dcfp };
        }

        // Detect Change
        const isValidContract = activeContract.ponta > 0 || activeContract.foraPonta > 0;
        const hasChanged = isValidContract && (b.dcp !== activeContract.ponta || b.dcfp !== activeContract.foraPonta);
        
        if (hasChanged) {
          // Calculate average of the PREVIOUS contract period
          const previousPeriodBills = processedBills.slice(currentContractStartIdx, index);
          if (previousPeriodBills.length > 0) {
             const sum = previousPeriodBills.reduce((acc, item) => acc + item.currentTotal, 0);
             previousContractAverageCost = sum / previousPeriodBills.length;
          } else {
             previousContractAverageCost = 0;
          }
          
          // Update for new contract
          currentContractStartIdx = index;
          activeContract = { ponta: b.dcp, foraPonta: b.dcfp };
        } else if (!isValidContract && (b.dcp > 0 || b.dcfp > 0)) {
           // First valid contract found after 0s
           activeContract = { ponta: b.dcp, foraPonta: b.dcfp };
           currentContractStartIdx = index;
           // previousContractAverageCost remains 0 as there was no valid previous contract
        }

        // Calculate Economy
        // If we have a valid previous average (meaning we are in a changed state relative to something valid)
        let economyFromChange = 0;
        let referenceTotal = 0;

        if (previousContractAverageCost > 0) {
            referenceTotal = previousContractAverageCost;
            
            // Cálculo da Economia para Monitoramento: (Ref. Anterior - Gasto Real)
            economyFromChange = referenceTotal - b.currentTotal;
            
            accumulatedEconomy += economyFromChange;
        }

        return {
          mes: b.mes,
          ano: b.ano,
          currentTotal: b.currentTotal,
          referenceTotal, // Fixed Average of Previous Contract
          economy: economyFromChange,
          accumulatedEconomy,
          dmp: b.dmp,
          dmfp: b.dmfp,
          dcp: b.dcp,
          dcfp: b.dcfp,
          hasChanged,
          referenceContract: null // Not used in new logic but kept for type compatibility if needed
        };
      });

      // Reverse to show newest first in UI
      monthlyData.reverse();

      const totalEconomy = accumulatedEconomy; // Total accumulated from changes
      const totalCurrent = monthlyData.reduce((acc, curr) => acc + curr.currentTotal, 0);
      const hasContractChange = monthlyData.some(m => m.hasChanged);

      return {
        uc,
        city,
        totalEconomy,
        totalCurrent,
        monthlyData,
        optPonta,
        optForaPonta,
        hasContractChange
      };
    });

    const changedUCs = allUcData.filter(uc => uc.hasContractChange);
    const unchangedUCs = allUcData.filter(uc => !uc.hasContractChange);

    const generalTotalEconomy = allUcData.reduce((acc, curr) => acc + curr.totalEconomy, 0);
    const generalTotalCurrent = allUcData.reduce((acc, curr) => acc + curr.totalCurrent, 0);

    // Group by city for the chart
    const cityMap: Record<string, { city: string, totalEconomy: number, positiveEconomy: number, negativeEconomy: number, totalCurrent: number, optimized: number, ucs: { uc: string, economy: number }[], positiveUcs: { uc: string, economy: number }[], negativeUcs: { uc: string, economy: number }[] }> = {};
    allUcData.forEach(uc => {
      if (!cityMap[uc.city]) {
        cityMap[uc.city] = { city: uc.city, totalEconomy: 0, positiveEconomy: 0, negativeEconomy: 0, totalCurrent: 0, optimized: 0, ucs: [], positiveUcs: [], negativeUcs: [] };
      }
      cityMap[uc.city].totalEconomy += uc.totalEconomy;
      if (uc.totalEconomy > 0) {
        cityMap[uc.city].positiveEconomy += uc.totalEconomy;
        cityMap[uc.city].positiveUcs.push({ uc: String(uc.uc), economy: uc.totalEconomy });
      } else if (uc.totalEconomy < 0) {
        cityMap[uc.city].negativeEconomy += uc.totalEconomy;
        cityMap[uc.city].negativeUcs.push({ uc: String(uc.uc), economy: uc.totalEconomy });
      }
      cityMap[uc.city].totalCurrent += uc.totalCurrent;
      cityMap[uc.city].ucs.push({ uc: String(uc.uc), economy: uc.totalEconomy });
    });
    const cityData = Object.values(cityMap)
      .map(c => ({
        ...c,
        optimized: Math.max(0, c.totalCurrent - c.totalEconomy),
        ucs: c.ucs.sort((a, b) => b.economy - a.economy),
        positiveUcs: c.positiveUcs.sort((a, b) => b.economy - a.economy),
        negativeUcs: c.negativeUcs.sort((a, b) => a.economy - b.economy)
      }))
      .sort((a, b) => b.totalCurrent - a.totalCurrent);

    setMonitoringResults({
      changedUCs,
      unchangedUCs,
      generalTotalEconomy,
      generalTotalCurrent,
      cityData
    });
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
      addFiles(files);
    }
  };

  const processFile = async (bill: BillData & { file: File }, retryCount = 0, currentUser: any = null) => {
    const apiKey = (typeof process !== 'undefined' ? (process.env.GEMINI_API_KEY || process.env.API_KEY) : '') || import.meta.env.VITE_GEMINI_API_KEY || '';
    const ai = new GoogleGenAI({ apiKey });
    
    try {
      if (!bill.file || !(bill.file instanceof Blob)) {
        throw new Error('Arquivo não encontrado na memória. Por favor, remova esta fatura e faça o upload novamente.');
      }

      let user = currentUser;

      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve, reject) => {
        const timeout = setTimeout(() => reject(new Error('Erro ao ler arquivo: Tempo limite excedido (60s)')), 60000);
        
        reader.onprogress = (data) => {
          if (data.lengthComputable) {
            const progress = Math.round((data.loaded / data.total) * 30);
            setBills(prev => prev.map(b => b.id === bill.id ? { ...b, progress } : b));
          }
        };
        reader.onload = () => {
          clearTimeout(timeout);
          const base64 = (reader.result as string).split(',')[1];
          setBills(prev => prev.map(b => b.id === bill.id ? { ...b, progress: 30 } : b));
          resolve(base64);
        };
        reader.onerror = () => {
          clearTimeout(timeout);
          console.error("FileReader error:", reader.error);
          reject(new Error(`Erro ao ler arquivo: ${reader.error?.message || 'Erro desconhecido'}`));
        };
        if (bill.file instanceof Blob) {
          reader.readAsDataURL(bill.file);
        } else {
          clearTimeout(timeout);
          console.error("bill.file is not a Blob:", bill.file);
          resolve(""); // Or handle error appropriately
        }
      });

      const base64Data = await base64Promise;

      if (bill.abortController?.signal.aborted) throw new Error('Upload cancelado');

      // History check by file_name removed to allow re-extraction of files with the same name.
      // Duplicates will still be caught after extraction by UC/Mes/Ano.

      const progressInterval = setInterval(() => {
        setBills(prev => {
          const currentBill = prev.find(b => b.id === bill.id);
          if (!currentBill || (currentBill.progress || 0) >= 95) {
            clearInterval(progressInterval);
            return prev;
          }
          return prev.map(b => b.id === bill.id ? { ...b, progress: (b.progress || 0) + 5 } : b);
        });
      }, 1000);

      let response;
      try {
        response = await generateContentWithRetry(ai, {
          model: "gemini-3.1-flash-lite-preview",
          contents: [
            {
              parts: [
                { text: "Você é um especialista em análise de faturas de energia elétrica brasileiras. Sua tarefa é extrair com precisão absoluta os dados técnicos e financeiros da fatura fornecida.\n\nREGRAS DE EXTRAÇÃO:\n1. UNIDADE CONSUMIDORA (UC):\n   - Para ENERGISA: Procure por 'CÓDIGO DO CLIENTE' (ex: 10/1069-4) ou 'MATRÍCULA'. Extraia o identificador completo que identifica esta conta.\n   - Para ELEKTRO: A UC é o 'Código da Instalação'.\n2. VALORES NUMÉRICOS: Capture todos os dígitos. Não ignore o primeiro dígito de valores altos.\n3. CONSUMO E DEMANDA: Diferencie 'Contratada' de 'Medida'.\n4. GERAÇÃO DISTRIBUÍDA: Capture créditos de energia, injeção e compensação.\n5. TRIBUTOS: Extraia PIS, COFINS e ICMS separadamente.\n\nSe um campo não estiver presente, deixe em branco. Retorne o JSON seguindo o schema.\n\nIMPORTANTE: SEMPRE RESPONDA EM PORTUGUÊS." },
                {
                  inlineData: {
                    mimeType: bill.file?.type || 'application/pdf',
                    data: base64Data
                  }
                }
              ]
            }
          ],
          config: {
            responseMimeType: "application/json",
            responseSchema: EXTRACTION_SCHEMA,
            thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
            systemInstruction: "Você é um especialista em análise de faturas de energia elétrica. Extraia os dados da fatura com precisão, especialmente UC, Mês, Ano e Valores Totais. SEMPRE RESPONDA EM PORTUGUÊS."
          }
        });
      } finally {
        clearInterval(progressInterval);
      }

      if (bill.abortController?.signal.aborted) throw new Error('Upload cancelado');

        let result: any = {};
        try {
          let text = response.text || '{}';
          // Remove markdown code blocks if present
          text = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
          result = JSON.parse(text);
        } catch (parseError) {
          console.error(`Erro ao fazer parse do JSON para o arquivo ${bill.fileName}:`, parseError);
          console.error('Texto retornado pela API:', response.text);
          throw new Error('Falha ao interpretar a resposta da IA. O formato retornado não é um JSON válido.');
        }
        
        // Force UC to be the filename numbers for Elektro as requested by the user
        const concessionaria = (result.concessionaria || '').toUpperCase();
        if (concessionaria.includes('ELEKTRO')) {
          const fileNameNumbers = bill.fileName.replace(/\.[^/.]+$/, "").replace(/\D/g, "");
          if (fileNameNumbers.length >= 5) {
            result.uc = fileNameNumbers;
          }
        }

        if (result.uc) {
          let cleanUc = result.uc.replace(/UC:?\s*/i, '').trim();
          
          // Rule for formats like "10/2941716-9" -> extract "2941716"
          if (cleanUc.includes('/') && cleanUc.includes('-')) {
            const afterSlash = cleanUc.split('/')[1];
            if (afterSlash) {
              cleanUc = afterSlash.split('-')[0].trim();
            }
          } else if (cleanUc.includes('-')) {
            cleanUc = cleanUc.split('-')[0].trim();
          }
          
          result.uc = cleanUc;
        }

        if (result.mesReferencia) {
          result.mesReferencia = formatMonth(result.mesReferencia);
        }

      // --- CHECK FOR DUPLICATES IN DB AFTER EXTRACTION ---
      let existingDbId: string | null = null;

      if (isSupabaseConfigured && user && result.uc && result.mesReferencia && result.anoLeitura) {
        try {
          const { data: existingData, error: dbCheckError } = await supabase
            .from('bills')
            .select('id')
            .eq('uc', result.uc)
            .eq('mes_referencia', result.mesReferencia)
            .eq('ano_leitura', result.anoLeitura)
            .limit(1);
          
          if (!dbCheckError && existingData && existingData.length > 0) {
            existingDbId = existingData[0].id;
          }
        } catch (dbCheckErr) {
          console.warn('Erro ao verificar duplicatas no banco de dados:', dbCheckErr);
        }
      }

      // Check duplicates in current list using functional state update to avoid stale closures
      let finalStatus: 'completed' | 'error' = 'completed';
      let finalError: string | undefined = undefined;
      let isDuplicateInCurrentList = false;
      
      setBills(prev => {
        isDuplicateInCurrentList = prev.some(b => {
          if (b.id === bill.id || b.status !== 'completed') return false;
          
          const normalize = (str: string) => (str || '').toString().trim().toLowerCase();
          
          const hasKeys = result.uc && result.mesReferencia && result.anoLeitura;
          if (!hasKeys) return false;

          return normalize(b.uc) === normalize(result.uc) && 
                 normalize(b.mesReferencia) === normalize(result.mesReferencia) &&
                 normalize(b.anoLeitura) === normalize(result.anoLeitura);
        });

        return prev;
      });

      if (isDuplicateInCurrentList) {
        // We allow duplicates to proceed to update the DB and list
      }

      if (result.uc && UCS_OPER.has(String(result.uc))) {
        result.tipo = 'OPER';
      } else if (result.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(result.uc))) {
        let mod = result.modalidadeTarifaria || '';
        if (!mod.toUpperCase().includes('LIVRE')) {
          result.modalidadeTarifaria = mod ? `${mod} - LIVRE` : 'LIVRE';
        }
        result.tipo = 'LIVRE';
      }
      result.mercado = result.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(result.uc)) ? 'LIVRE' : 'CATIVO';

      const updatedBill: BillData = {
        ...bill,
        ...result,
        status: finalStatus,
        error: finalError
      };

      if (isSupabaseConfigured && user) {
        const dbData = mapBillDataToDb(updatedBill, user.id);
        
        if (existingDbId) {
          // Delete existing record and insert new one to "give place" to the new one
          await supabase
            .from('bills')
            .delete()
            .eq('id', existingDbId);
            
          let { error: insertError } = await supabase
            .from('bills')
            .insert(dbData);
            
          if (insertError && (insertError.message.includes('data_vencimento') || insertError.message.includes('mercado') || insertError.details?.includes('data_vencimento') || insertError.details?.includes('mercado') || insertError.code === 'PGRST204')) {
            console.warn('Coluna data_vencimento ou mercado não encontrada. Inserindo sem elas...');
            const { data_vencimento, mercado, ...fallbackData } = dbData;
            const fallbackRes = await supabase.from('bills').insert(fallbackData);
            insertError = fallbackRes.error;
          }
            
          if (insertError) {
            console.error('Erro ao substituir fatura no Supabase:', insertError);
          }
        } else if (!isDuplicateInCurrentList) {
          // Insert new record
          let { error: insertError } = await supabase
            .from('bills')
            .insert(dbData);
            
          if (insertError && (insertError.message.includes('data_vencimento') || insertError.message.includes('mercado') || insertError.details?.includes('data_vencimento') || insertError.details?.includes('mercado') || insertError.code === 'PGRST204')) {
            console.warn('Coluna data_vencimento ou mercado não encontrada. Inserindo sem elas...');
            const { data_vencimento, mercado, ...fallbackData } = dbData;
            const fallbackRes = await supabase.from('bills').insert(fallbackData);
            insertError = fallbackRes.error;
          }
            
          if (insertError) {
            console.error('Erro ao salvar fatura no Supabase:', insertError);
          }
        }
      }
      
      setBills(prev => deduplicateBills(prev.map(b => b.id === bill.id ? { ...updatedBill, progress: 100 } : b)));

    } catch (error: any) {
      if (error.message === 'Upload cancelado') {
        setBills(prev => prev.map(b => b.id === bill.id ? {
          ...b,
          status: 'error',
          error: 'Cancelado',
          progress: 0
        } : b));
        return;
      }

      console.error("Erro na extração:", error);
      
      let isRateLimit = false;
      let isQuotaExhausted = false;
      let isLockError = false;
      let isTransientError = false;
      let retryAfter = 0;

      // Check for rate limit in various error formats
      const errorStr = error?.message || (typeof error === 'string' ? error : JSON.stringify(error));
      const nestedError = error?.error || error;
      const errorCode = nestedError?.code || error?.status || 0;
      const errorStatus = nestedError?.status || '';
      const msg = nestedError?.message || error?.message || '';

      if (msg.includes('spending cap') || errorStr.includes('spending cap') || errorStr.includes('limite de gastos') || errorStr.includes('monthly limit')) {
        isQuotaExhausted = true;
      } else if (errorCode === 429 || errorStatus === 'RESOURCE_EXHAUSTED' || errorStr.includes('429') || errorStr.includes('RESOURCE_EXHAUSTED') || msg.toLowerCase().includes('quota') || msg.toLowerCase().includes('rate limit') || errorStr.includes('billing details')) {
        isRateLimit = true;
        // Try to extract retry time from message
        const match = msg.match(/retry in ([\d.]+)s/);
        if (match && match[1]) {
          retryAfter = parseFloat(match[1]) * 1000;
        } else {
          // Default to a longer delay for quota issues
          retryAfter = 15000;
        }
      } else if (errorStr.includes('Lock broken by another request')) {
        isLockError = true;
      } else if (errorCode === 500 || errorCode === 502 || errorCode === 503 || errorCode === 504 || errorStatus === 'INTERNAL' || errorStatus === 'UNAVAILABLE' || errorStr.includes('500') || errorStr.includes('502') || errorStr.includes('503') || errorStr.includes('504') || errorStr.includes('INTERNAL') || errorStr.includes('UNAVAILABLE')) {
        isTransientError = true;
      }

      if (isQuotaExhausted) {
        setBills(prev => prev.map(b => b.id === bill.id ? {
          ...b,
          status: 'error',
          error: 'O limite de gastos do seu projeto foi atingido. O processamento foi interrompido para evitar cobranças excedentes.'
        } : b));
        
        // Stop the entire processing queue
        setIsProcessing(false);
        isProcessingRef.current = false;
        return;
      }

      if (isRateLimit || isLockError || isTransientError) {
        // Limit max retries to 8 to avoid freezing the app for too long
        if (retryCount < 8) {
          // Use retryAfter if found, otherwise exponential backoff starting at 10s for rate limits
          const delay = retryAfter > 0 
            ? retryAfter + 3000 // Adiciona 3s de margem
            : Math.pow(2, retryCount) * 5000 + Math.random() * 2000;
          
          console.log(`[Worker] ${isLockError ? 'Erro de trava' : isTransientError ? 'Erro temporário (' + errorCode + ')' : 'Limite de taxa/cota'} atingido para ${bill.fileName}. Tentando novamente em ${Math.round(delay/1000)}s... (Tentativa ${retryCount + 1}/8)`);
          
          setBills(prev => prev.map(b => b.id === bill.id ? {
            ...b,
            status: 'processing',
            error: `Aguardando ${isLockError ? 'liberação' : isTransientError ? 'servidor' : 'limite de cota'} da API... Tentativa ${retryCount + 1}/8 (${Math.round(delay/1000)}s)`
          } : b));

          await new Promise(resolve => setTimeout(resolve, delay));
          return await processFile(bill, retryCount + 1, currentUser);
        } else {
          console.error(`[Worker] Falha após ${retryCount} tentativas para ${bill.fileName} devido a ${isLockError ? 'erro de trava' : 'limite de taxa'}.`);
        }
      }

      setBills(prev => prev.map(b => b.id === bill.id ? {
        ...b,
        status: 'error',
        error: error.message || 'Erro na extração'
      } : b));
    }
  };

  const startProcessing = async () => {
    if (isProcessing) return;
    
    const pendingBills = bills.filter(b => b.status === 'pending');
    if (pendingBills.length === 0) return;

    setIsProcessing(true);
    isProcessingRef.current = true;

    let currentUser = null;
    if (isSupabaseConfigured) {
      try {
        const { data: { user } } = await supabase.auth.getUser();
        currentUser = user;
      } catch (authError: any) {
        console.warn('Erro ao obter usuário no início do processamento:', authError);
        if (authError.message?.includes('Refresh Token Not Found')) {
          supabase.auth.signOut();
          setIsAuthenticated(false);
          setIsProcessing(false);
          isProcessingRef.current = false;
          return;
        }
      }
    }

    // Worker pool approach to maintain concurrency limited to 10 for faster processing
    const queue = [...pendingBills];
    const maxConcurrency = 10;
    const initialWorkers = Math.min(maxConcurrency, queue.length);

    const runWorker = async (workerId: number) => {
      // Add a small staggered start for workers to avoid simultaneous requests
      await new Promise(resolve => setTimeout(resolve, workerId * 500));
      
      while (queue.length > 0 && isProcessingRef.current) {
        const bill = queue.shift();
        if (!bill) break;

        // Update status to processing
        const abortController = new AbortController();
        setBills(prev => prev.map(b => b.id === bill.id ? { ...b, status: 'processing', error: undefined, abortController, progress: 0 } : b));
        
        try {
          console.log(`[Worker ${workerId}] Iniciando processamento de: ${bill.fileName} (Restam: ${queue.length})`);
          await processFile({ ...bill, abortController } as any, 0, currentUser);
          console.log(`[Worker ${workerId}] Concluído processamento de: ${bill.fileName}`);
          
          // Add a 1s delay to avoid overwhelming the API
          if (queue.length > 0 && isProcessingRef.current) {
            console.log(`[Worker ${workerId}] Aguardando 1s para o próximo arquivo...`);
            await new Promise(resolve => setTimeout(resolve, 1000));
          }
        } catch (error: any) {
          console.error(`[Worker ${workerId}] Erro crítico no processamento de ${bill.fileName}:`, error);
          setBills(prev => prev.map(b => b.id === bill.id ? { ...b, status: 'error', error: error.message || 'Erro crítico' } : b));
        }
      }
    };

    const workers = [];
    for (let i = 0; i < initialWorkers; i++) {
      workers.push(runWorker(i));
    }

    await Promise.all(workers);
    setIsProcessing(false);
  };

  const resetStuckProcesses = () => {
    setBills(prev => prev.map(b => b.status === 'processing' ? { ...b, status: 'pending', progress: 0 } : b));
    setIsProcessing(false);
    isProcessingRef.current = false;
  };

  const [selectedBills, setSelectedBills] = useState<string[]>([]);
  const [monitoringResults, setMonitoringResults] = useState<any>(null);
  const [expandedUCs, setExpandedUCs] = useState<Set<string>>(new Set());
  const [expandedAnalysisUCs, setExpandedAnalysisUCs] = useState<Set<string>>(new Set());
  const [expandedSummaryCities, setExpandedSummaryCities] = useState<Set<string>>(new Set());
  const [expandedReactiveUcs, setExpandedReactiveUcs] = useState<Set<string>>(new Set());

  const toggleReactiveUc = (uc: string) => {
    const newSet = new Set(expandedReactiveUcs);
    if (newSet.has(uc)) newSet.delete(uc);
    else newSet.add(uc);
    setExpandedReactiveUcs(newSet);
  };

  const handleReactiveSort = (field: string) => {
    if (reactiveSortField === field) {
      setReactiveSortDirection(reactiveSortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setReactiveSortField(field);
      setReactiveSortDirection('desc');
    }
  };

  const toggleSummaryCity = (city: string) => {
    setExpandedSummaryCities(prev => {
      const next = new Set(prev);
      if (next.has(city)) next.delete(city);
      else next.add(city);
      return next;
    });
  };

  const toggleUCExpansion = (uc: string) => {
    setExpandedUCs(prev => {
      const next = new Set(prev);
      if (next.has(uc)) next.delete(uc);
      else next.add(uc);
      return next;
    });
  };

  const toggleAnalysisUCExpansion = (uc: string) => {
    setExpandedAnalysisUCs(prev => {
      const next = new Set(prev);
      if (next.has(uc)) next.delete(uc);
      else next.add(uc);
      return next;
    });
  };

  const toggleBillSelection = (id: string) => {
    setSelectedBills(prev => 
      prev.includes(id) ? prev.filter(s => s !== id) : [...prev, id]
    );
  };

  const removeBill = async (id: string) => {
    if (isSupabaseConfigured && isAuthenticated) {
      try {
        const { error } = await supabase.from('bills').delete().eq('id', id);
        if (error) console.error('Erro ao deletar fatura do Supabase:', error);
      } catch (err) {
        console.error('Erro inesperado ao deletar fatura:', err);
      }
    }
    setBills(prev => prev.filter(b => b.id !== id));
    setSelectedBills(prev => prev.filter(s => s !== id));
  };

  const removeSelectedBills = async () => {
    if (isSupabaseConfigured && isAuthenticated && selectedBills.length > 0) {
      try {
        const { error } = await supabase.from('bills').delete().in('id', selectedBills);
        if (error) console.error('Erro ao deletar faturas do Supabase:', error);
      } catch (err) {
        console.error('Erro inesperado ao deletar faturas:', err);
      }
    }
    setBills(prev => prev.filter(b => !selectedBills.includes(b.id)));
    setSelectedBills([]);
  };

  const deselectFirst223 = () => {
    const first223Ids = bills.slice(0, 223).map(b => b.id);
    setSelectedBills(prev => prev.filter(id => !first223Ids.includes(id)));
  };

  const exportAnalysisToCSV = () => {
    if (!analysisResults || analysisResults.length === 0) return;

    const headers = [
      "Nome do Arquivo", "UC", "Tipo", "Mercado", "Ano", "Mês", "Demanda Medida Ponta", "Demanda Medida Fora Ponta",
      "Demanda Ideal Ponta", "Demanda Ideal Fora Ponta", "Gasto Real (R$)", "Economia (R$)", "Status",
      "Grupo Tarifário", "Tarifa Branca", "Optante B"
    ];

    const rows = analysisResults.map((r: any) => {
      // Basic classification logic (placeholder, needs refinement based on actual data)
      const isGrupoA = r.dcp > 0 || r.dcfp > 0; // Simplified assumption
      const isGrupoB = !isGrupoA;
      const isSolar = r.solarInjetadaOUC > 0 || r.solarInjetadaMUC > 0; // Assuming these fields exist in analysisResults
      
      const grupo = isGrupoA ? "Grupo A (Verde/Azul)" : (isSolar ? "Grupo B (Solar)" : "Grupo B (Não Solar)");
      const tarifaBranca = "N/A"; // Need to determine how to identify this
      const optanteB = "N/A"; // Need to determine how to identify this

      return [
        r.fileName,
        r.uc, 
        r.tipo,
        r.mercado,
        r.ano, r.mes, r.dmp, r.dmfp,
        r.optimizedPonta, r.optimizedForaPonta, 
        String((r.currentTotal || 0).toFixed(2)).replace('.', ','),
        String((r.economy || 0).toFixed(2)).replace('.', ','),
        r.isOverrun ? 'Ultrapassagem' : (r.isSub ? 'Subutilização' : 'OK'),
        grupo, tarifaBranca, optanteB
      ];
    });

    // Use semicolon as delimiter for better compatibility with Excel in many locales (like Brazil)
    // Add UTF-8 BOM (\uFEFF) to ensure Excel recognizes the encoding
    const csvContent = "\uFEFF" + [
      headers.join(';'),
      ...rows.map(row => row.map(val => {
        const safeVal = String(val || '').replace(/;/g, ',');
        return `"${safeVal}"`;
      }).join(';'))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `analise_demanda_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportToCSV = () => {
    // Use sortedBills which are already filtered by the selected reference
    const completedBills = sortedBills.filter(b => b.status === 'completed');
    if (completedBills.length === 0) {
      showAlert('Exportação', 'Não há faturas concluídas para exportar nesta referência.');
      return;
    }

    const headers = [
      "Nome do Arquivo",
      "UC",
      "Tipo",
      "Mercado",
      "Concessionária",
      "Cidade",
      "Mês Referência",
      "Ano Leitura",
      "Vencimento",
      "Nota Fiscal",
      "Modalidade Tarifária",
      "Subgrupo",
      "Valor Total (R$)",
      "Demanda Contratada Ponta (kW)",
      "Demanda Contratada Fora Ponta (kW)",
      "Demanda Medida Ponta (kW)",
      "Valor Demanda Medida Ponta (R$)",
      "Demanda Medida Fora Ponta (kW)",
      "Valor Demanda Medida Fora Ponta (R$)",
      "Consumo Ponta (kWh)",
      "Valor Consumo Ponta (R$)",
      "Consumo Fora Ponta (kWh)",
      "Valor Consumo Fora Ponta (R$)",
      "Demanda Não Consumida Ponta (kW)",
      "Valor Demanda Não Consumida Ponta (R$)",
      "Demanda Não Consumida Fora Ponta (kW)",
      "Valor Demanda Não Consumida Fora Ponta (R$)",
      "Ultrapassagem Ponta (kW)",
      "Valor Ultrapassagem Ponta (R$)",
      "Ultrapassagem Fora Ponta (kW)",
      "Valor Ultrapassagem Fora Ponta (R$)",
      "Reativa Excedente Ponta (kVArh)",
      "Valor Reativa Excedente Ponta (R$)",
      "Reativa Excedente Fora Ponta (kVArh)",
      "Valor Reativa Excedente Fora Ponta (R$)",
      "Energia Injetada (kWh)",
      "Energia Compensada (kWh)",
      "GDI oUC (kWh)",
      "Valor GDI oUC (R$)",
      "GDI mUC (kWh)",
      "Valor GDI mUC (R$)",
      "CIP (R$)",
      "Outros Encargos (R$)",
      "PIS (R$)",
      "COFINS (R$)",
      "ICMS (R$)"
    ];

    const formatCSVValue = (val: any) => {
      if (val === null || val === undefined) return '';
      let str = String(val);
      // If it looks like a number with a dot, replace it with a comma for Brazilian Excel
      if (!isNaN(Number(val)) && str.includes('.') && !str.includes(',')) {
        str = str.replace('.', ',');
      }
      // Replace any existing semicolons to avoid breaking the CSV structure
      str = str.replace(/;/g, ',');
      return `"${str}"`;
    };

    const rows = completedBills.map(b => [
      b.fileName,
      b.uc,
      b.tipo || '',
      UCS_LIVRE_MERCADO_LIVRE.has(b.uc) ? 'LIVRE' : 'CATIVO',
      b.concessionaria 
        ? (b.concessionaria.toUpperCase().includes('ENERGISA') 
            ? 'ENERGISA' 
            : b.concessionaria.toUpperCase().includes('ELEKTRO') 
              ? 'ELEKTRO' 
              : b.concessionaria)
        : '',
      b.cidade,
      b.mesReferencia,
      b.anoLeitura,
      b.dataVencimento || '',
      b.numeroNotaFiscal || '',
      b.modalidadeTarifaria || '',
      b.subgrupo || '',
      b.valorTotal,
      b.demandaPontaKW,
      b.demandaForaPontaKW,
      b.demandaPotenciaMedidaPonta,
      b.valorDemandaPotenciaMedidaPonta,
      b.demandaPotenciaMedidaForaPonta,
      b.valorDemandaPotenciaMedidaForaPonta,
      b.consumoKwhPonta,
      b.valorConsumoKwhPonta,
      b.consumoKwhForaPonta,
      b.valorConsumoKwhForaPonta,
      b.demandaPotenciaNaoConsumidaPonta,
      b.valorDemandaPotenciaNaoConsumidaPonta,
      b.demandaPotenciaNaoConsumidaFPonta,
      b.valorDemandaPotenciaNaoConsumidaFPonta,
      b.demandaPotenciaAtivaUltrapPonta,
      b.valorDemandaPotenciaAtivaUltrapPonta,
      b.demandaPotenciaAtivaUltrapFPonta,
      b.valorDemandaPotenciaAtivaUltrapFPonta,
      b.energiaReativaExcedPonta,
      b.valorEnergiaReativaExcedPonta,
      b.energiaReativaExcedFPonta,
      b.valorEnergiaReativaExcedFPonta,
      b.energiaInjetadaKwh,
      b.energiaCompensadaKwh,
      b.energiaAtvInjetadaGDIOUC,
      b.valorEnergiaAtvInjetadaGDIOUC,
      b.energiaAtvInjetadaGDIMUC,
      b.valorEnergiaAtvInjetadaGDIMUC,
      b.cip,
      b.outrosEncargos,
      b.pis || '',
      b.cofins || '',
      b.icms || ''
    ]);

    // Use semicolon as delimiter for better compatibility with Excel in many locales (like Brazil)
    // Add UTF-8 BOM (\uFEFF) to ensure Excel recognizes the encoding
    const csvContent = "\uFEFF" + [
      headers.join(';'),
      ...rows.map(row => row.map(formatCSVValue).join(';'))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    const fileName = filterReference === 'all' 
      ? `extracao_faturas_consolidado_${new Date().toISOString().split('T')[0]}.csv`
      : `extracao_faturas_${filterReference.replace('/', '_')}.csv`;
    
    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportRelatorioToCSV = () => {
    if (filteredRelatorioData.length === 0) {
      showAlert('Exportação', 'Não há dados para exportar com os filtros selecionados.');
      return;
    }

    const headers = [
      "Nome do Arquivo",
      "Mês/Ano",
      "UC",
      "Tipo",
      "Mercado",
      "Concessionária",
      "Cidade",
      "Nota Fiscal",
      "Modalidade Tarifária",
      "Subgrupo",
      "Valor Total (R$)",
      "Consumo Ponta (kWh)",
      "Consumo Fora Ponta (kWh)",
      "Demanda Medida Ponta (kW)",
      "Demanda Medida Fora Ponta (kW)",
      "Demanda Contratada Ponta (kW)",
      "Demanda Contratada Fora Ponta (kW)",
      "Ultrapassagem Ponta (kW)",
      "Ultrapassagem Fora Ponta (kW)",
      "Reativa Ponta (kVArh)",
      "Reativa Fora Ponta (kVArh)",
      "Solar Injetada OUC (kWh)",
      "Solar Injetada MUC (kWh)",
      "CIP (R$)",
      "Outros Encargos (R$)",
      "PIS (R$)",
      "COFINS (R$)",
      "ICMS (R$)"
    ];

    const formatCSVValue = (val: any) => {
      if (val === null || val === undefined) return '';
      let str = String(val);
      if (!isNaN(Number(val)) && str.includes('.') && !str.includes(',')) {
        str = str.replace('.', ',');
      }
      str = str.replace(/;/g, ',');
      return `"${str}"`;
    };

    const rows = filteredRelatorioData.map(d => [
      d.fileName,
      d.name,
      d.uc,
      d.tipo,
      d.mercado,
      d.concessionaria,
      d.cidade,
      d.numeroNotaFiscal,
      d.modalidadeTarifaria,
      d.subgrupo,
      d.valorTotal,
      d.consumoPonta,
      d.consumoForaPonta,
      d.demandaMedidaPonta,
      d.demandaMedidaForaPonta,
      d.demandaContratadaPonta,
      d.demandaContratadaForaPonta,
      d.ultrapassagemPonta,
      d.ultrapassagemForaPonta,
      d.reativaPonta,
      d.reativaForaPonta,
      d.solarInjetadaOUC,
      d.solarInjetadaMUC,
      d.cip,
      d.outrosEncargos,
      d.pis,
      d.cofins,
      d.icms
    ]);

    const csvContent = "\uFEFF" + [
      headers.join(';'),
      ...rows.map(row => row.map(formatCSVValue).join(';'))
    ].join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", `Relatorio_Financeiro_${selectedRelatorioMonth.replace('/', '_')}.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // --- Dashboard Data Processing ---

  const completedBills = bills.filter(b => b.status === 'completed');

  const { multasTotals, multasMonthlyData, multasUcList } = useMemo(() => {
    let ultrapassagem = 0;
    let reativa = 0;
    let subutilizacao = 0;
    let total = 0;
    
    const monthlyBreakdown: Record<string, { month: string, sortKey: string, ultrapassagem: number, reativa: number, subutilizacao: number, total: number }> = {};
    const ucBreakdown: Record<string, { cidade: string, ultrapassagem: number, reativa: number, subutilizacao: number, total: number }> = {};

    completedBills.forEach(b => {
      const u = parseValue(b.valorDemandaPotenciaAtivaUltrapPonta) + parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta);
      const r = parseValue(b.valorEnergiaReativaExcedPonta) + parseValue(b.valorEnergiaReativaExcedFPonta);
      const s = parseValue(b.valorDemandaPotenciaNaoConsumidaPonta) + parseValue(b.valorDemandaPotenciaNaoConsumidaFPonta);
      const t = u + r + s;

      const monthName = `${formatMonth(b.mesReferencia)}/${b.anoLeitura}`;
      const sortKey = `${b.anoLeitura}${getMonthNumber(b.mesReferencia.toString()).toString().padStart(2, '0')}`;
      
      if (!monthlyBreakdown[sortKey]) {
        monthlyBreakdown[sortKey] = { month: monthName, sortKey, ultrapassagem: 0, reativa: 0, subutilizacao: 0, total: 0 };
      }
      monthlyBreakdown[sortKey].ultrapassagem += u;
      monthlyBreakdown[sortKey].reativa += r;
      monthlyBreakdown[sortKey].subutilizacao += s;
      monthlyBreakdown[sortKey].total += t;

      if (multasMonth === 'all' || monthName === multasMonth) {
        ultrapassagem += u;
        reativa += r;
        subutilizacao += s;
        total += t;

        if (!ucBreakdown[b.uc]) {
          ucBreakdown[b.uc] = { cidade: b.cidade || 'N/A', ultrapassagem: 0, reativa: 0, subutilizacao: 0, total: 0 };
        }
        ucBreakdown[b.uc].ultrapassagem += u;
        ucBreakdown[b.uc].reativa += r;
        ucBreakdown[b.uc].subutilizacao += s;
        ucBreakdown[b.uc].total += t;
      }
    });

    const monthlyData = Object.values(monthlyBreakdown)
      .sort((a, b) => a.sortKey.localeCompare(b.sortKey));

    const ucList = Object.entries(ucBreakdown)
      .map(([uc, data]) => ({ uc, ...data }))
      .filter(d => d[selectedMultaType] > 0)
      .sort((a, b) => {
        if (multasSortDirection === 'asc') {
          return a[selectedMultaType] - b[selectedMultaType];
        } else {
          return b[selectedMultaType] - a[selectedMultaType];
        }
      });

    return {
      multasTotals: { ultrapassagem, reativa, subutilizacao, total },
      multasMonthlyData: monthlyData,
      multasUcList: ucList
    };
  }, [completedBills, selectedMultaType, multasMonth, multasSortDirection]);

  const dashboardData = completedBills.map(b => ({
    name: `${formatMonth(b.mesReferencia)}/${b.anoLeitura}`,
    uc: b.uc,
    consumoPonta: parseValue(b.consumoKwhPonta),
    valorConsumoPonta: parseValue(b.valorConsumoKwhPonta),
    consumoForaPonta: parseValue(b.consumoKwhForaPonta),
    valorConsumoForaPonta: parseValue(b.valorConsumoKwhForaPonta),
    valorTotal: parseValue(b.valorTotal),
    demandaMedidaPonta: parseValue(b.demandaPotenciaMedidaPonta),
    demandaMedidaForaPonta: parseValue(b.demandaPotenciaMedidaForaPonta),
    valorDemandaPonta: parseValue(b.valorDemandaPotenciaMedidaPonta),
    valorDemandaForaPonta: parseValue(b.valorDemandaPotenciaMedidaForaPonta),
    demandaContratadaPonta: parseValue(b.demandaPontaKW),
    demandaContratadaForaPonta: parseValue(b.demandaForaPontaKW),
    ultrapassagemPonta: parseValue(b.demandaPotenciaAtivaUltrapPonta),
    ultrapassagemForaPonta: parseValue(b.demandaPotenciaAtivaUltrapFPonta),
    reativaPonta: parseValue(b.energiaReativaExcedPonta),
    reativaForaPonta: parseValue(b.energiaReativaExcedFPonta),
    solarInjetada: parseValue(b.energiaInjetadaKwh),
    solarCompensada: parseValue(b.energiaCompensadaKwh),
    solarInjetadaOUC: parseValue(b.energiaAtvInjetadaGDIOUC),
    solarInjetadaMUC: parseValue(b.energiaAtvInjetadaGDIMUC),
    valorUltrapassagemPonta: parseValue(b.valorDemandaPotenciaAtivaUltrapPonta),
    valorUltrapassagemForaPonta: parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta),
    valorReativaPonta: parseValue(b.valorEnergiaReativaExcedPonta),
    valorReativaForaPonta: parseValue(b.valorEnergiaReativaExcedFPonta),
    valorSolarOUC: parseValue(b.valorEnergiaAtvInjetadaGDIOUC),
    valorSolarMUC: parseValue(b.valorEnergiaAtvInjetadaGDIMUC),
    cip: parseValue(b.cip),
    outrosEncargos: parseValue(b.outrosEncargos),
    pis: parseValue(b.pis),
    cofins: parseValue(b.cofins),
    icms: parseValue(b.icms),
    concessionaria: b.concessionaria || '',
    numeroNotaFiscal: b.numeroNotaFiscal || '',
    cidade: b.cidade,
    tipo: b.tipo || '',
    mercado: UCS_LIVRE_MERCADO_LIVRE.has(b.uc) ? 'LIVRE' : 'CATIVO',
    fileName: b.fileName || '',
    modalidadeTarifaria: (b.modalidadeTarifaria || '').toString().toUpperCase(),
    subgrupo: (b.subgrupo || '').toString().toUpperCase()
  }));

  const ucs = Array.from(new Set(dashboardData.map(d => d.uc))).filter(Boolean);

  const availableMonths = Array.from(new Set(dashboardData.map(d => d.name))).filter(Boolean).sort((a, b) => {
    const [mA, yA] = String(a).split('/');
    const [mB, yB] = String(b).split('/');
    if (yA !== yB) return Number(yB) - Number(yA);
    return getMonthNumber(mB) - getMonthNumber(mA);
  });

  const availableRelatorioTypes = Array.from(new Set(dashboardData.map(d => d.tipo))).filter(Boolean).sort();

  const filteredDashboardData = dashboardData.filter(d => {
    const matchesUC = !selectedUC || selectedUC === 'all' || d.uc.toString().includes(selectedUC);
    const matchesMonth = selectedMonth === 'all' || d.name === selectedMonth;
    const matchesConcessionaria = selectedConcessionaria === 'all' || d.concessionaria === selectedConcessionaria;
    
    if (!matchesUC || !matchesMonth || !matchesConcessionaria) return false;

    if (dashboardSubTab === 'financeiro' && financialSubTab === 'energia_solar') {
      const totalCreditos = Math.abs(d.valorSolarOUC + d.valorSolarMUC);
      return totalCreditos > 0;
    }

    if (dashboardSubTab === 'operacionais' && operationalSubTab === 'solar') {
      const totalInjetada = Math.abs(d.solarInjetadaOUC + d.solarInjetadaMUC);
      return totalInjetada > 0;
    }

    if (dashboardSubTab === 'operacionais' && operationalSubTab === 'reativa') {
      return d.reativaPonta > 0 || d.reativaForaPonta > 0;
    }

    if (dashboardSubTab === 'operacionais' && operationalSubTab === 'ultrapassagem') {
      return (d.ultrapassagemPonta + d.ultrapassagemForaPonta) > 0;
    }

    if (dashboardSubTab === 'operacionais' && operationalSubTab === 'subutilizacao') {
      return d.demandaContratadaPonta > 0 || d.demandaContratadaForaPonta > 0;
    }

    if (dashboardSubTab === 'financeiro' && financialSubTab === 'multa_reativa') {
      return d.valorReativaPonta > 0 || d.valorReativaForaPonta > 0;
    }

    if (dashboardSubTab === 'financeiro' && financialSubTab === 'multa_ultrapassagem') {
      return (d.valorUltrapassagemPonta + d.valorUltrapassagemForaPonta) > 0;
    }

    return true;
  });

  const generalFilteredData = dashboardData.filter(d => {
    const matchesUC = !selectedUC || selectedUC === 'all' || d.uc.toString().includes(selectedUC);
    const matchesMonth = selectedMonth === 'all' || d.name === selectedMonth;
    const matchesConcessionaria = selectedConcessionaria === 'all' || d.concessionaria === selectedConcessionaria;
    return matchesUC && matchesMonth && matchesConcessionaria;
  });

  const filteredRelatorioData = dashboardData.filter(d => {
    const matchesMonth = selectedRelatorioMonth === 'all' || d.name === selectedRelatorioMonth;
    const matchesType = selectedRelatorioType.includes('all') || selectedRelatorioType.includes(d.tipo);
    return matchesMonth && matchesType && d.uc !== '31383580';
  });

  const filteredUcs = Array.from(new Set(filteredDashboardData.map(d => d.uc))).filter(Boolean);

  const sortedDashboardData = React.useMemo(() => {
    return [...filteredDashboardData].sort((a, b) => {
      let aValue: any;
      let bValue: any;

      switch (dashboardSort.key) {
        case 'name':
          const [mA, yA] = a.name.split('/');
          const [mB, yB] = b.name.split('/');
          aValue = Number(yA) * 12 + getMonthNumber(mA);
          bValue = Number(yB) * 12 + getMonthNumber(mB);
          break;
        case 'uc':
          aValue = a.uc;
          bValue = b.uc;
          break;
        case 'total_kw':
          aValue = a.ultrapassagemPonta + a.ultrapassagemForaPonta;
          bValue = b.ultrapassagemPonta + b.ultrapassagemForaPonta;
          break;
        case 'utilizacao':
          aValue = (a.demandaMedidaPonta / (a.demandaContratadaPonta || 1));
          bValue = (b.demandaMedidaPonta / (b.demandaContratadaPonta || 1));
          break;
        case 'total_kvarh':
          aValue = a.reativaPonta + a.reativaForaPonta;
          bValue = b.reativaPonta + b.reativaForaPonta;
          break;
        default:
          aValue = a[dashboardSort.key as keyof typeof a];
          bValue = b[dashboardSort.key as keyof typeof b];
      }

      if (aValue === undefined || aValue === null) return 1;
      if (bValue === undefined || bValue === null) return -1;

      if (aValue < bValue) return dashboardSort.direction === 'asc' ? -1 : 1;
      if (aValue > bValue) return dashboardSort.direction === 'asc' ? 1 : -1;
      return 0;
    });
  }, [filteredDashboardData, dashboardSort]);

  const memoData = React.useMemo(() => {
    const energisa = filteredRelatorioData.filter(d => (d.concessionaria || '').toUpperCase().includes('ENERGISA'));
    const elektro = filteredRelatorioData.filter(d => (d.concessionaria || '').toUpperCase().includes('ELEKTRO'));

    const sum = (arr: any[], field: string) => arr.reduce((acc, curr) => acc + (curr[field] || 0), 0);

    const energisaData = {
      total: sum(energisa, 'valorTotal'),
      pis: sum(energisa, 'pis'),
      cofins: sum(energisa, 'cofins'),
      icms: sum(energisa, 'icms'),
      cip: sum(energisa, 'cip'),
      nf: memoNfEnergisa || agrupadoraFiles['ENERGISA']?.numeroNotaFiscal || '-',
      mesRef: selectedRelatorioMonth === 'all' 
        ? (agrupadoraFiles['ENERGISA'] ? formatReference(agrupadoraFiles['ENERGISA'].mesReferencia) : '-')
        : selectedRelatorioMonth
    };

    const elektroData = {
      total: sum(elektro, 'valorTotal'),
      pis: sum(elektro, 'pis'),
      cofins: sum(elektro, 'cofins'),
      icms: sum(elektro, 'icms'),
      cip: agrupadoraFiles['ELEKTRO_DETALHADO']?.cip || sum(elektro, 'cip') || 0,
      nf: memoNfElektro || agrupadoraFiles['ELEKTRO']?.numeroNotaFiscal || '-',
      mesRef: selectedRelatorioMonth === 'all' 
        ? (agrupadoraFiles['ELEKTRO'] ? formatReference(agrupadoraFiles['ELEKTRO'].mesReferencia) : '-')
        : selectedRelatorioMonth
    };

    return {
      energisa: energisaData,
      elektro: elektroData
    };
  }, [filteredRelatorioData, agrupadoraFiles, memoNfEnergisa, memoNfElektro, selectedRelatorioMonth]);

  // Group by month/year for charts
  const timeSeriesData = Object.values(filteredDashboardData.reduce((acc: any, curr) => {
    const key = curr.name;
    if (!acc[key]) {
      acc[key] = { 
        name: key, 
        consumoPonta: 0, 
        valorConsumoPonta: 0,
        consumoForaPonta: 0, 
        valorConsumoForaPonta: 0,
        valorTotal: 0,
        demandaMedidaPonta: 0,
        demandaMedidaForaPonta: 0,
        demandaContratadaPonta: 0,
        demandaContratadaForaPonta: 0,
        ultrapassagemPonta: 0,
        ultrapassagemForaPonta: 0,
        reativaPonta: 0,
        reativaForaPonta: 0,
        solarInjetada: 0,
        solarCompensada: 0,
        solarInjetadaOUC: 0,
        solarInjetadaMUC: 0,
        valorUltrapassagemPonta: 0,
        valorUltrapassagemForaPonta: 0,
        valorReativaPonta: 0,
        valorReativaForaPonta: 0,
        valorSolarOUC: 0,
        valorSolarMUC: 0,
        cip: 0,
        outrosEncargos: 0
      };
    }
    acc[key].consumoPonta += curr.consumoPonta;
    acc[key].valorConsumoPonta += curr.valorConsumoPonta;
    acc[key].consumoForaPonta += curr.consumoForaPonta;
    acc[key].valorConsumoForaPonta += curr.valorConsumoForaPonta;
    acc[key].valorTotal += curr.valorTotal;
    acc[key].demandaMedidaPonta = Math.max(acc[key].demandaMedidaPonta, curr.demandaMedidaPonta);
    acc[key].demandaMedidaForaPonta = Math.max(acc[key].demandaMedidaForaPonta, curr.demandaMedidaForaPonta);
    acc[key].demandaContratadaPonta = Math.max(acc[key].demandaContratadaPonta, curr.demandaContratadaPonta);
    acc[key].demandaContratadaForaPonta = Math.max(acc[key].demandaContratadaForaPonta, curr.demandaContratadaForaPonta);
    acc[key].ultrapassagemPonta += curr.ultrapassagemPonta;
    acc[key].ultrapassagemForaPonta += curr.ultrapassagemForaPonta;
    acc[key].reativaPonta += curr.reativaPonta;
    acc[key].reativaForaPonta += curr.reativaForaPonta;
    acc[key].solarInjetada += curr.solarInjetada;
    acc[key].solarCompensada += curr.solarCompensada;
    acc[key].solarInjetadaOUC += curr.solarInjetadaOUC;
    acc[key].solarInjetadaMUC += curr.solarInjetadaMUC;
    acc[key].valorUltrapassagemPonta += curr.valorUltrapassagemPonta;
    acc[key].valorUltrapassagemForaPonta += curr.valorUltrapassagemForaPonta;
    acc[key].valorReativaPonta += curr.valorReativaPonta;
    acc[key].valorReativaForaPonta += curr.valorReativaForaPonta;
    acc[key].valorSolarOUC += curr.valorSolarOUC;
    acc[key].valorSolarMUC += curr.valorSolarMUC;
    acc[key].cip += curr.cip;
    acc[key].outrosEncargos += curr.outrosEncargos;
    acc[key].valorConsumo = acc[key].valorTotal + Math.abs(acc[key].valorSolarOUC + acc[key].valorSolarMUC) - acc[key].cip - acc[key].outrosEncargos;
    return acc;
  }, {}));

  const COLORS = ['#0054A6', '#00AEEF', '#1E293B', '#64748B'];

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-sanesul-bg flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-2xl shadow-xl w-full max-w-md border border-sanesul-primary/10">
          <div className="mb-8">
            <Logo className="h-16" isLogin={true} />
          </div>

          <form onSubmit={handleLogin} className="space-y-6">
            {!isSupabaseConfigured && (
              <div className="p-4 bg-amber-50 border border-amber-200 rounded-xl flex items-start gap-3">
                <AlertCircle className="text-amber-500 shrink-0" size={18} />
                <p className="text-xs text-amber-700 leading-relaxed">
                  <strong>Atenção:</strong> O banco de dados não está configurado. O login não funcionará até que as chaves do Supabase sejam adicionadas.
                </p>
              </div>
            )}
            <div>
              <label className="block text-xs font-bold text-sanesul-muted uppercase tracking-wider mb-2">Usuário</label>
              <input
                type="text"
                value={loginUsername}
                onChange={(e) => setLoginUsername(e.target.value)}
                className="w-full px-4 py-3 bg-sanesul-bg border border-sanesul-primary/20 rounded-xl text-sanesul-text focus:outline-none focus:ring-2 focus:ring-sanesul-primary/50 transition-all"
                placeholder="Digite seu usuário"
                required
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-sanesul-muted uppercase tracking-wider mb-2">Senha</label>
              <input
                type="password"
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
                className="w-full px-4 py-3 bg-sanesul-bg border border-sanesul-primary/20 rounded-xl text-sanesul-text focus:outline-none focus:ring-2 focus:ring-sanesul-primary/50 transition-all"
                placeholder="Digite sua senha"
                required
              />
            </div>
            
            {loginError && (
              <p className="text-red-500 text-sm font-medium text-center">{loginError}</p>
            )}

            <button
              type="submit"
              className="w-full py-4 bg-sanesul-primary text-white rounded-xl font-bold tracking-wider shadow-lg shadow-sanesul-primary/20 hover:bg-sanesul-primary/90 transition-all active:scale-95"
            >
              ENTRAR
            </button>
          </form>
        </div>
      </div>
    );
  }

  if (currentPage === 'visao_geral') {
    return <VisaoGeralDashboard 
      data={dashboardData} 
      setCurrentPage={setCurrentPage} 
      handleLogout={handleLogout} 
      hasApiKey={hasApiKey}
      handleSelectKey={handleSelectKey}
    />;
  }

  return (
    <div className="min-h-screen bg-sanesul-bg text-sanesul-text font-sans p-4 md:p-8">
      {/* Header */}
      <header className="max-w-[1600px] mx-auto mb-12 border-b border-sanesul-primary/10 pb-8">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-10">
          <Logo className="h-12" />
          <div className="flex flex-wrap gap-3">
            {activeTab === 'faturas' && (
              <></>
            )}
            <button
              onClick={() => setCurrentPage('visao_geral')}
              className="flex items-center gap-2 px-6 py-3 bg-white border border-sanesul-primary/20 text-sanesul-primary hover:bg-sanesul-primary/5 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
            >
              <ArrowLeft size={16} />
              Voltar para Visão Geral
            </button>
            <button
              onClick={handleSelectKey}
              className={`flex items-center gap-2 px-4 py-3 ${hasApiKey ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 'bg-amber-50 text-amber-700 border-amber-200'} border hover:opacity-80 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95`}
              title={hasApiKey ? "Trocar Chave de API" : "Selecionar Chave de API"}
            >
              <Key size={16} />
              {hasApiKey ? "Trocar Conta" : "Configurar API"}
            </button>
            <button
              onClick={handleLogout}
              className="flex items-center gap-2 px-4 py-3 bg-white border border-sanesul-primary/20 text-sanesul-primary hover:bg-sanesul-primary/5 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
              title="Sair"
            >
              <LogOut size={16} />
              Sair
            </button>
            {activeTab === 'faturas' && (
              <>
                <button
                  onClick={() => {
                    setEditingBill({
                      id: crypto.randomUUID(),
                      fileName: 'Fatura Manual',
                      status: 'completed',
                      tipo: 'OPERACIONAL',
                      concessionaria: 'ENERGISA',
                      mesReferencia: `${formatMonth((new Date().getMonth() + 1).toString().padStart(2, '0'))}/${new Date().getFullYear()}`,
                      anoLeitura: new Date().getFullYear().toString()
                    });
                    setIsBillModalOpen(true);
                  }}
                  className="flex items-center gap-2 px-6 py-3 bg-white border border-sanesul-primary/20 text-sanesul-primary hover:bg-sanesul-primary/5 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
                >
                  <Plus size={16} />
                  Nova Fatura Manual
                </button>
                <button
                  onClick={downloadExcelTemplate}
                  className="flex items-center gap-2 px-6 py-3 bg-emerald-50 border border-emerald-200 text-emerald-700 hover:bg-emerald-100 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
                >
                  <Download size={16} />
                  Baixar Modelo
                </button>
                <label className="flex items-center gap-2 px-6 py-3 bg-blue-50 border border-blue-200 text-blue-700 hover:bg-blue-100 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95 cursor-pointer">
                  <FileSpreadsheet size={16} />
                  Importar Planilha
                  <input 
                    type="file" 
                    accept=".xlsx, .xls" 
                    className="hidden" 
                    onChange={handleExcelImport} 
                  />
                </label>
                <button
                  onClick={() => fileInputEnergisaRef.current?.click()}
                  className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white hover:bg-sanesul-primary/90 transition-all rounded-xl text-xs font-bold tracking-wider shadow-lg shadow-sanesul-primary/20 active:scale-95"
                >
                  <Plus size={16} />
                  Adicionar Faturas - ENERGISA
                </button>
                <input
                  type="file"
                  ref={fileInputEnergisaRef}
                  onChange={(e) => handleFileUpload(e, 'ENERGISA')}
                  multiple
                  accept="application/pdf,image/*"
                  className="hidden"
                />
                <button
                  onClick={() => fileInputElektroRef.current?.click()}
                  className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white hover:bg-sanesul-primary/90 transition-all rounded-xl text-xs font-bold tracking-wider shadow-lg shadow-sanesul-primary/20 active:scale-95"
                >
                  <Plus size={16} />
                  Adicionar Faturas - ELEKTRO
                </button>
                <input
                  type="file"
                  ref={fileInputElektroRef}
                  onChange={(e) => handleFileUpload(e, 'ELEKTRO')}
                  multiple
                  accept="application/pdf,image/*"
                  className="hidden"
                />
              </>
            )}
            {bills.length > 0 && activeTab === 'faturas' && (
              <button
                onClick={startProcessing}
                disabled={isProcessing || !bills.some(b => b.status === 'pending')}
                className="flex items-center gap-2 px-6 py-3 bg-white border border-sanesul-primary/20 text-sanesul-primary hover:bg-sanesul-primary/5 transition-all rounded-xl text-xs font-bold tracking-wider disabled:opacity-30 disabled:cursor-not-allowed shadow-sm active:scale-95"
              >
                {isProcessing ? <Loader2 size={16} className="animate-spin" /> : <FileText size={16} />}
                Processar Arquivos
              </button>
            )}
            {bills.some(b => b.status === 'processing') && (
              <button
                onClick={resetStuckProcesses}
                className="flex items-center gap-2 px-6 py-3 bg-red-50 border border-red-200 text-red-600 hover:bg-red-100 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
                title="Reseta faturas que ficaram presas no status 'Extraindo'"
              >
                <RotateCcw size={16} />
                Resetar Travados
              </button>
            )}
            {bills.some(b => b.status === 'completed') && activeTab === 'faturas' && (
              <button
                onClick={exportToCSV}
                className="flex items-center gap-2 px-6 py-3 bg-sanesul-secondary text-white hover:bg-sanesul-secondary/90 transition-all rounded-xl text-xs font-bold tracking-wider shadow-lg shadow-sanesul-secondary/20 active:scale-95"
              >
                <Download size={16} />
                Exportar CSV
              </button>
            )}
            {bills.length > 0 && activeTab === 'faturas' && (
              <div className="flex items-center gap-2 bg-white border border-sanesul-primary/20 rounded-xl px-4 py-2 shadow-sm">
                <Filter size={14} className="text-sanesul-primary" />
                <select
                  value={filterReference}
                  onChange={(e) => setFilterReference(e.target.value)}
                  className="bg-transparent text-xs font-bold text-sanesul-primary outline-none cursor-pointer min-w-[140px]"
                >
                  <option value="all">Todas as Referências</option>
                  {availableReferences.map(ref => (
                    <option key={ref} value={ref}>{ref}</option>
                  ))}
                </select>
              </div>
            )}
            {bills.length > 0 && activeTab === 'faturas' && (
              <button
                onClick={() => setSelectedBills(bills.filter(b => b.status === 'error' || b.status === 'pending').map(b => b.id))}
                className="flex items-center gap-2 px-6 py-3 bg-white border border-sanesul-primary/20 text-sanesul-primary hover:bg-sanesul-primary/5 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
              >
                <CheckSquare size={16} />
                Selecionar Pendentes/Erro
              </button>
            )}
            {(bills.length > 0 || Object.keys(agrupadoraFiles).length > 0) && activeTab === 'faturas' && (
              <></>
            )}

          </div>
        </div>

        {/* Tabs */}
        <div className="flex gap-2 p-1 bg-white/50 backdrop-blur-sm border border-sanesul-primary/10 rounded-2xl w-fit">
          <button 
            onClick={() => setActiveTab('faturas')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'faturas' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <FileText size={14} />
            Gestão de Faturas
          </button>
          <button 
            onClick={() => setActiveTab('multas')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'multas' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <AlertTriangle size={14} />
            Multas
          </button>
          <button 
            onClick={() => setActiveTab('dashboard')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'dashboard' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <LayoutDashboard size={14} />
            Dashboard Analítico
          </button>
          <button 
            onClick={() => setActiveTab('analises')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'analises' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <BarChart3 size={14} />
            Análises de Dados
          </button>
          <button 
            onClick={() => setActiveTab('monitoramento')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'monitoramento' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <DollarSign size={14} />
            Monitoramento de Despesas
          </button>
          <button 
            onClick={() => setActiveTab('monitoramento_reativo')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'monitoramento_reativo' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <Zap size={14} />
            Monitoramento Reativo
          </button>
          <button 
            onClick={() => setActiveTab('relatorio')}
            className={`flex items-center gap-2 px-8 py-3 transition-all rounded-xl text-xs font-bold tracking-wide ${
              activeTab === 'relatorio' 
                ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                : 'text-sanesul-muted hover:text-sanesul-primary hover:bg-white'
            }`}
          >
            <FileText size={14} />
            Relatório Financeiro
          </button>
        </div>
      </header>

      {bills.some(b => b.error?.includes('limite de gastos') || b.error?.includes('Cota')) && (
        <div className="max-w-[1600px] mx-auto mb-6 p-4 bg-red-50 border border-red-200 rounded-2xl flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-3 text-red-700">
            <AlertCircle size={20} />
            <p className="text-sm font-medium">
              <strong>Atenção:</strong> O limite de gastos ou cota da sua chave API foi atingido. 
              Verifique seu faturamento no Google Cloud.
            </p>
          </div>
        </div>
      )}

      <main 
        className="max-w-[1600px] mx-auto"
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {activeTab === 'faturas' ? (
          bills.length === 0 ? (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center py-12">
              <div className="space-y-8">
                <div className="inline-flex items-center gap-2 px-4 py-2 bg-sanesul-primary/10 rounded-full">
                  <span className="w-2 h-2 rounded-full bg-sanesul-primary animate-pulse" />
                  <span className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">Pronto para processar</span>
                </div>
                <h2 className="text-5xl md:text-6xl font-display font-bold text-sanesul-primary leading-[1.1] tracking-tight">
                  Transforme suas <span className="text-sanesul-secondary">faturas</span> em inteligência.
                </h2>
                <p className="text-lg text-sanesul-muted max-w-md leading-relaxed">
                  Nossa IA extrai automaticamente todos os indicadores técnicos e financeiros das suas faturas de energia em segundos.
                </p>
                <div className="flex flex-wrap gap-4 pt-4">
                  <div className="flex items-center gap-3 px-6 py-4 bg-white rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <div className="w-10 h-10 rounded-xl bg-green-50 flex items-center justify-center">
                      <CheckCircle2 className="w-5 h-5 text-green-600" />
                    </div>
                    <span className="text-sm font-bold text-sanesul-primary">Extração Precisa</span>
                  </div>
                  <div className="flex items-center gap-3 px-6 py-4 bg-white rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <div className="w-10 h-10 rounded-xl bg-blue-50 flex items-center justify-center">
                      <Zap className="w-5 h-5 text-sanesul-primary" />
                    </div>
                    <span className="text-sm font-bold text-sanesul-primary">Análise em Tempo Real</span>
                  </div>
                </div>
              </div>

              <div 
                className={`relative aspect-square lg:aspect-auto lg:h-[500px] border-2 border-dashed rounded-[40px] flex flex-col items-center justify-center transition-all group overflow-hidden ${
                  isDragging 
                    ? 'border-sanesul-primary bg-sanesul-primary/5 scale-[0.98]' 
                    : 'border-sanesul-primary/20 bg-white/50 hover:border-sanesul-primary/40 hover:bg-white'
                }`}
              >
                <div className="absolute inset-0 bg-gradient-to-br from-sanesul-primary/5 to-transparent opacity-0 group-hover:opacity-100 transition-opacity" />
                <div className="relative z-10 flex flex-col items-center">
                  <div className="w-24 h-24 bg-sanesul-primary rounded-3xl flex items-center justify-center mb-8 shadow-2xl shadow-sanesul-primary/30 group-hover:scale-110 transition-transform duration-500">
                    <Upload size={40} className="text-white" />
                  </div>
                  <p className="text-2xl font-display font-bold text-sanesul-primary mb-3">
                    {isDragging ? 'Solte agora' : 'Arraste suas faturas'}
                  </p>
                  <p className="text-sm text-sanesul-muted text-center max-w-[240px] mb-8">
                    Suporta PDF, JPG e PNG. Processamento automático via Gemini AI.
                  </p>
                  <div className="flex flex-col sm:flex-row gap-4">
                    <button 
                      onClick={(e) => { e.stopPropagation(); fileInputEnergisaRef.current?.click(); }}
                      className="px-8 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest shadow-lg shadow-sanesul-primary/20 hover:bg-sanesul-primary/90 transition-all"
                    >
                      Selecionar ENERGISA
                    </button>
                    <button 
                      onClick={(e) => { e.stopPropagation(); fileInputElektroRef.current?.click(); }}
                      className="px-8 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest shadow-lg shadow-sanesul-primary/20 hover:bg-sanesul-primary/90 transition-all"
                    >
                      Selecionar ELEKTRO
                    </button>
                  </div>

                  <div className="mt-8 flex flex-col sm:flex-row items-center gap-4 pt-8 border-t border-slate-200 w-full max-w-md">
                    <p className="text-sm font-bold text-slate-500 w-full text-center sm:text-left">Ou preencha manualmente:</p>
                    <button
                      onClick={(e) => { e.stopPropagation(); downloadExcelTemplate(); }}
                      className="flex items-center gap-2 px-4 py-2 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold uppercase tracking-wider hover:bg-emerald-100 transition-colors w-full sm:w-auto justify-center"
                    >
                      <Download size={16} />
                      Baixar Modelo
                    </button>
                    <label className="flex items-center gap-2 px-4 py-2 bg-blue-50 text-blue-700 rounded-lg text-xs font-bold uppercase tracking-wider hover:bg-blue-100 transition-colors cursor-pointer w-full sm:w-auto justify-center">
                      <FileSpreadsheet size={16} />
                      Importar Planilha
                      <input 
                        type="file" 
                        accept=".xlsx, .xls" 
                        className="hidden" 
                        onChange={handleExcelImport} 
                        onClick={(e) => e.stopPropagation()}
                      />
                    </label>
                  </div>
                </div>
                
                {/* Decorative elements */}
                <div className="absolute -bottom-12 -right-12 w-48 h-48 bg-sanesul-primary/5 rounded-full blur-3xl" />
                <div className="absolute -top-12 -left-12 w-48 h-48 bg-sanesul-secondary/5 rounded-full blur-3xl" />
              </div>
            </div>
          ) : (
            <div className={`space-y-8 transition-all ${isDragging ? 'opacity-50 scale-[0.99]' : ''}`}>
              {/* Stats Bar */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                {[
                  { label: filterReference === 'all' ? 'Total de Arquivos' : `Arquivos (${filterReference})`, value: sortedBills.length, color: 'sanesul-primary', icon: FileText },
                  { label: 'Aguardando', value: sortedBills.filter(b => b.status === 'pending').length, color: 'slate-500', icon: Clock },
                  { label: 'Em Processamento', value: sortedBills.filter(b => b.status === 'processing').length, color: 'sanesul-secondary', icon: Loader2 },
                  { label: 'Concluídos', value: sortedBills.filter(b => b.status === 'completed').length, color: 'green-600', icon: CheckCircle2 }
                ].map((stat, i) => (
                  <div key={i} className="bg-white p-6 rounded-3xl border border-sanesul-primary/5 shadow-sm flex items-center gap-4">
                    <div className={`w-12 h-12 rounded-2xl bg-${stat.color === 'sanesul-primary' ? 'sanesul-primary' : stat.color === 'sanesul-secondary' ? 'sanesul-secondary' : stat.color}/10 flex items-center justify-center`}>
                      <stat.icon size={20} className={`text-${stat.color === 'sanesul-primary' ? 'sanesul-primary' : stat.color === 'sanesul-secondary' ? 'sanesul-secondary' : stat.color} ${stat.label === 'Em Processamento' ? 'animate-spin' : ''}`} />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-sanesul-muted uppercase tracking-widest">{stat.label}</p>
                      <p className={`text-2xl font-display font-bold text-${stat.color === 'sanesul-primary' ? 'sanesul-primary' : stat.color === 'sanesul-secondary' ? 'sanesul-secondary' : stat.color}`}>{stat.value}</p>
                    </div>
                  </div>
                ))}
              </div>

              {/* Search Bar */}
              <div className="flex flex-col sm:flex-row gap-4 items-center justify-between bg-white p-4 rounded-3xl border border-sanesul-primary/5 shadow-sm">
                <div className="relative w-full sm:w-96">
                  <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none">
                    <Search size={16} className="text-sanesul-muted" />
                  </div>
                  <input
                    type="text"
                    placeholder="Buscar por Unidade Consumidora (UC)..."
                    value={searchUC}
                    onChange={(e) => setSearchUC(e.target.value)}
                    className="block w-full pl-10 pr-3 py-2 border border-sanesul-primary/10 rounded-xl text-xs focus:ring-sanesul-primary focus:border-sanesul-primary bg-slate-50/50"
                  />
                </div>
                {searchUC && (
                  <button 
                    onClick={() => setSearchUC('')}
                    className="text-[10px] font-bold text-red-500 uppercase tracking-widest hover:text-red-600 transition-colors"
                  >
                    Limpar Busca
                  </button>
                )}
              </div>

              {/* Table Container */}
              <div className="bg-white rounded-[32px] border border-sanesul-primary/10 shadow-xl overflow-hidden">
                <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
                  <table className="w-full text-left border-collapse relative">
                    <thead className="sticky top-0 z-20 bg-slate-50 shadow-sm">
                      <tr className="bg-slate-50/50">
                        <th className="px-4 py-3 w-12">
                          <input 
                            type="checkbox"
                            checked={selectedBills.length > 0 && selectedBills.length === bills.length}
                            onChange={() => setSelectedBills(selectedBills.length === bills.length ? [] : bills.map(b => b.id))}
                            className="rounded border-sanesul-primary/20 text-sanesul-primary focus:ring-sanesul-primary"
                          />
                        </th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary" onClick={() => requestSort('fileName')}>Arquivo</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary" onClick={() => requestSort('uc')}>UC</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary" onClick={() => requestSort('concessionaria')}>Concessionária</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary" onClick={() => requestSort('referencia')}>Referência</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary" onClick={() => requestSort('dataVencimento')}>Vencimento</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">Demanda Medida</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">Demanda Contratada</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary" onClick={() => requestSort('mercado')}>Mercado</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">Tipo</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">Status</th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 text-right">
                          {selectedBills.length > 0 ? (
                            <div className="flex items-center justify-end gap-4">
                              <button onClick={removeSelectedBills} className="text-red-600 hover:text-red-700 font-bold">Excluir ({selectedBills.length})</button>
                              {bills.length >= 223 && (
                                <button onClick={deselectFirst223} className="text-sanesul-primary hover:text-sanesul-secondary font-bold">Deselecionar 223</button>
                              )}
                            </div>
                          ) : 'Ações'}
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-sanesul-primary/5">
                      <AnimatePresence initial={false}>
                        {sortedBills.map((bill) => (
                          <motion.tr
                            key={bill.id}
                            initial={{ opacity: 0, y: 10 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, x: -20 }}
                            className={`hover:bg-sanesul-primary/5 transition-colors group ${selectedBills.includes(bill.id) ? 'bg-sanesul-primary/5' : ''}`}
                          >
                            <td className="px-4 py-3">
                              <input 
                                type="checkbox"
                                checked={selectedBills.includes(bill.id)}
                                onChange={() => toggleBillSelection(bill.id)}
                                className="rounded border-sanesul-primary/20 text-sanesul-primary focus:ring-sanesul-primary"
                              />
                            </td>
                            <td className="px-4 py-3">
                              <div className="flex items-center gap-4">
                                <div className="w-8 h-8 bg-sanesul-primary/5 rounded-lg flex items-center justify-center text-sanesul-primary group-hover:bg-sanesul-primary group-hover:text-white transition-all">
                                  <FileText size={16} />
                                </div>
                                <div className="flex flex-col">
                                  <span className="text-xs font-bold text-sanesul-primary truncate max-w-[200px]" title={bill.fileName}>
                                    {bill.fileName}
                                  </span>
                                  <span className="text-[9px] text-sanesul-muted uppercase tracking-wider">
                                    {bill.file ? `${(bill.file.size / 1024 / 1024).toFixed(2)} MB • ${bill.file.type.split('/')[1].toUpperCase()}` : 'ARQUIVO SALVO'}
                                  </span>
                                </div>
                              </div>
                            </td>
                            <td className="px-4 py-3">
                              <div className="flex flex-col">
                                <span className="text-xs font-mono font-bold text-sanesul-primary">{bill.uc || '---'}</span>
                              </div>
                            </td>
                            <td className="px-4 py-3">
                              <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                                {bill.concessionaria 
                                  ? (bill.concessionaria.toUpperCase().includes('ENERGISA') 
                                      ? 'ENERGISA' 
                                      : bill.concessionaria.toUpperCase().includes('ELEKTRO') 
                                        ? 'ELEKTRO' 
                                        : bill.concessionaria)
                                  : '---'}
                              </span>
                            </td>
                            <td className="px-4 py-3">
                              <span className="text-xs text-slate-600">
                                {bill.mesReferencia && bill.anoLeitura ? `${formatMonth(bill.mesReferencia)}/${bill.anoLeitura}` : '---'}
                              </span>
                            </td>
                            <td className="px-4 py-3">
                              <span className="text-xs text-slate-600">
                                {bill.dataVencimento || '---'}
                              </span>
                            </td>
                            <td className="px-4 py-3">
                              <div className="flex flex-col">
                                <span className="text-[10px] text-sanesul-muted uppercase font-bold tracking-tight">P: {bill.demandaPotenciaMedidaPonta || '0'} kW</span>
                                <span className="text-[10px] text-sanesul-muted uppercase font-bold tracking-tight">FP: {bill.demandaPotenciaMedidaForaPonta || '0'} kW</span>
                              </div>
                            </td>
                            <td className="px-4 py-3">
                              <div className="flex flex-col">
                                <span className="text-[9px] text-sanesul-muted uppercase font-bold tracking-tight">P: {bill.demandaPontaKW || '0'} kW</span>
                                <span className="text-[9px] text-sanesul-muted uppercase font-bold tracking-tight">FP: {bill.demandaForaPontaKW || '0'} kW</span>
                              </div>
                            </td>
                            <td className="px-4 py-3">
                              <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold uppercase tracking-wider ${
                                (bill.mercado || (bill.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(bill.uc)) ? 'LIVRE' : 'CATIVO')) === 'LIVRE' 
                                  ? 'bg-blue-100 text-blue-700' 
                                  : 'bg-slate-100 text-slate-600'
                              }`}>
                                {bill.mercado || (bill.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(bill.uc)) ? 'LIVRE' : 'CATIVO')}
                              </span>
                            </td>
                            <td className="px-4 py-3">
                              <span className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold uppercase tracking-wider ${
                                bill.tipo === 'OPERACIONAL' 
                                  ? 'bg-blue-50 text-blue-600' 
                                  : bill.tipo === 'ADMINISTRATIVO'
                                    ? 'bg-purple-50 text-purple-600'
                                    : 'bg-slate-100 text-slate-500'
                              }`}>
                                {bill.tipo || 'N/A'}
                              </span>
                            </td>
                            <td className="px-4 py-3">
                              <div className="flex items-center gap-2">
                                {bill.status === 'pending' && (
                                  <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-slate-100 text-slate-600 rounded-full text-[9px] font-bold uppercase tracking-wider">
                                    <Clock size={10} />
                                    Aguardando
                                  </span>
                                )}
                                {bill.status === 'processing' && (
                                  <div className="flex flex-col gap-1 w-full min-w-[100px]">
                                    <div className="flex items-center justify-between">
                                      <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-blue-50 text-sanesul-primary rounded-full text-[9px] font-bold uppercase tracking-wider">
                                        <Loader2 size={10} className="animate-spin" />
                                        Extraindo... {bill.progress || 0}%
                                      </span>
                                      <button 
                                        onClick={() => bill.abortController?.abort()}
                                        className="text-red-500 hover:text-red-700 p-0.5 rounded-full hover:bg-red-50 transition-colors"
                                        title="Cancelar"
                                      >
                                        <X size={12} />
                                      </button>
                                    </div>
                                    <div className="w-full bg-slate-100 rounded-full h-1 overflow-hidden">
                                      <div 
                                        className="bg-sanesul-primary h-1 rounded-full transition-all duration-300 ease-out" 
                                        style={{ width: `${bill.progress || 0}%` }}
                                      ></div>
                                    </div>
                                    {bill.error && (
                                      <span className="text-[9px] text-amber-600 font-bold flex items-center gap-1 mt-1 bg-amber-50 px-2 py-0.5 rounded-full w-fit">
                                        <AlertCircle size={10} /> {bill.error}
                                      </span>
                                    )}
                                  </div>
                                )}
                                {bill.status === 'completed' && (
                                  <div className="flex flex-col gap-1">
                                    <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-green-50 text-green-600 rounded-full text-[10px] font-bold uppercase tracking-wider w-fit">
                                      <CheckCircle2 size={12} />
                                      Concluído
                                    </span>
                                    {bill.error && (
                                      <span className="text-[9px] text-amber-600 font-bold flex items-center gap-1 mt-1 bg-amber-50 px-2 py-0.5 rounded-full w-fit">
                                        <AlertCircle size={10} /> {bill.error}
                                      </span>
                                    )}
                                  </div>
                                )}
                                {bill.status === 'error' && (
                                  <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-red-50 text-red-600 rounded-full text-[10px] font-bold uppercase tracking-wider" title={bill.error}>
                                    <AlertCircle size={12} />
                                    {bill.error || 'Erro'}
                                  </span>
                                )}
                              </div>
                            </td>
                            <td className="px-8 py-5 text-right">
                              <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                                <button
                                  onClick={() => {
                                    setEditingBill(bill);
                                    setIsBillModalOpen(true);
                                  }}
                                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-slate-400 hover:bg-blue-50 hover:text-blue-500 transition-all"
                                  disabled={isProcessing && bill.status === 'processing'}
                                  title="Editar"
                                >
                                  <Pencil size={16} />
                                </button>
                                <button
                                  onClick={() => removeBill(bill.id)}
                                  className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-slate-400 hover:bg-red-50 hover:text-red-500 transition-all"
                                  disabled={isProcessing && bill.status === 'processing'}
                                  title="Excluir"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </div>
                            </td>
                          </motion.tr>
                        ))}
                      </AnimatePresence>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )
        ) : activeTab === 'multas' ? (
          <div className="space-y-8">
            <div className="bg-white p-10 rounded-3xl border border-sanesul-primary/10 shadow-xl">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-10">
                <div>
                  <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">Análise de Multas</h2>
                  <p className="text-sanesul-muted">Acompanhe as multas por ultrapassagem, reativa e subutilização.</p>
                </div>
              </div>

              {/* Chart Section */}
              <div className="bg-slate-50 rounded-3xl p-8 mb-10 border border-sanesul-primary/5">
                <h3 className="text-lg font-bold text-sanesul-primary mb-6">
                  {selectedMultaType === 'total' ? 'Evolução Mensal - Gasto Total' :
                   selectedMultaType === 'ultrapassagem' ? 'Evolução Mensal - Multa de Ultrapassagem' :
                   selectedMultaType === 'reativa' ? 'Evolução Mensal - Multa Reativa' :
                   'Evolução Mensal - Subutilização'}
                </h3>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart data={multasMonthlyData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                      <defs>
                        <linearGradient id="colorTotal" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#8b5cf6" stopOpacity={0.3}/>
                          <stop offset="95%" stopColor="#8b5cf6" stopOpacity={0}/>
                        </linearGradient>
                        <linearGradient id="colorUltrapassagem" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#ef4444" stopOpacity={0.3}/>
                          <stop offset="95%" stopColor="#ef4444" stopOpacity={0}/>
                        </linearGradient>
                        <linearGradient id="colorReativa" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.3}/>
                          <stop offset="95%" stopColor="#f59e0b" stopOpacity={0}/>
                        </linearGradient>
                        <linearGradient id="colorSubutilizacao" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.3}/>
                          <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                      <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} dy={10} />
                      <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 12, fill: '#64748b' }} tickFormatter={(value) => `R$ ${value.toLocaleString('pt-BR')}`} />
                      <Tooltip 
                        cursor={{ stroke: '#cbd5e1', strokeWidth: 1, strokeDasharray: '3 3' }}
                        contentStyle={{ borderRadius: '16px', border: 'none', boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)' }}
                        formatter={(value: number) => [`R$ ${value.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}`, 'Valor']}
                      />
                      <Area 
                        type="monotone"
                        dataKey={selectedMultaType} 
                        stroke={selectedMultaType === 'total' ? '#8b5cf6' : selectedMultaType === 'ultrapassagem' ? '#ef4444' : selectedMultaType === 'reativa' ? '#f59e0b' : '#3b82f6'} 
                        fillOpacity={1}
                        fill={`url(#color${selectedMultaType === 'total' ? 'Total' : selectedMultaType === 'ultrapassagem' ? 'Ultrapassagem' : selectedMultaType === 'reativa' ? 'Reativa' : 'Subutilizacao'})`}
                        strokeWidth={4}
                        dot={{ r: 4, strokeWidth: 2 }}
                        activeDot={{ r: 6, strokeWidth: 0 }}
                      />
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Cards Section */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6 mb-10">
                <div 
                  onClick={() => setSelectedMultaType('total')}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === 'total' 
                      ? 'bg-violet-50 border-violet-200 shadow-xl shadow-violet-100' 
                      : 'bg-white border-slate-100 hover:border-violet-100 hover:bg-violet-50/50'
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div className={`p-3 rounded-xl ${selectedMultaType === 'total' ? 'bg-violet-100 text-violet-600' : 'bg-slate-50 text-slate-400'}`}>
                      <Calculator size={24} />
                    </div>
                    <h3 className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === 'total' ? 'text-violet-900' : 'text-slate-500'}`}>
                      Gasto Total
                    </h3>
                  </div>
                  <p className={`text-3xl font-display font-bold ${selectedMultaType === 'total' ? 'text-violet-600' : 'text-slate-700'}`}>
                    R$ {multasTotals.total.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                  </p>
                </div>
                <div 
                  onClick={() => setSelectedMultaType('reativa')}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === 'reativa' 
                      ? 'bg-amber-50 border-amber-200 shadow-xl shadow-amber-100' 
                      : 'bg-white border-slate-100 hover:border-amber-100 hover:bg-amber-50/50'
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div className={`p-3 rounded-xl ${selectedMultaType === 'reativa' ? 'bg-amber-100 text-amber-600' : 'bg-slate-50 text-slate-400'}`}>
                      <Zap size={24} />
                    </div>
                    <h3 className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === 'reativa' ? 'text-amber-900' : 'text-slate-500'}`}>
                      Multa Reativa
                    </h3>
                  </div>
                  <p className={`text-3xl font-display font-bold ${selectedMultaType === 'reativa' ? 'text-amber-600' : 'text-slate-700'}`}>
                    R$ {multasTotals.reativa.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                  </p>
                </div>

                <div 
                  onClick={() => setSelectedMultaType('ultrapassagem')}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === 'ultrapassagem' 
                      ? 'bg-red-50 border-red-200 shadow-xl shadow-red-100' 
                      : 'bg-white border-slate-100 hover:border-red-100 hover:bg-red-50/50'
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div className={`p-3 rounded-xl ${selectedMultaType === 'ultrapassagem' ? 'bg-red-100 text-red-600' : 'bg-slate-50 text-slate-400'}`}>
                      <TrendingUp size={24} />
                    </div>
                    <h3 className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === 'ultrapassagem' ? 'text-red-900' : 'text-slate-500'}`}>
                      Ultrapassagem de Demanda
                    </h3>
                  </div>
                  <p className={`text-3xl font-display font-bold ${selectedMultaType === 'ultrapassagem' ? 'text-red-600' : 'text-slate-700'}`}>
                    R$ {multasTotals.ultrapassagem.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                  </p>
                </div>

                <div 
                  onClick={() => setSelectedMultaType('subutilizacao')}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === 'subutilizacao' 
                      ? 'bg-blue-50 border-blue-200 shadow-xl shadow-blue-100' 
                      : 'bg-white border-slate-100 hover:border-blue-100 hover:bg-blue-50/50'
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div className={`p-3 rounded-xl ${selectedMultaType === 'subutilizacao' ? 'bg-blue-100 text-blue-600' : 'bg-slate-50 text-slate-400'}`}>
                      <TrendingDown size={24} />
                    </div>
                    <h3 className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === 'subutilizacao' ? 'text-blue-900' : 'text-slate-500'}`}>
                      Subutilização
                    </h3>
                  </div>
                  <p className={`text-3xl font-display font-bold ${selectedMultaType === 'subutilizacao' ? 'text-blue-600' : 'text-slate-700'}`}>
                    R$ {multasTotals.subutilizacao.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                  </p>
                </div>
              </div>

              {/* UC List Section */}
              <div className="bg-white rounded-3xl border border-sanesul-primary/10 overflow-hidden">
                <div className="p-6 border-b border-sanesul-primary/10 bg-slate-50/50 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                  <div>
                    <h3 className="text-lg font-bold text-sanesul-primary">
                      Detalhamento por Unidade Consumidora (UC)
                    </h3>
                    <p className="text-sm text-sanesul-muted mt-1">
                      {selectedMultaType === 'total' ? 'UCs com gastos totais (ultrapassagem, reativa ou subutilização) no período.' :
                       selectedMultaType === 'ultrapassagem' ? 'UCs com multas de ultrapassagem no período.' :
                       selectedMultaType === 'reativa' ? 'UCs com multas reativas no período.' :
                       'UCs com subutilização no período.'}
                    </p>
                  </div>
                  <div className="flex items-center gap-4">
                    <select
                      value={multasMonth}
                      onChange={(e) => setMultasMonth(e.target.value)}
                      className="px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-bold text-sanesul-primary focus:ring-2 focus:ring-sanesul-primary/20"
                    >
                      <option value="all">Todos os Meses</option>
                      {Array.from(new Set(bills.filter(b => b.status === 'completed').map(b => `${formatMonth(b.mesReferencia)}/${b.anoLeitura}`))).map(month => (
                        <option key={month} value={month}>{month}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50/80 border-b border-sanesul-primary/10">
                        <th className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted">UC</th>
                        <th className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted">Cidade</th>
                        <th 
                          className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted text-right cursor-pointer hover:text-sanesul-primary transition-colors group"
                          onClick={() => setMultasSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')}
                        >
                          <div className="flex items-center justify-end gap-1">
                            Valor Total (R$)
                            <div className="flex flex-col">
                              <ArrowUp size={10} className={`${multasSortDirection === 'asc' ? 'text-sanesul-primary' : 'text-slate-300 group-hover:text-slate-400'}`} />
                              <ArrowDown size={10} className={`-mt-1 ${multasSortDirection === 'desc' ? 'text-sanesul-primary' : 'text-slate-300 group-hover:text-slate-400'}`} />
                            </div>
                          </div>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {multasUcList.length > 0 ? (
                        multasUcList.map((item, index) => (
                          <tr key={item.uc} className={`border-b border-slate-100 hover:bg-slate-50 transition-colors ${index % 2 === 0 ? 'bg-white' : 'bg-slate-50/30'}`}>
                            <td className="p-4 font-medium text-sanesul-primary">{item.uc}</td>
                            <td className="p-4 text-slate-600">{item.cidade}</td>
                            <td className="p-4 text-right font-bold text-sanesul-primary">
                              R$ {item[selectedMultaType].toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={2} className="p-8 text-center text-sanesul-muted">
                            Nenhuma UC com este tipo de multa no período.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        ) : activeTab === 'dashboard' ? (
          <div className="flex flex-col md:flex-row gap-8">
            {/* Dashboard Sidebar Navigation */}
            <aside className="w-full md:w-72 space-y-6">
              <div className="flex flex-col gap-2 p-3 bg-white rounded-3xl border border-sanesul-primary/10 shadow-xl">
                <button 
                  onClick={() => setDashboardSubTab('operacionais')}
                  className={`flex items-center gap-3 px-5 py-4 rounded-2xl text-xs font-bold uppercase tracking-wider transition-all ${
                    dashboardSubTab === 'operacionais' 
                      ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                      : 'text-sanesul-muted hover:bg-sanesul-primary/5 hover:text-sanesul-primary'
                  }`}
                >
                  <Zap size={16} />
                  Operacionais
                </button>
                
                {dashboardSubTab === 'operacionais' && (
                  <div className="ml-4 flex flex-col gap-1 border-l-2 border-sanesul-primary/10 pl-4 py-2">
                    {[
                      { id: 'consumo', label: 'Consumo de Energia' },
                      { id: 'ultrapassagem', label: 'Ultrapassagem de Demanda' },
                      { id: 'subutilizacao', label: 'Subutilização de Demanda' },
                      { id: 'reativa', label: 'Energia Reativa' },
                      { id: 'solar', label: 'Energia Solar' }
                    ].map(item => (
                      <button
                        key={item.id}
                        onClick={() => setOperationalSubTab(item.id as any)}
                        className={`text-left px-3 py-2.5 text-[11px] font-bold uppercase tracking-tight transition-all rounded-xl ${
                          operationalSubTab === item.id 
                            ? 'text-sanesul-primary bg-sanesul-primary/5' 
                            : 'text-sanesul-muted hover:text-sanesul-primary'
                        }`}
                      >
                        {item.label}
                      </button>
                    ))}
                  </div>
                )}
 
                <button 
                  onClick={() => setDashboardSubTab('financeiro')}
                  className={`flex items-center gap-3 px-5 py-4 rounded-2xl text-xs font-bold uppercase tracking-wider transition-all ${
                    dashboardSubTab === 'financeiro' 
                      ? 'bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20' 
                      : 'text-sanesul-muted hover:bg-sanesul-primary/5 hover:text-sanesul-primary'
                  }`}
                >
                  <DollarSign size={16} />
                  Financeiro
                </button>
 
                {dashboardSubTab === 'financeiro' && (
                  <div className="ml-4 flex flex-col gap-1 border-l-2 border-sanesul-primary/10 pl-4 py-2">
                    {[
                      { id: 'despesas', label: 'Despesas com Energia' },
                      { id: 'multa_ultrapassagem', label: 'Multa de Ultrapassagem' },
                      { id: 'multa_reativa', label: 'Multa de Energia Reativa' },
                      { id: 'tarifa_media', label: 'Tarifa Média' },
                      { id: 'energia_solar', label: 'Energia Solar' }
                    ].map(item => (
                      <button
                        key={item.id}
                        onClick={() => setFinancialSubTab(item.id as any)}
                        className={`text-left px-3 py-2.5 text-[11px] font-bold uppercase tracking-tight transition-all rounded-xl ${
                          financialSubTab === item.id 
                            ? 'text-sanesul-primary bg-sanesul-primary/5' 
                            : 'text-sanesul-muted hover:text-sanesul-primary'
                        }`}
                      >
                        {item.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
              
              <div className="p-6 bg-sanesul-primary rounded-3xl shadow-xl shadow-sanesul-primary/20 text-white hidden md:block">
                <h3 className="text-[10px] font-bold uppercase tracking-[0.2em] opacity-60 mb-4">Resumo Geral</h3>
                <div className="space-y-4">
                  <div className="flex justify-between items-end">
                    <span className="text-[10px] font-bold uppercase opacity-60">Processados</span>
                    <span className="text-2xl font-display font-bold leading-none">{generalFilteredData.length}</span>
                  </div>
                  <div className="flex justify-between items-end">
                    <span className="text-[10px] font-bold uppercase opacity-60">Unidades</span>
                    <span className="text-2xl font-display font-bold leading-none">{new Set(generalFilteredData.map(d => d.uc)).size}</span>
                  </div>
                  <div className="pt-4 border-t border-white/10 space-y-4">
                    <div className="flex flex-col">
                      <span className="text-[10px] font-bold uppercase opacity-60 mb-1">Custo Total</span>
                      <span className="text-xl font-display font-bold">R$ {generalFilteredData.reduce((acc, curr) => acc + curr.valorTotal, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                    </div>
                    <div className="flex flex-col">
                      <span className="text-[10px] font-bold uppercase opacity-60 mb-1">Consumo Total</span>
                      <span className="text-xl font-display font-bold">{generalFilteredData.reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0).toLocaleString('pt-BR')} <span className="text-xs opacity-60">kWh</span></span>
                    </div>
                  </div>
                </div>
              </div>
            </aside>

            <div className="flex-1 space-y-10">
              <div className="flex flex-col md:flex-row md:items-center justify-between border-b border-sanesul-primary/10 pb-6 gap-6">
                <div className="flex flex-col">
                  <h2 className="text-3xl font-display font-bold text-sanesul-primary">
                    {dashboardSubTab === 'financeiro' ? (
                      financialSubTab === 'despesas' ? 'Despesas com Energia' :
                      financialSubTab === 'multa_ultrapassagem' ? 'Multa de Ultrapassagem' :
                      financialSubTab === 'multa_reativa' ? 'Multa de Energia Reativa' : 
                      financialSubTab === 'energia_solar' ? 'Créditos de Energia Solar' : 'Tarifa Média'
                    ) : (
                      operationalSubTab === 'consumo' ? 'Consumo de Energia' :
                      operationalSubTab === 'ultrapassagem' ? 'Ultrapassagem de Demanda' :
                      operationalSubTab === 'subutilizacao' ? 'Subutilização de Demanda' :
                      operationalSubTab === 'reativa' ? 'Energia Reativa' : 'Energia Solar'
                    )}
                  </h2>
                  <p className="text-[11px] font-bold uppercase tracking-wider text-sanesul-muted mt-1">
                    {!selectedUC || selectedUC === 'all' ? 'Visão consolidada do grupo' : `Unidade Consumidora: ${selectedUC}`}
                  </p>
                </div>
                
                <div className="flex flex-wrap items-center gap-3">
                  <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-sanesul-muted ml-2">Concessionária:</span>
                    <select 
                      value={selectedConcessionaria}
                      onChange={(e) => setSelectedConcessionaria(e.target.value)}
                      className="bg-sanesul-bg border-none px-4 py-2 rounded-xl text-[11px] font-bold uppercase tracking-wider text-sanesul-primary outline-none focus:ring-2 focus:ring-sanesul-primary/20 transition-all cursor-pointer"
                    >
                      <option value="all">Todas</option>
                      <option value="ENERGISA">Energisa</option>
                      <option value="ELEKTRO">Elektro</option>
                    </select>
                  </div>
                  <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-sanesul-muted ml-2">Mês:</span>
                    <select 
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(e.target.value)}
                      className="bg-sanesul-bg border-none px-4 py-2 rounded-xl text-[11px] font-bold uppercase tracking-wider text-sanesul-primary outline-none focus:ring-2 focus:ring-sanesul-primary/20 transition-all cursor-pointer"
                    >
                      <option value="all">Todos os Meses</option>
                      {availableMonths.map(month => (
                        <option key={month} value={month}>{month}</option>
                      ))}
                    </select>
                  </div>
                  <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-sanesul-muted ml-2">Filtrar UC:</span>
                    <div className="relative flex items-center">
                      <Search size={14} className="absolute left-3 text-sanesul-primary/40" />
                      <input 
                        type="text"
                        value={selectedUC === 'all' ? '' : selectedUC}
                        onChange={(e) => setSelectedUC(e.target.value)}
                        placeholder="Buscar UC..."
                        className="bg-sanesul-bg border-none pl-9 pr-4 py-2 rounded-xl text-[11px] font-bold uppercase tracking-wider text-sanesul-primary outline-none focus:ring-2 focus:ring-sanesul-primary/20 transition-all w-48"
                      />
                    </div>
                  </div>
                </div>
              </div>

              {completedBills.length === 0 ? (
                <div className="p-24 text-center bg-white rounded-3xl border border-sanesul-primary/10 shadow-xl">
                  <div className="w-16 h-16 bg-sanesul-primary/5 rounded-full flex items-center justify-center mx-auto mb-6">
                    <AlertCircle size={32} className="text-sanesul-primary/30" />
                  </div>
                  <p className="text-xl font-display font-semibold text-sanesul-primary">Nenhum dado disponível</p>
                  <p className="text-sanesul-muted mt-2">Processe algumas faturas para visualizar o dashboard analítico</p>
                </div>
              ) : (
                <div className="space-y-10">
                  {/* Summary Cards */}
                  <div className={`grid grid-cols-1 gap-8 ${
                    (dashboardSubTab === 'financeiro' && financialSubTab === 'energia_solar') || (dashboardSubTab === 'operacionais' && operationalSubTab === 'consumo')
                      ? 'md:grid-cols-2 lg:grid-cols-3' 
                      : 'md:grid-cols-2 lg:grid-cols-3'
                  }`}>
                    {dashboardSubTab === 'operacionais' ? (
                      <>
                        {operationalSubTab === 'consumo' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total de Faturas</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">{filteredDashboardData.length} <span className="text-base font-sans font-medium opacity-40">Arquivos</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Consumo Total</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">{Math.round(filteredDashboardData.reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0)).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kWh</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Grupo A</p>
                              <div className="space-y-2">
                                <p className="text-lg font-bold text-green-600 flex justify-between">
                                  <span>Verde:</span> 
                                  <span>{Math.round(filteredDashboardData.filter(d => d.modalidadeTarifaria.includes('VERDE')).reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0)).toLocaleString('pt-BR')} kWh</span>
                                </p>
                                <p className="text-lg font-bold text-sanesul-primary flex justify-between">
                                  <span>Azul:</span> 
                                  <span>{Math.round(filteredDashboardData.filter(d => d.modalidadeTarifaria.includes('AZUL')).reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0)).toLocaleString('pt-BR')} kWh</span>
                                </p>
                              </div>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Grupo B</p>
                              <div className="space-y-2">
                                <p className="text-lg font-bold text-sanesul-primary flex justify-between">
                                  <span>Solar:</span> 
                                  <span>{Math.round(filteredDashboardData.filter(d => (d.solarInjetadaOUC > 0 || d.solarInjetadaMUC > 0)).reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0)).toLocaleString('pt-BR')} kWh</span>
                                </p>
                                <p className="text-lg font-bold text-sanesul-primary flex justify-between">
                                  <span>Não Solar:</span> 
                                  <span>{Math.round(filteredDashboardData.filter(d => !(d.solarInjetadaOUC > 0 || d.solarInjetadaMUC > 0) && (d.consumoPonta + d.consumoForaPonta > 0)).reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0)).toLocaleString('pt-BR')} kWh</span>
                                </p>
                              </div>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Tarifa Branca</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {filteredDashboardData.filter(d => d.modalidadeTarifaria.includes('BRANCA')).reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kWh</span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Optante B</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {filteredDashboardData.filter(d => d.subgrupo.startsWith('B') && d.demandaContratadaPonta > 0).length} <span className="text-base font-sans font-medium opacity-40">Faturas</span>
                              </p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === 'ultrapassagem' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <AlertCircle size={80} className="text-red-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Ultrapassagem</p>
                              <p className="text-4xl font-display font-bold text-red-600">{filteredDashboardData.reduce((acc, curr) => acc + curr.ultrapassagemPonta + curr.ultrapassagemForaPonta, 0).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kW</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Ocorrências</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">{new Set(filteredDashboardData.filter(d => d.ultrapassagemPonta > 0 || d.ultrapassagemForaPonta > 0).map(d => d.uc)).size} <span className="text-base font-sans font-medium opacity-40">Unidades</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-red-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Maior Desvio</p>
                              <p className="text-4xl font-display font-bold text-red-600">{Math.max(...filteredDashboardData.map(d => d.ultrapassagemPonta + d.ultrapassagemForaPonta), 0).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kW</span></p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === 'subutilizacao' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Média Utilização</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {((filteredDashboardData.reduce((acc, curr) => acc + (curr.demandaMedidaPonta / (curr.demandaContratadaPonta || 1)), 0) / filteredDashboardData.length || 0) * 100).toFixed(1)}%
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <AlertCircle size={80} className="text-orange-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Subutilizados (&lt;80%)</p>
                              <p className="text-4xl font-display font-bold text-orange-600">{new Set(filteredDashboardData.filter(d => d.demandaMedidaPonta < (d.demandaContratadaPonta * 0.8)).map(d => d.uc)).size} <span className="text-base font-sans font-medium opacity-40">Unidades</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-orange-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Mínima Utilização</p>
                              <p className="text-4xl font-display font-bold text-orange-600">{Math.min(...filteredDashboardData.map(d => (d.demandaMedidaPonta / (d.demandaContratadaPonta || 1)) * 100), 100).toFixed(1)}%</p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === 'reativa' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Excedente</p>
                              <p className="text-4xl font-display font-bold text-purple-600">{filteredDashboardData.reduce((acc, curr) => acc + curr.reativaPonta + curr.reativaForaPonta, 0).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kVArh</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Unidades com Excesso</p>
                              <p className="text-4xl font-display font-bold text-purple-600">{new Set(filteredDashboardData.filter(d => d.reativaPonta > 0 || d.reativaForaPonta > 0).map(d => d.uc)).size} <span className="text-base font-sans font-medium opacity-40">Unidades</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Média Excedente</p>
                              <p className="text-4xl font-display font-bold text-purple-600">{(filteredDashboardData.reduce((acc, curr) => acc + curr.reativaPonta + curr.reativaForaPonta, 0) / timeSeriesData.length || 0).toFixed(1)} <span className="text-base font-sans font-medium opacity-40">kVArh</span></p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === 'solar' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-sanesul-secondary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Consumo em kWh</p>
                              <p className="text-4xl font-display font-bold text-sanesul-secondary">{filteredDashboardData.reduce((acc, curr) => acc + (curr.consumoPonta + curr.consumoForaPonta), 0).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kWh</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-green-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Injetada</p>
                              <p className="text-4xl font-display font-bold text-green-600">{filteredDashboardData.reduce((acc, curr) => acc + (curr.solarInjetadaOUC + curr.solarInjetadaMUC), 0).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kWh</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-green-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Saldo Energia</p>
                              <p className="text-4xl font-display font-bold text-green-600">{(filteredDashboardData.reduce((acc, curr) => acc + (curr.solarInjetadaOUC + curr.solarInjetadaMUC) - (curr.consumoPonta + curr.consumoForaPonta), 0)).toLocaleString('pt-BR')} <span className="text-base font-sans font-medium opacity-40">kWh</span></p>
                            </div>
                          </>
                        )}
                      </>
                    ) : (
                      <>
                        {financialSubTab === 'despesas' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total de Faturas</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">{filteredDashboardData.length} <span className="text-base font-sans font-medium opacity-40">Arquivos</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Gasto Acumulado</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">R$ {filteredDashboardData.reduce((acc, curr) => acc + curr.valorTotal, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Média Mensal</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">R$ {(filteredDashboardData.reduce((acc, curr) => acc + curr.valorTotal, 0) / timeSeriesData.length || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <LayoutDashboard size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Unidades Ativas</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">{filteredUcs.length} <span className="text-base font-sans font-medium opacity-40">UCs</span></p>
                            </div>
                          </>
                        )}
                        {financialSubTab === 'multa_ultrapassagem' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <AlertCircle size={80} className="text-red-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Multas</p>
                              <p className="text-4xl font-display font-bold text-red-600">R$ {filteredDashboardData.reduce((acc, curr) => acc + curr.valorUltrapassagemPonta + curr.valorUltrapassagemForaPonta, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-red-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Média Mensal</p>
                              <p className="text-4xl font-display font-bold text-red-600">R$ {(filteredDashboardData.reduce((acc, curr) => acc + curr.valorUltrapassagemPonta + curr.valorUltrapassagemForaPonta, 0) / timeSeriesData.length || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign size={80} className="text-red-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Maior Penalidade</p>
                              <p className="text-4xl font-display font-bold text-red-600">R$ {Math.max(...filteredDashboardData.map(d => d.valorUltrapassagemPonta + d.valorUltrapassagemForaPonta), 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                          </>
                        )}
                        {financialSubTab === 'multa_reativa' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Multas</p>
                              <p className="text-4xl font-display font-bold text-purple-600">R$ {filteredDashboardData.reduce((acc, curr) => acc + curr.valorReativaPonta + curr.valorReativaForaPonta, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Média Mensal</p>
                              <p className="text-4xl font-display font-bold text-purple-600">R$ {(filteredDashboardData.reduce((acc, curr) => acc + curr.valorReativaPonta + curr.valorReativaForaPonta, 0) / timeSeriesData.length || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Maior Penalidade</p>
                              <p className="text-4xl font-display font-bold text-purple-600">R$ {Math.max(...filteredDashboardData.map(d => d.valorReativaPonta + d.valorReativaForaPonta), 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                          </>
                        )}
                        {financialSubTab === 'tarifa_media' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Tarifa Média</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">R$ {(filteredDashboardData.reduce((acc, curr) => acc + curr.valorTotal, 0) / (filteredDashboardData.reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0) || 1)).toLocaleString('pt-BR', { minimumFractionDigits: 3 })} <span className="text-base font-sans font-medium opacity-40">/kWh</span></p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-green-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Menor Tarifa</p>
                              <p className="text-4xl font-display font-bold text-green-600">R$ {Math.min(...filteredDashboardData.map(d => d.valorTotal / (d.consumoPonta + d.consumoForaPonta || 1)), 100).toLocaleString('pt-BR', { minimumFractionDigits: 3 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-red-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Maior Tarifa</p>
                              <p className="text-4xl font-display font-bold text-red-600">R$ {Math.max(...filteredDashboardData.map(d => d.valorTotal / (d.consumoPonta + d.consumoForaPonta || 1)), 0).toLocaleString('pt-BR', { minimumFractionDigits: 3 })}</p>
                            </div>
                          </>
                        )}
                        {financialSubTab === 'energia_solar' && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-sanesul-primary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Consumo (R$)</p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">R$ {filteredDashboardData.reduce((acc, curr) => acc + Math.abs(curr.valorSolarOUC + curr.valorSolarMUC) + (curr.valorTotal - curr.cip - curr.outrosEncargos), 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp size={80} className="text-green-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Créditos</p>
                              <p className="text-4xl font-display font-bold text-green-600">R$ {Math.abs(filteredDashboardData.reduce((acc, curr) => acc + curr.valorSolarOUC + curr.valorSolarMUC, 0)).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign size={80} className="text-sanesul-secondary" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Valor Total da Fatura</p>
                              <p className="text-4xl font-display font-bold text-sanesul-secondary">R$ {filteredDashboardData.reduce((acc, curr) => acc + curr.valorTotal, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                            </div>
                          </>
                        )}
                      </>
                    )}
                  </div>

                  {/* Charts removed as per user request */}

                  <div className="mt-12">
                    <div className="bg-white rounded-3xl border border-sanesul-primary/10 shadow-xl overflow-hidden">
                      <div className="p-8 border-b border-sanesul-primary/5 bg-slate-50/50 flex justify-between items-center">
                        <div>
                          <h3 className="text-lg font-display font-bold text-sanesul-primary">Detalhamento por Unidade Consumidora</h3>
                          <p className="text-sm text-sanesul-muted mt-1">Visão granular dos indicadores para cada registro no período selecionado.</p>
                        </div>
                        <div className="px-4 py-2 bg-sanesul-primary/10 rounded-full">
                          <span className="text-xs font-bold text-sanesul-primary uppercase tracking-widest">{filteredDashboardData.length} Registros</span>
                        </div>
                      </div>
                      <div className="overflow-auto max-h-[600px]">
                        <table className="w-full text-left border-collapse">
                          <thead className="sticky top-0 z-10 bg-slate-50">
                            <tr className="bg-slate-50/50">
                              <th 
                                onClick={() => setDashboardSort(prev => ({ key: 'uc', direction: prev.key === 'uc' && prev.direction === 'desc' ? 'asc' : 'desc' }))}
                                className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:bg-sanesul-primary/5 transition-colors"
                              >
                                UC {dashboardSort.key === 'uc' && (dashboardSort.direction === 'asc' ? '↑' : '↓')}
                              </th>
                              <th 
                                onClick={() => setDashboardSort(prev => ({ key: 'name', direction: prev.key === 'name' && prev.direction === 'desc' ? 'asc' : 'desc' }))}
                                className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:bg-sanesul-primary/5 transition-colors"
                              >
                                Mês/Ano {dashboardSort.key === 'name' && (dashboardSort.direction === 'asc' ? '↑' : '↓')}
                              </th>
                              <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Classificação</th>
                              <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Mercado</th>
                              <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Modalidade</th>
                              {dashboardSubTab === 'operacionais' ? (
                                <>
                                  {operationalSubTab === 'consumo' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Consumo Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Consumo F. Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Total (kWh)</th>
                                    </>
                                  )}
                                  {operationalSubTab === 'ultrapassagem' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Contratada Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Contratada F. Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Ultrap. Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Ultrap. F. Ponta</th>
                                      <th 
                                        onClick={() => setDashboardSort(prev => ({ key: 'total_kw', direction: prev.key === 'total_kw' && prev.direction === 'desc' ? 'asc' : 'desc' }))}
                                        className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-red-600 border-b border-sanesul-primary/5 text-right cursor-pointer hover:bg-red-50 transition-colors"
                                      >
                                        Total (kW) {dashboardSort.key === 'total_kw' && (dashboardSort.direction === 'asc' ? '↑' : '↓')}
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === 'subutilizacao' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Contratada Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Contratada F. Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Medida Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Medida F. Ponta</th>
                                      <th 
                                        onClick={() => setDashboardSort(prev => ({ key: 'utilizacao', direction: prev.key === 'utilizacao' && prev.direction === 'desc' ? 'asc' : 'desc' }))}
                                        className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-orange-600 border-b border-sanesul-primary/5 text-right cursor-pointer hover:bg-orange-50 transition-colors"
                                      >
                                        Utilização (%) {dashboardSort.key === 'utilizacao' && (dashboardSort.direction === 'asc' ? '↑' : '↓')}
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === 'reativa' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Reativa Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Reativa F. Ponta</th>
                                      <th 
                                        onClick={() => setDashboardSort(prev => ({ key: 'total_kvarh', direction: prev.key === 'total_kvarh' && prev.direction === 'desc' ? 'asc' : 'desc' }))}
                                        className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-purple-600 border-b border-sanesul-primary/5 text-right cursor-pointer hover:bg-purple-50 transition-colors"
                                      >
                                        Total (kVArh) {dashboardSort.key === 'total_kvarh' && (dashboardSort.direction === 'asc' ? '↑' : '↓')}
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === 'solar' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Consumo em kWh</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Injetada oUC</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Injetada mUC</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">Saldo Energia</th>
                                    </>
                                  )}
                                </>
                              ) : (
                                <>
                                  {financialSubTab === 'despesas' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Valor Total</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Cidade</th>
                                    </>
                                  )}
                                  {financialSubTab === 'multa_ultrapassagem' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Multa Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Multa F. Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-red-600 border-b border-sanesul-primary/5 text-right">Total (R$)</th>
                                    </>
                                  )}
                                  {financialSubTab === 'multa_reativa' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Multa Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Multa F. Ponta</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-purple-600 border-b border-sanesul-primary/5 text-right">Total (R$)</th>
                                    </>
                                  )}
                                  {financialSubTab === 'tarifa_media' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Valor Total</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">Consumo Total</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Tarifa (R$/kWh)</th>
                                    </>
                                  )}
                                  {financialSubTab === 'energia_solar' && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Consumo (R$)</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">Total Créditos</th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-secondary border-b border-sanesul-primary/5 text-right">Valor Total da Fatura</th>
                                    </>
                                  )}
                                </>
                              )}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-sanesul-primary/5">
                            {sortedDashboardData.map((row, idx) => (
                              <tr key={idx} className="hover:bg-sanesul-primary/5 transition-colors group">
                                <td className="px-8 py-5 text-sm font-bold text-sanesul-primary">{row.uc}</td>
                                <td className="px-8 py-5 text-sm text-slate-600">{row.name}</td>
                                <td className="px-8 py-5 text-sm font-medium text-slate-600">
                                  {UCS_PPP.has(String(row.uc)) ? 'PPP Fotovoltaica' : (UCS_USINA.has(String(row.uc)) ? 'Usinas SANESUL' : 'Geral')}
                                </td>
                                <td className="px-8 py-5 text-sm font-bold">
                                  <span className={`px-3 py-1 rounded-full text-[10px] uppercase tracking-wider ${row.mercado === 'LIVRE' ? 'bg-blue-100 text-blue-700' : 'bg-slate-100 text-slate-600'}`}>
                                    {row.mercado}
                                  </span>
                                </td>
                                <td className="px-8 py-5 text-sm font-medium text-slate-600">
                                  {row.modalidadeTarifaria || '-'}
                                </td>
                                {dashboardSubTab === 'operacionais' ? (
                                  <>
                                    {operationalSubTab === 'consumo' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.consumoPonta.toLocaleString('pt-BR')} kWh</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.consumoForaPonta.toLocaleString('pt-BR')} kWh</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">{(row.consumoPonta + row.consumoForaPonta).toLocaleString('pt-BR')} kWh</td>
                                      </>
                                    )}
                                    {operationalSubTab === 'ultrapassagem' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.demandaContratadaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.demandaContratadaForaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.ultrapassagemPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.ultrapassagemForaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-red-600">{(row.ultrapassagemPonta + row.ultrapassagemForaPonta).toLocaleString('pt-BR')} kW</td>
                                      </>
                                    )}
                                    {operationalSubTab === 'subutilizacao' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.demandaContratadaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.demandaContratadaForaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.demandaMedidaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.demandaMedidaForaPonta.toLocaleString('pt-BR')} kW</td>
                                        <td className={`px-8 py-5 text-sm font-bold text-right ${row.demandaMedidaPonta < (row.demandaContratadaPonta * 0.8) ? 'text-orange-600' : 'text-slate-600'}`}>
                                          {((row.demandaMedidaPonta / (row.demandaContratadaPonta || 1)) * 100).toFixed(1)}%
                                        </td>
                                      </>
                                    )}
                                    {operationalSubTab === 'reativa' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.reativaPonta.toLocaleString('pt-BR')} kVArh</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.reativaForaPonta.toLocaleString('pt-BR')} kVArh</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-purple-600">{(row.reativaPonta + row.reativaForaPonta).toLocaleString('pt-BR')} kVArh</td>
                                      </>
                                    )}
                                    {operationalSubTab === 'solar' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{(row.consumoPonta + row.consumoForaPonta).toLocaleString('pt-BR')} kWh</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.solarInjetadaOUC.toLocaleString('pt-BR')} kWh</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{row.solarInjetadaMUC.toLocaleString('pt-BR')} kWh</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-green-600">{((row.solarInjetadaOUC + row.solarInjetadaMUC) - (row.consumoPonta + row.consumoForaPonta)).toLocaleString('pt-BR')} kWh</td>
                                      </>
                                    )}
                                  </>
                                ) : (
                                  <>
                                    {financialSubTab === 'despesas' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">R$ {row.valorTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm text-slate-600">{row.cidade}</td>
                                      </>
                                    )}
                                    {financialSubTab === 'multa_ultrapassagem' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">R$ {row.valorUltrapassagemPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">R$ {row.valorUltrapassagemForaPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-red-600">R$ {(row.valorUltrapassagemPonta + row.valorUltrapassagemForaPonta).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                      </>
                                    )}
                                    {financialSubTab === 'multa_reativa' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">R$ {row.valorReativaPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">R$ {row.valorReativaForaPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-purple-600">R$ {(row.valorReativaPonta + row.valorReativaForaPonta).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                      </>
                                    )}
                                    {financialSubTab === 'tarifa_media' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">R$ {row.valorTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">{(row.consumoPonta + row.consumoForaPonta).toLocaleString('pt-BR')} kWh</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">R$ {(row.valorTotal / (row.consumoPonta + row.consumoForaPonta || 1)).toLocaleString('pt-BR', { minimumFractionDigits: 3 })}</td>
                                      </>
                                    )}
                                    {financialSubTab === 'energia_solar' && (
                                      <>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">R$ {(Math.abs(row.valorSolarOUC + row.valorSolarMUC) + (row.valorTotal - row.cip - row.outrosEncargos)).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-green-600">R$ {Math.abs(row.valorSolarOUC + row.valorSolarMUC).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                        <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-secondary">R$ {row.valorTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                      </>
                                    )}
                                  </>
                                )}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        ) : activeTab === 'analises' ? (
          <div className="space-y-8">
            <div className="bg-white p-10 rounded-3xl border border-sanesul-primary/10 shadow-xl">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-10">
                <div>
                  <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">Análises de Dados</h2>
                  <p className="text-sanesul-muted">Analise ultrapassagens e subutilização de demanda com base nas faturas processadas.</p>
                </div>
                <div className="flex items-center gap-4">
                  <button 
                    onClick={runAnalysis}
                    disabled={bills.filter(b => b.status === 'completed').length === 0}
                    className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
                  >
                    <BarChart3 size={16} />
                    Gerar Análise
                  </button>
                  {analysisResults && (
                    <div className="flex items-center gap-4">
                      <button 
                        onClick={exportAnalysisToCSV}
                        className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg shadow-green-600/20"
                      >
                        <Download size={16} />
                        Exportar
                      </button>
                      <button 
                        onClick={() => { setAnalysisResults(null); setAnalysisData([]); }}
                        className="flex items-center gap-2 px-6 py-3 bg-red-50 text-red-600 rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-red-100 transition-all"
                      >
                        <Trash2 size={16} />
                        Limpar
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {!analysisResults ? (
                <div className="border-2 border-dashed border-sanesul-primary/10 rounded-3xl p-20 text-center bg-slate-50/50">
                  <div className="w-20 h-20 bg-white rounded-2xl flex items-center justify-center mx-auto mb-6 shadow-sm">
                    <FileSpreadsheet size={40} className="text-sanesul-primary/40" />
                  </div>
                  <h3 className="text-xl font-bold text-sanesul-primary mb-2">Nenhuma análise gerada</h3>
                  <p className="text-sanesul-muted max-w-md mx-auto mb-8">
                    Clique em "Gerar Análise" para utilizar os dados das faturas processadas e calcular a demanda ideal.
                  </p>
                  <div className="flex justify-center gap-4">
                    <div className="p-4 rounded-xl bg-white border border-slate-100 text-left max-w-xs">
                      <div className="flex items-center gap-2 text-sanesul-primary font-bold text-xs mb-2">
                        <TrendingUp size={14} />
                        Ultrapassagem
                      </div>
                      <p className="text-[10px] text-sanesul-muted">Calculamos a demanda otimizada para eliminar multas de ultrapassagem (Resolução 1000 ANEEL).</p>
                    </div>
                    <div className="p-4 rounded-xl bg-white border border-slate-100 text-left max-w-xs">
                      <div className="flex items-center gap-2 text-sanesul-primary font-bold text-xs mb-2">
                        <Zap size={14} />
                        Subutilização
                      </div>
                      <p className="text-[10px] text-sanesul-muted">Identificamos a demanda ideal para evitar pagamentos por potência não utilizada.</p>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="space-y-10">
                  {/* Summary Cards */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8">
                    <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        {analysisResults.reduce((acc: any, curr: any) => acc + curr.economy, 0) >= 0 ? (
                          <TrendingUp size={80} className="text-green-600" />
                        ) : (
                          <TrendingDown size={80} className="text-red-600" />
                        )}
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Economia Potencial</p>
                      <p className={`text-4xl font-display font-bold ${analysisResults.reduce((acc: any, curr: any) => acc + curr.economy, 0) >= 0 ? 'text-green-600' : 'text-red-600'}`}>
                        R$ {analysisResults.reduce((acc: any, curr: any) => acc + curr.economy, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                      </p>
                    </div>
                    <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <AlertCircle size={80} className="text-red-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Ultrapassagens</p>
                      <p className="text-4xl font-display font-bold text-red-600">{analysisResults.filter((r: any) => r.isOverrun).length}</p>
                    </div>
                    <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <TrendingDown size={80} className="text-orange-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Subutilizações</p>
                      <p className="text-4xl font-display font-bold text-orange-600">{analysisResults.filter((r: any) => r.isSub).length}</p>
                    </div>
                    <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <Zap size={80} className="text-green-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Eficiência</p>
                      <p className="text-4xl font-display font-bold text-green-600">
                        {Math.round((analysisResults.filter((r: any) => !r.isOverrun && !r.isSub).length / analysisResults.length) * 100)}%
                      </p>
                    </div>
                  </div>

                  {/* Results Table */}
                  <div className="overflow-hidden rounded-2xl border border-sanesul-primary/5 shadow-sm">
                    <table className="w-full border-collapse bg-white text-left">
                      <thead>
                        <tr className="bg-slate-50">
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 w-10"></th>
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5">UC</th>
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Contratada (P/FP)</th>
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">Demanda Ideal</th>
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Gasto Real</th>
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">Economia</th>
                          <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">Meses Analisados</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {Object.values(analysisResults.reduce((acc: any, curr: any) => {
                          if (!acc[curr.uc]) {
                            acc[curr.uc] = {
                              uc: curr.uc,
                              months: [],
                              optimizedPonta: curr.optimizedPonta,
                              optimizedForaPonta: curr.optimizedForaPonta,
                              dcp: curr.dcp,
                              dcfp: curr.dcfp,
                              totalEconomy: 0,
                              totalCurrent: 0,
                              totalOptimized: 0
                            };
                          }
                          acc[curr.uc].months.push(curr);
                          acc[curr.uc].totalEconomy += curr.economy;
                          acc[curr.uc].totalCurrent += curr.currentTotal;
                          acc[curr.uc].totalOptimized += curr.optimizedTotal;
                          return acc;
                        }, {})).map((group: any, idx: number) => (
                          <React.Fragment key={idx}>
                            <tr 
                              className={`hover:bg-slate-50/50 transition-colors cursor-pointer ${expandedAnalysisUCs.has(group.uc) ? 'bg-slate-50/80' : ''}`}
                              onClick={() => toggleAnalysisUCExpansion(group.uc)}
                            >
                              <td className="px-6 py-4 text-center">
                                <ChevronRight size={16} className={`text-sanesul-muted transition-transform ${expandedAnalysisUCs.has(group.uc) ? 'rotate-90' : ''}`} />
                              </td>
                              <td className="px-6 py-4">
                                <div className="font-bold text-sanesul-primary text-xs">{group.uc}</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-mono text-slate-600">{group.dcp > 0 ? group.dcp.toFixed(2) : '-'} / {group.dcfp.toFixed(2)} kW</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-mono font-bold text-green-600">
                                  {group.optimizedPonta > 0 ? group.optimizedPonta.toFixed(2) : '-'} / {group.optimizedForaPonta.toFixed(2)} kW
                                </div>
                                <div className="text-[9px] text-green-500 uppercase font-bold tracking-tighter">Ideal Fixo (1 Ano)</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-mono font-bold text-sanesul-primary">R$ {group.totalCurrent.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className={`text-xs font-mono font-bold ${group.totalEconomy >= 0 ? 'text-green-600' : 'text-red-600'}`}>
                                  R$ {group.totalEconomy.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-bold text-slate-500">{group.months.length} meses</div>
                              </td>
                            </tr>
                            {expandedAnalysisUCs.has(group.uc) && (
                              <tr>
                                <td colSpan={7} className="px-10 py-4 bg-slate-50/30">
                                  <div className="overflow-hidden rounded-xl border border-slate-200 shadow-inner">
                                    <table className="w-full text-left border-collapse bg-white">
                                      <thead>
                                        <tr className="bg-slate-100/50">
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-slate-500">Mês/Ano</th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-slate-500 text-right">Medida (P/FP)</th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-red-600 text-right">Ultrapassagem</th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-orange-600 text-right">Subutilização</th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-sanesul-primary text-right">Gasto Real</th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-green-600 text-right">Economia</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {group.months.map((month: any, mIdx: number) => (
                                          <tr key={mIdx} className="hover:bg-slate-50/30 transition-colors">
                                            <td className="px-4 py-2 text-xs font-bold text-slate-700">{month.mes} {month.ano ? `/ ${month.ano}` : ''}</td>
                                            <td className="px-4 py-2 text-right text-xs font-mono font-bold text-sanesul-primary">{month.dcp > 0 ? month.dmp.toFixed(2) : '-'} / {month.dmfp.toFixed(2)} kW</td>
                                            <td className="px-4 py-2 text-right">
                                              {month.isOverrun ? (
                                                <div className="flex flex-col items-end gap-0.5">
                                                  {month.overrunPonta > 0 && <span className="text-[9px] font-bold text-red-600">P: +{month.overrunPonta.toFixed(2)}</span>}
                                                  {month.overrunForaPonta > 0 && <span className="text-[9px] font-bold text-red-600">FP: +{month.overrunForaPonta.toFixed(2)}</span>}
                                                </div>
                                              ) : <span className="text-slate-300 text-[9px]">-</span>}
                                            </td>
                                            <td className="px-4 py-2 text-right">
                                              {month.isSub ? (
                                                <div className="flex flex-col items-end gap-0.5">
                                                  {month.subPonta > 0 && <span className="text-[9px] font-bold text-orange-600">P: -{month.subPonta.toFixed(2)}</span>}
                                                  {month.subForaPonta > 0 && <span className="text-[9px] font-bold text-orange-600">FP: -{month.subForaPonta.toFixed(2)}</span>}
                                                </div>
                                              ) : <span className="text-slate-300 text-[9px]">-</span>}
                                            </td>
                                            <td className="px-4 py-2 text-right text-xs font-mono font-bold text-sanesul-primary">
                                              R$ {month.currentTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                            </td>
                                            <td className="px-4 py-2 text-right text-xs font-mono font-bold text-green-600">
                                              R$ {month.economy.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          </div>
        ) : activeTab === 'monitoramento' ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">Monitoramento de Despesas</h2>
                <p className="text-sanesul-muted">Acompanhamento detalhado de gastos e economia com demanda por cidade e UC.</p>
              </div>
              <button 
                onClick={runMonitoringAnalysis}
                disabled={bills.filter(b => b.status === 'completed').length === 0}
                className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <BarChart3 size={16} />
                Atualizar Monitoramento
              </button>
            </div>

            {!monitoringResults ? (
              <div className="border-2 border-dashed border-sanesul-primary/10 rounded-[40px] p-20 text-center bg-white/50 backdrop-blur-sm">
                <div className="w-24 h-24 bg-sanesul-primary/5 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-inner">
                  <DollarSign size={48} className="text-sanesul-primary/40" />
                </div>
                <h3 className="text-2xl font-display font-bold text-sanesul-primary mb-4">Pronto para analisar</h3>
                <p className="text-sanesul-muted max-w-md mx-auto mb-10 text-lg">
                  Clique no botão acima para processar os indicadores de despesa e economia baseados nas faturas extraídas.
                </p>
              </div>
            ) : (
              <div className="space-y-12">
                {/* Summary Cards */}
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8">
                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      <FileText size={80} className="text-sanesul-primary" />
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Faturas Analisadas</p>
                    <p className="text-4xl font-display font-bold text-sanesul-primary">{bills.filter(b => b.status === 'completed').length} <span className="text-base font-sans font-medium opacity-40">Arquivos</span></p>
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      {monitoringResults.generalTotalEconomy >= 0 ? (
                        <TrendingUp size={80} className="text-green-600" />
                      ) : (
                        <TrendingDown size={80} className="text-red-600" />
                      )}
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Economia Geral Total</p>
                    <p className={`text-4xl font-display font-bold ${monitoringResults.generalTotalEconomy >= 0 ? 'text-green-600' : 'text-red-600'}`}>
                      R$ {monitoringResults.generalTotalEconomy.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                    </p>
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      <DollarSign size={80} className="text-sanesul-primary" />
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Despesa Atual Total</p>
                    <p className="text-4xl font-display font-bold text-sanesul-primary">R$ {monitoringResults.generalTotalCurrent.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</p>
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      <Zap size={80} className="text-sanesul-secondary" />
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Potencial de Redução</p>
                    <p className="text-4xl font-display font-bold text-sanesul-secondary">
                      {((monitoringResults.generalTotalEconomy / monitoringResults.generalTotalCurrent) * 100).toFixed(1)}%
                    </p>
                  </div>
                </div>

                {/* Charts Section - Replaced with Text Summary */}
                <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl">
                  <div className="flex items-center justify-between mb-8">
                    <h3 className="text-xl font-display font-bold text-sanesul-primary">Resumo de Economia por Cidade</h3>
                    <div className="px-4 py-2 bg-sanesul-primary/5 rounded-full">
                      <span className="text-xs font-bold text-sanesul-primary uppercase tracking-wider">Top Impactos Financeiros</span>
                    </div>
                  </div>
                  
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-12">
                    {/* Top Economies */}
                    <div className="space-y-6">
                      <div className="flex items-center gap-3 pb-4 border-b border-green-100">
                        <div className="w-10 h-10 rounded-xl bg-green-50 flex items-center justify-center">
                          <TrendingUp className="text-green-600" size={20} />
                        </div>
                        <h4 className="font-display font-bold text-green-700 uppercase tracking-tight">Principais Economias</h4>
                      </div>
                      <div className="space-y-2">
                        {[...monitoringResults.cityData]
                          .filter(c => c.positiveEconomy > 0)
                          .sort((a, b) => b.positiveEconomy - a.positiveEconomy)
                          .slice(0, 10)
                          .map((city, idx) => (
                            <div key={idx} className="flex flex-col rounded-2xl bg-green-50/50 border border-green-100/50 hover:bg-green-50 transition-colors overflow-hidden">
                              <div 
                                className="flex items-center justify-between p-3 cursor-pointer"
                                onClick={() => toggleSummaryCity(`pos-${city.city}`)}
                              >
                                <div className="flex items-center gap-3">
                                  <span className="text-xs font-bold text-green-600/50 w-4">{String(idx + 1).padStart(2, '0')}</span>
                                  <span className="font-bold text-sanesul-primary text-sm">{city.city}</span>
                                </div>
                                <div className="flex items-center gap-3">
                                  <span className="font-mono font-bold text-green-600 text-sm">
                                    + R$ {city.positiveEconomy.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                  </span>
                                  <ChevronDown className={`w-4 h-4 text-green-600/50 transition-transform ${expandedSummaryCities.has(`pos-${city.city}`) ? 'rotate-180' : ''}`} />
                                </div>
                              </div>
                              {expandedSummaryCities.has(`pos-${city.city}`) && city.positiveUcs && (
                                <div className="px-3 pb-3 pt-1 border-t border-green-100/30 bg-green-50/30">
                                  <div className="space-y-1 mt-2">
                                    {city.positiveUcs.map((u: any, i: number) => (
                                      <div key={i} className="flex justify-between items-center text-xs">
                                        <span className="text-sanesul-muted font-medium">UC {u.uc}</span>
                                        <span className="font-mono text-green-600/80">
                                          + R$ {Math.abs(u.economy).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                        </span>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              )}
                            </div>
                          ))}
                        {monitoringResults.cityData.filter(c => c.positiveEconomy > 0).length === 0 && (
                          <p className="text-sm text-sanesul-muted italic p-4">Nenhuma economia significativa identificada.</p>
                        )}
                      </div>
                    </div>

                    {/* Top Losses */}
                    <div className="space-y-6">
                      <div className="flex items-center gap-3 pb-4 border-b border-red-100">
                        <div className="w-10 h-10 rounded-xl bg-red-50 flex items-center justify-center">
                          <TrendingDown className="text-red-600" size={20} />
                        </div>
                        <h4 className="font-display font-bold text-red-700 uppercase tracking-tight">Principais Prejuízos</h4>
                      </div>
                      <div className="space-y-2">
                        {[...monitoringResults.cityData]
                          .filter(c => c.negativeEconomy < 0)
                          .sort((a, b) => a.negativeEconomy - b.negativeEconomy)
                          .slice(0, 10)
                          .map((city, idx) => (
                            <div key={idx} className="flex flex-col rounded-2xl bg-red-50/50 border border-red-100/50 hover:bg-red-50 transition-colors overflow-hidden">
                              <div 
                                className="flex items-center justify-between p-3 cursor-pointer"
                                onClick={() => toggleSummaryCity(`neg-${city.city}`)}
                              >
                                <div className="flex items-center gap-3">
                                  <span className="text-xs font-bold text-red-600/50 w-4">{String(idx + 1).padStart(2, '0')}</span>
                                  <span className="font-bold text-sanesul-primary text-sm">{city.city}</span>
                                </div>
                                <div className="flex items-center gap-3">
                                  <span className="font-mono font-bold text-red-600 text-sm">
                                    - R$ {Math.abs(city.negativeEconomy).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                  </span>
                                  <ChevronDown className={`w-4 h-4 text-red-600/50 transition-transform ${expandedSummaryCities.has(`neg-${city.city}`) ? 'rotate-180' : ''}`} />
                                </div>
                              </div>
                              {expandedSummaryCities.has(`neg-${city.city}`) && city.negativeUcs && (
                                <div className="px-3 pb-3 pt-1 border-t border-red-100/30 bg-red-50/30">
                                  <div className="space-y-1 mt-2">
                                    {city.negativeUcs.map((u: any, i: number) => (
                                      <div key={i} className="flex justify-between items-center text-xs">
                                        <span className="text-sanesul-muted font-medium">UC {u.uc}</span>
                                        <span className="font-mono text-red-600/80">
                                          - R$ {Math.abs(u.economy).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                        </span>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              )}
                            </div>
                          ))}
                        {monitoringResults.cityData.filter(c => c.negativeEconomy < 0).length === 0 && (
                          <p className="text-sm text-sanesul-muted italic p-4">Nenhum prejuízo significativo identificado.</p>
                        )}
                      </div>
                    </div>
                  </div>
                </div>

                {/* City Breakdown */}
                <div className="space-y-8">
                  <div className="flex items-center justify-between">
                    <h3 className="text-2xl font-display font-bold text-sanesul-primary border-l-4 border-sanesul-secondary pl-4">Detalhamento por Cidade</h3>
                  </div>
                {/* Changed UCs Group */}
                {monitoringResults.changedUCs.length > 0 && (
                  <div className="bg-white rounded-[40px] border border-sanesul-primary/10 shadow-2xl overflow-hidden mb-8">
                    <div className="bg-slate-50/80 px-10 py-8 border-b border-sanesul-primary/5">
                      <h3 className="text-2xl font-display font-bold text-sanesul-primary border-l-4 border-yellow-500 pl-4">
                        Unidades com Alteração de Demanda
                      </h3>
                      <p className="text-xs font-bold text-sanesul-muted uppercase tracking-widest mt-2 pl-5">
                        {monitoringResults.changedUCs.length} Unidades Encontradas
                      </p>
                    </div>
                    
                    <div className="p-10">
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] w-10"></th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">Unidade Consumidora</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">Cidade</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">Contratada Atual (P/FP)</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">Demanda Ideal (P/FP)</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">Gasto Atual (Total)</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">Economia Acumulada</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">Status</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {monitoringResults.changedUCs.map((uc: any, uIdx: number) => (
                              <React.Fragment key={uIdx}>
                                <tr className={`group hover:bg-slate-50/50 transition-colors cursor-pointer ${expandedUCs.has(uc.uc) ? 'bg-slate-50/80' : ''}`} onClick={() => toggleUCExpansion(uc.uc)}>
                                  <td className="py-6 text-center">
                                    <ChevronRight size={16} className={`text-sanesul-muted transition-transform ${expandedUCs.has(uc.uc) ? 'rotate-90' : ''}`} />
                                  </td>
                                  <td className="py-6">
                                    <div className="flex items-center gap-3">
                                      <div className="w-10 h-10 rounded-xl bg-sanesul-primary/5 flex items-center justify-center font-bold text-sanesul-primary text-xs">
                                        {uIdx + 1}
                                      </div>
                                      <span className="font-bold text-sanesul-primary font-mono">{uc.uc}</span>
                                    </div>
                                  </td>
                                  <td className="py-6 text-xs font-bold text-slate-500 uppercase tracking-wider">
                                    {uc.city}
                                  </td>
                                  <td className="py-6 text-center">
                                    <span className="px-3 py-1.5 bg-slate-50 text-slate-700 rounded-lg text-xs font-bold font-mono">
                                      {uc.monthlyData[0]?.dcp} / {uc.monthlyData[0]?.dcfp} kW
                                    </span>
                                  </td>
                                  <td className="py-6 text-center">
                                    <span className="px-3 py-1.5 bg-green-50 text-green-700 rounded-lg text-xs font-bold font-mono">
                                      {uc.optPonta} / {uc.optForaPonta} kW
                                    </span>
                                  </td>
                                  <td className="py-6 text-right font-mono text-sm text-sanesul-muted">
                                    R$ {uc.totalCurrent.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                  </td>
                                  <td className="py-6 text-right font-bold text-green-600 text-lg">
                                    R$ {uc.totalEconomy.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                  </td>
                                  <td className="py-6 text-right">
                                    <div className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                                      uc.totalEconomy > 0 ? 'bg-green-100 text-green-700' : uc.totalEconomy < 0 ? 'bg-red-100 text-red-700' : 'bg-slate-100 text-slate-500'
                                    }`}>
                                      {uc.totalEconomy > 0 ? <TrendingUp size={12} /> : uc.totalEconomy < 0 ? <TrendingDown size={12} /> : <div className="w-3 h-3 rounded-full bg-slate-400" />}
                                      {uc.totalEconomy > 0 ? 'Economia' : uc.totalEconomy < 0 ? 'Prejuízo' : 'Neutro'}
                                    </div>
                                  </td>
                                </tr>
                                {expandedUCs.has(uc.uc) && (
                                  <tr>
                                    <td colSpan={8} className="px-10 py-0">
                                      <div className="bg-slate-50/50 rounded-2xl p-6 mb-6 border border-slate-100 animate-in fade-in slide-in-from-top-2 duration-300">
                                        <div className="flex items-center gap-2 mb-4">
                                          <Calendar size={14} className="text-sanesul-primary" />
                                          <h5 className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">Histórico de Alterações e Economia</h5>
                                        </div>
                                        <div className="grid grid-cols-6 gap-4 mb-2 px-4">
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest">Mês/Ano</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">Contratada (P/FP)</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">Medida (P/FP)</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">Gasto Real</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">Ref. Anterior</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">Economia</div>
                                        </div>
                                        <div className="space-y-2">
                                          {uc.monthlyData.map((month: any, mIdx: number) => (
                                            <div key={mIdx} className={`grid grid-cols-6 gap-4 px-4 py-3 rounded-xl border transition-colors ${month.hasChanged ? 'bg-yellow-50 border-yellow-200' : 'bg-white border-slate-100 hover:border-sanesul-primary/20'}`}>
                                              <div className="flex flex-col">
                                                <span className="text-xs font-bold text-slate-700">{formatMonth(month.mes)}/{month.ano}</span>
                                                {month.hasChanged && <span className="text-[9px] font-bold text-yellow-600 uppercase tracking-wider mt-1">Alteração de Contrato</span>}
                                              </div>
                                              <div className="text-xs font-mono text-center text-slate-500">{month.dcp} / {month.dcfp} kW</div>
                                              <div className="text-xs font-mono text-center text-slate-500">{month.dmp} / {month.dmfp} kW</div>
                                              <div className="text-xs font-mono text-right font-bold text-sanesul-primary">R$ {month.currentTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                                              <div className="text-xs font-mono text-right text-slate-400">
                                                {month.referenceTotal > 0 ? `R$ ${month.referenceTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}
                                              </div>
                                              <div className={`text-xs font-mono text-right font-bold ${month.economy > 0 ? 'text-green-600' : month.economy < 0 ? 'text-red-500' : 'text-slate-400'}`}>
                                                {month.economy !== 0 ? `R$ ${month.economy.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}
                                              </div>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}

                {/* Unchanged UCs Group */}
                {monitoringResults.unchangedUCs.length > 0 && (
                  <div className="bg-white rounded-[40px] border border-sanesul-primary/10 shadow-2xl overflow-hidden">
                    <div className="bg-slate-50/80 px-10 py-8 border-b border-sanesul-primary/5">
                      <h3 className="text-2xl font-display font-bold text-sanesul-muted border-l-4 border-slate-300 pl-4">
                        Unidades sem Alteração (Contrato Estável)
                      </h3>
                      <p className="text-xs font-bold text-sanesul-muted uppercase tracking-widest mt-2 pl-5">
                        {monitoringResults.unchangedUCs.length} Unidades Encontradas
                      </p>
                    </div>
                    
                    <div className="p-10">
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse opacity-80 hover:opacity-100 transition-opacity">
                          <thead>
                            <tr>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] w-10"></th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">Unidade Consumidora</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">Cidade</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">Contratada (P/FP)</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">Demanda Ideal (P/FP)</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">Gasto Atual (Total)</th>
                              <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">Status</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {monitoringResults.unchangedUCs.map((uc: any, uIdx: number) => (
                              <React.Fragment key={uIdx}>
                                <tr className={`group hover:bg-slate-50/50 transition-colors cursor-pointer ${expandedUCs.has(uc.uc) ? 'bg-slate-50/80' : ''}`} onClick={() => toggleUCExpansion(uc.uc)}>
                                  <td className="py-6 text-center">
                                    <ChevronRight size={16} className={`text-sanesul-muted transition-transform ${expandedUCs.has(uc.uc) ? 'rotate-90' : ''}`} />
                                  </td>
                                  <td className="py-6">
                                    <div className="flex items-center gap-3">
                                      <div className="w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center font-bold text-slate-500 text-xs">
                                        {uIdx + 1}
                                      </div>
                                      <span className="font-bold text-slate-600 font-mono">{uc.uc}</span>
                                    </div>
                                  </td>
                                  <td className="py-6 text-xs font-bold text-slate-500 uppercase tracking-wider">
                                    {uc.city}
                                  </td>
                                  <td className="py-6 text-center">
                                    <span className="px-3 py-1.5 bg-slate-50 text-slate-700 rounded-lg text-xs font-bold font-mono">
                                      {uc.monthlyData[0]?.dcp} / {uc.monthlyData[0]?.dcfp} kW
                                    </span>
                                  </td>
                                  <td className="py-6 text-center">
                                    <span className="px-3 py-1.5 bg-green-50 text-green-700 rounded-lg text-xs font-bold font-mono">
                                      {uc.optPonta} / {uc.optForaPonta} kW
                                    </span>
                                  </td>
                                  <td className="py-6 text-right font-mono text-sm text-sanesul-muted">
                                    R$ {uc.totalCurrent.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                                  </td>
                                  <td className="py-6 text-right">
                                    <div className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[10px] font-bold uppercase tracking-wider bg-slate-100 text-slate-500">
                                      <div className="w-3 h-3 rounded-full bg-slate-400" />
                                      Sem Alteração
                                    </div>
                                  </td>
                                </tr>
                                {expandedUCs.has(uc.uc) && (
                                  <tr>
                                    <td colSpan={8} className="px-10 py-0">
                                      <div className="bg-slate-50/50 rounded-2xl p-6 mb-6 border border-slate-100 animate-in fade-in slide-in-from-top-2 duration-300">
                                        <div className="flex items-center gap-2 mb-4">
                                          <Calendar size={14} className="text-sanesul-primary" />
                                          <h5 className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">Detalhamento Mensal</h5>
                                        </div>
                                        <div className="grid grid-cols-4 gap-4 mb-2 px-4">
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest">Mês/Ano</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">Contratada (P/FP)</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">Medida (P/FP)</div>
                                          <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">Gasto Real</div>
                                        </div>
                                        <div className="space-y-2">
                                          {uc.monthlyData.map((month: any, mIdx: number) => (
                                            <div key={mIdx} className="grid grid-cols-4 gap-4 px-4 py-3 bg-white rounded-xl border border-slate-100 hover:border-sanesul-primary/20 transition-colors">
                                              <div className="text-xs font-bold text-slate-700">{formatMonth(month.mes)}/{month.ano}</div>
                                              <div className="text-xs font-mono text-center text-slate-500">{month.dcp} / {month.dcfp} kW</div>
                                              <div className="text-xs font-mono text-center text-slate-500">{month.dmp} / {month.dmfp} kW</div>
                                              <div className="text-xs font-mono text-right font-bold text-sanesul-primary">R$ {month.currentTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                )}
                </div>
              </div>
            )}
          </div>
        ) : activeTab === 'monitoramento_reativo' ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">Monitoramento Reativo</h2>
                <p className="text-sanesul-muted">Acompanhamento do Valor da Energia Reativa Excedente por UC.</p>
              </div>
              <div className="flex flex-wrap gap-4 items-center">
                <div className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm">
                  <Calendar size={16} className="text-sanesul-primary" />
                  <select 
                    value={selectedReactiveMonth}
                    onChange={(e) => setSelectedReactiveMonth(e.target.value)}
                    className="bg-transparent text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                  >
                    <option value="all">Todos os Meses</option>
                    {availableMonths.map(month => (
                      <option key={month} value={month}>{month}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            {(() => {
              const reactiveBills = bills.filter(b => {
                if (b.status !== 'completed') return false;
                if (selectedReactiveMonth !== 'all' && `${b.mesReferencia}/${b.anoLeitura}` !== selectedReactiveMonth) return false;
                const totalReativo = parseValue(b.valorEnergiaReativaExcedPonta) + parseValue(b.valorEnergiaReativaExcedFPonta);
                return totalReativo > 100;
              });
              
              const grouped = reactiveBills.reduce((acc, bill) => {
                const uc = String(bill.uc);
                if (!acc[uc]) acc[uc] = { uc, cidade: bill.cidade || '', totalPonta: 0, totalFPonta: 0, totalFatura: 0, bills: [] };
                acc[uc].totalPonta += parseValue(bill.valorEnergiaReativaExcedPonta);
                acc[uc].totalFPonta += parseValue(bill.valorEnergiaReativaExcedFPonta);
                acc[uc].totalFatura += parseValue(bill.valorTotal);
                acc[uc].bills.push(bill);
                return acc;
              }, {} as Record<string, { uc: string, cidade: string, totalPonta: number, totalFPonta: number, totalFatura: number, bills: typeof bills }>);

              const reactiveData = (Object.values(grouped) as { uc: string, cidade: string, totalPonta: number, totalFPonta: number, totalFatura: number, bills: typeof bills }[]).sort((a, b) => {
                let valA: any = 0;
                let valB: any = 0;
                
                if (reactiveSortField === 'uc') {
                  valA = a.uc;
                  valB = b.uc;
                } else if (reactiveSortField === 'cidade') {
                  valA = a.cidade;
                  valB = b.cidade;
                } else if (reactiveSortField === 'totalPonta') {
                  valA = a.totalPonta;
                  valB = b.totalPonta;
                } else if (reactiveSortField === 'totalFPonta') {
                  valA = a.totalFPonta;
                  valB = b.totalFPonta;
                } else if (reactiveSortField === 'totalGeral') {
                  valA = a.totalPonta + a.totalFPonta;
                  valB = b.totalPonta + b.totalFPonta;
                } else if (reactiveSortField === 'percentual') {
                  valA = a.totalFatura > 0 ? (a.totalPonta + a.totalFPonta) / a.totalFatura : 0;
                  valB = b.totalFatura > 0 ? (b.totalPonta + b.totalFPonta) / b.totalFatura : 0;
                }

                if (typeof valA === 'string' && typeof valB === 'string') {
                  return reactiveSortDirection === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
                }
                return reactiveSortDirection === 'asc' ? (valA as number) - (valB as number) : (valB as number) - (valA as number);
              });
              
              const totalGeral = reactiveData.reduce((acc, curr) => acc + curr.totalPonta + curr.totalFPonta, 0);

              return (
                <div className="space-y-8">
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                    <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <Zap size={80} className="text-red-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">Total Multa Reativa</p>
                      <p className="text-4xl font-display font-bold text-red-600">
                        R$ {totalGeral.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                      <thead>
                        <tr>
                          <th 
                            className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary transition-colors"
                            onClick={() => handleReactiveSort('uc')}
                          >
                            <div className="flex items-center gap-1">
                              UC {reactiveSortField === 'uc' && (reactiveSortDirection === 'asc' ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
                            </div>
                          </th>
                          <th 
                            className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary transition-colors"
                            onClick={() => handleReactiveSort('cidade')}
                          >
                            <div className="flex items-center gap-1">
                              Cidade {reactiveSortField === 'cidade' && (reactiveSortDirection === 'asc' ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
                            </div>
                          </th>
                          <th 
                            className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                            onClick={() => handleReactiveSort('totalPonta')}
                          >
                            <div className="flex items-center justify-end gap-1">
                              Total Reativa Ponta (R$) {reactiveSortField === 'totalPonta' && (reactiveSortDirection === 'asc' ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
                            </div>
                          </th>
                          <th 
                            className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                            onClick={() => handleReactiveSort('totalFPonta')}
                          >
                            <div className="flex items-center justify-end gap-1">
                              Total Reativa F. Ponta (R$) {reactiveSortField === 'totalFPonta' && (reactiveSortDirection === 'asc' ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
                            </div>
                          </th>
                          <th 
                            className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                            onClick={() => handleReactiveSort('totalGeral')}
                          >
                            <div className="flex items-center justify-end gap-1">
                              Total Geral (R$) {reactiveSortField === 'totalGeral' && (reactiveSortDirection === 'asc' ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
                            </div>
                          </th>
                          <th 
                            className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                            onClick={() => handleReactiveSort('percentual')}
                          >
                            <div className="flex items-center justify-end gap-1">
                              % da Fatura {reactiveSortField === 'percentual' && (reactiveSortDirection === 'asc' ? <TrendingUp size={12} /> : <TrendingDown size={12} />)}
                            </div>
                          </th>
                          <th className="px-6 py-4 border-b border-sanesul-primary/5 w-10"></th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-sanesul-primary/5">
                        {reactiveData.map((data, idx) => (
                          <React.Fragment key={idx}>
                            <tr 
                              className="hover:bg-sanesul-primary/5 transition-colors cursor-pointer group"
                              onClick={() => toggleReactiveUc(data.uc)}
                            >
                              <td className="px-6 py-4">
                                <div className="font-bold text-sanesul-primary">{data.uc}</div>
                                <div className="text-[10px] text-sanesul-muted uppercase tracking-wider mt-1">{data.bills.length} faturas</div>
                              </td>
                              <td className="px-6 py-4">
                                <div className="text-sm text-slate-600">{data.cidade}</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-sm font-mono text-slate-600">R$ {data.totalPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-sm font-mono text-slate-600">R$ {data.totalFPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-sm font-bold text-red-600">R$ {(data.totalPonta + data.totalFPonta).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-sm font-bold text-orange-500">
                                  {data.totalFatura > 0 ? (((data.totalPonta + data.totalFPonta) / data.totalFatura) * 100).toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '0,00'}%
                                </div>
                              </td>
                              <td className="px-6 py-4 text-center">
                                <ChevronDown className={`w-5 h-5 text-sanesul-primary/40 transition-transform group-hover:text-sanesul-primary ${expandedReactiveUcs.has(data.uc) ? 'rotate-180' : ''}`} />
                              </td>
                            </tr>
                            {expandedReactiveUcs.has(data.uc) && (
                              <tr>
                                <td colSpan={7} className="p-0 bg-slate-50/50">
                                  <div className="px-12 py-6 border-t border-sanesul-primary/5 shadow-inner">
                                    <table className="w-full text-left">
                                      <thead>
                                        <tr>
                                          <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200">Mês/Ano</th>
                                          <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">Valor Ponta (R$)</th>
                                          <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">Valor F. Ponta (R$)</th>
                                          <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">Total Mês (R$)</th>
                                          <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">% da Fatura</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {data.bills
                                          .sort((a, b) => {
                                            const yearA = parseInt(a.anoLeitura || '0', 10);
                                            const yearB = parseInt(b.anoLeitura || '0', 10);
                                            if (yearA !== yearB) return yearB - yearA;
                                            return getMonthNumber(b.mesReferencia) - getMonthNumber(a.mesReferencia);
                                          })
                                          .map((bill, bIdx) => {
                                            const vPonta = parseValue(bill.valorEnergiaReativaExcedPonta);
                                            const vFPonta = parseValue(bill.valorEnergiaReativaExcedFPonta);
                                            const vFatura = parseValue(bill.valorTotal);
                                            const percentual = vFatura > 0 ? ((vPonta + vFPonta) / vFatura) * 100 : 0;
                                            return (
                                              <tr key={bIdx} className="hover:bg-white transition-colors">
                                                <td className="px-4 py-3 text-sm font-medium text-slate-600">{bill.mesReferencia}/{bill.anoLeitura}</td>
                                                <td className="px-4 py-3 text-sm font-mono text-slate-500 text-right">R$ {vPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                                <td className="px-4 py-3 text-sm font-mono text-slate-500 text-right">R$ {vFPonta.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                                <td className="px-4 py-3 text-sm font-bold text-red-500 text-right">R$ {(vPonta + vFPonta).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                                                <td className="px-4 py-3 text-sm font-bold text-orange-500 text-right">{percentual.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%</td>
                                              </tr>
                                            );
                                          })}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        ))}
                        {reactiveData.length === 0 && (
                          <tr>
                            <td colSpan={5} className="px-6 py-12 text-center text-sanesul-muted">
                              Nenhuma UC com energia reativa excedente encontrada.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              );
            })()}
          </div>
        ) : activeTab === 'relatorio' ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">Relatório Financeiro</h2>
                <p className="text-sanesul-muted">Visão consolidada de faturamento, impostos e indicadores financeiros.</p>
              </div>
              <div className="flex flex-wrap gap-4 items-center">
                <div className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm">
                  <Calendar size={16} className="text-sanesul-primary" />
                  <select 
                    value={selectedRelatorioMonth}
                    onChange={(e) => setSelectedRelatorioMonth(e.target.value)}
                    className="bg-transparent text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                  >
                    <option value="all">Todos os Meses</option>
                    {availableMonths.map(month => (
                      <option key={month} value={month}>{month}</option>
                    ))}
                  </select>
                </div>

                <div className="relative">
                  <button 
                    onClick={() => setIsRelatorioTypeDropdownOpen(!isRelatorioTypeDropdownOpen)}
                    className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                  >
                    <Filter size={16} className="text-sanesul-primary" />
                    <span>
                      {selectedRelatorioType.includes('all') 
                        ? 'Todos os Tipos' 
                        : selectedRelatorioType.length > 1 
                          ? `${selectedRelatorioType.length} Selecionados` 
                          : selectedRelatorioType[0]}
                    </span>
                    <ChevronDown size={14} className={`transition-transform ${isRelatorioTypeDropdownOpen ? 'rotate-180' : ''}`} />
                  </button>
                  
                  {isRelatorioTypeDropdownOpen && (
                    <>
                      <div className="fixed inset-0 z-40" onClick={() => setIsRelatorioTypeDropdownOpen(false)} />
                      <div className="absolute top-full left-0 mt-2 w-48 bg-white rounded-xl shadow-xl border border-slate-100 z-50 py-2">
                        <label className="flex items-center gap-3 px-4 py-2 hover:bg-slate-50 cursor-pointer transition-colors">
                          <input 
                            type="checkbox" 
                            checked={selectedRelatorioType.includes('all')}
                            onChange={() => {
                              setSelectedRelatorioType(['all']);
                            }}
                            className="rounded border-slate-300 text-sanesul-primary focus:ring-sanesul-primary"
                          />
                          <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">Todos os Tipos</span>
                        </label>
                        <div className="h-px bg-slate-100 my-1" />
                        {availableRelatorioTypes.map(type => (
                          <label key={type} className="flex items-center gap-3 px-4 py-2 hover:bg-slate-50 cursor-pointer transition-colors">
                            <input 
                              type="checkbox" 
                              checked={selectedRelatorioType.includes(type)}
                              onChange={() => {
                                let newTypes = [...selectedRelatorioType].filter(t => t !== 'all');
                                if (newTypes.includes(type)) {
                                  newTypes = newTypes.filter(t => t !== type);
                                } else {
                                  newTypes.push(type);
                                }
                                if (newTypes.length === 0) {
                                  setSelectedRelatorioType(['all']);
                                } else {
                                  setSelectedRelatorioType(newTypes);
                                }
                              }}
                              className="rounded border-slate-300 text-sanesul-primary focus:ring-sanesul-primary"
                            />
                            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">{type}</span>
                          </label>
                        ))}
                      </div>
                    </>
                  )}
                </div>
                <div className="flex gap-2">
                  <button 
                    onClick={exportRelatorioToCSV}
                    className="flex items-center gap-2 px-6 py-3 bg-white border border-sanesul-primary/10 text-sanesul-primary rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-primary/5 transition-all"
                  >
                    <Download size={16} />
                    Exportar CSV
                  </button>
                  <button 
                    id="btn-gerar-relatorio"
                    onClick={() => {
                      setTempMemoNumber(memoNumber);
                      setTempMemoNfEnergisa(memoNfEnergisa);
                      setTempMemoNfElektro(memoNfElektro);
                      setShowMemoNumberPrompt(true);
                    }}
                    className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20"
                  >
                    <BarChart3 size={16} />
                    Gerar Relatório
                  </button>
                </div>
              </div>
            </div>

            {Object.keys(uploadProgress).length > 0 && (
              <div className="mb-8 space-y-4">
                {Object.entries(uploadProgress).map(([fileId, progress]: [string, { status: string, percent: number, fileName: string, fileSize: number, abortController: AbortController | null }]) => (
                  <motion.div 
                    key={fileId}
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -20 }}
                    className="bg-white p-6 rounded-2xl border border-sanesul-primary/10 shadow-sm flex items-center gap-4"
                  >
                    <div className="flex-1">
                      <div className="flex justify-between items-center mb-2">
                        <div className="flex flex-col">
                          <span className="text-sm font-bold text-slate-900 truncate max-w-[200px]">{progress.fileName}</span>
                          <span className="text-[10px] text-slate-500">{(progress.fileSize / 1024).toFixed(1)} KB</span>
                        </div>
                        <span className="text-sm font-bold text-sanesul-primary">{progress.percent}%</span>
                      </div>
                      <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden">
                        <div 
                          className="bg-sanesul-primary h-2.5 rounded-full transition-all duration-300 ease-out" 
                          style={{ width: `${progress.percent}%` }}
                        ></div>
                      </div>
                      <span className="text-xs text-slate-500 mt-1 block">{progress.status}</span>
                    </div>
                    {progress.percent < 100 && progress.status !== 'Erro' && (
                      <button 
                        onClick={() => progress.abortController?.abort()} 
                        className="text-red-500 hover:text-red-700 p-2 rounded-full hover:bg-red-50 transition-colors"
                        title="Cancelar upload"
                      >
                        <X size={20} />
                      </button>
                    )}
                  </motion.div>
                ))}
              </div>
            )}

            {Object.keys(agrupadoraFiles).length > 0 && (
              <div className="bg-blue-50 p-6 rounded-2xl border border-blue-100">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-2">
                    <CheckCircle2 size={16} className="text-blue-600" />
                    <h3 className="text-sm font-bold text-blue-800 uppercase tracking-wider">
                      Dados Extraídos das Faturas
                    </h3>
                  </div>
                  <button
                    onClick={() => setAgrupadoraFiles({})}
                    className="flex items-center gap-1 px-3 py-1.5 bg-white border border-red-100 text-red-500 rounded-lg text-[10px] font-bold uppercase tracking-widest hover:bg-red-50 transition-all"
                  >
                    <Trash2 size={12} />
                    Limpar
                  </button>
                </div>
                <div className="flex flex-col lg:flex-row gap-6">
                  {(Object.values(agrupadoraFiles) as AgrupadoraData[]).map((data, idx) => {
                    const isDetailed = data.concessionaria.includes('DETALHADO');
                    return (
                    <div key={idx} className="flex-1 bg-white p-5 rounded-xl border border-blue-100 shadow-sm min-w-[300px]">
                      <div className="flex justify-between items-start mb-4 pb-3 border-b border-slate-100">
                        <span className="text-sm font-bold text-blue-600 uppercase tracking-wider">{data.concessionaria}</span>
                        <span className="text-[10px] text-slate-400 truncate max-w-[150px] text-right" title={data.fileName}>{data.fileName}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-x-4 gap-y-3">
                        <div className="flex flex-col">
                          <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Mês Ref</span>
                          <span className="font-mono font-bold text-slate-700 text-sm">{formatReference(data.mesReferencia) || '-'}</span>
                        </div>
                        {!isDetailed && (
                          <div className="flex flex-col">
                            <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Vencimento</span>
                            <span className="font-mono font-bold text-slate-700 text-sm">{data.vencimento || '-'}</span>
                          </div>
                        )}
                        {!isDetailed && (
                          <div className="flex flex-col col-span-2">
                            <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Nota Fiscal</span>
                            <span className="font-mono font-bold text-slate-700 text-sm">{data.numeroNotaFiscal || '-'}</span>
                          </div>
                        )}
                        {!isDetailed && (
                          <>
                            <div className="flex flex-col">
                              <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">ICMS</span>
                              <span className="font-mono font-bold text-slate-700 text-sm">R$ {data.icms.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                            </div>
                            <div className="flex flex-col">
                              <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">COFINS</span>
                              <span className="font-mono font-bold text-slate-700 text-sm">R$ {data.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                            </div>
                            <div className="flex flex-col">
                              <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">PIS</span>
                              <span className="font-mono font-bold text-slate-700 text-sm">R$ {data.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                            </div>
                          </>
                        )}
                        <div className={`flex flex-col ${isDetailed ? 'col-span-2' : ''}`}>
                          <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">CIP</span>
                          <span className={`font-mono font-bold text-slate-700 ${isDetailed ? 'text-2xl text-blue-700' : 'text-sm'}`}>R$ {data.cip.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                        </div>
                        {!isDetailed && (
                          <div className="flex flex-col col-span-2 pt-3 border-t border-slate-100 mt-1">
                            <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">Total Faturado</span>
                            <span className="font-mono font-bold text-blue-700 text-lg">R$ {data.valorTotal.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                          </div>
                        )}
                      </div>
                    </div>
                  )})}
                </div>
              </div>
            )}

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5">
                <h3 className="text-lg font-display font-bold text-sanesul-primary mb-8 flex items-center gap-3">
                  <DollarSign className="text-sanesul-primary" size={20} />
                  Resumo de Impostos e Contribuições - Energisa
                </h3>
                <div className="space-y-6">
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">PIS</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.energisa.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">COFINS</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.energisa.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">ICMS</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.energisa.icms.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-sanesul-primary/5 rounded-2xl border border-sanesul-primary/10">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-primary">CIP MUNICIPAL</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.energisa.cip.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                </div>
              </div>

              <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5">
                <h3 className="text-lg font-display font-bold text-sanesul-primary mb-8 flex items-center gap-3">
                  <DollarSign className="text-sanesul-primary" size={20} />
                  Resumo de Impostos e Contribuições - ELEKTRO
                </h3>
                <div className="space-y-6">
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">PIS</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.elektro.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">COFINS</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.elektro.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">ICMS</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.elektro.icms.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-sanesul-primary/5 rounded-2xl border border-sanesul-primary/10">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-primary">CIP MUNICIPAL</span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">R$ {memoData.elektro.cip.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-white rounded-[40px] border border-sanesul-primary/5 shadow-2xl overflow-hidden hidden">
              <div className="p-10 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-xl font-display font-bold text-sanesul-primary">Detalhamento por Unidade Consumidora</h3>
                <div className="text-[10px] font-bold text-sanesul-muted uppercase tracking-widest bg-slate-100 px-3 py-1 rounded-full">
                  {Array.from(new Set(filteredRelatorioData.map(d => d.uc))).filter(Boolean).length} Unidades
                </div>
              </div>
              <div className="overflow-auto max-h-[600px]">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 z-10 bg-slate-50">
                    <tr className="bg-slate-50/50">
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100">UC</th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100">Cidade</th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">PIS</th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">COFINS</th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">ICMS</th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">CIP</th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">Total</th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from(new Set(filteredRelatorioData.map(d => d.uc))).filter(Boolean).map(uc => {
                      const ucData = filteredRelatorioData.filter(d => d.uc === uc);
                      const totals = {
                        pis: ucData.reduce((acc, curr) => acc + curr.pis, 0),
                        cofins: ucData.reduce((acc, curr) => acc + curr.cofins, 0),
                        icms: ucData.reduce((acc, curr) => acc + curr.icms, 0),
                        cip: ucData.reduce((acc, curr) => acc + curr.cip, 0),
                        total: ucData.reduce((acc, curr) => acc + curr.valorTotal, 0)
                      };
                      return (
                        <tr key={uc} className="hover:bg-slate-50/50 transition-colors group">
                          <td className="px-8 py-6 border-b border-slate-50">
                            <span className="text-sm font-bold text-sanesul-primary group-hover:text-sanesul-secondary transition-colors">{uc}</span>
                          </td>
                          <td className="px-8 py-6 border-b border-slate-50">
                            <span className="text-xs font-medium text-sanesul-muted">{ucData[0]?.cidade || '-'}</span>
                          </td>
                          <td className="px-8 py-6 border-b border-slate-50 text-right">
                            <span className="text-xs font-mono font-bold text-slate-600">R$ {totals.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                          </td>
                          <td className="px-8 py-6 border-b border-slate-50 text-right">
                            <span className="text-xs font-mono font-bold text-slate-600">R$ {totals.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                          </td>
                          <td className="px-8 py-6 border-b border-slate-50 text-right">
                            <span className="text-xs font-mono font-bold text-slate-600">R$ {totals.icms.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                          </td>
                          <td className="px-8 py-6 border-b border-slate-50 text-right">
                            <span className="text-xs font-mono font-bold text-sanesul-primary">R$ {totals.cip.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                          </td>
                          <td className="px-8 py-6 border-b border-slate-50 text-right">
                            <span className="text-sm font-display font-bold text-sanesul-primary">R$ {totals.total.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</span>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : null}
      </main>

      {/* Custom Memo Number Prompt Modal */}
      {showMemoNumberPrompt && (
        <div className="fixed inset-0 z-[110] flex items-center justify-center bg-black/60 backdrop-blur-sm">
          <div className="bg-white w-full max-w-md rounded-3xl shadow-2xl p-8 animate-in zoom-in duration-200">
            <div className="flex items-center gap-3 mb-6">
              <div className="w-10 h-10 rounded-xl bg-sanesul-primary/10 flex items-center justify-center">
                <FileText className="text-sanesul-primary" size={20} />
              </div>
              <div>
                <h3 className="text-lg font-display font-bold text-sanesul-primary">Dados do Memorando</h3>
                <p className="text-[10px] text-sanesul-muted uppercase tracking-widest font-bold">Identificação e Notas Fiscais</p>
              </div>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-bold text-sanesul-muted uppercase tracking-widest mb-2 px-1">Número do Memorando</label>
                <input 
                  type="text" 
                  value={tempMemoNumber}
                  onChange={(e) => setTempMemoNumber(e.target.value)}
                  placeholder="Ex: 001447/2024/GEDEO/DCO"
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold text-sanesul-primary focus:outline-none focus:ring-2 focus:ring-sanesul-primary/20 focus:bg-white transition-all"
                  autoFocus
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      setMemoNumber(tempMemoNumber);
                      setMemoNfEnergisa(tempMemoNfEnergisa);
                      setMemoNfElektro(tempMemoNfElektro);
                      setShowMemoNumberPrompt(false);
                      setShowMemo(true);
                    }
                  }}
                />
              </div>

              <div>
                <label className="block text-[10px] font-bold text-sanesul-muted uppercase tracking-widest mb-2 px-1">NF Energisa</label>
                <input 
                  type="text" 
                  value={tempMemoNfEnergisa}
                  onChange={(e) => setTempMemoNfEnergisa(e.target.value)}
                  placeholder="Ex: 123456"
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold text-sanesul-primary focus:outline-none focus:ring-2 focus:ring-sanesul-primary/20 focus:bg-white transition-all"
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      setMemoNumber(tempMemoNumber);
                      setMemoNfEnergisa(tempMemoNfEnergisa);
                      setMemoNfElektro(tempMemoNfElektro);
                      setShowMemoNumberPrompt(false);
                      setShowMemo(true);
                    }
                  }}
                />
              </div>

              <div>
                <label className="block text-[10px] font-bold text-sanesul-muted uppercase tracking-widest mb-2 px-1">NF Elektro</label>
                <input 
                  type="text" 
                  value={tempMemoNfElektro}
                  onChange={(e) => setTempMemoNfElektro(e.target.value)}
                  placeholder="Ex: 789012"
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold text-sanesul-primary focus:outline-none focus:ring-2 focus:ring-sanesul-primary/20 focus:bg-white transition-all"
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      setMemoNumber(tempMemoNumber);
                      setMemoNfEnergisa(tempMemoNfEnergisa);
                      setMemoNfElektro(tempMemoNfElektro);
                      setShowMemoNumberPrompt(false);
                      setShowMemo(true);
                    }
                  }}
                />
              </div>
              
              <div className="flex gap-3 pt-4">
                <button 
                  onClick={() => setShowMemoNumberPrompt(false)}
                  className="flex-1 px-6 py-3 bg-slate-100 text-slate-500 rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-slate-200 transition-all"
                >
                  Cancelar
                </button>
                <button 
                  onClick={() => {
                    setMemoNumber(tempMemoNumber);
                    setMemoNfEnergisa(tempMemoNfEnergisa);
                    setMemoNfElektro(tempMemoNfElektro);
                    setShowMemoNumberPrompt(false);
                    setShowMemo(true);
                  }}
                  className="flex-1 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20"
                >
                  Confirmar
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Memo Modal */}
      {showMemo && (
        <div className="fixed inset-0 z-[100] overflow-y-auto bg-black/60 backdrop-blur-sm print:static print:bg-white print:overflow-visible">
          <div className="flex min-h-full items-start justify-center p-4 print:p-0 print:block">
            <div className="bg-white w-full max-w-4xl rounded-3xl shadow-2xl relative my-4 sm:my-8 print:my-0 print:shadow-none print:rounded-none">
              <div className="sticky top-0 bg-white/95 backdrop-blur-md border-b border-slate-100 p-4 sm:p-6 flex justify-between items-center z-20 print:hidden rounded-t-3xl">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-sanesul-primary/10 flex items-center justify-center">
                  <FileText className="text-sanesul-primary" size={20} />
                </div>
                <div>
                  <h3 className="text-lg font-display font-bold text-sanesul-primary">Memorando de Faturamento</h3>
                  <p className="text-[10px] text-sanesul-muted uppercase tracking-widest font-bold">GEDEO/DCO - Sanesul</p>
                </div>
              </div>
              <div className="flex gap-3">
                <button 
                  onClick={handleDownloadPDF}
                  disabled={isGeneratingPDF}
                  className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isGeneratingPDF ? (
                    <>
                      <Loader2 size={16} className="animate-spin" />
                      Gerando PDF...
                    </>
                  ) : (
                    <>
                      <Printer size={16} />
                      Baixar PDF
                    </>
                  )}
                </button>
                <button 
                  onClick={handleDownloadDocx}
                  className="flex items-center gap-2 px-6 py-3 bg-white border border-sanesul-primary text-sanesul-primary rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-primary/5 transition-all"
                >
                  <FileText size={16} />
                  Baixar DOCX
                </button>
                <button 
                  onClick={() => setShowMemo(false)}
                  className="p-3 hover:bg-slate-100 rounded-xl transition-colors text-slate-400 hover:text-red-500"
                >
                  <X size={20} />
                </button>
              </div>
            </div>

            <div id="memo-content" className="p-16 font-serif text-slate-800 leading-relaxed print:p-8 bg-white">
              {/* Header */}
              <div className="flex justify-between items-start mb-16">
                <div className="flex items-center gap-6">
                  <img src="https://www.sanesul.ms.gov.br/images/logo_sanesul.png" alt="Sanesul" className="h-20 object-contain" referrerPolicy="no-referrer" />
                  <div className="h-16 w-px bg-slate-200" />
                  <div className="text-[11px] font-bold text-sanesul-primary leading-tight uppercase tracking-tight">
                    Empresa de Saneamento de <br /> Mato Grosso do Sul S.A.<br />
                    <span className="text-sanesul-muted font-medium">Diretoria da Presidência</span>
                  </div>
                </div>
                <img src="https://www.ms.gov.br/wp-content/uploads/2023/01/logo-governo-ms.png" alt="Governo MS" className="h-20 object-contain" referrerPolicy="no-referrer" />
              </div>

              {/* Memo Info */}
              <div className="space-y-1 mb-12 text-sm">
                <p className="font-bold text-base memo-number-text">MEMO Nº {memoNumber}</p>
                <p className="text-slate-500 italic">Campo Grande, {new Intl.DateTimeFormat('pt-BR', { dateStyle: 'long' }).format(new Date())}.</p>
              </div>

              <div className="space-y-4 mb-12 text-sm border-l-4 border-sanesul-primary/20 pl-6 py-2">
                <p><span className="font-bold text-sanesul-primary uppercase tracking-wider text-[10px]">DE:</span> <br />GEDEO - Gerência de Desenvolvimento Operacional</p>
                <p><span className="font-bold text-sanesul-primary uppercase tracking-wider text-[10px]">PARA:</span> <br />GEFI - Gerência Financeira e Gestão de Recursos</p>
                <p><span className="font-bold text-sanesul-primary uppercase tracking-wider text-[10px]">ASSUNTO:</span> <br />Faturas Agrupadora Operacional Energisa e Agrupadora Elektro — {selectedRelatorioMonth === 'all' ? 'Consolidado' : selectedRelatorioMonth}{!selectedRelatorioType.includes('all') ? ` (${selectedRelatorioType.join(', ')})` : ''}.</p>
              </div>

              <p className="mb-6 font-medium">Prezado(a),</p>
              <p className="mb-10 text-justify">
                Seguem anexas para pagamento as faturas de energia elétrica Agrupadora da concessionária Energisa MS, e Agrupadora da concessionária Elektro — todas referentes ao mês de <span className="font-bold underline decoration-sanesul-primary/30 underline-offset-4">{selectedRelatorioMonth === 'all' ? 'todos os períodos' : selectedRelatorioMonth}</span>{!selectedRelatorioType.includes('all') ? ` (Tipo: ${selectedRelatorioType.join(', ')})` : ''} e correspondentes às unidades operacionais da SANESUL.
              </p>

              <div className="mb-6 flex items-center gap-3">
                <div className="h-px flex-1 bg-slate-200" />
                <p className="font-bold text-xs uppercase tracking-widest text-slate-400">Tabela 1 - Faturas Anexas</p>
                <div className="h-px flex-1 bg-slate-200" />
              </div>
              
              <table className="w-full border-collapse border border-slate-300 text-sm mb-12 shadow-sm">
                <thead>
                  <tr className="bg-slate-50">
                    <th className="border border-slate-300 p-3 text-left text-[10px] font-bold uppercase tracking-wider">LOCALIDADE</th>
                    <th className="border border-slate-300 p-3 text-right text-[10px] font-bold uppercase tracking-wider">VALOR (R$)</th>
                    <th className="border border-slate-300 p-3 text-center text-[10px] font-bold uppercase tracking-wider">NOTA FISCAL</th>
                    <th className="border border-slate-300 p-3 text-center text-[10px] font-bold uppercase tracking-wider">REF: MÊS / ANO</th>
                  </tr>
                </thead>
                <tbody>
                  {/* Energisa */}
                  <tr className="font-bold bg-sanesul-primary/5">
                    <td className="border border-slate-300 p-3 text-sanesul-primary">Agrupadora Energisa Operacional</td>
                    <td className="border border-slate-300 p-3 text-right text-sanesul-primary">R$ {memoData.energisa.total.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                    <td rowSpan={5} className="border border-slate-300 p-3 text-center align-middle font-mono text-xs max-w-[120px] break-all">{memoData.energisa.nf}</td>
                    <td rowSpan={5} className="border border-slate-300 p-3 text-center align-middle font-bold text-sanesul-primary">{memoData.energisa.mesRef}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">PIS</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.energisa.pis > 0 ? `R$ ${memoData.energisa.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">COFINS</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.energisa.cofins > 0 ? `R$ ${memoData.energisa.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">ICMS</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.energisa.icms > 0 ? `R$ ${memoData.energisa.icms.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600 italic">Tarifa de Iluminação Pública</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.energisa.cip > 0 ? `R$ ${memoData.energisa.cip.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>

                  {/* Elektro */}
                  <tr className="font-bold bg-sanesul-secondary/5">
                    <td className="border border-slate-300 p-3 text-sanesul-secondary">Agrupadora Elektro</td>
                    <td className="border border-slate-300 p-3 text-right text-sanesul-secondary">R$ {memoData.elektro.total.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                    <td rowSpan={5} className="border border-slate-300 p-3 text-center align-middle font-mono text-xs max-w-[120px] break-all">{memoData.elektro.nf}</td>
                    <td rowSpan={5} className="border border-slate-300 p-3 text-center align-middle font-bold text-sanesul-secondary">{memoData.elektro.mesRef}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">PIS</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.elektro.pis > 0 ? `R$ ${memoData.elektro.pis.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">COFINS</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.elektro.cofins > 0 ? `R$ ${memoData.elektro.cofins.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">ICMS</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.elektro.icms > 0 ? `R$ ${memoData.elektro.icms.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600 italic">Tarifa de Iluminação Pública</td>
                    <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">{memoData.elektro.cip > 0 ? `R$ ${memoData.elektro.cip.toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-'}</td>
                  </tr>

                  {/* Total */}
                  <tr className="font-bold bg-slate-100">
                    <td className="border border-slate-300 p-4 uppercase text-xs tracking-wider">TOTAL (Agrupadora ENERGISA + ELEKTRO)</td>
                    <td className="border border-slate-300 p-4 text-right text-base text-sanesul-primary">R$ {(memoData.energisa.total + memoData.elektro.total).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}</td>
                    <td className="border border-slate-300 p-4 text-center text-slate-300">-------------------</td>
                    <td className="border border-slate-300 p-4 text-center text-slate-300">-------------------</td>
                  </tr>
                </tbody>
              </table>

              <div className="flex justify-between items-end text-[10px] text-slate-400 mb-12 italic">
                <p>Proc. N.º 694/2018</p>
                <p>Nota Orçamentária Nº 003/2019</p>
              </div>

              <div className="mt-16 text-sm text-slate-600 mb-12">
                <p>A planilha contendo a estratificação dos dados apresentados neste memorando está disponível em <span className="font-mono text-[10px] bg-slate-50 px-2 py-1 rounded break-all">\\srv-fs 01\DADOS\DCO\GEDEO\OPERACAO_AGUA\COTAA\ENERGIA\FATURAS</span>.</p>
              </div>

              <p className="mb-12">Atenciosamente,</p>

              <div className="mt-32 text-center">
                <div className="w-72 h-px bg-slate-800 mx-auto mb-4" />
                <p className="font-bold text-lg text-slate-900">Fabio Roberto Alves da Silva</p>
                <p className="text-xs text-slate-500 uppercase tracking-widest font-bold mt-1">Engenheiro Eletricista/GEDEO</p>
                <p className="text-[10px] text-slate-400 mt-0.5">Gerência de Desenvolvimento Operacional</p>
              </div>
            </div>
          </div>
        </div>
        </div>
      )}

      {/* Footer Info */}
      {isBillModalOpen && editingBill && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between p-6 border-b border-slate-100">
              <h2 className="text-xl font-bold text-sanesul-primary flex items-center gap-2">
                <Pencil size={20} />
                {editingBill.id && bills.some(b => b.id === editingBill.id) ? 'Editar Fatura' : 'Nova Fatura Manual'}
              </h2>
              <button
                onClick={() => setIsBillModalOpen(false)}
                className="w-8 h-8 flex items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 transition-colors"
              >
                <X size={20} />
              </button>
            </div>
            <div className="p-6 overflow-y-auto flex-1">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {/* Basic Info */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-2 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">Informações Básicas</h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">UC</label>
                  <input
                    type="text"
                    value={editingBill.uc || ''}
                    onChange={e => setEditingBill({ ...editingBill, uc: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Concessionária</label>
                  <select
                    value={editingBill.concessionaria || 'ENERGISA'}
                    onChange={e => setEditingBill({ ...editingBill, concessionaria: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  >
                    <option value="ENERGISA">ENERGISA</option>
                    <option value="ELEKTRO">ELEKTRO</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Cidade</label>
                  <input
                    type="text"
                    value={editingBill.cidade || ''}
                    onChange={e => setEditingBill({ ...editingBill, cidade: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Mês Referência</label>
                  <input
                    type="text"
                    placeholder="Ex: Janeiro"
                    value={editingBill.mesReferencia || ''}
                    onChange={e => setEditingBill({ ...editingBill, mesReferencia: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Ano Leitura</label>
                  <input
                    type="text"
                    placeholder="Ex: 2024"
                    value={editingBill.anoLeitura || ''}
                    onChange={e => setEditingBill({ ...editingBill, anoLeitura: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Data de Vencimento</label>
                  <input
                    type="text"
                    placeholder="Ex: 15/08/2025"
                    value={editingBill.dataVencimento || ''}
                    onChange={e => setEditingBill({ ...editingBill, dataVencimento: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Nota Fiscal</label>
                  <input
                    type="text"
                    value={editingBill.numeroNotaFiscal || ''}
                    onChange={e => setEditingBill({ ...editingBill, numeroNotaFiscal: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Modalidade Tarifária</label>
                  <input
                    type="text"
                    value={editingBill.modalidadeTarifaria || ''}
                    onChange={e => setEditingBill({ ...editingBill, modalidadeTarifaria: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Subgrupo</label>
                  <input
                    type="text"
                    value={editingBill.subgrupo || ''}
                    onChange={e => setEditingBill({ ...editingBill, subgrupo: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Tipo</label>
                  <select
                    value={editingBill.tipo || 'OPERACIONAL'}
                    onChange={e => setEditingBill({ ...editingBill, tipo: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  >
                    <option value="OPERACIONAL">OPERACIONAL</option>
                    <option value="ADMINISTRATIVO">ADMINISTRATIVO</option>
                    <option value="LIVRE">LIVRE</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Total (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorTotal || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorTotal: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                
                {/* Demanda */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">Demanda</h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Contratada Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPontaKW || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPontaKW: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Contratada Fora Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaForaPontaKW || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaForaPontaKW: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Medida Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaMedidaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPotenciaMedidaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Medida Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaMedidaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorDemandaPotenciaMedidaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Medida Fora Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaMedidaForaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPotenciaMedidaForaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Medida Fora Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaMedidaForaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorDemandaPotenciaMedidaForaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Não Consumida Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaNaoConsumidaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPotenciaNaoConsumidaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Não Consumida Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaNaoConsumidaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorDemandaPotenciaNaoConsumidaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Não Consumida Fora Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaNaoConsumidaFPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPotenciaNaoConsumidaFPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Não Consumida Fora Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaNaoConsumidaFPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorDemandaPotenciaNaoConsumidaFPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Ultrapassagem Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaAtivaUltrapPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPotenciaAtivaUltrapPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Ultrapassagem Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaAtivaUltrapPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorDemandaPotenciaAtivaUltrapPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Ultrapassagem Fora Ponta (kW)</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaAtivaUltrapFPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, demandaPotenciaAtivaUltrapFPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Ultrapassagem Fora Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaAtivaUltrapFPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorDemandaPotenciaAtivaUltrapFPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Consumo */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">Consumo</h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Consumo Ponta (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.consumoKwhPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, consumoKwhPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Consumo Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorConsumoKwhPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorConsumoKwhPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Consumo Fora Ponta (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.consumoKwhForaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, consumoKwhForaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Consumo Fora Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorConsumoKwhForaPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorConsumoKwhForaPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Reativa */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">Energia Reativa</h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Reativa Ponta (kVArh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaReativaExcedPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, energiaReativaExcedPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Reativa Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaReativaExcedPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorEnergiaReativaExcedPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Reativa Fora Ponta (kVArh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaReativaExcedFPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, energiaReativaExcedFPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor Reativa Fora Ponta (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaReativaExcedFPonta || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorEnergiaReativaExcedFPonta: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Geração Distribuída */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">Geração Distribuída</h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Energia Injetada (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaInjetadaKwh || ''}
                    onChange={e => setEditingBill({ ...editingBill, energiaInjetadaKwh: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Energia Compensada (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaCompensadaKwh || ''}
                    onChange={e => setEditingBill({ ...editingBill, energiaCompensadaKwh: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">GDI oUC (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaAtvInjetadaGDIOUC || ''}
                    onChange={e => setEditingBill({ ...editingBill, energiaAtvInjetadaGDIOUC: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor GDI oUC (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaAtvInjetadaGDIOUC || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorEnergiaAtvInjetadaGDIOUC: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">GDI mUC (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaAtvInjetadaGDIMUC || ''}
                    onChange={e => setEditingBill({ ...editingBill, energiaAtvInjetadaGDIMUC: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Valor GDI mUC (R$)</label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaAtvInjetadaGDIMUC || ''}
                    onChange={e => setEditingBill({ ...editingBill, valorEnergiaAtvInjetadaGDIMUC: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Encargos e Impostos */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">Encargos e Impostos</h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">CIP (R$)</label>
                  <input
                    type="text"
                    value={editingBill.cip || ''}
                    onChange={e => setEditingBill({ ...editingBill, cip: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Outros Encargos (R$)</label>
                  <input
                    type="text"
                    value={editingBill.outrosEncargos || ''}
                    onChange={e => setEditingBill({ ...editingBill, outrosEncargos: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">PIS (R$)</label>
                  <input
                    type="text"
                    value={editingBill.pis || ''}
                    onChange={e => setEditingBill({ ...editingBill, pis: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">COFINS (R$)</label>
                  <input
                    type="text"
                    value={editingBill.cofins || ''}
                    onChange={e => setEditingBill({ ...editingBill, cofins: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">ICMS (R$)</label>
                  <input
                    type="text"
                    value={editingBill.icms || ''}
                    onChange={e => setEditingBill({ ...editingBill, icms: e.target.value })}
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
              </div>
            </div>
            <div className="p-6 border-t border-slate-100 flex justify-end gap-3 bg-slate-50 rounded-b-2xl">
              <button
                onClick={() => setIsBillModalOpen(false)}
                className="px-6 py-2.5 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-200 transition-colors"
              >
                Cancelar
              </button>
              <button
                onClick={async () => {
                  let billToSave = { ...editingBill } as BillData;
                  if (billToSave.uc && UCS_OPER.has(String(billToSave.uc))) {
                    billToSave.tipo = 'OPER';
                  } else if (billToSave.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(billToSave.uc))) {
                    let mod = billToSave.modalidadeTarifaria || '';
                    if (!mod.toUpperCase().includes('LIVRE')) {
                      billToSave.modalidadeTarifaria = mod ? `${mod} - LIVRE` : 'LIVRE';
                    }
                    billToSave.tipo = 'LIVRE';
                  }
                  billToSave.mercado = billToSave.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(billToSave.uc)) ? 'LIVRE' : 'CATIVO';
                  const isExisting = bills.some(b => b.id === billToSave.id);
                  
                  if (isSupabaseConfigured && isAuthenticated) {
                    try {
                      const { data: { user } } = await supabase.auth.getUser();
                      if (user) {
                        const dbData = mapBillDataToDb(billToSave, user.id);
                        if (isExisting) {
                          let { error } = await supabase.from('bills').update(dbData).eq('id', billToSave.id);
                          if (error && (error.message.includes('data_vencimento') || error.message.includes('mercado') || error.details?.includes('data_vencimento') || error.details?.includes('mercado') || error.code === 'PGRST204')) {
                            console.warn('Coluna data_vencimento ou mercado não encontrada. Atualizando sem elas...');
                            const { data_vencimento, mercado, ...fallbackData } = dbData;
                            const fallbackRes = await supabase.from('bills').update(fallbackData).eq('id', billToSave.id);
                            error = fallbackRes.error;
                          }
                          if (error) {
                            console.error('Erro ao atualizar fatura no Supabase:', error);
                            return;
                          }
                          setBills(prev => prev.map(b => b.id === billToSave.id ? billToSave : b));
                        } else {
                          let { data, error } = await supabase.from('bills').insert(dbData).select().single();
                          if (error && (error.message.includes('data_vencimento') || error.message.includes('mercado') || error.details?.includes('data_vencimento') || error.details?.includes('mercado') || error.code === 'PGRST204')) {
                            console.warn('Coluna data_vencimento ou mercado não encontrada. Inserindo sem elas...');
                            const { data_vencimento, mercado, ...fallbackData } = dbData;
                            const fallbackRes = await supabase.from('bills').insert(fallbackData).select().single();
                            error = fallbackRes.error;
                            data = fallbackRes.data;
                          }
                          if (error) {
                            console.error('Erro ao inserir fatura no Supabase:', error);
                            return;
                          }
                          const newBill = mapDbToBillData(data);
                          setBills(prev => {
                            if (prev.some(b => b.id === newBill.id)) return prev;
                            return [...prev, newBill];
                          });
                        }
                      }
                    } catch (err) {
                      console.error('Erro inesperado ao salvar fatura:', err);
                      return;
                    }
                  } else {
                    if (isExisting) {
                      setBills(prev => prev.map(b => b.id === billToSave.id ? billToSave : b));
                    } else {
                      setBills(prev => {
                        if (prev.some(b => b.id === billToSave.id)) return prev;
                        return [...prev, billToSave];
                      });
                    }
                  }
                  setIsBillModalOpen(false);
                }}
                className="flex items-center gap-2 px-6 py-2.5 bg-sanesul-primary text-white rounded-xl text-sm font-bold shadow-lg shadow-sanesul-primary/20 hover:bg-sanesul-primary/90 transition-all active:scale-95"
              >
                <Save size={16} />
                Salvar Fatura
              </button>
            </div>
          </div>
        </div>
      )}

      <footer className="max-w-[1600px] mx-auto mt-24 pb-12 px-8">
        <div className="pt-8 border-t border-sanesul-primary/10 flex flex-col md:flex-row justify-between items-center gap-6">
          <div className="flex items-center gap-4">
            <div className="w-8 h-8 rounded-lg bg-sanesul-primary/10 flex items-center justify-center">
              <Zap className="w-4 h-4 text-sanesul-primary" />
            </div>
            <div>
              <div className="text-xs font-bold text-sanesul-primary uppercase tracking-widest">Sanesul - Portal de Inteligência Energética</div>
              <div className="text-[10px] text-sanesul-muted mt-0.5">© 1979 Empresa de Saneamento de Mato Grosso do Sul</div>
              <div className="text-[10px] text-sanesul-muted mt-0.5">Developed by: Fabio Roberto alves da Silva</div>
            </div>
          </div>
          <div className="flex items-center gap-8">
            <div className="flex flex-col items-end">
              <span className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">COTAA</span>
              <span className="text-[10px] text-sanesul-muted">Coodenação: Alexandre Santos Andrade Monteiro</span>
            </div>
            <div className="flex flex-col items-end">
              <span className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">GEDEO</span>
              <span className="text-[10px] text-sanesul-muted">Gerência: Elthon Santos Teixeira</span>
            </div>
          </div>
        </div>
      </footer>
      {/* Custom Confirmation Modal */}
      <AnimatePresence>
        {showConfirmModal && confirmModalData && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="bg-white rounded-3xl shadow-2xl max-w-md w-full overflow-hidden border border-slate-100"
            >
              <div className="p-8">
                <div className={`w-14 h-14 rounded-2xl flex items-center justify-center mb-6 ${
                  confirmModalData.type === 'danger' ? 'bg-red-50 text-red-600' : 'bg-blue-50 text-blue-600'
                }`}>
                  {confirmModalData.type === 'danger' ? <Trash2 size={28} /> : <AlertCircle size={28} />}
                </div>
                <h3 className="text-xl font-display font-bold text-slate-900 mb-2">
                  {confirmModalData.title}
                </h3>
                <p className="text-slate-600 leading-relaxed">
                  {confirmModalData.message}
                </p>
              </div>
              <div className="p-6 bg-slate-50 flex gap-3 justify-end">
                {!confirmModalData.isAlert && (
                  <button
                    onClick={() => setShowConfirmModal(false)}
                    className="px-6 py-2.5 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-200 transition-colors"
                  >
                    Cancelar
                  </button>
                )}
                <button
                  onClick={confirmModalData.onConfirm}
                  className={`px-8 py-2.5 rounded-xl text-sm font-bold text-white transition-all active:scale-95 shadow-lg ${
                    confirmModalData.type === 'danger' 
                      ? 'bg-red-600 shadow-red-600/20 hover:bg-red-700' 
                      : 'bg-sanesul-primary shadow-sanesul-primary/20 hover:bg-sanesul-primary/90'
                  }`}
                >
                  {confirmModalData.isAlert ? 'Entendido' : 'Confirmar'}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
