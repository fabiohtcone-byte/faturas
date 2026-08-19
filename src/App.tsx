/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, {
  useState,
  useCallback,
  useRef,
  useMemo,
  useEffect,
} from "react";
import { extractTextFromPdf, parseBillText } from "./utils/pdfParser";
import domtoimage from "dom-to-image-more";
import { jsPDF } from "jspdf";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  AlignmentType,
  ShadingType,
  ImageRun,
  VerticalAlign,
} from "docx";
import { saveAs } from "file-saver";
import localforage from "localforage";
import {
  GoogleGenAI,
  Type,
  GenerateContentResponse,
  ThinkingLevel,
} from "@google/genai";
import * as XLSX from "xlsx";
import { supabase, isSupabaseConfigured } from "./lib/supabase";
import { UC_MIGRATION_MAP } from "./uc_migration";
import { REQUESTED_ADJUSTMENTS, ORIGINAL_CONTRATADAS } from "./requested_adjustments";
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
  Database,
  Home,
  BarChart2,
  GitCompare,
  Activity,
  Battery,
  ZapOff,
  Leaf,
  Key,
  AlertTriangle,
  Building,
  MapPin,
  CalendarX2,
  ListX,
  Cloud,
  FolderPlus,
  FolderUp,
  Sparkles,
  Layers,
  Settings2,
  SlidersHorizontal,
  RefreshCw,
  FileUp,
  FileCheck,
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
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
  LabelList,
} from "recharts";

const PDF_PREVISAO_BASE: Record<string, number> = {
  "16797005173": 4120.34,
  "24963105181": 22196.34,
  "22732005159": 38198.51,
  "160951205134": 6968.39,
  "38058605160": 13156.26,
  "141571705146": 7755.36,
  "167728305135": 60803.01,
  "179291005130": 16238.13,
  "120020005100": 15825.54,
  "116105605155": 17342.37,
  "85024005139": 29867.12,
  "151272005143": 8777.50,
  "14096605180": 2208.40,
  "11691105139": 18407.58,
  "163738105109": 13790.09,
  "153773805183": 91615.24,
  "181328805106": 3915.23,
  "91490205181": 10162.03,
  "96000605106": 10422.31,
  "127185505143": 3262.83,
  "148525805184": 34501.49,
  "126974105187": 17602.55,
  "156602805192": 1808.73,
  "101003105158": 69691.16,
  "10926205169": 2541.83,
  "31156405106": 3714.40,
  "101390305130": 40918.13,
  "100072705175": 87527.86,
  "101373605106": 4272.72,
  "35399605165": 12085.99,
  "89431905152": 10612.72,
  "87868105162": 53139.38,
  "142469005180": 7697.51,
  "37265905137": 194830.78,
  "144378605167": 60817.12,
  "3498071": 44169.93,
  "134798305148": 118594.86,
  "87312605145": 3962.79,
  "1626524": 5543.97,
  "138266605137": 6332.49,
};

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
  consumoGrupoB?: string;
  demandaTodosPeriodosKW?: string;
  consumoKwh?: string;
  demandaTodosPeriodos?: string;
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
  gerencia?: string;
  locin?: string;
  status: "pending" | "processing" | "completed" | "error";
  error?: string;
  file?: File;
  progress?: number;
  abortController?: AbortController;
  createdAt?: number;
}

export interface UCLocinMapping {
  uc: string;
  gerencia: string;
  locin: string;
  cidade: string;
}

// --- Constants ---

let UCS_LIVRE_MERCADO_LIVRE = new Set<string>();

// Carrega UCs do Mercado Livre salvas no localStorage ou usa a lista padrão
try {
  const savedML = localStorage.getItem("sanesul_mercado_livre_ucs");
  if (savedML) {
    const list = JSON.parse(savedML);
    UCS_LIVRE_MERCADO_LIVRE = new Set(list.map((uc: any) => String(uc).trim()));
  } else {
    // Lista padrão inicial
    const defaultList = [
      "65132005113",
      "158196505196",
      "103208705140",
      "14096605180",
      "164910105198",
      "151289605137",
      "25205905149",
      "141571705146",
      "87312605145",
      "101390305130",
      "16167605114",
      "142469005180",
      "163944605136",
      "14350505108",
      "160951205134",
      "22732005159",
      "1626524",
      "89982805122",
      "156602805192",
      "87868105162",
      "19995105197",
      "121387805187",
      "8851605161",
      "10926205169",
      "153773805183",
      "180833405162",
      "93446505190",
      "101003105158",
      "49633305113",
      "90408305117",
      "134798305148",
      "92807105114",
      "37265905137",
      "144378605167",
      "148159505104",
      "123125105129",
      "126974105187",
      "163738105109",
      "162836505110",
      "120020005100",
      "179291005130",
      "116105605155",
      "85024005139",
      "39038513",
      "41904974",
      "41905059",
      "18256767",
      "18256830"
    ];
    UCS_LIVRE_MERCADO_LIVRE = new Set(defaultList);
    localStorage.setItem("sanesul_mercado_livre_ucs", JSON.stringify(defaultList));
  }
} catch (err) {
  console.warn("Falha ao inicializar UCs do Mercado Livre:", err);
  UCS_LIVRE_MERCADO_LIVRE = new Set();
}

const UCS_OPER = new Set([
  "172701705150",
  "113100805164",
  "103217205180",
  "120020005100",
  "179291005130",
  "116105605155",
  "85024005139",
  "65132005113",
  "93022105109",
  "62131705107",
  "8396905117",
  "112499605106",
  "62131405130",
  "162517205196",
  "162517105103",
  "167009105148",
  "162517005105",
  "89131105135",
  "168499305191",
  "158196905158",
  "158196805162",
  "158196705177",
  "158196605181",
  "158196505196",
  "5079005140",
  "6288105102",
  "56699805185",
  "5677905183",
  "8092005187",
  "8390905126",
  "4185005108",
  "141574205134",
  "74799405160",
  "75395505121",
  "58205005124",
  "60612105100",
  "124508705180",
  "49794205103",
  "130254205109",
  "69364705182",
  "130762105138",
  "32416305160",
  "50106505161",
  "179585305159",
  "179585205163",
  "16797005173",
  "184977605141",
  "137971405106",
  "62120305195",
  "101410505103",
  "51602005150",
  "185281105180",
  "132263205107",
  "122406905120",
  "25209505104",
  "103208705140",
  "93604305181",
  "49789105159",
  "101405705112",
  "49789005163",
  "116098105132",
  "9888705108",
  "12896205113",
  "165205605195",
  "11691105139",
  "122098305147",
  "14096805161",
  "11691005143",
  "14096605180",
  "103804405124",
  "52508605180",
  "102911205180",
  "103202105107",
  "170890205128",
  "50105405194",
  "164910205183",
  "164910105198",
  "132555305179",
  "164910005105",
  "164909905180",
  "125402805170",
  "151289805118",
  "162806105134",
  "77796405188",
  "106496205175",
  "140660405123",
  "151289705122",
  "160391505116",
  "151289605137",
  "96601605175",
  "22192905120",
  "96601505180",
  "36884505166",
  "18598905157",
  "96601305101",
  "149608205156",
  "166104405119",
  "156087305142",
  "154300505157",
  "158483905128",
  "165805505187",
  "158483805132",
  "158483705147",
  "123908005149",
  "47986105120",
  "158475305170",
  "140359405100",
  "85324705160",
  "146901305102",
  "158475205185",
  "158475105190",
  "134058205118",
  "158475005107",
  "76286605117",
  "64830405102",
  "105912705181",
  "138266605137",
  "136171305153",
  "103507505171",
  "157885505127",
  "157885405131",
  "170287705105",
  "25205905149",
  "180174105119",
  "108599305182",
  "35399605165",
  "47375405160",
  "20092805168",
  "19211605190",
  "13489605155",
  "172675805116",
  "172675705120",
  "172675605135",
  "7495805153",
  "141571705146",
  "153979205150",
  "31513505109",
  "166101005101",
  "166100905149",
  "24604805179",
  "53694705118",
  "10181105197",
  "54899305145",
  "54899205150",
  "96000605106",
  "141263205190",
  "140063005120",
  "7492705112",
  "7492605127",
  "5983105194",
  "13175205107",
  "170884105191",
  "181987705142",
  "137964305198",
  "128403305168",
  "91511005161",
  "91510905101",
  "181052305186",
  "44674205128",
  "171479105125",
  "171479005130",
  "171478905177",
  "92106705133",
  "126316405103",
  "143071205188",
  "97217905103",
  "73596905192",
  "180450205124",
  "73596805100",
  "57295305106",
  "97217805118",
  "80814305179",
  "68452305192",
  "68452205100",
  "180756905147",
  "106795005109",
  "137061305171",
  "56688705130",
  "87312605145",
  "151272005143",
  "80210205179",
  "79315905121",
  "79315805136",
  "63911205170",
  "66330205110",
  "2305505194",
  "21901405192",
  "44670905139",
  "58480705168",
  "51008005164",
  "72072705164",
  "161909505155",
  "170561005110",
  "170560905158",
  "101400005175",
  "4167605165",
  "15587305191",
  "135250605105",
  "121199305132",
  "8070005177",
  "7202005104",
  "135250205143",
  "155479505116",
  "170878005155",
  "57294105130",
  "182274805123",
  "128710505165",
  "17088005196",
  "7201705158",
  "171472205148",
  "143352505183",
  "113381105112",
  "113381005127",
  "112478405100",
  "112478305101",
  "73278005100",
  "153079705137",
  "143054305172",
  "180163805100",
  "155181005199",
  "68147405125",
  "123897705194",
  "181043905132",
  "182272405100",
  "14985605104",
  "10463705140",
  "7773005103",
  "88515405190",
  "12855005168",
  "11369905109",
  "145361605177",
  "23103805113",
  "92705905110",
  "129915905116",
  "158136005102",
  "41047605170",
  "724705186",
  "180441505104",
  "156956605193",
  "156956505100",
  "122073105100",
  "160040605175",
  "123889305118",
  "73885405179",
  "163383405189",
  "23101705164",
  "22779305190",
  "71482905110",
  "161588405105",
  "164585505118",
  "93586805113",
  "17066905108",
  "72971205109",
  "60891905167",
  "180745205159",
  "180745105163",
  "153957405121",
  "169089205120",
  "17065305170",
  "110067905134",
  "36872105162",
  "65999705170",
  "4456305122",
  "101390305130",
  "164580305128",
  "143038605120",
  "11658305102",
  "181951005101",
  "129298705114",
  "142469005180",
  "74782905177",
  "29708405173",
  "41974905197",
  "164573505186",
  "54864705183",
  "117868005190",
  "157264005164",
  "170852805141",
  "178935205102",
  "144492505191",
  "127185505143",
  "125092705180",
  "7180005159",
  "128096905151",
  "91490205181",
  "163944605136",
  "181328805106",
  "38058605160",
  "41971405194",
  "170843505189",
  "154255305181",
  "90903005154",
  "156336405189",
  "110056305131",
  "31786805167",
  "89401305195",
  "4445805120",
  "99265805127",
  "20666005124",
  "151224905165",
  "153632305197",
  "153632105106",
  "46429705106",
  "117268905121",
  "93297905175",
  "14350505108",
  "21565905158",
  "156937805100",
  "89396805140",
  "106745005135",
  "110047705107",
  "113030705129",
  "113359005167",
  "119376305160",
  "113357305170",
  "128680005150",
  "181934205189",
  "49153005170",
  "20045305195",
  "152132105132",
  "130698005193",
  "23979905113",
  "32084205160",
  "84677305148",
  "89385005167",
  "51858605184",
  "30280505126",
  "101373605106",
  "29063805113",
  "160951205134",
  "140940705147",
  "103761405153",
  "4120305123",
  "178621705168",
  "22732005159",
  "89367005184",
  "117833105151",
  "169621005133",
  "81678805114",
  "121458405147",
  "74770205100",
  "65061305168",
  "117822205197",
  "130684205179",
  "158998605151",
  "124482705100",
  "126837305156",
  "1626524",
  "171409805156",
  "123242905118",
  "129871605197",
  "31156405106",
  "157503205118",
  "155977905137",
  "84376605149",
  "13715905180",
  "180681905198",
  "23645205140",
  "40126505160",
  "17623505182",
  "154800605165",
  "115744405101",
  "109185905122",
  "71756205189",
  "97493505129",
  "157496205152",
  "89982805122",
  "55411005131",
  "90853805184",
  "82865605102",
  "25772805138",
  "160927605172",
  "32936305129",
  "48226005105",
  "47913405144",
  "156602805192",
  "87868105162",
  "59348305154",
  "40117305193",
  "64774705134",
  "19995105197",
  "87863205149",
  "21243805169",
  "44879305151",
  "164849805112",
  "8596305106",
  "26906105102",
  "162113305109",
  "151502305109",
  "112139105126",
  "42490705163",
  "63855905164",
  "83799805119",
  "88479305109",
  "137896505192",
  "98337605101",
  "63842105117",
  "172258705171",
  "22687705105",
  "64761305159",
  "164819205155",
  "31136105180",
  "170158405129",
  "160892305130",
  "159254805121",
  "145268205113",
  "42184705129",
  "4372805134",
  "150111705145",
  "2339906",
  "110313705159",
  "25083205106",
  "98896105111",
  "65021505156",
  "139974505159",
  "114757605155",
  "70180605136",
  "149171105182",
  "35006805153",
  "181843505187",
  "169315705137",
  "167205105140",
  "169288805110",
  "168705205109",
  "94427705167",
  "161169005120",
  "34689405149",
  "12734205156",
  "37658705181",
  "119001505130",
  "2611984",
  "185438205107",
  "2656959",
  "20261705108",
  "2659601",
  "145558105143",
  "181502205142",
  "167767505127",
  "2713805",
  "38913505143",
  "2765031",
  "54241805102",
  "163212205144",
  "167728305135",
  "168642005100",
  "100072705175",
  "165649505178",
  "165649505178",
  "2916984",
  "95020605104",
  "164431805131",
  "2934674",
  "2936673",
  "113865805154",
  "107334405140",
  "11827605101",
  "89431905152",
  "148525805184",
  "190231505154",
  "136859405101",
  "2995984",
  "11522005153",
  "53530605104",
  "92869005121",
  "64719505123",
  "100412205137",
  "87539605193",
  "170688205185",
  "69581605117",
  "38233505148",
  "139331805100",
  "20548705115",
  "20231705180",
  "159790805100",
  "97442605171",
  "14816205140",
  "17222905104",
  "2214205140",
  "22397805172",
  "98285905108",
  "92285405125",
  "31980105167",
  "43332905165",
  "9127305170",
  "122877005159",
  "118380605162",
  "89013705130",
  "179064605184",
  "44791905121",
  "53516905112",
  "7948705165",
  "27728705184",
  "38225105171",
  "171568905163",
  "104674005104",
  "121387805187",
  "123166305184",
  "75325705108",
  "154990005130",
  "34967005160",
  "111796005150",
  "89258205104",
  "46044805180",
  "159474405135",
  "9733605175",
  "4321005168",
  "97098505180",
  "9473605174",
  "8851605161",
  "16616705180",
  "11808005191",
  "112282805117",
  "32854305147",
  "27402605168",
  "48141005188",
  "22611805190",
  "10926205169",
  "121707705164",
  "153163205185",
  "69159405118",
  "79202105110",
  "179052505157",
  "73502805116",
  "13290805149",
  "107920005160",
  "10008405196",
  "110828105106",
  "158573005108",
  "71109405167",
  "29602505103",
  "37917905181",
  "153773805183",
  "7929105158",
  "7059205197",
  "7315205160",
  "129726505185",
  "180833405162",
  "156772005130",
  "93446505190",
  "149959205192",
  "114126905148",
  "153460605195",
  "61050305194",
  "97733305105",
  "3341371",
  "19632105119",
  "58947605175",
  "21185305106",
  "171244905109",
  "170934305176",
  "25722405105",
  "84585205105",
  "89244605170",
  "92826705121",
  "95001305160",
  "25721505119",
  "66203305172",
  "160176605145",
  "142859405122",
  "142859305137",
  "180554305101",
  "169746905172",
  "88420905141",
  "112869505154",
  "181124405185",
  "131795605106",
  "77650905197",
  "61350505169",
  "101003105158",
  "61350105100",
  "49633305113",
  "157950805134",
  "9995105140",
  "9452805127",
  "90408305117",
  "103387605113",
  "11788405174",
  "155259305160",
  "133550005144",
  "122859605119",
  "20842105175",
  "144621905193",
  "169736905107",
  "28623105119",
  "182028705175",
  "178741305107",
  "91339405108",
  "82768005138",
  "134798305148",
  "106094305149",
  "17189205113",
  "3498071",
  "80951205195",
  "38190205103",
  "34075005107",
  "155844505140",
  "3545512",
  "146386505198",
  "3398405159",
  "119565605105",
  "356205159",
  "3633305",
  "130345505140",
  "3635126",
  "185896705198",
  "181029905187",
  "281170105186",
  "281338705110",
  "281406005168",
  "281435705139",
  "281469205123",
  "281580305193",
  "3762346",
  "92807105114",
  "37265905137",
  "144378605167",
  "148159505104",
  "24963105181",
  "123125105129",
  "126974105187",
  "163738105109",
]);

const UCS_ADM = new Set([
  "156087405138",
  "140660305138",
  "160703905146",
  "41291305106",
  "49196705147",
  "153079805122",
  "180464105134",
  "179883905121",
  "179351305188",
  "171541705163",
  "2233618",
  "159675205165",
  "159393105169",
  "158800405100",
  "158135905137",
  "157680905150",
  "148993805186",
  "140891605130",
  "117316005142",
  "116399605165",
  "117010305134",
  "115802805181",
  "156957605175",
  "224868905127",
  "280804105197",
  "118499105184",
  "102449105140",
  "101293205126",
  "100422905143",
  "99305705167",
  "89248505102",
  "89129705164",
  "75885305195",
  "65131905150",
  "62895505134",
  "44976605147",
  "43167505100",
  "42569205192",
  "41883505157",
  "37473705190",
  "34120105184",
  "26505605117",
  "60318605165",
  "60024105165",
  "22056905102",
  "16872205163",
  "23891505189",
  "3822405173",
  "826505123",
  "565605136",
  "65714305194",
  "3242",
  "96842605179",
  "148089205185",
  "16167605114",
  "162836505110",
]);

const extractUcFromFileName = (fileName: string, extractedUc: string): string => {
  if (!fileName) return extractedUc || "";
  const cleanName = fileName.replace(/\.[^/.]+$/, "");
  
  const knownUcs = new Set<string>();
  
  if (typeof UCS_LIVRE_MERCADO_LIVRE !== "undefined") {
    UCS_LIVRE_MERCADO_LIVRE.forEach((uc) => knownUcs.add(String(uc)));
  }
  if (typeof UCS_OPER !== "undefined") {
    UCS_OPER.forEach((uc) => knownUcs.add(String(uc)));
  }
  if (typeof UCS_ADM !== "undefined") {
    UCS_ADM.forEach((uc) => knownUcs.add(String(uc)));
  }
  
  try {
    const saved = localStorage.getItem("sanesul_uc_mappings");
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        parsed.forEach((m) => {
          if (m && m.uc) {
            knownUcs.add(String(m.uc));
          }
        });
      }
    }
  } catch (e) {}

  const sortedKnownUcs = Array.from(knownUcs).sort((a, b) => b.length - a.length);
  
  for (const uc of sortedKnownUcs) {
    if (uc.length >= 4 && cleanName.includes(uc)) {
      return uc;
    }
  }
  
  const fileNameNumbers = cleanName.replace(/\D/g, "");
  for (const uc of sortedKnownUcs) {
    if (uc.length >= 4 && fileNameNumbers.includes(uc)) {
      return uc;
    }
  }
  
  if (fileNameNumbers.length >= 5) {
    return fileNameNumbers;
  }
  
  return extractedUc || fileNameNumbers;
};

const EXCEL_COLUMNS = [
  { header: "UC", key: "uc" },
  { header: "Tipo", key: "tipo" },
  { header: "Concessionária", key: "concessionaria" },
  { header: "Cidade", key: "cidade" },
  { header: "Gerência", key: "gerencia" },
  { header: "LOCINS", key: "locin" },
  { header: "Mês Referência", key: "mesReferencia" },
  { header: "Ano Leitura", key: "anoLeitura" },
  { header: "Vencimento", key: "dataVencimento" },
  { header: "Nota Fiscal", key: "numeroNotaFiscal" },
  { header: "Modalidade Tarifária", key: "modalidadeTarifaria" },
  { header: "Subgrupo", key: "subgrupo" },
  { header: "Valor Total (R$)", key: "valorTotal" },
  { header: "Demanda Contratada Ponta (kW)", key: "demandaPontaKW" },
  { header: "Demanda Contratada Fora Ponta (kW)", key: "demandaForaPontaKW" },
  { header: "Demanda Medida Ponta (kW)", key: "demandaPotenciaMedidaPonta" },
  {
    header: "Valor Demanda Medida Ponta (R$)",
    key: "valorDemandaPotenciaMedidaPonta",
  },
  {
    header: "Demanda Medida Fora Ponta (kW)",
    key: "demandaPotenciaMedidaForaPonta",
  },
  {
    header: "Valor Demanda Medida Fora Ponta (R$)",
    key: "valorDemandaPotenciaMedidaForaPonta",
  },
  { header: "Consumo Ponta (kWh)", key: "consumoKwhPonta" },
  { header: "Valor Consumo Ponta (R$)", key: "valorConsumoKwhPonta" },
  { header: "Consumo Fora Ponta (kWh)", key: "consumoKwhForaPonta" },
  { header: "Valor Consumo Fora Ponta (R$)", key: "valorConsumoKwhForaPonta" },
  {
    header: "Demanda Não Consumida Ponta (kW)",
    key: "demandaPotenciaNaoConsumidaPonta",
  },
  {
    header: "Valor Demanda Não Consumida Ponta (R$)",
    key: "valorDemandaPotenciaNaoConsumidaPonta",
  },
  {
    header: "Demanda Não Consumida Fora Ponta (kW)",
    key: "demandaPotenciaNaoConsumidaFPonta",
  },
  {
    header: "Valor Demanda Não Consumida Fora Ponta (R$)",
    key: "valorDemandaPotenciaNaoConsumidaFPonta",
  },
  {
    header: "Ultrapassagem Ponta (kW)",
    key: "demandaPotenciaAtivaUltrapPonta",
  },
  {
    header: "Valor Ultrapassagem Ponta (R$)",
    key: "valorDemandaPotenciaAtivaUltrapPonta",
  },
  {
    header: "Ultrapassagem Fora Ponta (kW)",
    key: "demandaPotenciaAtivaUltrapFPonta",
  },
  {
    header: "Valor Ultrapassagem Fora Ponta (R$)",
    key: "valorDemandaPotenciaAtivaUltrapFPonta",
  },
  {
    header: "Reativa Excedente Ponta (kVArh)",
    key: "energiaReativaExcedPonta",
  },
  {
    header: "Valor Reativa Excedente Ponta (R$)",
    key: "valorEnergiaReativaExcedPonta",
  },
  {
    header: "Reativa Excedente Fora Ponta (kVArh)",
    key: "energiaReativaExcedFPonta",
  },
  {
    header: "Valor Reativa Excedente Fora Ponta (R$)",
    key: "valorEnergiaReativaExcedFPonta",
  },
  { header: "GDI oUC (kWh)", key: "energiaAtvInjetadaGDIOUC" },
  { header: "Valor GDI oUC (R$)", key: "valorEnergiaAtvInjetadaGDIOUC" },
  { header: "GDI mUC (kWh)", key: "energiaAtvInjetadaGDIMUC" },
  { header: "Valor GDI mUC (R$)", key: "valorEnergiaAtvInjetadaGDIMUC" },
  { header: "CIP (R$)", key: "cip" },
  { header: "Outros Encargos (R$)", key: "outrosEncargos" },
  { header: "PIS (R$)", key: "pis" },
  { header: "COFINS (R$)", key: "cofins" },
  { header: "ICMS (R$)", key: "icms" },
];

const EXTRACTION_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    uc: {
      type: Type.STRING,
      description:
        "Código da Unidade Consumidora (UC). Para ENERGISA, use o 'CÓDIGO DO CLIENTE' completo ou a 'MATRÍCULA'. Extraia apenas os números e caracteres identificadores (ex: 10/1069-4).",
    },
    demandaPontaKW: {
      type: Type.STRING,
      description:
        "Demanda contratada na Ponta em kW. Geralmente encontrada na seção 'Grandezas Contratadas' ou 'Dados da Unidade Consumidora'.",
    },
    demandaForaPontaKW: {
      type: Type.STRING,
      description:
        "Demanda contratada Fora Ponta em kW. Para UCs da ELEKTRO na modalidade VERDE (somente para elas), o valor da demanda contratada FORA PONTA é o valor de 'Demanda Todos os Períodos'. NÃO APLIQUE ESTA REGRA para ENERGISA.",
    },
    demandaPotenciaMedidaPonta: {
      type: Type.STRING,
      description:
        "Demanda de Potência Medida no horário de Ponta (kW). Procure na tabela de 'Itens de Fatura'.",
    },
    demandaPotenciaMedidaForaPonta: {
      type: Type.STRING,
      description:
        "Demanda de Potência Medida no horário Fora Ponta (kW). Para ELEKTRO modalidade Verde, caso não haja a discriminação Fora Ponta, utilize o valor de 'DEMANDA kW'. NÃO APLIQUE ESTA REGRA para ENERGISA.",
    },
    anoLeitura: {
      type: Type.STRING,
      description:
        "Ano de referência da fatura (ex: 2025). Extraia apenas os 4 dígitos do ano.",
    },
    mesReferencia: {
      type: Type.STRING,
      description:
        "Mês de referência da fatura (ex: Agosto). Extraia apenas o nome do mês, sem o ano.",
    },
    consumoGrupoB: {
      type: Type.STRING,
      description: "Quantidade de consumo em kWh para o Grupo B (se aplicável).",
    },
    demandaTodosPeriodosKW: {
      type: Type.STRING,
      description: "Demanda de Potência para todos os períodos (kW).",
    },
    consumoKwhPonta: {
      type: Type.STRING,
      description:
        "Quantidade de consumo em kWh no horário de Ponta. Procure por 'Consumo Ponta' ou 'Consumo Ativo Ponta'.",
    },
    valorConsumoKwhPonta: {
      type: Type.STRING,
      description: "Valor total em R$ do consumo no horário de Ponta.",
    },
    consumoKwhForaPonta: {
      type: Type.STRING,
      description:
        "Quantidade de consumo em kWh no horário Fora Ponta. Procure por 'Consumo Fora Ponta' ou 'Consumo Ativo Fora Ponta'.",
    },
    valorConsumoKwhForaPonta: {
      type: Type.STRING,
      description: "Valor total em R$ do consumo no horário Fora Ponta.",
    },
    valorTotal: {
      type: Type.STRING,
      description:
        "Valor total da fatura a pagar (R$). Geralmente em destaque.",
    },
    cidade: {
      type: Type.STRING,
      description: "Cidade onde se localiza a Unidade Consumidora.",
    },
    demandaPotenciaNaoConsumidaPonta: {
      type: Type.STRING,
      description: "Demanda de Potência Não Consumida - Ponta (kW).",
    },
    demandaPotenciaNaoConsumidaFPonta: {
      type: Type.STRING,
      description: "Demanda de Potência Não Consumida - Fora Ponta (kW).",
    },
    demandaPotenciaAtivaUltrapPonta: {
      type: Type.STRING,
      description: "Demanda de Potência Ativa - Ultrapassagem - Ponta (kW).",
    },
    demandaPotenciaAtivaUltrapFPonta: {
      type: Type.STRING,
      description:
        "Demanda de Potência Ativa - Ultrapassagem - Fora Ponta (kW).",
    },
    energiaReativaExcedPonta: {
      type: Type.STRING,
      description: "Energia Reativa Excedente - Ponta (kVArh).",
    },
    energiaReativaExcedFPonta: {
      type: Type.STRING,
      description: "Energia Reativa Excedente - Fora Ponta (kVArh).",
    },
    valorDemandaPotenciaMedidaPonta: {
      type: Type.STRING,
      description: "Valor em R$ da Demanda de Potência Medida - Ponta.",
    },
    valorDemandaPotenciaMedidaForaPonta: {
      type: Type.STRING,
      description: "Valor em R$ da Demanda de Potência Medida - Fora Ponta.",
    },
    valorDemandaPotenciaNaoConsumidaPonta: {
      type: Type.STRING,
      description: "Valor em R$ da Demanda Potência Não Consumida - Ponta.",
    },
    valorDemandaPotenciaNaoConsumidaFPonta: {
      type: Type.STRING,
      description:
        "Valor em R$ da Demanda Potência Não Consumida - Fora Ponta.",
    },
    valorDemandaPotenciaAtivaUltrapPonta: {
      type: Type.STRING,
      description:
        "Valor em R$ da Demanda Potência Ativa - Ultrapassagem - Ponta.",
    },
    valorDemandaPotenciaAtivaUltrapFPonta: {
      type: Type.STRING,
      description:
        "Valor em R$ da Demanda Potência Ativa - Ultrapassagem - Fora Ponta.",
    },
    valorEnergiaReativaExcedPonta: {
      type: Type.STRING,
      description: "Valor em R$ da Energia Reativa Excedente - Ponta.",
    },
    valorEnergiaReativaExcedFPonta: {
      type: Type.STRING,
      description: "Valor em R$ da Energia Reativa Excedente - Fora Ponta.",
    },
    energiaAtvInjetadaGDIOUC: {
      type: Type.STRING,
      description:
        "Energia Ativa Injetada GDI oUC (kWh). IMPORTANTE: Se houver mais de um valor para este item na fatura, SOME todos os valores em um único valor.",
    },
    valorEnergiaAtvInjetadaGDIOUC: {
      type: Type.STRING,
      description:
        "Valor em R$ da Energia Ativa Injetada GDI oUC. IMPORTANTE: Este valor é frequentemente negativo na fatura (ex: -3.822,92). Extraia EXATAMENTE como aparece, incluindo o sinal negativo. Se houver mais de um valor para este item na fatura, SOME todos os valores em um único valor.",
    },
    energiaAtvInjetadaGDIMUC: {
      type: Type.STRING,
      description:
        "Energia Ativa Injetada GDI mUC (kWh). IMPORTANTE: Se houver mais de um valor para este item na fatura, SOME todos os valores em um único valor.",
    },
    valorEnergiaAtvInjetadaGDIMUC: {
      type: Type.STRING,
      description:
        "Valor em R$ da Energia Ativa Injetada GDI mUC. IMPORTANTE: Este valor é frequentemente negativo na fatura (ex: -124,39). Extraia EXATAMENTE como aparece, incluindo o sinal negativo. Se houver mais de um valor para este item na fatura, SOME todos os valores em um único valor.",
    },
    cip: {
      type: Type.STRING,
      description:
        "Valor em R$ da Contribuição de Iluminação Pública (CIP ou COSIP).",
    },
    outrosEncargos: {
      type: Type.STRING,
      description:
        "Soma de outros encargos, multas, juros ou adicionais de bandeira tarifária.",
    },
    pis: { type: Type.STRING, description: "Valor em R$ do PIS." },
    cofins: { type: Type.STRING, description: "Valor em R$ do COFINS." },
    icms: { type: Type.STRING, description: "Valor em R$ do ICMS." },
    concessionaria: {
      type: Type.STRING,
      description:
        "Nome da empresa concessionária (ex: ENERGISA, ELEKTRO, CPFL).",
    },
    numeroNotaFiscal: {
      type: Type.STRING,
      description: "Número da Nota Fiscal ou Número da Fatura.",
    },
    dataVencimento: {
      type: Type.STRING,
      description: "Data de vencimento da fatura (ex: 15/08/2025).",
    },
    modalidadeTarifaria: {
      type: Type.STRING,
      description:
        "Modalidade Tarifária (ex: AZUL, VERDE, BRANCA, CONVENCIONAL).",
    },
    subgrupo: {
      type: Type.STRING,
      description: "Subgrupo tarifário (ex: A4, B1, B3).",
    },
  },
  required: ["uc", "anoLeitura", "mesReferencia"],
};

const AGRUPADORA_SCHEMA = {
  type: Type.OBJECT,
  properties: {
    concessionaria: {
      type: Type.STRING,
      description: "Nome da concessionária (ex: ELEKTRO, ENERGISA)",
    },
    valorTotal: {
      type: Type.STRING,
      description: "Valor total da fatura (R$)",
    },
    mesReferencia: {
      type: Type.STRING,
      description: "Mês/Ano de referência (ex: Fevereiro/2026)",
    },
    vencimento: { type: Type.STRING, description: "Data de vencimento" },
    numeroNotaFiscal: {
      type: Type.STRING,
      description: "Número da Nota Fiscal ou Fatura (ex: AGP-01...)",
    },
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

const Logo = ({
  className = "h-10",
  showText = true,
  isLogin = false,
}: {
  className?: string;
  showText?: boolean;
  isLogin?: boolean;
}) => {
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
    <div
      className={`flex ${isLogin ? "flex-col items-center" : "items-center gap-3"} ${className}`}
    >
      <div
        className={`${isLogin ? "w-16 h-16 rounded-2xl mb-4" : className.includes("h-12") ? "w-12 h-12 rounded-xl" : "w-10 h-10 rounded-xl"} bg-sanesul-primary flex items-center justify-center shadow-lg shadow-sanesul-primary/20 shrink-0`}
      >
        <Zap
          className="text-white"
          size={isLogin ? 32 : className.includes("h-12") ? 24 : 20}
        />
      </div>
      {showText && (
        <div className={isLogin ? "text-center" : ""}>
          <h1
            className={`${isLogin ? "text-3xl" : className.includes("h-12") ? "text-3xl md:text-4xl" : "text-2xl"} font-display font-bold tracking-tight text-sanesul-primary leading-none`}
          >
            Sanesul <span className="text-sanesul-secondary">Energy</span>
          </h1>
          <p className="text-[9px] uppercase tracking-[0.2em] text-sanesul-muted font-bold mt-1">
            {isLogin ? "Acesso Restrito" : "Portal de Inteligência Energética"}
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
    if (
      bill.status === "completed" &&
      bill.uc &&
      bill.mesReferencia &&
      bill.anoLeitura
    ) {
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
  if (
    window.aistudio &&
    typeof window.aistudio.hasSelectedApiKey === "function"
  ) {
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
  delay = 5000,
): Promise<GenerateContentResponse> => {
  try {
    // Add a timeout of 120 seconds to the API call
    let timeoutId: NodeJS.Timeout;
    const timeoutPromise = new Promise((_, reject) => {
      timeoutId = setTimeout(
        () =>
          reject(
            new Error(
              "TIMEOUT_API: A API demorou muito para responder (120s).",
            ),
          ),
        120000,
      );
    });

    try {
      const response = (await Promise.race([
        ai.models.generateContent(params),
        timeoutPromise,
      ])) as GenerateContentResponse;

      clearTimeout(timeoutId!);
      return response;
    } catch (error) {
      clearTimeout(timeoutId!);
      throw error;
    }
  } catch (error: any) {
    // Extract error details
    let errorStr = "";
    let errorCode = 0;
    let errorStatus = "";

    if (typeof error === "string") {
      errorStr = error;
    } else if (error && typeof error === "object") {
      // Handle the specific format provided by the user
      const nestedError = error.error || error;
      errorCode = nestedError.code || error.status || 0;
      errorStatus = nestedError.status || "";
      errorStr = nestedError.message || error.message || JSON.stringify(error);
    }

    console.error("generateContentWithRetry Error:", {
      params: JSON.stringify(params),
      errorCode,
      errorStatus,
      errorStr,
      retries,
    });

    const isTransientError =
      errorCode === 429 ||
      errorCode === 500 ||
      errorCode === 502 ||
      errorCode === 503 ||
      errorCode === 504 ||
      errorStatus === "RESOURCE_EXHAUSTED" ||
      errorStatus === "INTERNAL" ||
      errorStatus === "UNAVAILABLE" ||
      errorStr.includes("429") ||
      errorStr.includes("500") ||
      errorStr.includes("502") ||
      errorStr.includes("503") ||
      errorStr.includes("504") ||
      errorStr.includes("RESOURCE_EXHAUSTED") ||
      errorStr.includes("INTERNAL") ||
      errorStr.includes("UNAVAILABLE");

    const isTimeout = errorStr.includes("TIMEOUT_API");
    const isLockError = errorStr.includes("Lock broken by another request");
    const isHardQuota =
      errorStr.includes("spending cap") || errorStr.includes("monthly limit") || errorStr.includes("prepayment credits are depleted") || errorStr.includes("Failed to fetch");
    const isRateLimit =
      errorCode === 429 ||
      errorStatus === "RESOURCE_EXHAUSTED" ||
      errorStr.includes("429") ||
      errorStr.includes("RESOURCE_EXHAUSTED");
    const isExpired =
      errorStr.includes("API key expired") ||
      errorStr.includes("API_KEY_INVALID") ||
      errorStr.includes("expired");
    const isInvalid =
      errorStr.includes("invalid API key") ||
      errorStr.includes("invalid key") ||
      (errorCode === 401 && errorStr.includes("invalid"));
    const isNotFound =
      errorStr.includes("Requested entity was not found") ||
      errorStr.includes("API key not found");

    if (isNotFound || isExpired || isInvalid) {
      const msg = isNotFound
        ? 'Chave de API não encontrada ou inválida. Por favor, use o botão "Trocar Conta" para selecionar uma chave válida.'
        : 'A chave da API expirou ou é inválida. Por favor, use o botão "Trocar Conta" para selecionar uma nova chave.';
      throw new Error(msg);
    }

    if (
      retries > 0 &&
      (isTransientError || isTimeout || isLockError || isRateLimit) &&
      !isHardQuota
    ) {
      console.warn(
        `${isTimeout ? "Timeout" : isLockError ? "Lock error" : isRateLimit ? "Rate limit" : "Transient error (" + errorCode + ")"} hit, retrying in ${delay}ms... (${retries} retries left)`,
      );
      await new Promise((resolve) => setTimeout(resolve, delay));
      return generateContentWithRetry(ai, params, retries - 1, delay * 2);
    }

    // If it's a quota error or we're out of retries, throw
    if (isHardQuota || isRateLimit || isTransientError) {
      if (isHardQuota || isRateLimit) {
        const msg = isHardQuota
          ? "O limite de gastos do seu projeto foi atingido. Verifique sua conta do Google Cloud (https://ai.google.dev/gemini-api/docs/billing)."
          : "Cota da API excedida ou limite de taxa atingido. Verifique seu plano e detalhes de faturamento no Google AI Studio (https://ai.google.dev/gemini-api/docs/billing). " +
            errorStr;
        const quotaError = new Error(msg);
        (quotaError as any).isQuotaError = true;
        throw quotaError;
      }
    }

    throw error;
  }
};

const parseValue = (val: string | number) => {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return val;

  let str = val.toString().trim();

  const lastDot = str.lastIndexOf(".");
  const lastComma = str.lastIndexOf(",");

  if (lastComma > lastDot) {
    // Brazilian format: 1.234,56
    str = str.replace(/\./g, "").replace(",", ".");
  } else if (lastDot > lastComma && lastComma !== -1) {
    // US format: 1,234.56
    str = str.replace(/,/g, "");
  } else if (lastComma !== -1) {
    // Only comma: 1234,56
    str = str.replace(",", ".");
  }
  // If only dot, it's already in US format: 1234.56

  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
};

const formatNumber = (
  val: number,
  isCurrency: boolean = false,
  precision: number = 2,
) => {
  return val.toLocaleString("pt-BR", {
    minimumFractionDigits: isCurrency ? precision : 0,
    maximumFractionDigits: precision,
  });
};

const formatMonth = (month: string | number) => {
  if (!month) return "";
  let normalized = month.toString().toLowerCase().trim();

  // Se vier no formato MM/YYYY, pega apenas o mês
  if (normalized.includes("/")) {
    normalized = normalized.split("/")[0];
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

const getMonthNumber = (month: string | number) => {
  if (!month) return 0;
  const formatted = formatMonth(month).toLowerCase();
  const monthOrder: Record<string, number> = {
    janeiro: 1,
    fevereiro: 2,
    março: 3,
    marco: 3,
    abril: 4,
    maio: 5,
    junho: 6,
    julho: 7,
    agosto: 8,
    setembro: 9,
    outubro: 10,
    novembro: 11,
    dezembro: 12,
    "01": 1,
    "02": 2,
    "03": 3,
    "04": 4,
    "05": 5,
    "06": 6,
    "07": 7,
    "08": 8,
    "09": 9,
    "10": 10,
    "11": 11,
    "12": 12,
    "1": 1,
    "2": 2,
    "3": 3,
    "4": 4,
    "5": 5,
    "6": 6,
    "7": 7,
    "8": 8,
    "9": 9,
  };
  return monthOrder[formatted] || parseInt(formatted, 10) || 0;
};

const formatReference = (ref: string) => {
  if (!ref) return "";
  if (ref.includes("/")) {
    const parts = ref.split("/");
    if (parts.length === 2) {
      return `${formatMonth(parts[0])}/${parts[1]}`;
    }
  }
  return formatMonth(ref);
};

const mapDbToBillData = (dbBill: any): BillData => {
  let mod = dbBill.modalidade_tarifaria || "";
  let tipo = dbBill.tipo || "";
  if (UCS_OPER.has(String(dbBill.uc))) {
    tipo = "OPER";
  } else if (UCS_ADM.has(String(dbBill.uc))) {
    tipo = "ADM";
  } else if (UCS_LIVRE_MERCADO_LIVRE.has(String(dbBill.uc))) {
    tipo = "LIVRE";
    if (!mod.toUpperCase().includes("LIVRE")) {
      mod = mod ? `${mod} - LIVRE` : "LIVRE";
    }
  }

  return {
    id: dbBill.id,
    fileName: dbBill.file_name,
    uc: dbBill.uc || "",
    demandaPontaKW: dbBill.demanda_ponta_kw || "",
    demandaForaPontaKW: dbBill.demanda_fora_ponta_kw || "",
    demandaPotenciaMedidaPonta: dbBill.demanda_potencia_medida_ponta || "",
    demandaPotenciaMedidaForaPonta:
      dbBill.demanda_potencia_medida_fora_ponta || "",
    anoLeitura: dbBill.ano_leitura || "",
    mesReferencia: dbBill.mes_referencia || "",
    consumoGrupoB: dbBill.consumo_grupo_b || "",
    demandaTodosPeriodosKW: dbBill.demanda_todos_periodos_kw || "",
    consumoKwhPonta: dbBill.consumo_kwh_ponta || "",
    consumoKwhForaPonta: dbBill.consumo_kwh_fora_ponta || "",
    consumoKwhGrupoB: dbBill.consumo_kwh_grupo_b || dbBill.consumo_grupo_b || "",
    valorConsumoKwhPonta: dbBill.valor_consumo_kwh_ponta || "",
    valorConsumoKwhForaPonta: dbBill.valor_consumo_kwh_fora_ponta || "",
    valorConsumoKwhGrupoB: dbBill.valor_consumo_kwh_grupo_b || "",
    valorTotal: dbBill.valor_total || "",
    cidade: dbBill.cidade || "",
    demandaPotenciaNaoConsumidaPonta:
      dbBill.demanda_potencia_nao_consumida_ponta || "",
    demandaPotenciaNaoConsumidaFPonta:
      dbBill.demanda_potencia_nao_consumida_f_ponta || "",
    demandaPotenciaAtivaUltrapPonta:
      dbBill.demanda_potencia_ativa_ultrap_ponta || "",
    demandaPotenciaAtivaUltrapFPonta:
      dbBill.demanda_potencia_ativa_ultrap_f_ponta || "",
    energiaReativaExcedPonta: dbBill.energia_reativa_exced_ponta || "",
    energiaReativaExcedFPonta: dbBill.energia_reativa_exced_f_ponta || "",
    valorDemandaPotenciaMedidaPonta:
      dbBill.valor_demanda_potencia_medida_ponta || "",
    valorDemandaPotenciaMedidaForaPonta:
      dbBill.valor_demanda_potencia_medida_fora_ponta || "",
    valorDemandaPotenciaNaoConsumidaPonta:
      dbBill.valor_demanda_potencia_nao_consumida_ponta || "",
    valorDemandaPotenciaNaoConsumidaFPonta:
      dbBill.valor_demanda_potencia_nao_consumida_f_ponta || "",
    valorDemandaPotenciaAtivaUltrapPonta:
      dbBill.valor_demanda_potencia_ativa_ultrap_ponta || "",
    valorDemandaPotenciaAtivaUltrapFPonta:
      dbBill.valor_demanda_potencia_ativa_ultrap_f_ponta || "",
    valorEnergiaReativaExcedPonta:
      dbBill.valor_energia_reativa_exced_ponta || "",
    valorEnergiaReativaExcedFPonta:
      dbBill.valor_energia_reativa_exced_f_ponta || "",
    energiaAtvInjetadaGDIOUC: dbBill.energia_atv_injetada_gdi_ouc || "",
    valorEnergiaAtvInjetadaGDIOUC:
      dbBill.valor_energia_atv_injetada_gdi_ouc || "",
    energiaAtvInjetadaGDIMUC: dbBill.energia_atv_injetada_gdi_muc || "",
    valorEnergiaAtvInjetadaGDIMUC:
      dbBill.valor_energia_atv_injetada_gdi_muc || "",
    cip: dbBill.cip || "",
    outrosEncargos: dbBill.outros_encargos || "",
    pis: dbBill.pis || "",
    cofins: dbBill.cofins || "",
    icms: dbBill.icms || "",
    concessionaria: dbBill.concessionaria || "",
    numeroNotaFiscal: dbBill.numero_nota_fiscal || "",
    modalidadeTarifaria: mod,
    subgrupo: dbBill.subgrupo || "",
    tipo: tipo,
    mercado:
      dbBill.mercado ||
      (UCS_LIVRE_MERCADO_LIVRE.has(String(dbBill.uc)) ? "LIVRE" : "CATIVO"),
    gerencia: dbBill.gerencia || "",
    locin: dbBill.locin || dbBill.locins || "",
    dataVencimento: dbBill.data_vencimento || "",
    status: dbBill.status as any,
    error: dbBill.error || undefined,
    createdAt: dbBill.created_at
      ? new Date(dbBill.created_at).getTime()
      : Date.now(),
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
  consumo_grupo_b: bill.consumoGrupoB,
  demanda_todos_periodos_kw: bill.demandaTodosPeriodosKW,
  consumo_kwh_ponta: bill.consumoKwhPonta,
  consumo_kwh_fora_ponta: bill.consumoKwhForaPonta,
  consumo_kwh_grupo_b: bill.consumoKwhGrupoB,
  valor_consumo_kwh_ponta: bill.valorConsumoKwhPonta,
  valor_consumo_kwh_fora_ponta: bill.valorConsumoKwhForaPonta,
  valor_consumo_kwh_grupo_b: bill.valorConsumoKwhGrupoB,
  valor_total: bill.valorTotal,
  cidade: bill.cidade,
  demanda_potencia_nao_consumida_ponta: bill.demandaPotenciaNaoConsumidaPonta,
  demanda_potencia_nao_consumida_f_ponta:
    bill.demandaPotenciaNaoConsumidaFPonta,
  demanda_potencia_ativa_ultrap_ponta: bill.demandaPotenciaAtivaUltrapPonta,
  demanda_potencia_ativa_ultrap_f_ponta: bill.demandaPotenciaAtivaUltrapFPonta,
  energia_reativa_exced_ponta: bill.energiaReativaExcedPonta,
  energia_reativa_exced_f_ponta: bill.energiaReativaExcedFPonta,
  valor_demanda_potencia_medida_ponta: bill.valorDemandaPotenciaMedidaPonta,
  valor_demanda_potencia_medida_fora_ponta:
    bill.valorDemandaPotenciaMedidaForaPonta,
  valor_demanda_potencia_nao_consumida_ponta:
    bill.valorDemandaPotenciaNaoConsumidaPonta,
  valor_demanda_potencia_nao_consumida_f_ponta:
    bill.valorDemandaPotenciaNaoConsumidaFPonta,
  valor_demanda_potencia_ativa_ultrap_ponta:
    bill.valorDemandaPotenciaAtivaUltrapPonta,
  valor_demanda_potencia_ativa_ultrap_f_ponta:
    bill.valorDemandaPotenciaAtivaUltrapFPonta,
  valor_energia_reativa_exced_ponta: bill.valorEnergiaReativaExcedPonta,
  valor_energia_reativa_exced_f_ponta: bill.valorEnergiaReativaExcedFPonta,
  energia_atv_injetada_gdi_ouc: bill.energiaAtvInjetadaGDIOUC,
  valor_energia_atv_injetada_gdi_ouc: bill.valorEnergiaAtvInjetadaGDIOUC,
  energia_atv_injetada_gdi_muc: bill.energiaAtvInjetadaGDIMUC,
  valor_energia_atv_injetada_gdi_muc: bill.valorEnergiaAtvInjetadaGDIMUC,
  cip: bill.cip,
  outros_encargos: bill.outrosEncargos,
  pis: bill.pis || "",
  cofins: bill.cofins || "",
  icms: bill.icms || "",
  concessionaria: bill.concessionaria || "",
  numero_nota_fiscal: bill.numeroNotaFiscal || "",
  modalidade_tarifaria: bill.modalidadeTarifaria || "",
  subgrupo: bill.subgrupo || "",
  tipo: bill.tipo || "",
  mercado: bill.mercado || "",
  gerencia: bill.gerencia || "",
  locins: bill.locin || "",
  data_vencimento: bill.dataVencimento || "",
  status: bill.status,
  error: bill.error || null,
  user_id: userId,
});

// --- Components ---

const MetricCard = ({
  title,
  custo,
  consumo,
  isReference = false,
  rightElement,
  titleColorClass = "text-sanesul-primary",
}: {
  title: React.ReactNode;
  custo: number;
  consumo: number;
  isReference?: boolean;
  rightElement?: React.ReactNode;
  titleColorClass?: string;
}) => {
  const tarifaMedia = consumo > 0 ? custo / consumo : 0;
  const tarifaLabel = "Tarifa Média (R$/kWh)";
  const custoLabel = isReference ? "CUSTO (R$)" : "Custo (R$)";
  const consumoLabel = isReference ? "CONSUMO (kWh)" : "Consumo (kWh)";

  return (
    <div className="bg-white p-6 rounded-[24px] border border-sanesul-primary/10 shadow-lg hover:shadow-xl transition-all hover:border-sanesul-primary/30">
      <div className="flex justify-between items-center mb-4">
        <h3
          className={`text-sm font-display font-bold uppercase tracking-wider ${titleColorClass}`}
        >
          {title}
        </h3>
        {rightElement}
      </div>
      <div className="space-y-3">
        <div className="flex justify-between items-end border-b border-slate-100 pb-2">
          <span className="text-[10px] font-bold text-sanesul-muted uppercase tracking-wider">
            {custoLabel}
          </span>
          <span className="text-lg font-bold text-slate-800">
            R${" "}
            {custo.toLocaleString("pt-BR", {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })}
          </span>
        </div>
        <div className="flex justify-between items-end border-b border-slate-100 pb-2">
          <span className="text-[10px] font-bold text-sanesul-muted uppercase tracking-wider">
            {consumoLabel}
          </span>
          <span className="text-lg font-bold text-slate-800">
            {consumo.toLocaleString("pt-BR", { maximumFractionDigits: 0 })}
          </span>
        </div>
        <div className="flex justify-between items-end pt-1">
          <span className="text-[10px] font-bold text-sanesul-secondary uppercase tracking-wider">
            {tarifaLabel}
          </span>
          <span className="text-xl font-bold text-sanesul-secondary">
            R${" "}
            {tarifaMedia.toLocaleString("pt-BR", {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })}
          </span>
        </div>
      </div>
    </div>
  );
};

const UCS_PPP = new Set([
  "170878005155",
  "57294105130",
  "182274805123",
  "128710505165",
  "7201705158",
  "171472205148",
  "113381105112",
  "113381005127",
  "112478305101",
  "153079705137",
  "143054305172",
  "180163805100",
  "155181005199",
  "123897705194",
  "181043905132",
  "182272405100",
  "14985605104",
  "7773005103",
  "88515405190",
  "145361605177",
  "23103805113",
  "148993805186",
  "92705905110",
  "129915905116",
  "156957605175",
  "158136005102",
  "41047605170",
  "724705186",
  "180441505104",
  "156956605193",
  "156956505100",
  "122073105100",
  "160040605175",
  "123889305118",
  "73885405179",
  "23101705164",
  "22779305190",
  "71482905110",
  "17066905108",
  "72971205109",
  "60891905167",
  "180745205159",
  "180745105163",
  "153957405121",
  "17065305170",
  "110067905134",
  "36872105162",
  "65999705170",
  "143038605120",
  "129298705114",
  "29708405173",
  "164573505186",
  "117868005190",
  "144492505191",
  "7180005159",
  "41971405194",
  "154255305181",
  "156336405189",
  "110056305131",
  "31786805167",
  "89401305195",
  "151224905165",
  "153632305197",
  "117268905121",
  "110047705107",
  "113030705129",
  "113359005167",
  "113357305170",
  "128680005150",
  "181934205189",
  "49153005170",
  "130698005193",
  "23979905113",
  "84677305148",
  "51858605184",
  "30280505126",
  "29063805113",
  "140940705147",
  "103761405153",
  "4120305123",
  "178621705168",
  "89367005184",
  "169621005133",
  "81678805114",
  "121458405147",
  "74770205100",
  "130684205179",
  "123242905118",
  "129871605197",
  "84376605149",
  "13715905180",
  "180681905198",
  "17623505182",
  "115744405101",
  "109185905122",
  "71756205189",
  "97493505129",
  "157496205152",
  "160927605172",
  "32936305129",
  "47913405144",
  "59348305154",
  "87863205149",
  "21243805169",
  "44879305151",
  "8596305106",
  "26906105102",
  "112139105126",
  "42490705163",
  "63855905164",
  "88479305109",
  "137896505192",
  "98337605101",
  "63842105117",
  "172258705171",
  "64761305159",
  "170158405129",
  "2233618",
  "4372805134",
  "150111705145",
  "110313705159",
  "25083205106",
  "98896105111",
  "149171105182",
  "181843505187",
  "169315705137",
  "94427705167",
  "185438205107",
  "100422905143",
  "92869005121",
  "64719505123",
  "100412205137",
  "87539605193",
  "170688205185",
  "159790805100",
  "169736905107",
  "91339405108",
  "16872205163",
  "103217205180",
  "65131905150",
  "93022105109",
  "115802805181",
  "62131705107",
  "8396905117",
  "112499605106",
  "167009105148",
  "162517005105",
  "44976605147",
  "8390905126",
  "74799405160",
  "75395505121",
  "58205005124",
  "49794205103",
  "50106505161",
  "179585205163",
  "137971405106",
  "62120305195",
  "101410505103",
  "51602005150",
  "132263205107",
  "159393105169",
  "49789105159",
  "101405705112",
  "49789005163",
  "116098105132",
  "12896205113",
  "165205605195",
  "122098305147",
  "42569205192",
  "103804405124",
  "52508605180",
  "103202105107",
  "164910205183",
  "132555305179",
  "164909905180",
  "162806105134",
  "106496205175",
  "96601505180",
  "60318605165",
  "156087405138",
  "165805505187",
  "49196705147",
  "43167505100",
  "158475205185",
  "158475105190",
  "179883905121",
  "103507505171",
  "157885505127",
  "157885405131",
  "47375405160",
  "20092805168",
  "19211605190",
  "13489605155",
  "99305705167",
  "172675705120",
  "7495805153",
  "153979205150",
  "31513505109",
  "166101005101",
  "166100905149",
  "180464105134",
  "53694705118",
  "10181105197",
  "54899305145",
  "54899205150",
  "160703905146",
  "141263205190",
  "148089205185",
  "140063005120",
  "13175205107",
  "170884105191",
  "181987705142",
  "137964305198",
  "91511005161",
  "91510905101",
  "181052305186",
  "44674205128",
  "171479005130",
  "171478905177",
  "92106705133",
  "97217905103",
  "73596905192",
  "57295305106",
  "97217805118",
  "68452205100",
  "180756905147",
  "106795005109",
  "137061305171",
  "56688705130",
  "80210205179",
  "79315905121",
  "2305505194",
  "21901405192",
  "44670905139",
  "58480705168",
  "51008005164",
  "72072705164",
  "170561005110",
  "101400005175",
  "15587305191",
  "135250605105",
  "121199305132",
  "8070005177",
  "7202005104",
  "135250205143",
  "17088005196",
  "164585505118",
  "153632105106",
  "46429705106",
  "156937805100",
  "106745005135",
  "20045305195",
  "139974505159",
  "114757605155",
  "169288805110",
  "20261705108",
  "181502205142",
  "34120105184",
  "38913505143",
  "140891605130",
  "54241805102",
  "163212205144",
  "165649505178",
  "11827605101",
  "11522005153",
  "97442605171",
  "102449105140",
  "43332905165",
  "27728705184",
  "104674005104",
  "89258205104",
  "9733605175",
  "4321005168",
  "48141005188",
  "22611805190",
  "23891505189",
  "179052505157",
  "73502805116",
  "10008405196",
  "158573005108",
  "7929105158",
  "7059205197",
  "7315205160",
  "129726505185",
  "96842605179",
  "89248505102",
  "3341371",
  "19632105119",
  "58947605175",
  "21185305106",
  "171244905109",
  "142859405122",
  "169746905172",
  "112869505154",
  "181124405185",
  "157950805134",
  "103387605113",
  "155259305160",
  "133550005144",
  "122859605119",
  "20842105175",
  "144621905193",
  "28623105119",
  "17189205113",
  "80951205195",
]);

const UCS_USINA = new Set([
  "65021505156",
  "158998605151",
  "112478405100",
  "105912705181",
  "76286605117",
  "130254205109",
  "117316005142",
]);

if (typeof window !== "undefined") {
  try {
    const ppp = localStorage.getItem("PPP_UCS_OVERRIDE");
    if (ppp) {
      UCS_PPP.clear();
      JSON.parse(ppp).forEach((u: string) => UCS_PPP.add(u));
    }
    const usina = localStorage.getItem("USINA_UCS_OVERRIDE");
    if (usina) {
      UCS_USINA.clear();
      JSON.parse(usina).forEach((u: string) => UCS_USINA.add(u));
    }
  } catch (e) {}
}

const hasCompensacao = (d: any) =>
  d.solarInjetadaOUC > 0 ||
  d.solarInjetadaMUC > 0 ||
  UCS_PPP.has(String(d.uc)) ||
  UCS_USINA.has(String(d.uc));

const VisaoGeralDashboard = ({
  data,
  setCurrentPage,
  handleLogout,
  hasApiKey,
  handleSelectKey,
}: {
  data: any[];
  setCurrentPage: (page: string) => void;
  handleLogout: () => void;
  hasApiKey: boolean;
  handleSelectKey: () => void;
}) => {
  const [selectedMonth, setSelectedMonth] = useState<string>("all");
  const [hoveredLivreType, setHoveredLivreType] = useState<
    "azul" | "verde" | null
  >(null);

  const availableMonths = Array.from(new Set(data.map((d) => d.name)))
    .filter(Boolean)
    .sort((a, b) => {
      const [mA, yA] = String(a).split("/");
      const [mB, yB] = String(b).split("/");
      if (yA !== yB) return Number(yB) - Number(yA);
      return getMonthNumber(mB) - getMonthNumber(mA);
    });

  const filteredData =
    selectedMonth === "all"
      ? data
      : data.filter((d) => d.name === selectedMonth);

  const calc = (filterFn: (d: any) => boolean) => {
    const filtered = filteredData.filter(filterFn);
    const custo = filtered.reduce((acc, curr) => acc + curr.valorTotal, 0);
    const consumo = filtered.reduce(
      (acc, curr) =>
        acc + (curr.consumoPonta || 0) + (curr.consumoForaPonta || 0) + (curr.consumoGrupoB || 0) + (curr.consumoKwh || 0),
      0,
    );
    return { custo, consumo, tarifa: consumo > 0 ? custo / consumo : 0 };
  };

  const isGrupoA = (d: any) =>
    d.demandaContratadaPonta > 0 || d.demandaContratadaForaPonta > 0;
  const isGrupoB = (d: any) => !isGrupoA(d);
  const isLivre = (d: any) =>
    d.modalidadeTarifaria.includes("LIVRE") || d.tipo === "LIVRE";
  const isCativo = (d: any) => !isLivre(d);

  const isAzul = (d: any) => d.modalidadeTarifaria.includes("AZUL");
  const isVerde = (d: any) => d.modalidadeTarifaria.includes("VERDE");
  const isOutrosGrupoA = (d: any) => isGrupoA(d) && !isAzul(d) && !isVerde(d);

  const isConsumoMinimo = (d: any) =>
    isGrupoB(d) &&
    (d.consumoPonta || 0) + (d.consumoForaPonta || 0) + (d.consumoGrupoB || 0) + (d.consumoKwh || 0) <= 100 &&
    d.valorTotal < 150;
  const isPPP = (d: any) => UCS_PPP.has(String(d.uc));
  const isUsina = (d: any) => UCS_USINA.has(String(d.uc));
  const isOptanteB = (d: any) =>
    isGrupoB(d) &&
    (d.modalidadeTarifaria || "").toUpperCase().includes("OPTANTE");
  const isGeral = (d: any) =>
    isGrupoB(d) &&
    !isConsumoMinimo(d) &&
    !isPPP(d) &&
    !isUsina(d) &&
    !isOptanteB(d);

  const totalGeral = calc(() => true);
  const grupoA = calc(isGrupoA);
  const grupoB = calc(isGrupoB);

  const livre = calc((d) => isGrupoA(d) && isLivre(d));
  const livreAzul = calc((d) => isGrupoA(d) && isLivre(d) && isAzul(d));
  const livreVerde = calc((d) => isGrupoA(d) && isLivre(d) && isVerde(d));

  const cativo = calc((d) => isGrupoA(d) && isCativo(d));
  const cativoAzul = calc((d) => isGrupoA(d) && isCativo(d) && isAzul(d));
  const cativoVerde = calc((d) => isGrupoA(d) && isCativo(d) && isVerde(d));
  const cativoOutras = calc(
    (d) => isGrupoA(d) && isCativo(d) && isOutrosGrupoA(d),
  );

  const semCompensacao = calc((d) => isGrupoB(d) && !hasCompensacao(d));
  const geral = calc(isGeral);
  const consumosMinimos = calc(isConsumoMinimo);
  const optanteB = calc(isOptanteB);

  const comCompensacao = calc((d) => isGrupoB(d) && hasCompensacao(d));
  const ppp = calc(isPPP);
  const usinas = calc(isUsina);

  const totalSolarInjetada = filteredData.reduce(
    (acc, curr) =>
      acc + (curr.solarInjetadaOUC || 0) + (curr.solarInjetadaMUC || 0),
    0,
  );
  const emissoesEvitadas = totalSolarInjetada * 0.0426; // Fator médio do SIN

  const getMonthlyData = (
    sourceData: any[],
    filterFn?: (d: any) => boolean,
  ) => {
    interface GroupedItem {
      name: string;
      month: number;
      year: number;
      consumo: number;
      custo: number;
    }

    const filtered = filterFn ? sourceData.filter(filterFn) : sourceData;

    const grouped = filtered.reduce(
      (acc, curr) => {
        const name = curr.name; // e.g. "Janeiro/2026"
        if (!acc[name]) {
          const [month, year] = name.split("/");
          acc[name] = {
            name,
            month: getMonthNumber(month),
            year: parseInt(year),
            consumo: 0,
            custo: 0,
          };
        }
        acc[name].consumo += (curr.consumoPonta || 0) + (curr.consumoForaPonta || 0) + (curr.consumoGrupoB || 0) + (curr.consumoKwh || 0);
        acc[name].custo += curr.valorTotal;
        return acc;
      },
      {} as Record<string, GroupedItem>,
    );

    const sorted = (Object.values(grouped) as GroupedItem[]).sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      return a.month - b.month;
    });

    const ptAbbr: Record<string, string> = {
      Janeiro: "Jan",
      Fevereiro: "Fev",
      Março: "Mar",
      Abril: "Abr",
      Maio: "Mai",
      Junho: "Jun",
      Julho: "Jul",
      Agosto: "Ago",
      Setembro: "Set",
      Outubro: "Out",
      Novembro: "Nov",
      Dezembro: "Dez",
    };

    const years = new Set(sorted.map((item) => item.year));
    const multiYear = years.size > 1;

    return sorted.map((item) => {
      const [month] = item.name.split("/");
      const fullMonth = formatMonth(month);
      const abbr = ptAbbr[fullMonth] || month.substring(0, 3);
      return {
        name: multiYear ? `${abbr}/${String(item.year).slice(-2)}` : abbr,
        consumo: item.consumo,
        custo: item.custo,
      };
    });
  };

  const monthlyData = useMemo(() => getMonthlyData(data), [data]);

  const unifiedMax = useMemo(() => {
    const maxConsumo = Math.max(...monthlyData.map((d) => d.consumo), 0);
    const maxCusto = Math.max(...monthlyData.map((d) => d.custo), 0);
    return Math.max(maxConsumo, maxCusto);
  }, [monthlyData]);

  const chartDomainConsumo = useMemo(() => {
    return [0, Math.ceil(unifiedMax * 1.1)];
  }, [unifiedMax]);

  const chartDomainCusto = useMemo(() => {
    return [0, Math.ceil(unifiedMax * 1.1)];
  }, [unifiedMax]);

  const monthlyDataAzul = useMemo(
    () => getMonthlyData(data, (d) => isGrupoA(d) && isLivre(d) && isAzul(d)),
    [data],
  );
  const monthlyDataVerde = useMemo(
    () => getMonthlyData(data, (d) => isGrupoA(d) && isLivre(d) && isVerde(d)),
    [data],
  );

  const sparklineDataAzul = monthlyDataAzul.map((m) => ({ value: m.custo }));
  const sparklineDataVerde = monthlyDataVerde.map((m) => ({ value: m.custo }));

  const unifiedMaxAzul = useMemo(() => {
    const maxConsumo = Math.max(...monthlyDataAzul.map((d) => d.consumo), 0);
    const maxCusto = Math.max(...monthlyDataAzul.map((d) => d.custo), 0);
    return Math.max(maxConsumo, maxCusto);
  }, [monthlyDataAzul]);

  const unifiedMaxVerde = useMemo(() => {
    const maxConsumo = Math.max(...monthlyDataVerde.map((d) => d.consumo), 0);
    const maxCusto = Math.max(...monthlyDataVerde.map((d) => d.custo), 0);
    return Math.max(maxConsumo, maxCusto);
  }, [monthlyDataVerde]);

  const projDomain = useMemo(() => {
    const currentMax = hoveredLivreType === "azul" ? unifiedMaxAzul : unifiedMaxVerde;
    return [0, Math.ceil(currentMax * 1.1)];
  }, [hoveredLivreType, unifiedMaxAzul, unifiedMaxVerde]);

  const MetricRow = ({
    icon: Icon,
    label,
    value,
    unit,
    isCurrency,
  }: {
    icon: any;
    label: string;
    value: number;
    unit?: string;
    isCurrency?: boolean;
  }) => (
    <div className="flex items-center justify-between border-b border-slate-100 pb-1 mb-2 group/row">
      <div className="flex items-center gap-2">
        <Icon
          size={14}
          className="text-slate-400 group-hover/row:text-blue-500 transition-colors"
        />
        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">
          {label}
        </span>
      </div>
      <div className="text-right flex items-baseline gap-1">
        {isCurrency && (
          <span className="text-[10px] text-slate-400 font-bold">R$</span>
        )}
        <span className="text-lg font-black text-slate-900 tracking-tight">
          {formatNumber(value, isCurrency, 2)}
        </span>
        {unit && (
          <span className="text-[10px] text-slate-400 font-bold ml-1">
            {unit}
          </span>
        )}
      </div>
    </div>
  );

  const SummaryCard = ({
    title,
    data,
    icon: Icon,
    color = "blue",
    className = "",
  }: {
    title: string;
    data: any;
    icon: any;
    color?: "blue" | "indigo" | "sky";
    className?: string;
  }) => {
    const colorStyles = {
      blue: {
        icon: "text-blue-600",
        bg: "bg-blue-50",
        border: "border-blue-100",
      },
      indigo: {
        icon: "text-indigo-600",
        bg: "bg-indigo-50",
        border: "border-indigo-100",
      },
      sky: { icon: "text-sky-600", bg: "bg-sky-50", border: "border-sky-100" },
    }[color];

    return (
      <div
        className={`rounded-[2rem] p-5 relative overflow-hidden bg-white text-slate-900 border border-slate-200 group transition-all duration-500 hover:shadow-xl shadow-sm ${className}`}
      >
        <div className="absolute top-0 right-0 w-64 h-64 bg-slate-50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-30 group-hover:scale-110 transition-transform duration-700 blur-3xl"></div>

        <div className="flex justify-between items-center mb-4 relative z-10">
          <div className="flex items-center gap-4">
            <div
              className={`p-2.5 ${colorStyles.bg} rounded-xl border ${colorStyles.border} shadow-sm group-hover:bg-white transition-colors duration-300`}
            >
              <Icon size={20} className={colorStyles.icon} />
            </div>
            <h3 className="text-xs font-black uppercase tracking-[0.2em] text-slate-800">
              {title}
            </h3>
          </div>
        </div>

        <div className="space-y-1 relative z-10">
          <MetricRow
            icon={DollarSign}
            label="Custo Total"
            value={data.custo}
            isCurrency
          />
          <MetricRow
            icon={Zap}
            label="Consumo Total"
            value={data.consumo}
            unit="kWh"
          />
          <MetricRow
            icon={Calculator}
            label="Tarifa Média"
            value={data.tarifa}
            isCurrency
          />
        </div>
      </div>
    );
  };

  const DetailCard = ({
    title,
    data,
    color = "blue",
    icon: Icon,
  }: {
    title: string;
    data: any;
    color?: "blue" | "green" | "slate" | "indigo";
    icon?: any;
  }) => {
    const colorStyles = {
      blue: {
        text: "text-blue-600",
        bg: "bg-blue-50",
        border: "border-blue-100",
        iconBg: "bg-blue-100",
        iconText: "text-blue-600",
        hover: "hover:border-blue-200 hover:bg-blue-100/50",
        valueText: "text-slate-900",
      },
      green: {
        text: "text-emerald-600",
        bg: "bg-emerald-50",
        border: "border-emerald-100",
        iconBg: "bg-emerald-100",
        iconText: "text-emerald-600",
        hover: "hover:border-emerald-200 hover:bg-emerald-100/50",
        valueText: "text-slate-900",
      },
      slate: {
        text: "text-slate-600",
        bg: "bg-slate-50",
        border: "border-slate-200",
        iconBg: "bg-slate-100",
        iconText: "text-slate-600",
        hover: "hover:border-slate-300 hover:bg-slate-100/50",
        valueText: "text-slate-900",
      },
      indigo: {
        text: "text-indigo-600",
        bg: "bg-indigo-50",
        border: "border-indigo-100",
        iconBg: "bg-indigo-100",
        iconText: "text-indigo-600",
        hover: "hover:border-indigo-200 hover:bg-indigo-100/50",
        valueText: "text-slate-900",
      },
    }[color];

    return (
      <div
        className={`rounded-2xl p-3.5 border ${colorStyles.border} ${colorStyles.bg} flex-1 transition-all duration-300 ${colorStyles.hover} group relative overflow-hidden shadow-sm`}
      >
        <div className="absolute top-0 right-0 w-32 h-32 bg-white/50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-40 group-hover:scale-110 transition-transform duration-500 blur-xl"></div>
        <div className="flex items-center gap-3 mb-3 relative z-10">
          {Icon && (
            <div
              className={`p-2 rounded-xl ${colorStyles.iconBg} ${colorStyles.iconText} group-hover:scale-110 transition-transform shadow-sm border border-white`}
            >
              <Icon size={16} />
            </div>
          )}
          <h4
            className={`text-xs font-bold uppercase tracking-wider ${colorStyles.text}`}
          >
            {title}
          </h4>
        </div>
        <div className="space-y-2 relative z-10">
          <div className="flex justify-between items-end bg-white/50 p-1.5 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
              Custo
            </span>
            <span className={`text-sm font-bold ${colorStyles.valueText}`}>
              R$ {formatNumber(data.custo, true)}
            </span>
          </div>
          <div className="flex justify-between items-end bg-white/50 p-1.5 rounded-xl border border-slate-100">
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
              Consumo
            </span>
            <span className={`text-sm font-bold ${colorStyles.valueText}`}>
              {formatNumber(data.consumo, false, 2)}{" "}
              <span className="text-[10px] text-slate-400 font-medium">
                kWh
              </span>
            </span>
          </div>
          <div className="flex justify-between items-end pt-2 border-t border-slate-100">
            <span
              className={`text-[10px] font-bold uppercase tracking-wider ${colorStyles.text}`}
            >
              Tarifa Média
            </span>
            <span className={`text-sm font-bold ${colorStyles.text}`}>
              R$ {formatNumber(data.tarifa, true)}
            </span>
          </div>
        </div>
      </div>
    );
  };

  const SparklineCard = ({
    title,
    data,
    color = "blue",
    fullMonthlyData,
    onMouseEnter,
    onMouseLeave,
  }: {
    title: string;
    data: any;
    color?: "blue" | "green";
    sparklineData: any[];
    fullMonthlyData: any[];
    onMouseEnter?: () => void;
    onMouseLeave?: () => void;
  }) => {
    const bgClass = color === "blue" ? "bg-blue-50" : "bg-emerald-50";
    const textClass = color === "blue" ? "text-blue-600" : "text-emerald-600";
    const valueTextClass = "text-slate-900";
    const borderClass =
      color === "blue" ? "border-blue-100" : "border-emerald-100";
    const hoverBorderClass =
      color === "blue"
        ? "hover:border-blue-200 hover:bg-blue-100/50"
        : "hover:border-emerald-200 hover:bg-emerald-100/50";
    const iconBgClass = color === "blue" ? "bg-blue-100" : "bg-emerald-100";

    const maxConsumo = Math.max(...fullMonthlyData.map((d) => d.consumo), 0);
    const maxCusto = Math.max(...fullMonthlyData.map((d) => d.custo), 0);
    const maxVal = Math.max(maxConsumo, maxCusto);
    const domain = [0, Math.ceil(maxVal * 1.1)];

    return (
      <div
        onMouseEnter={onMouseEnter}
        onMouseLeave={onMouseLeave}
        className={`rounded-2xl p-4 border ${borderClass} ${bgClass} flex flex-col mt-3 transition-all duration-300 ${hoverBorderClass} group relative overflow-hidden shadow-sm h-[220px] cursor-pointer`}
      >
        <div className="absolute top-0 right-0 w-24 h-24 bg-white/50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-40 group-hover:scale-110 transition-transform duration-500 blur-xl"></div>
        <div className="flex items-center justify-between mb-4 relative z-10">
          <div className="flex items-center gap-3">
            <div
              className={`w-8 h-8 rounded-lg ${iconBgClass} flex items-center justify-center shadow-sm border border-white group-hover:scale-110 transition-transform`}
            >
              <TrendingUp size={16} className={textClass} />
            </div>
            <div>
              <h4
                className={`text-[10px] font-bold uppercase tracking-wider text-slate-500`}
              >
                {title}
              </h4>
              <p className={`text-sm font-bold ${valueTextClass}`}>
                R$ {formatNumber(data.custo, true, 2)}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-3 bg-white/50 px-2 py-1 rounded-lg border border-slate-100">
            <div className="flex items-center gap-1.5">
              <div className="w-1.5 h-1.5 rounded-full bg-[#0ea5e9]"></div>
              <span className="text-[8px] font-bold text-slate-500 uppercase">
                Consumo
              </span>
            </div>
            <div className="flex items-center gap-1.5">
              <div className="w-1.5 h-1.5 rounded-full bg-[#6366f1]"></div>
              <span className="text-[8px] font-bold text-slate-500 uppercase">
                Custo
              </span>
            </div>
          </div>
        </div>

        <div className="flex-1 w-full relative z-10">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart
              data={fullMonthlyData}
              margin={{ top: 5, right: 5, bottom: 0, left: -20 }}
            >
              <defs>
                <linearGradient
                  id={`colorConsumo-${color}`}
                  x1="0"
                  y1="0"
                  x2="0"
                  y2="1"
                >
                  <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.2} />
                  <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0} />
                </linearGradient>
                <linearGradient
                  id={`colorCusto-${color}`}
                  x1="0"
                  y1="0"
                  x2="0"
                  y2="1"
                >
                  <stop offset="5%" stopColor="#6366f1" stopOpacity={0.2} />
                  <stop offset="95%" stopColor="#6366f1" stopOpacity={0} />
                </linearGradient>
              </defs>
              <CartesianGrid
                strokeDasharray="3 3"
                vertical={false}
                stroke="#e2e8f0"
                opacity={0.5}
              />
              <XAxis
                dataKey="name"
                axisLine={false}
                tickLine={false}
                tick={{ fill: "#64748b", fontSize: 8, fontWeight: 600 }}
                dy={5}
              />
              <YAxis
                yAxisId="left"
                axisLine={false}
                tickLine={false}
                tick={{ fill: "#64748b", fontSize: 8, fontWeight: 600 }}
                tickFormatter={(val) => formatNumber(val, false, 0)}
                domain={domain}
                hide
              />
              <Tooltip
                contentStyle={{
                  backgroundColor: "rgba(255, 255, 255, 0.95)",
                  borderRadius: "12px",
                  border: "none",
                  boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)",
                  fontSize: "10px",
                }}
                itemStyle={{ padding: "2px 0" }}
                formatter={(value: number, name: string) => [
                  name === "custo"
                    ? `R$ ${formatNumber(value, true, 2)}`
                    : `${formatNumber(value, false, 2)} kWh`,
                  name === "custo" ? "Custo" : "Consumo",
                ]}
              />
              <Area
                yAxisId="left"
                type="monotone"
                dataKey="consumo"
                name="consumo"
                stroke="#0ea5e9"
                strokeWidth={2}
                fillOpacity={1}
                fill={`url(#colorConsumo-${color})`}
                isAnimationActive={false}
              />
              <Area
                yAxisId="left"
                type="monotone"
                dataKey="custo"
                name="custo"
                stroke="#6366f1"
                strokeWidth={2}
                fillOpacity={1}
                fill={`url(#colorCusto-${color})`}
                isAnimationActive={false}
              />
            </AreaChart>
          </ResponsiveContainer>
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
          <div
            className="flex items-center gap-2 px-3 py-2 border rounded-xl text-xs font-semibold shadow-sm bg-sky-50 text-sky-800 border-sky-200"
            title={`Banco de Dados Supabase Cloud: ${data.length.toLocaleString()} faturas carregadas`}
          >
            <Cloud size={14} className="text-sky-600" />
            <span>
              {`Supabase: ${data.length.toLocaleString()} faturas`}
            </span>
          </div>
          <button
            onClick={() => setCurrentPage("sistema")}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white hover:bg-blue-500 transition-all rounded-xl text-xs font-bold tracking-wider shadow-md active:scale-95"
          >
            <LayoutDashboard size={16} />
            Acessar Sistema
          </button>
          <button
            onClick={handleSelectKey}
            className={`flex items-center gap-2 px-4 py-2 ${hasApiKey ? "bg-emerald-50 text-emerald-700 border-emerald-200" : "bg-amber-50 text-amber-700 border-amber-200"} border hover:opacity-80 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95`}
            title={
              hasApiKey ? "Trocar Chave de API" : "Selecionar Chave de API"
            }
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
              <h2 className="text-2xl font-bold text-slate-900 tracking-tight">
                Análise de Insumo de Energia Elétrica
              </h2>
              <p className="text-sm text-slate-500 font-medium">
                Visão geral consolidada de custos e consumos por grupos
                tarifários.
              </p>
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
              {availableMonths.map((month) => (
                <option key={month} value={month}>
                  {month}
                </option>
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
                <h3 className="text-lg font-bold text-slate-900 uppercase tracking-widest">
                  Evolução de Consumo e Custo
                </h3>
                <p className="text-xs text-slate-500 font-medium mt-1">
                  Histórico Mensal Consolidado
                </p>
              </div>
            </div>
            <div className="flex items-center gap-6 bg-slate-50 px-5 py-3 rounded-2xl border border-slate-200 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="w-3 h-3 rounded-full bg-[#0ea5e9]"></div>
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
                  Consumo (kWh)
                </span>
              </div>
              <div className="flex items-center gap-3">
                <div className="w-3 h-3 rounded-full bg-[#6366f1]"></div>
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
                  Custo (R$)
                </span>
              </div>
            </div>
          </div>
          <div className="h-[280px] w-full relative z-10">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart
                data={monthlyData}
                margin={{ top: 20, right: 80, bottom: 10, left: 80 }}
              >
                <defs>
                  <linearGradient id="colorConsumo" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.8} />
                    <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0.3} />
                  </linearGradient>
                  <linearGradient id="colorCusto" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#6366f1" stopOpacity={0.8} />
                    <stop offset="95%" stopColor="#6366f1" stopOpacity={0.3} />
                  </linearGradient>
                </defs>
                <CartesianGrid
                  strokeDasharray="3 3"
                  vertical={false}
                  horizontal={true}
                  stroke="#e2e8f0"
                />
                <XAxis
                  dataKey="name"
                  axisLine={false}
                  tickLine={false}
                  tick={{ fill: "#64748b", fontSize: 11, fontWeight: 600 }}
                  dy={10}
                  padding={{ left: 10, right: 10 }}
                />
                <YAxis
                  axisLine={false}
                  tickLine={false}
                  tick={{ fill: "#64748b", fontSize: 11, fontWeight: 600 }}
                  dx={-10}
                  tickFormatter={(val) => formatNumber(val, false, 0)}
                  domain={chartDomainConsumo}
                  label={{
                    value: "Valores",
                    angle: -90,
                    position: "insideLeft",
                    offset: -55,
                    style: {
                      textAnchor: "middle",
                      fill: "#64748b",
                      fontSize: 10,
                      fontWeight: 700,
                      textTransform: "uppercase",
                    },
                  }}
                />
                <Tooltip
                  contentStyle={{
                    borderRadius: "20px",
                    border: "1px solid #e2e8f0",
                    boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)",
                    padding: "16px 20px",
                    fontWeight: "bold",
                    backgroundColor: "rgba(255, 255, 255, 0.95)",
                    backdropFilter: "blur(12px)",
                  }}
                  formatter={(value: number, name: string) => [
                    name === "consumo"
                      ? formatNumber(value, false, 0) + " kWh"
                      : formatNumber(value, true),
                    name === "consumo" ? "Consumo" : "Custo",
                  ]}
                  labelStyle={{
                    color: "#64748b",
                    marginBottom: "8px",
                    fontSize: "12px",
                    textTransform: "uppercase",
                    letterSpacing: "0.1em",
                  }}
                  itemStyle={{ color: "#1e293b" }}
                />
                <Bar
                  dataKey="consumo"
                  fill="url(#colorConsumo)"
                  radius={[4, 4, 0, 0]}
                  maxBarSize={40}
                />
                <Bar
                  dataKey="custo"
                  fill="url(#colorCusto)"
                  radius={[4, 4, 0, 0]}
                  maxBarSize={40}
                />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Summary Cards */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
          <div className="md:col-span-2">
            <SummaryCard
              title="TOTAL GERAL"
              data={totalGeral}
              icon={DollarSign}
              color="indigo"
            />
          </div>
          <SummaryCard
            title="GRUPO A (MT/AT)"
            data={grupoA}
            icon={Zap}
            color="blue"
          />
          <SummaryCard
            title="GRUPO B (BT)"
            data={grupoB}
            icon={Battery}
            color="sky"
          />
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
              <span className="text-xs font-bold text-blue-600 bg-blue-50 px-3 py-1.5 rounded-lg uppercase tracking-wider border border-blue-100">
                Alta Tensão
              </span>
            </div>

            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* LIVRE Card */}
              <div
                onMouseLeave={() => setHoveredLivreType(null)}
                className="bg-white rounded-[2rem] p-6 shadow-sm border border-blue-100 relative group hover:shadow-xl transition-all duration-500 min-h-[500px] flex flex-col"
              >
                <div className="absolute inset-0 overflow-hidden rounded-[2rem] pointer-events-none">
                  <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-blue-50 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>
                </div>

                <AnimatePresence>
                  {hoveredLivreType && (
                    <motion.div
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 20 }}
                      onMouseLeave={() => setHoveredLivreType(null)}
                      className="absolute inset-x-0 bottom-0 z-50 bg-white/98 backdrop-blur-md rounded-b-[2rem] p-8 border-t border-slate-100 shadow-2xl overflow-hidden"
                      style={{ height: "50%" }}
                    >
                      <div className="flex justify-between items-center mb-6">
                        <div>
                          <h4 className="text-lg font-bold text-slate-800 uppercase tracking-wider">
                            Projeção Detalhada:{" "}
                            {hoveredLivreType === "azul"
                              ? "Evolução Azul"
                              : "Evolução Verde"}
                          </h4>
                          <p className="text-xs text-slate-500">
                            Análise comparativa de consumo e custo mensal
                          </p>
                        </div>
                        <div className="flex gap-6">
                          <div className="flex items-center gap-2">
                            <div className="w-3 h-3 rounded-full bg-[#0ea5e9]"></div>
                            <span className="text-[10px] font-bold text-slate-600 uppercase">
                              Consumo (kWh)
                            </span>
                          </div>
                          <div className="flex items-center gap-2">
                            <div className="w-3 h-3 rounded-full bg-[#6366f1]"></div>
                            <span className="text-[10px] font-bold text-slate-600 uppercase">
                              Custo (R$)
                            </span>
                          </div>
                        </div>
                      </div>
                      <div className="flex-1 w-full h-[calc(100%-60px)]">
                        <ResponsiveContainer width="100%" height="100%">
                          <AreaChart
                            data={
                              hoveredLivreType === "azul"
                                ? monthlyDataAzul
                                : monthlyDataVerde
                            }
                            margin={{
                              top: 20,
                              right: 30,
                              left: 20,
                              bottom: 20,
                            }}
                          >
                            <defs>
                              <linearGradient
                                id="colorConsumoProj"
                                x1="0"
                                y1="0"
                                x2="0"
                                y2="1"
                              >
                                <stop
                                  offset="5%"
                                  stopColor="#0ea5e9"
                                  stopOpacity={0.3}
                                />
                                <stop
                                  offset="95%"
                                  stopColor="#0ea5e9"
                                  stopOpacity={0}
                                />
                              </linearGradient>
                              <linearGradient
                                id="colorCustoProj"
                                x1="0"
                                y1="0"
                                x2="0"
                                y2="1"
                              >
                                <stop
                                  offset="5%"
                                  stopColor="#6366f1"
                                  stopOpacity={0.3}
                                />
                                <stop
                                  offset="95%"
                                  stopColor="#6366f1"
                                  stopOpacity={0}
                                />
                              </linearGradient>
                            </defs>
                            <CartesianGrid
                              strokeDasharray="3 3"
                              vertical={false}
                              stroke="#f1f5f9"
                            />
                            <XAxis
                              dataKey="name"
                              axisLine={false}
                              tickLine={false}
                              tick={{
                                fontSize: 10,
                                fontWeight: 600,
                                fill: "#64748b",
                              }}
                            />
                            <YAxis
                              yAxisId="left"
                              orientation="left"
                              axisLine={false}
                              tickLine={false}
                              tick={{
                                fontSize: 10,
                                fontWeight: 600,
                                fill: "#0ea5e9",
                              }}
                              tickFormatter={(val) =>
                                formatNumber(val, false, 2)
                              }
                              domain={projDomain}
                              label={{
                                value: "Consumo (kWh)",
                                angle: -90,
                                position: "insideLeft",
                                style: {
                                  fontSize: "10px",
                                  fontWeight: "bold",
                                  fill: "#0ea5e9",
                                },
                              }}
                            />
                            <YAxis
                              yAxisId="right"
                              orientation="right"
                              axisLine={false}
                              tickLine={false}
                              tick={{
                                fontSize: 10,
                                fontWeight: 600,
                                fill: "#6366f1",
                              }}
                              tickFormatter={(val) =>
                                `R$ ${formatNumber(val, true, 2)}`
                              }
                              domain={projDomain}
                              label={{
                                value: "Custo (R$)",
                                angle: 90,
                                position: "insideRight",
                                style: {
                                  fontSize: "10px",
                                  fontWeight: "bold",
                                  fill: "#6366f1",
                                },
                              }}
                            />
                            <Tooltip
                              contentStyle={{
                                backgroundColor: "rgba(255, 255, 255, 0.95)",
                                borderRadius: "12px",
                                border: "none",
                                boxShadow: "0 10px 15px -3px rgb(0 0 0 / 0.1)",
                              }}
                              formatter={(val: any, name: string) => [
                                name === "custo"
                                  ? `R$ ${formatNumber(val, true, 2)}`
                                  : `${formatNumber(val, false, 2)} kWh`,
                                name === "custo" ? "Custo" : "Consumo",
                              ]}
                            />
                            <Area
                              yAxisId="left"
                              type="monotone"
                              dataKey="consumo"
                              name="consumo"
                              stroke="#0ea5e9"
                              fillOpacity={1}
                              fill="url(#colorConsumoProj)"
                              strokeWidth={3}
                            />
                            <Area
                              yAxisId="right"
                              type="monotone"
                              dataKey="custo"
                              name="custo"
                              stroke="#6366f1"
                              fillOpacity={1}
                              fill="url(#colorCustoProj)"
                              strokeWidth={3}
                            />
                          </AreaChart>
                        </ResponsiveContainer>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>

                <AnimatePresence mode="wait">
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
                          <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">
                            MERCADO LIVRE
                          </h3>
                          <p className="text-xs text-slate-500 font-medium mt-1">
                            Ambiente de Contratação Livre
                          </p>
                        </div>
                      </div>
                      <div className="p-3 bg-blue-50 rounded-2xl border border-blue-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                        <Activity size={24} className="text-blue-600" />
                      </div>
                    </div>

                    <div className="space-y-1 relative z-10 mb-8">
                      <MetricRow
                        icon={DollarSign}
                        label="Custo Total"
                        value={livre.custo}
                        isCurrency
                      />
                      <MetricRow
                        icon={Zap}
                        label="Consumo Total"
                        value={livre.consumo}
                        unit="kWh"
                      />
                      <MetricRow
                        icon={Calculator}
                        label="Tarifa Média"
                        value={livre.tarifa}
                        isCurrency
                      />
                    </div>

                    <div className="space-y-6 relative z-10 mt-auto">
                      <div className="grid grid-cols-2 gap-6">
                        <DetailCard
                          title="Faturas Azul"
                          data={livreAzul}
                          color="blue"
                          icon={Zap}
                        />
                        <DetailCard
                          title="Faturas Verde"
                          data={livreVerde}
                          color="green"
                          icon={Zap}
                        />
                      </div>
                      <div className="grid grid-cols-2 gap-6">
                        <SparklineCard
                          title="Evolução Azul"
                          data={livreAzul}
                          color="blue"
                          sparklineData={sparklineDataAzul}
                          fullMonthlyData={monthlyDataAzul}
                          onMouseEnter={() => setHoveredLivreType("azul")}
                        />
                        <SparklineCard
                          title="Evolução Verde"
                          data={livreVerde}
                          color="green"
                          sparklineData={sparklineDataVerde}
                          fullMonthlyData={monthlyDataVerde}
                          onMouseEnter={() => setHoveredLivreType("verde")}
                        />
                      </div>
                    </div>
                  </motion.div>
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
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">
                        Consumidor Cativo
                      </h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">
                        Ambiente de Contratação Regulada
                      </p>
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <ZapOff size={24} className="text-slate-600" />
                  </div>
                </div>

                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow
                    icon={DollarSign}
                    label="Custo Total"
                    value={cativo.custo}
                    isCurrency
                  />
                  <MetricRow
                    icon={Zap}
                    label="Consumo Total"
                    value={cativo.consumo}
                    unit="kWh"
                  />
                  <MetricRow
                    icon={Calculator}
                    label="Tarifa Média"
                    value={cativo.tarifa}
                    isCurrency
                  />
                </div>

                <div className="grid grid-cols-2 gap-6 relative z-10">
                  <DetailCard
                    title="Faturas Azul"
                    data={cativoAzul}
                    color="blue"
                    icon={Zap}
                  />
                  <DetailCard
                    title="Faturas Verde"
                    data={cativoVerde}
                    color="green"
                    icon={Zap}
                  />
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
              <span className="text-xs font-bold text-sky-600 bg-sky-50 px-3 py-1.5 rounded-lg uppercase tracking-wider border border-sky-100">
                Baixa Tensão
              </span>
            </div>

            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* Sem Compensação Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-slate-100 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>

                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-slate-400 to-slate-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">
                        UC's Sem Compensação
                      </h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">
                        Consumo Padrão da Rede
                      </p>
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Activity size={24} className="text-slate-600" />
                  </div>
                </div>

                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow
                    icon={DollarSign}
                    label="Custo Total"
                    value={semCompensacao.custo}
                    isCurrency
                  />
                  <MetricRow
                    icon={Zap}
                    label="Consumo Total"
                    value={semCompensacao.consumo}
                    unit="kWh"
                  />
                  <MetricRow
                    icon={Calculator}
                    label="Tarifa Média"
                    value={semCompensacao.tarifa}
                    isCurrency
                  />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10">
                  <DetailCard
                    title="Consumo Geral"
                    data={geral}
                    color="slate"
                    icon={Activity}
                  />
                  <DetailCard
                    title="Consumos Mínimos"
                    data={consumosMinimos}
                    color="indigo"
                    icon={ArrowDown}
                  />
                  <DetailCard
                    title="Optante B"
                    data={optanteB}
                    color="blue"
                    icon={Battery}
                  />
                </div>
              </div>

              {/* Com Compensação Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="absolute top-0 right-0 w-80 h-80 bg-gradient-to-br from-emerald-100 to-transparent rounded-full -translate-y-1/2 translate-x-1/2 opacity-50 group-hover:scale-110 transition-transform duration-700"></div>

                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-emerald-400 to-emerald-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">
                        UC's Com Compensação
                      </h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">
                        Geração Distribuída e Sustentável
                      </p>
                    </div>
                  </div>
                  <div className="p-3 bg-emerald-50 rounded-2xl border border-emerald-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Zap size={24} className="text-emerald-600" />
                  </div>
                </div>

                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow
                    icon={DollarSign}
                    label="Custo Total"
                    value={comCompensacao.custo}
                    isCurrency
                  />
                  <MetricRow
                    icon={Zap}
                    label="Consumo Total"
                    value={comCompensacao.consumo}
                    unit="kWh"
                  />
                  <MetricRow
                    icon={Calculator}
                    label="Tarifa Média"
                    value={comCompensacao.tarifa}
                    isCurrency
                  />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10 mb-6">
                  <DetailCard
                    title="PPP Fotovoltaica"
                    data={ppp}
                    color="green"
                    icon={Zap}
                  />
                  <DetailCard
                    title="Usinas Sanesul"
                    data={usinas}
                    color="blue"
                    icon={Activity}
                  />

                  {/* Crédito de Carbono Card */}
                  <div className="bg-emerald-50/50 rounded-2xl p-5 border border-emerald-100 relative overflow-hidden group">
                    <div className="flex items-center gap-3 mb-4">
                      <div className="p-2 bg-emerald-100 rounded-lg text-emerald-600">
                        <Leaf size={20} />
                      </div>
                      <h4 className="text-sm font-bold text-emerald-800 uppercase tracking-wider">
                        Crédito de Carbono
                      </h4>
                    </div>
                    <div className="space-y-3">
                      <div className="flex justify-between items-center bg-white/60 p-2 rounded-lg border border-emerald-100/50">
                        <span className="text-[10px] font-bold text-emerald-600/70 uppercase tracking-wider">
                          Mês/Ano
                        </span>
                        <span className="text-xs font-bold text-emerald-900">
                          {selectedMonth === "all"
                            ? "Todos os Meses"
                            : selectedMonth}
                        </span>
                      </div>
                      <div className="flex justify-between items-center bg-white/60 p-2 rounded-lg border border-emerald-100/50">
                        <span className="text-[10px] font-bold text-emerald-600/70 uppercase tracking-wider">
                          Geração Solar Abatida
                        </span>
                        <span className="text-xs font-bold text-emerald-900">
                          {formatNumber(totalSolarInjetada, false, 0)}{" "}
                          <span className="text-[10px] text-emerald-600/70">
                            kWh
                          </span>
                        </span>
                      </div>
                      <div className="flex justify-between items-center bg-white/60 p-2 rounded-lg border border-emerald-100/50">
                        <span className="text-[10px] font-bold text-emerald-600/70 uppercase tracking-wider">
                          Emissões Evitadas
                        </span>
                        <span className="text-xs font-bold text-emerald-900">
                          {formatNumber(emissoesEvitadas, false, 2)}{" "}
                          <span className="text-[10px] text-emerald-600/70">
                            KgCO₂
                          </span>
                        </span>
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
    // Auto-authenticate: data is protected by Supabase RLS at the row level.
    // The login screen is optional and only used for write operations control.
    return true;
  });

  const [searchUC, setSearchUC] = useState("");
  const [fetchTrigger, setFetchTrigger] = useState(0);
  const [isSyncing, setIsSyncing] = useState(false);

  React.useEffect(() => {
    if (!isSupabaseConfigured) {
      console.warn(
        "Supabase não configurado. Ignorando verificação de sessão.",
      );
      return;
    }

    supabase.auth
      .getSession()
      .then(({ data: { session } }) => {
        if (session) {
          setIsAuthenticated(true);
          localStorage.setItem("sanesul_auth", "true");
        }
      })
      .catch((err) => {
        console.warn("Aviso ao recuperar sessão do Supabase:", err);
      });

    const {
      data: { subscription },
    } = supabase.auth.onAuthStateChange((event, session) => {
      if (session) {
        localStorage.setItem("sanesul_auth", "true");
      }
      // Never force logout from the dashboard - data is always accessible via RLS anon policy
    });

    return () => subscription.unsubscribe();
  }, []);
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [isSignUp, setIsSignUp] = useState(false);

  const handleSignUp = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError("");
    
    if (!loginUsername || !loginPassword) {
      setLoginError("Por favor, preencha todos os campos.");
      return;
    }

    if (!isSupabaseConfigured) {
      setIsAuthenticated(true);
      localStorage.setItem("sanesul_auth", "true");
      return;
    }
    
    try {
      setIsSyncing(true);
      const { data, error } = await supabase.auth.signUp({
        email: loginUsername,
        password: loginPassword,
      });
      
      if (error) {
        if (error.message?.includes("Failed to fetch") || error.message?.includes("fetch failed")) {
          setLoginError("Erro de conexão com o Supabase (Failed to fetch). Verifique se o projeto no Supabase está ativo ou clique em 'Acessar Modo Local / Offline' abaixo.");
        } else if (error.message?.toLowerCase().includes("rate limit")) {
          setLoginError("Limite de envio de e-mails do Supabase atingido. Tente fazer login direto na aba 'Entrar' (caso já tenha cadastrado) ou desative 'Confirm email' em Authentication -> Providers -> Email no Supabase.");
        } else {
          setLoginError(`Erro ao cadastrar: ${error.message}`);
        }
      } else if (data.user) {
        if (data.session) {
          setIsAuthenticated(true);
          localStorage.setItem("sanesul_auth", "true");
          setFetchTrigger((prev) => prev + 1);
          showAlert("Sucesso", "Cadastro realizado e login efetuado com sucesso!");
        } else {
          showAlert("Sucesso", "Cadastro realizado! Por favor, confirme o e-mail em sua caixa de entrada.");
          setIsSignUp(false);
        }
      }
    } catch (err: any) {
      console.error(err);
      if (err.message?.includes("Failed to fetch") || err.message?.includes("fetch failed")) {
        setLoginError("Erro de conexão com o Supabase (Failed to fetch). Verifique se o projeto no Supabase está ativo ou clique em 'Acessar Modo Local / Offline' abaixo.");
      } else if (err.message?.toLowerCase().includes("rate limit")) {
        setLoginError("Limite de envio de e-mails do Supabase atingido. Tente fazer login direto na aba 'Entrar' ou desative 'Confirm email' em Authentication -> Providers -> Email no Supabase.");
      } else {
        setLoginError(`Erro inesperado: ${err.message || err}`);
      }
    } finally {
      setIsSyncing(false);
    }
  };

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError("");

    console.log("Tentando login com:", loginUsername);

    try {
      if (!isSupabaseConfigured) {
        console.log("Modo local: autenticado com sucesso.");
        setIsAuthenticated(true);
        localStorage.setItem("sanesul_auth", "true");
        return;
      }

      const { data, error } = await supabase.auth.signInWithPassword({
        email: loginUsername,
        password: loginPassword,
      });

      if (error) {
        console.error("Erro no Supabase Auth:", error);
        if (error.message?.includes("Failed to fetch") || error.message?.includes("fetch failed")) {
          setLoginError("Erro de conexão com o Supabase (Failed to fetch). Verifique se o projeto no Supabase está ativo ou clique em 'Acessar Modo Local / Offline' abaixo.");
        } else {
          setLoginError(`Erro de autenticação: ${error.message}`);
        }
      } else if (data.user) {
        console.log("Login bem-sucedido:", data.user);
        setIsAuthenticated(true);
        localStorage.setItem("sanesul_auth", "true");
        setFetchTrigger((prev) => prev + 1);
      }
    } catch (err: any) {
      console.error("Erro inesperado no handleLogin:", err);
      if (err.message?.includes("Failed to fetch") || err.message?.includes("fetch failed")) {
        setLoginError("Erro de conexão com o Supabase (Failed to fetch). Verifique se o projeto no Supabase está ativo ou clique em 'Acessar Modo Local / Offline' abaixo.");
      } else {
        setLoginError("Erro inesperado ao tentar logar. Verifique a conexão.");
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
    try {
      if (isSupabaseConfigured) {
        await supabase.auth.signOut();
      }
    } catch (err) {
      console.warn("Erro no signOut:", err);
    }
    localStorage.removeItem("sanesul_auth");
    // App stays open - no login wall. Data is always accessible via Supabase anon policy.
    showAlert("Sessão Encerrada", "Sua sessão foi encerrada. O painel permanece acessível em modo de visualização.");
  };

  const [bills, setBills] = useState<BillData[]>(() => {
    const saved = localStorage.getItem("sanesul_bills");
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Deduplicate on load
        const deduplicated = deduplicateBills(
          Array.isArray(parsed) ? parsed : [],
        );

        // One-time fix for Elektro/Energisa UC from localStorage & UC 12031 replacement
        return deduplicated.map((b) => {
          let currentUc = b.uc === "179291005130" ? "179291005130" : b.uc;
          const upperConc = (b.concessionaria || "").toUpperCase();
          const isElektro = upperConc.includes("ELEKTRO");
          const isEnergisa = upperConc.includes("ENERGISA");
          if ((isElektro || isEnergisa) && b.status === "completed") {
            const solvedUc = extractUcFromFileName(b.fileName, currentUc);
            if (solvedUc && currentUc !== solvedUc) {
              return { ...b, uc: solvedUc === "179291005130" ? "179291005130" : solvedUc };
            }
          }
          return { ...b, uc: currentUc };
        });
      } catch (e) {
        return [];
      }
    }
    return [];
  });


  const [customRequestedAdjustments, setCustomRequestedAdjustments] = useState<Record<string, { p: number; fp: number }>>(() => {
    try {
      const saved = localStorage.getItem("custom_requested_adjustments");
      let data = saved ? JSON.parse(saved) : { ...REQUESTED_ADJUSTMENTS };
      
      let needsSave = false;
      for (const [oldUc, newUc] of Object.entries(UC_MIGRATION_MAP)) {
        if (data[oldUc]) {
          data[newUc] = data[oldUc];
          delete data[oldUc];
          needsSave = true;
        }
      }
      if (needsSave) {
        try { localStorage.setItem("custom_requested_adjustments", JSON.stringify(data)); } catch (e) {}
      }

      Object.keys(REQUESTED_ADJUSTMENTS).forEach((k) => {
        if (!data[k]) data[k] = REQUESTED_ADJUSTMENTS[k];
      });
      return data;
    } catch {
      return { ...REQUESTED_ADJUSTMENTS };
    }
  });

  const [customOriginalContratadas, setCustomOriginalContratadas] = useState<Record<string, { p: number; fp: number }>>(() => {
    try {
      const saved = localStorage.getItem("custom_original_contratadas");
      let data = saved ? JSON.parse(saved) : { ...ORIGINAL_CONTRATADAS };
      
      let needsSave = false;
      for (const [oldUc, newUc] of Object.entries(UC_MIGRATION_MAP)) {
        if (data[oldUc]) {
          data[newUc] = data[oldUc];
          delete data[oldUc];
          needsSave = true;
        }
      }
      if (needsSave) {
        try { localStorage.setItem("custom_original_contratadas", JSON.stringify(data)); } catch (e) {}
      }

      Object.keys(ORIGINAL_CONTRATADAS).forEach((k) => {
        if (!data[k]) data[k] = ORIGINAL_CONTRATADAS[k];
      });
      return data;
    } catch {
      return { ...ORIGINAL_CONTRATADAS };
    }
  });

  const [customAdjustmentsMetadata, setCustomAdjustmentsMetadata] = useState<Record<string, {
    city?: string;
    gerencia?: string;
    dataSolicitacao?: string;
    dataAlteracao?: string;
    previsaoEconomia?: string;
    ecoRealizada?: string;
    status?: string;
  }>>(() => {
    try {
      const saved = localStorage.getItem("custom_adjustments_metadata");
      let data = saved ? JSON.parse(saved) : {};
      
      let needsSave = false;
      for (const [oldUc, newUc] of Object.entries(UC_MIGRATION_MAP)) {
        if (data[oldUc]) {
          data[newUc] = data[oldUc];
          delete data[oldUc];
          needsSave = true;
        }
      }
      if (needsSave) {
        try { localStorage.setItem("custom_adjustments_metadata", JSON.stringify(data)); } catch (e) {}
      }
      return data;
    } catch {
      return {};
    }
  });

  // Estado de sincronização com o Supabase Cloud
  const [supabaseHealth, setSupabaseHealth] = useState({
    connected: isSupabaseConfigured,
    totalBills: 0,
    totalUcs: 0,
    lastSync: new Date(),
  });

  // Limpeza e sanitização inicial de UCs no armazenamento
  useEffect(() => {
    try {
      ["PPP_UCS_OVERRIDE", "USINA_UCS_OVERRIDE"].forEach((key) => {
        const val = localStorage.getItem(key);
        if (val && val.includes("179291005130")) {
          const arr = JSON.parse(val);
          if (Array.isArray(arr)) {
            const replaced = arr.map((x) => (x === "179291005130" ? "179291005130" : x));
            localStorage.setItem(key, JSON.stringify(replaced));
          }
        }
      });
      const valMappings = localStorage.getItem("sanesul_uc_mappings");
      if (valMappings && valMappings.includes("179291005130")) {
        const arr = JSON.parse(valMappings);
        if (Array.isArray(arr)) {
          const replaced = arr.map((m: any) => (m.uc === "179291005130" ? { ...m, uc: "179291005130" } : m));
          localStorage.setItem("sanesul_uc_mappings", JSON.stringify(replaced));
        }
      }
    } catch (e) {}
  }, []);

  const [isAdjustmentModalOpen, setIsAdjustmentModalOpen] = useState(false);
  const [isNewAdjustment, setIsNewAdjustment] = useState(false);
  const [editingAdjustment, setEditingAdjustment] = useState<{
    uc: string;
    origP: number;
    origFP: number;
    reqP: number;
    reqFP: number;
    city?: string;
    gerencia?: string;
    dataSolicitacao?: string;
    dataAlteracao?: string;
    previsaoEconomia?: string;
    ecoRealizada?: string;
    status?: string;
  } | null>(null);

  React.useEffect(() => {
    localStorage.setItem("custom_requested_adjustments", JSON.stringify(customRequestedAdjustments));
  }, [customRequestedAdjustments]);

  React.useEffect(() => {
    localStorage.setItem("custom_original_contratadas", JSON.stringify(customOriginalContratadas));
  }, [customOriginalContratadas]);

  React.useEffect(() => {
    localStorage.setItem("custom_adjustments_metadata", JSON.stringify(customAdjustmentsMetadata));
  }, [customAdjustmentsMetadata]);

  const [monitoringResults, setMonitoringResults] = useState<any>(null);
  const [selectedMonitoramentoMes, setSelectedMonitoramentoMes] = useState<string>("Todos");
  const [expandedUCs, setExpandedUCs] = useState<Set<string>>(new Set());
  const [expandedAnalysisUCs, setExpandedAnalysisUCs] = useState<Set<string>>(new Set());
  const [expandedSummaryCities, setExpandedSummaryCities] = useState<Set<string>>(new Set());
  const [expandedReactiveUcs, setExpandedReactiveUcs] = useState<Set<string>>(new Set());

  React.useEffect(() => {
    const completedBills = bills.filter((b) => b.status === "completed");
    if (monitoringResults && completedBills.length > 0) {
      runMonitoringAnalysis();
    }
  }, [customRequestedAdjustments, customOriginalContratadas, customAdjustmentsMetadata]);

  const handleDeleteAdjustment = (ucId: string) => {
    setConfirmModalData({
      title: "Excluir Ajuste de Demanda",
      message: `Tem certeza que deseja remover a UC ${ucId} do monitoramento de ajuste de demanda? Esta ação não apagará as faturas da unidade descrita.`,
      onConfirm: () => {
        setCustomRequestedAdjustments(prev => {
          const next = { ...prev };
          delete next[ucId];
          return next;
        });
        setCustomOriginalContratadas(prev => {
          const next = { ...prev };
          delete next[ucId];
          return next;
        });
        setShowConfirmModal(false);
      },
      type: "danger",
    });
    setShowConfirmModal(true);
  };

  React.useEffect(() => {
    localforage
      .getItem<Record<string, File>>("sanesul_bills_files")
      .then((filesMap) => {
        if (filesMap) {
          setBills((prev) =>
            prev.map((b) => {
              if (filesMap[b.id]) {
                return { ...b, file: filesMap[b.id] };
              }
              return b;
            }),
          );
        }
      })
      .catch((err) => {
        console.warn("Failed to load files from localforage:", err);
      });
  }, []);

  React.useEffect(() => {
    let isCancelled = false;

    const SUPA_URL = 'https://yydvjgbfaapldtkhlqrh.supabase.co';
    const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inl5ZHZqZ2JmYWFwbGR0a2hscXJoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzMzMjQ1NTMsImV4cCI6MjA4ODkwMDU1M30.GqUoGviAYXveiEs7YmtN6SE5eZ3ZbiENaZtPUfy8blU';

    const supaFetch = async (table: string, params: string) => {
      const res = await fetch(`${SUPA_URL}/rest/v1/${table}?${params}`, {
        method: 'GET',
        headers: {
          'apikey': SUPA_KEY,
          'Authorization': `Bearer ${SUPA_KEY}`,
          'Content-Type': 'application/json',
          'Prefer': 'return=representation',
        },
      });
      if (!res.ok) throw new Error(`Supabase REST error: ${res.status} ${await res.text()}`);
      return res.json();
    };

    const fetchBills = async () => {
      try {
        setIsSyncing(true);
        console.log('[Supabase REST] Carregando dados da nuvem...');

        // 1. Carrega mapeamentos de UCs
        try {
          const cloudMappings = await supaFetch('uc_mappings', 'select=*');
          if (cloudMappings && cloudMappings.length > 0 && !isCancelled) {
            const sanitizedCloud = cloudMappings.map((m: any) => ({
              uc: String(m.uc),
              gerencia: m.gerencia || '',
              locin: m.locin || '',
              cidade: m.cidade || ''
            }));
            setUcMappings((prev) => {
              const map = new Map(prev.map((m) => [m.uc, m]));
              sanitizedCloud.forEach((m: any) => map.set(m.uc, m));
              return Array.from(map.values());
            });
          }
        } catch (err) {
          console.warn('[Supabase REST] Erro ao carregar uc_mappings:', err);
        }

        // 2. Carrega todas as faturas paginadas com fetch nativo
        let allData: any[] = [];
        let from = 0;
        let finished = false;

        while (!finished && !isCancelled) {
          const chunk = await supaFetch(
            'bills',
            `select=*&order=created_at.desc,id.desc&offset=${from}&limit=1000`
          );

          if (chunk && chunk.length > 0) {
            allData = [...allData, ...chunk];
            if (chunk.length < 1000) {
              finished = true;
            } else {
              from += 1000;
            }
          } else {
            finished = true;
          }
        }

        if (isCancelled) return;

        console.log(`[Supabase REST] Total de faturas recebidas: ${allData.length}`);

        if (allData.length > 0) {
          const mappedBills = allData.map(mapDbToBillData);

          // Extract unique mappings from the database
          const dbMappings: Record<string, UCLocinMapping> = {};
          mappedBills.forEach((b) => {
            if (b.uc && b.gerencia && b.locin) {
              dbMappings[b.uc] = {
                uc: b.uc,
                gerencia: b.gerencia,
                locin: b.locin,
                cidade: b.cidade || "",
              };
            }
          });

          if (Object.keys(dbMappings).length > 0) {
            setUcMappings((prev) => {
              const existingMap = new Map(prev.map((m) => [m.uc, m]));
              Object.values(dbMappings).forEach((m) => {
                existingMap.set(m.uc, m);
              });
              const updated = Array.from(existingMap.values());
              try {
                localStorage.setItem(
                  "sanesul_uc_mappings",
                  JSON.stringify(updated),
                );
              } catch (e) {}
              return updated;
            });
          }

          // Apply fixes & deduplicate
          const updatedBills = mappedBills.map((b) => {
            let updatedBill = { ...b };
            // Fix Elektro/Energisa UC - filename is the authoritative UC
            const upperConc2 = (b.concessionaria || "").toUpperCase();
            const isElektroOrEnergisa =
              upperConc2.includes("ELEKTRO") || upperConc2.includes("ENERGISA");
            if (isElektroOrEnergisa && b.status === "completed") {
              const solvedUc = extractUcFromFileName(b.fileName, b.uc);
              if (solvedUc && b.uc !== solvedUc) {
                updatedBill.uc = solvedUc;
              }
            }

            // Fix UC 117384 - FÁTIMA DO SUL
            if (b.uc === "93604305181" && b.cidade !== "FÁTIMA DO SUL") {
              updatedBill.cidade = "FÁTIMA DO SUL";
            }

            return updatedBill;
          });

          setBills((prev) => {
            const pending = prev.filter((b) => b.status !== "completed");
            return deduplicateBills([...updatedBills, ...pending]);
          });

          // Cache in localStorage as fallback
          try {
            localStorage.setItem("sanesul_bills", JSON.stringify(updatedBills));
          } catch (e) {}

          setSupabaseHealth({
            connected: true,
            totalBills: updatedBills.length,
            totalUcs: new Set(updatedBills.map(b => b.uc).filter(Boolean)).size,
            lastSync: new Date(),
          });
        }
      } catch (err) {
        console.error("[Supabase] Erro inesperado ao buscar faturas:", err);
      } finally {
        setIsSyncing(false);
      }
    };

    fetchBills();

    return () => {
      isCancelled = true;
    };
  }, [fetchTrigger]);

  const saveTimeoutRef = React.useRef<NodeJS.Timeout>();

  React.useEffect(() => {
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    saveTimeoutRef.current = setTimeout(async () => {
      try {
        const completedBillsList = bills.filter((b) => b.status === "completed");

        // Direct auto-save batch to Supabase Cloud
        if (isSupabaseConfigured && completedBillsList.length > 0) {
          try {
            const dbData = completedBillsList.map(b => mapBillDataToDb(b));
            for (let i = 0; i < dbData.length; i += 100) {
              const chunk = dbData.slice(i, i + 100);
              await supabase.from("bills").upsert(chunk, { onConflict: "id" });
            }
          } catch (cloudErr) {
            console.warn("[Supabase] Erro ao persistir faturas na nuvem:", cloudErr);
          }
        }

        const billsToSave = bills.map((b) => {
          const { file, ...rest } = b as any;
          return rest;
        });
        localStorage.setItem("sanesul_bills", JSON.stringify(billsToSave));

        // Save files to localforage ONLY if files changed to prevent UI freezing
        const currentFileIds = bills.filter(b => (b as any).file).map(b => b.id).sort().join(',');
        if (currentFileIds !== (window as any).__lastSavedFileIds) {
          (window as any).__lastSavedFileIds = currentFileIds;
          const filesMap: Record<string, File> = {};
          bills.forEach((b) => {
            if ((b as any).file) {
              filesMap[b.id] = (b as any).file;
            }
          });

          if (Object.keys(filesMap).length <= 50) {
            localforage.setItem("sanesul_bills_files", filesMap).catch((err) => {
              console.warn("Failed to save files to localforage:", err);
            });
          }
        }
      } catch (e) {
        console.warn("LocalStorage limit reached, skipping save:", e);
      }
    }, 1500);
  }, [bills, isAuthenticated]);

  const [isProcessing, setIsProcessing] = useState(false);
  const isProcessingRef = useRef(false);

  useEffect(() => {
    isProcessingRef.current = isProcessing;
  }, [isProcessing]);
  const [isDragging, setIsDragging] = useState(false);
  const [ucsTrigger, setUcsTrigger] = useState(0);
  const [currentPage, setCurrentPage] = useState<"visao_geral" | "sistema">(
    "visao_geral",
  );
  const [activeTab, setActiveTab] = useState<
    | "faturas"
    | "multas"
    | "dashboard"
    | "analises"
    | "monitoramento"
    | "monitoramento_ajustes"
    | "monitoramento_reativo"
    | "monitoramento_usinas"
    | "relatorio"
  >("faturas");
  const [multasMonth, setMultasMonth] = useState<string>("all");
  const [selectedMultaType, setSelectedMultaType] = useState<
    "ultrapassagem" | "reativa" | "subutilizacao" | "total"
  >("total");
  const [multasSortDirection, setMultasSortDirection] = useState<
    "asc" | "desc"
  >("desc");
  const [filterReference, setFilterReference] = useState<string>("all");
  const [sortConfig, setSortConfig] = useState<{
    key: keyof BillData | "referencia";
    direction: "asc" | "desc";
  } | null>(null);
  const [analysisData, setAnalysisData] = useState<any[]>([]);
  const [memoNumber, setMemoNumber] = useState(
    `001447/${new Date().getFullYear()}/GEDEO/DCO`,
  );
  const [memoNfEnergisa, setMemoNfEnergisa] = useState("");
  const [memoNfElektro, setMemoNfElektro] = useState("");
  const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);

  const handleDownloadPDF = async () => {
    const element = document.getElementById("memo-content");
    if (!element) return;

    setIsGeneratingPDF(true);
    try {
      // Create a clone of the element to modify it for PDF generation
      const clone = element.cloneNode(true) as HTMLElement;

      // Update the memo number in the clone immediately
      const memoNumEl = clone.querySelector(".memo-number-text");
      if (memoNumEl) {
        memoNumEl.textContent = `MEMO Nº ${memoNumber}`;
      }

      // Create a temporary container off-screen to render the clone
      const container = document.createElement("div");
      container.style.position = "absolute";
      container.style.left = "-9999px";
      container.style.top = "0";
      container.style.width = "210mm"; // A4 width
      container.appendChild(clone);
      document.body.appendChild(container);

      // Convert to image using dom-to-image-more (handles modern CSS like oklch better)
      const dataUrl = await domtoimage.toJpeg(clone, {
        quality: 0.98,
        bgcolor: "#ffffff",
        width: clone.offsetWidth,
        height: clone.offsetHeight,
        style: {
          transform: "scale(1)",
          transformOrigin: "top left",
        },
      });

      // Calculate dimensions for A4
      const pdf = new jsPDF("p", "mm", "a4");
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (clone.offsetHeight * pdfWidth) / clone.offsetWidth;

      pdf.addImage(dataUrl, "JPEG", 0, 0, pdfWidth, pdfHeight);
      pdf.save(`memorando_faturamento_${new Date().getTime()}.pdf`);

      // Cleanup
      document.body.removeChild(container);
    } catch (error) {
      console.error("Erro ao gerar PDF:", error);
      showAlert("Erro", "Erro ao gerar o PDF. Tente novamente.");
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
        sections: [
          {
            properties: {},
            children: [
              // Header Text (Safe fallback instead of images to prevent corruption)
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "EMPRESA DE SANEAMENTO DE MATO GROSSO DO SUL S.A.",
                    bold: true,
                    size: 24,
                    color: "0070C0",
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "DIRETORIA DA PRESIDÊNCIA",
                    bold: true,
                    size: 20,
                    color: "0070C0",
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({}),

              // Memo Number
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: `MEMO Nº ${memoNumber || "-"}`,
                    bold: true,
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),

              // Date
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: `Campo Grande, ${new Intl.DateTimeFormat("pt-BR", { dateStyle: "long" }).format(new Date())}.`,
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),

              // To/From/Subject
              new Paragraph({
                children: [
                  new TextRun({ text: "DE: ", bold: true, size: 24 }),
                  new TextRun({
                    text: "GEDEO - Gerência de Desenvolvimento Operacional",
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({ text: "PARA: ", bold: true, size: 24 }),
                  new TextRun({
                    text: "GEFI - Gerência Financeira e Gestão de Recursos",
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({ text: "ASSUNTO: ", bold: true, size: 24 }),
                  new TextRun({
                    text: `Faturas Agrupadora Operacional Energisa e Agrupadora Elektro — ${selectedRelatorioMonth === "all" ? "Consolidado" : selectedRelatorioMonth}${!selectedRelatorioType.includes("all") ? ` (${selectedRelatorioType.join(", ")})` : ""}.`,
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),

              // Body
              new Paragraph({
                children: [
                  new TextRun({ text: "        Prezado(a),", size: 24 }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.JUSTIFIED,
                children: [
                  new TextRun({
                    text: "        Seguem anexas para pagamento as faturas de energia elétrica Agrupadora da concessionária Energisa MS, e Agrupadora da concessionária Elektro — todas referentes ao mês de ",
                    size: 24,
                  }),
                  new TextRun({
                    text:
                      selectedRelatorioMonth === "all"
                        ? "todos os períodos"
                        : selectedRelatorioMonth,
                    bold: true,
                    color: "0070C0",
                    size: 24,
                  }),
                  new TextRun({
                    text: !selectedRelatorioType.includes("all")
                      ? ` (Tipo: ${selectedRelatorioType.join(", ")})`
                      : "",
                    size: 24,
                  }),
                  new TextRun({
                    text: " e correspondentes às unidades operacionais da SANESUL.",
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Na Tabela 1 são especificadas as faturas anexas.",
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Tabela 1 - Faturas Anexas",
                    bold: true,
                    size: 24,
                  }),
                ],
              }),

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
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "LOCALIDADE", bold: true }),
                            ],
                          }),
                        ],
                        shading: {
                          fill: "E0E0E0",
                          type: ShadingType.CLEAR,
                          color: "auto",
                        },
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "VALOR (R$)", bold: true }),
                            ],
                          }),
                        ],
                        shading: {
                          fill: "E0E0E0",
                          type: ShadingType.CLEAR,
                          color: "auto",
                        },
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "NOTA FISCAL", bold: true }),
                            ],
                          }),
                        ],
                        shading: {
                          fill: "E0E0E0",
                          type: ShadingType.CLEAR,
                          color: "auto",
                        },
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: "REF: MÊS / ANO",
                                bold: true,
                              }),
                            ],
                          }),
                        ],
                        shading: {
                          fill: "E0E0E0",
                          type: ShadingType.CLEAR,
                          color: "auto",
                        },
                      }),
                    ],
                  }),
                  // Energisa Main Row
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text: "Agrupadora Energisa Operacional",
                                bold: true,
                                color: "0070C0",
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: `R$ ${(memoData?.energisa?.total || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        rowSpan: 5,
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: memoData?.energisa?.nf || "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        rowSpan: 5,
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: memoData?.energisa?.mesRef || "-",
                                bold: true,
                                color: "0070C0",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  // Energisa Details
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: "PIS", size: 20 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.energisa?.pis || 0) > 0
                                    ? `R$ ${(memoData?.energisa?.pis || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "COFINS", size: 20 }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.energisa?.cofins || 0) > 0
                                    ? `R$ ${(memoData?.energisa?.cofins || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: "ICMS", size: 20 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.energisa?.icms || 0) > 0
                                    ? `R$ ${(memoData?.energisa?.icms || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: "Tarifa de Iluminação Pública",
                                size: 20,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.energisa?.cip || 0) > 0
                                    ? `R$ ${(memoData?.energisa?.cip || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  // Elektro Main Row
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text: "Agrupadora Elektro",
                                bold: true,
                                color: "ED7D31",
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: `R$ ${(memoData?.elektro?.total || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        rowSpan: 5,
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: memoData?.elektro?.nf || "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        rowSpan: 5,
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: memoData?.elektro?.mesRef || "-",
                                bold: true,
                                color: "ED7D31",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  // Elektro Details
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: "PIS", size: 20 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.elektro?.pis || 0) > 0
                                    ? `R$ ${(memoData?.elektro?.pis || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "COFINS", size: 20 }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.elektro?.cofins || 0) > 0
                                    ? `R$ ${(memoData?.elektro?.cofins || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [new TextRun({ text: "ICMS", size: 20 })],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.elektro?.icms || 0) > 0
                                    ? `R$ ${(memoData?.elektro?.icms || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: "Tarifa de Iluminação Pública",
                                size: 20,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text:
                                  (memoData?.elektro?.cip || 0) > 0
                                    ? `R$ ${(memoData?.elektro?.cip || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                    : "-",
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                  // Total Row
                  new TableRow({
                    children: [
                      new TableCell({
                        children: [
                          new Paragraph({
                            children: [
                              new TextRun({
                                text: "TOTAL ( Agrupadora ENERGISA + ELEKTRO)",
                                bold: true,
                                italics: true,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: `R$ ${((memoData?.energisa?.total || 0) + (memoData?.elektro?.total || 0)).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`,
                              }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "-------------------" }),
                            ],
                          }),
                        ],
                      }),
                      new TableCell({
                        children: [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({ text: "-------------------" }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: "Proc. N.º 694/2018    Nota Orçamentária Nº 003/2019",
                    italics: true,
                    size: 20,
                  }),
                ],
              }),
              new Paragraph({
                pageBreakBefore: true,
                alignment: AlignmentType.JUSTIFIED,
                children: [
                  new TextRun({
                    text: "        A planilha contendo a estratificação dos dados apresentados neste memorando está disponível em \\\\srv-fs-01\\DADOS\\DCO\\GEDEO\\OPERACAO_AGUA\\COTAA\\ENERGIA\\FATURAS.",
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({}),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [new TextRun({ text: "Atenciosamente,", size: 24 })],
              }),
              new Paragraph({}),
              new Paragraph({}),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: "Fabio Roberto Alves da Silva",
                    bold: true,
                    size: 24,
                  }),
                ],
              }),
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: "Engenheiro Eletricista/GEDEO/Gerência de Desenvolvimento Operacional",
                    size: 24,
                  }),
                ],
              }),
            ],
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      saveAs(blob, `memorando_faturamento_${new Date().getTime()}.docx`);
    } catch (error) {
      console.error("Erro ao gerar DOCX:", error);
      showAlert(
        "Erro",
        "Ocorreu um erro ao gerar o arquivo DOCX. Verifique o console para mais detalhes.",
      );
    }
  };

  const requestSort = (key: keyof BillData | "referencia") => {
    if (sortConfig && sortConfig.key === key) {
      if (sortConfig.direction === "asc") {
        setSortConfig({ key, direction: "desc" });
      } else {
        setSortConfig(null);
      }
    } else {
      setSortConfig({ key, direction: "asc" });
    }
  };

  const availableReferences = React.useMemo(() => {
    const refs = new Set<string>();
    bills.forEach((b) => {
      if (b.mesReferencia && b.anoLeitura) {
        refs.add(`${formatMonth(b.mesReferencia)}/${b.anoLeitura}`);
      }
    });
    return Array.from(refs).sort((a, b) => {
      const [mA, yA] = a.split("/");
      const [mB, yB] = b.split("/");
      if (yA !== yB) return parseInt(yB) - parseInt(yA);
      return getMonthNumber(mB) - getMonthNumber(mA);
    });
  }, [bills]);

  const sortedBills = React.useMemo(() => {
    let filtered = [...bills];
    if (filterReference !== "all") {
      filtered = filtered.filter(
        (b) =>
          `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` === filterReference,
      );
    }

    // UC Search Filter
    if (searchUC.trim() !== "") {
      const search = searchUC.toLowerCase().trim();
      filtered = filtered.filter((b) =>
        (b.uc || "").toLowerCase().includes(search),
      );
    }

    let sortableBills = filtered;

    sortableBills.sort((a, b) => {
      // Priority mapping for statuses
      const getStatusPriority = (status: string) => {
        if (status === "processing") return 0;
        if (status === "error") return 1;
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
          return matches ? parseInt(matches.join(""), 10) : 0;
        };

        let aValue: any;
        let bValue: any;

        if (sortConfig.key === "referencia") {
          const monthA = getMonthNumber(a.mesReferencia);
          const monthB = getMonthNumber(b.mesReferencia);
          aValue = parseInt(a.anoLeitura || "0", 10) * 100 + monthA;
          bValue = parseInt(b.anoLeitura || "0", 10) * 100 + monthB;
        } else if (sortConfig.key === "uc") {
          aValue = extractNumericValue(a.uc || "");
          bValue = extractNumericValue(b.uc || "");
        } else if (sortConfig.key === "fileName") {
          aValue = extractNumericValue(a.fileName || "");
          bValue = extractNumericValue(b.fileName || "");
        } else if (sortConfig.key === "concessionaria") {
          aValue = (a.concessionaria || "").toLowerCase();
          bValue = (b.concessionaria || "").toLowerCase();
        } else {
          aValue = a[sortConfig.key as keyof BillData];
          bValue = b[sortConfig.key as keyof BillData];
        }

        if (aValue < bValue) return sortConfig.direction === "asc" ? -1 : 1;
        if (aValue > bValue) return sortConfig.direction === "asc" ? 1 : -1;
      }

      // Default sort by createdAt descending (newest first)
      return (b.createdAt || 0) - (a.createdAt || 0);
    });

    return sortableBills;
  }, [bills, sortConfig, filterReference, searchUC]);

  const [analysisResults, setAnalysisResults] = useState<any>(null);
  const [dashboardSubTab, setDashboardSubTab] = useState<
    "operacionais" | "financeiro"
  >("operacionais");
  const [operationalSubTab, setOperationalSubTab] = useState<
    "consumo" | "ultrapassagem" | "subutilizacao" | "reativa" | "solar"
  >("consumo");
  const [financialSubTab, setFinancialSubTab] = useState<
    | "despesas"
    | "multa_ultrapassagem"
    | "multa_reativa"
    | "tarifa_media"
    | "energia_solar"
  >("despesas");
  const [selectedUC, setSelectedUC] = useState<string>("");
  const [selectedMonth, setSelectedMonth] = useState<string>("all");
  const [selectedConcessionaria, setSelectedConcessionaria] =
    useState<string>("all");
  const [selectedRelatorioMonth, setSelectedRelatorioMonth] =
    useState<string>("all");
  const [selectedRelatorioType, setSelectedRelatorioType] = useState<string[]>([
    "all",
  ]);
  const [isRelatorioTypeDropdownOpen, setIsRelatorioTypeDropdownOpen] =
    useState(false);
  const [selectedReactiveMonth, setSelectedReactiveMonth] =
    useState<string>("all");
  const [selectedUsinaMonth, setSelectedUsinaMonth] = useState<string>("all");
  const [selectedUsinaCity, setSelectedUsinaCity] = useState<string>("all");
  const [reactiveSortField, setReactiveSortField] =
    useState<string>("totalGeral");
  const [reactiveSortDirection, setReactiveSortDirection] = useState<
    "asc" | "desc"
  >("desc");
  const [dashboardSort, setDashboardSort] = useState<{
    key: string;
    direction: "asc" | "desc";
  }>({ key: "name", direction: "desc" });
  const [showMemo, setShowMemo] = useState(false);
  const [showMemoNumberPrompt, setShowMemoNumberPrompt] = useState(false);
  const [tempMemoNumber, setTempMemoNumber] = useState("");
  const [tempMemoNfEnergisa, setTempMemoNfEnergisa] = useState("");
  const [tempMemoNfElektro, setTempMemoNfElektro] = useState("");
  const [uploadProgress, setUploadProgress] = useState<
    Record<
      string,
      {
        status: string;
        percent: number;
        fileName: string;
        fileSize: number;
        abortController: AbortController | null;
      }
    >
  >({});
  const [isBillModalOpen, setIsBillModalOpen] = useState(false);
  const [editingBill, setEditingBill] = useState<Partial<BillData> | null>(
    null,
  );
  const fileInputEnergisaRef = useRef<HTMLInputElement>(null);
  const fileInputElektroRef = useRef<HTMLInputElement>(null);
  const folderInputEnergisaRef = useRef<HTMLInputElement>(null);
  const folderInputElektroRef = useRef<HTMLInputElement>(null);
  const [agrupadoraFiles, setAgrupadoraFiles] = useState<
    Record<string, AgrupadoraData>
  >(() => {
    const saved = localStorage.getItem("sanesul_agrupadora_files");
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
    type: "danger" | "info";
    isAlert?: boolean;
  } | null>(null);

  const [isGerenciasModalOpen, setIsGerenciasModalOpen] = useState(false);
  const [isDeleteByListModalOpen, setIsDeleteByListModalOpen] = useState(false);
  const [deleteUcListInput, setDeleteUcListInput] = useState("");
  const [isDeletingByList, setIsDeletingByList] = useState(false);

  const [isMercadoLivreModalOpen, setIsMercadoLivreModalOpen] = useState(false);
  const [mercadoLivreInput, setMercadoLivreInput] = useState("");



  const mirrorAppToSupabase = async () => {
    if (!isSupabaseConfigured) {
      showAlert("Atenção", "O Supabase não está configurado.");
      return;
    }

    const completedBills = bills.filter((b) => b.status === "completed");
    const countApp = completedBills.length;

    showConfirm(
      "Espelhar Faturas no Supabase (Sincronização 1:1)",
      `Esta ação irá garantir que o banco Supabase contenha EXATAMENTE as ${countApp.toLocaleString()} faturas presentes no aplicativo.\n\n` +
      `• Todas as ${countApp.toLocaleString()} faturas do app serão enviadas/atualizadas na nuvem.\n` +
      `• Qualquer fatura antiga que exista na nuvem mas NÃO esteja no app será excluída.\n\n` +
      `Deseja prosseguir com o espelhamento?`,
      async () => {
        setIsSyncing(true);
        try {
          // 1. Obter todos os IDs atualmente no Supabase
          let cloudIds: string[] = [];
          let from = 0;
          let to = 999;
          let finished = false;

          while (!finished) {
            const { data, error } = await supabase
              .from("bills")
              .select("id")
              .range(from, to);

            if (error) throw error;
            if (data && data.length > 0) {
              cloudIds = [...cloudIds, ...data.map((d) => d.id)];
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

          const currentIdsSet = new Set(completedBills.map((b) => b.id));

          // 2. Identifica registros órfãos no Supabase e remove
          const orphanIds = cloudIds.filter((id) => !currentIdsSet.has(id));
          if (orphanIds.length > 0) {
            for (let i = 0; i < orphanIds.length; i += 100) {
              const chunk = orphanIds.slice(i, i + 100);
              const { error: delError } = await supabase.from("bills").delete().in("id", chunk);
              if (delError) console.warn("[Supabase] Erro ao deletar órfãos:", delError);
            }
          }

          // 3. Envia/atualiza todas as faturas do app no Supabase em lotes
          const dbData = completedBills.map((b) => mapBillDataToDb(b));
          for (let i = 0; i < dbData.length; i += 100) {
            const chunk = dbData.slice(i, i + 100);
            const { error: upsertError } = await supabase.from("bills").upsert(chunk, { onConflict: "id" });
            if (upsertError) throw upsertError;
          }

          // 4. Mapeamentos de UCs
          if (ucMappings.length > 0) {
            const mappingsData = ucMappings.map((m) => ({
              uc: String(m.uc),
              gerencia: m.gerencia || "",
              locin: m.locin || "",
              cidade: m.cidade || "",
              updated_at: new Date().toISOString(),
            }));
            await supabase.from("uc_mappings").upsert(mappingsData, { onConflict: "uc" });
          }

          setSupabaseHealth({
            connected: true,
            totalBills: countApp,
            totalUcs: new Set(completedBills.map((b) => b.uc).filter(Boolean)).size,
            lastSync: new Date(),
          });

          showAlert(
            "Sucesso",
            `Espelhamento 1:1 concluído com sucesso!\n\n` +
            `• Faturas no App: ${countApp.toLocaleString()}\n` +
            `• Faturas no Supabase: ${countApp.toLocaleString()}\n` +
            (orphanIds.length > 0 ? `• Faturas órfãs removidas da nuvem: ${orphanIds.length.toLocaleString()}` : "• O banco de dados já estava 100% alinhado.")
          );
        } catch (err: any) {
          console.error("Erro no espelhamento 1:1:", err);
          showAlert("Erro", `Falha ao espelhar com Supabase: ${err.message || err}`);
        } finally {
          setIsSyncing(false);
        }
      },
      "info"
    );
  };

  const saveMercadoLivreUcs = () => {
    const tokens = mercadoLivreInput
      .split(/[\n,;\s\t]+/)
      .map((s) => s.trim())
      .filter((s) => /^\d+$/.test(s)); // apenas números
    const newList = Array.from(new Set(tokens));
    
    // Atualiza global
    UCS_LIVRE_MERCADO_LIVRE = new Set(newList);
    localStorage.setItem("sanesul_mercado_livre_ucs", JSON.stringify(newList));
    
    setIsMercadoLivreModalOpen(false);
    showAlert("Sucesso", "Lista de UCs do Mercado Livre atualizada com sucesso!");
    
    // Atualiza o estado das faturas localmente
    setBills(prev => prev.map(b => ({
      ...b,
      mercado: UCS_LIVRE_MERCADO_LIVRE.has(String(b.uc)) ? "LIVRE" : "CATIVO"
    })));
  };

  const parsedInputUcs = React.useMemo(() => {
    if (!deleteUcListInput.trim()) return [];
    const tokens = deleteUcListInput
      .split(/[\n,;\s\t]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    return Array.from(new Set(tokens));
  }, [deleteUcListInput]);

  const matchedBillsForDeleteList = React.useMemo(() => {
    if (parsedInputUcs.length === 0) {
      return { bills: [] as BillData[], matchedUcs: new Set<string>(), notFoundUcs: [] as string[] };
    }

    const cleanInputSet = new Set(parsedInputUcs.map((u) => u.replace(/[^\d]/g, "")));
    const rawInputSet = new Set(parsedInputUcs.map((u) => u.toLowerCase()));

    const matched: BillData[] = [];
    const matchedUcs = new Set<string>();

    bills.forEach((b) => {
      const cleanUc = (b.uc || "").replace(/[^\d]/g, "");
      const rawUc = (b.uc || "").toLowerCase();
      if ((cleanUc && cleanInputSet.has(cleanUc)) || (rawUc && rawInputSet.has(rawUc))) {
        matched.push(b);
        matchedUcs.add(b.uc);
      }
    });

    const notFoundUcs = parsedInputUcs.filter((u) => {
      const clean = u.replace(/[^\d]/g, "");
      const raw = u.toLowerCase();
      return !Array.from(matchedUcs).some(
        (mu) => (clean && mu.replace(/[^\d]/g, "") === clean) || mu.toLowerCase() === raw,
      );
    });

    return { bills: matched, matchedUcs, notFoundUcs };
  }, [bills, parsedInputUcs]);

  const handleDeleteByUcList = async () => {
    const { bills: targetBills, matchedUcs } = matchedBillsForDeleteList;
    if (targetBills.length === 0) {
      showAlert("Aviso", "Nenhuma fatura encontrada correspondente às UCs informadas.");
      return;
    }

    const targetBillIds = targetBills.map((b) => b.id);
    const ucsCount = matchedUcs.size;
    const billsCount = targetBillIds.length;

    showConfirm(
      "Excluir Faturas por Lista de UCs",
      `Tem certeza que deseja excluir permanentemente ${billsCount} fatura(s) pertencente(s) a ${ucsCount} UC(s)? Esta ação não pode ser desfeita.`,
      async () => {
        setIsDeletingByList(true);
        try {
          // Delete from Supabase if configured
          if (isSupabaseConfigured) {
            try {
              let error = null;
              for (let i = 0; i < targetBillIds.length; i += 100) {
                const chunk = targetBillIds.slice(i, i + 100);
                const { error: err } = await supabase.from("bills").delete().in("id", chunk);
                if (err) error = err;
              }
              if (error) console.error("Erro ao deletar faturas do Supabase:", error);
            } catch (err) {
              console.error("Erro inesperado ao deletar faturas:", err);
            }
          }

          // Remove files from localforage
          localforage
            .getItem<Record<string, File>>("sanesul_bills_files")
            .then((filesMap) => {
              if (filesMap) {
                let hasChanges = false;
                targetBillIds.forEach((id) => {
                  if (filesMap[id]) {
                    delete filesMap[id];
                    hasChanges = true;
                  }
                });
                if (hasChanges) localforage.setItem("sanesul_bills_files", filesMap);
              }
            })
            .catch((e) => console.warn("Erro ao deletar arquivos locais", e));

          // Update state
          setBills((prev) => prev.filter((b) => !targetBillIds.includes(b.id)));
          setSelectedBills((prev) => prev.filter((id) => !targetBillIds.includes(id)));
          if (analysisResults) {
            setAnalysisResults((prev: any) =>
              prev ? prev.filter((r: any) => !targetBillIds.includes(r.id) && !matchedUcs.has(r.uc)) : null,
            );
          }

          setIsDeleteByListModalOpen(false);
          setDeleteUcListInput("");
          showAlert("Sucesso", `${billsCount} fatura(s) de ${ucsCount} UC(s) foram excluídas com sucesso!`);
        } catch (error) {
          console.error("Erro ao excluir lista de UCs:", error);
          showAlert("Erro", "Ocorreu um erro ao excluir as faturas. Verifique o console.");
        } finally {
          setIsDeletingByList(false);
        }
      },
      "danger",
    );
  };
  const [ucMappingSearchTerm, setUcMappingSearchTerm] = useState("");
  const [ucMappings, setUcMappings] = useState<UCLocinMapping[]>(() => {
    const saved = localStorage.getItem("sanesul_uc_mappings");
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          return parsed.map((m: any) => (m.uc === "179291005130" ? { ...m, uc: "179291005130" } : m));
        }
      } catch (e) {}
    }
    return [];
  });

  const syncUcMappingsToSupabase = async (mappings: UCLocinMapping[]) => {
    if (!isSupabaseConfigured || !isAuthenticated) return;
    try {
      const { data: { user } } = await supabase.auth.getUser();
      if (!user) return;
      
      const dbMappings = mappings.map(m => ({
        uc: String(m.uc),
        gerencia: String(m.gerencia || ''),
        locin: String(m.locin || ''),
        cidade: String(m.cidade || ''),
        user_id: user.id,
        updated_at: new Date().toISOString()
      }));
      
      const { error } = await supabase
        .from("uc_mappings")
        .upsert(dbMappings);
        
      if (error) {
        console.warn("[Supabase] Erro ao salvar mapeamentos na nuvem:", error);
      }
    } catch (err) {
      console.warn("[Supabase] Erro na sincronização de mapeamentos:", err);
    }
  };

  const saveUcMappings = (mappings: UCLocinMapping[]) => {
    setUcMappings(mappings);
    syncUcMappingsToSupabase(mappings);
    try {
      localStorage.removeItem("sanesul_uc_mappings");
      if (mappings.length > 0) {
        localStorage.setItem("sanesul_uc_mappings", JSON.stringify(mappings));
      }
    } catch (error) {
      console.warn(
        "Não foi possível salvar os mapeamentos no localStorage",
        error,
      );
    }
  };

  const getGerencia = (uc: string) => {
    const mapping = ucMappings.find((m) => String(m.uc) === String(uc));
    if (mapping && mapping.gerencia) return mapping.gerencia;

    // Fallback: look into the loaded bills state for the gerencia
    const bill = bills.find(
      (b) => String(b.uc) === String(uc) && b.gerencia && b.gerencia !== "---",
    );
    if (bill && bill.gerencia) return bill.gerencia;

    return "---";
  };

  const getLocin = (uc: string) => {
    const mapping = ucMappings.find((m) => String(m.uc) === String(uc));
    if (mapping && mapping.locin) return mapping.locin;

    const bill = bills.find(
      (b) => String(b.uc) === String(uc) && b.locin && b.locin !== "---",
    );
    if (bill && bill.locin) return bill.locin;

    return "---";
  };

  const getCidade = (uc: string, fallback?: string) => {
    const mapping = ucMappings.find((m) => String(m.uc) === String(uc));
    if (mapping && mapping.cidade) return mapping.cidade;

    const bill = bills.find(
      (b) => String(b.uc) === String(uc) && b.cidade && b.cidade !== "---",
    );
    if (bill && bill.cidade) return bill.cidade;

    // Fix UC 117384 - FÁTIMA DO SUL
    if (String(uc) === "93604305181") return "FÁTIMA DO SUL";

    return fallback || "---";
  };

  const handleImportTxtGerencias = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result as string;
      const lines = content.split(/\r?\n/);
      const importedMappings: UCLocinMapping[] = [];
      let skippedLines = 0;

      lines.forEach((line) => {
        if (!line.trim()) return;
        const parts = line.split(/[;,\t]/).map((p) => p.trim());
        // Require at least UC, Gerência, and LOCIN
        if (parts.length >= 3) {
          const uc = parts[0];
          const gerencia = parts[1];
          const locin = parts[2];
          const cidade = parts[3] || "";
          if (uc && gerencia && locin) {
            importedMappings.push({ uc, gerencia, locin, cidade });
          } else {
            skippedLines++;
          }
        } else {
          skippedLines++;
        }
      });

      if (importedMappings.length > 0) {
        // Obter valores numéricos e fundir com os atuais
        setUcMappings((prev) => {
          const existingMap = new Map(prev.map((m) => [m.uc, m]));
          importedMappings.forEach((m) => existingMap.set(m.uc, m));
          const updated = Array.from(existingMap.values());
          syncUcMappingsToSupabase(updated);
          try {
            localStorage.setItem(
              "sanesul_uc_mappings",
              JSON.stringify(updated),
            );
          } catch (error) {
            console.warn(
              "Não foi possível salvar importação no localStorage",
              error,
            );
          }
          return updated;
        });
        showAlert(
          "Sucesso",
          `${importedMappings.length} mapeamentos importados com sucesso.${skippedLines > 0 ? ` (${skippedLines} linhas ignoradas por formato inválido)` : ""}`,
        );
      } else {
        showAlert(
          "Atenção",
          "Nenhum dado válido encontrado. O formato deve ser: UC;Gerência;LOCINS;Cidade",
        );
      }

      // Limpa input
      e.target.value = "";
    };
    reader.readAsText(file);
  };

  const [isSyncingMappings, setIsSyncingMappings] = useState(false);
  const [syncProgress, setSyncProgress] = useState(0);
  const [totalSyncItems, setTotalSyncItems] = useState(0);

  const syncSingleMappingToSupabase = async (mapping: UCLocinMapping) => {
    if (!isSupabaseConfigured || !isAuthenticated) {
      showAlert(
        "Erro",
        "Você precisa estar logado e com o Supabase configurado para sincronizar.",
      );
      return;
    }

    try {
      const { error } = await supabase
        .from("bills")
        .update({
          cidade: mapping.cidade || null,
          gerencia: mapping.gerencia,
          locins: mapping.locin,
        })
        .eq("uc", mapping.uc);

      if (error) {
        console.error(`Erro ao sincronizar UC ${mapping.uc}:`, error);
        showAlert(
          "Erro na Sincronização",
          `Ocorreu um erro ao sincronizar a UC ${mapping.uc}. Verifique as permissões.`,
        );
      } else {
        setBills((prev) =>
          prev.map((bill) => {
            if (bill.uc === mapping.uc) {
              return {
                ...bill,
                cidade: mapping.cidade || bill.cidade,
                gerencia: mapping.gerencia,
                locin: mapping.locin, // Changed to match local schema
              };
            }
            return bill;
          }),
        );
        showAlert("Sucesso", `UC ${mapping.uc} sincronizada com sucesso!`);
      }
    } catch (e) {
      console.error(e);
      showAlert("Erro", "Erro ao comunicar com o banco de dados.");
    }
  };

  const syncMappingsToSupabase = () => {
    if (!isSupabaseConfigured || !isAuthenticated) {
      showAlert(
        "Erro",
        "Você precisa estar logado e com o Supabase configurado para sincronizar.",
      );
      return;
    }

    showConfirm(
      "Sincronizar com Banco de Dados",
      "Isso irá atualizar as colunas Gerência e LOCIN de TODAS as faturas no banco de dados. Este processo pode demorar alguns minutos. Deseja continuar?",
      async () => {
        setIsSyncingMappings(true);
        setTotalSyncItems(ucMappings.length);
        setSyncProgress(0);
        let successCount = 0;
        let errorCount = 0;

        try {
          for (let i = 0; i < ucMappings.length; i++) {
            const mapping = ucMappings[i];
            const { error } = await supabase
              .from("bills")
              .update({
                cidade: mapping.cidade || null,
                gerencia: mapping.gerencia,
                locins: mapping.locin,
              })
              .eq("uc", mapping.uc);

            if (error) {
              console.error(`Erro ao sincronizar UC ${mapping.uc}:`, error);
              errorCount++;
            } else {
              successCount++;
            }
            setSyncProgress(i + 1);
          }

          // Update local state to reflect changes without reloading
          setBills((prev) =>
            prev.map((bill) => {
              const mapping = ucMappings.find((m) => m.uc === bill.uc);
              if (mapping) {
                return {
                  ...bill,
                  gerencia: mapping.gerencia,
                  locin: mapping.locin,
                };
              }
              return bill;
            }),
          );

          if (errorCount > 0) {
            showAlert(
              "Atenção: Sincronização com Falhas",
              `Sucesso: ${successCount}. Falhas: ${errorCount}. \nSe este for o seu primeiro sincronismo, verifique se as colunas 'gerencia', 'locins' e 'cidade' existem na tabela 'bills' do seu Supabase.`,
            );
          } else {
            showAlert(
              "Sincronização Finalizada",
              `Sucesso: ${successCount} UCs atualizadas com sucesso!`,
            );
          }
        } catch (err: any) {
          showAlert("Erro", `Falha durante a sincronização: ${err.message}`);
        } finally {
          setIsSyncingMappings(false);
          setSyncProgress(0);
          setTotalSyncItems(0);
        }
      },
      "info",
    );
  };

  const showAlert = (title: string, message: string) => {
    setConfirmModalData({
      title,
      message,
      onConfirm: () => setShowConfirmModal(false),
      type: "info",
      isAlert: true,
    });
    setShowConfirmModal(true);
  };

  const showConfirm = (
    title: string,
    message: string,
    onConfirm: () => void,
    type: "danger" | "info" = "info",
  ) => {
    setConfirmModalData({
      title,
      message,
      onConfirm: () => {
        onConfirm();
        setShowConfirmModal(false);
      },
      type,
      isAlert: false,
    });
    setShowConfirmModal(true);
  };

  React.useEffect(() => {
    localStorage.setItem(
      "sanesul_agrupadora_files",
      JSON.stringify(agrupadoraFiles),
    );
  }, [agrupadoraFiles]);
  const agrupadoraInputRef = useRef<HTMLInputElement>(null);
  const energisaInputRef = useRef<HTMLInputElement>(null);
  const detailedElektroInputRef = useRef<HTMLInputElement>(null);

  const handleAgrupadoraUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
    reportType: "summary" | "detailed" = "summary",
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const fileId = `${file.name}-${Date.now()}`;
    const abortController = new AbortController();
    const statusPrefix =
      reportType === "detailed" ? "Relatório Detalhado" : "Fatura Agrupadora";

    setUploadProgress((prev) => ({
      ...prev,
      [fileId]: {
        status: `Lendo ${statusPrefix}...`,
        percent: 0,
        fileName: file.name,
        fileSize: file.size,
        abortController,
      },
    }));

    const apiKey =
      (typeof process !== "undefined"
        ? process.env.GEMINI_API_KEY || process.env.API_KEY
        : "") ||
      import.meta.env.VITE_GEMINI_API_KEY ||
      "";
    const ai = new GoogleGenAI({ apiKey });

    try {
      const reader = new FileReader();
      const base64Promise = new Promise<string>((resolve, reject) => {
        const timeout = setTimeout(
          () =>
            reject(
              new Error("Erro ao ler arquivo: Tempo limite excedido (60s)"),
            ),
          60000,
        );

        reader.onprogress = (data) => {
          if (data.lengthComputable) {
            const progress = Math.round((data.loaded / data.total) * 30);
            setUploadProgress((prev) => ({
              ...prev,
              [fileId]: {
                ...prev[fileId],
                status: `Lendo ${statusPrefix}...`,
                percent: progress,
              },
            }));
          }
        };
        reader.onload = () => {
          clearTimeout(timeout);
          const base64 = (reader.result as string).split(",")[1];
          setUploadProgress((prev) => ({
            ...prev,
            [fileId]: {
              ...prev[fileId],
              status: "Processando com IA...",
              percent: 30,
            },
          }));
          resolve(base64);
        };
        reader.onerror = () => {
          clearTimeout(timeout);
          reject(new Error("Erro ao ler arquivo"));
        };
        reader.readAsDataURL(file);
      });

      const base64Data = await base64Promise;

      let prompt =
        "Você é um especialista em faturas agrupadoras de energia elétrica. Sua tarefa é extrair os dados consolidados desta fatura.\n\nINSTRUÇÕES:\n1. CONCESSIONÁRIA: Identifique se é ELEKTRO ou ENERGISA.\n2. VALOR TOTAL: Extraia o valor total a pagar da fatura agrupadora.\n3. REFERÊNCIA: Identifique o mês e ano de referência (ex: Fevereiro/2026).\n4. NOTA FISCAL: Procure pelo número da Nota Fiscal ou Fatura (ex: AGP-01... ou similar).\n5. IMPOSTOS: Extraia os valores de PIS, COFINS, ICMS e CIP. Para faturas da Energisa, os impostos federais (PIS/COFINS) podem estar agrupados como 'Imp. Fed.'.\n\nSe algum valor não for encontrado, retorne 0 ou string vazia.\n\nIMPORTANTE: SEMPRE RESPONDA EM PORTUGUÊS.";

      if (reportType === "detailed") {
        prompt =
          "VOCÊ É UM AUDITOR CONTÁBIL ESPECIALISTA EM FATURAS DE ENERGIA. Sua tarefa é analisar TODAS AS PÁGINAS deste relatório detalhado para consolidar o valor da CIP.\n\nINSTRUÇÕES DETALHADAS:\n1. Percorra TODAS as páginas do documento, sem exceção.\n2. Em cada página, localize a tabela de itens faturados.\n3. Procure pelas descrições: 'COBRANCA ILUM PUBLICA', 'CIP', 'ILUMINACAO PUBLICA' ou 'CONTRIBUIÇÃO DE ILUMINAÇÃO PÚBLICA'.\n4. Extraia o valor monetário associado a cada uma dessas linhas.\n5. SOMA TOTAL: Você deve somar TODOS os valores encontrados em todas as páginas para obter o total da CIP do grupo.\n6. RETORNO: Retorne o JSON preenchendo o campo 'cip' com a soma total calculada. Os campos 'valorTotal', 'pis', 'cofins', 'icms' devem ser preenchidos como 0, a menos que você encontre um valor consolidado claro para eles no documento.\n7. Identifique a 'concessionaria' e o 'mesReferencia'.\n\nIMPORTANTE: SEMPRE RESPONDA EM PORTUGUÊS.";
      }

      if (abortController.signal.aborted) throw new Error("Upload cancelado");

      const progressInterval = setInterval(() => {
        setUploadProgress((prev) => {
          if (!prev || !prev[fileId] || prev[fileId].percent >= 95) {
            clearInterval(progressInterval);
            return prev;
          }
          return {
            ...prev,
            [fileId]: { ...prev[fileId], percent: prev[fileId].percent + 5 },
          };
        });
      }, 1000);

      let response;
      try {
        response = await generateContentWithRetry(ai, {
          model: "gemini-2.5-flash",
          contents: [
            {
              parts: [
                {
                  inlineData: {
                    mimeType: file.type || "application/pdf",
                    data: base64Data,
                  },
                },
                {
                  text: "Extraia os dados desta fatura seguindo as instruções do sistema.",
                },
              ],
            },
          ],
          config: {
            responseMimeType: "application/json",
            responseSchema: AGRUPADORA_SCHEMA,
            thinkingConfig: { thinkingLevel: ThinkingLevel.LOW },
            systemInstruction: prompt,
          },
        });
      } finally {
        clearInterval(progressInterval);
      }

      if (abortController.signal.aborted) throw new Error("Upload cancelado");

      setUploadProgress((prev) => ({
        ...prev,
        [fileId]: { ...prev[fileId], status: "Concluído!", percent: 100 },
      }));
      setTimeout(
        () =>
          setUploadProgress((prev) => {
            const next = { ...prev };
            delete next[fileId];
            return next;
          }),
        2000,
      );

      let text = response.text || "{}";
      text = text
        .replace(/```json\n?/g, "")
        .replace(/```\n?/g, "")
        .trim();
      const result = JSON.parse(text);
      const concessionariaRaw = (
        result.concessionaria || "DESCONHECIDA"
      ).toUpperCase();
      const key = concessionariaRaw.includes("ENERGISA")
        ? "ENERGISA"
        : "ELEKTRO";

      if (reportType === "detailed") {
        const detailedKey = `${key}_DETALHADO`;
        setAgrupadoraFiles((prev) => ({
          ...prev,
          [detailedKey]: {
            cip:
              typeof result.cip === "string"
                ? parseValue(result.cip)
                : result.cip,
            concessionaria: `${concessionariaRaw} (DETALHADO)`,
            mesReferencia: formatMonth(result.mesReferencia || ""),
            valorTotal: 0,
            vencimento: "",
            numeroNotaFiscal: "",
            pis: 0,
            cofins: 0,
            icms: 0,
            fileName: file.name,
          },
        }));
      } else {
        const newData: AgrupadoraData = {
          concessionaria: concessionariaRaw,
          valorTotal:
            typeof result.valorTotal === "string"
              ? parseValue(result.valorTotal)
              : result.valorTotal,
          mesReferencia: formatMonth(result.mesReferencia || ""),
          vencimento: result.vencimento || "",
          numeroNotaFiscal: result.numeroNotaFiscal || "",
          pis:
            typeof result.pis === "string"
              ? parseValue(result.pis)
              : result.pis,
          cofins:
            typeof result.cofins === "string"
              ? parseValue(result.cofins)
              : result.cofins,
          icms:
            typeof result.icms === "string"
              ? parseValue(result.icms)
              : result.icms,
          cip:
            typeof result.cip === "string"
              ? parseValue(result.cip)
              : result.cip,
          fileName: file.name,
        };

        setAgrupadoraFiles((prev) => ({
          ...prev,
          [key]: newData,
        }));
      }
    } catch (error: any) {
      console.error("Agrupadora extraction error:", {
        fileName: file.name,
        fileType: file.type,
        fileSize: file.size,
        reportType,
        error,
      });
      const errorStr =
        error?.message ||
        (typeof error === "string" ? error : JSON.stringify(error));

      if (
        error?.isQuotaError ||
        errorStr.includes("quota") ||
        errorStr.includes("429") ||
        errorStr.includes("RESOURCE_EXHAUSTED")
      ) {
        showAlert(
          "Limite Atingido",
          "Cota da API excedida. Verifique seu plano e detalhes de faturamento no Google AI Studio. Se você já tem um plano pago, aguarde alguns minutos.",
        );
      } else {
        showAlert("Erro", "Erro ao processar fatura agrupadora: " + errorStr);
      }
      setUploadProgress(null);
    } finally {
      setIsProcessing(false);
      if (agrupadoraInputRef.current) agrupadoraInputRef.current.value = "";
      if (energisaInputRef.current) energisaInputRef.current.value = "";
      if (detailedElektroInputRef.current)
        detailedElektroInputRef.current.value = "";
    }
  };

  const downloadExcelTemplate = () => {
    const ws = XLSX.utils.json_to_sheet([{}], {
      header: EXCEL_COLUMNS.map((c) => c.header),
    });
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
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });

        if (data.length < 2) {
          alert("A planilha está vazia ou não possui dados.");
          return;
        }

        const headers = data[0] as string[];
        const rows = data.slice(1) as any[][];

        const newBills: BillData[] = rows
          .map((row, index) => {
            const billData: any = {
              id: crypto.randomUUID(),
              fileName: `Importado_${file.name}_Linha_${index + 1}`,
              status: "completed",
              createdAt: Date.now() + index,
              tipo: "normal",
            };

            headers.forEach((header, colIndex) => {
              const columnDef = EXCEL_COLUMNS.find((c) => c.header === header);
              if (columnDef) {
                const value = row[colIndex];
                billData[columnDef.key] =
                  value !== undefined && value !== null ? String(value) : "";
              }
            });

            return billData as BillData;
          })
          .map((bill) => {
            // Enriquecer com mapeamentos de Gerência e LOCIN
            const mapping = ucMappings.find((m) => m.uc === String(bill.uc));
            if (mapping) {
              bill.gerencia = mapping.gerencia;
              bill.locin = mapping.locin;
              if (!bill.cidade) bill.cidade = mapping.cidade;
            }

            // Fix UC 117384 - FÁTIMA DO SUL
            if (String(bill.uc) === "93604305181") {
              bill.cidade = "FÁTIMA DO SUL";
            }

            // Fix UC 9000943 - Demand Contratada Ponta is 0 (Verde)
            if (String(bill.uc) === "163738105109") {
              bill.demandaPontaKW = "0";
            }

            bill.mercado =
              bill.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(bill.uc))
                ? "LIVRE"
                : "CATIVO";
            if (bill.uc && UCS_OPER.has(String(bill.uc))) {
              bill.tipo = "OPER";
            } else if (
              bill.uc &&
              UCS_LIVRE_MERCADO_LIVRE.has(String(bill.uc))
            ) {
              let mod = bill.modalidadeTarifaria || "";
              if (!mod.toUpperCase().includes("LIVRE")) {
                bill.modalidadeTarifaria = mod ? `${mod} - LIVRE` : "LIVRE";
              }
              bill.tipo = "LIVRE";
            }
            return bill;
          });

        if (isSupabaseConfigured && isAuthenticated) {
          const { data: userData } = await supabase.auth.getUser();
          if (userData.user) {
            const dbData = newBills.map((bill) => ({
              id: bill.id,
              user_id: userData.user!.id,
              file_name: bill.fileName,
              uc: bill.uc || "",
              demanda_ponta_kw: bill.demandaPontaKW || "",
              demanda_fora_ponta_kw: bill.demandaForaPontaKW || "",
              demanda_potencia_medida_ponta:
                bill.demandaPotenciaMedidaPonta || "",
              demanda_potencia_medida_fora_ponta:
                bill.demandaPotenciaMedidaForaPonta || "",
              ano_leitura: bill.anoLeitura || "",
              mes_referencia: bill.mesReferencia || "",
              consumo_kwh_ponta: bill.consumoKwhPonta || "",
              consumo_kwh_fora_ponta: bill.consumoKwhForaPonta || "",
              consumo_kwh_grupo_b: bill.consumoKwhGrupoB || bill.consumoKwh || "",
              valor_consumo_kwh_ponta: bill.valorConsumoKwhPonta || "",
              valor_consumo_kwh_fora_ponta: bill.valorConsumoKwhForaPonta || "",
              valor_consumo_kwh_grupo_b: bill.valorConsumoKwhGrupoB || "",
              valor_total: bill.valorTotal || "",
              cidade: bill.cidade || "",
              demanda_potencia_nao_consumida_ponta:
                bill.demandaPotenciaNaoConsumidaPonta || "",
              demanda_potencia_nao_consumida_f_ponta:
                bill.demandaPotenciaNaoConsumidaFPonta || "",
              demanda_potencia_ativa_ultrap_ponta:
                bill.demandaPotenciaAtivaUltrapPonta || "",
              demanda_potencia_ativa_ultrap_f_ponta:
                bill.demandaPotenciaAtivaUltrapFPonta || "",
              energia_reativa_exced_ponta: bill.energiaReativaExcedPonta || "",
              energia_reativa_exced_f_ponta:
                bill.energiaReativaExcedFPonta || "",
              valor_demanda_potencia_medida_ponta:
                bill.valorDemandaPotenciaMedidaPonta || "",
              valor_demanda_potencia_medida_fora_ponta:
                bill.valorDemandaPotenciaMedidaForaPonta || "",
              valor_demanda_potencia_nao_consumida_ponta:
                bill.valorDemandaPotenciaNaoConsumidaPonta || "",
              valor_demanda_potencia_nao_consumida_f_ponta:
                bill.valorDemandaPotenciaNaoConsumidaFPonta || "",
              valor_demanda_potencia_ativa_ultrap_ponta:
                bill.valorDemandaPotenciaAtivaUltrapPonta || "",
              valor_demanda_potencia_ativa_ultrap_f_ponta:
                bill.valorDemandaPotenciaAtivaUltrapFPonta || "",
              valor_energia_reativa_exced_ponta:
                bill.valorEnergiaReativaExcedPonta || "",
              valor_energia_reativa_exced_f_ponta:
                bill.valorEnergiaReativaExcedFPonta || "",
              energia_atv_injetada_gdi_ouc: bill.energiaAtvInjetadaGDIOUC || "",
              valor_energia_atv_injetada_gdi_ouc:
                bill.valorEnergiaAtvInjetadaGDIOUC || "",
              energia_atv_injetada_gdi_muc: bill.energiaAtvInjetadaGDIMUC || "",
              valor_energia_atv_injetada_gdi_muc:
                bill.valorEnergiaAtvInjetadaGDIMUC || "",
              cip: bill.cip || "",
              outros_encargos: bill.outrosEncargos || "",
              pis: bill.pis || "",
              cofins: bill.cofins || "",
              icms: bill.icms || "",
              concessionaria: bill.concessionaria || "",
              numero_nota_fiscal: bill.numeroNotaFiscal || "",
              modalidade_tarifaria: bill.modalidadeTarifaria || "",
              subgrupo: bill.subgrupo || "",
              tipo: bill.tipo || "normal",
              mercado: bill.mercado || "",
              gerencia: bill.gerencia || "",
              locins: bill.locin || "",
              data_vencimento: bill.dataVencimento || "",
              status: bill.status,
              created_at: new Date(bill.createdAt || Date.now()).toISOString(),
            }));

            let { error } = await supabase.from("bills").insert(dbData);

            if (
              error &&
              (error.message.includes("data_vencimento") ||
                error.message.includes("mercado") ||
                error.message.includes("gerencia") ||
                error.message.includes("locins") ||
                error.message.includes("consumo_kwh_grupo_b") ||
                error.message.includes("valor_consumo_kwh_grupo_b") ||
                error.details?.includes("data_vencimento") ||
                error.details?.includes("mercado") ||
                error.details?.includes("gerencia") ||
                error.details?.includes("locins") ||
                error.details?.includes("consumo_kwh_grupo_b") ||
                error.details?.includes("valor_consumo_kwh_grupo_b") ||
                error.code === "PGRST204")
            ) {
              console.warn(
                "Colunas novas não encontradas no Supabase. Inserindo sem elas...",
              );
              const fallbackData = dbData.map((d: any) => {
                const { data_vencimento, mercado, gerencia, locins, consumo_kwh_grupo_b, valor_consumo_kwh_grupo_b, ...rest } = d;
                if (consumo_kwh_grupo_b) rest.consumo_kwh_ponta = consumo_kwh_grupo_b;
                if (valor_consumo_kwh_grupo_b) rest.valor_consumo_kwh_ponta = valor_consumo_kwh_grupo_b;
                return rest;
              });
              const fallbackRes = await supabase
                .from("bills")
                .insert(fallbackData);
              error = fallbackRes.error;
            }

            if (error) {
              console.error("Erro ao salvar no Supabase:", error);
              alert("Erro ao salvar os dados importados no banco de dados.");
            }
          }
        }

        setBills((prev) => [...prev, ...newBills]);
        alert(`${newBills.length} faturas importadas com sucesso!`);
      } catch (error) {
        console.error("Erro ao importar planilha:", error);
        alert("Erro ao processar o arquivo Excel.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const addFiles = (files: FileList | File[], concessionaria?: string) => {
    const now = Date.now();
    const validFiles = (Array.from(files) as File[]).filter(f => 
      f.name.toLowerCase().endsWith('.pdf') || 
      f.type.startsWith('image/') || 
      f.name.toLowerCase().match(/\.(png|jpe?g)$/)
    );
    
    if (validFiles.length === 0) return;

    const newBills: BillData[] = validFiles.map(
      (file, index) =>
        ({
          id: crypto.randomUUID(),
          createdAt: now + index,
          fileName: file.name,
          concessionaria: concessionaria || "",
          uc: "",
          demandaPontaKW: "",
          demandaForaPontaKW: "",
          demandaPotenciaMedidaPonta: "",
          demandaPotenciaMedidaForaPonta: "",
          anoLeitura: "",
          mesReferencia: "",
          consumoKwhPonta: "",
          valorConsumoKwhPonta: "",
          consumoKwhForaPonta: "",
          valorConsumoKwhForaPonta: "",
          valorTotal: "",
          cidade: "",
          demandaPotenciaNaoConsumidaPonta: "",
          demandaPotenciaNaoConsumidaFPonta: "",
          demandaPotenciaAtivaUltrapPonta: "",
          demandaPotenciaAtivaUltrapFPonta: "",
          energiaReativaExcedPonta: "",
          energiaReativaExcedFPonta: "",
          valorDemandaPotenciaMedidaPonta: "",
          valorDemandaPotenciaMedidaForaPonta: "",
          valorDemandaPotenciaNaoConsumidaPonta: "",
          valorDemandaPotenciaNaoConsumidaFPonta: "",
          valorDemandaPotenciaAtivaUltrapPonta: "",
          valorDemandaPotenciaAtivaUltrapFPonta: "",
          valorEnergiaReativaExcedPonta: "",
          valorEnergiaReativaExcedFPonta: "",
          energiaAtvInjetadaGDIOUC: "",
          valorEnergiaAtvInjetadaGDIOUC: "",
          energiaAtvInjetadaGDIMUC: "",
          valorEnergiaAtvInjetadaGDIMUC: "",
          cip: "",
          outrosEncargos: "",
          status: "pending",
          file: file,
        }) as any,
    );

    setBills((prev) => deduplicateBills([...prev, ...newBills]));
  };

  const handleFileUpload = (
    event: React.ChangeEvent<HTMLInputElement>,
    concessionaria?: string,
  ) => {
    const files = event.target.files;
    if (!files) return;
    addFiles(files, concessionaria);
    event.target.value = "";
  };

  const runAnalysis = () => {
    const completedBills = bills.filter((b) => b.status === "completed");
    if (completedBills.length === 0) return;

    // Filter out UCs that are clearly Grupo B (sem demanda contratada)
    const validUCs = new Set(
      Array.from(new Set(completedBills.map((b) => b.uc))).filter((uc) => {
        const ucBills = completedBills.filter((b) => b.uc === uc);
        const grupoBCount = ucBills.filter((b) =>
          (b.subgrupo || "").toUpperCase().startsWith("B"),
        ).length;
        const grupoACount = ucBills.filter((b) =>
          (b.subgrupo || "").toUpperCase().startsWith("A"),
        ).length;

        if (grupoBCount > 0 && grupoACount === 0) {
          return false;
        }
        return true;
      })
    );

    const validBills = completedBills.filter(b => validUCs.has(b.uc));

    const parsedData = validBills
      .map((b) => {
        let dmp = parseValue(b.demandaPotenciaMedidaPonta);

        // Correção manual para UC 3401807 em Dezembro/2025 conforme solicitado pelo usuário
        // O usuário informou que houve uma leitura incorreta da demanda medida ponta neste mês.
        const isUC3401807 = String(b.uc) === "9452805127";
        const isDec2025 =
          (b.mesReferencia?.toLowerCase() === "dezembro" ||
            b.mesReferencia === "12" ||
            b.mesReferencia === "12/2025") &&
          b.anoLeitura === "2025";

        if (isUC3401807 && isDec2025) {
          // Se a leitura foi incorreta (provavelmente um pico errôneo), ajustamos para 0
          // para que não influencie o cálculo da demanda ideal (que usa o máximo do período).
          dmp = 0;
        }

        const modalidade = (b.modalidadeTarifaria || "").toUpperCase();

        return {
          fileName: b.fileName,
          mes: b.mesReferencia || "N/A",
          ano: b.anoLeitura || "",
          uc: b.uc || "N/A",
          dcp: parseValue(b.demandaPontaKW),
          dmp: dmp,
          dcfp: parseValue(b.demandaForaPontaKW),
          dmfp: parseValue(b.demandaPotenciaMedidaForaPonta),
          modalidade: modalidade,
          // Valores financeiros para o Gasto Real
          vDmpP: parseValue(b.valorDemandaPotenciaMedidaPonta),
          vDmfpFP: parseValue(b.valorDemandaPotenciaMedidaForaPonta),
          vUltrapP: parseValue(b.valorDemandaPotenciaAtivaUltrapPonta),
          vUltrapFP: parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta),
          vNaoConsP: parseValue(b.valorDemandaPotenciaNaoConsumidaPonta),
          vNaoConsFP: parseValue(b.valorDemandaPotenciaNaoConsumidaFPonta),
          tipo: UCS_PPP.has(String(b.uc)) ? "PPP Fotovoltaica" : b.tipo || "",
          mercado: UCS_LIVRE_MERCADO_LIVRE.has(b.uc) ? "LIVRE" : "CATIVO",
          city: getCidade(String(b.uc)),
        };
      })
      .filter((d) => d.dcp > 0 || d.dcfp > 0);

    // 2. Calculate Optimal Fixed Demand (Resolução 1000)
    // The optimal fixed demand is the one that minimizes the total cost over the period.
    // Group by UC to find the optimal demand per UC
    const ucs = Array.from(new Set(parsedData.map((d) => d.uc)));
    const optimalDemands: Record<string, { ponta: number; foraPonta: number }> =
      {};

    ucs.forEach((uc) => {
      const ucData = parsedData.filter((d) => d.uc === uc);
      // Verifica se a UC tem contrato de ponta (ex: Tarifa Azul)
      // Se dcp for 0 em todos os meses, não sugerimos valor para ponta (ex: Tarifa Verde)
      // NOVO: Se a modalidade for VERDE, forçamos hasPontaContract para false
      const isVerde = ucData.some((d) => d.modalidade.includes("VERDE"));
      const isAzul = ucData.some((d) => d.modalidade.includes("AZUL"));
      const hasPontaContract = !isVerde && ucData.some((d) => d.dcp > 0);

      const getOptimalForDemandsSimulation = (
        measurements: number[],
        isVerdeCheck: boolean,
        isAzulCheck: boolean,
        isPonta: boolean
      ) => {
        if (measurements.length === 0) return 0;
        
        // Define as tarifas base e ultrapassagem para a simulação
        let tariff = 0;
        let penaltyTariff = 0;
        
        if (isAzulCheck) {
          if (isPonta) {
            tariff = 85.53;
            penaltyTariff = 171.07;
          } else {
            tariff = 42.90;
            penaltyTariff = 85.810360;
          }
        } else if (isVerdeCheck) {
          if (isPonta) {
            tariff = 0; // Verde doesn't have Ponta contracted demand
            penaltyTariff = 0;
          } else {
            tariff = 43.17715;
            penaltyTariff = 43.17715 * 2;
          }
        } else {
          tariff = 10.0;
          penaltyTariff = 20.0;
        }
        
        if (tariff === 0) return 0;
        
        const minM = Math.min(...measurements);
        const maxM = Math.max(...measurements);
        
        let bestDemand = 0;
        let minCost = Infinity;
        
        let startCand = Math.floor(Math.max(0, minM / 1.05));
        if (!isPonta && startCand < 30) startCand = 30; // Min demand A group
        const endCand = Math.ceil(maxM);
        const actualEndCand = (!isPonta && endCand < 30) ? 30 : endCand;

        for (let cand = startCand; cand <= actualEndCand; cand += 0.5) {
            let totalCost = 0;
            for (const m of measurements) {
                if (m > cand * 1.05) {
                    totalCost += (cand * tariff) + ((m - cand) * penaltyTariff);
                } else {
                    totalCost += Math.max(m, cand) * tariff;
                }
            }
            if (totalCost < minCost) {
                minCost = totalCost;
                bestDemand = cand;
            }
        }
        return bestDemand;
      };

      // Utilize os ultimos 12 meses para encontrar a demanda otimizada!
      const last12MonthsData = ucData.slice(-12);
      const measuredPonta = last12MonthsData.map((d) => d.dmp);
      const measuredForaPonta = last12MonthsData.map((d) => d.dmfp);

      optimalDemands[String(uc)] = {
        ponta: hasPontaContract ? getOptimalForDemandsSimulation(measuredPonta, isVerde, isAzul, true) : 0,
        foraPonta: getOptimalForDemandsSimulation(measuredForaPonta, isVerde, isAzul, false),
      };
    });

    // 3. Second Pass: Calculate costs and savings based on the fixed optimal demand
    const tp = 15.0;
    const tfp = 10.0;

    const results = parsedData.flatMap((row) => {
      const {
        mes,
        ano,
        uc,
        dcp,
        dmp,
        dcfp,
        dmfp,
        vDmpP,
        vDmfpFP,
        vUltrapP,
        vUltrapFP,
        vNaoConsP,
        vNaoConsFP,
        modalidade,
        city,
      } = row;
      const opt = optimalDemands[String(uc)];

      if (!opt) return [];

      // Gasto Real conforme solicitado: Soma dos valores financeiros da fatura
      const currentTotal =
        vDmpP + vDmfpFP + vUltrapP + vUltrapFP + vNaoConsP + vNaoConsFP;

      // Para o cálculo da economia, precisamos dos custos base (Ponta e Fora Ponta)
      // que compõem esse Gasto Real.
      const costPonta = vDmpP + vUltrapP + vNaoConsP;
      const costForaPonta = vDmfpFP + vUltrapFP + vNaoConsFP;

      // Definir tarifas ideais com base na modalidade (Valores solicitados pelo usuário)
      const isAzul = modalidade.includes("AZUL");
      const isVerde = modalidade.includes("VERDE");

      let tp_ideal = 0;
      let tfp_ideal = 0;
      let tp_ideal_multa = 0;
      let tfp_ideal_multa = 0;

      if (isAzul) {
          tp_ideal = 85.53;
          tp_ideal_multa = 171.07;
          tfp_ideal = 42.90;
          tfp_ideal_multa = 85.810360;
      } else if (isVerde) {
          tp_ideal = 0;
          tp_ideal_multa = 0;
          tfp_ideal = 43.17715;
          tfp_ideal_multa = 43.17715 * 2;
      } else {
          tp_ideal = 15.0;
          tp_ideal_multa = 30.0;
          tfp_ideal = 10.0;
          tfp_ideal_multa = 20.0;
      }

      let optimizedTotal = 0;
      let economy = 0;

      if (isVerde) {
        // Cálculo de economia considerando a regra tarifária ANEEL limitando em 1.05 e aplicando multa de atraso.
        const optCostForaPonta =
          dmfp > opt.foraPonta * 1.05
            ? opt.foraPonta * tfp_ideal + (dmfp - opt.foraPonta) * tfp_ideal_multa
            : Math.max(dmfp, opt.foraPonta) * tfp_ideal;

        optimizedTotal = vDmpP + optCostForaPonta;
        economy = currentTotal - optimizedTotal;
      } else {
        // Optimized Cost (using the FIXED optimal demand and specific rates)
        const optCostPonta =
          dcp > 0 && opt.ponta > 0 && tp_ideal > 0
            ? dmp > opt.ponta * 1.05
              ? opt.ponta * tp_ideal + (dmp - opt.ponta) * tp_ideal_multa
              : Math.max(dmp, opt.ponta) * tp_ideal
            : 0;

        const optCostForaPonta =
          dmfp > opt.foraPonta * 1.05
            ? opt.foraPonta * tfp_ideal + (dmfp - opt.foraPonta) * tfp_ideal_multa
            : Math.max(dmfp, opt.foraPonta) * tfp_ideal;

        optimizedTotal = optCostPonta + optCostForaPonta;
        economy = currentTotal - optimizedTotal;
      }

      // Ultrapassagem
      const overrunPonta = dcp > 0 && dmp > dcp * 1.05 ? dmp - dcp : 0;
      const overrunForaPonta = dmfp > dcfp * 1.05 ? dmfp - dcfp : 0;

      // Subutilização
      const subPonta = dcp > 0 && dmp < dcp ? dcp - dmp : 0;
      const subForaPonta = dmfp < dcfp ? dcfp - dmfp : 0;

      return [
        {
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
          mercado: row.mercado,
          city,
        },
      ];
    });

    results.sort((a, b) => {
      const yearA = parseInt(a.ano || "0", 10);
      const yearB = parseInt(b.ano || "0", 10);
      if (yearA !== yearB) return yearB - yearA;

      const monthA = getMonthNumber(a.mes);
      const monthB = getMonthNumber(b.mes);
      return monthB - monthA;
    });

    setAnalysisResults(results);
  };

  const exportMonitoramentoExcel = () => {
    if (!monitoringResults) return;

    let html = `
      <html xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <style>
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid black; padding: 5px; text-align: left; }
          th { background-color: #f2f2f2; font-weight: bold; }
          .header { font-size: 16px; font-weight: bold; background-color: #d9e1f2; }
          .negative { color: red; }
          .positive { color: green; }
        </style>
      </head>
      <body>
    `;

    // 1. Prejuízos Encontrados
    html += `
      <table>
        <tr><td colspan="4" class="header">PREJUÍZOS ENCONTRADOS</td></tr>
        <tr>
          <th>Cidade</th>
          <th>UC</th>
          <th>Prejuízo (R$)</th>
        </tr>
    `;

    const cityLosses = monitoringResults.cityData.filter(
      (c) => c.negativeEconomy < 0,
    );
    cityLosses.forEach((city) => {
      city.negativeUcs.forEach((ucData: any) => {
        html += `
          <tr>
            <td>${city.city}</td>
            <td>${ucData.uc}</td>
            <td class="negative">${ucData.economy.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</td>
          </tr>
        `;
      });
    });

    if (cityLosses.length === 0) {
      html += `<tr><td colspan="3">Nenhum prejuízo identificado.</td></tr>`;
    }

    html += `</table><br/><br/>`;

    // 2. Histórico de Alterações Mensais (Apenas UCs com Prejuízo)
    html += `
      <table>
        <tr><td colspan="10" class="header" style="font-size: 18px; text-align: center;">HISTÓRICO MENSAL INDIVIDUAL DAS UCs COM PREJUÍZO</td></tr>
    `;

    const lossesUCs = monitoringResults.changedUCs
      .filter((uc: any) => uc.totalEconomy < 0)
      .sort((a: any, b: any) => String(a.gerencia || "").localeCompare(String(b.gerencia || "")));

    lossesUCs.forEach((uc: any) => {
      // Linha de Cabeçalho da UC
      html += `
        <tr><td colspan="10" style="background-color: #e2efda; font-weight: bold; font-size: 16px; border-bottom: 2px solid #548235; border-top: 2px solid #548235;">UC: ${uc.uc} | Cidade: ${uc.city} | Gerência: ${uc.gerencia}</td></tr>
        <tr>
          <th>Mês/Ano</th>
          <th>Contratada P (kW)</th>
          <th>Contratada FP (kW)</th>
          <th>Medida P (kW)</th>
          <th>Medida FP (kW)</th>
          <th>Gasto Real (R$)</th>
          <th>Ref. Anterior (R$)</th>
          <th>Economia (R$)</th>
          <th>Demanda Ideal P (kW)</th>
          <th>Demanda Ideal FP (kW)</th>
        </tr>
      `;

      uc.monthlyData.forEach((monthData: any) => {
        const economyClass =
          monthData.economy < 0
            ? "negative"
            : monthData.economy > 0
              ? "positive"
              : "";
        html += `
          <tr>
            <td>${monthData.mes}/${monthData.ano}</td>
            <td>${monthData.dcp}</td>
            <td>${monthData.dcfp}</td>
            <td>${monthData.dmp}</td>
            <td>${monthData.dmfp}</td>
            <td>${monthData.currentTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</td>
            <td>${monthData.referenceTotal > 0 ? monthData.referenceTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2 }) : "-"}</td>
            <td class="${economyClass}">${monthData.economy.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</td>
            <td>${uc.optPonta || 0}</td>
            <td>${uc.optForaPonta || 0}</td>
          </tr>
        `;
      });
      // Linha de Totalizador da UC ou espaço em branco
      html += `
          <tr>
            <td colspan="7" style="text-align: right; font-weight: bold;">Prejuízo Acumulado Total (${uc.uc}):</td>
            <td style="font-weight: bold;" class="negative">${uc.totalEconomy.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</td>
            <td></td>
            <td></td>
          </tr>
          <tr><td colspan="10" style="border: none; height: 20px;"></td></tr>
      `;
    });

    if (lossesUCs.length === 0) {
      html += `<tr><td colspan="10">Nenhuma UC com prejuízo identificada.</td></tr>`;
    }

    html += `
      </table>
      </body>
      </html>
    `;

    const blob = new Blob([html], { type: "application/vnd.ms-excel" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "Monitoramento_Despesas_Resultados.xls";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };


  const exportStableContractsExcel = () => {
    if (!monitoringResults || !monitoringResults.unchangedUCs || monitoringResults.unchangedUCs.length === 0) return;

    let html = `
      <html xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <style>
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid black; padding: 5px; text-align: left; }
          th { background-color: #f2f2f2; font-weight: bold; }
          .header { font-size: 16px; font-weight: bold; background-color: #d9e1f2; }
        </style>
      </head>
      <body>
    `;

    html += `
      <table>
        <tr><td colspan="8" class="header" style="font-size: 18px; text-align: center;">UNIDADES SEM ALTERAÇÃO (CONTRATO ESTÁVEL)</td></tr>
    `;

    [...monitoringResults.unchangedUCs]
      .sort((a: any, b: any) => String(a.gerencia || "").localeCompare(String(b.gerencia || "")))
      .forEach((uc: any) => {
      html += `
        <tr><td colspan="8" style="background-color: #e2efda; font-weight: bold; font-size: 16px; border-bottom: 2px solid #548235; border-top: 2px solid #548235;">UC: ${uc.uc} | Cidade: ${uc.city} - Gerência: ${uc.gerencia}</td></tr>
        <tr>
          <th>Mês/Ano</th>
          <th>Contratada P (kW)</th>
          <th>Contratada FP (kW)</th>
          <th>Medida P (kW)</th>
          <th>Medida FP (kW)</th>
          <th>Gasto Real (R$)</th>
          <th>Demanda Ideal P (kW)</th>
          <th>Demanda Ideal FP (kW)</th>
        </tr>
      `;

      uc.monthlyData.forEach((monthData: any) => {
        html += `
          <tr>
            <td>${monthData.mes}/${monthData.ano}</td>
            <td>${monthData.dcp}</td>
            <td>${monthData.dcfp}</td>
            <td>${monthData.dmp}</td>
            <td>${monthData.dmfp}</td>
            <td>${monthData.currentTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</td>
            <td>${uc.optPonta || 0}</td>
            <td>${uc.optForaPonta || 0}</td>
          </tr>
        `;
      });
      html += `
          <tr><td colspan="8" style="border: none; height: 20px;"></td></tr>
      `;
    });

    html += `
      </table>
      </body>
      </html>
    `;

    const blob = new Blob([html], { type: "application/vnd.ms-excel" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "Monitoramento_Contratos_Estaveis.xls";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const runMonitoringAnalysis = () => {
    const completedBills = bills.filter((b) => b.status === "completed");
    if (completedBills.length === 0) return;

    const tp = 15.0;
    const tfp = 10.0;

    const getYear = (year: string) => parseInt(year || "0", 10);

    // Get all unique UCs
    const ucs = Array.from(new Set(completedBills.map((b) => b.uc))).filter(
      (uc) => {
        const ucBills = completedBills.filter((b) => b.uc === uc);

        // Remove UCs that are clearly Grupo B (to avoid parser hallucination of demand on Grupo B)
        const grupoBCount = ucBills.filter((b) =>
          (b.subgrupo || "").toUpperCase().startsWith("B"),
        ).length;
        const grupoACount = ucBills.filter((b) =>
          (b.subgrupo || "").toUpperCase().startsWith("A"),
        ).length;

        if (grupoBCount > 0 && grupoACount === 0) {
          return false;
        }

        return ucBills.some(
          (b) =>
            parseValue(b.demandaPontaKW) > 0 ||
            parseValue(b.demandaForaPontaKW) > 0,
        );
      },
    );

    const allUcData = ucs.map((uc) => {
      let ucBills = completedBills.filter((b) => b.uc === uc);
      const city = getCidade(String(uc), ucBills[0]?.cidade || "N/A");

      // Sort bills chronologically (Oldest to Newest) for change detection
      ucBills.sort((a, b) => {
        const yearA = getYear(a.anoLeitura);
        const yearB = getYear(b.anoLeitura);
        if (yearA !== yearB) return yearA - yearB;
        return (
          getMonthNumber(a.mesReferencia) - getMonthNumber(b.mesReferencia)
        );
      });

      // Calculate optimal demand (Ideal) based on the last 12 months (or fewer if less data available)
      // LEI 1000: Demanda ideal deve garantir vantagem financeira, considerando arredondamento para cima em 0.5kW
      const isVerde = ucBills.some((b) =>
        (b.modalidadeTarifaria || "").toUpperCase().includes("VERDE"),
      );
      const isAzul = ucBills.some((b) =>
        (b.modalidadeTarifaria || "").toUpperCase().includes("AZUL"),
      );

      const last12MonthsBills = ucBills.slice(-12);
      
      const getOptimalForDemands = (billsToAnalyze: any[]) => {
        const measuredPonta = billsToAnalyze.map((b) => parseValue(b.demandaPotenciaMedidaPonta));
        const measuredForaPonta = billsToAnalyze.map((b) => parseValue(b.demandaPotenciaMedidaForaPonta));
        
        const optimize = (measurements: number[], isPonta: boolean) => {
          if (measurements.length === 0) return 0;
          
          let tariff = 0;
          let penaltyTariff = 0;
          
          if (isAzul) {
            if (isPonta) {
              tariff = 85.53;
              penaltyTariff = 171.07;
            } else {
              tariff = 42.90;
              penaltyTariff = 85.810360;
            }
          } else if (isVerde) {
            if (isPonta) {
              tariff = 0;
              penaltyTariff = 0;
            } else {
              tariff = 43.17715;
              penaltyTariff = 43.17715 * 2;
            }
          } else {
            tariff = 10.0;
            penaltyTariff = 20.0;
          }
          
          if (tariff === 0) return 0;
          
          const minM = Math.min(...measurements);
          const maxM = Math.max(...measurements);
          let bestDemand = 0;
          let minCost = Infinity;
          
          let startCand = Math.floor(Math.max(0, minM / 1.05));
          if (!isPonta && startCand < 30) startCand = 30;
          const endCand = Math.ceil(maxM);
          const actualEndCand = (!isPonta && endCand < 30) ? 30 : endCand;

          for (let cand = startCand; cand <= actualEndCand; cand += 0.5) {
              let totalCost = 0;
              for (const m of measurements) {
                  if (m > cand * 1.05) {
                      totalCost += (cand * tariff) + ((m - cand) * penaltyTariff);
                  } else {
                      totalCost += Math.max(m, cand) * tariff;
                  }
              }
              if (totalCost < minCost) {
                  minCost = totalCost;
                  bestDemand = cand;
              }
          }
          return bestDemand;
        };
        
        const optPonta = isVerde ? 0 : optimize(measuredPonta, true);
        const optForaPonta = optimize(measuredForaPonta, false);
        return { optPonta, optForaPonta };
      };

      const { optPonta, optForaPonta } = getOptimalForDemands(last12MonthsBills);

      // 1st Pass: Calculate Current Total for ALL bills first
      const processedBills = ucBills.map((b) => {
        const modalidade = (b.modalidadeTarifaria || "").toUpperCase();
        const bIsVerde = modalidade.includes("VERDE");
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

        const currentTotal =
          vDmpP + vDmpFp + vDncP + vDncFp + vUltrapP + vUltrapFp;

        return {
          originalBill: b,
          dcp,
          dmp,
          dcfp,
          dmfp,
          currentTotal,
          mes: b.mesReferencia,
          ano: b.anoLeitura,
        };
      });

      // 2nd Pass: Identify all contract transitions and select the LAST (última) contract alteration
      const orig = customOriginalContratadas[uc] || ORIGINAL_CONTRATADAS[uc] || null;
      const meta = customAdjustmentsMetadata[uc] || {};
      const metaAlteracao = meta.dataAlteracao;

      // Scan all contract transitions in chronological order
      const contractTransitions: {
        idx: number;
        prevStart: number;
        prevEnd: number;
        prevContract: { ponta: number; foraPonta: number };
        newContract: { ponta: number; foraPonta: number };
      }[] = [];

      let runningContract: { ponta: number; foraPonta: number } | null = null;
      let runningContractStart = 0;

      for (let i = 0; i < processedBills.length; i++) {
        const b = processedBills[i];
        if (b.dcp > 0 || b.dcfp > 0) {
          if (runningContract === null) {
            runningContract = { ponta: b.dcp, foraPonta: b.dcfp };
            runningContractStart = i;
          } else if (b.dcp !== runningContract.ponta || b.dcfp !== runningContract.foraPonta) {
            contractTransitions.push({
              idx: i,
              prevStart: runningContractStart,
              prevEnd: i,
              prevContract: runningContract,
              newContract: { ponta: b.dcp, foraPonta: b.dcfp },
            });
            runningContract = { ponta: b.dcp, foraPonta: b.dcfp };
            runningContractStart = i;
          }
        }
      }

      let alterationIdx = -1;
      let previousContractAverageCost = 0;

      // Check if explicit meta dataAlteracao is defined
      if (metaAlteracao && metaAlteracao !== "-" && metaAlteracao.trim() !== "") {
        const metaParts = metaAlteracao.trim().split("/");
        let targetMonth = 0;
        let targetYear = 0;
        if (metaParts.length === 3) {
          targetMonth = parseInt(metaParts[1], 10);
          const yStr = metaParts[2].trim();
          targetYear = parseInt(yStr.length === 2 ? "20" + yStr : yStr, 10);
        } else if (metaParts.length === 2) {
          targetMonth = getMonthNumber(metaParts[0]);
          const yStr = metaParts[1].trim();
          targetYear = parseInt(yStr.length === 2 ? "20" + yStr : yStr, 10);
        }
        if (targetMonth > 0 && targetYear > 0) {
          alterationIdx = processedBills.findIndex((b) => {
            const bM = getMonthNumber(b.mes);
            const bY = parseInt(String(b.ano || "0"), 10);
            return bY > targetYear || (bY === targetYear && bM >= targetMonth);
          });
          if (alterationIdx > 0) {
            const preBills = processedBills.slice(0, alterationIdx);
            previousContractAverageCost = preBills.reduce((acc, item) => acc + item.currentTotal, 0) / preBills.length;
          }
        }
      }

      // If not set by metadata, select the LAST (última) contract alteration
      if (alterationIdx === -1 && contractTransitions.length > 0) {
        const lastTransition = contractTransitions[contractTransitions.length - 1];
        alterationIdx = lastTransition.idx;
        const preBills = processedBills.slice(lastTransition.prevStart, lastTransition.prevEnd);
        if (preBills.length > 0) {
          previousContractAverageCost = preBills.reduce((acc, item) => acc + item.currentTotal, 0) / preBills.length;
        } else if (alterationIdx > 0) {
          const preAll = processedBills.slice(0, alterationIdx);
          previousContractAverageCost = preAll.reduce((acc, item) => acc + item.currentTotal, 0) / preAll.length;
        }
      }

      // If still not found, check change against original baseline contract
      if (alterationIdx === -1 && orig && (orig.p > 0 || orig.fp > 0)) {
        // Find last bill that differs from orig
        for (let i = processedBills.length - 1; i >= 0; i--) {
          const b = processedBills[i];
          if (b.dcp !== orig.p || b.dcfp !== orig.fp) {
            alterationIdx = i;
            break;
          }
        }
        if (alterationIdx > 0) {
          const preBills = processedBills.slice(0, alterationIdx);
          previousContractAverageCost = preBills.reduce((acc, item) => acc + item.currentTotal, 0) / preBills.length;
        }
      }

      // Fallback if all bills uploaded are already from the alteration period onwards
      if (alterationIdx === 0 && previousContractAverageCost === 0) {
        const baseDcp = orig ? orig.p : processedBills[0].dcp;
        const baseDcfp = orig ? orig.fp : processedBills[0].dcfp;
        const avgTariffP = isAzul ? 85.53 : 0;
        const avgTariffFp = isAzul ? 42.90 : 43.17715;
        const simulatedBase = (baseDcp * avgTariffP) + (baseDcfp * avgTariffFp);
        previousContractAverageCost = simulatedBase > 0 ? simulatedBase : processedBills[0].currentTotal;
      }

      // 3rd Pass: Calculate monthly economy and accumulated economy starting strictly from the LAST alteration month onwards
      let accumulatedEconomy = 0;
      const hasContractChange = alterationIdx !== -1;

      const monthlyData = processedBills.map((b, index) => {
        const isFromAlterationOnwards = hasContractChange && index >= alterationIdx;
        const hasChanged = hasContractChange && index === alterationIdx;

        let economyFromChange = 0;
        let referenceTotal = 0;

        if (isFromAlterationOnwards && previousContractAverageCost > 0) {
          referenceTotal = previousContractAverageCost;
          // Economia no mês: (Ref. Anterior - Gasto Real)
          economyFromChange = referenceTotal - b.currentTotal;
          // Economia Acumulada a partir da última alteração do contrato em diante
          accumulatedEconomy += economyFromChange;
        }

        return {
          mes: b.mes,
          ano: b.ano,
          currentTotal: b.currentTotal,
          referenceTotal,
          economy: economyFromChange,
          accumulatedEconomy: isFromAlterationOnwards ? accumulatedEconomy : 0,
          dmp: b.dmp,
          dmfp: b.dmfp,
          dcp: b.dcp,
          dcfp: b.dcfp,
          hasChanged,
          referenceContract: null,
        };
      });

      // Reverse to show newest first in UI
      monthlyData.reverse();

      // Total economy = accumulated economy from alteration month onwards
      const totalEconomy = accumulatedEconomy;
      const totalCurrent = monthlyData.reduce(
        (acc, curr) => acc + curr.currentTotal,
        0,
      );

      return {
        uc,
        city,
        gerencia: getGerencia(String(uc)),
        locin: getLocin(String(uc)),
        totalEconomy,
        totalCurrent,
        monthlyData,
        optPonta,
        optForaPonta,
        hasContractChange,
      };
    });

    const changedUCs = allUcData.filter((uc) => uc.hasContractChange);
    const unchangedUCs = allUcData.filter((uc) => !uc.hasContractChange);

    const adjustmentUCs = Object.keys(customRequestedAdjustments).map((ucId) => {
      const req = customRequestedAdjustments[ucId];
      const ucData = allUcData.find((u) => String(u.uc) === String(ucId));
      return {
        uc: ucId,
        reqP: req.p,
        reqFP: req.fp,
        ucData: ucData || null,
      };
    });

    const generalTotalEconomy = allUcData.reduce(
      (acc, curr) => acc + curr.totalEconomy,
      0,
    );
    const generalTotalCurrent = allUcData.reduce(
      (acc, curr) => acc + curr.totalCurrent,
      0,
    );

    // Group by city for the chart
    const cityMap: Record<
      string,
      {
        city: string;
        totalEconomy: number;
        positiveEconomy: number;
        negativeEconomy: number;
        totalCurrent: number;
        optimized: number;
        ucs: { uc: string; economy: number }[];
        positiveUcs: { uc: string; economy: number }[];
        negativeUcs: { uc: string; economy: number }[];
      }
    > = {};
    allUcData.forEach((uc) => {
      if (!cityMap[uc.city]) {
        cityMap[uc.city] = {
          city: uc.city,
          totalEconomy: 0,
          positiveEconomy: 0,
          negativeEconomy: 0,
          totalCurrent: 0,
          optimized: 0,
          ucs: [],
          positiveUcs: [],
          negativeUcs: [],
        };
      }
      cityMap[uc.city].totalEconomy += uc.totalEconomy;
      if (uc.totalEconomy > 0) {
        cityMap[uc.city].positiveEconomy += uc.totalEconomy;
        cityMap[uc.city].positiveUcs.push({
          uc: String(uc.uc),
          economy: uc.totalEconomy,
        });
      } else if (uc.totalEconomy < 0) {
        cityMap[uc.city].negativeEconomy += uc.totalEconomy;
        cityMap[uc.city].negativeUcs.push({
          uc: String(uc.uc),
          economy: uc.totalEconomy,
        });
      }
      cityMap[uc.city].totalCurrent += uc.totalCurrent;
      cityMap[uc.city].ucs.push({
        uc: String(uc.uc),
        economy: uc.totalEconomy,
      });
    });
    const cityData = Object.values(cityMap)
      .map((c) => ({
        ...c,
        optimized: Math.max(0, c.totalCurrent - c.totalEconomy),
        ucs: c.ucs.sort((a, b) => b.economy - a.economy),
        positiveUcs: c.positiveUcs.sort((a, b) => b.economy - a.economy),
        negativeUcs: c.negativeUcs.sort((a, b) => a.economy - b.economy),
      }))
      .sort((a, b) => b.totalCurrent - a.totalCurrent);

    const timelineDataMap: Record<string, { monthYear: string; sortKey: string; currentTotal: number; referenceTotal: number; economy: number }> = {};
    changedUCs.forEach((uc) => {
      uc.monthlyData.forEach((md) => {
        if (md.referenceTotal > 0) {
          const monthName = formatMonth(md.mes);
          const monthYear = `${monthName}/${md.ano}`;
          const sortKey = `${md.ano}${getMonthNumber(md.mes.toString()).toString().padStart(2, "0")}`;
          if (!timelineDataMap[sortKey]) {
            timelineDataMap[sortKey] = {
              monthYear,
              sortKey,
              currentTotal: 0,
              referenceTotal: 0,
              economy: 0,
            };
          }
          timelineDataMap[sortKey].currentTotal += md.currentTotal;
          timelineDataMap[sortKey].referenceTotal += md.referenceTotal;
          timelineDataMap[sortKey].economy += md.economy;
        }
      });
    });
    const timelineData = Object.values(timelineDataMap).sort((a, b) => a.sortKey.localeCompare(b.sortKey));

    setMonitoringResults({
      changedUCs,
      unchangedUCs,
      adjustmentUCs,
      generalTotalEconomy,
      generalTotalCurrent,
      cityData,
      timelineData,
    });
  };

  
  const monitoramentoMeses = useMemo(() => {
    if (!monitoringResults) return [];
    const meses = new Set<string>();
    monitoringResults.changedUCs.forEach((uc: any) => {
      uc.monthlyData.forEach((m: any) => meses.add(formatMonth(m.mes) + '/' + m.ano));
    });
    monitoringResults.unchangedUCs.forEach((uc: any) => {
      uc.monthlyData.forEach((m: any) => meses.add(formatMonth(m.mes) + '/' + m.ano));
    });
    return Array.from(meses).sort((a, b) => {
      const partsA = a.split('/');
      const partsB = b.split('/');
      if (partsA[1] !== partsB[1]) return parseInt(partsB[1]) - parseInt(partsA[1]);
      const m = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
      return m.indexOf(partsB[0]) - m.indexOf(partsA[0]);
    });
  }, [monitoringResults]);

  const filteredMonitoringResults = useMemo(() => {
    if (!monitoringResults) return null;
    if (selectedMonitoramentoMes === "Todos") return monitoringResults;

    const processUcData = (ucDataList: any[]) => {
      return ucDataList.map((uc) => {
        const filteredMonthly = uc.monthlyData.filter((m: any) => (formatMonth(m.mes) + '/' + m.ano) === selectedMonitoramentoMes);
        const totalEconomy = filteredMonthly.reduce((acc: number, curr: any) => acc + curr.economy, 0);
        const totalCurrent = filteredMonthly.reduce((acc: number, curr: any) => acc + curr.currentTotal, 0);
        return {
          ...uc,
          totalEconomy,
          totalCurrent,
          monthlyData: filteredMonthly,
        };
      }).filter((uc) => uc.monthlyData.length > 0);
    };

    const changedUCs = processUcData(monitoringResults.changedUCs);
    const unchangedUCs = processUcData(monitoringResults.unchangedUCs);
    const allUcData = [...changedUCs, ...unchangedUCs];

    const generalTotalEconomy = allUcData.reduce((acc, curr) => acc + curr.totalEconomy, 0);
    const generalTotalCurrent = allUcData.reduce((acc, curr) => acc + curr.totalCurrent, 0);

    const cityMap: Record<string, any> = {};
    allUcData.forEach((uc: any) => {
      if (!cityMap[uc.city]) {
        cityMap[uc.city] = {
          city: uc.city,
          totalEconomy: 0,
          positiveEconomy: 0,
          negativeEconomy: 0,
          totalCurrent: 0,
          optimized: 0,
          ucs: [],
          positiveUcs: [],
          negativeUcs: [],
        };
      }
      const c = cityMap[uc.city];
      c.totalEconomy += uc.totalEconomy;
      c.totalCurrent += uc.totalCurrent;
      c.ucs.push(uc);
      if (uc.totalEconomy > 0) {
        c.positiveEconomy += uc.totalEconomy;
        c.positiveUcs.push(uc);
      } else if (uc.totalEconomy < 0) {
        c.negativeEconomy += Math.abs(uc.totalEconomy);
        c.negativeUcs.push(uc);
      }
    });

    const cityData = Object.values(cityMap)
      .map((c: any) => ({
        ...c,
        optimized: Math.max(0, c.totalCurrent - c.totalEconomy),
        ucs: c.ucs.sort((a: any, b: any) => b.totalEconomy - a.totalEconomy),
        positiveUcs: c.positiveUcs.sort((a: any, b: any) => b.totalEconomy - a.totalEconomy),
        negativeUcs: c.negativeUcs.sort((a: any, b: any) => a.totalEconomy - b.totalEconomy),
      }))
      .sort((a, b) => b.totalCurrent - a.totalCurrent);

    return {
      changedUCs,
      unchangedUCs,
      adjustmentUCs: monitoringResults.adjustmentUCs,
      generalTotalEconomy,
      generalTotalCurrent,
      cityData,
    };
  }, [monitoringResults, selectedMonitoramentoMes]);

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

  const processFile = async (
    bill: BillData & { file: File },
    retryCount = 0,
    currentUser: any = null,
  ) => {
    try {
      if (!bill.file || !(bill.file instanceof Blob)) {
        throw new Error(
          "Arquivo não encontrado na memória. Por favor, remova esta fatura e faça o upload novamente."
        );
      }

      let user = currentUser;

      setBills((prev) =>
        prev.map((b) => (b.id === bill.id ? { ...b, progress: 30 } : b)),
      );

      const text = await extractTextFromPdf(bill.file);

      if (bill.abortController?.signal.aborted)
        throw new Error("Upload cancelado");

      setBills((prev) =>
        prev.map((b) => (b.id === bill.id ? { ...b, progress: 80 } : b)),
      );

      let result: any = parseBillText(text);

      // Resolve UC: For Energisa and Elektro, the filename IS the authoritative UC
      // (CÓDIGO DO CLIENTE), e.g. 7315205160.PDF -> UC 7315205160.
      // The parser often picks up internal codes (e.g. 33100900) instead of the real customer code.
      const concessionaria = (result.concessionaria || "").toUpperCase();
      const solvedUc = extractUcFromFileName(bill.fileName, result.uc);
      if (solvedUc) {
        if (
          concessionaria.includes("ELEKTRO") ||
          concessionaria.includes("ENERGISA") ||
          !result.uc ||
          result.uc.length < 8 ||
          (solvedUc.length >= 8 && result.uc.length < 8)
        ) {
          result.uc = solvedUc;
        }
      }

      if (result.uc) {
        let cleanUc = result.uc.replace(/UC:?\s*/i, "").trim();

        // Rule for formats like "10/2941716-9" -> extract "113865805154"
        if (cleanUc.includes("/") && cleanUc.includes("-")) {
          const afterSlash = cleanUc.split("/")[1];
          if (afterSlash) {
            cleanUc = afterSlash.split("-")[0].trim();
          }
        } else if (cleanUc.includes("-")) {
          cleanUc = cleanUc.split("-")[0].trim();
        }

        // Remover zeros à esquerda da UC (ex: 0000238385 -> 238385)
        cleanUc = cleanUc.replace(/^0+(?=\d)/, "");

        // If after cleaning it ended up too short while filename has valid UC, restore
        const fnNumbers = (bill.fileName || "").replace(/\.[^/.]+$/, "").replace(/\D/g, "");
        if (fnNumbers.length >= 8 && cleanUc.length < 8) {
          cleanUc = fnNumbers;
        }

        result.uc = cleanUc;
      }

      if (result.mesReferencia) {
        result.mesReferencia = formatMonth(result.mesReferencia);
      }

      // --- CHECK FOR DUPLICATES IN DB AFTER EXTRACTION ---
      // existingDbId logic removed
      if (
        isSupabaseConfigured &&
        user &&
        result.uc &&
        result.mesReferencia &&
        result.anoLeitura
      ) {
      // No longer querying for existingDbId. We will unconditionally delete any matching records in Supabase.
      
      // Check duplicates in current list to delete them from SQLite
      let finalStatus: "completed" | "error" = "completed";
      let finalError: string | undefined = undefined;

      const duplicateInClosure = bills.find((b) => {
        if (b.id === bill.id || b.status !== "completed") return false;

        const normalize = (str: string) =>
          (str || "").toString().trim().toLowerCase();

        const hasKeys =
          result.uc && result.mesReferencia && result.anoLeitura;
        if (!hasKeys) return false;

        return (
          normalize(b.uc) === normalize(result.uc) &&
          normalize(b.mesReferencia) === normalize(result.mesReferencia) &&
          normalize(b.anoLeitura) === normalize(result.anoLeitura)
        );
      });

      let isDuplicateInCurrentList = !!duplicateInClosure;

      if (duplicateInClosure) {
        if (isSupabaseConfigured) {
          try {
            await supabase.from("bills").delete().eq("id", duplicateInClosure.id);
          } catch (e) {}
        }
      }

      if (result.uc && UCS_OPER.has(String(result.uc))) {
        result.tipo = "OPER";
      } else if (result.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(result.uc))) {
        let mod = result.modalidadeTarifaria || "";
        if (!mod.toUpperCase().includes("LIVRE")) {
          result.modalidadeTarifaria = mod ? `${mod} - LIVRE` : "LIVRE";
        }
        result.tipo = "LIVRE";
      }
      result.mercado =
        result.uc && UCS_LIVRE_MERCADO_LIVRE.has(String(result.uc))
          ? "LIVRE"
          : "CATIVO";

      // Enriquecer com Gerência e LOCIN do mapeamento
      const mapping = ucMappings.find(
        (m) => m.uc === String(result.uc || bill.uc),
      );
      if (mapping) {
        result.gerencia = mapping.gerencia;
        result.locin = mapping.locin;
        if (!result.cidade) result.cidade = mapping.cidade;
      }

      // Fix UC 117384 - FÁTIMA DO SUL
      if (result.uc === "93604305181") {
        result.cidade = "FÁTIMA DO SUL";
      }

      const { file, ...billWithoutFile } = bill as any;
      const updatedBill: BillData = {
        ...billWithoutFile,
        ...result,
        status: finalStatus,
        error: finalError,
      };

      if (isSupabaseConfigured && user) {
        const dbData = mapBillDataToDb(updatedBill, user.id);

        if (dbData.uc && dbData.mes_referencia && dbData.ano_leitura) {
          // Unconditionally delete all older duplicates to guarantee only this latest one remains
          await supabase.from("bills").delete()
            .eq("uc", dbData.uc)
            .eq("mes_referencia", dbData.mes_referencia)
            .eq("ano_leitura", dbData.ano_leitura);
        }

        // Insert new record
          let { error: insertError } = await supabase
            .from("bills")
            .insert(dbData);

          if (
            insertError &&
            (insertError.message.includes("data_vencimento") ||
              insertError.message.includes("mercado") ||
              insertError.message.includes("gerencia") ||
              insertError.message.includes("locins") ||
              insertError.message.includes("consumo_kwh_grupo_b") ||
              insertError.message.includes("valor_consumo_kwh_grupo_b") ||
              insertError.details?.includes("data_vencimento") ||
              insertError.details?.includes("mercado") ||
              insertError.code === "PGRST204")
          ) {
            console.warn(
              "Colunas novas não encontradas no Supabase. Inserindo sem elas...",
            );
            const {
              data_vencimento,
              mercado,
              gerencia,
              locins,
              consumo_kwh_grupo_b,
              valor_consumo_kwh_grupo_b,
              ...fallbackData
            } = dbData;
            if (consumo_kwh_grupo_b) fallbackData.consumo_kwh_ponta = consumo_kwh_grupo_b;
            if (valor_consumo_kwh_grupo_b) fallbackData.valor_consumo_kwh_ponta = valor_consumo_kwh_grupo_b;
            const fallbackRes = await supabase
              .from("bills")
              .insert(fallbackData);
            insertError = fallbackRes.error;
          }

          if (insertError) {
            console.error("Erro ao salvar fatura no Supabase:", insertError);
            updatedBill.status = "error";
            updatedBill.error = insertError.message || "Erro ao salvar no banco de dados.";
          }
        }
      }

      setBills((prev) =>
        deduplicateBills(
          prev.map((b) =>
            b.id === bill.id ? { ...updatedBill, progress: 100 } : b,
          ),
        ),
      );
    } catch (error: any) {
      if (error.message === "Upload cancelado") {
        setBills((prev) =>
          prev.map((b) =>
            b.id === bill.id
              ? {
                  ...b,
                  status: "error",
                  error: "Cancelado",
                  progress: 0,
                }
              : b,
          ),
        );
        return;
      }

      console.error("Erro na extração:", error);

      let isRateLimit = false;
      let isQuotaExhausted = false;
      let isLockError = false;
      let isTransientError = false;
      let retryAfter = 0;

      // Check for rate limit in various error formats
      const errorStr =
        error?.message ||
        (typeof error === "string" ? error : JSON.stringify(error));
      const nestedError = error?.error || error;
      const errorCode = nestedError?.code || error?.status || 0;
      const errorStatus = nestedError?.status || "";
      const msg = nestedError?.message || error?.message || "";

      if (
        msg.includes("spending cap") ||
        errorStr.includes("spending cap") ||
        errorStr.includes("limite de gastos") ||
        errorStr.includes("monthly limit") ||
        errorStr.includes("prepayment credits are depleted") ||
        errorStr.includes("Failed to fetch")
      ) {
        isQuotaExhausted = true;
      } else if (
        errorCode === 429 ||
        errorStatus === "RESOURCE_EXHAUSTED" ||
        errorStr.includes("429") ||
        errorStr.includes("RESOURCE_EXHAUSTED") ||
        msg.toLowerCase().includes("quota") ||
        msg.toLowerCase().includes("rate limit") ||
        errorStr.includes("billing details")
      ) {
        isRateLimit = true;
        // Try to extract retry time from message
        const match = msg.match(/retry in ([\d.]+)s/);
        if (match && match[1]) {
          retryAfter = parseFloat(match[1]) * 1000;
        } else {
          // Default to a longer delay for quota issues
          retryAfter = 15000;
        }
      } else if (errorStr.includes("Lock broken by another request")) {
        isLockError = true;
      } else if (
        errorCode === 500 ||
        errorCode === 502 ||
        errorCode === 503 ||
        errorCode === 504 ||
        errorStatus === "INTERNAL" ||
        errorStatus === "UNAVAILABLE" ||
        errorStr.includes("500") ||
        errorStr.includes("502") ||
        errorStr.includes("503") ||
        errorStr.includes("504") ||
        errorStr.includes("INTERNAL") ||
        errorStr.includes("UNAVAILABLE")
      ) {
        isTransientError = true;
      }

      if (isQuotaExhausted) {
        setBills((prev) =>
          prev.map((b) =>
            b.id === bill.id
              ? {
                  ...b,
                  status: "error",
                  error:
                    "O limite de gastos do seu projeto foi atingido. O processamento foi interrompido para evitar cobranças excedentes.",
                }
              : b,
          ),
        );

        // Stop the entire processing queue
        setIsProcessing(false);
        isProcessingRef.current = false;
        return;
      }

      if (isRateLimit || isLockError || isTransientError) {
        // Limit max retries to 8 to avoid freezing the app for too long
        if (retryCount < 8) {
          // Use retryAfter if found, otherwise exponential backoff starting at 10s for rate limits
          const delay =
            retryAfter > 0
              ? retryAfter + 3000 // Adiciona 3s de margem
              : Math.pow(2, retryCount) * 5000 + Math.random() * 2000;

          console.log(
            `[Worker] ${isLockError ? "Erro de trava" : isTransientError ? "Erro temporário (" + errorCode + ")" : "Limite de taxa/cota"} atingido para ${bill.fileName}. Tentando novamente em ${Math.round(delay / 1000)}s... (Tentativa ${retryCount + 1}/8)`,
          );

          setBills((prev) =>
            prev.map((b) =>
              b.id === bill.id
                ? {
                    ...b,
                    status: "processing",
                    error: `Aguardando ${isLockError ? "liberação" : isTransientError ? "servidor" : "limite de cota"} da API... Tentativa ${retryCount + 1}/8 (${Math.round(delay / 1000)}s)`,
                  }
                : b,
            ),
          );

          await new Promise((resolve) => setTimeout(resolve, delay));
          return await processFile(bill, retryCount + 1, currentUser);
        } else {
          console.error(
            `[Worker] Falha após ${retryCount} tentativas para ${bill.fileName} devido a ${isLockError ? "erro de trava" : "limite de taxa"}.`,
          );
        }
      }

      setBills((prev) =>
        prev.map((b) =>
          b.id === bill.id
            ? {
                ...b,
                status: "error",
                error: error.message || "Erro na extração",
              }
            : b,
        ),
      );
    }
  };

  const startProcessing = async () => {
    if (isProcessing) return;

    const pendingBills = bills.filter((b) => b.status === "pending");
    if (pendingBills.length === 0) return;

    setIsProcessing(true);
    isProcessingRef.current = true;

    let currentUser = null;
    if (isSupabaseConfigured) {
      try {
        const {
          data: { user },
        } = await supabase.auth.getUser();
        currentUser = user;
      } catch (authError: any) {
        console.warn(
          "Erro ao obter usuário no início do processamento:",
          authError,
        );
        if (authError.message?.includes("Refresh Token Not Found")) {
          supabase.auth.signOut();
          setIsAuthenticated(false);
          setIsProcessing(false);
          isProcessingRef.current = false;
          return;
        }
      }
    }

    // Worker pool approach to maintain concurrency limited to 8 to speed up extraction now that memory leaks are fixed
    const queue = [...pendingBills];
    const maxConcurrency = 8;
    const initialWorkers = Math.min(maxConcurrency, queue.length);

    const runWorker = async (workerId: number) => {
      // Add a small staggered start for workers to avoid simultaneous requests
      await new Promise((resolve) => setTimeout(resolve, workerId * 500));

      while (queue.length > 0 && isProcessingRef.current) {
        // Yield to the React UI thread to update progress bars
        await new Promise((resolve) => setTimeout(resolve, 50));

        const bill = queue.shift();
        if (!bill) break;

        // Update status to processing
        const abortController = new AbortController();
        setBills((prev) =>
          prev.map((b) =>
            b.id === bill.id
              ? {
                  ...b,
                  status: "processing",
                  error: undefined,
                  abortController,
                  progress: 0,
                }
              : b,
          ),
        );

        try {
          console.log(
            `[Worker ${workerId}] Iniciando processamento de: ${bill.fileName} (Restam: ${queue.length})`,
          );
          await processFile(
            { ...bill, abortController } as any,
            0,
            currentUser,
          );
          console.log(
            `[Worker ${workerId}] Concluído processamento de: ${bill.fileName}`,
          );

          // Add a 1s delay to avoid overwhelming the API
          if (queue.length > 0 && isProcessingRef.current) {
            console.log(
              `[Worker ${workerId}] Aguardando 1s para o próximo arquivo...`,
            );
            await new Promise((resolve) => setTimeout(resolve, 1000));
          }
        } catch (error: any) {
          console.error(
            `[Worker ${workerId}] Erro crítico no processamento de ${bill.fileName}:`,
            error,
          );
          setBills((prev) =>
            prev.map((b) =>
              b.id === bill.id
                ? {
                    ...b,
                    status: "error",
                    error: error.message || "Erro crítico",
                  }
                : b,
            ),
          );
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
    setBills((prev) =>
      prev.map((b) =>
        b.status === "processing"
          ? { ...b, status: "pending", progress: 0 }
          : b,
      ),
    );
    setIsProcessing(false);
    isProcessingRef.current = false;
  };

  const [selectedBills, setSelectedBills] = useState<string[]>([]);

  const completedBills = bills.filter((b) => b.status === "completed");

  const adjustmentsList = useMemo(() => {
    return Object.keys(customRequestedAdjustments).map((ucId) => {
      const req = customRequestedAdjustments[ucId];
      const orig = customOriginalContratadas[ucId] || { p: 0, fp: 0 };
      
      const ucData = monitoringResults?.adjustmentUCs?.find((a: any) => String(a.uc) === String(ucId))?.ucData || null;

      let city = "-";
      let gerencia = "-";
      
      const ucBills = completedBills.filter((b) => String(b.uc) === String(ucId));

      if (ucData) {
        city = ucData.city || "-";
        gerencia = ucData.gerencia || "-";
      } else if (ucBills.length > 0) {
        city = ucBills[0]?.cidade || getCidade(String(ucId), "N/A");
        gerencia = ucBills[0]?.gerencia || getGerencia(String(ucId));
      } else {
        const mapping = ucMappings.find(m => String(m.uc) === String(ucId));
        if (mapping) {
          city = mapping.cidade || "-";
          gerencia = mapping.gerencia || "-";
        } else {
          city = getCidade(String(ucId), "-");
          gerencia = getGerencia(String(ucId)) || "-";
        }
      }

      if (city === "-" || city === "N/A") {
        city = getCidade(String(ucId), "-");
      }
      if (gerencia === "-" || gerencia === "N/A") {
        gerencia = getGerencia(String(ucId)) || "-";
      }

      // Default values
      let dcpBefore = orig.p;
      let dcfpBefore = orig.fp;
      let dcpAfter = req.p;
      let dcfpAfter = req.fp;
      let dataAlteracao = "-";
      let economia = 0;

      // Helper to parse DD/MM/YYYY dates
      const parseBrlDate = (dateStr?: string) => {
        if (!dateStr) return null;
        const parts = dateStr.trim().split("/");
        if (parts.length !== 3) return null;
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // 0-indexed month in JS Date
        const yearStr = parts[2].trim();
        const year = parseInt(yearStr.length === 2 ? "20" + yearStr : yearStr, 10);
        if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
        return new Date(year, month, day);
      };

      // Helper to check if a bill's date is from 03/06/2026 (June 3rd, 2026) onwards
      const isFromJune03_2026Onwards = (b: any) => {
        const dt = parseBrlDate(b.dataVencimento);
        if (dt) {
          return dt >= new Date(2026, 5, 3); // Month 5 is June
        }
        // Fallback to reference month
        const m = getMonthNumber(b.mesReferencia);
        const yStr = b.anoLeitura || "";
        const y = parseInt(yStr.length === 2 ? "20" + yStr : yStr, 10);
        if (!isNaN(m) && !isNaN(y)) {
          if (y > 2026) return true;
          if (y === 2026 && m >= 6) return true; // reference June 2026 is considered >= 03/06/2026
        }
        return false;
      };

      // Try reading actual contracts from history of bills loaded in app if available
      if (ucBills.length > 0) {
        const sortedUcBills = [...ucBills].sort((a, b) => {
          const yearA = parseInt(a.anoLeitura || "0", 10);
          const yearB = parseInt(b.anoLeitura || "0", 10);
          if (yearA !== yearB) return yearA - yearB;
          return getMonthNumber(a.mesReferencia) - getMonthNumber(b.mesReferencia);
        });

        // Use the value of the last month inserted in the app for BEFORE/Baseline columns as requested (Contratada P and Contratada FP)
        const lastBill = sortedUcBills[sortedUcBills.length - 1];
        dcpBefore = parseValue(lastBill.demandaPontaKW);
        dcfpBefore = parseValue(lastBill.demandaForaPontaKW);

        // Keep the requested_adjustments values for the AFTER/Altered columns as requested (Contratada P Alt and Contratada FP Alt)
        dcpAfter = req.p;
        dcfpAfter = req.fp;

        // Try reading actual contract data alteration month from the bills
        // It's the first month where the contract demand actually changed from the original values
        const changeBill = sortedUcBills.find(b => 
          parseValue(b.demandaPontaKW) !== orig.p || 
          parseValue(b.demandaForaPontaKW) !== orig.fp
        );

        if (changeBill && isFromJune03_2026Onwards(changeBill)) {
          dataAlteracao = `${changeBill.mesReferencia}/${changeBill.anoLeitura}`;
        }
      }

      if (ucData) {
        economia = ucData.totalEconomy;
        if (dataAlteracao === "-") {
          const changedMonth = ucData.monthlyData.find((m: any) => m.hasChanged);
          if (changedMonth) {
            const mNum = getMonthNumber(changedMonth.mes);
            const yNum = parseInt(changedMonth.ano || "0", 10);
            const belongsToBoundary = (yNum > 2026) || (yNum === 2026 && mNum >= 6);
            if (belongsToBoundary) {
              dataAlteracao = `${changedMonth.mes}/${changedMonth.ano}`;
            }
          }
        }
      }

      // Metadata overrides
      const meta = customAdjustmentsMetadata[ucId] || {};
      if (meta.city !== undefined && meta.city !== "") city = meta.city;
      if (meta.gerencia !== undefined && meta.gerencia !== "") gerencia = meta.gerencia;
      
      const dataSolicitacao = meta.dataSolicitacao !== undefined && meta.dataSolicitacao !== "" ? meta.dataSolicitacao : (ucId === "10926205169" ? "-" : "03/06/2026");
      if (meta.dataAlteracao !== undefined && meta.dataAlteracao !== "") dataAlteracao = meta.dataAlteracao;
      
      const status = meta.status !== undefined && meta.status !== "" ? meta.status : (dataAlteracao !== "-" ? "Concluída" : "Aguardando");

      let finalEconomia = economia;
      if (meta.ecoRealizada !== undefined && meta.ecoRealizada !== "") {
        finalEconomia = parseFloat(meta.ecoRealizada) || 0;
      } else if (dataAlteracao === "-") {
        finalEconomia = 0;
      }
      
      let previsaoEconomia = meta.previsaoEconomia !== undefined && meta.previsaoEconomia !== "" ? meta.previsaoEconomia : "-";
      if (previsaoEconomia === "-" && PDF_PREVISAO_BASE[ucId] !== undefined) {
        previsaoEconomia = (PDF_PREVISAO_BASE[ucId] * 0.7).toFixed(2);
      }

      return {
        uc: ucId,
        reqP: req.p,
        reqFP: req.fp,
        origP: orig.p,
        origFP: orig.fp,
        dcpBefore,
        dcfpBefore,
        dcpAfter,
        dcfpAfter,
        dataAlteracao,
        economia: finalEconomia,
        city,
        gerencia,
        ucData,
        dataSolicitacao,
        previsaoEconomia,
        status
      };
    });
  }, [monitoringResults, ucMappings, completedBills, customRequestedAdjustments, customOriginalContratadas, customAdjustmentsMetadata]);

  const exportAdjustmentsExcel = () => {
    if (adjustmentsList.length === 0) return;

    let html = `
      <html xmlns:x="urn:schemas-microsoft-com:office:excel">
      <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <style>
          table { border-collapse: collapse; width: 100%; }
          th, td { border: 1px solid black; padding: 5px; text-align: center; }
          th { background-color: #f2f2f2; font-weight: bold; }
        </style>
      </head>
      <body>
        <table>
          <tr>
            <th colspan="12" style="font-size: 18px;">Relação de UCs com Alteração de Demanda</th>
          </tr>
          <tr>
            <th>Nº UC</th>
            <th>Município</th>
            <th>Localidade</th>
            <th>Data da solicitação</th>
            <th>Data da alteração</th>
            <th>Status</th>
            <th>Demanda contratada Ponta</th>
            <th>Demanda contratada Ponta Alterada</th>
            <th>Demanda contratada F Ponta</th>
            <th>Demanda contratada F Ponta Alterada</th>
            <th>Previsão de economia</th>
            <th>Economia realizada</th>
          </tr>
    `;

    [...adjustmentsList].forEach((adj: any) => {
        const {
          uc: ucId,
          city,
          gerencia,
          dcpBefore,
          dcfpBefore,
          dcpAfter,
          dcfpAfter,
          dataAlteracao,
          economia,
          dataSolicitacao,
          previsaoEconomia,
          status
        } = adj;

        html += `
          <tr>
            <td style="mso-number-format:'\\@';">${adj.uc}</td>
            <td>${city}</td>
            <td>${gerencia}</td>
            <td>${dataSolicitacao}</td>
            <td>${dataAlteracao}</td>
            <td>${status}</td>
            <td>${String(dcpBefore).replace(".", ",")}</td>
            <td>${String(dcpAfter).replace(".", ",")}</td>
            <td>${String(dcfpBefore).replace(".", ",")}</td>
            <td>${String(dcfpAfter).replace(".", ",")}</td>
            <td>${previsaoEconomia}</td>
            <td>${String(economia.toFixed(2)).replace(".", ",")}</td>
          </tr>
        `;
    });

    html += `
        </table>
      </body>
      </html>
    `;

    const blob = new Blob([html], { type: "application/vnd.ms-excel" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "monitoramento_ajustes.xls";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const toggleReactiveUc = (uc: string) => {
    const newSet = new Set(expandedReactiveUcs);
    if (newSet.has(uc)) newSet.delete(uc);
    else newSet.add(uc);
    setExpandedReactiveUcs(newSet);
  };

  const handleReactiveSort = (field: string) => {
    if (reactiveSortField === field) {
      setReactiveSortDirection(
        reactiveSortDirection === "asc" ? "desc" : "asc",
      );
    } else {
      setReactiveSortField(field);
      setReactiveSortDirection("desc");
    }
  };

  const toggleSummaryCity = (city: string) => {
    setExpandedSummaryCities((prev) => {
      const next = new Set(prev);
      if (next.has(city)) next.delete(city);
      else next.add(city);
      return next;
    });
  };

  const toggleUCExpansion = (uc: string) => {
    setExpandedUCs((prev) => {
      const next = new Set(prev);
      if (next.has(uc)) next.delete(uc);
      else next.add(uc);
      return next;
    });
  };

  const toggleAnalysisUCExpansion = (uc: string) => {
    setExpandedAnalysisUCs((prev) => {
      const next = new Set(prev);
      if (next.has(uc)) next.delete(uc);
      else next.add(uc);
      return next;
    });
  };

  const toggleBillSelection = (id: string) => {
    setSelectedBills((prev) =>
      prev.includes(id) ? prev.filter((s) => s !== id) : [...prev, id],
    );
  };


  // Helper to immediately and synchronously purge deleted bills from localStorage
  const purgeBillsFromLocalStorage = (deletedIds: string[]) => {
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    try {
      const saved = localStorage.getItem("sanesul_bills");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          const filtered = parsed.filter((b: any) => !deletedIds.includes(b.id));
          localStorage.setItem("sanesul_bills", JSON.stringify(filtered));
        }
      }
    } catch (e) {}
  };

  const removeBill = async (id: string) => {
    if (isSupabaseConfigured) {
      try {
        const { error } = await supabase.from("bills").delete().eq("id", id);
        if (error) console.error("Erro ao deletar fatura do Supabase:", error);
      } catch (err) {
        console.error("Erro inesperado ao deletar fatura:", err);
      }
    }

    // Remove individual file from localforage to free space immediately
    localforage
      .getItem<Record<string, File>>("sanesul_bills_files")
      .then((filesMap) => {
        if (filesMap && filesMap[id]) {
          delete filesMap[id];
          localforage.setItem("sanesul_bills_files", filesMap);
        }
      })
      .catch((e) => console.warn("Erro ao deletar arquivo local", e));

    purgeBillsFromLocalStorage([id]);
    setBills((prev) => prev.filter((b) => b.id !== id));
    setSelectedBills((prev) => prev.filter((s) => s !== id));
  };

  const removeSelectedBills = async () => {
    if (isSupabaseConfigured && selectedBills.length > 0) {
      try {
        let error = null;
        for (let i = 0; i < selectedBills.length; i += 100) {
          const chunk = selectedBills.slice(i, i + 100);
          const { error: err } = await supabase.from("bills").delete().in("id", chunk);
          if (err) error = err;
        }
        if (error) console.error("Erro ao deletar faturas do Supabase:", error);
      } catch (err) {
        console.error("Erro inesperado ao deletar faturas:", err);
      }
    }

    // Remove files from localforage
    if (selectedBills.length > 0) {
      localforage
        .getItem<Record<string, File>>("sanesul_bills_files")
        .then((filesMap) => {
          if (filesMap) {
            let Changed = false;
            selectedBills.forEach((id) => {
              if (filesMap[id]) {
                delete filesMap[id];
                Changed = true;
              }
            });
            if (Changed) {
              localforage.setItem("sanesul_bills_files", filesMap);
            }
          }
        })
        .catch((e) => console.warn("Erro ao deletar arquivos locais", e));
    }

    purgeBillsFromLocalStorage(selectedBills);
    setBills((prev) => prev.filter((b) => !selectedBills.includes(b.id)));
    setSelectedBills([]);
  };

  const removeUCBills = async (uc: string) => {
    const ucBills = bills.filter((b) => b.uc === uc);
    const ucBillIds = ucBills.map((b) => b.id);
    
    if (ucBillIds.length === 0) return;

    if (isSupabaseConfigured) {
      try {
        let error = null;
        for (let i = 0; i < ucBillIds.length; i += 100) {
          const chunk = ucBillIds.slice(i, i + 100);
          const { error: err } = await supabase.from("bills").delete().in("id", chunk);
          if (err) error = err;
        }
        if (error) console.error("Erro ao deletar faturas do Supabase:", error);
      } catch (err) {
        console.error("Erro inesperado ao deletar faturas:", err);
      }
    }

    localforage
      .getItem<Record<string, File>>("sanesul_bills_files")
      .then((filesMap) => {
        if (filesMap) {
          let hasChanges = false;
          ucBillIds.forEach(id => {
            if (filesMap[id]) {
              delete filesMap[id];
              hasChanges = true;
            }
          });
          if (hasChanges) localforage.setItem("sanesul_bills_files", filesMap);
        }
      });
      
    purgeBillsFromLocalStorage(ucBillIds);
    setBills((prev) => prev.filter((b) => b.uc !== uc));
    setSelectedBills((prev) => prev.filter((id) => !ucBillIds.includes(id)));
    if (analysisResults) {
      setAnalysisResults((prev: any) => prev?.filter((r: any) => r.uc !== uc) || null);
    }
  };

  const removeMonthBills = async (monthRef: string) => {
    const monthBills = bills.filter(
      (b) => `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` === monthRef,
    );
    const monthBillIds = monthBills.map((b) => b.id);

    if (monthBillIds.length === 0) {
      alert(`Nenhuma fatura encontrada para o mês ${monthRef}.`);
      return;
    }

    const confirmed = window.confirm(
      `Tem certeza que deseja excluir TODAS as ${monthBillIds.length} faturas do mês ${monthRef}? Esta ação não pode ser desfeita.`,
    );
    if (!confirmed) return;

    if (isSupabaseConfigured) {
      try {
        let error = null;
        for (let i = 0; i < monthBillIds.length; i += 100) {
          const chunk = monthBillIds.slice(i, i + 100);
          const { error: err } = await supabase.from("bills").delete().in("id", chunk);
          if (err) error = err;
        }
        if (error) console.error("Erro ao deletar faturas do Supabase:", error);
      } catch (err) {
        console.error("Erro inesperado ao deletar faturas:", err);
      }
    }

    // Remove files from localforage
    localforage
      .getItem<Record<string, File>>("sanesul_bills_files")
      .then((filesMap) => {
        if (filesMap) {
          let hasChanges = false;
          monthBillIds.forEach((id) => {
            if (filesMap[id]) {
              delete filesMap[id];
              hasChanges = true;
            }
          });
          if (hasChanges) localforage.setItem("sanesul_bills_files", filesMap);
        }
      })
      .catch((e) => console.warn("Erro ao deletar arquivos locais", e));

    purgeBillsFromLocalStorage(monthBillIds);
    setBills((prev) => prev.filter((b) => !monthBillIds.includes(b.id)));
    setSelectedBills((prev) => prev.filter((id) => !monthBillIds.includes(id)));
    if (analysisResults) {
      setAnalysisResults((prev: any) =>
        prev ? prev.filter((r: any) => !monthBillIds.includes(r.id)) : null,
      );
    }
    // Reset filter if the deleted month was the active filter
    if (filterReference === monthRef) {
      setFilterReference("all");
    }
  };

  const deselectFirst223 = () => {
    const first223Ids = bills.slice(0, 223).map((b) => b.id);
    setSelectedBills((prev) => prev.filter((id) => !first223Ids.includes(id)));
  };

  const exportAnalysisToCSV = () => {
    if (!analysisResults || analysisResults.length === 0) return;

    const headers = [
      "Nome do Arquivo",
      "UC",
      "Gerência",
      "LOCINS",
      "Cidade",
      "Tipo",
      "Mercado",
      "Ano",
      "Mês",
      "Demanda Medida Ponta",
      "Demanda Medida Fora Ponta",
      "Demanda Ideal Ponta",
      "Demanda Ideal Fora Ponta",
      "Gasto Real (R$)",
      "Economia (R$)",
      "Status",
      "Grupo Tarifário",
      "Tarifa Branca",
      "Optante B",
    ];

    const rows = analysisResults.map((r: any) => {
      // Basic classification logic (placeholder, needs refinement based on actual data)
      const isGrupoA = r.dcp > 0 || r.dcfp > 0; // Simplified assumption
      const isGrupoB = !isGrupoA;
      const isSolar = hasCompensacao(r); // Assuming these fields exist in analysisResults

      const grupo = isGrupoA
        ? "Grupo A (Verde/Azul)"
        : isSolar
          ? "Grupo B (Solar)"
          : "Grupo B (Não Solar)";
      const tarifaBranca = "N/A"; // Need to determine how to identify this
      const optanteB = "N/A"; // Need to determine how to identify this

      return [
        r.fileName,
        r.uc,
        getGerencia(String(r.uc)),
        getLocin(String(r.uc)),
        getCidade(String(r.uc), r.city),
        r.tipo,
        r.mercado,
        r.ano,
        r.mes,
        r.dmp,
        r.dmfp,
        r.optimizedPonta,
        r.optimizedForaPonta,
        String((r.currentTotal || 0).toFixed(2)).replace(".", ","),
        String((r.economy || 0).toFixed(2)).replace(".", ","),
        r.isOverrun ? "Ultrapassagem" : r.isSub ? "Subutilização" : "OK",
        grupo,
        tarifaBranca,
        optanteB,
      ];
    });

    // Use semicolon as delimiter for better compatibility with Excel in many locales (like Brazil)
    // Add UTF-8 BOM (\uFEFF) to ensure Excel recognizes the encoding
    const csvContent =
      "\uFEFF" +
      [
        headers.join(";"),
        ...rows.map((row) =>
          row
            .map((val) => {
              const safeVal = String(val || "").replace(/;/g, ",");
              return `"${safeVal}"`;
            })
            .join(";"),
        ),
      ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `analise_demanda_${new Date().toISOString().split("T")[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportReactiveToCSV = () => {
    const reactiveBills = bills.filter((b) => {
      if (b.status !== "completed") return false;
      if (
        selectedReactiveMonth !== "all" &&
        `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` !== selectedReactiveMonth
      )
        return false;
      const totalReativo =
        parseValue(b.valorEnergiaReativaExcedPonta) +
        parseValue(b.valorEnergiaReativaExcedFPonta);
      return totalReativo > 100;
    });

    const grouped = reactiveBills.reduce(
      (acc, bill) => {
        const uc = String(bill.uc);
        if (!acc[uc]) {
          acc[uc] = {
            uc,
            cidade: getCidade(uc, bill.cidade),
            totalPonta: 0,
            totalFPonta: 0,
            totalFatura: 0,
            bills: [],
          };
        }
        acc[uc].totalPonta += parseValue(bill.valorEnergiaReativaExcedPonta);
        acc[uc].totalFPonta += parseValue(bill.valorEnergiaReativaExcedFPonta);
        acc[uc].totalFatura += parseValue(bill.valorTotal);
        acc[uc].bills.push(bill);
        return acc;
      },
      {} as Record<
        string,
        {
          uc: string;
          cidade: string;
          totalPonta: number;
          totalFPonta: number;
          totalFatura: number;
          bills: typeof bills;
        }
      >,
    );

    const reactiveData = Object.values(grouped);

    if (reactiveData.length === 0) {
      showAlert(
        "Exportação",
        "Não há dados de monitoramento reativo para exportar.",
      );
      return;
    }

    const headers = [
      "UC",
      "Gerência",
      "Cidade",
      "Total Reativa Ponta (R$)",
      "Total Reativa F. Ponta (R$)",
      "Total Geral Reativo (R$)",
      "Total Faturas (R$)",
      "% da Fatura",
    ];

    const rows = (reactiveData as any[]).map((data) => [
      data.uc,
      getGerencia(data.uc),
      data.cidade,
      String(data.totalPonta.toFixed(2)).replace(".", ","),
      String(data.totalFPonta.toFixed(2)).replace(".", ","),
      String((data.totalPonta + data.totalFPonta).toFixed(2)).replace(".", ","),
      String(data.totalFatura.toFixed(2)).replace(".", ","),
      String(
        data.totalFatura > 0
          ? (
              ((data.totalPonta + data.totalFPonta) / data.totalFatura) *
              100
            ).toFixed(2)
          : "0",
      ).replace(".", ",") + "%",
    ]);

    const csvContent =
      "\uFEFF" +
      [
        headers.join(";"),
        ...rows.map((row) =>
          row
            .map((val) => {
              const safeVal = String(val || "").replace(/;/g, ",");
              return `"${safeVal}"`;
            })
            .join(";"),
        ),
      ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `monitoramento_reativo_${selectedReactiveMonth.replace("/", "_")}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportReactiveGroupedToCSV = () => {
    const reactiveBills = bills.filter((b) => {
      if (b.status !== "completed") return false;
      if (
        selectedReactiveMonth !== "all" &&
        `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` !== selectedReactiveMonth
      )
        return false;
      const totalReativo =
        parseValue(b.valorEnergiaReativaExcedPonta) +
        parseValue(b.valorEnergiaReativaExcedFPonta);
      return totalReativo > 100;
    });

    const ucConsolidated: Record<string, any> = {};

    reactiveBills.forEach((bill) => {
      const uc = String(bill.uc);
      if (!ucConsolidated[uc]) {
        ucConsolidated[uc] = {
          uc,
          gerencia: getGerencia(uc),
          cidade: getCidade(uc, bill.cidade),
          locin: getLocin(uc),
          totalReativo: 0,
        };
      }
      ucConsolidated[uc].totalReativo +=
        parseValue(bill.valorEnergiaReativaExcedPonta) +
        parseValue(bill.valorEnergiaReativaExcedFPonta);
    });

    const gerenciaGroups: Record<string, any[]> = {};
    Object.values(ucConsolidated).forEach((item) => {
      const g = item.gerencia;
      if (!gerenciaGroups[g]) gerenciaGroups[g] = [];
      gerenciaGroups[g].push(item);
    });

    if (Object.keys(gerenciaGroups).length === 0) {
      showAlert(
        "Exportação",
        "Não há dados de monitoramento reativo para exportar.",
      );
      return;
    }

    const headers = [
      "UC",
      "GERÊNCIA",
      "CIDADE",
      "LOCINS",
      "VALOR REATIVO (R$)",
    ];
    const rows: string[][] = [];
    let grandTotal = 0;

    const sortedGerencias = Object.keys(gerenciaGroups).sort();

    sortedGerencias.forEach((g) => {
      let groupTotal = 0;
      gerenciaGroups[g]
        .sort((a, b) => b.totalReativo - a.totalReativo)
        .forEach((item) => {
        rows.push([
          item.uc,
          item.gerencia,
          item.cidade,
          item.locin,
          String(item.totalReativo.toFixed(2)).replace(".", ","),
        ]);
        groupTotal += item.totalReativo;
      });
      rows.push([
        "",
        "SUBTOTAL " + g,
        "",
        "",
        String(groupTotal.toFixed(2)).replace(".", ","),
      ]);
      rows.push(["", "", "", "", ""]);
      grandTotal += groupTotal;
    });

    rows.push([
      "",
      "SOMA TOTAL GERAL",
      "",
      "",
      String(grandTotal.toFixed(2)).replace(".", ","),
    ]);

    const csvContent =
      "\uFEFF" +
      [
        headers.join(";"),
        ...rows.map((row) =>
          row
            .map((val) => {
              const safeVal = String(val || "").replace(/;/g, ",");
              return `"${safeVal}"`;
            })
            .join(";"),
        ),
      ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `relatorio_reativo_por_gerencia_${selectedReactiveMonth.replace("/", "_")}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportToCSV = () => {
    // Export all completed bills without filtering by the selected reference
    const completedBills = bills.filter((b) => b.status === "completed");
    if (completedBills.length === 0) {
      showAlert(
        "Exportação",
        "Não há faturas concluídas para exportar.",
      );
      return;
    }

    const headers = [
      "Nome do Arquivo",
      "UC",
      "Gerência",
      "LOCINS",
      "Cidade",
      "Tipo",
      "Mercado",
      "Concessionária",
      "Mês Referência",
      "Ano Leitura",
      "Vencimento",
      "Nota Fiscal",
      "Modalidade Tarifária",
      "Subgrupo",
      "Valor Total (R$)",
      "Consumo Grupo B (kWh)",
      "Demanda Contratada Ponta (kW)",
      "Demanda Contratada Fora Ponta (kW)",
      "Demanda Medida Ponta (kW)",
      "Valor Demanda Medida Ponta (R$)",
      "Demanda Medida Fora Ponta (kW)",
      "Valor Demanda Medida Fora Ponta (R$)",
      "Consumo Grupo B (kWh)",
      "Valor Consumo Grupo B (R$)",
      "Consumo Ponta (kWh)",
      "Valor Consumo Ponta (R$)",
      "Consumo Fora Ponta (kWh)",
      "Valor Consumo Fora Ponta (R$)",
      "Demanda Todos os Períodos (kW)",
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
      "GDI oUC (kWh)",
      "Valor GDI oUC (R$)",
      "GDI mUC (kWh)",
      "Valor GDI mUC (R$)",
      "CIP (R$)",
      "Outros Encargos (R$)",
      "PIS (R$)",
      "COFINS (R$)",
      "ICMS (R$)",
    ];

    const formatCSVValue = (val: any) => {
      if (val === null || val === undefined) return "";
      let str = String(val);
      // If it looks like a number with a dot, replace it with a comma for Brazilian Excel
      if (!isNaN(Number(val)) && str.includes(".") && !str.includes(",")) {
        str = str.replace(".", ",");
      }
      // Replace any existing semicolons to avoid breaking the CSV structure
      str = str.replace(/;/g, ",");
      return `"${str}"`;
    };

    const rows = completedBills.map((b) => {
      return [
        b.fileName,
        b.uc,
        getGerencia(String(b.uc)),
        getLocin(String(b.uc)),
        getCidade(String(b.uc), b.cidade),
        b.tipo || "",
        UCS_LIVRE_MERCADO_LIVRE.has(String(b.uc)) ? "LIVRE" : "CATIVO",
        b.concessionaria
          ? b.concessionaria.toUpperCase().includes("ENERGISA")
            ? "ENERGISA"
            : b.concessionaria.toUpperCase().includes("ELEKTRO")
              ? "ELEKTRO"
              : b.concessionaria
          : "",
        b.mesReferencia,
        b.anoLeitura,
        b.dataVencimento || "",
        b.numeroNotaFiscal || "",
        b.modalidadeTarifaria || "",
        b.subgrupo || "",
        b.valorTotal,
        b.consumoGrupoB || "",
        b.demandaPontaKW,
        b.demandaForaPontaKW,
        b.demandaPotenciaMedidaPonta,
        b.valorDemandaPotenciaMedidaPonta,
        b.demandaPotenciaMedidaForaPonta,
        b.valorDemandaPotenciaMedidaForaPonta,
        b.consumoKwhGrupoB || b.consumoKwh || "",
        b.valorConsumoKwhGrupoB || "",
        b.consumoKwhPonta,
        b.valorConsumoKwhPonta,
        b.consumoKwhForaPonta,
        b.valorConsumoKwhForaPonta,
        b.demandaTodosPeriodosKW || "",
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
        b.energiaAtvInjetadaGDIOUC,
        b.valorEnergiaAtvInjetadaGDIOUC,
        b.energiaAtvInjetadaGDIMUC,
        b.valorEnergiaAtvInjetadaGDIMUC,
        b.cip,
        b.outrosEncargos,
        b.pis || "",
        b.cofins || "",
        b.icms || "",
      ];
    });

    // Use semicolon as delimiter for better compatibility with Excel in many locales (like Brazil)
    // Add UTF-8 BOM (\uFEFF) to ensure Excel recognizes the encoding
    const csvContent =
      "\uFEFF" +
      [
        headers.join(";"),
        ...rows.map((row) => row.map(formatCSVValue).join(";")),
      ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    const fileName = `extracao_faturas_consolidado_${new Date().toISOString().split("T")[0]}.csv`;

    link.setAttribute("href", url);
    link.setAttribute("download", fileName);
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const exportRelatorioToCSV = () => {
    if (filteredRelatorioData.length === 0) {
      showAlert(
        "Exportação",
        "Não há dados para exportar com os filtros selecionados.",
      );
      return;
    }

    const headers = [
      "Nome do Arquivo",
      "Mês/Ano",
      "UC",
      "Gerência",
      "LOCINS",
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
      "Demanda ponta - kW",
      "Demanda fora ponta - kW",
      "Ultrapassagem Ponta (kW)",
      "Ultrapassagem Fora Ponta (kW)",
      "Energia Reativa Exced em KWh - Ponta",
      "Energia Reativa Exced em KWh - Fponta",
      "Solar Injetada OUC (kWh)",
      "Solar Injetada MUC (kWh)",
      "Valor GDI oUC (R$)",
      "Valor GDI mUC (R$)",
      "CIP (R$)",
      "Outros Encargos (R$)",
      "PIS (R$)",
      "COFINS (R$)",
      "ICMS (R$)",
    ];

    const formatCSVValue = (val: any) => {
      if (val === null || val === undefined) return "";
      let str = String(val);
      if (!isNaN(Number(val)) && str.includes(".") && !str.includes(",")) {
        str = str.replace(".", ",");
      }
      str = str.replace(/;/g, ",");
      return `"${str}"`;
    };

    const rows = filteredRelatorioData.map((d) => {
      return [
        d.fileName,
        d.name,
        d.uc,
        getGerencia(String(d.uc)),
        getLocin(String(d.uc)),
        d.tipo,
        d.mercado,
        d.concessionaria,
        getCidade(String(d.uc), d.cidade),
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
        d.valorSolarOUC,
        d.valorSolarMUC,
        d.cip,
        d.outrosEncargos,
        d.pis,
        d.cofins,
        d.icms,
      ];
    });

    const csvContent =
      "\uFEFF" +
      [
        headers.join(";"),
        ...rows.map((row) => row.map(formatCSVValue).join(";")),
      ].join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute(
      "download",
      `Relatorio_Financeiro_${selectedRelatorioMonth.replace("/", "_")}.csv`,
    );
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };
  // --- Dashboard Data Processing ---

  const { multasTotals, multasMonthlyData, multasUcList } = useMemo(() => {
    let ultrapassagem = 0;
    let reativa = 0;
    let subutilizacao = 0;
    let total = 0;

    const monthlyBreakdown: Record<
      string,
      {
        month: string;
        sortKey: string;
        ultrapassagem: number;
        reativa: number;
        subutilizacao: number;
        total: number;
      }
    > = {};
    const ucBreakdown: Record<
      string,
      {
        cidade: string;
        ultrapassagem: number;
        reativa: number;
        subutilizacao: number;
        total: number;
      }
    > = {};

    completedBills.forEach((b) => {
      const u =
        parseValue(b.valorDemandaPotenciaAtivaUltrapPonta) +
        parseValue(b.valorDemandaPotenciaAtivaUltrapFPonta);
      const r =
        parseValue(b.valorEnergiaReativaExcedPonta) +
        parseValue(b.valorEnergiaReativaExcedFPonta);
      const s =
        parseValue(b.valorDemandaPotenciaNaoConsumidaPonta) +
        parseValue(b.valorDemandaPotenciaNaoConsumidaFPonta);
      const t = u + r + s;

      const monthName = `${formatMonth(b.mesReferencia)}/${b.anoLeitura}`;
      const sortKey = `${b.anoLeitura}${getMonthNumber(b.mesReferencia.toString()).toString().padStart(2, "0")}`;

      if (!monthlyBreakdown[sortKey]) {
        monthlyBreakdown[sortKey] = {
          month: monthName,
          sortKey,
          ultrapassagem: 0,
          reativa: 0,
          subutilizacao: 0,
          total: 0,
        };
      }
      monthlyBreakdown[sortKey].ultrapassagem += u;
      monthlyBreakdown[sortKey].reativa += r;
      monthlyBreakdown[sortKey].subutilizacao += s;
      monthlyBreakdown[sortKey].total += t;

      if (multasMonth === "all" || monthName === multasMonth) {
        ultrapassagem += u;
        reativa += r;
        subutilizacao += s;
        total += t;

        if (!ucBreakdown[b.uc]) {
          ucBreakdown[b.uc] = {
            cidade: getCidade(String(b.uc), b.cidade || "N/A"),
            ultrapassagem: 0,
            reativa: 0,
            subutilizacao: 0,
            total: 0,
          };
        }
        ucBreakdown[b.uc].ultrapassagem += u;
        ucBreakdown[b.uc].reativa += r;
        ucBreakdown[b.uc].subutilizacao += s;
        ucBreakdown[b.uc].total += t;
      }
    });

    const monthlyData = Object.values(monthlyBreakdown).sort((a, b) =>
      a.sortKey.localeCompare(b.sortKey),
    );

    const ucList = Object.entries(ucBreakdown)
      .map(([uc, data]) => ({ uc, ...data }))
      .filter((d) => d[selectedMultaType] > 0)
      .sort((a, b) => {
        if (multasSortDirection === "asc") {
          return a[selectedMultaType] - b[selectedMultaType];
        } else {
          return b[selectedMultaType] - a[selectedMultaType];
        }
      });

    return {
      multasTotals: { ultrapassagem, reativa, subutilizacao, total },
      multasMonthlyData: monthlyData,
      multasUcList: ucList,
    };
  }, [completedBills, selectedMultaType, multasMonth, multasSortDirection]);

  const dashboardData = completedBills.map((b) => ({
    name: `${formatMonth(b.mesReferencia)}/${b.anoLeitura}`,
    uc: b.uc,
    consumoPonta: parseValue(b.consumoKwhPonta),
    valorConsumoPonta: parseValue(b.valorConsumoKwhPonta),
    consumoForaPonta: parseValue(b.consumoKwhForaPonta),
    valorConsumoForaPonta: parseValue(b.valorConsumoKwhForaPonta),
    consumoGrupoB: parseValue(b.consumoKwhGrupoB || b.consumoKwh),
    valorConsumoGrupoB: parseValue(b.valorConsumoKwhGrupoB),
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
    solarInjetadaOUC: parseValue(b.energiaAtvInjetadaGDIOUC),
    solarInjetadaMUC: parseValue(b.energiaAtvInjetadaGDIMUC),
    valorUltrapassagemPonta: parseValue(b.valorDemandaPotenciaAtivaUltrapPonta),
    valorUltrapassagemForaPonta: parseValue(
      b.valorDemandaPotenciaAtivaUltrapFPonta,
    ),
    valorReativaPonta: parseValue(b.valorEnergiaReativaExcedPonta),
    valorReativaForaPonta: parseValue(b.valorEnergiaReativaExcedFPonta),
    valorSolarOUC: parseValue(b.valorEnergiaAtvInjetadaGDIOUC),
    valorSolarMUC: parseValue(b.valorEnergiaAtvInjetadaGDIMUC),
    cip: parseValue(b.cip),
    outrosEncargos: parseValue(b.outrosEncargos),
    pis: parseValue(b.pis),
    cofins: parseValue(b.cofins),
    icms: parseValue(b.icms),
    concessionaria: b.concessionaria || "",
    numeroNotaFiscal: b.numeroNotaFiscal || "",
    cidade: getCidade(String(b.uc), b.cidade),
    tipo: UCS_PPP.has(String(b.uc)) ? "PPP Fotovoltaica" : b.tipo || "",
    mercado: UCS_LIVRE_MERCADO_LIVRE.has(String(b.uc)) ? "LIVRE" : "CATIVO",
    fileName: b.fileName || "",
    modalidadeTarifaria: (b.modalidadeTarifaria || "").toString().toUpperCase(),
    subgrupo: (b.subgrupo || "").toString().toUpperCase(),
  }));

  const ucs = Array.from(new Set(dashboardData.map((d) => d.uc))).filter(
    Boolean,
  );

  const availableMonths = Array.from(new Set(dashboardData.map((d) => d.name)))
    .filter(Boolean)
    .sort((a, b) => {
      const [mA, yA] = String(a).split("/");
      const [mB, yB] = String(b).split("/");
      if (yA !== yB) return Number(yB) - Number(yA);
      return getMonthNumber(mB) - getMonthNumber(mA);
    });

  const availableRelatorioTypes = React.useMemo(() => {
    const types = Array.from(new Set(dashboardData.map((d) => d.tipo)))
      .filter(Boolean)
      .sort();
    if (types.includes("OPER") && !types.includes("INJETADO")) {
      const operIndex = types.indexOf("OPER");
      types.splice(operIndex + 1, 0, "INJETADO");
    } else if (!types.includes("INJETADO")) {
      types.push("INJETADO");
    }
    return types;
  }, [dashboardData]);

  const filteredDashboardData = dashboardData.filter((d) => {
    const matchesUC =
      !selectedUC ||
      selectedUC === "all" ||
      d.uc.toString().includes(selectedUC);
    const matchesMonth = selectedMonth === "all" || d.name === selectedMonth;
    const matchesConcessionaria =
      selectedConcessionaria === "all" ||
      d.concessionaria === selectedConcessionaria;

    if (!matchesUC || !matchesMonth || !matchesConcessionaria) return false;

    if (
      dashboardSubTab === "financeiro" &&
      financialSubTab === "energia_solar"
    ) {
      const totalCreditos = Math.abs(d.valorSolarOUC + d.valorSolarMUC);
      return totalCreditos > 0;
    }

    if (dashboardSubTab === "operacionais" && operationalSubTab === "solar") {
      const totalInjetada = Math.abs(d.solarInjetadaOUC + d.solarInjetadaMUC);
      return totalInjetada > 0;
    }

    if (dashboardSubTab === "operacionais" && operationalSubTab === "reativa") {
      return d.reativaPonta > 0 || d.reativaForaPonta > 0;
    }

    if (
      dashboardSubTab === "operacionais" &&
      operationalSubTab === "ultrapassagem"
    ) {
      return d.ultrapassagemPonta + d.ultrapassagemForaPonta > 0;
    }

    if (
      dashboardSubTab === "operacionais" &&
      operationalSubTab === "subutilizacao"
    ) {
      return d.demandaContratadaPonta > 0 || d.demandaContratadaForaPonta > 0;
    }

    if (
      dashboardSubTab === "financeiro" &&
      financialSubTab === "multa_reativa"
    ) {
      return d.valorReativaPonta > 0 || d.valorReativaForaPonta > 0;
    }

    if (
      dashboardSubTab === "financeiro" &&
      financialSubTab === "multa_ultrapassagem"
    ) {
      return d.valorUltrapassagemPonta + d.valorUltrapassagemForaPonta > 0;
    }

    return true;
  });

  const generalFilteredData = dashboardData.filter((d) => {
    const matchesUC =
      !selectedUC ||
      selectedUC === "all" ||
      d.uc.toString().includes(selectedUC);
    const matchesMonth = selectedMonth === "all" || d.name === selectedMonth;
    const matchesConcessionaria =
      selectedConcessionaria === "all" ||
      d.concessionaria === selectedConcessionaria;
    return matchesUC && matchesMonth && matchesConcessionaria;
  });

  const filteredRelatorioData = dashboardData.filter((d) => {
    const matchesMonth =
      selectedRelatorioMonth === "all" || d.name === selectedRelatorioMonth;
    const matchesType =
      selectedRelatorioType.includes("all") ||
      selectedRelatorioType.includes(d.tipo) ||
      (selectedRelatorioType.includes("INJETADO") && UCS_PPP.has(String(d.uc)));
    return matchesMonth && matchesType && d.uc !== "31383580";
  });

  const relatorioTotals = React.useMemo(() => {
    return filteredRelatorioData.reduce(
      (acc, curr) => {
        acc.valorTotal += curr.valorTotal || 0;
        acc.consumoTotal +=
          (curr.consumoPonta || 0) + (curr.consumoForaPonta || 0) + (curr.consumoGrupoB || 0);

        // For Relatório Financeiro cards, only include PPP Fotovoltaica UCs for injected energy
        if (UCS_PPP.has(String(curr.uc))) {
          acc.valorInjetado +=
            (curr.valorSolarOUC || 0) + (curr.valorSolarMUC || 0);
          acc.totalInjetadoKwh +=
            (curr.solarInjetadaOUC || 0) + (curr.solarInjetadaMUC || 0);
        }

        return acc;
      },
      { valorTotal: 0, consumoTotal: 0, valorInjetado: 0, totalInjetadoKwh: 0 },
    );
  }, [filteredRelatorioData]);

  const filteredUcs = Array.from(
    new Set(filteredDashboardData.map((d) => d.uc)),
  ).filter(Boolean);

  const sortedDashboardData = React.useMemo(() => {
    return [...filteredDashboardData].sort((a, b) => {
      let aValue: any;
      let bValue: any;

      switch (dashboardSort.key) {
        case "name":
          const [mA, yA] = a.name.split("/");
          const [mB, yB] = b.name.split("/");
          aValue = Number(yA) * 12 + getMonthNumber(mA);
          bValue = Number(yB) * 12 + getMonthNumber(mB);
          break;
        case "uc":
          aValue = a.uc;
          bValue = b.uc;
          break;
        case "total_kw":
          aValue = a.ultrapassagemPonta + a.ultrapassagemForaPonta;
          bValue = b.ultrapassagemPonta + b.ultrapassagemForaPonta;
          break;
        case "utilizacao":
          aValue = a.demandaMedidaPonta / (a.demandaContratadaPonta || 1);
          bValue = b.demandaMedidaPonta / (b.demandaContratadaPonta || 1);
          break;
        case "total_kvarh":
          aValue = a.reativaPonta + a.reativaForaPonta;
          bValue = b.reativaPonta + b.reativaForaPonta;
          break;
        default:
          aValue = a[dashboardSort.key as keyof typeof a];
          bValue = b[dashboardSort.key as keyof typeof b];
      }

      if (aValue === undefined || aValue === null) return 1;
      if (bValue === undefined || bValue === null) return -1;

      if (aValue < bValue) return dashboardSort.direction === "asc" ? -1 : 1;
      if (aValue > bValue) return dashboardSort.direction === "asc" ? 1 : -1;
      return 0;
    });
  }, [filteredDashboardData, dashboardSort]);

  const memoData = React.useMemo(() => {
    const energisa = filteredRelatorioData.filter((d) =>
      (d.concessionaria || "").toUpperCase().includes("ENERGISA"),
    );
    const elektro = filteredRelatorioData.filter((d) =>
      (d.concessionaria || "").toUpperCase().includes("ELEKTRO"),
    );

    const sum = (arr: any[], field: string) =>
      arr.reduce((acc, curr) => acc + (curr[field] || 0), 0);

    const energisaData = {
      total: sum(energisa, "valorTotal"),
      pis: sum(energisa, "pis"),
      cofins: sum(energisa, "cofins"),
      icms: sum(energisa, "icms"),
      cip: sum(energisa, "cip"),
      nf:
        memoNfEnergisa || agrupadoraFiles["ENERGISA"]?.numeroNotaFiscal || "-",
      mesRef:
        selectedRelatorioMonth === "all"
          ? agrupadoraFiles["ENERGISA"]
            ? formatReference(agrupadoraFiles["ENERGISA"].mesReferencia)
            : "-"
          : selectedRelatorioMonth,
    };

    const elektroData = {
      total: sum(elektro, "valorTotal"),
      pis: sum(elektro, "pis"),
      cofins: sum(elektro, "cofins"),
      icms: sum(elektro, "icms"),
      cip:
        agrupadoraFiles["ELEKTRO_DETALHADO"]?.cip || sum(elektro, "cip") || 0,
      nf: memoNfElektro || agrupadoraFiles["ELEKTRO"]?.numeroNotaFiscal || "-",
      mesRef:
        selectedRelatorioMonth === "all"
          ? agrupadoraFiles["ELEKTRO"]
            ? formatReference(agrupadoraFiles["ELEKTRO"].mesReferencia)
            : "-"
          : selectedRelatorioMonth,
    };

    return {
      energisa: energisaData,
      elektro: elektroData,
    };
  }, [
    filteredRelatorioData,
    agrupadoraFiles,
    memoNfEnergisa,
    memoNfElektro,
    selectedRelatorioMonth,
  ]);

  // Group by month/year for charts
  const timeSeriesData = Object.values(
    filteredDashboardData.reduce((acc: any, curr) => {
      const key = curr.name;
      if (!acc[key]) {
        acc[key] = {
          name: key,
          consumoPonta: 0,
          valorConsumoPonta: 0,
          consumoForaPonta: 0,
          valorConsumoForaPonta: 0,
          consumoGrupoB: 0,
          valorConsumoGrupoB: 0,
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
          outrosEncargos: 0,
        };
      }
      acc[key].consumoPonta += curr.consumoPonta;
      acc[key].valorConsumoPonta += curr.valorConsumoPonta;
      acc[key].consumoForaPonta += curr.consumoForaPonta;
      acc[key].valorConsumoForaPonta += curr.valorConsumoForaPonta;
      acc[key].consumoGrupoB += curr.consumoGrupoB;
      acc[key].valorConsumoGrupoB += curr.valorConsumoGrupoB;
      acc[key].valorTotal += curr.valorTotal;
      acc[key].demandaMedidaPonta = Math.max(
        acc[key].demandaMedidaPonta,
        curr.demandaMedidaPonta,
      );
      acc[key].demandaMedidaForaPonta = Math.max(
        acc[key].demandaMedidaForaPonta,
        curr.demandaMedidaForaPonta,
      );
      acc[key].demandaContratadaPonta = Math.max(
        acc[key].demandaContratadaPonta,
        curr.demandaContratadaPonta,
      );
      acc[key].demandaContratadaForaPonta = Math.max(
        acc[key].demandaContratadaForaPonta,
        curr.demandaContratadaForaPonta,
      );
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
      acc[key].valorConsumo =
        acc[key].valorTotal +
        Math.abs(acc[key].valorSolarOUC + acc[key].valorSolarMUC) -
        acc[key].cip -
        acc[key].outrosEncargos;
      return acc;
    }, {}),
  );

  const COLORS = ["#0054A6", "#00AEEF", "#1E293B", "#64748B"];



  // Dashboard is always accessible - Supabase RLS anon policy handles read security.
  // Login via Supabase Auth is optional (used only for write-protected ops if needed).

  // Show full-screen loading while fetching from Supabase
  if (currentPage === "visao_geral" && isSyncing && bills.length === 0) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center gap-6">
        <Logo className="h-14" />
        <div className="flex flex-col items-center gap-3">
          <Loader2 size={36} className="animate-spin text-sanesul-primary" />
          <p className="text-sm font-bold text-slate-600">Carregando faturas do Supabase Cloud...</p>
          <p className="text-xs text-slate-400">Aguarde enquanto buscamos os dados</p>
        </div>
      </div>
    );
  }

  if (currentPage === "visao_geral") {
    return (
      <VisaoGeralDashboard
        data={dashboardData}
        setCurrentPage={setCurrentPage}
        handleLogout={handleLogout}
        hasApiKey={hasApiKey}
        handleSelectKey={handleSelectKey}
      />
    );
  }

  return (
    <div className="min-h-screen bg-sanesul-bg text-sanesul-text font-sans p-4 md:p-8">
      {/* Header */}
      <header className="max-w-[1600px] mx-auto mb-8 border-b border-slate-200/80 pb-6 space-y-6">
        {/* Top Navbar */}
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
          <div className="flex items-center gap-4">
            <Logo className="h-12" />
            <div className="hidden sm:flex items-center gap-2 pl-4 border-l border-slate-200">
              <span className="flex h-2.5 w-2.5 relative">
                <span className="animate-ping absolute inline-flex h-full w-full rounded-full opacity-75 bg-emerald-400"></span>
                <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500"></span>
              </span>
              <span className="text-[11px] font-bold text-slate-600 tracking-wider uppercase">
                Sistema Operacional (Nuvem Supabase)
              </span>
            </div>
          </div>

          {/* Quick Actions & System Controls */}
          <div className="flex flex-wrap gap-2.5 items-center">
            {/* Supabase Database Badge */}
            <div
              className="flex items-center gap-2 px-3.5 py-2 border rounded-xl text-xs font-semibold shadow-xs transition-all bg-sky-50/90 text-sky-800 border-sky-200"
              title={`Banco de Dados Supabase Cloud: ${bills.length.toLocaleString()} faturas`}
            >
              <Cloud size={15} className="text-sky-600" />
              <span className="font-mono font-bold">
                {`Supabase: ${bills.length.toLocaleString()} faturas`}
              </span>
            </div>


            {/* Espelhar no Supabase (1:1) */}
            <button
              onClick={mirrorAppToSupabase}
              disabled={isSyncing}
              className="flex items-center gap-2 px-3.5 py-2 bg-gradient-to-r from-sky-600 to-blue-700 text-white hover:from-sky-500 hover:to-blue-600 transition-all rounded-xl text-xs font-bold shadow-sm shadow-blue-500/20 active:scale-95 disabled:opacity-50"
              title="Garante que o Supabase contenha exatamente as faturas presentes no App (Sincronização Estrita 1:1)"
            >
              {isSyncing ? (
                <Loader2 size={15} className="animate-spin text-white" />
              ) : (
                <FileUp size={15} />
              )}
              <span>Espelhar no Supabase (1:1)</span>
            </button>

            {/* Configurar API */}
            <button
              onClick={handleSelectKey}
              className={`flex items-center gap-2 px-4 py-2 border transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95 ${
                hasApiKey
                  ? "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
                  : "bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100"
              }`}
              title={hasApiKey ? "Chave de IA Ativa - Trocar Chave" : "Configurar Chave Gemini AI"}
            >
              <Key size={15} />
              <span>{hasApiKey ? "IA Conectada" : "Configurar API"}</span>
            </button>

            {/* Voltar para Visão Geral */}
            <button
              onClick={() => setCurrentPage("visao_geral")}
              className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 hover:border-sanesul-primary/40 text-slate-700 hover:text-sanesul-primary transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
            >
              <ArrowLeft size={15} />
              <span>Visão Geral</span>
            </button>

            {/* Sair */}
            <button
              onClick={handleLogout}
              className="flex items-center gap-2 px-3.5 py-2 bg-white border border-slate-200 hover:border-red-300 text-slate-600 hover:text-red-600 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
              title="Encerrar Sessão"
            >
              <LogOut size={15} />
              <span>Sair</span>
            </button>
          </div>
        </div>

        {/* Navigation Tabs Bar */}
        <div className="flex items-center overflow-x-auto no-scrollbar py-1">
          <div className="flex gap-1.5 p-1.5 bg-slate-100/80 backdrop-blur-sm border border-slate-200/80 rounded-2xl w-fit">
            {[
              { id: "faturas", label: "Gestão de Faturas", icon: FileText },
              { id: "multas", label: "Multas", icon: AlertTriangle },
              { id: "dashboard", label: "Dashboard Analítico", icon: LayoutDashboard },
              { id: "analises", label: "Análises de Dados", icon: BarChart3 },
              { id: "monitoramento", label: "Monitoramento Demanda", icon: DollarSign },
              { id: "monitoramento_ajustes", label: "Ajuste de Demanda", icon: FileCheck },
              { id: "monitoramento_reativo", label: "Monitoramento Reativo", icon: Zap },
              { id: "monitoramento_usinas", label: "Monitoramento Usinas", icon: Leaf },
              { id: "relatorio", label: "Relatório Financeiro", icon: FileSpreadsheet },
            ].map((tab) => {
              const Icon = tab.icon;
              const isActive = activeTab === tab.id;
              return (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id as any)}
                  className={`flex items-center gap-2 px-5 py-2.5 transition-all rounded-xl text-xs font-bold tracking-wide whitespace-nowrap ${
                    isActive
                      ? "bg-sanesul-primary text-white shadow-md shadow-sanesul-primary/20 scale-[1.02]"
                      : "text-slate-600 hover:text-sanesul-primary hover:bg-white/80"
                  }`}
                >
                  <Icon size={15} />
                  <span>{tab.label}</span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Modern Executive Action Toolbar (Displayed for activeTab === "faturas") */}
        {activeTab === "faturas" && (
          <div className="bg-white rounded-2xl border border-slate-200/90 shadow-sm p-4 space-y-4">
            {/* Top Toolbar Row: Ingestion, Process & Strategy */}
            <div className="flex flex-wrap items-center justify-between gap-4">
              
              {/* Group 1: Adicionar & Importar Faturas */}
              <div className="flex flex-wrap items-center gap-2">
                <div className="flex items-center gap-1 bg-slate-50 border border-slate-200/80 rounded-xl p-1">
                  {/* Energisa Group */}
                  <span className="text-[10px] font-black text-blue-800 bg-blue-100/80 px-2 py-1 rounded-lg uppercase tracking-wider">
                    Energisa
                  </span>
                  <button
                    onClick={() => fileInputEnergisaRef.current?.click()}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-sanesul-primary text-white hover:bg-sanesul-primary/90 transition-all rounded-lg text-xs font-bold shadow-xs active:scale-95"
                    title="Adicionar arquivos de faturas Energisa (PDF)"
                  >
                    <Plus size={14} />
                    <span>Faturas</span>
                  </button>
                  <button
                    onClick={() => folderInputEnergisaRef.current?.click()}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 text-white hover:bg-emerald-700 transition-all rounded-lg text-xs font-bold shadow-xs active:scale-95"
                    title="Adicionar pasta completa de faturas Energisa"
                  >
                    <FolderPlus size={14} />
                    <span>Pasta</span>
                  </button>
                </div>

                <div className="flex items-center gap-1 bg-slate-50 border border-slate-200/80 rounded-xl p-1">
                  {/* Elektro Group */}
                  <span className="text-[10px] font-black text-emerald-800 bg-emerald-100/80 px-2 py-1 rounded-lg uppercase tracking-wider">
                    Elektro
                  </span>
                  <button
                    onClick={() => fileInputElektroRef.current?.click()}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-sanesul-primary text-white hover:bg-sanesul-primary/90 transition-all rounded-lg text-xs font-bold shadow-xs active:scale-95"
                    title="Adicionar arquivos de faturas Elektro (PDF)"
                  >
                    <Plus size={14} />
                    <span>Faturas</span>
                  </button>
                  <button
                    onClick={() => folderInputElektroRef.current?.click()}
                    className="flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 text-white hover:bg-emerald-700 transition-all rounded-lg text-xs font-bold shadow-xs active:scale-95"
                    title="Adicionar pasta completa de faturas Elektro"
                  >
                    <FolderPlus size={14} />
                    <span>Pasta</span>
                  </button>
                </div>

                {/* Planilhas & Manual Group */}
                <div className="flex items-center gap-1.5">
                  <label className="flex items-center gap-1.5 px-3 py-2 bg-blue-50/80 border border-blue-200/80 text-blue-700 hover:bg-blue-100 transition-all rounded-xl text-xs font-bold cursor-pointer active:scale-95 shadow-xs">
                    <FileSpreadsheet size={15} />
                    <span>Importar Planilha</span>
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      className="hidden"
                      onChange={handleExcelImport}
                    />
                  </label>
                  <button
                    onClick={downloadExcelTemplate}
                    className="flex items-center gap-1.5 px-3 py-2 bg-emerald-50/80 border border-emerald-200/80 text-emerald-700 hover:bg-emerald-100 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                    title="Baixar modelo de planilha Excel para preenchimento"
                  >
                    <Download size={15} />
                    <span>Baixar Modelo</span>
                  </button>
                  <button
                    onClick={() => {
                      setEditingBill({
                        id: crypto.randomUUID(),
                        fileName: "Fatura Manual",
                        status: "completed",
                        tipo: "OPERACIONAL",
                        concessionaria: "ENERGISA",
                        mesReferencia: `${formatMonth((new Date().getMonth() + 1).toString().padStart(2, "0"))}/${new Date().getFullYear()}`,
                        anoLeitura: new Date().getFullYear().toString(),
                      });
                      setIsBillModalOpen(true);
                    }}
                    className="flex items-center gap-1.5 px-3 py-2 bg-white border border-slate-200 hover:border-sanesul-primary/40 text-slate-700 hover:text-sanesul-primary transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                    title="Cadastrar fatura manualmente"
                  >
                    <Plus size={15} />
                    <span>Nova Manual</span>
                  </button>
                </div>
              </div>

              {/* Group 2: Processamento IA (Gemini) Highlight Button */}
              {bills.length > 0 && (
                <div className="flex items-center gap-2">
                  <button
                    onClick={startProcessing}
                    disabled={isProcessing || !bills.some((b) => b.status === "pending")}
                    className={`flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-bold tracking-wide transition-all shadow-md active:scale-95 ${
                      bills.some((b) => b.status === "pending") && !isProcessing
                        ? "bg-gradient-to-r from-blue-600 via-indigo-600 to-blue-700 text-white shadow-blue-500/25 hover:shadow-blue-500/40 hover:scale-[1.02] ring-2 ring-blue-400/30"
                        : "bg-slate-100 text-slate-400 border border-slate-200 cursor-not-allowed shadow-none"
                    }`}
                  >
                    {isProcessing ? (
                      <>
                        <Loader2 size={16} className="animate-spin text-white" />
                        <span>Processando IA...</span>
                      </>
                    ) : (
                      <>
                        <Sparkles size={16} className={bills.some((b) => b.status === "pending") ? "text-amber-300 animate-pulse" : ""} />
                        <span>
                          Processar Arquivos
                          {bills.some((b) => b.status === "pending") && (
                            <span className="ml-1.5 px-1.5 py-0.5 bg-white/20 rounded-md text-[10px] font-black">
                              {bills.filter((b) => b.status === "pending").length}
                            </span>
                          )}
                        </span>
                      </>
                    )}
                  </button>
                </div>
              )}
            </div>

            {/* Bottom Toolbar Row: Filters, Cadastros, Export & Batch Operations */}
            <div className="flex flex-wrap items-center justify-between gap-3 pt-3 border-t border-slate-100">
              
              {/* Left Subgroup: Cadastros & Gestão */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setIsGerenciasModalOpen(true)}
                  className="flex items-center gap-2 px-3.5 py-2 bg-slate-50 border border-slate-200/90 text-slate-700 hover:bg-slate-100 hover:text-sanesul-primary transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                >
                  <Building size={15} className="text-sanesul-primary" />
                  <span>Gerências & LOCINs</span>
                </button>

                <button
                  onClick={() => {
                    setMercadoLivreInput(Array.from(UCS_LIVRE_MERCADO_LIVRE).join("\n"));
                    setIsMercadoLivreModalOpen(true);
                  }}
                  className="flex items-center gap-2 px-3.5 py-2 bg-slate-50 border border-slate-200/90 text-slate-700 hover:bg-slate-100 hover:text-sanesul-primary transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                >
                  <DollarSign size={15} className="text-emerald-600" />
                  <span>Mercado Livre</span>
                </button>
              </div>

              {/* Right Subgroup: Filtering, Export, Selection & Deletion */}
              <div className="flex flex-wrap items-center gap-2">
                {/* Reference Filter */}
                {bills.length > 0 && (
                  <div className="flex items-center gap-2 bg-slate-50 border border-slate-200/90 rounded-xl px-3 py-1.5 shadow-xs">
                    <Filter size={14} className="text-sanesul-primary" />
                    <select
                      value={filterReference}
                      onChange={(e) => setFilterReference(e.target.value)}
                      className="bg-transparent text-xs font-bold text-slate-700 outline-none cursor-pointer pr-2"
                    >
                      <option value="all">Todas as Referências</option>
                      {availableReferences.map((ref) => (
                        <option key={ref} value={ref}>
                          {ref}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                {/* Exportar CSV */}
                {bills.some((b) => b.status === "completed") && (
                  <button
                    onClick={exportToCSV}
                    className="flex items-center gap-2 px-4 py-2 bg-sanesul-secondary text-white hover:bg-sanesul-secondary/90 transition-all rounded-xl text-xs font-bold shadow-xs shadow-sanesul-secondary/20 active:scale-95"
                    title="Exportar todas as faturas para formato CSV"
                  >
                    <Download size={15} />
                    <span>Exportar CSV</span>
                  </button>
                )}

                {/* Espelhar Supabase 1:1 */}
                {isSupabaseConfigured && bills.some((b) => b.status === "completed") && (
                  <button
                    onClick={mirrorAppToSupabase}
                    disabled={isSyncing}
                    className="flex items-center gap-2 px-3.5 py-2 bg-sky-50 border border-sky-200 text-sky-700 hover:bg-sky-100 hover:border-sky-300 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95 disabled:opacity-50"
                    title="Espelhar faturas no Supabase mantendo a contagem 100% idêntica ao aplicativo"
                  >
                    {isSyncing ? (
                      <Loader2 size={15} className="animate-spin text-sky-600" />
                    ) : (
                      <FileUp size={15} className="text-sky-600" />
                    )}
                    <span>Espelhar 1:1</span>
                  </button>
                )}

                {/* Selecionar Pendentes / Erros */}
                {bills.length > 0 && (
                  <button
                    onClick={() =>
                      setSelectedBills(
                        bills
                          .filter((b) => b.status === "error" || b.status === "pending")
                          .map((b) => b.id),
                      )
                    }
                    className="flex items-center gap-2 px-3.5 py-2 bg-white border border-slate-200 hover:border-slate-300 text-slate-600 hover:text-slate-900 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                    title="Selecionar faturas com erro ou aguardando processamento"
                  >
                    <CheckSquare size={15} />
                    <span>Selecionar Pendentes/Erro</span>
                  </button>
                )}

                {/* Resetar Travados (Condicional) */}
                {bills.some((b) => b.status === "processing") && (
                  <button
                    onClick={resetStuckProcesses}
                    className="flex items-center gap-2 px-3.5 py-2 bg-amber-50 border border-amber-200 text-amber-700 hover:bg-amber-100 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                    title="Reseta faturas que ficaram presas no status 'Processando'"
                  >
                    <RotateCcw size={15} />
                    <span>Resetar Travados</span>
                  </button>
                )}

                {/* Excluir Mês Filtrado */}
                {bills.length > 0 && filterReference !== "all" && (
                  <button
                    onClick={() => removeMonthBills(filterReference)}
                    className="flex items-center gap-2 px-3.5 py-2 bg-red-50 border border-red-200 text-red-600 hover:bg-red-100 hover:border-red-300 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                    title={`Excluir todas as faturas do mês ${filterReference}`}
                  >
                    <CalendarX2 size={15} />
                    <span>Excluir Mês ({filterReference})</span>
                  </button>
                )}

                {/* Excluir por Lista de UCs */}
                {bills.length > 0 && (
                  <button
                    onClick={() => {
                      setDeleteUcListInput("");
                      setIsDeleteByListModalOpen(true);
                    }}
                    className="flex items-center gap-2 px-3.5 py-2 bg-red-50 border border-red-200 text-red-600 hover:bg-red-100 hover:border-red-300 transition-all rounded-xl text-xs font-bold shadow-xs active:scale-95"
                    title="Excluir faturas colando uma lista de UCs"
                  >
                    <ListX size={15} />
                    <span>Excluir p/ Lista</span>
                  </button>
                )}
              </div>
            </div>

            {/* Hidden native file inputs (Required for file picker triggers) */}
            <input
              type="file"
              ref={fileInputEnergisaRef}
              onChange={(e) => handleFileUpload(e, "ENERGISA")}
              multiple
              accept="application/pdf,image/*"
              className="hidden"
            />
            <input
              type="file"
              ref={folderInputEnergisaRef}
              onChange={(e) => handleFileUpload(e, "ENERGISA")}
              webkitdirectory=""
              directory=""
              multiple
              className="hidden"
            />
            <input
              type="file"
              ref={fileInputElektroRef}
              onChange={(e) => handleFileUpload(e, "ELEKTRO")}
              multiple
              accept="application/pdf,image/*"
              className="hidden"
            />
            <input
              type="file"
              ref={folderInputElektroRef}
              onChange={(e) => handleFileUpload(e, "ELEKTRO")}
              webkitdirectory=""
              directory=""
              multiple
              className="hidden"
            />
          </div>
        )}
      </header>



      <main
        className="max-w-[1600px] mx-auto"
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        {activeTab === "faturas" ? (
          bills.length === 0 ? (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-12 items-center py-12">
              <div className="space-y-8">
                <div className="inline-flex items-center gap-2 px-4 py-2 bg-sanesul-primary/10 rounded-full">
                  <span className="w-2 h-2 rounded-full bg-sanesul-primary animate-pulse" />
                  <span className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">
                    Pronto para processar
                  </span>
                </div>
                <h2 className="text-5xl md:text-6xl font-display font-bold text-sanesul-primary leading-[1.1] tracking-tight">
                  Transforme suas{" "}
                  <span className="text-sanesul-secondary">faturas</span> em
                  inteligência.
                </h2>
                <p className="text-lg text-sanesul-muted max-w-md leading-relaxed">
                  Nossa IA extrai automaticamente todos os indicadores técnicos
                  e financeiros das suas faturas de energia em segundos.
                </p>
                <div className="flex flex-wrap gap-4 pt-4">
                  <div className="flex items-center gap-3 px-6 py-4 bg-white rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <div className="w-10 h-10 rounded-xl bg-green-50 flex items-center justify-center">
                      <CheckCircle2 className="w-5 h-5 text-green-600" />
                    </div>
                    <span className="text-sm font-bold text-sanesul-primary">
                      Extração Precisa
                    </span>
                  </div>
                  <div className="flex items-center gap-3 px-6 py-4 bg-white rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <div className="w-10 h-10 rounded-xl bg-blue-50 flex items-center justify-center">
                      <Zap className="w-5 h-5 text-sanesul-primary" />
                    </div>
                    <span className="text-sm font-bold text-sanesul-primary">
                      Análise em Tempo Real
                    </span>
                  </div>
                </div>
              </div>

              <div
                className={`relative aspect-square lg:aspect-auto lg:h-[500px] border-2 border-dashed rounded-[40px] flex flex-col items-center justify-center transition-all group overflow-hidden ${
                  isDragging
                    ? "border-sanesul-primary bg-sanesul-primary/5 scale-[0.98]"
                    : "border-sanesul-primary/20 bg-white/50 hover:border-sanesul-primary/40 hover:bg-white"
                }`}
              >
                <div className="absolute inset-0 bg-gradient-to-br from-sanesul-primary/5 to-transparent opacity-0 group-hover:opacity-100 transition-opacity" />
                <div className="relative z-10 flex flex-col items-center">
                  <div className="w-24 h-24 bg-sanesul-primary rounded-3xl flex items-center justify-center mb-8 shadow-2xl shadow-sanesul-primary/30 group-hover:scale-110 transition-transform duration-500">
                    <Upload size={40} className="text-white" />
                  </div>
                  <p className="text-2xl font-display font-bold text-sanesul-primary mb-3">
                    {isDragging ? "Solte agora" : "Arraste suas faturas"}
                  </p>
                  <p className="text-sm text-sanesul-muted text-center max-w-[240px] mb-8">
                    Suporta PDF, JPG e PNG. Processamento automático via Gemini
                    AI.
                  </p>
                  <div className="flex flex-col sm:flex-row gap-4">
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        fileInputEnergisaRef.current?.click();
                      }}
                      className="px-8 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest shadow-lg shadow-sanesul-primary/20 hover:bg-sanesul-primary/90 transition-all"
                    >
                      Selecionar ENERGISA
                    </button>
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        fileInputElektroRef.current?.click();
                      }}
                      className="px-8 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest shadow-lg shadow-sanesul-primary/20 hover:bg-sanesul-primary/90 transition-all"
                    >
                      Selecionar ELEKTRO
                    </button>
                  </div>

                  <div className="mt-8 flex flex-col sm:flex-row items-center gap-4 pt-8 border-t border-slate-200 w-full max-w-md">
                    <p className="text-sm font-bold text-slate-500 w-full text-center sm:text-left">
                      Ou preencha manualmente:
                    </p>
                    <button
                      onClick={(e) => {
                        e.stopPropagation();
                        downloadExcelTemplate();
                      }}
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
            <div
              className={`space-y-8 transition-all ${isDragging ? "opacity-50 scale-[0.99]" : ""}`}
            >
              {/* Stats Bar */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                {[
                  {
                    label:
                      filterReference === "all"
                        ? "Total de Arquivos"
                        : `Arquivos (${filterReference})`,
                    value: sortedBills.length,
                    color: "sanesul-primary",
                    icon: FileText,
                  },
                  {
                    label: "Aguardando",
                    value: sortedBills.filter((b) => b.status === "pending")
                      .length,
                    color: "slate-500",
                    icon: Clock,
                  },
                  {
                    label: "Em Processamento",
                    value: sortedBills.filter((b) => b.status === "processing")
                      .length,
                    color: "sanesul-secondary",
                    icon: Loader2,
                  },
                  {
                    label: "Concluídos",
                    value: sortedBills.filter((b) => b.status === "completed")
                      .length,
                    color: "green-600",
                    icon: CheckCircle2,
                  },
                ].map((stat, i) => (
                  <div
                    key={i}
                    className="bg-white p-6 rounded-3xl border border-sanesul-primary/5 shadow-sm flex items-center gap-4"
                  >
                    <div
                      className={`w-12 h-12 rounded-2xl bg-${stat.color === "sanesul-primary" ? "sanesul-primary" : stat.color === "sanesul-secondary" ? "sanesul-secondary" : stat.color}/10 flex items-center justify-center`}
                    >
                      <stat.icon
                        size={20}
                        className={`text-${stat.color === "sanesul-primary" ? "sanesul-primary" : stat.color === "sanesul-secondary" ? "sanesul-secondary" : stat.color} ${stat.label === "Em Processamento" ? "animate-spin" : ""}`}
                      />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-sanesul-muted uppercase tracking-widest">
                        {stat.label}
                      </p>
                      <p
                        className={`text-2xl font-display font-bold text-${stat.color === "sanesul-primary" ? "sanesul-primary" : stat.color === "sanesul-secondary" ? "sanesul-secondary" : stat.color}`}
                      >
                        {stat.value}
                      </p>
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
                    onClick={() => setSearchUC("")}
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
                            checked={
                              selectedBills.length > 0 &&
                              selectedBills.length === bills.length
                            }
                            onChange={() =>
                              setSelectedBills(
                                selectedBills.length === bills.length
                                  ? []
                                  : bills.map((b) => b.id),
                              )
                            }
                            className="rounded border-sanesul-primary/20 text-sanesul-primary focus:ring-sanesul-primary"
                          />
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("fileName")}
                        >
                          Arquivo
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("uc")}
                        >
                          UC
                        </th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">
                          Gerência
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("cidade")}
                        >
                          Cidade
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("concessionaria")}
                        >
                          Concessionária
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("referencia")}
                        >
                          Referência
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("dataVencimento")}
                        >
                          Vencimento
                        </th>
                        <th
                          className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary"
                          onClick={() => requestSort("mercado")}
                        >
                          Mercado
                        </th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">
                          Tipo
                        </th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5">
                          Status
                        </th>
                        <th className="px-4 py-3 text-[9px] font-bold text-sanesul-muted uppercase tracking-widest border-b border-sanesul-primary/5 text-right">
                          {selectedBills.length > 0 ? (
                            <div className="flex items-center justify-end gap-4">
                              <button
                                onClick={removeSelectedBills}
                                className="text-red-600 hover:text-red-700 font-bold"
                              >
                                Excluir ({selectedBills.length})
                              </button>
                              {bills.length >= 223 && (
                                <button
                                  onClick={deselectFirst223}
                                  className="text-sanesul-primary hover:text-sanesul-secondary font-bold"
                                >
                                  Deselecionar 223
                                </button>
                              )}
                            </div>
                          ) : (
                            "Ações"
                          )}
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-sanesul-primary/5">
                      <AnimatePresence initial={false}>
                        {(() => {
                          const paginatedSortedBills = sortedBills.slice(
                            0,
                            100,
                          );
                          return (
                            <>
                              {paginatedSortedBills.map((bill) => (
                                <motion.tr
                                  key={bill.id}
                                  initial={{ opacity: 0, y: 10 }}
                                  animate={{ opacity: 1, y: 0 }}
                                  exit={{ opacity: 0, x: -20 }}
                                  className={`hover:bg-sanesul-primary/5 transition-colors group ${selectedBills.includes(bill.id) ? "bg-sanesul-primary/5" : ""}`}
                                >
                                  <td className="px-4 py-3">
                                    <input
                                      type="checkbox"
                                      checked={selectedBills.includes(bill.id)}
                                      onChange={() =>
                                        toggleBillSelection(bill.id)
                                      }
                                      className="rounded border-sanesul-primary/20 text-sanesul-primary focus:ring-sanesul-primary"
                                    />
                                  </td>
                                  <td className="px-4 py-3">
                                    <div className="flex items-center gap-4">
                                      <div className="w-8 h-8 bg-sanesul-primary/5 rounded-lg flex items-center justify-center text-sanesul-primary group-hover:bg-sanesul-primary group-hover:text-white transition-all">
                                        <FileText size={16} />
                                      </div>
                                      <div className="flex flex-col">
                                        <span
                                          className="text-xs font-bold text-sanesul-primary truncate max-w-[200px]"
                                          title={bill.fileName}
                                        >
                                          {bill.fileName}
                                        </span>
                                        <span className="text-[9px] text-sanesul-muted uppercase tracking-wider">
                                          {bill.file
                                            ? `${(bill.file.size / 1024 / 1024).toFixed(2)} MB • ${bill.file.type.split("/")[1].toUpperCase()}`
                                            : "ARQUIVO SALVO"}
                                        </span>
                                      </div>
                                    </div>
                                  </td>
                                  <td className="px-4 py-3">
                                    <div className="flex flex-col">
                                      <span className="text-xs font-mono font-bold text-sanesul-primary">
                                        {bill.uc || "---"}
                                      </span>
                                    </div>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span className="text-xs font-bold text-slate-800">
                                      {bill.gerencia && bill.gerencia !== "---" ? bill.gerencia : getGerencia(bill.uc || "")}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                                      {bill.cidade && bill.cidade !== "---" ? bill.cidade : getCidade(String(bill.uc), bill.cidade)}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                                      {bill.concessionaria
                                        ? bill.concessionaria
                                            .toUpperCase()
                                            .includes("ENERGISA")
                                          ? "ENERGISA"
                                          : bill.concessionaria
                                                .toUpperCase()
                                                .includes("ELEKTRO")
                                            ? "ELEKTRO"
                                            : bill.concessionaria
                                        : "---"}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span className="text-xs text-slate-600">
                                      {bill.mesReferencia && bill.anoLeitura
                                        ? `${formatMonth(bill.mesReferencia)}/${bill.anoLeitura}`
                                        : "---"}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span className="text-xs text-slate-600">
                                      {bill.dataVencimento || "---"}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span
                                      className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold uppercase tracking-wider ${
                                        (bill.mercado ||
                                          (bill.uc &&
                                          UCS_LIVRE_MERCADO_LIVRE.has(
                                            String(bill.uc),
                                          )
                                            ? "LIVRE"
                                            : "CATIVO")) === "LIVRE"
                                          ? "bg-blue-100 text-blue-700"
                                          : "bg-slate-100 text-slate-600"
                                      }`}
                                    >
                                      {bill.mercado ||
                                        (bill.uc &&
                                        UCS_LIVRE_MERCADO_LIVRE.has(
                                          String(bill.uc),
                                        )
                                          ? "LIVRE"
                                          : "CATIVO")}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <span
                                      className={`inline-flex items-center px-2 py-0.5 rounded-full text-[9px] font-bold uppercase tracking-wider ${
                                        bill.tipo === "OPERACIONAL"
                                          ? "bg-blue-50 text-blue-600"
                                          : bill.tipo === "ADMINISTRATIVO"
                                            ? "bg-purple-50 text-purple-600"
                                            : "bg-slate-100 text-slate-500"
                                      }`}
                                    >
                                      {bill.tipo || "N/A"}
                                    </span>
                                  </td>
                                  <td className="px-4 py-3">
                                    <div className="flex items-center gap-2">
                                      {bill.status === "pending" && (
                                        <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-slate-100 text-slate-600 rounded-full text-[9px] font-bold uppercase tracking-wider">
                                          <Clock size={10} />
                                          Aguardando
                                        </span>
                                      )}
                                      {bill.status === "processing" && (
                                        <div className="flex flex-col gap-1 w-full min-w-[100px]">
                                          <div className="flex items-center justify-between">
                                            <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-blue-50 text-sanesul-primary rounded-full text-[9px] font-bold uppercase tracking-wider">
                                              <Loader2
                                                size={10}
                                                className="animate-spin"
                                              />
                                              Extraindo... {bill.progress || 0}%
                                            </span>
                                            <button
                                              onClick={() =>
                                                bill.abortController?.abort()
                                              }
                                              className="text-red-500 hover:text-red-700 p-0.5 rounded-full hover:bg-red-50 transition-colors"
                                              title="Cancelar"
                                            >
                                              <X size={12} />
                                            </button>
                                          </div>
                                          <div className="w-full bg-slate-100 rounded-full h-1 overflow-hidden">
                                            <div
                                              className="bg-sanesul-primary h-1 rounded-full transition-all duration-300 ease-out"
                                              style={{
                                                width: `${bill.progress || 0}%`,
                                              }}
                                            ></div>
                                          </div>
                                          {bill.error && (
                                            <span className="text-[9px] text-amber-600 font-bold flex items-center gap-1 mt-1 bg-amber-50 px-2 py-0.5 rounded-full w-fit">
                                              <AlertCircle size={10} />{" "}
                                              {bill.error}
                                            </span>
                                          )}
                                        </div>
                                      )}
                                      {bill.status === "completed" && (
                                        <div className="flex flex-col gap-1">
                                          <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-green-50 text-green-600 rounded-full text-[10px] font-bold uppercase tracking-wider w-fit">
                                            <CheckCircle2 size={12} />
                                            Concluído
                                          </span>
                                          {bill.error && (
                                            <span className="text-[9px] text-amber-600 font-bold flex items-center gap-1 mt-1 bg-amber-50 px-2 py-0.5 rounded-full w-fit">
                                              <AlertCircle size={10} />{" "}
                                              {bill.error}
                                            </span>
                                          )}
                                        </div>
                                      )}
                                      {bill.status === "error" && (
                                        <span
                                          className="inline-flex items-center gap-1.5 px-3 py-1 bg-red-50 text-red-600 rounded-full text-[10px] font-bold uppercase tracking-wider"
                                          title={bill.error}
                                        >
                                          <AlertCircle size={12} />
                                          {bill.error || "Erro"}
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
                                        disabled={
                                          isProcessing &&
                                          bill.status === "processing"
                                        }
                                        title="Editar"
                                      >
                                        <Pencil size={16} />
                                      </button>
                                      <button
                                        onClick={() => removeBill(bill.id)}
                                        className="w-8 h-8 inline-flex items-center justify-center rounded-lg text-slate-400 hover:bg-red-50 hover:text-red-500 transition-all"
                                        disabled={
                                          isProcessing &&
                                          bill.status === "processing"
                                        }
                                        title="Excluir"
                                      >
                                        <Trash2 size={16} />
                                      </button>
                                    </div>
                                  </td>
                                </motion.tr>
                              ))}
                              {sortedBills.length > 100 && (
                                <tr>
                                  <td
                                    colSpan={13}
                                    className="px-4 py-6 text-center text-sm text-slate-500 italic bg-slate-50 border-t border-slate-100"
                                  >
                                    Exibindo os primeiros 100 de{" "}
                                    {sortedBills.length} resultados. Use a
                                    pesquisa para ver mais.
                                  </td>
                                </tr>
                              )}
                            </>
                          );
                        })()}
                      </AnimatePresence>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )
        ) : activeTab === "multas" ? (
          <div className="space-y-8">
            <div className="bg-white p-10 rounded-3xl border border-sanesul-primary/10 shadow-xl">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-10">
                <div>
                  <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                    Análise de Multas
                  </h2>
                  <p className="text-sanesul-muted">
                    Acompanhe as multas por ultrapassagem, reativa e
                    subutilização.
                  </p>
                </div>
              </div>

              {/* Chart Section */}
              <div className="bg-slate-50 rounded-3xl p-8 mb-10 border border-sanesul-primary/5">
                <h3 className="text-lg font-bold text-sanesul-primary mb-6">
                  {selectedMultaType === "total"
                    ? "Evolução Mensal - Gasto Total"
                    : selectedMultaType === "ultrapassagem"
                      ? "Evolução Mensal - Multa de Ultrapassagem"
                      : selectedMultaType === "reativa"
                        ? "Evolução Mensal - Multa Reativa"
                        : "Evolução Mensal - Subutilização"}
                </h3>
                <div className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <AreaChart
                      data={multasMonthlyData}
                      margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                    >
                      <defs>
                        <linearGradient
                          id="colorTotal"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="5%"
                            stopColor="#8b5cf6"
                            stopOpacity={0.3}
                          />
                          <stop
                            offset="95%"
                            stopColor="#8b5cf6"
                            stopOpacity={0}
                          />
                        </linearGradient>
                        <linearGradient
                          id="colorUltrapassagem"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="5%"
                            stopColor="#ef4444"
                            stopOpacity={0.3}
                          />
                          <stop
                            offset="95%"
                            stopColor="#ef4444"
                            stopOpacity={0}
                          />
                        </linearGradient>
                        <linearGradient
                          id="colorReativa"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="5%"
                            stopColor="#f59e0b"
                            stopOpacity={0.3}
                          />
                          <stop
                            offset="95%"
                            stopColor="#f59e0b"
                            stopOpacity={0}
                          />
                        </linearGradient>
                        <linearGradient
                          id="colorSubutilizacao"
                          x1="0"
                          y1="0"
                          x2="0"
                          y2="1"
                        >
                          <stop
                            offset="5%"
                            stopColor="#3b82f6"
                            stopOpacity={0.3}
                          />
                          <stop
                            offset="95%"
                            stopColor="#3b82f6"
                            stopOpacity={0}
                          />
                        </linearGradient>
                      </defs>
                      <CartesianGrid
                        strokeDasharray="3 3"
                        vertical={false}
                        stroke="#e2e8f0"
                      />
                      <XAxis
                        dataKey="month"
                        axisLine={false}
                        tickLine={false}
                        tick={{ fontSize: 12, fill: "#64748b" }}
                        dy={10}
                      />
                      <YAxis
                        axisLine={false}
                        tickLine={false}
                        tick={{ fontSize: 12, fill: "#64748b" }}
                        tickFormatter={(value) =>
                          `R$ ${value.toLocaleString("pt-BR")}`
                        }
                      />
                      <Tooltip
                        cursor={{
                          stroke: "#cbd5e1",
                          strokeWidth: 1,
                          strokeDasharray: "3 3",
                        }}
                        contentStyle={{
                          borderRadius: "16px",
                          border: "none",
                          boxShadow:
                            "0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)",
                        }}
                        formatter={(value: number) => [
                          `R$ ${value.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`,
                          "Valor",
                        ]}
                      />
                      <Area
                        type="monotone"
                        dataKey={selectedMultaType}
                        stroke={
                          selectedMultaType === "total"
                            ? "#8b5cf6"
                            : selectedMultaType === "ultrapassagem"
                              ? "#ef4444"
                              : selectedMultaType === "reativa"
                                ? "#f59e0b"
                                : "#3b82f6"
                        }
                        fillOpacity={1}
                        fill={`url(#color${selectedMultaType === "total" ? "Total" : selectedMultaType === "ultrapassagem" ? "Ultrapassagem" : selectedMultaType === "reativa" ? "Reativa" : "Subutilizacao"})`}
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
                  onClick={() => setSelectedMultaType("total")}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === "total"
                      ? "bg-violet-50 border-violet-200 shadow-xl shadow-violet-100"
                      : "bg-white border-slate-100 hover:border-violet-100 hover:bg-violet-50/50"
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div
                      className={`p-3 rounded-xl ${selectedMultaType === "total" ? "bg-violet-100 text-violet-600" : "bg-slate-50 text-slate-400"}`}
                    >
                      <Calculator size={24} />
                    </div>
                    <h3
                      className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === "total" ? "text-violet-900" : "text-slate-500"}`}
                    >
                      Gasto Total
                    </h3>
                  </div>
                  <p
                    className={`text-3xl font-display font-bold ${selectedMultaType === "total" ? "text-violet-600" : "text-slate-700"}`}
                  >
                    R${" "}
                    {multasTotals.total.toLocaleString("pt-BR", {
                      minimumFractionDigits: 2,
                    })}
                  </p>
                </div>
                <div
                  onClick={() => setSelectedMultaType("reativa")}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === "reativa"
                      ? "bg-amber-50 border-amber-200 shadow-xl shadow-amber-100"
                      : "bg-white border-slate-100 hover:border-amber-100 hover:bg-amber-50/50"
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div
                      className={`p-3 rounded-xl ${selectedMultaType === "reativa" ? "bg-amber-100 text-amber-600" : "bg-slate-50 text-slate-400"}`}
                    >
                      <Zap size={24} />
                    </div>
                    <h3
                      className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === "reativa" ? "text-amber-900" : "text-slate-500"}`}
                    >
                      Multa Reativa
                    </h3>
                  </div>
                  <p
                    className={`text-3xl font-display font-bold ${selectedMultaType === "reativa" ? "text-amber-600" : "text-slate-700"}`}
                  >
                    R${" "}
                    {multasTotals.reativa.toLocaleString("pt-BR", {
                      minimumFractionDigits: 2,
                    })}
                  </p>
                </div>

                <div
                  onClick={() => setSelectedMultaType("ultrapassagem")}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === "ultrapassagem"
                      ? "bg-red-50 border-red-200 shadow-xl shadow-red-100"
                      : "bg-white border-slate-100 hover:border-red-100 hover:bg-red-50/50"
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div
                      className={`p-3 rounded-xl ${selectedMultaType === "ultrapassagem" ? "bg-red-100 text-red-600" : "bg-slate-50 text-slate-400"}`}
                    >
                      <TrendingUp size={24} />
                    </div>
                    <h3
                      className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === "ultrapassagem" ? "text-red-900" : "text-slate-500"}`}
                    >
                      Ultrapassagem de Demanda
                    </h3>
                  </div>
                  <p
                    className={`text-3xl font-display font-bold ${selectedMultaType === "ultrapassagem" ? "text-red-600" : "text-slate-700"}`}
                  >
                    R${" "}
                    {multasTotals.ultrapassagem.toLocaleString("pt-BR", {
                      minimumFractionDigits: 2,
                    })}
                  </p>
                </div>

                <div
                  onClick={() => setSelectedMultaType("subutilizacao")}
                  className={`p-8 rounded-3xl border cursor-pointer transition-all ${
                    selectedMultaType === "subutilizacao"
                      ? "bg-blue-50 border-blue-200 shadow-xl shadow-blue-100"
                      : "bg-white border-slate-100 hover:border-blue-100 hover:bg-blue-50/50"
                  }`}
                >
                  <div className="flex items-center gap-4 mb-4">
                    <div
                      className={`p-3 rounded-xl ${selectedMultaType === "subutilizacao" ? "bg-blue-100 text-blue-600" : "bg-slate-50 text-slate-400"}`}
                    >
                      <TrendingDown size={24} />
                    </div>
                    <h3
                      className={`font-bold text-sm uppercase tracking-wider ${selectedMultaType === "subutilizacao" ? "text-blue-900" : "text-slate-500"}`}
                    >
                      Subutilização
                    </h3>
                  </div>
                  <p
                    className={`text-3xl font-display font-bold ${selectedMultaType === "subutilizacao" ? "text-blue-600" : "text-slate-700"}`}
                  >
                    R${" "}
                    {multasTotals.subutilizacao.toLocaleString("pt-BR", {
                      minimumFractionDigits: 2,
                    })}
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
                      {selectedMultaType === "total"
                        ? "UCs com gastos totais (ultrapassagem, reativa ou subutilização) no período."
                        : selectedMultaType === "ultrapassagem"
                          ? "UCs com multas de ultrapassagem no período."
                          : selectedMultaType === "reativa"
                            ? "UCs com multas reativas no período."
                            : "UCs com subutilização no período."}
                    </p>
                  </div>
                  <div className="flex items-center gap-4">
                    <select
                      value={multasMonth}
                      onChange={(e) => setMultasMonth(e.target.value)}
                      className="px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-bold text-sanesul-primary focus:ring-2 focus:ring-sanesul-primary/20"
                    >
                      <option value="all">Todos os Meses</option>
                      {Array.from(
                        new Set(
                          bills
                            .filter((b) => b.status === "completed")
                            .map(
                              (b) =>
                                `${formatMonth(b.mesReferencia)}/${b.anoLeitura}`,
                            ),
                        ),
                      ).map((month) => (
                        <option key={month} value={month}>
                          {month}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-slate-50/80 border-b border-sanesul-primary/10">
                        <th className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted">
                          UC
                        </th>
                        <th className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted">
                          Gerência
                        </th>
                        <th className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted">
                          Cidade
                        </th>
                        <th
                          className="p-4 font-bold text-xs uppercase tracking-wider text-sanesul-muted text-right cursor-pointer hover:text-sanesul-primary transition-colors group"
                          onClick={() =>
                            setMultasSortDirection((prev) =>
                              prev === "asc" ? "desc" : "asc",
                            )
                          }
                        >
                          <div className="flex items-center justify-end gap-1">
                            Valor Total (R$)
                            <div className="flex flex-col">
                              <ArrowUp
                                size={10}
                                className={`${multasSortDirection === "asc" ? "text-sanesul-primary" : "text-slate-300 group-hover:text-slate-400"}`}
                              />
                              <ArrowDown
                                size={10}
                                className={`-mt-1 ${multasSortDirection === "desc" ? "text-sanesul-primary" : "text-slate-300 group-hover:text-slate-400"}`}
                              />
                            </div>
                          </div>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {multasUcList.length > 0 ? (
                        multasUcList.map((item, index) => (
                          <tr
                            key={item.uc}
                            className={`border-b border-slate-100 hover:bg-slate-50 transition-colors ${index % 2 === 0 ? "bg-white" : "bg-slate-50/30"}`}
                          >
                            <td className="p-4 font-medium text-sanesul-primary">
                              {item.uc}
                            </td>
                            <td className="p-4 text-slate-800 font-bold text-sm">
                              {getGerencia(item.uc)}
                            </td>
                            <td className="p-4 text-slate-600 font-bold text-xs uppercase">
                              {item.cidade}
                            </td>
                            <td className="p-4 text-right font-bold text-sanesul-primary">
                              R${" "}
                              {item[selectedMultaType].toLocaleString("pt-BR", {
                                minimumFractionDigits: 2,
                              })}
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td
                            colSpan={4}
                            className="p-8 text-center text-sanesul-muted"
                          >
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
        ) : activeTab === "dashboard" ? (
          <div className="flex flex-col md:flex-row gap-8">
            {/* Dashboard Sidebar Navigation */}
            <aside className="w-full md:w-72 space-y-6">
              <div className="flex flex-col gap-2 p-3 bg-white rounded-3xl border border-sanesul-primary/10 shadow-xl">
                <button
                  onClick={() => setDashboardSubTab("operacionais")}
                  className={`flex items-center gap-3 px-5 py-4 rounded-2xl text-xs font-bold uppercase tracking-wider transition-all ${
                    dashboardSubTab === "operacionais"
                      ? "bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20"
                      : "text-sanesul-muted hover:bg-sanesul-primary/5 hover:text-sanesul-primary"
                  }`}
                >
                  <Zap size={16} />
                  Operacionais
                </button>

                {dashboardSubTab === "operacionais" && (
                  <div className="ml-4 flex flex-col gap-1 border-l-2 border-sanesul-primary/10 pl-4 py-2">
                    {[
                      { id: "consumo", label: "Consumo de Energia" },
                      {
                        id: "ultrapassagem",
                        label: "Ultrapassagem de Demanda",
                      },
                      {
                        id: "subutilizacao",
                        label: "Subutilização de Demanda",
                      },
                      { id: "reativa", label: "Energia Reativa" },
                      { id: "solar", label: "Energia Solar" },
                    ].map((item) => (
                      <button
                        key={item.id}
                        onClick={() => setOperationalSubTab(item.id as any)}
                        className={`text-left px-3 py-2.5 text-[11px] font-bold uppercase tracking-tight transition-all rounded-xl ${
                          operationalSubTab === item.id
                            ? "text-sanesul-primary bg-sanesul-primary/5"
                            : "text-sanesul-muted hover:text-sanesul-primary"
                        }`}
                      >
                        {item.label}
                      </button>
                    ))}
                  </div>
                )}

                <button
                  onClick={() => setDashboardSubTab("financeiro")}
                  className={`flex items-center gap-3 px-5 py-4 rounded-2xl text-xs font-bold uppercase tracking-wider transition-all ${
                    dashboardSubTab === "financeiro"
                      ? "bg-sanesul-primary text-white shadow-lg shadow-sanesul-primary/20"
                      : "text-sanesul-muted hover:bg-sanesul-primary/5 hover:text-sanesul-primary"
                  }`}
                >
                  <DollarSign size={16} />
                  Financeiro
                </button>

                {dashboardSubTab === "financeiro" && (
                  <div className="ml-4 flex flex-col gap-1 border-l-2 border-sanesul-primary/10 pl-4 py-2">
                    {[
                      { id: "despesas", label: "Despesas com Energia" },
                      {
                        id: "multa_ultrapassagem",
                        label: "Multa de Ultrapassagem",
                      },
                      {
                        id: "multa_reativa",
                        label: "Multa de Energia Reativa",
                      },
                      { id: "tarifa_media", label: "Tarifa Média" },
                      { id: "energia_solar", label: "Energia Solar" },
                    ].map((item) => (
                      <button
                        key={item.id}
                        onClick={() => setFinancialSubTab(item.id as any)}
                        className={`text-left px-3 py-2.5 text-[11px] font-bold uppercase tracking-tight transition-all rounded-xl ${
                          financialSubTab === item.id
                            ? "text-sanesul-primary bg-sanesul-primary/5"
                            : "text-sanesul-muted hover:text-sanesul-primary"
                        }`}
                      >
                        {item.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>

              <div className="p-6 bg-sanesul-primary rounded-3xl shadow-xl shadow-sanesul-primary/20 text-white hidden md:block">
                <h3 className="text-[10px] font-bold uppercase tracking-[0.2em] opacity-60 mb-4">
                  Resumo Geral
                </h3>
                <div className="space-y-4">
                  <div className="flex justify-between items-end">
                    <span className="text-[10px] font-bold uppercase opacity-60">
                      Processados
                    </span>
                    <span className="text-2xl font-display font-bold leading-none">
                      {generalFilteredData.length}
                    </span>
                  </div>
                  <div className="flex justify-between items-end">
                    <span className="text-[10px] font-bold uppercase opacity-60">
                      Unidades
                    </span>
                    <span className="text-2xl font-display font-bold leading-none">
                      {new Set(generalFilteredData.map((d) => d.uc)).size}
                    </span>
                  </div>
                  <div className="pt-4 border-t border-white/10 space-y-4">
                    <div className="flex flex-col">
                      <span className="text-[10px] font-bold uppercase opacity-60 mb-1">
                        Custo Total
                      </span>
                      <span className="text-xl font-display font-bold">
                        R${" "}
                        {generalFilteredData
                          .reduce((acc, curr) => acc + curr.valorTotal, 0)
                          .toLocaleString("pt-BR", {
                            minimumFractionDigits: 2,
                          })}
                      </span>
                    </div>
                    <div className="flex flex-col">
                      <span className="text-[10px] font-bold uppercase opacity-60 mb-1">
                        Consumo Total
                      </span>
                      <span className="text-xl font-display font-bold">
                        {generalFilteredData
                          .reduce(
                            (acc, curr) =>
                              acc + (curr.consumoPonta || 0) + (curr.consumoForaPonta || 0) + (curr.consumoGrupoB || 0) + (curr.consumoKwh || 0),
                            0,
                          )
                          .toLocaleString("pt-BR")}{" "}
                        <span className="text-xs opacity-60">kWh</span>
                      </span>
                    </div>
                  </div>
                </div>
              </div>
            </aside>

            <div className="flex-1 space-y-10">
              <div className="flex flex-col md:flex-row md:items-center justify-between border-b border-sanesul-primary/10 pb-6 gap-6">
                <div className="flex flex-col">
                  <h2 className="text-3xl font-display font-bold text-sanesul-primary">
                    {dashboardSubTab === "financeiro"
                      ? financialSubTab === "despesas"
                        ? "Despesas com Energia"
                        : financialSubTab === "multa_ultrapassagem"
                          ? "Multa de Ultrapassagem"
                          : financialSubTab === "multa_reativa"
                            ? "Multa de Energia Reativa"
                            : financialSubTab === "energia_solar"
                              ? "Créditos de Energia Solar"
                              : "Tarifa Média"
                      : operationalSubTab === "consumo"
                        ? "Consumo de Energia"
                        : operationalSubTab === "ultrapassagem"
                          ? "Ultrapassagem de Demanda"
                          : operationalSubTab === "subutilizacao"
                            ? "Subutilização de Demanda"
                            : operationalSubTab === "reativa"
                              ? "Energia Reativa"
                              : "Energia Solar"}
                  </h2>
                  <p className="text-[11px] font-bold uppercase tracking-wider text-sanesul-muted mt-1">
                    {!selectedUC || selectedUC === "all"
                      ? "Visão consolidada do grupo"
                      : `Unidade Consumidora: ${selectedUC}`}
                  </p>
                </div>

                <div className="flex flex-wrap items-center gap-3">
                  <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-sanesul-muted ml-2">
                      Concessionária:
                    </span>
                    <select
                      value={selectedConcessionaria}
                      onChange={(e) =>
                        setSelectedConcessionaria(e.target.value)
                      }
                      className="bg-sanesul-bg border-none px-4 py-2 rounded-xl text-[11px] font-bold uppercase tracking-wider text-sanesul-primary outline-none focus:ring-2 focus:ring-sanesul-primary/20 transition-all cursor-pointer"
                    >
                      <option value="all">Todas</option>
                      <option value="ENERGISA">Energisa</option>
                      <option value="ELEKTRO">Elektro</option>
                    </select>
                  </div>
                  <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-sanesul-muted ml-2">
                      Mês:
                    </span>
                    <select
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(e.target.value)}
                      className="bg-sanesul-bg border-none px-4 py-2 rounded-xl text-[11px] font-bold uppercase tracking-wider text-sanesul-primary outline-none focus:ring-2 focus:ring-sanesul-primary/20 transition-all cursor-pointer"
                    >
                      <option value="all">Todos os Meses</option>
                      {availableMonths.map((month) => (
                        <option key={month} value={month}>
                          {month}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="flex items-center gap-3 bg-white p-2 rounded-2xl border border-sanesul-primary/10 shadow-sm">
                    <span className="text-[10px] font-bold uppercase tracking-wider text-sanesul-muted ml-2">
                      Filtrar UC:
                    </span>
                    <div className="relative flex items-center">
                      <Search
                        size={14}
                        className="absolute left-3 text-sanesul-primary/40"
                      />
                      <input
                        type="text"
                        value={selectedUC === "all" ? "" : selectedUC}
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
                    <AlertCircle
                      size={32}
                      className="text-sanesul-primary/30"
                    />
                  </div>
                  <p className="text-xl font-display font-semibold text-sanesul-primary">
                    Nenhum dado disponível
                  </p>
                  <p className="text-sanesul-muted mt-2">
                    Processe algumas faturas para visualizar o dashboard
                    analítico
                  </p>
                </div>
              ) : (
                <div className="space-y-10">
                  {/* Summary Cards */}
                  <div
                    className={`grid grid-cols-1 gap-8 ${
                      (dashboardSubTab === "financeiro" &&
                        financialSubTab === "energia_solar") ||
                      (dashboardSubTab === "operacionais" &&
                        operationalSubTab === "consumo")
                        ? "md:grid-cols-2 lg:grid-cols-3"
                        : "md:grid-cols-2 lg:grid-cols-3"
                    }`}
                  >
                    {dashboardSubTab === "operacionais" ? (
                      <>
                        {operationalSubTab === "consumo" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total de Faturas
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {filteredDashboardData.length}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  Arquivos
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Consumo Total
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {Math.round(
                                  filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.consumoPonta +
                                      curr.consumoForaPonta,
                                    0,
                                  ),
                                ).toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kWh
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Grupo A
                              </p>
                              <div className="space-y-2">
                                <p className="text-lg font-bold text-green-600 flex justify-between">
                                  <span>Verde:</span>
                                  <span>
                                    {Math.round(
                                      filteredDashboardData
                                        .filter((d) =>
                                          d.modalidadeTarifaria.includes(
                                            "VERDE",
                                          ),
                                        )
                                        .reduce(
                                          (acc, curr) =>
                                            acc +
                                            curr.consumoPonta +
                                            curr.consumoForaPonta,
                                          0,
                                        ),
                                    ).toLocaleString("pt-BR")}{" "}
                                    kWh
                                  </span>
                                </p>
                                <p className="text-lg font-bold text-sanesul-primary flex justify-between">
                                  <span>Azul:</span>
                                  <span>
                                    {Math.round(
                                      filteredDashboardData
                                        .filter((d) =>
                                          d.modalidadeTarifaria.includes(
                                            "AZUL",
                                          ),
                                        )
                                        .reduce(
                                          (acc, curr) =>
                                            acc +
                                            curr.consumoPonta +
                                            curr.consumoForaPonta,
                                          0,
                                        ),
                                    ).toLocaleString("pt-BR")}{" "}
                                    kWh
                                  </span>
                                </p>
                              </div>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Grupo B
                              </p>
                              <div className="space-y-2">
                                <p className="text-lg font-bold text-sanesul-primary flex justify-between">
                                  <span>Solar:</span>
                                  <span>
                                    {Math.round(
                                      filteredDashboardData
                                        .filter((d) => hasCompensacao(d))
                                        .reduce(
                                          (acc, curr) =>
                                            acc +
                                            curr.consumoPonta +
                                            curr.consumoForaPonta,
                                          0,
                                        ),
                                    ).toLocaleString("pt-BR")}{" "}
                                    kWh
                                  </span>
                                </p>
                                <p className="text-lg font-bold text-sanesul-primary flex justify-between">
                                  <span>Não Solar:</span>
                                  <span>
                                    {Math.round(
                                      filteredDashboardData
                                        .filter(
                                          (d) =>
                                            !hasCompensacao(d) &&
                                            d.consumoPonta +
                                              d.consumoForaPonta >
                                              0,
                                        )
                                        .reduce(
                                          (acc, curr) =>
                                            acc +
                                            curr.consumoPonta +
                                            curr.consumoForaPonta,
                                          0,
                                        ),
                                    ).toLocaleString("pt-BR")}{" "}
                                    kWh
                                  </span>
                                </p>
                              </div>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Tarifa Branca
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {filteredDashboardData
                                  .filter((d) =>
                                    d.modalidadeTarifaria.includes("BRANCA"),
                                  )
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.consumoPonta +
                                      curr.consumoForaPonta,
                                    0,
                                  )
                                  .toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kWh
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Optante B
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {
                                  filteredDashboardData.filter(
                                    (d) =>
                                      d.subgrupo.startsWith("B") &&
                                      d.demandaContratadaPonta > 0,
                                  ).length
                                }{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  Faturas
                                </span>
                              </p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === "ultrapassagem" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <AlertCircle
                                  size={80}
                                  className="text-red-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total Ultrapassagem
                              </p>
                              <p className="text-4xl font-display font-bold text-red-600">
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.ultrapassagemPonta +
                                      curr.ultrapassagemForaPonta,
                                    0,
                                  )
                                  .toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kW
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Ocorrências
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {
                                  new Set(
                                    filteredDashboardData
                                      .filter(
                                        (d) =>
                                          d.ultrapassagemPonta > 0 ||
                                          d.ultrapassagemForaPonta > 0,
                                      )
                                      .map((d) => d.uc),
                                  ).size
                                }{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  Unidades
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-red-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Maior Desvio
                              </p>
                              <p className="text-4xl font-display font-bold text-red-600">
                                {Math.max(
                                  ...filteredDashboardData.map(
                                    (d) =>
                                      d.ultrapassagemPonta +
                                      d.ultrapassagemForaPonta,
                                  ),
                                  0,
                                ).toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kW
                                </span>
                              </p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === "subutilizacao" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Média Utilização
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {(
                                  (filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.demandaMedidaPonta /
                                        (curr.demandaContratadaPonta || 1),
                                    0,
                                  ) / filteredDashboardData.length || 0) * 100
                                ).toFixed(1)}
                                %
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <AlertCircle
                                  size={80}
                                  className="text-orange-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Subutilizados (&lt;80%)
                              </p>
                              <p className="text-4xl font-display font-bold text-orange-600">
                                {
                                  new Set(
                                    filteredDashboardData
                                      .filter(
                                        (d) =>
                                          d.demandaMedidaPonta <
                                          d.demandaContratadaPonta * 0.8,
                                      )
                                      .map((d) => d.uc),
                                  ).size
                                }{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  Unidades
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-orange-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Mínima Utilização
                              </p>
                              <p className="text-4xl font-display font-bold text-orange-600">
                                {Math.min(
                                  ...filteredDashboardData.map(
                                    (d) =>
                                      (d.demandaMedidaPonta /
                                        (d.demandaContratadaPonta || 1)) *
                                      100,
                                  ),
                                  100,
                                ).toFixed(1)}
                                %
                              </p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === "reativa" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total Excedente
                              </p>
                              <p className="text-4xl font-display font-bold text-purple-600">
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.reativaPonta +
                                      curr.reativaForaPonta,
                                    0,
                                  )
                                  .toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kVArh
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText
                                  size={80}
                                  className="text-purple-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Unidades com Excesso
                              </p>
                              <p className="text-4xl font-display font-bold text-purple-600">
                                {
                                  new Set(
                                    filteredDashboardData
                                      .filter(
                                        (d) =>
                                          d.reativaPonta > 0 ||
                                          d.reativaForaPonta > 0,
                                      )
                                      .map((d) => d.uc),
                                  ).size
                                }{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  Unidades
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-purple-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Média Excedente
                              </p>
                              <p className="text-4xl font-display font-bold text-purple-600">
                                {(
                                  filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.reativaPonta +
                                      curr.reativaForaPonta,
                                    0,
                                  ) / timeSeriesData.length || 0
                                ).toFixed(1)}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kVArh
                                </span>
                              </p>
                            </div>
                          </>
                        )}
                        {operationalSubTab === "solar" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap
                                  size={80}
                                  className="text-sanesul-secondary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Consumo em kWh
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-secondary">
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      (curr.consumoPonta +
                                        curr.consumoForaPonta),
                                    0,
                                  )
                                  .toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kWh
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-green-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total Injetada
                              </p>
                              <p className="text-4xl font-display font-bold text-green-600">
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      (curr.solarInjetadaOUC +
                                        curr.solarInjetadaMUC),
                                    0,
                                  )
                                  .toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kWh
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-green-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Saldo Energia
                              </p>
                              <p className="text-4xl font-display font-bold text-green-600">
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      (curr.solarInjetadaOUC +
                                        curr.solarInjetadaMUC) -
                                      (curr.consumoPonta +
                                        curr.consumoForaPonta),
                                    0,
                                  )
                                  .toLocaleString("pt-BR")}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  kWh
                                </span>
                              </p>
                            </div>
                          </>
                        )}
                      </>
                    ) : (
                      <>
                        {financialSubTab === "despesas" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <FileText
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total de Faturas
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {filteredDashboardData.length}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  Arquivos
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Gasto Acumulado
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                R${" "}
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) => acc + curr.valorTotal,
                                    0,
                                  )
                                  .toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Média Mensal
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                R${" "}
                                {(
                                  filteredDashboardData.reduce(
                                    (acc, curr) => acc + curr.valorTotal,
                                    0,
                                  ) / timeSeriesData.length || 0
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <LayoutDashboard
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Unidades Ativas
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                {filteredUcs.length}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  UCs
                                </span>
                              </p>
                            </div>
                          </>
                        )}
                        {financialSubTab === "multa_ultrapassagem" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <AlertCircle
                                  size={80}
                                  className="text-red-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total Multas
                              </p>
                              <p className="text-4xl font-display font-bold text-red-600">
                                R${" "}
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.valorUltrapassagemPonta +
                                      curr.valorUltrapassagemForaPonta,
                                    0,
                                  )
                                  .toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-red-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Média Mensal
                              </p>
                              <p className="text-4xl font-display font-bold text-red-600">
                                R${" "}
                                {(
                                  filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.valorUltrapassagemPonta +
                                      curr.valorUltrapassagemForaPonta,
                                    0,
                                  ) / timeSeriesData.length || 0
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign
                                  size={80}
                                  className="text-red-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Maior Penalidade
                              </p>
                              <p className="text-4xl font-display font-bold text-red-600">
                                R${" "}
                                {Math.max(
                                  ...filteredDashboardData.map(
                                    (d) =>
                                      d.valorUltrapassagemPonta +
                                      d.valorUltrapassagemForaPonta,
                                  ),
                                  0,
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </p>
                            </div>
                          </>
                        )}
                        {financialSubTab === "multa_reativa" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap size={80} className="text-purple-600" />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total Multas
                              </p>
                              <p className="text-4xl font-display font-bold text-purple-600">
                                R${" "}
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.valorReativaPonta +
                                      curr.valorReativaForaPonta,
                                    0,
                                  )
                                  .toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-purple-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Média Mensal
                              </p>
                              <p className="text-4xl font-display font-bold text-purple-600">
                                R${" "}
                                {(
                                  filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.valorReativaPonta +
                                      curr.valorReativaForaPonta,
                                    0,
                                  ) / timeSeriesData.length || 0
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign
                                  size={80}
                                  className="text-purple-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Maior Penalidade
                              </p>
                              <p className="text-4xl font-display font-bold text-purple-600">
                                R${" "}
                                {Math.max(
                                  ...filteredDashboardData.map(
                                    (d) =>
                                      d.valorReativaPonta +
                                      d.valorReativaForaPonta,
                                  ),
                                  0,
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </p>
                            </div>
                          </>
                        )}
                        {financialSubTab === "tarifa_media" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Tarifa Média
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                R${" "}
                                {(
                                  filteredDashboardData.reduce(
                                    (acc, curr) => acc + curr.valorTotal,
                                    0,
                                  ) /
                                  (filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.consumoPonta +
                                      curr.consumoForaPonta,
                                    0,
                                  ) || 1)
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 3,
                                })}{" "}
                                <span className="text-base font-sans font-medium opacity-40">
                                  /kWh
                                </span>
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-green-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Menor Tarifa
                              </p>
                              <p className="text-4xl font-display font-bold text-green-600">
                                R${" "}
                                {Math.min(
                                  ...filteredDashboardData.map(
                                    (d) =>
                                      d.valorTotal /
                                      (d.consumoPonta + d.consumoForaPonta ||
                                        1),
                                  ),
                                  100,
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 3,
                                })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-red-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Maior Tarifa
                              </p>
                              <p className="text-4xl font-display font-bold text-red-600">
                                R${" "}
                                {Math.max(
                                  ...filteredDashboardData.map(
                                    (d) =>
                                      d.valorTotal /
                                      ((d.consumoPonta || 0) + (d.consumoForaPonta || 0) + (d.consumoGrupoB || 0) + (d.consumoKwh || 0) ||
                                        1),
                                  ),
                                  0,
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 3,
                                })}
                              </p>
                            </div>
                          </>
                        )}
                        {financialSubTab === "energia_solar" && (
                          <>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <Zap
                                  size={80}
                                  className="text-sanesul-primary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Consumo (R$)
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-primary">
                                R${" "}
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) =>
                                      acc +
                                      Math.abs(
                                        curr.valorSolarOUC + curr.valorSolarMUC,
                                      ) +
                                      (curr.valorTotal -
                                        curr.cip -
                                        curr.outrosEncargos),
                                    0,
                                  )
                                  .toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <TrendingUp
                                  size={80}
                                  className="text-green-600"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Total Créditos
                              </p>
                              <p className="text-4xl font-display font-bold text-green-600">
                                R${" "}
                                {Math.abs(
                                  filteredDashboardData.reduce(
                                    (acc, curr) =>
                                      acc +
                                      curr.valorSolarOUC +
                                      curr.valorSolarMUC,
                                    0,
                                  ),
                                ).toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </p>
                            </div>
                            <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                              <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                                <DollarSign
                                  size={80}
                                  className="text-sanesul-secondary"
                                />
                              </div>
                              <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                                Valor Total da Fatura
                              </p>
                              <p className="text-4xl font-display font-bold text-sanesul-secondary">
                                R${" "}
                                {filteredDashboardData
                                  .reduce(
                                    (acc, curr) => acc + curr.valorTotal,
                                    0,
                                  )
                                  .toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                              </p>
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
                          <h3 className="text-lg font-display font-bold text-sanesul-primary">
                            Detalhamento por Unidade Consumidora
                          </h3>
                          <p className="text-sm text-sanesul-muted mt-1">
                            Visão granular dos indicadores para cada registro no
                            período selecionado.
                          </p>
                        </div>
                        <div className="px-4 py-2 bg-sanesul-primary/10 rounded-full">
                          <span className="text-xs font-bold text-sanesul-primary uppercase tracking-widest">
                            {filteredDashboardData.length} Registros
                          </span>
                        </div>
                      </div>
                      <div className="overflow-auto max-h-[600px]">
                        <table className="w-full text-left border-collapse">
                          <thead className="sticky top-0 z-10 bg-slate-50">
                            <tr className="bg-slate-50/50">
                              <th
                                onClick={() =>
                                  setDashboardSort((prev) => ({
                                    key: "uc",
                                    direction:
                                      prev.key === "uc" &&
                                      prev.direction === "desc"
                                        ? "asc"
                                        : "desc",
                                  }))
                                }
                                className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:bg-sanesul-primary/5 transition-colors"
                              >
                                UC{" "}
                                {dashboardSort.key === "uc" &&
                                  (dashboardSort.direction === "asc"
                                    ? "↑"
                                    : "↓")}
                              </th>
                              <th
                                onClick={() =>
                                  setDashboardSort((prev) => ({
                                    key: "name",
                                    direction:
                                      prev.key === "name" &&
                                      prev.direction === "desc"
                                        ? "asc"
                                        : "desc",
                                  }))
                                }
                                className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:bg-sanesul-primary/5 transition-colors"
                              >
                                Mês/Ano{" "}
                                {dashboardSort.key === "name" &&
                                  (dashboardSort.direction === "asc"
                                    ? "↑"
                                    : "↓")}
                              </th>
                              <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">
                                Classificação
                              </th>
                              <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">
                                Mercado
                              </th>
                              <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">
                                Modalidade
                              </th>
                              {dashboardSubTab === "operacionais" ? (
                                <>
                                  {operationalSubTab === "consumo" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Consumo Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Consumo F. Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">
                                        Total (kWh)
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === "ultrapassagem" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Contratada Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Contratada F. Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Ultrap. Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Ultrap. F. Ponta
                                      </th>
                                      <th
                                        onClick={() =>
                                          setDashboardSort((prev) => ({
                                            key: "total_kw",
                                            direction:
                                              prev.key === "total_kw" &&
                                              prev.direction === "desc"
                                                ? "asc"
                                                : "desc",
                                          }))
                                        }
                                        className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-red-600 border-b border-sanesul-primary/5 text-right cursor-pointer hover:bg-red-50 transition-colors"
                                      >
                                        Total (kW){" "}
                                        {dashboardSort.key === "total_kw" &&
                                          (dashboardSort.direction === "asc"
                                            ? "↑"
                                            : "↓")}
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === "subutilizacao" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Contratada Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Contratada F. Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Medida Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Medida F. Ponta
                                      </th>
                                      <th
                                        onClick={() =>
                                          setDashboardSort((prev) => ({
                                            key: "utilizacao",
                                            direction:
                                              prev.key === "utilizacao" &&
                                              prev.direction === "desc"
                                                ? "asc"
                                                : "desc",
                                          }))
                                        }
                                        className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-orange-600 border-b border-sanesul-primary/5 text-right cursor-pointer hover:bg-orange-50 transition-colors"
                                      >
                                        Utilização (%){" "}
                                        {dashboardSort.key === "utilizacao" &&
                                          (dashboardSort.direction === "asc"
                                            ? "↑"
                                            : "↓")}
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === "reativa" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Reativa Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Reativa F. Ponta
                                      </th>
                                      <th
                                        onClick={() =>
                                          setDashboardSort((prev) => ({
                                            key: "total_kvarh",
                                            direction:
                                              prev.key === "total_kvarh" &&
                                              prev.direction === "desc"
                                                ? "asc"
                                                : "desc",
                                          }))
                                        }
                                        className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-purple-600 border-b border-sanesul-primary/5 text-right cursor-pointer hover:bg-purple-50 transition-colors"
                                      >
                                        Total (kVArh){" "}
                                        {dashboardSort.key === "total_kvarh" &&
                                          (dashboardSort.direction === "asc"
                                            ? "↑"
                                            : "↓")}
                                      </th>
                                    </>
                                  )}
                                  {operationalSubTab === "solar" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Consumo em kWh
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Injetada oUC
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Injetada mUC
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">
                                        Saldo Energia
                                      </th>
                                    </>
                                  )}
                                </>
                              ) : (
                                <>
                                  {financialSubTab === "despesas" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">
                                        Valor Total
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">
                                        Cidade
                                      </th>
                                    </>
                                  )}
                                  {financialSubTab ===
                                    "multa_ultrapassagem" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Multa Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Multa F. Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-red-600 border-b border-sanesul-primary/5 text-right">
                                        Total (R$)
                                      </th>
                                    </>
                                  )}
                                  {financialSubTab === "multa_reativa" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Multa Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Multa F. Ponta
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-purple-600 border-b border-sanesul-primary/5 text-right">
                                        Total (R$)
                                      </th>
                                    </>
                                  )}
                                  {financialSubTab === "tarifa_media" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Valor Total
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right">
                                        Consumo Total
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">
                                        Tarifa (R$/kWh)
                                      </th>
                                    </>
                                  )}
                                  {financialSubTab === "energia_solar" && (
                                    <>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-right">
                                        Consumo (R$)
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">
                                        Total Créditos
                                      </th>
                                      <th className="px-8 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-secondary border-b border-sanesul-primary/5 text-right">
                                        Valor Total da Fatura
                                      </th>
                                    </>
                                  )}
                                </>
                              )}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-sanesul-primary/5">
                            {(() => {
                              const paginatedDashboard =
                                sortedDashboardData.slice(0, 100);
                              return (
                                <>
                                  {paginatedDashboard.map((row, idx) => (
                                    <tr
                                      key={idx}
                                      className="hover:bg-sanesul-primary/5 transition-colors group"
                                    >
                                      <td className="px-8 py-5 text-sm font-bold text-sanesul-primary">
                                        {row.uc}
                                      </td>
                                      <td className="px-8 py-5 text-sm text-slate-600">
                                        {row.name}
                                      </td>
                                      <td className="px-8 py-5 text-sm font-medium text-slate-600">
                                        {UCS_PPP.has(String(row.uc))
                                          ? "PPP Fotovoltaica"
                                          : UCS_USINA.has(String(row.uc))
                                            ? "Usinas SANESUL"
                                            : "Geral"}
                                      </td>
                                      <td className="px-8 py-5 text-sm font-bold">
                                        <span
                                          className={`px-3 py-1 rounded-full text-[10px] uppercase tracking-wider ${row.mercado === "LIVRE" ? "bg-blue-100 text-blue-700" : "bg-slate-100 text-slate-600"}`}
                                        >
                                          {row.mercado}
                                        </span>
                                      </td>
                                      <td className="px-8 py-5 text-sm font-medium text-slate-600">
                                        {row.modalidadeTarifaria || "-"}
                                      </td>
                                      {dashboardSubTab === "operacionais" ? (
                                        <>
                                          {operationalSubTab === "consumo" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.consumoPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kWh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.consumoForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kWh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">
                                                {(
                                                  row.consumoPonta +
                                                  row.consumoForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}{" "}
                                                kWh
                                              </td>
                                            </>
                                          )}
                                          {operationalSubTab ===
                                            "ultrapassagem" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.demandaContratadaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.demandaContratadaForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.ultrapassagemPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.ultrapassagemForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-red-600">
                                                {(
                                                  row.ultrapassagemPonta +
                                                  row.ultrapassagemForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}{" "}
                                                kW
                                              </td>
                                            </>
                                          )}
                                          {operationalSubTab ===
                                            "subutilizacao" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.demandaContratadaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.demandaContratadaForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.demandaMedidaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.demandaMedidaForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kW
                                              </td>
                                              <td
                                                className={`px-8 py-5 text-sm font-bold text-right ${row.demandaMedidaPonta < row.demandaContratadaPonta * 0.8 ? "text-orange-600" : "text-slate-600"}`}
                                              >
                                                {(
                                                  (row.demandaMedidaPonta /
                                                    (row.demandaContratadaPonta ||
                                                      1)) *
                                                  100
                                                ).toFixed(2)}
                                                %
                                              </td>
                                            </>
                                          )}
                                          {operationalSubTab === "reativa" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.reativaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kVArh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.reativaForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kVArh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-purple-600">
                                                {(
                                                  row.reativaPonta +
                                                  row.reativaForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}{" "}
                                                kVArh
                                              </td>
                                            </>
                                          )}
                                          {operationalSubTab === "solar" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {(
                                                  row.consumoPonta +
                                                  row.consumoForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}{" "}
                                                kWh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.solarInjetadaOUC.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kWh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {row.solarInjetadaMUC.toLocaleString(
                                                  "pt-BR",
                                                  {
                                                    minimumFractionDigits: 2,
                                                    maximumFractionDigits: 2,
                                                  },
                                                )}{" "}
                                                kWh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-green-600">
                                                {(
                                                  row.solarInjetadaOUC +
                                                  row.solarInjetadaMUC -
                                                  (row.consumoPonta +
                                                    row.consumoForaPonta)
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}{" "}
                                                kWh
                                              </td>
                                            </>
                                          )}
                                        </>
                                      ) : (
                                        <>
                                          {financialSubTab === "despesas" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">
                                                R${" "}
                                                {row.valorTotal.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-8 py-5 text-sm text-slate-600">
                                                {row.cidade}
                                              </td>
                                            </>
                                          )}
                                          {financialSubTab ===
                                            "multa_ultrapassagem" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                R${" "}
                                                {row.valorUltrapassagemPonta.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                R${" "}
                                                {row.valorUltrapassagemForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-red-600">
                                                R${" "}
                                                {(
                                                  row.valorUltrapassagemPonta +
                                                  row.valorUltrapassagemForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                })}
                                              </td>
                                            </>
                                          )}
                                          {financialSubTab ===
                                            "multa_reativa" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                R${" "}
                                                {row.valorReativaPonta.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                R${" "}
                                                {row.valorReativaForaPonta.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-purple-600">
                                                R${" "}
                                                {(
                                                  row.valorReativaPonta +
                                                  row.valorReativaForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                })}
                                              </td>
                                            </>
                                          )}
                                          {financialSubTab ===
                                            "tarifa_media" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                R${" "}
                                                {row.valorTotal.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-mono text-right text-slate-600">
                                                {(
                                                  row.consumoPonta +
                                                  row.consumoForaPonta
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}{" "}
                                                kWh
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">
                                                R${" "}
                                                {(
                                                  row.valorTotal /
                                                  (row.consumoPonta +
                                                    row.consumoForaPonta || 1)
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                  maximumFractionDigits: 2,
                                                })}
                                              </td>
                                            </>
                                          )}
                                          {financialSubTab ===
                                            "energia_solar" && (
                                            <>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-primary">
                                                R${" "}
                                                {(
                                                  Math.abs(
                                                    row.valorSolarOUC +
                                                      row.valorSolarMUC,
                                                  ) +
                                                  (row.valorTotal -
                                                    row.cip -
                                                    row.outrosEncargos)
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                })}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-green-600">
                                                R${" "}
                                                {Math.abs(
                                                  row.valorSolarOUC +
                                                    row.valorSolarMUC,
                                                ).toLocaleString("pt-BR", {
                                                  minimumFractionDigits: 2,
                                                })}
                                              </td>
                                              <td className="px-8 py-5 text-sm font-bold text-right text-sanesul-secondary">
                                                R${" "}
                                                {row.valorTotal.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                            </>
                                          )}
                                        </>
                                      )}
                                    </tr>
                                  ))}
                                  {sortedDashboardData.length > 100 && (
                                    <tr>
                                      <td
                                        colSpan={13}
                                        className="px-8 py-6 text-center text-sm text-slate-500 italic bg-slate-50 border-t border-sanesul-primary/5"
                                      >
                                        Exibindo os primeiros 100 de{" "}
                                        {sortedDashboardData.length} resultados.
                                        Use a pesquisa para ver mais.
                                      </td>
                                    </tr>
                                  )}
                                </>
                              );
                            })()}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>
        ) : activeTab === "analises" ? (
          <div className="space-y-8">
            <div className="bg-white p-10 rounded-3xl border border-sanesul-primary/10 shadow-xl">
              <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-10">
                <div>
                  <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                    Análises de Dados
                  </h2>
                  <p className="text-sanesul-muted">
                    Analise ultrapassagens e subutilização de demanda com base
                    nas faturas processadas.
                  </p>
                </div>
                <div className="flex items-center gap-4">
                  <button
                    onClick={runAnalysis}
                    disabled={
                      bills.filter((b) => b.status === "completed").length === 0
                    }
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
                        onClick={() => {
                          setAnalysisResults(null);
                          setAnalysisData([]);
                        }}
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
                    <FileSpreadsheet
                      size={40}
                      className="text-sanesul-primary/40"
                    />
                  </div>
                  <h3 className="text-xl font-bold text-sanesul-primary mb-2">
                    Nenhuma análise gerada
                  </h3>
                  <p className="text-sanesul-muted max-w-md mx-auto mb-8">
                    Clique em "Gerar Análise" para utilizar os dados das faturas
                    processadas e calcular a demanda ideal.
                  </p>
                  <div className="flex justify-center gap-4">
                    <div className="p-4 rounded-xl bg-white border border-slate-100 text-left max-w-xs">
                      <div className="flex items-center gap-2 text-sanesul-primary font-bold text-xs mb-2">
                        <TrendingUp size={14} />
                        Ultrapassagem
                      </div>
                      <p className="text-[10px] text-sanesul-muted">
                        Calculamos a demanda otimizada para eliminar multas de
                        ultrapassagem (Resolução 1000 ANEEL).
                      </p>
                    </div>
                    <div className="p-4 rounded-xl bg-white border border-slate-100 text-left max-w-xs">
                      <div className="flex items-center gap-2 text-sanesul-primary font-bold text-xs mb-2">
                        <Zap size={14} />
                        Subutilização
                      </div>
                      <p className="text-[10px] text-sanesul-muted">
                        Identificamos a demanda ideal para evitar pagamentos por
                        potência não utilizada.
                      </p>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="space-y-10">
                  {/* Summary Cards */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-6">
                    <div className="bg-white p-8 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <TrendingUp size={80} className="text-green-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                        Economia Positiva
                      </p>
                      <p className="text-3xl font-display font-bold text-green-600">
                        R${" "}
                        {analysisResults
                          .filter((r: any) => r.economy > 0)
                          .reduce(
                            (acc: any, curr: any) => acc + curr.economy,
                            0,
                          )
                          .toLocaleString("pt-BR", {
                            minimumFractionDigits: 2,
                          })}
                      </p>
                    </div>

                    <div className="bg-white p-8 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <TrendingDown size={80} className="text-red-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                        Economia Negativa
                      </p>
                      <p className="text-3xl font-display font-bold text-red-600">
                        R${" "}
                        {analysisResults
                          .filter((r: any) => r.economy < 0)
                          .reduce(
                            (acc: any, curr: any) => acc + curr.economy,
                            0,
                          )
                          .toLocaleString("pt-BR", {
                            minimumFractionDigits: 2,
                          })}
                      </p>
                    </div>

                    <div className="bg-white p-8 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <AlertCircle size={80} className="text-red-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                        Ultrapassagens
                      </p>
                      <p className="text-3xl font-display font-bold text-red-600">
                        {analysisResults.filter((r: any) => r.isOverrun).length}
                      </p>
                    </div>
                    <div className="bg-white p-8 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <TrendingDown size={80} className="text-orange-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                        Subutilizações
                      </p>
                      <p className="text-3xl font-display font-bold text-orange-600">
                        {analysisResults.filter((r: any) => r.isSub).length}
                      </p>
                    </div>
                    <div className="bg-white p-8 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <Zap size={80} className="text-green-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                        Eficiência
                      </p>
                      <p className="text-3xl font-display font-bold text-green-600">
                        {Math.round(
                          (analysisResults.filter(
                            (r: any) => !r.isOverrun && !r.isSub,
                          ).length /
                            analysisResults.length) *
                            100,
                        )}
                        %
                      </p>
                    </div>
                  </div>

                  {/* Results Table */}
                  <div className="overflow-hidden rounded-2xl border border-sanesul-primary/5 shadow-sm">
                    <div className="overflow-y-auto max-h-[520px] scrollbar-thin scrollbar-thumb-sanesul-primary/20 scrollbar-track-transparent">
                      <table className="w-full border-collapse bg-white text-left relative">
                        <thead className="sticky top-0 z-10 bg-slate-50 shadow-sm border-b border-sanesul-primary/10">
                          <tr>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary w-10"></th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary">
                              UC
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary">
                              Gerência
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary">
                              Cidade
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary text-right">
                              Contratada (P/FP)
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 text-right">
                              Demanda Ideal
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary text-right">
                              Gasto Real
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 text-right">
                              Economia
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary text-right">
                              Meses Analisados
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary text-center w-16">
                              Ações
                            </th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                        {Object.values(
                          analysisResults.reduce((acc: any, curr: any) => {
                            if (!acc[curr.uc]) {
                              acc[curr.uc] = {
                                uc: curr.uc,
                                city: curr.city || "",
                                months: [],
                                optimizedPonta: curr.optimizedPonta,
                                optimizedForaPonta: curr.optimizedForaPonta,
                                dcp: curr.dcp,
                                dcfp: curr.dcfp,
                                totalEconomy: 0,
                                totalCurrent: 0,
                                totalOptimized: 0,
                              };
                            }
                            acc[curr.uc].months.push(curr);
                            acc[curr.uc].totalEconomy += curr.economy;
                            acc[curr.uc].totalCurrent += curr.currentTotal;
                            acc[curr.uc].totalOptimized += curr.optimizedTotal;
                            return acc;
                          }, {}),
                        ).map((group: any, idx: number) => (
                          <React.Fragment key={idx}>
                            <tr
                              className={`hover:bg-slate-50/50 transition-colors cursor-pointer ${expandedAnalysisUCs.has(group.uc) ? "bg-slate-50/80" : ""}`}
                              onClick={() =>
                                toggleAnalysisUCExpansion(group.uc)
                              }
                            >
                              <td className="px-6 py-4 text-center">
                                <ChevronRight
                                  size={16}
                                  className={`text-sanesul-muted transition-transform ${expandedAnalysisUCs.has(group.uc) ? "rotate-90" : ""}`}
                                />
                              </td>
                              <td className="px-6 py-4">
                                <div className="font-bold text-sanesul-primary text-xs">
                                  {group.uc}
                                </div>
                              </td>
                              <td className="px-6 py-4">
                                <div className="font-bold text-slate-800 text-xs">
                                  {getGerencia(group.uc)}
                                </div>
                              </td>
                              <td className="px-6 py-4">
                                <div className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                                  {group.city || "---"}
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-mono text-slate-600">
                                  {group.dcp > 0 ? group.dcp.toFixed(2) : "-"} /{" "}
                                  {group.dcfp.toFixed(2)} kW
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-mono font-bold text-green-600">
                                  {group.optimizedPonta > 0
                                    ? group.optimizedPonta.toFixed(2)
                                    : "-"}{" "}
                                  / {group.optimizedForaPonta.toFixed(2)} kW
                                </div>
                                <div className="text-[9px] text-green-500 uppercase font-bold tracking-tighter">
                                  Ideal Fixo (1 Ano)
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-mono font-bold text-sanesul-primary">
                                  R${" "}
                                  {group.totalCurrent.toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div
                                  className={`text-xs font-mono font-bold ${group.totalEconomy >= 0 ? "text-green-600" : "text-red-600"}`}
                                >
                                  R${" "}
                                  {group.totalEconomy.toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                                </div>
                              </td>
                              <td className="px-6 py-4 text-right">
                                <div className="text-xs font-bold text-slate-500">
                                  {group.months.length} meses
                                </div>
                              </td>
                              <td className="px-6 py-4 text-center">
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    showConfirm(
                                      "Excluir UC do Sistema",
                                      `Tem certeza que deseja excluir TODAS as faturas da UC ${group.uc}? Esta ação não pode ser desfeita.`,
                                      () => removeUCBills(group.uc),
                                      "danger"
                                    );
                                  }}
                                  className="p-2 text-red-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors"
                                  title="Excluir UC"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </td>
                            </tr>
                            {expandedAnalysisUCs.has(group.uc) && (
                              <tr>
                                <td
                                  colSpan={10}
                                  className="px-10 py-4 bg-slate-50/30"
                                >
                                  <div className="overflow-hidden rounded-xl border border-slate-200 shadow-inner">
                                    <table className="w-full text-left border-collapse bg-white">
                                      <thead>
                                        <tr className="bg-slate-100/50">
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-slate-500">
                                            Mês/Ano
                                          </th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-slate-500 text-right">
                                            Medida (P/FP)
                                          </th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-red-600 text-right">
                                            Ultrapassagem
                                          </th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-orange-600 text-right">
                                            Subutilização
                                          </th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-sanesul-primary text-right">
                                            Gasto Real
                                          </th>
                                          <th className="px-4 py-2 text-[9px] font-bold uppercase tracking-widest text-green-600 text-right">
                                            Economia
                                          </th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-slate-100">
                                        {group.months.map(
                                          (month: any, mIdx: number) => (
                                            <tr
                                              key={mIdx}
                                              className="hover:bg-slate-50/30 transition-colors"
                                            >
                                              <td className="px-4 py-2 text-xs font-bold text-slate-700">
                                                {month.mes}{" "}
                                                {month.ano
                                                  ? `/ ${month.ano}`
                                                  : ""}
                                              </td>
                                              <td className="px-4 py-2 text-right text-xs font-mono font-bold text-sanesul-primary">
                                                {month.dcp > 0
                                                  ? month.dmp.toFixed(2)
                                                  : "-"}{" "}
                                                / {month.dmfp.toFixed(2)} kW
                                              </td>
                                              <td className="px-4 py-2 text-right">
                                                {month.isOverrun ? (
                                                  <div className="flex flex-col items-end gap-0.5">
                                                    {month.overrunPonta > 0 && (
                                                      <span className="text-[9px] font-bold text-red-600">
                                                        P: +
                                                        {month.overrunPonta.toFixed(
                                                          2,
                                                        )}
                                                      </span>
                                                    )}
                                                    {month.overrunForaPonta >
                                                      0 && (
                                                      <span className="text-[9px] font-bold text-red-600">
                                                        FP: +
                                                        {month.overrunForaPonta.toFixed(
                                                          2,
                                                        )}
                                                      </span>
                                                    )}
                                                  </div>
                                                ) : (
                                                  <span className="text-slate-300 text-[9px]">
                                                    -
                                                  </span>
                                                )}
                                              </td>
                                              <td className="px-4 py-2 text-right">
                                                {month.isSub ? (
                                                  <div className="flex flex-col items-end gap-0.5">
                                                    {month.subPonta > 0 && (
                                                      <span className="text-[9px] font-bold text-orange-600">
                                                        P: -
                                                        {month.subPonta.toFixed(
                                                          2,
                                                        )}
                                                      </span>
                                                    )}
                                                    {month.subForaPonta > 0 && (
                                                      <span className="text-[9px] font-bold text-orange-600">
                                                        FP: -
                                                        {month.subForaPonta.toFixed(
                                                          2,
                                                        )}
                                                      </span>
                                                    )}
                                                  </div>
                                                ) : (
                                                  <span className="text-slate-300 text-[9px]">
                                                    -
                                                  </span>
                                                )}
                                              </td>
                                              <td className="px-4 py-2 text-right text-xs font-mono font-bold text-sanesul-primary">
                                                R${" "}
                                                {month.currentTotal.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                              <td className="px-4 py-2 text-right text-xs font-mono font-bold text-green-600">
                                                R${" "}
                                                {month.economy.toLocaleString(
                                                  "pt-BR",
                                                  { minimumFractionDigits: 2 },
                                                )}
                                              </td>
                                            </tr>
                                          ),
                                        )}
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
                </div>
              )}
            </div>
          </div>
        ) : activeTab === "monitoramento" ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                  Monitoramento Demanda
                </h2>
                <p className="text-sanesul-muted">
                  Acompanhamento detalhado de gastos e economia com demanda por
                  cidade e UC.
                </p>
              </div>
              <div className="flex items-center gap-4">
                {monitoringResults && (
                  <>
                    <button
                      onClick={exportMonitoramentoExcel}
                      className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg shadow-green-600/20"
                    >
                      <Download size={16} />
                      Exportar Prejuízo
                    </button>
                    <button
                      onClick={exportStableContractsExcel}
                      className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg shadow-green-600/20"
                    >
                      <Download size={16} />
                      Exportar Sem Alteração
                    </button>
                  </>
                )}
                
                {monitoramentoMeses.length > 0 && (
                  <select
                    value={selectedMonitoramentoMes}
                    onChange={(e) => setSelectedMonitoramentoMes(e.target.value)}
                    className="px-4 py-3 bg-white border border-sanesul-primary/20 rounded-xl text-sanesul-primary text-xs font-bold uppercase tracking-wider focus:outline-none focus:ring-2 focus:ring-sanesul-primary shadow-sm"
                  >
                    <option value="Todos">Todos os meses</option>
                    {monitoramentoMeses.map((mes) => (
                      <option key={mes} value={mes}>{mes}</option>
                    ))}
                  </select>
                )}

                <button
                  onClick={runMonitoringAnalysis}
                  disabled={
                    bills.filter((b) => b.status === "completed").length === 0
                  }
                  className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <BarChart3 size={16} />
                  Atualizar Monitoramento
                </button>
              </div>
            </div>

            {!monitoringResults ? (
              <div className="border-2 border-dashed border-sanesul-primary/10 rounded-[40px] p-20 text-center bg-white/50 backdrop-blur-sm">
                <div className="w-24 h-24 bg-sanesul-primary/5 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-inner">
                  <DollarSign size={48} className="text-sanesul-primary/40" />
                </div>
                <h3 className="text-2xl font-display font-bold text-sanesul-primary mb-4">
                  Pronto para analisar
                </h3>
                <p className="text-sanesul-muted max-w-md mx-auto mb-10 text-lg">
                  Clique no botão acima para processar os indicadores de despesa
                  e economia baseados nas faturas extraídas.
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
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                      Faturas Analisadas
                    </p>
                    <p className="text-4xl font-display font-bold text-sanesul-primary">
                      {bills.filter((b) => b.status === "completed").length}{" "}
                      <span className="text-base font-sans font-medium opacity-40">
                        Arquivos
                      </span>
                    </p>
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      {filteredMonitoringResults.generalTotalEconomy >= 0 ? (
                        <TrendingUp size={80} className="text-green-600" />
                      ) : (
                        <TrendingDown size={80} className="text-red-600" />
                      )}
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                      Economia Geral Total
                    </p>
                    <p
                      className={`text-4xl font-display font-bold ${filteredMonitoringResults.generalTotalEconomy >= 0 ? "text-green-600" : "text-red-600"}`}
                    >
                      R${" "}
                      {filteredMonitoringResults.generalTotalEconomy.toLocaleString(
                        "pt-BR",
                        { minimumFractionDigits: 2 },
                      )}
                    </p>
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      <DollarSign size={80} className="text-sanesul-primary" />
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                      Despesa Atual Total
                    </p>
                    <p className="text-4xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {filteredMonitoringResults.generalTotalCurrent.toLocaleString(
                        "pt-BR",
                        { minimumFractionDigits: 2 },
                      )}
                    </p>
                  </div>

                  <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                    <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                      <Zap size={80} className="text-sanesul-secondary" />
                    </div>
                    <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                      Potencial de Redução
                    </p>
                    <p className="text-4xl font-display font-bold text-sanesul-secondary">
                      {(
                        (filteredMonitoringResults.generalTotalEconomy /
                          filteredMonitoringResults.generalTotalCurrent) *
                        100
                      ).toFixed(1)}
                      %
                    </p>
                  </div>
                </div>

                {monitoringResults.timelineData && monitoringResults.timelineData.length > 0 && (
                  <div className="grid grid-cols-1 xl:grid-cols-3 gap-8">
                    {/* Left Column (Span 2): Gráfico de Evolução Mensal */}
                    <div className="xl:col-span-2 bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl flex flex-col justify-between">
                      <div className="flex items-center justify-between mb-6">
                        <h3 className="text-xl font-display font-bold text-sanesul-primary">
                          Evolução Mensal Pós Alteração
                        </h3>
                        <div className="px-4 py-2 bg-sanesul-primary/5 rounded-full">
                          <span className="text-xs font-bold text-sanesul-primary uppercase tracking-wider">
                            Economia x Valor Simulado
                          </span>
                        </div>
                      </div>
                      <div className="h-96 w-full mt-2">
                        <ResponsiveContainer width="100%" height="100%">
                          <ComposedChart
                            data={monitoringResults.timelineData}
                            margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
                          >
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis 
                              dataKey="monthYear" 
                              interval={0}
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fontSize: 11, fill: "#64748b", fontWeight: 600 }} 
                              dy={10} 
                              tickFormatter={(val: string) => {
                                if (!val) return "";
                                const parts = val.split("/");
                                if (parts.length === 2) {
                                  const shortM = parts[0].substring(0, 3);
                                  const shortY = parts[1].length === 4 ? parts[1].substring(2) : parts[1];
                                  return `${shortM}/${shortY}`;
                                }
                                return val;
                              }}
                            />
                            <YAxis 
                              yAxisId="left"
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fontSize: 12, fill: "#64748b" }}
                              tickFormatter={(value) => `R$ ${(value / 1000).toFixed(0)}k`}
                            />
                            <YAxis 
                              yAxisId="right"
                              orientation="right"
                              axisLine={false} 
                              tickLine={false} 
                              tick={{ fontSize: 12, fill: "#64748b" }}
                              tickFormatter={(value) => `R$ ${(value / 1000).toFixed(0)}k`}
                            />
                            <Tooltip
                              contentStyle={{ borderRadius: "16px", border: "none", boxShadow: "0 10px 25px -5px rgba(0, 0, 0, 0.05)" }}
                              formatter={(value: number) => `R$ ${value.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`}
                            />
                            <Legend wrapperStyle={{ paddingTop: "20px" }} />
                            <Bar 
                              yAxisId="left"
                              dataKey="referenceTotal" 
                              name="Valor Sem Alteração (Simulado)" 
                              fill="#cbd5e1" 
                              radius={[6, 6, 0, 0]} 
                              barSize={36}
                            />
                            <Bar 
                              yAxisId="left"
                              dataKey="currentTotal" 
                              name="Valor Real Pago" 
                              fill="#0d2551" 
                              radius={[6, 6, 0, 0]} 
                              barSize={36}
                            />
                            <Line 
                              yAxisId="right"
                              type="monotone" 
                              dataKey="economy" 
                              name="Valor Economizado" 
                              stroke="#16a34a" 
                              strokeWidth={3}
                              dot={{ r: 4, strokeWidth: 2, fill: "#fff" }}
                              activeDot={{ r: 6, strokeWidth: 0 }}
                            />
                          </ComposedChart>
                        </ResponsiveContainer>
                      </div>
                    </div>

                    {/* Right Column (Span 1): Card Somatório das Economias Mês a Mês */}
                    <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl flex flex-col justify-between">
                      <div>
                        <div className="flex items-center justify-between mb-6 pb-4 border-b border-slate-100">
                          <div className="flex items-center gap-3">
                            <div className="w-10 h-10 rounded-xl bg-green-50 text-green-600 flex items-center justify-center">
                              <TrendingUp size={20} />
                            </div>
                            <div>
                              <h3 className="text-lg font-display font-bold text-sanesul-primary">
                                Somatório Mês a Mês
                              </h3>
                              <p className="text-[11px] font-bold text-sanesul-muted uppercase tracking-wider">
                                Economia Realizada Consolidada
                              </p>
                            </div>
                          </div>
                          <div className="px-3 py-1 bg-green-50 text-green-700 rounded-full text-[10px] font-black uppercase tracking-wider">
                            {monitoringResults.timelineData.length} {monitoringResults.timelineData.length === 1 ? "Mês" : "Meses"}
                          </div>
                        </div>

                        {/* Totalizador Geral */}
                        <div className="bg-gradient-to-br from-green-50/80 to-emerald-50/40 p-5 rounded-2xl border border-green-100/80 mb-6">
                          <span className="text-[10px] font-bold uppercase tracking-widest text-green-700 block mb-1">
                            Economia Total Acumulada
                          </span>
                          <p className="text-3xl font-display font-extrabold text-green-700">
                            R${" "}
                            {monitoringResults.timelineData
                              .reduce((sum: number, item: any) => sum + item.economy, 0)
                              .toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                          </p>
                          <div className="flex items-center justify-between mt-2 pt-2 border-t border-green-200/50 text-xs">
                            <span className="text-slate-500 font-medium">Média mensal:</span>
                            <span className="font-bold font-mono text-green-800">
                              R${" "}
                              {(
                                monitoringResults.timelineData.reduce((sum: number, item: any) => sum + item.economy, 0) /
                                (monitoringResults.timelineData.length || 1)
                              ).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              /mês
                            </span>
                          </div>
                        </div>

                        {/* Lista Mês a Mês */}
                        <div className="space-y-3 max-h-[280px] overflow-y-auto pr-1">
                          {(() => {
                            let runningSum = 0;
                            // Calculate cumulative sum chronologically
                            const itemsWithAccumulated = monitoringResults.timelineData.map((item: any) => {
                              runningSum += item.economy;
                              return {
                                ...item,
                                accumulated: runningSum,
                              };
                            });
                            // Reverse so current/most recent month is first
                            const reversedItems = [...itemsWithAccumulated].reverse();

                            return reversedItems.map((item: any, idx: number) => {
                              const isPositive = item.economy >= 0;
                              return (
                                <div
                                  key={idx}
                                  className="p-3.5 bg-slate-50/80 hover:bg-slate-50 rounded-xl border border-slate-100 flex items-center justify-between transition-colors"
                                >
                                  <div className="space-y-0.5">
                                    <div className="flex items-center gap-2">
                                      <span className={`w-2 h-2 rounded-full ${isPositive ? "bg-green-500" : "bg-red-500"}`} />
                                      <span className="text-xs font-bold text-sanesul-primary">
                                        {item.monthYear}
                                      </span>
                                      {idx === 0 && (
                                        <span className="px-2 py-0.5 bg-blue-50 text-blue-700 rounded-md text-[9px] font-extrabold uppercase tracking-wider">
                                          Atual / Recente
                                        </span>
                                      )}
                                    </div>
                                    <span className="text-[10px] font-mono text-slate-400 block pl-4">
                                      Acumulado: R${" "}
                                      {item.accumulated.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                    </span>
                                  </div>

                                  <div className="text-right">
                                    <span
                                      className={`text-sm font-display font-extrabold ${
                                        isPositive ? "text-green-600" : "text-red-600"
                                      }`}
                                    >
                                      {isPositive ? "+" : ""}R${" "}
                                      {item.economy.toLocaleString("pt-BR", {
                                        minimumFractionDigits: 2,
                                        maximumFractionDigits: 2,
                                      })}
                                    </span>
                                    <span className="text-[10px] text-slate-400 block font-mono">
                                      Ref: R${" "}
                                      {(item.referenceTotal / 1000).toFixed(1)}k
                                    </span>
                                  </div>
                                </div>
                              );
                            });
                          })()}
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {/* Charts Section - Replaced with Text Summary */}
                <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl">
                  <div className="flex items-center justify-between mb-8">
                    <h3 className="text-xl font-display font-bold text-sanesul-primary">
                      Resumo de Economia por Cidade
                    </h3>
                    <div className="px-4 py-2 bg-sanesul-primary/5 rounded-full">
                      <span className="text-xs font-bold text-sanesul-primary uppercase tracking-wider">
                        Top Impactos Financeiros
                      </span>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-12">
                    {/* Top Economies */}
                    <div className="space-y-6">
                      <div className="flex items-center gap-3 pb-4 border-b border-green-100">
                        <div className="w-10 h-10 rounded-xl bg-green-50 flex items-center justify-center">
                          <TrendingUp className="text-green-600" size={20} />
                        </div>
                        <h4 className="font-display font-bold text-green-700 uppercase tracking-tight">
                          Principais Economias
                        </h4>
                      </div>
                      <div className="space-y-2">
                        {[...filteredMonitoringResults.cityData]
                          .filter((c) => c.positiveEconomy > 0)
                          .sort((a, b) => b.positiveEconomy - a.positiveEconomy)
                          .slice(0, 10)
                          .map((city, idx) => (
                            <div
                              key={idx}
                              className="flex flex-col rounded-2xl bg-green-50/50 border border-green-100/50 hover:bg-green-50 transition-colors overflow-hidden"
                            >
                              <div
                                className="flex items-center justify-between p-3 cursor-pointer"
                                onClick={() =>
                                  toggleSummaryCity(`pos-${city.city}`)
                                }
                              >
                                <div className="flex items-center gap-3">
                                  <span className="text-xs font-bold text-green-600/50 w-4">
                                    {String(idx + 1).padStart(2, "0")}
                                  </span>
                                  <span className="font-bold text-sanesul-primary text-sm">
                                    {city.city}
                                  </span>
                                </div>
                                <div className="flex items-center gap-3">
                                  <span className="font-mono font-bold text-green-600 text-sm">
                                    + R${" "}
                                    {city.positiveEconomy.toLocaleString(
                                      "pt-BR",
                                      { minimumFractionDigits: 2 },
                                    )}
                                  </span>
                                  <ChevronDown
                                    className={`w-4 h-4 text-green-600/50 transition-transform ${expandedSummaryCities.has(`pos-${city.city}`) ? "rotate-180" : ""}`}
                                  />
                                </div>
                              </div>
                              {expandedSummaryCities.has(`pos-${city.city}`) &&
                                city.positiveUcs && (
                                  <div className="px-3 pb-3 pt-1 border-t border-green-100/30 bg-green-50/30">
                                    <div className="space-y-1 mt-2">
                                      {city.positiveUcs.map(
                                        (u: any, i: number) => (
                                          <div
                                            key={i}
                                            className="flex justify-between items-center text-xs"
                                          >
                                            <span className="text-sanesul-muted font-medium">
                                              UC {u.uc}{" "}
                                              <span className="opacity-60 ml-2">
                                                ({getGerencia(u.uc)})
                                              </span>
                                            </span>
                                            <span className="font-mono text-green-600/80">
                                              + R${" "}
                                              {Math.abs(
                                                u.economy,
                                              ).toLocaleString("pt-BR", {
                                                minimumFractionDigits: 2,
                                              })}
                                            </span>
                                          </div>
                                        ),
                                      )}
                                    </div>
                                  </div>
                                )}
                            </div>
                          ))}
                        {filteredMonitoringResults.cityData.filter(
                          (c) => c.positiveEconomy > 0,
                        ).length === 0 && (
                          <p className="text-sm text-sanesul-muted italic p-4">
                            Nenhuma economia significativa identificada.
                          </p>
                        )}
                      </div>
                    </div>

                    {/* Top Losses */}
                    <div className="space-y-6">
                      <div className="flex items-center gap-3 pb-4 border-b border-red-100">
                        <div className="w-10 h-10 rounded-xl bg-red-50 flex items-center justify-center">
                          <TrendingDown className="text-red-600" size={20} />
                        </div>
                        <h4 className="font-display font-bold text-red-700 uppercase tracking-tight">
                          Principais Prejuízos
                        </h4>
                      </div>
                      <div className="space-y-2">
                        {[...filteredMonitoringResults.cityData]
                          .filter((c) => c.negativeEconomy < 0)
                          .sort((a, b) => a.negativeEconomy - b.negativeEconomy)
                          .slice(0, 10)
                          .map((city, idx) => (
                            <div
                              key={idx}
                              className="flex flex-col rounded-2xl bg-red-50/50 border border-red-100/50 hover:bg-red-50 transition-colors overflow-hidden"
                            >
                              <div
                                className="flex items-center justify-between p-3 cursor-pointer"
                                onClick={() =>
                                  toggleSummaryCity(`neg-${city.city}`)
                                }
                              >
                                <div className="flex items-center gap-3">
                                  <span className="text-xs font-bold text-red-600/50 w-4">
                                    {String(idx + 1).padStart(2, "0")}
                                  </span>
                                  <span className="font-bold text-sanesul-primary text-sm">
                                    {city.city}
                                  </span>
                                </div>
                                <div className="flex items-center gap-3">
                                  <span className="font-mono font-bold text-red-600 text-sm">
                                    - R${" "}
                                    {Math.abs(
                                      city.negativeEconomy,
                                    ).toLocaleString("pt-BR", {
                                      minimumFractionDigits: 2,
                                    })}
                                  </span>
                                  <ChevronDown
                                    className={`w-4 h-4 text-red-600/50 transition-transform ${expandedSummaryCities.has(`neg-${city.city}`) ? "rotate-180" : ""}`}
                                  />
                                </div>
                              </div>
                              {expandedSummaryCities.has(`neg-${city.city}`) &&
                                city.negativeUcs && (
                                  <div className="px-3 pb-3 pt-1 border-t border-red-100/30 bg-red-50/30">
                                    <div className="space-y-1 mt-2">
                                      {city.negativeUcs.map(
                                        (u: any, i: number) => (
                                          <div
                                            key={i}
                                            className="flex justify-between items-center text-xs"
                                          >
                                            <span className="text-sanesul-muted font-medium">
                                              UC {u.uc}{" "}
                                              <span className="opacity-60 ml-2">
                                                ({getGerencia(u.uc)})
                                              </span>
                                            </span>
                                            <span className="font-mono text-red-600/80">
                                              - R${" "}
                                              {Math.abs(
                                                u.economy,
                                              ).toLocaleString("pt-BR", {
                                                minimumFractionDigits: 2,
                                              })}
                                            </span>
                                          </div>
                                        ),
                                      )}
                                    </div>
                                  </div>
                                )}
                            </div>
                          ))}
                        {filteredMonitoringResults.cityData.filter(
                          (c) => c.negativeEconomy < 0,
                        ).length === 0 && (
                          <p className="text-sm text-sanesul-muted italic p-4">
                            Nenhum prejuízo significativo identificado.
                          </p>
                        )}
                      </div>
                    </div>
                  </div>
                </div>

                {/* City Breakdown */}
                <div className="space-y-8">
                  <div className="flex items-center justify-between">
                    <h3 className="text-2xl font-display font-bold text-sanesul-primary border-l-4 border-sanesul-secondary pl-4">
                      Detalhamento por Cidade
                    </h3>
                  </div>
                  {/* Changed UCs Group */}
                  {filteredMonitoringResults.changedUCs.length > 0 && (
                    <div className="bg-white rounded-[40px] border border-sanesul-primary/10 shadow-2xl overflow-hidden mb-8">
                      <div className="bg-slate-50/80 px-10 py-8 border-b border-sanesul-primary/5">
                        <h3 className="text-2xl font-display font-bold text-sanesul-primary border-l-4 border-yellow-500 pl-4">
                          Unidades com Alteração de Demanda
                        </h3>
                        <p className="text-xs font-bold text-sanesul-muted uppercase tracking-widest mt-2 pl-5">
                          {filteredMonitoringResults.changedUCs.length} Unidades
                          Encontradas
                        </p>
                      </div>

                      <div className="p-10">
                        <div className="overflow-x-auto overflow-y-auto max-h-[520px]">
                          <table className="w-full text-left border-collapse">
                            <thead className="sticky top-0 z-10 bg-white">
                              <tr>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] w-10"></th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">
                                  Unidade Consumidora
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">
                                  Gerência
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">
                                  Cidade
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">
                                  Contratada Atual (P/FP)
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">
                                  Demanda Ideal (P/FP)
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">
                                  Gasto Atual (Total)
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">
                                  Economia Acumulada
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">
                                  Status
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {filteredMonitoringResults.changedUCs.map(
                                (uc: any, uIdx: number) => (
                                  <React.Fragment key={uIdx}>
                                    <tr
                                      className={`group hover:bg-slate-50/50 transition-colors cursor-pointer ${expandedUCs.has(uc.uc) ? "bg-slate-50/80" : ""}`}
                                      onClick={() => toggleUCExpansion(uc.uc)}
                                    >
                                      <td className="py-6 text-center">
                                        <ChevronRight
                                          size={16}
                                          className={`text-sanesul-muted transition-transform ${expandedUCs.has(uc.uc) ? "rotate-90" : ""}`}
                                        />
                                      </td>
                                      <td className="py-6">
                                        <div className="flex items-center gap-3">
                                          <div className="w-10 h-10 rounded-xl bg-sanesul-primary/5 flex items-center justify-center font-bold text-sanesul-primary text-xs">
                                            {uIdx + 1}
                                          </div>
                                          <span className="font-bold text-sanesul-primary font-mono">
                                            {uc.uc}
                                          </span>
                                        </div>
                                      </td>
                                      <td className="py-6 text-xs font-bold text-slate-800">
                                        {getGerencia(uc.uc)}
                                      </td>
                                      <td className="py-6 text-xs font-bold text-slate-500 uppercase tracking-wider">
                                        {uc.city}
                                      </td>
                                      <td className="py-6 text-center">
                                        <span className="px-3 py-1.5 bg-slate-50 text-slate-700 rounded-lg text-xs font-bold font-mono">
                                          {uc.monthlyData[0]?.dcp} /{" "}
                                          {uc.monthlyData[0]?.dcfp} kW
                                        </span>
                                      </td>
                                      <td className="py-6 text-center">
                                        <span className="px-3 py-1.5 bg-green-50 text-green-700 rounded-lg text-xs font-bold font-mono">
                                          {uc.optPonta} / {uc.optForaPonta} kW
                                        </span>
                                      </td>
                                      <td className="py-6 text-right font-mono text-sm text-sanesul-muted">
                                        R${" "}
                                        {uc.totalCurrent.toLocaleString(
                                          "pt-BR",
                                          { minimumFractionDigits: 2 },
                                        )}
                                      </td>
                                      <td className="py-6 text-right font-bold text-green-600 text-lg">
                                        R${" "}
                                        {uc.totalEconomy.toLocaleString(
                                          "pt-BR",
                                          { minimumFractionDigits: 2 },
                                        )}
                                      </td>
                                      <td className="py-6 text-right">
                                        <div
                                          className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                                            uc.totalEconomy > 0
                                              ? "bg-green-100 text-green-700"
                                              : uc.totalEconomy < 0
                                                ? "bg-red-100 text-red-700"
                                                : "bg-slate-100 text-slate-500"
                                          }`}
                                        >
                                          {uc.totalEconomy > 0 ? (
                                            <TrendingUp size={12} />
                                          ) : uc.totalEconomy < 0 ? (
                                            <TrendingDown size={12} />
                                          ) : (
                                            <div className="w-3 h-3 rounded-full bg-slate-400" />
                                          )}
                                          {uc.totalEconomy > 0
                                            ? "Economia"
                                            : uc.totalEconomy < 0
                                              ? "Prejuízo"
                                              : "Neutro"}
                                        </div>
                                      </td>
                                    </tr>
                                    {expandedUCs.has(uc.uc) && (
                                      <tr>
                                        <td colSpan={9} className="px-10 py-0">
                                          <div className="bg-slate-50/50 rounded-2xl p-6 mb-6 border border-slate-100 animate-in fade-in slide-in-from-top-2 duration-300">
                                            <div className="flex items-center gap-2 mb-4">
                                              <Calendar
                                                size={14}
                                                className="text-sanesul-primary"
                                              />
                                              <h5 className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">
                                                Histórico de Alterações e
                                                Economia
                                              </h5>
                                            </div>
                                            <div className="grid grid-cols-7 gap-4 mb-2 px-4">
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest">
                                                Mês/Ano
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">
                                                Contratada (P/FP)
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">
                                                Medida (P/FP)
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">
                                                Gasto Real
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">
                                                Ref. Anterior
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">
                                                Economia Mensal
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">
                                                Economia Acumulada
                                              </div>
                                            </div>
                                            <div className="space-y-2">
                                              {uc.monthlyData.map(
                                                (month: any, mIdx: number) => (
                                                  <div
                                                    key={mIdx}
                                                    className={`grid grid-cols-7 gap-4 px-4 py-3 rounded-xl border transition-colors ${month.hasChanged ? "bg-yellow-50 border-yellow-200 ring-1 ring-yellow-300" : "bg-white border-slate-100 hover:border-sanesul-primary/20"}`}
                                                  >
                                                    <div className="flex flex-col">
                                                      <span className="text-xs font-bold text-slate-700">
                                                        {formatMonth(month.mes)}
                                                        /{month.ano}
                                                      </span>
                                                      {month.hasChanged && (
                                                        <span className="text-[9px] font-bold text-yellow-700 uppercase tracking-wider mt-0.5">
                                                          Alteração de Contrato
                                                        </span>
                                                      )}
                                                    </div>
                                                    <div className="text-xs font-mono text-center text-slate-500">
                                                      {month.dcp} / {month.dcfp}{" "}
                                                      kW
                                                    </div>
                                                    <div className="text-xs font-mono text-center text-slate-500">
                                                      {month.dmp} / {month.dmfp}{" "}
                                                      kW
                                                    </div>
                                                    <div className="text-xs font-mono text-right font-bold text-sanesul-primary">
                                                      R${" "}
                                                      {month.currentTotal.toLocaleString(
                                                        "pt-BR",
                                                        {
                                                          minimumFractionDigits: 2,
                                                        },
                                                      )}
                                                    </div>
                                                    <div className="text-xs font-mono text-right text-slate-400">
                                                      {month.referenceTotal > 0
                                                        ? `R$ ${month.referenceTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                                        : "-"}
                                                    </div>
                                                    <div
                                                      className={`text-xs font-mono text-right font-bold ${month.economy > 0 ? "text-green-600" : month.economy < 0 ? "text-red-500" : "text-slate-400"}`}
                                                    >
                                                      {month.referenceTotal > 0
                                                        ? `R$ ${month.economy.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                                        : "-"}
                                                    </div>
                                                    <div
                                                      className={`text-xs font-mono text-right font-extrabold ${month.accumulatedEconomy > 0 ? "text-green-700" : month.accumulatedEconomy < 0 ? "text-red-600" : "text-slate-400"}`}
                                                    >
                                                      {month.referenceTotal > 0
                                                        ? `R$ ${month.accumulatedEconomy.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                                                        : "-"}
                                                    </div>
                                                  </div>
                                                ),
                                              )}
                                            </div>
                                          </div>
                                        </td>
                                      </tr>
                                    )}
                                  </React.Fragment>
                                ),
                              )}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* Unchanged UCs Group */}
                  {filteredMonitoringResults.unchangedUCs.length > 0 && (
                    <div className="bg-white rounded-[40px] border border-sanesul-primary/10 shadow-2xl overflow-hidden">
                      <div className="bg-slate-50/80 px-10 py-8 border-b border-sanesul-primary/5 flex justify-between items-center">
                        <div>
                          <h3 className="text-2xl font-display font-bold text-sanesul-muted border-l-4 border-slate-300 pl-4">
                            Unidades sem Alteração (Contrato Estável)
                          </h3>
                          <p className="text-xs font-bold text-sanesul-muted uppercase tracking-widest mt-2 pl-5">
                            {filteredMonitoringResults.unchangedUCs.length} Unidades
                            Encontradas
                          </p>
                        </div>
                      </div>

                      <div className="p-10">
                        <div className="overflow-x-auto overflow-y-auto max-h-[520px]">
                          <table className="w-full text-left border-collapse opacity-80 hover:opacity-100 transition-opacity">
                            <thead className="sticky top-0 z-10 bg-white">
                              <tr>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] w-10"></th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">
                                  Unidade Consumidora
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">
                                  Gerência
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em]">
                                  Cidade
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">
                                  Contratada (P/FP)
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-center">
                                  Demanda Ideal (P/FP)
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">
                                  Gasto Atual (Total)
                                </th>
                                <th className="pb-6 text-[10px] font-bold text-sanesul-muted uppercase tracking-[0.2em] text-right">
                                  Status
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                              {filteredMonitoringResults.unchangedUCs.map(
                                (uc: any, uIdx: number) => (
                                  <React.Fragment key={uIdx}>
                                    <tr
                                      className={`group hover:bg-slate-50/50 transition-colors cursor-pointer ${expandedUCs.has(uc.uc) ? "bg-slate-50/80" : ""}`}
                                      onClick={() => toggleUCExpansion(uc.uc)}
                                    >
                                      <td className="py-6 text-center">
                                        <ChevronRight
                                          size={16}
                                          className={`text-sanesul-muted transition-transform ${expandedUCs.has(uc.uc) ? "rotate-90" : ""}`}
                                        />
                                      </td>
                                      <td className="py-6">
                                        <div className="flex items-center gap-3">
                                          <div className="w-10 h-10 rounded-xl bg-slate-100 flex items-center justify-center font-bold text-slate-500 text-xs">
                                            {uIdx + 1}
                                          </div>
                                          <span className="font-bold text-slate-600 font-mono">
                                            {uc.uc}
                                          </span>
                                        </div>
                                      </td>
                                      <td className="py-6 text-xs font-bold text-slate-800">
                                        {getGerencia(uc.uc)}
                                      </td>
                                      <td className="py-6 text-xs font-bold text-slate-500 uppercase tracking-wider">
                                        {uc.city}
                                      </td>
                                      <td className="py-6 text-center">
                                        <span className="px-3 py-1.5 bg-slate-50 text-slate-700 rounded-lg text-xs font-bold font-mono">
                                          {uc.monthlyData[0]?.dcp} /{" "}
                                          {uc.monthlyData[0]?.dcfp} kW
                                        </span>
                                      </td>
                                      <td className="py-6 text-center">
                                        <span className="px-3 py-1.5 bg-green-50 text-green-700 rounded-lg text-xs font-bold font-mono">
                                          {uc.optPonta} / {uc.optForaPonta} kW
                                        </span>
                                      </td>
                                      <td className="py-6 text-right font-mono text-sm text-sanesul-muted">
                                        R${" "}
                                        {uc.totalCurrent.toLocaleString(
                                          "pt-BR",
                                          { minimumFractionDigits: 2 },
                                        )}
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
                                        <td colSpan={9} className="px-10 py-0">
                                          <div className="bg-slate-50/50 rounded-2xl p-6 mb-6 border border-slate-100 animate-in fade-in slide-in-from-top-2 duration-300">
                                            <div className="flex items-center gap-2 mb-4">
                                              <Calendar
                                                size={14}
                                                className="text-sanesul-primary"
                                              />
                                              <h5 className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">
                                                Detalhamento Mensal
                                              </h5>
                                            </div>
                                            <div className="grid grid-cols-4 gap-4 mb-2 px-4">
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest">
                                                Mês/Ano
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">
                                                Contratada (P/FP)
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-center">
                                                Medida (P/FP)
                                              </div>
                                              <div className="text-[9px] font-bold text-sanesul-muted uppercase tracking-widest text-right">
                                                Gasto Real
                                              </div>
                                            </div>
                                            <div className="space-y-2">
                                              {uc.monthlyData.map(
                                                (month: any, mIdx: number) => (
                                                  <div
                                                    key={mIdx}
                                                    className="grid grid-cols-4 gap-4 px-4 py-3 bg-white rounded-xl border border-slate-100 hover:border-sanesul-primary/20 transition-colors"
                                                  >
                                                    <div className="text-xs font-bold text-slate-700">
                                                      {formatMonth(month.mes)}/
                                                      {month.ano}
                                                    </div>
                                                    <div className="text-xs font-mono text-center text-slate-500">
                                                      {month.dcp} / {month.dcfp}{" "}
                                                      kW
                                                    </div>
                                                    <div className="text-xs font-mono text-center text-slate-500">
                                                      {month.dmp} / {month.dmfp}{" "}
                                                      kW
                                                    </div>
                                                    <div className="text-xs font-mono text-right font-bold text-sanesul-primary">
                                                      R${" "}
                                                      {month.currentTotal.toLocaleString(
                                                        "pt-BR",
                                                        {
                                                          minimumFractionDigits: 2,
                                                        },
                                                      )}
                                                    </div>
                                                  </div>
                                                ),
                                              )}
                                            </div>
                                          </div>
                                        </td>
                                      </tr>
                                    )}
                                  </React.Fragment>
                                ),
                              )}
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
        ) : activeTab === "monitoramento_ajustes" ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                  Monitoramento Ajuste de Demanda
                </h2>
                <p className="text-sanesul-muted">
                  Acompanhamento das unidades com solicitações de alteração de contrato neste ano.
                </p>
              </div>
              <div className="flex flex-wrap gap-4 items-center">
                <button
                  onClick={() => {
                    setIsNewAdjustment(true);
                    setEditingAdjustment({
                      uc: "",
                      origP: 0,
                      origFP: 0,
                      reqP: 0,
                      reqFP: 0,
                      city: "",
                      gerencia: "",
                      dataSolicitacao: "",
                      dataAlteracao: "",
                      previsaoEconomia: "",
                      ecoRealizada: "",
                      status: "",
                    });
                    setIsAdjustmentModalOpen(true);
                  }}
                  className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20"
                >
                  <Plus size={16} />
                  Inserir Ajuste de UC
                </button>
                <button
                  onClick={exportAdjustmentsExcel}
                  className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg shadow-green-600/20"
                >
                  <FileSpreadsheet size={16} />
                  Exportar Solicitações
                </button>
              </div>
            </div>

            {/* Cards com Totais formatados */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              {/* Card Previsão de Economia Total */}
              <div className="bg-white rounded-[24px] p-8 border border-sanesul-primary/10 shadow-xl flex items-center justify-between hover:scale-[1.01] transition-transform duration-300">
                <div className="space-y-1">
                  <span className="text-xs font-bold text-slate-400 uppercase tracking-widest block">
                    Previsão de Economia Total Anual
                  </span>
                  <p className="text-3xl font-display font-extrabold text-sanesul-primary">
                    R$ {adjustmentsList.reduce((sum, adj) => {
                      const val = adj.previsaoEconomia;
                      if (val === "-" || !val) return sum;
                      const num = parseFloat(val);
                      return isNaN(num) ? sum : sum + num;
                    }, 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </p>
                </div>
                <div className="w-14 h-14 bg-blue-50 text-blue-600 rounded-2xl flex items-center justify-center shadow-inner">
                  <TrendingUp size={24} />
                </div>
              </div>

              {/* Card Economia Realizada Total */}
              <div className="bg-white rounded-[24px] p-8 border border-sanesul-primary/10 shadow-xl flex items-center justify-between hover:scale-[1.01] transition-transform duration-300">
                <div className="space-y-1">
                  <span className="text-xs font-bold text-slate-400 uppercase tracking-widest block">
                    Economia Realizada Total Acumulada
                  </span>
                  <p className="text-3xl font-display font-extrabold text-green-600">
                    R$ {adjustmentsList.reduce((sum, adj) => {
                      return sum + (adj.economia || 0);
                    }, 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                  </p>
                </div>
                <div className="w-14 h-14 bg-green-50 text-green-600 rounded-2xl flex items-center justify-center shadow-inner">
                  <DollarSign size={24} />
                </div>
              </div>
            </div>

            {(adjustmentsList.length === 0) ? (
              <div className="bg-white rounded-3xl p-16 text-center border border-slate-100 shadow-xl shadow-slate-200/50">
                <div className="w-24 h-24 bg-slate-50 rounded-full flex items-center justify-center mx-auto mb-6">
                  <FileText className="text-slate-300" size={48} />
                </div>
                <h3 className="text-xl font-bold text-slate-700 mb-2 font-display">Sem alterações</h3>
                <p className="text-slate-500 max-w-sm mx-auto">Nenhuma unidade com alteração de contrato registrada. Clique em "Inserir Ajuste de UC" para adicionar.</p>
              </div>
            ) : (
              <div className="bg-white rounded-[40px] border border-sanesul-primary/10 shadow-2xl overflow-hidden">
                <div className="bg-slate-50/80 px-10 py-8 border-b border-sanesul-primary/5">
                  <h3 className="text-2xl font-display font-bold text-sanesul-primary border-l-4 border-yellow-500 pl-4">
                    Relação das UCs
                  </h3>
                  <p className="text-xs font-bold text-sanesul-muted uppercase tracking-widest mt-2 pl-5">
                    {adjustmentsList.length} Unidades Solicitadas
                  </p>
                </div>
                <div className="p-10 overflow-x-auto overflow-y-auto max-h-[520px]">
                  <table className="w-full min-w-[1240px] text-left border-collapse">
                    <thead className="sticky top-0 z-10 bg-white shadow-sm">
                      <tr>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Nº UC</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Município</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Localidade</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-center">Contrato Assinado</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-center">Data Alteração</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">Status</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-center">Contratada P</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-center">Contratada P Alt.</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-center">Contratada FP</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-primary border-b border-sanesul-primary/5 text-center">Contratada FP Alt.</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-green-600 border-b border-sanesul-primary/5 text-right">Previsão Economia</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-green-700 border-b border-sanesul-primary/5 text-right">Eco. Realizada</th>
                        <th className="px-4 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-center">Ações</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {adjustmentsList.map((adj: any, idx: number) => {
                        const {
                          uc: ucId,
                          city,
                          gerencia,
                          dcpBefore,
                          dcfpBefore,
                          dcpAfter,
                          dcfpAfter,
                          dataAlteracao,
                          economia,
                          dataSolicitacao,
                          previsaoEconomia,
                          status
                        } = adj;

                        return (
                          <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-4 py-4 font-bold text-sanesul-primary font-mono whitespace-nowrap">{adj.uc}</td>
                            <td className="px-4 py-4 font-bold text-slate-800 text-xs">{city}</td>
                            <td className="px-4 py-4 font-bold text-slate-500 text-xs">{gerencia}</td>
                            <td className="px-4 py-4 text-center font-mono text-xs text-sanesul-muted">{dataSolicitacao}</td>
                            <td className="px-4 py-4 text-center font-mono text-xs text-sanesul-muted">{dataAlteracao}</td>
                            <td className="px-4 py-4 text-xs font-bold">
                              <span className={`px-2 py-0.5 rounded text-[10px] uppercase tracking-wider ${
                                status === "Concluída" 
                                  ? "bg-green-100 text-green-800" 
                                  : status === "Aguardando" 
                                  ? "bg-amber-100 text-amber-800" 
                                  : "bg-slate-100 text-slate-800"
                              }`}>
                                {status}
                              </span>
                            </td>
                            <td className="px-4 py-4 text-center">
                              <span className="px-2 py-1 bg-slate-50 text-slate-500 rounded text-xs font-mono">{dcpBefore}</span>
                            </td>
                            <td className="px-4 py-4 text-center">
                              <span className="px-2 py-1 bg-blue-50 text-blue-700 font-bold rounded text-xs font-mono">{dcpAfter}</span>
                            </td>
                            <td className="px-4 py-4 text-center">
                              <span className="px-2 py-1 bg-slate-50 text-slate-500 rounded text-xs font-mono">{dcfpBefore}</span>
                            </td>
                            <td className="px-4 py-4 text-center">
                              <span className="px-2 py-1 bg-blue-50 text-blue-700 font-bold rounded text-xs font-mono">{dcfpAfter}</span>
                            </td>
                            <td className="px-4 py-4 text-right font-mono text-xs text-sanesul-muted font-bold">
                              {previsaoEconomia === "-" ? "-" : isNaN(Number(previsaoEconomia)) ? previsaoEconomia : `R$ ${Number(previsaoEconomia).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`}
                            </td>
                            <td className="px-4 py-4 text-right font-bold text-green-600">
                              R$ {economia.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                            </td>
                            <td className="px-4 py-4 text-center">
                              <div className="flex items-center justify-center gap-2">
                                <button
                                  onClick={() => {
                                    setIsNewAdjustment(false);
                                    setEditingAdjustment({
                                      uc: adj.uc,
                                      origP: dcpBefore || 0,
                                      origFP: dcfpBefore || 0,
                                      reqP: dcpAfter || 0,
                                      reqFP: dcfpAfter || 0,
                                      city: city !== "-" ? city : "",
                                      gerencia: gerencia !== "-" ? gerencia : "",
                                      dataSolicitacao: dataSolicitacao !== "-" ? dataSolicitacao : "",
                                      dataAlteracao: dataAlteracao !== "-" ? dataAlteracao : "",
                                      previsaoEconomia: previsaoEconomia !== "-" ? previsaoEconomia : "",
                                      ecoRealizada: economia ? String(economia) : "",
                                      status: status,
                                    });
                                    setIsAdjustmentModalOpen(true);
                                  }}
                                  title="Editar"
                                  className="p-1.5 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-colors"
                                >
                                  <Pencil size={14} />
                                </button>
                                <button
                                  onClick={() => handleDeleteAdjustment(adj.uc)}
                                  title="Remover"
                                  className="p-1.5 bg-red-50 text-red-600 rounded-lg hover:bg-red-100 transition-colors"
                                >
                                  <Trash2 size={14} />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        ) : activeTab === "monitoramento_reativo" ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                  Monitoramento Reativo
                </h2>
                <p className="text-sanesul-muted">
                  Acompanhamento do Valor da Energia Reativa Excedente por UC.
                </p>
              </div>
              <div className="flex flex-wrap gap-4 items-center">
                <button
                  onClick={exportReactiveToCSV}
                  className="flex items-center gap-2 px-6 py-3 bg-green-600 text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-green-700 transition-all shadow-lg shadow-green-600/20"
                >
                  <Download size={16} />
                  Exportar Excel
                </button>
                <button
                  onClick={exportReactiveGroupedToCSV}
                  className="flex items-center gap-2 px-6 py-3 bg-sanesul-primary text-white rounded-xl text-xs font-bold uppercase tracking-widest hover:bg-sanesul-secondary transition-all shadow-lg shadow-sanesul-primary/20"
                >
                  <FileText size={16} />
                  Relatório por Gerência
                </button>
                <div className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm">
                  <Calendar size={16} className="text-sanesul-primary" />
                  <select
                    value={selectedReactiveMonth}
                    onChange={(e) => setSelectedReactiveMonth(e.target.value)}
                    className="bg-transparent text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                  >
                    <option value="all">Todos os Meses</option>
                    {availableMonths.map((month) => (
                      <option key={month} value={month}>
                        {month}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            {(() => {
              const reactiveBills = bills.filter((b) => {
                if (b.status !== "completed") return false;
                if (
                  selectedReactiveMonth !== "all" &&
                  `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` !== selectedReactiveMonth
                )
                  return false;
                const totalReativo =
                  parseValue(b.valorEnergiaReativaExcedPonta) +
                  parseValue(b.valorEnergiaReativaExcedFPonta);
                return totalReativo > 100;
              });

              const grouped = reactiveBills.reduce(
                (acc, bill) => {
                  const uc = String(bill.uc);
                  if (!acc[uc])
                    acc[uc] = {
                      uc,
                      cidade: getCidade(uc, bill.cidade),
                      totalPonta: 0,
                      totalFPonta: 0,
                      totalFatura: 0,
                      bills: [],
                    };
                  acc[uc].totalPonta += parseValue(
                    bill.valorEnergiaReativaExcedPonta,
                  );
                  acc[uc].totalFPonta += parseValue(
                    bill.valorEnergiaReativaExcedFPonta,
                  );
                  acc[uc].totalFatura += parseValue(bill.valorTotal);
                  acc[uc].bills.push(bill);
                  return acc;
                },
                {} as Record<
                  string,
                  {
                    uc: string;
                    cidade: string;
                    totalPonta: number;
                    totalFPonta: number;
                    totalFatura: number;
                    bills: typeof bills;
                  }
                >,
              );

              const reactiveData = (
                Object.values(grouped) as {
                  uc: string;
                  cidade: string;
                  totalPonta: number;
                  totalFPonta: number;
                  totalFatura: number;
                  bills: typeof bills;
                }[]
              ).sort((a, b) => {
                let valA: any = 0;
                let valB: any = 0;

                if (reactiveSortField === "uc") {
                  valA = a.uc;
                  valB = b.uc;
                } else if (reactiveSortField === "cidade") {
                  valA = a.cidade;
                  valB = b.cidade;
                } else if (reactiveSortField === "totalPonta") {
                  valA = a.totalPonta;
                  valB = b.totalPonta;
                } else if (reactiveSortField === "totalFPonta") {
                  valA = a.totalFPonta;
                  valB = b.totalFPonta;
                } else if (reactiveSortField === "totalGeral") {
                  valA = a.totalPonta + a.totalFPonta;
                  valB = b.totalPonta + b.totalFPonta;
                } else if (reactiveSortField === "percentual") {
                  valA =
                    a.totalFatura > 0
                      ? (a.totalPonta + a.totalFPonta) / a.totalFatura
                      : 0;
                  valB =
                    b.totalFatura > 0
                      ? (b.totalPonta + b.totalFPonta) / b.totalFatura
                      : 0;
                }

                if (typeof valA === "string" && typeof valB === "string") {
                  return reactiveSortDirection === "asc"
                    ? valA.localeCompare(valB)
                    : valB.localeCompare(valA);
                }
                return reactiveSortDirection === "asc"
                  ? (valA as number) - (valB as number)
                  : (valB as number) - (valA as number);
              });

              const totalGeral = reactiveData.reduce(
                (acc, curr) => acc + curr.totalPonta + curr.totalFPonta,
                0,
              );

              return (
                <div className="space-y-8">
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                    <div className="bg-white p-10 rounded-[40px] border border-sanesul-primary/5 shadow-2xl shadow-sanesul-primary/5 relative overflow-hidden group hover:border-sanesul-primary/20 transition-all">
                      <div className="absolute top-0 right-0 p-6 opacity-10 group-hover:scale-110 transition-transform">
                        <Zap size={80} className="text-red-600" />
                      </div>
                      <p className="text-[11px] font-bold uppercase tracking-[0.3em] text-sanesul-muted mb-4">
                        Total Multa Reativa
                      </p>
                      <p className="text-4xl font-display font-bold text-red-600">
                        R${" "}
                        {totalGeral.toLocaleString("pt-BR", {
                          minimumFractionDigits: 2,
                        })}
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl">
                    <div className="overflow-x-auto overflow-y-auto max-h-[520px]">
                      <table className="w-full text-left border-collapse">
                        <thead className="sticky top-0 z-10 bg-white">
                          <tr>
                            <th
                              className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary transition-colors"
                              onClick={() => handleReactiveSort("uc")}
                            >
                              <div className="flex items-center gap-1">
                                UC{" "}
                                {reactiveSortField === "uc" &&
                                  (reactiveSortDirection === "asc" ? (
                                    <TrendingUp size={12} />
                                  ) : (
                                    <TrendingDown size={12} />
                                  ))}
                              </div>
                            </th>
                            <th className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5">
                              Gerência
                            </th>
                            <th
                              className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 cursor-pointer hover:text-sanesul-primary transition-colors"
                              onClick={() => handleReactiveSort("cidade")}
                            >
                              <div className="flex items-center gap-1">
                                Cidade{" "}
                                {reactiveSortField === "cidade" &&
                                  (reactiveSortDirection === "asc" ? (
                                    <TrendingUp size={12} />
                                  ) : (
                                    <TrendingDown size={12} />
                                  ))}
                              </div>
                            </th>
                            <th
                              className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                              onClick={() => handleReactiveSort("totalPonta")}
                            >
                              <div className="flex items-center justify-end gap-1">
                                Total Reativa Ponta (R$){" "}
                                {reactiveSortField === "totalPonta" &&
                                  (reactiveSortDirection === "asc" ? (
                                    <TrendingUp size={12} />
                                  ) : (
                                    <TrendingDown size={12} />
                                  ))}
                              </div>
                            </th>
                            <th
                              className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                              onClick={() => handleReactiveSort("totalFPonta")}
                            >
                              <div className="flex items-center justify-end gap-1">
                                Total Reativa F. Ponta (R$){" "}
                                {reactiveSortField === "totalFPonta" &&
                                  (reactiveSortDirection === "asc" ? (
                                    <TrendingUp size={12} />
                                  ) : (
                                    <TrendingDown size={12} />
                                  ))}
                              </div>
                            </th>
                            <th
                              className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                              onClick={() => handleReactiveSort("totalGeral")}
                            >
                              <div className="flex items-center justify-end gap-1">
                                Total Geral (R$){" "}
                                {reactiveSortField === "totalGeral" &&
                                  (reactiveSortDirection === "asc" ? (
                                    <TrendingUp size={12} />
                                  ) : (
                                    <TrendingDown size={12} />
                                  ))}
                              </div>
                            </th>
                            <th
                              className="px-6 py-4 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-sanesul-primary/5 text-right cursor-pointer hover:text-sanesul-primary transition-colors"
                              onClick={() => handleReactiveSort("percentual")}
                            >
                              <div className="flex items-center justify-end gap-1">
                                % da Fatura{" "}
                                {reactiveSortField === "percentual" &&
                                  (reactiveSortDirection === "asc" ? (
                                    <TrendingUp size={12} />
                                  ) : (
                                    <TrendingDown size={12} />
                                  ))}
                              </div>
                            </th>
                            <th className="px-6 py-4 border-b border-sanesul-primary/5 w-10"></th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-sanesul-primary/5">
                          {(() => {
                            const paginatedReactive = reactiveData.slice(
                              0,
                              100,
                            );
                            return (
                              <>
                                {paginatedReactive.map((data, idx) => (
                                  <React.Fragment key={idx}>
                                    <tr
                                      className="hover:bg-sanesul-primary/5 transition-colors cursor-pointer group"
                                      onClick={() => toggleReactiveUc(data.uc)}
                                    >
                                      <td className="px-6 py-4">
                                        <div className="font-bold text-sanesul-primary">
                                          {data.uc}
                                        </div>
                                        <div className="text-[10px] text-sanesul-muted uppercase tracking-wider mt-1">
                                          {data.bills.length} faturas
                                        </div>
                                      </td>
                                      <td className="px-6 py-4">
                                        <div className="text-sm font-bold text-slate-800">
                                          {getGerencia(data.uc)}
                                        </div>
                                      </td>
                                      <td className="px-6 py-4">
                                        <div className="text-sm text-slate-600">
                                          {data.cidade}
                                        </div>
                                      </td>
                                      <td className="px-6 py-4 text-right">
                                        <div className="text-sm font-mono text-slate-600">
                                          R${" "}
                                          {data.totalPonta.toLocaleString(
                                            "pt-BR",
                                            {
                                              minimumFractionDigits: 2,
                                            },
                                          )}
                                        </div>
                                      </td>
                                      <td className="px-6 py-4 text-right">
                                        <div className="text-sm font-mono text-slate-600">
                                          R${" "}
                                          {data.totalFPonta.toLocaleString(
                                            "pt-BR",
                                            {
                                              minimumFractionDigits: 2,
                                            },
                                          )}
                                        </div>
                                      </td>
                                      <td className="px-6 py-4 text-right">
                                        <div className="text-sm font-bold text-red-600">
                                          R${" "}
                                          {(
                                            data.totalPonta + data.totalFPonta
                                          ).toLocaleString("pt-BR", {
                                            minimumFractionDigits: 2,
                                          })}
                                        </div>
                                      </td>
                                      <td className="px-6 py-4 text-right">
                                        <div className="text-sm font-bold text-orange-500">
                                          {data.totalFatura > 0
                                            ? (
                                                ((data.totalPonta +
                                                  data.totalFPonta) /
                                                  data.totalFatura) *
                                                100
                                              ).toLocaleString("pt-BR", {
                                                minimumFractionDigits: 2,
                                                maximumFractionDigits: 2,
                                              })
                                            : "0,00"}
                                          %
                                        </div>
                                      </td>
                                      <td className="px-6 py-4 text-center">
                                        <ChevronDown
                                          className={`w-5 h-5 text-sanesul-primary/40 transition-transform group-hover:text-sanesul-primary ${expandedReactiveUcs.has(data.uc) ? "rotate-180" : ""}`}
                                        />
                                      </td>
                                    </tr>
                                    {expandedReactiveUcs.has(data.uc) && (
                                      <tr>
                                        <td
                                          colSpan={8}
                                          className="p-0 bg-slate-50/50"
                                        >
                                          <div className="px-12 py-6 border-t border-sanesul-primary/5 shadow-inner">
                                            <table className="w-full text-left">
                                              <thead>
                                                <tr>
                                                  <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200">
                                                    Mês/Ano
                                                  </th>
                                                  <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">
                                                    Valor Ponta (R$)
                                                  </th>
                                                  <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">
                                                    Valor F. Ponta (R$)
                                                  </th>
                                                  <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">
                                                    Total Mês (R$)
                                                  </th>
                                                  <th className="px-4 py-2 text-[10px] font-bold uppercase tracking-widest text-slate-400 border-b border-slate-200 text-right">
                                                    % da Fatura
                                                  </th>
                                                </tr>
                                              </thead>
                                              <tbody className="divide-y divide-slate-100">
                                                {data.bills
                                                  .sort((a, b) => {
                                                    const yearA = parseInt(
                                                      a.anoLeitura || "0",
                                                      10,
                                                    );
                                                    const yearB = parseInt(
                                                      b.anoLeitura || "0",
                                                      10,
                                                    );
                                                    if (yearA !== yearB)
                                                      return yearB - yearA;
                                                    return (
                                                      getMonthNumber(
                                                        b.mesReferencia,
                                                      ) -
                                                      getMonthNumber(
                                                        a.mesReferencia,
                                                      )
                                                    );
                                                  })
                                                  .map((bill, bIdx) => {
                                                    const vPonta = parseValue(
                                                      bill.valorEnergiaReativaExcedPonta,
                                                    );
                                                    const vFPonta = parseValue(
                                                      bill.valorEnergiaReativaExcedFPonta,
                                                    );
                                                    const vFatura = parseValue(
                                                      bill.valorTotal,
                                                    );
                                                    const percentual =
                                                      vFatura > 0
                                                        ? ((vPonta + vFPonta) /
                                                            vFatura) *
                                                          100
                                                        : 0;
                                                    return (
                                                      <tr
                                                        key={bIdx}
                                                        className="hover:bg-white transition-colors"
                                                      >
                                                        <td className="px-4 py-3 text-sm font-medium text-slate-600">
                                                          {bill.mesReferencia}/
                                                          {bill.anoLeitura}
                                                        </td>
                                                        <td className="px-4 py-3 text-sm font-mono text-slate-500 text-right">
                                                          R${" "}
                                                          {vPonta.toLocaleString(
                                                            "pt-BR",
                                                            {
                                                              minimumFractionDigits: 2,
                                                            },
                                                          )}
                                                        </td>
                                                        <td className="px-4 py-3 text-sm font-mono text-slate-500 text-right">
                                                          R${" "}
                                                          {vFPonta.toLocaleString(
                                                            "pt-BR",
                                                            {
                                                              minimumFractionDigits: 2,
                                                            },
                                                          )}
                                                        </td>
                                                        <td className="px-4 py-3 text-sm font-bold text-red-500 text-right">
                                                          R${" "}
                                                          {(
                                                            vPonta + vFPonta
                                                          ).toLocaleString(
                                                            "pt-BR",
                                                            {
                                                              minimumFractionDigits: 2,
                                                            },
                                                          )}
                                                        </td>
                                                        <td className="px-4 py-3 text-sm font-bold text-orange-500 text-right">
                                                          {percentual.toLocaleString(
                                                            "pt-BR",
                                                            {
                                                              minimumFractionDigits: 2,
                                                              maximumFractionDigits: 2,
                                                            },
                                                          )}
                                                          %
                                                        </td>
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
                                {reactiveData.length > 100 && (
                                  <tr>
                                    <td
                                      colSpan={5}
                                      className="px-6 py-6 text-center text-sm text-slate-500 italic bg-slate-50 border-t border-slate-100"
                                    >
                                      Exibindo os primeiros 100 de{" "}
                                      {reactiveData.length} resultados.
                                    </td>
                                  </tr>
                                )}
                              </>
                            );
                          })()}
                          {reactiveData.length === 0 && (
                            <tr>
                              <td
                                colSpan={5}
                                className="px-6 py-12 text-center text-sanesul-muted"
                              >
                                Nenhuma UC com energia reativa excedente
                                encontrada.
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
        ) : activeTab === "monitoramento_usinas" ? (
          <div className="py-12 space-y-12">
            {(() => {
              // Trigger reactive re-evaluations when classification overrides happen
              const _trigger = ucsTrigger; 
              
              const injectedBills = bills.filter((b) => {
                if (b.status !== "completed") return false;
                if (
                  selectedUsinaMonth !== "all" &&
                  `${formatMonth(b.mesReferencia)}/${b.anoLeitura}` !== selectedUsinaMonth
                ) return false;
                const injOuc = parseValue(b.energiaAtvInjetadaGDIOUC);
                const injMuc = parseValue(b.energiaAtvInjetadaGDIMUC);
                return injOuc > 0 || injMuc > 0;
              });

              const usinasDataObj = injectedBills.reduce((acc, bill) => {
                 const uc = String(bill.uc);
                 if (!acc[uc]) {
                     const classification = UCS_PPP.has(uc) ? "PPP" : (UCS_USINA.has(uc) ? "SANESUL" : "OUTRO");
                     acc[uc] = {
                        uc,
                        locin: getLocin(uc),
                        gerencia: getGerencia(uc),
                        cidade: getCidade(uc, bill.cidade),
                        classification,
                        totalInjetadaOuc: 0,
                        totalValorOuc: 0,
                        totalInjetadaMuc: 0,
                        totalValorMuc: 0,
                        billsCount: 0
                     };
                 }
                 acc[uc].totalInjetadaOuc += parseValue(bill.energiaAtvInjetadaGDIOUC);
                 acc[uc].totalValorOuc += parseValue(bill.valorEnergiaAtvInjetadaGDIOUC);
                 acc[uc].totalInjetadaMuc += parseValue(bill.energiaAtvInjetadaGDIMUC);
                 acc[uc].totalValorMuc += parseValue(bill.valorEnergiaAtvInjetadaGDIMUC);
                 acc[uc].billsCount += 1;
                 return acc;
              }, {} as Record<string, any>);

              const usinasArray = (Object.values(usinasDataObj) as any[]).sort((a,b) => {
                  if (a.classification !== b.classification) return a.classification.localeCompare(b.classification);
                  return String(a.gerencia).localeCompare(String(b.gerencia));
              });

              const availableUsinaCities = Array.from(new Set(usinasArray.map(u => u.cidade))).filter(Boolean).sort();

              const filteredUsinasArray = selectedUsinaCity === "all" ? usinasArray : usinasArray.filter(u => u.cidade === selectedUsinaCity);

              const pppArray = filteredUsinasArray.filter(u => u.classification === "PPP");
              const sanesulArray = filteredUsinasArray.filter(u => u.classification !== "PPP");

              const handleExportUsinas = () => {
                if (filteredUsinasArray.length === 0) {
                  showAlert("Aviso", "Nenhum dado para exportar.");
                  return;
                }
                const exportData = filteredUsinasArray.map((u) => ({
                  "UC": u.uc,
                  "LOCIN": u.locin,
                  "Gerência": u.gerencia,
                  "Cidade": u.cidade,
                  "Classificação": u.classification,
                  "Injetada OUC (kWh)": u.totalInjetadaOuc,
                  "Valor Injetada OUC (R$)": u.totalValorOuc,
                  "Injetada MUC (kWh)": u.totalInjetadaMuc,
                  "Valor Injetada MUC (R$)": u.totalValorMuc,
                  "Quantidade Faturas": u.billsCount,
                }));
                const ws = XLSX.utils.json_to_sheet(exportData);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Usinas");
                XLSX.writeFile(wb, `monitoramento_usinas_${selectedUsinaMonth.replace("/", "_")}.xlsx`);
              };

              const toggleClass = (uc: string, current: string) => {
                 const targetClass = current === "PPP" ? "SANESUL" : "PPP";
                 showConfirm(
                   "Alterar Classificação",
                   `Tem certeza que deseja alterar a classificação da UC ${uc} de ${current} para ${targetClass}?`,
                   () => {
                     if (current === "PPP") {
                         UCS_PPP.delete(uc);
                         UCS_USINA.add(uc);
                     } else if (current === "SANESUL") {
                         UCS_USINA.delete(uc);
                         UCS_PPP.add(uc);
                     } else {
                         UCS_USINA.delete(uc);
                         UCS_PPP.add(uc);
                     }
                     if (typeof window !== "undefined") {
                         localStorage.setItem("PPP_UCS_OVERRIDE", JSON.stringify([...UCS_PPP]));
                         localStorage.setItem("USINA_UCS_OVERRIDE", JSON.stringify([...UCS_USINA]));
                     }
                     setUcsTrigger(t => t + 1);
                   }
                 );
              };

              const renderTable = (dataArray: any[], title: string, colorLine: string) => (
                  <div className="bg-white rounded-3xl border border-slate-200 overflow-hidden shadow-sm relative">
                     <div className={`absolute top-0 left-0 w-full h-1 ${colorLine}`}></div>
                     <div className="px-6 py-5 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
                        <h3 className="font-bold text-slate-800 uppercase tracking-widest text-sm">{title}</h3>
                        <span className="text-xs font-semibold text-slate-500 bg-white px-3 py-1 rounded-full border border-slate-200 shadow-sm">{dataArray.length} Registros</span>
                     </div>
                     <div className="overflow-x-auto overflow-y-auto max-h-[480px]">
                         <table className="w-full">
                            <thead className="sticky top-0 z-10">
                               <tr className="bg-white border-b border-slate-200">
                                   <th className="px-6 py-4 text-left text-xs font-bold text-slate-500 uppercase tracking-wider">UC</th>
                                   <th className="px-6 py-4 text-left text-xs font-bold text-slate-500 uppercase tracking-wider">LOCIN</th>
                                   <th className="px-6 py-4 text-left text-xs font-bold text-slate-500 uppercase tracking-wider">Gerência/Cidade</th>
                                   <th className="px-6 py-4 text-left text-xs font-bold text-slate-500 uppercase tracking-wider">Classificação</th>
                                   <th className="px-6 py-4 text-right text-xs font-bold text-slate-500 uppercase tracking-wider">Injetada oUC (kWh/R$)</th>
                                   <th className="px-6 py-4 text-right text-xs font-bold text-slate-500 uppercase tracking-wider">Injetada mUC (kWh/R$)</th>
                                   <th className="px-6 py-4 text-center text-xs font-bold text-slate-500 uppercase tracking-wider">Ações</th>
                               </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {dataArray.map(item => (
                                    <tr key={item.uc} className="hover:bg-slate-50 transition-colors bg-white">
                                       <td className="px-6 py-4 font-bold text-sanesul-primary">{item.uc}</td>
                                       <td className="px-6 py-4 text-sm text-slate-600">{item.locin}</td>
                                       <td className="px-6 py-4">
                                          <div className="text-sm font-semibold text-slate-700">{item.gerencia}</div>
                                          <div className="text-xs text-slate-500">{item.cidade}</div>
                                       </td>
                                       <td className="px-6 py-4">
                                           <span className={`px-3 py-1 text-xs font-bold uppercase tracking-wider flex items-center justify-center min-w-28 gap-1 w-fit rounded-full ${item.classification === "PPP" ? "bg-green-100 text-green-700" : item.classification === "SANESUL" ? "bg-blue-100 text-blue-700" : "bg-slate-100 text-slate-600"}`}>
                                              {item.classification}
                                           </span>
                                       </td>
                                       <td className="px-6 py-4 text-right">
                                          <div className="text-sm font-bold text-emerald-600">{item.totalInjetadaOuc.toLocaleString("pt-BR")} kWh</div>
                                          <div className="text-xs text-rose-600">{(item.totalValorOuc).toLocaleString("pt-BR", {style: "currency", currency: "BRL"})}</div>
                                       </td>
                                       <td className="px-6 py-4 text-right">
                                          <div className="text-sm font-bold text-emerald-600">{item.totalInjetadaMuc.toLocaleString("pt-BR")} kWh</div>
                                          <div className="text-xs text-rose-600">{(item.totalValorMuc).toLocaleString("pt-BR", {style: "currency", currency: "BRL"})}</div>
                                       </td>
                                       <td className="px-6 py-4 text-center">
                                          <button onClick={() => toggleClass(item.uc, item.classification)} className="text-[10px] px-3 py-2 rounded-lg border border-sanesul-primary/20 text-sanesul-primary hover:bg-sanesul-primary hover:text-white font-bold transition-colors uppercase tracking-widest whitespace-nowrap">
                                             Mudar p/ {item.classification === "PPP" ? "SANESUL" : "PPP"}
                                          </button>
                                       </td>
                                    </tr>
                                ))}
                                {dataArray.length === 0 && (
                                    <tr>
                                       <td colSpan={7} className="px-6 py-12 text-center text-slate-500 italic bg-white">
                                           Nenhum registro encontrado nesta categoria.
                                       </td>
                                    </tr>
                                )}
                            </tbody>
                         </table>
                     </div>
                  </div>
              );

              return (
                  <>
                    <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-8">
                      <div>
                        <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                          Monitoramento Usinas
                        </h2>
                        <p className="text-sanesul-muted">
                          Usinas com energia injetada e suas classificações (PPP Fotovoltaica e Usinas Sanesul).
                        </p>
                      </div>
                      <div className="flex flex-wrap gap-4 items-center">
                        <div className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm">
                          <MapPin size={16} className="text-sanesul-primary" />
                          <select
                            value={selectedUsinaCity}
                            onChange={(e) => setSelectedUsinaCity(e.target.value)}
                            className="bg-transparent text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                          >
                            <option value="all">Todas as Cidades</option>
                            {availableUsinaCities.map((cidade) => (
                              <option key={String(cidade)} value={String(cidade)}>
                                {String(cidade)}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm">
                          <Calendar size={16} className="text-sanesul-primary" />
                          <select
                            value={selectedUsinaMonth}
                            onChange={(e) => setSelectedUsinaMonth(e.target.value)}
                            className="bg-transparent text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                          >
                            <option value="all">Todos os Meses</option>
                            {availableMonths.map((month) => (
                              <option key={month} value={month}>
                                {month}
                              </option>
                            ))}
                          </select>
                        </div>
                        <button
                          onClick={handleExportUsinas}
                          className="flex items-center gap-2 bg-sanesul-primary text-white px-5 py-2 rounded-xl text-xs font-bold uppercase tracking-wider hover:bg-sanesul-primary/90 transition-colors shadow-sm"
                        >
                          <Download size={16} />
                          Exportar Excel
                        </button>
                      </div>
                    </div>
                    <div className="space-y-8">
                       {renderTable(pppArray, "Parceria Público-Privada (PPP Fotovoltaica)", "bg-emerald-500")}
                       {renderTable(sanesulArray, "Usinas Sanesul & Outros", "bg-blue-500")}
                    </div>
                  </>
              );
            })()}
          </div>
        ) : activeTab === "relatorio" ? (
          <div className="py-12 space-y-12">
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
              <div>
                <h2 className="text-3xl font-display font-bold text-sanesul-primary mb-2">
                  Relatório Financeiro
                </h2>
                <p className="text-sanesul-muted">
                  Visão consolidada de faturamento, impostos e indicadores
                  financeiros.
                </p>
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
                    {availableMonths.map((month) => (
                      <option key={month} value={month}>
                        {month}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="relative">
                  <button
                    onClick={() =>
                      setIsRelatorioTypeDropdownOpen(
                        !isRelatorioTypeDropdownOpen,
                      )
                    }
                    className="flex items-center gap-3 bg-white px-4 py-2 rounded-xl border border-sanesul-primary/10 shadow-sm text-xs font-bold text-sanesul-primary uppercase tracking-wider outline-none cursor-pointer"
                  >
                    <Filter size={16} className="text-sanesul-primary" />
                    <span>
                      {selectedRelatorioType.includes("all")
                        ? "Todos os Tipos"
                        : selectedRelatorioType.length > 1
                          ? `${selectedRelatorioType.length} Selecionados`
                          : selectedRelatorioType[0]}
                    </span>
                    <ChevronDown
                      size={14}
                      className={`transition-transform ${isRelatorioTypeDropdownOpen ? "rotate-180" : ""}`}
                    />
                  </button>

                  {isRelatorioTypeDropdownOpen && (
                    <>
                      <div
                        className="fixed inset-0 z-40"
                        onClick={() => setIsRelatorioTypeDropdownOpen(false)}
                      />
                      <div className="absolute top-full left-0 mt-2 w-48 bg-white rounded-xl shadow-xl border border-slate-100 z-50 py-2">
                        <label className="flex items-center gap-3 px-4 py-2 hover:bg-slate-50 cursor-pointer transition-colors">
                          <input
                            type="checkbox"
                            checked={selectedRelatorioType.includes("all")}
                            onChange={() => {
                              setSelectedRelatorioType(["all"]);
                            }}
                            className="rounded border-slate-300 text-sanesul-primary focus:ring-sanesul-primary"
                          />
                          <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                            Todos os Tipos
                          </span>
                        </label>
                        <div className="h-px bg-slate-100 my-1" />
                        {availableRelatorioTypes.map((type) => (
                          <label
                            key={type}
                            className="flex items-center gap-3 px-4 py-2 hover:bg-slate-50 cursor-pointer transition-colors"
                          >
                            <input
                              type="checkbox"
                              checked={selectedRelatorioType.includes(type)}
                              onChange={() => {
                                let newTypes = [
                                  ...selectedRelatorioType,
                                ].filter((t) => t !== "all");
                                if (newTypes.includes(type)) {
                                  newTypes = newTypes.filter((t) => t !== type);
                                } else {
                                  newTypes.push(type);
                                }
                                if (newTypes.length === 0) {
                                  setSelectedRelatorioType(["all"]);
                                } else {
                                  setSelectedRelatorioType(newTypes);
                                }
                              }}
                              className="rounded border-slate-300 text-sanesul-primary focus:ring-sanesul-primary"
                            />
                            <span className="text-xs font-bold text-slate-700 uppercase tracking-wider">
                              {type}
                            </span>
                          </label>
                        ))}
                      </div>
                    </>
                  )}
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={exportToCSV}
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

            <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
              <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl shadow-sanesul-primary/5 flex flex-col gap-2">
                <span className="text-[10px] font-bold uppercase tracking-widest text-sanesul-muted">
                  Valor Total
                </span>
                <span className="text-3xl font-display font-bold text-sanesul-primary">
                  R${" "}
                  {relatorioTotals.valorTotal.toLocaleString("pt-BR", {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2,
                  })}
                </span>
              </div>
              <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl shadow-sanesul-primary/5 flex flex-col gap-2">
                <span className="text-[10px] font-bold uppercase tracking-widest text-sanesul-muted">
                  Consumo Total
                </span>
                <span className="text-3xl font-display font-bold text-sanesul-primary">
                  {relatorioTotals.consumoTotal.toLocaleString("pt-BR", {
                    minimumFractionDigits: 0,
                    maximumFractionDigits: 0,
                  })}{" "}
                  <span className="text-sm font-sans font-medium opacity-40">
                    kWh
                  </span>
                </span>
              </div>
              <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl shadow-sanesul-primary/5 flex flex-col gap-2">
                <span className="text-[10px] font-bold uppercase tracking-widest text-sanesul-muted">
                  Total Injetado kWh
                </span>
                <span className="text-3xl font-display font-bold text-green-600">
                  {relatorioTotals.totalInjetadoKwh.toLocaleString("pt-BR", {
                    minimumFractionDigits: 0,
                    maximumFractionDigits: 0,
                  })}{" "}
                  <span className="text-sm font-sans font-medium opacity-40">
                    kWh
                  </span>
                </span>
              </div>
              <div className="bg-white p-8 rounded-[32px] border border-sanesul-primary/5 shadow-xl shadow-sanesul-primary/5 flex flex-col gap-2">
                <span className="text-[10px] font-bold uppercase tracking-widest text-sanesul-muted">
                  Valor Total Injetado
                </span>
                <span className="text-3xl font-display font-bold text-green-600">
                  R${" "}
                  {relatorioTotals.valorInjetado.toLocaleString("pt-BR", {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2,
                  })}
                </span>
              </div>
            </div>

            {Object.keys(uploadProgress).length > 0 && (
              <div className="mb-8 space-y-4">
                {Object.entries(uploadProgress).map(
                  ([fileId, progress]: [
                    string,
                    {
                      status: string;
                      percent: number;
                      fileName: string;
                      fileSize: number;
                      abortController: AbortController | null;
                    },
                  ]) => (
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
                            <span className="text-sm font-bold text-slate-900 truncate max-w-[200px]">
                              {progress.fileName}
                            </span>
                            <span className="text-[10px] text-slate-500">
                              {(progress.fileSize / 1024).toFixed(1)} KB
                            </span>
                          </div>
                          <span className="text-sm font-bold text-sanesul-primary">
                            {progress.percent}%
                          </span>
                        </div>
                        <div className="w-full bg-slate-100 rounded-full h-2.5 overflow-hidden">
                          <div
                            className="bg-sanesul-primary h-2.5 rounded-full transition-all duration-300 ease-out"
                            style={{ width: `${progress.percent}%` }}
                          ></div>
                        </div>
                        <span className="text-xs text-slate-500 mt-1 block">
                          {progress.status}
                        </span>
                      </div>
                      {progress.percent < 100 && progress.status !== "Erro" && (
                        <button
                          onClick={() => progress.abortController?.abort()}
                          className="text-red-500 hover:text-red-700 p-2 rounded-full hover:bg-red-50 transition-colors"
                          title="Cancelar upload"
                        >
                          <X size={20} />
                        </button>
                      )}
                    </motion.div>
                  ),
                )}
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
                  {(Object.values(agrupadoraFiles) as AgrupadoraData[]).map(
                    (data, idx) => {
                      const isDetailed =
                        data.concessionaria.includes("DETALHADO");
                      return (
                        <div
                          key={idx}
                          className="flex-1 bg-white p-5 rounded-xl border border-blue-100 shadow-sm min-w-[300px]"
                        >
                          <div className="flex justify-between items-start mb-4 pb-3 border-b border-slate-100">
                            <span className="text-sm font-bold text-blue-600 uppercase tracking-wider">
                              {data.concessionaria}
                            </span>
                            <span
                              className="text-[10px] text-slate-400 truncate max-w-[150px] text-right"
                              title={data.fileName}
                            >
                              {data.fileName}
                            </span>
                          </div>
                          <div className="grid grid-cols-2 gap-x-4 gap-y-3">
                            <div className="flex flex-col">
                              <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                Mês Ref
                              </span>
                              <span className="font-mono font-bold text-slate-700 text-sm">
                                {formatReference(data.mesReferencia) || "-"}
                              </span>
                            </div>
                            {!isDetailed && (
                              <div className="flex flex-col">
                                <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                  Vencimento
                                </span>
                                <span className="font-mono font-bold text-slate-700 text-sm">
                                  {data.vencimento || "-"}
                                </span>
                              </div>
                            )}
                            {!isDetailed && (
                              <div className="flex flex-col col-span-2">
                                <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                  Nota Fiscal
                                </span>
                                <span className="font-mono font-bold text-slate-700 text-sm">
                                  {data.numeroNotaFiscal || "-"}
                                </span>
                              </div>
                            )}
                            {!isDetailed && (
                              <>
                                <div className="flex flex-col">
                                  <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                    ICMS
                                  </span>
                                  <span className="font-mono font-bold text-slate-700 text-sm">
                                    R${" "}
                                    {data.icms.toLocaleString("pt-BR", {
                                      minimumFractionDigits: 2,
                                    })}
                                  </span>
                                </div>
                                <div className="flex flex-col">
                                  <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                    COFINS
                                  </span>
                                  <span className="font-mono font-bold text-slate-700 text-sm">
                                    R${" "}
                                    {data.cofins.toLocaleString("pt-BR", {
                                      minimumFractionDigits: 2,
                                    })}
                                  </span>
                                </div>
                                <div className="flex flex-col">
                                  <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                    PIS
                                  </span>
                                  <span className="font-mono font-bold text-slate-700 text-sm">
                                    R${" "}
                                    {data.pis.toLocaleString("pt-BR", {
                                      minimumFractionDigits: 2,
                                    })}
                                  </span>
                                </div>
                              </>
                            )}
                            <div
                              className={`flex flex-col ${isDetailed ? "col-span-2" : ""}`}
                            >
                              <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                CIP
                              </span>
                              <span
                                className={`font-mono font-bold text-slate-700 ${isDetailed ? "text-2xl text-blue-700" : "text-sm"}`}
                              >
                                R${" "}
                                {data.cip.toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </span>
                            </div>
                            {!isDetailed && (
                              <div className="flex flex-col col-span-2 pt-3 border-t border-slate-100 mt-1">
                                <span className="text-[10px] uppercase tracking-wider text-slate-400 font-bold">
                                  Total Faturado
                                </span>
                                <span className="font-mono font-bold text-blue-700 text-lg">
                                  R${" "}
                                  {data.valorTotal.toLocaleString("pt-BR", {
                                    minimumFractionDigits: 2,
                                  })}
                                </span>
                              </div>
                            )}
                          </div>
                        </div>
                      );
                    },
                  )}
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
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">
                      PIS
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.energisa.pis.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">
                      COFINS
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.energisa.cofins.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">
                      ICMS
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.energisa.icms.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-sanesul-primary/5 rounded-2xl border border-sanesul-primary/10">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-primary">
                      CIP MUNICIPAL
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.energisa.cip.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
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
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">
                      PIS
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.elektro.pis.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">
                      COFINS
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.elektro.cofins.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-slate-50 rounded-2xl">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-muted">
                      ICMS
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.elektro.icms.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                  <div className="flex justify-between items-center p-4 bg-sanesul-primary/5 rounded-2xl border border-sanesul-primary/10">
                    <span className="text-xs font-bold uppercase tracking-wider text-sanesul-primary">
                      CIP MUNICIPAL
                    </span>
                    <span className="text-xl font-display font-bold text-sanesul-primary">
                      R${" "}
                      {memoData.elektro.cip.toLocaleString("pt-BR", {
                        minimumFractionDigits: 2,
                      })}
                    </span>
                  </div>
                </div>
              </div>
            </div>

            <div className="bg-white rounded-[40px] border border-sanesul-primary/5 shadow-2xl overflow-hidden hidden">
              <div className="p-10 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-xl font-display font-bold text-sanesul-primary">
                  Detalhamento por Unidade Consumidora
                </h3>
                <div className="text-[10px] font-bold text-sanesul-muted uppercase tracking-widest bg-slate-100 px-3 py-1 rounded-full">
                  {
                    Array.from(
                      new Set(filteredRelatorioData.map((d) => d.uc)),
                    ).filter(Boolean).length
                  }{" "}
                  Unidades
                </div>
              </div>
              <div className="overflow-auto max-h-[600px]">
                <table className="w-full text-left border-collapse">
                  <thead className="sticky top-0 z-10 bg-slate-50">
                    <tr className="bg-slate-50/50">
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100">
                        UC
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100">
                        Gerência
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100">
                        Cidade
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">
                        PIS
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">
                        COFINS
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">
                        ICMS
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">
                        CIP
                      </th>
                      <th className="px-8 py-5 text-[10px] font-bold uppercase tracking-widest text-sanesul-muted border-b border-slate-100 text-right">
                        Total
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {Array.from(new Set(filteredRelatorioData.map((d) => d.uc)))
                      .filter(Boolean)
                      .map((ucDataRaw) => {
                        const uc = ucDataRaw as string;
                        const ucData = filteredRelatorioData.filter(
                          (d) => d.uc === uc,
                        );
                        const totals = {
                          pis: ucData.reduce((acc, curr) => acc + curr.pis, 0),
                          cofins: ucData.reduce(
                            (acc, curr) => acc + curr.cofins,
                            0,
                          ),
                          icms: ucData.reduce(
                            (acc, curr) => acc + curr.icms,
                            0,
                          ),
                          cip: ucData.reduce((acc, curr) => acc + curr.cip, 0),
                          total: ucData.reduce(
                            (acc, curr) => acc + curr.valorTotal,
                            0,
                          ),
                        };
                        return (
                          <tr
                            key={uc}
                            className="hover:bg-slate-50/50 transition-colors group"
                          >
                            <td className="px-8 py-6 border-b border-slate-50">
                              <span className="text-sm font-bold text-sanesul-primary group-hover:text-sanesul-secondary transition-colors">
                                {uc}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50">
                              <span className="text-xs font-bold text-slate-800">
                                {getGerencia(uc || "")}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50">
                              <span className="text-xs font-medium text-sanesul-muted">
                                {ucData[0]?.cidade || "-"}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50 text-right">
                              <span className="text-xs font-mono font-bold text-slate-600">
                                R${" "}
                                {totals.pis.toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50 text-right">
                              <span className="text-xs font-mono font-bold text-slate-600">
                                R${" "}
                                {totals.cofins.toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50 text-right">
                              <span className="text-xs font-mono font-bold text-slate-600">
                                R${" "}
                                {totals.icms.toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50 text-right">
                              <span className="text-xs font-mono font-bold text-sanesul-primary">
                                R${" "}
                                {totals.cip.toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </span>
                            </td>
                            <td className="px-8 py-6 border-b border-slate-50 text-right">
                              <span className="text-sm font-display font-bold text-sanesul-primary">
                                R${" "}
                                {totals.total.toLocaleString("pt-BR", {
                                  minimumFractionDigits: 2,
                                })}
                              </span>
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
                <h3 className="text-lg font-display font-bold text-sanesul-primary">
                  Dados do Memorando
                </h3>
                <p className="text-[10px] text-sanesul-muted uppercase tracking-widest font-bold">
                  Identificação e Notas Fiscais
                </p>
              </div>
            </div>

            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-bold text-sanesul-muted uppercase tracking-widest mb-2 px-1">
                  Número do Memorando
                </label>
                <input
                  type="text"
                  value={tempMemoNumber}
                  onChange={(e) => setTempMemoNumber(e.target.value)}
                  placeholder="Ex: 001447/2024/GEDEO/DCO"
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold text-sanesul-primary focus:outline-none focus:ring-2 focus:ring-sanesul-primary/20 focus:bg-white transition-all"
                  autoFocus
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
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
                <label className="block text-[10px] font-bold text-sanesul-muted uppercase tracking-widest mb-2 px-1">
                  NF Energisa
                </label>
                <input
                  type="text"
                  value={tempMemoNfEnergisa}
                  onChange={(e) => setTempMemoNfEnergisa(e.target.value)}
                  placeholder="Ex: 123456"
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold text-sanesul-primary focus:outline-none focus:ring-2 focus:ring-sanesul-primary/20 focus:bg-white transition-all"
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
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
                <label className="block text-[10px] font-bold text-sanesul-muted uppercase tracking-widest mb-2 px-1">
                  NF Elektro
                </label>
                <input
                  type="text"
                  value={tempMemoNfElektro}
                  onChange={(e) => setTempMemoNfElektro(e.target.value)}
                  placeholder="Ex: 789012"
                  className="w-full px-4 py-3 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold text-sanesul-primary focus:outline-none focus:ring-2 focus:ring-sanesul-primary/20 focus:bg-white transition-all"
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
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
                    <h3 className="text-lg font-display font-bold text-sanesul-primary">
                      Memorando de Faturamento
                    </h3>
                    <p className="text-[10px] text-sanesul-muted uppercase tracking-widest font-bold">
                      GEDEO/DCO - Sanesul
                    </p>
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

              <div
                id="memo-content"
                className="p-16 font-serif text-slate-800 leading-relaxed print:p-8 bg-white"
              >
                {/* Header */}
                <div className="flex justify-between items-start mb-16">
                  <div className="flex items-center gap-6">
                    <img
                      src="https://www.sanesul.ms.gov.br/images/logo_sanesul.png"
                      alt="Sanesul"
                      className="h-20 object-contain"
                      referrerPolicy="no-referrer"
                    />
                    <div className="h-16 w-px bg-slate-200" />
                    <div className="text-[11px] font-bold text-sanesul-primary leading-tight uppercase tracking-tight">
                      Empresa de Saneamento de <br /> Mato Grosso do Sul S.A.
                      <br />
                      <span className="text-sanesul-muted font-medium">
                        Diretoria da Presidência
                      </span>
                    </div>
                  </div>
                  <img
                    src="https://www.ms.gov.br/wp-content/uploads/2023/01/logo-governo-ms.png"
                    alt="Governo MS"
                    className="h-20 object-contain"
                    referrerPolicy="no-referrer"
                  />
                </div>

                {/* Memo Info */}
                <div className="space-y-1 mb-12 text-sm">
                  <p className="font-bold text-base memo-number-text">
                    MEMO Nº {memoNumber}
                  </p>
                  <p className="text-slate-500 italic">
                    Campo Grande,{" "}
                    {new Intl.DateTimeFormat("pt-BR", {
                      dateStyle: "long",
                    }).format(new Date())}
                    .
                  </p>
                </div>

                <div className="space-y-4 mb-12 text-sm border-l-4 border-sanesul-primary/20 pl-6 py-2">
                  <p>
                    <span className="font-bold text-sanesul-primary uppercase tracking-wider text-[10px]">
                      DE:
                    </span>{" "}
                    <br />
                    GEDEO - Gerência de Desenvolvimento Operacional
                  </p>
                  <p>
                    <span className="font-bold text-sanesul-primary uppercase tracking-wider text-[10px]">
                      PARA:
                    </span>{" "}
                    <br />
                    GEFI - Gerência Financeira e Gestão de Recursos
                  </p>
                  <p>
                    <span className="font-bold text-sanesul-primary uppercase tracking-wider text-[10px]">
                      ASSUNTO:
                    </span>{" "}
                    <br />
                    Faturas Agrupadora Operacional Energisa e Agrupadora Elektro
                    —{" "}
                    {selectedRelatorioMonth === "all"
                      ? "Consolidado"
                      : selectedRelatorioMonth}
                    {!selectedRelatorioType.includes("all")
                      ? ` (${selectedRelatorioType.join(", ")})`
                      : ""}
                    .
                  </p>
                </div>

                <p className="mb-6 font-medium">Prezado(a),</p>
                <p className="mb-10 text-justify">
                  Seguem anexas para pagamento as faturas de energia elétrica
                  Agrupadora da concessionária Energisa MS, e Agrupadora da
                  concessionária Elektro — todas referentes ao mês de{" "}
                  <span className="font-bold underline decoration-sanesul-primary/30 underline-offset-4">
                    {selectedRelatorioMonth === "all"
                      ? "todos os períodos"
                      : selectedRelatorioMonth}
                  </span>
                  {!selectedRelatorioType.includes("all")
                    ? ` (Tipo: ${selectedRelatorioType.join(", ")})`
                    : ""}{" "}
                  e correspondentes às unidades operacionais da SANESUL.
                </p>

                <div className="mb-6 flex items-center gap-3">
                  <div className="h-px flex-1 bg-slate-200" />
                  <p className="font-bold text-xs uppercase tracking-widest text-slate-400">
                    Tabela 1 - Faturas Anexas
                  </p>
                  <div className="h-px flex-1 bg-slate-200" />
                </div>

                <table className="w-full border-collapse border border-slate-300 text-sm mb-12 shadow-sm">
                  <thead>
                    <tr className="bg-slate-50">
                      <th className="border border-slate-300 p-3 text-left text-[10px] font-bold uppercase tracking-wider">
                        LOCALIDADE
                      </th>
                      <th className="border border-slate-300 p-3 text-right text-[10px] font-bold uppercase tracking-wider">
                        VALOR (R$)
                      </th>
                      <th className="border border-slate-300 p-3 text-center text-[10px] font-bold uppercase tracking-wider">
                        NOTA FISCAL
                      </th>
                      <th className="border border-slate-300 p-3 text-center text-[10px] font-bold uppercase tracking-wider">
                        REF: MÊS / ANO
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {/* Energisa */}
                    <tr className="font-bold bg-sanesul-primary/5">
                      <td className="border border-slate-300 p-3 text-sanesul-primary">
                        Agrupadora Energisa Operacional
                      </td>
                      <td className="border border-slate-300 p-3 text-right text-sanesul-primary">
                        R${" "}
                        {memoData.energisa.total.toLocaleString("pt-BR", {
                          minimumFractionDigits: 2,
                        })}
                      </td>
                      <td
                        rowSpan={5}
                        className="border border-slate-300 p-3 text-center align-middle font-mono text-xs max-w-[120px] break-all"
                      >
                        {memoData.energisa.nf}
                      </td>
                      <td
                        rowSpan={5}
                        className="border border-slate-300 p-3 text-center align-middle font-bold text-sanesul-primary"
                      >
                        {memoData.energisa.mesRef}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">
                        PIS
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.energisa.pis > 0
                          ? `R$ ${memoData.energisa.pis.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">
                        COFINS
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.energisa.cofins > 0
                          ? `R$ ${memoData.energisa.cofins.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">
                        ICMS
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.energisa.icms > 0
                          ? `R$ ${memoData.energisa.icms.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600 italic">
                        Tarifa de Iluminação Pública
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.energisa.cip > 0
                          ? `R$ ${memoData.energisa.cip.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>

                    {/* Elektro */}
                    <tr className="font-bold bg-sanesul-secondary/5">
                      <td className="border border-slate-300 p-3 text-sanesul-secondary">
                        Agrupadora Elektro
                      </td>
                      <td className="border border-slate-300 p-3 text-right text-sanesul-secondary">
                        R${" "}
                        {memoData.elektro.total.toLocaleString("pt-BR", {
                          minimumFractionDigits: 2,
                        })}
                      </td>
                      <td
                        rowSpan={5}
                        className="border border-slate-300 p-3 text-center align-middle font-mono text-xs max-w-[120px] break-all"
                      >
                        {memoData.elektro.nf}
                      </td>
                      <td
                        rowSpan={5}
                        className="border border-slate-300 p-3 text-center align-middle font-bold text-sanesul-secondary"
                      >
                        {memoData.elektro.mesRef}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">
                        PIS
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.elektro.pis > 0
                          ? `R$ ${memoData.elektro.pis.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">
                        COFINS
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.elektro.cofins > 0
                          ? `R$ ${memoData.elektro.cofins.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600">
                        ICMS
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.elektro.icms > 0
                          ? `R$ ${memoData.elektro.icms.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-300 p-2 pl-8 text-[11px] text-slate-600 italic">
                        Tarifa de Iluminação Pública
                      </td>
                      <td className="border border-slate-300 p-2 text-right text-[11px] text-slate-600 font-mono">
                        {memoData.elektro.cip > 0
                          ? `R$ ${memoData.elektro.cip.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          : "-"}
                      </td>
                    </tr>

                    {/* Total */}
                    <tr className="font-bold bg-slate-100">
                      <td className="border border-slate-300 p-4 uppercase text-xs tracking-wider">
                        TOTAL (Agrupadora ENERGISA + ELEKTRO)
                      </td>
                      <td className="border border-slate-300 p-4 text-right text-base text-sanesul-primary">
                        R${" "}
                        {(
                          memoData.energisa.total + memoData.elektro.total
                        ).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                      </td>
                      <td className="border border-slate-300 p-4 text-center text-slate-300">
                        -------------------
                      </td>
                      <td className="border border-slate-300 p-4 text-center text-slate-300">
                        -------------------
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div className="flex justify-between items-end text-[10px] text-slate-400 mb-12 italic">
                  <p>Proc. N.º 694/2018</p>
                  <p>Nota Orçamentária Nº 003/2019</p>
                </div>

                <div className="mt-16 text-sm text-slate-600 mb-12">
                  <p>
                    A planilha contendo a estratificação dos dados apresentados
                    neste memorando está disponível em{" "}
                    <span className="font-mono text-[10px] bg-slate-50 px-2 py-1 rounded break-all">
                      \\srv-fs
                      01\DADOS\DCO\GEDEO\OPERACAO_AGUA\COTAA\ENERGIA\FATURAS
                    </span>
                    .
                  </p>
                </div>

                <p className="mb-12">Atenciosamente,</p>

                <div className="mt-32 text-center">
                  <div className="w-72 h-px bg-slate-800 mx-auto mb-4" />
                  <p className="font-bold text-lg text-slate-900">
                    Fabio Roberto Alves da Silva
                  </p>
                  <p className="text-xs text-slate-500 uppercase tracking-widest font-bold mt-1">
                    Engenheiro Eletricista/GEDEO
                  </p>
                  <p className="text-[10px] text-slate-400 mt-0.5">
                    Gerência de Desenvolvimento Operacional
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Footer Info */}
      {isGerenciasModalOpen && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[60] flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-5xl h-[80vh] flex flex-col">
            <div className="flex items-center justify-between p-6 border-b border-slate-100">
              <h2 className="text-xl font-bold text-sanesul-primary flex items-center gap-2">
                <Building size={20} />
                Gestão de Gerências e LOCINS
              </h2>
              <button
                onClick={() => setIsGerenciasModalOpen(false)}
                className="w-8 h-8 flex items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 transition-colors"
              >
                <X size={20} />
              </button>
            </div>
            <div className="p-6 flex-1 flex flex-col gap-6 overflow-hidden">
              <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 shrink-0">
                <div className="flex items-center justify-between mb-4 border-b border-slate-200 pb-2">
                  <h3 className="text-sm font-bold text-sanesul-primary">
                    Cadastrar / Editar
                  </h3>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => {
                        showConfirm(
                          "Excluir Todos os Mapeamentos",
                          "Tem certeza que deseja excluir TODOS os mapeamentos cadastrados? Esta ação não pode ser desfeita.",
                          () => saveUcMappings([]),
                          "danger",
                        );
                      }}
                      className="px-4 py-1.5 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg text-xs font-bold transition-colors flex items-center gap-2"
                    >
                      <Trash2 size={14} />
                      Excluir Todos
                    </button>
                    <button
                      onClick={syncMappingsToSupabase}
                      disabled={isSyncingMappings}
                      className={`px-4 py-1.5 rounded-lg text-xs font-bold transition-colors flex items-center gap-2 relative overflow-hidden ${isSyncingMappings ? "bg-slate-100 text-slate-400 cursor-not-allowed" : "bg-green-50 text-green-600 hover:bg-green-100"}`}
                    >
                      {isSyncingMappings && totalSyncItems > 0 && (
                        <div
                          className="absolute left-0 top-0 bottom-0 bg-green-100 z-0 transition-all duration-300"
                          style={{
                            width: `${(syncProgress / totalSyncItems) * 100}%`,
                          }}
                        />
                      )}
                      <Database
                        size={14}
                        className={`z-10 ${isSyncingMappings ? "opacity-50" : ""}`}
                      />
                      <span className="z-10">
                        {isSyncingMappings
                          ? `Sincronizando ${syncProgress}/${totalSyncItems}...`
                          : "Sincronizar com Banco"}
                      </span>
                    </button>
                    <label className="cursor-pointer px-4 py-1.5 bg-blue-50 text-blue-600 hover:bg-blue-100 rounded-lg text-xs font-bold transition-colors flex items-center gap-2">
                      <Upload size={14} />
                      Importar TXT/CSV
                      <input
                        type="file"
                        accept=".txt,.csv"
                        className="hidden"
                        onChange={handleImportTxtGerencias}
                      />
                    </label>
                  </div>
                </div>
                <div className="flex flex-col md:flex-row gap-4 items-end">
                  <div className="space-y-1 flex-[0.5]">
                    <label className="text-xs font-bold text-slate-500 uppercase">
                      Pesquisar UC
                    </label>
                    <div className="relative">
                      <Search
                        size={14}
                        className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                      />
                      <input
                        type="text"
                        className="w-full pl-9 pr-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                        placeholder="Filtrar por UC..."
                        value={ucMappingSearchTerm}
                        onChange={(e) => setUcMappingSearchTerm(e.target.value)}
                      />
                    </div>
                  </div>
                  <div className="h-10 w-px bg-slate-200 hidden md:block mx-2 self-center mb-1"></div>
                  <div className="space-y-1 flex-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">
                      UC
                    </label>
                    <input
                      type="text"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                      placeholder="Ex: 3005931"
                      id="newGerenciaUC"
                    />
                  </div>
                  <div className="space-y-1 flex-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">
                      Gerência
                    </label>
                    <input
                      type="text"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                      placeholder="Ex: GRS"
                      id="newGerenciaName"
                    />
                  </div>
                  <div className="space-y-1 flex-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">
                      LOCINS
                    </label>
                    <input
                      type="text"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                      placeholder="Ex: SMJ"
                      id="newGerenciaLOCIN"
                    />
                  </div>
                  <div className="space-y-1 flex-1">
                    <label className="text-xs font-bold text-slate-500 uppercase">
                      Cidade
                    </label>
                    <input
                      type="text"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                      placeholder="Ex: Bonito"
                      id="newGerenciaCidade"
                    />
                  </div>
                  <button
                    onClick={() => {
                      const uc = (
                        document.getElementById(
                          "newGerenciaUC",
                        ) as HTMLInputElement
                      ).value;
                      const gerencia = (
                        document.getElementById(
                          "newGerenciaName",
                        ) as HTMLInputElement
                      ).value;
                      const locin = (
                        document.getElementById(
                          "newGerenciaLOCIN",
                        ) as HTMLInputElement
                      ).value;
                      const cidade = (
                        document.getElementById(
                          "newGerenciaCidade",
                        ) as HTMLInputElement
                      ).value;
                      if (!uc || !gerencia || !locin) {
                        showAlert("Atenção", "Preencha UC, Gerência e LOCINS.");
                        return;
                      }
                      const newMapping = { uc, gerencia, locin, cidade };
                      const newMappings = [
                        ...ucMappings.filter((m) => m.uc !== uc),
                        newMapping,
                      ];
                      saveUcMappings(newMappings);
                      (
                        document.getElementById(
                          "newGerenciaUC",
                        ) as HTMLInputElement
                      ).value = "";
                      (
                        document.getElementById(
                          "newGerenciaName",
                        ) as HTMLInputElement
                      ).value = "";
                      (
                        document.getElementById(
                          "newGerenciaLOCIN",
                        ) as HTMLInputElement
                      ).value = "";
                      (
                        document.getElementById(
                          "newGerenciaCidade",
                        ) as HTMLInputElement
                      ).value = "";
                    }}
                    className="px-6 py-2 bg-sanesul-primary text-white hover:bg-sanesul-primary/90 rounded-lg text-sm font-bold transition-colors shadow-lg active:scale-95"
                  >
                    Salvar
                  </button>
                </div>
              </div>

              <div className="bg-white rounded-xl border border-sanesul-primary/10 overflow-hidden flex-1 flex flex-col">
                <div className="overflow-y-auto flex-1">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-slate-50 border-b border-slate-200 sticky top-0 z-10">
                      <tr>
                        <th className="px-4 py-3 text-[10px] font-bold text-sanesul-muted uppercase tracking-widest">
                          UC
                        </th>
                        <th className="px-4 py-3 text-[10px] font-bold text-sanesul-muted uppercase tracking-widest">
                          Gerência
                        </th>
                        <th className="px-4 py-3 text-[10px] font-bold text-sanesul-muted uppercase tracking-widest">
                          Cidade
                        </th>
                        <th className="px-4 py-3 text-[10px] font-bold text-sanesul-muted uppercase tracking-widest">
                          LOCINS
                        </th>
                        <th className="px-4 py-3 text-[10px] font-bold text-sanesul-muted uppercase tracking-widest text-right">
                          Ações
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {(() => {
                        const filtered = ucMappings.filter(
                          (m) =>
                            (m.uc || "")
                              .toLowerCase()
                              .includes(ucMappingSearchTerm.toLowerCase()) ||
                            (m.gerencia || "")
                              .toLowerCase()
                              .includes(ucMappingSearchTerm.toLowerCase()) ||
                            (m.cidade || "")
                              .toLowerCase()
                              .includes(ucMappingSearchTerm.toLowerCase()),
                        );

                        if (filtered.length === 0) {
                          return (
                            <tr>
                              <td
                                colSpan={5}
                                className="px-4 py-8 text-center text-sm text-slate-500"
                              >
                                {ucMappingSearchTerm
                                  ? "Nenhum resultado para a pesquisa."
                                  : "Nenhum mapeamento cadastrado."}
                              </td>
                            </tr>
                          );
                        }

                        const paginatedFiltered = filtered.slice(0, 100);

                        return (
                          <>
                            {paginatedFiltered.map((m, i) => (
                              <tr key={i} className="hover:bg-slate-50">
                                <td className="px-4 py-3 text-sm font-bold text-sanesul-primary">
                                  {m.uc}
                                </td>
                                <td className="px-4 py-3 text-sm text-slate-700">
                                  {m.gerencia}
                                </td>
                                <td className="px-4 py-3 text-sm text-slate-700">
                                  {m.cidade || "---"}
                                </td>
                                <td className="px-4 py-3 text-sm text-slate-700">
                                  {m.locin}
                                </td>
                                <td className="px-4 py-3 text-sm text-right">
                                  <button
                                    onClick={() => {
                                      (
                                        document.getElementById(
                                          "newGerenciaUC",
                                        ) as HTMLInputElement
                                      ).value = m.uc;
                                      (
                                        document.getElementById(
                                          "newGerenciaName",
                                        ) as HTMLInputElement
                                      ).value = m.gerencia;
                                      (
                                        document.getElementById(
                                          "newGerenciaLOCIN",
                                        ) as HTMLInputElement
                                      ).value = m.locin;
                                      (
                                        document.getElementById(
                                          "newGerenciaCidade",
                                        ) as HTMLInputElement
                                      ).value = m.cidade || "";

                                      const nameInput =
                                        document.getElementById(
                                          "newGerenciaName",
                                        );
                                      if (nameInput) {
                                        nameInput.scrollIntoView({
                                          behavior: "smooth",
                                          block: "center",
                                        });
                                        nameInput.focus();
                                      }
                                    }}
                                    className="text-blue-500 hover:text-blue-700 mx-2 font-bold text-[10px] uppercase tracking-wider"
                                    title="Editar"
                                  >
                                    Editar
                                  </button>
                                  <button
                                    onClick={() =>
                                      syncSingleMappingToSupabase(m)
                                    }
                                    className="text-green-500 hover:text-green-700 mx-2 font-bold text-[10px] uppercase tracking-wider"
                                    title="Sincronizar com Banco de Dados"
                                  >
                                    Sincronizar
                                  </button>
                                  <button
                                    onClick={() => {
                                      showConfirm(
                                        "Remover Mapeamento",
                                        `Tem certeza que deseja remover o mapeamento da UC ${m.uc}?`,
                                        () =>
                                          saveUcMappings(
                                            ucMappings.filter(
                                              (x) => x.uc !== m.uc,
                                            ),
                                          ),
                                        "danger",
                                      );
                                    }}
                                    className="text-red-500 hover:text-red-700 mx-2"
                                    title="Remover"
                                  >
                                    <Trash2 size={16} />
                                  </button>
                                </td>
                              </tr>
                            ))}
                            {filtered.length > 100 && (
                              <tr>
                                <td
                                  colSpan={5}
                                  className="px-4 py-6 text-center text-sm text-slate-500 italic bg-slate-50 border-t border-slate-100"
                                >
                                  Exibindo os primeiros 100 de {filtered.length}{" "}
                                  resultados. Use a pesquisa para ver mais.
                                </td>
                              </tr>
                            )}
                          </>
                        );
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {isMercadoLivreModalOpen && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[60] flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl h-[70vh] flex flex-col">
            <div className="flex items-center justify-between p-6 border-b border-slate-100">
              <h2 className="text-xl font-bold text-sanesul-primary flex items-center gap-2">
                <DollarSign size={20} />
                Lista de UCs do Mercado Livre
              </h2>
              <button
                onClick={() => setIsMercadoLivreModalOpen(false)}
                className="w-8 h-8 flex items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 transition-colors"
              >
                <X size={20} />
              </button>
            </div>
            <div className="p-6 flex-1 flex flex-col gap-4 overflow-hidden">
              <p className="text-xs text-sanesul-muted leading-relaxed">
                Insira abaixo a lista de Unidades Consumidoras (UCs) que pertencem ao <strong>Mercado Livre</strong>. 
                Informe apenas os números das UCs, separados por quebra de linha, vírgulas ou espaços. 
                Faturas carregadas com UCs cadastradas nesta lista serão identificadas automaticamente como 
                "Mercado Livre" e seus formatos e relatórios serão ajustados de acordo.
              </p>
              
              <div className="flex-1 flex flex-col min-h-0 relative">
                <textarea
                  value={mercadoLivreInput}
                  onChange={(e) => setMercadoLivreInput(e.target.value)}
                  className="w-full flex-1 p-4 bg-slate-50 border border-slate-200 rounded-xl text-sm font-mono focus:ring-2 focus:ring-sanesul-primary/50 outline-none resize-none overflow-y-auto"
                  placeholder="Exemplo:&#10;65132005113&#10;158196505196&#10;103208705140"
                />
                
                <div className="absolute bottom-3 right-3 bg-slate-800/80 text-[10px] text-white font-bold px-3 py-1.5 rounded-lg shadow-sm backdrop-blur-sm pointer-events-none">
                  {mercadoLivreInput.split(/[\n,;\s\t]+/).map(s => s.trim()).filter(s => /^\d+$/.test(s)).length} UCs Válidas
                </div>
              </div>
            </div>
            <div className="p-6 border-t border-slate-100 flex justify-end gap-3 bg-slate-50 rounded-b-2xl">
              <button
                onClick={() => setIsMercadoLivreModalOpen(false)}
                className="px-6 py-2.5 bg-white border border-slate-200 hover:bg-slate-50 text-slate-700 rounded-xl text-xs font-bold uppercase tracking-wider transition-all"
              >
                Cancelar
              </button>
              <button
                onClick={saveMercadoLivreUcs}
                className="px-6 py-2.5 bg-sanesul-primary hover:bg-sanesul-secondary text-white rounded-xl text-xs font-bold uppercase tracking-wider transition-all shadow-md shadow-sanesul-primary/20"
              >
                Salvar Alterações
              </button>
            </div>
          </div>
        </div>
      )}

      {isBillModalOpen && editingBill && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-50 flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between p-6 border-b border-slate-100">
              <h2 className="text-xl font-bold text-sanesul-primary flex items-center gap-2">
                <Pencil size={20} />
                {editingBill.id && bills.some((b) => b.id === editingBill.id)
                  ? "Editar Fatura"
                  : "Nova Fatura Manual"}
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
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">
                    Informações Básicas
                  </h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    UC
                  </label>
                  <input
                    type="text"
                    value={editingBill.uc || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, uc: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Gerência
                  </label>
                  <input
                    type="text"
                    value={editingBill.gerencia || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, gerencia: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    LOCINS
                  </label>
                  <input
                    type="text"
                    value={editingBill.locin || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, locin: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Mercado
                  </label>
                  <select
                    value={editingBill.mercado || "CATIVO"}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, mercado: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  >
                    <option value="CATIVO">CATIVO</option>
                    <option value="LIVRE">LIVRE</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Concessionária
                  </label>
                  <select
                    value={editingBill.concessionaria || "ENERGISA"}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        concessionaria: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  >
                    <option value="ENERGISA">ENERGISA</option>
                    <option value="ELEKTRO">ELEKTRO</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Cidade
                  </label>
                  <input
                    type="text"
                    value={editingBill.cidade || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, cidade: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Mês Referência
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: Janeiro"
                    value={editingBill.mesReferencia || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        mesReferencia: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Ano Leitura
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: 2024"
                    value={editingBill.anoLeitura || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        anoLeitura: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Data de Vencimento
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: 15/08/2025"
                    value={editingBill.dataVencimento || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        dataVencimento: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Nota Fiscal
                  </label>
                  <input
                    type="text"
                    value={editingBill.numeroNotaFiscal || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        numeroNotaFiscal: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Modalidade Tarifária
                  </label>
                  <input
                    type="text"
                    value={editingBill.modalidadeTarifaria || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        modalidadeTarifaria: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Subgrupo
                  </label>
                  <input
                    type="text"
                    value={editingBill.subgrupo || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        subgrupo: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Tipo
                  </label>
                  <select
                    value={editingBill.tipo || "OPERACIONAL"}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, tipo: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  >
                    <option value="OPERACIONAL">OPERACIONAL</option>
                    <option value="ADMINISTRATIVO">ADMINISTRATIVO</option>
                    <option value="LIVRE">LIVRE</option>
                  </select>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Total (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorTotal || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorTotal: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Demanda */}
                {editingBill.concessionaria === "ELEKTRO" && (
                  <div className="space-y-1 mb-4">
                    <label className="text-xs font-bold text-slate-500 uppercase text-blue-600">Demanda Todos os Períodos (ELEKTRO)</label>
                    <input
                      type="text"
                      value={editingBill.demandaTodosPeriodos || ""}
                      onChange={(e) =>
                        setEditingBill({
                          ...editingBill,
                          demandaTodosPeriodos: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 bg-blue-50 border border-blue-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                    />
                  </div>
                )}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">
                    Demanda
                  </h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Todos os Períodos (kW)
                  </label>
                  <input
                    type="text"
                    value={editingBill.demandaTodosPeriodosKW || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaTodosPeriodosKW: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Contratada Ponta (kW)
                  </label>
                  <input
                    type="text"
                    value={editingBill.demandaPontaKW || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPontaKW: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Demanda fora ponta - kW
                  </label>
                  <input
                    type="text"
                    value={editingBill.demandaForaPontaKW || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaForaPontaKW: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Demanda de Potência Medida - Ponta</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaMedidaPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPotenciaMedidaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Medida Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorDemandaPotenciaMedidaPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorDemandaPotenciaMedidaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Demanda de Potência Medida - Fora Ponta</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaMedidaForaPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPotenciaMedidaForaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Medida Fora Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={
                      editingBill.valorDemandaPotenciaMedidaForaPonta || ""
                    }
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorDemandaPotenciaMedidaForaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Demanda Potência Não Consumida - Ponta</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaNaoConsumidaPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPotenciaNaoConsumidaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Não Consumida Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={
                      editingBill.valorDemandaPotenciaNaoConsumidaPonta || ""
                    }
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorDemandaPotenciaNaoConsumidaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Demanda Potência Não Consumida - Fora Ponta</label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaNaoConsumidaFPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPotenciaNaoConsumidaFPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Não Consumida Fora Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={
                      editingBill.valorDemandaPotenciaNaoConsumidaFPonta || ""
                    }
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorDemandaPotenciaNaoConsumidaFPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Ultrapassagem Ponta (kW)
                  </label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaAtivaUltrapPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPotenciaAtivaUltrapPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Ultrapassagem Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={
                      editingBill.valorDemandaPotenciaAtivaUltrapPonta || ""
                    }
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorDemandaPotenciaAtivaUltrapPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Ultrapassagem Fora Ponta (kW)
                  </label>
                  <input
                    type="text"
                    value={editingBill.demandaPotenciaAtivaUltrapFPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        demandaPotenciaAtivaUltrapFPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Ultrapassagem Fora Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={
                      editingBill.valorDemandaPotenciaAtivaUltrapFPonta || ""
                    }
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorDemandaPotenciaAtivaUltrapFPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Consumo */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">
                    Consumo
                  </h3>
                </div>
                {((editingBill.modalidadeTarifaria && editingBill.modalidadeTarifaria.toUpperCase().includes("CONVENCIONAL")) || 
                  (editingBill.subgrupo && editingBill.subgrupo.toUpperCase().includes("B")) || 
                  (editingBill.consumoKwhGrupoB !== undefined && editingBill.consumoKwhGrupoB !== "")) && (
                  <>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase text-blue-600">Consumo em kWh (Grupo B)</label>
                    <input
                      type="text"
                      value={editingBill.consumoKwhGrupoB || editingBill.consumoKwh || ""}
                      onChange={(e) =>
                        setEditingBill({
                          ...editingBill,
                          consumoKwhGrupoB: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 bg-blue-50 border border-blue-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-slate-500 uppercase text-blue-600">Consumo (R$) (Grupo B)</label>
                    <input
                      type="text"
                      value={editingBill.valorConsumoKwhGrupoB || ""}
                      onChange={(e) =>
                        setEditingBill({
                          ...editingBill,
                          valorConsumoKwhGrupoB: e.target.value,
                        })
                      }
                      className="w-full px-3 py-2 bg-blue-50 border border-blue-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                    />
                  </div>
                  </>
                )}
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Consumo Grupo B (kWh)
                  </label>
                  <input
                    type="text"
                    value={editingBill.consumoGrupoB || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        consumoGrupoB: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Consumo Ponta (kWh)
                  </label>
                  <input
                    type="text"
                    value={editingBill.consumoKwhPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        consumoKwhPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Consumo Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorConsumoKwhPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorConsumoKwhPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Consumo em kWh - Fora Ponta</label>
                  <input
                    type="text"
                    value={editingBill.consumoKwhForaPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        consumoKwhForaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Consumo Fora Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorConsumoKwhForaPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorConsumoKwhForaPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Reativa */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">
                    Energia Reativa
                  </h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Energia Reativa Exced em KWh - Ponta
                  </label>
                  <input
                    type="text"
                    value={editingBill.energiaReativaExcedPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        energiaReativaExcedPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Reativa Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaReativaExcedPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorEnergiaReativaExcedPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Energia Reativa Exced em KWh - Fponta
                  </label>
                  <input
                    type="text"
                    value={editingBill.energiaReativaExcedFPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        energiaReativaExcedFPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor Reativa Fora Ponta (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaReativaExcedFPonta || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorEnergiaReativaExcedFPonta: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Geração Distribuída */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">
                    Geração Distribuída
                  </h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Energia Atv Injetada GDI oUC (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaAtvInjetadaGDIOUC || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        energiaAtvInjetadaGDIOUC: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor GDI oUC (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaAtvInjetadaGDIOUC || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorEnergiaAtvInjetadaGDIOUC: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">Energia Atv Injetada GDI mUC (kWh)</label>
                  <input
                    type="text"
                    value={editingBill.energiaAtvInjetadaGDIMUC || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        energiaAtvInjetadaGDIMUC: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Valor GDI mUC (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.valorEnergiaAtvInjetadaGDIMUC || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        valorEnergiaAtvInjetadaGDIMUC: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>

                {/* Encargos e Impostos */}
                <div className="col-span-1 md:col-span-2 lg:col-span-3 mt-4 mb-1">
                  <h3 className="text-sm font-bold text-sanesul-primary border-b border-slate-200 pb-1">
                    Encargos e Impostos
                  </h3>
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    CIP (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.cip || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, cip: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    Outros Encargos (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.outrosEncargos || ""}
                    onChange={(e) =>
                      setEditingBill({
                        ...editingBill,
                        outrosEncargos: e.target.value,
                      })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    PIS (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.pis || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, pis: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    COFINS (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.cofins || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, cofins: e.target.value })
                    }
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase">
                    ICMS (R$)
                  </label>
                  <input
                    type="text"
                    value={editingBill.icms || ""}
                    onChange={(e) =>
                      setEditingBill({ ...editingBill, icms: e.target.value })
                    }
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
                    billToSave.tipo = "OPER";
                  } else if (
                    billToSave.uc &&
                    UCS_LIVRE_MERCADO_LIVRE.has(String(billToSave.uc))
                  ) {
                    let mod = billToSave.modalidadeTarifaria || "";
                    if (!mod.toUpperCase().includes("LIVRE")) {
                      billToSave.modalidadeTarifaria = mod
                        ? `${mod} - LIVRE`
                        : "LIVRE";
                    }
                    billToSave.tipo = "LIVRE";
                  }
                  billToSave.mercado =
                    billToSave.uc &&
                    UCS_LIVRE_MERCADO_LIVRE.has(String(billToSave.uc))
                      ? "LIVRE"
                      : "CATIVO";

                  // Enriquecer com Gerência e LOCIN
                  const mapping = ucMappings.find(
                    (m) => m.uc === String(billToSave.uc),
                  );
                  if (mapping) {
                    billToSave.gerencia = mapping.gerencia;
                    billToSave.locin = mapping.locin;
                    if (!billToSave.cidade) billToSave.cidade = mapping.cidade;
                  }

                  const isExisting = bills.some((b) => b.id === billToSave.id);

                  if (isSupabaseConfigured && isAuthenticated) {
                    try {
                      const {
                        data: { user },
                      } = await supabase.auth.getUser();
                      if (user) {
                        const dbData = mapBillDataToDb(billToSave, user.id);
                        if (isExisting) {
                          let { error } = await supabase
                            .from("bills")
                            .update(dbData)
                            .eq("id", billToSave.id);
                          if (
                            error &&
                            (error.message.includes("data_vencimento") ||
                              error.message.includes("mercado") ||
                              error.message.includes("gerencia") ||
                              error.message.includes("locins") ||
                error.message.includes("consumo_kwh_grupo_b") ||
                error.message.includes("valor_consumo_kwh_grupo_b") ||
                              error.details?.includes("data_vencimento") ||
                              error.details?.includes("mercado") ||
                              error.code === "PGRST204")
                          ) {
                            console.warn(
                              "Colunas novas não encontradas no Supabase. Atualizando sem elas...",
                            );
                            const {
              data_vencimento,
              mercado,
              gerencia,
              locins,
              consumo_kwh_grupo_b,
              valor_consumo_kwh_grupo_b,
              ...fallbackData
            } = dbData;
            if (consumo_kwh_grupo_b) fallbackData.consumo_kwh_ponta = consumo_kwh_grupo_b;
            if (valor_consumo_kwh_grupo_b) fallbackData.valor_consumo_kwh_ponta = valor_consumo_kwh_grupo_b;
                            const fallbackRes = await supabase
                              .from("bills")
                              .update(fallbackData)
                              .eq("id", billToSave.id);
                            error = fallbackRes.error;
                          }
                          if (error) {
                            console.error(
                              "Erro ao atualizar fatura no Supabase:",
                              error,
                            );
                            return;
                          }
                          setBills((prev) =>
                            prev.map((b) =>
                              b.id === billToSave.id ? billToSave : b,
                            ),
                          );
                        } else {
                          let { data, error } = await supabase
                            .from("bills")
                            .insert(dbData)
                            .select()
                            .single();
                          if (
                            error &&
                            (error.message.includes("data_vencimento") ||
                              error.message.includes("mercado") ||
                              error.message.includes("gerencia") ||
                              error.message.includes("locins") ||
                error.message.includes("consumo_kwh_grupo_b") ||
                error.message.includes("valor_consumo_kwh_grupo_b") ||
                              error.details?.includes("data_vencimento") ||
                              error.details?.includes("mercado") ||
                              error.code === "PGRST204")
                          ) {
                            console.warn(
                              "Colunas novas não encontradas no Supabase. Inserindo sem elas...",
                            );
                            const {
              data_vencimento,
              mercado,
              gerencia,
              locins,
              consumo_kwh_grupo_b,
              valor_consumo_kwh_grupo_b,
              ...fallbackData
            } = dbData;
            if (consumo_kwh_grupo_b) fallbackData.consumo_kwh_ponta = consumo_kwh_grupo_b;
            if (valor_consumo_kwh_grupo_b) fallbackData.valor_consumo_kwh_ponta = valor_consumo_kwh_grupo_b;
                            const fallbackRes = await supabase
                              .from("bills")
                              .insert(fallbackData)
                              .select()
                              .single();
                            error = fallbackRes.error;
                            data = fallbackRes.data;
                          }
                          if (error) {
                            console.error(
                              "Erro ao inserir fatura no Supabase:",
                              error,
                            );
                            return;
                          }
                          const newBill = mapDbToBillData(data);
                          setBills((prev) => {
                            if (prev.some((b) => b.id === newBill.id))
                              return prev;
                            return [...prev, newBill];
                          });
                        }
                      }
                    } catch (err) {
                      console.error("Erro inesperado ao salvar fatura:", err);
                      return;
                    }
                  } else {
                    if (isExisting) {
                      setBills((prev) =>
                        prev.map((b) =>
                          b.id === billToSave.id ? billToSave : b,
                        ),
                      );
                    } else {
                      setBills((prev) => {
                        if (prev.some((b) => b.id === billToSave.id))
                          return prev;
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

      {/* Modal para Adicionar/Editar Ajuste de Demanda */}
      {isAdjustmentModalOpen && editingAdjustment && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-sm z-[90] flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-lg flex flex-col my-8">
            <div className="flex items-center justify-between p-6 border-b border-slate-100">
              <h2 className="text-xl font-bold text-sanesul-primary flex items-center gap-2 font-display">
                <Pencil size={20} />
                {isNewAdjustment ? "Inserir UC no Monitoramento" : "Alterar Valores de Ajuste"}
              </h2>
              <button
                onClick={() => setIsAdjustmentModalOpen(false)}
                className="w-8 h-8 flex items-center justify-center rounded-lg text-slate-400 hover:bg-slate-100 transition-colors"
              >
                <X size={20} />
              </button>
            </div>
            
            {/* Scrollable container for many fields */}
            <div className="p-6 space-y-4 max-h-[70vh] overflow-y-auto">
              {/* UC Block */}
              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                  Nº da Unidade Consumidora (UC)
                </label>
                <input
                  type="text"
                  placeholder="Ex: 179291005130"
                  value={editingAdjustment.uc}
                  disabled={!isNewAdjustment}
                  onChange={(e) =>
                    setEditingAdjustment({ ...editingAdjustment, uc: e.target.value.trim() })
                  }
                  className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none disabled:opacity-60 disabled:cursor-not-allowed font-mono font-bold"
                />
                {isNewAdjustment && (
                  <p className="text-[10px] text-slate-400">
                    Insira o código numérico da UC tal como cadastrado nas faturas do sistema.
                  </p>
                )}
              </div>

              {/* Grid 1: City & Gerencia */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                    Município
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: Aquidauana"
                    value={editingAdjustment.city || ""}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        city: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                    Localidade / Gerência
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: ETA SUL"
                    value={editingAdjustment.gerencia || ""}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        gerencia: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none"
                  />
                </div>
              </div>

              {/* Grid 2: Original demand fields */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                    Contratada P (Original)
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={editingAdjustment.origP}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        origP: parseFloat(e.target.value) || 0,
                      })
                    }
                    className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                    Contratada FP (Original)
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={editingAdjustment.origFP}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        origFP: parseFloat(e.target.value) || 0,
                      })
                    }
                    className="w-full px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono"
                  />
                </div>
              </div>

              {/* Grid 3: Adjusted demand fields */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider text-sanesul-primary">
                    Contratada P Alt.
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={editingAdjustment.reqP}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        reqP: parseFloat(e.target.value) || 0,
                      })
                    }
                    className="w-full px-4 py-2 border border-blue-200 bg-blue-50/20 text-blue-900 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono font-bold"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider text-sanesul-primary">
                    Contratada FP Alt.
                  </label>
                  <input
                    type="number"
                    step="any"
                    value={editingAdjustment.reqFP}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        reqFP: parseFloat(e.target.value) || 0,
                      })
                    }
                    className="w-full px-4 py-2 border border-blue-200 bg-blue-50/20 text-blue-900 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono font-bold"
                  />
                </div>
              </div>

              {/* Grid 4: Dates */}
              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                    Contrato Assinado
                  </label>
                  <input
                    type="text"
                    placeholder="DD/MM/AAAA"
                    value={editingAdjustment.dataSolicitacao || ""}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        dataSolicitacao: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                    Data Alteração
                  </label>
                  <input
                    type="text"
                    placeholder="DD/MM/AAAA"
                    value={editingAdjustment.dataAlteracao || ""}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        dataAlteracao: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono"
                  />
                </div>
              </div>

              {/* Grid 5: Status & Economy values */}
              <div className="space-y-1">
                <label className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                  Status da Solicitação
                </label>
                <select
                  value={editingAdjustment.status || "Aguardando"}
                  onChange={(e) =>
                    setEditingAdjustment({
                      ...editingAdjustment,
                      status: e.target.value,
                    })
                  }
                  className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-bold"
                >
                  <option value="Aguardando">Aguardando</option>
                  <option value="Concluída">Concluída</option>
                  <option value="Cancelada">Cancelada</option>
                </select>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-1">
                  <label className="text-xs font-bold text-green-600 uppercase tracking-wider">
                    Previsão Economia (R$)
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: 1450.25 ou -"
                    value={editingAdjustment.previsaoEconomia || ""}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        previsaoEconomia: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono font-bold text-green-700"
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-xs font-bold text-green-700 uppercase tracking-wider">
                    Eco. Realizada (R$)
                  </label>
                  <input
                    type="text"
                    placeholder="Ex: 1200"
                    value={editingAdjustment.ecoRealizada || ""}
                    onChange={(e) =>
                      setEditingAdjustment({
                        ...editingAdjustment,
                        ecoRealizada: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm focus:ring-2 focus:ring-sanesul-primary/50 outline-none font-mono font-bold text-green-800"
                  />
                </div>
              </div>
            </div>
            
            <div className="flex items-center justify-end gap-3 px-6 py-4 border-t border-slate-100">
              <button
                onClick={() => setIsAdjustmentModalOpen(false)}
                className="px-5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition-all"
              >
                Cancelar
              </button>
              <button
                onClick={() => {
                  const {
                    uc,
                    origP,
                    origFP,
                    reqP,
                    reqFP,
                    city,
                    gerencia,
                    dataSolicitacao,
                    dataAlteracao,
                    previsaoEconomia,
                    ecoRealizada,
                    status
                  } = editingAdjustment;
                  
                  if (!uc) {
                    alert("Por favor, preencha o número de UC.");
                    return;
                  }
                  
                  // Save requested adjustments
                  setCustomRequestedAdjustments((prev) => ({
                    ...prev,
                    [uc]: { p: reqP, fp: reqFP },
                  }));
                  
                  // Save original adjustments
                  setCustomOriginalContratadas((prev) => ({
                    ...prev,
                    [uc]: { p: origP, fp: origFP },
                  }));

                  // Save metadata adjustments
                  setCustomAdjustmentsMetadata((prev) => ({
                    ...prev,
                    [uc]: {
                      city: city || "",
                      gerencia: gerencia || "",
                      dataSolicitacao: dataSolicitacao || "",
                      dataAlteracao: dataAlteracao || "",
                      previsaoEconomia: previsaoEconomia || "",
                      ecoRealizada: ecoRealizada || "",
                      status: status || "Aguardando",
                    },
                  }));
                  
                  setIsAdjustmentModalOpen(false);
                }}
                className="px-6 py-2 bg-sanesul-primary text-white hover:bg-sanesul-secondary rounded-xl text-xs font-bold transition-all shadow-lg shadow-sanesul-primary/20"
              >
                Salvar Alterações
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
              <div className="text-xs font-bold text-sanesul-primary uppercase tracking-widest">
                Sanesul - Portal de Inteligência Energética
              </div>
              <div className="text-[10px] text-sanesul-muted mt-0.5">
                © 1979 Empresa de Saneamento de Mato Grosso do Sul
              </div>
              <div className="text-[10px] text-sanesul-muted mt-0.5">
                Developed by: Fabio Roberto alves da Silva
              </div>
            </div>
          </div>
          <div className="flex items-center gap-8">
            <div className="flex flex-col items-end">
              <span className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">
                COTAA
              </span>
              <span className="text-[10px] text-sanesul-muted">
                Coodenação: Alexandre Santos Andrade Monteiro
              </span>
            </div>
            <div className="flex flex-col items-end">
              <span className="text-[10px] font-bold text-sanesul-primary uppercase tracking-widest">
                GEDEO
              </span>
              <span className="text-[10px] text-sanesul-muted">
                Gerência: Elthon Santos Teixeira
              </span>
            </div>
          </div>
        </div>
      </footer>
      {/* Modal: Excluir Faturas por Lista de UCs */}
      {isDeleteByListModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4 overflow-y-auto">
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 15 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95, y: 15 }}
            className="bg-white rounded-3xl shadow-2xl w-full max-w-2xl overflow-hidden border border-slate-100 flex flex-col max-h-[90vh]"
          >
            {/* Header */}
            <div className="flex items-center justify-between p-6 border-b border-slate-100 bg-slate-50/50">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-red-50 text-red-600 flex items-center justify-center shadow-sm">
                  <ListX size={22} />
                </div>
                <div>
                  <h2 className="text-lg font-bold text-slate-900">
                    Excluir Faturas por Lista de UCs
                  </h2>
                  <p className="text-xs text-slate-500">
                    Cole uma lista de UCs para remover todas as faturas correspondentes
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsDeleteByListModalOpen(false)}
                className="w-8 h-8 flex items-center justify-center rounded-lg text-slate-400 hover:bg-slate-200 transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            {/* Content */}
            <div className="p-6 flex-1 overflow-y-auto space-y-5">
              {/* Instructions */}
              <div className="bg-amber-50 border border-amber-200/70 rounded-2xl p-4 text-xs text-amber-800 space-y-1">
                <div className="font-bold flex items-center gap-1.5">
                  <AlertTriangle size={15} className="text-amber-600" />
                  Instruções de Formatação
                </div>
                <p className="text-amber-700 leading-relaxed">
                  Você pode colar ou digitar as UCs separadas por <strong>linha</strong>, <strong>vírgula</strong>, <strong>ponto e vírgula</strong> ou <strong>espaço</strong>.
                  Pontos, traços e caracteres não numéricos são normalizados automaticamente.
                </p>
              </div>

              {/* Textarea */}
              <div className="space-y-2">
                <div className="flex justify-between items-center">
                  <label className="text-xs font-bold text-slate-600 uppercase tracking-wider">
                    Lista de Unidades Consumidoras (UCs)
                  </label>
                  {deleteUcListInput && (
                    <button
                      onClick={() => setDeleteUcListInput("")}
                      className="text-[11px] font-bold text-slate-400 hover:text-red-500 transition-colors"
                    >
                      Limpar Texto
                    </button>
                  )}
                </div>
                <textarea
                  value={deleteUcListInput}
                  onChange={(e) => setDeleteUcListInput(e.target.value)}
                  placeholder="Cole as UCs aqui, exemplo:&#10;92106705133&#10;181951005101&#10;3005931&#10;2.822.635.051-30"
                  rows={6}
                  className="w-full p-4 border border-slate-200 rounded-2xl text-xs font-mono focus:ring-2 focus:ring-red-500/20 focus:border-red-500 outline-none bg-slate-50/50 resize-y"
                />
              </div>

              {/* Live Preview Stats */}
              {parsedInputUcs.length > 0 && (
                <div className="space-y-3">
                  <div className="grid grid-cols-3 gap-3">
                    <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-3.5 text-center">
                      <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                        UCs Informadas
                      </p>
                      <p className="text-xl font-bold text-slate-700 mt-0.5">
                        {parsedInputUcs.length}
                      </p>
                    </div>

                    <div className="bg-emerald-50 border border-emerald-200/80 rounded-2xl p-3.5 text-center">
                      <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider">
                        UCs Encontradas
                      </p>
                      <p className="text-xl font-bold text-emerald-700 mt-0.5">
                        {matchedBillsForDeleteList.matchedUcs.size}
                      </p>
                    </div>

                    <div className="bg-red-50 border border-red-200/80 rounded-2xl p-3.5 text-center">
                      <p className="text-[10px] font-bold text-red-600 uppercase tracking-wider">
                        Faturas a Excluir
                      </p>
                      <p className="text-xl font-bold text-red-700 mt-0.5">
                        {matchedBillsForDeleteList.bills.length}
                      </p>
                    </div>
                  </div>

                  {/* Matching list summary */}
                  {matchedBillsForDeleteList.matchedUcs.size > 0 && (
                    <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-3.5">
                      <p className="text-[11px] font-bold text-slate-700 mb-2">
                        UCs identificadas na base de dados:
                      </p>
                      <div className="flex flex-wrap gap-1.5 max-h-32 overflow-y-auto">
                        {Array.from(matchedBillsForDeleteList.matchedUcs).map((uc) => {
                          const count = matchedBillsForDeleteList.bills.filter((b) => b.uc === uc).length;
                          return (
                            <span
                              key={uc}
                              className="px-2.5 py-1 bg-white border border-slate-200 rounded-lg text-[11px] font-mono font-semibold text-slate-700 shadow-xs flex items-center gap-1.5"
                            >
                              <span>{uc}</span>
                              <span className="text-[10px] px-1.5 py-0.2 bg-red-100 text-red-700 rounded-full font-sans">
                                {count} fat.
                              </span>
                            </span>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* UCs not found notice */}
                  {matchedBillsForDeleteList.notFoundUcs.length > 0 && (
                    <div className="bg-amber-50/70 border border-amber-200/60 rounded-2xl p-3">
                      <p className="text-[11px] font-bold text-amber-800 mb-1">
                        {matchedBillsForDeleteList.notFoundUcs.length} UC(s) informada(s) não possuem faturas no sistema:
                      </p>
                      <p className="text-[10px] font-mono text-amber-700 break-words line-clamp-3">
                        {matchedBillsForDeleteList.notFoundUcs.join(", ")}
                      </p>
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* Footer Actions */}
            <div className="p-6 bg-slate-50 border-t border-slate-100 flex items-center justify-between">
              <button
                onClick={() => setIsDeleteByListModalOpen(false)}
                className="px-6 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-200 transition-colors"
              >
                Cancelar
              </button>

              <button
                onClick={handleDeleteByUcList}
                disabled={matchedBillsForDeleteList.bills.length === 0 || isDeletingByList}
                className={`flex items-center gap-2 px-7 py-3 rounded-xl text-xs font-bold text-white transition-all shadow-lg active:scale-95 ${
                  matchedBillsForDeleteList.bills.length === 0 || isDeletingByList
                    ? "bg-slate-300 shadow-none cursor-not-allowed opacity-60"
                    : "bg-red-600 hover:bg-red-700 shadow-red-600/20"
                }`}
              >
                {isDeletingByList ? (
                  <Loader2 size={16} className="animate-spin" />
                ) : (
                  <Trash2 size={16} />
                )}
                Excluir {matchedBillsForDeleteList.bills.length} Fatura(s) ({matchedBillsForDeleteList.matchedUcs.size} UCs)
              </button>
            </div>
          </motion.div>
        </div>
      )}

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
                <div
                  className={`w-14 h-14 rounded-2xl flex items-center justify-center mb-6 ${
                    confirmModalData.type === "danger"
                      ? "bg-red-50 text-red-600"
                      : "bg-blue-50 text-blue-600"
                  }`}
                >
                  {confirmModalData.type === "danger" ? (
                    <Trash2 size={28} />
                  ) : (
                    <AlertCircle size={28} />
                  )}
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
                    confirmModalData.type === "danger"
                      ? "bg-red-600 shadow-red-600/20 hover:bg-red-700"
                      : "bg-sanesul-primary shadow-sanesul-primary/20 hover:bg-sanesul-primary/90"
                  }`}
                >
                  {confirmModalData.isAlert ? "Entendido" : "Confirmar"}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
