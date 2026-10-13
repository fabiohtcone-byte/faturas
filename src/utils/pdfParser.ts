import * as pdfjsLib from "pdfjs-dist";
import pdfjsWorker from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { UC_MIGRATION_MAP } from "../uc_migration";

// Set the workerSrc to the local file
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfjsWorker;

export async function extractTextFromPdf(file: File | Blob): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ 
    data: arrayBuffer,
    disableFontFace: true, // Prevents loading fonts into memory (huge memory saver)
    useSystemFonts: true
  });
  const pdf = await loadingTask.promise;
  let fullText = "";

  try {
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items
        .map((item: any) => item.str)
        .join(" ");
      fullText += pageText + "\n";
      
      // Crucial for memory management with large batches
      if (typeof page.cleanup === "function") {
        page.cleanup();
      }
    }
  } finally {
    try {
      if (typeof loadingTask.destroy === "function") {
        await loadingTask.destroy();
      } else if (typeof (pdf as any).destroy === "function") {
        await (pdf as any).destroy();
      } else if (typeof pdf.cleanup === "function") {
        await pdf.cleanup();
      }
    } catch (e) {
      console.warn("Could not destroy PDF instance", e);
    }
  }

  return fullText;
}

const formatNumberToBR = (num: number) => {
  if (num === 0) return "0";
  return num.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
};

export function parseBillText(text: string): Record<string, string> {
  // A Regex-based parser for energy bills (Energisa / Elektro)
  // This is a naive implementation that will need tuning based on actual PDFs
  
  // Extractors
  const extract = (regex: RegExp, fallback = "") => {
    const match = text.match(regex);
    return match && match[1] ? match[1].trim() : fallback;
  };

  const cleanNumber = (val: string) => {
    if (!val) return "0";
    return val.replace(/[^\d,\.-]/g, "");
  };

  // Base fields
  // Priority 1: Explicit CÓDIGO DO CLIENTE / CÓDIGO DA UC / CÓDIGO DA UNIDADE CONSUMIDORA
  let uc = "";
  const codClienteMatch = text.match(/(?:CÓDIGO\s+DO\s+CLIENTE|CODIGO\s+DO\s+CLIENTE|CÓDIGO\s+DA\s+UC|CODIGO\s+DA\s+UC|CÓDIGO\s+DA\s+UNIDADE\s+CONSUMIDORA|CODIGO\s+DA\s+UNIDADE\s+CONSUMIDORA)[\s:]*([\d\.\-\/\s]+)/i);
  if (codClienteMatch && codClienteMatch[1]) {
    const cleaned = codClienteMatch[1].replace(/[^\d]/g, "");
    if (cleaned.length >= 8 && cleaned.length <= 13) {
      uc = cleaned;
    }
  }

  // Priority 2: Generic UC / Instalação / Seu Código / Unidade Consumidora
  if (!uc) {
    uc = extract(/(?:UC|Instalação|CÓDIGO DO CLIENTE|Seu C[oó]digo|Unidade Consumidora)[\s:]*([\d\.\-\/]+)/i);
    if (uc) uc = uc.replace(/[^\d]/g, "");
  }
  
  // Secondary: try MATRÍCULA / FATURAMENTO AGRUPADO but filter out billing references
  if (!uc) {
    const rawUc = extract(/(?:MATRÍCULA|FATURAMENTO AGRUPADO)[\s:]*([\d\.\-\/]+(?:\s+[\d\.\-\/]+)*)/i);
    if (rawUc) {
      // Filter out billing references like "3795488-2026-5-2" (contain year 202X)
      if (!/20[2-3]\d/.test(rawUc)) {
        uc = rawUc.replace(/[^\d]/g, "");
      }
    }
  }
  
  const rawConcessionaria = extract(/(ELEKTRO|ENERGISA|NEOENERGIA|ELEK[AB]|DANFELEK[AB])/i);
  let concessionaria = rawConcessionaria.toUpperCase();
  if (/NEOENERGIA|ELEK[AB]|DANFELEK[AB]/i.test(concessionaria)) {
    concessionaria = "ELEKTRO";
  }

  if (concessionaria === "ENERGISA" || text.includes("ENERGISA")) {
    concessionaria = "ENERGISA";
    
    // Strategy 1 (PRIMARY): Look for formatted Energisa UC pattern
    // Supports 2-dot (181.951.005-101 or 83.969.051-17) and 3-dot (2.822.635.051-30) formats
    // Also handles cases where pdf.js strips the dots but leaves the dash (e.g. 71800051-59)
    let formattedUc = "";
    const energisaUcMatch = text.match(/(\d{1,4}(?:[\.\s]*\d{1,3}){2,3}-\s*\d{1,3})/g);
    if (energisaUcMatch) {
      for (const match of energisaUcMatch) {
        let parsedUc = match.replace(/[^\d]/g, "");
        // Ignore Inscrição Estadual MS which starts with 28 and has 9 digits
        if (parsedUc.startsWith("28") && parsedUc.length === 9) continue;
        // Strip year prefix contamination: if UC has >12 digits and starts with
        // a year (2020-2039), the regex likely captured a nearby date as part of the UC
        if (parsedUc.length > 12 && /^20[2-3]\d/.test(parsedUc)) {
          parsedUc = parsedUc.substring(4);
        }
        // Valid ENERGISA UCs have 9-12 digits; skip anything outside this range
        if (parsedUc.length < 9 || parsedUc.length > 12) continue;
        // Prefer longest formatted match within valid range
        if (!formattedUc || parsedUc.length >= formattedUc.length) {
          formattedUc = parsedUc;
        }
      }
    }
    
    // Formatted UC always takes priority over base keyword extraction
    if (formattedUc && formattedUc.length >= 9) {
      uc = formattedUc;
    }
    
    // Strategy 2: If UC still seems incomplete, try to recover detached leading digits
    if (uc && uc.length >= 10 && uc.length <= 11) {
      const first6 = uc.substring(0, 6);
      try {
        const detachedRegex = new RegExp("(\\d{1,2})[\\s\\n]+(" + first6.split("").join("\\s*") + ")");
        const detachedMatch = text.match(detachedRegex);
        if (detachedMatch) {
          const candidate = detachedMatch[1] + uc;
          if (candidate.length > uc.length) {
            uc = candidate;
          }
        }
      } catch (e) {
        // regex construction failed, skip
      }
    }
    
    // Strategy 3: Direct search for long digit sequences (9-12 digits) ONLY if no UC found at all
    if (!uc || uc.length < 6) {
      // Find all standalone sequences of 9-12 digits
      const fallbackMatches = text.match(/(?<![\d\.,])\d{9,12}(?![\d\.,])/g);
      if (fallbackMatches && fallbackMatches.length > 0) {
        for (const num of fallbackMatches) {
          if (!uc || num.length > uc.length) {
            uc = num;
          }
        }
      }
    }
  }

  let cidade = extract(/([A-ZÀ-Ÿ\s]+)\s*\(AG:\s*\d+\)/i);
  if (cidade) {
    if (cidade.toUpperCase().includes("SANESUL")) {
      cidade = cidade.replace(/.*SANESUL\s*/i, "").trim();
    }
    // Caso ainda tenha capturado muitas coisas divididas por espaço duplo, pega a última parte
    const parts = cidade.split(/\s{2,}/);
    cidade = parts[parts.length - 1].trim();
  }
  let valorTotal = extract(/(?:Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro|Janeiro|Fevereiro|Março|Marco|Abril|[A-Za-z]+)\s*\/\s*\d{4}\s+(?:\d{2}\/\d{2}\/\d{4}\s+)?R\$\s*([\d\.,]+)/i)
    || extract(/\d{2}\/\d{2}\/\d{4}\s+R\$\s*([\d\.,]+)/i)
    || extract(/R\$\s*([\d\.,]+)\s*NOTA FISCAL/i)
    || extract(/TOTAL:\s*([\d\.,]+)/i)
    || extract(/(?:Total a Pagar|TOTAL:|Total R\$)[^\d]*([\d\.,]+)/i)
    || extract(/R\$\s*([\d\.,]+)\s*Nº FATURA/i);

  if (valorTotal) {
    const rawVal = parseFloat(valorTotal.replace(/\./g, "").replace(",", "."));
    if (!isNaN(rawVal) && rawVal > 0) {
      valorTotal = formatNumberToBR(rawVal);
    }
  }
  
  // Dates
  let mesReferencia = "";
  let anoLeitura = "";
  
  const MONTH_MAP_EXTRACT: Record<string, string> = {
    jan: "Janeiro", janeiro: "Janeiro",
    fev: "Fevereiro", fevereiro: "Fevereiro",
    mar: "Março", marco: "Março", março: "Março",
    abr: "Abril", abril: "Abril",
    mai: "Maio", maio: "Maio",
    jun: "Junho", junho: "Junho",
    jul: "Julho", julho: "Julho",
    ago: "Agosto", agosto: "Agosto",
    set: "Setembro", setembro: "Setembro",
    out: "Outubro", outubro: "Outubro",
    nov: "Novembro", novembro: "Novembro",
    dez: "Dezembro", dezembro: "Dezembro",
    "01": "Janeiro", "1": "Janeiro",
    "02": "Fevereiro", "2": "Fevereiro",
    "03": "Março", "3": "Março",
    "04": "Abril", "4": "Abril",
    "05": "Maio", "5": "Maio",
    "06": "Junho", "6": "Junho",
    "07": "Julho", "7": "Julho",
    "08": "Agosto", "8": "Agosto",
    "09": "Setembro", "9": "Setembro",
    "10": "Outubro",
    "11": "Novembro",
    "12": "Dezembro",
  };

  // Prioridade 1: Formato explícito ancorado por palavras-chave (Mês/Ano, Referência, REF, Competência)
  let matchNum = text.match(/(?:Referência:|M[êe]s\s*(?:\/|\s*e\s*|\s+)\s*Ano|REF:?|COMPET[ÊE]NCIA:?)\s*(?:VENCIMENTO\s*)?(\d{2}|Jan(?:eiro)?|Fev(?:ereiro)?|Mar(?:[çc]o)?|Abr(?:il)?|Mai(?:o)?|Jun(?:ho)?|Jul(?:ho)?|Ago(?:sto)?|Set(?:embro)?|Out(?:ubro)?|Nov(?:embro)?|Dez(?:embro)?)\s*(?:\/|-|\s)?\s*(202\d)/i);
  if (matchNum) {
    const rawM = matchNum[1].toLowerCase().trim();
    mesReferencia = MONTH_MAP_EXTRACT[rawM] || matchNum[1];
    anoLeitura = matchNum[2];
  } else {
    // Prioridade 2: Formato textual explícito com separador: "Setembro / 2026", "SET/2026", "Setembro de 2026"
    let matchMesAno = text.match(/(Janeiro|Fevereiro|Mar[çc]o|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro|Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez)\s*(?:\/|-|de|\s)\s*(202\d)/i);
    if (matchMesAno) {
      const rawM = matchMesAno[1].toLowerCase().trim();
      mesReferencia = MONTH_MAP_EXTRACT[rawM] || matchMesAno[1];
      anoLeitura = matchMesAno[2];
    } else {
      // Prioridade 3: Qualquer "MM/202X" perto do topo/vencimento/apresentação
      let matchFallback = text.match(/(?:DATA DE EMISSÃO|VENCIMENTO|Data de apresenta[\s\S]{1,3}o)[\s\S]{0,150}?(0[1-9]|1[0-2])\s*\/\s*(202\d)/i);
      if (matchFallback) {
        const rawM = matchFallback[1].toLowerCase().trim();
        mesReferencia = MONTH_MAP_EXTRACT[rawM] || matchFallback[1];
        anoLeitura = matchFallback[2];
      }
    }
  }

  const dataVencimento = extract(/(\d{2}\/\d{2}\/\d{4})\s*VENCIMENTO/i) || extract(/VENCIMENTO.*?(\d{2}\/\d{2}\/\d{4})/i) || extract(/(\d{2}\/\d{2}\/\d{4})\s*R\$/i);
  const numeroNotaFiscal = extract(/NOTA FISCAL Nº:?\s*([\d\.]+)/i);
  let modalidadeTarifaria = extract(/Classifica[\s\S]{1,3}o:\s*([^\/]+)/i);
  const subgrupo = extract(/Classifica[\s\S]{1,3}o:[^\/]+\/\s*([^\s]+)/i);
  const classificacaoLinha = extract(/Classifica[\s\S]{1,3}o:\s*([^\n\r]+)/i);
  const isVerdeGlobal = modalidadeTarifaria ? /VERDE/i.test(modalidadeTarifaria) : (/VERDE/i.test(text) || /HOR[ÁA]RIA\s+VERDE/i.test(text) || /Classifica[\s\S]{1,3}o:[^\n\r]*VERDE/i.test(text));
  let isGrupoB = (modalidadeTarifaria && modalidadeTarifaria.toUpperCase().includes("CONVENCIONAL")) || (subgrupo && subgrupo.toUpperCase().includes("B3"));
  if (!isGrupoB && classificacaoLinha && (classificacaoLinha.toUpperCase().includes("CONVENCIONAL") || classificacaoLinha.toUpperCase().includes("B3"))) {
    isGrupoB = true;
  }
  if (!isGrupoB && /Consumo em kWh[\s\S]{0,100}KWH/i.test(text) && !/AZUL|VERDE|A3|A4/i.test(modalidadeTarifaria || "")) {
      isGrupoB = true; // Forte indício de Energisa Grupo B (Grupo A não tem essa coluna exata)
  }

  // Consumo (Tratando B3 que não tem ponta/fora ponta explicito como verde)
  let consumoKwh = "";
  let consumoKwhPonta = extract(/Consumo[\s\S]{0,60}?(?<!Fora\s)(?<!F)Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let consumoKwhIntermediaria = extract(/Consumo[\s\S]{0,60}?Intermedi[aá]ri[ao][\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let consumoKwhForaPonta = extract(/Consumo[\s\S]{0,60}?Fora Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let valorConsumoKwhPonta = extract(/Valor Consumo[\s\S]{0,60}?(?<!Fora\s)(?<!F)Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let valorConsumoKwhForaPonta = extract(/Valor Consumo[\s\S]{0,60}?Fora Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let consumoKwhGrupoB = "";
  let valorConsumoKwhGrupoB = "";

  // Demandas e Reativos
  let demandaPontaKW: string | null = null;
  let demandaForaPontaKW: string | null = null;

  // 1. Padrão Invertido do PDF Energisa Page 2: "50 Fora Ponta: 47 KW Ponta:"
  const mInverted = text.match(/(\d+(?:[\.,]\d+)?)\s+Fora\s+Ponta:\s*(\d+(?:[\.,]\d+)?)\s+KW\s+Ponta/i);
  if (mInverted) {
    demandaPontaKW = mInverted[2];
    demandaForaPontaKW = mInverted[1];
  }

  // 2. Padrão Sequencial do PDF Energisa Page 2: "KW Ponta: 18,4 Fora Ponta: 30"
  if (!demandaPontaKW || !demandaForaPontaKW) {
    const mSeq = text.match(/KW\s*Ponta:\s*(\d+(?:[\.,]\d+)?)\s*(?:KW\s*)?Fora\s*Ponta:\s*(\d+(?:[\.,]\d+)?)/i);
    if (mSeq) {
      demandaPontaKW = mSeq[1];
      demandaForaPontaKW = mSeq[2];
    }
  }

  // 3. Padrão Page 1 / Grandezas Contratadas: "Demanda ponta - kW Demanda fora ponta - kW 36 36" ou "61"
  if (!demandaPontaKW || !demandaForaPontaKW) {
    const mSeqContratadas = text.match(/Demanda\s*ponta\s*-\s*kW\s*Demanda\s*fora\s*ponta\s*-\s*kW\s*(?:Demanda\s*TUSDG\s*-\s*kW\s*)?(\d+(?:[\.,]\d+)?)(?:\s+(\d+(?:[\.,]\d+)?))?/i);
    if (mSeqContratadas) {
      if (mSeqContratadas[2]) {
        demandaPontaKW = mSeqContratadas[1];
        demandaForaPontaKW = mSeqContratadas[2];
      } else {
        const val = mSeqContratadas[1];
        const isVerdeCheck = isVerdeGlobal;
        if (isVerdeCheck) {
          demandaPontaKW = "0";
          demandaForaPontaKW = val;
        } else {
          demandaPontaKW = val;
          demandaForaPontaKW = val;
        }
      }
    }
  }

  // 4. Outros padrões explícitos
  if (isVerdeGlobal) {
    demandaPontaKW = "0";
  } else if (!demandaPontaKW || demandaPontaKW === "0") {
    demandaPontaKW = extract(/Demanda\s*(?:de\s+)?ponta\s*(?:-\s*kW)?[\s:]+([\d\.,]+)/i)
      || extract(/KW\s*Ponta:\s*([\d\.,]+)/i)
      || extract(/Demanda\s+Contratada[^\d]*(?<!Fora\s)(?<!F)Ponta[^\d]*([\d\.,]+)/i)
      || extract(/([\d\.,]+)\s+Demanda ponta/i);
  }
  if (!demandaForaPontaKW || demandaForaPontaKW === "0") {
    demandaForaPontaKW = extract(/Demanda\s*(?:de\s+)?fora\s*ponta\s*(?:-\s*kW)?[\s:]+([\d\.,]+)/i)
      || extract(/(?:KW\s*)?Fora\s*Ponta:\s*([\d\.,]+)/i)
      || extract(/Demanda\s+Contratada[^\d]*(?:Fora Ponta|Fponta)[^\d]*([\d\.,]+)/i)
      || extract(/([\d\.,]+)\s+Demanda fora ponta/i);
  }
  let demandaTodosPeriodos = extract(/Demanda Todos os Períodos:\s*([\d\.,]+)/i) || "0";

  let demandaContratadaPonta = isVerdeGlobal ? "0,00" : (demandaPontaKW || extract(/(\d+(?:,\d+)?)\s*Demanda\s+ponta/i) || extract(/Demanda\s+ponta.*?[\D](\d+(?:,\d+)?)/i) || "0");
  let demandaContratadaForaPonta = demandaForaPontaKW || extract(/(\d+(?:,\d+)?)\s*Demanda\s+fora\s+ponta/i) || extract(/Demanda\s+fora\s+ponta.*?[\D](\d+(?:,\d+)?)/i) || "0";
  const getExplicitMedida = (type: string) => {
    const regex = /Demanda[\s\S]{0,100}?Medida[\s\S]{0,100}?[-\u2010-\u2015]?[\s\S]{0,30}?Ponta[^\d]{0,100}?(\d+(?:[\.,]\d+)*)/ig;
    let match;
    while ((match = regex.exec(text)) !== null) {
      let isFora = /Fora|F\s/i.test(match[0]);
      let val = match[1];
      if (val && (val.includes(',') || val.includes('.'))) {
        const decPart = val.includes(',') ? val.split(',')[1] : val.split('.')[1];
        if (decPart && decPart.length > 2) continue; // Ignora tarifas com 4 ou 6 casas decimais
      }
      const numVal = parseFloat(val.replace(/\./g, "").replace(",", "."));
      if (isNaN(numVal) || numVal < 1.0 || numVal === 2.5) continue;

      if (type === 'ponta' && !isFora) {
         return val;
      }
      if (type === 'fora ponta' && isFora) {
         return val;
      }
    }
    return null;
  };
  
  const directKwMedPonta = extract(/KW\s+Ponta\s+(?:[\d\.,]+\s+){6}([\d\.,]+)/i);
  const directKwMedForaPonta = extract(/KW\s+FPonta\s+(?:[\d\.,]+\s+){6}([\d\.,]+)/i);
  const directUltpPonta = extract(/ULTP\s+Ponta\s+(?:[\d\.,]+\s+){6}([\d\.,]+)/i);
  const directUltpForaPonta = extract(/ULTP\s+FPonta\s+(?:[\d\.,]+\s+){6}([\d\.,]+)/i);

  let demandaPotenciaMedidaPonta = directKwMedPonta || extract(/Demanda\s+de\s+Potência\s+Medida\s+-\s+Ponta[^\d]{0,50}?(\d+(?:,\d+)?)/i) || getExplicitMedida('ponta') || extract(/Demanda Medida.*Ponta[^\d]*([\d\.,]+)/i) || "0";
  let demandaPotenciaMedidaForaPonta = directKwMedForaPonta || extract(/Demanda\s+de\s+Potência\s+Medida\s+-\s+Fora\s+Ponta[^\d]{0,50}?(\d+(?:,\d+)?)/i) || getExplicitMedida('fora ponta') || extract(/Demanda Medida.*Fora Ponta[^\d]*([\d\.,]+)/i) || "0";
  // Demanda Pot.ncia N.o Consumida – limit to 30 chars gap to avoid matching far-away numbers
  let demandaPotenciaNaoConsumidaPonta = extract(/Demanda\s+Pot.ncia\s+N.o\s+Consumida\s*-\s*Ponta[^\d]{0,30}(\d+(?:[\.,]\d+)*)/i) || "0";

  let demandaPotenciaNaoConsumidaForaPonta = extract(/Demanda\s+Pot.ncia\s+N.o\s+Consumida\s*-\s*F(?:ora)?\s*Ponta[^\d]{0,30}(\d+(?:[\.,]\d+)*)/i) || "0";
  let ultrapassagemPontaKW = directUltpPonta || "0";
  let ultrapassagemForaPontaKW = directUltpForaPonta || "0";
  const demandaLines = text.match(/.*Demanda.*/ig) || [];
  let energiaReativaExcedPonta = extract(/Reativa Exced[^-]*-\s*(?<!F)(?<!Fora\s)Ponta[^\d]{0,30}([\d\.,]+)/i) || "0";
  let energiaReativaExcedFPonta = extract(/Reativa Exced[^-]*-\s*F(?:ora)?\s*Ponta[^\d]{0,30}([\d\.,]+)/i) || "0";
  let valorDemandaPotenciaMedidaPonta = "0";
  let valorDemandaPotenciaMedidaForaPonta = "0";
  let valorDemandaPotenciaNaoConsumidaPonta = "0";
  let valorDemandaPotenciaNaoConsumidaForaPonta = "0";
  let valorUltrapassagemPonta = "0";
  let valorUltrapassagemForaPonta = "0";
  let valorEnergiaReativaExcedPonta = "0";
  let valorEnergiaReativaExcedFPonta = "0";

  if (concessionaria === "ENERGISA" && !isGrupoB) {
    // Faturas Grupo A geralmente listam tarifas com 6 casas decimais
    let cleanTextForPairing = text;
    const historyIdx = text.search(/(HIST.RICO|CONSUMO\s+DOS\s+ÚLTIMOS|INDICADORES\s+DE\s+QUALIDADE|LIMITES\s+DE\s+CONTINUIDADE|FIC\s+DMIC\s+DICRI|CONSUMO\s+DOS\s+13)/i);
    if (historyIdx !== -1) {
        cleanTextForPairing = text.substring(0, historyIdx);
    }

    const allTariffMatches = Array.from(cleanTextForPairing.matchAll(/\d{1,3}[\.,]\d{4,6}/g)).map(m => parseFloat(m[0].replace(",", ".")));
    const tariffMatches = allTariffMatches.filter(t => t > 0.05 && t < 300);
    // Expandimos para \d{1,7} para suportar consumos > 99.999 kWh
    const numbers = Array.from(cleanTextForPairing.matchAll(/(?<![\d\.,])\d{1,7}(?:[\.,]\d{3})*[\.,]\d{2}(?![\d\.,])/g))
         .map(m => parseFloat(m[0].replace(/[^\d]/g, "")) / 100);

    if (tariffMatches.length > 0) {
      // Reset initial regex values - pair matching will set correct values
      consumoKwhPonta = null;
      consumoKwhForaPonta = null;
      valorConsumoKwhPonta = null;
      valorConsumoKwhForaPonta = null;
      energiaReativaExcedPonta = "0";
      energiaReativaExcedFPonta = "0";
      valorEnergiaReativaExcedPonta = "0";
      valorEnergiaReativaExcedFPonta = "0";

      interface ConsumoPair { q: number; v: number; t: number; error?: number; }
      const pairs: ConsumoPair[] = [];

      for (const t of tariffMatches) {
         if (t > 0.05 && t < 300) { // Tarifas variam de 0.1 a 150+ para Demandas
            for (const q of numbers) {
               if (q > 0.5) { // Ignora minúsculos para evitar ruído, mas pega pequenas não-consumidas
                  const expectedV = q * t;
                  for (const v of numbers) {
                     if (v > expectedV * 0.95 && v < expectedV * 1.05) {
                        const absDiff = Math.abs(v - expectedV);
                        const relDiff = absDiff / expectedV;
                  if (relDiff < 0.001 || absDiff < 2.0) {
                     if (!pairs.find(p => p.q === q && p.v === v)) {
                         pairs.push({ q, v, t, error: relDiff });
                     }
                  }
               }
            }
         }
      }
      }
      }

      pairs.sort((a, b) => (a.error || 0) - (b.error || 0));

      let ponta = null;
      let foraPonta = null;

      let qPontaRegexStr = extract(/(?<!Fora\s)(?<!F)Consumo em kWh - Ponta[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                           extract(/KWH\s+Ponta\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)/i) ||
                           extract(/Energia\s+ativa\s+em\s+kWh\s+(?:-\s*)?Ponta[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                           extract(/(?<!Fora\s)(?<!F)Ponta Energia ativa em kWh[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i);

      let qForaPontaRegexStr = extract(/(?<!Fora\s)(?<!F)Consumo em kWh - Fora Ponta[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                               extract(/Consumo em kWh - Fponta[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                               extract(/KWH\s+FPonta\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)/i) ||
                               extract(/Energia\s+ativa\s+em\s+kWh\s+(?:-\s*)?Fora\s+Ponta[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                               extract(/Fora Ponta Energia ativa em kWh[\s\S]{0,250}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i);

      // Ignore regex if we have valid mathematical pairs, because column-by-column OCR makes regex match wrong numbers
      if (pairs.length > 0) {
          qPontaRegexStr = null;
          qForaPontaRegexStr = null;
          ponta = null;
          foraPonta = null;

          // Consumo DEVE ter tarifa < 10.0 e quantidade > 5.0 para evitar ruídos minúsculos
          let validPairs = pairs.filter(p => p.t < 10.0 && p.t > 0.10 && p.q > 5.0);
          
          let findPairInBlock = (block: string) => {
              for (let p of validPairs) {
                  let qStr2 = p.q.toFixed(2).replace('.', ',');
                  let qStr1 = p.q.toFixed(1).replace('.', ',');
                  let qStr2Br = p.q.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
                  let qStr1Br = p.q.toLocaleString('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
                  if (block.includes(qStr2) || block.includes(qStr1) || block.includes(qStr2Br) || block.includes(qStr1Br)) {
                      return p;
                  }
              }
              return null;
          };

          // Tentativa 1: Tabela de Medidores (À prova de falhas de OCR colunar e line-bleeding)
          let getMedidorBlock = (textStr: string, startIndex: number): string => {
              let block = textStr.substring(startIndex, startIndex + 150);
              let nextStart = block.slice(30).search(/(Ponta|Fora|W00)/i);
              if (nextStart !== -1) {
                  return block.substring(0, nextStart + 30);
              }
              return block;
          };

          let pontaMedidorIndex = text.search(/(Ponta\s+Energia\s+ativa\s+em\s+kWh|Energia\s+ativa\s+em\s+kWh\s+Ponta)/i);
          if (pontaMedidorIndex !== -1) ponta = findPairInBlock(getMedidorBlock(text, pontaMedidorIndex));
          
          let foraPontaMedidorIndex = text.search(/(Fora\s*Ponta\s+Energia\s+ativa\s+em\s+kWh|Energia\s+ativa\s+em\s+kWh\s+Fora\s*Ponta)/i);
          if (foraPontaMedidorIndex !== -1) foraPonta = findPairInBlock(getMedidorBlock(text, foraPontaMedidorIndex));

          // Tentativa 2: Agrupamento Inteligente por Tarifa (Remove Injetada e prioriza maiores volumes)
          if (!ponta || !foraPonta) {
              let tariffGroups = new Map<number, typeof validPairs[0]>();
              for (let p of validPairs) {
                  let tKey = Math.round(p.t * 1000); // agrupa tarifas idênticas
                  let existing = tariffGroups.get(tKey);
                  // Mantemos apenas a MAIOR quantidade por tarifa (Isso elimina Energia Injetada, que tem a mesma tarifa do Consumo)
                  if (!existing || p.q > existing.q) {
                      tariffGroups.set(tKey, p);
                  }
              }
              
              let uniquePairs = Array.from(tariffGroups.values());
              // Ordena pelas maiores quantidades para eliminar Energia Reativa (que tem q minúsculo)
              uniquePairs.sort((a, b) => b.q - a.q); 
              
              let expectedPonta = /(?<!Fora\s)(?<!F)Consumo\s+(?:em\s+)?kWh\s*(?:-)?\s*Ponta/i.test(text);
              if (uniquePairs.length >= 2 && expectedPonta) {
                  let top2 = uniquePairs.slice(0, 2);
                  ponta = ponta || (top2[0].t > top2[1].t ? top2[0] : top2[1]);
                  foraPonta = foraPonta || (top2[0].t > top2[1].t ? top2[1] : top2[0]);
              } else if (uniquePairs.length >= 1) {
                  foraPonta = foraPonta || uniquePairs[0];
              }
          }
      }

      let matchPonta = text.match(/(?<!Fora\s)(?<!F)Consumo\s+(?:em\s+)?kWh\s*(?:-)?\s*Ponta[\sA-Za-z]{0,30}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)[\sA-Za-z]{0,20}?(?<![\d\,])\d{1,3},\d+[\sA-Za-z]{0,20}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                       text.match(/Energia\s+ativa\s+em\s+kWh\s+(?:-\s*)?Ponta[\sA-Za-z]{0,30}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)[\sA-Za-z]{0,20}?(?<![\d\,])\d{1,3},\d+[\sA-Za-z]{0,20}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i);

      let matchForaPonta = text.match(/(?<!Fora\s)(?<!F)Consumo\s+(?:em\s+)?kWh\s*(?:-)?\s*Fora\s*Ponta[\sA-Za-z]{0,30}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)[\sA-Za-z]{0,20}?(?<![\d\,])\d{1,3},\d+[\sA-Za-z]{0,20}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i) ||
                           text.match(/Energia\s+ativa\s+em\s+kWh\s+(?:-\s*)?Fora\s+Ponta[\sA-Za-z]{0,30}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)[\sA-Za-z]{0,20}?(?<![\d\,])\d{1,3},\d+[\sA-Za-z]{0,20}?(?<![\d\,])(\d{1,7}(?:\.\d{3})*,\d+)/i);

      // Se a matemática agrupou errado (ex: pegou Energia Injetada no lugar de Consumo porque o OCR falhou no Consumo), a regex salva.
      if (ponta && matchPonta) {
          let qRegex = parseFloat(matchPonta[1].replace(/\./g, "").replace(",", "."));
          if (Math.abs(qRegex - ponta.q) > 1.0) ponta = null;
      }
      if (foraPonta && matchForaPonta) {
          let qRegex = parseFloat(matchForaPonta[1].replace(/\./g, "").replace(",", "."));
          if (Math.abs(qRegex - foraPonta.q) > 1.0) foraPonta = null;
      }

      if (ponta) {
         consumoKwhPonta = formatNumberToBR(ponta.q);
         valorConsumoKwhPonta = formatNumberToBR(ponta.v);
      } else if (matchPonta) {
         consumoKwhPonta = matchPonta[1];
         valorConsumoKwhPonta = matchPonta[2];
      }
      
      // Tentativa 3: Extração Direta da Tabela de Medidores
      if (!consumoKwhPonta || consumoKwhPonta === "0") {
          let medPonta = extract(/(?:Ponta\s+Energia\s+ativa|Energia\s+ativa\s+em\s+kWh\s+Ponta)[\s\S]{0,120}?(?<![\d\,])(\d{1,7}(?:[\.,]\d{1,2})?)\s*$/im);
          if (medPonta) consumoKwhPonta = formatNumberToBR(parseFloat(medPonta.replace(/\./g, "").replace(",", ".")));
      }

      if (foraPonta) {
         consumoKwhForaPonta = formatNumberToBR(foraPonta.q);
         valorConsumoKwhForaPonta = formatNumberToBR(foraPonta.v);
      } else if (matchForaPonta) {
         consumoKwhForaPonta = matchForaPonta[1];
         valorConsumoKwhForaPonta = matchForaPonta[2];
      }
      
      if (!consumoKwhForaPonta || consumoKwhForaPonta === "0") {
          let medForaPonta = extract(/(?:Fora\s*Ponta\s+Energia\s+ativa|Energia\s+ativa\s+em\s+kWh\s+Fora\s*Ponta)[\s\S]{0,120}?(?<![\d\,])(\d{1,7}(?:[\.,]\d{1,2})?)\s*$/im);
          if (medForaPonta) consumoKwhForaPonta = formatNumberToBR(parseFloat(medForaPonta.replace(/\./g, "").replace(",", ".")));
      }
      
      let reativas = pairs.filter(p => {
          if (p === ponta || p === foraPonta) return false;
          if (p.t >= 1.5 || p.t < 0.1) return false;
          if (ponta && (Math.abs(p.v - ponta.v) < 1.0 || Math.abs(p.q - ponta.q) < 1.0 || Math.abs(p.v - ponta.q) < 1.0 || Math.abs(p.q - ponta.v) < 1.0)) return false;
          if (foraPonta && (Math.abs(p.v - foraPonta.v) < 1.0 || Math.abs(p.q - foraPonta.q) < 1.0 || Math.abs(p.v - foraPonta.q) < 1.0 || Math.abs(p.q - foraPonta.v) < 1.0)) return false;
          return true;
      });
      
      let reativaPonta = null;
      let reativaForaPonta = null;

      let qReativaPontaRegexStr = extract(/(?<!Fora\s)(?<!F)Ponta Energia reativa exced[^\d]*kVArh.*?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) || extract(/Energia Reativa Exced[^\d]{0,20}(?<!Fora\s)(?<!F)Ponta[^\d]{0,30}(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i);
      let qReativaForaPontaRegexStr = extract(/Fora Ponta Energia reativa exced[^\d]*kVArh.*?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) || extract(/Energia Reativa Exced[^\d]{0,20}(?:Fora Ponta|Fponta)[^\d]{0,30}(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i);

      if (qReativaPontaRegexStr) {
          let qRPonta = parseFloat(qReativaPontaRegexStr.replace(/\./g, "").replace(",", "."));
          reativaPonta = reativas.find(p => Math.abs(p.q - qRPonta) < 2.0);
      }
      if (qReativaForaPontaRegexStr) {
          let qRForaPonta = parseFloat(qReativaForaPontaRegexStr.replace(/\./g, "").replace(",", "."));
          reativaForaPonta = reativas.find(p => Math.abs(p.q - qRForaPonta) < 2.0);
      }

      if (!reativaPonta && !reativaForaPonta && reativas.length > 0) {
         let reativaTariff = null;
         for(let i=0; i<tariffMatches.length-1; i++) {
             if (tariffMatches[i] === tariffMatches[i+1] && tariffMatches[i] < 1.0) {
                 reativaTariff = tariffMatches[i];
                 break;
             }
         }

         let validReativas = reativaTariff ? reativas.filter(r => r.t === reativaTariff) : reativas.filter(r => r.q > 15.0);

         // Deduplicate by v: keep only the best-error pair for each unique value
         let seenV = new Map<number, typeof validReativas[0]>();
         for (const r of validReativas) {
             let existing = seenV.get(r.v);
             if (!existing || (r.error || 0) < (existing.error || 0)) {
                 seenV.set(r.v, r);
             }
         }
         validReativas = Array.from(seenV.values());

         validReativas.sort((a, b) => {
             let aIdx = text.indexOf(formatNumberToBR(a.v));
             let bIdx = text.indexOf(formatNumberToBR(b.v));
             if (aIdx === -1) aIdx = 999999;
             if (bIdx === -1) bIdx = 999999;
             return aIdx - bIdx;
         });
         
         if (validReativas.length >= 2) {
            // Só atribuir Ponta se o texto realmente menciona "Reativa Exced...Ponta" (e não apenas Fponta)
            let textHasReativaPonta = /Reativa Exced[^-]*-\s*(?!F)Ponta/i.test(text);
            if (textHasReativaPonta) {
                reativaPonta = validReativas[0];
                reativaForaPonta = validReativas[1];
            } else {
                // Texto só menciona Fponta → regex direta deveria ter encontrado.
                // Se não encontrou, valores são muito pequenos. Deixar zerado.
            }
         } else if (validReativas.length === 1) {
            reativaForaPonta = validReativas[0];
         }
      }

      if (reativaPonta) {
         energiaReativaExcedPonta = formatNumberToBR(reativaPonta.q);
         valorEnergiaReativaExcedPonta = formatNumberToBR(reativaPonta.v);
      }
      if (reativaForaPonta) {
         energiaReativaExcedFPonta = formatNumberToBR(reativaForaPonta.q);
         valorEnergiaReativaExcedFPonta = formatNumberToBR(reativaForaPonta.v);
      }
    }
  }

  // Se não achou ponta/fora ponta na extração inicial, tenta pegar o consumo geral (B3)
  // N.o aplicar para Energisa Grupo A (que usa pair matching) para evitar capturar valores errados
  if (!consumoKwhPonta && !consumoKwhForaPonta && concessionaria !== "ENERGISA") {
    const consumoGeral = extract(/Energia ativa em kWh\s+\d+\s+([\d\.,]+)/i) || extract(/Consumo kWh\s+[A-Z]{3}\/\d{2}\s+([\d\.,]+)/i);
    if (consumoGeral) {
      consumoKwhForaPonta = consumoGeral;
    }
  }

  if (concessionaria === "ENERGISA" && !isGrupoB) {
    const isVerde = isVerdeGlobal;

    let cleanTextForPairing = text;
    const historyIdx = text.search(/(CONSUMO\s+DOS\s+ÚLTIMOS|INDICADORES\s+DE\s+QUALIDADE|LIMITES\s+DE\s+CONTINUIDADE|FIC\s+DMIC\s+DICRI|CONSUMO\s+DOS\s+13)/i);
    if (historyIdx !== -1) {
        cleanTextForPairing = text.substring(0, historyIdx);
    }

    const tariffMatches = Array.from(cleanTextForPairing.matchAll(/(?<![\d\.,])\d{1,3}[\.,]\d{4,6}(?![\d\.,])/g)).map(m => parseFloat(m[0].replace(",", ".")));
    const numbers = Array.from(cleanTextForPairing.matchAll(/(?<![\d\.,])\d{1,7}(?:[\.,]\d{3})*[\.,]\d{2}(?![\d\.,])/g))
         .map(m => parseFloat(m[0].replace(/[^\d]/g, "")) / 100);

    let contPonta = parseFloat((demandaPontaKW || "0").replace(/\./g, "").replace(",", ".")) || 0;
    let contForaPonta = parseFloat((demandaForaPontaKW || "0").replace(/\./g, "").replace(",", ".")) || 0;
    const referenceDemand = contForaPonta > 0 ? contForaPonta : (contPonta > 0 ? contPonta : 0);

    // Add calculated implied q values to numbers
    const impliedNumbers = [...numbers];
    if (referenceDemand > 0) {
      for (const t of tariffMatches) {
        if (t >= 10.0 && t < 300.0) {
          for (const v of numbers) {
            if (v > 1.0) {
              const impliedQ = parseFloat((v / t).toFixed(2));
              const diffRatio = Math.abs(impliedQ - referenceDemand) / referenceDemand;
              if (diffRatio <= 0.6) { // allow up to 60% variation
                if (!impliedNumbers.includes(impliedQ)) {
                  impliedNumbers.push(impliedQ);
                }
              }
            }
          }
        }
      }
    }

    const demandaPairs: { q: number; v: number; t: number; error: number }[] = [];
    for (const t of tariffMatches) {
      if (t >= 10.0 && t < 300.0) {
        for (const q of impliedNumbers) {
          if (q > 2.0) { // Demanda é sempre maior que 2.0 kW (evita falsos positivos com PIS/COFINS como 0.38)
            const expectedV = q * t;
            for (const v of numbers) {
              if (v > expectedV * 0.95 && v < expectedV * 1.05) {
                const absDiff = Math.abs(v - expectedV);
                const relDiff = absDiff / expectedV;
                if (absDiff < 2.0 || relDiff < 0.01) {
                  if (!demandaPairs.find(p => p.q === q && p.v === v && p.t === t)) {
                    demandaPairs.push({ q, v, t, error: relDiff });
                  }
                }
              }
            }
          }
        }
      }
    }

    // Sort: Perfect matches first, then proximity to contracted demand, then error
    demandaPairs.sort((a, b) => {
      const isPerfectA = a.error < 0.001;
      const isPerfectB = b.error < 0.001;
      
      if (isPerfectA && !isPerfectB) return -1;
      if (!isPerfectA && isPerfectB) return 1;
      
      if (referenceDemand > 0) {
        const distA = Math.abs(a.q - referenceDemand);
        const distB = Math.abs(b.q - referenceDemand);
        if (Math.abs(distA - distB) > 0.01) {
          return distA - distB;
        }
      }
      
      return a.error - b.error;
    });

    const hasNcPonta = /Não\s+Consumida\s*-\s*Ponta|N\.o\s+Consumida\s*-\s*Ponta/i.test(cleanTextForPairing);
    const hasNcForaPonta = /Não\s+Consumida\s*-\s*F|N\.o\s+Consumida\s*-\s*F/i.test(cleanTextForPairing);
    const hasUltPonta = /Ultrapassagem\s*-\s*Ponta|ULTP\s*Ponta|Ultrap\s*-\s*Ponta/i.test(cleanTextForPairing);
    const hasUltForaPonta = /Ultrapassagem\s*-\s*F|ULTP\s*FPonta|Ultrap\s*-\s*F/i.test(cleanTextForPairing);

    if (isVerde) {
      contPonta = 0;
      demandaContratadaPonta = "0,00";
      demandaPotenciaMedidaPonta = "0,00";
      valorDemandaPotenciaMedidaPonta = "0,00";
      demandaPotenciaNaoConsumidaPonta = "0,00";
      valorDemandaPotenciaNaoConsumidaPonta = "0,00";
      ultrapassagemPontaKW = "0,00";
      valorUltrapassagemPonta = "0,00";

      if (demandaPairs.length > 0) {
        let knownQ = demandaPotenciaMedidaForaPonta !== "0" && demandaPotenciaMedidaForaPonta !== "0,00" ? parseFloat(demandaPotenciaMedidaForaPonta.replace(/\./g, "").replace(",", ".")) : 0;
        let matched = knownQ > 0 ? demandaPairs.find(p => Math.abs(p.q - knownQ) < 0.1) : null;

        if (matched) {
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(matched.v);
            
            if (hasUltForaPonta) {
                let bestUlt = demandaPairs.find(p => p !== matched && Math.abs(Math.max(p.q, matched.q) - Math.min(p.q, matched.q) - contForaPonta) < 1.0);
                if (bestUlt) {
                    ultrapassagemForaPontaKW = formatNumberToBR(bestUlt.q);
                    valorUltrapassagemForaPonta = formatNumberToBR(bestUlt.v);
                }
            } else if (hasNcForaPonta) {
                let bestNc = demandaPairs.find(p => p !== matched && Math.abs(p.q + matched.q - contForaPonta) < 1.0);
                if (bestNc) {
                    demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNc.q);
                    valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNc.v);
                }
            }
        } else {
            if (knownQ === 0) {
                demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
                valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
            }
        }
      }
    } else {
      // Modalidade AZUL
      const pontaPairs = demandaPairs.filter(p => p.t >= 60.0);
      const foraPairs = demandaPairs.filter(p => p.t < 60.0);

      // PONTA
      if (pontaPairs.length > 0) {
        let knownQ = demandaPotenciaMedidaPonta !== "0" && demandaPotenciaMedidaPonta !== "0,00" ? parseFloat(demandaPotenciaMedidaPonta.replace(/\./g, "").replace(",", ".")) : 0;
        let matched = knownQ > 0 ? pontaPairs.find(p => Math.abs(p.q - knownQ) < 0.1) : null;
        
        if (matched) {
            valorDemandaPotenciaMedidaPonta = formatNumberToBR(matched.v);
            
            if (hasUltPonta) {
                let bestUlt = pontaPairs.find(p => p !== matched && Math.abs(Math.max(p.q, matched.q) - Math.min(p.q, matched.q) - contPonta) < 1.0);
                if (bestUlt) {
                    ultrapassagemPontaKW = formatNumberToBR(Math.min(bestUlt.q, matched.q));
                    valorUltrapassagemPonta = formatNumberToBR(bestUlt.v); // Wait, bestUlt.v is the ultrapassagem value? Yes, bestUlt is the ultrapassagem pair!
                    // Actually, if matched is Medida, then bestUlt is Ultrapassagem!
                    ultrapassagemPontaKW = formatNumberToBR(bestUlt.q);
                    valorUltrapassagemPonta = formatNumberToBR(bestUlt.v);
                }
            } else if (hasNcPonta) {
                let bestNc = pontaPairs.find(p => p !== matched && Math.abs(p.q + matched.q - contPonta) < 1.0);
                if (bestNc) {
                    demandaPotenciaNaoConsumidaPonta = formatNumberToBR(bestNc.q);
                    valorDemandaPotenciaNaoConsumidaPonta = formatNumberToBR(bestNc.v);
                }
            }
        } else {
            if (knownQ === 0) {
                demandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].q);
                valorDemandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].v);
            }
        }
      }

      // FORA PONTA
      if (foraPairs.length > 0) {
        let knownQ = demandaPotenciaMedidaForaPonta !== "0" && demandaPotenciaMedidaForaPonta !== "0,00" ? parseFloat(demandaPotenciaMedidaForaPonta.replace(/\./g, "").replace(",", ".")) : 0;
        let matched = knownQ > 0 ? foraPairs.find(p => Math.abs(p.q - knownQ) < 0.1) : null;
        
        if (matched) {
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(matched.v);
            
            if (hasUltForaPonta) {
                let bestUlt = foraPairs.find(p => p !== matched && Math.abs(Math.max(p.q, matched.q) - Math.min(p.q, matched.q) - contForaPonta) < 1.0);
                if (bestUlt) {
                    ultrapassagemForaPontaKW = formatNumberToBR(bestUlt.q);
                    valorUltrapassagemForaPonta = formatNumberToBR(bestUlt.v);
                }
            } else if (hasNcForaPonta) {
                let bestNc = foraPairs.find(p => p !== matched && Math.abs(p.q + matched.q - contForaPonta) < 1.0);
                if (bestNc) {
                    demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNc.q);
                    valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNc.v);
                }
            }
        } else {
            if (knownQ === 0) {
                demandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].q);
                valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].v);
            }
        }
      }
    }
  }
  
  // Tributos (PIS, COFINS, ICMS)
  let pis = "0,00";
  let cofins = "0,00";
  let icms = "0,00";

  // 1. Grupo A / Layout Principal Energisa:
  // "PIS/ COFINS (R$)  <ValPIS> <ValCOFINS> <ValICMS> <AliqPIS> <AliqCOFINS> <AliqICMS>"
  // Ex: "PIS/ COFINS (R$) 894,61 4.120,80 11.350,42 1,6019 7,3786 17,00"
  const mPISCOFINS = text.match(/PIS\/\s*COFINS\s*\(R\$\)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+[\d\.,]+\s+[\d\.,]+\s+(?:17|25|27)/i);
  if (mPISCOFINS) {
    pis = mPISCOFINS[1];
    cofins = mPISCOFINS[2];
    icms = mPISCOFINS[3];
  } else {
    // 2. Grupo B com rodapé PIS ICMS COFINS:
    // "<ValPIS> <ValCOFINS> <AliqPIS> <AliqCOFINS> 17,00 <BasePIS> <BaseCOFINS> <BaseICMS> <ValICMS> PIS ICMS COFINS"
    // Ex: "1,38 6,36 0,9660 4,4493 17,00 143,02 143,02 172,32 29,29 PIS ICMS COFINS"
    const mGrupoB = text.match(/([\d\.,]+)\s+([\d\.,]+)\s+[\d\.,]+\s+[\d\.,]+\s+17(?:,00)?\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)\s+PIS\s+ICMS\s+COFINS/i);
    if (mGrupoB) {
      pis = mGrupoB[1];
      cofins = mGrupoB[2];
      icms = mGrupoB[3];
    } else {
      // 3. Fallback genérico para 3 valores antes de alíquotas com 17% ou 25%
      const mGeneric = text.match(/([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s+\d{1,2}(?:,\d+)?\s+\d{1,2}(?:,\d+)?\s+(?:17(?:,00)?|25(?:,00)?)\s+(?:TOTAL|TOTAL:|PIS|ICMS|Base|Consulte)/i);
      if (mGeneric) {
        pis = mGeneric[1];
        cofins = mGeneric[2];
        icms = mGeneric[3];
      } else {
        // 4. Linha a linha explícita
        const mPis = text.match(/PIS\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)/i);
        const mCofins = text.match(/COFINS\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)/i);
        const mIcms = text.match(/ICMS\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)/i);
        if (mPis && mCofins && mIcms) {
          pis = mPis[3];
          cofins = mCofins[3];
          icms = mIcms[3];
        } else {
          // Fallback genérico para outros layouts ou concessionárias
          icms = extract(/ICMS\s*\(R\$\).*?\d+\,\d+\s+([\d\.,]+)/i) || extract(/TOTAL:.*?\n.*?([\d\.,]+)\s+PIS/i) || "0,00";
          pis = extract(/PIS\s*\(R\$\).*?\d+\,\d+\s+([\d\.,]+)/i) || "0,00";
          cofins = extract(/COFINS\s*\(R\$\).*?\d+\,\d+\s+([\d\.,]+)/i) || "0,00";
        }
      }
    }
  }

  let cip = extract(/CONT\.IL\.PUB-CIP[^\d]*R\$[\s]*([\d\.,]+)/i) || extract(/Iluminação Pública[^\d]*R\$[\s]*([\d\.,]+)/i) || "0";
  // O Leitor agrupa o valor do CIP antes do mês de referência para a Energisa
  const energisaCipMatch = text.match(/[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)\s+[\d\.,]+\s+[A-Za-z]+\s*\/\s*\d{4}/i);
  if (concessionaria === "ENERGISA" && energisaCipMatch) {
    cip = energisaCipMatch[1];
  }
  // Fallback Energisa: CIP é isento de tributos (COFINS=0,00) e aparece antes da linha "Demanda"
  if (concessionaria === "ENERGISA" && cip === "0" && /CIP/i.test(text)) {
    const cipFallbackWithNegative = text.match(/([\d\.,]+)\s+0,00\s+-[\d\.,]+\s+0,00\s+[\d\.,]+\s+Demanda/i);
    const cipFallbackNormal = text.match(/(?<!-)(?<![\d\.,])([\d\.,]+)\s+0,00\s+[\d\.,]+\s+Demanda/i);
    if (cipFallbackWithNegative) {
      cip = cipFallbackWithNegative[1];
    } else if (cipFallbackNormal) {
      cip = cipFallbackNormal[1];
    }
  }

  // Extração específica para ELEKTRO
  if (concessionaria === "ELEKTRO") {

    // ── Identificar Grupo B (Convencional / B3) ─────────────────────────────
    const isElektroGrupoB =
      /CONVENCIONAL/i.test(text) ||
      /B3/i.test(text) ||
      /Consumo em kWh/i.test(text) ||
      !/AZUL|VERDE|A3|A4/i.test(modalidadeTarifaria || "");
    if (isElektroGrupoB) {
      isGrupoB = true;
      modalidadeTarifaria = modalidadeTarifaria || "CONVENCIONAL";
    }

    // ── UC (Código de instalação) ────────────────────────────────────────────
    const mUcElektro =
      extract(/CÓDIGO\s+DA\s+INSTALAÇÃO[^\d]*([\d]+)/i) ||
      extract(/Seu\s+Código[^\d]*([\d]+)/i) ||
      extract(/N[°º]\s+DA\s+INSTALA[ÇC][ÃA]O[^\d]*([\d]+)/i) ||
      extract(/INSTALAÇÃO[^\d]*([\d]{6,13})/i);
    if (mUcElektro) uc = mUcElektro.replace(/\D/g, "");

    // Reseta valores sujos de fallbacks anteriores
    consumoKwhPonta = "";
    consumoKwhForaPonta = "";
    valorConsumoKwhPonta = "";
    valorConsumoKwhForaPonta = "";
    demandaPontaKW = "";
    demandaForaPontaKW = "";
    consumoKwhGrupoB = "";
    valorConsumoKwhGrupoB = "";

    // ── Consumo e Demanda Elektro ─────────────────────────────────────────────
    const isGrupoA = /HORÁRIA\s+(?:VERDE|AZUL)|CONSUMO\s+PONTA/i.test(text);

    if (isGrupoA) {
      isGrupoB = false;

      // CONSUMO PONTA TE kWh 1.270 0,603898 766,95
      const mPonta = text.match(/CONSUMO\s+PONTA\s+TE[^\d]+([\d\.,]+)\s+[\d\.,]+\s+([\d\.,]+)/i);
      if (mPonta) {
        let qVal = mPonta[1].replace(/\./g, ''); // 1.270 -> 1270
        let vVal = mPonta[2]; // 766,95
        consumoKwhPonta = formatNumberToBR(parseFloat(qVal));
        valorConsumoKwhPonta = vVal;
      }

      // CONSUMO FORA PONTA TE kWh 8.515 0,374036 3.184,92
      const mForaPonta = text.match(/CONSUMO\s+FORA\s+PONTA\s+TE[^\d]+([\d\.,]+)\s+[\d\.,]+\s+([\d\.,]+)/i);
      if (mForaPonta) {
        let qVal = mForaPonta[1].replace(/\./g, ''); // 8.515 -> 8515
        let vVal = mForaPonta[2]; // 3.184,92
        consumoKwhForaPonta = formatNumberToBR(parseFloat(qVal));
        valorConsumoKwhForaPonta = vVal;
      }

      // Demanda Todos os Períodos (vai exclusivamente para demandaForaPontaKW)
      const mDemanda = text.match(/Demanda\s+Todos\s+os\s+Per[íi]odos:\s*([\d\.,]+)/i) || 
                       text.match(/DEMANDA\s+TUSD[^\d]*([\d\.,]+)/i);
      if (mDemanda) {
        let dVal = mDemanda[1];
        if (!dVal.includes(',')) {
          dVal = formatNumberToBR(parseFloat(dVal));
        }
        demandaForaPontaKW = dVal;
      }
    } else {
      // Pode aparecer como "CONSUMO TE kWh 50", "CONSUMO TE kWh 10.640"
      const extractConsumoElektro = (): string | null => {
        const match = text.match(/CONSUMO\s+(?:TE|TUSD)[^\d]*([\d\.,]+)/i) ||
                      text.match(/CONSUMO\s+kWh[^\d]*([\d\.,]+)/i) ||
                      text.match(/Energia\s+Ativa\s+kWh[^\d]*([\d\.,]+)/i);
        if (match && match[1]) {
          let valStr = match[1];
          if (valStr.includes('.') && !valStr.includes(',')) {
            valStr = valStr.replace(/\./g, '');
          } else if (valStr.includes(',')) {
            valStr = valStr.replace(/\./g, '').replace(',', '.');
          }
          const rawKwh = parseFloat(valStr);
          if (!isNaN(rawKwh) && rawKwh > 0) {
            return formatNumberToBR(rawKwh);
          }
        }
        return null;
      };

      const cElektro = extractConsumoElektro();
      if (cElektro) {
        consumoKwhGrupoB = cElektro;
        isGrupoB = true; 
      }
    }

    // ── Cidade ───────────────────────────────────────────────────────────────
    // Procura o padrão de CEP associado à cidade do cliente para evitar a sede da Elektro (Três Lagoas)
    const cidadeElektroMatch = text.match(/([A-ZÀ-Ÿ\s]+)\s*\([A-Z]{2}\)\s*-\s*[A-Z]{2}\s*(?:-\s*)?CEP\s*[\d.-]+/i) ||
                               text.match(/([A-ZÀ-Ÿ\s]+)\s*\([A-Z]{2}\)\s*-\s*[A-Z]{2}\s*(?:-|CNPJ|IE|INSC|CEP)/i);
    if (cidadeElektroMatch) {
      const rawCidade = cidadeElektroMatch[1].trim().toUpperCase();
      
      // Identifica com precisão as cidades atendidas pela Elektro em MS
      if (rawCidade.includes("SELVIRIA") || rawCidade.includes("SELVÍRIA")) {
        cidade = "SELVIRIA";
      } else if (rawCidade.includes("TRES LAGOAS") || rawCidade.includes("TRÊS LAGOAS")) {
        cidade = "TRES LAGOAS";
      } else if (rawCidade.includes("BRASILANDIA") || rawCidade.includes("BRASILÂNDIA")) {
        cidade = "BRASILANDIA";
      } else if (rawCidade.includes("SANTA RITA DO PARDO")) {
        cidade = "SANTA RITA DO PARDO";
      } else if (rawCidade.includes("ANAURILANDIA") || rawCidade.includes("ANAURILÂNDIA")) {
        cidade = "ANAURILANDIA";
      } else {
        const parts = cidadeElektroMatch[1].trim().split(/\s{2,}/);
        cidade = parts[parts.length - 1];
      }
    }

    // ── Modalidade tarifária ─────────────────────────────────────────────────
    const modElektroMatch = text.match(/(HORÁRIA\s+VERDE|HORÁRIA\s+AZUL|CONVENCIONAL\s*\/\s*BIFASICO|CONVENCIONAL\s*\/\s*TRIFASICO|CONVENCIONAL\s*\/\s*MONOFASICO|CONVENCIONAL)/i);
    if (modElektroMatch) modalidadeTarifaria = modElektroMatch[1];

    // ── CIP ──────────────────────────────────────────────────────────────────
    const cipMatch =
      extract(/COBRANCA ILUM PUBLICA[^\d]+([\d]+\,[\d]{2})/i) ||
      extract(/ILUM PUBLICA[^\d]+([\d]+\,[\d]{2})/i) ||
      extract(/CONTRIB\.\s*ILUM\s*PUB[^\d]+([\d]+\,[\d]{2})/i);
    if (cipMatch) cip = cipMatch;

    // ── Tributos Elektro ──────────────────────────────────────────────────────
    // A extração foca no padrão exato da linha de imposto para evitar cabeçalhos.
    // Layout padrão: ICMS 51,77 17% 8,79
    const extractTributoSeguro = (nome: string): string => {
      // 1. Tenta o padrão completo: NOME + Base + Alíquota + Valor
      const regex3 = new RegExp('\\b' + nome + '\\b\\s+([\\d\\.]+,\\d{2})\\s+([\\d\\.,]+%?)\\s+([\\d\\.]+,\\d{2})', 'i');
      const m3 = text.match(regex3);
      if (m3 && m3[3]) return m3[3];

      // 2. Tenta o padrão sem Alíquota: NOME + Base + Valor
      const regex2 = new RegExp('\\b' + nome + '\\b\\s+([\\d\\.]+,\\d{2})\\s+([\\d\\.]+,\\d{2})', 'i');
      const m2 = text.match(regex2);
      if (m2 && m2[2]) return m2[2];

      // 3. Fallback: Procura o NOME e olha apenas os próximos caracteres sem letras
      // Como é um loop global (ig), ele pula o cabeçalho (que tem letras logo depois) e acha a linha certa.
      const regexFallback = new RegExp('\\b' + nome + '\\b([^a-zA-Z]{1,40})', 'ig');
      let match;
      while ((match = regexFallback.exec(text)) !== null) {
        const window = match[1];
        const valores = window.match(/[\d\.]+,\d{2}/g);
        if (valores && valores.length > 0) {
           return valores[valores.length - 1]; // Retorna o último (Valor)
        }
      }

      return "0,00";
    };

    // Reseta valores sujos de fallbacks anteriores
    icms = "0,00";
    cofins = "0,00";
    pis = "0,00";

    // 1. Prioridade: Leitura em Colunas (O PDFjs-dist extrai essa tabela de cima para baixo na maioria das faturas)
    // Se o texto vier como "ICMS COFINS PIS 51,77 42,98 42,98 17% 4,31% 0,93% 8,79 1,85 0,40"
    const mCol = text.match(/ICMS\s+COFINS\s+PIS\s+([\d\.]+,\d{2})\s+([\d\.]+,\d{2})\s+([\d\.]+,\d{2})\s+[\d\.,]+%?\s+[\d\.,]+%?\s+[\d\.,]+%?\s+([\d\.]+,\d{2})\s+([\d\.]+,\d{2})\s+([\d\.]+,\d{2})/i);
    if (mCol) {
      icms = mCol[4];
      cofins = mCol[5];
      pis = mCol[6];
    } else {
      // 2. Fallback: Leitura Horizontal por tributo individual
      const icmsV  = extractTributoSeguro('ICMS');
      const cofinsV = extractTributoSeguro('COFINS');
      const pisV   = extractTributoSeguro('PIS');

      if (icmsV  !== "0,00") icms   = icmsV;
      if (cofinsV !== "0,00") cofins = cofinsV;
      if (pisV   !== "0,00") pis    = pisV;
    }

    // ── Valor Total ──────────────────────────────────────────────────────────
    const mTotal =
      extract(/Total\s+R\$[^\d]*([\d\.]+,\d{2})/i) ||
      extract(/TOTAL\s+A\s+PAGAR[^\d]*([\d\.]+,\d{2})/i) ||
      extract(/(?:Jan|Fev|Mar|Abr|Mai|Jun|Jul|Ago|Set|Out|Nov|Dez|Janeiro|Fevereiro|Mar[çc]o|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro)\/\d{4}\s+R\$\s*([\d\.]+,\d{2})/i) ||
      extract(/\d{2}\/\d{2}\/\d{4}\s+R\$\s*([\d\.]+,\d{2})/i) ||
      extract(/R\$\s*([\d\.]+,\d{2})\s*(?:NOTA FISCAL|N[°º]\s*FATURA)/i);
    if (mTotal) {
      const raw = parseFloat(mTotal.replace(/\./g, "").replace(",", "."));
      if (!isNaN(raw) && raw > 0) valorTotal = formatNumberToBR(raw);
    }
  }

  // Extração de GDI (Geração Distribuída Injetada) para Energisa / Grupo A / Grupo B
  let energiaAtvInjetadaGDIOUC = 0;
  let valorEnergiaAtvInjetadaGDIOUC = 0;
  let energiaAtvInjetadaGDIMUC = 0;
  let valorEnergiaAtvInjetadaGDIMUC = 0;

  const hasGdiInDoc = /GDI|GD[_\s]?I|Energia\s+(?:Atv|Ativa)?\s*Injetada|Injetada\s+GDI/i.test(text);

  if ((concessionaria === "ENERGISA" || /ENERGISA/i.test(text) || !concessionaria) && hasGdiInDoc) {
    // Normalização inicial para desmembrar números colados e unificar sinais de menos
    const normalizedText = text
      .replace(/[\u2010-\u2015\u2212\uFE63\uFF0D]/g, "-")
      .replace(/-\s+(\d)/g, "-$1")
      .replace(/(\d{1,2},\d{6})(\d{1,7}(?:\.\d{3})*,\d{2})/g, "$1 $2")
      .replace(/(KWH|kWh)(\d)/gi, "$1 $2");

    // Identificar contagens e tipos de GDI presentes no documento
    const oUcMatches = normalizedText.match(/GDI\s*oUC|GDI\s*OUC|Injetada\s+GDI\s*oUC|GD[_\s]?I\s*oUC/gi) || [];
    const mUcMatches = normalizedText.match(/GDI\s*mUC|GDI\s*MUC|Injetada\s+GDI\s*mUC|GD[_\s]?I\s*mUC/gi) || [];
    const genericGdiMatches = normalizedText.match(/Energia\s+(?:Atv|Ativa)?\s*Injetada\s+GDI|Injetada\s+GDI|GD[_\s]?I\s*oUC|GD[_\s]?I\s*mUC/gi) || [];
    
    const oUcCount = oUcMatches.length;
    const mUcCount = mUcMatches.length;
    const isExplicitOUC = oUcCount > 0 || /GDI\s*oUC/i.test(normalizedText);
    const isExplicitMUC = mUcCount > 0 || /GDI\s*mUC/i.test(normalizedText);

    let foundGdiItems: { type: 'OUC' | 'MUC'; kwh: number; valor: number; rawText?: string }[] = [];

    // ── Estratégia 1: Regex Direta de Linhas (captura TODAS as ocorrências de mUC e oUC) ────────
    // Ex: "Energia Atv Injetada GDI mUC 01/2026 mPT KWH 1.200,00 0,907120 -1.088,54"
    // Ex: "Energia Atv Injetada GDI mUC KWH 3.451,00 0,907120 -3.130,52"
    const lineRegexPatterns = [
      /(?:Energia\s+(?:Atv|Ativa)?\s*)?Injetada\s+(GDI\s*(?:oUC|mUC|OUC|MUC)?[^\n\r]*?)(?:KWH|kWh)?\s+([\d\.]+,\d{2})\s+[\d\.]+,\d+\s+(-[\d\.]+,\d{2})/ig,
      /(?:Energia\s+(?:Atv|Ativa)?\s*)?Injetada\s+(GDI\s*(?:oUC|mUC|OUC|MUC)?[^\n\r]*?)(?:KWH|kWh)?\s+([\d\.]+,\d{2})[^\n\r]*?(-[\d\.]+,\d{2})/ig,
      /(GDI\s*(?:oUC|mUC|OUC|MUC)[^\n\r]*?)(?:KWH|kWh)?\s+([\d\.]+,\d{2})\s+[\d\.]+,\d+\s+(-[\d\.]+,\d{2})/ig,
      /(GDI\s*(?:oUC|mUC|OUC|MUC)[^\n\r]*?)(?:KWH|kWh)?\s+([\d\.]+,\d{2})[^\n\r]*?(-[\d\.]+,\d{2})/ig,
      /(GDI\s*(?:oUC|mUC|OUC|MUC)[^\n\r]*?)(-[\d\.]+,\d{2})[^\n\r]*?([\d\.]+,\d{2})/ig,
    ];

    for (const regex of lineRegexPatterns) {
      let m;
      while ((m = regex.exec(normalizedText)) !== null) {
        const fullMatch = m[0];
        const labelPart = (m[1] || fullMatch).toUpperCase();
        const type: 'OUC' | 'MUC' = labelPart.includes("MUC") ? "MUC" : "OUC";
        
        let kwhStr = m[2];
        let valStr = m[3];
        
        if (kwhStr && valStr) {
          if (kwhStr.startsWith("-") && !valStr.startsWith("-")) {
            const temp = kwhStr;
            kwhStr = valStr;
            valStr = temp;
          }

          const kwh = parseFloat(kwhStr.replace(/\./g, "").replace(",", "."));
          const rawV = valStr.replace(/[^\d\.,-]/g, "");
          const valor = -Math.abs(parseFloat(rawV.replace(/\./g, "").replace(",", ".")));

          if (!isNaN(kwh) && kwh > 0 && !isNaN(valor) && valor < 0) {
            // Evita duplicatas pelo índice aproximado ou valor exato já registrado
            const alreadyAdded = foundGdiItems.some(
              item => Math.abs(item.kwh - kwh) < 0.01 && Math.abs(item.valor - valor) < 0.01 && item.type === type
            );
            if (!alreadyAdded) {
              foundGdiItems.push({ type, kwh, valor, rawText: fullMatch });
            }
          }
        }
      }
    }

    // ── Estratégia 2: Extração Estruturada por Correlação Matemática (DANF3E Colunar) ───
    const expectedGdiCount = Math.max(1, oUcCount + mUcCount, genericGdiMatches.length);
    if (foundGdiItems.length < expectedGdiCount) {
      // 1. Coleta todas as tarifas (4 a 6 casas decimais)
      const tariffs = Array.from(new Set(
        Array.from(normalizedText.matchAll(/(?<![\d,])\d{1,3},\d{4,6}(?![\d,])/g))
          .map(m => parseFloat(m[0].replace(",", ".")))
      )).filter(t => t > 0.05 && t < 10.0);

      // 2. Coleta todos os números positivos formatados como quantidade/valor
      const allPositiveNumbers = Array.from(
        normalizedText.matchAll(/(?<!-)(?<![\d,])(\d{1,7}(?:\.\d{3})*,\d{2})(?![\d,])/g)
      ).map(m => ({
        str: m[1],
        num: parseFloat(m[1].replace(/\./g, "").replace(",", "."))
      })).filter(n => n.num > 1.0);

      // 3. Coleta todos os valores negativos (ex: -3.130,52, -100,19)
      const allNegativeMatches = Array.from(
        normalizedText.matchAll(/(-(?:\d{1,7}(?:\.\d{3})*|\d+),\d{2})/g)
      ).map(m => ({
        str: m[1],
        val: -Math.abs(parseFloat(m[1].replace(/\./g, "").replace(",", ".")))
      })).filter(n => Math.abs(n.val) > 0.5);

      // Agrupa valores negativos excluindo os tributos de linha que são ~2% a ~15% de um valor maior (ex: PIS/COFINS)
      const mainNegatives: { val: number; matchedKwh?: number }[] = [];
      for (const neg of allNegativeMatches) {
        const absVal = Math.abs(neg.val);
        const isTaxOfAnother = allNegativeMatches.some(other => {
          if (other === neg) return false;
          const otherAbs = Math.abs(other.val);
          if (otherAbs <= absVal) return false;
          const ratio = absVal / otherAbs;
          return ratio >= 0.02 && ratio <= 0.15;
        });

        if (!isTaxOfAnother) {
          mainNegatives.push({ val: neg.val });
        }
      }

      // 4. Para cada valor negativo principal, encontra a quantidade (kWh) e tarifa correspondentes
      for (const mainNeg of mainNegatives) {
        const absV = Math.abs(mainNeg.val);
        let bestKwh: number | null = null;
        let bestDiff = 999999;

        // Procura entre as quantidades positivas presentes no PDF qual multiplicada pela tarifa dá absV
        for (const pos of allPositiveNumbers) {
          for (const t of tariffs) {
            const expectedV = pos.num * t;
            const diff = Math.abs(expectedV - absV);
            if (diff < 2.0 && diff < bestDiff) {
              bestDiff = diff;
              bestKwh = pos.num;
            }
          }
        }

        // Se não achou quantidade exata no texto, calcula pelo divisor da tarifa mais próxima de número inteiro
        if (bestKwh === null) {
          for (const t of tariffs) {
            const calculatedKwh = absV / t;
            const roundedKwh = Math.round(calculatedKwh);
            const diffFromInt = Math.abs(calculatedKwh - roundedKwh);
            if (diffFromInt < 0.1) {
              bestKwh = roundedKwh;
              break;
            }
          }
        }

        if (bestKwh !== null && bestKwh > 0) {
          mainNeg.matchedKwh = bestKwh;
        }
      }

      // 5. Atribui aos itens encontrados
      for (let i = 0; i < mainNegatives.length; i++) {
        const neg = mainNegatives[i];
        if (!neg.matchedKwh) continue;

        let type: 'OUC' | 'MUC' = "OUC";
        if (isExplicitMUC && !isExplicitOUC) {
          type = "MUC";
        } else if (isExplicitOUC && !isExplicitMUC) {
          type = "OUC";
        } else if (i < oUcCount) {
          type = "OUC";
        } else {
          type = "MUC";
        }

        if (!foundGdiItems.some(item => Math.abs(item.kwh - neg.matchedKwh!) < 0.01 && Math.abs(item.valor - neg.val) < 0.01)) {
          foundGdiItems.push({
            type,
            kwh: neg.matchedKwh,
            valor: neg.val
          });
        }
      }
    }

    // ── Consolidação e SOMA DE TODOS OS VALORES de GDI mUC e GDI oUC ───────────
    for (const item of foundGdiItems) {
      if (item.type === "MUC") {
        energiaAtvInjetadaGDIMUC += item.kwh;
        valorEnergiaAtvInjetadaGDIMUC += item.valor;
      } else {
        energiaAtvInjetadaGDIOUC += item.kwh;
        valorEnergiaAtvInjetadaGDIOUC += item.valor;
      }
    }
  }

  // ── Extração de Saldo Acumulado GD e Totais de Energia Injetada/Compensada ──
  let saldoAcumuladoGD = "";
  let energiaInjetadaKwh = "";
  let energiaCompensadaKwh = "";

  // 1. Extração de Saldo Acumulado GD
  const saldoPatterns = [
    /(?:Saldo\s+(?:Acumulado\s+GD|GD\s+Acumulado|Acumulado|Atual|Restante|do\s+Ciclo|para\s+o\s+pr[óo]ximo\s+m[êe]s|para\s+o\s+m[êe]s\s+seguinte|p\/\s*pr[óo]x(?:imo)?\s*m[êe]s|p\/\s*m[êe]s\s*seg(?:uinte)?|a\s+expirar|Total\s+Acumulado|a\s+Compensar|Dispon[íi]vel|Remanescente)[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i,
    /(?:SALDO\s+(?:ACUMULADO|ATUAL|RESTANTE|DE\s+ENERGIA|DO\s+CICLO|A\s+COMPENSAR|DISPON[ÍI]VEL|DO\s+M[ÊE]S\s+SEGUINTE)[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i,
    /(?:Saldo\s+(?:acumulado|atual)\s+(?:do\s+)?ciclo\s+(?:anterior|atual)[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i,
    /(?:Saldo\s+(?:Acumulado|Atual|a\s+Compensar|Dispon[íi]vel)\s*\(kWh\)[^\d:\n\r]{0,35}[:\s]*)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i,
    /(?:Saldo\s+de\s+(?:Energia\s+Acumulada|Cr[ée]dito[s]?|GD)[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i,
    /(?:Cr[ée]dito[s]?\s+Acumulado[s]?(?:\s+GD)?[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i,
    /(?:SALDO\s+ACUMULADO\s*\(KWH\)[\s\S]{0,30}?)([\d\.]+(?:,\d{2})?)/i,
    /(?:SALDO\s+ATUAL\s*\(KWH\)[\s\S]{0,30}?)([\d\.]+(?:,\d{2})?)/i,
    /(?:SALDO\s+A\s+COMPENSAR\s*\(KWH\)[\s\S]{0,30}?)([\d\.]+(?:,\d{2})?)/i,
  ];

  for (const pattern of saldoPatterns) {
    const m = text.match(pattern);
    if (m && m[1] && m[1].trim() !== "") {
      saldoAcumuladoGD = m[1].trim().replace(/\./g, "");
      break;
    }
  }

  // 2. Extração de Energia Injetada e Compensada
  const injetadaMatch =
    text.match(/(?:Energia\s+(?:Ativa\s+)?Injetada\s*(?:\(kWh\))?[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i) ||
    text.match(/(?:Total\s+(?:de\s+)?Energia\s+Injetada[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i);
  if (injetadaMatch && injetadaMatch[1]) {
    energiaInjetadaKwh = injetadaMatch[1].trim();
  } else if (energiaAtvInjetadaGDIOUC > 0 || energiaAtvInjetadaGDIMUC > 0) {
    energiaInjetadaKwh = formatNumberToBR(energiaAtvInjetadaGDIOUC + energiaAtvInjetadaGDIMUC);
  }

  const compensadaMatch =
    text.match(/(?:Energia\s+(?:Ativa\s+)?Compensada\s*(?:\(kWh\))?[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i) ||
    text.match(/(?:Total\s+(?:de\s+)?Energia\s+Compensada[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i) ||
    text.match(/(?:Energia\s+Compensada\s+GDI[^\d:\n\r]{0,35}[:\s]+)(?:KWH|kWh)?\s*([\d\.]+(?:,\d{2})?)/i);
  if (compensadaMatch && compensadaMatch[1]) {
    energiaCompensadaKwh = compensadaMatch[1].trim();
  }

  // Set consumoKwh for Grupo B (B3, Convencional)
    if (isGrupoB) {
      const isCustoDisponibilidadeOnly = /Custo\s+de\s+Disponibilidade/i.test(text);
      if (isCustoDisponibilidadeOnly) {
          consumoKwhGrupoB = "0";
          valorConsumoKwhGrupoB = "0";
      } else {
          let extractedConsumo = extract(/Energia ativa em kWh[\s\S]{0,30}?(?:\d[\d\.,]*)\s+([\d\.,]+)/i) || 
                                 extract(/Consumo\s+(?:em\s+)?kWh[\s\S]{0,80}?(?:KWH|kWh)?\s*([\d\.]+,\d{2})/i) ||
                                 extract(/Consumo em kWh[\s\S]{0,30}?([\d\.,]+)/i) || "";
          
          // If regex failed or extracted a bad value like '1', prefer foraPonta (where Energisa B3 table extraction usually lands)
          if ((!extractedConsumo || extractedConsumo === "0" || extractedConsumo === "1" || extractedConsumo === "1,00") && consumoKwhForaPonta && consumoKwhForaPonta !== "0") {
              extractedConsumo = consumoKwhForaPonta;
          }

          // Special fallback for Energisa corrupted column extractions (e.g., "KWH 1.250,00")
          if (!extractedConsumo || extractedConsumo === "0" || extractedConsumo === "1" || extractedConsumo === "1,00") {
              let kwhMatch = text.match(/KWH\s+(\d{1,5}(?:\.\d{3})*,\d{2})/i);
              if (kwhMatch) extractedConsumo = kwhMatch[1];
          }

          // If still nothing, check Base de Calc (R$) 3.551,00 or Consumo Faturado
          if (!extractedConsumo || extractedConsumo === "0" || extractedConsumo === "1" || extractedConsumo === "1,00") {
              let baseMatch = text.match(/Base\s+de\s+C[áa]lc\.\s*\(R\$\)\s*([\d\.]+,\d{2})/i);
              if (baseMatch) extractedConsumo = baseMatch[1];
          }

          // If still nothing, fallback to ponta
          if ((!extractedConsumo || extractedConsumo === "0" || extractedConsumo === "1" || extractedConsumo === "1,00") && consumoKwhPonta && consumoKwhPonta !== "0" && consumoKwhPonta !== "1" && consumoKwhPonta !== "1,00") {
              extractedConsumo = consumoKwhPonta;
          }
          
          if (extractedConsumo && extractedConsumo !== "0") {
              consumoKwhGrupoB = extractedConsumo;
          }

          // Mathematical fallback for valorConsumoKwhGrupoB
          let extractedValor = valorConsumoKwhPonta || valorConsumoKwhForaPonta || "";
          if ((!extractedValor || extractedValor === "0") && consumoKwhGrupoB && consumoKwhGrupoB !== "0") {
              const numKwh = parseFloat(consumoKwhGrupoB.replace(/\./g, "").replace(",", "."));
              let bestDiff = 99999;
              let bestV = "0";
              const tariffMatches = Array.from(text.matchAll(/\d{1,3},\d{4,6}/g)).map(m => parseFloat(m[0].replace(",", ".")));
              const numbers = Array.from(text.matchAll(/(?<![\d\.,])\d{1,5}(?:[\.,]\d{3})*[\.,]\d{2}(?![\d\.,])/g)).map(m => parseFloat(m[0].replace(/[^\d]/g, "")) / 100);
              for (let v of numbers) {
                  for (let t of tariffMatches) {
                      let expectedV = numKwh * t;
                      let diff = Math.abs(v - expectedV);
                      if (diff < bestDiff && diff < 1.0) {
                          bestDiff = diff;
                          bestV = formatNumberToBR(v);
                      }
                  }
              }
              if (bestV !== "0") valorConsumoKwhGrupoB = bestV;
          } else if (extractedValor && extractedValor !== "0") {
              valorConsumoKwhGrupoB = extractedValor;
          }
      }
      
      // As requested, Group B should never have Ponta/ForaPonta consumption
      consumoKwhPonta = "";
      consumoKwhForaPonta = "";
      valorConsumoKwhPonta = "";
      valorConsumoKwhForaPonta = "";
  }

  if (uc && UC_MIGRATION_MAP[uc]) {
    uc = UC_MIGRATION_MAP[uc];
  }

  const isFinalVerde = (modalidadeTarifaria || "").toUpperCase().includes("VERDE") || isVerdeGlobal;
  if (isFinalVerde) {
    demandaPontaKW = "0";
  }

  return {
    uc,
    concessionaria,
    cidade,
    anoLeitura,
    mesReferencia,
    dataVencimento,
    numeroNotaFiscal,
    modalidadeTarifaria,
    subgrupo,
    valorTotal: cleanNumber(valorTotal),
    consumoKwh: cleanNumber(consumoKwh),
    consumoKwhPonta: cleanNumber(consumoKwhPonta),
    consumoKwhIntermediaria: cleanNumber(consumoKwhIntermediaria),
    consumoKwhForaPonta: cleanNumber(consumoKwhForaPonta),
    consumoKwhGrupoB: cleanNumber(consumoKwhGrupoB),
    valorConsumoKwhPonta: cleanNumber(valorConsumoKwhPonta),
    valorConsumoKwhForaPonta: cleanNumber(valorConsumoKwhForaPonta),
    valorConsumoKwhGrupoB: cleanNumber(valorConsumoKwhGrupoB),
    demandaPontaKW: cleanNumber(demandaPontaKW),
    demandaTodosPeriodos: cleanNumber(demandaTodosPeriodos),
    demandaForaPontaKW: cleanNumber(demandaForaPontaKW),
    demandaPotenciaMedidaPonta: cleanNumber(demandaPotenciaMedidaPonta),
    demandaPotenciaMedidaForaPonta: cleanNumber(demandaPotenciaMedidaForaPonta),
    demandaPotenciaNaoConsumidaPonta: cleanNumber(demandaPotenciaNaoConsumidaPonta),
    demandaPotenciaNaoConsumidaFPonta: cleanNumber(demandaPotenciaNaoConsumidaForaPonta),
    demandaPotenciaAtivaUltrapPonta: cleanNumber(ultrapassagemPontaKW),
    demandaPotenciaAtivaUltrapFPonta: cleanNumber(ultrapassagemForaPontaKW),
    energiaReativaExcedPonta: energiaReativaExcedPonta,
    energiaReativaExcedFPonta: energiaReativaExcedFPonta,
    valorDemandaPotenciaMedidaPonta: valorDemandaPotenciaMedidaPonta,
    valorDemandaPotenciaMedidaForaPonta: valorDemandaPotenciaMedidaForaPonta,
    valorDemandaPotenciaNaoConsumidaPonta: valorDemandaPotenciaNaoConsumidaPonta,
    valorDemandaPotenciaNaoConsumidaFPonta: valorDemandaPotenciaNaoConsumidaForaPonta,
    valorDemandaPotenciaAtivaUltrapPonta: valorUltrapassagemPonta,
    valorDemandaPotenciaAtivaUltrapFPonta: valorUltrapassagemForaPonta,
    valorEnergiaReativaExcedPonta: valorEnergiaReativaExcedPonta,
    valorEnergiaReativaExcedFPonta: valorEnergiaReativaExcedFPonta,
    energiaAtvInjetadaGDIOUC: formatNumberToBR(energiaAtvInjetadaGDIOUC),
    valorEnergiaAtvInjetadaGDIOUC: formatNumberToBR(valorEnergiaAtvInjetadaGDIOUC),
    energiaAtvInjetadaGDIMUC: formatNumberToBR(energiaAtvInjetadaGDIMUC),
    valorEnergiaAtvInjetadaGDIMUC: formatNumberToBR(valorEnergiaAtvInjetadaGDIMUC),
    energiaInjetadaKwh: cleanNumber(energiaInjetadaKwh),
    energiaCompensadaKwh: cleanNumber(energiaCompensadaKwh),
    saldoAcumuladoGD: cleanNumber(saldoAcumuladoGD).replace(/\./g, ""),
    cip: cleanNumber(cip),
    pis: cleanNumber(pis),
    cofins: cleanNumber(cofins),
    icms: cleanNumber(icms),
  };
}
