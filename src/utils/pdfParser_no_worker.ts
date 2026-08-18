import * as pdfjsLib from "pdfjs-dist";


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
      } else if (typeof pdf.destroy === "function") {
        await pdf.destroy();
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
  
  const rawConcessionaria = extract(/(ELEKTRO|ENERGISA|NEOENERGIA|ELEKB|DANFELEKB)/i);
  let concessionaria = rawConcessionaria.toUpperCase();
  if (concessionaria === "NEOENERGIA" || concessionaria === "ELEKB" || concessionaria === "DANFELEKB") concessionaria = "ELEKTRO";

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
  const mesReferenciaMatch = extract(/([A-Za-z]+)\s*\/\s*(\d{4})/) || extract(/Referência:\s*(\d{2}\/\d{4})/);
  let mesReferencia = "";
  let anoLeitura = "";
  if (mesReferenciaMatch) {
    if (mesReferenciaMatch.includes("/")) {
      mesReferencia = mesReferenciaMatch.split("/")[0];
      anoLeitura = mesReferenciaMatch.split("/")[1];
    } else {
      mesReferencia = mesReferenciaMatch;
      anoLeitura = extract(/[A-Za-z]+\s*\/\s*(\d{4})/) || extract(/(202[0-9])/);
    }
  }

  const dataVencimento = extract(/(\d{2}\/\d{2}\/\d{4})\s*VENCIMENTO/i) || extract(/VENCIMENTO.*?(\d{2}\/\d{2}\/\d{4})/i) || extract(/(\d{2}\/\d{2}\/\d{4})\s*R\$/i);
  const numeroNotaFiscal = extract(/NOTA FISCAL Nº:?\s*([\d\.]+)/i);
  
  let modalidadeTarifaria = extract(/Classifica[\s\S]{1,3}o:\s*([^\/]+)/i);
  const subgrupo = extract(/Classifica[\s\S]{1,3}o:[^\/]+\/\s*([^\s]+)/i);
  const classificacaoLinha = extract(/Classifica[\s\S]{1,3}o:\s*([^\n\r]+)/i);
  const isVerdeGlobal = classificacaoLinha ? /VERDE/i.test(classificacaoLinha) : false;
  let isGrupoB = (modalidadeTarifaria && modalidadeTarifaria.toUpperCase().includes("CONVENCIONAL")) || (subgrupo && subgrupo.toUpperCase().includes("B3"));
  if (!isGrupoB && (/CONVENCIONAL/i.test(text) || /\bB3\b/i.test(text))) {
    isGrupoB = true;
  }

  // Consumo (Tratando B3 que não tem ponta/fora ponta explicito como verde)
  let consumoKwh = "";
  let consumoKwhPonta = extract(/Consumo[\s\S]{0,60}?(?<!Fora\s)(?<!F)Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let consumoKwhForaPonta = extract(/Consumo[\s\S]{0,60}?Fora Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let valorConsumoKwhPonta = extract(/Valor Consumo[\s\S]{0,60}?(?<!Fora\s)(?<!F)Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);
  let valorConsumoKwhForaPonta = extract(/Valor Consumo[\s\S]{0,60}?Fora Ponta[\s\S]{0,60}?(?<![\d\,])([\d\.,]+)/i);

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

  let demandaContratadaPonta = demandaPontaKW || extract(/(\d+(?:,\d+)?)\s*Demanda\s+ponta/i) || extract(/Demanda\s+ponta.*?[\D](\d+(?:,\d+)?)/i) || "0";
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

  let demandaPotenciaMedidaPonta = directKwMedPonta || getExplicitMedida('ponta') || extract(/Demanda Medida.*Ponta[^\d]*([\d\.,]+)/i) || "0";
  let demandaPotenciaMedidaForaPonta = directKwMedForaPonta || getExplicitMedida('fora ponta') || extract(/Demanda Medida.*Fora Ponta[^\d]*([\d\.,]+)/i) || "0";
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
    // Faturas Grupo A geralmente listam tarifas com 6 casas decimais      // Pega todas as tarifas poss.veis (números como 0,494240 ou 0,49424)
    const allTariffMatches = Array.from(text.matchAll(/\d{1,3}[\.,]\d{4,6}/g)).map(m => parseFloat(m[0].replace(",", ".")));
    const tariffMatches = allTariffMatches.filter(t => t > 0.05 && t < 300);
    const numbers = Array.from(text.matchAll(/(?<![\d\.,])\d{1,5}(?:[\.,]\d{3})*[\.,]\d{2}(?![\d\.,])/g))
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
                  if (relDiff < 0.005 || absDiff < 2.0) {
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

      let qPontaRegexStr = extract(/(?<!Fora\s)(?<!F)Consumo em kWh - Ponta[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) ||
                           extract(/KWH\s+Ponta\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)/i) ||
                           extract(/Energia\s+ativa\s+em\s+kWh\s+(?:-\s*)?Ponta[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) ||
                           extract(/(?<!Fora\s)(?<!F)Ponta Energia ativa em kWh[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i);

      let qForaPontaRegexStr = extract(/(?<!Fora\s)(?<!F)Consumo em kWh - Fora Ponta[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) ||
                               extract(/Consumo em kWh - Fponta[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) ||
                               extract(/KWH\s+FPonta\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+[\d\.,]+\s+([\d\.,]+)/i) ||
                               extract(/Energia\s+ativa\s+em\s+kWh\s+(?:-\s*)?Fora\s+Ponta[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) ||
                               extract(/Fora Ponta Energia ativa em kWh[\s\S]{0,60}?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i);

      if (qPontaRegexStr) {
          let val = parseFloat(qPontaRegexStr.replace(/\./g, "").replace(",", "."));
          ponta = pairs.find(p => p.t < 10.0 && p.t > 0.05 && p.q > 0.5 && (Math.abs(p.q - val) < 2.0));
      }
      if (qForaPontaRegexStr) {
          let val = parseFloat(qForaPontaRegexStr.replace(/\./g, "").replace(",", "."));
          foraPonta = pairs.find(p => p.t < 10.0 && p.t > 0.05 && p.q > 0.5 && (Math.abs(p.q - val) < 2.0));
      }

        if ((!ponta || !foraPonta) && pairs.length > 0) {
          // Consumo DEVE ter tarifa < 10.0 (energia ativa kWh) e nunca tarifa de demanda (>= 10.0)
          let validPairs = pairs.filter(p => p.t < 10.0 && p.t > 0.05 && p.q > 5.0);
          validPairs.sort((a, b) => {
             let aIdx = text.indexOf(formatNumberToBR(a.v));
             let bIdx = text.indexOf(formatNumberToBR(b.v));
             if (aIdx === -1) aIdx = 999999;
             if (bIdx === -1) bIdx = 999999;
             return aIdx - bIdx;
          });
          
          let availablePairs = validPairs.filter(p => {
              if (p === ponta || p === foraPonta) return false;
              if (ponta && (Math.abs(p.v - ponta.v) < 1.0 || Math.abs(p.q - ponta.q) < 1.0 || Math.abs(p.v - ponta.q) < 1.0 || Math.abs(p.q - ponta.v) < 1.0)) return false;
              if (foraPonta && (Math.abs(p.v - foraPonta.v) < 1.0 || Math.abs(p.q - foraPonta.q) < 1.0 || Math.abs(p.v - foraPonta.q) < 1.0 || Math.abs(p.q - foraPonta.v) < 1.0)) return false;
              return true;
          });

                    if (!ponta && !foraPonta) {              let bestPairs = availablePairs.slice(0, 2);
              if (bestPairs.length >= 2) {
                  bestPairs.sort((a, b) => b.t - a.t);
                  if (bestPairs[0].t > 0.8 || bestPairs[0].t > bestPairs[1].t * 1.5) {
                      ponta = bestPairs[0];
                      foraPonta = bestPairs[1];
                  } else {
                      // Both are Fora Ponta tiers
                      foraPonta = bestPairs[0];
                      foraPonta.q += bestPairs[1].q;
                      foraPonta.v += bestPairs[1].v;
                  }
              } else if (bestPairs.length === 1) {
                  foraPonta = bestPairs[0];
              }
          } else if (!ponta && availablePairs.length > 0) {
                // If foraPonta already exists, only assign ponta if tariff is noticeably higher
                if (foraPonta && availablePairs[0].t > foraPonta.t * 1.5) {
                    ponta = availablePairs[0];
                } else if (foraPonta && Math.abs(availablePairs[0].t - foraPonta.t) < 0.05) {
                    // It's just another foraPonta tier, sum it
                    foraPonta.q += availablePairs[0].q;
                    foraPonta.v += availablePairs[0].v;
                } else {
                    ponta = availablePairs[0];
                }
            } else if (!foraPonta && availablePairs.length > 0) {
                foraPonta = availablePairs[0];
            }
      }

      if (ponta) {
         consumoKwhPonta = formatNumberToBR(ponta.q);
         valorConsumoKwhPonta = formatNumberToBR(ponta.v);
      } else if (qPontaRegexStr) {
         consumoKwhPonta = formatNumberToBR(parseFloat(qPontaRegexStr.replace(/\./g, "").replace(",", ".")));
      }
      if (foraPonta) {
         consumoKwhForaPonta = formatNumberToBR(foraPonta.q);
         valorConsumoKwhForaPonta = formatNumberToBR(foraPonta.v);
      } else if (qForaPontaRegexStr) {
         consumoKwhForaPonta = formatNumberToBR(parseFloat(qForaPontaRegexStr.replace(/\./g, "").replace(",", ".")));
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

    const tariffMatches = Array.from(text.matchAll(/\d{1,3}[\.,]\d{4,6}/g)).map(m => parseFloat(m[0].replace(",", ".")));
    const numbers = Array.from(text.matchAll(/(?<![\d\.,])\d{1,5}(?:[\.,]\d{3})*[\.,]\d{2}(?![\d\.,])/g))
         .map(m => parseFloat(m[0].replace(/[^\d]/g, "")) / 100);

    const demandaPairs: { q: number; v: number; t: number; error: number }[] = [];
    for (const t of tariffMatches) {
      if (t >= 10.0 && t < 300.0) {
        for (const q of numbers) {
          if (q > 0.1) {
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

    demandaPairs.sort((a, b) => a.error - b.error);

    let contPonta = parseFloat((demandaPontaKW || "0").replace(/\./g, "").replace(",", ".")) || 0;
    let contForaPonta = parseFloat((demandaForaPontaKW || "0").replace(/\./g, "").replace(",", ".")) || 0;

    const hasNcPonta = /Não\s+Consumida\s*-\s*Ponta|N\.o\s+Consumida\s*-\s*Ponta/i.test(text);
    const hasNcForaPonta = /Não\s+Consumida\s*-\s*F|N\.o\s+Consumida\s*-\s*F/i.test(text);
    const hasUltPonta = /Ultrapassagem\s*-\s*Ponta|ULTP\s*Ponta/i.test(text);
    const hasUltForaPonta = /Ultrapassagem\s*-\s*F|ULTP\s*FPonta/i.test(text);

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
        if (hasNcForaPonta) {
          let bestPair: { med: any; nc: any } | null = null;
          for (let i = 0; i < demandaPairs.length; i++) {
            for (let j = i + 1; j < demandaPairs.length; j++) {
              const p1 = demandaPairs[i], p2 = demandaPairs[j];
              if (contForaPonta > 0 && Math.abs(p1.q + p2.q - contForaPonta) < 1.0) {
                bestPair = p1.q > p2.q ? { med: p1, nc: p2 } : { med: p2, nc: p1 };
                break;
              }
            }
            if (bestPair) break;
          }

          if (bestPair) {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(bestPair.med.q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(bestPair.med.v);
            demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestPair.nc.q);
            valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestPair.nc.v);
          } else {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
          }
        } else if (hasUltForaPonta) {
          let bestUlt: { med: any; ult: any } | null = null;
          for (let i = 0; i < demandaPairs.length; i++) {
            for (let j = i + 1; j < demandaPairs.length; j++) {
              const p1 = demandaPairs[i], p2 = demandaPairs[j];
              const larger = p1.q > p2.q ? p1 : p2;
              const smaller = p1.q > p2.q ? p2 : p1;
              if (contForaPonta > 0 && Math.abs(larger.q - smaller.q - contForaPonta) < 1.0) {
                bestUlt = { med: larger, ult: smaller };
                break;
              }
            }
            if (bestUlt) break;
          }
          if (bestUlt) {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(bestUlt.med.q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(bestUlt.med.v);
            ultrapassagemForaPontaKW = formatNumberToBR(bestUlt.ult.q);
            valorUltrapassagemForaPonta = formatNumberToBR(bestUlt.ult.v);
          } else {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
          }
        } else {
          demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
          valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
        }
      }
    } else {
      // Modalidade AZUL
      const pontaPairs = demandaPairs.filter(p => p.t >= 60.0);
      const foraPairs = demandaPairs.filter(p => p.t < 60.0);

      // PONTA
      if (pontaPairs.length > 0) {
        if (hasNcPonta) {
          let bestNcPonta: { med: any; nc: any } | null = null;
          for (let i = 0; i < pontaPairs.length; i++) {
            for (let j = i + 1; j < pontaPairs.length; j++) {
              const p1 = pontaPairs[i], p2 = pontaPairs[j];
              if (contPonta > 0 && Math.abs(p1.q + p2.q - contPonta) < 1.0) {
                bestNcPonta = p1.q > p2.q ? { med: p1, nc: p2 } : { med: p2, nc: p1 };
                break;
              }
            }
            if (bestNcPonta) break;
          }
          if (bestNcPonta) {
            demandaPotenciaMedidaPonta = formatNumberToBR(bestNcPonta.med.q);
            valorDemandaPotenciaMedidaPonta = formatNumberToBR(bestNcPonta.med.v);
            demandaPotenciaNaoConsumidaPonta = formatNumberToBR(bestNcPonta.nc.q);
            valorDemandaPotenciaNaoConsumidaPonta = formatNumberToBR(bestNcPonta.nc.v);
          } else {
            demandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].q);
            valorDemandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].v);
          }
        } else if (hasUltPonta) {
          let bestUltPonta: { med: any; ult: any } | null = null;
          for (let i = 0; i < pontaPairs.length; i++) {
            for (let j = i + 1; j < pontaPairs.length; j++) {
              const p1 = pontaPairs[i], p2 = pontaPairs[j];
              const larger = p1.q > p2.q ? p1 : p2;
              const smaller = p1.q > p2.q ? p2 : p1;
              if (contPonta > 0 && Math.abs(larger.q - smaller.q - contPonta) < 1.0) {
                bestUltPonta = { med: larger, ult: smaller };
                break;
              }
            }
            if (bestUltPonta) break;
          }
          if (bestUltPonta) {
            demandaPotenciaMedidaPonta = formatNumberToBR(bestUltPonta.med.q);
            valorDemandaPotenciaMedidaPonta = formatNumberToBR(bestUltPonta.med.v);
            ultrapassagemPontaKW = formatNumberToBR(bestUltPonta.ult.q);
            valorUltrapassagemPonta = formatNumberToBR(bestUltPonta.ult.v);
          } else {
            demandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].q);
            valorDemandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].v);
          }
        } else {
          demandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].q);
          valorDemandaPotenciaMedidaPonta = formatNumberToBR(pontaPairs[0].v);
        }
      }

      // FORA PONTA
      if (foraPairs.length > 0) {
        if (hasNcForaPonta) {
          let bestNcFora: { med: any; nc: any } | null = null;
          for (let i = 0; i < foraPairs.length; i++) {
            for (let j = i + 1; j < foraPairs.length; j++) {
              const p1 = foraPairs[i], p2 = foraPairs[j];
              if (contForaPonta > 0 && Math.abs(p1.q + p2.q - contForaPonta) < 1.0) {
                bestNcFora = p1.q > p2.q ? { med: p1, nc: p2 } : { med: p2, nc: p1 };
                break;
              }
            }
            if (bestNcFora) break;
          }
          if (bestNcFora) {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(bestNcFora.med.q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(bestNcFora.med.v);
            demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNcFora.nc.q);
            valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNcFora.nc.v);
          } else {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].v);
          }
        } else if (hasUltForaPonta) {
          let bestUltFora: { med: any; ult: any } | null = null;
          for (let i = 0; i < foraPairs.length; i++) {
            for (let j = i + 1; j < foraPairs.length; j++) {
              const p1 = foraPairs[i], p2 = foraPairs[j];
              const larger = p1.q > p2.q ? p1 : p2;
              const smaller = p1.q > p2.q ? p2 : p1;
              if (contForaPonta > 0 && Math.abs(larger.q - smaller.q - contForaPonta) < 1.0) {
                bestUltFora = { med: larger, ult: smaller };
                break;
              }
            }
            if (bestUltFora) break;
          }
          if (bestUltFora) {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(bestUltFora.med.q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(bestUltFora.med.v);
            ultrapassagemForaPontaKW = formatNumberToBR(bestUltFora.ult.q);
            valorUltrapassagemForaPonta = formatNumberToBR(bestUltFora.ult.v);
          } else {
            demandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].v);
          }
        } else {
          demandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].q);
          valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(foraPairs[0].v);
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
    // Na Elektro (B3), lemos como "Consumo Fora Ponta" para compatibilidade com Sanesul.
    let consumoMatch = text.match(/CONSUMO(?: TE| TUSD)?\s+kWh\s+(\d+)/i) || text.match(/kWh\s+(?:kWh\s+)?(?:kWh\s+)?(\d+)/);
      if ((!consumoKwhForaPonta || consumoKwhForaPonta === "0") && qForaPontaRegexStr) { consumoKwhForaPonta = formatNumberToBR(parseFloat(qForaPontaRegexStr.replace(/\./g, "").replace(",", "."))); }
        if (!consumoKwhForaPonta || consumoKwhForaPonta === "0") {
          if (consumoMatch) {
            consumoKwhForaPonta = consumoMatch[1];
          } else {
            let fallback = text.match(/kWh[^\d]+(\d+)/);
            if (fallback) consumoKwhForaPonta = fallback[1];
          }
      }
    
    let cipMatch = extract(/COBRANCA ILUM PUBLICA[^\d]+([\d]+\,[\d]{2})/i) || extract(/ILUM PUBLICA[^\d]+([\d]+\,[\d]{2})/i);
    if (cipMatch) cip = cipMatch;

    let cidadeElektroMatch = text.match(/([A-ZÀ-Ÿ\s]+)\s*\([A-Z]{2}\)\s*-\s*[A-Z]{2}\s*(?:-|CNPJ|IE|INSC|CEP)/i);
    if (cidadeElektroMatch) {
      cidade = cidadeElektroMatch[1].trim();
      const parts = cidade.split(/\s{2,}/);
      cidade = parts[parts.length - 1];
    }

    let modElektroMatch = text.match(/(CONVENCIONAL\s*\/\s*BIFASICO|CONVENCIONAL\s*\/\s*TRIFASICO|CONVENCIONAL\s*\/\s*MONOFASICO)/i);
    if (modElektroMatch) {
      modalidadeTarifaria = modElektroMatch[1];
    }
  }

  // Extração de GDI (Geração Distribuída Injetada) para Energisa
  let energiaAtvInjetadaGDIOUC = 0;
  let valorEnergiaAtvInjetadaGDIOUC = 0;
  let energiaAtvInjetadaGDIMUC = 0;
  let valorEnergiaAtvInjetadaGDIMUC = 0;

  if (concessionaria === "ENERGISA" && text.includes("GDI")) {
    const oUcCount = (text.match(/GDI oUC/g) || []).length;
    const mUcCount = (text.match(/GDI mUC/g) || []).length;
    const expectedTotal = oUcCount + mUcCount;
    let foundCount = 0;

    // Estratégia 1: "GDI mUC ... KWH <kWh> <tarifa> -<R$>" (Grupo A multi-mês com KWH inline)
    const gdiLineRegex = /(GDI mUC|GDI oUC)[\s\S]{0,60}?(?:KWH|kWh)\s*([\d\.]+,\d{2})\s+[\d\.]+,\d+\s+([-\u2010-\u2015\u2212][\d\.]+,\d{2})/ig;
    let gdiMatch;
    while ((gdiMatch = gdiLineRegex.exec(text)) !== null) {
      foundCount++;
      const type = gdiMatch[1].toUpperCase();
      const kwhVal = parseFloat(gdiMatch[2].replace(/\./g, "").replace(",", "."));
      const rawR = gdiMatch[3].replace(/[^\d\.,-]/g, "");
      const rVal = -Math.abs(parseFloat(rawR.replace(/\./g, "").replace(",", ".")));
      if (type.includes("MUC")) {
        energiaAtvInjetadaGDIMUC += kwhVal;
        valorEnergiaAtvInjetadaGDIMUC += rVal;
      } else {
        energiaAtvInjetadaGDIOUC += kwhVal;
        valorEnergiaAtvInjetadaGDIOUC += rVal;
      }
    }

    // Estratégia 2: varredura por janela (Grupo B - labels primeiro, valores depois)
    // Também usado quando Estratégia 1 capturou menos itens do que o esperado
    if (foundCount < expectedTotal) {
      if (foundCount === 0) {
        energiaAtvInjetadaGDIMUC = 0;
        valorEnergiaAtvInjetadaGDIMUC = 0;
        energiaAtvInjetadaGDIOUC = 0;
        valorEnergiaAtvInjetadaGDIOUC = 0;
      }

      // Isola janela: do primeiro ao último label GDI + 800 chars
      const firstGdiIdx = text.search(/GDI (?:mUC|oUC)/i);
      const lastGdiMatch = [...text.matchAll(/GDI (?:mUC|oUC)/ig)].pop();
      const lastGdiIdx = lastGdiMatch ? (lastGdiMatch.index ?? firstGdiIdx) : firstGdiIdx;
      const gdiWindow = text.substring(Math.max(0, firstGdiIdx - 50), Math.min(text.length, lastGdiIdx + 800));

      // Coleta todos os pares <tarifa_6dec> + <negativo> dentro da janela
      const pairRegex = /(\d{1,3},\d{4,6})\s+([-\u2010-\u2015\u2212][\d\.]+,\d{2})/g;
      const pairs: {tariff: number, rVal: number}[] = [];
      let pairM;
      while ((pairM = pairRegex.exec(gdiWindow)) !== null) {
        const tariff = parseFloat(pairM[1].replace(",", "."));
        const rawR = pairM[2].replace(/[^\d.,]/g, "");
        const rVal = -Math.abs(parseFloat(rawR.replace(/\./g, "").replace(",", ".")));
        pairs.push({ tariff, rVal });
      }

      if (pairs.length >= expectedTotal) {
        // Janela continha pares suficientes — distribui normalmente
        let pIdx = 0;
        for (let i = 0; i < oUcCount && pIdx < pairs.length; i++, pIdx++) {
          const { tariff, rVal } = pairs[pIdx];
          valorEnergiaAtvInjetadaGDIOUC += rVal;
          energiaAtvInjetadaGDIOUC += tariff > 0 ? Math.round((Math.abs(rVal) / tariff) * 100) / 100 : 0;
        }
        for (let i = 0; i < mUcCount && pIdx < pairs.length; i++, pIdx++) {
          const { tariff, rVal } = pairs[pIdx];
          valorEnergiaAtvInjetadaGDIMUC += rVal;
          energiaAtvInjetadaGDIMUC += tariff > 0 ? Math.round((Math.abs(rVal) / tariff) * 100) / 100 : 0;
        }
      } else {
        // Estratégia 3: valores GDI aparecem ANTES dos labels no texto (padrão intercalado)
        // Ex: "-151,75 -13,63 -158,25 -14,21 ..." = pares (principal, PIS+COFINS) intercalados
        //
        // Passo 1: Detecta a tarifa GDI avaliando todas as tarifas com 6 casas decimais
        const tariffs = Array.from(new Set(Array.from(text.matchAll(/\d{1,3},\d{4,6}/g)).map(m => Math.round(parseFloat(m[0].replace(",", ".")) * 1000000) / 1000000)));
        const possibleTariffs = tariffs.filter(t => t > 0.5 && t < 5);

        // Passo 2: Coleta todos os valores negativos no texto inteiro
        const allNegRaw = Array.from(text.matchAll(/([-\u2010-\u2015\u2212][\d\.]+,\d{2})/g))
          .map(m => Math.abs(parseFloat(m[0].replace(/[^\d.,]/g, "").replace(".", "").replace(",", "."))));

        // Passo 3: Encontra a tarifa que divide os maiores valores negativos num kWh inteiro
        let gdiTariff = 0;
        let maxMatches = -1;
        for (const t of possibleTariffs) {
          let matches = 0;
          for (const v of allNegRaw) {
            if (v < 1) continue;
            const kwhEst = v / t;
            if (Math.abs(Math.round(kwhEst) - kwhEst) <= 0.05) matches++;
          }
          if (matches > maxMatches) {
            maxMatches = matches;
            gdiTariff = t;
          }
        }

        // Passo 3: Detecta se valores estão intercalados (main, cofins, main, cofins...)
        // Sinal: segundo valor é ~9% do primeiro (PIS+COFINS ≈ 9%)
        const isInterleaved = allNegRaw.length >= 4 &&
          allNegRaw[0] > 1 &&
          (allNegRaw[1] / allNegRaw[0]) < 0.20 &&
          (allNegRaw[3] / allNegRaw[2]) < 0.20;

        // Passo 4: Filtra apenas os valores "principais" (ignora os de imposto)
        const mainNeg: number[] = [];
        for (let i = 0; i < allNegRaw.length && mainNeg.length < expectedTotal; i++) {
          if (isInterleaved && i % 2 !== 0) continue; // pula valores de imposto
          const v = allNegRaw[i];
          if (v < 0.5) continue; // ignora valores mínimos (ruído)
          // Valida: se temos a tarifa, o kWh deve ser próximo de um inteiro
          if (gdiTariff > 0) {
            const kwhEst = v / gdiTariff;
            if (Math.abs(Math.round(kwhEst) - kwhEst) > 0.05) continue;
          }
          mainNeg.push(-v);
        }

        // Passo 5: Distribui
        let pIdx3 = 0;
        for (let i = 0; i < oUcCount && pIdx3 < mainNeg.length; i++, pIdx3++) {
          const rVal = mainNeg[pIdx3];
          valorEnergiaAtvInjetadaGDIOUC += rVal;
          energiaAtvInjetadaGDIOUC += gdiTariff > 0 ? Math.round((Math.abs(rVal) / gdiTariff) * 100) / 100 : 0;
        }
        for (let i = 0; i < mUcCount && pIdx3 < mainNeg.length; i++, pIdx3++) {
          const rVal = mainNeg[pIdx3];
          valorEnergiaAtvInjetadaGDIMUC += rVal;
          energiaAtvInjetadaGDIMUC += gdiTariff > 0 ? Math.round((Math.abs(rVal) / gdiTariff) * 100) / 100 : 0;
        }
      }
    }
  }
  // Set consumoKwh for Grupo B (B3, Convencional)
    if (isGrupoB) {
      if (!consumoKwhPonta || consumoKwhPonta === "0") {
        let extractedConsumo = extract(/Energia ativa em kWh[\s\S]{0,10}?(?:\d[\d\.,]*)\s+([\d\.,]+)/i) || extract(/Consumo em kWh[\s\S]{0,10}?([\d\.,]+)/i);
        if (extractedConsumo && extractedConsumo !== "0") {
          consumoKwhPonta = extractedConsumo;
        }
      }
      
      // Mathematical fallback for valorConsumoKwhPonta
      if ((!valorConsumoKwhPonta || valorConsumoKwhPonta === "0") && consumoKwhPonta && consumoKwhPonta !== "0") {
          const numKwh = parseFloat(consumoKwhPonta.replace(/\./g, "").replace(",", "."));
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
          if (bestV !== "0") valorConsumoKwhPonta = bestV;
      }
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
    consumoKwhForaPonta: cleanNumber(consumoKwhForaPonta),
    valorConsumoKwhPonta: cleanNumber(valorConsumoKwhPonta),
    valorConsumoKwhForaPonta: cleanNumber(valorConsumoKwhForaPonta),
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
    cip: cleanNumber(cip),
    pis: cleanNumber(pis),
    cofins: cleanNumber(cofins),
    icms: cleanNumber(icms),
  };
}
