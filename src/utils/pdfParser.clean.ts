import * as pdfjsLib from "pdfjs-dist";


// Set the workerSrc to the local file


export async function extractTextFromPdf(file: File | Blob): Promise<string> {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = "";

  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const pageText = textContent.items
      .map((item: any) => item.str)
      .join(" ");
    fullText += pageText + "\n";
  }

  console.log("=== TEXTO EXTRAÍDO DO PDF ===");
  console.log(fullText);
  console.log("=============================");

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
  let uc = extract(/(?:UC|Instalação|CÓDIGO DO CLIENTE|Seu C[oó]digo|Unidade Consumidora)[\s:]*([\d\.\-\/]+)/i);
  if (uc) uc = uc.replace(/[^\d]/g, "");
  
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
    // Supports 2-dot (181.951.005-101) and 3-dot (2.822.635.051-30) formats
    // Each digit group limited to 1-4 digits to avoid capturing dates
    let formattedUc = "";
    const energisaUcMatch = text.match(/(\d{1,4}(?:\.\s*\d{1,3}){2,3}-\s*\d{1,3})/g);
    if (energisaUcMatch) {
      for (const match of energisaUcMatch) {
        const parsedUc = match.replace(/[^\d]/g, "");
        // Ignore Inscrição Estadual MS which starts with 28 and has 9 digits
        if (parsedUc.startsWith("28") && parsedUc.length === 9) continue;
        // Prefer longest formatted match
        if (!formattedUc || parsedUc.length > formattedUc.length) {
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
    
    // Strategy 3: Direct search for long digit sequences (11-12 digits) ONLY if no UC found at all
    if (!uc || uc.length < 6) {
      // Use word boundaries to avoid matching parts of longer strings like 16-digit protocol numbers
      const longNumberMatch = text.match(/\b(\d{11,12})\b/g);
      if (longNumberMatch) {
        for (const num of longNumberMatch) {
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
  const valorTotal = extract(/(?:Total a Pagar|TOTAL:|Total R\$[^\d]*R\$)[^\d]*([\d\.,]+)/i) || extract(/R\$\s*([\d\.,]+)\s*Nº FATURA/i);
  
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
  
  let modalidadeTarifaria = extract(/Classificação:\s*([^\/]+)/i);
  const subgrupo = extract(/Classificação:[^\/]+\/\s*([^\s]+)/i);

  // Consumo (Tratando B3 que não tem ponta/fora ponta explicito como verde)
  let consumoKwh = "";
  let consumoKwhPonta = extract(/Consumo.*(?<!Fora )(?<!F)Ponta[^\d]*([\d\.,]+)/i);
  let consumoKwhForaPonta = extract(/Consumo.*Fora Ponta[^\d]*([\d\.,]+)/i);
  let valorConsumoKwhPonta = extract(/Valor Consumo.*(?<!Fora )(?<!F)Ponta[^\d]*([\d\.,]+)/i);
  let valorConsumoKwhForaPonta = extract(/Valor Consumo.*Fora Ponta[^\d]*([\d\.,]+)/i);

  // Demandas e Reativos
  let demandaPontaKW = extract(/Demanda Contratada[^\d]*(?<!Fora\s)(?<!F)Ponta[^\d]*([\d\.,]+)/i) || extract(/KW Ponta:\s*([\d\.,]+)/i) || extract(/Ponta:\s*([\d\.,]+)\s*Fora\s*Ponta/i) || extract(/([\d\.,]+)\s+Demanda ponta/i);
  let demandaTodosPeriodos = extract(/Demanda Todos os Períodos:\s*([\d\.,]+)/i) || "0";
  let demandaForaPontaKW = extract(/Demanda Contratada[^\d]*(?:Fora Ponta|Fponta)[^\d]*([\d\.,]+)/i) || extract(/Fora Ponta:\s*([\d\.,]+)\s*TUSDG/i) || extract(/KW Fora Ponta:\s*([\d\.,]+)/i) || extract(/([\d\.,]+)\s+Demanda fora ponta/i);
  let demandaPotenciaMedidaPonta = extract(/Demanda Medida.*Ponta[^\d]*([\d\.,]+)/i) || "0";
  let demandaPotenciaMedidaForaPonta = extract(/Demanda Medida.*Fora Ponta[^\d]*([\d\.,]+)/i) || "0";
  // Demanda Potência Não Consumida – limit to 30 chars gap to avoid matching far-away numbers
  let demandaPotenciaNaoConsumidaPonta = extract(/Demanda\s+Potência\s+Não\s+Consumida\s*-\s*Ponta[^\d]{0,30}([\d\.,]+)/i) || "0";
  let demandaPotenciaNaoConsumidaForaPonta = extract(/Demanda\s+Potência\s+Não\s+Consumida\s*-\s*F(?:ora)?\s*Ponta[^\d]{0,30}([\d\.,]+)/i) || "0";
  let ultrapassagemPontaKW = "0";
  let ultrapassagemForaPontaKW = "0";
  let energiaReativaExcedPonta = extract(/Reativa Exced[^-]*-\s*(?!F)Ponta[^\d]{0,30}([\d\.,]+)/i) || "0";
  let energiaReativaExcedFPonta = extract(/Reativa Exced[^-]*-\s*F(?:ora)?\s*Ponta[^\d]{0,30}([\d\.,]+)/i) || "0";
  let valorDemandaPotenciaMedidaPonta = "0";
  let valorDemandaPotenciaMedidaForaPonta = "0";
  let valorDemandaPotenciaNaoConsumidaPonta = "0";
  let valorDemandaPotenciaNaoConsumidaForaPonta = "0";
  let valorUltrapassagemPonta = "0";
  let valorUltrapassagemForaPonta = "0";
  let valorEnergiaReativaExcedPonta = "0";
  let valorEnergiaReativaExcedFPonta = "0";

  if (concessionaria === "ENERGISA") {
    // Heurística matemática para Consumo Ponta e Fora Ponta em faturas Grupo A
    // Faturas Grupo A geralmente listam tarifas com 6 casas decimais e os valores espalhados.
    // Excluir tarifas do TUSDG (são concatenações de números, ex: "0,100222" = "0,10" + "0222")
    const tusdgSection = text.match(/TUSDG.*?(?=Importante|SEGUNDA|Insc\.|$)/is)?.[0] || "";
    const allTariffMatches = Array.from(text.matchAll(/\d{1,3},\d{6}/g)).map(m => ({ val: parseFloat(m[0].replace(",", ".")), pos: m.index || 0 }));
    const tusdgEnd = tusdgSection.length > 0 ? text.indexOf(tusdgSection) + tusdgSection.length : 0;
    const tariffMatches = allTariffMatches.filter(t => t.pos >= tusdgEnd).map(t => t.val);
    const numbers = Array.from(text.matchAll(/(?<![\d\,])\d{1,5}(?:\.\d{3})*,\d{2}(?![\d\,])/g))
         .map(m => parseFloat(m[0].replace(/\./g, "").replace(",", ".")));

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
         if (t > 0.1 && t < 300) { // Tarifas variam de 0.2 a 150+ para Demandas
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

      let qPontaRegexStr = extract(/(?<!Fora\s)(?<!F)Ponta Energia ativa em kWh.*?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) || extract(/(?<!Fora\s)(?<!F)Consumo em kWh - Ponta.*?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i);
      let qForaPontaRegexStr = extract(/(?<!Ponta.*)Fora Ponta Energia ativa em kWh.*?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i) || extract(/Consumo em kWh - Fora Ponta.*?(?<![\d\,])(\d{1,5}(?:\.\d{3})*,\d+)/i);

      if (qPontaRegexStr) {
          let val = parseFloat(qPontaRegexStr.replace(/\./g, "").replace(",", "."));
          ponta = pairs.find(p => p.t < 10.0 && (p.t < 0.9 || p.t > 1.1) && p.q > 50.0 && (Math.abs(p.q - val) < 2.0));
      }
      if (qForaPontaRegexStr) {
          let val = parseFloat(qForaPontaRegexStr.replace(/\./g, "").replace(",", "."));
          foraPonta = pairs.find(p => p.t < 10.0 && (p.t < 0.9 || p.t > 1.1) && p.q > 50.0 && (Math.abs(p.q - val) < 2.0));
      }

      // Se ambos os regex do TUSDG encontraram valores mas nenhum correspondeu a um par,
      // é porque os valores são muito pequenos ou concatenados. Não usar fallback neste caso.
      let bothRegexFoundButNeitherMatched = qPontaRegexStr && qForaPontaRegexStr && !ponta && !foraPonta;
        if ((!ponta || !foraPonta) && pairs.length > 0 && !bothRegexFoundButNeitherMatched) {
          let validPairs = pairs.filter(p => p.t < 10.0 && (p.t < 0.9 || p.t > 1.1) && p.q > 50.0);
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
  // Não aplicar para Energisa Grupo A (que usa pair matching) para evitar capturar valores errados
  if (!consumoKwhPonta && !consumoKwhForaPonta && concessionaria !== "ENERGISA") {
    const consumoGeral = extract(/Energia ativa em kWh\s+\d+\s+([\d\.,]+)/i) || extract(/Consumo kWh\s+[A-Z]{3}\/\d{2}\s+([\d\.,]+)/i);
    if (consumoGeral) {
      consumoKwhForaPonta = consumoGeral;
    }
  }

  if (concessionaria === "ENERGISA") {
    // Demanda Contratada


    if (modalidadeTarifaria && modalidadeTarifaria.toUpperCase().includes('VERDE')) {
        demandaPontaKW = "0";
        demandaPotenciaMedidaPonta = "0";
        valorDemandaPotenciaMedidaPonta = "0";
        demandaPotenciaNaoConsumidaPonta = "0";
        valorDemandaPotenciaNaoConsumidaPonta = "0";
        ultrapassagemPontaKW = "0";
        valorUltrapassagemPonta = "0";
    }
    if (!demandaPontaKW) demandaPontaKW = "0";
    if (!demandaForaPontaKW) demandaForaPontaKW = "0";

    // Heurística matemática robusta para Demandas
    const tariffMatches = Array.from(text.matchAll(/\d{1,3},\d{6}/g)).map(m => parseFloat(m[0].replace(",", ".")));
    const numbers = Array.from(text.matchAll(/(?<![\d\,])\d{1,5}(?:\.\d{3})*,\d{2}(?![\d\,])/g))
         .map(m => parseFloat(m[0].replace(/\./g, "").replace(",", ".")));

    const demandaPairs: {q: number, v: number, t: number, error: number}[] = [];
    for (const t of tariffMatches) {
       if (t >= 10.0 && t < 300.0) { 
          for (const q of numbers) {
             if (q > 0.5) { 
                const expectedV = q * t;
                for (const v of numbers) {
                   if (v > expectedV * 0.98 && v < expectedV * 1.02) {
                      const absDiff = Math.abs(v - expectedV);
                      const relDiff = absDiff / expectedV;
                      if (absDiff < 0.5) {
                         if (!demandaPairs.find(p => p.q === q && p.v === v)) {
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

    let contPonta = parseFloat(demandaPontaKW.replace(/\./g, "").replace(",", ".")) || 0;
    let contForaPonta = parseFloat(demandaForaPontaKW.replace(/\./g, "").replace(",", ".")) || 0;
    let usedPairs = new Set();

    const getCandidates = (contratada: number) => {
        let candidates = [];
        for (let p of demandaPairs) {
            if (Math.abs(p.q - contratada) < 1.0) candidates.push({ type: 'single', p1: p, p2: null });
        }
        for (let i = 0; i < demandaPairs.length; i++) {
            for (let j = i + 1; j < demandaPairs.length; j++) {
                let p1 = demandaPairs[i], p2 = demandaPairs[j];
                if (Math.abs(p1.q + p2.q - contratada) < 1.0) {
                    // nc: p1 = medida (larger q, closer to contracted), p2 = não consumida (smaller q, the residual)
                    candidates.push({ type: 'nc', p1: p1.q > p2.q ? p1 : p2, p2: p1.q > p2.q ? p2 : p1 });
                }
            }
        }
        for (let i = 0; i < demandaPairs.length; i++) {
            for (let j = 0; j < demandaPairs.length; j++) {
                if (i !== j) {
                    let p1 = demandaPairs[i], p2 = demandaPairs[j];
                    if (Math.abs(p1.q - contratada - p2.q) < 1.0) {
                        candidates.push({ type: 'ult', p1: p1.t < p2.t ? p1 : p2, p2: p1.t > p2.t ? p1 : p2 });
                    }
                }
            }
        }
        return candidates;
    }

    let isVerde = modalidadeTarifaria && modalidadeTarifaria.toUpperCase().includes('VERDE');

    if (isVerde && demandaPairs.length > 0) {
        // VERDE: handle demand with possible "Não Consumida" or "Ultrapassagem" split
        let textHasNcVerde = /Não\s+Consumida/i.test(text);
        let textHasUltrapVerde = /Ultrap/i.test(text);

        if (contForaPonta > 0) {
            if (textHasNcVerde) {
                // NC has priority: look for two pairs summing to contratada (medida + não consumida)
                let bestNcPair: { medida: any; nc: any } | null = null;
                let bestNcError = Infinity;
                for (let i = 0; i < demandaPairs.length; i++) {
                    for (let j = i + 1; j < demandaPairs.length; j++) {
                        let pi = demandaPairs[i], pj = demandaPairs[j];
                        let sumError = Math.abs(pi.q + pj.q - contForaPonta);
                        if (sumError <= 0.01 && sumError < bestNcError) {
                            bestNcError = sumError;
                            bestNcPair = pi.q > pj.q
                                ? { medida: pi, nc: pj }
                                : { medida: pj, nc: pi };
                        }
                    }
                }
                if (bestNcPair) {
                    demandaPotenciaMedidaForaPonta = formatNumberToBR(bestNcPair.medida.q);
                    valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(bestNcPair.medida.v);
                    demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNcPair.nc.q);
                    valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(bestNcPair.nc.v);
                } else {
                    // No two pairs sum to contratada. Check if "Medida" even exists in Itens
                    let textHasMedidaFP = /Medida\s*-\s*F(?:ora)?\s*Ponta/i.test(text);
                    let singleNcMatch = demandaPairs.find(p => Math.abs(p.q - contForaPonta) < 2.0 && p.q < contForaPonta);
                    if (!singleNcMatch) singleNcMatch = demandaPairs.find(p => Math.abs(p.q - contForaPonta) < 1.0);

                    if (singleNcMatch && singleNcMatch.q < contForaPonta) {
                        demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(singleNcMatch.q);
                        valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(singleNcMatch.v);
                        if (textHasMedidaFP) {
                            demandaPotenciaMedidaForaPonta = formatNumberToBR(contForaPonta - singleNcMatch.q);
                            valorDemandaPotenciaMedidaForaPonta = "0,00";
                        }
                    } else if (!textHasMedidaFP) {
                        if (singleNcMatch) {
                            demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(singleNcMatch.q);
                            valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(singleNcMatch.v);
                        }
                    } else {
                        demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
                        valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
                    }
                }
            } else if (textHasUltrapVerde) {
                // Ultrap has priority: medida > contratada. Find two pairs where larger.q - smaller.q ≈ contratada
                let bestUltPair: { medida: any; ultrap: any } | null = null;
                let bestUltError = Infinity;
                for (let i = 0; i < demandaPairs.length; i++) {
                    for (let j = i + 1; j < demandaPairs.length; j++) {
                        let pi = demandaPairs[i], pj = demandaPairs[j];
                        let larger = pi.q > pj.q ? pi : pj;
                        let smaller = pi.q > pj.q ? pj : pi;
                        let diffError = Math.abs(larger.q - smaller.q - contForaPonta);
                        if (diffError < 1.0 && diffError < bestUltError) {
                            bestUltError = diffError;
                            bestUltPair = { medida: larger, ultrap: smaller };
                        }
                    }
                }
                if (bestUltPair) {
                    demandaPotenciaMedidaForaPonta = formatNumberToBR(bestUltPair.medida.q);
                    valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(bestUltPair.medida.v);
                    ultrapassagemForaPontaKW = formatNumberToBR(bestUltPair.ultrap.q);
                    valorUltrapassagemForaPonta = formatNumberToBR(bestUltPair.ultrap.v);
                } else {
                    demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
                    valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
                }
            } else {
                // No nc or ultrap signal: try single match (medida ≈ contratada), then fallback
                let singleMatch = demandaPairs.find(p => Math.abs(p.q - contForaPonta) < 1.0);
                if (singleMatch) {
                    demandaPotenciaMedidaForaPonta = formatNumberToBR(singleMatch.q);
                    valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(singleMatch.v);
                } else {
                    demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
                    valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
                }
            }
        } else {
            // No contracted demand found, fallback to lowest error
            demandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].q);
            valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(demandaPairs[0].v);
        }
    } else {
        // AZUL: use candidate matching for Ponta + Fora Ponta
        let candPonta = getCandidates(contPonta);
        let candFora = getCandidates(contForaPonta);

        let validCombinations: any[] = [];

        if (candPonta.length > 0 && candFora.length > 0) {
            for (let cp of candPonta) {
                for (let cf of candFora) {
                    if (cp.p1 === cf.p1 || cp.p1 === cf.p2 || (cp.p2 && (cp.p2 === cf.p1 || cp.p2 === cf.p2))) continue;
                    if (cp.p1.t <= cf.p1.t) continue;
                    validCombinations.push({ cp, cf });
                }
            }
        } else if (candPonta.length > 0) {
            for (let cp of candPonta) validCombinations.push({ cp, cf: null });
        } else if (candFora.length > 0) {
            for (let cf of candFora) validCombinations.push({ cp: null, cf });
        }

        validCombinations.sort((a: any, b: any) => {
            const getDiff = (c: any, cont: number) => {
                if (!c) return 0;
                let sumQ = c.type === 'nc' ? c.p1.q + c.p2.q : c.type === 'ult' ? c.p1.q - c.p2.q : c.p1.q;
                return Math.abs(sumQ - cont);
            };

            let diffA = getDiff(a.cp, contPonta) + getDiff(a.cf, contForaPonta);
            let diffB = getDiff(b.cp, contPonta) + getDiff(b.cf, contForaPonta);

            if (Math.abs(diffA - diffB) > 0.01) {
                return diffA - diffB;
            }

            let tA = a.cp ? a.cp.p1.t : 0;
            let tB = b.cp ? b.cp.p1.t : 0;
            if (tB !== tA) return tB - tA;
            return 0;
        });

        // Check if the PDF text explicitly mentions "Não Consumida" for Ponta and/or Fora Ponta
        let textHasNcPonta = /Não\s+Consumida\s*-\s*Ponta/i.test(text);
        let textHasNcForaPonta = /Não\s+Consumida\s*-\s*F(?:ora)?\s*Ponta/i.test(text);

        if (validCombinations.length > 0) {
            let best = validCombinations[0];
            if (best.cp) {
                usedPairs.add(best.cp.p1);
                if (best.cp.p2) usedPairs.add(best.cp.p2);
                demandaPotenciaMedidaPonta = formatNumberToBR(best.cp.p1.q);
                valorDemandaPotenciaMedidaPonta = formatNumberToBR(best.cp.p1.v);
                if (best.cp.type === 'nc' && textHasNcPonta && demandaPotenciaNaoConsumidaPonta === "0") {
                    demandaPotenciaNaoConsumidaPonta = formatNumberToBR(best.cp.p2.q);
                    valorDemandaPotenciaNaoConsumidaPonta = formatNumberToBR(best.cp.p2.v);
                }
                if (best.cp.type === 'ult') {
                    ultrapassagemPontaKW = formatNumberToBR(best.cp.p2.q);
                    valorUltrapassagemPonta = formatNumberToBR(best.cp.p2.v);
                }
            }
            if (best.cf) {
                usedPairs.add(best.cf.p1);
                if (best.cf.p2) usedPairs.add(best.cf.p2);
                demandaPotenciaMedidaForaPonta = formatNumberToBR(best.cf.p1.q);
                valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(best.cf.p1.v);
                if (best.cf.type === 'nc' && textHasNcForaPonta && demandaPotenciaNaoConsumidaForaPonta === "0") {
                    demandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(best.cf.p2.q);
                    valorDemandaPotenciaNaoConsumidaForaPonta = formatNumberToBR(best.cf.p2.v);
                }
                if (best.cf.type === 'ult') {
                    ultrapassagemForaPontaKW = formatNumberToBR(best.cf.p2.q);
                    valorUltrapassagemForaPonta = formatNumberToBR(best.cf.p2.v);
                }
            }
        } else if (demandaPairs.length > 0) {
            let validDemandaPairs = [...demandaPairs].sort((a, b) => {
                let aIdx = text.indexOf(formatNumberToBR(a.v));
                let bIdx = text.indexOf(formatNumberToBR(b.v));
                if (aIdx === -1) aIdx = 999999;
                if (bIdx === -1) bIdx = 999999;
                if (aIdx !== bIdx) return aIdx - bIdx;
                return Math.abs(a.q * a.t - a.v) - Math.abs(b.q * b.t - b.v);
            });

            let uniquePairs = [];
            for (let p of validDemandaPairs) {
                if (!uniquePairs.some(up => up.v === p.v)) {
                    uniquePairs.push(p);
                }
            }

            let topAzul = uniquePairs.slice(0, 2);
            topAzul.sort((a, b) => b.t - a.t);
            if (topAzul.length >= 2 && topAzul[0].t !== topAzul[1].t) {
                demandaPotenciaMedidaPonta = formatNumberToBR(topAzul[0].q);
                valorDemandaPotenciaMedidaPonta = formatNumberToBR(topAzul[0].v);
                demandaPotenciaMedidaForaPonta = formatNumberToBR(topAzul[1].q);
                valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(topAzul[1].v);
            } else {
                demandaPotenciaMedidaForaPonta = formatNumberToBR(topAzul[0].q);
                valorDemandaPotenciaMedidaForaPonta = formatNumberToBR(topAzul[0].v);
            }
        }
    }
  }
  
  // Tributos (Aproximação baseada na leitura bagunçada do PDF)
  let pis = "0";
  let cofins = "0";
  let icms = "0";
  
  // Para faturas da Energisa, os valores de PIS, COFINS e ICMS costumam aparecer nesta sequência específica no rodapé
  const tributosMatch = text.match(/([\d\.,]+)\s+([\d\.,]+)(?:\s+[\d\.,]+){6}\s+([\d\.,]+)\s+PIS ICMS COFINS/i);
  if (tributosMatch) {
    pis = tributosMatch[1];
    cofins = tributosMatch[2];
    icms = tributosMatch[3];
  } else {
    // Fallback genérico para outros layouts ou concessionárias
    icms = extract(/ICMS\s*\(R\$\).*?\d+\,\d+\s+([\d\.,]+)/i) || extract(/TOTAL:.*?\n.*?([\d\.,]+)\s+PIS/i) || "0";
    pis = extract(/PIS\s*\(R\$\).*?\d+\,\d+\s+([\d\.,]+)/i) || "0";
    cofins = extract(/COFINS\s*\(R\$\).*?\d+\,\d+\s+([\d\.,]+)/i) || "0";
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
    
    const negativeMatches = Array.from(text.matchAll(/-\d{1,3}(?:\.\d{3})*,\d{2}/g)).map(m => m[0]);
    // Valores reais da energia injetada costumam ser intercalados com valores menores (deduções de impostos)
    const gdiValuesR$ = negativeMatches.filter((_, i) => i % 2 === 0).map(v => parseFloat(v.replace(/\./g, "").replace(",", ".")));
    
    // Para achar o kWh, procuramos os números terminados em ,00 mais próximos do valor em R$ (tarifa é ~1.0)
    const kwhCandidates = Array.from(text.matchAll(/(?<![\d\.-])\d{1,5}(?:\.\d{3})*,\d{2}(?![\d\,])/g)).map(m => parseFloat(m[0].replace(/\./g, "").replace(",", ".")));

    let currentIndex = 0;
    
    // Extract all potential tariffs
    const tariffMatches = Array.from(text.matchAll(/\d{1,3},\d{6}/g)).map(m => parseFloat(m[0].replace(",", ".")));
    const uniqueTariffs = Array.from(new Set(tariffMatches)).filter(t => t > 0);

    // Processa oUC
    for (let i = 0; i < oUcCount; i++) {
      if (currentIndex < gdiValuesR$.length) {
        const rVal = gdiValuesR$[currentIndex];
        valorEnergiaAtvInjetadaGDIOUC += rVal;
        
        const targetKwh = Math.abs(rVal);
        let bestKwh = 0;
        let minDiff = Infinity;
        
        // Pass 1: Tenta casar a quantidade com uma tarifa da fatura (precisão alta)
        for (const cand of kwhCandidates) {
            for (const t of uniqueTariffs) {
                const expectedVal = cand * t;
                const diff = Math.abs(expectedVal - targetKwh);
                if (diff < minDiff && diff < 0.5) { 
                    minDiff = diff;
                    bestKwh = cand;
                }
            }
        }
        
        // Pass 2: Fallback para o valor mais próximo que seja razoável
        if (minDiff === Infinity) {
            for (const cand of kwhCandidates) {
              const diff = Math.abs(cand - targetKwh);
              if (diff < minDiff && diff < targetKwh * 0.4) {
                minDiff = diff;
                bestKwh = cand;
              }
            }
        }
        
        energiaAtvInjetadaGDIOUC += bestKwh;
        currentIndex++;
      }
    }
    
    // Processa mUC
    for (let i = 0; i < mUcCount; i++) {
      if (currentIndex < gdiValuesR$.length) {
        const rVal = gdiValuesR$[currentIndex];
        valorEnergiaAtvInjetadaGDIMUC += rVal;
        
        console.log(`[GDI ALGO V3] mUC[${i}] -> rVal:`, rVal, "kwhCandidates:", kwhCandidates);
        
        const targetKwh = Math.abs(rVal);
        let bestKwh = 0;
        let minDiff = Infinity;
        
        // Pass 1: Tenta casar a quantidade com uma tarifa da fatura (precisão alta)
        for (const cand of kwhCandidates) {
            for (const t of uniqueTariffs) {
                const expectedVal = cand * t;
                const diff = Math.abs(expectedVal - targetKwh);
                if (diff < minDiff && diff < 0.5) { 
                    minDiff = diff;
                    bestKwh = cand;
                }
            }
        }
        
        // Pass 2: Fallback para o valor mais próximo que seja razoável
        if (minDiff === Infinity) {
            for (const cand of kwhCandidates) {
              const diff = Math.abs(cand - targetKwh);
              if (diff < minDiff && diff < targetKwh * 0.4) {
                minDiff = diff;
                bestKwh = cand;
              }
            }
        }
        
        energiaAtvInjetadaGDIMUC += bestKwh;
        currentIndex++;
      }
    }
  }

  // Set consumoKwh for Grupo B (B3, Convencional)
  const isGrupoB = (modalidadeTarifaria && modalidadeTarifaria.toUpperCase().includes("CONVENCIONAL")) || 
                   (subgrupo && subgrupo.toUpperCase().includes("B3"));
  if (isGrupoB && (!consumoKwhForaPonta || consumoKwhForaPonta === "0")) {
    let extractedConsumo = extract(/Energia ativa em kWh\s+[\d\.,]+\s+([\d\.,]+)/i) || extract(/Consumo em kWh\s+([\d\.,]+)/i);
    if (extractedConsumo && extractedConsumo !== "0") {
      consumoKwh = extractedConsumo;
    } else {
      consumoKwh = consumoKwhForaPonta;
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
