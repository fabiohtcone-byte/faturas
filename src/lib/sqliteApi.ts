import { BillData } from '../types';

const getApiBase = () => {
  if (typeof window !== 'undefined' && window.location.port === '3000') {
    return '/api';
  }
  return 'http://localhost:3001/api';
};

const API_BASE = getApiBase();

export interface SqliteHealth {
  connected: boolean;
  totalBills: number;
  totalUcs: number;
  totalMappings: number;
  databasePath?: string;
  error?: string;
}

export async function checkSqliteHealth(): Promise<SqliteHealth> {
  try {
    const res = await fetch(`${API_BASE}/health`, { signal: AbortSignal.timeout(3000) });
    if (!res.ok) throw new Error(`HTTP error ${res.status}`);
    const data = await res.json();
    return {
      connected: data.status === 'ok',
      totalBills: data.totalBills || 0,
      totalUcs: data.totalUcs || 0,
      totalMappings: data.totalMappings || 0,
      databasePath: data.databasePath,
    };
  } catch (err: any) {
    return {
      connected: false,
      totalBills: 0,
      totalUcs: 0,
      totalMappings: 0,
      error: err.message || 'SQLite server indisponível',
    };
  }
}

export async function fetchBillsSqlite(): Promise<BillData[]> {
  try {
    const res = await fetch(`${API_BASE}/bills`, { signal: AbortSignal.timeout(10000) });
    if (!res.ok) throw new Error(`HTTP error ${res.status}`);
    const rows = await res.json();
    if (!Array.isArray(rows)) return [];
    
    return rows.map((dbBill: any): BillData => ({
      id: String(dbBill.id),
      fileName: dbBill.file_name || 'Sem nome',
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
      consumoKwhGrupoB: dbBill.consumo_kwh_grupo_b || '',
      valorConsumoKwhGrupoB: dbBill.valor_consumo_kwh_grupo_b || '',
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
      modalidadeTarifaria: dbBill.modalidade_tarifaria || '',
      subgrupo: dbBill.subgrupo || '',
      tipo: dbBill.tipo || '',
      mercado: dbBill.mercado || '',
      gerencia: dbBill.gerencia || '',
      locin: dbBill.locins || '',
      dataVencimento: dbBill.data_vencimento || '',
      status: (dbBill.status as any) || 'completed',
      error: dbBill.error || undefined,
      createdAt: dbBill.created_at ? new Date(dbBill.created_at).getTime() : Date.now(),
    }));
  } catch (err) {
    console.warn('[SQLite] Não foi possível carregar faturas do SQLite:', err);
    return [];
  }
}

export async function saveBillSqlite(bill: BillData): Promise<boolean> {
  try {
    const res = await fetch(`${API_BASE}/bills`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bill),
    });
    return res.ok;
  } catch (err) {
    console.warn('[SQLite] Erro ao salvar fatura:', err);
    return false;
  }
}

export async function saveBillsBatchSqlite(bills: BillData[]): Promise<boolean> {
  if (!bills || bills.length === 0) return true;
  try {
    const res = await fetch(`${API_BASE}/bills/batch`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(bills),
    });
    return res.ok;
  } catch (err) {
    console.warn('[SQLite] Erro ao salvar lote de faturas:', err);
    return false;
  }
}

export async function deleteBillSqlite(id: string): Promise<boolean> {
  try {
    const res = await fetch(`${API_BASE}/bills/${id}`, {
      method: 'DELETE',
    });
    return res.ok;
  } catch (err) {
    console.warn('[SQLite] Erro ao deletar fatura:', err);
    return false;
  }
}

export async function deleteBillsBatchSqlite(ids: string[]): Promise<boolean> {
  if (!ids || ids.length === 0) return true;
  try {
    const res = await fetch(`${API_BASE}/bills/delete-batch`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ids }),
    });
    return res.ok;
  } catch (err) {
    console.warn('[SQLite] Erro ao deletar lote de faturas:', err);
    return false;
  }
}

export async function clearAllBillsSqlite(): Promise<boolean> {
  try {
    const res = await fetch(`${API_BASE}/bills/all/clear`, {
      method: 'DELETE',
    });
    return res.ok;
  } catch (err) {
    console.warn('[SQLite] Erro ao limpar faturas:', err);
    return false;
  }
}

export async function fetchUcMappingsSqlite(): Promise<any[]> {
  try {
    const res = await fetch(`${API_BASE}/uc-mappings`);
    if (!res.ok) return [];
    return await res.json();
  } catch {
    return [];
  }
}

export async function saveUcMappingsSqlite(mappings: any[]): Promise<boolean> {
  try {
    const res = await fetch(`${API_BASE}/uc-mappings`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(mappings),
    });
    return res.ok;
  } catch {
    return false;
  }
}
