export interface BillData {
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
  status: 'pending' | 'processing' | 'completed' | 'error';
  error?: string;
  file?: File;
  progress?: number;
  abortController?: AbortController;
  name?: string; // This is added during grouping/mapping
}

export interface AgrupadoraData {
  id: string;
  nome: string;
  concessionaria: string;
  ucs: string[];
}
