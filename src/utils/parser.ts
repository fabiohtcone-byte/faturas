import Papa from 'papaparse';

export interface EnergyData {
  uc: string;
  ano: string;
  mes: string;
  consumoPonta: number;
  valorPonta: number;
  consumoForaPonta: number;
  valorForaPonta: number;
  valorTotal: number;
  cidade: string;
}

const parseNumber = (val: string) => {
  if (!val) return 0;
  // Remove dots (thousand separators) and replace comma with dot
  const cleanVal = val.replace(/\./g, '').replace(',', '.');
  const num = parseFloat(cleanVal);
  return isNaN(num) ? 0 : num;
};

export const parseCSV = (file: File): Promise<EnergyData[]> => {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      delimiter: ';',
      complete: (results) => {
        const data = results.data.map((row: any) => ({
          uc: row['UC'] || '',
          ano: row['Ano Leitura'] || '',
          mes: row['Mês Referência'] || '',
          consumoPonta: parseNumber(row['Consumo em kWh - Ponta']),
          valorPonta: parseNumber(row['Valor (R$) - Consumo em kWh - Ponta']),
          consumoForaPonta: parseNumber(row['Consumo em kWh - Fora Ponta']),
          valorForaPonta: parseNumber(row['Valor (R$) - Consumo em kWh - Fora Ponta']),
          valorTotal: parseNumber(row['Valor R$']),
          cidade: row['CIDADE'] || '',
        }));
        resolve(data);
      },
      error: (error) => {
        reject(error);
      }
    });
  });
};
