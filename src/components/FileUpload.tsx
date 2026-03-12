import React, { useCallback, useState } from 'react';
import { UploadCloud, AlertCircle, FileText } from 'lucide-react';
import { parseCSV, EnergyData } from '../utils/parser';
import Papa from 'papaparse';
import { supabase, isSupabaseConfigured } from '../lib/supabase';

interface FileUploadProps {
  onDataLoaded: (data: EnergyData[]) => void;
}

export default function FileUpload({ onDataLoaded }: FileUploadProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [csvText, setCsvText] = useState('');

  const handleFile = async (file: File) => {
    if (!file.name.endsWith('.csv')) {
      setError('Por favor, envie um arquivo CSV válido.');
      return;
    }

    setIsLoading(true);
    setError(null);
    try {
      let user = null;
      if (isSupabaseConfigured) {
        const { data: { user: supabaseUser } } = await supabase.auth.getUser();
        user = supabaseUser;
      }
      
      if (!user && isSupabaseConfigured) throw new Error("Usuário não autenticado");

      const data = await parseCSV(file);
      
      if (isSupabaseConfigured && user) {
        const { error: supabaseError } = await supabase
          .from('energy_invoices')
          .insert(data.map(d => ({
            uc: d.uc,
            ano: d.ano,
            mes: d.mes,
            consumo_ponta: d.consumoPonta,
            valor_ponta: d.valorPonta,
            consumo_fora_ponta: d.consumoForaPonta,
            valor_fora_ponta: d.valorForaPonta,
            valor_total: d.valorTotal,
            cidade: d.cidade,
            user_id: user.id
          })));
        
        if (supabaseError) throw supabaseError;
      }
      
      onDataLoaded(data);
    } catch (err) {
      setError('Erro ao processar o arquivo. Verifique o formato.');
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  };

  const handleTextSubmit = () => {
    if (!csvText.trim()) {
      setError('Cole os dados CSV na área de texto.');
      return;
    }
    setIsLoading(true);
    setError(null);
    
    const parseNumber = (val: string) => {
      if (!val) return 0;
      const cleanVal = val.replace(/\./g, '').replace(',', '.');
      const num = parseFloat(cleanVal);
      return isNaN(num) ? 0 : num;
    };

    Papa.parse(csvText, {
      header: true,
      skipEmptyLines: true,
      delimiter: ';',
      complete: async (results) => {
        try {
          let user = null;
          if (isSupabaseConfigured) {
            const { data: { user: supabaseUser } } = await supabase.auth.getUser();
            user = supabaseUser;
          }
          
          if (!user && isSupabaseConfigured) throw new Error("Usuário não autenticado");

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
          
          if (isSupabaseConfigured && user) {
            const { error: supabaseError } = await supabase
              .from('energy_invoices')
              .insert(data.map(d => ({
                uc: d.uc,
                ano: d.ano,
                mes: d.mes,
                consumo_ponta: d.consumoPonta,
                valor_ponta: d.valorPonta,
                consumo_fora_ponta: d.consumoForaPonta,
                valor_fora_ponta: d.valorForaPonta,
                valor_total: d.valorTotal,
                cidade: d.cidade,
                user_id: user.id
              })));
              
            if (supabaseError) {
              setError('Erro ao salvar no banco de dados.');
              console.error(supabaseError);
              setIsLoading(false);
              return;
            }
          }
          
          onDataLoaded(data);
          setIsLoading(false);
        } catch (err) {
          setError('Erro ao processar os dados. Verifique se o formato está correto.');
          setIsLoading(false);
        }
      },
      error: (err) => {
        setError('Erro ao processar os dados. Verifique se o formato está correto.');
        setIsLoading(false);
      }
    });
  };

  const onDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFile(e.dataTransfer.files[0]);
    }
  }, []);

  return (
    <div className="flex flex-col items-center justify-center py-10">
      <div className="text-center mb-8">
        <h2 className="text-3xl font-bold text-slate-800 mb-3">Análise de Faturas de Energia</h2>
        <p className="text-slate-500 max-w-lg mx-auto">
          Faça o upload do seu arquivo CSV ou cole os dados diretamente abaixo para visualizar os dashboards.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full max-w-5xl">
        {/* Upload Area */}
        <div
          onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
          onDragLeave={() => setIsDragging(false)}
          onDrop={onDrop}
          className={`border-2 border-dashed rounded-2xl p-10 text-center transition-all duration-200 flex flex-col items-center justify-center ${
            isDragging 
              ? 'border-emerald-500 bg-emerald-50' 
              : 'border-slate-300 bg-white hover:border-emerald-400 hover:bg-slate-50'
          }`}
        >
          <div className={`p-4 rounded-full mb-4 ${isDragging ? 'bg-emerald-100' : 'bg-slate-100'}`}>
            <UploadCloud className={`w-10 h-10 ${isDragging ? 'text-emerald-600' : 'text-slate-400'}`} />
          </div>
          
          <p className="text-lg font-medium text-slate-700 mb-1">
            Arraste e solte seu arquivo CSV
          </p>
          <p className="text-sm text-slate-500 mb-6">
            ou clique para selecionar
          </p>

          <label className="relative cursor-pointer bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-3 rounded-xl font-medium transition-colors shadow-sm">
            <span>Selecionar Arquivo</span>
            <input 
              type="file" 
              accept=".csv" 
              className="hidden" 
              onChange={(e) => e.target.files && handleFile(e.target.files[0])}
              disabled={isLoading}
            />
          </label>
        </div>

        {/* Paste Area */}
        <div className="bg-white border border-slate-200 rounded-2xl p-6 flex flex-col shadow-sm">
          <div className="flex items-center gap-3 mb-4">
            <div className="p-2 bg-blue-100 rounded-lg">
              <FileText className="w-5 h-5 text-blue-600" />
            </div>
            <h3 className="text-lg font-medium text-slate-800">Colar Dados CSV</h3>
          </div>
          <textarea
            className="flex-1 w-full border border-slate-200 rounded-xl p-4 text-sm font-mono text-slate-600 focus:ring-2 focus:ring-emerald-500 focus:border-emerald-500 outline-none resize-none min-h-[200px]"
            placeholder="Cole os dados do seu arquivo CSV aqui (incluindo o cabeçalho)..."
            value={csvText}
            onChange={(e) => setCsvText(e.target.value)}
            disabled={isLoading}
          />
          <button
            onClick={handleTextSubmit}
            disabled={isLoading || !csvText.trim()}
            className="mt-4 w-full bg-slate-800 hover:bg-slate-900 disabled:bg-slate-300 text-white px-6 py-3 rounded-xl font-medium transition-colors shadow-sm"
          >
            Processar Texto
          </button>
        </div>
      </div>

      {error && (
        <div className="mt-6 flex items-center gap-2 text-red-600 bg-red-50 px-4 py-3 rounded-lg max-w-2xl w-full">
          <AlertCircle className="w-5 h-5 flex-shrink-0" />
          <span className="font-medium">{error}</span>
        </div>
      )}
      
      {isLoading && (
        <div className="mt-6 flex items-center gap-3 text-emerald-600">
          <div className="w-5 h-5 border-2 border-emerald-600 border-t-transparent rounded-full animate-spin"></div>
          <span className="font-medium">Processando dados...</span>
        </div>
      )}
    </div>
  );
}
