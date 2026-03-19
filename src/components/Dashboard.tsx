import React, { useMemo } from 'react';
import { EnergyData } from '../utils/parser';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  Cell, LabelList
} from 'recharts';
import { Zap, DollarSign, TrendingUp, TrendingDown } from 'lucide-react';

interface DashboardProps {
  data: EnergyData[];
}

const formatCurrency = (value: number) => 
  new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);

const formatNumber = (value: number) => 
  new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 }).format(value);

const COLORS = {
  blue: '#3b82f6',
  green: '#10b981',
  orange: '#f59e0b',
  purple: '#8b5cf6',
  slate: '#64748b'
};

export default function Dashboard({ data }: DashboardProps) {
  
  const categories = useMemo(() => {
    // Classification logic based on the image categories
    const cats: Record<string, { name: string, custo: number, consumo: number }> = {
      semCompensacao: { name: 'UC\'s sem Compensação', custo: 0, consumo: 0 },
      geral: { name: 'Geral', custo: 0, consumo: 0 },
      consumosMinimos: { name: 'Consumos Mínimos', custo: 0, consumo: 0 },
      ppp: { name: 'PPP Fotovoltaica', custo: 0, consumo: 0 },
      comCompensacao: { name: 'UC\'s com Compensação', custo: 0, consumo: 0 },
      usinas: { name: 'Usinas Sanesul', custo: 0, consumo: 0 },
    };

    data.forEach(row => {
      // Simplified classification for demonstration
      const consumo = row.consumoPonta + row.consumoForaPonta;
      if (row.uc.startsWith('PPP')) {
        cats.ppp.custo += row.valorTotal;
        cats.ppp.consumo += consumo;
      } else if (row.uc.startsWith('USINA')) {
        cats.usinas.custo += row.valorTotal;
        cats.usinas.consumo += consumo;
      } else if (row.valorTotal < 150 && consumo <= 100) {
        cats.consumosMinimos.custo += row.valorTotal;
        cats.consumosMinimos.consumo += consumo;
      } else {
        cats.geral.custo += row.valorTotal;
        cats.geral.consumo += consumo;
      }
    });
    return cats;
  }, [data]);

  const chartData = Object.entries(categories).map(([key, val]: [string, { name: string, custo: number, consumo: number }]) => ({
    name: val.name,
    custo: val.custo,
    consumo: val.consumo,
    tarifa: val.consumo > 0 ? val.custo / val.consumo : 0
  }));

  return (
    <div className="space-y-8 p-6 bg-slate-50">
      {/* Top Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Custos Chart */}
        <div className="bg-white rounded-3xl p-6 shadow-sm border border-slate-100">
          <h3 className="text-lg font-bold text-slate-800 mb-6 text-center">Custos (R$)</h3>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="name" hide />
              <Tooltip formatter={(value: number) => formatCurrency(value)} />
              <Bar dataKey="custo" fill={COLORS.blue} radius={[4, 4, 0, 0]}>
                {chartData.map((entry, index) => <Cell key={index} fill={Object.values(COLORS)[index % 5]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Consumo Chart */}
        <div className="bg-white rounded-3xl p-6 shadow-sm border border-slate-100">
          <h3 className="text-lg font-bold text-slate-800 mb-6 text-center">Consumo (kWh)</h3>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="name" hide />
              <Tooltip formatter={(value: number) => formatNumber(value)} />
              <Bar dataKey="consumo" fill={COLORS.green} radius={[4, 4, 0, 0]}>
                {chartData.map((entry, index) => <Cell key={index} fill={Object.values(COLORS)[index % 5]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Tarifa Chart */}
        <div className="bg-white rounded-3xl p-6 shadow-sm border border-slate-100">
          <h3 className="text-lg font-bold text-slate-800 mb-6 text-center">Tarifa Média (R$/kWh)</h3>
          <ResponsiveContainer width="100%" height={250}>
            <BarChart data={chartData}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="name" hide />
              <Tooltip formatter={(value: number) => formatCurrency(value)} />
              <Bar dataKey="tarifa" fill={COLORS.slate} radius={[4, 4, 0, 0]}>
                {chartData.map((entry, index) => <Cell key={index} fill={Object.values(COLORS)[index % 5]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4">
        {Object.entries(categories).map(([key, cat]: [string, { name: string, custo: number, consumo: number }]) => (
          <div key={key} className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
            <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-2">{cat.name}</p>
            <p className="text-lg font-bold text-slate-900">{formatCurrency(cat.custo)}</p>
            <p className="text-xs text-slate-500">{formatNumber(cat.consumo)} kWh</p>
          </div>
        ))}
      </div>
    </div>
  );
}
