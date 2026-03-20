import React, { useMemo } from 'react';
import { EnergyData } from '../utils/parser';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  LineChart, Line, Cell
} from 'recharts';
import { Zap, DollarSign, MapPin, Building2, TrendingUp, TrendingDown } from 'lucide-react';

interface DashboardProps {
  data: EnergyData[];
}

const formatCurrency = (value: number) => 
  new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);

const formatNumber = (value: number) => 
  new Intl.NumberFormat('pt-BR').format(value);

const COLORS = ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#14b8a6', '#f97316'];

const MetricCard = ({ title, custo, consumo, subValue, icon: Icon, colorClass, bgClass }: any) => (
  <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100 flex items-center gap-4 group hover:shadow-md transition-all">
    <div className={`w-12 h-12 rounded-xl ${bgClass} flex items-center justify-center ${colorClass} group-hover:scale-110 transition-transform`}>
      <Icon size={24} />
    </div>
    <div>
      <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider mb-1">{title}</h4>
      {custo !== undefined && (
        <p className="text-xl font-bold text-slate-900">{formatCurrency(custo)}</p>
      )}
      {consumo !== undefined && (
        <p className="text-xl font-bold text-slate-900">{formatNumber(consumo)} <span className="text-xs text-slate-400 font-medium">kWh</span></p>
      )}
      {subValue !== undefined && (
        <p className="text-xl font-bold text-slate-900">{subValue}</p>
      )}
    </div>
  </div>
);

export default function Dashboard({ data }: DashboardProps) {
  
  const stats = useMemo(() => {
    let totalGasto = 0;
    let totalConsumoPonta = 0;
    let totalConsumoForaPonta = 0;
    const ucs = new Set<string>();
    const cidades = new Set<string>();

    data.forEach(row => {
      totalGasto += row.valorTotal;
      totalConsumoPonta += row.consumoPonta;
      totalConsumoForaPonta += row.consumoForaPonta;
      if (row.uc) ucs.add(row.uc);
      if (row.cidade) cidades.add(row.cidade.toUpperCase());
    });

    return {
      totalGasto,
      totalConsumo: totalConsumoPonta + totalConsumoForaPonta,
      totalConsumoPonta,
      totalConsumoForaPonta,
      numUcs: ucs.size,
      numCidades: cidades.size
    };
  }, [data]);

  const cityData = useMemo(() => {
    const map = new Map<string, number>();
    data.forEach(row => {
      if (!row.cidade) return;
      const city = row.cidade.toUpperCase();
      map.set(city, (map.get(city) || 0) + row.valorTotal);
    });
    return Array.from(map.entries())
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 10);
  }, [data]);

  const monthData = useMemo(() => {
    const map = new Map<string, { name: string, gasto: number, consumo: number }>();
    data.forEach(row => {
      if (!row.mes || !row.ano) return;
      const key = `${row.mes}/${row.ano}`;
      const existing = map.get(key) || { name: key, gasto: 0, consumo: 0 };
      existing.gasto += row.valorTotal;
      existing.consumo += (row.consumoPonta + row.consumoForaPonta);
      map.set(key, existing);
    });
    return Array.from(map.values());
  }, [data]);

  const ucData = useMemo(() => {
    const map = new Map<string, { uc: string, cidade: string, gasto: number, consumo: number }>();
    data.forEach(row => {
      if (!row.uc) return;
      const existing = map.get(row.uc) || { uc: row.uc, cidade: row.cidade.toUpperCase(), gasto: 0, consumo: 0 };
      existing.gasto += row.valorTotal;
      existing.consumo += (row.consumoPonta + row.consumoForaPonta);
      map.set(row.uc, existing);
    });
    return Array.from(map.values()).sort((a, b) => b.gasto - a.gasto).slice(0, 50);
  }, [data]);

  return (
    <div className="space-y-8 p-6 bg-slate-50 min-h-screen">
      <header>
        <h1 className="text-3xl font-bold text-slate-900">Análise de Insumo de Energia Elétrica</h1>
        <p className="text-slate-500 mt-1">Visão gerencial consolidada de custos e consumos.</p>
      </header>

      {/* KPI Cards */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <MetricCard title="Custo Total" custo={stats.totalGasto} icon={DollarSign} colorClass="text-emerald-600" bgClass="bg-emerald-100" />
        <MetricCard title="Consumo Total" consumo={stats.totalConsumo} icon={Zap} colorClass="text-blue-600" bgClass="bg-blue-100" />
        <MetricCard title="Unidades Consumidoras" subValue={stats.numUcs.toString()} icon={Building2} colorClass="text-purple-600" bgClass="bg-purple-100" />
        <MetricCard title="Cidades Atendidas" subValue={stats.numCidades.toString()} icon={MapPin} colorClass="text-orange-600" bgClass="bg-orange-100" />
      </div>

      {/* Charts */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100">
          <h3 className="text-lg font-semibold text-slate-800 mb-6">Tendência de Custo</h3>
          <div className="h-80">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={monthData}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} />
                <YAxis 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{fill: '#64748b'}}
                  tickFormatter={(val) => `R$ ${(val/1000).toFixed(0)}k`}
                />
                <Tooltip 
                  formatter={(value: number) => formatCurrency(value)}
                  contentStyle={{borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}}
                />
                <Line type="monotone" dataKey="gasto" stroke="#10b981" strokeWidth={3} dot={{r: 4, fill: '#10b981', strokeWidth: 2, stroke: '#fff'}} activeDot={{r: 6}} />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white rounded-2xl p-6 shadow-sm border border-slate-100">
          <h3 className="text-lg font-semibold text-slate-800 mb-6">Top 10 Cidades por Custo</h3>
          <div className="h-80">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={cityData} layout="vertical" margin={{ left: 40 }}>
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                <XAxis type="number" axisLine={false} tickLine={false} tick={{fill: '#64748b'}} tickFormatter={(val) => `R$ ${(val/1000).toFixed(0)}k`} />
                <YAxis dataKey="name" type="category" axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 12}} width={100} />
                <Tooltip 
                  formatter={(value: number) => formatCurrency(value)}
                  cursor={{fill: '#f8fafc'}}
                  contentStyle={{borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}}
                />
                <Bar dataKey="value" fill="#3b82f6" radius={[0, 4, 4, 0]}>
                  {cityData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="px-6 py-5 border-b border-slate-100">
          <h3 className="text-lg font-semibold text-slate-800">Detalhamento por Unidade Consumidora</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50 text-slate-500 text-xs font-medium uppercase tracking-wider">
                <th className="px-6 py-4">UC</th>
                <th className="px-6 py-4">Cidade</th>
                <th className="px-6 py-4 text-right">Consumo (kWh)</th>
                <th className="px-6 py-4 text-right">Custo Total</th>
              </tr>
            </thead>
            <tbody className="text-slate-700 text-sm">
              {ucData.map((uc, i) => (
                <tr key={i} className="hover:bg-slate-50 transition-colors border-b border-slate-50 last:border-0">
                  <td className="px-6 py-4 font-medium text-slate-900">{uc.uc}</td>
                  <td className="px-6 py-4">{uc.cidade}</td>
                  <td className="px-6 py-4 text-right">{formatNumber(uc.consumo)}</td>
                  <td className="px-6 py-4 text-right font-semibold text-slate-900">{formatCurrency(uc.gasto)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
