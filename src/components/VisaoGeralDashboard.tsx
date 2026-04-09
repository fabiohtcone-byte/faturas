import React, { useState, useMemo, useCallback } from 'react';
import { 
  Activity, 
  Calendar, 
  DollarSign, 
  Zap, 
  Battery, 
  ZapOff, 
  Calculator, 
  ArrowDown, 
  TrendingUp, 
  Home, 
  BarChart2, 
  GitCompare, 
  LayoutDashboard, 
  LogOut,
  Leaf
} from 'lucide-react';
import { 
  ResponsiveContainer, 
  AreaChart, 
  Area, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  LineChart, 
  Line 
} from 'recharts';
import { UCS_PPP, UCS_USINA } from '../constants';
import { formatMonth, getMonthNumber, formatNumber } from '../utils/formatters';

interface VisaoGeralDashboardProps {
  data: any[];
  setCurrentPage: (page: 'visao_geral' | 'sistema') => void;
  handleLogout: () => void;
}

const VisaoGeralDashboard = ({ data, setCurrentPage, handleLogout }: VisaoGeralDashboardProps) => {
  const [selectedMonth, setSelectedMonth] = useState<string>('all');

  const availableMonths = Array.from(new Set(data.map(d => d.name))).filter(Boolean).sort((a, b) => {
    const [mA, yA] = String(a).split('/');
    const [mB, yB] = String(b).split('/');
    if (yA !== yB) return Number(yB) - Number(yA);
    return getMonthNumber(mB) - getMonthNumber(mA);
  });

  const filteredData = selectedMonth === 'all' ? data : data.filter(d => d.name === selectedMonth);

  const calc = (filterFn: (d: any) => boolean) => {
    const filtered = filteredData.filter(filterFn);
    const custo = filtered.reduce((acc, curr) => acc + curr.valorTotal, 0);
    const consumo = filtered.reduce((acc, curr) => acc + curr.consumoPonta + curr.consumoForaPonta, 0);
    return { custo, consumo, tarifa: consumo > 0 ? custo / consumo : 0 };
  };

  const isGrupoA = (d: any) => d.demandaContratadaPonta > 0 || d.demandaContratadaForaPonta > 0;
  const isGrupoB = (d: any) => !isGrupoA(d);
  const isLivre = (d: any) => (d.modalidadeTarifaria || '').includes('LIVRE') || d.tipo === 'LIVRE';
  const isCativo = (d: any) => !isLivre(d);
  
  const isAzul = (d: any) => (d.modalidadeTarifaria || '').includes('AZUL');
  const isVerde = (d: any) => (d.modalidadeTarifaria || '').includes('VERDE');
  const isOutrosGrupoA = (d: any) => isGrupoA(d) && !isAzul(d) && !isVerde(d);
  
  const hasCompensacao = (d: any) => d.solarInjetadaOUC > 0 || d.solarInjetadaMUC > 0;
  
  const isConsumoMinimo = (d: any) => isGrupoB(d) && (d.consumoPonta + d.consumoForaPonta) <= 100 && d.valorTotal < 150;
  const isPPP = (d: any) => isGrupoB(d) && UCS_PPP.has(String(d.uc));
  const isUsina = (d: any) => isGrupoB(d) && UCS_USINA.has(String(d.uc));
  const isOptanteB = (d: any) => isGrupoB(d) && (d.modalidadeTarifaria || '').toUpperCase().includes('OPTANTE');
  const isGeral = (d: any) => isGrupoB(d) && !isConsumoMinimo(d) && !isPPP(d) && !isUsina(d) && !isOptanteB(d);

  const totalGeral = calc(() => true);
  const grupoA = calc(isGrupoA);
  const grupoB = calc(isGrupoB);

  const livre = calc(d => isGrupoA(d) && isLivre(d));
  const livreAzul = calc(d => isGrupoA(d) && isLivre(d) && isAzul(d));
  const livreVerde = calc(d => isGrupoA(d) && isLivre(d) && isVerde(d));

  const cativo = calc(d => isGrupoA(d) && isCativo(d));
  const cativoAzul = calc(d => isGrupoA(d) && isCativo(d) && isAzul(d));
  const cativoVerde = calc(d => isGrupoA(d) && isCativo(d) && isVerde(d));
  const cativoOutras = calc(d => isGrupoA(d) && isCativo(d) && isOutrosGrupoA(d));

  const semCompensacao = calc(d => isGrupoB(d) && !hasCompensacao(d));
  const geral = calc(isGeral);
  const consumosMinimos = calc(isConsumoMinimo);
  const optanteB = calc(isOptanteB);

  const comCompensacao = calc(d => isGrupoB(d) && hasCompensacao(d));
  const ppp = calc(isPPP);
  const usinas = calc(isUsina);

  const totalSolarInjetada = filteredData.reduce((acc, curr) => acc + (curr.solarInjetadaOUC || 0) + (curr.solarInjetadaMUC || 0), 0);
  const emissoesEvitadas = totalSolarInjetada * 0.0426; // Fator médio do SIN

  const monthlyData = useMemo(() => {
    interface GroupedItem {
      name: string;
      month: number;
      year: number;
      consumo: number;
      custo: number;
    }

    const grouped = data.reduce((acc, curr) => {
      const name = curr.name; 
      if (!acc[name]) {
        const [month, year] = name.split('/');
        acc[name] = { 
          name, 
          month: getMonthNumber(month), 
          year: parseInt(year), 
          consumo: 0, 
          custo: 0 
        };
      }
      acc[name].consumo += curr.consumoPonta + curr.consumoForaPonta;
      acc[name].custo += curr.valorTotal;
      return acc;
    }, {} as Record<string, GroupedItem>);

    const sorted = (Object.values(grouped) as GroupedItem[]).sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      return a.month - b.month;
    });

    const ptAbbr: Record<string, string> = {
      'Janeiro': 'Jan', 'Fevereiro': 'Fev', 'Março': 'Mar', 'Abril': 'Abr',
      'Maio': 'Mai', 'Junho': 'Jun', 'Julho': 'Jul', 'Agosto': 'Ago',
      'Setembro': 'Set', 'Outubro': 'Out', 'Novembro': 'Nov', 'Dezembro': 'Dez'
    };

    return sorted.map(item => {
      const [month] = item.name.split('/');
      const fullMonth = formatMonth(month);
      const abbr = ptAbbr[fullMonth] || month.substring(0, 3);
      return {
        name: abbr,
        consumo: item.consumo,
        custo: item.custo
      };
    });
  }, [data]);

  const chartDomain = useMemo(() => {
    const maxConsumo = Math.max(...monthlyData.map(d => d.consumo), 0);
    const maxCusto = Math.max(...monthlyData.map(d => d.custo), 0);
    const max = Math.max(maxConsumo, maxCusto);
    return [0, Math.ceil(max * 1.1)];
  }, [monthlyData]);

  const sparklineDataAzul = monthlyData.map(m => ({ value: m.custo * 0.6 })); 
  const sparklineDataVerde = monthlyData.map(m => ({ value: m.custo * 0.4 })); 

  const MetricRow = ({ icon: Icon, label, value, unit, isCurrency }: { icon: any, label: string, value: number, unit?: string, isCurrency?: boolean }) => (
    <div className="flex items-center justify-between border-b border-slate-200 py-4 group/row">
      <div className="flex items-center gap-4">
        <div className="w-10 h-10 rounded-xl bg-slate-50 flex items-center justify-center border border-slate-100 group-hover/row:bg-blue-50 group-hover/row:border-blue-100 transition-all">
          <Icon size={20} className="text-slate-400 group-hover/row:text-blue-600 transition-colors" />
        </div>
        <span className="text-sm font-bold text-slate-600 uppercase tracking-wider">{label}</span>
      </div>
      <div className="text-right">
        <span className="text-2xl font-black text-slate-900 tracking-tight">
          {isCurrency && <span className="text-base text-slate-400 font-bold mr-1.5">R$</span>}
          {formatNumber(value, isCurrency, 2)}
          {unit && <span className="text-base text-slate-400 font-bold ml-1.5">{unit}</span>}
        </span>
      </div>
    </div>
  );

  const SummaryCard = ({ title, data, icon: Icon, color = "blue", className = "" }: { title: string, data: any, icon: any, color?: "blue" | "indigo" | "sky", className?: string }) => {
    const colorStyles = {
      blue: { icon: "text-blue-600", bg: "bg-blue-50", border: "border-blue-100" },
      indigo: { icon: "text-indigo-600", bg: "bg-indigo-50", border: "border-indigo-100" },
      sky: { icon: "text-sky-600", bg: "bg-sky-50", border: "border-sky-100" },
    }[color];

    return (
      <div className={`rounded-[2rem] p-5 relative overflow-hidden bg-white text-slate-900 border border-slate-200 group transition-all duration-500 hover:shadow-xl shadow-sm ${className}`}>
        <div className="absolute top-0 right-0 w-64 h-64 bg-slate-50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-30 group-hover:scale-110 transition-transform duration-700 blur-3xl"></div>
        
        <div className="flex justify-between items-center mb-4 relative z-10">
          <div className="flex items-center gap-4">
            <div className={`p-2.5 ${colorStyles.bg} rounded-xl border ${colorStyles.border} shadow-sm group-hover:bg-white transition-colors duration-300`}>
              <Icon size={20} className={colorStyles.icon} />
            </div>
            <h3 className="text-xs font-black uppercase tracking-[0.2em] text-slate-800">{title}</h3>
          </div>
        </div>
        
        <div className="space-y-1 relative z-10">
          <MetricRow icon={DollarSign} label="Custo Total" value={data.custo} isCurrency />
          <MetricRow icon={Zap} label="Consumo Total" value={data.consumo} unit="kWh" />
          <MetricRow icon={Calculator} label="Tarifa Média" value={data.tarifa} isCurrency />
        </div>
      </div>
    );
  };

  const DetailCard = ({ title, data, color = "blue", icon: Icon }: { title: string, data: any, color?: "blue" | "green" | "slate" | "indigo", icon?: any }) => {
    const colorStyles = {
      blue: { text: "text-blue-600", bg: "bg-blue-50", border: "border-blue-100", iconBg: "bg-blue-100", iconText: "text-blue-600", hover: "hover:border-blue-200 hover:bg-blue-100/50", valueText: "text-slate-900" },
      green: { text: "text-emerald-600", bg: "bg-emerald-50", border: "border-emerald-100", iconBg: "bg-emerald-100", iconText: "text-emerald-600", hover: "hover:border-emerald-200 hover:bg-emerald-100/50", valueText: "text-slate-900" },
      slate: { text: "text-slate-600", bg: "bg-slate-50", border: "border-slate-200", iconBg: "bg-slate-100", iconText: "text-slate-600", hover: "hover:border-slate-300 hover:bg-slate-100/50", valueText: "text-slate-900" },
      indigo: { text: "text-indigo-600", bg: "bg-indigo-50", border: "border-indigo-100", iconBg: "bg-indigo-100", iconText: "text-indigo-600", hover: "hover:border-indigo-200 hover:bg-indigo-100/50", valueText: "text-slate-900" },
    }[color];
    
    return (
      <div className={`rounded-2xl p-2.5 border ${colorStyles.border} ${colorStyles.bg} flex-1 transition-all duration-300 ${colorStyles.hover} group relative overflow-hidden shadow-sm`}>
        <div className="flex items-center gap-2 mb-2 relative z-10">
          {Icon && (
            <div className={`p-1.5 rounded-lg ${colorStyles.iconBg} ${colorStyles.iconText} group-hover:scale-110 transition-transform shadow-sm border border-white`}>
              <Icon size={14} />
            </div>
          )}
          <h4 className={`text-[10px] font-bold uppercase tracking-wider ${colorStyles.text}`}>{title}</h4>
        </div>
        <div className="space-y-1.5 relative z-10">
          <div className="flex justify-between items-center bg-white/50 p-1.5 rounded-lg border border-slate-100">
            <span className="text-[9px] font-bold text-slate-500 uppercase tracking-wider">Custo</span>
            <span className={`text-xs font-bold ${colorStyles.valueText}`}>R$ {formatNumber(data.custo, true)}</span>
          </div>
          <div className="flex justify-between items-center bg-white/50 p-1.5 rounded-lg border border-slate-100">
            <span className="text-[9px] font-bold text-slate-500 uppercase tracking-wider">Consumo</span>
            <span className={`text-xs font-bold ${colorStyles.valueText}`}>{formatNumber(data.consumo, false, 2)} <span className="text-[9px] text-slate-400 font-medium">kWh</span></span>
          </div>
          <div className="flex justify-between items-center pt-1.5 border-t border-slate-100">
            <span className={`text-[9px] font-bold uppercase tracking-wider ${colorStyles.text}`}>Tarifa Média</span>
            <span className={`text-xs font-bold ${colorStyles.text}`}>R$ {formatNumber(data.tarifa, true)}</span>
          </div>
        </div>
      </div>
    );
  };

  const SparklineCard = ({ title, data, color = "blue", sparklineData }: { title: string, data: any, color?: "blue" | "green", sparklineData: any[] }) => {
    const colorHex = color === "blue" ? "#3b82f6" : "#10b981";
    const bgClass = color === "blue" ? "bg-blue-50" : "bg-emerald-50";
    const textClass = color === "blue" ? "text-blue-600" : "text-emerald-600";
    const valueTextClass = "text-slate-900";
    const borderClass = color === "blue" ? "border-blue-100" : "border-emerald-100";
    const hoverBorderClass = color === "blue" ? "hover:border-blue-200 hover:bg-blue-100/50" : "hover:border-emerald-200 hover:bg-emerald-100/50";
    const iconBgClass = color === "blue" ? "bg-blue-100" : "bg-emerald-100";
    
    return (
      <div className={`rounded-2xl p-4 border ${borderClass} ${bgClass} flex items-center justify-between mt-3 transition-all duration-300 ${hoverBorderClass} group relative overflow-hidden shadow-sm`}>
        <div className="absolute top-0 right-0 w-24 h-24 bg-white/50 rounded-full -translate-y-1/2 translate-x-1/2 opacity-40 group-hover:scale-110 transition-transform duration-500 blur-xl"></div>
        <div className="flex items-center gap-4 relative z-10">
          <div className={`w-10 h-10 rounded-xl ${iconBgClass} flex items-center justify-center shadow-sm border border-white group-hover:scale-110 transition-transform`}>
            <TrendingUp size={20} className={textClass} />
          </div>
          <div>
            <h4 className={`text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1`}>{title}</h4>
            <p className={`text-lg font-bold ${valueTextClass}`}>R$ {formatNumber(data.custo, true)}</p>
          </div>
        </div>
        <div className="w-24 h-12 opacity-80 relative z-10 group-hover:opacity-100 transition-opacity">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={sparklineData}>
              <Line type="monotone" dataKey="value" stroke={colorHex} strokeWidth={2.5} dot={false} isAnimationActive={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-[#f8fafc] pb-8 text-slate-600 selection:bg-blue-500/30">
      {/* Header */}
      <header className="flex items-center justify-between px-8 py-4 bg-white border-b border-slate-200">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-blue-600 rounded-lg flex items-center justify-center shadow-lg">
            <Zap className="text-white" size={20} />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900 leading-tight">Sanesul Energy</h1>
            <p className="text-[9px] text-blue-600 uppercase tracking-widest font-bold">Portal de Inteligência Energética</p>
          </div>
        </div>
        <div className="flex items-center gap-4">
          <button
            onClick={() => setCurrentPage('sistema')}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white hover:bg-blue-500 transition-all rounded-xl text-xs font-bold tracking-wider shadow-md active:scale-95"
          >
            <LayoutDashboard size={16} />
            Acessar Sistema
          </button>
          <button
            onClick={handleLogout}
            className="flex items-center gap-2 px-4 py-2 bg-white border border-slate-200 text-slate-600 hover:bg-slate-50 transition-all rounded-xl text-xs font-bold tracking-wider shadow-sm active:scale-95"
            title="Sair"
          >
            <LogOut size={16} />
            Sair
          </button>
          <div className="w-10 h-10 rounded-full bg-slate-200 overflow-hidden border-2 border-white shadow-sm ml-2">
            <img src="https://i.pravatar.cc/150?img=11" alt="User" className="w-full h-full object-cover" />
          </div>
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
              <h2 className="text-2xl font-bold text-slate-900 tracking-tight">Análise de Insumo de Energia Elétrica</h2>
              <p className="text-sm text-slate-500 font-medium">Visão geral consolidada de custos e consumos por grupos tarifários.</p>
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
              {availableMonths.map(month => (
                <option key={month} value={month}>{month}</option>
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
                <h3 className="text-lg font-bold text-slate-900 uppercase tracking-widest">Evolução de Consumo e Custo</h3>
                <p className="text-xs text-slate-500 font-medium mt-1">Histórico Mensal Consolidado</p>
              </div>
            </div>
            <div className="flex items-center gap-6 bg-slate-50 px-5 py-3 rounded-2xl border border-slate-200 shadow-sm">
              <div className="flex items-center gap-3">
                <div className="w-3 h-3 rounded-full bg-[#0ea5e9]"></div>
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Consumo (kWh)</span>
              </div>
              <div className="flex items-center gap-3">
                <div className="w-3 h-3 rounded-full bg-[#6366f1]"></div>
                <span className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">Custo (R$)</span>
              </div>
            </div>
          </div>
          <div className="h-[165px] w-full relative z-10">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={monthlyData} margin={{ top: 20, right: 80, bottom: 10, left: 80 }}>
                <defs>
                  <linearGradient id="colorConsumo" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#0ea5e9" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#0ea5e9" stopOpacity={0}/>
                  </linearGradient>
                  <linearGradient id="colorCusto" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#6366f1" stopOpacity={0.3}/>
                    <stop offset="95%" stopColor="#6366f1" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} horizontal={true} stroke="#e2e8f0" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fill: '#64748b', fontSize: 11, fontWeight: 600 }} dy={10} padding={{ left: 10, right: 10 }} />
                <YAxis 
                  yAxisId="left" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748b', fontSize: 11, fontWeight: 600 }} 
                  dx={-10} 
                  domain={chartDomain}
                  tickFormatter={(val) => formatNumber(val, false, 0)}
                  label={{ value: 'Consumo (kWh)', angle: -90, position: 'insideLeft', offset: -55, style: { textAnchor: 'middle', fill: '#64748b', fontSize: 10, fontWeight: 700, textTransform: 'uppercase' } }}
                />
                <YAxis 
                  yAxisId="right" 
                  orientation="right" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#64748b', fontSize: 11, fontWeight: 600 }} 
                  dx={10} 
                  domain={chartDomain}
                  tickFormatter={(val) => formatNumber(val, false, 0)}
                  label={{ value: 'Custo (R$)', angle: 90, position: 'insideRight', offset: -55, style: { textAnchor: 'middle', fill: '#64748b', fontSize: 10, fontWeight: 700, textTransform: 'uppercase' } }}
                />
                <Tooltip 
                  contentStyle={{ borderRadius: '20px', border: '1px solid #e2e8f0', boxShadow: '0 10px 15px -3px rgba(0,0,0,0.1)', padding: '16px 20px', fontWeight: 'bold', backgroundColor: 'rgba(255, 255, 255, 0.95)', backdropFilter: 'blur(12px)' }}
                  formatter={(value: number, name: string) => [
                    name === 'consumo' ? formatNumber(value, false, 0) + ' kWh' : formatNumber(value, true),
                    name === 'consumo' ? 'Consumo' : 'Custo'
                  ]}
                  labelStyle={{ color: '#64748b', marginBottom: '8px', fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.1em' }}
                  itemStyle={{ color: '#1e293b' }}
                />
                <Area 
                  yAxisId="left" 
                  type="monotone" 
                  dataKey="consumo" 
                  stroke="#0ea5e9" 
                  strokeWidth={4} 
                  fillOpacity={1} 
                  fill="url(#colorConsumo)" 
                  dot={{ r: 4, fill: '#fff', strokeWidth: 2, stroke: '#0ea5e9' }} 
                  activeDot={{ r: 6, strokeWidth: 0, fill: '#0ea5e9' }} 
                />
                <Area 
                  yAxisId="right" 
                  type="monotone" 
                  dataKey="custo" 
                  stroke="#6366f1" 
                  strokeWidth={4} 
                  fillOpacity={1} 
                  fill="url(#colorCusto)" 
                  dot={{ r: 4, fill: '#fff', strokeWidth: 2, stroke: '#6366f1' }} 
                  activeDot={{ r: 6, strokeWidth: 0, fill: '#6366f1' }} 
                />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Summary Cards */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
          <div className="md:col-span-2">
            <SummaryCard title="TOTAL GERAL" data={totalGeral} icon={DollarSign} color="indigo" />
          </div>
          <SummaryCard title="GRUPO A (MT/AT)" data={grupoA} icon={Zap} color="blue" />
          <SummaryCard title="GRUPO B (BT)" data={grupoB} icon={Battery} color="sky" />
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
              <span className="text-xs font-bold text-blue-600 bg-blue-50 px-3 py-1.5 rounded-lg uppercase tracking-wider border border-blue-100">Alta Tensão</span>
            </div>
            
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* LIVRE Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-blue-500 to-blue-700 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">MERCADO LIVRE</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Ambiente de Contratação Livre</p>
                    </div>
                  </div>
                  <div className="p-3 bg-blue-50 rounded-2xl border border-blue-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Activity size={24} className="text-blue-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={livre.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={livre.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={livre.tarifa} isCurrency />
                </div>

                <div className="space-y-6 relative z-10">
                  <div className="grid grid-cols-2 gap-6">
                    <DetailCard title="Faturas Azul" data={livreAzul} color="blue" icon={Zap} />
                    <DetailCard title="Faturas Verde" data={livreVerde} color="green" icon={Zap} />
                  </div>
                  <div className="grid grid-cols-2 gap-6">
                    <SparklineCard title="Evolução Azul" data={livreAzul} color="blue" sparklineData={sparklineDataAzul} />
                    <SparklineCard title="Evolução Verde" data={livreVerde} color="green" sparklineData={sparklineDataVerde} />
                  </div>
                </div>
              </div>

              {/* Cativo Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-slate-400 to-slate-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">Consumidor Cativo</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Ambiente de Contratação Regulada</p>
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <ZapOff size={24} className="text-slate-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={cativo.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={cativo.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={cativo.tarifa} isCurrency />
                </div>

                <div className="grid grid-cols-2 gap-6 relative z-10">
                  <DetailCard title="Faturas Azul" data={cativoAzul} color="blue" icon={Zap} />
                  <DetailCard title="Faturas Verde" data={cativoVerde} color="green" icon={Zap} />
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
              <span className="text-xs font-bold text-sky-600 bg-sky-50 px-3 py-1.5 rounded-lg uppercase tracking-wider border border-sky-100">Baixa Tensão</span>
            </div>
            
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-8">
              {/* Sem Compensação Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-slate-400 to-slate-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">UC's Sem Compensação</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Consumo Padrão da Rede</p>
                    </div>
                  </div>
                  <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Activity size={24} className="text-slate-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={semCompensacao.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={semCompensacao.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={semCompensacao.tarifa} isCurrency />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10">
                  <DetailCard title="Consumo Geral" data={geral} color="slate" icon={Activity} />
                  <DetailCard title="Consumos Mínimos" data={consumosMinimos} color="indigo" icon={ArrowDown} />
                  <DetailCard title="Optante B" data={optanteB} color="blue" icon={Battery} />
                </div>
              </div>

              {/* Com Compensação Card */}
              <div className="bg-white rounded-[2rem] p-6 shadow-sm border border-slate-100 relative overflow-hidden group hover:shadow-xl transition-all duration-500">
                <div className="flex justify-between items-center mb-8 relative z-10">
                  <div className="flex items-center gap-5">
                    <div className="w-3 h-12 bg-gradient-to-b from-emerald-400 to-emerald-600 rounded-full shadow-md"></div>
                    <div>
                      <h3 className="text-lg font-bold text-slate-800 uppercase tracking-widest">UC's Com Compensação</h3>
                      <p className="text-xs text-slate-500 font-medium mt-1">Geração Distribuída e Sustentável</p>
                    </div>
                  </div>
                  <div className="p-3 bg-emerald-50 rounded-2xl border border-emerald-100/50 shadow-sm group-hover:scale-110 transition-transform duration-300">
                    <Zap size={24} className="text-emerald-600" />
                  </div>
                </div>
                
                <div className="space-y-1 relative z-10 mb-8">
                  <MetricRow icon={DollarSign} label="Custo Total" value={comCompensacao.custo} isCurrency />
                  <MetricRow icon={Zap} label="Consumo Total" value={comCompensacao.consumo} unit="kWh" />
                  <MetricRow icon={Calculator} label="Tarifa Média" value={comCompensacao.tarifa} isCurrency />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-6 relative z-10">
                  <DetailCard title="PPP Fotovoltaica" data={ppp} color="green" icon={Zap} />
                  <DetailCard title="Usinas Sanesul" data={usinas} color="blue" icon={Activity} />
                  
                  {/* Crédito de Carbono Card */}
                  <div className="bg-emerald-50/50 rounded-2xl p-4 border border-emerald-100 relative overflow-hidden group flex flex-col justify-between shadow-sm">
                    <div className="flex items-center gap-3 mb-3">
                      <div className="p-1.5 bg-emerald-100 rounded-lg text-emerald-600 border border-white shadow-sm">
                        <Leaf size={16} />
                      </div>
                      <h4 className="text-[10px] font-bold text-emerald-800 uppercase tracking-wider">Crédito de Carbono</h4>
                    </div>
                    <div className="space-y-2">
                      <div className="flex justify-between items-center bg-white/60 p-1.5 rounded-lg border border-emerald-100/50">
                        <span className="text-[9px] font-bold text-emerald-600/70 uppercase tracking-wider">Mês/Ano</span>
                        <span className="text-xs font-bold text-emerald-900">{selectedMonth === 'all' ? 'Todos os Meses' : selectedMonth}</span>
                      </div>
                      <div className="flex justify-between items-center bg-white/60 p-1.5 rounded-lg border border-emerald-100/50">
                        <span className="text-[9px] font-bold text-emerald-600/70 uppercase tracking-wider">Geração Solar</span>
                        <span className="text-xs font-bold text-emerald-900">{formatNumber(totalSolarInjetada, false, 2)} <span className="text-[9px] text-emerald-600/70">kWh</span></span>
                      </div>
                      <div className="flex justify-between items-center bg-white/60 p-1.5 rounded-lg border border-emerald-100/50">
                        <span className="text-[9px] font-bold text-emerald-600/70 uppercase tracking-wider">Emissões Evitadas</span>
                        <span className="text-xs font-bold text-emerald-900">{formatNumber(emissoesEvitadas, false, 2)} <span className="text-[9px] text-emerald-600/70">KgCO₂</span></span>
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

export default VisaoGeralDashboard;
