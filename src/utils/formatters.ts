export const formatMonth = (month: string | number) => {
  const months = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'
  ];
  if (typeof month === 'number') return months[month - 1] || '';
  const m = month.toString().trim();
  if (m.length <= 2) return months[parseInt(m) - 1] || '';
  return m.charAt(0).toUpperCase() + m.slice(1).toLowerCase();
};

export const getMonthNumber = (month: string) => {
  const months: Record<string, number> = {
    'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4,
    'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
    'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12,
    'Jan': 1, 'Fev': 2, 'Mar': 3, 'Abr': 4, 'Mai': 5, 'Jun': 6,
    'Jul': 7, 'Ago': 8, 'Set': 9, 'Out': 10, 'Nov': 11, 'Dez': 12
  };
  const m = month.trim();
  if (months[m]) return months[m];
  const capitalized = m.charAt(0).toUpperCase() + m.slice(1).toLowerCase();
  return months[capitalized] || 0;
};

export const parseValue = (val: string | number) => {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const clean = val.toString().replace('R$', '').replace('kWh', '').replace('.', '').replace(',', '.').trim();
  const parsed = parseFloat(clean);
  return isNaN(parsed) ? 0 : parsed;
};

export const formatReference = (ref: string) => {
  if (ref && ref.includes('/')) {
    const parts = ref.split('/');
    if (parts.length === 2) {
      return `${formatMonth(parts[0])}/${parts[1]}`;
    }
  }
  return formatMonth(ref);
};

export const formatNumber = (val: number, isCurrency: boolean = false, precision: number = 2) => {
  return val.toLocaleString('pt-BR', {
    minimumFractionDigits: isCurrency ? precision : 0,
    maximumFractionDigits: precision
  });
};
