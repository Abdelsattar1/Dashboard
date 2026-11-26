export function formatEGP(value: number | null | undefined, decimals = 2): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return '—';
  return (
    value.toLocaleString('en-US', {
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }) + ' EGP'
  );
}

export function formatNum(value: number | null | undefined, decimals = 0): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return '—';
  return value.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

export function formatPct(value: number | null | undefined): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return 'غير متاح';
  return `${value.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}%`;
}

export function formatPctOneDecimal(value: number | null | undefined): string {
  if (value === null || value === undefined || !Number.isFinite(value)) return 'غير متاح';
  return `${value.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
}

export function formatMillionOrFull(value: number | null | undefined): { numText: string; unit: string } {
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return { numText: '—', unit: '' };
  }
  if (Math.abs(value) >= 1_000_000) {
    const millions = value / 1_000_000;
    return {
      numText: millions.toLocaleString('en-US', { minimumFractionDigits: 1, maximumFractionDigits: 1 }),
      unit: 'مليون',
    };
  }
  return {
    numText: value.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 }),
    unit: 'EGP',
  };
}

export function formatDateTime(iso: string | null | undefined): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
}
