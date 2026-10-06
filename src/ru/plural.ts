// plural(5, 'слайд', 'слайда', 'слайдов') -> 'слайдов'
export function plural(n: number, one: string, few: string, many: string): string {
  const m10 = n % 10;
  const m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
  return many;
}

export const count = (n: number, one: string, few: string, many: string): string =>
  `${n.toLocaleString('ru-RU')} ${plural(n, one, few, many)}`;
