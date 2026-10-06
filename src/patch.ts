import RU_JS from '../build/ru.txt';
import { LITERALS } from './ru/literals';

// Меняется при любом изменении перевода или логики патча — так уже пропатченный
// Office Reader пропатчится заново после обновления этого плагина.
const PATCH_VERSION = 1;
const PREFIX = '/*office-reader-ru:';

function hash(s: string): string {
  let h = 5381;
  for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

export const SIGNATURE = `${PREFIX}${PATCH_VERSION}.${hash(RU_JS + JSON.stringify(LITERALS))}*/`;

export const isPatched = (src: string) => src.startsWith(PREFIX);
export const isCurrentPatch = (src: string) => src.startsWith(SIGNATURE);

class PatchError extends Error {
  constructor(what: string) {
    super(`не найдено в main.js Office Reader: ${what} — возможно, плагин обновился и изменил формат`);
  }
}

// Русский перевод занимает место английского: Office Reader поддерживает только
// "en" и "zh-CN", и подмена "en" требует минимум правок в чужом коде.
export function patchOfficeReader(original: string): string {
  let out = original;

  // 1. Общие тексты и XLSX. В сборке это
  //    function Pt(r){var t;return(t=yh[r])!=null?t:yh["zh-CN"]}
  //    (или с нативным ??). Какой из двух — определяем по содержимому словаря.
  const getters = [...out.matchAll(/function ([\w$]+)\(([\w$]+)\)\{([^{}]{0,80}?([\w$]+)\["zh-CN"\])\}/g)];
  const patched = new Set<string>();
  for (const m of getters.reverse()) {
    const [found, name, arg, body, dict] = m;
    const head = out.slice(out.indexOf(`${dict}={en:{`), out.indexOf(`${dict}={en:{`) + 40);
    const key = head.includes('{en:{settings:') ? 'word' : head.includes('{en:{displayName:') ? 'xlsx' : null;
    if (!key) continue;
    const fn = `function ${name}(${arg}){if(${arg}==="en")return __officeRu.${key};${body}}`;
    out = out.slice(0, m.index) + fn + out.slice(m.index! + found.length);
    patched.add(key);
  }
  if (!patched.has('word')) throw new PatchError('выбор языка для общих текстов');
  if (!patched.has('xlsx')) throw new PatchError('выбор языка для таблиц');

  // 2. PPTX: function bi(r){return r==="en"?Cb:kb}, где Cb — английский словарь.
  const pptx = [...out.matchAll(/function ([\w$]+)\(([\w$]+)\)\{return \2==="en"\?([\w$]+):([\w$]+)\}/g)].find((m) =>
    out.includes(`${m[3]}={displayName:"PowerPoint Reader"`)
  );
  if (!pptx) throw new PatchError('выбор языка для презентаций');
  out = out.replace(pptx[0], `function ${pptx[1]}(${pptx[2]}){return ${pptx[2]}==="en"?__officeRu.pptx:${pptx[4]}}`);

  // 3. Язык по умолчанию — русский (слот "en") вместо китайского.
  //    DEFAULT_SETTINGS: {language:"zh-CN",...}
  //    normalizeLanguage: return r==="en"||r==="zh-CN"?r:"zh-CN"
  const defaults: [RegExp, string, string][] = [
    [/language:"zh-CN"/, 'language:"en"', 'настройки по умолчанию'],
    [/(==="zh-CN"\?[\w$]+:)"zh-CN"/, '$1"en"', 'нормализация языка'],
  ];
  for (const [re, to, what] of defaults) {
    if (!re.test(out)) throw new PatchError(`язык по умолчанию (${what})`);
    out = out.replace(re, to);
  }

  // 4. В списке языков "English" превращается в "Русский" — и в китайском интерфейсе тоже.
  if (!out.includes('languageEn:"English"')) throw new PatchError('название языка English');
  out = out.split('languageEn:"English"').join('languageEn:"Русский"');

  // 5. Строки, вшитые в код мимо i18n.
  for (const [en, ru] of Object.entries(LITERALS)) {
    const q = JSON.stringify(en);
    if (!out.includes(q)) throw new PatchError(`строка ${q}`);
    out = out.split(q).join(JSON.stringify(ru));
  }

  // esbuild ставит "use strict" в начало IIFE; в начале чужого main.js он включил бы
  // строгий режим для всего Office Reader.
  const ru = RU_JS.replace(/^"use strict";/, '');
  return `${SIGNATURE}\n${ru}\n${out}`;
}
