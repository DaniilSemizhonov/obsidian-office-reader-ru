// Строки из src/reader/recovery.ts — они вшиты в код в обход i18n.
export const LITERALS: Record<string, string> = {
  'Office file format changed': 'Формат файла Office изменился',
  'Office file is no longer available': 'Файл Office больше недоступен',
  'Office file changed': 'Файл Office изменился',
  'The restored view no longer matches this file format. Its stale preview, search results, and caches were discarded.':
    'Восстановленный вид больше не соответствует формату файла. Устаревшее превью, результаты поиска и кэш сброшены.',
  'The restored source file could not be found. The reader released its pending work and cached resources.':
    'Исходный файл не найден. Незавершённые задачи и кэш освобождены.',
  'The source changed while it was being read. The stale session was invalidated before it could update the preview.':
    'Файл изменился во время чтения. Устаревший сеанс сброшен до обновления превью.',
  'Reader session invalidated safely': 'Сеанс просмотра безопасно сброшен',
  'Reload current file': 'Перезагрузить текущий файл',
};
