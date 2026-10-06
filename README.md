# Office Reader: русский язык

Плагин для Obsidian, который переводит на русский язык интерфейс плагина [Office Reader](https://github.com/qianwei4712/obsidian-word-reader) (`word-reader`) — просмотрщика `.docx`, `.pptx` и `.xlsx`.

Переведено всё: настройки, команды, панели инструментов, уведомления, ошибки, просмотр изображений, навигация по слайдам и листам, заметки-конспекты. Числа склоняются по-русски («3 слайда», «1 250 строк»).

## Как работает

У Office Reader есть только английский и китайский языки. Плагин при запуске Obsidian:

1. сохраняет оригинальный `.obsidian/plugins/word-reader/main.js` в `main.js.office-reader-ru-backup`;
2. встраивает русский перевод на место английского: в настройках Office Reader выбор языка будет «Русский» / «中文»;
3. при первой установке переключает язык Office Reader с китайского на русский;
4. перезапускает Office Reader и заново открывает открытые файлы.

После обновления Office Reader перевод применяется снова автоматически. Если новая версия изменила формат и патч не встаёт, файл не трогается, а плагин показывает уведомление.

## Установка

Через [BRAT](https://github.com/TfTHacker/obsidian42-brat): «Add Beta plugin» → `DaniilSemizhonov/obsidian-office-reader-ru`, затем включите плагин в разделе «Сторонние плагины».

Вручную: скачайте `main.js` и `manifest.json` из [последнего релиза](https://github.com/DaniilSemizhonov/obsidian-office-reader-ru/releases/latest) в `<хранилище>/.obsidian/plugins/office-reader-ru/`.

## Откат

Команда «Вернуть оригинальный Office Reader и отключить перевод» или кнопка в настройках плагина.

## Сборка

```bash
npm install
npm run build
```

Перевод лежит в `src/ru/`. Типы в `src/ru/types.ts` скопированы из исходников Office Reader, поэтому `npm run check` ловит пропущенные строки.
