import { count } from './plural';
import type { XlsxReaderText } from './types';

const num = (n: number) => n.toLocaleString('ru-RU');

export const xlsx: XlsxReaderText = {
  displayName: 'Просмотр таблиц',
  commands: {
    focusSearch: 'Поиск по текущей книге',
    goToCell: 'Перейти к ячейке или именованному диапазону',
    copyValues: 'Скопировать выделенные значения как TSV',
    copyFormulas: 'Скопировать выделенные формулы',
    copyMarkdown: 'Скопировать выделенный диапазон как Markdown',
    createSummaryNote: 'Создать заметку-конспект таблицы',
  },
  toolbar: {
    reload: 'Перезагрузить',
    searchPlaceholder: 'Поиск по книге',
    searchWorkbook: 'Поиск по всем листам',
    previousResult: 'Предыдущий результат',
    nextResult: 'Следующий результат',
    zoomPercentage: 'Масштаб, %',
    fitWidth: 'Вписать лист по ширине',
    copyValues: 'Скопировать значения как TSV',
    copyFormulas: 'Скопировать формулы',
    copyMarkdown: 'Скопировать как таблицу Markdown',
    createSummaryNote: 'Создать заметку-конспект книги',
    openExternally: 'Открыть во внешней программе',
  },
  labels: {
    formula: 'Формула',
    cachedValue: 'Сохранённое значение',
    openLink: 'Открыть ссылку',
    sheetTabs: 'Листы книги',
    row: (row) => `Строка ${row}`,
    column: (column) => `Столбец ${column}`,
    selectedRange: (range) => `Выделено: ${range}`,
    formulaSafety: 'Формулы не пересчитываются: показываются только сохранённый текст и последние вычисленные значения.',
    nameBox: 'Ячейка, диапазон или имя',
    hiddenSheets: (n) => count(n, 'скрытый лист', 'скрытых листа', 'скрытых листов'),
    hiddenSheet: 'Скрытый лист',
    veryHiddenSheet: 'Очень скрытый лист',
    comment: (author, text) => `${author}: ${text}`,
    unknownCommentAuthor: 'Неизвестный автор',
    worksheetImage: 'Изображение на листе',
    chartKind: (kind) => {
      switch (kind) {
        case 'area':
          return 'Диаграмма с областями';
        case 'bar':
          return 'Столбчатая диаграмма';
        case 'line':
          return 'График';
        case 'pie':
          return 'Круговая диаграмма';
        default:
          return 'Диаграмма';
      }
    },
    unsupportedChart: 'Этот тип диаграммы пока не показывается в режиме просмотра.',
  },
  status: {
    reading: (fileName) => `Чтение ${fileName}…`,
    parsingSheet: (sheetName) => `Загрузка листа ${sheetName}…`,
    parsingSheetProgress: (sheetName, percent) => `Загрузка листа ${sheetName}… ${percent}%`,
    ready: (fileName, sheetName, rows, columns) =>
      `${fileName} · ${sheetName} · ${count(rows, 'строка', 'строки', 'строк')} × ${count(columns, 'столбец', 'столбца', 'столбцов')}`,
    searching: (sheetName, current, total) => `Поиск на листе ${sheetName} (${current}/${total})…`,
    searchResults: (n) => count(n, 'результат', 'результата', 'результатов'),
    creatingSummary: (sheetName, current, total) => `Конспект листа ${sheetName} (${current}/${total})…`,
    legacy: (fileName) => `${fileName} нужно сконвертировать в .xlsx`,
  },
  notices: {
    copiedValues: 'Значения ячеек скопированы как TSV',
    copiedFormulas: 'Формулы и значения скопированы',
    copiedMarkdown: 'Выделение скопировано как таблица Markdown',
    copyFailed: (message) => `Не удалось скопировать ячейки: ${message}`,
    selectionTooLarge: (n, limit) =>
      `В выделении ${count(n, 'ячейка', 'ячейки', 'ячеек')}. За раз можно скопировать не больше ${num(limit)}.`,
    invalidReference: (value) => `Не найдена ячейка, диапазон или имя: ${value}`,
    openedExistingSummaryNote: 'Открыта существующая заметка-конспект таблицы',
    createdSummaryNote: 'Заметка-конспект таблицы создана',
    summaryFailed: (message) => `Не удалось создать конспект таблицы: ${message}`,
    externalDesktopOnly: 'Открытие во внешней программе доступно только в Obsidian для компьютера',
    externalLocalVaultOnly: 'Для открытия во внешней программе нужно локальное хранилище на компьютере',
    externalFailed: (message) => `Не удалось открыть книгу: ${message}`,
    linkConfirmation: (target) => `Открыть ссылку из книги?\n\n${target}`,
    blockedLink: 'Открывать можно только безопасные веб-ссылки, адреса почты и ссылки внутри книги',
  },
  summaryNote: {
    sourceLabel: 'Источник',
    currentSheetLabel: 'Текущий лист',
    selectedRangeLabel: 'Выделенный диапазон',
    workbookOverviewHeading: 'Обзор книги',
    namedRangesHeading: 'Именованные диапазоны',
    selectedRangeHeading: 'Снимок выделенного диапазона',
    worksheetPreviewsHeading: 'Превью листов',
    keyFindingsHeading: 'Главное',
    followUpsHeading: 'Что сделать дальше',
    cellHeader: 'Ячейка',
    displayedValueHeader: 'Значение',
    noNamedRanges: 'Поддерживаемых именованных диапазонов нет.',
    noPreview: 'Для превью нет заполненных ячеек.',
    visibility: {
      visible: 'видимый',
      hidden: 'скрытый',
      veryHidden: 'очень скрытый',
    },
    sheetSummary: (name, visibility, rows, columns, populatedCells, comments, images, charts, conditionalRules) =>
      `${name} (${visibility}) — ${count(rows, 'строка', 'строки', 'строк')} × ${count(columns, 'столбец', 'столбца', 'столбцов')}; ` +
      `заполнено: ${count(populatedCells, 'ячейка', 'ячейки', 'ячеек')}; ` +
      `комментарии: ${num(comments)}, изображения: ${num(images)}, диаграммы: ${num(charts)}, правила условного форматирования: ${num(conditionalRules)}`,
  },
  legacy: {
    title: 'Старый формат Excel',
    body: 'Двоичные файлы .xls нельзя безопасно отрисовать в Obsidian.',
    convert: 'Откройте файл во внешней программе и сохраните копию в .xlsx, чтобы смотреть её здесь.',
  },
  errors: {
    formatMismatchTitle: 'Файл не является корректной книгой XLSX',
    formatMismatchBody: 'Это не поддерживаемая книга Excel в ZIP-формате.',
    encryptedTitle: 'Книга зашифрована',
    encryptedBody: 'Зашифрованные или защищённые паролем книги просмотреть нельзя.',
    limitTitle: 'Книга превышает безопасные лимиты',
    limitBody: 'Книга выходит за безопасные ограничения по размеру архива или памяти.',
    damagedTitle: 'Книга повреждена',
    damagedBody: 'Не удалось безопасно прочитать архив книги или XML листа.',
    unsupportedTitle: 'Неподдерживаемое содержимое книги',
    unsupportedBody: 'В книге есть структуры, которые режим просмотра не поддерживает.',
    unknownTitle: 'Не удалось открыть книгу',
    unknownBody: 'Просмотр таблицы прервался из-за непредвиденной ошибки.',
    tryAgain: 'Проверьте исходный файл и перезагрузите книгу.',
    openExternally: 'Откройте исходный файл в надёжной программе для таблиц.',
    details: 'Диагностика',
  },
};
