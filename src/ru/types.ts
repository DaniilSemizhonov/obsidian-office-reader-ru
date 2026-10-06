// Интерфейсы скопированы из qianwei4712/obsidian-word-reader 3.5.0
// (src/i18n.ts, src/pptx/pptxI18n.ts, src/xlsx/xlsxI18n.ts), чтобы tsc проверял полноту перевода.

export interface WordReaderText {
  settings: {
    languageName: string;
    languageDesc: string;
    languageZh: string;
    languageEn: string;
    defaultZoomName: string;
    defaultZoomDesc: string;
    defaultFitWidthName: string;
    defaultFitWidthDesc: string;
    xlsxFitWidthName: string;
    xlsxFitWidthDesc: string;
    showOutlineName: string;
    showOutlineDesc: string;
    imagePreviewName: string;
    imagePreviewDesc: string;
    largeFileWarningName: string;
    largeFileWarningDesc: string;
    externalOpeningName: string;
    externalOpeningDesc: string;
  };
  commands: {
    reload: string;
    copyText: string;
    copyMarkdown: string;
    createNote: string;
    openExternal: string;
  };
  toolbar: {
    reload: string;
    zoomPercentage: string;
    fitWidth: string;
    showOutline: string;
    hideOutline: string;
    searchPlaceholder: string;
    searchDocument: string;
    previousSearchResult: string;
    nextSearchResult: string;
    copyText: string;
    copyMarkdown: string;
    createSummaryNote: string;
    openExternally: string;
  };
  notices: {
    legacyDoc: string;
    copiedSelectedText: string;
    copiedPlainText: string;
    copiedSelectedMarkdown: string;
    copiedMarkdown: string;
    copiedImage: string;
    couldNotCopyImage: (message: string) => string;
    startedImageDownload: string;
    savedImage: string;
    couldNotSaveImage: (message: string) => string;
    copiedDiagnostics: string;
    couldNotCopyDiagnostics: (message: string) => string;
    openedExistingSummaryNote: string;
    createdSummaryNote: string;
    externalDesktopOnly: string;
    externalLocalVaultOnly: string;
    largeDocument: (sizeMb: string) => string;
  };
  status: {
    extractingPlainText: string;
    copiedPlainTextFrom: (fileName: string) => string;
    extractingMarkdown: string;
    copiedMarkdownFrom: (fileName: string) => string;
    preview: (fileName: string) => string;
    reading: (fileName: string) => string;
    rendering: (fileName: string) => string;
    preparingPreview: (fileName: string, percent: number) => string;
    buildingNavigation: (fileName: string) => string;
    largeDocument: (sizeMb: string) => string;
    legacyDoc: (fileName: string) => string;
  };
  outline: {
    title: string;
    empty: string;
    expandSection: string;
    collapseSection: string;
  };
  legacyDoc: {
    title: string;
    body: string;
    openExternally: string;
    convertToDocx: string;
  };
  errors: {
    encryptedTitle: string;
    encryptedBody: string;
    encryptedStatus: string;
    encryptedTips: string[];
    formatMismatchTitle: string;
    formatMismatchBody: string;
    formatMismatchStatus: string;
    formatMismatchTips: string[];
    zipContainerTitle: string;
    zipContainerBody: string;
    zipContainerStatus: string;
    zipContainerTips: string[];
    xmlStructureTitle: string;
    xmlStructureBody: string;
    xmlStructureStatus: string;
    xmlStructureTips: string[];
    unsupportedStructureTitle: string;
    unsupportedStructureBody: string;
    unsupportedStructureStatus: string;
    unsupportedStructureTips: string[];
    genericTitle: string;
    genericBody: string;
    genericStatus: string;
    genericTips: string[];
    diagnosticDetails: string;
    diagnosticCategory: string;
    diagnosticFileName: string;
    diagnosticFileSize: string;
    diagnosticModified: string;
    diagnosticErrorSummary: string;
    copyDiagnostics: string;
  };
  imagePreview: {
    fitToWindow: string;
    actualSize: string;
    copyImage: string;
    saveImageAs: string;
    loading: string;
    loadFailed: string;
    unknownSize: string;
    saveDialogTitle: string;
    imageFilterName: (extension: string) => string;
    allFiles: string;
  };
  summaryNote: {
    sourceLabel: string;
    summaryHeading: string;
    keyFindingsHeading: string;
    followUpsHeading: string;
    quotesHeading: string;
  };
}

export interface PptxReaderText {
  displayName: string;
  commands: {
    previousSlide: string;
    nextSlide: string;
    toggleFullscreen: string;
    copySlideText: string;
    copyRenderDiagnostics: string;
    createSummaryNote: string;
    toggleNotes: string;
    focusSearch: string;
  };
  toolbar: {
    reload: string;
    previousSlide: string;
    nextSlide: string;
    pageNumber: string;
    zoomPercentage: string;
    fitWindow: string;
    enterFullscreen: string;
    exitFullscreen: string;
    showNavigation: string;
    hideNavigation: string;
    searchPlaceholder: string;
    searchPresentation: string;
    copyText: string;
    copyRenderDiagnostics: string;
    createSummaryNote: string;
    showNotes: string;
    hideNotes: string;
    openExternally: string;
  };
  status: {
    reading: (fileName: string) => string;
    indexing: (fileName: string) => string;
    rendering: (page: number, total: number) => string;
    ready: (fileName: string, page: number, total: number) => string;
  };
  navigation: {
    title: string;
    slideLabel: (page: number) => string;
    slideCount: (slides: number) => string;
    searchCount: (matches: number, slides: number) => string;
    notesMatch: string;
    noMatches: string;
  };
  notes: {
    title: string;
    empty: string;
  };
  errors: {
    formatMismatchTitle: string;
    formatMismatchBody: string;
    encryptedTitle: string;
    encryptedBody: string;
    damagedTitle: string;
    damagedBody: string;
    limitTitle: string;
    limitBody: string;
    unsupportedTitle: string;
    unsupportedBody: string;
    unknownTitle: string;
    unknownBody: string;
    tryAgain: string;
    openExternally: string;
    details: string;
  };
  notices: {
    externalDesktopOnly: string;
    externalLocalVaultOnly: string;
    externalFailed: (message: string) => string;
    fullscreenFailed: (message: string) => string;
    copiedSelectedText: string;
    copiedSlideText: string;
    copiedRenderDiagnostics: string;
    noRenderDiagnostics: string;
    copyFailed: (message: string) => string;
    openedExistingSummaryNote: string;
    createdSummaryNote: string;
  };
  summaryNote: {
    sourceLabel: string;
    currentSlideLabel: string;
    summaryHeading: string;
    keySlidesHeading: string;
    followUpsHeading: string;
    slideReferencesHeading: string;
    slideReference: (page: number, title: string) => string;
  };
}

export interface XlsxReaderText {
  displayName: string;
  commands: {
    focusSearch: string;
    goToCell: string;
    copyValues: string;
    copyFormulas: string;
    copyMarkdown: string;
    createSummaryNote: string;
  };
  toolbar: {
    reload: string;
    searchPlaceholder: string;
    searchWorkbook: string;
    previousResult: string;
    nextResult: string;
    zoomPercentage: string;
    fitWidth: string;
    copyValues: string;
    copyFormulas: string;
    copyMarkdown: string;
    createSummaryNote: string;
    openExternally: string;
  };
  labels: {
    formula: string;
    cachedValue: string;
    openLink: string;
    sheetTabs: string;
    row: (row: number) => string;
    column: (column: string) => string;
    selectedRange: (range: string) => string;
    formulaSafety: string;
    nameBox: string;
    hiddenSheets: (count: number) => string;
    hiddenSheet: string;
    veryHiddenSheet: string;
    comment: (author: string, text: string) => string;
    unknownCommentAuthor: string;
    worksheetImage: string;
    chartKind: (kind: string) => string;
    unsupportedChart: string;
  };
  status: {
    reading: (fileName: string) => string;
    parsingSheet: (sheetName: string) => string;
    parsingSheetProgress: (sheetName: string, percent: number) => string;
    ready: (
      fileName: string,
      sheetName: string,
      rows: number,
      columns: number,
    ) => string;
    searching: (
      sheetName: string,
      current: number,
      total: number,
    ) => string;
    searchResults: (count: number) => string;
    creatingSummary: (
      sheetName: string,
      current: number,
      total: number,
    ) => string;
    legacy: (fileName: string) => string;
  };
  notices: {
    copiedValues: string;
    copiedFormulas: string;
    copiedMarkdown: string;
    copyFailed: (message: string) => string;
    selectionTooLarge: (count: number, limit: number) => string;
    invalidReference: (value: string) => string;
    openedExistingSummaryNote: string;
    createdSummaryNote: string;
    summaryFailed: (message: string) => string;
    externalDesktopOnly: string;
    externalLocalVaultOnly: string;
    externalFailed: (message: string) => string;
    linkConfirmation: (target: string) => string;
    blockedLink: string;
  };
  summaryNote: {
    sourceLabel: string;
    currentSheetLabel: string;
    selectedRangeLabel: string;
    workbookOverviewHeading: string;
    namedRangesHeading: string;
    selectedRangeHeading: string;
    worksheetPreviewsHeading: string;
    keyFindingsHeading: string;
    followUpsHeading: string;
    cellHeader: string;
    displayedValueHeader: string;
    noNamedRanges: string;
    noPreview: string;
    visibility: {
      visible: string;
      hidden: string;
      veryHidden: string;
    };
    sheetSummary: (
      name: string,
      visibility: string,
      rows: number,
      columns: number,
      populatedCells: number,
      comments: number,
      images: number,
      charts: number,
      conditionalRules: number,
    ) => string;
  };
  legacy: {
    title: string;
    body: string;
    convert: string;
  };
  errors: {
    formatMismatchTitle: string;
    formatMismatchBody: string;
    encryptedTitle: string;
    encryptedBody: string;
    limitTitle: string;
    limitBody: string;
    damagedTitle: string;
    damagedBody: string;
    unsupportedTitle: string;
    unsupportedBody: string;
    unknownTitle: string;
    unknownBody: string;
    tryAgain: string;
    openExternally: string;
    details: string;
  };
}
