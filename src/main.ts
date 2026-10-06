import { App, Notice, Plugin, PluginSettingTab, Setting, WorkspaceLeaf, ViewState } from 'obsidian';
import { isCurrentPatch, isPatched, patchOfficeReader } from './patch';

const TARGET_ID = 'word-reader';
const TARGET_VIEWS = ['word-reader-view', 'pptx-reader-view', 'xlsx-reader-view'];
const BACKUP_NAME = 'main.js.office-reader-ru-backup';

// app.plugins — недокументированный, но стабильный API Obsidian.
interface PluginsApi {
  manifests: Record<string, { dir?: string; version: string }>;
  plugins: Record<string, unknown>;
  enabledPlugins: Set<string>;
  enablePlugin(id: string): Promise<boolean>;
  disablePlugin(id: string): Promise<void>;
  disablePluginAndSave(id: string): Promise<void>;
}

type Status = 'not-installed' | 'original' | 'patched' | 'outdated';

export default class OfficeReaderRuPlugin extends Plugin {
  private get plugins(): PluginsApi {
    return (this.app as unknown as { plugins: PluginsApi }).plugins;
  }

  private paths(): { main: string; backup: string; data: string } | null {
    const dir = this.plugins.manifests[TARGET_ID]?.dir;
    return dir ? { main: `${dir}/main.js`, backup: `${dir}/${BACKUP_NAME}`, data: `${dir}/data.json` } : null;
  }

  async onload() {
    this.addSettingTab(new OfficeReaderRuSettingTab(this.app, this));

    this.addCommand({
      id: 'apply',
      name: 'Применить перевод Office Reader заново',
      callback: () => this.apply(true),
    });
    this.addCommand({
      id: 'restore',
      name: 'Вернуть оригинальный Office Reader и отключить перевод',
      callback: () => this.restoreAndDisable(),
    });

    // К этому моменту все плагины уже загружены, можно спокойно перезапускать Office Reader.
    this.app.workspace.onLayoutReady(() => this.apply(false));
  }

  async status(): Promise<Status> {
    const p = this.paths();
    if (!p || !(await this.app.vault.adapter.exists(p.main))) return 'not-installed';
    const src = await this.app.vault.adapter.read(p.main);
    if (isCurrentPatch(src)) return 'patched';
    return isPatched(src) ? 'outdated' : 'original';
  }

  async apply(force: boolean) {
    const p = this.paths();
    if (!p) {
      if (force) new Notice('Office Reader не установлен');
      return;
    }
    const adapter = this.app.vault.adapter;

    try {
      const src = await adapter.read(p.main);
      if (isCurrentPatch(src) && !force) return;

      const firstInstall = !(await adapter.exists(p.backup));
      let original: string;
      if (isPatched(src)) {
        // Пропатчено прошлой версией перевода — берём оригинал из бэкапа.
        if (firstInstall) throw new Error('main.js уже изменён, а резервной копии нет. Переустановите Office Reader.');
        original = await adapter.read(p.backup);
      } else {
        original = src;
      }

      // Сначала патчим в памяти: если формат не распознан, на диске ничего не меняется.
      const patched = patchOfficeReader(original);
      // Первый запуск или Office Reader обновился: сохраняем свежий оригинал.
      if (original === src) await adapter.write(p.backup, original);
      await adapter.write(p.main, patched);
      if (firstInstall) await this.switchLanguageToRussian(p.data);
      await this.reloadTarget();
      new Notice('Office Reader переведён на русский');
    } catch (e) {
      console.error('[office-reader-ru]', e);
      new Notice(`Office Reader: не удалось применить перевод — ${(e as Error).message}`, 10000);
    }
  }

  // Русский живёт в слоте "en". Без этого пользователю пришлось бы искать
  // переключатель языка в китайском интерфейсе. Делается только при первой установке,
  // чтобы потом не перебивать осознанный выбор китайского.
  private async switchLanguageToRussian(dataPath: string) {
    const adapter = this.app.vault.adapter;
    if (!(await adapter.exists(dataPath))) return; // нет настроек — сработает новый язык по умолчанию
    try {
      const data = JSON.parse(await adapter.read(dataPath));
      let changed = false;
      if (data?.language === 'zh-CN') {
        data.language = 'en';
        changed = true;
      }
      if (data?.common?.language === 'zh-CN') {
        data.common.language = 'en';
        changed = true;
      }
      if (changed) await adapter.write(dataPath, JSON.stringify(data, null, 2));
    } catch (e) {
      console.warn('[office-reader-ru] не удалось переключить язык в data.json', e);
    }
  }

  async restoreAndDisable() {
    const p = this.paths();
    const adapter = this.app.vault.adapter;
    if (p && (await adapter.exists(p.backup))) {
      await adapter.write(p.main, await adapter.read(p.backup));
      await adapter.remove(p.backup);
      await this.reloadTarget();
    }
    new Notice('Оригинальный Office Reader восстановлен, перевод отключён');
    // Иначе при следующем запуске перевод применится снова.
    await this.plugins.disablePluginAndSave(this.manifest.id);
  }

  // Перезапускает Office Reader, чтобы он перечитал main.js, и возвращает открытые файлы.
  private async reloadTarget() {
    const plugins = this.plugins;
    if (!plugins.enabledPlugins.has(TARGET_ID) || !plugins.plugins[TARGET_ID]) return;

    const open: [WorkspaceLeaf, ViewState][] = TARGET_VIEWS.flatMap((type) =>
      this.app.workspace.getLeavesOfType(type).map((leaf): [WorkspaceLeaf, ViewState] => [leaf, leaf.getViewState()])
    );

    await plugins.disablePlugin(TARGET_ID);
    await plugins.enablePlugin(TARGET_ID);

    for (const [leaf, state] of open) await leaf.setViewState(state);
  }
}

const STATUS_TEXT: Record<Status, string> = {
  'not-installed': 'Плагин Office Reader не найден. Установите его из каталога сторонних плагинов.',
  original: 'Перевод ещё не применён.',
  patched: 'Перевод применён. Язык переключается в настройках Office Reader: «Русский» / «中文».',
  outdated: 'Перевод устарел и будет обновлён при следующем запуске.',
};

class OfficeReaderRuSettingTab extends PluginSettingTab {
  constructor(app: App, private plugin: OfficeReaderRuPlugin) {
    super(app, plugin);
  }

  async display() {
    const { containerEl } = this;
    containerEl.empty();

    const status = await this.plugin.status();
    new Setting(containerEl).setName('Состояние').setDesc(STATUS_TEXT[status]);

    new Setting(containerEl)
      .setName('Применить заново')
      .setDesc('Обычно не требуется: после обновления Office Reader перевод применяется автоматически при запуске.')
      .addButton((b) =>
        b.setButtonText('Применить').onClick(async () => {
          await this.plugin.apply(true);
          this.display();
        })
      );

    new Setting(containerEl)
      .setName('Вернуть оригинал')
      .setDesc('Восстанавливает исходный main.js Office Reader и отключает этот плагин.')
      .addButton((b) =>
        b
          .setButtonText('Вернуть')
          .setWarning()
          .onClick(() => this.plugin.restoreAndDisable())
      );
  }
}
