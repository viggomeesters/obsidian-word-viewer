import {
  FileView,
  Notice,
  Plugin,
  TFile,
  WorkspaceLeaf,
  setIcon,
} from "obsidian";
import {
  DocxBlock,
  DocxComment,
  DocxMedia,
  DocxNote,
  DocxPackageEntry,
  ParsedDocx,
  filterBlocks,
  filterComments,
  filterMedia,
  filterNotes,
  parseDocx,
} from "./parser";

const VIEW_TYPE_DOCX_VIEWER = "word-viewer";
const DOCX_EXTENSIONS = ["docx"];
const MEDIA_RENDER_LIMIT = 120;
const PACKAGE_RENDER_LIMIT = 240;

export default class DocxViewerPlugin extends Plugin {
  async onload(): Promise<void> {
    this.registerView(
      VIEW_TYPE_DOCX_VIEWER,
      (leaf) => new DocxViewerView(leaf),
    );
    this.registerExtensions(DOCX_EXTENSIONS, VIEW_TYPE_DOCX_VIEWER);

    this.addCommand({
      id: "open-current-docx-in-viewer",
      name: "Open current DOCX file in viewer",
      checkCallback: (checking) => {
        const file = this.app.workspace.getActiveFile();
        if (!isDocxFile(file)) return false;

        if (!checking) {
          void this.openDocxFile(file);
        }
        return true;
      },
    });
  }

  async openDocxFile(file: TFile): Promise<void> {
    const leaf = this.app.workspace.getLeaf(false);
    await leaf.setViewState({
      active: true,
      state: { file: file.path },
      type: VIEW_TYPE_DOCX_VIEWER,
    });
  }
}

class DocxViewerView extends FileView {
  private document: ParsedDocx | null = null;
  private activeBlockIndex = 1;
  private filterValue = "";
  private errorMessage = "";

  constructor(leaf: WorkspaceLeaf) {
    super(leaf);
  }

  getViewType(): string {
    return VIEW_TYPE_DOCX_VIEWER;
  }

  getDisplayText(): string {
    return this.file?.basename ?? "DOCX viewer";
  }

  getIcon(): string {
    return "file-text";
  }

  async onLoadFile(file: TFile): Promise<void> {
    await this.loadDocument(file);
  }

  async onUnloadFile(): Promise<void> {
    this.document = null;
    this.activeBlockIndex = 1;
    this.errorMessage = "";
    this.contentEl.empty();
  }

  private async loadDocument(file: TFile): Promise<void> {
    try {
      const data = await this.app.vault.readBinary(file);
      this.document = await parseDocx(data);
      this.activeBlockIndex = this.document.renderedBlocks[0]?.index ?? 1;
      this.errorMessage = "";
    } catch (error) {
      this.document = null;
      this.activeBlockIndex = 1;
      this.errorMessage = `Unable to read document: ${getErrorMessage(error)}`;
    }
    this.render();
  }

  private render(): void {
    const container = this.contentEl;
    container.empty();
    container.addClass("word-viewer");

    const header = container.createDiv({ cls: "word-viewer__header" });
    this.renderTitle(header);
    this.renderToolbar(header);

    if (!this.file) {
      renderMessage(container, "No DOCX file is attached to this viewer.");
      return;
    }
    if (!isDocxFile(this.file)) {
      renderMessage(container, "This viewer only supports .docx files.");
      return;
    }
    if (this.errorMessage) {
      renderMessage(container, this.errorMessage);
      return;
    }
    if (!this.document) {
      renderMessage(container, "Document is not loaded.");
      return;
    }

    renderSummary(container, this.document);
    renderWarnings(container, this.document.warnings, "Document warnings");

    const body = container.createDiv({ cls: "word-viewer__body" });
    this.renderOutline(body, this.document);
    this.renderDetail(body, this.document);
  }

  private renderTitle(parent: HTMLElement): void {
    const title = parent.createDiv({ cls: "word-viewer__title" });
    title.createDiv({
      cls: "word-viewer__filename",
      text: this.file?.name ?? "DOCX file",
    });
    title.createDiv({
      cls: "word-viewer__path",
      text: this.file?.path ?? "",
    });
  }

  private renderToolbar(parent: HTMLElement): void {
    const toolbar = parent.createDiv({ cls: "word-viewer__toolbar" });
    const searchWrap = toolbar.createDiv({ cls: "word-viewer__search" });
    setIcon(searchWrap.createSpan({ cls: "word-viewer__search-icon" }), "search");
    const searchInput = searchWrap.createEl("input", {
      attr: {
        "aria-label": "Filter document content",
        placeholder: "Filter",
        spellcheck: "false",
        type: "search",
        value: this.filterValue,
      },
    });
    searchInput.addEventListener("input", () => {
      this.filterValue = searchInput.value;
      this.render();
    });

    const refreshButton = createIconButton(toolbar, "refresh-cw", "Refresh document");
    refreshButton.addEventListener("click", () => {
      void this.reloadFile();
    });
  }

  private renderOutline(parent: HTMLElement, document: ParsedDocx): void {
    const sidebar = parent.createDiv({ cls: "word-viewer__sidebar" });
    sidebar.createDiv({ cls: "word-viewer__section-title", text: "Outline" });

    const filteredBlocks = filterBlocks(document.renderedBlocks, this.filterValue);
    const outlineBlocks = filteredBlocks.filter((block) => block.headingLevel !== null || block.type === "table");
    const blocks = outlineBlocks.length > 0 ? outlineBlocks : filteredBlocks.slice(0, 80);
    if (blocks.length === 0) {
      sidebar.createDiv({ cls: "word-viewer__empty", text: "No document blocks match the filter." });
      return;
    }

    blocks.forEach((block) => {
      const button = sidebar.createEl("button", {
        cls: "word-viewer__outline-button",
        attr: { type: "button" },
      });
      button.toggleClass("is-active", block.index === this.activeBlockIndex);
      button.toggleClass("is-heading", block.headingLevel !== null);
      button.toggleClass(`is-level-${block.headingLevel ?? 0}`, block.headingLevel !== null);
      button.createSpan({ cls: "word-viewer__outline-index", text: String(block.index) });
      const label = button.createSpan({ cls: "word-viewer__outline-label" });
      label.createSpan({ cls: "word-viewer__outline-title", text: block.label || block.text || block.type });
      label.createSpan({
        cls: "word-viewer__outline-meta",
        text: block.type === "table" ? `${block.rowCount} rows, ${block.columnCount} columns` : block.style || "paragraph",
      });
      button.addEventListener("click", () => {
        this.activeBlockIndex = block.index;
        this.render();
      });
    });
  }

  private renderDetail(parent: HTMLElement, document: ParsedDocx): void {
    const detail = parent.createDiv({ cls: "word-viewer__detail" });
    renderMetadata(detail, document);

    const filteredBlocks = filterBlocks(document.renderedBlocks, this.filterValue);
    renderBlocks(detail, filteredBlocks, this.activeBlockIndex);
    renderComments(detail, filterComments(document.renderedComments, this.filterValue));
    renderNotes(detail, filterNotes(document.renderedNotes, this.filterValue));
    renderMedia(detail, filterMedia(document.media, this.filterValue));
    renderPackageEntries(detail, document.packageEntries);
  }

  private async reloadFile(): Promise<void> {
    if (!this.file) {
      new Notice("No DOCX file to refresh");
      return;
    }
    await this.loadDocument(this.file);
  }
}

function renderSummary(parent: HTMLElement, document: ParsedDocx): void {
  const summary = parent.createDiv({ cls: "word-viewer__summary" });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.blockCount} blocks` });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.headingCount} headings` });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.tableCount} tables` });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.commentCount} comments` });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.noteCount} notes` });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.mediaCount} media` });
  summary.createSpan({ cls: "word-viewer__pill", text: `${document.summary.packageEntryCount} package entries` });
  if (document.summary.externalRelationshipCount > 0) {
    summary.createSpan({ cls: "word-viewer__note", text: `${document.summary.externalRelationshipCount} external relationships listed` });
  }
  if (document.summary.trackedChangeCount > 0) {
    summary.createSpan({ cls: "word-viewer__note", text: `${document.summary.trackedChangeCount} tracked change markers` });
  }
  if (document.summary.renderedBlockCount < document.summary.blockCount) {
    summary.createSpan({ cls: "word-viewer__note", text: `${document.summary.renderedBlockCount} blocks rendered` });
  }
}

function renderMetadata(parent: HTMLElement, document: ParsedDocx): void {
  const section = parent.createDiv({ cls: "word-viewer__card" });
  const heading = section.createDiv({ cls: "word-viewer__detail-heading" });
  heading.createDiv({ cls: "word-viewer__detail-title", text: document.title || "Untitled document" });
  heading.createDiv({ cls: "word-viewer__detail-path", text: [document.creator, document.application].filter(Boolean).join(" · ") || "Metadata unavailable" });
}

function renderBlocks(parent: HTMLElement, blocks: DocxBlock[], activeBlockIndex: number): void {
  const section = parent.createDiv({ cls: "word-viewer__card" });
  section.createDiv({ cls: "word-viewer__section-title", text: "Document content" });
  if (blocks.length === 0) {
    section.createDiv({ cls: "word-viewer__empty", text: "No document content matches the filter." });
    return;
  }

  blocks.forEach((block) => {
    const item = section.createDiv({ cls: "word-viewer__block" });
    item.toggleClass("is-active", block.index === activeBlockIndex);
    item.toggleClass("is-heading", block.headingLevel !== null);
    const meta = item.createDiv({ cls: "word-viewer__block-meta" });
    meta.createSpan({ text: `${block.index}. ${block.type}` });
    if (block.headingLevel !== null) meta.createSpan({ text: `H${block.headingLevel}` });
    if (block.list) meta.createSpan({ text: "list" });
    if (block.style) meta.createSpan({ text: block.style });

    if (block.type === "table") {
      renderTable(item, block);
    } else {
      item.createDiv({ cls: "word-viewer__paragraph", text: block.text || "(empty paragraph)" });
    }
    renderHyperlinks(item, block.hyperlinks);
  });
}

function renderTable(parent: HTMLElement, block: DocxBlock): void {
  const wrap = parent.createDiv({ cls: "word-viewer__table-wrap" });
  const table = wrap.createEl("table", { cls: "word-viewer__table" });
  block.rows.forEach((row) => {
    const tr = table.createEl("tr");
    row.forEach((cell) => {
      tr.createEl("td", { text: cell || " " });
    });
  });
}

function renderHyperlinks(parent: HTMLElement, links: DocxBlock["hyperlinks"]): void {
  if (links.length === 0) return;
  const list = parent.createDiv({ cls: "word-viewer__links" });
  links.forEach((link) => {
    const pill = list.createSpan({ cls: "word-viewer__link-pill" });
    pill.createSpan({ text: link.text });
    pill.createSpan({ cls: "word-viewer__link-target", text: link.target });
    if (link.external) pill.createSpan({ cls: "word-viewer__link-external", text: "external" });
  });
}

function renderComments(parent: HTMLElement, comments: DocxComment[]): void {
  const section = parent.createDiv({ cls: "word-viewer__card" });
  section.createDiv({ cls: "word-viewer__section-title", text: "Comments" });
  if (comments.length === 0) {
    section.createDiv({ cls: "word-viewer__empty", text: "No comments found." });
    return;
  }
  comments.forEach((comment) => {
    const item = section.createDiv({ cls: "word-viewer__annotation" });
    item.createDiv({ cls: "word-viewer__annotation-meta", text: `#${comment.id} ${comment.author || "Unknown author"} ${comment.date}`.trim() });
    item.createDiv({ cls: "word-viewer__annotation-text", text: comment.text || "(empty comment)" });
  });
}

function renderNotes(parent: HTMLElement, notes: DocxNote[]): void {
  const section = parent.createDiv({ cls: "word-viewer__card" });
  section.createDiv({ cls: "word-viewer__section-title", text: "Footnotes and endnotes" });
  if (notes.length === 0) {
    section.createDiv({ cls: "word-viewer__empty", text: "No footnotes or endnotes found." });
    return;
  }
  notes.forEach((note) => {
    const item = section.createDiv({ cls: "word-viewer__annotation" });
    item.createDiv({ cls: "word-viewer__annotation-meta", text: `${note.kind} #${note.id}` });
    item.createDiv({ cls: "word-viewer__annotation-text", text: note.text });
  });
}

function renderMedia(parent: HTMLElement, media: DocxMedia[]): void {
  const section = parent.createDiv({ cls: "word-viewer__card" });
  section.createDiv({ cls: "word-viewer__section-title", text: "Media" });
  if (media.length === 0) {
    section.createDiv({ cls: "word-viewer__empty", text: "No embedded media found." });
    return;
  }

  const grid = section.createDiv({ cls: "word-viewer__media-grid" });
  media.slice(0, MEDIA_RENDER_LIMIT).forEach((item) => {
    const card = grid.createDiv({ cls: "word-viewer__media-card" });
    card.createDiv({ cls: "word-viewer__media-name", text: item.name });
    card.createDiv({ cls: "word-viewer__media-path", text: item.path });
    card.createDiv({ cls: "word-viewer__media-meta", text: `${item.contentType} · ${formatBytes(item.size)}` });
  });
  if (media.length > MEDIA_RENDER_LIMIT) {
    section.createDiv({ cls: "word-viewer__note", text: `${media.length - MEDIA_RENDER_LIMIT} more media files hidden by render cap.` });
  }
}

function renderPackageEntries(parent: HTMLElement, entries: DocxPackageEntry[]): void {
  const section = parent.createDiv({ cls: "word-viewer__card" });
  section.createDiv({ cls: "word-viewer__section-title", text: "Package diagnostics" });
  const table = section.createEl("table", { cls: "word-viewer__package-table" });
  const head = table.createEl("thead").createEl("tr");
  head.createEl("th", { text: "Path" });
  head.createEl("th", { text: "Size" });
  entries.slice(0, PACKAGE_RENDER_LIMIT).forEach((entry) => {
    const row = table.createEl("tr");
    row.createEl("td", { text: entry.path });
    row.createEl("td", { text: entry.directory ? "directory" : formatBytes(entry.size) });
  });
  if (entries.length > PACKAGE_RENDER_LIMIT) {
    section.createDiv({ cls: "word-viewer__note", text: `${entries.length - PACKAGE_RENDER_LIMIT} package entries hidden by render cap.` });
  }
}

function renderWarnings(parent: HTMLElement, warnings: string[], title: string): void {
  if (warnings.length === 0) return;
  const box = parent.createDiv({ cls: "word-viewer__warnings" });
  box.createDiv({ cls: "word-viewer__warnings-title", text: title });
  const list = box.createEl("ul");
  warnings.forEach((warning) => list.createEl("li", { text: warning }));
}

function renderMessage(parent: HTMLElement, message: string): void {
  parent.createDiv({ cls: "word-viewer__message", text: message });
}

function createIconButton(parent: HTMLElement, icon: string, label: string): HTMLButtonElement {
  const button = parent.createEl("button", {
    cls: "word-viewer__icon-button",
    attr: {
      "aria-label": label,
      title: label,
      type: "button",
    },
  });
  setIcon(button, icon);
  return button;
}

function isDocxFile(file: TFile | null): file is TFile {
  return file?.extension.toLowerCase() === "docx";
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
