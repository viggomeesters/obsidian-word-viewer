import { strFromU8, unzipSync } from "fflate";

const LARGE_DOCUMENT_BYTES = 100 * 1024 * 1024;
const MANY_BLOCK_THRESHOLD = 1000;
const MANY_MEDIA_THRESHOLD = 500;
const BLOCK_RENDER_LIMIT = 500;
const COMMENT_RENDER_LIMIT = 120;
const NOTE_RENDER_LIMIT = 160;
const WARNING_RENDER_LIMIT = 100;
const ZIP_SIGNATURE_1 = 0x50;
const ZIP_SIGNATURE_2 = 0x4b;
const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0];

export interface DocxHyperlink {
  text: string;
  target: string;
  external: boolean;
}

export interface DocxBlock {
  index: number;
  type: "paragraph" | "table";
  label: string;
  text: string;
  style: string;
  headingLevel: number | null;
  list: boolean;
  hyperlinks: DocxHyperlink[];
  rows: string[][];
  rowCount: number;
  columnCount: number;
}

export interface DocxMedia {
  path: string;
  name: string;
  extension: string;
  size: number;
  contentType: string;
}

export interface DocxPackageEntry {
  path: string;
  size: number;
  directory: boolean;
}

export interface DocxComment {
  id: string;
  author: string;
  date: string;
  text: string;
}

export interface DocxNote {
  kind: "footnote" | "endnote";
  id: string;
  text: string;
}

export interface DocxSummary {
  blockCount: number;
  renderedBlockCount: number;
  headingCount: number;
  tableCount: number;
  hyperlinkCount: number;
  commentCount: number;
  noteCount: number;
  mediaCount: number;
  relationshipCount: number;
  externalRelationshipCount: number;
  packageEntryCount: number;
  trackedChangeCount: number;
}

export interface ParsedDocx {
  title: string;
  creator: string;
  application: string;
  blocks: DocxBlock[];
  renderedBlocks: DocxBlock[];
  comments: DocxComment[];
  renderedComments: DocxComment[];
  notes: DocxNote[];
  renderedNotes: DocxNote[];
  media: DocxMedia[];
  packageEntries: DocxPackageEntry[];
  warnings: string[];
  summary: DocxSummary;
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

interface ContentTypes {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

type ZipEntries = Record<string, Uint8Array>;

export async function parseDocx(data: ArrayBuffer): Promise<ParsedDocx> {
  validateDocxSignature(data);

  const bytes = new Uint8Array(data);
  const zip = unzipSync(bytes);
  const fileNames = Object.keys(zip);
  const warnings: string[] = [];

  if (!zip["[Content_Types].xml"]) {
    throw new Error("Package is missing [Content_Types].xml.");
  }
  if (!zip["word/document.xml"]) {
    throw new Error("Package is missing word/document.xml.");
  }
  if (bytes.byteLength > LARGE_DOCUMENT_BYTES) {
    warnings.push(`Document is ${formatBytes(bytes.byteLength)}; rendering is capped for responsiveness.`);
  }

  const contentTypes = parseContentTypes(await readText(zip, "[Content_Types].xml"));
  const documentXml = await readText(zip, "word/document.xml");
  const docProps = await readOptionalText(zip, "docProps/core.xml");
  const appProps = await readOptionalText(zip, "docProps/app.xml");
  const documentRelationships = await readRelationships(zip, "word/_rels/document.xml.rels");
  const relationshipFiles = await Promise.all(
    fileNames.filter((path) => path.endsWith(".rels")).map((path) => readRelationships(zip, path)),
  );
  const allRelationships = relationshipFiles.flat();
  const media = collectMedia(zip, contentTypes);
  const comments = parseComments(await readOptionalText(zip, "word/comments.xml") ?? "");
  const notes = [
    ...parseNotes(await readOptionalText(zip, "word/footnotes.xml") ?? "", "footnote"),
    ...parseNotes(await readOptionalText(zip, "word/endnotes.xml") ?? "", "endnote"),
  ];
  const blocks = parseDocumentBlocks(documentXml, documentRelationships);
  const packageEntries = fileNames
    .map((path) => ({
      directory: path.endsWith("/"),
      path,
      size: path.endsWith("/") ? 0 : zip[path]?.byteLength ?? 0,
    }))
    .sort((a, b) => a.path.localeCompare(b.path));

  const externalRelationshipCount = allRelationships.filter((rel) => rel.targetMode.toLowerCase() === "external").length;
  const trackedChangeCount = countMatches(documentXml, /<w:(?:ins|del)\b/g);
  const macroContentTypes = Array.from(contentTypes.overrides.values()).filter((type) => type.toLowerCase().includes("macroenabled"));
  const embeddedObjectCount = fileNames.filter((path) => /^word\/embeddings\//i.test(path)).length;

  if (blocks.length === 0) warnings.push("Document has no extracted paragraphs or tables.");
  if (blocks.length > MANY_BLOCK_THRESHOLD) warnings.push(`${blocks.length} content blocks found; rendered document list is capped.`);
  if (media.length > MANY_MEDIA_THRESHOLD) warnings.push(`${media.length} media files found; media list is capped.`);
  if (externalRelationshipCount > 0) warnings.push(`${externalRelationshipCount} external relationships are listed as metadata only.`);
  if (trackedChangeCount > 0) warnings.push(`${trackedChangeCount} tracked change markers detected; changes are shown as extracted text only.`);
  if (comments.length > COMMENT_RENDER_LIMIT) warnings.push(`${comments.length} comments found; rendered comment list is capped.`);
  if (notes.length > NOTE_RENDER_LIMIT) warnings.push(`${notes.length} footnotes/endnotes found; rendered note list is capped.`);
  if (embeddedObjectCount > 0) warnings.push(`${embeddedObjectCount} embedded object files detected; embedded objects are not opened or executed.`);
  if (macroContentTypes.length > 0 || fileNames.some((path) => /vbaProject\.bin$/i.test(path))) {
    warnings.push("Macro-enabled package content detected; macros are not opened or executed.");
  }
  documentRelationships
    .filter((rel) => rel.type.toLowerCase().includes("oleobject") || rel.type.toLowerCase().includes("package"))
    .forEach(() => warnings.push("Embedded object relationship detected; it is not opened or executed."));

  const renderedBlocks = blocks.slice(0, BLOCK_RENDER_LIMIT);

  return {
    application: firstXmlText(appProps ?? "", "Application"),
    blocks,
    comments,
    creator: firstXmlText(docProps ?? "", "dc:creator"),
    media,
    notes,
    packageEntries,
    renderedBlocks,
    renderedComments: comments.slice(0, COMMENT_RENDER_LIMIT),
    renderedNotes: notes.slice(0, NOTE_RENDER_LIMIT),
    summary: {
      blockCount: blocks.length,
      commentCount: comments.length,
      externalRelationshipCount,
      headingCount: blocks.filter((block) => block.headingLevel !== null).length,
      hyperlinkCount: blocks.reduce((total, block) => total + block.hyperlinks.length, 0),
      mediaCount: media.length,
      noteCount: notes.length,
      packageEntryCount: packageEntries.length,
      relationshipCount: allRelationships.length,
      renderedBlockCount: renderedBlocks.length,
      tableCount: blocks.filter((block) => block.type === "table").length,
      trackedChangeCount,
    },
    title: firstXmlText(docProps ?? "", "dc:title") || inferDocumentTitle(blocks),
    warnings: unique(warnings).slice(0, WARNING_RENDER_LIMIT),
  };
}

export function filterBlocks(blocks: DocxBlock[], query: string): DocxBlock[] {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return blocks;
  return blocks.filter((block) => {
    const haystack = [
      block.label,
      block.text,
      block.style,
      ...block.hyperlinks.map((link) => `${link.text} ${link.target}`),
      ...block.rows.flat(),
    ].join(" ").toLowerCase();
    return haystack.includes(normalized);
  });
}

export function filterMedia(media: DocxMedia[], query: string): DocxMedia[] {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return media;
  const extensionQuery = normalized.startsWith(".") ? normalized.slice(1) : "";
  return media.filter((item) => {
    return (
      item.path.toLowerCase().includes(normalized) ||
      item.contentType.toLowerCase().includes(normalized) ||
      Boolean(extensionQuery && item.extension.toLowerCase() === extensionQuery)
    );
  });
}

export function filterComments(comments: DocxComment[], query: string): DocxComment[] {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return comments;
  return comments.filter((comment) => [comment.id, comment.author, comment.text].join(" ").toLowerCase().includes(normalized));
}

export function filterNotes(notes: DocxNote[], query: string): DocxNote[] {
  const normalized = query.trim().toLowerCase();
  if (!normalized) return notes;
  return notes.filter((note) => [note.kind, note.id, note.text].join(" ").toLowerCase().includes(normalized));
}

function validateDocxSignature(data: ArrayBuffer): void {
  if (data.byteLength < 4) {
    throw new Error("File is too small to be a .docx package.");
  }
  const bytes = new Uint8Array(data, 0, 4);
  if (CFB_SIGNATURE.every((value, index) => bytes[index] === value)) {
    throw new Error("Encrypted, password-protected, or legacy Word files are not supported in v0.1.");
  }
  if (bytes[0] !== ZIP_SIGNATURE_1 || bytes[1] !== ZIP_SIGNATURE_2) {
    throw new Error("File is not a valid .docx zip package.");
  }
}

function parseDocumentBlocks(xml: string, relationships: Relationship[]): DocxBlock[] {
  const blocks: DocxBlock[] = [];
  const blockRegex = /<w:(p|tbl)\b[\s\S]*?<\/w:\1>/g;
  let index = 1;
  for (const match of xml.matchAll(blockRegex)) {
    const raw = match[0];
    if (match[1] === "tbl") {
      const rows = parseTableRows(raw);
      const text = rows.flat().join(" ").trim();
      blocks.push({
        columnCount: rows.reduce((max, row) => Math.max(max, row.length), 0),
        headingLevel: null,
        hyperlinks: extractHyperlinks(raw, relationships),
        index,
        label: `Table ${blocks.filter((block) => block.type === "table").length + 1}`,
        list: false,
        rowCount: rows.length,
        rows,
        style: "",
        text,
        type: "table",
      });
      index += 1;
      continue;
    }

    const text = extractTextRuns(raw).join("").replace(/\s+/g, " ").trim();
    const style = attributeValue(raw.match(/<w:pStyle\b[^>]*>/)?.[0] ?? "", "w:val");
    const headingLevel = headingLevelForStyle(style);
    const label = headingLevel ? text || `Heading ${index}` : `Paragraph ${index}`;
    if (text || raw.includes("<w:drawing") || raw.includes("<w:hyperlink")) {
      blocks.push({
        columnCount: 0,
        headingLevel,
        hyperlinks: extractHyperlinks(raw, relationships),
        index,
        label,
        list: raw.includes("<w:numPr"),
        rowCount: 0,
        rows: [],
        style,
        text,
        type: "paragraph",
      });
      index += 1;
    }
  }
  return blocks;
}

function parseTableRows(xml: string): string[][] {
  return [...xml.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map((rowMatch) => {
    return [...rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((cellMatch) => {
      return extractTextRuns(cellMatch[0]).join("").replace(/\s+/g, " ").trim();
    });
  });
}

function extractHyperlinks(xml: string, relationships: Relationship[]): DocxHyperlink[] {
  const links: DocxHyperlink[] = [];
  for (const match of xml.matchAll(/<w:hyperlink\b([^>]*)>([\s\S]*?)<\/w:hyperlink>/g)) {
    const attrs = match[1] ?? "";
    const rid = attributeValue(attrs, "r:id");
    const anchor = attributeValue(attrs, "w:anchor");
    const rel = rid ? relationships.find((item) => item.id === rid) : null;
    const text = extractTextRuns(match[2] ?? "").join("").trim();
    const target = rel ? normalizePackagePath("word/document.xml", rel.target) : anchor ? `#${anchor}` : "";
    links.push({
      external: rel?.targetMode.toLowerCase() === "external",
      target,
      text: text || target || "Hyperlink",
    });
  }
  return links;
}

function parseComments(xml: string): DocxComment[] {
  return [...xml.matchAll(/<w:comment\b([^>]*)>([\s\S]*?)<\/w:comment>/g)].map((match) => ({
    author: xmlDecode(attributeValue(match[1] ?? "", "w:author")),
    date: xmlDecode(attributeValue(match[1] ?? "", "w:date")),
    id: attributeValue(match[1] ?? "", "w:id"),
    text: extractTextRuns(match[2] ?? "").join("").replace(/\s+/g, " ").trim(),
  }));
}

function parseNotes(xml: string, kind: DocxNote["kind"]): DocxNote[] {
  const tag = kind === "footnote" ? "footnote" : "endnote";
  return [...xml.matchAll(new RegExp(`<w:${tag}\\b([^>]*)>([\\s\\S]*?)<\\/w:${tag}>`, "g"))]
    .map((match) => ({
      id: attributeValue(match[1] ?? "", "w:id"),
      kind,
      text: extractTextRuns(match[2] ?? "").join("").replace(/\s+/g, " ").trim(),
    }))
    .filter((note) => note.text.length > 0 && !note.id.startsWith("-"));
}

function collectMedia(zip: ZipEntries, contentTypes: ContentTypes): DocxMedia[] {
  return Object.entries(zip)
    .filter(([path]) => /^word\/media\//i.test(path) && !path.endsWith("/"))
    .map(([path, content]) => {
      const extension = extensionForPath(path);
      return {
        contentType: contentTypeForPath(path, contentTypes),
        extension,
        name: path.split("/").pop() ?? path,
        path,
        size: content.byteLength,
      };
    })
    .sort((a, b) => a.path.localeCompare(b.path));
}

function extractTextRuns(xml: string): string[] {
  return [...xml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:delText\b[^>]*>([\s\S]*?)<\/w:delText>/g)].map((match) => xmlDecode(match[1] ?? match[2] ?? ""));
}

function headingLevelForStyle(style: string): number | null {
  const normalized = style.toLowerCase();
  const heading = normalized.match(/^heading([1-6])$/);
  if (heading) return Number(heading[1]);
  if (normalized === "title") return 1;
  return null;
}

function inferDocumentTitle(blocks: DocxBlock[]): string {
  return blocks.find((block) => block.headingLevel !== null && block.text)?.text ?? blocks.find((block) => block.text)?.text.slice(0, 80) ?? "";
}

async function readText(zip: ZipEntries, path: string): Promise<string> {
  const value = zip[path];
  if (!value) throw new Error(`Package entry not found: ${path}`);
  return strFromU8(value);
}

async function readOptionalText(zip: ZipEntries, path: string): Promise<string | null> {
  const value = zip[path];
  return value ? strFromU8(value) : null;
}

async function readRelationships(zip: ZipEntries, path: string): Promise<Relationship[]> {
  const xml = await readOptionalText(zip, path);
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b([^>]*)\/?>/g)].map((match) => {
    const attrs = match[1] ?? "";
    return {
      id: attributeValue(attrs, "Id"),
      target: xmlDecode(attributeValue(attrs, "Target")),
      targetMode: attributeValue(attrs, "TargetMode"),
      type: attributeValue(attrs, "Type"),
    };
  });
}

function parseContentTypes(xml: string): ContentTypes {
  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();

  for (const match of xml.matchAll(/<Default\b([^>]*)\/?>/g)) {
    const attrs = match[1] ?? "";
    defaults.set(attributeValue(attrs, "Extension").toLowerCase(), attributeValue(attrs, "ContentType"));
  }
  for (const match of xml.matchAll(/<Override\b([^>]*)\/?>/g)) {
    const attrs = match[1] ?? "";
    overrides.set(stripLeadingSlash(attributeValue(attrs, "PartName")), attributeValue(attrs, "ContentType"));
  }

  return { defaults, overrides };
}

function contentTypeForPath(path: string, contentTypes: ContentTypes): string {
  return contentTypes.overrides.get(path) ?? contentTypes.defaults.get(extensionForPath(path).toLowerCase()) ?? "application/octet-stream";
}

function normalizePackagePath(basePath: string, target: string): string {
  if (!target) return "";
  if (/^[a-z][a-z0-9+.-]*:/i.test(target)) return target;
  if (target.startsWith("/")) return stripLeadingSlash(target);
  const baseParts = basePath.split("/");
  baseParts.pop();
  for (const part of target.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") baseParts.pop();
    else baseParts.push(part);
  }
  return baseParts.join("/");
}

function attributeValue(attrs: string, name: string): string {
  const escaped = name.replace(":", "\\:");
  const match = attrs.match(new RegExp(`(?:^|\\s)${escaped}=["']([^"']*)["']`));
  return match ? match[1] : "";
}

function firstXmlText(xml: string, tagName: string): string {
  const escaped = tagName.replace(":", "\\:");
  const match = xml.match(new RegExp(`<${escaped}\\b[^>]*>([\\s\\S]*?)<\\/${escaped}>`));
  return match ? xmlDecode(match[1]) : "";
}

function stripLeadingSlash(value: string): string {
  return value.startsWith("/") ? value.slice(1) : value;
}

function extensionForPath(path: string): string {
  const name = path.split("/").pop() ?? path;
  const dot = name.lastIndexOf(".");
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : "";
}

function xmlDecode(value: string): string {
  return value
    .replace(/<[^>]+>/g, "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function countMatches(value: string, pattern: RegExp): number {
  return [...value.matchAll(pattern)].length;
}

function unique(values: string[]): string[] {
  return [...new Set(values)];
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}
