import fs from "node:fs";

const manifest = JSON.parse(fs.readFileSync("manifest.json", "utf8"));
const versions = JSON.parse(fs.readFileSync("versions.json", "utf8"));
const main = fs.readFileSync("src/main.ts", "utf8");
const parser = fs.readFileSync("src/parser.ts", "utf8");
const styles = fs.readFileSync("styles.css", "utf8");
const bundle = fs.readFileSync("main.js", "utf8");

const token = (...parts) => parts.join("");
const forbidden = [
  token("fet", "ch("),
  token("XML", "HttpRequest"),
  token("Web", "Socket"),
  token("navigator", ".clipboard"),
  token("child", "_process"),
  token("sp", "awn("),
  token("ex", "ec("),
  token("ev", "al("),
  token("new", " Function"),
  token("vault", ".modify"),
  token("vault", ".adapter"),
];

const assertions = [
  [manifest.id === "word-viewer", "manifest id is word-viewer"],
  [manifest.name === "Word Viewer", "manifest name is Word Viewer"],
  [manifest.version === "0.1.1", "manifest version is 0.1.1"],
  [versions[manifest.version] === manifest.minAppVersion, "versions.json maps manifest version"],
  [!/obsidian/i.test(manifest.description), "manifest description avoids product name"],
  [/^[a-z-]+$/.test(manifest.id) && !manifest.id.includes("obsidian") && !manifest.id.endsWith("plugin"), "manifest id follows directory rules"],
  [main.includes('const DOCX_EXTENSIONS = ["docx"]'), "docx extension is the v0.1 scope"],
  [main.includes("registerExtensions(DOCX_EXTENSIONS"), "docx extension is registered"],
  [main.includes("extends FileView"), "FileView is used"],
  [main.includes("this.app.vault.readBinary(file)"), "vault binary reader is used"],
  [parser.includes("unzipSync") && parser.includes("strFromU8"), "local zip parser dependency is used"],
  [parser.includes("Encrypted, password-protected, or legacy Word files are not supported"), "encrypted/legacy state exists"],
  [parser.includes("Macro-enabled package content detected"), "macro warning exists"],
  [parser.includes("Embedded object relationship detected"), "embedded object warning exists"],
  [parser.includes("BLOCK_RENDER_LIMIT = 500"), "document block render cap exists"],
  [forbidden.every((token) => !main.includes(token) && !parser.includes(token) && !bundle.includes(token)), "forbidden runtime APIs are absent"],
  [!styles.includes(`!${"important"}`), "styles do not use important overrides"],
  [fs.existsSync("fixtures/simple.docx"), "simple docx fixture exists"],
  [fs.existsSync("fixtures/no-comments.docx"), "no-comments docx fixture exists"],
  [fs.existsSync("fixtures/embedded-object.docx"), "embedded object docx fixture exists"],
  [fs.existsSync("fixtures/large.docx"), "large docx fixture exists"],
  [fs.existsSync("fixtures/malformed.docx"), "malformed docx fixture exists"],
  [fs.existsSync("fixtures/encrypted.docx"), "encrypted docx fixture exists"],
  [fs.existsSync("README.md") && fs.existsSync("SECURITY.md") && fs.existsSync("LICENSE"), "docs and license exist"],
  [fs.existsSync("assets/hero.svg") && fs.existsSync("assets/social-preview.svg") && fs.existsSync("assets/screenshot.svg"), "visual assets exist"],
  [fs.existsSync("main.js") && fs.existsSync("styles.css") && fs.existsSync("manifest.json"), "release assets exist"],
];

const failures = assertions.filter(([passes]) => !passes).map(([, label]) => label);

if (failures.length > 0) {
  throw new Error(`Smoke checks failed: ${failures.join("; ")}`);
}

console.log("Word Viewer smoke checks passed.");
