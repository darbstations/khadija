// In-memory backing store for the `window.fs.readFile` shim that App.jsx expects.
//
// In the Claude artifact sandbox, `window.fs.readFile(name)` returns the bytes of
// an uploaded file. The browser has no such API, so we recreate it: the upload
// gate (FileGate.jsx) drops selected files in here keyed by filename, and App.jsx
// reads them back unchanged.

const files = new Map();

function normalize(name) {
  // Excel filenames in FILE_GROUPS sometimes carry double spaces / odd casing.
  // Collapse whitespace and lowercase so "MK072  3.xlsx" matches "MK072 3.xlsx".
  return String(name).replace(/\s+/g, " ").trim().toLowerCase();
}

export function setFiles(fileList) {
  files.clear();
  for (const file of fileList) {
    files.set(normalize(file.name), file);
  }
}

export function fileCount() {
  return files.size;
}

// Mirrors window.fs.readFile: resolves to a Uint8Array of the file's bytes.
export async function readFile(name) {
  const file = files.get(normalize(name));
  if (!file) {
    throw new Error(`الملف غير موجود (لم يتم رفعه): ${name}`);
  }
  const buf = await file.arrayBuffer();
  return new Uint8Array(buf);
}

// Install the shim App.jsx relies on, without modifying App.jsx itself.
export function installFsShim() {
  if (typeof window === "undefined") return;
  window.fs = window.fs || {};
  window.fs.readFile = readFile;
}
