// Loading and saving the OPC (zip) package around word/document.xml.
import JSZip from 'jszip';
import { NS, parseXml, serializeXml } from './xml.js';

const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const IMAGE_MIME = {
  png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
  bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml',
};

function resolvePath(baseDir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = (baseDir + target).split('/');
  const out = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

async function readXml(zip, path) {
  const file = zip.file(path);
  return file ? parseXml(await file.async('string')) : null;
}

/** Parse a .rels part into Map<id, {type, target, external}>; `target` is a package path unless external. */
async function readRels(zip, partPath) {
  const dir = partPath.slice(0, partPath.lastIndexOf('/') + 1);
  const relsPath = dir + '_rels/' + partPath.slice(dir.length) + '.rels';
  const doc = await readXml(zip, relsPath);
  const rels = new Map();
  if (!doc) return rels;
  for (const el of doc.getElementsByTagNameNS(NS.rel, 'Relationship')) {
    const external = el.getAttribute('TargetMode') === 'External';
    const target = el.getAttribute('Target');
    rels.set(el.getAttribute('Id'), {
      type: el.getAttribute('Type'),
      target: external ? target : resolvePath(dir, target),
      external,
    });
  }
  return rels;
}

export async function loadDocx(data) {
  const zip = await JSZip.loadAsync(data);
  const rootRels = await readRels(zip, '');
  const main = [...rootRels.values()].find((r) => r.type === REL_TYPE + 'officeDocument');
  if (!main) {
    const strict = [...rootRels.values()].some((r) => r.type.includes('purl.oclc.org'));
    throw new Error(strict ? 'Strict Open XML documents are not supported.' : 'Not a Word document.');
  }
  const mainPath = main.target;
  const doc = await readXml(zip, mainPath);
  if (doc?.documentElement.namespaceURI !== NS.w) throw new Error('Unrecognised main document part.');
  const rels = await readRels(zip, mainPath);
  const partOfType = (type) => [...rels.values()].find((r) => r.type === REL_TYPE + type)?.target;

  // Images are needed synchronously while rendering, so turn them into blob URLs up front.
  const media = new Map();
  for (const rel of rels.values()) {
    if (rel.external || rel.type !== REL_TYPE + 'image') continue;
    const mime = IMAGE_MIME[rel.target.split('.').pop().toLowerCase()];
    const file = zip.file(rel.target);
    if (mime && file && !media.has(rel.target)) {
      media.set(rel.target, URL.createObjectURL(new Blob([await file.async('uint8array')], { type: mime })));
    }
  }

  return {
    zip,
    mainPath,
    doc,
    rels,
    media,
    stylesDoc: await readXml(zip, partOfType('styles')),
    numberingDoc: await readXml(zip, partOfType('numbering')),
    themeDoc: await readXml(zip, partOfType('theme')),
    /** URL of an image relationship, or null if it can't be displayed. */
    imageUrl(rId) {
      const rel = rels.get(rId);
      return rel && !rel.external ? media.get(rel.target) ?? null : null;
    },
  };
}

/** Re-zip the package. Only the main document part is rewritten, and only if it was modified. */
export async function saveDocx(pkg, { modified }) {
  if (modified) pkg.zip.file(pkg.mainPath, serializeXml(pkg.doc), { createFolders: false });
  return pkg.zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

export function disposeDocx(pkg) {
  for (const url of pkg.media.values()) URL.revokeObjectURL(url);
}
