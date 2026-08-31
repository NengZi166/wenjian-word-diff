import JSZip from 'jszip';

const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STORY_PART = /^word\/(document|footnotes|endnotes|header\d+|footer\d+)\.xml$/i;
const MAX_FILE_BYTES = 100 * 1024 * 1024;

export type FormulaKind = 'omml';
export type FormulaChangeKind = 'modified' | 'added' | 'deleted';
export type ObjectKind = 'ole' | 'image';

export interface FormulaItem {
  id: string;
  kind: FormulaKind;
  index: number;
  location: string;
  part: string;
  paragraph: number | null;
  latex: string;
  text: string;
  signature: string;
}

export interface EmbeddedObject {
  id: string;
  index: number;
  kind: ObjectKind;
  label: string;
  location: string;
  programId?: string;
  formulaLike: boolean;
  signature: string;
}

export interface DocxAnalysis {
  name: string;
  size: number;
  formulas: FormulaItem[];
  objects: EmbeddedObject[];
  warnings: string[];
}

export interface FormulaChange {
  id: string;
  kind: FormulaChangeKind;
  before?: FormulaItem;
  after?: FormulaItem;
}

export interface ObjectChange {
  id: string;
  kind: FormulaChangeKind;
  before?: EmbeddedObject;
  after?: EmbeddedObject;
}

export interface FormulaComparison {
  changes: FormulaChange[];
  unchanged: number;
  beforeCount: number;
  afterCount: number;
}

export interface ObjectComparison {
  changes: ObjectChange[];
  unchanged: number;
}

function childElements(node: Element): Element[] {
  return Array.from(node.children);
}

function directChild(node: Element, localName: string): Element | undefined {
  return childElements(node).find((child) => child.localName === localName);
}

function descendant(node: Element, localName: string): Element | undefined {
  return Array.from(node.getElementsByTagNameNS('*', localName))[0] as Element | undefined;
}

function attributeByLocalName(node: Element | undefined, localName: string): string {
  if (!node) return '';
  return Array.from(node.attributes).find((attribute) => attribute.localName === localName)?.value ?? '';
}

function childrenLatex(node: Element | undefined): string {
  if (!node) return '';
  return childElements(node).map(ommlElementToLatex).join('');
}

const symbolMap: Record<string, string> = {
  '−': '-', '×': '\\times ', '÷': '\\div ', '·': '\\cdot ', '±': '\\pm ',
  '≤': '\\le ', '≥': '\\ge ', '≠': '\\ne ', '≈': '\\approx ', '∞': '\\infty ',
  '∂': '\\partial ', '∈': '\\in ', '∉': '\\notin ', '∪': '\\cup ', '∩': '\\cap ',
  '→': '\\to ', '←': '\\leftarrow ', '↔': '\\leftrightarrow ', '⇒': '\\Rightarrow ',
  'α': '\\alpha ', 'β': '\\beta ', 'γ': '\\gamma ', 'δ': '\\delta ', 'ε': '\\varepsilon ',
  'θ': '\\theta ', 'λ': '\\lambda ', 'μ': '\\mu ', 'π': '\\pi ', 'ρ': '\\rho ',
  'σ': '\\sigma ', 'τ': '\\tau ', 'φ': '\\varphi ', 'ω': '\\omega ',
  'Γ': '\\Gamma ', 'Δ': '\\Delta ', 'Θ': '\\Theta ', 'Λ': '\\Lambda ', 'Π': '\\Pi ',
  'Σ': '\\Sigma ', 'Φ': '\\Phi ', 'Ω': '\\Omega ',
};

function escapeMathText(value: string): string {
  let output = '';
  for (const char of value) {
    if (symbolMap[char]) {
      output += symbolMap[char];
    } else if (char === '{' || char === '}' || char === '#' || char === '%' || char === '&' || char === '_') {
      output += `\\${char}`;
    } else if (char === '\\') {
      output += '\\backslash ';
    } else if (/\s/u.test(char)) {
      output += '\\, ';
    } else if (/[^\x00-\x7F]/u.test(char)) {
      output += `\\text{${char}}`;
    } else {
      output += char;
    }
  }
  return output;
}

function operatorLatex(value: string): string {
  const operators: Record<string, string> = {
    '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint',
    '∭': '\\iiint', '∮': '\\oint', '⋂': '\\bigcap', '⋃': '\\bigcup',
  };
  return operators[value] ?? escapeMathText(value || '∫');
}

function ommlElementToLatex(node: Element): string {
  const name = node.localName;
  if (name.endsWith('Pr') || name === 'ctrlPr') return '';

  switch (name) {
    case 'oMath':
      return childElements(node).map(ommlElementToLatex).join('');
    case 'oMathPara':
      return childElements(node).filter((child) => child.localName === 'oMath').map(ommlElementToLatex).join('\\quad ');
    case 't':
      return escapeMathText(node.textContent ?? '');
    case 'r': {
      const textNodes = Array.from(node.getElementsByTagNameNS(MATH_NS, 't'));
      return textNodes.map((textNode) => escapeMathText(textNode.textContent ?? '')).join('');
    }
    case 'f': {
      const numerator = directChild(node, 'num');
      const denominator = directChild(node, 'den');
      return `\\frac{${childrenLatex(numerator)}}{${childrenLatex(denominator)}}`;
    }
    case 'sSup':
      return `{${childrenLatex(directChild(node, 'e'))}}^{${childrenLatex(directChild(node, 'sup'))}}`;
    case 'sSub':
      return `{${childrenLatex(directChild(node, 'e'))}}_{${childrenLatex(directChild(node, 'sub'))}}`;
    case 'sSubSup':
      return `{${childrenLatex(directChild(node, 'e'))}}_{${childrenLatex(directChild(node, 'sub'))}}^{${childrenLatex(directChild(node, 'sup'))}}`;
    case 'sPre':
      return `{}_{${childrenLatex(directChild(node, 'sub'))}}^{${childrenLatex(directChild(node, 'sup'))}}${childrenLatex(directChild(node, 'e'))}`;
    case 'rad': {
      const degree = childrenLatex(directChild(node, 'deg'));
      const body = childrenLatex(directChild(node, 'e'));
      return degree ? `\\sqrt[${degree}]{${body}}` : `\\sqrt{${body}}`;
    }
    case 'nary': {
      const properties = directChild(node, 'naryPr');
      const character = attributeByLocalName(descendant(properties ?? node, 'chr'), 'val');
      const sub = childrenLatex(directChild(node, 'sub'));
      const sup = childrenLatex(directChild(node, 'sup'));
      const body = childrenLatex(directChild(node, 'e'));
      return `${operatorLatex(character)}${sub ? `_{${sub}}` : ''}${sup ? `^{${sup}}` : ''}{${body}}`;
    }
    case 'd': {
      const properties = directChild(node, 'dPr');
      const begin = attributeByLocalName(descendant(properties ?? node, 'begChr'), 'val') || '(';
      const end = attributeByLocalName(descendant(properties ?? node, 'endChr'), 'val') || ')';
      const separator = attributeByLocalName(descendant(properties ?? node, 'sepChr'), 'val') || '|';
      const entries = childElements(node).filter((child) => child.localName === 'e').map(childrenLatex);
      return `\\left${escapeMathText(begin)}${entries.join(`\\mathrel{${escapeMathText(separator)}}`)}\\right${escapeMathText(end)}`;
    }
    case 'func':
      return `\\operatorname{${childrenLatex(directChild(node, 'fName'))}}${childrenLatex(directChild(node, 'e'))}`;
    case 'limLow':
      return `{${childrenLatex(directChild(node, 'e'))}}_{${childrenLatex(directChild(node, 'lim'))}}`;
    case 'limUpp':
      return `{${childrenLatex(directChild(node, 'e'))}}^{${childrenLatex(directChild(node, 'lim'))}}`;
    case 'acc': {
      const character = attributeByLocalName(descendant(directChild(node, 'accPr') ?? node, 'chr'), 'val');
      const command: Record<string, string> = {
        '̂': 'hat', '^': 'hat', '¯': 'bar', '̅': 'bar', '→': 'vec', '˙': 'dot', '¨': 'ddot', '~': 'tilde',
      };
      return `\\${command[character] ?? 'widehat'}{${childrenLatex(directChild(node, 'e'))}}`;
    }
    case 'bar': {
      const position = attributeByLocalName(descendant(directChild(node, 'barPr') ?? node, 'pos'), 'val');
      return `\\${position === 'bot' ? 'underline' : 'overline'}{${childrenLatex(directChild(node, 'e'))}}`;
    }
    case 'groupChr': {
      const position = attributeByLocalName(descendant(directChild(node, 'groupChrPr') ?? node, 'pos'), 'val');
      return `\\${position === 'bot' ? 'underbrace' : 'overbrace'}{${childrenLatex(directChild(node, 'e'))}}`;
    }
    case 'm': {
      const rows = childElements(node).filter((child) => child.localName === 'mr').map((row) =>
        childElements(row).filter((child) => child.localName === 'e').map(childrenLatex).join(' & '),
      );
      return `\\begin{matrix}${rows.join(' \\\\ ')}\\end{matrix}`;
    }
    case 'eqArr': {
      const rows = childElements(node).filter((child) => child.localName === 'e').map(childrenLatex);
      return `\\begin{aligned}${rows.join(' \\\\ ')}\\end{aligned}`;
    }
    case 'box':
    case 'borderBox':
    case 'phant':
      return `{${childrenLatex(directChild(node, 'e'))}}`;
    default:
      return childElements(node).map(ommlElementToLatex).join('');
  }
}

function normalizeFormula(latex: string, text: string): string {
  return (latex || text)
    .replace(/\\left|\\right/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function storyLabel(path: string): string {
  if (/document\.xml$/i.test(path)) return '正文';
  if (/header/i.test(path)) return '页眉';
  if (/footer/i.test(path)) return '页脚';
  if (/footnotes/i.test(path)) return '脚注';
  if (/endnotes/i.test(path)) return '尾注';
  return path;
}

function closestParagraph(element: Element): Element | null {
  let current: Element | null = element.parentElement;
  while (current) {
    if (current.namespaceURI === WORD_NS && current.localName === 'p') return current;
    current = current.parentElement;
  }
  return null;
}

function parseXml(xml: string, path: string): XMLDocument {
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  if (document.getElementsByTagName('parsererror').length) {
    throw new Error(`无法解析 ${path}，文件可能已损坏。`);
  }
  return document;
}

function normalizeZipPath(path: string): string {
  const output: string[] = [];
  for (const segment of path.replace(/\\/g, '/').split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') output.pop();
    else output.push(segment);
  }
  return output.join('/');
}

function relationshipPath(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  return `${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`;
}

async function sha256(bytes: Uint8Array): Promise<string> {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  const digest = await crypto.subtle.digest('SHA-256', copy.buffer);
  return Array.from(new Uint8Array(digest)).map((byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function relationshipsForPart(zip: JSZip, partPath: string): Promise<Map<string, string>> {
  const relationshipFile = zip.file(relationshipPath(partPath));
  const map = new Map<string, string>();
  if (!relationshipFile) return map;
  const xml = parseXml(await relationshipFile.async('string'), relationshipPath(partPath));
  for (const relation of Array.from(xml.getElementsByTagNameNS('*', 'Relationship'))) {
    const id = relation.getAttribute('Id') ?? '';
    const target = relation.getAttribute('Target') ?? '';
    const mode = relation.getAttribute('TargetMode') ?? '';
    if (id && target && mode.toLowerCase() !== 'external') {
      const base = partPath.slice(0, partPath.lastIndexOf('/') + 1);
      map.set(id, target.startsWith('/') ? target.slice(1) : normalizeZipPath(`${base}${target}`));
    }
  }
  return map;
}

async function signatureForTarget(zip: JSZip, target: string, cache: Map<string, string>): Promise<string> {
  if (!target) return '';
  if (cache.has(target)) return cache.get(target) ?? '';
  const entry = zip.file(target);
  if (!entry) return '';
  const signature = await sha256(await entry.async('uint8array'));
  cache.set(target, signature);
  return signature;
}

function relId(element: Element): string {
  const direct = element.getAttributeNS(REL_NS, 'id') || element.getAttributeNS(REL_NS, 'embed');
  if (direct) return direct;
  const nested = Array.from(element.getElementsByTagNameNS('*', '*')).find((child) =>
    child.hasAttributeNS(REL_NS, 'id') || child.hasAttributeNS(REL_NS, 'embed'),
  );
  return nested?.getAttributeNS(REL_NS, 'id') || nested?.getAttributeNS(REL_NS, 'embed') || '';
}

function paragraphMap(document: XMLDocument): Map<Element, number> {
  const map = new Map<Element, number>();
  Array.from(document.getElementsByTagNameNS(WORD_NS, 'p')).forEach((paragraph, index) => map.set(paragraph, index + 1));
  return map;
}

export async function analyzeDocx(file: File): Promise<DocxAnalysis> {
  if (!file.name.toLowerCase().endsWith('.docx')) throw new Error('目前仅支持 .docx 文件。');
  if (file.size > MAX_FILE_BYTES) throw new Error('单个文件不能超过 100 MB。');

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(file);
  } catch {
    throw new Error('无法打开该 DOCX；文件可能损坏、加密或设有打开密码。');
  }
  if (!zip.file('[Content_Types].xml') || !zip.file('word/document.xml')) {
    throw new Error('这不是有效的 Word DOCX 文件。');
  }

  const formulas: FormulaItem[] = [];
  const objects: EmbeddedObject[] = [];
  const warnings: string[] = [];
  const hashCache = new Map<string, string>();
  const storyPaths = Object.keys(zip.files).filter((path) => STORY_PART.test(path)).sort((a, b) => {
    if (a === 'word/document.xml') return -1;
    if (b === 'word/document.xml') return 1;
    return a.localeCompare(b);
  });

  for (const partPath of storyPaths) {
    const entry = zip.file(partPath);
    if (!entry) continue;
    const document = parseXml(await entry.async('string'), partPath);
    const paragraphNumbers = paragraphMap(document);
    const relationships = await relationshipsForPart(zip, partPath);
    const label = storyLabel(partPath);

    const formulaNodes = Array.from(document.getElementsByTagNameNS(MATH_NS, 'oMath')).filter((formula) => {
      let parent = formula.parentElement;
      while (parent) {
        if (parent.namespaceURI === MATH_NS && parent.localName === 'oMath') return false;
        parent = parent.parentElement;
      }
      return true;
    });

    for (const formula of formulaNodes) {
      const paragraph = closestParagraph(formula);
      const paragraphNumber = paragraph ? paragraphNumbers.get(paragraph) ?? null : null;
      const latex = ommlElementToLatex(formula).trim();
      const text = Array.from(formula.getElementsByTagNameNS(MATH_NS, 't')).map((node) => node.textContent ?? '').join('');
      const index = formulas.length + 1;
      formulas.push({
        id: `formula-${index}`,
        kind: 'omml',
        index,
        part: partPath,
        paragraph: paragraphNumber,
        location: paragraphNumber ? `${label} · 第 ${paragraphNumber} 段` : label,
        latex: latex || escapeMathText(text) || '\\text{空公式}',
        text,
        signature: normalizeFormula(latex, text),
      });
    }

    const allElements = Array.from(document.getElementsByTagNameNS('*', '*'));
    for (const ole of allElements.filter((element) => element.localName === 'OLEObject')) {
      const programId = attributeByLocalName(ole, 'ProgID');
      const id = relId(ole);
      const target = relationships.get(id) ?? '';
      const paragraph = closestParagraph(ole);
      const paragraphNumber = paragraph ? paragraphNumbers.get(paragraph) ?? null : null;
      const formulaLike = /equation|mathtype|dsm?t/i.test(programId) || /equation|mathtype/i.test(target);
      const index = objects.length + 1;
      objects.push({
        id: `object-${index}`,
        index,
        kind: 'ole',
        label: formulaLike ? 'MathType / OLE 公式' : '嵌入式 OLE 对象',
        location: paragraphNumber ? `${label} · 第 ${paragraphNumber} 段` : label,
        programId,
        formulaLike,
        signature: (await signatureForTarget(zip, target, hashCache)) || `${programId}:${target}`,
      });
    }

    for (const drawing of allElements.filter((element) => element.localName === 'drawing' || element.localName === 'pict')) {
      const id = relId(drawing);
      if (!id) continue;
      const target = relationships.get(id) ?? '';
      if (!/\.(png|jpe?g|gif|bmp|tiff?|emf|wmf|svg)$/i.test(target)) continue;
      const descriptor = Array.from(drawing.getElementsByTagNameNS('*', 'docPr'))[0] as Element | undefined;
      const alt = [attributeByLocalName(descriptor, 'name'), attributeByLocalName(descriptor, 'title'), attributeByLocalName(descriptor, 'descr')].join(' ');
      const formulaLike = /formula|equation|math|公式|方程/i.test(alt);
      const paragraph = closestParagraph(drawing);
      const paragraphNumber = paragraph ? paragraphNumbers.get(paragraph) ?? null : null;
      const index = objects.length + 1;
      objects.push({
        id: `object-${index}`,
        index,
        kind: 'image',
        label: formulaLike ? '图片公式' : '图片对象',
        location: paragraphNumber ? `${label} · 第 ${paragraphNumber} 段` : label,
        formulaLike,
        signature: (await signatureForTarget(zip, target, hashCache)) || target,
      });
    }
  }

  if (objects.some((object) => object.kind === 'ole' && object.formulaLike)) {
    warnings.push('检测到 MathType/OLE 公式：可判断对象是否变化，但不能解析公式内部的单个符号。');
  }
  if (objects.some((object) => object.kind === 'image' && object.formulaLike)) {
    warnings.push('检测到图片公式：可判断图片是否变化，但不进行 OCR 公式识别。');
  }

  return { name: file.name, size: file.size, formulas, objects, warnings };
}

function lcsMatches<T extends { signature: string }>(before: T[], after: T[]): Array<[number, number]> {
  const rows = before.length + 1;
  const columns = after.length + 1;
  const table = Array.from({ length: rows }, () => new Uint16Array(columns));
  for (let i = 1; i < rows; i += 1) {
    for (let j = 1; j < columns; j += 1) {
      table[i][j] = before[i - 1].signature === after[j - 1].signature
        ? table[i - 1][j - 1] + 1
        : Math.max(table[i - 1][j], table[i][j - 1]);
    }
  }
  const matches: Array<[number, number]> = [];
  let i = before.length;
  let j = after.length;
  while (i > 0 && j > 0) {
    if (before[i - 1].signature === after[j - 1].signature) {
      matches.push([i - 1, j - 1]);
      i -= 1;
      j -= 1;
    } else if (table[i - 1][j] >= table[i][j - 1]) {
      i -= 1;
    } else {
      j -= 1;
    }
  }
  return matches.reverse();
}

function compareSequences<T extends { id: string; signature: string; kind?: string }>(before: T[], after: T[]) {
  const matches = lcsMatches(before, after);
  const output: Array<{ kind: FormulaChangeKind; before?: T; after?: T }> = [];
  let unchanged = 0;
  let beforeStart = 0;
  let afterStart = 0;

  const addGap = (beforeEnd: number, afterEnd: number) => {
    const removed = before.slice(beforeStart, beforeEnd);
    const added = after.slice(afterStart, afterEnd);
    const pairedBefore = new Set<number>();
    const pairedAfter = new Set<number>();
    for (let left = 0; left < removed.length; left += 1) {
      const right = added.findIndex((item, index) => !pairedAfter.has(index) && (!removed[left].kind || item.kind === removed[left].kind));
      if (right >= 0) {
        pairedBefore.add(left);
        pairedAfter.add(right);
        output.push({ kind: 'modified', before: removed[left], after: added[right] });
      }
    }
    removed.forEach((item, index) => { if (!pairedBefore.has(index)) output.push({ kind: 'deleted', before: item }); });
    added.forEach((item, index) => { if (!pairedAfter.has(index)) output.push({ kind: 'added', after: item }); });
  };

  for (const [beforeIndex, afterIndex] of [...matches, [before.length, after.length] as [number, number]]) {
    addGap(beforeIndex, afterIndex);
    if (beforeIndex < before.length && afterIndex < after.length) unchanged += 1;
    beforeStart = beforeIndex + 1;
    afterStart = afterIndex + 1;
  }

  return { output, unchanged };
}

export function compareFormulas(before: FormulaItem[], after: FormulaItem[]): FormulaComparison {
  const { output, unchanged } = compareSequences(before, after);
  return {
    changes: output.map((change, index) => ({ id: `formula-change-${index + 1}`, ...change })),
    unchanged,
    beforeCount: before.length,
    afterCount: after.length,
  };
}

export function compareObjects(before: EmbeddedObject[], after: EmbeddedObject[]): ObjectComparison {
  const { output, unchanged } = compareSequences(before, after);
  return {
    changes: output.map((change, index) => ({ id: `object-change-${index + 1}`, ...change })),
    unchanged,
  };
}
