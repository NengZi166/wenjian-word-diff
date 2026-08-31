'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import katex from 'katex';
import 'katex/dist/katex.min.css';
import {
  analyzeDocx,
  compareFormulas,
  compareObjects,
  type DocxAnalysis,
  type FormulaChange,
  type FormulaComparison,
  type FormulaItem,
  type ObjectChange,
  type ObjectComparison,
} from './lib/docx-analysis';

type ResultTab = 'preview' | 'formula' | 'object' | 'notes';

interface ComparisonResult {
  original: DocxAnalysis;
  modified: DocxAnalysis;
  formulas: FormulaComparison;
  objects: ObjectComparison;
  previewHtml: string;
  previewError: string;
  downloadUrl: string;
  downloadName: string;
  revisionCount: number;
  insertedCount: number;
  deletedCount: number;
  warnings: string[];
}

function formatBytes(size: number): string {
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(0)} KB`;
  return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

function safeBaseName(name: string): string {
  return name.replace(/\.docx$/i, '').replace(/[<>:"/\\|?*]+/g, '_').slice(0, 48);
}

function FilePicker({
  label,
  hint,
  marker,
  file,
  accent,
  onSelect,
}: {
  label: string;
  hint: string;
  marker: string;
  file: File | null;
  accent?: boolean;
  onSelect: (file: File | null) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const acceptFiles = (files: FileList | null) => {
    const selected = files?.[0] ?? null;
    if (selected) onSelect(selected);
  };

  return (
    <div
      className={`drop-card${dragging ? ' dragging' : ''}${file ? ' has-file' : ''}`}
      onDragEnter={(event) => { event.preventDefault(); setDragging(true); }}
      onDragOver={(event) => event.preventDefault()}
      onDragLeave={(event) => {
        if (!event.currentTarget.contains(event.relatedTarget as Node)) setDragging(false);
      }}
      onDrop={(event) => {
        event.preventDefault();
        setDragging(false);
        acceptFiles(event.dataTransfer.files);
      }}
    >
      <span className="step">{marker}</span>
      <span className={`file-icon${accent ? ' accent' : ''}`}>{file ? '✓' : accent ? 'B' : 'A'}</span>
      {file ? (
        <>
          <strong className="file-name" title={file.name}>{file.name}</strong>
          <span className="file-meta">{formatBytes(file.size)} · DOCX</span>
          <div className="file-actions">
            <button type="button" onClick={() => inputRef.current?.click()}>更换</button>
            <button type="button" onClick={() => onSelect(null)}>移除</button>
          </div>
        </>
      ) : (
        <>
          <strong>{label}</strong>
          <span>{hint}</span>
          <button className="select-file" type="button" onClick={() => inputRef.current?.click()}>选择文件</button>
        </>
      )}
      <input
        ref={inputRef}
        className="hidden-input"
        type="file"
        accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        onChange={(event) => acceptFiles(event.target.files)}
      />
    </div>
  );
}

function Formula({ item }: { item?: FormulaItem }) {
  const html = useMemo(() => {
    if (!item) return '';
    try {
      return katex.renderToString(item.latex, {
        displayMode: true,
        throwOnError: false,
        trust: false,
        strict: false,
        output: 'htmlAndMathml',
      });
    } catch {
      return '';
    }
  }, [item]);

  if (!item) return <div className="formula-empty">无</div>;
  return (
    <div className="formula-block">
      <div className="formula-location">{item.location}</div>
      {html ? <div className="formula-render" dangerouslySetInnerHTML={{ __html: html }} /> : <code>{item.text || item.latex}</code>}
    </div>
  );
}

function changeLabel(kind: FormulaChange['kind']): string {
  if (kind === 'modified') return '已修改';
  if (kind === 'added') return '新增';
  return '删除';
}

function FormulaChanges({ comparison }: { comparison: FormulaComparison }) {
  if (!comparison.beforeCount && !comparison.afterCount) {
    return (
      <div className="empty-state">
        <span className="empty-symbol">∅</span>
        <h3>没有检测到 Office 原生公式</h3>
        <p>如果文档中看得到公式，它可能是 MathType/OLE 或图片，请查看“对象检查”。</p>
      </div>
    );
  }
  if (!comparison.changes.length) {
    return (
      <div className="empty-state success">
        <span className="empty-symbol">✓</span>
        <h3>{comparison.unchanged} 个原生公式均未变化</h3>
        <p>公式结构与内容的规范化结果一致。</p>
      </div>
    );
  }
  return (
    <div className="change-list">
      {comparison.changes.map((change, index) => (
        <article className={`change-item ${change.kind}`} key={change.id}>
          <header>
            <span className="change-number">{String(index + 1).padStart(2, '0')}</span>
            <strong>公式{changeLabel(change.kind)}</strong>
            <span className={`change-badge ${change.kind}`}>{changeLabel(change.kind)}</span>
          </header>
          <div className="formula-compare">
            <section>
              <small>原始版本</small>
              <Formula item={change.before} />
            </section>
            <div className="formula-arrow">→</div>
            <section>
              <small>修改版本</small>
              <Formula item={change.after} />
            </section>
          </div>
        </article>
      ))}
    </div>
  );
}

function objectName(item: ObjectChange['before'] | ObjectChange['after']): string {
  if (!item) return '无';
  return `${item.label}${item.programId ? ` · ${item.programId}` : ''}`;
}

function ObjectChanges({ comparison }: { comparison: ObjectComparison }) {
  if (!comparison.changes.length) {
    return (
      <div className="empty-state success">
        <span className="empty-symbol">✓</span>
        <h3>嵌入对象未发现变化</h3>
        <p>共匹配 {comparison.unchanged} 个图片或 OLE 对象。</p>
      </div>
    );
  }
  return (
    <div className="object-list">
      {comparison.changes.map((change, index) => {
        const item = change.after ?? change.before;
        return (
          <article className="object-item" key={change.id}>
            <span className={`object-icon ${item?.kind ?? 'image'}`}>{item?.kind === 'ole' ? 'ƒx' : '图'}</span>
            <div>
              <div className="object-title">
                <strong>{objectName(item)}</strong>
                <span className={`change-badge ${change.kind}`}>{changeLabel(change.kind)}</span>
              </div>
              <p>{change.before?.location ?? change.after?.location}</p>
              <small>{item?.formulaLike ? '公式相关对象 · 采用文件指纹比对' : '普通嵌入对象 · 采用文件指纹比对'}</small>
            </div>
          </article>
        );
      })}
    </div>
  );
}

function ResultPanel({ result }: { result: ComparisonResult }) {
  const [tab, setTab] = useState<ResultTab>('preview');
  const tabs: Array<{ id: ResultTab; label: string; count?: number }> = [
    { id: 'preview', label: '文档差异' },
    { id: 'formula', label: '公式变化', count: result.formulas.changes.length },
    { id: 'object', label: '对象检查', count: result.objects.changes.length },
    { id: 'notes', label: '支持说明' },
  ];

  return (
    <section className="result-section" aria-live="polite">
      <div className="result-heading">
        <div>
          <span className="result-kicker">比对完成</span>
          <h2>发现 <em>{result.revisionCount}</em> 处 Word 修订</h2>
          <p>其中原生公式变化 {result.formulas.changes.length} 处，嵌入对象变化 {result.objects.changes.length} 处。</p>
        </div>
        <a className="download-button" href={result.downloadUrl} download={result.downloadName}>
          <span>↓</span>
          <span><b>下载修订版 Word</b><small>{result.downloadName}</small></span>
        </a>
      </div>

      <div className="stats-grid">
        <div><span className="stat-dot insert" /><b>{result.insertedCount}</b><small>插入修订</small></div>
        <div><span className="stat-dot delete" /><b>{result.deletedCount}</b><small>删除修订</small></div>
        <div><span className="stat-dot formula" /><b>{result.original.formulas.length} → {result.modified.formulas.length}</b><small>原生公式数量</small></div>
        <div><span className="stat-dot object" /><b>{result.objects.changes.length}</b><small>对象变化</small></div>
      </div>

      {result.warnings.length > 0 && (
        <div className="warning-box">
          <b>注意</b>
          <div>{result.warnings.map((warning) => <p key={warning}>{warning}</p>)}</div>
        </div>
      )}

      <div className="result-workspace">
        <nav className="result-tabs" aria-label="结果视图">
          {tabs.map((item) => (
            <button className={tab === item.id ? 'active' : ''} type="button" key={item.id} onClick={() => setTab(item.id)}>
              {item.label}{typeof item.count === 'number' && <span>{item.count}</span>}
            </button>
          ))}
        </nav>
        <div className="result-body">
          {tab === 'preview' && (
            result.previewHtml ? (
              <iframe className="document-preview" title="Word 文档差异预览" sandbox="" srcDoc={result.previewHtml} />
            ) : (
              <div className="empty-state"><span className="empty-symbol">!</span><h3>网页预览未能生成</h3><p>{result.previewError}。修订版 Word 仍可正常下载查看。</p></div>
            )
          )}
          {tab === 'formula' && <FormulaChanges comparison={result.formulas} />}
          {tab === 'object' && <ObjectChanges comparison={result.objects} />}
          {tab === 'notes' && (
            <div className="support-grid">
              <article><span>01</span><h3>Office 原生公式</h3><p>解析 OMML 结构，可对分式、上下标、根式、积分、求和、矩阵等常见结构进行网页排版和变化定位。</p></article>
              <article><span>02</span><h3>MathType / OLE</h3><p>作为嵌入对象比较二进制指纹，可以判断新增、删除或替换；暂不解析对象内部的单个公式符号。</p></article>
              <article><span>03</span><h3>图片公式</h3><p>可判断图片是否变化。由于本工具不上传文件，也不调用云端 OCR，因此不能判断图片内部哪一个符号变化。</p></article>
              <article><span>04</span><h3>隐私与内网</h3><p>解析、比较与生成结果全部发生在浏览器内存。关闭或刷新页面后，所选文件和结果会被清除。</p></article>
            </div>
          )}
        </div>
      </div>
    </section>
  );
}

export default function Home() {
  const [originalFile, setOriginalFile] = useState<File | null>(null);
  const [modifiedFile, setModifiedFile] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [phase, setPhase] = useState('');
  const [error, setError] = useState('');
  const [result, setResult] = useState<ComparisonResult | null>(null);

  useEffect(() => () => {
    if (result?.downloadUrl) URL.revokeObjectURL(result.downloadUrl);
  }, [result]);

  const resetResult = () => {
    setError('');
    setResult((current) => {
      if (current?.downloadUrl) URL.revokeObjectURL(current.downloadUrl);
      return null;
    });
  };

  const setFile = (side: 'original' | 'modified', file: File | null) => {
    resetResult();
    if (side === 'original') setOriginalFile(file);
    else setModifiedFile(file);
  };

  const swapFiles = () => {
    resetResult();
    setOriginalFile(modifiedFile);
    setModifiedFile(originalFile);
  };

  const runComparison = async () => {
    if (!originalFile || !modifiedFile || busy) return;
    setBusy(true);
    setError('');
    resetResult();
    try {
      setPhase('正在读取文档结构与公式…');
      const [original, modified, engine] = await Promise.all([
        analyzeDocx(originalFile),
        analyzeDocx(modifiedFile),
        import('docxodus'),
      ]);
      const formulas = compareFormulas(original.formulas, modified.formulas);
      const objects = compareObjects(original.objects, modified.objects);

      setPhase('正在启动本地 Word 比对引擎…');
      await engine.initialize('/vendor/docxodus/wasm/');
      setPhase('正在生成带修订痕迹的 Word…');
      const redline = await engine.compareDocuments(originalFile, modifiedFile, {
        authorName: '文鉴',
        detailThreshold: 0,
        caseInsensitive: false,
        engine: engine.ComparisonEngine.WmlComparer,
      });

      setPhase('正在整理差异与网页预览…');
      const revisions = await engine.getRevisions(redline, { detectMoves: false });
      let previewHtml = '';
      let previewError = '';
      try {
        previewHtml = await engine.convertDocxToHtml(redline, {
          pageTitle: '文鉴 · 文档差异预览',
          renderTrackedChanges: true,
          showDeletedContent: true,
          renderMoveOperations: false,
          renderUnsupportedContentPlaceholders: true,
          renderHeadersAndFooters: true,
          renderFootnotesAndEndnotes: true,
          documentLanguage: 'zh-CN',
          additionalCss: 'body{background:#f5f3ed!important;padding:24px!important} ins{background:#dff4e7!important;text-decoration:none!important} del{background:#ffe8df!important;color:#8f3f25!important}',
        });
      } catch (previewFailure) {
        previewError = previewFailure instanceof Error ? previewFailure.message : '未知预览错误';
      }

      const revisionKinds = revisions.map((revision) => String(revision.revisionType).toLowerCase());
      const insertedCount = revisionKinds.filter((kind) => kind.includes('insert')).length;
      const deletedCount = revisionKinds.filter((kind) => kind.includes('delet')).length;
      const downloadName = `${safeBaseName(originalFile.name)}_与_${safeBaseName(modifiedFile.name)}_对比结果.docx`;
      const redlineBuffer = new Uint8Array(redline).buffer;
      const downloadUrl = URL.createObjectURL(new Blob([redlineBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }));
      setResult({
        original,
        modified,
        formulas,
        objects,
        previewHtml,
        previewError,
        downloadUrl,
        downloadName,
        revisionCount: revisions.length,
        insertedCount,
        deletedCount,
        warnings: Array.from(new Set([...original.warnings, ...modified.warnings])),
      });
      setTimeout(() => document.querySelector('.result-section')?.scrollIntoView({ behavior: 'smooth', block: 'start' }), 80);
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : '比对失败，请确认文件可正常用 Word 打开。');
    } finally {
      setBusy(false);
      setPhase('');
    }
  };

  return (
    <main className="app-shell">
      <header className="topbar">
        <a className="brand" href="#top" aria-label="文鉴首页">
          <span className="brand-mark">文</span>
          <span><strong>文鉴</strong><small>Word 文档比对</small></span>
        </a>
        <div className="privacy-pill"><span /> 文件仅在本机浏览器中处理</div>
      </header>

      <section className="hero" id="top">
        <div className="eyebrow">离线 · 免费 · 公式友好</div>
        <h1>两份 Word，<em>差异一目了然。</em></h1>
        <p>支持文字、表格与 Office 原生公式比对。无需上传，无需登录，内网环境也能使用。</p>
      </section>

      <section className="compare-card" aria-label="选择待比对文档">
        <div className="file-grid">
          <FilePicker marker="01" label="原始版本" hint="拖入或选择原始 .docx 文件" file={originalFile} onSelect={(file) => setFile('original', file)} />
          <button className="swap-files" type="button" aria-label="交换两份文档" title="交换两份文档" onClick={swapFiles} disabled={!originalFile && !modifiedFile}>⇄</button>
          <FilePicker marker="02" label="修改版本" hint="拖入或选择修改后的 .docx" file={modifiedFile} accent onSelect={(file) => setFile('modified', file)} />
        </div>

        <div className="action-row">
          <div className="support-note"><b>可识别</b> 文字 · 格式 · 表格 · Office 公式 · 图片/OLE 对象变化</div>
          <button className="compare-button" type="button" disabled={!originalFile || !modifiedFile || busy} onClick={runComparison}>
            {busy ? '正在比对…' : '开始比对'} <span>↗</span>
          </button>
        </div>
        {error && <div className="error-box" role="alert"><b>比对没有完成</b><span>{error}</span></div>}
      </section>

      {!result && (
        <section className="feature-strip" aria-label="产品特点">
          <article><b>本地处理</b><span>文件不离开电脑</span></article>
          <article><b>公式识别</b><span>支持常见 OMML 结构</span></article>
          <article><b>Word 修订版</b><span>结果可接受或拒绝</span></article>
        </section>
      )}

      {busy && (
        <div className="busy-overlay" role="status" aria-live="polite">
          <div className="busy-dialog">
            <span className="busy-spinner" />
            <div><b>{phase}</b><small>所有处理都在这台电脑上完成，请不要关闭页面。</small></div>
          </div>
        </div>
      )}

      {result && <ResultPanel result={result} />}

      <footer className="footer-note">
        <span>文鉴 v0.1</span><span>·</span><span>开源离线工具</span><span>·</span><span>单个文件上限 100 MB</span>
      </footer>
    </main>
  );
}
