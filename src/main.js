import './style.css';
import { marked } from 'marked';

const SUPPORTED_EXTS = ['pdf', 'xlsx', 'docx', 'pptx'];
const FILE_ACCEPT = '.pdf,.xlsx,.docx,.pptx';

const NO_SUMMARIZER_MSG =
  'Chrome’s on-device Summarizer API is not available in this browser. Open this page in a recent version of Google Chrome to generate document summaries.';
const SUMMARIZER_UNUSABLE_MSG =
  'The on-device Summarizer API is not usable in this browser or on this device. Try a recent Google Chrome build with on-device AI enabled.';
const SUMMARIZER_TOO_LARGE_MSG =
  'Unfortunately this document was too large to summarize.';
const SUMMARIZER_FAILED_MSG =
  'Could not generate a summary for this document.';

const NO_PROMPT_MSG =
  'Chrome’s on-device Prompt API is not available in this browser. Open this page in a recent version of Google Chrome to ask questions about documents.';
const PROMPT_UNUSABLE_MSG =
  'The on-device Prompt API is not usable in this browser or on this device. Try a recent Google Chrome build with on-device AI enabled.';
const PROMPT_FAILED_MSG = 'Could not answer that question.';
const PROMPT_SYSTEM_INSTRUCTION =
  'You are a helpful document assistant. Answer questions using only the document the user provides. If the document does not contain the answer, say that clearly. Be concise and accurate.';

const fileInput = document.querySelector('#file-input');
const fileInputEmpty = document.querySelector('#file-input-empty');
const docEmpty = document.querySelector('#doc-empty');
const docFilename = document.querySelector('#doc-filename');
const documentScroll = document.querySelector('#document-scroll');
const documentView = document.querySelector('#document-view');
const summaryEmpty = document.querySelector('#summary-empty');
const qaThread = document.querySelector('#qa-thread');
const qaForm = document.querySelector('#qa-form');
const qaInput = document.querySelector('#qa-input');
const qaSubmit = document.querySelector('#qa-submit');

let summaryEl = document.querySelector('#summary');
let currentFile = null;
let currentAst = null;
let currentDocMarkdown = '';
let answering = false;
let qaReady = false;
let loadToken = 0;
let officeParserPromise = null;

/** Core Prompt API session with the system instruction (cloned per document). */
let basePromptSession = null;
/** Per-document cloned session with markdown context appended. */
let docPromptSession = null;

fileInput.accept = FILE_ACCEPT;
fileInputEmpty.accept = FILE_ACCEPT;

marked.setOptions({ breaks: true });

function loadOfficeParser() {
  if (!officeParserPromise) {
    officeParserPromise = import('officeparser');
  }
  return officeParserPromise;
}

async function getAST(file, config = {}) {
  const { OfficeParser } = await loadOfficeParser();
  return OfficeParser.parseOffice(file, config);
}

async function parseDocument(file) {
  const arrayBuffer = await file.arrayBuffer();
  const ast = await getAST(arrayBuffer, {});
  currentAst = ast;

  const [html, markdown] = await Promise.all([
    ast.to('html', {
      htmlConfig: {
        standalone: { document: false, styles: 'scoped' },
      },
    }),
    ast.to('markdown'),
  ]);

  currentDocMarkdown = markdown.value ?? '';
  return html.value;
}

function renderDocumentHtml(html) {
  documentView.innerHTML = html;
}

function destroyDocPromptSession() {
  if (!docPromptSession) return;
  try {
    docPromptSession.destroy();
  } catch {
    // Session may already be destroyed.
  }
  docPromptSession = null;
}

function showQaStatus(message, { isError = false } = {}) {
  qaThread.replaceChildren();
  const hint = document.createElement('p');
  hint.id = 'qa-empty';
  hint.className = `qa-hint${isError ? ' qa-hint-error' : ''}`;
  hint.textContent = message;
  qaThread.appendChild(hint);
}

function clearQaThread() {
  showQaStatus('Ask a question about this document.');
}

function setQaEnabled(enabled) {
  const on = Boolean(enabled) && qaReady && !answering;
  qaInput.disabled = !on;
  qaSubmit.disabled = !on;
}

function appendMessage(role, text, { pending = false, markdown = false } = {}) {
  const hint = qaThread.querySelector('#qa-empty');
  if (hint) hint.remove();

  const msg = document.createElement('div');
  msg.className = `qa-msg qa-msg-${role}${pending ? ' pending' : ''}`;
  if (markdown) {
    msg.innerHTML = marked.parse(text);
  } else {
    msg.textContent = text;
  }
  qaThread.appendChild(msg);
  qaThread.scrollTop = qaThread.scrollHeight;
  return msg;
}

function renderAiMarkdown(el, markdown) {
  el.classList.remove('pending');
  el.innerHTML = marked.parse(markdown);
}

function setVisible(el, visible) {
  el.classList.toggle('is-hidden', !visible);
  el.hidden = !visible;
  el.setAttribute('aria-hidden', visible ? 'false' : 'true');
}

function ensureSummaryPanel() {
  setVisible(summaryEmpty, false);
  if (!summaryEl.isConnected) {
    summaryEl = document.querySelector('#summary') ?? summaryEl;
  }
  setVisible(summaryEl, true);
  return summaryEl;
}

function showSummaryStatus(message, { isError = false } = {}) {
  const el = ensureSummaryPanel();
  el.className = `summary-body summary-status${isError ? ' summary-status-error' : ''}`;
  el.replaceChildren();
  const p = document.createElement('p');
  p.textContent = message;
  el.appendChild(p);
}

function showSummaryHtml(html) {
  const el = ensureSummaryPanel();
  el.className = 'summary-body';
  el.innerHTML = html;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function sharedContextFor(ext) {
  if (ext === 'pptx') {
    return 'This is extracted text from a PowerPoint file.';
  }
  if (ext === 'xlsx') {
    return 'This is extracted text from an Excel file.';
  }
  if (ext === 'pdf') {
    return 'This is extracted text from a PDF document.';
  }
  if (ext === 'docx') {
    return 'This is extracted text from a Word document.';
  }
  return null;
}

function isGestureOrActivationError(err) {
  const message = String(err?.message || err || '').toLowerCase();
  return (
    err?.name === 'NotAllowedError' ||
    err?.name === 'InvalidStateError' ||
    message.includes('user gesture') ||
    message.includes('user activation')
  );
}

function isBogusModelResult(result, sourceText = '') {
  const normalized = String(result ?? '').trim();
  const lower = normalized.toLowerCase();
  if (
    lower.startsWith('model not available') ||
    lower.includes('model not available in chromium')
  ) {
    return true;
  }

  const source = sourceText.trim();
  if (!source) return false;
  if (normalized === source) return true;
  const sampleLen = Math.min(120, source.length);
  if (sampleLen >= 40 && normalized.includes(source.slice(0, sampleLen))) {
    return true;
  }
  return false;
}

/**
 * Create the summarizer during the file-input user gesture.
 * Chrome requires activation when the model is still downloadable.
 */
async function prepareSummarizer(file, token) {
  if (!('Summarizer' in self)) {
    return { error: 'no-api' };
  }

  let availability = 'unavailable';
  try {
    availability = await Summarizer.availability();
  } catch (err) {
    console.error(err);
    return { error: 'no-api' };
  }

  if (availability === 'unavailable') {
    return { error: 'unavailable' };
  }

  if (token !== loadToken) return { error: 'cancelled' };

  showSummaryStatus(
    availability === 'downloadable' || availability === 'downloading'
      ? 'Preparing the Summary model…'
      : 'Summary model ready. Waiting for document parse…',
  );

  try {
    const summarizer = await Summarizer.create({
      type: 'tldr',
      length: 'long',
      format: 'markdown',
      sharedContext: sharedContextFor(fileExtension(file)),
      monitor(m) {
        m.addEventListener('downloadprogress', (e) => {
          if (e.loaded === 0 || e.loaded === 1) return;
          if (token !== loadToken) return;
          showSummaryStatus(
            `Downloading the Summary model, currently at ${Math.floor(e.loaded * 100)}%`,
          );
        });
      },
    });
    return { summarizer };
  } catch (err) {
    console.error(err);
    return { error: 'create-failed', detail: err };
  }
}

async function runSummary(summarizer, file, token) {
  if (token !== loadToken) return;

  showSummaryStatus('Working on the summary…');

  const title = currentAst?.metadata?.title || file.name || 'No Title';
  const text = currentDocMarkdown;

  if (!text?.trim()) {
    showSummaryStatus('No text was found in this document to summarize.', { isError: true });
    return;
  }

  try {
    const summary = await summarizer.summarize(text);
    if (token !== loadToken) return;
    if (typeof summary !== 'string' || !summary.trim()) {
      showSummaryStatus(SUMMARIZER_FAILED_MSG, { isError: true });
      return;
    }
    if (isBogusModelResult(summary, text)) {
      console.warn('Ignoring non-functional Summarizer result:', summary.slice(0, 120));
      showSummaryStatus(SUMMARIZER_UNUSABLE_MSG, { isError: true });
      return;
    }
    showSummaryHtml(
      `<h3>Summary for ${escapeHtml(title)}</h3>${marked.parse(summary)}`,
    );
  } catch (err) {
    if (token !== loadToken) return;
    console.error(err);
    if (err?.name === 'QuotaExceededError') {
      showSummaryStatus(SUMMARIZER_TOO_LARGE_MSG, { isError: true });
    } else {
      showSummaryStatus(SUMMARIZER_FAILED_MSG, { isError: true });
    }
  }
}

function showSummarizerPrepError(result) {
  if (result.error === 'no-api') {
    showSummaryStatus(NO_SUMMARIZER_MSG, { isError: true });
    return;
  }
  if (result.error === 'unavailable') {
    showSummaryStatus(SUMMARIZER_UNUSABLE_MSG, { isError: true });
    return;
  }
  if (result.error === 'create-failed') {
    console.error('Summarizer.create failed:', result.detail);
    if (isGestureOrActivationError(result.detail)) {
      showSummaryStatus(SUMMARIZER_UNUSABLE_MSG, { isError: true });
      return;
    }
    showSummaryStatus(SUMMARIZER_FAILED_MSG, { isError: true });
  }
}

const promptCreateOptions = {
  initialPrompts: [{ role: 'system', content: PROMPT_SYSTEM_INSTRUCTION }],
  expectedInputs: [{ type: 'text', languages: ['en'] }],
  expectedOutputs: [{ type: 'text', languages: ['en'] }],
};

/**
 * Ensure the core Prompt API session exists (system instruction only).
 * Created during the file-input gesture so model download is allowed.
 */
async function ensureBasePromptSession(token) {
  if (basePromptSession) {
    return { session: basePromptSession };
  }

  if (!('LanguageModel' in self)) {
    return { error: 'no-api' };
  }

  let availability = 'unavailable';
  try {
    availability = await LanguageModel.availability(promptCreateOptions);
  } catch (err) {
    console.error(err);
    return { error: 'no-api' };
  }

  if (availability === 'unavailable') {
    return { error: 'unavailable' };
  }

  if (token !== loadToken) return { error: 'cancelled' };

  if (availability === 'downloadable' || availability === 'downloading') {
    showQaStatus('Preparing the on-device language model…');
  } else {
    showQaStatus('Language model ready. Waiting for document parse…');
  }

  try {
    basePromptSession = await LanguageModel.create({
      ...promptCreateOptions,
      monitor(m) {
        m.addEventListener('downloadprogress', (e) => {
          if (e.loaded === 0 || e.loaded === 1) return;
          if (token !== loadToken) return;
          showQaStatus(
            `Downloading the language model, currently at ${Math.floor(e.loaded * 100)}%`,
          );
        });
      },
    });
    return { session: basePromptSession };
  } catch (err) {
    console.error(err);
    return { error: 'create-failed', detail: err };
  }
}

function showPromptPrepError(result) {
  qaReady = false;
  setQaEnabled(false);
  if (result.error === 'no-api') {
    showQaStatus(NO_PROMPT_MSG, { isError: true });
    return;
  }
  if (result.error === 'unavailable') {
    showQaStatus(PROMPT_UNUSABLE_MSG, { isError: true });
    return;
  }
  if (result.error === 'create-failed') {
    console.error('LanguageModel.create failed:', result.detail);
    if (isGestureOrActivationError(result.detail)) {
      showQaStatus(PROMPT_UNUSABLE_MSG, { isError: true });
      return;
    }
    showQaStatus(PROMPT_FAILED_MSG, { isError: true });
  }
}

/**
 * Clone the core session and append this document's markdown as context.
 * Ask stays disabled until this completes successfully.
 */
async function prepareDocPromptSession(markdown, token) {
  qaReady = false;
  setQaEnabled(false);
  destroyDocPromptSession();

  const base = await ensureBasePromptSession(token);
  if (token !== loadToken) return { error: 'cancelled' };
  if (base.error) return base;

  if (!markdown?.trim()) {
    return { error: 'no-text' };
  }

  showQaStatus('Loading document into the Q&A session…');

  try {
    const cloned = await base.session.clone();
    if (token !== loadToken) {
      try {
        cloned.destroy();
      } catch {
        // ignore
      }
      return { error: 'cancelled' };
    }

    await cloned.append([
      {
        role: 'user',
        content:
          'Here is the document in Markdown. Use it as the sole source of truth when answering questions.\n\n' +
          markdown,
      },
    ]);

    if (token !== loadToken) {
      try {
        cloned.destroy();
      } catch {
        // ignore
      }
      return { error: 'cancelled' };
    }

    docPromptSession = cloned;
    return { session: docPromptSession };
  } catch (err) {
    console.error(err);
    if (isGestureOrActivationError(err)) {
      return { error: 'create-failed', detail: err };
    }
    return { error: 'create-failed', detail: err };
  }
}

async function answerQuestion(question) {
  if (!docPromptSession) {
    throw new Error('No document Q&A session');
  }
  const result = await docPromptSession.prompt(question);
  if (typeof result !== 'string' || !result.trim()) {
    throw new Error('Empty prompt result');
  }
  if (isBogusModelResult(result, currentDocMarkdown)) {
    const err = new Error('Prompt API stubbed or unavailable');
    err.code = 'unusable';
    throw err;
  }
  return result;
}

function showDocStatus(message, { isError = false } = {}) {
  setVisible(docEmpty, false);
  setVisible(documentScroll, true);
  const p = document.createElement('p');
  p.className = `doc-status${isError ? ' doc-status-error' : ''}`;
  p.textContent = message;
  documentView.replaceChildren(p);
}

function fileExtension(file) {
  const name = file?.name ?? '';
  const parts = name.split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : '';
}

async function loadDocument(file) {
  if (!file) return;

  const ext = fileExtension(file);
  if (!SUPPORTED_EXTS.includes(ext)) {
    currentFile = null;
    currentAst = null;
    currentDocMarkdown = '';
    qaReady = false;
    destroyDocPromptSession();
    setQaEnabled(false);
    setVisible(docFilename, true);
    docFilename.textContent = file.name;
    docFilename.title = file.name;
    showDocStatus(
      'Unsupported file type. Please choose a PDF, DOCX, XLSX, or PPTX file.',
      { isError: true },
    );
    setVisible(summaryEmpty, true);
    setVisible(summaryEl, false);
    showQaStatus('Load a supported document, then ask a question.');
    return;
  }

  const token = ++loadToken;
  currentFile = file;
  currentAst = null;
  currentDocMarkdown = '';
  answering = false;
  qaReady = false;
  destroyDocPromptSession();
  setQaEnabled(false);

  setVisible(docEmpty, false);
  setVisible(documentScroll, true);
  setVisible(docFilename, true);
  docFilename.textContent = file.name;
  docFilename.title = file.name;
  showDocStatus('Parsing document…');
  showQaStatus('Parsing document…');

  // Start model setup while the file-input gesture is still active.
  const summarizerPromise = prepareSummarizer(file, token);
  const promptBasePromise = ensureBasePromptSession(token);

  try {
    const html = await parseDocument(file);
    if (token !== loadToken) return;

    renderDocumentHtml(html);

    const [summarizerPrep, promptBase] = await Promise.all([
      summarizerPromise,
      promptBasePromise,
    ]);
    if (token !== loadToken) return;

    // Build a fresh cloned Q&A session with this document's markdown.
    if (promptBase.error) {
      if (promptBase.error !== 'cancelled') showPromptPrepError(promptBase);
    } else {
      const docSession = await prepareDocPromptSession(currentDocMarkdown, token);
      if (token !== loadToken) return;
      if (docSession.error) {
        if (docSession.error === 'no-text') {
          qaReady = false;
          setQaEnabled(false);
          showQaStatus('No text was found in this document for Q&A.', { isError: true });
        } else if (docSession.error !== 'cancelled') {
          showPromptPrepError(docSession);
        }
      } else {
        qaReady = true;
        showQaStatus('Ask a question about this document.');
        setQaEnabled(true);
        qaInput.focus();
      }
    }

    if (summarizerPrep.error) {
      if (summarizerPrep.error !== 'cancelled') showSummarizerPrepError(summarizerPrep);
    } else {
      await runSummary(summarizerPrep.summarizer, file, token);
    }
  } catch (err) {
    if (token !== loadToken) return;
    console.error(err);
    currentAst = null;
    currentDocMarkdown = '';
    qaReady = false;
    destroyDocPromptSession();
    showDocStatus('Could not parse that document. Try another file.', { isError: true });
    setVisible(summaryEmpty, true);
    setVisible(summaryEl, false);
    showQaStatus('Load a document, then ask a question.');
    setQaEnabled(false);
  }
}

function onFileChosen(event) {
  const [file] = event.target.files ?? [];
  if (!file) return;
  loadDocument(file);
  if (event.target !== fileInput) fileInput.value = '';
  if (event.target !== fileInputEmpty) fileInputEmpty.value = '';
}

fileInput.addEventListener('change', onFileChosen);
fileInputEmpty.addEventListener('change', onFileChosen);

qaForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  if (!currentFile || !qaReady || answering || !docPromptSession) return;

  const question = qaInput.value.trim();
  if (!question) return;

  appendMessage('user', question);
  qaInput.value = '';
  answering = true;
  setQaEnabled(false);

  const pending = appendMessage('ai', 'Thinking…', { pending: true });

  try {
    const answer = await answerQuestion(question);
    renderAiMarkdown(pending, answer);
  } catch (err) {
    console.error(err);
    pending.classList.remove('pending');
    if (err?.code === 'unusable') {
      pending.textContent = PROMPT_UNUSABLE_MSG;
      qaReady = false;
    } else if (err?.name === 'QuotaExceededError') {
      pending.textContent = 'This question or document is too large for the model context.';
    } else {
      pending.textContent = PROMPT_FAILED_MSG;
    }
  } finally {
    answering = false;
    setQaEnabled(true);
    if (qaReady) qaInput.focus();
    qaThread.scrollTop = qaThread.scrollHeight;
  }
});
