document.addEventListener('DOMContentLoaded', () => {
    const hostName = window.location.hostname || 'localhost';
    const configuredApiBase = window.TRANSLATION_CONFIG?.API_BASE_URL?.trim();
    const localHost = hostName === 'localhost' || hostName === '127.0.0.1';
    const apiBase = (localHost ? `http://${hostName}:8000` : configuredApiBase).replace(/\/+$/, '');
    const ngrokHeaders = { 'ngrok-skip-browser-warning': 'true' };

    const translationForm = document.getElementById('translationForm');
    const statusBox = document.getElementById('statusBox');
    const submitBtn = document.getElementById('submitBtn');
    const useGlossaryToggle = document.getElementById('useGlossaryToggle');
    const glossaryUploadGroup = document.getElementById('glossaryUploadGroup');
    const glossaryFileInput = document.getElementById('glossaryFile');
    const docFileInput = document.getElementById('documentFile');
    const targetLangSelect = document.getElementById('targetLanguage');
    const modelNameSelect = document.getElementById('modelSelect');
    const reviewContainer = document.getElementById('reviewContainer');
    const progressPanel = document.getElementById('progressPanel');
    const progressBar = document.getElementById('progressBar');
    const progressLabel = document.getElementById('progressLabel');
    const progressValue = document.getElementById('progressValue');
    const translationView = document.getElementById('translationView');
    const editorView = document.getElementById('editorView');
    const editorShell = document.querySelector('.editor-shell');
    const documentEditor = document.getElementById('documentEditor');
    const editorTitle = document.getElementById('editorTitle');
    const editorStatus = document.getElementById('editorStatus');
    const exportBtn = document.getElementById('exportBtn');
    const previewBtn = document.getElementById('previewBtn');
    const previewModal = document.getElementById('previewModal');
    const previewPdfViewer = document.getElementById('previewPdfViewer');
    const sourcePdfViewer = document.getElementById('sourcePdfViewer');
    const importEditorFileBtn = document.getElementById('importEditorFileBtn');
    const editorFileInput = document.getElementById('editorFileInput');
    const quickImportEditorBtn = document.getElementById('quickImportEditorBtn');
    // Keep the comparison/editor flow available for a future rollout.
    const ENABLE_COMPARISON_UI = false;
    const quickEditorFileInput = document.getElementById('quickEditorFileInput');
    // const translationMemoryFileInput = document.getElementById('translationMemoryFile');
    // const translationMemoryFileName = document.getElementById('translationMemoryFileName');
    let translatedFilename = 'Translated_Document.docx';
    let translatedBlob = null;
    let sourcePdfBlob = null;
    let translatedPdfBlob = null;
    let editorDirty = false;
    let previewRequestId = 0;
    let previewObjectUrl = null;
    let previewPdfDocument = null;
    let pdfjsPromise = null;
    let pdfPageTexts = [];
    let activePreviewPage = 0;
    const ABBREVIATIONS_EDITABLE = true;
    const MAX_DOCUMENT_SIZE_BYTES = 50 * 1024 * 1024;

    function loadPdfJs() {
        if (!pdfjsPromise) {
            pdfjsPromise = import('./vendor/pdfjs/pdf.min.mjs').then((pdfjs) => {
                pdfjs.GlobalWorkerOptions.workerSrc = './vendor/pdfjs/pdf.worker.min.mjs';
                return pdfjs;
            });
        }
        return pdfjsPromise;
    }

    async function renderPdfPreview(previewBlob, viewer = previewPdfViewer) {
        const previousScrollTop = viewer.scrollTop;
        const pdfjs = await loadPdfJs();
        const bytes = new Uint8Array(await previewBlob.arrayBuffer());
        viewer.replaceChildren();
        activePreviewPage = -1;
        const loadingMessage = document.createElement('p');
        loadingMessage.className = 'preview-pdf-loading';
        loadingMessage.textContent = 'Đang render PDF...';
        viewer.appendChild(loadingMessage);
        try {
            previewPdfDocument = await pdfjs.getDocument({ data: bytes }).promise;
        } catch (workerError) {
            previewPdfDocument = await pdfjs.getDocument({ data: bytes, disableWorker: true }).promise;
        }
        viewer.replaceChildren();

        const pageTexts = await Promise.all(
            Array.from({ length: previewPdfDocument.numPages }, async (_, index) => {
                const page = await previewPdfDocument.getPage(index + 1);
                const content = await page.getTextContent();
                return content.items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim().toLowerCase();
            })
        );
        let pageCount = pageTexts.length;
        while (pageCount > 1 && !pageTexts[pageCount - 1]) pageCount -= 1;

        const viewerWidth = Math.max(viewer.clientWidth - 24, 320);
        for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
            const page = await previewPdfDocument.getPage(pageNumber);
            const baseViewport = page.getViewport({ scale: 1 });
            const scale = viewerWidth / baseViewport.width;
            const viewport = page.getViewport({ scale });
            const pageElement = document.createElement('div');
            pageElement.className = 'preview-pdf-page';
            pageElement.dataset.pageNumber = String(pageNumber);
            const canvas = document.createElement('canvas');
            const outputScale = 1;
            canvas.width = Math.ceil(viewport.width);
            canvas.height = Math.ceil(viewport.height);
            canvas.setAttribute('aria-label', `PDF page ${pageNumber}`);
            pageElement.appendChild(canvas);
            viewer.appendChild(pageElement);
            await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        }
        pdfPageTexts = pageTexts.slice(0, pageCount);
        viewer.scrollTop = Math.min(previousScrollTop, viewer.scrollHeight);
    }

    function normalizeSearchText(value) {
        return (value || '').replace(/\s+/g, ' ').trim().toLowerCase();
    }

    function syncPreviewToSection(block) {
        if (!previewPdfViewer || !pdfPageTexts.length) return;
        const text = normalizeSearchText(block.textContent);
        if (text.length < 8) return;
        const words = [...new Set(text.split(' ').filter((word) => word.length > 3))].slice(0, 8);
        const minimumMatches = Math.min(3, words.length);
        const pageIndex = pdfPageTexts.findIndex((pageText) =>
            words.filter((word) => pageText.includes(word)).length >= minimumMatches);
        if (pageIndex < 0) return;
        if (pageIndex === activePreviewPage) return;
        activePreviewPage = pageIndex;
        const pageElement = previewPdfViewer.querySelector(`[data-page-number="${pageIndex + 1}"]`);
        pageElement?.scrollIntoView({ behavior: 'smooth', block: 'start', inline: 'nearest' });
    }

    function setActiveNavigationItem(items, activeItem) {
        items.forEach((item) => item.classList.toggle('is-active', item === activeItem));
    }

    function focusEditorSection(block, shouldFocus = true) {
        const documentStage = document.querySelector('.document-stage');
        documentEditor.querySelectorAll('.navigation-focus-target')
            .forEach((node) => node.classList.remove('navigation-focus-target'));
        block.classList.add('navigation-focus-target');
        if (documentStage) {
            const stageRect = documentStage.getBoundingClientRect();
            const blockRect = block.getBoundingClientRect();
            const blockTop = documentStage.scrollTop
                + (blockRect.top - stageRect.top)
                - (documentStage.clientHeight - blockRect.height) / 2;
            documentStage.scrollTo({ top: Math.max(0, blockTop), behavior: 'smooth' });
        } else {
            block.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
        if (shouldFocus && block.isContentEditable) block.focus({ preventScroll: true });
    }

    function navigateToSection(block, navigationItems, item, shouldFocus = false) {
        focusEditorSection(block, shouldFocus);
        setActiveNavigationItem(navigationItems, item);
        syncPreviewToSection(block);
    }

    function renderReviewStructure(data) {
        if (!reviewContainer || !data || !Array.isArray(data.components)) {
            return;
        }

        reviewContainer.innerHTML = '';

        data.components.forEach((component) => {
            const item = document.createElement('div');
            item.className = 'review-item';
            item.innerHTML = `
                <div class="review-item-top">
                    <span class="component-badge">${component.type}</span>
                    <strong>${component.title}</strong>
                </div>
                <div class="review-body">
                    <div class="review-column">
                        <label>Source</label>
                        <p>${component.source}</p>
                    </div>
                    <div class="review-column">
                        <label>Target</label>
                        <p>${component.translation}</p>
                    </div>
                </div>
            `;
            reviewContainer.appendChild(item);
        });

        if (Array.isArray(data.diagram) && data.diagram.length > 0) {
            const diagramHeader = document.createElement('div');
            diagramHeader.className = 'review-item';
            diagramHeader.innerHTML = `
                <div class="review-item-top">
                    <span class="component-badge">diagram</span>
                    <strong>Diagram review</strong>
                </div>
                <div class="review-body">
                    <div class="review-column">
                        <label>Source</label>
                        <p>${data.diagram.map((d) => d.source || d.title).join(' • ')}</p>
                    </div>
                    <div class="review-column">
                        <label>Target</label>
                        <p>${data.diagram.map((d) => d.translation || d.target || d.title).join(' • ')}</p>
                    </div>
                </div>
            `;
            reviewContainer.appendChild(diagramHeader);
        }

    }

    function updateStatus(message, type) {
        statusBox.innerText = message;
        statusBox.className = `status-box ${type}`;
        statusBox.classList.remove('hidden');
    }

    function updateProgress(percent, label) {
        progressPanel.classList.remove('hidden');
        progressBar.style.width = '42%';
        progressValue.textContent = '';
        progressLabel.textContent = label;
    }

    function hideProgress() {
        progressPanel.classList.add('hidden');
        progressBar.style.width = '42%';
        progressValue.textContent = '';
        progressLabel.textContent = 'Đang chuẩn bị...';
    }

    function setProcessingState(isProcessing) {
        document.querySelectorAll('#translationForm input, #translationForm select, #translationForm button')
            .forEach((control) => {
                control.disabled = isProcessing;
            });
        translationForm.classList.toggle('is-processing', isProcessing);
    }

    function downloadBlob(blob, filename) {
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        window.URL.revokeObjectURL(url);
    }

    function updateEditorStatus(message, type) {
        if (!editorStatus) return;
        editorStatus.textContent = message;
        editorStatus.className = `status-box ${type}`;
        editorStatus.classList.remove('hidden');
    }

    function scopeDocumentStyles(cssText) {
        return cssText.replace(/([^{}]+)\{([^{}]*)\}/g, (match, selectors, rules) => {
            if (selectors.trim().startsWith('@')) return match;
            const scopedSelectors = selectors.split(',').map((selector) => {
                const cleanSelector = selector.trim();
                if (!cleanSelector || cleanSelector.startsWith('@')) return cleanSelector;
                if (/^(html|body)$/i.test(cleanSelector)) return '#documentEditor';
                return `#documentEditor ${cleanSelector}`;
            }).join(', ');
            return `${scopedSelectors} { ${rules} }`;
        });
    }

    function openEditor(html, filename) {
        window.editorTemplateHtml = html;
        editorDirty = false;
        const parsed = new DOMParser().parseFromString(html, 'text/html');
        const sourceStyles = [...parsed.querySelectorAll('style')]
            .map((style) => scopeDocumentStyles(style.textContent || ''))
            .join('\n');
        let styleTag = document.getElementById('document-editor-styles');
        if (!styleTag) {
            styleTag = document.createElement('style');
            styleTag.id = 'document-editor-styles';
            document.head.appendChild(styleTag);
        }
        styleTag.textContent = sourceStyles;
        const root = parsed.getElementById('spire-document-root') || parsed.body;
        if (root) {
            root.querySelectorAll('.docx-diagram-editors').forEach((node) => node.remove());
            documentEditor.className = root.className || '';
            documentEditor.innerHTML = root.innerHTML;
        } else {
            documentEditor.className = '';
            documentEditor.innerHTML = html;
        }
        window.editorTemplateHtml = parsed.documentElement.outerHTML;
        assignEditorNodeIds(documentEditor);
        markProtectedDocumentRegions();
        markAbbreviationsProtected();
        setupTextBoxes();
        // window.attachTranslationMemory?.(documentEditor, editorStatus, apiBase);
        // generateTranslationMemoryFromEditor().catch((error) => {
        //     updateEditorStatus(error.message, 'error');
        // });
        documentEditor.querySelectorAll('table').forEach((table) => table.classList.add('editable-table'));
        hideRenderedFootnoteCopies();
        buildNavigationPanel();
        translatedFilename = filename.replace(/^Translated_[^_]+_/, 'Translated_');
        editorTitle.textContent = filename;
        translationView.classList.add('hidden');
        editorView.classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    function assignEditorNodeIds(container) {
        const editableNodes = container.querySelectorAll('p, h1, h2, h3, h4, h5, h6, td, th, [data-footnote-id]');
        editableNodes.forEach((node, index) => {
            if (!node.dataset.editorNodeId) {
                node.dataset.editorNodeId = `text-${index}`;
            }
            if (!node.dataset.originalText) {
                node.dataset.originalText = node.textContent || '';
            }
        });
    }

    function markProtectedDocumentRegions() {
        const protectedNodes = documentEditor.querySelectorAll(
            '.TOC, [class^="TOC"], [class*=" TOC"], .zcontents, .Header, .Footer, .docx-footnotes, .docx-footnote, .docx-diagram-editors, .docx-diagram-editor'
        );
        protectedNodes.forEach((node) => {
            node.dataset.nonEditable = 'true';
            node.querySelectorAll('[data-editor-node-id]').forEach((child) => {
                child.dataset.nonEditable = 'true';
            });
        });
    }

    function markAbbreviationsProtected() {
        const abbreviationHeading = [...documentEditor.querySelectorAll('h1, h2, h3, h4, h5, h6, p')]
            .find((node) => /^(abbreviations|chữ viết tắt)$/i.test(
                (node.textContent || '').replace(/\s+/g, ' ').trim()
            ));
        if (!abbreviationHeading) return;

        const blocks = [...documentEditor.querySelectorAll('h1, h2, h3, h4, h5, h6, p, table')];
        const startIndex = blocks.indexOf(abbreviationHeading);
        if (startIndex < 0) return;
        const nextHeadingIndex = blocks.findIndex(
            (node, index) => index > startIndex && /^H[1-6]$/i.test(node.tagName)
        );
        const endIndex = nextHeadingIndex < 0 ? blocks.length : nextHeadingIndex;
        const protectedBlocks = blocks.slice(startIndex, endIndex);

        protectedBlocks.forEach((block) => {
            block.dataset.nonEditable = 'true';
            block.removeAttribute('contenteditable');
            block.classList.remove('editor-text-box');
            block.querySelectorAll('[data-editor-node-id], [data-xml-id]').forEach((node) => {
                node.dataset.nonEditable = 'true';
                node.removeAttribute('contenteditable');
                node.classList.remove('editor-text-box');
                node.removeAttribute('tabindex');
            });
        });

        if (ABBREVIATIONS_EDITABLE) {
            protectedBlocks.forEach((block) => {
                delete block.dataset.nonEditable;
                block.querySelectorAll('p').forEach((node) => {
                    if (node === abbreviationHeading) return;
                    delete node.dataset.nonEditable;
                    node.dataset.abbreviationEditable = 'true';
                });
            });
        }
    }

    function setupTextBoxes() {
        documentEditor.querySelectorAll('[data-editor-node-id]').forEach((node) => {
            const isEditableDocumentNode = node.matches('p')
                && (isBodyOrNormalStyle(node) || node.dataset.abbreviationEditable === 'true');
            const isProtected = node.dataset.nonEditable === 'true' || node.closest('[data-non-editable="true"]');
            const isTextBox = isEditableDocumentNode && !isProtected;
            node.contentEditable = isTextBox ? 'true' : 'false';
            if (isTextBox) {
                node.classList.add('editor-text-box');
                if (node.dataset.debugListenersAttached !== 'true') {
                    node.addEventListener('focus', () => logEditorNodeDebug(node, 'focus'));
                    node.addEventListener('input', () => logEditorNodeDebug(node, 'input'));
                    node.dataset.debugListenersAttached = 'true';
                }
            }
        });
    }

    function isBodyOrNormalStyle(node) {
        const styleClasses = [...node.classList].map((value) => value.toLowerCase());
        return styleClasses.some((value) => (
            /^normal(?:[a-z0-9_-]*)$/.test(value)
            || value === 'bodytext'
            || value.startsWith('bodytext-')
            || /^bodytext\d/.test(value)
        ));
    }

    function logEditorNodeDebug(node, eventName) {
        const computedStyle = window.getComputedStyle(node);
        console.groupCollapsed(`[Editor ${eventName}] ${(node.textContent || '').trim().slice(0, 80)}`);
        console.table({
            tag: node.tagName.toLowerCase(),
            className: node.className || '',
            editorNodeId: node.dataset.editorNodeId || '',
            xmlId: node.dataset.xmlId || '',
            originalText: node.dataset.originalText || '',
            contentEditable: node.contentEditable,
            isContentEditable: node.isContentEditable,
            bodyOrNormalStyle: isBodyOrNormalStyle(node),
            abbreviationEditable: node.dataset.abbreviationEditable === 'true',
            parentTag: node.parentElement?.tagName.toLowerCase() || '',
            parentClass: node.parentElement?.className || '',
            fontFamily: computedStyle.fontFamily,
            fontSize: computedStyle.fontSize,
            fontWeight: computedStyle.fontWeight,
            fontStyle: computedStyle.fontStyle,
            textAlign: computedStyle.textAlign
        });
        console.groupEnd();
    }

    function logEditorParagraphSnapshot(node, eventName, details = {}) {
        if (!node) return;
        console.groupCollapsed(`[Edit trace ${eventName}] ${(node.textContent || '').trim().slice(0, 100)}`);
        console.debug('Paragraph', {
            nodeId: node.dataset.editorNodeId || '',
            xmlId: node.dataset.xmlId || '',
            originalText: node.dataset.originalText || '',
            textContent: node.textContent || '',
            innerHTML: node.innerHTML,
            ...details
        });
        console.table([...node.querySelectorAll('sup, [style*="vertical-align"], span')].map((child, index) => ({
            index,
            tag: child.tagName.toLowerCase(),
            text: child.textContent || '',
            className: child.className || '',
            style: child.getAttribute('style') || '',
            html: child.outerHTML
        })));
        console.groupEnd();
    }

    function hideRenderedFootnoteCopies() {
        const editableFootnotes = [...documentEditor.querySelectorAll('.docx-footnote')];
        editableFootnotes.forEach((footnote) => {
            const sourceText = (footnote.dataset.originalText || '').replace(/\s+/g, ' ').trim();
            if (!sourceText) return;
            const copy = [...documentEditor.querySelectorAll('p, li, td, div')]
                .filter((node) => !node.closest('.docx-footnotes') && node !== documentEditor)
                .filter((node) => (node.textContent || '').replace(/\s+/g, ' ').includes(sourceText))
                .sort((first, second) => (first.textContent || '').length - (second.textContent || '').length)[0];
            if (copy) copy.hidden = true;
        });
    }

    function getNavigationLabel(text) {
        return (text || '')
            .replace(/\[\d+\]/g, '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function buildNavigationPanel() {
        const list = document.getElementById('navigationList');
        const count = document.getElementById('navigationCount');
        if (!list || !count) return;
        list.innerHTML = '';

        if (documentEditor.classList.contains('pptx-document')) {
            buildPptxNavigation(list, count);
            return;
        }

        const blocks = [...documentEditor.querySelectorAll('h1, h2, h3, h4, h5, h6, p, table')]
            .filter((node) => !node.closest('table') || node.tagName === 'TABLE')
            .filter((node) => getNavigationLabel(node.textContent));

        count.textContent = blocks.length;
        if (!blocks.length) {
            list.innerHTML = '<p class="navigation-empty">Chưa tìm thấy section.</p>';
            return;
        }

        const navigationItems = [];
        blocks.forEach((block) => {
            block.id = `document-block-${blocks.indexOf(block)}`;
            const item = document.createElement('div');
            const isTable = block.tagName === 'TABLE';
            const isHeading = /^H[1-6]$/i.test(block.tagName);
            const text = getNavigationLabel(block.textContent);
            const label = isTable ? `Table: ${text}` : text;
            const level = isHeading ? Number(block.tagName.substring(1)) : (isTable ? 2 : 4);
            item.className = `navigation-item level-${level}`;
            item.title = label;
            item.setAttribute('role', 'button');
            item.tabIndex = 0;
            const labelElement = document.createElement('span');
            labelElement.className = 'navigation-label';
            labelElement.textContent = label;
            item.append(labelElement);
            labelElement.title = label;
            item.addEventListener('click', () => {
                navigateToSection(block, navigationItems, item, true);
            });
            item.addEventListener('keydown', (event) => {
                if (event.key === 'Enter' && event.target === item) item.click();
            });
            list.appendChild(item);
            navigationItems.push(item);
        });

        const observer = new IntersectionObserver((entries) => {
            const visible = entries
                .filter((entry) => entry.isIntersecting)
                .sort((first, second) => second.intersectionRatio - first.intersectionRatio)[0];
            if (!visible) return;
            const index = blocks.indexOf(visible.target);
            setActiveNavigationItem(navigationItems, navigationItems[index]);
        }, { root: document.querySelector('.document-stage'), threshold: [0.35, 0.7] });
        blocks.forEach((block) => observer.observe(block));
        setActiveNavigationItem(navigationItems, navigationItems[0]);
    }

    function buildPptxNavigation(list, count) {
        const slides = [...documentEditor.querySelectorAll('.pptx-slide')];
        const navigationItems = [];
        count.textContent = slides.length;
        if (!slides.length) {
            list.innerHTML = '<p class="navigation-empty">Chưa tìm thấy slide.</p>';
            return;
        }

        slides.forEach((slide, index) => {
            const item = document.createElement('button');
            const thumbnail = document.createElement('span');
            const label = document.createElement('span');
            const preview = slide.cloneNode(true);

            item.type = 'button';
            item.className = 'slide-navigation-item';
            item.title = `Slide ${index + 1}`;
            thumbnail.className = 'slide-thumbnail';
            label.className = 'slide-navigation-label';
            label.textContent = `Slide ${index + 1}`;
            preview.removeAttribute('contenteditable');
            preview.querySelectorAll('[contenteditable]').forEach((node) => node.removeAttribute('contenteditable'));
            thumbnail.appendChild(preview);
            item.appendChild(thumbnail);
            item.appendChild(label);
            item.addEventListener('click', () => {
                navigateToSection(slide, navigationItems, item, true);
            });
            list.appendChild(item);
            navigationItems.push(item);
        });

        list.querySelector('.slide-navigation-item')?.classList.add('is-active');

        const observer = new IntersectionObserver((entries) => {
            const visible = entries
                .filter((entry) => entry.isIntersecting)
                .sort((first, second) => second.intersectionRatio - first.intersectionRatio)[0];
            if (!visible) return;
            const index = slides.indexOf(visible.target);
            setActiveNavigationItem(navigationItems, navigationItems[index]);
        }, { root: document.querySelector('.document-stage'), threshold: [0.35, 0.7] });
        slides.forEach((slide) => observer.observe(slide));
    }

    async function blobToBase64(blob) {
        const bytes = new Uint8Array(await blob.arrayBuffer());
        let binary = '';
        const chunkSize = 0x8000;
        for (let index = 0; index < bytes.length; index += chunkSize) {
            binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
        }
        return btoa(binary);
    }

    function base64ToBlob(value, type) {
        const binary = atob(value);
        const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
        return new Blob([bytes], { type });
    }

    function collectTextEdits() {
        const edits = [...documentEditor.querySelectorAll('[data-editor-node-id], [data-xml-id]')]
            .filter((node) => !node.dataset.nonEditable && !node.closest('[data-non-editable="true"]'))
            .filter((node) => node.matches('p'))
            .map((node) => {
                let newText = node.textContent || '';
                return {
                    node_id: node.dataset.editorNodeId || '',
                    xml_id: node.dataset.xmlId || '',
                    old_text: node.dataset.originalText || node.textContent || '',
                    new_text: newText
                };
            })
            .filter((edit) => edit.old_text !== edit.new_text && edit.old_text.trim());
        console.groupCollapsed(`[Preview debug] Collected ${edits.length} edit(s)`);
        edits.forEach((edit, index) => {
            console.debug(`Edit ${index + 1}`, {
                nodeId: edit.node_id,
                xmlId: edit.xml_id,
                oldText: edit.old_text,
                newText: edit.new_text
            });
        });
        console.groupEnd();
        return edits;
    }

    async function requestPreviewPdf() {
        const edits = collectTextEdits();
        console.debug('[Preview debug] Sending preview request', {
            filename: translatedFilename,
            editCount: edits.length,
            edits
        });
        const response = await fetch(`${apiBase}/api/v1/document/preview`, {
            method: 'POST',
            headers: { ...ngrokHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify({
                document_base64: await blobToBase64(translatedBlob),
                edits,
                filename: translatedFilename
            })
        });
        if (!response.ok) {
            const error = await response.json().catch(() => null);
            console.error('[Preview debug] Preview request failed', {
                status: response.status,
                error
            });
            throw new Error(error?.detail || 'Không thể tạo PDF preview.');
        }
        console.debug('[Preview debug] Preview PDF received', {
            status: response.status,
            contentType: response.headers.get('content-type'),
            contentLength: response.headers.get('content-length')
        });
        return response.blob();
    }

    async function refreshPreview() {
        if (!translatedBlob) return;
        const requestId = ++previewRequestId;
        try {
            const previewBlob = (!editorDirty && translatedPdfBlob) || await requestPreviewPdf();
            if (requestId !== previewRequestId || previewModal.classList.contains('hidden')) return;
            if (previewObjectUrl) window.URL.revokeObjectURL(previewObjectUrl);
            previewObjectUrl = window.URL.createObjectURL(previewBlob);
            const sourceUrl = sourcePdfBlob ? window.URL.createObjectURL(sourcePdfBlob) : '';
            const installPdfFrame = (viewer, url, label) => {
                if (!viewer || !url) return;
                viewer.replaceChildren();
                const frame = document.createElement('iframe');
                frame.className = 'pdf-preview-frame';
                frame.title = label;
                frame.src = url;
                viewer.appendChild(frame);
            };
            installPdfFrame(sourcePdfViewer, sourceUrl, 'Original document with glossary highlights');
            installPdfFrame(previewPdfViewer, previewObjectUrl, 'Translated document with glossary highlights');
        } catch (error) {
            if (requestId !== previewRequestId) return;
            previewPdfViewer.replaceChildren();
            updateEditorStatus(`Preview lỗi: ${String(error.message || error)}`, 'error');
        }
    }

    function openPreview() {
        if (!translatedBlob) return;
        previewModal.classList.remove('hidden');
        editorShell?.classList.add('preview-is-open');
        document.body.classList.add('preview-open');
        refreshPreview();
    }

    function closePreview() {
        previewRequestId += 1;
        previewModal.classList.add('hidden');
        editorShell?.classList.remove('preview-is-open');
        document.body.classList.remove('preview-open');
        previewPdfViewer.replaceChildren();
        previewPdfDocument = null;
        if (previewObjectUrl) {
            window.URL.revokeObjectURL(previewObjectUrl);
            previewObjectUrl = null;
        }
    }

    async function renderTranslatedDocument(blob, filename) {
        const formData = new FormData();
        formData.append('document_file', blob, filename);
        const response = await fetch(`${apiBase}/api/v1/document/render`, {
            method: 'POST',
            headers: ngrokHeaders,
            body: formData
        });
        const result = await response.json().catch(() => null);
        if (!response.ok) throw new Error(result?.detail || 'Không thể mở trình chỉnh sửa tài liệu.');
        openEditor(result.html, result.filename || filename);
    }

    async function generateTranslationMemoryFromEditor() {
        const entries = [...documentEditor.querySelectorAll('[data-editor-node-id]')]
            .map((node) => ({
                source: node.dataset.originalText || node.textContent.replace(/\s+/g, ' ').trim(),
                target: ''
            }))
            .filter((entry) => entry.source);
        if (!entries.length) return;
        const response = await fetch(`${apiBase}/api/v1/translation-memory/generate`, {
            method: 'POST',
            headers: { ...ngrokHeaders, 'Content-Type': 'application/json' },
            body: JSON.stringify({ entries })
        });
        if (!response.ok) {
            const result = await response.json().catch(() => null);
            throw new Error(result?.detail || 'Không thể tạo translation_memory.json.');
        }
        if (translationMemoryFileName) translationMemoryFileName.textContent = 'translation_memory.json';
    }

    /* translationMemoryFileInput?.addEventListener('change', async () => {
        const file = translationMemoryFileInput.files?.[0];
        if (!file) return;
        const formData = new FormData();
        formData.append('memory_file', file, file.name);
        try {
            const response = await fetch(`${apiBase}/api/v1/translation-memory/load`, {
                method: 'POST',
                headers: ngrokHeaders,
                body: formData
            });
            const result = await response.json().catch(() => null);
            if (!response.ok) throw new Error(result?.detail || 'Không thể load Translation Memory JSON.');
            if (translationMemoryFileName) translationMemoryFileName.textContent = result.filename;
            updateEditorStatus(`Đã load ${result.filename} (${result.entry_count} entries).`, 'success');
        } catch (error) {
            updateEditorStatus(error.message, 'error');
        } finally {
            translationMemoryFileInput.value = '';
        }
    }); */

    async function importEditorFile(file, input) {
        if (!file) return;
        if (!file.name.toLowerCase().endsWith('.docx')) {
            updateEditorStatus('Live Editor test hiện chỉ hỗ trợ file DOCX.', 'error');
            return;
        }

        importEditorFileBtn.disabled = true;
        updateEditorStatus('Đang mở file test trong Live Editor...', 'info');
        try {
            translatedBlob = file;
            editorDirty = false;
            await renderTranslatedDocument(file, file.name);
            updateEditorStatus('Đã mở file test. Bạn có thể sửa Body Text/Normal rồi Preview hoặc Export.', 'success');
        } catch (error) {
            translatedBlob = null;
            updateEditorStatus(error.message, 'error');
        } finally {
            importEditorFileBtn.disabled = false;
            if (input) input.value = '';
        }
    }

    if (importEditorFileBtn && editorFileInput) {
        importEditorFileBtn.addEventListener('click', () => editorFileInput.click());
        editorFileInput.addEventListener('change', () => importEditorFile(editorFileInput.files?.[0], editorFileInput));
    }
    if (quickImportEditorBtn && quickEditorFileInput) {
        quickImportEditorBtn.addEventListener('click', () => quickEditorFileInput.click());
        quickEditorFileInput.addEventListener('change', () => importEditorFile(quickEditorFileInput.files?.[0], quickEditorFileInput));
    }

    function isEditableInBrowser(filename) {
        return filename.toLowerCase().endsWith('.docx');
    }

    function setupDropInput(input) {
        if (!input) return;
        ['dragenter', 'dragover'].forEach((eventName) => {
            input.addEventListener(eventName, (event) => {
                event.preventDefault();
                input.classList.add('is-dragging');
            });
        });

        ['dragleave', 'drop'].forEach((eventName) => {
            input.addEventListener(eventName, (event) => {
                event.preventDefault();
                input.classList.remove('is-dragging');
            });
        });
    }

    function toggleGlossaryUI() {
        if (useGlossaryToggle && glossaryUploadGroup) {
            if (useGlossaryToggle.checked) {
                glossaryUploadGroup.style.display = 'block';
            } else {
                glossaryUploadGroup.style.display = 'none';
                if (glossaryFileInput) glossaryFileInput.value = '';
            }
        }
    }

    async function checkBackendConnection() {
        try {
            const response = await fetch(`${apiBase}/api/v1/health`, {
                method: 'GET',
                headers: ngrokHeaders
            });
            if (!response.ok) throw new Error('Backend offline');
            return true;
        } catch (error) {
            updateStatus(`Backend is not reachable. Start it at ${apiBase} before translation.`, 'error');
            return false;
        }
    }

    async function parseDocumentStructure(file) {
        const formData = new FormData();
        formData.append('document_file', file);

        const response = await fetch(`${apiBase}/api/v1/document/structure`, {
            method: 'POST',
            headers: ngrokHeaders,
            body: formData
        });

        if (!response.ok) {
            const err = await response.json().catch(() => null);
            throw new Error(err?.detail || 'Failed to parse document structure');
        }

        return await response.json();
    }

    translationForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        if (!docFileInput.files || docFileInput.files.length === 0) {
            updateStatus('Vui lòng chọn 1 file tài liệu!', 'error');
            return;
        }

        if (!(await checkBackendConnection())) {
            return;
        }

        const file = docFileInput.files[0];
        const fileExtension = file.name.toLowerCase().split('.').pop();
        if (!['docx', 'pptx'].includes(fileExtension)) {
            updateStatus('Please select a DOCX or PPTX file.', 'error');
            return;
        }
        if (file.size > MAX_DOCUMENT_SIZE_BYTES) {
            updateStatus('The document exceeds the 50 MB demo limit.', 'error');
            return;
        }
        const formData = new FormData();
        formData.append('document_file', file);

        const isGlossaryEnabled = useGlossaryToggle ? useGlossaryToggle.checked : false;
        if (isGlossaryEnabled && glossaryFileInput && glossaryFileInput.files.length > 0) {
            formData.append('glossary_file', glossaryFileInput.files[0]);
        }

        formData.append('target_language', targetLangSelect.value);
        formData.append('model_name', modelNameSelect.value);

        updateProgress(10, 'Đang tải tài liệu lên...');
        updateStatus('Đang kích hoạt AI Translation Pipeline...', 'info');
        submitBtn.disabled = true;
        submitBtn.textContent = 'Đang dịch...';
        setProcessingState(true);

        try {
            const includeComparison = ENABLE_COMPARISON_UI && isEditableInBrowser(file.name);
            const response = await fetch(
                `${apiBase}/api/v1/translate?include_comparison=${includeComparison}`,
                {
                    method: 'POST',
                    headers: ngrokHeaders,
                    body: formData
                }
            );

            if (!response.ok) {
                const errJson = await response.json().catch(() => null);
                throw new Error(errJson?.detail || 'Lỗi xử lý dịch thuật từ Server.');
            }

            updateProgress(82, includeComparison
                ? 'Đang chuẩn bị tài liệu để chỉnh sửa...'
                : 'Đang chuẩn bị file để tải xuống...');
            const responseType = response.headers.get('content-type') || '';
            const result = responseType.includes('application/json')
                ? await response.json()
                : null;
            const blob = result
                ? base64ToBlob(result.document_base64, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                : await response.blob();
            translatedBlob = blob;
            if (includeComparison) {
                sourcePdfBlob = result?.source_pdf_base64
                    ? base64ToBlob(result.source_pdf_base64, 'application/pdf') : null;
                translatedPdfBlob = result?.translated_pdf_base64
                    ? base64ToBlob(result.translated_pdf_base64, 'application/pdf') : null;
                translatedFilename = `Translated_${file.name}`;
                translationView.classList.add('hidden');
                editorView.classList.remove('hidden');
                hideProgress();
                updateEditorStatus('Translation is ready in the editor and preview.', 'success');
                openPreview();
            } else {
                downloadBlob(blob, `Translated_${file.name}`);
                translatedBlob = null;
                hideProgress();
                updateStatus('Dịch hoàn tất. File đã được tự động tải xuống.', 'success');
            }
            if (includeComparison) {
                updateStatus('Dịch xong. Bạn có thể rà soát và export khi sẵn sàng.', 'success');
            }
        } catch (err) {
            console.error(err);
            const message = `Lỗi trong quá trình dịch: ${err.message || 'Đã xảy ra lỗi không xác định.'}`;
            updateStatus(message, 'error');
            window.alert(message);
            window.location.reload();
        } finally {
            submitBtn.disabled = false;
            submitBtn.textContent = 'Dịch tài liệu';
            setProcessingState(false);
        }
    });

    if (useGlossaryToggle) {
        useGlossaryToggle.addEventListener('change', toggleGlossaryUI);
        toggleGlossaryUI();
    }

    setupDropInput(docFileInput);
    setupDropInput(glossaryFileInput);

    documentEditor.addEventListener('input', () => {
        editorDirty = true;
        exportBtn.classList.add('is-needed');
        updateEditorStatus('Tài liệu đã được chỉnh sửa. Bấm Export để tải bản đã sửa.', 'info');
        const target = window.getSelection()?.anchorNode?.parentElement?.closest('[data-editor-node-id], [data-xml-id]');
        console.debug('[Editor debug] Input changed', {
            nodeId: target?.dataset.editorNodeId || '',
            xmlId: target?.dataset.xmlId || '',
            originalText: target?.dataset.originalText || '',
            currentText: target?.textContent || ''
        });
        logEditorParagraphSnapshot(target, 'after-input');
    });

    documentEditor.addEventListener('beforeinput', (event) => {
        const target = window.getSelection()?.anchorNode?.parentElement?.closest('[data-editor-node-id], [data-xml-id]');
        logEditorParagraphSnapshot(target, 'before-input', {
            inputType: event.inputType,
            data: event.data,
            isComposing: event.isComposing
        });
    });

    document.querySelectorAll('.tool-button').forEach((button) => {
        button.addEventListener('mousedown', (event) => event.preventDefault());
        button.addEventListener('click', () => {
            if (!documentEditor.contains(window.getSelection()?.anchorNode)) return;
            document.execCommand(button.dataset.command, false);
            documentEditor.dispatchEvent(new Event('input', { bubbles: true }));
        });
    });

    previewBtn?.addEventListener('click', openPreview);
    previewModal?.querySelectorAll('[data-close-preview]').forEach((element) => {
        element.addEventListener('click', closePreview);
    });
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && previewModal && !previewModal.classList.contains('hidden')) closePreview();
    });

    document.getElementById('backToTranslationBtn').addEventListener('click', () => {
        editorView.classList.add('hidden');
        translationView.classList.remove('hidden');
        hideProgress();
        setProcessingState(false);
        documentEditor.innerHTML = '';
        translatedBlob = null;
        editorDirty = false;
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    exportBtn.addEventListener('click', async () => {
        if (!documentEditor.innerHTML.trim() || !translatedBlob) return;
        exportBtn.disabled = true;
        updateEditorStatus(`Đang tạo ${translatedFilename.split('.').pop().toUpperCase()} từ nội dung đã chỉnh sửa...`, 'info');
        try {
            const documentBase64 = await blobToBase64(translatedBlob);
            const edits = collectTextEdits();
            updateEditorStatus(`Đang export ${edits.length} thay đổi...`, 'info');
            const response = await fetch(`${apiBase}/api/v1/document/export`, {
                method: 'POST',
                headers: { ...ngrokHeaders, 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    document_base64: documentBase64,
                    edits,
                    filename: translatedFilename
                })
            });
            if (!response.ok) {
                const error = await response.json().catch(() => null);
                throw new Error(error?.detail || 'Export tài liệu thất bại.');
            }
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = translatedFilename;
            link.click();
            window.URL.revokeObjectURL(url);
            updateEditorStatus('Export thành công. File đã được tải xuống.', 'success');
        } catch (error) {
            updateEditorStatus(error.message, 'error');
        } finally {
            exportBtn.disabled = false;
        }
    });
});