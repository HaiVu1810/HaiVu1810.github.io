document.addEventListener('DOMContentLoaded', async () => {
    const hostName = window.location.hostname || 'localhost';
    const configuredApiBase = window.TRANSLATION_CONFIG?.API_BASE_URL?.trim();
    const apiBase = (configuredApiBase || `http://${hostName}:8000`).replace(/\/+$/, '');
    const ngrokHeaders = { 'ngrok-skip-browser-warning': 'true' };
    const stage = document.getElementById('stage');
    const slideList = document.getElementById('slideList');
    const slideCount = document.getElementById('slideCount');
    const editCount = document.getElementById('editCount');
    const status = document.getElementById('status');
    const fileTitle = document.getElementById('fileTitle');
    const exportButton = document.getElementById('exportButton');
    let filename = 'Translated_Presentation.pptx';
    let sourceBlob = null;

    function showStatus(message, type = 'info') {
        status.textContent = message;
        status.className = `status ${type}`;
        status.classList.remove('hidden');
    }

    function base64ToBlob(base64) {
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let index = 0; index < binary.length; index += 1) {
            bytes[index] = binary.charCodeAt(index);
        }
        return new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    }

    function assignEditorIds(root) {
        root.querySelectorAll('[data-editor-node-id]').forEach((node) => {
            if (!node.dataset.originalText) node.dataset.originalText = node.textContent || '';
        });
    }

    function collectEdits() {
        return [...stage.querySelectorAll('[data-editor-node-id]')]
            .map((node) => ({
                node_id: node.dataset.editorNodeId,
                old_text: node.dataset.originalText || '',
                new_text: node.textContent || ''
            }))
            .filter((edit) => edit.old_text !== edit.new_text && edit.old_text.trim());
    }

    function buildSlideNavigation(slides) {
        slideList.innerHTML = '';
        slideCount.textContent = slides.length;
        slides.forEach((slide, index) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'slide-item';
            button.innerHTML = `<span class="thumbnail"></span><span class="slide-label">Slide ${index + 1}</span>`;
            const preview = slide.cloneNode(true);
            preview.removeAttribute('contenteditable');
            preview.querySelectorAll('[contenteditable]').forEach((node) => node.removeAttribute('contenteditable'));
            button.querySelector('.thumbnail').appendChild(preview);
            button.addEventListener('click', () => {
                slide.scrollIntoView({ behavior: 'smooth', block: 'center' });
                document.querySelectorAll('.slide-item.is-active').forEach((item) => item.classList.remove('is-active'));
                button.classList.add('is-active');
            });
            slideList.appendChild(button);
        });
        slideList.querySelector('.slide-item')?.classList.add('is-active');

        const observer = new IntersectionObserver((entries) => {
            const visible = entries.filter((entry) => entry.isIntersecting)
                .sort((first, second) => second.intersectionRatio - first.intersectionRatio)[0];
            if (!visible) return;
            const index = slides.indexOf(visible.target);
            slideList.querySelectorAll('.slide-item').forEach((item, itemIndex) => {
                item.classList.toggle('is-active', itemIndex === index);
            });
        }, { root: document.querySelector('.stage'), threshold: 0.55 });
        slides.forEach((slide) => observer.observe(slide));
    }

    function updateEditCount() {
        editCount.textContent = `${collectEdits().length} thay đổi`;
    }

    async function loadPresentation() {
        const payload = sessionStorage.getItem('pptx-editor-payload');
        if (!payload) {
            showStatus('Không tìm thấy file PowerPoint cần mở.', 'error');
            return;
        }
        try {
            const data = JSON.parse(payload);
            sessionStorage.removeItem('pptx-editor-payload');
            filename = data.filename || filename;
            fileTitle.textContent = filename;
            sourceBlob = base64ToBlob(data.document_base64);
            const formData = new FormData();
            formData.append('document_file', sourceBlob, filename);
            const response = await fetch(`${apiBase}/api/v1/document/render`, {
                method: 'POST',
                headers: ngrokHeaders,
                body: formData
            });
            const result = await response.json().catch(() => null);
            if (!response.ok) throw new Error(result?.detail || 'Không thể render PowerPoint.');
            const parsed = new DOMParser().parseFromString(result.html, 'text/html');
            const root = parsed.getElementById('document-editor-root');
            if (!root) throw new Error('Backend không trả về slide editor.');
            stage.className = `stage ${root.className || ''}`;
            stage.innerHTML = root.innerHTML;
            assignEditorIds(stage);
            const slides = [...stage.querySelectorAll('.pptx-slide')];
            buildSlideNavigation(slides);
            stage.addEventListener('input', updateEditCount);
            showStatus('PowerPoint đã sẵn sàng để chỉnh sửa.', 'success');
        } catch (error) {
            showStatus(error.message, 'error');
        }
    }

    exportButton.addEventListener('click', async () => {
        if (!sourceBlob) return;
        const edits = collectEdits();
        exportButton.disabled = true;
        showStatus(`Đang export ${edits.length} thay đổi...`, 'info');
        try {
            const bytes = new Uint8Array(await sourceBlob.arrayBuffer());
            let binary = '';
            const chunkSize = 0x8000;
            for (let index = 0; index < bytes.length; index += chunkSize) {
                binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
            }
            const response = await fetch(`${apiBase}/api/v1/document/export`, {
                method: 'POST',
                headers: { ...ngrokHeaders, 'Content-Type': 'application/json' },
                body: JSON.stringify({ document_base64: btoa(binary), edits, filename })
            });
            if (!response.ok) {
                const error = await response.json().catch(() => null);
                throw new Error(error?.detail || 'Export PPTX thất bại.');
            }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            link.click();
            URL.revokeObjectURL(url);
            showStatus('Export PPTX thành công.', 'success');
        } catch (error) {
            showStatus(error.message, 'error');
        } finally {
            exportButton.disabled = false;
        }
    });

    document.getElementById('backButton').addEventListener('click', () => {
        window.location.href = '../index.html';
    });

    await loadPresentation();
});
