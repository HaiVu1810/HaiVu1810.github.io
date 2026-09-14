(() => {
    function attachTranslationMemory(documentEditor, editorStatus, apiBase) {
        if (!documentEditor || documentEditor.dataset.translationMemoryAttached === 'true') return;
        documentEditor.dataset.translationMemoryAttached = 'true';

        let pending = null;
        const popup = document.createElement('div');
        popup.className = 'translation-memory-suggestion hidden';
        popup.innerHTML = `
            <strong>Translation Memory suggestion</strong>
            <p class="translation-memory-text"></p>
            <div class="translation-memory-actions">
                <button type="button" data-tm-accept>Yes</button>
                <button type="button" data-tm-reject>No</button>
            </div>
        `;
        document.body.appendChild(popup);
        const message = popup.querySelector('.translation-memory-text');

        function close() {
            pending = null;
            popup.classList.add('hidden');
            if (window.CSS?.highlights) window.CSS.highlights.delete('translation-memory-selection');
        }

        function show(entry, range, paragraph) {
            pending = { entry, range, paragraph };
            if (window.CSS?.highlights) {
                window.CSS.highlights.set(
                    'translation-memory-selection',
                    new Highlight(range)
                );
            }
            message.textContent = `"${entry.source}" -> "${entry.target}"`;
            const rect = range.getBoundingClientRect();
            popup.style.left = `${Math.min(window.innerWidth - 360, Math.max(12, rect.left))}px`;
            popup.style.top = `${Math.min(window.innerHeight - 130, Math.max(12, rect.bottom + 8))}px`;
            popup.classList.remove('hidden');
        }

        async function checkParagraph(paragraph) {
            const sourceText = paragraph.dataset.originalText || paragraph.textContent || '';
            if (!sourceText.trim()) return;
            const range = document.createRange();
            range.selectNodeContents(paragraph);
            const response = await fetch(`${apiBase}/api/v1/translation-memory/suggest`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ text: sourceText, exact_only: true })
            });
            const result = await response.json().catch(() => null);
            if (!response.ok) throw new Error(result?.detail || 'Translation Memory lookup failed.');
            if (result?.entry) {
                show(result.entry, range, paragraph);
            } else if (editorStatus) {
                editorStatus.textContent = 'No matching Translation Memory term found.';
                editorStatus.className = 'status-box info';
                editorStatus.classList.remove('hidden');
            }
        }

        documentEditor.addEventListener('click', (event) => {
            const paragraph = event.target.closest('p');
            if (!paragraph || !documentEditor.contains(paragraph)) return;
            checkParagraph(paragraph).catch((error) => {
                if (editorStatus) {
                    editorStatus.textContent = error.message;
                    editorStatus.className = 'status-box error';
                }
            });
        });

        popup.querySelector('[data-tm-accept]').addEventListener('click', () => {
            if (!pending) return;
            const { entry, range, paragraph } = pending;
            range.deleteContents();
            range.insertNode(document.createTextNode(entry.target));
            paragraph.normalize();
            documentEditor.dispatchEvent(new Event('input', { bubbles: true }));
            close();
        });
        popup.querySelector('[data-tm-reject]').addEventListener('click', close);
    }

    window.attachTranslationMemory = attachTranslationMemory;
})();
