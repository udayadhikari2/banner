document.addEventListener('DOMContentLoaded', () => {
    const fileInput      = document.getElementById('fileInput');
    const imageInput     = document.getElementById('imageInput');
    const downloadBtn    = document.getElementById('downloadBtn');
    const clearSessionBtn = document.getElementById('clearSessionBtn');
    const bannerContent  = document.getElementById('bannerContent');

    let studentData = [];
    let imageMap    = new Map();
    let gridLayout  = [1, 16, 16, 16, 16, 7];

    // ── Session helpers ───────────────────────────────────────────────────────
    const KEYS = {
        layout: 'banner_gridLayout',
        data:   'banner_studentData',
        images: 'banner_imageMap'
    };

    function saveSession() {
        localStorage.setItem(KEYS.layout, JSON.stringify(gridLayout));
        localStorage.setItem(KEYS.data,   JSON.stringify(studentData));
        const imageObj = {};
        imageMap.forEach((v, k) => { imageObj[k] = v; });
        localStorage.setItem(KEYS.images, JSON.stringify(imageObj));
    }

    function loadSession() {
        const savedLayout = localStorage.getItem(KEYS.layout);
        const savedData   = localStorage.getItem(KEYS.data);
        const savedImages = localStorage.getItem(KEYS.images);
        if (savedLayout) gridLayout  = JSON.parse(savedLayout);
        if (savedData)   studentData = JSON.parse(savedData);
        if (savedImages) {
            const obj = JSON.parse(savedImages);
            imageMap = new Map(Object.entries(obj).map(([k, v]) => [parseInt(k), v]));
        }
    }

    function resetState() {
        Object.values(KEYS).forEach(k => localStorage.removeItem(k));
        gridLayout  = [1, 16, 16, 16, 16, 7];
        studentData = [];
        imageMap    = new Map();
        fileInput.value  = '';
        imageInput.value = '';
    }

    // ── Init ──────────────────────────────────────────────────────────────────
    loadSession();
    renderGrid();

    // ── 1. CSV / XLSX Upload ──────────────────────────────────────────────────
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
            const workbook  = XLSX.read(evt.target.result, { type: 'binary' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const data      = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            const startIndex = (typeof data[0][1] === 'string' && data[0][1].toLowerCase().includes('gpa')) ? 1 : 0;
            const totalSlots = gridLayout.reduce((a, b) => a + b, 0);
            studentData = data.slice(startIndex).map(row => ({
                name: row[0] || 'N/A',
                gpa:  row[1] || '0.00'
            })).slice(0, totalSlots);
            saveSession();
            renderGrid();
        };
        reader.readAsBinaryString(file);
    });

    // ── 2. Bulk Image Upload (name-matched) ───────────────────────────────────
    imageInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        if (files.length === 0) return;

        let processed = 0, matchedCount = 0;
        const normalize = (str) => str.toLowerCase().replace(/[^a-z0-9]/g, '').trim();
        
        // Track which students in THIS upload batch get matched to avoid multiple files matching same student
        // and to support matching duplicate student names with multiple files.
        const matchedIndicesInBatch = new Set();

        files.forEach((file) => {
            const reader = new FileReader();
            reader.onload = (evt) => {
                const fileName = file.name.split('.')[0];
                const normalizedFileName = normalize(fileName);
                
                // Priority 1: Exact Match (normalized)
                let matchIndex = studentData.findIndex((s, idx) => 
                    !matchedIndicesInBatch.has(idx) && normalize(s.name) === normalizedFileName
                );

                // Priority 2: Fuzzy Match (one contains the other)
                if (matchIndex === -1) {
                    matchIndex = studentData.findIndex((s, idx) => {
                        if (matchedIndicesInBatch.has(idx)) return false;
                        const sName = normalize(s.name);
                        // Avoid matching empty names or very short names to everything
                        if (sName.length < 2 || normalizedFileName.length < 2) return false;
                        return normalizedFileName.includes(sName) || sName.includes(normalizedFileName);
                    });
                }

                if (matchIndex !== -1) {
                    imageMap.set(matchIndex, evt.target.result);
                    matchedIndicesInBatch.add(matchIndex);
                    matchedCount++;
                }
                
                processed++;
                if (processed === files.length) {
                    console.log(`Matched ${matchedCount}/${files.length} images.`);
                    saveSession();
                    renderGrid();
                    // Clear input so same files can be re-uploaded if needed
                    imageInput.value = '';
                }
            };
            reader.readAsDataURL(file);
        });
    });

    // ── 3. Render Grid ────────────────────────────────────────────────────────
    function renderGrid() {
        bannerContent.innerHTML = '';
        const gridContainer = document.createElement('div');
        gridContainer.className = 'banner-grid';
        const totalSlots = gridLayout.reduce((a, b) => a + b, 0);
        let currentIndex = 0;

        gridLayout.forEach((count, rowIndex) => {
            const rowGroup = document.createElement('div');
            rowGroup.className = 'row-group';

            // Row control bar
            const rowCtrl = document.createElement('div');
            rowCtrl.className = 'row-ctrl';

            const label = document.createElement('span');
            label.className = 'row-label';
            label.textContent = `Row ${rowIndex + 1}`;

            const numInput = document.createElement('input');
            numInput.type      = 'number';
            numInput.className = 'row-count-input';
            numInput.value     = count;
            numInput.min       = 1;
            numInput.max       = 50;
            numInput.title     = 'Slots in this row';
            numInput.addEventListener('change', () => {
                const v = Math.max(1, Math.min(50, parseInt(numInput.value) || 1));
                numInput.value = v;
                gridLayout[rowIndex] = v;
                saveSession();
                renderGrid();
            });

            const slotsLabel = document.createElement('span');
            slotsLabel.className   = 'row-slots-label';
            slotsLabel.textContent = 'slots';

            rowCtrl.appendChild(label);
            rowCtrl.appendChild(numInput);
            rowCtrl.appendChild(slotsLabel);

            if (gridLayout.length > 1) {
                const removeBtn = document.createElement('button');
                removeBtn.className   = 'row-remove-btn';
                removeBtn.title       = 'Remove this row';
                removeBtn.textContent = '✕';
                removeBtn.addEventListener('click', () => {
                    gridLayout.splice(rowIndex, 1);
                    saveSession();
                    renderGrid();
                });
                rowCtrl.appendChild(removeBtn);
            }

            rowGroup.appendChild(rowCtrl);

            // Student cards
            const rowDiv = document.createElement('div');
            rowDiv.className = `banner-row ${rowIndex === 0 ? 'row-hero' : ''}`;
            for (let i = 0; i < count; i++) {
                const student   = studentData[currentIndex] || { name: `Student ${currentIndex + 1}`, gpa: '-' };
                const imageData = imageMap.get(currentIndex);
                rowDiv.appendChild(createStudentCard(student, imageData, currentIndex));
                currentIndex++;
            }

            rowGroup.appendChild(rowDiv);
            gridContainer.appendChild(rowGroup);
        });

        // "+ Add Row" button
        const addBtnRow = document.createElement('div');
        addBtnRow.className = 'add-slot-row';
        const addBtn = document.createElement('button');
        addBtn.className = 'slot-add-btn';
        addBtn.id        = 'addRowBtn';
        addBtn.innerHTML = `<span class="slot-add-icon">+</span> Add Row <span class="slot-count-badge">${totalSlots} slots total</span>`;
        addBtn.addEventListener('click', () => {
            gridLayout.push(1);
            saveSession();
            renderGrid();
        });
        addBtnRow.appendChild(addBtn);
        gridContainer.appendChild(addBtnRow);

        bannerContent.appendChild(gridContainer);
    }

    // ── 4. Student Card ───────────────────────────────────────────────────────
    function createStudentCard(student, imageData, index) {
        const card = document.createElement('div');
        card.className = 'student-card';
        card.innerHTML = `
            <div class="photo-container" data-index="${index}">
                ${imageData
                    ? `<img src="${imageData}" alt="Photo">`
                    : '<span class="placeholder-text">Click to Add</span>'}
            </div>
            <div class="student-info">
                <span class="student-name">${student.name}</span>
                <span class="student-gpa">${formatGPA(student.gpa)}</span>
            </div>
        `;
        card.querySelector('.photo-container').addEventListener('click', () => {
            const inp    = document.createElement('input');
            inp.type     = 'file';
            inp.accept   = 'image/*';
            inp.onchange = (e) => {
                const reader = new FileReader();
                reader.onload = (evt) => {
                    imageMap.set(index, evt.target.result);
                    saveSession();
                    renderGrid();
                };
                reader.readAsDataURL(e.target.files[0]);
            };
            inp.click();
        });
        return card;
    }

    function formatGPA(val) {
        if (!val || val === 'GPA' || val === '-') return '';
        return `GPA : ${val}`;
    }

    // ── 5. Download — hide controls during capture ────────────────────────────
    downloadBtn.addEventListener('click', () => {
        downloadBtn.textContent = 'Generating...';
        downloadBtn.disabled    = true;

        const hideEls = document.querySelectorAll('.row-ctrl, .add-slot-row');
        hideEls.forEach(el => el.classList.add('capture-hidden'));

        setTimeout(() => {
            html2canvas(bannerContent, {
                scale:        3,
                useCORS:      true,
                backgroundColor: '#ffffff',
                scrollX:      -window.scrollX,
                scrollY:      -window.scrollY,
                windowWidth:  bannerContent.scrollWidth,
                windowHeight: bannerContent.scrollHeight
            }).then(canvas => {
                const link      = document.createElement('a');
                link.download   = 'achievement-banner.png';
                link.href       = canvas.toDataURL('image/png');
                link.click();
            }).catch(err => {
                console.error(err);
                alert('Error generating image.');
            }).finally(() => {
                hideEls.forEach(el => el.classList.remove('capture-hidden'));
                downloadBtn.textContent = 'Download Banner';
                downloadBtn.disabled    = false;
            });
        }, 50);
    });

    // ── 6. Clear Session ──────────────────────────────────────────────────────
    clearSessionBtn.addEventListener('click', () => {
        if (confirm('Clear all saved session data? This will reset the banner to its default state.')) {
            resetState();
            renderGrid();
        }
    });
});
