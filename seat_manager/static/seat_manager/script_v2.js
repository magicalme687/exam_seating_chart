document.addEventListener('DOMContentLoaded', () => {
    // --- Theme Toggle ---
    const themeToggleBtn = document.getElementById('theme-toggle');
    const themeIcon = document.getElementById('theme-icon');

    // Check local storage or system preference
    const currentTheme = localStorage.getItem('theme') || 'dark';
    if (currentTheme === 'light') {
        document.documentElement.setAttribute('data-theme', 'light');
        if (themeIcon) {
            themeIcon.classList.replace('fa-sun', 'fa-moon');
        }
    }

    if (themeToggleBtn) {
        themeToggleBtn.addEventListener('click', () => {
            console.log("Theme toggle clicked!");
            let theme = document.documentElement.getAttribute('data-theme');
            console.log("Current theme attribute:", theme);
            if (theme === 'light') {
                document.documentElement.removeAttribute('data-theme');
                localStorage.setItem('theme', 'dark');
                themeIcon.classList.replace('fa-moon', 'fa-sun');
                console.log("Switched to dark mode.");
            } else {
                document.documentElement.setAttribute('data-theme', 'light');
                localStorage.setItem('theme', 'light');
                themeIcon.classList.replace('fa-sun', 'fa-moon');
                console.log("Switched to light mode.");
            }
        });
    }

    // --- Elements ---
    const roomCountInput = document.getElementById('room-count');
    const generateRoomsBtn = document.getElementById('generate-rooms-btn');
    const setupError = document.getElementById('setup-error');

    const mainContent = document.getElementById('main-content');
    const batchRoomCheckboxes = document.getElementById('batch-room-checkboxes');
    const roomsContainer = document.getElementById('rooms-container');

    // Batch Config Elements
    const batchRowsInput = document.getElementById('batch-rows');
    const batchColsInput = document.getElementById('batch-cols');
    const selectAllBtn = document.getElementById('select-all-btn');
    const deselectAllBtn = document.getElementById('deselect-all-btn');
    const applyBatchBtn = document.getElementById('apply-batch-btn');
    const batchError = document.getElementById('batch-error');

    // Final Generate Elements
    const finalGenerateBtn = document.getElementById('final-generate-btn');
    const finalError = document.getElementById('final-error');
    const outputContent = document.getElementById('output-content');
    const seatingResultsContainer = document.getElementById('seating-results-container');
    const backToConfigBtn = document.getElementById('back-to-config');

    let roomData = [];

    // --- Tab Navigation Setup ---
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            // Remove active from all tabs
            document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));

            // Set clicked as active
            btn.classList.add('active');
            const targetId = btn.getAttribute('data-target');
            document.getElementById(targetId).classList.remove('hidden');
        });
    });

    // --- Subject Parsing logic ---
    window.subjectCodes = {
        "I Yr": [],
        "II Yr": [],
        "III Yr": [],
        "IV Yr": []
    };

    const subjectFileInput = document.getElementById('subject-file');
    if (subjectFileInput) {
        subjectFileInput.addEventListener('change', function (e) {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function (evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    // Clear previous
                    window.subjectCodes = { "I Yr": [], "II Yr": [], "III Yr": [], "IV Yr": [] };

                    // Parse columns (assuming no headers, or skipping row 0 if it's headers. Let's assume row 0 might be headers like '1st Year')
                    // Start from row 0, if it looks like a subject code, keep it.
                    json.forEach(row => {
                        // Combines Subject Code and Name into a single string if available
                        if (row[0]) {
                            let txt = String(row[0]).trim();
                            if (row[1]) txt += " - " + String(row[1]).trim();
                            window.subjectCodes["I Yr"].push(txt);
                        }
                        if (row[2]) {
                            let txt = String(row[2]).trim();
                            if (row[3]) txt += " - " + String(row[3]).trim();
                            window.subjectCodes["II Yr"].push(txt);
                        }
                        if (row[4]) {
                            let txt = String(row[4]).trim();
                            if (row[5]) txt += " - " + String(row[5]).trim();
                            window.subjectCodes["III Yr"].push(txt);
                        }
                        if (row[6]) {
                            let txt = String(row[6]).trim();
                            if (row[7]) txt += " - " + String(row[7]).trim();
                            window.subjectCodes["IV Yr"].push(txt);
                        }
                    });

                    // Filter out likely headers (like "1st year") if needed... but just relying on user format.
                    console.log("Parsed Subject Codes:", window.subjectCodes);
                    alert("Subject Codes Loaded Successfully!");

                    // Re-populate any active session dropdowns
                    document.querySelectorAll('.subject-select').forEach(select => {
                        const yr = select.dataset.year;
                        populateDropdown(select, yr);
                    });

                } catch (err) {
                    console.error("Error parsing subjects excel:", err);
                    alert("Failed to parse Subject Excel file.");
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // Toggle Subject Source
    const subjectSourceRadios = document.querySelectorAll('input[name="subject-source"]');
    const excelSourceDiv = document.querySelector('.subject-source-excel');
    const manualSourceDiv = document.querySelector('.subject-source-manual');
    const subjectFile = document.getElementById('subject-file');

    // Update pill toggle visual state
    const pillExcel = document.getElementById('pill-excel');
    const pillManual = document.getElementById('pill-manual');
    const ACTIVE_PILL_STYLE = 'background:linear-gradient(135deg,var(--primary-color),var(--primary-hover));color:white;box-shadow:0 2px 8px rgba(139,92,246,0.35);';
    const INACTIVE_PILL_STYLE = 'background:transparent;color:var(--text-muted);box-shadow:none;';

    function updatePillStyles() {
        const val = document.querySelector('input[name="subject-source"]:checked')?.value;
        if (pillExcel && pillManual) {
            pillExcel.style.cssText += val === 'excel' ? ACTIVE_PILL_STYLE : INACTIVE_PILL_STYLE;
            pillManual.style.cssText += val === 'manual' ? ACTIVE_PILL_STYLE : INACTIVE_PILL_STYLE;
        }
    }
    updatePillStyles(); // Set correct state on load

    subjectSourceRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            updatePillStyles();
            if (e.target.value === 'excel') {
                excelSourceDiv.classList.remove('hidden');
                manualSourceDiv.classList.add('hidden');
                subjectFile.required = true;
                if (subjectFile.files.length > 0) processSubjectExcel(subjectFile.files[0]);
            } else {
                excelSourceDiv.classList.add('hidden');
                manualSourceDiv.classList.remove('hidden');
                subjectFile.required = false;
                subjectFile.value = ''; // clear input
                // Populate from manual inputs on blur
                updateManualSubjects();
            }
        });
    });

    const manualInputs = ['manual-subj-i', 'manual-subj-ii', 'manual-subj-iii', 'manual-subj-iv'];
    manualInputs.forEach(id => {
        document.getElementById(id).addEventListener('blur', updateManualSubjects);
    });

    function updateManualSubjects() {
        if (document.querySelector('input[name="subject-source"]:checked').value !== 'manual') return;

        window.subjectCodes = {
            "I Yr": [], "II Yr": [], "III Yr": [], "IV Yr": []
        };

        const map = {
            'manual-subj-i': 'I Yr',
            'manual-subj-ii': 'II Yr',
            'manual-subj-iii': 'III Yr',
            'manual-subj-iv': 'IV Yr'
        };

        manualInputs.forEach(id => {
            const val = document.getElementById(id).value;
            if (val.trim()) {
                window.subjectCodes[map[id]] = val.split(',').map(s => s.trim()).filter(s => s !== '');
            }
        });

        // re-populate all dropdowns if active
        document.querySelectorAll('.subject-select').forEach(select => {
            const yr = select.dataset.year;
            populateDropdown(select, yr);
        });
    }

    // Department Selection Logic
    const departmentSelect = document.getElementById('department-name');
    const customDepartmentWrapper = document.getElementById('custom-department-wrapper');
    const customDepartmentInput = document.getElementById('custom-department-name');

    departmentSelect.addEventListener('change', (e) => {
        if (e.target.value === 'Custom') {
            customDepartmentWrapper.classList.remove('hidden');
            customDepartmentInput.required = true;
        } else {
            customDepartmentWrapper.classList.add('hidden');
            customDepartmentInput.required = false;
            customDepartmentInput.value = ''; // clear custom text
        }
    });

    // --- Timetable Builder Logic ---
    const datesContainer = document.getElementById('dates-container');
    const addDateBtn = document.getElementById('add-date-btn');
    let dateCount = 0;

    function populateDropdown(selectEl, year) {
        const val = selectEl.value;
        selectEl.innerHTML = '<option value="" disabled selected>Select Subject</option>';
        if (window.subjectCodes[year]) {
            window.subjectCodes[year].forEach(code => {
                const opt = document.createElement('option');
                opt.value = code;
                opt.textContent = code;
                selectEl.appendChild(opt);
            });
        }
        if (val) selectEl.value = val;
    }

    // Update global subject options to disable selected subjects in other dropdowns
    function updateSubjectOptions() {
        const allSelects = document.querySelectorAll('.subject-select');
        const selectedValues = new Set();

        allSelects.forEach(select => {
            const wrapper = select.closest('.input-wrapper');
            if (wrapper && !wrapper.classList.contains('hidden') && select.value) {
                selectedValues.add(select.value);
            }
        });

        allSelects.forEach(select => {
            const options = select.querySelectorAll('option');
            options.forEach(opt => {
                if (opt.value && opt.value !== select.value) {
                    opt.disabled = selectedValues.has(opt.value);
                }
            });
        });
    }

    function createShiftBlock(dateId, shiftCount, shiftsContainer) {
        const shiftId = `${dateId}-shift-${Date.now()}`;
        const shiftDiv = document.createElement('div');
        shiftDiv.className = 'shift-block';
        shiftDiv.style.background = 'rgba(0, 0, 0, 0.2)';
        shiftDiv.style.border = '1px solid var(--border-color)';
        shiftDiv.style.borderRadius = '8px';
        shiftDiv.style.padding = '1rem';
        shiftDiv.style.marginBottom = '1rem';
        shiftDiv.style.position = 'relative';

        shiftDiv.innerHTML = `
            <button type="button" class="btn text-btn remove-shift-btn" style="position: absolute; top:0.5rem; right:0.5rem; color: var(--error-color); padding: 5px;">
                <i class="fa-solid fa-xmark"></i>
            </button>
            <h4 style="margin-bottom: 1rem; font-size: 1rem; color: var(--text-main);">Shift Timing</h4>
            
            <div style="display: flex; gap: 0.75rem; margin-bottom: 1.5rem; align-items: flex-end;">
                <div style="flex: 1;">
                    <label style="font-size: 0.8rem; color: var(--text-muted); display: block; margin-bottom: 5px;"><i class="fa-regular fa-clock" style="margin-right:3px;"></i>Start Time</label>
                    <div style="display: flex; gap: 5px;">
                        <select class="shift-start-hr" required style="flex:1; padding: 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main);">
                            <option value="" disabled selected>HH</option>
                            ${Array.from({ length: 12 }, (_, i) => `<option value="${String(i + 1).padStart(2, '0')}">${String(i + 1).padStart(2, '0')}</option>`).join('')}
                        </select>
                        <select class="shift-start-min" required style="flex:1; padding: 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main);">
                            <option value="" disabled selected>MM</option>
                            ${Array.from({ length: 60 }, (_, i) => `<option value="${String(i).padStart(2, '0')}">${String(i).padStart(2, '0')}</option>`).join('')}
                        </select>
                        <select class="shift-start-ampm" required style="flex:1; padding: 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main);">
                            <option value="AM" selected>AM</option>
                            <option value="PM">PM</option>
                        </select>
                    </div>
                </div>
                <div style="color: var(--text-muted); font-size: 1.2rem; padding-bottom: 0.4rem; flex-shrink: 0;">→</div>
                <div style="flex: 1;">
                    <label style="font-size: 0.8rem; color: var(--text-muted); display: block; margin-bottom: 5px;"><i class="fa-regular fa-clock" style="margin-right:3px;"></i>End Time</label>
                    <div style="display: flex; gap: 5px;">
                        <select class="shift-end-hr" required style="flex:1; padding: 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main);">
                            <option value="" disabled selected>HH</option>
                            ${Array.from({ length: 12 }, (_, i) => `<option value="${String(i + 1).padStart(2, '0')}">${String(i + 1).padStart(2, '0')}</option>`).join('')}
                        </select>
                        <select class="shift-end-min" required style="flex:1; padding: 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main);">
                            <option value="" disabled selected>MM</option>
                            ${Array.from({ length: 60 }, (_, i) => `<option value="${String(i).padStart(2, '0')}">${String(i).padStart(2, '0')}</option>`).join('')}
                        </select>
                        <select class="shift-end-ampm" required style="flex:1; padding: 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main);">
                            <option value="AM">AM</option>
                            <option value="PM" selected>PM</option>
                        </select>
                    </div>
                </div>
            </div>
            
            <h4 style="margin-bottom: 0.5rem; font-size: 0.9rem; color: var(--text-muted);"><i class="fa-solid fa-users" style="margin-right:5px;"></i>Participating Years</h4>
            <div class="years-container" style="display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem;">
                <!-- Dynamically added year rows here... -->
            </div>
        `;

        const yearsContainer = shiftDiv.querySelector('.years-container');
        const years = ["I Yr", "II Yr", "III Yr", "IV Yr"];

        years.forEach((yr, idx) => {
            const yrId = `${shiftId}-yr-${idx}`;
            const yrRow = document.createElement('div');
            yrRow.style.display = 'flex';
            yrRow.style.flexDirection = 'column';
            yrRow.style.gap = '0.4rem';
            yrRow.style.background = 'rgba(0,0,0,0.15)';
            yrRow.style.border = '1px solid var(--border-color)';
            yrRow.style.borderRadius = '8px';
            yrRow.style.padding = '0.5rem 0.75rem';

            yrRow.innerHTML = `
                <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer; font-size: 0.95rem; font-weight: 500; color: var(--text-main);">
                    <input type="checkbox" class="year-checkbox custom-checkbox" data-year="${yr}"> ${yr}
                </label>
                <div class="input-wrapper subject-wrapper hidden" style="margin-bottom: 0;">
                    <select class="subject-select" data-year="${yr}" style="width: 100%; padding: 0.4rem 0.5rem; border-radius: 8px; background: var(--bg-card); border: 1px solid var(--border-color); color: var(--text-main); font-size: 0.85rem;">
                        <option value="" disabled selected>Select Subject</option>
                    </select>
                </div>
            `;

            yearsContainer.appendChild(yrRow);

            const cb = yrRow.querySelector('.year-checkbox');
            const subWrapper = yrRow.querySelector('.subject-wrapper');
            const subSelect = yrRow.querySelector('.subject-select');

            populateDropdown(subSelect, yr);

            cb.addEventListener('change', (e) => {
                if (e.target.checked) {
                    subWrapper.classList.remove('hidden');
                    subSelect.required = true;
                } else {
                    subWrapper.classList.add('hidden');
                    subSelect.required = false;
                    subSelect.value = "";
                }
                updateSubjectOptions();
            });

            subSelect.addEventListener('change', updateSubjectOptions);
        });

        shiftDiv.querySelector('.remove-shift-btn').addEventListener('click', () => {
            shiftDiv.remove();
            updateSubjectOptions();
        });

        // --- Auto-fill End Time from Start Time + Global Exam Duration ---
        const autoFillEndTime = () => {
            const durHr = parseInt(document.getElementById('exam-dur-hr')?.value || '0', 10);
            const durMin = parseInt(document.getElementById('exam-dur-min')?.value || '0', 10);
            const totalDurMins = durHr * 60 + durMin;
            if (totalDurMins === 0) return; // No duration set, skip

            const sHr = shiftDiv.querySelector('.shift-start-hr').value;
            const sMin = shiftDiv.querySelector('.shift-start-min').value;
            const sAmPm = shiftDiv.querySelector('.shift-start-ampm').value;
            if (!sHr || !sMin || !sAmPm) return; // Start time incomplete

            // Convert start to 24h minutes
            let startH24 = parseInt(sHr, 10) % 12; // 12 AM/PM → 0
            if (sAmPm === 'PM') startH24 += 12;
            const startTotalMins = startH24 * 60 + parseInt(sMin, 10);

            // Add duration
            const endTotalMins = startTotalMins + totalDurMins;
            const endH24 = Math.floor(endTotalMins / 60) % 24;
            const endMin = endTotalMins % 60;

            // Convert back to 12h
            const endAmPm = endH24 >= 12 ? 'PM' : 'AM';
            const endH12 = endH24 % 12 || 12;

            const endHrStr = String(endH12).padStart(2, '0');
            const endMinStr = String(endMin).padStart(2, '0');

            // Set end-time dropdowns
            shiftDiv.querySelector('.shift-end-hr').value = endHrStr;
            shiftDiv.querySelector('.shift-end-min').value = endMinStr;
            shiftDiv.querySelector('.shift-end-ampm').value = endAmPm;
        };

        // Trigger auto-fill whenever any start-time dropdown changes
        shiftDiv.querySelector('.shift-start-hr').addEventListener('change', autoFillEndTime);
        shiftDiv.querySelector('.shift-start-min').addEventListener('change', autoFillEndTime);
        shiftDiv.querySelector('.shift-start-ampm').addEventListener('change', autoFillEndTime);

        shiftsContainer.appendChild(shiftDiv);
    }

    function createDateBlock() {
        dateCount++;
        const dateId = `date-${Date.now()}`;
        const div = document.createElement('div');
        div.className = 'date-block slide-up delay-1';
        div.style.background = 'var(--bg-card)';
        div.style.padding = '1.5rem';
        div.style.borderRadius = '12px';
        div.style.border = '2px solid var(--border-color)';
        div.style.position = 'relative';

        div.innerHTML = `
            <button type="button" class="btn text-btn remove-date-btn" style="position: absolute; top:1rem; right:1rem; color: var(--error-color);">
                <i class="fa-solid fa-trash"></i> Remove Date
            </button>
            <h3 style="margin-bottom: 1.5rem; color: var(--primary-color); border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem;">
                 <i class="fa-regular fa-calendar" style="margin-right: 8px;"></i> Exam Date Configuration
            </h3>
            
            <div class="input-wrapper floating-label" style="max-width: 300px; margin-bottom: 1.5rem;">
                <input type="date" class="date-input" required>
                <label style="top: -10px; background: var(--bg-card); font-size: 0.8rem">Select Date</label>
            </div>
            
            <div class="shifts-container">
                <!-- Shifts go here -->
            </div>
            
            <button type="button" class="btn outline-btn add-shift-btn" style="margin-top: 0.5rem; font-size: 0.9rem; padding: 0.5rem 1rem;">
                <i class="fa-solid fa-plus"></i> Add a Shift to this Date
            </button>
        `;

        const shiftsContainer = div.querySelector('.shifts-container');
        let shiftCount = 0;

        div.querySelector('.add-shift-btn').addEventListener('click', () => {
            shiftCount++;
            createShiftBlock(dateId, shiftCount, shiftsContainer);
        });

        div.querySelector('.remove-date-btn').addEventListener('click', () => {
            div.remove();
            updateSubjectOptions();
        });

        // Add first shift by default
        shiftCount++;
        createShiftBlock(dateId, shiftCount, shiftsContainer);

        datesContainer.appendChild(div);
    }

    if (addDateBtn) {
        addDateBtn.addEventListener('click', createDateBlock);
        // Create first block by default
        createDateBlock();
    }

    // --- Generate Setup Configuration ---
    generateRoomsBtn.addEventListener('click', () => {
        const count = parseInt(roomCountInput.value);

        if (isNaN(count) || count < 1 || count > 50) {
            setupError.classList.remove('hidden');
            return;
        }

        setupError.classList.add('hidden');
        initializeRooms(count);
        mainContent.classList.remove('hidden');

        // Scroll to batch section
        setTimeout(() => {
            document.getElementById('batch-apply-section').scrollIntoView({ behavior: 'smooth' });
        }, 100);
    });

    // --- Room Name Auto-Increment Logic ---
    // Parse a room name string into { prefix (letters at start), number (trailing digits), suffix (any trailing non-digits) }
    function parseRoomName(name) {
        // Match: optional leading letters, then digits at the end
        const match = name.match(/^([A-Za-z]*)([0-9]+)([^0-9]*)$/);
        if (match) {
            return {
                prefix: match[1],
                number: parseInt(match[2], 10),
                numWidth: match[2].length, // preserve original digit width (e.g. "003" → width 3)
                suffix: match[3],
                parsed: true
            };
        }
        return { parsed: false };
    }

    // Called when a room name input is changed — cascades incremented names to all following rooms
    window.cascadeRoomNames = function (changedId) {
        const changedInput = document.getElementById(`room-name-input-${changedId}`);
        if (!changedInput) return;
        const changedName = changedInput.value.trim();
        const parsed = parseRoomName(changedName);
        if (!parsed.parsed) return; // Not a parseable numbered format; skip cascade

        // Find the index of the changed room in roomData
        const startIdx = roomData.findIndex(r => r.id === changedId);
        if (startIdx === -1) return;

        // Update all subsequent rooms, preserving leading-zero width
        const numWidth = parsed.numWidth || 1;
        let nextNum = parsed.number + 1;
        for (let i = startIdx + 1; i < roomData.length; i++) {
            const room = roomData[i];
            const paddedNum = String(nextNum).padStart(numWidth, '0');
            const newName = `${parsed.prefix}${paddedNum}${parsed.suffix}`;

            // Update the input
            const input = document.getElementById(`room-name-input-${room.id}`);
            if (input) input.value = newName;

            // Update the checkbox label
            const checkboxLabel = batchRoomCheckboxes.querySelector(`input[value="${room.id}"]`);
            if (checkboxLabel) {
                const labelText = checkboxLabel.closest('label').querySelector('.label-text');
                if (labelText) labelText.textContent = newName;
            }

            nextNum++;
        }
    };

    function initializeRooms(count) {
        roomData = [];
        roomsContainer.innerHTML = '';
        batchRoomCheckboxes.innerHTML = '';

        for (let i = 1; i <= count; i++) {
            const roomName = `N${100 + i}`;

            // Add to data structure
            roomData.push({
                id: i,
                name: roomName,
                rows: null,
                cols: null
            });

            // 1. Create Checkbox for Batch Section
            createBatchCheckbox(i, roomName);

            // 2. Create Room Card
            createRoomCard(i, roomName);
        }
    }

    // --- Room Cards & Visualizations ---
    function createRoomCard(id, name) {
        const card = document.createElement('div');
        card.className = 'room-card';
        card.dataset.roomId = id;

        card.innerHTML = `
            <div class="room-card-header">
                <div class="room-card-title" style="display:flex; align-items:center;">
                    <i class="fa-solid fa-chalkboard-user"></i> 
                    <input type="text" id="room-name-input-${id}" value="${name}" 
                           style="background:transparent; border:none; border-bottom: 2px dashed rgba(255,255,255,0.2); 
                                  color:var(--text-main); font-family:inherit; font-size:inherit; font-weight:bold; 
                                  width: 140px; outline:none; margin-left:8px; padding-bottom: 2px; transition: border-color 0.2s;"
                           onfocus="this.style.borderBottomColor='var(--primary-color)'"
                           onblur="if(!this.value.trim()) this.value='${name}'; this.style.borderBottomColor='rgba(255,255,255,0.2)'; cascadeRoomNames(${id});">
                </div>
                <div class="room-stats"><span id="stats-${id}">0</span> Seats</div>
            </div>
            
            <div class="room-config-row">
                <div class="input-wrapper floating-label">
                    <input type="number" id="rows-${id}" min="1" placeholder=" " required>
                    <label for="rows-${id}">Rows</label>
                    <i class="fa-solid fa-arrows-up-down input-icon"></i>
                </div>
                <div class="input-wrapper floating-label">
                    <input type="number" id="cols-${id}" min="1" placeholder=" " required>
                    <label for="cols-${id}">Columns</label>
                    <i class="fa-solid fa-arrows-left-right input-icon"></i>
                </div>
            </div>
            
            <div class="room-config-row" style="margin-bottom: 1rem;">
            <div class="input-wrapper floating-label" style="flex: 1;">
                <select id="pattern-${id}" required style="width: 100%; background: rgba(0, 0, 0, 0.2); border: 1px solid var(--border-color); border-radius: 12px; padding: 1.5rem 1rem 0.5rem 3rem; color: var(--text-main); font-size: 1.1rem; font-family: 'Outfit', sans-serif; appearance: none;">
                    <option value="IV Yr, III Yr, II Yr, I Yr" selected>IV Yr, III Yr, II Yr, I Yr</option>
                    <option value="I Yr, II Yr, III Yr, IV Yr">I Yr, II Yr, III Yr, IV Yr</option>
                    <option value="IV Yr, II Yr, III Yr, I Yr">IV Yr, II Yr, III Yr, I Yr</option>
                    <option value="III Yr, I Yr, IV Yr, II Yr">III Yr, I Yr, IV Yr, II Yr</option>
                    <option value="II Yr, IV Yr, I Yr, III Yr">II Yr, IV Yr, I Yr, III Yr</option>
                    <option value="IV Yr, I Yr, III Yr, II Yr">IV Yr, I Yr, III Yr, II Yr</option>
                    <option value="III Yr, II Yr, IV Yr, I Yr">III Yr, II Yr, IV Yr, I Yr</option>
                    <option value="II Yr, III Yr, I Yr, IV Yr">II Yr, III Yr, I Yr, IV Yr</option>
                </select>
                <i class="fa-solid fa-chevron-down" style="position: absolute; right: 1rem; top: 50%; transform: translateY(-50%); color: var(--text-muted); pointer-events: none;"></i>
                <label for="pattern-${id}">Seating Pattern</label>
                <i class="fa-solid fa-list-ol input-icon"></i>
            </div>
        </div>
        
        <div class="room-config-row" style="margin-bottom: 1rem;">
            <div class="input-wrapper floating-label">
                <select id="door-${id}">
                    <option value="top-right" selected>Top Right</option>
                        <option value="top-left">Top Left</option>
                        <option value="bottom-right">Bottom Right</option>
                        <option value="bottom-left">Bottom Left</option>
                    </select>
                    <label for="door-${id}">Entry Door</label>
                    <i class="fa-solid fa-door-open input-icon"></i>
                </div>
            </div>
            
            <button class="btn outline-btn update-room-btn" onclick="updateRoomPreview(${id})">
                <i class="fa-solid fa-rotate"></i> Update Preview
            </button>
            
            <div class="preview-wrapper">
                <div id="preview-${id}" class="grid-preview empty-state">
                    <i class="fa-solid fa-border-all"></i>
                    <span>Enter dimensions to see preview</span>
                </div>
            </div>
            `;

        roomsContainer.appendChild(card);

        // Add direct input listeners for instant feel (optional, but requested logic uses button)
        document.getElementById(`rows-${id}`).addEventListener('change', () => updateRoomPreview(id));
        document.getElementById(`cols-${id}`).addEventListener('change', () => updateRoomPreview(id));
    }

    // Expose update function to window for the inline onclick handler
    window.updateRoomPreview = function (id) {
        const rowsInput = document.getElementById(`rows-${id}`);
        const colsInput = document.getElementById(`cols-${id}`);
        const previewGrid = document.getElementById(`preview-${id}`);
        const statsEl = document.getElementById(`stats-${id}`);

        const rows = parseInt(rowsInput.value);
        const cols = parseInt(colsInput.value);

        if (isNaN(rows) || isNaN(cols) || rows < 1 || cols < 1) {
            return; // Invalid input
        }

        // Update data array
        const roomIndex = roomData.findIndex(r => r.id === id);
        if (roomIndex > -1) {
            roomData[roomIndex].rows = rows;
            roomData[roomIndex].cols = cols;
        }

        // Update Stats
        statsEl.textContent = rows * cols;

        // Remove empty state
        previewGrid.classList.remove('empty-state');
        previewGrid.innerHTML = '';

        // CSS Grid dynamic configuration
        previewGrid.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;

        // Generate seats
        // Cap preview visualization to prevent browser lag (e.g. max 400 seats mapped visually)
        const cellCount = Math.min(rows * cols, 400);

        for (let i = 0; i < cellCount; i++) {
            const seat = document.createElement('div');
            seat.className = 'seat-cell';
            previewGrid.appendChild(seat);
        }

        if (rows * cols > 400) {
            const notice = document.createElement('div');
            notice.style.gridColumn = `1 / -1`;
            notice.style.textAlign = 'center';
            notice.style.fontSize = '0.8rem';
            notice.style.color = 'var(--text-muted)';
            notice.style.marginTop = '10px';
            notice.textContent = `+ ${rows * cols - 400} more seats(preview limited)`;
            previewGrid.appendChild(notice);
        }
    };

    // --- Batch Configuration ---
    function createBatchCheckbox(id, name) {
        const wrapper = document.createElement('label');
        wrapper.className = 'custom-checkbox';
        wrapper.innerHTML = `
                <input type="checkbox" class="room-checkbox" value="${id}">
            <div class="checkmark"></div>
            <span class="label-text">${name}</span>
            `;
        batchRoomCheckboxes.appendChild(wrapper);
    }

    selectAllBtn.addEventListener('click', () => {
        document.querySelectorAll('.room-checkbox').forEach(cb => cb.checked = true);
    });

    deselectAllBtn.addEventListener('click', () => {
        document.querySelectorAll('.room-checkbox').forEach(cb => cb.checked = false);
    });

    applyBatchBtn.addEventListener('click', () => {
        const rows = parseInt(batchRowsInput.value);
        const cols = parseInt(batchColsInput.value);
        const pattern = document.getElementById('batch-pattern').value;
        const selectedCheckboxes = document.querySelectorAll('.room-checkbox:checked');

        if (isNaN(rows) || rows < 1 || isNaN(cols) || cols < 1 || !pattern || selectedCheckboxes.length === 0) {
            batchError.classList.remove('hidden');
            setTimeout(() => batchError.classList.add('hidden'), 3000);
            return;
        }

        batchError.classList.add('hidden');

        // Apply to selected rooms
        selectedCheckboxes.forEach(cb => {
            const id = parseInt(cb.value);

            // Update individual inputs
            document.getElementById(`rows-${id}`).value = rows;
            document.getElementById(`cols-${id}`).value = cols;
            document.getElementById(`pattern-${id}`).value = pattern;

            // Trigger visual update
            updateRoomPreview(id);
        });

        // Flash effect on button to confirm action
        const originalText = applyBatchBtn.innerHTML;
        applyBatchBtn.innerHTML = '<i class="fa-solid fa-check"></i> Applied Successfully';
        applyBatchBtn.style.background = 'linear-gradient(135deg, var(--success-color), #059669)';

        setTimeout(() => {
            applyBatchBtn.innerHTML = originalText;
            applyBatchBtn.style.background = '';
        }, 2000);
    });


    // --- Final Generation & API Call ---
    function getCSRFToken() {
        let cookieValue = null;
        if (document.cookie && document.cookie !== '') {
            const cookies = document.cookie.split(';');
            for (let i = 0; i < cookies.length; i++) {
                const cookie = cookies[i].trim();
                // Does this cookie string begin with the name we want?
                if (cookie.substring(0, 'csrftoken'.length + 1) === ('csrftoken' + '=')) {
                    cookieValue = decodeURIComponent(cookie.substring('csrftoken'.length + 1));
                    break;
                }
            }
        }
        return cookieValue;
    }

    finalGenerateBtn.addEventListener('click', async () => {
        finalError.classList.add('hidden');
        finalError.textContent = '';

        // 1. Gather all data
        const studentFile = document.getElementById('student-file').files[0];
        const subjectFile = document.getElementById('subject-file').files[0];

        const departmentNameSelect = document.getElementById('department-name').value;
        let departmentName;
        if (departmentNameSelect === 'Custom') {
            departmentName = document.getElementById('custom-department-name').value.trim();
            if (!departmentName) {
                finalError.textContent = 'Please enter a valid Custom Department Name.';
                finalError.classList.remove('hidden');
                return;
            }
        } else {
            departmentName = departmentNameSelect;
        }

        if (!studentFile) {
            finalError.textContent = 'Please upload a Student Excel file First.';
            finalError.classList.remove('hidden');
            document.getElementById('config-section').scrollIntoView({ behavior: 'smooth' });
            return;
        }

        if (!departmentName) { // This check is now for departmentName
            finalError.textContent = 'Please enter a Department Name.';
            finalError.classList.remove('hidden');
            return;
        }

        if (!subjectFile && Object.values(window.subjectCodes).every(arr => arr.length === 0)) {
            finalError.textContent = 'Please upload a Subject Codes Excel file first to populate subjects.';
            finalError.classList.remove('hidden');
            return;
        }

        // Validate Timetable Dates
        const scheduleConfig = [];
        let validationFailed = false;

        // Clear previous highlights
        document.querySelectorAll('.error-highlight').forEach(el => el.classList.remove('error-highlight'));
        const globalDates = new Set();
        const globalSubjects = new Set();
        let scrollTarget = null;

        const markError = (element) => {
            if (element) {
                element.classList.add('error-highlight');
                if (!scrollTarget) scrollTarget = element;
            }
        };

        document.querySelectorAll('.date-block').forEach(dateBlock => {
            const dateInput = dateBlock.querySelector('.date-input');
            const dateVal = dateInput.value;

            if (!dateVal) {
                validationFailed = true;
                markError(dateInput);
            } else if (globalDates.has(dateVal)) {
                // Must be unique date
                validationFailed = true;
                markError(dateInput);
                finalError.innerHTML = `Duplicate Exam Date found: <b>${dateVal}</b>. Please consolidate shifts under a single Date Block.`;
            } else {
                globalDates.add(dateVal);
            }

            const shifts = [];
            dateBlock.querySelectorAll('.shift-block').forEach(shiftBlock => {
                const sHr = shiftBlock.querySelector('.shift-start-hr').value;
                const sMin = shiftBlock.querySelector('.shift-start-min').value;
                const sAmPm = shiftBlock.querySelector('.shift-start-ampm').value;

                const eHr = shiftBlock.querySelector('.shift-end-hr').value;
                const eMin = shiftBlock.querySelector('.shift-end-min').value;
                const eAmPm = shiftBlock.querySelector('.shift-end-ampm').value;

                let timeStr = "";

                if (!sHr || !sMin || !sAmPm) { validationFailed = true; markError(shiftBlock.querySelector('.shift-start-hr')); }
                if (!eHr || !eMin || !eAmPm) { validationFailed = true; markError(shiftBlock.querySelector('.shift-end-hr')); }

                if (sHr && sMin && sAmPm && eHr && eMin && eAmPm) {
                    timeStr = `${parseInt(sHr, 10)}:${sMin} ${sAmPm} - ${parseInt(eHr, 10)}:${eMin} ${eAmPm}`;
                }

                const participatingYears = [];
                let yearSelected = false;

                shiftBlock.querySelectorAll('.year-checkbox').forEach(cb => {
                    if (cb.checked) {
                        yearSelected = true;
                        const yr = cb.dataset.year;
                        const subjectSelect = shiftBlock.querySelector(`.subject-select[data-year="${yr}"]`);
                        const subject = subjectSelect.value;

                        if (!subject) {
                            validationFailed = true;
                            markError(subjectSelect);
                        } else if (globalSubjects.has(subject)) {
                            validationFailed = true;
                            markError(subjectSelect);
                            finalError.innerHTML = `Duplicate Subject found: <b>${subject}</b>. Subjects can only be scheduled once.`;
                        } else {
                            globalSubjects.add(subject);
                            participatingYears.push({
                                year: yr,
                                subject: subject
                            });
                        }
                    }
                });

                if (!yearSelected) {
                    validationFailed = true;
                    markError(shiftBlock.querySelector('.years-container'));
                }

                shifts.push({
                    time: timeStr,
                    years: participatingYears
                });
            });

            if (shifts.length === 0) {
                validationFailed = true;
                markError(dateBlock);
            }

            scheduleConfig.push({
                date: dateVal,
                shifts: shifts
            });
        });

        if (scheduleConfig.length === 0) {
            finalError.textContent = 'Please add at least one Exam Date with a Shift.';
            finalError.classList.remove('hidden');
            return;
        }

        if (validationFailed) {
            if (!finalError.innerHTML.includes('Duplicate')) {
                finalError.textContent = 'Please fill all highlighted mandatory fields correctly.';
            }
            finalError.classList.remove('hidden');
            if (scrollTarget) {
                scrollTarget.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
            return;
        }

        // Validate that all initialized rooms have a configuration
        const payloadRooms = [];
        let missingConfig = false;

        for (const room of roomData) {
            // grab latest values from inputs in case preview wasn't explicitly clicked
            const nameEl = document.getElementById(`room-name-input-${room.id}`);
            const updatedRoomName = nameEl ? nameEl.value.trim() : room.name;

            const rowsInput = document.getElementById(`rows-${room.id}`).value;
            const colsInput = document.getElementById(`cols-${room.id}`).value;
            const doorInput = document.getElementById(`door-${room.id}`) ? document.getElementById(`door-${room.id}`).value : 'top-right';
            const patternInput = document.getElementById(`pattern-${room.id}`).value;

            if (!rowsInput || !colsInput || parseInt(rowsInput) < 1 || parseInt(colsInput) < 1 || !patternInput) {
                missingConfig = true;
                break;
            }

            payloadRooms.push({
                name: updatedRoomName || room.name,
                rows: rowsInput ? parseInt(rowsInput) : 0,
                cols: colsInput ? parseInt(colsInput) : 0,
                door: doorInput,
                seating_pattern: patternInput
            });
        }

        if (missingConfig) {
            finalError.textContent = 'Please make sure all rooms have valid Rows, Columns, and Patterns assigned.';
            finalError.classList.remove('hidden');
            return;
        }

        const isManualObjectsMode = document.querySelector('input[name="subject-source"]:checked').value === 'manual';

        // 2. Prepare FormData
        const formData = new FormData();
        if (!isManualObjectsMode) {
            formData.append('student_file', studentFile);
        } else {
            // Need to still pass student file, but maybe tell backend which mode we are in if we want to bypass subject logic there.
            // Actually, backend only reads Student File (enrolment numbers). Subject codes are passed via `schedule_config` which is generated by the UI dropdowns.
            // The python backend DOES NOT read subjects from excel, only students! So this form payload doesn't need to change for backend, UI already handled `subjectCodes` injection!
            formData.append('student_file', studentFile);
        }

        formData.append('branch_name', document.getElementById('department-name').value);
        formData.append('schedule_config', JSON.stringify(scheduleConfig));
        formData.append('room_config', JSON.stringify(payloadRooms));

        const originalBtnHtml = finalGenerateBtn.innerHTML;
        finalGenerateBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating...';
        finalGenerateBtn.disabled = true;

        // 3. Send API Request
        try {
            const response = await fetch('/generate/', {
                method: 'POST',
                // CSRF Token header normally needed, but we used @csrf_exempt on the view for simplicity in this script setup. 
                // However, adding it is best practice if omitting decorator:
                // headers: { 'X-CSRFToken': getCSRFToken() },
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || 'Something went wrong on the server.');
            }

            // 4. Handle Success
            const instituteName = document.getElementById('institute-name').value || 'Institute Name';
            const instituteLogoFile = document.getElementById('institute-logo').files[0];

            const deptSelectVal = document.getElementById('department-name').value;
            const customDeptVal = document.getElementById('custom-department-name').value;
            const finalDepartmentName = (deptSelectVal === 'Custom' && customDeptVal.trim() !== '') ? customDeptVal.trim() : deptSelectVal;

            let instituteLogoBase64 = null;
            if (instituteLogoFile) {
                instituteLogoBase64 = await new Promise(resolve => {
                    const reader = new FileReader();
                    reader.onload = e => resolve(e.target.result);
                    reader.readAsDataURL(instituteLogoFile);
                });
            }

            renderResults(data, instituteName, instituteLogoBase64, finalDepartmentName);

        } catch (error) {
            finalError.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Error: ${error.message}`;
            finalError.classList.remove('hidden');
        } finally {
            finalGenerateBtn.innerHTML = originalBtnHtml;
            finalGenerateBtn.disabled = false;
        }
    });

    backToConfigBtn.addEventListener('click', () => {
        document.getElementById('output-content').classList.add('hidden');
        document.getElementById('config-section').classList.remove('hidden');
        document.getElementById('timetable-builder-section').classList.remove('hidden');
        document.getElementById('setup-section').classList.remove('hidden');
        document.getElementById('main-content').classList.remove('hidden');
    });

    // --- Export Functionality ---

    document.getElementById('export-excel-btn').addEventListener('click', () => {
        const wb = XLSX.utils.book_new();
        // Export just the active tab's tables
        const activeTab = document.querySelector('.tab-content:not(.hidden)');
        if (!activeTab) return;

        activeTab.classList.add('exporting');
        const containers = activeTab.querySelectorAll('.print-container');
        if (containers.length === 0) {
            activeTab.classList.remove('exporting');
            return;
        }

        containers.forEach((container, index) => {
            // Find preceding heading to name the sheet
            let sheetName = `Sheet ${index + 1}`;

            // Try to name the sheet intelligently based on the context
            const h3 = container.querySelector('h3');
            if (h3) sheetName = h3.textContent.replace('SEATING CHART', '').trim().substring(0, 30).replace(/[\\/?*[\]]/g, '');
            else {
                const h1 = container.querySelector('h1');
                if (h1 && h1.textContent !== 'IPS Academy, Institute of Engineering and Science') sheetName = h1.textContent.substring(0, 30).replace(/[\\/?*[\]]/g, '');
            }

            // Ensure unique names
            let count = 1;
            let finalName = sheetName;
            while (wb.SheetNames.includes(finalName)) {
                finalName = `${sheetName.substring(0, 25)} (${count})`;
                count++;
            }

            // Build array of rows manually
            const sheetData = [];

            // 1. Add Header Information (H1 and P from print-header)
            const headerDiv = container.querySelector('.print-header');
            if (headerDiv) {
                const headerH1 = headerDiv.querySelector('h1')?.textContent || '';
                const headerP = headerDiv.querySelector('p')?.textContent || '';
                if (headerH1) sheetData.push([headerH1]);
                if (headerP) sheetData.push([headerP]);
                sheetData.push([]); // blank row
            }

            // 2. Add sub-headers (H2, H3, H4) excluding the global ones already in print-header
            const headings = container.querySelectorAll('h2, h3, h4');
            headings.forEach(h => {
                if (!h.closest('.print-header')) {
                    sheetData.push([h.textContent.trim()]);
                }
            });

            // 3. Add Room Stats if they exist (from seating charts)
            const statsCards = container.querySelectorAll('.year-badge');
            if (statsCards.length > 0) {
                const statsRow = [];
                container.querySelectorAll('span').forEach(span => {
                    const text = span.textContent.trim();
                    if (text.includes('Yr:') || text.includes('Total')) {
                        statsRow.push(text);
                    }
                });
                if (statsRow.length > 0) {
                    sheetData.push(statsRow);
                    sheetData.push([]); // blank row
                }
            }

            // 4. Add the Table Data
            const table = container.querySelector('table');
            if (table) {
                // Manually parse table avoiding complex parsing
                const rows = table.querySelectorAll('tr');
                rows.forEach(tr => {
                    const rowData = [];
                    tr.querySelectorAll('th, td').forEach(cell => {
                        // Clean up text content and remove extra whitespaces
                        rowData.push(cell.textContent.trim().replace(/\s+/g, ' '));
                    });
                    sheetData.push(rowData);
                });
            }

            const ws = XLSX.utils.aoa_to_sheet(sheetData);

            // Basic styling - auto-width columns
            if (table) {
                const colWidths = [];
                table.querySelectorAll('tr:first-child th, tr:first-child td').forEach(cell => {
                    colWidths.push({ wch: Math.max(20, cell.textContent.length + 5) });
                });
                ws['!cols'] = colWidths;
            }

            XLSX.utils.book_append_sheet(wb, ws, finalName);
        });

        let filenameStr = 'Exam_Result.xlsx';
        if (activeTab.id === 'tab-timetable') filenameStr = 'Master_Timetable.xlsx';
        if (activeTab.id === 'tab-seating') filenameStr = 'Seating_Charts.xlsx';

        XLSX.writeFile(wb, filenameStr);
        activeTab.classList.remove('exporting');
    });

    // Global UI toggle functions for the nested tabs
    window.switchDateTab = function (btn, dateId, prefix) {
        document.querySelectorAll(`.${prefix}-date-panel`).forEach(el => el.classList.add('hidden'));
        document.querySelectorAll(`.${prefix}-date-btn`).forEach(el => el.classList.remove('active'));
        document.getElementById(`${prefix}-date-${dateId}`).classList.remove('hidden');
        btn.classList.add('active');
    };

    window.switchShiftTab = function (btn, shiftId, prefix) {
        document.querySelectorAll(`.${prefix}-shift-panel`).forEach(el => el.classList.add('hidden'));
        document.querySelectorAll(`.${prefix}-shift-btn`).forEach(el => el.classList.remove('active'));
        document.getElementById(`${prefix}-shift-${shiftId}`).classList.remove('hidden');
        btn.classList.add('active');
    };

    window.switchRoomTab = function (btn, roomId, prefix = 'attendance') {
        document.querySelectorAll(`.${prefix}-room-panel`).forEach(el => el.classList.add('hidden'));
        document.querySelectorAll(`.${prefix}-room-btn`).forEach(el => el.classList.remove('active'));
        document.getElementById(`${prefix}-room-${roomId}`).classList.remove('hidden');
        btn.classList.add('active');
    };

    function renderResults(data, instituteName, instituteLogoBase64, departmentName) {
        // Hide config, show results
        document.getElementById('config-section').classList.add('hidden');
        document.getElementById('timetable-builder-section').classList.add('hidden');
        document.getElementById('setup-section').classList.add('hidden');
        document.getElementById('main-content').classList.add('hidden');
        outputContent.classList.remove('hidden');

        const buildPrintHeader = (subtitle) => `
            <div class="print-header" style="display: flex; align-items: center; justify-content: center; gap: 2rem; margin-bottom: 2rem; border-bottom: 3px double var(--border-color); padding-bottom: 1rem;">
                ${instituteLogoBase64 ? `<img src="${instituteLogoBase64}" alt="Logo" style="max-height: 80px; max-width: 80px; object-fit: contain;">` : ''}
                <div style="text-align: center;">
                    <h1 style="font-size: 2rem; margin: 0; color: var(--text-main); text-transform: uppercase; font-family: 'Times New Roman', Times, serif;">${instituteName}</h1>
                    <p style="margin: 0.15rem 0 0 0; font-size: 1rem; color: var(--text-muted); font-weight: 500; text-transform: uppercase; letter-spacing: 0.03em;">Department of ${departmentName}</p>
                    <p style="margin: 0.2rem 0 0 0; font-size: 1.2rem; color: var(--text-muted); font-weight: 600; text-transform: uppercase;">${subtitle}</p>
                </div>
            </div>
        `;

        // Build Timetable Tab
        // Build Master Timetable Tab
        let masterHtml = `
        <div class="glass-card print-container portrait-table" style="margin-bottom: 3rem; overflow-x: auto; position: relative;">
            ${buildPrintHeader('Master Examination Timetable')}
            <table class="timetable-table" style="width: 100%; border-collapse: collapse;">
                <thead>
                    <tr>
                        <th style="min-width: 150px; border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">Date & Time</th>`;

        const allExpectedYears = ["IV Yr", "III Yr", "II Yr", "I Yr"].filter(y => {
            return data.master_timetable.some(entry => entry[y] && entry[y] !== '-');
        });

        allExpectedYears.forEach(y => {
            masterHtml += `<th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">${y}</th>`;
        });

        masterHtml += `</tr>
                </thead>
                <tbody>`;

        data.master_timetable.forEach(entry => {
            masterHtml += `
            <tr>
                <td style="border: 1px solid var(--border-color); padding: 12px;">
                    <div style="font-weight: 600; color: var(--text-main); font-size: 1.1rem;">${entry.date}</div>
                    <div style="color: var(--accent-color); font-size: 0.9rem; margin-top: 4px;"><i class="fa-regular fa-clock"></i> ${entry.shift}</div>
                </td>`;

            allExpectedYears.forEach(y => {
                const subj = entry[y];
                if (subj && subj !== '-') {
                    masterHtml += `<td style="border: 1px solid var(--border-color); padding: 12px;"><span class="subject-badge">${subj}</span></td>`;
                } else {
                    masterHtml += `<td style="border: 1px solid var(--border-color); padding: 12px;"><span style="color: var(--text-muted); opacity: 0.5;">-</span></td>`;
                }
            });

            masterHtml += `</tr>`;
        });

        masterHtml += `
                </tbody>
            </table>
        </div>`;
        document.getElementById('tab-timetable').innerHTML = masterHtml;

        // Build Seating Charts Tab
        let seatingHtml = `<h1 style="font-size: 2.5rem; color: var(--primary-color); text-align: center; margin-bottom: 2rem;">Seating Charts</h1>`;

        const consolidatedSeating = {};

        data.seating_plans.forEach(plan => {
            // Hash the matrix to easily identify identical seating arrangements in the same room
            const matrixHash = JSON.stringify(plan.matrix);
            const hashKey = `${plan.room_name}___${matrixHash}`;

            if (!consolidatedSeating[hashKey]) {
                consolidatedSeating[hashKey] = {
                    room_name: plan.room_name,
                    rows: plan.rows,
                    cols: plan.cols,
                    door: plan.door,
                    headers: plan.headers,
                    matrix: plan.matrix,
                    counts: plan.counts,
                    total_in_room: plan.total_in_room,
                    sessions: []
                };
            }
            // Add the date and shift to this seating variation
            consolidatedSeating[hashKey].sessions.push({ date: plan.date, shift: plan.shift });
        });

        // Preserve the exact room order the user entered (UI tile order) — no sorting
        const seatingRoomsList = [...new Set(data.seating_plans.map(p => p.room_name))];

        if (seatingRoomsList.length > 0) {
            // Create pill nav for Rooms (flat tab structure exactly like Attendance)
            seatingHtml += `<div class="pill-nav" style="display:flex; justify-content:center; gap:10px; margin-bottom:2rem; flex-wrap:wrap;">`;
            seatingRoomsList.forEach((room, i) => {
                const safeRoom = room.replace(/\s/g, '_');
                seatingHtml += `<button class="btn outline-btn tab-btn seating-room-btn ${i === 0 ? 'active' : ''}" onclick="switchRoomTab(this, '${safeRoom}', 'seating')">${room}</button>`;
            });
            seatingHtml += `</div>`;

            seatingRoomsList.forEach((room, i) => {
                const safeRoom = room.replace(/\s/g, '_');
                seatingHtml += `<div id="seating-room-${safeRoom}" class="seating-room-panel sub-panel ${i === 0 ? '' : 'hidden'}">`;

                const plansForRoom = Object.values(consolidatedSeating).filter(c => c.room_name === room);

                plansForRoom.forEach((plan, idx) => {
                    // Smart Orientation Logic
                    const orientationClass = plan.cols > plan.rows ? 'landscape-table' : 'portrait-table';

                    // Construct a compact single-line string of all dates/shifts sharing this layout
                    const sessionsHtml = plan.sessions.map(s => `<span style="white-space:nowrap;">${s.date} &nbsp;${s.shift}</span>`).join('<span style="margin:0 6px;opacity:0.5;">•</span>');

                    seatingHtml += `
            <div class="glass-card ${orientationClass} print-container" style="margin-bottom: 4rem; overflow-x: auto; position: relative;">
                                ${buildPrintHeader(`Seating Chart • ${plan.room_name}`)}
                                <div style="color: var(--text-main); margin-bottom: 0.75rem; text-align: center; font-size: 0.85rem; font-weight: 600; line-height: 1.4; display: flex; flex-wrap: wrap; justify-content: center; align-items: center; gap: 4px;">
                                    <i class="fa-regular fa-calendar" style="margin-right:4px; opacity:0.7;"></i>${sessionsHtml}
                                </div>
                                
                                <div style="position: relative; display: inline-block; width: 100%; margin-top: 1.5rem;">
                                    <div class="door-indicator ${plan.door}"><i class="fa-solid fa-door-open"></i> Entry</div>
                                    <table class="seating-table">
                                        <thead>
                                            <tr>
                                                ${plan.headers.map(h => `<th>${h}</th>`).join('')}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${plan.matrix.slice(1).map(row => `
                                                <tr>
                                                    ${row.map(cell => {
                        if (!cell.student) return `<td class="empty-seat"></td>`;

                        // Handle both old format (string) and new format (object)
                        const enrollmentStr = typeof cell.student === 'object' ? (cell.student.enrollment || '') : cell.student;
                        const nameStr = typeof cell.student === 'object' ? (cell.student.name || '') : '';

                        return `<td>
                                    <div class="student-id">${enrollmentStr}</div>
                                    <div class="year-badge ${cell.year.replace(' ', '-')}">${cell.year}</div>
                                </td>`;
                    }).join('')}
                                                </tr>
                                            `).join('')}
                                        </tbody>
                                    </table>
                                </div>
                                
                                <div class="stats-container" style="margin-top: 2rem; padding-top: 1.5rem; border-top: 1px solid var(--border-color);">
                                    <div style="display: flex; gap: 2rem; justify-content: center; flex-wrap: wrap;">
                                        ${Object.entries(plan.counts).map(([yr, count]) => `
                                            <div style="display: flex; align-items: center; gap: 0.5rem;">
                                                <div class="year-badge ${yr.replace(' ', '-')}"></div>
                                                <span style="color: var(--text-muted);">${yr}: <strong style="color: var(--text-main);">${count}</strong></span>
                                            </div>
                                        `).join('')}
                                    </div>
                                    <div style="text-align: center; margin-top: 1rem; color: var(--text-main); font-weight: 600;">
                                        Total Students in Room: <span style="color: var(--primary-color);">${plan.total_in_room} / ${plan.rows * plan.cols} Capacity</span>
                                    </div>
                                </div>
                            </div>
                    `;
                });

                seatingHtml += `</div>`; // Close room content
            });
        }
        document.getElementById('tab-seating').innerHTML = seatingHtml;

        // Build Master Attendance Tab
        let masterAttendanceHtml = `<h1 style="font-size: 2.5rem; color: var(--primary-color); text-align: center; margin-bottom: 2rem;">Master Attendance</h1>`;
        ['I Yr', 'II Yr', 'III Yr', 'IV Yr'].forEach(yr => {
            const students = data.attendance_data[yr] || [];
            const yearExamDates = data.exam_dates_map[yr] || [];

            // Only show master attendance for a year if there are students AND they actually have an exam scheduled
            if (students.length > 0 && yearExamDates.length > 0) {
                masterAttendanceHtml += `
            <div class="glass-card print-container portrait-table" style="margin-bottom: 3rem;">
                    ${buildPrintHeader(`GLOBAL ATTENDANCE • YEAR: ${yr}`)}
                    <table class="attendance-table" style="width: 100%; border-collapse: collapse;">
                        <thead>
                            <tr>
                                <th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">S.No</th>
                                <th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">Enrollment No</th>
                                <th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">Student Name</th>
                `;

                if (yearExamDates.length > 0) {
                    yearExamDates.forEach(dateLabel => {
                        masterAttendanceHtml += `<th>Sign <br><span style="font-size: 0.8rem; font-weight: normal">${dateLabel}</span></th>`;
                    });
                } else {
                    masterAttendanceHtml += `<th>Signature</th>`;
                }

                masterAttendanceHtml += `
                            </tr>
                        </thead>
                        <tbody>
                            ${students.map((stu, idx) => {
                    const enrollmentStr = typeof stu === 'object' ? (stu.enrollment || '') : stu;
                    const nameStr = typeof stu === 'object' ? (stu.name || '') : '';

                    return `
                                <tr>
                                    <td style="border: 1px solid var(--border-color); padding: 12px;">${idx + 1}</td>
                                    <td style="font-size: 10pt; font-weight: bold; border: 1px solid var(--border-color); padding: 12px;">
                                        ${enrollmentStr}
                                    </td>
                                    <td style="font-size: 10pt; border: 1px solid var(--border-color); padding: 12px;">
                                        ${nameStr}
                                    </td>
                                    ${yearExamDates.length > 0 ?
                            yearExamDates.map(() => `<td></td>`).join('') :
                            `<td></td>`
                        }
                                </tr>
                                `;
                }).join('')}
                        </tbody>
                    </table>
                </div >
            `;
            }
        });
        document.getElementById('tab-attendance-master').innerHTML = masterAttendanceHtml;

        // Build Room-Wise Attendance Tab
        // Build Room-Wise Attendance Tab
        let roomAttendanceHtml = `<h1 style="font-size: 2.5rem; color: var(--primary-color); text-align: center; margin-bottom: 2rem;">Room - Wise Attendance</h1>`;

        const consolidatedAttendance = {};

        if (data.room_attendance_data && data.room_attendance_data.length > 0) {
            data.room_attendance_data.forEach(roomSheet => {
                const studentsByYear = {};
                roomSheet.students.forEach(stu => {
                    if (!studentsByYear[stu.year]) studentsByYear[stu.year] = [];
                    studentsByYear[stu.year].push(stu);
                });

                Object.keys(studentsByYear).forEach(yr => {
                    const studentList = studentsByYear[yr];
                    if (studentList.length === 0) return;

                    // Create a deterministic hash string
                    const enrollments = studentList.map(s => typeof s === 'object' ? (s.enrollment || '') : s).join(',');
                    const hashKey = `${roomSheet.room_name}___${yr}___${enrollments}`;

                    if (!consolidatedAttendance[hashKey]) {
                        consolidatedAttendance[hashKey] = {
                            room_name: roomSheet.room_name,
                            year: yr,
                            students: studentList,
                            sessions: []
                        };
                    }
                    // Add this date/shift to the sessions array
                    consolidatedAttendance[hashKey].sessions.push({ date: roomSheet.date, shift: roomSheet.shift });
                });
            });

            // Preserve the exact room order the user entered (UI tile order) — no sorting
            const roomsList = [...new Set(data.room_attendance_data.map(r => r.room_name))];

            // Create pill nav for Rooms (flat tab structure)
            roomAttendanceHtml += `<div class="pill-nav" style="display:flex; justify-content:center; gap:10px; margin-bottom:2rem; flex-wrap:wrap;">`;
            roomsList.forEach((room, i) => {
                const safeRoom = room.replace(/\s/g, '_');
                roomAttendanceHtml += `<button class="btn outline-btn tab-btn attendance-room-btn ${i === 0 ? 'active' : ''}" onclick="switchRoomTab(this, '${safeRoom}')">${room}</button>`;
            });
            roomAttendanceHtml += `</div>`;

            const yearOrder = ['I Yr', 'II Yr', 'III Yr', 'IV Yr'];

            roomsList.forEach((room, i) => {
                const safeRoom = room.replace(/\s/g, '_');
                roomAttendanceHtml += `<div id="attendance-room-${safeRoom}" class="attendance-room-panel sub-panel ${i === 0 ? '' : 'hidden'}">`;

                const sheetsForRoom = Object.values(consolidatedAttendance).filter(c => c.room_name === room);
                sheetsForRoom.sort((a, b) => yearOrder.indexOf(a.year) - yearOrder.indexOf(b.year));

                sheetsForRoom.forEach(sheet => {
                    let subtitle = `ROOM ATTENDANCE • ${sheet.room_name} • ${sheet.year}`;

                    roomAttendanceHtml += `
            <div class="glass-card print-container portrait-table" style="margin-bottom: 3rem; overflow-x: auto; position: relative;">
                                        ${buildPrintHeader(subtitle)}
                                        <table class="attendance-table" style="width: 100%; border-collapse: collapse;">
                                            <thead>
                                                <tr>
                                                    <th style="width: 80px; border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">S.No</th>
                                                    <th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">Enrollment No</th>
                                                    <th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">Student Name</th>
                                                    ${sheet.sessions.map(s => `<th style="border: 1px solid var(--border-color); padding: 12px; background: rgba(0,0,0,0.05);">Sign<br><span style="font-size: 0.8rem; font-weight: normal;">${s.date}<br>${s.shift}</span></th>`).join('')}
                                                </tr>
                                            </thead>
                                            <tbody>
                                                ${sheet.students.map((stu, idx) => {
                        const enrollmentStr = typeof stu === 'object' ? (stu.enrollment || '') : stu;
                        const nameStr = typeof stu === 'object' ? (stu.name || '') : '';

                        return `
                                                    <tr>
                                                        <td style="border: 1px solid var(--border-color); padding: 12px;">${idx + 1}</td>
                                                        <td style="border: 1px solid var(--border-color); padding: 12px; font-size: 10pt; font-weight: bold;">
                                                            ${enrollmentStr}
                                                        </td>
                                                        <td style="border: 1px solid var(--border-color); padding: 12px; font-size: 10pt;">
                                                            ${nameStr}
                                                        </td>
                                                        ${sheet.sessions.map(() => `<td style="border: 1px solid var(--border-color); padding: 12px;"></td>`).join('')}
                                                    </tr>
                                                    `;
                    }).join('')}
                                            </tbody>
                                        </table>
                                    </div>
                    `;
                });

                roomAttendanceHtml += `</div>`; // end room panel
            });
        } else {
            roomAttendanceHtml += `<p style="text-align:center; color: var(--text-muted);">Generating...</p>`;
        }
        document.getElementById('tab-attendance-room').innerHTML = roomAttendanceHtml;

        // Reset to first tab on generation
        document.querySelector('.tab-btn[data-target="tab-timetable"]').click();
        window.scrollTo(0, 0);
    }
});
