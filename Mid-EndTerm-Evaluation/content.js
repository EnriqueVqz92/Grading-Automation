chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === "RUN_AUTOMATION") {
        const tasks = request.tasks;
        const commentPref = request.commentPref; // "1" overwrite, "2" skip original
        const learnerProfilePref = request.learnerProfilePref;
        const partEffortPref = request.partEffortPref;

        const containers = document.querySelectorAll("div.js-student-grade");
        if (containers.length === 0) {
            sendResponse({ success: false, error: "No student containers found. Are you on the class list page?" });
            return true;
        }

        function flexMatch(nameA, nameB) {
            if (!nameA || !nameB) return false;
            let normalize = s => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9\s]/g, "");
            let wordsA = normalize(nameA).split(/\s+/).filter(w => w.length >= 2);
            let wordsB = normalize(nameB).split(/\s+/).filter(w => w.length >= 2);
            if (wordsA.length === 0 || wordsB.length === 0) return true;
            
            for (let wa of wordsA) {
                for (let wb of wordsB) {
                    if (wb === wa || (wa.length >= 3 && wb.length >= 3 && wb.substring(0, 3) === wa.substring(0, 3)) || wb.includes(wa) || wa.includes(wb)) {
                        return true;
                    }
                }
            }
            return false;
        }

        let processedCount = 0;

        // --- PRE-CHECK: VERIFICAR SI HAY AL MENOS UNA COINCIDENCIA DE 3 VÍAS ---
        let hasAnyMatch = false;
        for (let t of tasks) {
            if (!t.excel_found) continue;
            for (let c of containers) {
                let nameElem = c.querySelector("h4.student-name a");
                let mbName = nameElem ? nameElem.innerText.trim() : "";
                if (flexMatch(t.txt_name, mbName)) {
                    hasAnyMatch = true;
                    break;
                }
            }
            if (hasAnyMatch) break;
        }

        if (!hasAnyMatch && tasks.length > 0) {
            sendResponse({ success: false, error: "Error CRÍTICO: Las listas no coinciden. Ningún alumno coincide simultáneamente en el .txt, Excel y esta página de ManageBac. Proceso abortado." });
            return true;
        }

        let idxMb = 0;
        let missedTasks = [];

        for (let t = 0; t < tasks.length; t++) {
            let task = tasks[t];

            if (!task.excel_found) {
                console.warn(`[MB Automator Error] ❌ SALTADO: El alumno '${task.txt_name}' está en el .txt pero NO existe en el EXCEL. Revisa tus archivos.`);
                missedTasks.push(task);
                continue;
            }

            let container = null;
            if (idxMb < containers.length) {
                let nameElem = containers[idxMb].querySelector("h4.student-name a");
                let mbName = nameElem ? nameElem.innerText.trim() : "";
                
                if (flexMatch(task.txt_name, mbName)) {
                    container = containers[idxMb];
                    idxMb++;
                } else {
                    let foundMatch = false;
                    const MAX_LOOKAHEAD = 4;
                    
                    for (let totalOffset = 1; totalOffset <= MAX_LOOKAHEAD * 2 && !foundMatch; totalOffset++) {
                        for (let offsetTask = 0; offsetTask <= totalOffset; offsetTask++) {
                            let offsetMb = totalOffset - offsetTask;
                            if (offsetTask > MAX_LOOKAHEAD || offsetMb > MAX_LOOKAHEAD) continue;
                            
                            let nextTaskIdx = t + offsetTask;
                            let nextMbIdx = idxMb + offsetMb;
                            
                            if (nextTaskIdx < tasks.length && nextMbIdx < containers.length) {
                                let nextTask = tasks[nextTaskIdx];
                                let nElem = containers[nextMbIdx].querySelector("h4.student-name a");
                                let nMbName = nElem ? nElem.innerText.trim() : "";
                                
                                if (nextTask.excel_found && flexMatch(nextTask.txt_name, nMbName)) {
                                    if (offsetTask === 0) {
                                        container = containers[nextMbIdx];
                                        idxMb = nextMbIdx + 1;
                                    }
                                    foundMatch = true;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (!container) {
                missedTasks.push(task);
                continue;
            }

            // --- 1. PASTE COMMENT ---
            let skipIfNotEmpty = (commentPref === "2");
            let hasContent = false;

            let textareas = container.querySelectorAll('textarea');
            for (let t of textareas) {
                if (t.id.includes('comment') || t.name.includes('comment')) {
                    if (t.value.trim() !== '') hasContent = true;
                }
            }
            if (window.CKEDITOR) {
                for (let instance in window.CKEDITOR.instances) {
                    let editorNode = document.getElementById('cke_' + instance);
                    if (editorNode && container.contains(editorNode)) {
                        let data = window.CKEDITOR.instances[instance].getData().replace(/<[^>]*>/g, '').trim();
                        if (data !== '') hasContent = true;
                    }
                }
            }
            let editables = container.querySelectorAll('[contenteditable="true"]');
            for (let ed of editables) {
                if (ed.innerText.trim() !== '') hasContent = true;
            }

            if (!(skipIfNotEmpty && hasContent)) {
                // Limpieza de caracteres invisibles y retornos de carro (Windows)
                let cleanComment = task.comment.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/[\u200B-\u200D\uFEFF]/g, '');
                let htmlComment = `<p>${cleanComment.replace(/\n/g, '<br>')}</p>`;

                // Do the pasting logic
                for (let t of textareas) {
                    if (t.id.includes('comment') || t.name.includes('comment')) {
                        t.value = cleanComment;
                        t.dispatchEvent(new Event('change', { bubbles: true }));
                        t.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                }
                if (window.CKEDITOR) {
                    for (let instance in window.CKEDITOR.instances) {
                        let editorNode = document.getElementById('cke_' + instance);
                        if (editorNode && container.contains(editorNode)) {
                            window.CKEDITOR.instances[instance].setData(htmlComment);
                        }
                    }
                }
                for (let ed of editables) {
                    ed.innerHTML = htmlComment;
                    ed.dispatchEvent(new Event('input', { bubbles: true }));
                }
            }

            // --- 2. SET GRADES ---
            for (let crit in task.grades) {
                let targetScore = task.grades[crit];
                let buttons = container.querySelectorAll('[data-final-grade]');
                let found = false;

                for (let btn of buttons) {
                    try {
                        let attr = btn.getAttribute('data-final-grade');
                        if (!attr) continue;
                        let data = JSON.parse(attr);

                        if (data.criterion && data.criterion.toLowerCase() === crit.toLowerCase()) {
                            let isMatch = false;
                            if (targetScore === null) {
                                if (data.score === null || data.score === undefined || data.score === "") isMatch = true;
                            } else {
                                if (data.score !== null && data.score !== undefined && data.score == targetScore) isMatch = true;
                            }

                            if (isMatch) {
                                btn.click();
                                btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
                                btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
                                found = true;
                                break;
                            }
                        }
                    } catch (e) { }
                }

                // If N/A and not found by JSON, fallback to clicking N/A text
                if (targetScore === null && !found) {
                    let allElems = container.querySelectorAll('div, span, a, button');
                    for (let el of allElems) {
                        if (el.innerText && el.innerText.trim().toUpperCase() === 'N/A') {
                            let parentData = el.closest('[data-final-grade]');
                            if (parentData) {
                                let d = JSON.parse(parentData.getAttribute('data-final-grade'));
                                if (d.criterion.toLowerCase() === crit.toLowerCase()) {
                                    el.click();
                                    break;
                                }
                            } else {
                                el.click();
                            }
                        }
                    }
                }
            }

            // --- 3. SET PARTICIPATION & EFFORT ---
            if (partEffortPref && partEffortPref !== 'skip') {
                for (let pe of ['Participation', 'Effort']) {
                    let labels = container.querySelectorAll('label.control-label');
                    let targetLabel = Array.from(labels).find(l => l.innerText.trim() === pe);
                    
                    if (targetLabel) {
                        let selectId = targetLabel.getAttribute('for');
                        if (selectId) {
                            let selectElem = document.getElementById(selectId);
                            if (selectElem) {
                                if (partEffortPref === 'clear') {
                                    selectElem.value = "";
                                    selectElem.dispatchEvent(new Event('change', { bubbles: true }));
                                } else if (partEffortPref === 'overwrite') {
                                    let score = task.partEffort && task.partEffort[pe];
                                    if (score !== undefined && score !== null && score !== "") {
                                        let textValue = "";
                                        if (score == 4) textValue = "Proficient";
                                        else if (score == 3) textValue = "Capable";
                                        else if (score == 2) textValue = "Approaching expectations";
                                        else if (score == 1) textValue = "Below expectations";
                                        
                                        if (textValue !== "") {
                                            let opts = Array.from(selectElem.options);
                                            let opt = opts.find(o => o.value === textValue || o.text === textValue);
                                            if (opt) {
                                                selectElem.value = opt.value;
                                                selectElem.dispatchEvent(new Event('change', { bubbles: true }));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // --- 4. SET LEARNER PROFILE ---
            if (learnerProfilePref && learnerProfilePref.length > 0) {
                let lpSelect = container.querySelector('select.learner-profile-items');
                if (lpSelect) {
                    let changed = false;
                    
                    // Deseleccionamos todos los previos primero
                    Array.from(lpSelect.options).forEach(opt => {
                        if (opt.selected && opt.value !== "") {
                            opt.selected = false;
                            changed = true;
                        }
                    });

                    // Seleccionamos los nuevos
                    Array.from(lpSelect.options).forEach(opt => {
                        if (learnerProfilePref.includes(opt.value)) {
                            opt.selected = true;
                            changed = true;
                        }
                    });

                    if (changed) {
                        lpSelect.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                }
            }

            processedCount++;
        }

        // --- 3. REPORTE DE DESFASE (ALUMNOS NO ENCONTRADOS EN MANAGEBAC) ---
        if (missedTasks.length > 0) {
            let missedNames = missedTasks.map(t => t.txt_name).join(", ");
            console.warn(`[MB Automator Info] ℹ️ Alumnos del .txt saltados por no estar en MB: ${missedNames}`);
        }

        sendResponse({ success: true, count: processedCount });
        return true;
    }
});
