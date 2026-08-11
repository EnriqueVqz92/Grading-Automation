let workbook = null;
let _headers = [];
let _sheetData = [];
let _parsedComments = [];

document.addEventListener('DOMContentLoaded', () => {
  const txtUpload = document.getElementById('txtUpload');
  const excelUpload = document.getElementById('excelUpload');
  const txtUploadLbl = document.getElementById('txtUploadLbl');
  const excelUploadLbl = document.getElementById('excelUploadLbl');
  const sheetSelect = document.getElementById('sheetSelect');
  const step2 = document.getElementById('step2');
  const runBtn = document.getElementById('runBtn');
  const statusMsg = document.getElementById('statusMessage');

  // Handle TXT Upload
  txtUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    txtUploadLbl.textContent = file.name;
    txtUploadLbl.style.color = 'var(--text-main)';

    const reader = new FileReader();
    reader.onload = (evt) => {
      _parsedComments = parseComments(evt.target.result);
      if(_parsedComments.length > 0) {
        setStatus(`Loaded ${_parsedComments.length} comments.`, 'success');
      } else {
        setStatus(`Error: No comments matched the format.`, 'error');
      }
      checkReady();
    };
    reader.readAsText(file);
  });

  // Handle Excel Upload
  excelUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    excelUploadLbl.textContent = file.name;
    excelUploadLbl.style.color = 'var(--text-main)';

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      try {
        workbook = XLSX.read(data, {type: 'array'});
        populateSheetDropdown();
        step2.classList.remove('hidden');
        setStatus(`Excel loaded successfully.`, 'success');
      } catch (err) {
        setStatus(`Error reading Excel: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  });

  // Handle Sheet Selection
  sheetSelect.addEventListener('change', () => {
    loadSheetData(sheetSelect.value);
  });

  function populateSheetDropdown() {
    sheetSelect.innerHTML = '';
    workbook.SheetNames.forEach(name => {
      const opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      sheetSelect.appendChild(opt);
    });
    if(workbook.SheetNames.length > 0) {
      loadSheetData(workbook.SheetNames[0]);
    }
  }

  function loadSheetData(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ""});
    if(jsonData.length < 1) {
      setStatus("Sheet is empty.", "error"); return;
    }
    _headers = jsonData[0].map(h => h.toString().trim());
    
    // Create an array of objects
    _sheetData = [];
    for(let i = 1; i < jsonData.length; i++) {
      let rowObj = {};
      for(let j=0; j<_headers.length; j++) {
        rowObj[_headers[j]] = jsonData[i][j];
      }
      
      let fullNameChk = rowObj['NOMBRE'] !== undefined ? rowObj['NOMBRE'] : rowObj[_headers[0]];
      if (fullNameChk !== undefined && fullNameChk !== null && String(fullNameChk).trim() !== '') {
        _sheetData.push(rowObj);
      }
    }
    
    populateCriteriaDropdowns();
    checkReady();
  }

  function populateCriteriaDropdowns() {
    ['critA', 'critB', 'critC', 'critD', 'colPart', 'colEffort'].forEach(id => {
      const select = document.getElementById(id);
      select.innerHTML = '<option value="">-- Skip --</option>';
      _headers.forEach(header => {
        if(header && header.toUpperCase() !== 'NOMBRE' && header.toUpperCase() !== 'NAME') {
          const opt = document.createElement('option');
          opt.value = header;
          opt.textContent = header;
          select.appendChild(opt);
        }
      });
    });
  }

  function checkReady() {
    const commentPrefEl = document.getElementById('commentPref');
    let usePref3 = commentPrefEl && (commentPrefEl.value === '3' || commentPrefEl.value === '4');
    
    if (usePref3) {
      if (_sheetData.length > 0) {
        runBtn.disabled = false;
      } else {
        runBtn.disabled = true;
      }
    } else {
      if (_parsedComments.length > 0 && _sheetData.length > 0) {
        runBtn.disabled = false;
      } else {
        runBtn.disabled = true;
      }
    }
  }

  function setStatus(msg, type) {
    statusMsg.textContent = msg;
    statusMsg.className = 'status ' + type;
  }

  // --- Regex Parser ---
  function normalizeText(str) {
    if(!str) return "";
    return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  }

  function parseComments(text) {
    const content = '\n' + text;
    const blocks = content.split(/\n\d+\.\s*([^:]+):/);
    let stComments = [];
    for(let i = 1; i < blocks.length; i += 2) {
      if (i + 1 < blocks.length) {
        stComments.push({
          name: blocks[i].trim(),
          normalized_name: normalizeText(blocks[i].trim()),
          comment: blocks[i + 1].trim()
        });
      }
    }
    return stComments;
  }

  // --- Preferences Event Listeners ---
  const commentPrefEl = document.getElementById('commentPref');
  const extraTextRow = document.getElementById('extraTextRow');
  if (commentPrefEl && extraTextRow) {
    commentPrefEl.addEventListener('change', (e) => {
      if (e.target.value === '3' || e.target.value === '4') {
        extraTextRow.classList.remove('hidden');
      } else {
        extraTextRow.classList.add('hidden');
      }
      checkReady();
    });
  }

  // Learner Profile Multiselect Logic
  const learnerProfileHeader = document.getElementById('learnerProfileHeader');
  const learnerProfileDropdown = document.getElementById('learnerProfileDropdown');
  if (learnerProfileHeader && learnerProfileDropdown) {
    learnerProfileHeader.addEventListener('click', () => {
      learnerProfileDropdown.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
      if (!learnerProfileHeader.contains(e.target) && !learnerProfileDropdown.contains(e.target)) {
        learnerProfileDropdown.classList.add('hidden');
      }
    });

    const checkboxes = learnerProfileDropdown.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(cb => {
      cb.addEventListener('change', () => {
        const selected = Array.from(checkboxes).filter(c => c.checked).map(c => c.parentElement.textContent.trim());
        if (selected.length === 0) {
          learnerProfileHeader.textContent = '-- Do not select --';
        } else if (selected.length === 1) {
          learnerProfileHeader.textContent = selected[0];
        } else {
          learnerProfileHeader.textContent = `${selected.length} options selected`;
        }
      });
    });
  }

  // --- Action ---
  runBtn.addEventListener('click', () => {
    runBtn.disabled = true;
    runBtn.textContent = 'Processing...';
    
    // Retrieve preferences
    const emptyPref = document.getElementById('emptyPref').value;
    const skipPref = document.getElementById('skipPref').value;
    const commentPref = document.getElementById('commentPref').value;
    const partEffortPref = document.getElementById('partEffortPref').value;
    const extraText = document.getElementById('extraText') ? document.getElementById('extraText').value.trim() : "";
    const learnerProfileCheckboxes = document.querySelectorAll('#learnerProfileDropdown input[type="checkbox"]:checked');
    const learnerProfilePref = Array.from(learnerProfileCheckboxes).map(cb => cb.value);
    
    const critMap = {
      'A': document.getElementById('critA').value,
      'B': document.getElementById('critB').value,
      'C': document.getElementById('critC').value,
      'D': document.getElementById('critD').value
    };

    const partEffortMap = {
      'Participation': document.getElementById('colPart').value,
      'Effort': document.getElementById('colEffort').value
    };

    // Función de comprobación secundaria (muy flexible)
    function flexMatch(nameA, nameB) {
      if (!nameA || !nameB) return false;
      const normalize = s => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9\s]/g, "");
      const wordsA = normalize(nameA).split(/\s+/).filter(w => w.length >= 2);
      const wordsB = normalize(nameB).split(/\s+/).filter(w => w.length >= 2);
      if (wordsA.length === 0 || wordsB.length === 0) return true; // Si está vacío pero están en la misma posición, lo damos por bueno
      
      for (let wa of wordsA) {
        for (let wb of wordsB) {
          // Comparten 3 letras iniciales O uno contiene al otro (Ej: abi en abigail)
          if (wb === wa || (wa.length >= 3 && wb.length >= 3 && wb.substring(0, 3) === wa.substring(0, 3)) || wb.includes(wa) || wa.includes(wb)) {
            return true; // Ya que el ORDEN es el filtro principal, 1 sola coincidencia parcial es suficiente
          }
        }
      }
      return false;
    }

    // Build the finalized array of tasks for the selected students
    let finalTasks = [];
    let excelMatchCount = 0;

    if (commentPref === '3' || commentPref === '4') {
      // Usar solo datos de Excel/Numbers e ignorar .txt
      for (let i = 0; i < _sheetData.length; i++) {
        let rowItem = _sheetData[i];
        let excelName = rowItem['NOMBRE'] !== undefined ? rowItem['NOMBRE'] : rowItem[_headers[0]];
        
        let studentPayload = {
          txt_name: excelName, // Usamos el nombre del excel como base
          excel_found: true,
          comment: extraText,
          grades: {},
          partEffort: {}
        };
        excelMatchCount++;

        // Extraer calificaciones
        for(let crit of ['A', 'B', 'C', 'D']) {
          let headerName = critMap[crit];
          if(headerName) { 
            let val = rowItem[headerName];
            if(val !== undefined && val !== null && String(val).trim() !== '') {
              studentPayload.grades[crit] = parseInt(val, 10);
            } else if (emptyPref === "2") {
              studentPayload.grades[crit] = null; 
            }
          } else if (skipPref === "2") {
            studentPayload.grades[crit] = null; 
          }
        }

        for (let pe of ['Participation', 'Effort']) {
          let headerName = partEffortMap[pe];
          if (headerName) {
            let val = rowItem[headerName];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              studentPayload.partEffort[pe] = parseInt(val, 10);
            }
          }
        }

        finalTasks.push(studentPayload);
      }
    } else {
      // Ahora iteramos por cada comentario del .txt y buscamos su fila correspondiente en Excel
      let idxExcel = 0;
      for(let i=0; i < _parsedComments.length; i++) {
        let commentItem = _parsedComments[i];
        let rowItem = null;
        
        if (idxExcel < _sheetData.length) {
          let excelName = _sheetData[idxExcel]['NOMBRE'] !== undefined ? _sheetData[idxExcel]['NOMBRE'] : _sheetData[idxExcel][_headers[0]];
          
          if (flexMatch(commentItem.name, excelName)) {
            rowItem = _sheetData[idxExcel];
            idxExcel++;
          } else {
            // Cierre Inteligente (Lookahead)
            let foundMatch = false;
            const MAX_LOOKAHEAD = 4;
            
            for (let totalOffset = 1; totalOffset <= MAX_LOOKAHEAD * 2 && !foundMatch; totalOffset++) {
              for (let offsetTxt = 0; offsetTxt <= totalOffset; offsetTxt++) {
                let offsetExc = totalOffset - offsetTxt;
                if (offsetTxt > MAX_LOOKAHEAD || offsetExc > MAX_LOOKAHEAD) continue;
                
                let nextTxtIdx = i + offsetTxt;
                let nextExcIdx = idxExcel + offsetExc;
                
                if (nextTxtIdx < _parsedComments.length && nextExcIdx < _sheetData.length) {
                  let nextTxt = _parsedComments[nextTxtIdx];
                  let nextExcName = _sheetData[nextExcIdx]['NOMBRE'] !== undefined ? _sheetData[nextExcIdx]['NOMBRE'] : _sheetData[nextExcIdx][_headers[0]];
                  
                  if (flexMatch(nextTxt.name, nextExcName)) {
                    if (offsetTxt === 0) {
                      rowItem = _sheetData[nextExcIdx];
                      idxExcel = nextExcIdx + 1;
                    }
                    foundMatch = true;
                    break;
                  }
                }
              }
            }
          }
        }
        
        // Si no encuentra fila, crea un objeto vacío para que no falle la extracción de notas
        let excelFound = false;
        if (!rowItem) {
          rowItem = {};
        } else {
          excelFound = true;
          excelMatchCount++;
        }
        
        let finalCommentText = commentItem.comment;

        let studentPayload = {
          txt_name: commentItem.name,
          excel_found: excelFound,
          comment: finalCommentText,
          grades: {},
          partEffort: {}
        };

        // Extract grades based on selected headers
        for(let crit of ['A', 'B', 'C', 'D']) {
          let headerName = critMap[crit];
          if(headerName) { // Not skipped
            let val = rowItem[headerName];
            if(val !== undefined && val !== null && String(val).trim() !== '') {
              studentPayload.grades[crit] = parseInt(val, 10);
            } else if (emptyPref === "2") {
              studentPayload.grades[crit] = null; // Mark N/A
            }
          } else if (skipPref === "2") {
            studentPayload.grades[crit] = null; // Mark N/A
          }
        }

        for (let pe of ['Participation', 'Effort']) {
          let headerName = partEffortMap[pe];
          if (headerName) {
            let val = rowItem[headerName];
            if (val !== undefined && val !== null && String(val).trim() !== '') {
              studentPayload.partEffort[pe] = parseInt(val, 10);
            }
          }
        }

        finalTasks.push(studentPayload);
      }
    }

    if (excelMatchCount === 0 && finalTasks.length > 0) {
      setStatus("Error: Ningún alumno del .txt coincide con el Excel. Revisa tus archivos.", "error");
      runBtn.textContent = 'Run Automation';
      runBtn.disabled = false;
      return;
    }

    // Connect to the active Chrome tab
    chrome.tabs.query({active: true, currentWindow: true}, function(tabs) {
      if(!tabs[0].url.includes('managebac.com')) {
        setStatus("Target tab must be ManageBac class list.", "error");
        runBtn.textContent = 'Run Automation';
        runBtn.disabled = false;
        return;
      }

      chrome.scripting.executeScript({
        target: {tabId: tabs[0].id},
        files: ['content.js']
      }, () => {
        // Send our data to content.js
        let jsCommentPref = commentPref;
        if (commentPref === '3') jsCommentPref = '2'; // Paste ONLY if empty
        if (commentPref === '4') jsCommentPref = '1'; // Overwrite existing

        chrome.tabs.sendMessage(tabs[0].id, {
          action: "RUN_AUTOMATION", 
          tasks: finalTasks,
          commentPref: jsCommentPref,
          learnerProfilePref: learnerProfilePref,
          partEffortPref: partEffortPref
        }, function(response) {
            if (response) {
                if (response.success) {
                    setStatus(`Processed ${response.count} students!`, "success");
                } else if (response.error) {
                    setStatus(response.error, "error");
                } else {
                    setStatus("Ocurrió un error desconocido durante el proceso.", "error");
                }
            } else {
                setStatus("Error de conexión. Por favor refresca la página de ManageBac.", "error");
            }
            runBtn.textContent = 'Run Automation';
            runBtn.disabled = false;
        });
      });
    });
  });

});
