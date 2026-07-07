/* ===========================================================
 *  MühendisCalc v4.0 — Dashboard Scientific Calculator SaaS
 *  IIFE — depends on globals: math, functionPlot, jspdf, XLSX
 * =========================================================== */
window.exportPanelToPDF = function(panelId, filename) {
    var element = document.getElementById(panelId);
    if (!element) return;
    var exportBtns = element.querySelectorAll('.btn--export');
    exportBtns.forEach(function(btn) { btn.style.display = 'none'; });
    
    element.classList.add('pdf-export-mode');
    var emptyRows = [];
    if (panelId === 'panel-conv') {
        var convRows = element.querySelectorAll('.conv-row');
        convRows.forEach(function(row) {
            var inp1 = row.querySelector('.conv-inp-1');
            var inp2 = row.querySelector('.conv-inp-2');
            if (inp1 && inp2 && inp1.value === '' && inp2.value === '') {
                row.classList.add('pdf-hide');
                emptyRows.push(row);
            }
        });
    }
    
    html2pdf().set({
        margin: 10,
        filename: filename,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, logging: false },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).from(element).save().then(function() {
        exportBtns.forEach(function(btn) { btn.style.display = ''; });
        element.classList.remove('pdf-export-mode');
        emptyRows.forEach(function(row) { row.classList.remove('pdf-hide'); });
    });
};

(function () {
    'use strict';

    // =========================================================
    // 1. STATE
    // =========================================================
    var state = {
        result: '0',
        history: [],
        historyIndex: -1,
        angleMode: 'DEG',
        lastAnswer: 0,
        justEvaluated: false,
        currentLang: 'tr'
    };

    var mathField = null;
    var resultEl = null;

    // =========================================================
    // 2. i18n TRANSLATIONS
    // =========================================================
    var translations = {
        tr: {
            appTitle: 'Mühendislik Hesap Makinesi',
            logoText: 'MühendisCalc',
            tabStandard: 'Standart & Bilimsel',
            tabPlot: 'Grafik Çizici (PLOT)',
            tabMatrix: 'Matris (MATRIX)',
            tabEqn: 'Denklem (EQN)',
            tabCmplx: 'Kompleks (CMPLX)',
            tabConv: 'Birim Çevirici (CONV)',
            tabMacro: 'Makrolar (MACRO)',
            tabToolkit: 'Araç Kiti (TOOLKIT)',
            tabBase: 'Taban (BASE-N)',
            tabConst: 'Sabitler (CONST)',
            infoStandard: 'Temel hesaplamalar...',
            unit_km: 'km (Kilometre)',
            unit_m: 'm (Metre)',
            unit_cm: 'cm (Santimetre)',
            unit_mm: 'mm (Milimetre)',
            unit_um: 'μm (Mikrometre)',
            unit_nm: 'nm (Nanometre)',
            unit_inch: 'inch (İnç)',
            unit_ft: 'ft (Feet / Ayak)',
            unit_Pa: 'Pa (Pascal)',
            unit_kPa: 'kPa (Kilopascal)',
            unit_MPa: 'MPa (Megapascal)',
            unit_GPa: 'GPa (Gigapascal)',
            unit_bar: 'bar (Bar)',
            unit_atm: 'atm (Atmosfer)',
            unit_psi: 'psi (PSI)',
            'unit_kN/m²': 'kN/m² (Kilonewton/m²)',
            unit_W: 'W (Watt)',
            unit_kW: 'kW (Kilowatt)',
            unit_MW: 'MW (Megawatt)',
            unit_GW: 'GW (Gigawatt)',
            unit_mW: 'mW (Miliwatt)',
            unit_hp: 'HP (Beygir Gücü)',
            unit_Nm: 'Nm (Newton-metre)',
            unit_kNm: 'kNm (Kilonewton-metre)',
            unit_Nmm: 'Nmm (Newton-milimetre)',
            unit_kgfm: 'kgfm (Kilogram-kuvvet metre)',
            'unit_lb-ft': 'lb-ft (Pound-foot)',
            unit_C: '°C (Santigrat)',
            unit_F: '°F (Fahrenheit)',
            unit_K: 'K (Kelvin)'
        },
        en: {
            appTitle: 'Engineering Calculator',
            logoText: 'EngineerCalc',
            tabStandard: 'Standard & Scientific',
            tabPlot: 'Plotter (PLOT)',
            tabMatrix: 'Matrix (MATRIX)',
            tabEqn: 'Equation (EQN)',
            tabCmplx: 'Complex (CMPLX)',
            tabConv: 'Unit Converter (CONV)',
            tabMacro: 'Macros (MACRO)',
            tabToolkit: 'Toolkit (TOOLKIT)',
            tabBase: 'Base (BASE-N)',
            tabConst: 'Constants (CONST)',
            infoStandard: 'Basic calculations...',
            unit_km: 'km (Kilometer)',
            unit_m: 'm (Meter)',
            unit_cm: 'cm (Centimeter)',
            unit_mm: 'mm (Millimeter)',
            unit_um: 'μm (Micrometer)',
            unit_nm: 'nm (Nanometer)',
            unit_inch: 'inch (Inch)',
            unit_ft: 'ft (Foot)',
            unit_Pa: 'Pa (Pascal)',
            unit_kPa: 'kPa (Kilopascal)',
            unit_MPa: 'MPa (Megapascal)',
            unit_GPa: 'GPa (Gigapascal)',
            unit_bar: 'bar (Bar)',
            unit_atm: 'atm (Atmosphere)',
            unit_psi: 'psi (PSI)',
            'unit_kN/m²': 'kN/m² (Kilonewton/m²)',
            unit_W: 'W (Watt)',
            unit_kW: 'kW (Kilowatt)',
            unit_MW: 'MW (Megawatt)',
            unit_GW: 'GW (Gigawatt)',
            unit_mW: 'mW (Milliwatt)',
            unit_hp: 'HP (Horsepower)',
            unit_Nm: 'Nm (Newton-meter)',
            unit_kNm: 'kNm (Kilonewton-meter)',
            unit_Nmm: 'Nmm (Newton-millimeter)',
            unit_kgfm: 'kgfm (Kilogram-force meter)',
            'unit_lb-ft': 'lb-ft (Pound-foot)',
            unit_C: '°C (Celsius)',
            unit_F: '°F (Fahrenheit)',
            unit_K: 'K (Kelvin)'
        }
    };

    function changeLanguage(lang) {
        state.currentLang = lang;
        var t = translations[lang] || translations.tr;
        document.querySelectorAll('[data-i18n]').forEach(function (el) {
            var key = el.getAttribute('data-i18n');
            if (t[key] !== undefined) {
                el.textContent = t[key];
            }
        });
        document.querySelectorAll('[data-i18n-ph]').forEach(function (el) {
            var key = el.getAttribute('data-i18n-ph');
            if (t[key] !== undefined) {
                el.setAttribute('placeholder', t[key]);
            }
        });
        if (typeof setupConversions === 'function') {
            setupConversions();
        }
    }

    // =========================================================
    // 3. TAB MANAGER
    // =========================================================
    function switchTab(tabId) {
        document.querySelectorAll('.sidebar__item').forEach(function (btn) {
            btn.classList.remove('sidebar__item--active');
        });
        var activeBtn = document.querySelector('.sidebar__item[data-tab="' + tabId + '"]');
        if (activeBtn) activeBtn.classList.add('sidebar__item--active');

        document.querySelectorAll('.workspace__panel').forEach(function (p) {
            p.classList.remove('workspace__panel--active');
        });
        var targetPanel = document.getElementById('panel-' + tabId);
        if (targetPanel) targetPanel.classList.add('workspace__panel--active');

        var sidebar = document.getElementById('sidebar');
        if (sidebar) sidebar.classList.remove('sidebar--open');
    }

    // =========================================================
    // 4. DISPLAY & MATHLIVE
    // =========================================================
    function initMathLive() {
        if (window.customElements) {
            customElements.whenDefined('math-field').then(function() {
                mathField = document.getElementById('math-input');
                if (mathField) {
                    mathField.addEventListener('input', livePreview);
                    mathField.addEventListener('beforeinput', function(ev) {
                        if (ev.inputType === 'insertText' && ev.data === '/') {
                            ev.preventDefault();
                            mathField.executeCommand(['insert', '\\frac{#@}{#0}']);
                        } else if (ev.inputType === 'insertText' && (ev.data === ',' || ev.data === '.')) {
                            ev.preventDefault();
                            mathField.executeCommand(['insert', '.']);
                        }
                    });
                }
            });
        }
    }

    function updateResultDisplay() {
        if (!resultEl) return;
        resultEl.textContent = state.result;
        if (String(state.result).length > 14) {
            resultEl.classList.add('display__result--small');
        } else {
            resultEl.classList.remove('display__result--small');
        }
    }

    function formatResult(value) {
        if (typeof value === 'undefined' || value === null) return 'Hata';
        if (typeof value === 'object' && value.isBigNumber) value = value.toNumber();
        if (typeof value !== 'number') value = Number(value);
        if (isNaN(value)) return 'Hata';
        if (!isFinite(value)) return 'Tanımsız';
        if (Math.abs(value) > 1e15 || (Math.abs(value) < 1e-10 && value !== 0)) {
            return value.toExponential(8);
        }
        return String(parseFloat(value.toPrecision(10)));
    }

    function formatNum(n) {
        if (typeof n === 'object' && n.isBigNumber) n = n.toNumber();
        if (typeof n !== 'number') n = Number(n);
        if (isNaN(n)) return 'NaN';
        if (!isFinite(n)) return '∞';
        if (Math.abs(n) > 1e12 || (Math.abs(n) < 1e-8 && n !== 0)) return n.toExponential(4);
        return parseFloat(n.toPrecision(8)).toString();
    }

    function preprocessExpression(expr) {
        var s = expr;
        
        // MathLive Parser Bug: missing parens for trigonometric/logarithmic functions
        // Sadece yanında parantez olmayan sayıları yakalar, sin(90) bozulmaz.
        s = s.replace(/\\(sin|cos|tan|log|ln|sqrt|arcsin|arccos|arctan|asin|acos|atan)/g, '$1');
        s = s.replace(/arcsin/g, 'asin');
        s = s.replace(/arccos/g, 'acos');
        s = s.replace(/arctan/g, 'atan');
        s = s.replace(/(sin|cos|tan|log|ln|sqrt|asin|acos|atan)\s*(-?[0-9.,]+|[a-zA-Z]+)/g, '$1($2)');

        s = s.replace(/×/g, '*');
        s = s.replace(/÷/g, '/');
        s = s.replace(/\u2212/g, '-');
        s = s.replace(/pi/g, '(pi)');
        s = s.replace(/ANS/g, '(' + state.lastAnswer + ')');
        s = s.replace(/log\s*\(/g, 'log10(');
        s = s.replace(/ln\s*\(/g, 'log(');
        s = s.replace(/sqrt\s*\(/g, 'sqrt(');
        s = s.replace(/%/g, '/100');

        s = s.replace(/\)\(/g, ')*(');
        s = s.replace(/\)(\d)/g, ')*$1');
        s = s.replace(/(\d)\(/g, '$1*(');

        if (state.angleMode === 'DEG') {
            s = s.replace(/asin\s*\(/g, '__asd(');
            s = s.replace(/acos\s*\(/g, '__acd(');
            s = s.replace(/atan\s*\(/g, '__atd(');
            s = s.replace(/sin\s*\(/g, '__sd(');
            s = s.replace(/cos\s*\(/g, '__cd(');
            s = s.replace(/tan\s*\(/g, '__td(');
        }
        return s;
    }

    function getEvalScope() {
        return {
            __sd: function (x) { return Math.sin(x * Math.PI / 180); },
            __cd: function (x) { return Math.cos(x * Math.PI / 180); },
            __td: function (x) { return Math.tan(x * Math.PI / 180); },
            __asd: function (x) { return Math.asin(x) * 180 / Math.PI; },
            __acd: function (x) { return Math.acos(x) * 180 / Math.PI; },
            __atd: function (x) { return Math.atan(x) * 180 / Math.PI; }
        };
    }

    function livePreview() {
        if (!mathField) return;
        try {
            var expr = mathField.getValue('ascii-math') || '';
            if (expr.trim() === '') {
                state.result = '0';
                updateResultDisplay();
                return;
            }
            var processed = preprocessExpression(expr);
            var evalResult = math.evaluate(processed, getEvalScope());
            if (evalResult !== undefined && typeof evalResult !== 'function') {
                var num = evalResult;
                if (typeof num === 'object' && num.isBigNumber) num = num.toNumber();
                if (typeof num === 'number' && isFinite(num)) {
                    state.result = formatResult(num);
                    updateResultDisplay();
                }
            }
        } catch (e) { }
    }

    function evaluateExpression() {
        if (!mathField) return;
        var ascii = mathField.getValue('ascii-math') || '';
        var latex = mathField.getValue() || '';
        if (ascii.trim() === '') return;
        
        // Akıllı Yönlendirme: İntegral içeriyorsa nerdamer'a gönder
        if (ascii.indexOf('int ') !== -1 || latex.indexOf('\\int') !== -1) {
            if (typeof nerdamer === 'undefined') {
                state.result = 'Nerdamer yüklenemedi';
                state.justEvaluated = true;
                updateResultDisplay();
                return;
            }
            try {
                // ascii-math 'int x^2 dx' formatını yakalar
                var intMatch = ascii.match(/int\s+(.*?)\s+d([a-zA-Z])/);
                if (intMatch) {
                    var integrand = preprocessExpression(intMatch[1]);
                    var variable = intMatch[2];
                    var integrated = nerdamer('integrate(' + integrand + ', ' + variable + ')').toTeX();
                    state.result = integrated + ' + C';
                    state.justEvaluated = true;
                    state.historyIndex = -1;
                    updateResultDisplay();
                    addToHistory(latex, state.result);
                    return;
                }
                
                // Türev yakalama: (d)/(dx)(x^2) veya d/dx x^2 vb.
                var diffMatch = ascii.match(/(?:\(d\)|d)\/\(?d([a-zA-Z])\)?\s*(.*)/);
                if (diffMatch) {
                    var variable = diffMatch[1];
                    var expression = diffMatch[2];
                    if (expression.startsWith('(') && expression.endsWith(')')) {
                        expression = expression.substring(1, expression.length - 1);
                    }
                    var processedExpr = preprocessExpression(expression);
                    var derivative = nerdamer('diff(' + processedExpr + ', ' + variable + ')').toTeX();
                    state.result = derivative;
                    state.justEvaluated = true;
                    state.historyIndex = -1;
                    updateResultDisplay();
                    addToHistory(latex, state.result);
                    return;
                }
            } catch(e) {
                state.result = 'Sembolik Hesaplama Hatası';
                state.justEvaluated = true;
                updateResultDisplay();
                return;
            }
        }
        
        try {
            var processed = preprocessExpression(ascii);
            var evalResult = math.evaluate(processed, getEvalScope());
            if (typeof evalResult === 'object' && evalResult.isBigNumber) {
                evalResult = evalResult.toNumber();
            }
            var formatted = formatResult(evalResult);
            addToHistory(latex, formatted);
            state.lastAnswer = (typeof evalResult === 'number') ? evalResult : Number(evalResult);
            state.result = formatted;
            state.justEvaluated = true;
            state.historyIndex = -1;
        } catch (e) {
            state.result = 'Geçersiz/Eksik İfade';
            state.justEvaluated = true;
        }
        updateResultDisplay();
    }

    function handleButtonAction(action, btn) {
        if (!mathField) return;

        if (state.justEvaluated) {
            if (['add','subtract','multiply','divide'].indexOf(action) === -1 && action !== 'evaluate') {
                mathField.setValue('');
            }
            state.justEvaluated = false;
        }

        switch (action) {
            case 'digit': mathField.insert(btn.getAttribute('data-value')); break;
            case 'decimal': mathField.insert('.'); break;
            case 'add': mathField.insert('+'); break;
            case 'subtract': mathField.insert('-'); break;
            case 'multiply': mathField.insert('*'); break;
            case 'divide': mathField.insert('\\frac{#@}{#0}'); break;
            case 'open-paren': mathField.insert('('); break;
            case 'close-paren': mathField.insert(')'); break;
            case 'clear': mathField.setValue(''); state.result = '0'; updateResultDisplay(); break;
            case 'backspace': mathField.executeCommand('deleteBackward'); break;
            case 'evaluate': evaluateExpression(); break;
            case 'sin': mathField.insert('\\sin('); break;
            case 'cos': mathField.insert('\\cos('); break;
            case 'tan': mathField.insert('\\tan('); break;
            case 'asin': mathField.insert('\\arcsin('); break;
            case 'acos': mathField.insert('\\arccos('); break;
            case 'atan': mathField.insert('\\arctan('); break;
            case 'log': mathField.executeCommand(['insert', '\\log(']); break;
            case 'ln': mathField.executeCommand(['insert', '\\ln(']); break;
            case 'sqrt': mathField.executeCommand(['insert', '\\sqrt{#0}']); break;
            case 'fraction': mathField.executeCommand(['insert', '\\frac{#@}{#0}']); break;
            case 'square': mathField.executeCommand(['insert', '#@^{2}']); break;
            case 'power': mathField.executeCommand(['insert', '^{#0}']); break;
            case 'factorial': mathField.executeCommand(['insert', '!']); break;
            case 'abs': mathField.executeCommand(['insert', '|#0|']); break;
            case 'pow10': mathField.executeCommand(['insert', '10^{#0}']); break;
            case 'exp': mathField.executeCommand(['insert', 'e^{#0}']); break;
            case 'integral': mathField.executeCommand(['insert', '\\int {#0} dx']); break;
            case 'sym_deriv_main': mathField.executeCommand(['insert', '\\frac{d}{dx}({#0})']); break;
            case 'pi': mathField.executeCommand(['insert', '\\pi']); break;
            case 'euler': mathField.executeCommand(['insert', 'e']); break;
            case 'ans': mathField.executeCommand(['insert', 'ANS']); break;
            case 'negate': mathField.insert('-'); break;
            case 'percent': mathField.insert('\\%'); break;
        }
        mathField.focus();
        livePreview();
    }

    // =========================================================
    // 5. HISTORY & ANGLE
    // =========================================================
    function addToHistory(exprLatex, result) {
        state.history.push({ expr: exprLatex, result: result });
        if (state.history.length > 50) state.history.shift();
    }

    function navigateHistory(direction) {
        if (state.history.length === 0 || !mathField) return;
        if (direction === 'up') {
            if (state.historyIndex === -1) {
                state.historyIndex = state.history.length - 1;
            } else if (state.historyIndex > 0) {
                state.historyIndex--;
            }
        } else if (direction === 'down') {
            if (state.historyIndex === -1) return;
            if (state.historyIndex < state.history.length - 1) {
                state.historyIndex++;
            } else {
                state.historyIndex = -1;
                mathField.setValue('');
                state.result = '0';
                state.justEvaluated = false;
                updateResultDisplay();
                return;
            }
        }
        var entry = state.history[state.historyIndex];
        if (entry) {
            mathField.setValue(entry.expr);
            state.result = entry.result;
            state.justEvaluated = false;
            updateResultDisplay();
        }
    }

    function toggleAngleMode() {
        var modeBtn = document.getElementById('mode-toggle');
        if (state.angleMode === 'DEG') {
            state.angleMode = 'RAD';
            if (modeBtn) {
                modeBtn.textContent = 'RAD';
                modeBtn.classList.add('topbar-btn--rad');
            }
        } else {
            state.angleMode = 'DEG';
            if (modeBtn) {
                modeBtn.textContent = 'DEG';
                modeBtn.classList.remove('topbar-btn--rad');
            }
        }
        livePreview();
    }

    // =========================================================
    // 6. PLOT MODULE
    // =========================================================
    function setupPlot() {
        var plotBtn = document.getElementById('plot-btn');
        if (!plotBtn) return;
        plotBtn.addEventListener('click', function() {
            var func = document.getElementById('plot-func').value.trim();
            if (!func) return;
            try {
                var container = document.getElementById('plot-container');
                var w = container.clientWidth || 400;
                var h = container.clientHeight || 350;
                
                functionPlot({
                    target: '#plot-container',
                    data: [{ fn: func }],
                    grid: true,
                    width: w,
                    height: h,
                    xAxis: { domain: [-10, 10] },
                    yAxis: { domain: [-15, 15] }
                });
            } catch (e) {
                console.error('Plot error', e);
            }
        });
    }

    // =========================================================
    // 7. EXPORT MODULE
    // =========================================================
    function exportToPdf(text, filename) {
        if (!window.jspdf || !window.jspdf.jsPDF) return;
        var doc = new window.jspdf.jsPDF();
        var lines = doc.splitTextToSize(text, 180);
        doc.text(lines, 10, 10);
        doc.save(filename);
    }

    function exportToExcel(text, filename) {
        if (!window.XLSX) return;
        var rows = text.split('\n').map(function(line) {
            return line.split(/\s+/).filter(Boolean);
        });
        var wb = XLSX.utils.book_new();
        var ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, "Result");
        XLSX.writeFile(wb, filename);
    }

    function bindExport() {
        var matrixPdf = document.getElementById('matrix-pdf');
        if (matrixPdf) {
            matrixPdf.addEventListener('click', function() {
                var res = document.getElementById('matrix-result').textContent;
                exportToPdf(res, 'matrix_result.pdf');
            });
        }
        var matrixExcel = document.getElementById('matrix-excel');
        if (matrixExcel) {
            matrixExcel.addEventListener('click', function() {
                var res = document.getElementById('matrix-result').textContent;
                exportToExcel(res, 'matrix_result.xlsx');
            });
        }
        var eqnPdf = document.getElementById('eqn-pdf');
        if (eqnPdf) {
            eqnPdf.addEventListener('click', function() {
                var res = document.getElementById('eqn-result').textContent;
                exportToPdf(res, 'eqn_result.pdf');
            });
        }
        var eqnExcel = document.getElementById('eqn-excel');
        if (eqnExcel) {
            eqnExcel.addEventListener('click', function() {
                var res = document.getElementById('eqn-result').textContent;
                exportToExcel(res, 'eqn_result.xlsx');
            });
        }
    }

    // =========================================================
    // 8. MACRO MODULE
    // =========================================================
    var macros = [];
    var activeMacro = null;

    function loadMacros() {
        try {
            var saved = localStorage.getItem('calc-macros');
            return saved ? JSON.parse(saved) : [];
        } catch(e) { return []; }
    }
    function saveMacros(ms) {
        localStorage.setItem('calc-macros', JSON.stringify(ms));
    }
    
    function renderMacros() {
        var list = document.getElementById('macro-list');
        if (!list) return;
        list.innerHTML = '';
        macros.forEach(function(m, i) {
            var d = document.createElement('div');
            d.className = 'macro-item';
            d.style.display = 'flex';
            d.style.justifyContent = 'space-between';
            d.style.alignItems = 'center';
            d.style.padding = '8px';
            d.style.borderBottom = '1px solid rgba(255,255,255,0.1)';
            
            var nameSpan = document.createElement('span');
            nameSpan.textContent = m.name;
            nameSpan.style.cursor = 'pointer';
            nameSpan.style.flex = '1';
            nameSpan.addEventListener('click', function() {
                selectMacro(i);
            });
            
            var delBtn = document.createElement('button');
            delBtn.className = 'btn--macro-del';
            delBtn.innerHTML = '&times;';
            delBtn.style.background = '#e74c3c';
            delBtn.style.color = '#fff';
            delBtn.style.border = 'none';
            delBtn.style.borderRadius = '4px';
            delBtn.style.padding = '4px 8px';
            delBtn.style.cursor = 'pointer';
            delBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                macros.splice(i, 1);
                saveMacros(macros);
                renderMacros();
            });
            
            d.appendChild(nameSpan);
            d.appendChild(delBtn);
            list.appendChild(d);
        });
    }

    function selectMacro(index) {
        activeMacro = macros[index];
        var title = document.getElementById('macro-exec-title');
        if (title) title.textContent = 'Makro Çalıştır: ' + activeMacro.name;
        
        var varsDiv = document.getElementById('macro-vars');
        varsDiv.innerHTML = '';
        
        var mathFuncs = ['sin','cos','tan','log','ln','sqrt','abs','exp','asin','acos','atan','pi','e'];
        var matches = activeMacro.formula.match(/[A-Za-z_][A-Za-z0-9_]*/g) || [];
        var uniqueVars = [];
        matches.forEach(function(v) {
            if (mathFuncs.indexOf(v) === -1 && uniqueVars.indexOf(v) === -1) {
                uniqueVars.push(v);
            }
        });
        
        uniqueVars.forEach(function(v) {
            var field = document.createElement('div');
            field.className = 'panel-field';
            field.innerHTML = '<label class="panel-label">' + v + '</label>' +
                              '<input type="number" class="panel-input macro-var-input" data-var="' + v + '" step="any">';
            varsDiv.appendChild(field);
        });
        
        var btn = document.getElementById('macro-run');
        if (btn) btn.style.display = 'block';
    }

    function setupMacro() {
        macros = loadMacros();
        var saveBtn = document.getElementById('macro-save');
        if (saveBtn) {
            saveBtn.addEventListener('click', function() {
                var name = document.getElementById('macro-name').value.trim();
                var form = document.getElementById('macro-formula').value.trim();
                if (name && form) {
                    macros.push({ name: name, formula: form });
                    saveMacros(macros);
                    renderMacros();
                    document.getElementById('macro-name').value = '';
                    document.getElementById('macro-formula').value = '';
                }
            });
        }
        var runBtn = document.getElementById('macro-run');
        if (runBtn) {
            runBtn.addEventListener('click', function() {
                if (!activeMacro) return;
                var scope = {};
                document.querySelectorAll('.macro-var-input').forEach(function(inp) {
                    var v = inp.getAttribute('data-var');
                    scope[v] = parseFloat(inp.value) || 0;
                });
                try {
                    var res = math.evaluate(activeMacro.formula, scope);
                    document.getElementById('macro-result').textContent = 'Sonuç: ' + formatNum(res);
                } catch(e) {
                    document.getElementById('macro-result').textContent = 'Hata: ' + e.message;
                }
            });
        }
        renderMacros();
    }

    // =========================================================
    // 9. TOOLKIT MODULE
    // =========================================================
    var resColors = [
        { name: 'Siyah', val: 0, mult: 1, tol: null },
        { name: 'Kahverengi', val: 1, mult: 10, tol: 1 },
        { name: 'Kırmızı', val: 2, mult: 100, tol: 2 },
        { name: 'Turuncu', val: 3, mult: 1000, tol: null },
        { name: 'Sarı', val: 4, mult: 10000, tol: null },
        { name: 'Yeşil', val: 5, mult: 100000, tol: 0.5 },
        { name: 'Mavi', val: 6, mult: 1000000, tol: 0.25 },
        { name: 'Mor', val: 7, mult: 10000000, tol: 0.1 },
        { name: 'Gri', val: 8, mult: 100000000, tol: 0.05 },
        { name: 'Beyaz', val: 9, mult: 1000000000, tol: null },
        { name: 'Altın', val: null, mult: 0.1, tol: 5 },
        { name: 'Gümüş', val: null, mult: 0.01, tol: 10 }
    ];

    function buildResistorSelects(bands) {
        var container = document.getElementById('resistor-inputs');
        if (!container) return;
        container.innerHTML = '';
        for (var i = 0; i < bands; i++) {
            var sel = document.createElement('select');
            sel.className = 'panel-select res-band-select';
            sel.style.marginBottom = '8px';
            resColors.forEach(function(c) {
                var isTol = (i === bands - 1);
                var isMult = (i === bands - 2);
                if (isTol && c.tol === null) return;
                if (!isTol && !isMult && c.val === null) return;
                var opt = document.createElement('option');
                opt.value = c.name;
                opt.textContent = c.name;
                sel.appendChild(opt);
            });
            container.appendChild(sel);
        }
    }

    function setupToolkit() {
        var resBands = document.getElementById('res-bands');
        if (resBands) {
            resBands.addEventListener('change', function() {
                buildResistorSelects(parseInt(resBands.value, 10));
            });
            buildResistorSelects(parseInt(resBands.value, 10));
        }
        var resCalc = document.getElementById('res-calc');
        if (resCalc) {
            resCalc.addEventListener('click', function() {
                var selects = document.querySelectorAll('.res-band-select');
                var bands = selects.length;
                if (bands === 4) {
                    var b1 = resColors.find(c => c.name === selects[0].value).val;
                    var b2 = resColors.find(c => c.name === selects[1].value).val;
                    var mult = resColors.find(c => c.name === selects[2].value).mult;
                    var tol = resColors.find(c => c.name === selects[3].value).tol;
                    var val = (b1 * 10 + b2) * mult;
                    document.getElementById('res-result').textContent = formatNum(val) + ' Ω ±' + tol + '%';
                } else if (bands === 5) {
                    var b1 = resColors.find(c => c.name === selects[0].value).val;
                    var b2 = resColors.find(c => c.name === selects[1].value).val;
                    var b3 = resColors.find(c => c.name === selects[2].value).val;
                    var mult = resColors.find(c => c.name === selects[3].value).mult;
                    var tol = resColors.find(c => c.name === selects[4].value).tol;
                    var val = (b1 * 100 + b2 * 10 + b3) * mult;
                    document.getElementById('res-result').textContent = formatNum(val) + ' Ω ±' + tol + '%';
                }
            });
        }
        
        var opampCalc = document.getElementById('opamp-calc');
        if (opampCalc) {
            opampCalc.addEventListener('click', function() {
                var type = document.getElementById('opamp-type').value;
                var rin = parseFloat(document.getElementById('opamp-rin').value) || 0;
                var rf = parseFloat(document.getElementById('opamp-rf').value) || 0;
                if (rin === 0) return;
                var gain = type === 'inv' ? -rf / rin : 1 + rf / rin;
                document.getElementById('opamp-result').textContent = 'Kazanç (Av) = ' + formatNum(gain);
            });
        }
        
        var gearCalc = document.getElementById('gear-calc');
        if (gearCalc) {
            gearCalc.addEventListener('click', function() {
                var driver = parseFloat(document.getElementById('gear-driver').value) || 0;
                var driven = parseFloat(document.getElementById('gear-driven').value) || 0;
                if (driver === 0) return;
                var ratio = driven / driver;
                document.getElementById('gear-result').textContent = 'Oran = ' + formatNum(ratio) + ' : 1';
            });
        }
        
        var vdCalc = document.getElementById('vd-calc');
        if (vdCalc) {
            vdCalc.addEventListener('click', function() {
                var vin = parseFloat(document.getElementById('vd-vin').value) || 0;
                var r1 = parseFloat(document.getElementById('vd-r1').value) || 0;
                var r2 = parseFloat(document.getElementById('vd-r2').value) || 0;
                if ((r1 + r2) === 0) return;
                var vout = vin * (r2 / (r1 + r2));
                var resEl = document.getElementById('vd-result');
                if (resEl) resEl.textContent = 'V_out = ' + formatNum(vout) + ' V';
            });
        }
        
        var motorCalc = document.getElementById('motor-calc');
        if (motorCalc) {
            motorCalc.addEventListener('click', function() {
                var t = parseFloat(document.getElementById('motor-torque').value) || 0;
                var rpm = parseFloat(document.getElementById('motor-rpm').value) || 0;
                var kw = (t * rpm) / 9550;
                var hp = kw * 1.34102;
                var resEl = document.getElementById('motor-result');
                if (resEl) resEl.textContent = formatNum(kw) + ' kW / ' + formatNum(hp) + ' HP';
            });
        }
        
        var vdropCalc = document.getElementById('vdrop-calc');
        if (vdropCalc) {
            vdropCalc.addEventListener('click', function() {
                var sys = document.getElementById('vdrop-sys').value; // mono, tri
                var cond = document.getElementById('vdrop-cond').value; // cu, al
                var vSrc = parseFloat(document.getElementById('vdrop-volt').value) || 220;
                var len = parseFloat(document.getElementById('vdrop-len').value) || 50;
                var amp = parseFloat(document.getElementById('vdrop-amp').value) || 10;
                var sec = parseFloat(document.getElementById('vdrop-sec').value) || 2.5;
                
                var rho = (cond === 'al') ? 0.0282 : 0.0175;
                var vd = 0;
                if (sys === 'mono') {
                    vd = (2 * len * amp * rho) / sec;
                } else {
                    vd = (1.73205 * len * amp * rho) / sec; // sqrt(3) ~ 1.73205
                }
                
                var perc = (vd / vSrc) * 100;
                
                var resEl = document.getElementById('vdrop-result');
                if (resEl) resEl.textContent = formatNum(vd) + ' V (% ' + formatNum(perc) + ' Düşüm)';
            });
        }
        
        var ledCalc = document.getElementById('led-calc');
        if (ledCalc) {
            ledCalc.addEventListener('click', function() {
                var vs = parseFloat(document.getElementById('led-vs').value) || 0;
                var vf = parseFloat(document.getElementById('led-vf').value) || 0;
                var ledIfEl = document.getElementById('led-if');
                var i_f = parseFloat(ledIfEl ? ledIfEl.value : 0) || 0;
                if (i_f === 0) return;
                
                var currentA = i_f;
                var unitSel = document.getElementById('led-if-unit');
                if (unitSel && unitSel.value === 'mA') {
                    currentA = i_f / 1000;
                } else if (!unitSel) {
                    if (ledIfEl && ledIfEl.dataset.unit === 'mA') {
                        currentA = i_f / 1000;
                    } else if (i_f >= 1) { 
                        currentA = i_f / 1000;
                    }
                }
                
                var r = (vs - vf) / currentA;
                var resEl = document.getElementById('led-result');
                if (resEl) resEl.textContent = 'R = ' + formatNum(r) + ' Ω';
            });
        }
    }

    // =========================================================
    // 10. MATRIX, EQN, CMPLX, BASE, CONST, CONV (Preserved)
    // =========================================================

    // --- BASE-N ---
    var activeBase = 'dec';
    function getBaseRadix(base) {
        switch (base) {
            case 'dec': return 10;
            case 'hex': return 16;
            case 'oct': return 8;
            case 'bin': return 2;
            default: return 10;
        }
    }
    function convertBase() {
        var input = document.getElementById('base-input');
        if (!input) return;
        var val = input.value.trim();
        if (val === '') {
            document.getElementById('base-dec').textContent = '0';
            document.getElementById('base-hex').textContent = '0';
            document.getElementById('base-oct').textContent = '0';
            document.getElementById('base-bin').textContent = '0';
            return;
        }
        var radix = getBaseRadix(activeBase);
        var num = parseInt(val, radix);
        if (isNaN(num)) {
            document.getElementById('base-dec').textContent = 'Hata';
            document.getElementById('base-hex').textContent = 'Hata';
            document.getElementById('base-oct').textContent = 'Hata';
            document.getElementById('base-bin').textContent = 'Hata';
            return;
        }
        document.getElementById('base-dec').textContent = num.toString(10).toUpperCase();
        document.getElementById('base-hex').textContent = num.toString(16).toUpperCase();
        document.getElementById('base-oct').textContent = num.toString(8);
        document.getElementById('base-bin').textContent = (num >>> 0).toString(2);
    }
    function setActiveBase(btn) {
        document.querySelectorAll('.btn--base').forEach(function (b) { b.classList.remove('btn--base-active'); });
        btn.classList.add('btn--base-active');
        activeBase = btn.getAttribute('data-base');
        convertBase();
    }
    function calcLogic() {
        var a = parseInt(document.getElementById('logic-a').value, 10);
        var b = parseInt(document.getElementById('logic-b').value, 10);
        var op = document.getElementById('logic-op').value;
        var resultEl = document.getElementById('logic-result');
        if (!resultEl) return;
        if (isNaN(a)) { resultEl.textContent = 'A değeri geçersiz'; return; }
        var result;
        switch (op) {
            case 'and': if (isNaN(b)) { resultEl.textContent = 'B değeri geçersiz'; return; } result = a & b; break;
            case 'or': if (isNaN(b)) { resultEl.textContent = 'B değeri geçersiz'; return; } result = a | b; break;
            case 'xor': if (isNaN(b)) { resultEl.textContent = 'B değeri geçersiz'; return; } result = a ^ b; break;
            case 'not': result = ~a; break;
            case 'shl': if (isNaN(b)) { resultEl.textContent = 'B değeri geçersiz'; return; } result = a << b; break;
            case 'shr': if (isNaN(b)) { resultEl.textContent = 'B değeri geçersiz'; return; } result = a >> b; break;
            default: result = 0;
        }
        resultEl.textContent = 'DEC: ' + result + ' | HEX: ' + (result >>> 0).toString(16).toUpperCase() +
            ' | BIN: ' + (result >>> 0).toString(2);
    }

    // --- MATRIX ---
    function buildMatrixGrids(size) {
        var gridA = document.getElementById('matrix-a-grid');
        var gridB = document.getElementById('matrix-b-grid');
        if (!gridA || !gridB) return;
        gridA.innerHTML = ''; gridB.innerHTML = '';
        gridA.style.gridTemplateColumns = 'repeat(' + size + ', 1fr)';
        gridB.style.gridTemplateColumns = 'repeat(' + size + ', 1fr)';
        for (var i = 0; i < size * size; i++) {
            var inpA = document.createElement('input'); inpA.type = 'number'; inpA.className = 'matrix-cell'; inpA.value = '0'; inpA.step = 'any'; gridA.appendChild(inpA);
            var inpB = document.createElement('input'); inpB.type = 'number'; inpB.className = 'matrix-cell'; inpB.value = '0'; inpB.step = 'any'; gridB.appendChild(inpB);
        }
    }
    function getMatrix(gridEl) {
        if (!gridEl) return [];
        var inputs = gridEl.querySelectorAll('input');
        var size = Math.round(Math.sqrt(inputs.length));
        var matrix = [];
        for (var r = 0; r < size; r++) {
            var row = [];
            for (var c = 0; c < size; c++) row.push(parseFloat(inputs[r * size + c].value) || 0);
            matrix.push(row);
        }
        return matrix;
    }
    function formatMatrixResult(m) {
        if (typeof m === 'number' || (typeof m === 'object' && m.isBigNumber)) return formatNum(m);
        if (!Array.isArray(m)) return String(m);
        var lines = [];
        for (var r = 0; r < m.length; r++) {
            if (Array.isArray(m[r])) {
                var row = [];
                for (var c = 0; c < m[r].length; c++) row.push(formatNum(m[r][c]));
                lines.push('[ ' + row.join(',  ') + ' ]');
            } else {
                lines.push(formatNum(m[r]));
            }
        }
        return lines.join('\n');
    }
    function doMatrixOp(op) {
        var resultEl = document.getElementById('matrix-result');
        if (!resultEl) return;
        try {
            var A = getMatrix(document.getElementById('matrix-a-grid'));
            var B = getMatrix(document.getElementById('matrix-b-grid'));
            var result;
            switch (op) {
                case 'det': result = math.det(A); break;
                case 'inv': result = math.inv(A); break;
                case 'trans': result = math.transpose(A); break;
                case 'detB': result = math.det(B); break;
                case 'invB': result = math.inv(B); break;
                case 'transB': result = math.transpose(B); break;
                case 'add': result = math.add(A, B); break;
                case 'mul': result = math.multiply(A, B); break;
                default: result = 'Bilinmeyen işlem';
            }
            resultEl.textContent = formatMatrixResult(result);
        } catch (e) { resultEl.textContent = 'Hata: ' + e.message; }
    }

    // --- EQUATION ---
    function updateEqnInputs(type) {
        var container = document.getElementById('eqn-inputs');
        var linearSizeContainer = document.getElementById('eqn-linear-size-container');
        var polyDegreeContainer = document.getElementById('eqn-poly-degree-container');
        var resultEl = document.getElementById('eqn-result');
        if (!container) return;
        container.innerHTML = '';
        container.style.gridTemplateColumns = '';
        if (resultEl) resultEl.textContent = '—';
        
        if (linearSizeContainer) linearSizeContainer.style.display = 'none';
        if (polyDegreeContainer) polyDegreeContainer.style.display = 'none';

        switch (type) {
            case 'quad':
                container.innerHTML = '<input type="number" class="panel-input" id="eqn-a" placeholder="a" step="any"><input type="number" class="panel-input" id="eqn-b" placeholder="b" step="any"><input type="number" class="panel-input" id="eqn-c" placeholder="c" step="any">';
                container.style.gridTemplateColumns = 'repeat(3, 1fr)';
                break;
            case 'cubic':
                container.innerHTML = '<input type="number" class="panel-input" id="eqn-a" placeholder="a" step="any"><input type="number" class="panel-input" id="eqn-b" placeholder="b" step="any"><input type="number" class="panel-input" id="eqn-c" placeholder="c" step="any"><input type="number" class="panel-input" id="eqn-d" placeholder="d" step="any">';
                container.style.gridTemplateColumns = 'repeat(4, 1fr)';
                break;
            case 'poly':
                if (polyDegreeContainer) polyDegreeContainer.style.display = '';
                buildPolyInputs();
                break;
            case 'linear':
                if (linearSizeContainer) linearSizeContainer.style.display = '';
                buildLinearInputs();
                break;
            case 'integral':
                container.innerHTML = '<input type="text" class="panel-input" id="eqn-fx" placeholder="f(x)  örn: x^2+1"><input type="number" class="panel-input" id="eqn-lower" placeholder="a (alt sınır)" step="any"><input type="number" class="panel-input" id="eqn-upper" placeholder="b (üst sınır)" step="any">';
                container.style.gridTemplateColumns = 'repeat(3, 1fr)';
                break;
            case 'sym_integral':
                container.innerHTML = '<div style="grid-column: 1 / -1; margin-bottom: 10px;"><label class="panel-label">İntegrali Alınacak İfade (x\'e göre)</label></div><math-field id="eqn-sym-int" class="panel-input" style="font-size:24px; padding:12px; grid-column: 1 / -1;"></math-field>';
                break;
            case 'derivative':
                container.innerHTML = '<input type="text" class="panel-input" id="eqn-fx" placeholder="f(x)  örn: sin(x)+x^2"><input type="number" class="panel-input" id="eqn-x0" placeholder="x₀" step="any">';
                container.style.gridTemplateColumns = 'repeat(2, 1fr)';
                break;
            case 'sym_deriv':
                container.innerHTML = '<div style="grid-column: 1 / -1; margin-bottom: 10px;"><label class="panel-label">Türevi Alınacak İfade (x\'e göre)</label></div><math-field id="eqn-sym-der" class="panel-input" style="font-size:24px; padding:12px; grid-column: 1 / -1;"></math-field>';
                break;
        }
    }
    
    function buildPolyInputs() {
        var container = document.getElementById('eqn-inputs');
        var degreeInput = document.getElementById('eqn-poly-degree');
        if (!container || !degreeInput) return;
        var n = parseInt(degreeInput.value, 10);
        if (isNaN(n) || n < 2) n = 2; if (n > 10) n = 10;
        
        container.innerHTML = '';
        container.style.gridTemplateColumns = 'repeat(' + (n <= 5 ? n + 1 : 4) + ', 1fr)';
        for (var i = n; i >= 0; i--) {
            var label = i === 0 ? 'Sabit' : 'x^' + i;
            var wrapper = document.createElement('div');
            wrapper.style.display = 'flex';
            wrapper.style.flexDirection = 'column';
            wrapper.style.gap = '4px';
            wrapper.innerHTML = '<label style="font-size: 11px; color: var(--text-secondary);">' + label + ' Katsayısı</label>' +
                                '<input type="number" class="panel-input eqn-poly-coeff" data-degree="' + i + '" value="0" step="any">';
            container.appendChild(wrapper);
        }
    }
    function buildLinearInputs() {
        var container = document.getElementById('eqn-inputs');
        var sizeInput = document.getElementById('eqn-linear-size');
        if (!container || !sizeInput) return;
        var n = parseInt(sizeInput.value, 10);
        if (isNaN(n) || n < 2) n = 2; if (n > 10) n = 10;
        container.innerHTML = '';
        container.style.gridTemplateColumns = 'repeat(' + (n + 1) + ', 1fr)';
        for (var c = 0; c < n; c++) {
            var hdr = document.createElement('span'); hdr.className = 'eqn-header'; hdr.textContent = 'x' + (c + 1); container.appendChild(hdr);
        }
        var hdrConst = document.createElement('span'); hdrConst.className = 'eqn-header'; hdrConst.textContent = '= b'; container.appendChild(hdrConst);
        for (var r = 0; r < n; r++) {
            for (var c2 = 0; c2 <= n; c2++) {
                var inp = document.createElement('input'); inp.type = 'number'; inp.className = 'panel-input eqn-cell'; inp.step = 'any'; inp.value = '0';
                if (c2 < n) { inp.placeholder = 'a' + (r + 1) + (c2 + 1); } else { inp.placeholder = 'b' + (r + 1); }
                container.appendChild(inp);
            }
        }
    }
    function solveEquation(type) {
        var resultEl = document.getElementById('eqn-result');
        if (!resultEl) return;
        try {
            switch (type) {
                case 'quad': {
                    var a = parseFloat(document.getElementById('eqn-a').value) || 0;
                    var b = parseFloat(document.getElementById('eqn-b').value) || 0;
                    var c = parseFloat(document.getElementById('eqn-c').value) || 0;
                    if (a === 0) { resultEl.textContent = 'a ≠ 0 olmalıdır'; return; }
                    var disc = b * b - 4 * a * c;
                    if (disc >= 0) {
                        var x1 = (-b + Math.sqrt(disc)) / (2 * a);
                        var x2 = (-b - Math.sqrt(disc)) / (2 * a);
                        resultEl.textContent = 'x₁ = ' + formatNum(x1) + '\nx₂ = ' + formatNum(x2) + '\nΔ = ' + formatNum(disc);
                    } else {
                        var realPart = -b / (2 * a);
                        var imagPart = Math.sqrt(-disc) / (2 * a);
                        resultEl.textContent = 'x₁ = ' + formatNum(realPart) + ' + ' + formatNum(imagPart) + 'i' +
                            '\nx₂ = ' + formatNum(realPart) + ' − ' + formatNum(imagPart) + 'i' + '\nΔ = ' + formatNum(disc) + ' (Karmaşık kökler)';
                    }
                    break;
                }
                case 'cubic': {
                    var ca = parseFloat(document.getElementById('eqn-a').value) || 0;
                    var cb = parseFloat(document.getElementById('eqn-b').value) || 0;
                    var cc = parseFloat(document.getElementById('eqn-c').value) || 0;
                    var cd = parseFloat(document.getElementById('eqn-d').value) || 0;
                    if (ca === 0) { resultEl.textContent = 'a ≠ 0 olmalıdır'; return; }
                    var p = (3 * ca * cc - cb * cb) / (3 * ca * ca);
                    var q = (2 * cb * cb * cb - 9 * ca * cb * cc + 27 * ca * ca * cd) / (27 * ca * ca * ca);
                    var discrim = q * q / 4 + p * p * p / 27;
                    var offset = -cb / (3 * ca);
                    var roots = [];
                    if (discrim > 1e-12) {
                        var sqrtD = Math.sqrt(discrim);
                        var u = Math.cbrt(-q / 2 + sqrtD);
                        var v = Math.cbrt(-q / 2 - sqrtD);
                        roots.push(u + v + offset);
                        var realP = -(u + v) / 2 + offset;
                        var imagP = Math.sqrt(3) / 2 * (u - v);
                        resultEl.textContent = 'x₁ = ' + formatNum(roots[0]) + '\nx₂ = ' + formatNum(realP) + ' + ' + formatNum(Math.abs(imagP)) + 'i\nx₃ = ' + formatNum(realP) + ' − ' + formatNum(Math.abs(imagP)) + 'i';
                    } else if (Math.abs(discrim) <= 1e-12) {
                        if (Math.abs(q) < 1e-15) { roots = [offset, offset, offset]; } else { var u2 = Math.cbrt(-q / 2); roots = [2 * u2 + offset, -u2 + offset, -u2 + offset]; }
                        resultEl.textContent = 'x₁ = ' + formatNum(roots[0]) + '\nx₂ = ' + formatNum(roots[1]) + '\nx₃ = ' + formatNum(roots[2]);
                    } else {
                        var m = 2 * Math.sqrt(-p / 3);
                        var theta = Math.acos(3 * q / (p * m)) / 3;
                        roots = [m * Math.cos(theta) + offset, m * Math.cos(theta - 2 * Math.PI / 3) + offset, m * Math.cos(theta - 4 * Math.PI / 3) + offset];
                        resultEl.textContent = 'x₁ = ' + formatNum(roots[0]) + '\nx₂ = ' + formatNum(roots[1]) + '\nx₃ = ' + formatNum(roots[2]);
                    }
                    break;
                }
                case 'linear': {
                    var sizeInput = document.getElementById('eqn-linear-size');
                    var n = parseInt(sizeInput.value, 10);
                    if (isNaN(n) || n < 2) n = 2;
                    var container = document.getElementById('eqn-inputs');
                    var inputs = container.querySelectorAll('input.eqn-cell');
                    var coeffMatrix = []; var constVector = []; var idx = 0;
                    for (var r = 0; r < n; r++) {
                        var row = [];
                        for (var c2 = 0; c2 < n; c2++) { row.push(parseFloat(inputs[idx].value) || 0); idx++; }
                        coeffMatrix.push(row); constVector.push([parseFloat(inputs[idx].value) || 0]); idx++;
                    }
                    var solution = math.lusolve(coeffMatrix, constVector);
                    var lines = [];
                    for (var i = 0; i < solution.length; i++) lines.push('x' + (i + 1) + ' = ' + formatNum(solution[i][0]));
                    resultEl.textContent = lines.join('\n');
                    break;
                }
                case 'integral': {
                    var fx = document.getElementById('eqn-fx').value.trim();
                    var lower = parseFloat(document.getElementById('eqn-lower').value);
                    var upper = parseFloat(document.getElementById('eqn-upper').value);
                    if (!fx || isNaN(lower) || isNaN(upper)) { resultEl.textContent = 'Lütfen f(x), a ve b değerlerini girin'; return; }
                    var compiled = math.compile(fx);
                    var nSteps = 1000; var h = (upper - lower) / nSteps;
                    var sum = compiled.evaluate({ x: lower }) + compiled.evaluate({ x: upper });
                    for (var si = 1; si < nSteps; si++) {
                        var xi = lower + si * h;
                        var fxi = compiled.evaluate({ x: xi });
                        if (si % 2 === 0) { sum += 2 * fxi; } else { sum += 4 * fxi; }
                    }
                    var integral = (h / 3) * sum;
                    if (typeof integral === 'object' && integral.isBigNumber) integral = integral.toNumber();
                    resultEl.textContent = '∫f(x)dx = ' + formatNum(integral) + '\n[' + formatNum(lower) + ', ' + formatNum(upper) + '] aralığında';
                    break;
                }
                case 'derivative': {
                    var fx2 = document.getElementById('eqn-fx').value.trim();
                    var x0 = parseFloat(document.getElementById('eqn-x0').value);
                    if (!fx2 || isNaN(x0)) { resultEl.textContent = 'Lütfen f(x) ve x₀ değerlerini girin'; return; }
                    var compiled2 = math.compile(fx2);
                    var hd = 1e-8;
                    var fPlus = compiled2.evaluate({ x: x0 + hd });
                    var fMinus = compiled2.evaluate({ x: x0 - hd });
                    if (typeof fPlus === 'object' && fPlus.isBigNumber) fPlus = fPlus.toNumber();
                    if (typeof fMinus === 'object' && fMinus.isBigNumber) fMinus = fMinus.toNumber();
                    var deriv = (fPlus - fMinus) / (2 * hd);
                    resultEl.textContent = "f'(" + formatNum(x0) + ') = ' + formatNum(deriv);
                    break;
                }
                case 'poly': {
                    var degreeInput = document.getElementById('eqn-poly-degree');
                    var n = parseInt(degreeInput.value, 10) || 2;
                    var container = document.getElementById('eqn-inputs');
                    var inputs = container.querySelectorAll('.eqn-poly-coeff');
                    var coeffs = [];
                    for(var i = 0; i <= n; i++) coeffs.push(0);
                    inputs.forEach(function(inp) {
                        var d = parseInt(inp.getAttribute('data-degree'), 10);
                        coeffs[d] = parseFloat(inp.value) || 0;
                    });
                    if (coeffs[n] === 0) { resultEl.textContent = 'En yüksek dereceli katsayı 0 olamaz!'; return; }
                    
                    for (var i = 0; i < n; i++) coeffs[i] = coeffs[i] / coeffs[n];
                    coeffs[n] = 1;
                    
                    var roots = [];
                    var r0 = math.complex(0.4, 0.9);
                    for (var i = 0; i < n; i++) roots.push(math.pow(r0, i));
                    
                    var evalPoly = function(x, cfs) {
                        var res = math.complex(0,0);
                        for (var i = 0; i <= n; i++) {
                            res = math.add(res, math.multiply(cfs[i], math.pow(x, i)));
                        }
                        return res;
                    };
                    
                    var maxIter = 200;
                    for (var iter = 0; iter < maxIter; iter++) {
                        var maxDiff = 0;
                        var newRoots = [];
                        for (var i = 0; i < n; i++) {
                            var p_xi = evalPoly(roots[i], coeffs);
                            var denom = math.complex(1,0);
                            for (var j = 0; j < n; j++) {
                                if (i !== j) denom = math.multiply(denom, math.subtract(roots[i], roots[j]));
                            }
                            var correction = math.divide(p_xi, denom);
                            newRoots.push(math.subtract(roots[i], correction));
                            
                            var diff = math.abs(correction);
                            if (diff > maxDiff) maxDiff = diff;
                        }
                        roots = newRoots;
                        if (maxDiff < 1e-12) break;
                    }
                    
                    var lines = [];
                    for (var i = 0; i < n; i++) {
                        var r = roots[i];
                        if (Math.abs(r.im) < 1e-8) {
                            lines.push('x' + (i+1) + ' = ' + formatNum(r.re));
                        } else {
                            var sign = r.im >= 0 ? '+' : '-';
                            lines.push('x' + (i+1) + ' = ' + formatNum(r.re) + ' ' + sign + ' ' + formatNum(Math.abs(r.im)) + 'i');
                        }
                    }
                    resultEl.textContent = lines.join('\n');
                    break;
                }
                case 'sym_integral': {
                    if (typeof nerdamer === 'undefined') { resultEl.textContent = 'Nerdamer kütüphanesi yüklenemedi.'; return; }
                    var mf = document.getElementById('eqn-sym-int');
                    if (!mf) return;
                    var expr = mf.getValue('ascii-math');
                    if (!expr) { resultEl.textContent = 'Lütfen bir ifade girin'; return; }
                    var exprProcessed = preprocessExpression(expr);
                    var integrated = nerdamer('integrate(' + exprProcessed + ', x)').toTeX();
                    resultEl.innerHTML = '<math-field readonly style="font-size:24px; padding:12px; border:none; background:transparent;">' + integrated + ' + C</math-field>';
                    break;
                }
                case 'sym_deriv': {
                    if (typeof nerdamer === 'undefined') { resultEl.textContent = 'Nerdamer kütüphanesi yüklenemedi.'; return; }
                    var mf = document.getElementById('eqn-sym-der');
                    if (!mf) return;
                    var expr = mf.getValue('ascii-math');
                    if (!expr) { resultEl.textContent = 'Lütfen bir ifade girin'; return; }
                    var exprProcessed = preprocessExpression(expr);
                    var derivative = nerdamer('diff(' + exprProcessed + ', x)').toTeX();
                    resultEl.innerHTML = '<math-field readonly style="font-size:24px; padding:12px; border:none; background:transparent;">' + derivative + '</math-field>';
                    break;
                }
                default: resultEl.textContent = 'Bilinmeyen tip';
            }
        } catch (e) { resultEl.textContent = 'Hata: ' + e.message; }
    }

    // --- COMPLEX ---
    function formatComplex(z) {
        if (typeof z === 'number') return formatNum(z);
        if (typeof z === 'object' && z.isBigNumber) return formatNum(z.toNumber());
        var re = math.re(z); var im = math.im(z);
        if (typeof re === 'object' && re.isBigNumber) re = re.toNumber();
        if (typeof im === 'object' && im.isBigNumber) im = im.toNumber();
        if (im === 0) return formatNum(re);
        if (re === 0) return formatNum(im) + 'i';
        var sign = im >= 0 ? ' + ' : ' − ';
        return formatNum(re) + sign + formatNum(Math.abs(im)) + 'i';
    }
    function calcComplex() {
        var resultEl = document.getElementById('cmplx-result');
        if (!resultEl) return;
        try {
            var a1 = parseFloat(document.getElementById('cmplx-a1').value) || 0;
            var b1 = parseFloat(document.getElementById('cmplx-b1').value) || 0;
            var a2 = parseFloat(document.getElementById('cmplx-a2').value) || 0;
            var b2 = parseFloat(document.getElementById('cmplx-b2').value) || 0;
            var z1 = math.complex(a1, b1); var z2 = math.complex(a2, b2);
            var op = document.getElementById('cmplx-op').value;
            var result;
            switch (op) {
                case 'add': result = math.add(z1, z2); resultEl.textContent = formatComplex(result); break;
                case 'sub': result = math.subtract(z1, z2); resultEl.textContent = formatComplex(result); break;
                case 'mul': result = math.multiply(z1, z2); resultEl.textContent = formatComplex(result); break;
                case 'div': result = math.divide(z1, z2); resultEl.textContent = formatComplex(result); break;
                case 'parallel': result = math.divide(math.multiply(z1, z2), math.add(z1, z2)); resultEl.textContent = 'Z₁∥Z₂ = ' + formatComplex(result); break;
                case 'mod': var modulus = math.abs(z1); if (typeof modulus === 'object' && modulus.isBigNumber) modulus = modulus.toNumber(); resultEl.textContent = '|Z₁| = ' + formatNum(modulus); break;
                case 'conj': result = math.conj(z1); resultEl.textContent = 'Z₁* = ' + formatComplex(result); break;
                case 'mod2': var modulus2 = math.abs(z2); if (typeof modulus2 === 'object' && modulus2.isBigNumber) modulus2 = modulus2.toNumber(); resultEl.textContent = '|Z₂| = ' + formatNum(modulus2); break;
                case 'conj2': result = math.conj(z2); resultEl.textContent = 'Z₂* = ' + formatComplex(result); break;
                case 'polar': {
                    var r = math.abs(z1); if (typeof r === 'object' && r.isBigNumber) r = r.toNumber();
                    var theta = Math.atan2(b1, a1);
                    if (state.angleMode === 'DEG') { var thetaDeg = theta * 180 / Math.PI; resultEl.textContent = formatNum(r) + ' ∠ ' + formatNum(thetaDeg) + '°'; }
                    else { resultEl.textContent = formatNum(r) + ' ∠ ' + formatNum(theta) + ' rad'; }
                    break;
                }
                case 'euler': {
                    var rE = math.abs(z1); if (typeof rE === 'object' && rE.isBigNumber) rE = rE.toNumber();
                    var thetaE = Math.atan2(b1, a1); resultEl.textContent = formatNum(rE) + ' · e^(i·' + formatNum(thetaE) + ')';
                    break;
                }
                default: resultEl.textContent = 'Bilinmeyen işlem';
            }
        } catch (e) { resultEl.textContent = 'Hata: ' + e.message; }
    }

    // --- CONSTANTS ---
    var physicsConstants = [
        { symbol: 'c', name: 'Işık Hızı', value: 299792458, unit: 'm/s' },
        { symbol: 'G', name: 'Yerçekimi Sabiti', value: 6.674e-11, unit: 'N·m²/kg²' },
        { symbol: 'h', name: 'Planck Sabiti', value: 6.626e-34, unit: 'J·s' },
        { symbol: 'ℏ', name: 'Dirac Sabiti (ℏ)', value: 1.0546e-34, unit: 'J·s' },
        { symbol: 'e', name: 'Elektron Yükü', value: 1.602e-19, unit: 'C' },
        { symbol: 'mₑ', name: 'Elektron Kütlesi', value: 9.109e-31, unit: 'kg' },
        { symbol: 'mₚ', name: 'Proton Kütlesi', value: 1.673e-27, unit: 'kg' },
        { symbol: 'mₙ', name: 'Nötron Kütlesi', value: 1.675e-27, unit: 'kg' },
        { symbol: 'Nₐ', name: 'Avogadro Sayısı', value: 6.022e23, unit: '1/mol' },
        { symbol: 'kB', name: 'Boltzmann Sabiti', value: 1.381e-23, unit: 'J/K' },
        { symbol: 'R', name: 'Gaz Sabiti', value: 8.314, unit: 'J/(mol·K)' },
        { symbol: 'σ', name: 'Stefan-Boltzmann', value: 5.670e-8, unit: 'W/(m²·K⁴)' },
        { symbol: 'ε₀', name: 'Vakum Geçirgenliği', value: 8.854e-12, unit: 'F/m' },
        { symbol: 'μ₀', name: 'Vakum Manyetik Gç.', value: 1.2566e-6, unit: 'H/m' },
        { symbol: 'α', name: 'İnce Yapı Sabiti', value: 7.297e-3, unit: '' },
        { symbol: 'Φ₀', name: 'Manyetik Akı Kuantumu', value: 2.068e-15, unit: 'Wb' },
        { symbol: 'g', name: 'Yerçekimi İvmesi', value: 9.80665, unit: 'm/s²' },
        { symbol: 'atm', name: 'Standart Atmosfer', value: 101325, unit: 'Pa' },
        { symbol: 'eV', name: 'Elektron-Volt', value: 1.602e-19, unit: 'J' },
        { symbol: 'u', name: 'Atomik Kütle Birimi', value: 1.661e-27, unit: 'kg' }
    ];
    var customConsts = [];
    var allConstants = [];
    function loadCustomConstants() {
        try {
            var saved = localStorage.getItem('calc-custom-constants');
            return saved ? JSON.parse(saved) : [];
        } catch (e) { return []; }
    }
    function saveCustomConstants(arr) {
        localStorage.setItem('calc-custom-constants', JSON.stringify(arr));
    }
    function renderConstants(filter) {
        var listEl = document.getElementById('const-list');
        if (!listEl) return;
        listEl.innerHTML = '';
        var filterLower = (filter || '').toLowerCase();
        allConstants.forEach(function (c) {
            var searchStr = (c.symbol + ' ' + c.name + ' ' + c.unit).toLowerCase();
            if (filterLower && searchStr.indexOf(filterLower) === -1) return;
            var item = document.createElement('div'); item.className = 'const-item';
            item.innerHTML = '<span class="const-symbol">' + c.symbol + '</span><span class="const-name">' + c.name + '</span><span class="const-value">' + c.value.toExponential(4) + '</span><span class="const-unit">' + c.unit + '</span>';
            item.addEventListener('click', function () {
                switchTab('standard');
                if (mathField) {
                    mathField.insert(String(c.value));
                    mathField.focus();
                    livePreview();
                }
            });
            listEl.appendChild(item);
        });
    }

    // --- CONVERSIONS ---
    var Categories = {
        Length: { base: 'm', units: { km: 1000, m: 1, cm: 0.01, mm: 0.001, um: 0.000001, nm: 0.000000001, inch: 0.0254, ft: 0.3048 } },
        Pressure: { base: 'Pa', units: { GPa: 1000000000, MPa: 1000000, kPa: 1000, 'kN/m²': 1000, Pa: 1, bar: 100000, psi: 6894.76, atm: 101325 } },
        Power: { base: 'W', units: { GW: 1000000000, MW: 1000000, kW: 1000, W: 1, mW: 0.001, hp: 745.7 } },
        Torque: { base: 'Nm', units: { kNm: 1000, Nm: 1, Nmm: 0.001, kgfm: 9.80665, 'lb-ft': 1.355818 } },
        Temperature: {
            base: 'K',
            units: {
                K: { toBase: function(v){return v;}, fromBase: function(v){return v;} },
                C: { toBase: function(v){return v + 273.15;}, fromBase: function(v){return v - 273.15;} },
                F: { toBase: function(v){return (v - 32) * 5/9 + 273.15;}, fromBase: function(v){return (v - 273.15) * 9/5 + 32;} }
            }
        }
    };
    function setupConversions() {
        var grid = document.getElementById('conv-grid');
        if (!grid) return;
        grid.innerHTML = '';
        for (var cat in Categories) {
            var div = document.createElement('div');
            div.className = 'conv-card conv-row';
            div.innerHTML = '<label class="panel-label" style="font-weight:bold; margin-bottom:12px; display:block;">' + cat + '</label>' +
                '<div style="display:flex; gap:12px; align-items:center; flex-wrap:wrap; position:relative; z-index:10;">' +
                '<input type="number" class="panel-input conv-inp-1" step="any" style="flex: 1 1 150px; min-width: 120px;">' +
                '<select class="panel-select conv-sel-1" style="flex: 1 1 150px; min-width: 120px; z-index:20; position:relative;"></select>' +
                '<span style="color:var(--neon-blue); font-weight:bold; font-size:18px;">=</span>' +
                '<input type="number" class="panel-input conv-inp-2" step="any" style="flex: 1 1 150px; min-width: 120px;">' +
                '<select class="panel-select conv-sel-2" style="flex: 1 1 150px; min-width: 120px; z-index:20; position:relative;"></select>' +
                '</div>';
            
            var sel1 = div.querySelector('.conv-sel-1');
            var sel2 = div.querySelector('.conv-sel-2');
            var inp1 = div.querySelector('.conv-inp-1');
            var inp2 = div.querySelector('.conv-inp-2');
            
            // Add active focus logic for PDF export
            [inp1, inp2, sel1, sel2].forEach(function(el) {
                el.addEventListener('focus', function() {
                    document.querySelectorAll('.conv-row').forEach(function(r) { r.classList.remove('active'); });
                    div.classList.add('active');
                });
            });
            
            var units = Categories[cat].units;
            var unitNames = Object.keys(units);
            var t = translations[state.currentLang] || translations.tr;
            unitNames.forEach(function(u) {
                var label = t['unit_' + u] || u;
                sel1.innerHTML += '<option value="' + u + '">' + label + '</option>';
                sel2.innerHTML += '<option value="' + u + '">' + label + '</option>';
            });
            sel2.selectedIndex = unitNames.length > 1 ? 1 : 0;
            
            (function(c, unitsObj, i1, i2, s1, s2) {
                function convert(sourceVal, sourceUnit, targetUnit) {
                    var u1 = unitsObj[sourceUnit];
                    var u2 = unitsObj[targetUnit];
                    var baseVal;
                    if (typeof u1 === 'object' && u1.toBase) baseVal = u1.toBase(sourceVal);
                    else baseVal = sourceVal * u1;
                    
                    if (typeof u2 === 'object' && u2.fromBase) return u2.fromBase(baseVal);
                    else return baseVal / u2;
                }
                function update2() {
                    var v = parseFloat(i1.value);
                    if(isNaN(v)) { i2.value = ''; return; }
                    i2.value = parseFloat(convert(v, s1.value, s2.value).toPrecision(10));
                }
                function update1() {
                    var v = parseFloat(i2.value);
                    if(isNaN(v)) { i1.value = ''; return; }
                    i1.value = parseFloat(convert(v, s2.value, s1.value).toPrecision(10));
                }
                i1.addEventListener('input', update2);
                s1.addEventListener('change', update2);
                i2.addEventListener('input', update1);
                s2.addEventListener('change', update1);
            })(cat, units, inp1, inp2, sel1, sel2);
            
            grid.appendChild(div);
        }
    }

    // =========================================================
    // 11. MOBILE MENU
    // =========================================================
    function setupMobileMenu() {
        var menuBtn = document.getElementById('mobile-menu-btn');
        var sidebar = document.getElementById('sidebar');
        if (menuBtn && sidebar) {
            menuBtn.addEventListener('click', function () {
                sidebar.classList.toggle('sidebar--open');
            });
        }
    }

    // =========================================================
    // 12. INIT
    // =========================================================
    function init() {
        resultEl = document.getElementById('result');

        // Sidebar
        var sidebarMenu = document.getElementById('sidebar-menu');
        if (sidebarMenu) {
            sidebarMenu.addEventListener('click', function (e) {
                var btn = e.target.closest('.sidebar__item');
                if (!btn) return;
                var tabId = btn.getAttribute('data-tab');
                if (tabId) switchTab(tabId);
            });
        }

        var langToggle = document.getElementById('lang-toggle');
        if (langToggle) {
            langToggle.addEventListener('click', function () {
                changeLanguage(state.currentLang === 'tr' ? 'en' : 'tr');
            });
        }

        var modeToggle = document.getElementById('mode-toggle');
        if (modeToggle) modeToggle.addEventListener('click', toggleAngleMode);

        var histUp = document.getElementById('history-up');
        var histDown = document.getElementById('history-down');
        if (histUp) histUp.addEventListener('click', function () { navigateHistory('up'); });
        if (histDown) histDown.addEventListener('click', function () { navigateHistory('down'); });

        var standardPanel = document.getElementById('panel-standard');
        if (standardPanel) {
            standardPanel.addEventListener('click', function (e) {
                var btn = e.target.closest('[data-action]');
                if (!btn) return;
                handleButtonAction(btn.getAttribute('data-action'), btn);
            });
        }

        // Matrix
        buildMatrixGrids(3);
        var matrixSizeInput = document.getElementById('matrix-size');
        if (matrixSizeInput) {
            matrixSizeInput.addEventListener('change', function () {
                var s = parseInt(matrixSizeInput.value, 10);
                if (isNaN(s) || s < 1) s = 1; if (s > 10) s = 10;
                buildMatrixGrids(s);
            });
        }
        var matrixPanel = document.getElementById('panel-matrix');
        if (matrixPanel) {
            matrixPanel.addEventListener('click', function (e) {
                var btn = e.target.closest('[data-mop]');
                if (!btn) return;
                doMatrixOp(btn.getAttribute('data-mop'));
            });
        }

        // Equation
        var eqnType = document.getElementById('eqn-type');
        if (eqnType) {
            eqnType.addEventListener('change', function () { updateEqnInputs(eqnType.value); });
            updateEqnInputs(eqnType.value);
        }
        var eqnLinearSize = document.getElementById('eqn-linear-size');
        if (eqnLinearSize) {
            eqnLinearSize.addEventListener('change', function () {
                if (eqnType && eqnType.value === 'linear') buildLinearInputs();
            });
        }
        var eqnPolyDegree = document.getElementById('eqn-poly-degree');
        if (eqnPolyDegree) {
            eqnPolyDegree.addEventListener('change', function () {
                if (eqnType && eqnType.value === 'poly') buildPolyInputs();
            });
        }
        var eqnSolveBtn = document.getElementById('eqn-solve');
        if (eqnSolveBtn) eqnSolveBtn.addEventListener('click', function () { if (eqnType) solveEquation(eqnType.value); });

        // Complex
        var cmplxCalcBtn = document.getElementById('cmplx-calc');
        if (cmplxCalcBtn) cmplxCalcBtn.addEventListener('click', calcComplex);

        // Base-N
        var baseInput = document.getElementById('base-input');
        if (baseInput) baseInput.addEventListener('input', convertBase);
        document.querySelectorAll('.btn--base').forEach(function (btn) {
            btn.addEventListener('click', function () { setActiveBase(btn); });
        });
        var logicCalcBtn = document.getElementById('logic-calc');
        if (logicCalcBtn) logicCalcBtn.addEventListener('click', calcLogic);

        // Init new features
        initMathLive();
        setupPlot();
        setupMacro();
        setupToolkit();
        bindExport();

        customConsts = loadCustomConstants();
        allConstants = physicsConstants.concat(customConsts);

        var btnAddConst = document.getElementById('custom-const-add');
        if (btnAddConst) {
            btnAddConst.addEventListener('click', function() {
                var name = document.getElementById('custom-const-name').value.trim();
                var sym = document.getElementById('custom-const-sym').value.trim();
                var val = parseFloat(document.getElementById('custom-const-val').value);
                var unit = document.getElementById('custom-const-unit').value.trim();
                if (name && sym && !isNaN(val)) {
                    var newConst = { symbol: sym, name: name, value: val, unit: unit };
                    customConsts.push(newConst);
                    saveCustomConstants(customConsts);
                    allConstants = physicsConstants.concat(customConsts);
                    renderConstants('');
                    document.getElementById('custom-const-name').value = '';
                    document.getElementById('custom-const-sym').value = '';
                    document.getElementById('custom-const-val').value = '';
                    document.getElementById('custom-const-unit').value = '';
                }
            });
        }

        renderConstants('');
        var constSearch = document.getElementById('const-search');
        if (constSearch) constSearch.addEventListener('input', function () { renderConstants(constSearch.value); });

        setupConversions();
        setupMobileMenu();

        updateResultDisplay();
        changeLanguage('tr');

        setupCodeExport();
        setupDigitalLogic();
        setupBOM();
        setupPCB();
        setup3DPrint();
    }

    // =========================================================
    // 13. CODE EXPORT
    // =========================================================
    function setupCodeExport() {
        var btn = document.getElementById('code-export-btn');
        var modal = document.getElementById('code-export-modal');
        var closeBtn = document.getElementById('code-modal-close');
        if (!btn || !modal) return;

        btn.addEventListener('click', function() {
            var expr = (mathField ? mathField.getValue('ascii-math') : '') || '';
            var result = state.result || '0';
            if (!expr || expr.trim() === '') {
                expr = 'expression';
            }
            // Build a safe variable name
            var safeName = 'calculate';

            // Build Python code
            var pyCode = '# MühendisCalc - Python Kod Çıktısı\n';
            pyCode += 'import math\n\n';
            pyCode += 'def ' + safeName + '():\n';
            pyCode += '    # İfade: ' + expr + '\n';
            pyCode += '    result = ' + buildPythonExpr(expr) + '\n';
            pyCode += '    return result\n\n';
            pyCode += '# Çağrı\n';
            pyCode += 'print("Sonuç:", ' + safeName + '())\n';
            pyCode += '# Beklenen çıktı: ' + result;

            // Build C++ / Arduino code
            var cppCode = '// MühendisCalc - C++ / Arduino Kod Çıktısı\n';
            cppCode += '#include <math.h>\n\n';
            cppCode += 'double ' + safeName + '() {\n';
            cppCode += '    // İfade: ' + expr + '\n';
            cppCode += '    double result = ' + buildCppExpr(expr) + ';\n';
            cppCode += '    return result;\n';
            cppCode += '}\n\n';
            cppCode += 'void setup() {\n';
            cppCode += '    Serial.begin(9600);\n';
            cppCode += '    Serial.print("Sonuc: ");\n';
            cppCode += '    Serial.println(' + safeName + '());\n';
            cppCode += '    // Beklenen: ' + result + '\n';
            cppCode += '}\n\n';
            cppCode += 'void loop() {}';

            document.getElementById('code-python').textContent = pyCode;
            document.getElementById('code-cpp').textContent = cppCode;
            modal.style.display = 'flex';
        });

        closeBtn.addEventListener('click', function() {
            modal.style.display = 'none';
        });

        modal.addEventListener('click', function(e) {
            if (e.target === modal) modal.style.display = 'none';
        });

        document.querySelectorAll('.btn--copy-code').forEach(function(b) {
            b.addEventListener('click', function() {
                var targetId = b.getAttribute('data-target');
                var text = document.getElementById(targetId).textContent;
                navigator.clipboard.writeText(text).then(function() {
                    var orig = b.textContent;
                    b.textContent = '✅ Kopyalandı!';
                    setTimeout(function() { b.textContent = orig; }, 1800);
                });
            });
        });
    }

    function buildPythonExpr(expr) {
        var s = expr;
        s = s.replace(/\\\\/g, '');
        s = s.replace(/\\pi/g, 'math.pi');
        s = s.replace(/pi/g, 'math.pi');
        s = s.replace(/\\sin/g, 'math.sin').replace(/\bsin\b/g, 'math.sin');
        s = s.replace(/\\cos/g, 'math.cos').replace(/\bcos\b/g, 'math.cos');
        s = s.replace(/\\tan/g, 'math.tan').replace(/\btan\b/g, 'math.tan');
        s = s.replace(/\\arcsin|asin/g, 'math.asin');
        s = s.replace(/\\arccos|acos/g, 'math.acos');
        s = s.replace(/\\arctan|atan/g, 'math.atan');
        s = s.replace(/\\sqrt/g, 'math.sqrt').replace(/\bsqrt\b/g, 'math.sqrt');
        s = s.replace(/\\log/g, 'math.log10').replace(/\blog\b(?!10)/g, 'math.log10');
        s = s.replace(/\\ln|ln/g, 'math.log');
        s = s.replace(/\^/g, '**');
        s = s.replace(/×/g, '*').replace(/÷/g, '/');
        return s || '0';
    }

    function buildCppExpr(expr) {
        var s = expr;
        s = s.replace(/\\\\/g, '');
        s = s.replace(/\\pi|pi/g, 'M_PI');
        s = s.replace(/\\sin|\bsin\b/g, 'sin');
        s = s.replace(/\\cos|\bcos\b/g, 'cos');
        s = s.replace(/\\tan|\btan\b/g, 'tan');
        s = s.replace(/\\arcsin|asin/g, 'asin');
        s = s.replace(/\\arccos|acos/g, 'acos');
        s = s.replace(/\\arctan|atan/g, 'atan');
        s = s.replace(/\\sqrt|\bsqrt\b/g, 'sqrt');
        s = s.replace(/\\log|\blog\b/g, 'log10');
        s = s.replace(/\\ln|\bln\b/g, 'log');
        s = s.replace(/\^([^{])/g, function(m, p) { return '/*^*/ pow(BASE,' + p + ')'; });
        s = s.replace(/\^{([^}]*)}/g, function(m, p) { return '/*^*/ pow(BASE,' + p + ')'; });
        s = s.replace(/×/g, '*').replace(/÷/g, '/');
        return s || '0';
    }

    // =========================================================
    // 14. DIGITAL LOGIC ANALYZER
    // =========================================================
    function setupDigitalLogic() {
        var calcBtn = document.getElementById('logic-calc');
        if (!calcBtn) return;
        calcBtn.addEventListener('click', function() {
            var exprRaw = (document.getElementById('logic-expr').value || '').trim();
            var resultEl = document.getElementById('logic-result');
            if (!exprRaw) { resultEl.innerHTML = '<p style="color:#ff6b6b;">Lütfen bir lojik ifade girin.</p>'; return; }

            // Extract unique variables
            var varSet = {};
            exprRaw.replace(/\b([A-Z])\b/g, function(m, v) { varSet[v] = true; });
            var vars = Object.keys(varSet).sort();
            if (vars.length === 0 || vars.length > 6) {
                resultEl.innerHTML = '<p style="color:#ff6b6b;">1-6 arası büyük harf değişken kullanın (A-F).</p>';
                return;
            }

            var n = vars.length;
            var rows = [];
            var totalRows = Math.pow(2, n);

            for (var i = 0; i < totalRows; i++) {
                var assignment = {};
                for (var j = 0; j < n; j++) {
                    assignment[vars[j]] = (i >> (n - 1 - j)) & 1;
                }
                var output = evaluateLogic(exprRaw, assignment);
                rows.push({ assignment: assignment, output: output });
            }

            // Build HTML table
            var html = '<table class="truth-table"><thead><tr>';
            vars.forEach(function(v) { html += '<th>' + v + '</th>'; });
            html += '<th>ÇIKIŞ</th></tr></thead><tbody>';
            rows.forEach(function(row) {
                html += '<tr>';
                vars.forEach(function(v) {
                    var val = row.assignment[v];
                    html += '<td class="' + (val ? 'tt-true' : 'tt-false') + '">' + val + '</td>';
                });
                var out = row.output;
                html += '<td class="' + (out ? 'tt-true' : 'tt-false') + '"><strong>' + out + '</strong></td>';
                html += '</tr>';
            });
            html += '</tbody></table>';
            resultEl.innerHTML = html;
        });
    }

    function evaluateLogic(expr, vars) {
        var s = expr;
        // Replace NAND, NOR, XNOR first (before AND/OR/NOT)
        s = s.replace(/\bXNOR\b/g, '___XNOR___');
        s = s.replace(/\bNAND\b/g, '___NAND___');
        s = s.replace(/\bNOR\b/g, '___NOR___');
        s = s.replace(/\bXOR\b/g, '___XOR___');
        s = s.replace(/\bNOT\s+/g, '!');
        s = s.replace(/\bAND\b/g, '&&');
        s = s.replace(/\bOR\b/g, '||');

        // Substitute variable values
        Object.keys(vars).forEach(function(v) {
            s = s.replace(new RegExp('\\b' + v + '\\b', 'g'), vars[v]);
        });

        // Handle compound operators
        s = s.replace(/(\d|\))\s*___NAND___\s*(\d|\()/g, function(m, a, b) { return '!(' + a + '&&' + b + ')'; });
        s = s.replace(/(\d|\))\s*___NOR___\s*(\d|\()/g, function(m, a, b) { return '!(' + a + '||' + b + ')'; });
        s = s.replace(/(\d|\))\s*___XOR___\s*(\d|\()/g, function(m, a, b) { return '(' + a + '^' + b + ')'; });
        s = s.replace(/(\d|\))\s*___XNOR___\s*(\d|\()/g, function(m, a, b) { return '!(' + a + '^' + b + ')'; });

        try {
            // eslint-disable-next-line no-eval
            var r = eval(s);
            return r ? 1 : 0;
        } catch(e) { return '?'; }
    }

    // =========================================================
    // 15. BOM CALCULATOR (with Local Storage)
    // =========================================================
    function setupBOM() {
        var addBtn = document.getElementById('bom-add-row');
        var rowsContainer = document.getElementById('bom-rows');
        var saveBtn = document.getElementById('bom-save-btn');
        if (!addBtn || !rowsContainer) return;

        // Start with 3 default rows if empty
        if (rowsContainer.children.length === 0) {
            addBOMRow(); addBOMRow(); addBOMRow();
        }

        addBtn.addEventListener('click', function() { addBOMRow(); });
        if (saveBtn) saveBtn.addEventListener('click', saveBOMProject);

        loadBOMProjectList();
    }

    function addBOMRow(name, qty, price) {
        var rowsContainer = document.getElementById('bom-rows');
        if (!rowsContainer) return;
        var div = document.createElement('div');
        div.className = 'bom-row bom-row--large';
        div.innerHTML =
            '<input type="text" class="bom-name" placeholder="Malzeme adı..." value="' + (name || '') + '" />' +
            '<input type="number" class="bom-qty" placeholder="Adet" min="0" value="' + (qty !== undefined ? qty : 1) + '" />' +
            '<input type="number" class="bom-price" placeholder="Fiyat" min="0" step="0.01" value="' + (price !== undefined ? price : 0) + '" />' +
            '<span class="bom-row-total">0.00 TL</span>' +
            '<button class="bom-row-del" title="Sil">✕</button>';

        var qEl = div.querySelector('.bom-qty');
        var pEl = div.querySelector('.bom-price');
        var tEl = div.querySelector('.bom-row-total');
        var delBtn = div.querySelector('.bom-row-del');

        function updateRow() {
            var q = parseFloat(qEl.value) || 0;
            var p = parseFloat(pEl.value) || 0;
            tEl.textContent = (q * p).toFixed(2) + ' TL';
            updateBOMTotal();
        }

        qEl.addEventListener('input', updateRow);
        pEl.addEventListener('input', updateRow);
        delBtn.addEventListener('click', function() {
            div.remove();
            updateBOMTotal();
        });

        rowsContainer.appendChild(div);
        updateRow();
    }

    function updateBOMTotal() {
        var rows = document.querySelectorAll('#bom-rows .bom-row');
        var sum = 0;
        rows.forEach(function(row) {
            var q = parseFloat(row.querySelector('.bom-qty').value) || 0;
            var p = parseFloat(row.querySelector('.bom-price').value) || 0;
            sum += q * p;
        });
        var totalEl = document.getElementById('bom-total');
        if (totalEl) {
            totalEl.innerHTML = 'Toplam Proje Maliyeti: <strong>' + sum.toFixed(2) + ' TL</strong>';
        }
    }

    function getBOMProjects() {
        try {
            return JSON.parse(localStorage.getItem('bom_projects') || '{}');
        } catch(e) { return {}; }
    }

    function saveBOMProject() {
        var nameInput = document.getElementById('bom-project-name');
        var name = (nameInput.value || '').trim();
        if (!name) { alert('Lütfen proje adı girin.'); return; }

        var items = [];
        document.querySelectorAll('#bom-rows .bom-row').forEach(function(row) {
            var n = row.querySelector('.bom-name').value;
            var q = row.querySelector('.bom-qty').value;
            var p = row.querySelector('.bom-price').value;
            if (n || q || p) items.push({ name: n, qty: q, price: p });
        });

        var projects = getBOMProjects();
        projects[name] = items;
        localStorage.setItem('bom_projects', JSON.stringify(projects));
        loadBOMProjectList();
        alert('Proje kaydedildi!');
    }

    function loadBOMProjectList() {
        var listEl = document.getElementById('bom-project-list');
        if (!listEl) return;
        listEl.innerHTML = '';
        var projects = getBOMProjects();
        
        Object.keys(projects).forEach(function(pName) {
            var li = document.createElement('li');
            li.className = 'bom-project-item';
            li.innerHTML = '<span>&#128194; ' + pName + '</span> <button class="btn-del-project" title="Sil">🗑️</button>';
            
            li.querySelector('span').addEventListener('click', function() {
                loadBOMProject(pName);
            });
            
            li.querySelector('.btn-del-project').addEventListener('click', function(e) {
                e.stopPropagation();
                if (confirm(pName + ' projesini silmek istediğinize emin misiniz?')) {
                    var prj = getBOMProjects();
                    delete prj[pName];
                    localStorage.setItem('bom_projects', JSON.stringify(prj));
                    loadBOMProjectList();
                }
            });
            
            listEl.appendChild(li);
        });
    }

    function loadBOMProject(name) {
        var projects = getBOMProjects();
        var items = projects[name];
        if (!items) return;
        
        var rowsContainer = document.getElementById('bom-rows');
        rowsContainer.innerHTML = '';
        document.getElementById('bom-project-name').value = name;
        
        if (items.length === 0) {
            addBOMRow(); addBOMRow(); addBOMRow();
        } else {
            items.forEach(function(item) {
                addBOMRow(item.name, item.qty, item.price);
            });
        }
    }

    // =========================================================
    // 16. NEW TOOLKIT TOOLS
    // =========================================================
    function setupPCB() {
        var calcBtn = document.getElementById('pcb-calc');
        if (!calcBtn) return;
        calcBtn.addEventListener('click', function() {
            var i = parseFloat(document.getElementById('pcb-current').value) || 0;
            var oz = parseFloat(document.getElementById('pcb-copper').value) || 0;
            var res = document.getElementById('pcb-result');
            
            if (i <= 0 || oz <= 0) {
                res.textContent = 'Lütfen geçerli değerler girin.';
                return;
            }
            
            // IPC-2221 External Trace formula (Temp Rise = 10C)
            // Area = ( I / (0.048 * (10^0.44)) ) ^ (1/0.725)
            // Width = Area / (Thickness * 1.378)  [in mils]
            var area = Math.pow(i / (0.048 * Math.pow(10, 0.44)), 1 / 0.725);
            var widthMils = area / (oz * 1.378);
            var widthMm = widthMils * 0.0254;
            
            res.innerHTML = widthMm.toFixed(3) + ' mm <br><small style="font-size:12px;color:gray;">(' + widthMils.toFixed(1) + ' mil)</small>';
        });
    }

    function setup3DPrint() {
        var calcBtn = document.getElementById('print-calc');
        if (!calcBtn) return;
        calcBtn.addEventListener('click', function() {
            var weight = parseFloat(document.getElementById('print-weight').value) || 0;
            var price = parseFloat(document.getElementById('print-price').value) || 0;
            var res = document.getElementById('print-result');
            
            var cost = weight * (price / 1000);
            res.textContent = cost.toFixed(2) + ' TL';
        });
    }

    document.addEventListener('DOMContentLoaded', init);
})();

