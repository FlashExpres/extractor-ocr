document.addEventListener('DOMContentLoaded', () => {
    // --- Google Drive Configuration ---
    let CLIENT_ID = localStorage.getItem('google_client_id') || ""; 
    let API_KEY = localStorage.getItem('google_api_key') || "";
    const DISCOVERY_DOCS = ["https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"];
    const SCOPES = 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.readonly';

    let tokenClient;
    let gapiInited = false;
    let gisInited = false;
    let selectedDriveFile = null;

    // --- State & Constants ---
    let currentFiles = [];
    let dbCounter = 1;

    const tucumanMapping = {
        "YERBA BUENA": "4107",
        "SAN MIGUEL DE TUCUMAN": "4000",
        "SAN MIGUEL": "4000",
        "LA BANDA DEL RIO SALI": "4109",
        "CEVIL REDONDO": "4107",
        "TAFI VIEJO": "4103",
        "LAS TALITAS": "4101"
    };

    const getCP = (loc) => tucumanMapping[loc.toUpperCase()] || "";
    const getLoc = (cp) => Object.keys(tucumanMapping).find(k => tucumanMapping[k] === cp) || "";
    const cleanPhone = (tel) => {
        let t = tel.replace(/[^\d]/g, '');
        if (t.startsWith('5')) t = t.substring(1);
        return t;
    };
    
    // Default Drivers & Clients (initial state)
    let drivers = JSON.parse(localStorage.getItem('ers_drivers')) || ["MATIAS BRAVO", "ANGEL DIAZ"];
    let clients = JSON.parse(localStorage.getItem('ers_clients')) || [
        { name: "DIAZ MATIAS JAVIER", dom: "BOULEVARD 9 DE JULIO Y ZAVALIA 386", loc: "YERBA BUENA", cp: "4107" },
        { name: "KURANDA SAS", dom: "PERU 1706", loc: "YERBA BUENA", cp: "4107" },
        { name: "TUS TECNOLOGIAS", dom: "LAVALLE 470", loc: "SAN MIGUEL DE TUCUMAN", cp: "4000" },
        { name: "CUANDO TE ENAMORAS", dom: "VIRGEN DE LA MERCED 395", loc: "SAN MIGUEL DE TUCUMAN", cp: "4000" }
    ];

    // --- DOM Elements ---
    const themeToggle = document.getElementById('theme-toggle');
    const dateInput = document.getElementById('date-input');
    const driverSelect = document.getElementById('driver-select');
    const clientSelect = document.getElementById('client-select');
    const apiKeyInput = document.getElementById('api-key-input');
    const rawTextInput = document.getElementById('raw-text-input');
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadPreview = document.getElementById('upload-preview');
    const tableBody = document.getElementById('table-body');
    const processBtn = document.getElementById('process-btn');
    const exportBtn = document.getElementById('export-btn');
    
    // Modals
    const modals = {
        settings: document.getElementById('settings-modal'),
        driver: document.getElementById('driver-modal'),
        client: document.getElementById('client-modal')
    };

    // --- Initial Setup ---
    const init = () => {
        // Theme
        const prefersDark = window.matchMedia("(prefers-color-scheme: dark)");
        setTheme(prefersDark.matches ? 'dark' : 'light');

        // Date (Yesterday)
        const d = new Date();
        d.setDate(d.getDate() - 1);
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        dateInput.value = `${y}-${m}-${day}`;

        // API Keys
        apiKeyInput.value = localStorage.getItem('gemini_api_key') || "";
        document.getElementById('google-client-id').value = localStorage.getItem('google_client_id') || "";
        document.getElementById('google-api-key').value = localStorage.getItem('google_api_key') || "";

        // Populate Selects
        renderDriverOptions();
        renderClientOptions();
        
        // Google Drive
        initDrive();
    };

    const setTheme = (theme) => {
        document.body.setAttribute('data-theme', theme);
        const icon = themeToggle.querySelector('i');
        icon.className = theme === 'dark' ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
    };

    const renderDriverOptions = () => {
        const val = driverSelect.value;
        driverSelect.innerHTML = '<option value="">Seleccione un Driver...</option>';
        drivers.forEach(d => {
            const opt = document.createElement('option');
            opt.value = d;
            opt.textContent = d;
            driverSelect.appendChild(opt);
        });
        driverSelect.value = val;
    };

    const renderClientOptions = () => {
        const val = clientSelect.value;
        clientSelect.innerHTML = '<option value="">Seleccione un Cliente...</option>';
        clients.forEach(c => {
            const opt = document.createElement('option');
            opt.value = c.name;
            opt.textContent = c.name;
            clientSelect.appendChild(opt);
        });
        clientSelect.value = val;
    };

    // --- Modal Logic ---
    const showModal = (id) => {
        if (modals[id]) modals[id].classList.remove('hidden');
    };
    const hideModals = () => {
        Object.values(modals).forEach(m => {
            if (m) m.classList.add('hidden');
        });
    };

    document.getElementById('open-settings-btn').addEventListener('click', () => showModal('settings'));
    document.getElementById('add-driver-btn').addEventListener('click', () => showModal('driver'));
    document.getElementById('add-client-btn').addEventListener('click', () => showModal('client'));

    document.querySelectorAll('.close-modal, .close-modal-btn').forEach(btn => {
        btn.addEventListener('click', hideModals);
    });

    window.addEventListener('click', (e) => {
        if (Object.values(modals).includes(e.target)) hideModals();
    });

    // Save Settings
    document.getElementById('save-settings-btn').onclick = () => {
        localStorage.setItem('gemini_api_key', apiKeyInput.value.trim());
        localStorage.setItem('google_client_id', document.getElementById('google-client-id').value.trim());
        localStorage.setItem('google_api_key', document.getElementById('google-api-key').value.trim());
        
        // Update local variables and re-init Drive
        CLIENT_ID = localStorage.getItem('google_client_id');
        API_KEY = localStorage.getItem('google_api_key');
        if (CLIENT_ID && API_KEY) initDrive();

        hideModals();
        alert("Configuración guardada. Recarga la página si el botón de Drive aún no funciona.");
    };

    // Delete Driver
    document.getElementById('delete-driver-btn').onclick = () => {
        const val = driverSelect.value;
        if (!val) return alert("Selecciona un driver para eliminar.");
        if (confirm(`¿Estás seguro de eliminar a ${val}?`)) {
            drivers = drivers.filter(d => d !== val);
            localStorage.setItem('ers_drivers', JSON.stringify(drivers));
            renderDriverOptions();
            driverSelect.value = "";
        }
    };

    // Delete Client
    document.getElementById('delete-client-btn').onclick = () => {
        const val = clientSelect.value;
        if (!val) return alert("Selecciona un cliente para eliminar.");
        if (confirm(`¿Estás seguro de eliminar a ${val}?`)) {
            clients = clients.filter(c => c.name !== val);
            localStorage.setItem('ers_clients', JSON.stringify(clients));
            renderClientOptions();
            clientSelect.value = "";
        }
    };

    // Save New Driver
    document.getElementById('save-driver-btn').onclick = () => {
        const name = document.getElementById('new-driver-name').value.trim().toUpperCase();
        if (name && !drivers.includes(name)) {
            drivers.push(name);
            localStorage.setItem('ers_drivers', JSON.stringify(drivers));
            renderDriverOptions();
            driverSelect.value = name;
            document.getElementById('new-driver-name').value = "";
            hideModals();
        }
    };

    // --- Client Modal Auto-fill ---
    const newClientLoc = document.getElementById('new-client-localidad');
    const newClientCP = document.getElementById('new-client-cp');

    newClientLoc.addEventListener('blur', () => {
        const cp = getCP(newClientLoc.value.trim());
        if (cp && !newClientCP.value) newClientCP.value = cp;
        newClientLoc.value = newClientLoc.value.toUpperCase();
    });

    newClientCP.addEventListener('blur', () => {
        const loc = getLoc(newClientCP.value.trim());
        if (loc && !newClientLoc.value) newClientLoc.value = loc;
    });

    // Save New Client
    document.getElementById('save-client-btn').onclick = () => {
        const name = document.getElementById('new-client-name').value.trim().toUpperCase();
        const dom = document.getElementById('new-client-domicilio').value.trim().toUpperCase();
        const loc = document.getElementById('new-client-localidad').value.trim().toUpperCase() || getLoc(newClientCP.value.trim());
        const cp = document.getElementById('new-client-cp').value.trim() || getCP(loc);
        
        if (name && dom && loc) {
            clients.push({ name, dom, loc, cp });
            localStorage.setItem('ers_clients', JSON.stringify(clients));
            renderClientOptions();
            clientSelect.value = name;
            // Clear inputs
            ['new-client-name', 'new-client-domicilio', 'new-client-localidad', 'new-client-cp'].forEach(id => {
                document.getElementById(id).value = "";
            });
            hideModals();
        } else {
            alert("Completa nombre, domicilio y localidad.");
        }
    };

    const parseExcelFromDrive = async (arrayBuffer, rowNums, dateStr, driver, clientName, clientData) => {
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        rowNums.forEach(num => {
            const row = data[num - 1]; // 1-indexed to 0-indexed
            if (!row || row.length < 2) return;

            addTableRow({
                id: String(row[1] || "").toUpperCase(), // B is index 1
                fecha: dateStr, driver, cliente: clientName,
                domOrigen: clientData.dom, locOrigen: clientData.loc, cpOrigen: clientData.cp,
                destinatario: String(row[5] || "").toUpperCase(), // F is index 5
                domDestino: String(row[6] || "").toUpperCase(), // G
                locDestino: String(row[7] || "").toUpperCase(), // H
                cpDestino: String(row[8] || getCP(String(row[7] || ""))), // I
                telefono: cleanPhone(String(row[9] || "")), // J
                valorDeclarado: String(row[10] || "0") // K
            });
        });
    };

    // --- Google Drive Integration logic ---
    const initDrive = () => {
        if (!CLIENT_ID || !API_KEY) return; // Don't init without keys

        const script = document.createElement('script');
        script.src = "https://apis.google.com/js/api.js";
        script.onload = () => {
            gapi.load('client:picker', async () => {
                await gapi.client.init({
                    apiKey: API_KEY,
                    discoveryDocs: DISCOVERY_DOCS,
                });
                gapiInited = true;
            });
        };
        document.body.appendChild(script);

        const gisScript = document.createElement('script');
        gisScript.src = "https://accounts.google.com/gsi/client";
        gisScript.onload = () => {
            tokenClient = google.accounts.oauth2.initTokenClient({
                client_id: CLIENT_ID,
                scope: SCOPES,
                callback: '', // defined later
            });
            gisInited = true;
        };
        document.body.appendChild(gisScript);
    };

    document.getElementById('drive-btn').onclick = () => {
        if (!gapiInited || !gisInited) return alert("Google API no cargada.");
        
        tokenClient.callback = async (response) => {
            if (response.error !== undefined) throw (response);
            createPicker(response.access_token);
        };

        if (gapi.client.getToken() === null) {
            tokenClient.requestAccessToken({prompt: 'consent'});
        } else {
            tokenClient.requestAccessToken({prompt: ''});
        }
    };

    const createPicker = (accessToken) => {
        const view = new google.picker.View(google.picker.ViewId.DOCS);
        view.setMimeTypes("application/vnd.google-apps.spreadsheet,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        const picker = new google.picker.PickerBuilder()
            .addView(view)
            .setOAuthToken(accessToken)
            .setDeveloperKey(API_KEY)
            .setCallback(pickerCallback)
            .build();
        picker.setVisible(true);
    };

    const pickerCallback = (data) => {
        if (data.action == google.picker.Action.PICKED) {
            const doc = data.docs[0];
            selectedDriveFile = { id: doc.id, name: doc.name, mimeType: doc.mimeType };
            alert(`Archivo seleccionado: ${doc.name}. Ahora indica las filas en el cuadro de texto (ej: 6, 7).`);
            if (!rawTextInput.value) rawTextInput.value = "FILAS: ";
            rawTextInput.focus();
        }
    };

    // --- File Handling (Images/Excel) ---
    const handleFiles = (files) => {
        currentFiles = [...currentFiles, ...Array.from(files)];
        updatePreview();
    };

    fileInput.onchange = (e) => handleFiles(e.target.files);

    const updatePreview = () => {
        const dropZoneContent = document.querySelector('.drop-zone-content');
        if (currentFiles.length) {
            uploadPreview.style.display = 'grid';
            dropZoneContent.style.display = 'none';
        } else {
            uploadPreview.style.display = 'none';
            dropZoneContent.style.display = 'block';
        }
        uploadPreview.innerHTML = '';
        currentFiles.forEach((file, i) => {
            const item = document.createElement('div');
            item.className = 'preview-item';
            const isImg = file.type.startsWith('image/');
            item.innerHTML = `
                ${isImg ? `<img src="${URL.createObjectURL(file)}">` : `<i class="fa-solid fa-file-excel" style="font-size:2rem;color:var(--success)"></i>`}
                <span>${file.name}</span>
                <button class="remove-file" onclick="event.stopPropagation(); window.removeFile(${i})"><i class="fa-solid fa-xmark"></i></button>
            `;
            uploadPreview.appendChild(item);
        });
    };

    window.removeFile = (i) => {
        currentFiles.splice(i, 1);
        updatePreview();
    };

    // Dropzone events
    const prevent = (e) => { e.preventDefault(); e.stopPropagation(); };
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => dropZone.addEventListener(evt, prevent));
    dropZone.ondragenter = () => dropZone.classList.add('active');
    dropZone.ondragleave = () => dropZone.classList.remove('active');
    dropZone.ondrop = (e) => {
        dropZone.classList.remove('active');
        handleFiles(e.dataTransfer.files);
    };

    // --- Extraction Logic ---
    async function extractWithGemini(apiKey, base64Image, mimeType) {
        const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
        const prompt = `Eres un escáner de extracción de datos para logística. Devuelve JSON: {"id":"PEDIDO","destinatario":"NOMBRE","domicilio_destino":"CALLE 123","localidad_destino":"CIUDAD","telefono":"NUMERO"}. Si no ves un dato, deja el campo VACIO "". Todo en MAYUSCULAS.`;
        
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType, data: base64Image } }] }],
                generationConfig: { temperature: 0.2, response_mime_type: "application/json" }
            })
        });

        if (!response.ok) throw new Error("Error en API Gemini");
        const data = await response.json();
        const rawText = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
        
        try {
            return JSON.parse(rawText);
        } catch (e) {
            // Fallback regex
            const extract = (k) => (rawText.match(new RegExp(`"${k}"\\s*:\\s*"([\\s\\S]*?)"`, 'i')) || [])[1] || '';
            return { id: extract('id'), destinatario: extract('destinatario'), domicilio_destino: extract('domicilio_destino'), localidad_destino: extract('localidad_destino'), telefono: extract('telefono') };
        }
    }

    // --- Processing ---
    processBtn.onclick = async () => {
        const apiKey = apiKeyInput.value.trim();
        const driver = driverSelect.value;
        const clientName = clientSelect.value;
        const rawText = rawTextInput.value.trim();

        if (!driver || !clientName) return alert("Selecciona Driver y Cliente");
        if (currentFiles.length === 0 && !rawText) return alert("Sube imágenes o pega texto");
        if (currentFiles.length > 0 && !apiKey) return alert("Falta API Key para las imágenes");

        processBtn.disabled = true;
        processBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Procesando...';

        const clientData = clients.find(c => c.name === clientName);
        const dateStr = dateInput.value.split('-').reverse().join('/');

        // Advanced Text Override Logic
        // Support common formats and the specific delivery format
        const cleanID = (id) => id.replace(/^#[A-Z0-9]/i, '').replace(/^#/i, '').trim().toUpperCase();

        const textBlocks = rawText.split(/(?=\nNúmero de pedido:|\nID:|\nPEDIDO:)/i)
            .filter(b => b.trim().length > 0)
            .map(b => b.trim());

        const parsedOverrides = textBlocks.map(block => {
            const overrides = {};
            const keywords = {
                id: [/id[:\s]+([A-Z0-9#]+)/i, /pedido[:\s]+([A-Z0-9#]+)/i, /Número de pedido: ([A-Z0-9#]+)/i],
                destinatario: [/destinatario[:\s]+([^,\n(\[]+)/i, /nombre[:\s]+([^,\n(\[]+)/i, /Datos del cliente: ([^,\n(\[]+)/i],
                domDestino: [/domicilio destino[:\s]+([^,\n]+)/i, /domicilio[:\s]+([^,\n]+)/i, /direccion[:\s]+([^,\n]+)/i, /Dirección de entrega: ([^,\n]+)/i],
                locDestino: [/localidad destino[:\s]+([^,\n\r]+)/i, /localidad[:\s]+([^,\n\r]+)/i, /ciudad[:\s]+([^,\n\r]+)/i],
                cpDestino: [/cp destino[:\s]+(\d+)/i, /cp[:\s]+(\d+)/i, /codigo postal[:\s]+(\d+)/i],
                telefono: [/telefono[:\s]+([\d\-\s\(\)]+)/i, /tel[:\s]+([\d\-\s\(\)]+)/i, /\((\d+)\)/],
                valorDeclarado: [/valor[:\s]+(\d+)/i, /declarado[:\s]+(\d+)/i]
            };

            const usedText = [];
            for (const [key, regexes] of Object.entries(keywords)) {
                for (const reg of regexes) {
                    const match = block.match(reg);
                    if (match) {
                        let val = match[1].trim();
                        if (key === 'id') val = cleanID(val);
                        if (key === 'telefono') val = cleanPhone(val);
                        
                        overrides[key] = val.toUpperCase();
                        usedText.push(match[0]);
                        
                        // Smart Address Splitting: If it's a Dirección de entrega line with a comma
                        if (key === 'domDestino' && match[0].includes('Dirección de entrega:')) {
                            const fullAddr = match[1].split(',');
                            if (fullAddr.length > 1) {
                                overrides.domDestino = fullAddr[0].trim().toUpperCase();
                                overrides.locDestino = fullAddr[1].trim().toUpperCase();
                            }
                        }

                        // Auto-fill CP/Locality if one is provided
                        if (key === 'locDestino' && !overrides.cpDestino) {
                            overrides.cpDestino = getCP(overrides.locDestino);
                        }
                        if (key === 'cpDestino' && !overrides.locDestino) {
                            overrides.locDestino = getLoc(overrides.cpDestino);
                        }
                        break;
                    }
                }
            }
            
            // If no ID found, but there's an ID-like word NOT used in other keywords
            if (!overrides.id) {
                let remaining = block;
                usedText.forEach(t => remaining = remaining.replace(t, ''));
                const idMatch = remaining.match(/\b([A-Z0-9]{3,})\b/i);
                if (idMatch) overrides.id = cleanID(idMatch[1].toUpperCase());
            }
            
            return overrides;
        });

        let overrideIndex = 0;

        // 0. Process Drive File if selected
        if (selectedDriveFile) {
            try {
                const accessToken = gapi.client.getToken().access_token;
                let url = `https://www.googleapis.com/drive/v3/files/${selectedDriveFile.id}?alt=media`;
                
                // If it's a Google Sheet, we MUST export it
                if (selectedDriveFile.mimeType === 'application/vnd.google-apps.spreadsheet') {
                    url = `https://www.googleapis.com/drive/v3/files/${selectedDriveFile.id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`;
                }

                const response = await fetch(url, {
                    headers: { Authorization: `Bearer ${accessToken}` }
                });

                if (!response.ok) throw new Error("Error al descargar de Drive");
                const buffer = await response.arrayBuffer();
                
                const rowNums = rawText.match(/\d+/g)?.map(Number) || [];
                if (rowNums.length === 0) {
                    alert("Por favor indica los números de fila en el texto.");
                } else {
                    await parseExcelFromDrive(buffer, rowNums, dateStr, driver, clientName, clientData);
                }
                
                selectedDriveFile = null; // Clear after processing
                rawTextInput.value = "";
            } catch (e) {
                alert("Error procesando Drive: " + e.message);
            }
        }

        // 1. Process Manual Entry (if text blocks exist and NO images AND NO Drive active)
        const hasImages = currentFiles.some(f => f.type.startsWith('image/'));
        
        if (rawText && !hasImages && !selectedDriveFile) {
            parsedOverrides.forEach(ov => {
                addTableRow({
                    id: ov.id || "TEXTO",
                    fecha: dateStr, driver, cliente: clientName,
                    domOrigen: clientData.dom, locOrigen: clientData.loc, cpOrigen: clientData.cp,
                    destinatario: ov.destinatario || "MANUAL", 
                    domDestino: ov.domDestino || "VER TEXTO", 
                    locDestino: ov.locDestino || "TUCUMAN", 
                    cpDestino: ov.cpDestino || getCP(ov.locDestino || "TUCUMAN"),
                    telefono: ov.telefono || "1111", 
                    valorDeclarado: ov.valorDeclarado || "0"
                });
            });
            rawTextInput.value = "";
        }

        // 2. Process Files with Text Overrides (only images use Gemini)
        for (const file of currentFiles) {
            if (file.type.startsWith('image/')) {
                try {
                    const b64 = await new Promise(r => {
                        const rd = new FileReader();
                        rd.onload = () => r(rd.result.split(',')[1]);
                        rd.readAsDataURL(file);
                    });
                    const res = await extractWithGemini(apiKey, b64, file.type);
                    
                    // Override OCR results with text data if available
                    const ov = parsedOverrides[overrideIndex] || {};
                    overrideIndex++;

                    addTableRow({
                        id: ov.id || res.id.toUpperCase(),
                        fecha: dateStr, driver, cliente: clientName,
                        domOrigen: clientData.dom, locOrigen: clientData.loc, cpOrigen: clientData.cp,
                        destinatario: ov.destinatario || res.destinatario.toUpperCase(),
                        domDestino: ov.domDestino || res.domicilio_destino.toUpperCase(),
                        locDestino: ov.locDestino || res.localidad_destino.toUpperCase(),
                        cpDestino: ov.cpDestino || getCP(ov.locDestino || res.localidad_destino), 
                        telefono: ov.telefono || cleanPhone(res.telefono) || "1111",
                        valorDeclarado: ov.valorDeclarado || "0"
                    });
                } catch (e) {
                    addTableRow({ id: "ERROR", fecha: dateStr, driver, cliente: clientName, domOrigen: clientData.dom, locOrigen: clientData.loc, cpOrigen: clientData.cp, destinatario: "ERROR API", domDestino: "", locDestino: "", cpDestino: "", telefono: "1111", valorDeclarado: "0" });
                }
            }
        }

        rawTextInput.value = ""; // Clear IDs after use
        currentFiles = [];
        updatePreview();
        processBtn.disabled = false;
        processBtn.innerHTML = '<i class="fa-solid fa-file-import"></i> Cargar Datos';
    };

    // --- Table & Rows ---
    const addTableRow = (data) => {
        document.getElementById('empty-state').style.display = 'none';
        const tr = document.createElement('tr');
        tr.draggable = true;

        // Force uppercase on manual edits
        tr.addEventListener('input', (e) => {
            if (e.target.hasAttribute('contenteditable')) {
                const sel = window.getSelection();
                if (!sel.rangeCount) return;
                const range = sel.getRangeAt(0);
                const offset = range.startOffset;
                
                e.target.innerText = e.target.innerText.toUpperCase();
                
                // Restore cursor position
                try {
                    const newRange = document.createRange();
                    newRange.setStart(e.target.childNodes[0], offset);
                    newRange.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(newRange);
                } catch(err) {}
            }
        });

        // Double click to select all text
        tr.addEventListener('dblclick', (e) => {
            if (e.target.hasAttribute('contenteditable')) {
                const range = document.createRange();
                range.selectNodeContents(e.target);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
            }
        });

        tr.innerHTML = `
            <td class="row-num">${dbCounter++}</td>
            <td contenteditable="true">${data.id}</td>
            <td contenteditable="true">${data.fecha}</td>
            <td class="driver-cell" title="Doble clic para cambiar">${data.driver}</td>
            <td class="client-cell" title="Doble clic para cambiar">${data.cliente}</td>
            <td contenteditable="true">${data.domOrigen}</td>
            <td contenteditable="true">${data.locOrigen}</td>
            <td contenteditable="true">${data.cpOrigen}</td>
            <td contenteditable="true">${data.destinatario}</td>
            <td contenteditable="true">${data.domDestino}</td>
            <td contenteditable="true">${data.locDestino}</td>
            <td contenteditable="true">${data.cpDestino}</td>
            <td contenteditable="true">${data.telefono}</td>
            <td contenteditable="true">${data.valorDeclarado}</td>
            <td>
                <div class="action-btns">
                    <button class="btn-delete" onclick="this.parentElement.parentElement.parentElement.remove(); window.renumberRows();"><i class="fa-solid fa-trash"></i></button>
                    <button class="btn-drag" style="background:var(--text-secondary);"><i class="fa-solid fa-grip-vertical"></i></button>
                </div>
            </td>
        `;
        tableBody.appendChild(tr);
        setupRowInteractions(tr);
    };

    window.renumberRows = () => {
        const rows = tableBody.querySelectorAll('tr:not(.empty-row)');
        dbCounter = 1;
        rows.forEach(r => {
            r.querySelector('.row-num').textContent = dbCounter++;
        });
        if (rows.length === 0) document.getElementById('empty-state').style.display = 'table-row';
    };

    const setupRowInteractions = (row) => {
        // Drag and Drop
        row.ondragstart = (e) => {
            row.classList.add('dragging');
            e.dataTransfer.setData('text/plain', ''); // Required for Firefox
        };
        row.ondragend = () => {
            row.classList.remove('dragging');
            window.renumberRows();
        };
        row.ondragover = (e) => {
            e.preventDefault();
            const dragging = tableBody.querySelector('.dragging');
            if (dragging && dragging !== row) {
                const rect = row.getBoundingClientRect();
                const next = (e.clientY - rect.top) > (rect.height / 2);
                tableBody.insertBefore(dragging, next ? row.nextSibling : row);
            }
        };

        // Double click client to change sender
        const clientCell = row.querySelector('.client-cell');
        clientCell.ondblclick = () => {
            const select = document.createElement('select');
            select.className = 'clean-select';
            clients.forEach(c => {
                const opt = document.createElement('option');
                opt.value = c.name;
                opt.textContent = c.name;
                if (c.name === clientCell.textContent) opt.selected = true;
                select.appendChild(opt);
            });
            clientCell.appendChild(select);
            select.focus();
            select.onchange = () => {
                const newClient = clients.find(c => c.name === select.value);
                clientCell.textContent = newClient.name;
                const cells = row.querySelectorAll('td');
                cells[5].textContent = newClient.dom;
                cells[6].textContent = newClient.loc;
                cells[7].textContent = newClient.cp;
                select.remove();
            };
            select.onblur = () => { select.remove(); if (!clientCell.textContent) clientCell.textContent = data.cliente; };
        };

        // Double click driver to change
        const driverCell = row.querySelector('.driver-cell');
        driverCell.ondblclick = () => {
            const select = document.createElement('select');
            select.className = 'clean-select';
            drivers.forEach(d => {
                const opt = document.createElement('option');
                opt.value = d;
                opt.textContent = d;
                if (d === driverCell.textContent) opt.selected = true;
                select.appendChild(opt);
            });
            driverCell.appendChild(select);
            select.focus();
            select.onchange = () => {
                driverCell.textContent = select.value;
                select.remove();
            };
            select.onblur = () => { select.remove(); if (!driverCell.textContent) driverCell.textContent = data.driver; };
        };
    };

    // --- Export XLSX ---
    exportBtn.onclick = () => {
        const rows = tableBody.querySelectorAll('tr:not(.empty-row)');
        if (!rows.length) return alert("No hay datos");
        
        const wb = XLSX.utils.table_to_book(document.getElementById('data-table'), { sheet: "Logistica ERS" });
        XLSX.writeFile(wb, `ERS_Logistica_${new Date().toISOString().slice(0,10)}.xlsx`);
    };

    init();
});
