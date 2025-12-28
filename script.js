// Global variables to store loaded data
let clients = [];
let drivers = [];
let convoyeurs = [];
let products = [];
let history = [];
let localClients = [];

// Global variable to store the data directory handle for direct file saving
let dataDirectoryHandle = null;

// Load Excel data on page load
document.addEventListener('DOMContentLoaded', async function() {
    try {
        // Check if SheetJS is available
        if (typeof XLSX === 'undefined') {
            console.warn('⚠️ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
            alert('⚠️ جاري تحميل مكتبة Excel...\n\nيرجى الانتظار قليلاً ثم أعد تحميل الصفحة.');
        }
        
        // Check if localStorage is empty - if so, try to load from Excel files
        const hasClients = localStorage.getItem('clientsData');
        const hasDrivers = localStorage.getItem('driversData');
        const hasConvoyeurs = localStorage.getItem('convoyeursData');
        const hasProducts = localStorage.getItem('productsData');
        
        if (!hasClients && !hasDrivers && !hasConvoyeurs && !hasProducts) {
            console.log('ℹ️ localStorage فارغ. جاري محاولة تحميل البيانات من ملفات Excel في مجلد data/');
        }
        
        await loadClients();
        await loadDrivers();
        await loadConvoyeurs();
        await loadProducts();
        await loadHistory();
        setupEventListeners();
        updateStepDisplay();
        // Initialize document number after history is loaded
        setTimeout(() => {
            initializeDocumentNumber();
        }, 100);
    } catch (error) {
        console.error('Error during initialization:', error);
    }
});

// Helper function to reconstruct nested vehicle object from flattened Excel structure
function reconstructDriver(driver) {
    const reconstructed = { ...driver };
    if (driver['vehicle.matricule'] || driver['vehicle.model']) {
        reconstructed.vehicle = {
            matricule: driver['vehicle.matricule'] || '',
            model: driver['vehicle.model'] || ''
        };
        delete reconstructed['vehicle.matricule'];
        delete reconstructed['vehicle.model'];
    }
    return reconstructed;
}

// Helper function to reconstruct client data (convert itineraire string to array)
function reconstructClient(client) {
    const reconstructed = { ...client };
    // Convert itineraire from string to array if needed
    if (typeof reconstructed.itineraire === 'string' && reconstructed.itineraire.trim()) {
        reconstructed.itineraire = reconstructed.itineraire.split(',').map(s => s.trim()).filter(Boolean);
    } else if (!Array.isArray(reconstructed.itineraire)) {
        reconstructed.itineraire = [];
    }
    return reconstructed;
}

// Load Excel file using fs (Electron) or fetch (web server)
async function loadExcelFile(filename) {
    try {
        // Check if SheetJS is available
        if (typeof XLSX === 'undefined') {
            console.log('SheetJS not available yet');
            return null;
        }

        // Try Electron fs first
        if (typeof window !== 'undefined' && window.require) {
            try {
                const fs = window.require('fs');
                const path = window.require('path');
                
                let appPath = null;
                try {
                    const { app } = window.require('electron').remote || window.require('@electron/remote') || {};
                    if (app) {
                        appPath = app.getAppPath();
                    }
                } catch (e) {
                    try {
                        const process = window.require('process');
                        appPath = process.cwd();
                    } catch (e2) {
                        appPath = __dirname || '.';
                    }
                }
                
                if (appPath) {
                    const filePath = path.join(appPath, 'data', filename);
                    if (fs.existsSync(filePath)) {
                        const fileBuffer = fs.readFileSync(filePath);
                        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
                        const firstSheet = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheet];
                        return XLSX.utils.sheet_to_json(worksheet);
                    }
                }
            } catch (e) {
                console.log('Electron fs not available, trying fetch:', e);
            }
        }
        
        // Try fetch (works with web server)
        try {
        const response = await fetch(`data/${filename}`);
        if (response.ok) {
                const arrayBuffer = await response.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const firstSheet = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheet];
                return XLSX.utils.sheet_to_json(worksheet);
            }
        } catch (fetchError) {
            console.log(`⚠️ Cannot load ${filename} directly (file:// protocol). Use "Gestion Données" → Import Excel to load data.`);
            return null;
        }
        
        return null;
    } catch (error) {
        console.error(`Error loading ${filename}:`, error);
        return null;
    }
}

// JSON loading removed - Excel only

// Load clients from localStorage first (works offline), then try Excel file
async function loadClients() {
    try {
        // Priority 1: Load from localStorage (works offline, persists when PC is off)
        const saved = localStorage.getItem('clientsData');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                if (Array.isArray(parsed) && parsed.length > 0) {
                    clients = parsed;
                    console.log('✅ Clients loaded from localStorage:', clients.length);
                    populateClientSelect();
                    setTimeout(() => {
                        makeSelectSearchable('clientSelect');
                    }, 100);
                    return; // Use localStorage data
                }
            } catch (e) {
                console.error('Error parsing clients from localStorage:', e);
            }
        }
        
        // Priority 2: Try to load from Excel file (if available via server/Electron)
        const excelClients = await loadExcelFile('clients.xlsx');
        if (Array.isArray(excelClients) && excelClients.length > 0) {
            // Reconstruct client data (convert itineraire string to array)
            clients = excelClients.map(reconstructClient);
            // Save to localStorage for offline use
            localStorage.setItem('clientsData', JSON.stringify(clients));
            console.log('✅ Clients loaded from Excel file and saved to localStorage:', clients.length);
        } else {
            clients = [];
            console.warn('⚠️ No clients found. Use "Gestion Données" → Import Excel to load clients.xlsx or add clients manually.');
        }
        
        populateClientSelect();
        setTimeout(() => {
            makeSelectSearchable('clientSelect');
        }, 100);
    } catch (error) {
        console.error('❌ Error loading clients:', error);
        clients = [];
        populateClientSelect();
        setTimeout(() => {
            makeSelectSearchable('clientSelect');
        }, 100);
    }
}

function mergeLocalClients() {
    const saved = localStorage.getItem('clientsOverride');
    if (saved) {
        localClients = JSON.parse(saved);
    } else {
        localClients = [];
    }
    if (!Array.isArray(clients)) clients = [];
    const combined = [...clients, ...localClients];
    // remove duplicates by id
    const unique = combined.filter((item, index, self) =>
        index === self.findIndex(t => t.id === item.id)
    );
    clients = unique;
}

function getNextClientId() {
    let maxId = 0;
    clients.forEach(c => { if (c.id > maxId) maxId = c.id; });
    return maxId + 1;
}

// Load drivers from localStorage first, then try Excel file
async function loadDrivers() {
    try {
        // Priority 1: Load from localStorage
        const saved = localStorage.getItem('driversData');
        if (saved) {
            try {
                drivers = JSON.parse(saved);
                if (Array.isArray(drivers) && drivers.length > 0) {
                    console.log('✅ Drivers loaded from localStorage:', drivers.length);
                    populateDriverSelect();
                    setTimeout(() => {
                        makeSelectSearchable('driverSelect');
                    }, 100);
                    return;
                }
            } catch (e) {
                console.error('Error parsing drivers from localStorage:', e);
            }
        }
        
        // Priority 2: Try Excel file
        const excelDrivers = await loadExcelFile('drivers.xlsx');
        if (Array.isArray(excelDrivers) && excelDrivers.length > 0) {
            // Reconstruct nested vehicle objects from flattened Excel structure
            drivers = excelDrivers.map(reconstructDriver);
            localStorage.setItem('driversData', JSON.stringify(drivers));
            console.log('✅ Drivers loaded from Excel file and saved to localStorage:', drivers.length);
        } else {
            drivers = [];
            console.warn('⚠️ No drivers found, starting with empty list');
        }
        
        populateDriverSelect();
        setTimeout(() => {
            makeSelectSearchable('driverSelect');
        }, 100);
    } catch (error) {
        console.error('❌ Error loading drivers:', error);
        drivers = [];
        populateDriverSelect();
        setTimeout(() => {
            makeSelectSearchable('driverSelect');
        }, 100);
    }
}

// Load convoyeurs from localStorage first, then try Excel file
async function loadConvoyeurs() {
    try {
        // Priority 1: Load from localStorage
        const saved = localStorage.getItem('convoyeursData');
        if (saved) {
            try {
                convoyeurs = JSON.parse(saved);
                if (Array.isArray(convoyeurs) && convoyeurs.length > 0) {
                    console.log('✅ Convoyeurs loaded from localStorage:', convoyeurs.length);
                    populateConvoyeurSelect();
                    setTimeout(() => {
                        makeSelectSearchable('convoyeurSelect');
                    }, 100);
                    return;
                }
            } catch (e) {
                console.error('Error parsing convoyeurs from localStorage:', e);
            }
        }
        
        // Priority 2: Try Excel file
        const excelConvoyeurs = await loadExcelFile('convoyeurs.xlsx');
        if (Array.isArray(excelConvoyeurs) && excelConvoyeurs.length > 0) {
            convoyeurs = excelConvoyeurs;
            localStorage.setItem('convoyeursData', JSON.stringify(convoyeurs));
            console.log('✅ Convoyeurs loaded from Excel file and saved to localStorage:', convoyeurs.length);
        } else {
            convoyeurs = [];
            console.warn('⚠️ No convoyeurs found, starting with empty list');
        }
        
        populateConvoyeurSelect();
        setTimeout(() => {
            makeSelectSearchable('convoyeurSelect');
        }, 100);
    } catch (error) {
        console.error('❌ Error loading convoyeurs:', error);
        convoyeurs = [];
        populateConvoyeurSelect();
        setTimeout(() => {
            makeSelectSearchable('convoyeurSelect');
        }, 100);
    }
}

// Load products from localStorage first, then try Excel file
async function loadProducts() {
    try {
        // Priority 1: Load from localStorage
        const saved = localStorage.getItem('productsData');
        if (saved) {
            try {
                products = JSON.parse(saved);
                if (Array.isArray(products) && products.length > 0) {
                    console.log('✅ Products loaded from localStorage:', products.length);
                    populateProductSelects();
                    setTimeout(() => {
                        document.querySelectorAll('.product-select').forEach(select => {
                            makeSelectSearchable(select.id || null, select);
                        });
                    }, 100);
                    return;
                }
            } catch (e) {
                console.error('Error parsing products from localStorage:', e);
            }
        }
        
        // Priority 2: Try Excel file
        const excelProducts = await loadExcelFile('products.xlsx');
        if (Array.isArray(excelProducts) && excelProducts.length > 0) {
            products = excelProducts;
            localStorage.setItem('productsData', JSON.stringify(products));
            console.log('✅ Products loaded from Excel file and saved to localStorage:', products.length);
        } else {
            products = [];
            console.warn('⚠️ No products found, starting with empty list');
        }
        
        populateProductSelects();
        setTimeout(() => {
            document.querySelectorAll('.product-select').forEach(select => {
                makeSelectSearchable(select.id || null, select);
            });
        }, 100);
    } catch (error) {
        console.error('❌ Error loading products:', error);
        products = [];
        populateProductSelects();
        setTimeout(() => {
            document.querySelectorAll('.product-select').forEach(select => {
                makeSelectSearchable(select.id || null, select);
            });
        }, 100);
    }
}

// Load history from localStorage first (works offline), then try JSON file
async function loadHistory() {
    try {
        let localHistory = [];
        let jsonHistory = [];
        
        // Priority 1: Load from localStorage (works offline, persists when PC is off)
        const historyData = localStorage.getItem('declarationHistory');
        if (historyData) {
            try {
                localHistory = JSON.parse(historyData);
                if (Array.isArray(localHistory) && localHistory.length > 0) {
                    history = localHistory;
                    // Sort by timestamp (newest first)
                    history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
                    console.log('✅ History loaded from localStorage:', history.length, 'declarations');
                    return; // Use localStorage data
            }
        } catch (e) {
                console.error('Error parsing history from localStorage:', e);
            }
        }
        
        // Priority 2: Try to load from Excel file (if available via server/Electron)
        try {
            const loadedHistory = await loadExcelFile('history.xlsx');
            if (loadedHistory && Array.isArray(loadedHistory)) {
                jsonHistory = loadedHistory;
                history = jsonHistory;
                // Save to localStorage for offline use
                localStorage.setItem('declarationHistory', JSON.stringify(history));
                console.log('✅ History loaded from Excel file and saved to localStorage:', history.length, 'declarations');
            } else {
                history = [];
                console.warn('⚠️ No history found. Use "Gestion Données" → Import Excel to load history.xlsx.');
            }
        } catch (e) {
            console.error('❌ Could not load history from Excel file:', e);
            history = [];
        }
        
        // Sort by timestamp (newest first)
        history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        console.log('Total history loaded:', history.length, 'declarations');
    } catch (error) {
        console.error('Error loading history:', error);
        history = [];
    }
}

// Save history to localStorage and Excel file
async function saveHistory(newDeclaration = null) {
    // Sort by timestamp (newest first) before saving
    history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    // Save to localStorage (this always works, even if Excel save fails)
    localStorage.setItem('declarationHistory', JSON.stringify(history));
    console.log('✅ History saved to localStorage:', history.length, 'declarations');
    
    // Try to save to Excel file (optional - app works fine without it)
    try {
        // If a new declaration is provided, append only that row to Excel
        if (newDeclaration) {
            // Prepare declaration for Excel (convert arrays/objects to strings)
            const excelDeclaration = { ...newDeclaration };
            // Convert itineraire array to comma-separated string
            if (Array.isArray(excelDeclaration.itineraire)) {
                excelDeclaration.itineraire = excelDeclaration.itineraire.join(', ');
            }
            // Convert products array to JSON string for Excel
            if (Array.isArray(excelDeclaration.products)) {
                excelDeclaration.products = JSON.stringify(excelDeclaration.products);
            }
            
            // Try to append to Excel file first (preserves existing data)
            const appended = await appendRowToExcelFile('history.xlsx', excelDeclaration);
            if (!appended) {
                console.warn('⚠️ Failed to append to history.xlsx, trying full save...');
                // Fallback: Save all data if append failed
                // But first prepare all history data for Excel
                const excelHistory = history.map(decl => {
                    const excelDecl = { ...decl };
                    if (Array.isArray(excelDecl.itineraire)) {
                        excelDecl.itineraire = excelDecl.itineraire.join(', ');
                    }
                    if (Array.isArray(excelDecl.products)) {
                        excelDecl.products = JSON.stringify(excelDecl.products);
                    }
                    return excelDecl;
                });
                await saveToExcelFile('history.xlsx', excelHistory);
            } else {
                console.log('✅ Declaration appended successfully to data/history.xlsx');
            }
        } else {
            // No new declaration, save all data (for cases like import/merge)
            await saveToExcelFile('history.xlsx', history);
        }
    } catch (excelError) {
        // Excel save failed, but that's OK - data is in localStorage
        console.warn('⚠️ Could not save to Excel file. Data is safely stored in localStorage.', excelError);
        // Don't show error to user - the app works fine with localStorage only
    }
}

// Get next document number (auto-increment)
function getNextDocumentNumber() {
    // Check history to find the highest document number
    let maxNumber = 0;
    if (history && history.length > 0) {
        history.forEach(decl => {
            if (decl.documentNumber) {
                const match = decl.documentNumber.match(/^(\d+)\/(\d+)$/);
                if (match) {
                    const num = parseInt(match[1]);
                    const year = parseInt(match[2]);
                    const currentYear = new Date().getFullYear();
                    // Only count numbers from current year
                    if (year === currentYear && num > maxNumber) {
                        maxNumber = num;
                    }
                }
            }
        });
    }
    
    // Also check localStorage
    const storedNumber = parseInt(localStorage.getItem('lastDocumentNumber')) || 0;
    if (storedNumber > maxNumber) {
        maxNumber = storedNumber;
    }
    
    maxNumber++;
    localStorage.setItem('lastDocumentNumber', maxNumber);
    const currentYear = new Date().getFullYear();
    return `${maxNumber}/${currentYear}`;
}

// Initialize document number on page load
function initializeDocumentNumber() {
    const docNumberInput = document.getElementById('documentNumber');
    if (docNumberInput && !docNumberInput.value) {
        // Make sure history is loaded before getting next number
        if (history && history.length >= 0) {
            docNumberInput.value = getNextDocumentNumber();
        } else {
            // If history not loaded yet, use localStorage
            const storedNumber = parseInt(localStorage.getItem('lastDocumentNumber')) || 0;
            const currentYear = new Date().getFullYear();
            docNumberInput.value = `${storedNumber + 1}/${currentYear}`;
        }
    }
}

// Convert declaration to CSV row
function declarationToCSVRow(declaration) {
    const productsStr = declaration.products ? declaration.products.map(p => `${p.name} (${p.quantity} ${p.unit})`).join('; ') : '';
    const itineraireStr = Array.isArray(declaration.itineraire) ? declaration.itineraire.join(' - ') : '';
    
    return [
        declaration.documentNumber || '',
        declaration.date || '',
        declaration.dateDepart || '',
        declaration.clientName || '',
        declaration.destination || '',
        itineraireStr,
        declaration.driverName || '',
        declaration.driverCIN || '',
        declaration.driverPhone || '',
        declaration.vehicleModel || '',
        declaration.vehicleMatricule || '',
        declaration.convoyeurName || '',
        declaration.convoyeurCard || '',
        declaration.convoyeurCIN || '',
        declaration.convoyeurPhone || '',
        productsStr,
        declaration.passavantNumber || '',
        declaration.passavantExpiry || '',
        declaration.bonLivraison || '',
        declaration.timestamp || ''
    ].map(field => {
        // Escape commas and quotes in CSV
        const str = String(field || '');
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
            return `"${str.replace(/"/g, '""')}"`;
        }
        return str;
    }).join(',');
}

// Get CSV header
function getCSVHeader() {
    return [
        'N° Document',
        'Date',
        'Date Départ',
        'Client',
        'Destination',
        'Itinéraire',
        'Conducteur',
        'CIN Conducteur',
        'Téléphone Conducteur',
        'Modèle Véhicule',
        'Matricule Véhicule',
        'Convoyeur',
        'Carte de Contrôle Convoyeur',
        'CIN Convoyeur',
        'Téléphone Convoyeur',
        'Produits',
        'N° Passavant',
        'Expiration Passavant',
        'Bon de Livraison',
        'Date de Création'
    ].join(',');
}

// Generate CSV from history
function generateCSVFromHistory() {
    // Load history from localStorage (which contains merged data from history.xlsx)
    const historyData = localStorage.getItem('declarationHistory');
    let declarations = [];
    if (historyData) {
        declarations = JSON.parse(historyData);
        // Sort by timestamp (newest first)
        declarations.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    }
    
    if (declarations.length === 0) {
        return getCSVHeader() + '\n';
    }
    
    // Build CSV
    let csvData = getCSVHeader() + '\n';
    declarations.forEach(declaration => {
        csvData += declarationToCSVRow(declaration) + '\n';
    });
    
    return csvData;
}

// Download CSV file from history
function downloadCSV() {
    const csvData = generateCSVFromHistory();
    
    // Create blob and download
    const blob = new Blob(['\ufeff' + csvData], { type: 'text/csv;charset=utf-8;' }); // BOM for Excel
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `declarations_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Show history modal
async function showHistory(page = 1, filterText = '') {
    try {
        const modal = document.getElementById('historyModal');
        const historyContent = document.getElementById('historyContent');
        
        if (!modal || !historyContent) {
            console.error('History modal elements not found');
            return;
        }
        
        // Update current page and filter
        currentHistoryPage = page;
        historyFilter = filterText !== undefined ? filterText : (historyFilter || '');
        
        // Store current input state to preserve it
        const currentInput = document.getElementById('historyFilterInput');
        const cursorPosition = currentInput ? currentInput.selectionStart : null;
        
        // Only show loading if it's the first load
        if (!currentInput) {
            historyContent.innerHTML = '<p style="text-align: center; color: #666; padding: 20px;">Chargement...</p>';
        }
        modal.style.display = 'block';
        
        // Load and sort history
        await loadHistory();
        
        // Ensure history is an array
        if (!Array.isArray(history)) {
            history = [];
        }
        
        if (history.length === 0) {
            historyContent.innerHTML = '<p style="text-align: center; color: #666; padding: 20px;">Aucune déclaration dans l\'historique.</p>';
        } else {
            // Sort by timestamp (newest first) - this is the full sorted list
            const fullSortedHistory = [...history].sort((a, b) => {
                try {
                    const dateA = a.timestamp ? new Date(a.timestamp) : new Date(0);
                    const dateB = b.timestamp ? new Date(b.timestamp) : new Date(0);
                    return dateB - dateA;
                } catch (e) {
                    return 0;
                }
            });
            
            // Apply filter if provided
            let filteredHistory = fullSortedHistory;
            if (historyFilter && historyFilter.trim() !== '') {
                const filterLower = historyFilter.toLowerCase().trim();
                filteredHistory = fullSortedHistory.filter(decl => {
                    // Search in all fields
                    const docNumber = (decl.documentNumber || '').toLowerCase();
                    const date = decl.date ? new Date(decl.date).toLocaleDateString('fr-FR').toLowerCase() : '';
                    const client = (decl.clientName || '').toLowerCase();
                    const driver = (decl.driverName || '').toLowerCase();
                    const convoyeur = (decl.convoyeurName || '').toLowerCase();
                    const destination = (decl.destination || '').toLowerCase();
                    const products = decl.products && Array.isArray(decl.products) 
                        ? decl.products.map(p => `${p.name || ''} ${p.quantity || ''} ${p.unit || ''}`).join(' ').toLowerCase()
                        : '';
                    
                    return docNumber.includes(filterLower) ||
                           date.includes(filterLower) ||
                           client.includes(filterLower) ||
                           driver.includes(filterLower) ||
                           convoyeur.includes(filterLower) ||
                           destination.includes(filterLower) ||
                           products.includes(filterLower);
                });
            }
            
            // Calculate pagination
            const totalPages = Math.ceil(filteredHistory.length / historyItemsPerPage);
            const startIndex = (page - 1) * historyItemsPerPage;
            const endIndex = startIndex + historyItemsPerPage;
            const paginatedHistory = filteredHistory.slice(startIndex, endIndex);
            
            // Calculate the original index in fullSortedHistory for each item
            const historyHTML = paginatedHistory.map((decl, localIndex) => {
                try {
                    // Find the original index in the full sorted history array (not filtered)
                    const globalIndex = fullSortedHistory.findIndex(d => 
                        (d.id && decl.id && d.id === decl.id) || 
                        (d.timestamp === decl.timestamp && d.documentNumber === decl.documentNumber)
                    );
                    const safeIndex = globalIndex >= 0 ? globalIndex : startIndex + localIndex;
                    const productsStr = decl.products && Array.isArray(decl.products) 
                        ? decl.products.map(p => `${p.name || 'N/A'} (${p.quantity || '0'} ${p.unit || ''})`).join(', ') 
                        : 'Aucun';
                    let date = 'N/A';
                    try {
                        date = decl.date ? new Date(decl.date).toLocaleDateString('fr-FR') : 'N/A';
                    } catch (e) {
                        date = decl.date || 'N/A';
                    }
                    const rowClass = localIndex % 2 === 0 ? 'history-row-even' : 'history-row-odd';
                    const declId = decl.id || decl.timestamp || globalIndex;
                    return `
                        <tr class="${rowClass}" data-declaration-id="${declId}">
                            <td><strong>${decl.documentNumber || 'N/A'}</strong></td>
                            <td>${date}</td>
                            <td>${decl.clientName || 'N/A'}</td>
                            <td>${decl.driverName || 'N/A'}</td>
                            <td>${decl.convoyeurName || 'N/A'}</td>
                            <td>${productsStr}</td>
                            <td>
                                <button class="btn btn-sm btn-primary" onclick="viewDeclarationByIndex(${safeIndex})">👁️ Voir</button>
                                <button class="btn btn-sm btn-success" onclick="editDeclarationByIndex(${safeIndex})">✏️ Modifier</button>
                                <button class="btn btn-sm btn-danger" onclick="deleteDeclarationByIndex(${safeIndex})">🗑️</button>
                            </td>
                        </tr>
                    `;
                } catch (e) {
                    console.error('Error rendering history row:', e, decl);
                    return '';
                }
            }).join('');
            
            // Pagination controls
            let paginationHTML = '';
            if (totalPages > 1) {
                paginationHTML = `
                    <div class="history-pagination">
                        <button class="btn btn-sm btn-secondary" onclick="showHistory(${page - 1}, '${historyFilter.replace(/'/g, "\\'")}')" ${page === 1 ? 'disabled' : ''}>
                            ← Précédent
                        </button>
                        <span class="history-page-info">
                            Page ${page} sur ${totalPages} (${filteredHistory.length} déclaration${filteredHistory.length > 1 ? 's' : ''}${historyFilter ? ' filtrée' + (filteredHistory.length > 1 ? 's' : '') : ''})
                        </span>
                        <button class="btn btn-sm btn-secondary" onclick="showHistory(${page + 1}, '${historyFilter.replace(/'/g, "\\'")}')" ${page === totalPages ? 'disabled' : ''}>
                            Suivant →
                        </button>
                    </div>
                `;
            } else if (filteredHistory.length > 0) {
                paginationHTML = `
                    <div class="history-pagination">
                        <span class="history-page-info">
                            ${filteredHistory.length} déclaration${filteredHistory.length > 1 ? 's' : ''}${historyFilter ? ' trouvée' + (filteredHistory.length > 1 ? 's' : '') : ''}
                        </span>
                    </div>
                `;
            }
            
            // Filter input - Professional design
            const filterHTML = `
                <div class="history-filter-container">
                    <div style="display: flex; align-items: center; gap: 10px; flex: 1;">
                        <span style="font-size: 1.2em; color: #667eea;">🔍</span>
                        <input 
                            type="text" 
                            id="historyFilterInput" 
                            class="history-filter-input" 
                            placeholder="Rechercher par N° Document, Date, Client, Conducteur, Convoyeur, Produits..." 
                            value="${historyFilter.replace(/"/g, '&quot;').replace(/'/g, '&#39;')}"
                            oninput="filterHistory(this.value)"
                            onkeydown="handleHistoryFilterKeydown(event)"
                            autocomplete="off"
                        >
                    </div>
                    ${historyFilter ? `
                        <button class="btn btn-sm btn-secondary" onclick="clearHistoryFilter()" style="white-space: nowrap; padding: 10px 16px;">
                            <span style="margin-right: 5px;">✖️</span>Effacer
                        </button>
                    ` : ''}
                </div>
            `;
            
            // Check if filter container already exists - if so, only update table and pagination
            const existingFilterContainer = historyContent.querySelector('.history-filter-container');
            const existingTableContainer = historyContent.querySelector('.history-list-container');
            
            if (existingFilterContainer && existingTableContainer) {
                // Only update table body and pagination, preserve input
                const tbody = existingTableContainer.querySelector('tbody');
                if (tbody) {
                    tbody.innerHTML = historyHTML || '<tr><td colspan="7" style="text-align: center; padding: 20px; color: #666;">Aucune déclaration trouvée</td></tr>';
                }
                
                // Update pagination
                const existingPagination = historyContent.querySelector('.history-pagination');
                if (existingPagination) {
                    existingPagination.outerHTML = paginationHTML;
                } else if (paginationHTML) {
                    existingTableContainer.insertAdjacentHTML('afterend', paginationHTML);
                }
                
                // Update filter input value without losing focus
                const input = document.getElementById('historyFilterInput');
                if (input && input.value !== historyFilter) {
                    const cursorPos = input.selectionStart;
                    input.value = historyFilter;
                    setTimeout(() => {
                        if (cursorPos <= input.value.length) {
                            input.setSelectionRange(cursorPos, cursorPos);
                        }
                        input.focus();
                    }, 0);
                }
                
                // Update clear button visibility
                const clearBtn = existingFilterContainer.querySelector('button[onclick="clearHistoryFilter()"]');
                if (historyFilter && !clearBtn) {
                    existingFilterContainer.insertAdjacentHTML('beforeend', `
                        <button class="btn btn-sm btn-secondary" onclick="clearHistoryFilter()" style="white-space: nowrap; padding: 10px 16px;">
                            <span style="margin-right: 5px;">✖️</span>Effacer
                        </button>
                    `);
                } else if (!historyFilter && clearBtn) {
                    clearBtn.remove();
                }
            } else {
                // First time - create everything
                historyContent.innerHTML = `
                    ${filterHTML}
                    <div class="history-list-container">
                        <table class="history-table">
                            <thead>
                                <tr>
                                    <th>N° Document</th>
                                    <th>Date</th>
                                    <th>Client</th>
                                    <th>Conducteur</th>
                                    <th>Convoyeur</th>
                                    <th>Produits</th>
                                    <th>Actions</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${historyHTML || '<tr><td colspan="7" style="text-align: center; padding: 20px; color: #666;">Aucune déclaration trouvée</td></tr>'}
                            </tbody>
                        </table>
                    </div>
                    ${paginationHTML}
                `;
                
                // Focus input after first render
                setTimeout(() => {
                    const input = document.getElementById('historyFilterInput');
                    if (input) input.focus();
                }, 1000);
            }
        }
    } catch (error) {
        console.error('Error showing history:', error);
        const historyContent = document.getElementById('historyContent');
        if (historyContent) {
            historyContent.innerHTML = '<p style="text-align: center; color: #d32f2f; padding: 20px;">Erreur lors du chargement de l\'historique. Veuillez réessayer.</p>';
        }
    }
}

// Close history modal
function closeHistory() {
    document.getElementById('historyModal').style.display = 'none';
    // Reset filter when closing
    historyFilter = '';
}

// Filter history - instant filtering without delay, non-blocking
function filterHistory(filterText) {
    // Store cursor position before filtering (use window to survive DOM updates)
    const input = document.getElementById('historyFilterInput');
    if (input) {
        window.historyFilterCursorPos = input.selectionStart;
    }
    
    // Use requestAnimationFrame to avoid blocking the UI thread
    if (window.filterAnimationFrame) {
        cancelAnimationFrame(window.filterAnimationFrame);
    }
    window.filterAnimationFrame = requestAnimationFrame(() => {
        showHistory(1, filterText); // Reset to page 1 when filtering
    });
}

// Handle keyboard events in filter input
function handleHistoryFilterKeydown(event) {
    // Escape key to clear filter
    if (event.key === 'Escape' || event.key === 'Esc') {
        event.preventDefault();
        event.stopPropagation();
        clearHistoryFilter();
        return false;
    }
    // Allow all other keys to work normally - don't block typing
    return true;
}

// Clear history filter
function clearHistoryFilter() {
    historyFilter = '';
    const input = document.getElementById('historyFilterInput');
    if (input) {
        input.value = '';
    }
    showHistory(1, '');
    // Refocus the input after clearing
    setTimeout(() => {
        const newInput = document.getElementById('historyFilterInput');
        if (newInput) {
            newInput.focus();
        }
    }, 50);
}

// View specific declaration
async function viewDeclaration(index) {
    await loadHistory();
    const sortedHistory = [...history].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    if (sortedHistory[index]) {
        localStorage.setItem('currentDeclaration', JSON.stringify(sortedHistory[index]));
        window.location.href = 'declaration.html';
    }
}

// View declaration by index (wrapper to ensure history is loaded)
async function viewDeclarationByIndex(index) {
    await loadHistory();
    await viewDeclaration(index);
}

// Edit declaration - load into form
async function editDeclaration(index) {
    await loadHistory();
    const sortedHistory = [...history].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    if (!sortedHistory[index]) {
        alert('Déclaration non trouvée.');
        return;
    }
    
    const declaration = sortedHistory[index];
    
    // Store the declaration ID for update tracking
    localStorage.setItem('editingDeclarationId', declaration.id);
    
    // Close history modal
    closeHistory();
    
    // Wait a bit for modal to close
    setTimeout(() => {
        populateFormFromDeclaration(declaration);
        
        // Reset to step 1
        currentStep = 1;
        updateStepDisplay();
        
        // Scroll to top of form
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }, 300);
}

// Edit declaration by index (wrapper to ensure history is loaded)
async function editDeclarationByIndex(index) {
    await loadHistory();
    await editDeclaration(index);
}

// Populate form with declaration data
function populateFormFromDeclaration(declaration) {
    // Step 1: Client
    if (declaration.clientId) {
        document.getElementById('clientSelect').value = declaration.clientId;
        // Trigger change event to auto-fill related fields
        document.getElementById('clientSelect').dispatchEvent(new Event('change'));
    }
    if (declaration.destination) {
        document.getElementById('destination').value = declaration.destination;
    }
    
    // Step 2: Driver
    if (declaration.driverId) {
        document.getElementById('driverSelect').value = declaration.driverId;
        document.getElementById('driverSelect').dispatchEvent(new Event('change'));
    }
    if (declaration.driverCIN) {
        document.getElementById('driverCIN').value = declaration.driverCIN;
    }
    if (declaration.driverPhone) {
        document.getElementById('driverPhone').value = declaration.driverPhone;
    }
    if (declaration.vehicleMatricule) {
        document.getElementById('vehicleMatricule').value = declaration.vehicleMatricule;
    }
    if (declaration.vehicleModel) {
        document.getElementById('vehicleModel').value = declaration.vehicleModel;
    }
    
    // Step 3: Convoyeur
    if (declaration.convoyeurId) {
        document.getElementById('convoyeurSelect').value = declaration.convoyeurId;
        document.getElementById('convoyeurSelect').dispatchEvent(new Event('change'));
    }
    if (declaration.convoyeurCard) {
        document.getElementById('convoyeurCard').value = declaration.convoyeurCard;
    }
    if (declaration.convoyeurCIN) {
        document.getElementById('convoyeurCIN').value = declaration.convoyeurCIN;
    }
    if (declaration.convoyeurPhone) {
        document.getElementById('convoyeurPhone').value = declaration.convoyeurPhone;
    }
    
    // Step 4: Declaration details
    if (declaration.documentNumber) {
        document.getElementById('documentNumber').value = declaration.documentNumber;
    }
    if (declaration.date) {
        document.getElementById('date').value = declaration.date;
    }
    if (declaration.dateDepart) {
        // Handle datetime-local format
        let dateDepartValue = declaration.dateDepart;
        if (!dateDepartValue.includes('T')) {
            // If it's just a date, add time
            dateDepartValue = dateDepartValue + 'T00:00';
        }
        document.getElementById('dateDepart').value = dateDepartValue;
    }
    
    // Products
    const productsContainer = document.getElementById('productsContainer');
    // Clear existing products (keep first row)
    const existingRows = productsContainer.querySelectorAll('.product-row');
    existingRows.forEach((row, idx) => {
        if (idx > 0) {
            row.remove();
        } else {
            // Clear first row
            row.querySelector('.product-select').value = '';
            row.querySelector('.product-quantity').value = '';
            row.querySelector('.product-unit').value = '';
        }
    });
    
    // Add products from declaration
    if (declaration.products && declaration.products.length > 0) {
        // Ensure product selects are populated
        populateProductSelects();
        
        declaration.products.forEach((product, index) => {
            let productRow;
            if (index === 0) {
                // Use first row
                productRow = productsContainer.querySelector('.product-row');
            } else {
                // Add new row
                addProduct();
                productRow = productsContainer.querySelectorAll('.product-row')[index];
            }
            
            // Find product by name
            const productSelect = productRow.querySelector('.product-select');
            
            // Try to find product by exact name match first
            const matchingProduct = products.find(p => p.name === product.name);
            if (matchingProduct) {
                productSelect.value = matchingProduct.id;
                productSelect.dispatchEvent(new Event('change'));
            } else {
                // Try to find by partial match in options
                const productOption = Array.from(productSelect.options).find(opt => 
                    opt.textContent === product.name || opt.textContent.includes(product.name)
                );
                if (productOption && productOption.value) {
                    productSelect.value = productOption.value;
                    productSelect.dispatchEvent(new Event('change'));
                }
            }
            
            // Set quantity and unit (override auto-filled unit if needed)
            productRow.querySelector('.product-quantity').value = product.quantity || '';
            if (product.unit) {
                productRow.querySelector('.product-unit').value = product.unit;
            }
        });
        
        // Make selects searchable
        setTimeout(() => {
            document.querySelectorAll('.product-select').forEach(select => {
                makeSelectSearchable(select.id || null, select);
            });
        }, 100);
    }
    
    if (declaration.passavantNumber) {
        document.getElementById('passavantNumber').value = declaration.passavantNumber;
    }
    if (declaration.passavantExpiry) {
        document.getElementById('passavantExpiry').value = declaration.passavantExpiry;
    }
    if (declaration.bonLivraison) {
        document.getElementById('bonLivraison').value = declaration.bonLivraison;
    }
}

// Delete declaration
async function deleteDeclaration(index) {
    if (confirm('Êtes-vous sûr de vouloir supprimer cette déclaration ?')) {
        await loadHistory();
        const sortedHistory = [...history].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
        sortedHistory.splice(index, 1);
        history = sortedHistory;
        await saveHistory(); // Save to both localStorage and JSON file
        await showHistory(1); // Refresh the display (reset to page 1)
    }
}

// Delete declaration by index (wrapper to ensure history is loaded)
async function deleteDeclarationByIndex(index) {
    await loadHistory();
    await deleteDeclaration(index);
}

// Load history from file manually (using file input) - Excel only
async function loadHistoryFromFile() {
    if (typeof XLSX === 'undefined') {
        alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
        return;
    }

    // Create a file input element
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.onchange = async function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const firstSheet = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheet];
            const excelHistory = XLSX.utils.sheet_to_json(worksheet);
            
            // Merge with existing localStorage history
            const localHistoryData = localStorage.getItem('declarationHistory');
            let localHistory = [];
            if (localHistoryData) {
                localHistory = JSON.parse(localHistoryData);
            }
            
            // Combine and remove duplicates
            const combinedHistory = [...excelHistory, ...localHistory];
            const uniqueHistory = combinedHistory.filter((item, index, self) =>
                index === self.findIndex(t => t.id === item.id)
            );
            
            history = uniqueHistory;
            history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
            localStorage.setItem('declarationHistory', JSON.stringify(history));
            // Also save to Excel file
            await saveToExcelFile('history.xlsx', history);
            
            await showHistory(1); // Refresh the display (reset to page 1)
            alert(`Historique chargé avec succès ! ${history.length} déclaration(s) trouvée(s).`);
        } catch (error) {
            console.error('Error loading history file:', error);
            alert('Erreur lors du chargement du fichier: ' + error.message);
        }
    };
    input.click();
}

// Clear all history
async function clearHistory() {
    if (confirm('Êtes-vous sûr de vouloir effacer tout l\'historique ? Cette action est irréversible.')) {
        history = [];
        localStorage.removeItem('declarationHistory');
        localStorage.removeItem('lastDocumentNumber');
        // Also clear the Excel file
        await saveToExcelFile('history.xlsx', []);
        await showHistory(); // Refresh the display
        alert('Historique effacé avec succès.');
    }
}

// Populate client dropdown
function populateClientSelect() {
    const select = document.getElementById('clientSelect');
    if (!select) {
        console.error('clientSelect element not found');
        return;
    }
    
    // Clear existing options except the first one
    while (select.children.length > 1) {
        select.removeChild(select.lastChild);
    }
    
    // Show/hide empty message
    const emptyMessage = document.getElementById('clientEmptyMessage');
    
    if (clients && clients.length > 0) {
        clients.forEach(client => {
            const option = document.createElement('option');
            option.value = client.id;
            option.textContent = client.name;
            select.appendChild(option);
        });
        console.log('Clients loaded:', clients.length);
        if (emptyMessage) emptyMessage.style.display = 'none';
    } else {
        console.warn('No clients to populate');
        if (emptyMessage) emptyMessage.style.display = 'block';
    }
}

// Populate driver dropdown
function populateDriverSelect() {
    const select = document.getElementById('driverSelect');
    if (!select) {
        console.error('driverSelect element not found');
        return;
    }
    
    // Clear existing options except the first one
    while (select.children.length > 1) {
        select.removeChild(select.lastChild);
    }
    
    if (drivers && drivers.length > 0) {
    drivers.forEach(driver => {
        const option = document.createElement('option');
        option.value = driver.id;
        option.textContent = driver.name;
        select.appendChild(option);
    });
        console.log('Drivers loaded:', drivers.length);
    } else {
        console.warn('No drivers to populate');
    }
}

// Populate convoyeur dropdown
function populateConvoyeurSelect() {
    const select = document.getElementById('convoyeurSelect');
    if (!select) {
        console.error('convoyeurSelect element not found');
        return;
    }
    
    // Clear existing options except the first one
    while (select.children.length > 1) {
        select.removeChild(select.lastChild);
    }
    
    if (convoyeurs && convoyeurs.length > 0) {
    convoyeurs.forEach(convoyeur => {
        const option = document.createElement('option');
        option.value = convoyeur.id;
        option.textContent = convoyeur.name;
        select.appendChild(option);
    });
        console.log('Convoyeurs loaded:', convoyeurs.length);
    } else {
        console.warn('No convoyeurs to populate');
    }
}

// Setup event listeners for auto-fill
function setupEventListeners() {
    // Client selection
    document.getElementById('clientSelect').addEventListener('change', function() {
        const clientId = parseInt(this.value);
        const client = clients.find(c => c.id === clientId);
        if (client) {
            document.getElementById('destination').value = client.destination || '';
            
            // Display itinéraire
            const itineraireBox = document.getElementById('itineraireBox');
            if (client.itineraire && client.itineraire.length > 0) {
                itineraireBox.innerHTML = client.itineraire.map(point => 
                    `<div class="itineraire-item">${point}</div>`
                ).join('');
                itineraireBox.classList.remove('empty');
            } else {
                itineraireBox.innerHTML = '';
                itineraireBox.classList.add('empty');
            }
        } else {
            clearClientFields();
        }
    });

    // Driver selection
    document.getElementById('driverSelect').addEventListener('change', function() {
        const driverId = parseInt(this.value);
        const driver = drivers.find(d => d.id === driverId);
        if (driver) {
            document.getElementById('driverCIN').value = driver.cin || '';
            document.getElementById('driverPhone').value = driver.phone || '';
            document.getElementById('vehicleMatricule').value = driver.vehicle?.matricule || '';
            document.getElementById('vehicleModel').value = driver.vehicle?.model || '';
        } else {
            clearDriverFields();
        }
    });

    // Convoyeur selection
    document.getElementById('convoyeurSelect').addEventListener('change', function() {
        const convoyeurId = parseInt(this.value);
        const convoyeur = convoyeurs.find(c => c.id === convoyeurId);
        if (convoyeur) {
            document.getElementById('convoyeurCard').value = convoyeur.cce || convoyeur.card || '';
            document.getElementById('convoyeurCIN').value = convoyeur.cin || '';
            document.getElementById('convoyeurPhone').value = convoyeur.phone || '';
        } else {
            clearConvoyeurFields();
        }
    });

    // Form submission
    document.getElementById('declarationForm').addEventListener('submit', async function(e) {
        e.preventDefault();
        await generateDeclaration();
    });
    
    // Product selection change handlers
    document.querySelectorAll('.product-select').forEach(select => {
        select.addEventListener('change', function() {
            const selectedOption = this.options[this.selectedIndex];
            const unit = selectedOption.dataset.unit || '';
            const productRow = this.closest('.product-row');
            const unitInput = productRow.querySelector('.product-unit');
            unitInput.value = unit;
        });
    });
}

// Clear client fields
function clearClientFields() {
    document.getElementById('destination').value = '';
    const itineraireBox = document.getElementById('itineraireBox');
    itineraireBox.innerHTML = '';
    itineraireBox.classList.add('empty');
}

// Add client modal controls
function showAddClientModal() {
    const modal = document.getElementById('addClientModal');
    if (modal) modal.style.display = 'block';
}

function closeAddClientModal() {
    const modal = document.getElementById('addClientModal');
    if (modal) modal.style.display = 'none';
}

// Store file handles for direct file writing
const fileHandles = {};

// Excel file operations - Save and append functions

// Append a new row to an existing Excel file in data/ folder
async function appendRowToExcelFile(filename, newRow) {
    try {
        // Check if SheetJS is available
        if (typeof XLSX === 'undefined') {
            console.warn('SheetJS not available, skipping Excel append');
            return false;
        }

        // Prepare new row for Excel (convert arrays to strings where needed)
        const excelRow = { ...newRow };
        // Convert itineraire array to comma-separated string for Excel
        if (Array.isArray(excelRow.itineraire)) {
            excelRow.itineraire = excelRow.itineraire.join(', ');
        }
        // Convert products array to JSON string for Excel (for history declarations)
        if (Array.isArray(excelRow.products)) {
            excelRow.products = JSON.stringify(excelRow.products);
        }
        // Handle nested vehicle object for drivers
        if (excelRow.vehicle && typeof excelRow.vehicle === 'object') {
            excelRow['vehicle.matricule'] = excelRow.vehicle.matricule || '';
            excelRow['vehicle.model'] = excelRow.vehicle.model || '';
            delete excelRow.vehicle;
        }

        // Try to append using File System Access API
        if ('showDirectoryPicker' in window && dataDirectoryHandle) {
            try {
                // Get file handle
                let fileHandle;
                try {
                    fileHandle = await dataDirectoryHandle.getFileHandle(filename, { create: false });
                } catch (e) {
                    // File doesn't exist, create it
                    fileHandle = await dataDirectoryHandle.getFileHandle(filename, { create: true });
                }

                // Read existing file
                const file = await fileHandle.getFile();
                let existingData = [];
                let sheetName = 'Clients';
                
                if (file.size > 0) {
                    // File exists and has content, read it
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                    
                    // Determine sheet name
                    const firstSheetName = workbook.SheetNames[0];
                    sheetName = firstSheetName;
                    
                    // Get existing data
                    const worksheet = workbook.Sheets[firstSheetName];
                    existingData = XLSX.utils.sheet_to_json(worksheet);
                } else {
                    // New file, determine sheet name from filename
                    const name = filename.replace('.xlsx', '').replace('.xls', '');
                    if (name === 'clients') sheetName = 'Clients';
                    else if (name === 'drivers') sheetName = 'Conducteurs';
                    else if (name === 'convoyeurs') sheetName = 'Convoyeurs';
                    else if (name === 'products') sheetName = 'Produits';
                    else if (name === 'history') sheetName = 'Historique';
                }

                // Check if row already exists (by id or documentNumber for history)
                let rowExists = false;
                if (filename === 'history.xlsx' && excelRow.id) {
                    rowExists = existingData.some(row => row.id === excelRow.id);
                } else if (filename === 'history.xlsx' && excelRow.documentNumber) {
                    rowExists = existingData.some(row => row.documentNumber === excelRow.documentNumber);
                } else if (excelRow.id) {
                    rowExists = existingData.some(row => row.id === excelRow.id);
                }

                // Only append if row doesn't exist
                if (!rowExists) {
                    existingData.push(excelRow);
                } else {
                    console.log(`⚠️ Row already exists in ${filename}, skipping duplicate`);
                }

                // Create new workbook with updated data
                const worksheet = XLSX.utils.json_to_sheet(existingData);
                const workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);

                // Write back to file
                const excelBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
                const writable = await fileHandle.createWritable();
                await writable.write(new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
                await writable.close();

                console.log(`✅ تم إضافة صف جديد إلى data/${filename}`);
                return true;
            } catch (e) {
                console.error(`Error appending to Excel file ${filename}:`, e);
                // Fall through to fallback method
            }
        }

        // Fallback: If File System Access API is not available or failed,
        // we need to read from localStorage and save all data (not ideal, but works)
        console.warn(`⚠️ File System Access API غير متاح. سيتم حفظ جميع البيانات بدلاً من الإضافة فقط.`);
        return false; // Will trigger full save instead
    } catch (error) {
        console.error(`Error appending row to Excel file ${filename}:`, error);
        return false;
    }
}

// Save data to Excel file (tries to save directly to data/ folder, falls back to download)
async function saveToExcelFile(filename, data) {
    try {
        // Check if SheetJS is available
        if (typeof XLSX === 'undefined') {
            console.warn('SheetJS not available, skipping Excel save');
            return false;
        }

        // Prepare data for Excel (convert arrays to strings where needed)
        const excelData = data.map(item => {
            const excelItem = { ...item };
            // Convert itineraire array to comma-separated string for Excel
            if (Array.isArray(excelItem.itineraire)) {
                excelItem.itineraire = excelItem.itineraire.join(', ');
            }
            // Handle nested vehicle object for drivers
            if (excelItem.vehicle && typeof excelItem.vehicle === 'object') {
                excelItem['vehicle.matricule'] = excelItem.vehicle.matricule || '';
                excelItem['vehicle.model'] = excelItem.vehicle.model || '';
                delete excelItem.vehicle;
            }
            return excelItem;
        });

        // Create worksheet
        const worksheet = XLSX.utils.json_to_sheet(excelData);
        const workbook = XLSX.utils.book_new();
        
        // Determine sheet name from filename
        let sheetName = filename.replace('.xlsx', '').replace('.xls', '');
        if (sheetName === 'clients') sheetName = 'Clients';
        else if (sheetName === 'drivers') sheetName = 'Conducteurs';
        else if (sheetName === 'convoyeurs') sheetName = 'Convoyeurs';
        else if (sheetName === 'products') sheetName = 'Produits';
        else if (sheetName === 'history') sheetName = 'Historique';
        
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        // Try to save directly to data/ folder using File System Access API
        if ('showDirectoryPicker' in window) {
            try {
                // If we don't have a directory handle, ask user to select data/ folder (only once)
                if (!dataDirectoryHandle) {
                    // Check if we have a saved preference
                    const savedPath = sessionStorage.getItem('dataFolderPath');
                    if (!savedPath) {
                        // Ask user to select data folder (only first time per session)
                        const userChoice = confirm(
                            `💾 للحفظ المباشر في مجلد data/\n\n` +
                            `يرجى اختيار مجلد "data" من مشروعك.\n\n` +
                            `سيتم تذكر هذا الاختيار لهذه الجلسة فقط.\n\n` +
                            `هل تريد المتابعة؟`
                        );
                        if (userChoice) {
                            dataDirectoryHandle = await window.showDirectoryPicker({
                                startIn: 'desktop',
                                mode: 'readwrite'
                            });
                            // Verify it's the data folder
                            if (dataDirectoryHandle.name !== 'data') {
                                const confirmData = confirm(
                                    `⚠️ المجلد المحدد ليس "data".\n\n` +
                                    `اسم المجلد: "${dataDirectoryHandle.name}"\n\n` +
                                    `هل تريد المتابعة مع هذا المجلد؟`
                                );
                                if (!confirmData) {
                                    dataDirectoryHandle = null;
                                }
                            }
                        }
                    }
                }
                
                // If we have a directory handle, save directly
                if (dataDirectoryHandle) {
                    try {
                        // Convert workbook to buffer
                        const excelBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
                        
                        // Get file handle and write
                        const fileHandle = await dataDirectoryHandle.getFileHandle(filename, { create: true });
                        const writable = await fileHandle.createWritable();
                        await writable.write(new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
                        await writable.close();
                        
                        console.log(`✅ تم الحفظ مباشرة في data/${filename}`);
                        return true;
                    } catch (e) {
                        console.warn(`⚠️ فشل الحفظ المباشر، سيتم التنزيل بدلاً من ذلك:`, e);
                        // Fall through to download
                    }
                }
            } catch (e) {
                if (e.name !== 'AbortError') {
                    console.warn('File System Access API error:', e);
                }
                // Fall through to download
            }
        }
        
        // Fallback: Try to download the file (only if user interaction is allowed)
        try {
            XLSX.writeFile(workbook, filename);
            console.log(`✅ تم تنزيل الملف: ${filename}\n💡 انسخه إلى مجلد data/ في مشروعك`);
            return true;
        } catch (downloadError) {
            console.warn(`⚠️ Could not download ${filename}. Data is saved in localStorage.`, downloadError);
            // Data is still saved in localStorage, so the app continues to work
            return false;
        }
    } catch (error) {
        console.error(`Error saving Excel file ${filename}:`, error);
        return false;
    }
}

// Save all data directly to project files using File System Access API
async function saveAllDataToProjectFilesDirect() {
    // Check if File System Access API is available
    if (!('showDirectoryPicker' in window)) {
        alert('⚠️ Cette fonctionnalité nécessite Chrome ou Edge (version récente).\n\nUtilisez "Télécharger tous les fichiers" à la place.');
        return;
    }
    
    try {
        // Ask user to select the data folder (only first time)
        let dataDirHandle = null;
        const savedDirHandle = localStorage.getItem('dataDirHandle');
        
        if (savedDirHandle) {
            // Try to use saved handle (won't work after page reload, but good for session)
            try {
                // Note: File handles can't be stored in localStorage, so we'll ask again
                // But we can remember the path preference
            } catch (e) {
                console.log('Saved handle invalid, asking user again');
            }
        }
        
        // Ask user to select data folder
        dataDirHandle = await window.showDirectoryPicker({
            startIn: 'desktop',
            mode: 'readwrite'
        });
        
        // Ensure we have latest data from localStorage
        const savedClients = localStorage.getItem('clientsData');
        const savedDrivers = localStorage.getItem('driversData');
        const savedConvoyeurs = localStorage.getItem('convoyeursData');
        const savedProducts = localStorage.getItem('productsData');
        const savedHistory = localStorage.getItem('declarationHistory');
        
        if (savedClients) try { clients = JSON.parse(savedClients); } catch(e) {}
        if (savedDrivers) try { drivers = JSON.parse(savedDrivers); } catch(e) {}
        if (savedConvoyeurs) try { convoyeurs = JSON.parse(savedConvoyeurs); } catch(e) {}
        if (savedProducts) try { products = JSON.parse(savedProducts); } catch(e) {}
        if (savedHistory) try { history = JSON.parse(savedHistory); } catch(e) {}
        
        if (typeof XLSX === 'undefined') {
            alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
            return;
        }

        const filesToSave = [
            { name: 'clients.xlsx', data: clients || [] },
            { name: 'drivers.xlsx', data: drivers || [] },
            { name: 'convoyeurs.xlsx', data: convoyeurs || [] },
            { name: 'products.xlsx', data: products || [] },
            { name: 'history.xlsx', data: history || [] }
        ];
        
        let savedCount = 0;
        let totalItems = 0;
        
        // Save each file
        for (const file of filesToSave) {
            try {
                const fileHandle = await dataDirHandle.getFileHandle(file.name, { create: true });
                const writable = await fileHandle.createWritable();
                await writable.write(JSON.stringify(file.data, null, 2));
                await writable.close();
                
                savedCount++;
                totalItems += file.data.length;
                console.log(`✅ Saved ${file.name} directly to data folder`);
            } catch (e) {
                console.error(`Error saving ${file.name}:`, e);
            }
        }
        
        if (savedCount > 0) {
            alert(`✅ ${savedCount} fichier(s) sauvegardé(s) directement dans le dossier data/ !\n\n📊 Total: ${totalItems} élément(s)\n\n✅ Les fichiers de votre projet sont maintenant à jour !`);
        } else {
            alert('⚠️ Aucun fichier n\'a pu être sauvegardé.');
        }
    } catch (error) {
        if (error.name === 'AbortError') {
            console.log('User cancelled folder selection');
        } else {
            console.error('Error saving files:', error);
            alert(`❌ Erreur lors de la sauvegarde:\n\n${error.message}\n\nUtilisez "Télécharger tous les fichiers" à la place.`);
        }
    }
}

// Save all data to project files (download Excel files to replace data/ folder)
function saveAllDataToProjectFiles() {
    if (typeof XLSX === 'undefined') {
        alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
        return;
    }

    // Ensure we have latest data from localStorage
    const savedClients = localStorage.getItem('clientsData');
    const savedDrivers = localStorage.getItem('driversData');
    const savedConvoyeurs = localStorage.getItem('convoyeursData');
    const savedProducts = localStorage.getItem('productsData');
    const savedHistory = localStorage.getItem('declarationHistory');
    
    if (savedClients) try { clients = JSON.parse(savedClients); } catch(e) {}
    if (savedDrivers) try { drivers = JSON.parse(savedDrivers); } catch(e) {}
    if (savedConvoyeurs) try { convoyeurs = JSON.parse(savedConvoyeurs); } catch(e) {}
    if (savedProducts) try { products = JSON.parse(savedProducts); } catch(e) {}
    if (savedHistory) try { history = JSON.parse(savedHistory); } catch(e) {}
    
    const filesToSave = [
        { name: 'clients.xlsx', data: clients || [], count: (clients || []).length },
        { name: 'drivers.xlsx', data: drivers || [], count: (drivers || []).length },
        { name: 'convoyeurs.xlsx', data: convoyeurs || [], count: (convoyeurs || []).length },
        { name: 'products.xlsx', data: products || [], count: (products || []).length },
        { name: 'history.xlsx', data: history || [], count: (history || []).length }
    ];
    
    let totalItems = filesToSave.reduce((sum, f) => sum + f.count, 0);
    
    if (totalItems === 0) {
        const proceed = confirm('⚠️ Tous les fichiers sont vides.\n\nVoulez-vous quand même télécharger les fichiers vides ?');
        if (!proceed) return;
    }
    
    let savedCount = 0;
    let downloadIndex = 0;
    
    filesToSave.forEach((file, index) => {
        // Create Excel file for each data type
        setTimeout(() => {
            try {
                const worksheet = XLSX.utils.json_to_sheet(file.data);
                const workbook = XLSX.utils.book_new();
                const sheetName = file.name.replace('.xlsx', '').replace('.xls', '');
                XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
                XLSX.writeFile(workbook, file.name);
                savedCount++;
                console.log(`✅ Downloaded ${file.name} (${file.count} items)`);
            } catch (e) {
                console.error(`Error downloading ${file.name}:`, e);
            }
        }, index * 500); // 500ms delay between each file
    });
    
    // Show message after all downloads
    setTimeout(() => {
        const summary = filesToSave.map(f => `  • ${f.name}: ${f.count} élément(s)`).join('\n');
        alert(`✅ ${savedCount} fichier(s) Excel téléchargé(s) !\n\n📁 Fichiers téléchargés:\n${summary}\n\n📊 Total: ${totalItems} élément(s)\n\n💡 Instructions:\n1. Ouvrez le dossier "data" de votre projet\n2. Remplacez les fichiers existants par les fichiers Excel téléchargés\n3. Les fichiers du projet seront mis à jour\n\n💾 Emplacement: C:\\Users\\DELL\\Desktop\\project mvp\\data\\`);
    }, (filesToSave.length * 500) + 1000);
}

function getNextDriverId() {
    let maxId = 0;
    drivers.forEach(d => { if (d.id > maxId) maxId = d.id; });
    return maxId + 1;
}

function getNextConvoyeurId() {
    let maxId = 0;
    convoyeurs.forEach(c => { if (c.id > maxId) maxId = c.id; });
    return maxId + 1;
}

function getNextProductId() {
    let maxId = 0;
    products.forEach(p => { if (p.id > maxId) maxId = p.id; });
    return maxId + 1;
}

async function saveNewClient() {
    const name = document.getElementById('addClientName').value.trim();
    const destination = document.getElementById('addClientDestination').value.trim();
    const itineraireStr = document.getElementById('addClientItineraire').value.trim();

    if (!name || !destination) {
        alert('Veuillez remplir les champs Nom et Destination.');
        return;
    }

    const itineraire = itineraireStr ? itineraireStr.split(',').map(s => s.trim()).filter(Boolean) : [];
    const newClient = {
        id: getNextClientId(),
        name,
        destination,
        itineraire
    };

    // Update lists
    clients.push(newClient);
    
    // Save to localStorage (primary storage - works offline)
    localStorage.setItem('clientsData', JSON.stringify(clients));

    // Save to Excel file (will try to save directly or download)
    await saveToExcelFile('clients.xlsx', clients);

    populateClientSelect();
    makeSelectSearchable('clientSelect');

    // Select the newly added client
    const select = document.getElementById('clientSelect');
    if (select) {
        select.value = newClient.id;
        select.dispatchEvent(new Event('change'));
    }

    // Clear modal inputs
    document.getElementById('addClientName').value = '';
    document.getElementById('addClientDestination').value = '';
    document.getElementById('addClientItineraire').value = '';

    closeAddClientModal();
    alert('Client ajouté avec succès !');
}

// Convoyeur modal functions
function showAddConvoyeurModal() {
    const modal = document.getElementById('addConvoyeurModal');
    if (modal) modal.style.display = 'block';
}

function closeAddConvoyeurModal() {
    const modal = document.getElementById('addConvoyeurModal');
    if (modal) modal.style.display = 'none';
}

async function saveNewConvoyeur() {
    const name = document.getElementById('addConvoyeurName').value.trim();
    const cin = document.getElementById('addConvoyeurCIN').value.trim();
    const phone = document.getElementById('addConvoyeurPhone').value.trim();
    const cce = document.getElementById('addConvoyeurCCE').value.trim();

    if (!name || !cin || !phone) {
        alert('Veuillez remplir les champs Nom, CIN et Téléphone.');
        return;
    }

    const newConvoyeur = {
        id: getNextConvoyeurId(),
        name,
        cin,
        phone
    };

    if (cce) {
        newConvoyeur.cce = cce;
    }

    // Update lists
    convoyeurs.push(newConvoyeur);

    // Save to localStorage (primary storage)
    localStorage.setItem('convoyeursData', JSON.stringify(convoyeurs));

    // Try to append to Excel file first (preserves existing data)
    const appended = await appendRowToExcelFile('convoyeurs.xlsx', newConvoyeur);
    if (!appended) {
        // Fallback: Save all data if append failed
        await saveToExcelFile('convoyeurs.xlsx', convoyeurs);
    }

    populateConvoyeurSelect();
    makeSelectSearchable('convoyeurSelect');

    // Select the newly added convoyeur
    const select = document.getElementById('convoyeurSelect');
    if (select) {
        select.value = newConvoyeur.id;
        select.dispatchEvent(new Event('change'));
    }

    // Clear modal inputs
    document.getElementById('addConvoyeurName').value = '';
    document.getElementById('addConvoyeurCIN').value = '';
    document.getElementById('addConvoyeurPhone').value = '';
    document.getElementById('addConvoyeurCCE').value = '';

    closeAddConvoyeurModal();
    alert('Convoyeur ajouté avec succès !');
}

// Driver modal functions
function showAddDriverModal() {
    const modal = document.getElementById('addDriverModal');
    if (modal) modal.style.display = 'block';
}

function closeAddDriverModal() {
    const modal = document.getElementById('addDriverModal');
    if (modal) modal.style.display = 'none';
}

async function saveNewDriver() {
    const name = document.getElementById('addDriverName').value.trim();
    const cin = document.getElementById('addDriverCIN').value.trim();
    const phone = document.getElementById('addDriverPhone').value.trim();
    const matricule = document.getElementById('addDriverMatricule').value.trim();
    const model = document.getElementById('addDriverModel').value.trim();

    if (!name || !cin || !phone || !matricule || !model) {
        alert('Veuillez remplir tous les champs obligatoires.');
        return;
    }

    const newDriver = {
        id: getNextDriverId(),
        name,
        cin,
        phone,
        vehicle: {
            matricule,
            model
        }
    };

    // Update lists
    drivers.push(newDriver);

    // Save to localStorage (primary storage)
    localStorage.setItem('driversData', JSON.stringify(drivers));

    // Try to append to Excel file first (preserves existing data)
    const appended = await appendRowToExcelFile('drivers.xlsx', newDriver);
    if (!appended) {
        // Fallback: Save all data if append failed
        await saveToExcelFile('drivers.xlsx', drivers);
    }

    populateDriverSelect();
    makeSelectSearchable('driverSelect');

    // Select the newly added driver
    const select = document.getElementById('driverSelect');
    if (select) {
        select.value = newDriver.id;
        select.dispatchEvent(new Event('change'));
    }

    // Clear modal inputs
    document.getElementById('addDriverName').value = '';
    document.getElementById('addDriverCIN').value = '';
    document.getElementById('addDriverPhone').value = '';
    document.getElementById('addDriverMatricule').value = '';
    document.getElementById('addDriverModel').value = '';

    closeAddDriverModal();
    alert('Conducteur ajouté avec succès !');
}

// Product modal functions
function showAddProductModal() {
    const modal = document.getElementById('addProductModal');
    if (modal) modal.style.display = 'block';
}

function closeAddProductModal() {
    const modal = document.getElementById('addProductModal');
    if (modal) modal.style.display = 'none';
}

async function saveNewProduct() {
    const name = document.getElementById('addProductName').value.trim();
    const unit = document.getElementById('addProductUnit').value.trim();

    if (!name || !unit) {
        alert('Veuillez remplir tous les champs.');
        return;
    }

    const newProduct = {
        id: getNextProductId(),
        name,
        unit
    };

    // Update lists
    products.push(newProduct);

    // Save to localStorage (primary storage)
    localStorage.setItem('productsData', JSON.stringify(products));

    // Try to append to Excel file first (preserves existing data)
    const appended = await appendRowToExcelFile('products.xlsx', newProduct);
    if (!appended) {
        // Fallback: Save all data if append failed
        await saveToExcelFile('products.xlsx', products);
    }

    populateProductSelects();
    document.querySelectorAll('.product-select').forEach(select => {
        makeSelectSearchable(select.id || null, select);
    });

    // Clear modal inputs
    document.getElementById('addProductName').value = '';
    document.getElementById('addProductUnit').value = '';

    closeAddProductModal();
    alert('Produit ajouté avec succès !');
}

// Clear driver fields
function clearDriverFields() {
    document.getElementById('driverCIN').value = '';
    document.getElementById('driverPhone').value = '';
    document.getElementById('vehicleMatricule').value = '';
    document.getElementById('vehicleModel').value = '';
}

// Clear convoyeur fields
function clearConvoyeurFields() {
    document.getElementById('convoyeurCard').value = '';
    document.getElementById('convoyeurCIN').value = '';
    document.getElementById('convoyeurPhone').value = '';
}

// Populate product dropdowns
function populateProductSelects() {
    const selects = document.querySelectorAll('.product-select');
    selects.forEach(select => {
        // Clear existing options except the first one
        while (select.children.length > 1) {
            select.removeChild(select.lastChild);
        }
        // Add products
        products.forEach(product => {
            const option = document.createElement('option');
            option.value = product.id;
            option.textContent = product.name;
            option.dataset.unit = product.unit;
            select.appendChild(option);
        });
    });
}

// Add product row
function addProduct() {
    const container = document.getElementById('productsContainer');
    const productRow = document.createElement('div');
    productRow.className = 'product-row';
    productRow.innerHTML = `
        <div class="form-group">
            <label>Produit</label>
            <select class="form-control product-select searchable-select" required>
                <option value="">Sélectionner...</option>
            </select>
        </div>
        <div class="form-group">
            <label>Quantité</label>
            <input type="number" class="form-control product-quantity" placeholder="0" min="0" step="0.01" required>
        </div>
        <div class="form-group">
            <label>Unité</label>
            <input type="text" class="form-control product-unit" readonly>
        </div>
        <div class="form-group">
            <label style="opacity: 0;">Action</label>
            <button type="button" class="btn btn-danger" onclick="removeProduct(this)">🗑️</button>
        </div>
    `;
    container.appendChild(productRow);
    
    // Populate the new select
    const newSelect = productRow.querySelector('.product-select');
    products.forEach(product => {
        const option = document.createElement('option');
        option.value = product.id;
        option.textContent = product.name;
        option.dataset.unit = product.unit;
        newSelect.appendChild(option);
    });
    
    // Make the select searchable
    makeSelectSearchable(null, newSelect);
    
    // Add event listener for product selection
    newSelect.addEventListener('change', function() {
        const selectedOption = this.options[this.selectedIndex];
        const unit = selectedOption.dataset.unit || '';
        const unitInput = productRow.querySelector('.product-unit');
        unitInput.value = unit;
    });
}

// Remove product row
function removeProduct(button) {
    const container = document.getElementById('productsContainer');
    const productRow = button.closest('.product-row');
    if (container.children.length > 1) {
        productRow.remove();
    } else {
        alert('Vous devez avoir au moins un produit.');
    }
}

// Step navigation
let currentStep = 1;
const totalSteps = 5;

// History pagination
let currentHistoryPage = 1;
const historyItemsPerPage = 5;
let historyFilter = '';

function updateStepDisplay() {
    // Update form sections
    document.querySelectorAll('.form-section').forEach((section, index) => {
        if (index + 1 === currentStep) {
            section.classList.add('active');
        } else {
            section.classList.remove('active');
        }
    });
    
    // Update progress steps
    document.querySelectorAll('.step').forEach((step, index) => {
        const stepNum = index + 1;
        step.classList.remove('active', 'completed');
        if (stepNum < currentStep) {
            step.classList.add('completed');
        } else if (stepNum === currentStep) {
            step.classList.add('active');
        }
    });
    
    // Update progress line
    const progress = ((currentStep - 1) / (totalSteps - 1)) * 100;
    document.getElementById('progressLine').style.width = progress + '%';
    
    // Update navigation buttons
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    const submitBtn = document.getElementById('submitBtn');
    
    prevBtn.style.display = currentStep > 1 ? 'inline-block' : 'none';
    
    if (currentStep === totalSteps) {
        nextBtn.style.display = 'none';
        submitBtn.style.display = 'inline-block';
        generateSummary();
    } else {
        nextBtn.style.display = 'inline-block';
        submitBtn.style.display = 'none';
    }
}

function nextStep() {
    const currentSection = document.querySelector(`.form-section[data-section="${currentStep}"]`);
    const requiredFields = currentSection.querySelectorAll('[required]');
    let isValid = true;
    
    requiredFields.forEach(field => {
        if (!field.value) {
            isValid = false;
            field.classList.add('error');
        } else {
            field.classList.remove('error');
        }
    });
    
    if (isValid) {
        if (currentStep < totalSteps) {
            currentStep++;
            updateStepDisplay();
        }
    } else {
        alert('Veuillez remplir tous les champs obligatoires.');
    }
}

function prevStep() {
    if (currentStep > 1) {
        currentStep--;
        updateStepDisplay();
    }
}

function generateSummary() {
    const summaryContent = document.getElementById('summaryContent');
    const clientId = parseInt(document.getElementById('clientSelect').value);
    const driverId = parseInt(document.getElementById('driverSelect').value);
    const convoyeurId = parseInt(document.getElementById('convoyeurSelect').value);
    
    const client = clients.find(c => c.id === clientId);
    const driver = drivers.find(d => d.id === driverId);
    const convoyeur = convoyeurs.find(c => c.id === convoyeurId);
    
    const productRows = document.querySelectorAll('.product-row');
    const productsList = [];
    productRows.forEach(row => {
        const select = row.querySelector('.product-select');
        const quantity = row.querySelector('.product-quantity').value;
        const unit = row.querySelector('.product-unit').value;
        if (select.value && quantity && unit) {
            const product = products.find(p => p.id === parseInt(select.value));
            productsList.push(`${product ? product.name : select.options[select.selectedIndex].text}: ${quantity} ${unit}`);
        }
    });
    
    summaryContent.innerHTML = `
        <div class="summary-item">
            <strong>Client:</strong> ${client ? client.name : 'N/A'}<br>
            <strong>Destination:</strong> ${document.getElementById('destination').value || 'N/A'}
        </div>
        <div class="summary-item">
            <strong>Conducteur:</strong> ${driver ? driver.name : 'N/A'}<br>
            <strong>CIN:</strong> ${document.getElementById('driverCIN').value || 'N/A'}<br>
            <strong>Véhicule:</strong> ${document.getElementById('vehicleModel').value || 'N/A'} - ${document.getElementById('vehicleMatricule').value || 'N/A'}
        </div>
        <div class="summary-item">
            <strong>Convoyeur:</strong> ${convoyeur ? convoyeur.name : 'N/A'}<br>
            <strong>Carte de Contrôle:</strong> ${document.getElementById('convoyeurCard')?.value || 'N/A'}<br>
            <strong>CIN:</strong> ${document.getElementById('convoyeurCIN').value || 'N/A'}
        </div>
        <div class="summary-item">
            <strong>Produits:</strong><br>
            ${productsList.length > 0 ? productsList.map(p => `• ${p}`).join('<br>') : 'Aucun produit'}
        </div>
        <div class="summary-item">
            <strong>N° Document:</strong> ${document.getElementById('documentNumber').value || 'N/A'}<br>
            <strong>Date:</strong> ${document.getElementById('date').value || 'N/A'}<br>
            <strong>Date Départ:</strong> ${document.getElementById('dateDepart').value || 'N/A'}
        </div>
    `;
}

// Generate declaration
async function generateDeclaration() {
    // Ensure we have the latest history loaded
    await loadHistory();
    
    // Get form values
    const clientId = parseInt(document.getElementById('clientSelect').value);
    const driverId = parseInt(document.getElementById('driverSelect').value);
    const convoyeurId = parseInt(document.getElementById('convoyeurSelect').value);

    const client = clients.find(c => c.id === clientId);
    const driver = drivers.find(d => d.id === driverId);
    const convoyeur = convoyeurs.find(c => c.id === convoyeurId);

    // Collect products
    const formProducts = [];
    const productRows = document.querySelectorAll('.product-row');
    productRows.forEach(row => {
        const select = row.querySelector('.product-select');
        const productId = parseInt(select.value);
        const quantity = row.querySelector('.product-quantity').value;
        const unit = row.querySelector('.product-unit').value;
        
        if (productId && quantity && unit) {
            const selectedProduct = products.find(p => p.id === productId);
            const productName = selectedProduct ? selectedProduct.name : select.options[select.selectedIndex].text;
            formProducts.push({
                name: productName,
                quantity: quantity,
                unit: unit
            });
        }
    });

    // Get or generate document number (auto-increment)
    let docNumber = document.getElementById('documentNumber').value;
    if (!docNumber) {
        docNumber = getNextDocumentNumber();
    }

    // Check if we're editing an existing declaration
    const editingId = localStorage.getItem('editingDeclarationId');
    let isEditing = false;
    let existingDeclarationIndex = -1;
    
    if (editingId) {
        existingDeclarationIndex = history.findIndex(d => d.id == editingId || d.id === parseInt(editingId));
        isEditing = existingDeclarationIndex !== -1;
    }

    // Create declaration object
    const declaration = {
        id: isEditing ? history[existingDeclarationIndex].id : Date.now(), // Keep original ID if editing
        timestamp: isEditing ? history[existingDeclarationIndex].timestamp : new Date().toISOString(), // Keep original timestamp if editing
        documentNumber: docNumber,
        date: document.getElementById('date').value,
        dateDepart: document.getElementById('dateDepart').value,
        clientId: clientId,
        clientName: client ? client.name : '',
        destination: document.getElementById('destination').value,
        itineraire: client ? client.itineraire : [],
        driverId: driverId,
        driverName: driver ? driver.name : '',
        driverCIN: document.getElementById('driverCIN').value,
        driverPhone: document.getElementById('driverPhone').value,
        vehicleMatricule: document.getElementById('vehicleMatricule').value,
        vehicleModel: document.getElementById('vehicleModel').value,
        convoyeurId: convoyeurId,
        convoyeurName: convoyeur ? convoyeur.name : '',
        convoyeurCard: document.getElementById('convoyeurCard').value,
        convoyeurCIN: document.getElementById('convoyeurCIN').value,
        convoyeurPhone: document.getElementById('convoyeurPhone').value,
        products: formProducts,
        passavantNumber: document.getElementById('passavantNumber').value,
        passavantExpiry: document.getElementById('passavantExpiry').value,
        bonLivraison: document.getElementById('bonLivraison').value
    };

    if (isEditing) {
        // Update existing declaration
        history[existingDeclarationIndex] = declaration;
        // Remove editing flag
        localStorage.removeItem('editingDeclarationId');
    } else {
        // Add new declaration to history
        history.push(declaration);
    }
    
    // Ensure dataDirectoryHandle is set before saving
    if (!dataDirectoryHandle && 'showDirectoryPicker' in window) {
        try {
            const userChoice = confirm(
                `💾 للحفظ المباشر في مجلد data/\n\n` +
                `يرجى اختيار مجلد "data" من مشروعك.\n\n` +
                `سيتم تذكر هذا الاختيار لهذه الجلسة.\n\n` +
                `هل تريد المتابعة؟`
            );
            if (userChoice) {
                dataDirectoryHandle = await window.showDirectoryPicker({
                    startIn: 'desktop',
                    mode: 'readwrite'
                });
                // Verify it's the data folder
                if (dataDirectoryHandle.name !== 'data') {
                    const confirmData = confirm(
                        `⚠️ المجلد المحدد ليس "data".\n\n` +
                        `اسم المجلد: "${dataDirectoryHandle.name}"\n\n` +
                        `هل تريد المتابعة مع هذا المجلد؟`
                    );
                    if (!confirmData) {
                        dataDirectoryHandle = null;
                    }
                }
            }
        } catch (e) {
            if (e.name !== 'AbortError') {
                console.warn('Error selecting data folder:', e);
            }
        }
    }
    
    // Save to both localStorage and Excel file
    if (isEditing) {
        // For updates, save all history (since we're modifying an existing entry)
        await saveHistory();
        console.log('✅ Declaration updated in history.xlsx');
    } else {
        // For new declarations, append only
        await saveHistory(declaration);
        console.log('✅ New declaration saved to history.xlsx');
    }

    // Save current declaration to localStorage for declaration.html
    localStorage.setItem('currentDeclaration', JSON.stringify(declaration));

    // Open declaration page
    window.location.href = 'declaration.html';
}

// Make select searchable
function makeSelectSearchable(selectId, selectElement = null) {
    const select = selectElement || document.getElementById(selectId);
    if (!select || select.classList.contains('searchable-converted')) {
        return; // Already converted or doesn't exist
    }
    
    select.classList.add('searchable-converted');
    
    // Create wrapper
    const wrapper = document.createElement('div');
    wrapper.className = 'searchable-select-wrapper';
    wrapper.style.position = 'relative';
    
    // Create input for search
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'searchable-select-input form-control';
    input.placeholder = select.options[0]?.text || 'Rechercher...';
    input.autocomplete = 'off';
    
    // Create dropdown
    const dropdown = document.createElement('div');
    dropdown.className = 'searchable-select-dropdown';
    
    // Store original select
    const originalSelect = select;
    
    // Function to update input and dropdown
    function updateDisplay() {
        const selectedOption = originalSelect.options[originalSelect.selectedIndex];
        if (selectedOption && selectedOption.value) {
            input.value = selectedOption.text;
        } else {
            input.value = '';
        }
        updateDropdown();
    }
    
    // Function to update dropdown options
    function updateDropdown() {
        const searchTerm = input.value.toLowerCase();
        dropdown.innerHTML = '';
        
        let hasVisibleOptions = false;
        for (let i = 0; i < originalSelect.options.length; i++) {
            const option = originalSelect.options[i];
            const optionText = option.text.toLowerCase();
            
            if (option.value === '' || optionText.includes(searchTerm)) {
                const optionDiv = document.createElement('div');
                optionDiv.className = 'searchable-select-option';
                if (option.value === originalSelect.value) {
                    optionDiv.classList.add('selected');
                }
                optionDiv.textContent = option.text;
                optionDiv.dataset.value = option.value;
                
                optionDiv.addEventListener('click', function() {
                    originalSelect.value = option.value;
                    input.value = option.text;
                    dropdown.classList.remove('show');
                    originalSelect.dispatchEvent(new Event('change'));
                    updateDisplay();
                });
                
                dropdown.appendChild(optionDiv);
                hasVisibleOptions = true;
            }
        }
        
        if (!hasVisibleOptions) {
            const noResults = document.createElement('div');
            noResults.className = 'searchable-select-option';
            noResults.textContent = 'Aucun résultat';
            noResults.style.color = '#999';
            noResults.style.cursor = 'default';
            dropdown.appendChild(noResults);
        }
    }
    
    // Input event listeners
    input.addEventListener('focus', function() {
        dropdown.classList.add('show');
        updateDropdown();
    });
    
    input.addEventListener('input', function() {
        updateDropdown();
        dropdown.classList.add('show');
    });
    
    input.addEventListener('keydown', function(e) {
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp' || e.key === 'Enter') {
            e.preventDefault();
            const options = dropdown.querySelectorAll('.searchable-select-option:not(.hidden)');
            const selected = dropdown.querySelector('.searchable-select-option.selected');
            let currentIndex = -1;
            
            if (selected) {
                options.forEach((opt, idx) => {
                    if (opt === selected) currentIndex = idx;
                });
            }
            
            if (e.key === 'ArrowDown') {
                currentIndex = (currentIndex + 1) % options.length;
            } else if (e.key === 'ArrowUp') {
                currentIndex = (currentIndex - 1 + options.length) % options.length;
            } else if (e.key === 'Enter' && selected) {
                selected.click();
                return;
            }
            
            options.forEach(opt => opt.classList.remove('selected'));
            if (options[currentIndex]) {
                options[currentIndex].classList.add('selected');
                options[currentIndex].scrollIntoView({ block: 'nearest' });
            }
        } else if (e.key === 'Escape') {
            dropdown.classList.remove('show');
            updateDisplay();
        }
    });
    
    // Close dropdown when clicking outside
    document.addEventListener('click', function(e) {
        if (!wrapper.contains(e.target)) {
            dropdown.classList.remove('show');
            updateDisplay();
        }
    });
    
    // Store the parent and next sibling
    const parent = select.parentNode;
    const nextSibling = select.nextSibling;
    
    // Insert wrapper before select
    parent.insertBefore(wrapper, select);
    
    // Move select inside wrapper and hide it (but keep it accessible for form validation)
    wrapper.appendChild(select);
    wrapper.appendChild(input);
    wrapper.appendChild(dropdown);
    select.style.position = 'absolute';
    select.style.left = '-9999px';
    select.style.width = '1px';
    select.style.height = '1px';
    select.style.opacity = '0';
    select.tabIndex = -1;
    
    // Initialize display
    updateDisplay();
    
    // Update when select changes externally
    select.addEventListener('change', updateDisplay);
}

// ==================== EXPORT/IMPORT EXCEL FUNCTIONS ====================
// JSON functions removed - Excel only

// JSON import removed - Excel only. Use importDataFromExcel instead.

// Export all data (backup) - Excel only
function exportAllData() {
    exportDataToExcel('all');
}

// Merge data from another machine (combines instead of replacing) - Excel only
function mergeDataFromFile() {
    if (typeof XLSX === 'undefined') {
        alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
        return;
    }

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    
    input.onchange = async function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        try {
            const data = await file.arrayBuffer();
            const workbook = XLSX.read(data);
            const jsonData = {};
            
            // Read all sheets
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const sheetData = XLSX.utils.sheet_to_json(worksheet);
                if (sheetName === 'Clients' || sheetName === 'clients') {
                    // Reconstruct client data (convert itineraire string to array)
                    jsonData.clients = sheetData.map(reconstructClient);
                } else if (sheetName === 'Conducteurs' || sheetName === 'drivers') {
                    jsonData.drivers = sheetData;
                } else if (sheetName === 'Convoyeurs' || sheetName === 'convoyeurs') {
                    jsonData.convoyeurs = sheetData;
                } else if (sheetName === 'Produits' || sheetName === 'products') {
                    jsonData.products = sheetData;
                } else if (sheetName === 'Historique' || sheetName === 'history') {
                    jsonData.history = sheetData;
                }
            });
            
            let mergedCount = 0;
            let mergedTypes = [];
            
            // Merge clients
            if (Array.isArray(jsonData.clients) && jsonData.clients.length > 0) {
                const existingClients = clients || [];
                // Reconstruct client data (convert itineraire string to array)
                const newClients = jsonData.clients.map(reconstructClient);
                
                // Merge and remove duplicates by id
                const merged = [...existingClients, ...newClients];
                clients = merged.filter((item, index, self) =>
                    index === self.findIndex(t => t.id === item.id)
                );
                
                localStorage.setItem('clientsData', JSON.stringify(clients));
                populateClientSelect();
                setTimeout(() => makeSelectSearchable('clientSelect'), 100);
                mergedCount += (clients.length - existingClients.length);
                mergedTypes.push(`✅ ${clients.length - existingClients.length} nouveaux clients (Total: ${clients.length})`);
            }
            
            // Merge drivers
            if (Array.isArray(jsonData.drivers) && jsonData.drivers.length > 0) {
                const existingDrivers = drivers || [];
                const newDrivers = jsonData.drivers;
                
                const merged = [...existingDrivers, ...newDrivers];
                drivers = merged.filter((item, index, self) =>
                    index === self.findIndex(t => t.id === item.id)
                );
                
                localStorage.setItem('driversData', JSON.stringify(drivers));
                populateDriverSelect();
                setTimeout(() => makeSelectSearchable('driverSelect'), 100);
                mergedCount += (drivers.length - existingDrivers.length);
                mergedTypes.push(`✅ ${drivers.length - existingDrivers.length} nouveaux conducteurs (Total: ${drivers.length})`);
            }
            
            // Merge convoyeurs
            if (Array.isArray(jsonData.convoyeurs) && jsonData.convoyeurs.length > 0) {
                const existingConvoyeurs = convoyeurs || [];
                const newConvoyeurs = jsonData.convoyeurs;
                
                const merged = [...existingConvoyeurs, ...newConvoyeurs];
                convoyeurs = merged.filter((item, index, self) =>
                    index === self.findIndex(t => t.id === item.id)
                );
                
                localStorage.setItem('convoyeursData', JSON.stringify(convoyeurs));
                populateConvoyeurSelect();
                setTimeout(() => makeSelectSearchable('convoyeurSelect'), 100);
                mergedCount += (convoyeurs.length - existingConvoyeurs.length);
                mergedTypes.push(`✅ ${convoyeurs.length - existingConvoyeurs.length} nouveaux convoyeurs (Total: ${convoyeurs.length})`);
            }
            
            // Merge products
            if (Array.isArray(jsonData.products) && jsonData.products.length > 0) {
                const existingProducts = products || [];
                const newProducts = jsonData.products;
                
                const merged = [...existingProducts, ...newProducts];
                products = merged.filter((item, index, self) =>
                    index === self.findIndex(t => t.id === item.id)
                );
                
                localStorage.setItem('productsData', JSON.stringify(products));
                populateProductSelects();
                setTimeout(() => {
                    document.querySelectorAll('.product-select').forEach(select => {
                        makeSelectSearchable(select.id || null, select);
                    });
                }, 100);
                mergedCount += (products.length - existingProducts.length);
                mergedTypes.push(`✅ ${products.length - existingProducts.length} nouveaux produits (Total: ${products.length})`);
            }
            
            // Merge history
            if (Array.isArray(jsonData.history) && jsonData.history.length > 0) {
                const existingHistory = history || [];
                const newHistory = jsonData.history;
                
                const merged = [...existingHistory, ...newHistory];
                history = merged.filter((item, index, self) =>
                    index === self.findIndex(t => t.id === item.id)
                );
                
                localStorage.setItem('declarationHistory', JSON.stringify(history));
                mergedCount += (history.length - existingHistory.length);
                mergedTypes.push(`✅ ${history.length - existingHistory.length} nouvelles déclarations (Total: ${history.length})`);
            }
            
            if (mergedCount > 0) {
                const message = `✅ Fusion réussie !\n\n${mergedTypes.join('\n')}\n\n📊 Total: ${mergedCount} nouvel(le)(s) élément(s) ajouté(s).\n\n💡 Les données ont été fusionnées, aucune donnée n'a été perdue.`;
                alert(message);
                
                // Ask if user wants to save merged data to files
                const saveToFiles = confirm('💾 Voulez-vous sauvegarder les données fusionnées dans les fichiers du projet ?');
                if (saveToFiles) {
                    saveAllDataToProjectFiles();
                }
                
                // Reload page
                setTimeout(() => {
                    location.reload();
                }, saveToFiles ? 3000 : 500);
            } else {
                alert('ℹ️ Aucune nouvelle donnée à fusionner. Toutes les données existent déjà.');
            }
        } catch (error) {
            console.error('Error merging data:', error);
            alert('❌ Erreur lors de la fusion:\n\n' + error.message);
        }
    };
    
    input.click();
}

// Import all data (restore) - Excel only
function importAllData() {
    if (typeof XLSX === 'undefined') {
        alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
        return;
    }

    if (confirm('⚠️ Cette action va remplacer toutes vos données actuelles. Êtes-vous sûr ?\n\n💡 Pour fusionner les données au lieu de les remplacer, utilisez "Fusionner les données" à la place.')) {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        
        input.onchange = async function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            try {
                const data = await file.arrayBuffer();
                const workbook = XLSX.read(data);
                const jsonData = {};
                
                // Read all sheets
                workbook.SheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const sheetData = XLSX.utils.sheet_to_json(worksheet);
                    if (sheetName === 'Clients' || sheetName === 'clients') {
                        // Reconstruct client data (convert itineraire string to array)
                        jsonData.clients = sheetData.map(reconstructClient);
                    } else if (sheetName === 'Conducteurs' || sheetName === 'drivers') {
                        jsonData.drivers = sheetData;
                    } else if (sheetName === 'Convoyeurs' || sheetName === 'convoyeurs') {
                        jsonData.convoyeurs = sheetData;
                    } else if (sheetName === 'Produits' || sheetName === 'products') {
                        jsonData.products = sheetData;
                    } else if (sheetName === 'Historique' || sheetName === 'history') {
                        jsonData.history = sheetData;
                    }
                });
                
                let totalImported = 0;
                let importedTypes = [];
                
                // Handle backup file with all data (from Export Tout)
                if (jsonData.clients !== undefined || jsonData.drivers !== undefined || 
                    jsonData.convoyeurs !== undefined || jsonData.products !== undefined || 
                    jsonData.history !== undefined) {
                    
                    // This is a backup file with all data
                    if (Array.isArray(jsonData.clients)) {
                        // Reconstruct client data (convert itineraire string to array)
                        clients = jsonData.clients.map(reconstructClient);
                        localStorage.setItem('clientsData', JSON.stringify(clients));
                        populateClientSelect();
                        setTimeout(() => makeSelectSearchable('clientSelect'), 100);
                        totalImported += clients.length;
                        importedTypes.push(`✅ ${clients.length} clients`);
                    }
                    
                    if (Array.isArray(jsonData.drivers)) {
                        drivers = jsonData.drivers;
                        localStorage.setItem('driversData', JSON.stringify(drivers));
                        populateDriverSelect();
                        setTimeout(() => makeSelectSearchable('driverSelect'), 100);
                        totalImported += drivers.length;
                        importedTypes.push(`✅ ${drivers.length} conducteurs`);
                    }
                    
                    if (Array.isArray(jsonData.convoyeurs)) {
                        convoyeurs = jsonData.convoyeurs;
                        localStorage.setItem('convoyeursData', JSON.stringify(convoyeurs));
                        populateConvoyeurSelect();
                        setTimeout(() => makeSelectSearchable('convoyeurSelect'), 100);
                        totalImported += convoyeurs.length;
                        importedTypes.push(`✅ ${convoyeurs.length} convoyeurs`);
                    }
                    
                    if (Array.isArray(jsonData.products)) {
                        products = jsonData.products;
                        localStorage.setItem('productsData', JSON.stringify(products));
                        populateProductSelects();
                        setTimeout(() => {
                            document.querySelectorAll('.product-select').forEach(select => {
                                makeSelectSearchable(select.id || null, select);
                            });
                        }, 100);
                        totalImported += products.length;
                        importedTypes.push(`✅ ${products.length} produits`);
                    }
                    
                    if (Array.isArray(jsonData.history)) {
                        history = jsonData.history;
                        localStorage.setItem('declarationHistory', JSON.stringify(history));
                        totalImported += history.length;
                        importedTypes.push(`✅ ${history.length} déclarations`);
                    }
                    
                    if (totalImported > 0) {
                        const message = `✅ Importation réussie !\n\n${importedTypes.join('\n')}\n\n📊 Total: ${totalImported} élément(s) importé(s).`;
                        alert(message);
                        
                        // Ask if user wants to save to project files
                        const saveToFiles = confirm(`💾 Voulez-vous sauvegarder ces données dans les fichiers du projet (data/) ?\n\n(Cela téléchargera tous les fichiers pour que vous puissiez les copier dans le dossier data/)`);
                        
                        if (saveToFiles) {
                            saveAllDataToProjectFiles();
                        }
                        
                        // Reload page to refresh all data
                        setTimeout(() => {
                            location.reload();
                        }, saveToFiles ? 3000 : 500);
                    } else {
                        alert('⚠️ Aucune donnée valide trouvée dans le fichier.');
                    }
                } else if (Array.isArray(jsonData)) {
                    // Single array - try to detect type or ask user
                    alert('⚠️ Le fichier contient un tableau simple.\n\n💡 Utilisez les boutons d\'importation individuels (Clients, Drivers, etc.)\n   ou utilisez un fichier de sauvegarde complet (Export Tout).');
                } else {
                    alert('⚠️ Format de fichier non reconnu.\n\n💡 Utilisez un fichier de sauvegarde complet créé avec "Export Tout".');
                }
            } catch (error) {
                console.error('Error importing all data:', error);
                alert('❌ Erreur lors de l\'importation:\n\n' + error.message + '\n\nVérifiez que le fichier JSON est valide.');
            }
        };
        
        input.click();
    }
}

// Quick Add Menu functions
function toggleQuickAddMenu() {
    const menu = document.getElementById('quickAddMenu');
    if (menu) {
        if (menu.style.display === 'none' || menu.style.display === '') {
            menu.style.display = 'block';
            // Close menu when clicking outside
            setTimeout(() => {
                document.addEventListener('click', closeQuickAddMenuOnOutsideClick, true);
            }, 100);
        } else {
            closeQuickAddMenu();
        }
    }
}

function closeQuickAddMenu() {
    const menu = document.getElementById('quickAddMenu');
    if (menu) {
        menu.style.display = 'none';
        document.removeEventListener('click', closeQuickAddMenuOnOutsideClick, true);
    }
}

function closeQuickAddMenuOnOutsideClick(e) {
    const menu = document.getElementById('quickAddMenu');
    const button = e.target.closest('.btn-primary');
    if (menu && !menu.contains(e.target) && !button) {
        closeQuickAddMenu();
    }
}

// Show data management modal
function showDataManagement() {
    const modal = document.getElementById('dataManagementModal');
    if (modal) modal.style.display = 'block';
}

// Close data management modal
function closeDataManagement() {
    const modal = document.getElementById('dataManagementModal');
    if (modal) modal.style.display = 'none';
}

// Create sample Excel files with test data
function createSampleExcelFiles() {
    if (typeof XLSX === 'undefined') {
        alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أول مرة.');
        return;
    }

    // Sample data
    const sampleClients = [
        { id: 1, name: "SFI", destination: "SFI Depot", itineraire: "Point A, Point B, Point C" },
        { id: 2, name: "Client B", destination: "Warehouse B", itineraire: "Point D, Point E, Point F" },
        { id: 3, name: "Client Test", destination: "Test Destination", itineraire: "Start, Middle, End" }
    ];

    const sampleDrivers = [
        { id: 1, name: "Ahmed Benali", cin: "AB123456", phone: "0612345678", "vehicle.matricule": "12345-A-56", "vehicle.model": "Mercedes Actros" },
        { id: 2, name: "Mohamed Alami", cin: "MA789012", phone: "0623456789", "vehicle.matricule": "67890-B-12", "vehicle.model": "Volvo FH" },
        { id: 3, name: "Hassan Idrissi", cin: "HI345678", phone: "0634567890", "vehicle.matricule": "11111-C-22", "vehicle.model": "Scania R" }
    ];

    const sampleConvoyeurs = [
        { id: 1, name: "Omar Bensaid", cin: "OB111222", phone: "0645678901", cce: "CCE001" },
        { id: 2, name: "Youssef Amrani", cin: "YA333444", phone: "0656789012", cce: "CCE002" },
        { id: 3, name: "Karim Tazi", cin: "KT555666", phone: "0667890123", cce: "CCE003" }
    ];

    const sampleProducts = [
        { id: 1, name: "Dynamite", unit: "kg" },
        { id: 2, name: "Cordeau détonant", unit: "m" },
        { id: 3, name: "Détonateurs", unit: "unité" },
        { id: 4, name: "Explosif ANFO", unit: "kg" },
        { id: 5, name: "Emulsif", unit: "kg" }
    ];

    const sampleHistory = [];

    // Create and download Excel files
    const files = [
        { name: 'clients.xlsx', data: sampleClients },
        { name: 'drivers.xlsx', data: sampleDrivers },
        { name: 'convoyeurs.xlsx', data: sampleConvoyeurs },
        { name: 'products.xlsx', data: sampleProducts },
        { name: 'history.xlsx', data: sampleHistory }
    ];

    let downloaded = 0;
    files.forEach((file, index) => {
        setTimeout(() => {
            const worksheet = XLSX.utils.json_to_sheet(file.data);
            const workbook = XLSX.utils.book_new();
            const sheetName = file.name.replace('.xlsx', '').replace('.xls', '');
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
            XLSX.writeFile(workbook, file.name);
            downloaded++;
            console.log(`✅ Created ${file.name} with ${file.data.length} rows`);
            
            if (downloaded === files.length) {
                alert(`✅ تم إنشاء ${downloaded} ملف Excel تجريبي!\n\n📁 الملفات:\n${files.map(f => `- ${f.name}`).join('\n')}\n\n💡 Instructions:\n1. انسخ الملفات إلى مجلد data/\n2. استخدم "Import Excel" لتحميل البيانات\n3. أو افتح الملفات في Excel للتعديل`);
            }
        }, index * 500);
    });
}

// ==================== EXCEL EXPORT/IMPORT FUNCTIONS ====================

// Export data to Excel file
function exportDataToExcel(dataType) {
    try {
        // Check if SheetJS is available
        if (typeof XLSX === 'undefined') {
            alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أو استخدم JSON بدلاً من ذلك.');
            return;
        }

        let data = [];
        let filename = '';
        let sheetName = '';
        
        // Load data from localStorage
        const savedClients = localStorage.getItem('clientsData');
        const savedDrivers = localStorage.getItem('driversData');
        const savedConvoyeurs = localStorage.getItem('convoyeursData');
        const savedProducts = localStorage.getItem('productsData');
        const savedHistory = localStorage.getItem('declarationHistory');
        
        if (savedClients) try { clients = JSON.parse(savedClients); } catch(e) {}
        if (savedDrivers) try { drivers = JSON.parse(savedDrivers); } catch(e) {}
        if (savedConvoyeurs) try { convoyeurs = JSON.parse(savedConvoyeurs); } catch(e) {}
        if (savedProducts) try { products = JSON.parse(savedProducts); } catch(e) {}
        if (savedHistory) try { history = JSON.parse(savedHistory); } catch(e) {}
        
        switch(dataType) {
            case 'clients':
                data = clients || [];
                filename = 'clients.xlsx';
                sheetName = 'Clients';
                break;
            case 'drivers':
                data = drivers || [];
                filename = 'drivers.xlsx';
                sheetName = 'Conducteurs';
                break;
            case 'convoyeurs':
                data = convoyeurs || [];
                filename = 'convoyeurs.xlsx';
                sheetName = 'Convoyeurs';
                break;
            case 'products':
                data = products || [];
                filename = 'products.xlsx';
                sheetName = 'Produits';
                break;
            case 'history':
                data = history || [];
                filename = 'history.xlsx';
                sheetName = 'Historique';
                break;
            case 'all':
                // Export all data in multiple sheets
                const workbook = XLSX.utils.book_new();
                
                if (clients && clients.length > 0) {
                    const wsClients = XLSX.utils.json_to_sheet(clients);
                    XLSX.utils.book_append_sheet(workbook, wsClients, 'Clients');
                }
                if (drivers && drivers.length > 0) {
                    const wsDrivers = XLSX.utils.json_to_sheet(drivers);
                    XLSX.utils.book_append_sheet(workbook, wsDrivers, 'Conducteurs');
                }
                if (convoyeurs && convoyeurs.length > 0) {
                    const wsConvoyeurs = XLSX.utils.json_to_sheet(convoyeurs);
                    XLSX.utils.book_append_sheet(workbook, wsConvoyeurs, 'Convoyeurs');
                }
                if (products && products.length > 0) {
                    const wsProducts = XLSX.utils.json_to_sheet(products);
                    XLSX.utils.book_append_sheet(workbook, wsProducts, 'Produits');
                }
                if (history && history.length > 0) {
                    const wsHistory = XLSX.utils.json_to_sheet(history);
                    XLSX.utils.book_append_sheet(workbook, wsHistory, 'Historique');
                }
                
                filename = `backup_all_data_${new Date().toISOString().split('T')[0]}.xlsx`;
                XLSX.writeFile(workbook, filename);
                alert(`✅ Export Excel réussi !\n\n📁 Fichier: ${filename}\n\n✅ Ouvrez le fichier dans Excel pour voir toutes les données.`);
                return;
            default:
                alert('❌ Type de données invalide');
                return;
        }
        
        if (!data || data.length === 0) {
            alert(`⚠️ Aucune donnée à exporter pour ${dataType}.`);
            return;
        }
        
        // Convert to worksheet
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        // Download
        XLSX.writeFile(workbook, filename);
        alert(`✅ Export Excel réussi !\n\n📁 Fichier: ${filename}\n📊 ${data.length} élément(s) exporté(s)\n\n✅ Ouvrez le fichier dans Excel pour voir/modifier les données.`);
        
    } catch (error) {
        console.error('Excel export error:', error);
        alert(`❌ Erreur lors de l'export Excel:\n\n${error.message}`);
    }
}

// Import data from Excel file
function importDataFromExcel(dataType) {
    try {
        // Check if SheetJS is available
        if (typeof XLSX === 'undefined') {
            alert('❌ مكتبة Excel غير متاحة. تأكد من الاتصال بالإنترنت أو استخدم JSON بدلاً من ذلك.');
            return;
        }

        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xls';
        
        input.onchange = async function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            try {
                const data = await file.arrayBuffer();
                const workbook = XLSX.read(data);
                
                let imported = false;
                let count = 0;
                
                if (dataType === 'all') {
                    // Import all sheets
                    workbook.SheetNames.forEach(sheetName => {
                        const worksheet = workbook.Sheets[sheetName];
                        const jsonData = XLSX.utils.sheet_to_json(worksheet);
                        
                        if (sheetName === 'Clients' || sheetName === 'clients') {
                            // Reconstruct client data (convert itineraire string to array)
                            clients = jsonData.map(reconstructClient);
                            localStorage.setItem('clientsData', JSON.stringify(clients));
                            populateClientSelect();
                            setTimeout(() => makeSelectSearchable('clientSelect'), 100);
                            count += clients.length;
                            imported = true;
                        } else if (sheetName === 'Conducteurs' || sheetName === 'drivers') {
                            // Reconstruct nested vehicle object from flattened Excel structure
                            drivers = jsonData.map(driver => {
                                const reconstructed = { ...driver };
                                if (driver['vehicle.matricule'] || driver['vehicle.model']) {
                                    reconstructed.vehicle = {
                                        matricule: driver['vehicle.matricule'] || '',
                                        model: driver['vehicle.model'] || ''
                                    };
                                    delete reconstructed['vehicle.matricule'];
                                    delete reconstructed['vehicle.model'];
                                }
                                return reconstructed;
                            });
                            localStorage.setItem('driversData', JSON.stringify(drivers));
                            populateDriverSelect();
                            setTimeout(() => makeSelectSearchable('driverSelect'), 100);
                            count += drivers.length;
                            imported = true;
                        } else if (sheetName === 'Convoyeurs' || sheetName === 'convoyeurs') {
                            convoyeurs = jsonData;
                            localStorage.setItem('convoyeursData', JSON.stringify(convoyeurs));
                            populateConvoyeurSelect();
                            setTimeout(() => makeSelectSearchable('convoyeurSelect'), 100);
                            count += convoyeurs.length;
                            imported = true;
                        } else if (sheetName === 'Produits' || sheetName === 'products') {
                            products = jsonData;
                            localStorage.setItem('productsData', JSON.stringify(products));
                            populateProductSelects();
                            setTimeout(() => {
                                document.querySelectorAll('.product-select').forEach(select => {
                                    makeSelectSearchable(select.id || null, select);
                                });
                            }, 100);
                            count += products.length;
                            imported = true;
                        } else if (sheetName === 'Historique' || sheetName === 'history') {
                            history = jsonData;
                            localStorage.setItem('declarationHistory', JSON.stringify(history));
                            count += history.length;
                            imported = true;
                        }
                    });
                } else {
                    // Import single sheet
                    const firstSheet = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheet];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    switch(dataType) {
                        case 'clients':
                            // Reconstruct client data (convert itineraire string to array)
                            clients = jsonData.map(reconstructClient);
                            localStorage.setItem('clientsData', JSON.stringify(clients));
                            populateClientSelect();
                            setTimeout(() => makeSelectSearchable('clientSelect'), 100);
                            count = clients.length;
                            imported = true;
                            break;
                        case 'drivers':
                            drivers = jsonData;
                            localStorage.setItem('driversData', JSON.stringify(drivers));
                            populateDriverSelect();
                            setTimeout(() => makeSelectSearchable('driverSelect'), 100);
                            count = drivers.length;
                            imported = true;
                            break;
                        case 'convoyeurs':
                            convoyeurs = jsonData;
                            localStorage.setItem('convoyeursData', JSON.stringify(convoyeurs));
                            populateConvoyeurSelect();
                            setTimeout(() => makeSelectSearchable('convoyeurSelect'), 100);
                            count = convoyeurs.length;
                            imported = true;
                            break;
                        case 'products':
                            products = jsonData;
                            localStorage.setItem('productsData', JSON.stringify(products));
                            populateProductSelects();
                            setTimeout(() => {
                                document.querySelectorAll('.product-select').forEach(select => {
                                    makeSelectSearchable(select.id || null, select);
                                });
                            }, 100);
                            count = products.length;
                            imported = true;
                            break;
                        case 'history':
                            history = jsonData;
                            localStorage.setItem('declarationHistory', JSON.stringify(history));
                            count = history.length;
                            imported = true;
                            break;
                    }
                }
                
                if (imported) {
                    alert(`✅ Import Excel réussi !\n\n📊 ${count} élément(s) importé(s).\n\n✅ Les données ont été chargées avec succès.`);
                    setTimeout(() => {
                        location.reload();
                    }, 500);
                } else {
                    alert('⚠️ Aucune donnée valide trouvée dans le fichier Excel.');
                }
            } catch (error) {
                console.error('Excel import error:', error);
                alert(`❌ Erreur lors de l'import Excel:\n\n${error.message}`);
            }
        };
        
        input.click();
    } catch (error) {
        console.error('Excel import error:', error);
        alert(`❌ Erreur: ${error.message}`);
    }
}
