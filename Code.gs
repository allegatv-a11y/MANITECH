/***** Code.gs - Manitech Sistema Completo *****/

/* ===== CONFIGURAZIONE ===== */
const SS_ID = "1nCKiZLgVaAd_1ndXPoyDRnu7tbJgR2dZbWGHY82UDhI";
const FOLDER_MANUALS_ID = "1YCwG5BDC7Vrkd5tsVl5sISKkcZ9Lsn1";
const FOLDER_MANUALI_DRIVE_ID = "1YCwG5BDC7Vrkd5tsVl5sISKkcZ9Lsn1-";
const LOGO_FILE_ID = "12uc32Cdwyxk9oJ2FlY3DgOIiPriJbuI4";
const APP_TITLE = "Manitech S.r.l.";
const ADMIN_NOTIFY_EMAIL = "info@manitech.it";
const SECRET_JWT = "MANITECH_SECRET_KEY_2025";
const FOLDER_VERIFICHE_ID = "1hWDxw5rI2JikQWgYE0KtD0oGf3GlCgHB";
const FOLDER_GARANZIE_ID  = "1iSsZDIfN9utQyjXOM3sCTMt7ThWVAZw9";

/* ===== UTILITY ===== */
function _getSheet(name){ return SpreadsheetApp.openById(SS_ID).getSheetByName(name); }
function _now(){ return new Date().toISOString(); }
function _genId(prefix='id'){ return prefix + "-" + Utilities.getUuid(); }
function _json(obj){ return typeof obj === 'string' ? obj : JSON.stringify(obj || {}); }
function _hash(p){ 
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p + SECRET_JWT); 
  return d.map(b=>('0'+(b&0xFF).toString(16)).slice(-2)).join(''); 
}

function hashPassword(password){
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + SECRET_JWT); 
  return d.map(b=>('0'+(b&0xFF).toString(16)).slice(-2)).join('');
}
/* ===== AUTENTICAZIONE ===== */
function loginUser(email, password){
  try {
    const sh = _getSheet('Utenti');
    const rows = sh.getDataRange().getValues();
    
    const hashedPwd = _hash(password);
    
    for(let i = 1; i < rows.length; i++){
      if(rows[i][0].toLowerCase() === email.toLowerCase()){
        if(rows[i][4] === hashedPwd){
          const token = _genId('token');
          sh.getRange(i+1, 6).setValue(token);
          
          return { 
            ok: true, 
            token: token,
            email: email,
            name: rows[i][1],
            role: rows[i][2]
          };
        } else {
          return { ok: false, error: 'Password errata' };
        }
      }
    }
    
    return { ok: false, error: 'Utente non trovato' };
    
  } catch(e) {
    Logger.log('ERRORE loginUser: ' + e);
    return { ok: false, error: 'Errore di sistema' };
  }
}


function validateToken(token){
  try {
    if (!token) return null;
    const rows = _getSheet('Utenti').getDataRange().getValues();
    for (let i=1; i<rows.length; i++) {
      if ((rows[i][5]||'') === token) return rows[i][0];
    }
    return null;
  } catch(e) {
    Logger.log('ERRORE validateToken: ' + e);
    return null;
  }
}

function getRoleByEmail(email){
  const sh = _getSheet('Utenti');
  const data = sh.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++){
    if(data[i][0].toLowerCase() === email.toLowerCase()){
      return data[i][2];
    }
  }
  return null;
}

function getNameByEmail(email){
  const rows = _getSheet('Utenti').getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if ((rows[i][0]||'').toLowerCase() === email.toLowerCase()) {
      return rows[i][1] || '';
    }
  }
  return '';
}

function loginStatus(token){ 
  const user = validateToken(token); 
  if (!user) return { ok: false }; 
  return { 
    ok: true, 
    email: user, 
    role: getRoleByEmail(user), 
    name: getNameByEmail(user) 
  }; 
}
/* ===== GESTIONE UTENTI (ADMIN) ===== */
function adminListUtenti(token){
  const user = validateToken(token);
  if(!user || getRoleByEmail(user) !== 'admin') throw 'Only admin';
  
  const sh = _getSheet('Utenti');
  const data = sh.getDataRange().getValues();
  
  const Utenti = [];
  for(let i = 1; i < data.length; i++){
    Utenti.push({
      email: data[i][0],
      name: data[i][1],
      role: data[i][2],
      company: data[i][3] || '',
      createdAt: data[i][6] ? Utilities.formatDate(new Date(data[i][6]), 'Europe/Rome', 'dd/MM/yyyy HH:mm') : 'N/A'
    });
  }
  
  Logger.log('✅ Lista utenti: ' + Utenti.length);
  return Utenti;
}

/**
 * Crea un nuovo utente (solo admin)
 */
function adminCreateUser(token, name, email, password, role, company){
  const user = validateToken(token);
  if(!user || getRoleByEmail(user) !== 'admin') {
    return { ok: false, error: 'Solo gli admin possono creare utenti' };
  }
  
  const sh = _getSheet('Utenti');
  const data = sh.getDataRange().getValues();
  
  // Controlla se email esiste già
  for(let i = 1; i < data.length; i++){
    if(data[i][0] && data[i][0].toLowerCase() === email.toLowerCase()){
      return { ok: false, error: 'Email già registrata' };
    }
  }
  
  // Validazione
  if(!name || name.trim().length < 2){
    return { ok: false, error: 'Nome non valido (minimo 2 caratteri)' };
  }
  
  if(!email || !email.includes('@')){
    return { ok: false, error: 'Email non valida' };
  }
  
  if(!password || password.length < 6){
    return { ok: false, error: 'Password troppo corta (minimo 6 caratteri)' };
  }
  
  if(!role || !['admin', 'tecnico', 'cliente'].includes(role)){
    return { ok: false, error: 'Ruolo non valido' };
  }
  
  const hashedPassword = hashPassword(password);
  
  // ⭐ GENERA IL TOKEN AL MOMENTO DELLA CREAZIONE
  const sessionToken = 'token-' + Utilities.getUuid();
  
  // ⭐ ORDINE CORRETTO: Email | Nome | Ruolo | Company | PasswordHash | SessionToken | CreatedAt
  sh.appendRow([
    email.toLowerCase(),      // Col A: Email
    name,                     // Col B: Nome
    role,                     // Col C: Ruolo
    company || '',            // Col D: Company
    hashedPassword,           // Col E: PasswordHash
    sessionToken,             // Col F: SessionToken ← ⭐ GENERATO QUI!
    _now()                    // Col G: CreatedAt
  ]);
  
  Logger.log('✅ Utente creato: ' + email + ' (' + role + ') - Token: ' + sessionToken);
  return { 
    ok: true,
    message: 'Utente creato con successo',
    user: {
      name: name,
      email: email,
      role: role,
      company: company || '',
      token: sessionToken
    }
  };
}




/**
 * Elimina un utente (solo admin)
 */
function adminDeleteUser(token, emailToDelete){
  const user = validateToken(token);
  if(!user || getRoleByEmail(user) !== 'admin') {
    return { ok: false, error: 'Solo gli admin possono eliminare utenti' };
  }
  
  if(user.toLowerCase() === emailToDelete.toLowerCase()){
    return { ok: false, error: 'Non puoi eliminare il tuo account' };
  }
  
  const sh = _getSheet('Utenti');
  const data = sh.getDataRange().getValues();
  
  // ⭐ CORREZIONE: Email è nella colonna B (indice 1)
  for(let i = 1; i < data.length; i++){
    if(data[i][1] && data[i][1].toLowerCase() === emailToDelete.toLowerCase()){
      sh.deleteRow(i + 1);
      Logger.log('✅ Utente eliminato: ' + emailToDelete);
      return { ok: true, message: 'Utente eliminato con successo' };
    }
  }
  
  return { ok: false, error: 'Utente non trovato' };
}


/**
 * Aggiorna un utente (solo admin)
 */
function adminUpdateUser(token, email, updates){
  const user = validateToken(token);
  if(!user || getRoleByEmail(user) !== 'admin') {
    return { ok: false, error: 'Solo gli admin possono modificare utenti' };
  }
  
  const sh = _getSheet('Utenti');
  const data = sh.getDataRange().getValues();
  
  // ⭐ CORREZIONE: Email è nella colonna B (indice 1)
  for(let i = 1; i < data.length; i++){
    if(data[i][1] && data[i][1].toLowerCase() === email.toLowerCase()){
      // Aggiorna solo i campi forniti
      if(updates.name) sh.getRange(i + 1, 1).setValue(updates.name);        // Col A
      if(updates.role) sh.getRange(i + 1, 4).setValue(updates.role);        // Col D
      if(updates.company !== undefined) sh.getRange(i + 1, 5).setValue(updates.company); // Col E
      if(updates.password) sh.getRange(i + 1, 3).setValue(hashPassword(updates.password)); // Col C
      
      Logger.log('✅ Utente aggiornato: ' + email);
      return { ok: true, message: 'Utente aggiornato con successo' };
    }
  }
  
  return { ok: false, error: 'Utente non trovato' };
}

/* ===== MACCHINE ===== */
function getActiveRentals(){
  try {
    const sh = _getSheet('Rentals');
    if (!sh) return [];

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const rows = sh.getRange(2, 1, lastRow-1, 8).getValues();
    const now = new Date();
    const actives = [];

    for (let i=0; i<rows.length; i++){
      const status = rows[i][5];
      if (!rows[i][4]) continue;

      const endDate = new Date(rows[i][4]);

      if ((status === 'active' || status === 'approved') && endDate > now) {
        actives.push({
          id: rows[i][0],
          machineId: rows[i][1],
          requestedBy: rows[i][2],
          startDate: rows[i][3],
          endDate: rows[i][4],
          status: status
        });
      }
    }

    return actives;

  } catch(e) {
    Logger.log('ERRORE getActiveRentals: ' + e);
    return [];
  }
}

/**
 * Lista macchine con ricerca e paginazione ottimizzata per grandi dataset
 * @param {string} token Token autenticazione
 * @param {number} page Numero pagina (0-based)
 * @param {number} pageSize Elementi per pagina
 * @param {string} searchTerm Termine di ricerca (opzionale)
 * @returns {Object} Risultato con macchine, totale e info paginazione
 */
function listMachinesPaged(token, page, pageSize, searchTerm){
  try {
    const email = validateToken(token); 
    if (!email) return { machines: [], total: 0, page: 0, totalPages: 0 };

    const role = getRoleByEmail(email);
    const sh = _getSheet('Macchine'); 
    if (!sh) return { machines: [], total: 0, page: 0, totalPages: 0 };

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return { machines: [], total: 0, page: 0, totalPages: 0 };

    // ⭐ OTTIMIZZAZIONE: Leggi solo le colonne necessarie
    const rows = sh.getRange(2, 1, lastRow-1, 9).getValues();
    const rentals = getActiveRentals() || [];

    let allMachines = [];

    // ⭐ FILTRO RICERCA (se presente)
    const searchLower = (searchTerm || '').toLowerCase().trim();

    for (let i = 0; i < rows.length; i++){
      const machineId = rows[i][0];
      if (!machineId) continue;

      const model = String(rows[i][2] || '');
      const serial = String(rows[i][1] || '');
      
      // ⭐ Se c'è un termine di ricerca, filtra prima di aggiungere
      if (searchLower && 
          !machineId.toString().toLowerCase().includes(searchLower) &&
          !model.toLowerCase().includes(searchLower) &&
          !serial.toLowerCase().includes(searchLower)) {
        continue; // Salta questa macchina se non corrisponde alla ricerca
      }

      const rental = rentals.find(r => r.machineId === machineId);

      const rec = { 
        id: String(machineId),
        serial: serial,
        model: model,
        ownerType: String(rows[i][3] || 'manitech'),
        cliente: rows[i][4] || '',
        noleggiabile: rows[i][8] === true ? 'Sì' : 'No',
        status: String(rows[i][5] || 'available')
      };

      if (rental && rental.status === 'active') {
        rec.status = 'rented';
      }

      // Admin e Tecnici vedono tutte le macchine
      if (role === 'admin' || role === 'tecnico') {
        allMachines.push(rec);
      }
    }

    const total = allMachines.length;
    const totalPages = Math.ceil(total / pageSize);
    const start = page * pageSize;
    const end = start + pageSize;
    const pagedMachines = allMachines.slice(start, end);

    return {
      machines: pagedMachines,
      total: total,
      page: page,
      totalPages: totalPages
    };

  } catch(e) {
    Logger.log('❌ ERRORE listMachinesPaged: ' + e);
    return { machines: [], total: 0, page: 0, totalPages: 0 };
  }
}



function formatDateForDisplay(dateValue){
  if (!dateValue) return 'N/A';
  try {
    const d = new Date(dateValue);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch(e) {
    return String(dateValue);
  }
}

function extractCategoria(model){
  if (!model || model === null || model === undefined) return 'Altro';
  
  const modelStr = String(model).trim();
  if (modelStr === '' || modelStr === 'null' || modelStr === 'undefined') return 'Altro';
  
  const match = modelStr.match(/^([A-Z]+\d+)/i);
  if (match) return match[1].toUpperCase();
  
  return 'Altro';
}

function listAvailableMachinesForRental(token){
  try {
    const email = validateToken(token); 
    if (!email) return [];
    
    const sh = _getSheet('Macchine'); 
    if (!sh) return [];
    
    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];
    
    const rows = sh.getRange(2, 1, lastRow-1, 10).getValues();
    const rentals = getActiveRentals() || [];
    
    const available = [];
    
    for (let i=0; i<rows.length; i++){
      const machineId = rows[i][0];
      if (!machineId) continue;
      
      const noleggiabile = rows[i][8];
      const ownerType = rows[i][3];
      const status = rows[i][5];
      const modelValue = rows[i][2];
      
      const isNoleggiabile = (noleggiabile === true || 
                              noleggiabile === 'TRUE' || 
                              noleggiabile === 'true');
      
      if (ownerType === 'manitech' && isNoleggiabile && status === 'available') {
        const hasActiveRental = rentals.some(function(r){ 
          return r.machineId === machineId; 
        });
        
        if (!hasActiveRental) {
          available.push({
            id: String(machineId),
            serial: String(rows[i][1] || ''),
            model: String(modelValue || 'N/A'),
            categoria: extractCategoria(modelValue),
            status: 'available'
          });
        }
      }
    }
    
    return available;
    
  } catch(e) {
    Logger.log('ERRORE listAvailableMachinesForRental: ' + e);
    return [];
  }
}

function updateMachineStatus(machineId, newStatus){
  const sh = _getSheet('Macchine');
  const rows = sh.getDataRange().getValues();

  for (let i=1; i<rows.length; i++){
    if (rows[i][0] === machineId) {
      sh.getRange(i+1, 6).setValue(newStatus);
      return true;
    }
  }
  return false;
}
/* ===== NOLEGGI ===== */
function requestRentalAdvanced(token, machineId, startDate, endDate, notes){
  try {
    const user = validateToken(token); 
    if (!user) throw 'Not authenticated';

    const available = checkRentalAvailability(machineId, startDate, endDate);

    if (!available) {
      return { ok: false, error: 'Macchina non disponibile nelle date selezionate' };
    }

    const sh = _getSheet('Macchine');
    const rows = sh.getDataRange().getValues();
    let machineFound = false;

    for (let i=1; i<rows.length; i++){
      if (rows[i][0] === machineId) {
        machineFound = true;

        if (rows[i][8] !== true) {
          return { ok: false, error: 'Questa macchina non è disponibile per il noleggio' };
        }

        if (rows[i][3] !== 'manitech') {
          return { ok: false, error: 'Solo le macchine Manitech sono noleggiabili' };
        }

        break;
      }
    }

    if (!machineFound) {
      return { ok: false, error: 'Macchina non trovata' };
    }

    const shRental = _getSheet('Rentals');
    const id = _genId('rental');

    shRental.appendRow([
      id,
      machineId,
      user,
      startDate,
      endDate,
      'requested',
      _now(),
      notes || ''
    ]);

    try { 
      MailApp.sendEmail(
        ADMIN_NOTIFY_EMAIL, 
        'Nuova richiesta noleggio ' + id, 
        'Utente: ' + user + '\nMacchina: ' + machineId + '\nDal: ' + startDate + '\nAl: ' + endDate
      ); 
    } catch(e){}

    return { ok: true, id: id };

  } catch(e) {
    Logger.log('ERRORE requestRentalAdvanced: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}

function checkRentalAvailability(machineId, startDate, endDate){
  const sh = _getSheet('Rentals'); 
  const rows = sh.getDataRange().getValues();
  const s = new Date(startDate), e = new Date(endDate);
  
  for (let i=1; i<rows.length; i++){
    if (rows[i][1] === machineId && (rows[i][5] === 'approved' || rows[i][5] === 'active')){
      const rs = new Date(rows[i][3]), re = new Date(rows[i][4]);
      if (!(e < rs || s > re)) return false;
    }
  }
  return true;
}

function listRentalRequests(token){
  try {
    const user = validateToken(token);
    if (!user) throw new Error('Not authenticated');

    const sh = _getSheet('Rentals');
    if (!sh) return [];

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const rows = sh.getRange(2, 1, lastRow-1, 8).getValues();
    const out = [];
    const role = getRoleByEmail(user);

    for (let i=0; i<rows.length; i++){
      if (!rows[i][0]) continue;

      const rental = {
        id: rows[i][0],
        machineId: rows[i][1],
        requestedBy: rows[i][2],
        startDate: rows[i][3],
        endDate: rows[i][4],
        status: rows[i][5] || 'requested',
        createdAt: rows[i][6],
        notes: rows[i][7] || ''
      };

      if (role === 'admin' || rental.requestedBy === user) {
        out.push(rental);
      }
    }

    return out;

  } catch(e) {
    Logger.log('ERRORE listRentalRequests: ' + e);
    return [];
  }
}

function adminApproveRentalAdvanced(token, rentalId){
  try {
    const user = validateToken(token);
    if (!user || getRoleByEmail(user) !== 'admin') throw 'Only admin';

    const sh = _getSheet('Rentals');
    const rows = sh.getDataRange().getValues();

    for (let i=1; i<rows.length; i++){
      if (rows[i][0] === rentalId) {
        const machineId = rows[i][1];
        const clientEmail = rows[i][2];

        sh.getRange(i+1, 6).setValue('active');
        updateMachineStatus(machineId, 'rented');

        try {
          MailApp.sendEmail(
            clientEmail, 
            'Noleggio approvato - ' + rentalId, 
            'Il tuo noleggio è stato approvato.'
          );
        } catch(e){}

        return { ok: true };
      }
    }

    throw 'Noleggio non trovato';

  } catch(e) {
    Logger.log('ERRORE adminApproveRentalAdvanced: ' + e);
    throw e;
  }
}

function adminRejectRental(token, rentalId){
  const user = validateToken(token);
  if (!user || getRoleByEmail(user) !== 'admin') throw 'Only admin';

  const sh = _getSheet('Rentals');
  const rows = sh.getDataRange().getValues();

  for (let i=1; i<rows.length; i++){
    if (rows[i][0] === rentalId) {
      sh.getRange(i+1, 6).setValue('rejected');
      return { ok: true };
    }
  }

  throw 'Noleggio non trovato';
}

function adminEndRentalAdvanced(token, rentalId){
  try {
    const user = validateToken(token);
    if (!user || getRoleByEmail(user) !== 'admin') throw 'Only admin';

    const sh = _getSheet('Rentals');
    const rows = sh.getDataRange().getValues();

    for (let i=1; i<rows.length; i++){
      if (rows[i][0] === rentalId) {
        const machineId = rows[i][1];

        sh.getRange(i+1, 6).setValue('completed');
        updateMachineStatus(machineId, 'available');

        return { ok: true };
      }
    }

    throw 'Noleggio non trovato';

  } catch(e) {
    Logger.log('ERRORE adminEndRentalAdvanced: ' + e);
    throw e;
  }
}
/* ===== MANUALI ===== */
function formatBytes(bytes){
  if (!bytes) return '0 KB';
  const k = 1024;
  const mb = bytes / (k * k);
  if (mb > 1) return Math.round(mb * 10) / 10 + ' MB';
  return Math.round(bytes / k) + ' KB';
}

function refreshManualsCache(){
  const folder = DriveApp.getFolderById(FOLDER_MANUALI_DRIVE_ID);
  const cacheSheet = _getSheet('ManualsCache');
  
  cacheSheet.clear();
  cacheSheet.appendRow(['FileID', 'Title', 'Path', 'Size']);
  
  function scan(currentFolder, path) {
    const files = currentFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getMimeType() === 'application/pdf') {
        cacheSheet.appendRow([
          f.getId(),
          f.getName(),
          path,
          f.getSize()
        ]);
      }
    }
    
    const subfolders = currentFolder.getFolders();
    while (subfolders.hasNext()) {
      const sub = subfolders.next();
      scan(sub, path + ' › ' + sub.getName());
    }
  }
  
  scan(folder, 'macchine industriali');
  Logger.log('✅ Cache aggiornata');
}

function listManualsFromDrive(token){
  const user = validateToken(token); 
  if (!user) throw 'Not authenticated';
  
  const cacheSheet = _getSheet('ManualsCache');
  const data = cacheSheet.getDataRange().getValues();
  
  if (data.length <= 1) return [];
  
  const out = [];
  for (let i = 1; i < data.length; i++) {
    out.push({
      fileId: data[i][0],
      title: data[i][1],
      desc: data[i][2],
      size: formatBytes(data[i][3])
    });
  }
  
  return out;
}


/* ===== FRONTEND ===== */
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('index');
  return tpl.evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function updatePasswordForAlessandro(){
  const sh = _getSheet('Utenti');
  const rows = sh.getDataRange().getValues();
  
  for(let i = 1; i < rows.length; i++){
    if(rows[i][0] === 'ricambi@manitech.it'){
      const newHash = '4qHZJrF1c7yoIF2936XdLNQwzZmsxfPWJSoIxp0Ya3E=';
      sh.getRange(i+1, 5).setValue(newHash); // Colonna E
      Logger.log('✅ Password aggiornata per ' + rows[i][0]);
      return;
    }
  }
  
  Logger.log('❌ Utente non trovato');
}
function debugLogin(){
  const sh = _getSheet('Utenti');
  const rows = sh.getDataRange().getValues();
  
  Logger.log('=== DEBUG LOGIN ===');
  Logger.log('Foglio: ' + sh.getName());
  Logger.log('Righe totali: ' + rows.length);
  
  // Mostra prima riga utente
  if(rows.length > 1){
    Logger.log('\n--- UTENTE RIGA 2 ---');
    Logger.log('Email (col A): ' + rows[1][0]);
    Logger.log('Nome (col B): ' + rows[1][1]);
    Logger.log('Ruolo (col C): ' + rows[1][2]);
    Logger.log('Company (col D): ' + rows[1][3]);
    Logger.log('PasswordHash (col E): ' + rows[1][4]);
    Logger.log('SessionToken (col F): ' + rows[1][5]);
  }
  
  // Test hash
  const testPassword = 'Manitech01';
  const hash1 = hashPassword(testPassword);
  
  Logger.log('\n--- TEST HASH ---');
  Logger.log('Password: ' + testPassword);
  Logger.log('Hash generato: ' + hash1);
  Logger.log('Hash nel foglio: ' + rows[1][4]);
  Logger.log('MATCH: ' + (hash1 === rows[1][4]));
  
  return {
    storedHash: rows[1][4],
    generatedHash: hash1,
    match: hash1 === rows[1][4]
  };
}

/* ===== TICKETS E ASSISTENZA ===== */

function createTicket(token, machineId, clientName, description, priority){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';

    const role = getRoleByEmail(user);
    if (role !== 'tecnico' && role !== 'admin') {
      return { ok: false, error: 'Solo tecnici e admin possono creare ticket' };
    }

    // Verifica che la macchina esista
    const shMachine = _getSheet('Macchine');
    const machinesData = shMachine.getDataRange().getValues();
    let machineFound = null;

    for (let i=1; i<machinesData.length; i++){
      if (machinesData[i][0] === machineId) {
        machineFound = {
          id: machinesData[i][0],
          serial: machinesData[i][1],
          model: machinesData[i][2]
        };
        break;
      }
    }

    if (!machineFound) {
      return { ok: false, error: 'Macchina non trovata' };
    }

    const sh = _getSheet('Tickets');
    const ticketId = _genId('ticket');

    sh.appendRow([
      ticketId,
      machineId,
      machineFound.serial,
      machineFound.model,
      clientName,
      description,
      priority || 'media',
      'aperto',
      user,
      _now(),
      '',
      ''
    ]);

    // Notifica email all'admin
    try {
      MailApp.sendEmail(
        ADMIN_NOTIFY_EMAIL,
        '🔧 Nuovo Ticket Assistenza - ' + ticketId,
        'NUOVO TICKET DI ASSISTENZA\n\n' +
        'ID Ticket: ' + ticketId + '\n' +
        'Cliente: ' + clientName + '\n' +
        'Macchina: ' + machineFound.model + ' (Serial: ' + machineFound.serial + ')\n' +
        'Descrizione: ' + description + '\n' +
        'Priorità: ' + priority + '\n' +
        'Creato da: ' + user + '\n' +
        'Data: ' + _now()
      );
    } catch(e){
      Logger.log('Errore invio email: ' + e);
    }

    Logger.log('✅ Ticket creato: ' + ticketId);
    return { ok: true, ticketId: ticketId };

  } catch(e) {
    Logger.log('❌ ERRORE createTicket: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}


function listTickets(token, filterStatus){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';

    const role = getRoleByEmail(user);
    const sh = _getSheet('Tickets');

    if (!sh) return [];

    const lastRow = sh.getLastRow();
    if (lastRow <= 1) return [];

    const rows = sh.getRange(2, 1, lastRow-1, 12).getValues();
    const tickets = [];

    for (let i=0; i<rows.length; i++){
      if (!rows[i][0]) continue;

      const ticket = {
        id: rows[i][0],
        machineId: rows[i][1],
        machineSerial: rows[i][2],
        machineModel: rows[i][3],
        clientName: rows[i][4],
        description: rows[i][5],
        priority: rows[i][6],
        status: rows[i][7],
        createdBy: rows[i][8],
        createdAt: rows[i][9],
        closedAt: rows[i][10] || null,
        notes: rows[i][11] || ''
      };

      // Filtra per ruolo
      if (role === 'tecnico' && ticket.createdBy !== user) {
        continue;
      }

      // Filtra per stato se richiesto
      if (filterStatus && ticket.status !== filterStatus) {
        continue;
      }

      tickets.push(ticket);
    }

    return tickets;

  } catch(e) {
    Logger.log('❌ ERRORE listTickets: ' + e);
    return [];
  }
}


function updateTicketStatus(token, ticketId, newStatus, notes){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';

    const role = getRoleByEmail(user);
    if (role !== 'admin') {
      return { ok: false, error: 'Solo admin può aggiornare i ticket' };
    }

    const sh = _getSheet('Tickets');
    const rows = sh.getDataRange().getValues();

    for (let i=1; i<rows.length; i++){
      if (rows[i][0] === ticketId) {
        sh.getRange(i+1, 8).setValue(newStatus);

        if (notes) {
          sh.getRange(i+1, 12).setValue(notes);
        }

        if (newStatus === 'chiuso' || newStatus === 'risolto') {
          sh.getRange(i+1, 11).setValue(_now());
        }

        Logger.log('✅ Ticket aggiornato: ' + ticketId + ' -> ' + newStatus);
        return { ok: true };
      }
    }

    return { ok: false, error: 'Ticket non trovato' };

  } catch(e) {
    Logger.log('❌ ERRORE updateTicketStatus: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}


function getTicketStats(token){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';

    const role = getRoleByEmail(user);
    if (role !== 'admin') {
      return { ok: false, error: 'Solo admin può vedere le statistiche' };
    }

    const sh = _getSheet('Tickets');
    const lastRow = sh.getLastRow();

    if (lastRow <= 1) {
      return {
        totale: 0,
        aperti: 0,
        inLavorazione: 0,
        chiusi: 0,
        urgenti: 0
      };
    }

    const rows = sh.getRange(2, 1, lastRow-1, 12).getValues();

    const stats = {
      totale: rows.length,
      aperti: 0,
      inLavorazione: 0,
      chiusi: 0,
      urgenti: 0
    };

    for (let i=0; i<rows.length; i++){
      const status = rows[i][7];
      const priority = rows[i][6];

      if (status === 'aperto') stats.aperti++;
      if (status === 'in_lavorazione') stats.inLavorazione++;
      if (status === 'chiuso' || status === 'risolto') stats.chiusi++;
      if (priority === 'alta' || priority === 'urgente') stats.urgenti++;
    }

    return stats;

  } catch(e) {
    Logger.log('❌ ERRORE getTicketStats: ' + e);
    return null;
  }
}
function includeHtml(name, token){
  const user = validateToken(token);
  if(!user) return '<div class="alert alert-danger">Non autorizzato</div>';
  
  // Restituisce HTML vuoto perché usi sezioni dinamiche
  return '';
}
/* ===== SISTEMA TICKETS ASSISTENZA ===== */

/**
 * Crea un nuovo ticket di assistenza
 * Solo tecnici e admin possono creare ticket
 */
function createTicket(token, machineId, description){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';
    
    const role = getRoleByEmail(user);
    if (role !== 'tecnico' && role !== 'admin') {
      return { ok: false, error: 'Solo tecnici e admin possono creare ticket' };
    }
    
    // Recupera dati macchina
    const shMachines = _getSheet('Macchine');
    const machinesData = shMachines.getDataRange().getValues();
    let machineInfo = null;
    
    for (let i = 1; i < machinesData.length; i++) {
      if (machinesData[i][0] === machineId) {
        machineInfo = {
          id: machinesData[i][0],
          serial: machinesData[i][1],
          model: machinesData[i][2],
          cliente: machinesData[i][4] || 'N/A'  // Colonna E = Cliente
        };
        break;
      }
    }
    
    if (!machineInfo) {
      return { ok: false, error: 'Macchina non trovata' };
    }
    
    // Crea ticket
    const shTickets = _getSheet('Tickets');
    const ticketId = _genId('ticket');
    
    shTickets.appendRow([
      ticketId,                    // A: ID
      machineInfo.id,              // B: MachineID
      machineInfo.serial,          // C: MachineSerial
      machineInfo.model,           // D: MachineModel
      machineInfo.cliente,         // E: ClientName
      description,                 // F: Description
      'aperto',                    // G: Status
      user,                        // H: CreatedBy (email tecnico)
      _now(),                      // I: CreatedAt
      '',                          // J: ClosedAt
      ''                           // K: Notes
    ]);
    
    // Notifica admin via email
    try {
      MailApp.sendEmail(
        ADMIN_NOTIFY_EMAIL,
        '🔧 Nuovo Ticket Assistenza - ' + ticketId,
        'NUOVO TICKET DI ASSISTENZA\n\n' +
        'ID: ' + ticketId + '\n' +
        'Macchina: ' + machineInfo.model + ' (' + machineInfo.serial + ')\n' +
        'Cliente: ' + machineInfo.cliente + '\n' +
        'Difetto: ' + description + '\n' +
        'Creato da: ' + user + '\n' +
        'Data: ' + _now()
      );
    } catch(e) {
      Logger.log('Errore invio email: ' + e);
    }
    
    Logger.log('✅ Ticket creato: ' + ticketId);
    return { ok: true, ticketId: ticketId };
    
  } catch(e) {
    Logger.log('❌ ERRORE createTicket: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}


/**
 * Lista tickets con filtro per ruolo
 * - Tecnico: vede solo i PROPRI ticket
 * - Admin: vede TUTTI i ticket
 */
function listTickets(token){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';
    
    const role = getRoleByEmail(user);
    const shTickets = _getSheet('Tickets');
    
    if (!shTickets) return [];
    
    const lastRow = shTickets.getLastRow();
    if (lastRow <= 1) return [];
    
    const rows = shTickets.getRange(2, 1, lastRow - 1, 11).getValues();
    const tickets = [];
    
    for (let i = 0; i < rows.length; i++) {
      if (!rows[i][0]) continue;
      
      const ticket = {
        id: rows[i][0],
        machineId: rows[i][1],
        machineSerial: rows[i][2],
        machineModel: rows[i][3],
        clientName: rows[i][4],
        description: rows[i][5],
        status: rows[i][6],
        createdBy: rows[i][7],
        createdAt: rows[i][8],
        closedAt: rows[i][9] || null,
        notes: rows[i][10] || ''
      };
      
      // Filtro per ruolo
      if (role === 'tecnico') {
        // Tecnico vede solo i propri ticket
        if (ticket.createdBy === user) {
          tickets.push(ticket);
        }
      } else if (role === 'admin') {
        // Admin vede tutti i ticket
        tickets.push(ticket);
      }
    }
    
    Logger.log('✅ Tickets trovati: ' + tickets.length + ' (ruolo: ' + role + ')');
    return tickets;
    
  } catch(e) {
    Logger.log('❌ ERRORE listTickets: ' + e);
    return [];
  }
}


/**
 * Aggiorna stato ticket
 * Solo ADMIN può modificare lo stato
 */
function updateTicketStatus(token, ticketId, newStatus, notes){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';
    
    const role = getRoleByEmail(user);
    if (role !== 'admin') {
      return { ok: false, error: 'Solo admin può modificare i ticket' };
    }
    
    // Verifica stato valido
    const validStates = ['aperto', 'in_lavorazione', 'chiuso'];
    if (!validStates.includes(newStatus)) {
      return { ok: false, error: 'Stato non valido' };
    }
    
    const shTickets = _getSheet('Tickets');
    const rows = shTickets.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === ticketId) {
        // Aggiorna stato
        shTickets.getRange(i + 1, 7).setValue(newStatus);  // Colonna G: Status
        
        // Aggiorna note se presenti
        if (notes) {
          shTickets.getRange(i + 1, 11).setValue(notes);  // Colonna K: Notes
        }
        
        // Se chiuso, aggiungi data chiusura
        if (newStatus === 'chiuso') {
          shTickets.getRange(i + 1, 10).setValue(_now());  // Colonna J: ClosedAt
        }
        
        Logger.log('✅ Ticket aggiornato: ' + ticketId + ' -> ' + newStatus);
        return { ok: true };
      }
    }
    
    return { ok: false, error: 'Ticket non trovato' };
    
  } catch(e) {
    Logger.log('❌ ERRORE updateTicketStatus: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}


/**
 * Statistiche tickets per admin
 */
function getTicketStats(token){
  try {
    const user = validateToken(token);
    if (!user) throw 'Not authenticated';
    
    const role = getRoleByEmail(user);
    if (role !== 'admin') {
      return { ok: false, error: 'Solo admin può vedere le statistiche' };
    }
    
    const shTickets = _getSheet('Tickets');
    const lastRow = shTickets.getLastRow();
    
    if (lastRow <= 1) {
      return {
        totale: 0,
        aperti: 0,
        inLavorazione: 0,
        chiusi: 0
      };
    }
    
    const rows = shTickets.getRange(2, 1, lastRow - 1, 7).getValues();
    
    const stats = {
      totale: rows.length,
      aperti: 0,
      inLavorazione: 0,
      chiusi: 0
    };
    
    for (let i = 0; i < rows.length; i++) {
      const status = rows[i][6];  // Colonna G: Status
      
      if (status === 'aperto') stats.aperti++;
      else if (status === 'in_lavorazione') stats.inLavorazione++;
      else if (status === 'chiuso') stats.chiusi++;
    }
    
    return stats;
    
  } catch(e) {
    Logger.log('❌ ERRORE getTicketStats: ' + e);
    return null;
  }
}
/**
 * Recupera dati di una specifica macchina per ID
 */
function getMachineById(token, machineId){
  const user = validateToken(token);
  if (!user) throw 'Not authenticated';
  
  const sh = _getSheet('Macchine');
  const rows = sh.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === machineId) {
      return {
        id: rows[i][0],
        serial: rows[i][1],
        model: rows[i][2],
        cliente: rows[i][4] || 'N/A'  // Colonna E = Cliente
      };
    }
  }
  
  Logger.log('⚠️ Macchina non trovata: ' + machineId);
  return null;
}
function logout() {
  // Questa funzione deve creare un template e poi valutarlo per ottenere l'HTML
  const tpl = HtmlService.createTemplateFromFile('index');
  return tpl.evaluate().getContent();
}
/***** GESTIONE PDF VERIFICHE E GARANZIE *****/

/**
 * Genera un PDF a partire da un template HTML e lo salva su Drive
 * @param {string} templateName Nome del template ('verifiche' o 'garanzia')
 * @param {Object} data Dati da inserire nel PDF
 * @param {string} folderId ID della cartella Drive dove salvare il PDF
 * @returns {Object} Oggetto con URL del PDF generato
 */
function generatePdfFromTemplate(templateName, data, folderId) {
  try {
    // Prepara i metadati
    const meta = {
      id: _genId('doc'),
      company: APP_TITLE,
      createdAt: Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy HH:mm'),
      logoUrl: 'https://drive.google.com/uc?id=' + LOGO_FILE_ID
    };

    // Seleziona il template corretto
    let template;
    if (templateName === 'verifiche') {
      template = HtmlService.createTemplateFromFile('pdf_template_verifiche');
    } else if (templateName === 'garanzia') {
      template = HtmlService.createTemplateFromFile('pdf_template_garanzia');
    } else {
      throw 'Template non valido';
    }

    // Passa i dati al template
    template.meta = meta;
    template.data = data;

    // Genera l'HTML dal template
    const htmlContent = template.evaluate().getContent();

    // Converti HTML in PDF
    const blob = Utilities.newBlob(htmlContent, 'text/html', 'temp.html')
      .getAs('application/pdf')
      .setName(templateName + '_' + data.machineId + '_' + meta.id + '.pdf');

    // Salva il PDF nella cartella Drive
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);

    Logger.log('✅ PDF generato: ' + file.getName());

    return {
      ok: true,
      pdfUrl: file.getUrl(),
      fileId: file.getId(),
      fileName: file.getName()
    };

  } catch(e) {
    Logger.log('❌ ERRORE generatePdfFromTemplate: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}
/**
 * Crea una verifica di sicurezza per una macchina
 * @param {string} token Token di autenticazione
 * @param {string} machineId ID della macchina
 * @param {Object} formData Dati del form compilato
 * @returns {Object} Risultato con URL del PDF
 */
/**
 * Crea una verifica di sicurezza (Check-List) per una macchina
 * @param {string} token Token di autenticazione
 * @param {string} machineId ID della macchina
 * @param {Object} formData Dati del form, inclusa la lista di controlli
 * @returns {Object} Risultato con URL del PDF
 */
/**
 * Crea un Rapporto di Controllo Periodico (Check-List) per una macchina
 * @param {string} token Token di autenticazione
 * @param {string} machineId ID della macchina
 * @param {Object} formData Dati del form, inclusa la checklist, i dati cliente e i ricambi
 * @returns {Object} Risultato con URL del PDF
 */
function createVerificaSicurezza(token, machineId, formData) {
  try {
    const user = validateToken(token);
    if (!user) throw 'Non autorizzato';

    const role = getRoleByEmail(user);
    if (role !== 'tecnico' && role !== 'admin') {
      return { ok: false, error: 'Solo tecnici e admin possono creare verifiche' };
    }

    const machineData = getMachineById(token, machineId);
    if (!machineData) {
      return { ok: false, error: 'Macchina non trovata' };
    }

    // Prepara i dati per il PDF
    const pdfData = {
      // Dati Cliente (da form) e Macchina (automatici)
      cliente: formData.cliente,
      contattoNome: formData.contattoNome,
      contattoTel: formData.contattoTel,
      contattoEmail: formData.contattoEmail,
      marca: extractMarca(machineData.model),
      modello: machineData.model,
      matricola: machineData.serial,
      
      // Meta
      tecnico: getNameByEmail(user),
      dataVerifica: Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy'),
      
      // Contenuti dinamici
      verificheSicurezze: formData.verificheSicurezze || [],
      manutenzionePreventiva: formData.manutenzionePreventiva || [],
      verificheElettrico: formData.verificheElettrico || [],
      verificheDiesel: formData.verificheDiesel || [],
      verificheTrilaterale: formData.verificheTrilaterale || [],
      noteGenerali: formData.noteGenerali || '',
      ricambi: formData.ricambi || [],
      
      // Firme
      firmaTecnicoUrl: formData.firmaTecnicoUrl || '',
      firmaClienteUrl: formData.firmaClienteUrl || ''
    };

    const result = generatePdfFromTemplate('verifiche', pdfData, FOLDER_VERIFICHE_ID);

    if (!result.ok) {
      return result;
    }

    const shVerifiche = _getSheet('Verifiche');
    if (shVerifiche) {
      shVerifiche.appendRow([
        result.fileId, machineId, machineData.model, 'Verifica Periodica',
        user, _now(), result.pdfUrl
      ]);
    }

    Logger.log('✅ Rapporto di Controllo creato: ' + result.fileName);
    return result;

  } catch(e) {
    Logger.log('❌ ERRORE createVerificaSicurezza: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}




/**
 * Crea una garanzia per una macchina
 * @param {string} token Token di autenticazione
 * @param {string} machineId ID della macchina
 * @param {Object} formData Dati del form compilato
 * @returns {Object} Risultato con URL del PDF
 */
/**
 * Crea una garanzia completa per una macchina con allegati
 * @param {string} token Token di autenticazione
 * @param {string} machineId ID della macchina
 * @param {Object} formData Dati del form compilato
 * @returns {Object} Risultato con URL del PDF
 */
/**
 * Crea una garanzia per una macchina
 * @param {string} token Token di autenticazione
 * @param {string} machineId ID della macchina
 * @param {Object} formData Dati del form compilato
 * @returns {Object} Risultato con URL del PDF
 */
function createGaranzia(token, machineId, formData) {
  try {
    const user = validateToken(token);
    if (!user) throw 'Non autorizzato';

    const role = getRoleByEmail(user);
    if (role !== 'tecnico' && role !== 'admin') {
      return { ok: false, error: 'Solo tecnici e admin possono creare garanzie' };
    }

    // Recupera dati macchina
    const machineData = getMachineById(token, machineId);
    if (!machineData) {
      return { ok: false, error: 'Macchina non trovata' };
    }

    // Gestisci gli allegati fotografici/video
    const allegati = [];
    
    if (formData.photos && formData.photos.length > 0) {
      for (let i = 0; i < formData.photos.length; i++) {
        const photo = formData.photos[i];
        const photoBlob = Utilities.newBlob(
          Utilities.base64Decode(photo.data), 
          photo.mimeType, 
          photo.name
        );
        
        const folder = DriveApp.getFolderById(FOLDER_GARANZIE_ID);
        const file = folder.createFile(photoBlob);
        
        allegati.push({
          type: photo.mimeType.startsWith('image/') ? 'image' : 'video',
          name: file.getName(),
          url: file.getUrl(),
          fileId: file.getId()
        });
      }
    }

    // Gestisci la firma digitale
    let firmaUrl = '';
    if (formData.signature) {
      const signatureBlob = Utilities.newBlob(
        Utilities.base64Decode(formData.signature.data), 
        'image/png', 
        'firma_' + machineId + '_' + Date.now() + '.png'
      );
      
      const folder = DriveApp.getFolderById(FOLDER_GARANZIE_ID);
      const signatureFile = folder.createFile(signatureBlob);
      firmaUrl = signatureFile.getUrl();
    }

    // Prepara i dati completi per il PDF
    const pdfData = {
      // Dati automatici
      cliente: machineData.cliente || 'Non specificato',
      marca: extractMarca(machineData.model),
      modello: machineData.model,
      matricola: machineData.serial,
      
      // Dati compilati dal tecnico
      difettoRiscontrato: formData.difettoRiscontrato || '',
      codiceErrore: formData.codiceErrore || '',
      azioneIntrapresa: formData.azioneIntrapresa || '',
      difettoRisolto: formData.difettoRisolto || 'No',
      oreMacchina: formData.oreMacchina || '',
      
      // Meta
      tecnico: getNameByEmail(user),
      dataIntervento: Utilities.formatDate(new Date(), 'Europe/Rome', 'dd/MM/yyyy'),
      
      // Allegati
      firmaUrl: firmaUrl,
      allegati: allegati
    };

    // Genera il PDF
    const result = generatePdfFromTemplate('garanzia', pdfData, FOLDER_GARANZIE_ID);

    if (!result.ok) {
      return result;
    }

    // Registra nel foglio "Garanzie" (opzionale)
    const shGaranzie = _getSheet('Garanzie');
    if (shGaranzie) {
      shGaranzie.appendRow([
        result.fileId,
        machineId,
        machineData.model,
        'Garanzia',
        user,
        _now(),
        result.pdfUrl,
        allegati.length
      ]);
    }

    Logger.log('✅ Garanzia creata con ' + allegati.length + ' allegati: ' + result.fileName);
    return result;

  } catch(e) {
    Logger.log('❌ ERRORE createGaranzia: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}

// Funzione helper per estrarre la marca dal modello
function extractMarca(model) {
  if (!model) return 'N/A';
  const match = model.match(/^([A-Z]+)/i);
  return match ? match[1].toUpperCase() : 'N/A';
}


/**
 * Lista tutti i documenti di verifiche e garanzie dalla cartella Drive
 * @param {string} token Token di autenticazione
 * @returns {Array} Lista dei documenti
 */
function listVerificheDocuments(token) {
  try {
    const user = validateToken(token);
    if (!user) throw 'Non autorizzato';

    const documents = [];

    // Leggi verifiche
    try {
      const folderVerifiche = DriveApp.getFolderById(FOLDER_VERIFICHE_ID);
      const filesVerifiche = folderVerifiche.getFiles();
      
      while (filesVerifiche.hasNext()) {
        const file = filesVerifiche.next();
        if (file.getMimeType() === 'application/pdf') {
          documents.push({
            fileId: file.getId(),
            title: file.getName(),
            type: 'Verifica Sicurezza',
            size: formatBytes(file.getSize()),
            createdAt: Utilities.formatDate(file.getDateCreated(), 'Europe/Rome', 'dd/MM/yyyy HH:mm')
          });
        }
      }
    } catch(e) {
      Logger.log('⚠️ Errore lettura verifiche: ' + e);
    }

    // Leggi garanzie
    try {
      const folderGaranzie = DriveApp.getFolderById(FOLDER_GARANZIE_ID);
      const filesGaranzie = folderGaranzie.getFiles();
      
      while (filesGaranzie.hasNext()) {
        const file = filesGaranzie.next();
        if (file.getMimeType() === 'application/pdf') {
          documents.push({
            fileId: file.getId(),
            title: file.getName(),
            type: 'Garanzia',
            size: formatBytes(file.getSize()),
            createdAt: Utilities.formatDate(file.getDateCreated(), 'Europe/Rome', 'dd/MM/yyyy HH:mm')
          });
        }
      }
    } catch(e) {
      Logger.log('⚠️ Errore lettura garanzie: ' + e);
    }

    // Ordina per data (più recenti prima)
    documents.sort((a, b) => {
      return new Date(b.createdAt) - new Date(a.createdAt);
    });

    Logger.log('✅ Trovati ' + documents.length + ' documenti');
    return documents;

  } catch(e) {
    Logger.log('❌ ERRORE listVerificheDocuments: ' + e);
    return [];
  }
}
function uploadFileToDrive(token, base64Data, fileName, mimeType, folderId) {
  try {
    const user = validateToken(token);
    if (!user) throw 'Non autorizzato';

    // Decodifica base64
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data), 
      mimeType, 
      fileName
    );

    // Carica il file nella cartella specificata
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);

    Logger.log('✅ File caricato: ' + file.getName());

    return {
      ok: true,
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      fileName: file.getName()
    };

  } catch(e) {
    Logger.log('❌ ERRORE uploadFileToDrive: ' + e);
    return { ok: false, error: e.message || String(e) };
  }
}
// ===============================================
//  FUNZIONI AREA CLIENTE
// ===============================================

/**
 * 1️⃣ DASHBOARD - Statistiche Cliente
 */
function getClientStats(token){
  try {
    const userEmail = validateToken(token);
    if(!userEmail) return {machines: 0, activeRentals: 0, openTickets: 0};
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Utenti');
    const usersData = usersSheet.getDataRange().getValues();
    
    // Trova Company del cliente
    let clientCompany = null;
    for(let i = 1; i < usersData.length; i++){
      if(usersData[i][1] && usersData[i][1].toLowerCase() === userEmail.toLowerCase()){
        clientCompany = usersData[i][4]; // Col E: Company
        break;
      }
    }
    
    if(!clientCompany) return {machines: 0, activeRentals: 0, openTickets: 0};
    
    // Conta macchine del cliente
    const machinesSheet = ss.getSheetByName('Macchine');
    const machinesData = machinesSheet.getDataRange().getValues();
    const headers = machinesData[0];
    const ownerTypeIdx = headers.indexOf('OwnerType');
    
    let machinesCount = 0;
    for(let i = 1; i < machinesData.length; i++){
      if(machinesData[i][ownerTypeIdx] && 
         machinesData[i][ownerTypeIdx].toLowerCase() === clientCompany.toLowerCase()){
        machinesCount++;
      }
    }
    
    // Conta noleggi attivi (opzionale - se hai il foglio Noleggi)
    let activeRentals = 0;
    const rentalsSheet = ss.getSheetByName('Noleggi');
    if(rentalsSheet){
      const rentalsData = rentalsSheet.getDataRange().getValues();
      for(let i = 1; i < rentalsData.length; i++){
        if(rentalsData[i][2] === userEmail && 
           (rentalsData[i][5] === 'active' || rentalsData[i][5] === 'requested')){
          activeRentals++;
        }
      }
    }
    
    // Conta ticket aperti (opzionale - se hai il foglio Ticket)
    let openTickets = 0;
    const ticketsSheet = ss.getSheetByName('Ticket');
    if(ticketsSheet){
      const ticketsData = ticketsSheet.getDataRange().getValues();
      for(let i = 1; i < ticketsData.length; i++){
        if(ticketsData[i][4] === userEmail && ticketsData[i][6] !== 'chiuso'){
          openTickets++;
        }
      }
    }
    
    return {
      machines: machinesCount,
      activeRentals: activeRentals,
      openTickets: openTickets
    };
    
  } catch(error){
    Logger.log('❌ Errore getClientStats: ' + error);
    return {machines: 0, activeRentals: 0, openTickets: 0};
  }
}


/**
 * Lista macchine del cliente loggato
 */
function getClientMachines(token){
  try {
    const userEmail = validateToken(token);
    if(!userEmail) {
      Logger.log('❌ Token non valido');
      return [];
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Utenti');
    const usersData = usersSheet.getDataRange().getValues();
    
    // Trova Company del cliente (Colonna D nei Utenti)
    let clientCompany = null;
    for(let i = 1; i < usersData.length; i++){
      if(usersData[i][0] && usersData[i][0].toLowerCase() === userEmail.toLowerCase()){
        clientCompany = usersData[i][3]; // Col D: Company
        Logger.log('✅ Cliente trovato: ' + clientCompany);
        break;
      }
    }
    
    if(!clientCompany){
      Logger.log('⚠️ Cliente senza azienda: ' + userEmail);
      return [];
    }
    
    // Carica macchine
    const machinesSheet = ss.getSheetByName('Macchine');
    if(!machinesSheet){
      Logger.log('❌ Foglio Macchine non trovato');
      return [];
    }
    
    const machinesData = machinesSheet.getDataRange().getValues();
    
    // ⭐ NON USARE headers.indexOf perché le colonne sono fisse
    // Colonna A: ID
    // Colonna B: Serial
    // Colonna C: Model
    // Colonna D: OwnerType ← QUI c'è "APETRII ANA MARIA"
    // Colonna E: Cliente
    // Colonna F: Status
    
    Logger.log('🔍 Cerco macchine con OwnerType (Col D) = "' + clientCompany + '"');
    
    const clientMachines = [];
    
    // Filtra solo macchine del cliente
    for(let i = 1; i < machinesData.length; i++){
      const row = machinesData[i];
      
      const machineId = row[0] || '';          // Col A: ID
      const serial = row[1] || '';              // Col B: Serial
      const model = row[2] || '';               // Col C: Model
      const ownerType = row[3] ? String(row[3]).trim() : ''; // Col D: OwnerType
      const status = row[5] || 'unknown';       // Col F: Status
      
      // ⭐ CONFRONTA OwnerType con Company del cliente
      if(ownerType === clientCompany){
        clientMachines.push({
          id: machineId,
          serial: serial,
          model: model,
          status: status,
          ownerType: ownerType,
          categoria: 'N/A'
        });
        
        Logger.log('✅ Macchina trovata: ' + machineId + ' - ' + model);
      }
    }
    
    Logger.log('🎉 Totale macchine del cliente ' + clientCompany + ': ' + clientMachines.length);
    return clientMachines;
    
  } catch(error){
    Logger.log('❌ Errore getClientMachines: ' + error);
    return [];
  }
}



/**
 * 3️⃣ NOLEGGIO - Lista macchine noleggiabili (solo Manitech)
 */
function getClientAvailableRentals(token){
  try {
    const userEmail = validateToken(token);
    if(!userEmail) return [];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const machinesSheet = ss.getSheetByName('Macchine');
    const machinesData = machinesSheet.getDataRange().getValues();
    const headers = machinesData[0];
    
    const idIdx = headers.indexOf('ID');
    const serialIdx = headers.indexOf('Serial');
    const modelIdx = headers.indexOf('Model');
    const statusIdx = headers.indexOf('Status');
    const ownerTypeIdx = headers.indexOf('OwnerType');
    const noleggioIdx = headers.indexOf('Noleggiabile');
    const categoriaIdx = headers.indexOf('Categoria');
    
    const availableMachines = [];
    
    // Filtra: OwnerType = "manitech" E Noleggiabile = "Sì" E Status = "available"
    for(let i = 1; i < machinesData.length; i++){
      const row = machinesData[i];
      
      if(row[ownerTypeIdx] && 
         row[ownerTypeIdx].toLowerCase() === 'manitech' &&
         row[noleggioIdx] === 'Sì' &&
         row[statusIdx] === 'available'){
        
        availableMachines.push({
          id: row[idIdx] || '',
          serial: row[serialIdx] || '',
          model: row[modelIdx] || '',
          status: row[statusIdx] || '',
          categoria: row[categoriaIdx] || 'N/A'
        });
      }
    }
    
    Logger.log('✅ Macchine disponibili per noleggio: ' + availableMachines.length);
    return availableMachines;
    
  } catch(error){
    Logger.log('❌ Errore getClientAvailableRentals: ' + error);
    return [];
  }
}


/**
 * 3️⃣ NOLEGGIO - Richieste noleggio del cliente
 */
function getClientRentals(token){
  try {
    const userEmail = validateToken(token);
    if(!userEmail) return [];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rentalsSheet = ss.getSheetByName('Noleggi');
    
    if(!rentalsSheet) return [];
    
    const rentalsData = rentalsSheet.getDataRange().getValues();
    const clientRentals = [];
    
    // Filtra solo noleggi del cliente (col C = requestedBy)
    for(let i = 1; i < rentalsData.length; i++){
      if(rentalsData[i][2] === userEmail){
        clientRentals.push({
          id: rentalsData[i][0],
          machineId: rentalsData[i][1],
          requestedBy: rentalsData[i][2],
          startDate: rentalsData[i][3],
          endDate: rentalsData[i][4],
          status: rentalsData[i][5],
          notes: rentalsData[i][6] || ''
        });
      }
    }
    
    return clientRentals;
    
  } catch(error){
    Logger.log('❌ Errore getClientRentals: ' + error);
    return [];
  }
}


/**
 * 4️⃣ TICKET - Ticket del cliente
 */
function getClientTickets(token){
  try {
    const userEmail = validateToken(token);
    if(!userEmail) return [];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Utenti');
    const usersData = usersSheet.getDataRange().getValues();
    
    // Trova Company del cliente
    let clientCompany = null;
    for(let i = 1; i < usersData.length; i++){
      if(usersData[i][1] && usersData[i][1].toLowerCase() === userEmail.toLowerCase()){
        clientCompany = usersData[i][4];
        break;
      }
    }
    
    if(!clientCompany) return [];
    
    // Carica ticket
    const ticketsSheet = ss.getSheetByName('Ticket');
    if(!ticketsSheet) return [];
    
    const ticketsData = ticketsSheet.getDataRange().getValues();
    const machinesSheet = ss.getSheetByName('Macchine');
    const machinesData = machinesSheet.getDataRange().getValues();
    const headers = machinesData[0];
    const ownerTypeIdx = headers.indexOf('OwnerType');
    
    const clientTickets = [];
    
    // Filtra ticket relativi alle macchine del cliente
    for(let i = 1; i < ticketsData.length; i++){
      const ticketMachineId = ticketsData[i][1];
      
      // Verifica se la macchina appartiene al cliente
      for(let j = 1; j < machinesData.length; j++){
        if(machinesData[j][0] === ticketMachineId && 
           machinesData[j][ownerTypeIdx] && 
           machinesData[j][ownerTypeIdx].toLowerCase() === clientCompany.toLowerCase()){
          
          clientTickets.push({
            id: ticketsData[i][0],
            machineId: ticketsData[i][1],
            machineModel: ticketsData[i][2],
            machineSerial: ticketsData[i][3],
            clientName: ticketsData[i][4],
            description: ticketsData[i][5],
            status: ticketsData[i][6],
            createdAt: ticketsData[i][7]
          });
          break;
        }
      }
    }
    
    return clientTickets;
    
  } catch(error){
    Logger.log('❌ Errore getClientTickets: ' + error);
    return [];
  }
}


/**
 * 5️⃣ DOCUMENTI - Verifiche e garanzie delle macchine del cliente
 */
function getClientDocuments(token){
  try {
    const userEmail = validateToken(token);
    if(!userEmail) return [];
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Utenti');
    const usersData = usersSheet.getDataRange().getValues();
    
    // Trova Company del cliente
    let clientCompany = null;
    for(let i = 1; i < usersData.length; i++){
      if(usersData[i][1] && usersData[i][1].toLowerCase() === userEmail.toLowerCase()){
        clientCompany = usersData[i][4];
        break;
      }
    }
    
    if(!clientCompany) return [];
    
    // Lista tutte le macchine del cliente
    const machinesSheet = ss.getSheetByName('Macchine');
    const machinesData = machinesSheet.getDataRange().getValues();
    const headers = machinesData[0];
    const ownerTypeIdx = headers.indexOf('OwnerType');
    const idIdx = headers.indexOf('ID');
    
    const clientMachineIds = [];
    for(let i = 1; i < machinesData.length; i++){
      if(machinesData[i][ownerTypeIdx] && 
         machinesData[i][ownerTypeIdx].toLowerCase() === clientCompany.toLowerCase()){
        clientMachineIds.push(machinesData[i][idIdx]);
      }
    }
    
    if(clientMachineIds.length === 0) return [];
    
    // Carica documenti (da Drive)
    // Nota: Implementa qui la logica per recuperare i PDF da Drive
    // filtrati per machineId appartenenti al cliente
    
    return []; // Da implementare con listVerificheDocuments filtrato
    
  } catch(error){
    Logger.log('❌ Errore getClientDocuments: ' + error);
    return [];
  }
}
