function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate().setTitle("Stella's RAB Portal").addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, shrink-to-fit=no').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function getSheetByNameRobust(ss, targetName) {
  if (!ss) return null;
  const sheets = ss.getSheets();
  const targetNorm = targetName.toLowerCase().replace(/s$/, "").trim();
  
  let sheet = ss.getSheetByName(targetName);
  if (sheet) return sheet;
  
  for (let i = 0; i < sheets.length; i++) {
    const sheetName = sheets[i].getName();
    const sheetNorm = sheetName.toLowerCase().replace(/s$/, "").trim();
    if (sheetNorm === targetNorm) {
      return sheets[i];
    }
  }
  return null;
}

function getSpreadsheetDiag() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: "Active spreadsheet is null" };
    
    const sheets = ss.getSheets();
    const list = sheets.map(s => {
      const name = s.getName();
      const rows = s.getLastRow();
      const cols = s.getLastColumn();
      let sample = [];
      if (rows > 0 && cols > 0) {
        const rawValues = s.getRange(1, 1, Math.min(rows, 3), Math.min(cols, 8)).getValues();
        sample = rawValues.map(row => 
          row.map(cell => {
            if (cell instanceof Date) return cell.toLocaleString('id-ID');
            if (cell === null || cell === undefined) return "";
            return cell.toString();
          })
        );
      }
      return { name: name, rows: rows, cols: cols, sample: sample };
    });
    return { success: true, sheets: list, url: ss.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let masterSheet = getSheetByNameRobust(ss, "Master_Harga") || ss.insertSheet("Master_Harga");
  const masterHeaders = ["Material_ID", "Category", "Item_Name", "Specification", "Unit", "Unit_Price", "Ref_URL", "Conversion_Factor"];
  masterSheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders]).setFontWeight("bold").setBackground("#2563eb").setFontColor("#ffffff");
  masterSheet.setFrozenRows(1);
  if (masterSheet.getLastRow() === 1) masterSheet.getRange(2, 1, 2, masterHeaders.length).setValues([["MAT001", "Kayu", "Multipleks 15mm", "Tunas / Setara", "Lembar", 245000, "#", 1], ["MAT002", "Akrilik", "Akrilik Bening 5mm", "Marga Cipta", "Lembar", 850000, "#", 1]]);

  let projectSheet = getSheetByNameRobust(ss, "Data_Projects") || ss.insertSheet("Data_Projects");
  const projectHeaders = ["Project_ID", "Timestamp", "Created_By", "Project_Name", "Client_Email", "Total_Budget", "User_Note"];
  projectSheet.getRange(1, 1, 1, projectHeaders.length).setValues([projectHeaders]).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
  projectSheet.setFrozenRows(1);

  let detailsSheet = getSheetByNameRobust(ss, "RAB_Details") || ss.insertSheet("RAB_Details");
  const detailsHeaders = ["Project_ID", "Fixture_Name", "Material_Name", "Unit", "QTY", "Unit_Price", "Subtotal"];
  detailsSheet.getRange(1, 1, 1, detailsHeaders.length).setValues([detailsHeaders]).setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
  detailsSheet.setFrozenRows(1);

  let userSheet = getSheetByNameRobust(ss, "Config_Users") || ss.insertSheet("Config_Users");
  const userHeaders = ["Username", "Password", "Role"];
  userSheet.getRange(1, 1, 1, userHeaders.length).setValues([userHeaders]).setFontWeight("bold").setBackground("#dc2626").setFontColor("#ffffff");
  userSheet.setFrozenRows(1);
  if (userSheet.getLastRow() === 1) userSheet.getRange(2, 1, 2, userHeaders.length).setValues([["stella", "owner123", "Owner"], ["guest1", "guest123", "Guest"]]);

  let logSheet = getSheetByNameRobust(ss, "Activity_Logs") || ss.insertSheet("Activity_Logs");
  const logHeaders = ["Timestamp", "Username", "Role", "Action", "Details"];
  logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]).setFontWeight("bold").setBackground("#4b5563").setFontColor("#ffffff");
  logSheet.setFrozenRows(1);

  return "Database Berhasil Dikonfigurasi!";
}

function loginUser(u, p) {
  try {
    const normalizedUser = (u || "").toString().trim().toLowerCase();
    const cleanPass = (p || "").toString().trim();
    
    // 1. Emergency Fallback Bypass (Garansi 100% Bisa Masuk)
    if (normalizedUser === "stella" && cleanPass === "owner123") {
      const fallbackUser = { username: "stella", role: "Owner" };
      try { writeLog("stella", "Owner", "LOGIN_FALLBACK", "Berhasil masuk via emergency fallback"); } catch(e){}
      return { success: true, user: fallbackUser };
    }
    if (normalizedUser === "guest1" && cleanPass === "guest123") {
      const fallbackUser = { username: "guest1", role: "Guest" };
      try { writeLog("guest1", "Guest", "LOGIN_FALLBACK", "Berhasil masuk via emergency fallback"); } catch(e){}
      return { success: true, user: fallbackUser };
    }

    // 2. Database-backed login
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: "Active Spreadsheet tidak ditemukan! Pastikan skrip ini terikat (container-bound) ke Google Sheets." };
    
    let userSheet = getSheetByNameRobust(ss, "Config_Users");
    
    // Auto-setup database jika sheet user tidak ditemukan
    if (!userSheet) {
      setupDatabase();
      userSheet = getSheetByNameRobust(ss, "Config_Users");
    }
    
    if (!userSheet) return { success: false, error: "Gagal menginisialisasi database secara otomatis. Periksa izin akses spreadsheet Anda." };
    
    const data = userSheet.getDataRange().getValues();
    let authUser = null;
    for(let i = 1; i < data.length; i++) {
      if((data[i][0]||"").toString().trim().toLowerCase() === normalizedUser && (data[i][1]||"").toString().trim() === cleanPass) {
        authUser = { username: data[i][0], role: (data[i][2]||"").toString().trim() }; 
        break;
      }
    }
    if (authUser) { 
      writeLog(authUser.username, authUser.role, "LOGIN", "Berhasil masuk"); 
      return { success: true, user: authUser }; 
    } else {
      return { success: false, error: "Username atau Password salah! (Default: stella / owner123)" };
    }
  } catch (e) { return { success: false, error: e.toString() }; }
}

function writeLog(u, r, a, d) { try { const ss = SpreadsheetApp.getActiveSpreadsheet(); getSheetByNameRobust(ss, "Activity_Logs").appendRow([new Date(), u, r, a, d]); } catch(e){} }

function getMaterialData() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const sheet = getSheetByNameRobust(ss, "Master_Harga");
  if(!sheet) return [];
  const d = sheet.getDataRange().getValues(); 
  if(d.length<=1) return []; 
  const h = d[0]; 
  return d.slice(1).map(r => { 
    let o = {}; 
    h.forEach((k, i) => o[k] = r[i]); 
    return o; 
  }); 
}

function getProjectHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let projectSheet = getSheetByNameRobust(ss, "Data_Projects") || ss.insertSheet("Data_Projects");
    let detailsSheet = getSheetByNameRobust(ss, "RAB_Details") || ss.insertSheet("RAB_Details");
    
    const pData = projectSheet.getDataRange().getValues();
    const dData = detailsSheet.getDataRange().getValues();
    const existingProjects = new Set();
    
    if (pData.length > 1) {
      for (let i = 1; i < pData.length; i++) {
        if (pData[i][0]) existingProjects.add(pData[i][0].toString().trim());
      }
    }
    
    const detailsProjects = {};
    if (dData.length > 1) {
      const headers = dData[0];
      const subtotalIdx = headers.indexOf("Subtotal") !== -1 ? headers.indexOf("Subtotal") : 5;
      
      for (let i = 1; i < dData.length; i++) {
        const pId = (dData[i][0] || "").toString().trim();
        if (!pId) continue;
        const subtotal = parseFloat(dData[i][subtotalIdx]) || 0;
        if (!detailsProjects[pId]) {
          detailsProjects[pId] = { totalBudget: 0, timestamp: new Date() };
        }
        detailsProjects[pId].totalBudget += subtotal;
      }
    }
    
    let syncedAny = false;
    for (const pId in detailsProjects) {
      if (!existingProjects.has(pId)) {
        projectSheet.appendRow([pId, detailsProjects[pId].timestamp, "System Sync", "Proyek " + pId, "client@example.com", detailsProjects[pId].totalBudget, "Auto-synced from RAB_Details"]);
        syncedAny = true;
      }
    }
    
    let finalPData = pData;
    if (syncedAny) finalPData = projectSheet.getDataRange().getValues();
    if (finalPData.length <= 1) return [];
    const h = finalPData[0];
    return finalPData.slice(1).map(r => {
      let o = {};
      h.forEach((k, i) => {
        let val = r[i];
        if (val instanceof Date) val = val.toISOString();
        o[k] = val;
      });
      return o;
    }).reverse();
  } catch (e) {
    Logger.log("Error in getProjectHistory: " + e.toString());
    return [];
  }
}

function processAIExtraction(base64Data, mimeType, currentUser, currentRole) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) throw new Error("API Key belum diset!");
    const masterMaterials = getMaterialData();
    const materialListString = masterMaterials.map(m => `${m.Material_ID}: ${m.Item_Name} (Rp ${m.Unit_Price})`).join("\n");
    const prompt = `Kamu QS. Ekstrak gambar Gamker ini. Kombinasikan dengan MASTER MATERIAL:\n${materialListString}\nPecah komponen. Berikan estimasi harga satuan umum di Indonesia. OUTPUT JSON MURNI TANPA MARKDOWN: { "rabItems": [ { "fixtureName": "Nama Bagian", "materialName": "Nama Bahan", "unit": "ea/m2/liter/lembar/sak/kg/dll", "qty": 1.00, "estimatedPrice": 0 } ] }`;
    const payload = { contents: [{ parts: [ { text: prompt }, { inline_data: { mime_type: mimeType, data: base64Data.split(',')[1] } } ] }] };
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
    const jsonRes = JSON.parse(res.getContentText());
    if (jsonRes.error) throw new Error(`Google API Error: ${jsonRes.error.message}`);
    let aiText = jsonRes.candidates[0].content.parts[0].text;
    const parsedData = cleanAndParseJSON(aiText);
    return { success: true, items: parsedData.rabItems };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function processAIExtractionMultiple(files, currentUser, currentRole) {
  try {
    const masterMaterials = getMaterialData();
    const materialListString = masterMaterials.map(m => `${m.Material_ID}: ${m.Item_Name} (Rp ${m.Unit_Price})`).join("\n");
    const prompt = `Kamu QS. Ekstrak seluruh dokumen referensi/daftar material terlampir ini secara mendalam dan menyeluruh. Kombinasikan dengan MASTER MATERIAL:\n${materialListString}\nPecah komponen secara detail. Berikan estimasi harga satuan umum di Indonesia berdasarkan seluruh dokumen yang diunggah. OUTPUT JSON MURNI TANPA MARKDOWN: { "rabItems": [ { "fixtureName": "Nama Bagian", "materialName": "Nama Bahan", "unit": "ea/m2/liter/lembar/sak/kg/dll", "qty": 1.00, "estimatedPrice": 0 } ] }`;
    
    // Map files for general failover engine
    const imagesToPass = files.map(f => ({
      mimeType: f.mimeType,
      base64Data: f.base64Data
    }));
    
    const failoverResult = callAIChatFailover(prompt, "Kamu adalah Quantity Surveyor (QS) Indonesia yang ahli.", imagesToPass, true);
    if (!failoverResult.success) throw new Error(failoverResult.error);
    
    const parsedData = cleanAndParseJSON(failoverResult.text);
    return { success: true, items: parsedData.rabItems, providerUsed: failoverResult.provider };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function deepResearchMaterial(materialName) {
  try {
    const serpApiKey = PropertiesService.getScriptProperties().getProperty('SERPAPI_KEY');
    if (!serpApiKey) throw new Error("SERPAPI_KEY belum diset!");
    const url = `https://serpapi.com/search.json?engine=google_shopping&q=${encodeURIComponent(materialName)}&google_domain=google.co.id&gl=id&hl=id&api_key=${serpApiKey}`;
    const res = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    const jsonRes = JSON.parse(res.getContentText());
    if (jsonRes.error) throw new Error(`SerpApi Error: ${jsonRes.error}`);
    const shoppingResults = jsonRes.shopping_results;
    if (!shoppingResults || shoppingResults.length === 0) return { success: true, data: { actualPrice: 0, sourceUrl: "#" } };
    const topProduct = shoppingResults[0];
    const cleanPrice = parseFloat(topProduct.price.replace(/[^0-9]/g, '')) || 0;
    
    return processResearchResult(materialName, { price: cleanPrice, unit: "", sourceUrl: topProduct.link || "#" });
  } catch (e) { return { success: false, error: e.toString() }; }
}

function researchMaterialCSE(materialName) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GCSE_API_KEY');
    const cx = PropertiesService.getScriptProperties().getProperty('GCSE_CX');
    if (!apiKey || !cx) throw new Error("GCSE_API_KEY atau GCSE_CX belum dikonfigurasi di Script Properties!");
    
    const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${cx}&q=${encodeURIComponent("harga " + materialName + " indonesia")}`;
    const res = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    if (json.error) throw new Error(`Google CSE Error: ${json.error.message}`);
    const items = json.items;
    if (!items || items.length === 0) return { success: true, data: { actualPrice: 0, sourceUrl: "#" } };
    
    let foundPrice = 0;
    let foundLink = "#";
    const priceRegex = /(?:Rp|IDR)\s?(\d{1,3}(?:\.\d{3})+|\d+)/i;
    for (let i = 0; i < items.length; i++) {
      const textToScan = `${items[i].title} ${items[i].snippet}`;
      const match = textToScan.match(priceRegex);
      if (match) {
        const cleanPriceStr = match[1].replace(/\./g, '').trim();
        const price = parseFloat(cleanPriceStr);
        if (price > 1000) { foundPrice = price; foundLink = items[i].link; break; }
      }
    }
    if (foundPrice === 0) foundLink = items[0].link;
    return processResearchResult(materialName, { price: foundPrice, unit: "", sourceUrl: foundLink });
  } catch (e) { return { success: false, error: e.toString() }; }
}

function researchMaterialGeminiSearch(materialName) {
  try {
    const prompt = `Analisis dan cari perkiraan harga pasar aktual terbaru untuk material berikut di Indonesia: "${materialName}". Bandingkan beberapa harga pasar umum (contoh: Tokopedia, Shopee, distributor material lokal), lalu ambil harga rata-rata wajar yang paling akurat. Berikan jawaban dalam format JSON MURNI:\n{\n  "price": 150000,\n  "unit": "dus/lembar/m2/btg/sak/dll",\n  "sourceUrl": "https://sumber-url.com"\n}\nJika tidak menemukan harga, set "price" ke 0, "unit" ke "", dan "sourceUrl" ke "#".`;
    
    // We pass Gemini grounding search tool ONLY if provider is GEMINI.
    const failoverResult = callAIChatFailover(prompt, "Kamu adalah asisten riset QS Indonesia yang akurat.", null, true, true);
    if (!failoverResult.success) throw new Error(failoverResult.error);
    
    const parsedData = cleanAndParseJSON(failoverResult.text);
    return processResearchResult(materialName, parsedData);
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function brainstormDesign(chatInput, currentItemsJson, imageBase64, chatHistoryTranscript) {
  try {
    let prompt = `Klien Stella ingin mendiskusikan budget, konsep visual, tata ruang (layouting), display barang, atau modifikasi material desain.
Permintaan Diskusi Klien: "${chatInput}"
RAB Saat Ini: ${JSON.stringify(currentItemsJson)}`;

    if (chatHistoryTranscript) {
      prompt += `\n\nRiwayat Percakapan Sebelumnya (Simpan memori kontekstual ini):\n${chatHistoryTranscript}`;
    }

    prompt += `\n\nTugas Anda sebagai World-Class Interior Designer & VM di Indonesia:\n1. Berikan opini desain yang berkelas, estetis, dan profesional internasional namun disesuaikan dengan konteks pasar ritel/interior Indonesia.\n2. Berikan ide-ide konkret tentang elemen desain apa saja yang bisa di-update, di-upgrade, disesuaikan, atau dikurangi budgetnya untuk efisiensi biaya.\n3. Ajukan pertanyaan aktif kepada user/klien: "Apakah Anda ingin melakukan update desain secara otomatis pada tabel kalkulasi?"\n4. PENTING: Akhiri seluruh tanggapan/opini Anda dengan sebuah pertanyaan tegas yang menanyakan usulan atau ide rekomendasi mana yang ingin dibahas lebih lanjut oleh klien untuk mulai dikembangkan.\nFormat jawaban harus dalam JSON MURNI:\n{\n  "reply": "Pesan edukatif...",\n  "newRab": [ { "fixtureName": "Nama Bagian", "materialName": "Bahan Baru", "qty": 1.00, "unitPrice": 120000 } ]\n}`;

    let images = null;
    if (imageBase64 && imageBase64.indexOf("data:") !== -1) {
      const mimeType = imageBase64.substring(imageBase64.indexOf("data:") + 5, imageBase64.indexOf(";base64"));
      images = [{ mimeType: mimeType, base64Data: imageBase64 }];
    }
    
    const failoverResult = callAIChatFailover(prompt, "Anda adalah Co-Designer, World-Class Interior Designer, dan Visual Merchandiser (VM) ahli di Indonesia.", images, true);
    if (!failoverResult.success) throw new Error(failoverResult.error);
    
    const parsedData = cleanAndParseJSON(failoverResult.text);
    return { success: true, data: parsedData, providerUsed: failoverResult.provider };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function processRAB(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projectId = "PRJ-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    getSheetByNameRobust(ss, "Data_Projects").appendRow([projectId, new Date(), payload.currentUser, payload.projectName, payload.clientEmail, payload.totalBudget, payload.userNote || ""]);
    const detailsSheet = getSheetByNameRobust(ss, "RAB_Details");
    const rows = payload.items.map(item => [projectId, item.fixtureName, item.materialName, item.unit || "ea", item.qty, item.unitPrice, item.subtotal]);
    detailsSheet.getRange(detailsSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    generateAndSendFiles(projectId, payload);
    return { success: true, projectId: projectId };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function saveRABVersion(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projectId = payload.projectId || "PRJ-" + Utilities.getUuid().substring(0, 8).toUpperCase();
    
    // Ensure the new header columns exist in Data_Projects
    let projectSheet = getSheetByNameRobust(ss, "Data_Projects");
    if (projectSheet) {
      const colCount = projectSheet.getLastColumn();
      if (colCount > 0) {
        const headers = projectSheet.getRange(1, 1, 1, colCount).getValues()[0];
        if (headers.indexOf("Target_Budget") === -1) {
          projectSheet.getRange(1, 8, 1, 3).setValues([["Target_Budget", "Contingency_Rate", "Markup_Rate"]]).setFontWeight("bold").setBackground("#1e3a8a").setFontColor("#ffffff");
        }
      }
    }

    const targetBudget = payload.targetBudget || 0;
    const contingencyRate = payload.contingencyRate || 0;
    const markupRate = payload.markupRate || 0;

    projectSheet.appendRow([
      projectId, 
      new Date(), 
      payload.currentUser, 
      payload.projectName, 
      payload.clientEmail, 
      payload.totalBudget, 
      payload.userNote || "",
      targetBudget,
      contingencyRate,
      markupRate
    ]);

    const detailsSheet = getSheetByNameRobust(ss, "RAB_Details");
    const rows = payload.items.map(item => [projectId, item.fixtureName, item.materialName, item.unit || "ea", item.qty, item.unitPrice, item.subtotal]);
    detailsSheet.getRange(detailsSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    writeLog(payload.currentUser, payload.currentRole, "SAVE_VERSION", `Saved RAB version for ${payload.projectName}`);
    return { success: true, projectId: projectId };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteProjectArchive(projectId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projectSheet = getSheetByNameRobust(ss, "Data_Projects");
    const pData = projectSheet.getDataRange().getValues();
    let projectFound = false;
    for (let i = pData.length - 1; i >= 1; i--) {
      if (pData[i][0] === projectId) { projectSheet.deleteRow(i + 1); projectFound = true; }
    }
    const detailsSheet = getSheetByNameRobust(ss, "RAB_Details");
    const dData = detailsSheet.getDataRange().getValues();
    let detailsDeletedCount = 0;
    for (let i = dData.length - 1; i >= 1; i--) {
      if (dData[i][0] === projectId) { detailsSheet.deleteRow(i + 1); detailsDeletedCount++; }
    }
    return { success: true, projectFound: projectFound, detailsDeletedCount: detailsDeletedCount };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function sendRABEmail(payload) {
  try {
    const projectId = payload.projectId || "PRJ-TEMP";
    generateAndSendFiles(projectId, payload);
    writeLog(payload.currentUser, payload.currentRole, "SEND_EMAIL", `Sent PDF to ${payload.clientEmail}`);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function getProjectDetails(projectId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterMap = {};
    try {
      const masterData = getMaterialData();
      masterData.forEach(m => {
        if (m.Item_Name && m.Unit) masterMap[m.Item_Name.toString().trim().toLowerCase()] = m.Unit.toString().trim();
      });
    } catch(e){}
    
    const detailsSheet = getSheetByNameRobust(ss, "RAB_Details");
    const dData = detailsSheet.getDataRange().getValues();
    const items = [];
    if (dData.length > 1) {
      const headers = dData[0];
      const fixIdx = headers.indexOf("Fixture_Name") !== -1 ? headers.indexOf("Fixture_Name") : 1;
      const matIdx = headers.indexOf("Material_Name") !== -1 ? headers.indexOf("Material_Name") : 2;
      const unitIdx = headers.indexOf("Unit");
      const qtyIdx = headers.indexOf("QTY") !== -1 ? headers.indexOf("QTY") : 3;
      const priceIdx = headers.indexOf("Unit_Price") !== -1 ? headers.indexOf("Unit_Price") : 4;
      const subtotalIdx = headers.indexOf("Subtotal") !== -1 ? headers.indexOf("Subtotal") : 5;
      
      for (let i = 1; i < dData.length; i++) {
        if (dData[i][0] === projectId) {
          const matName = dData[i][matIdx] || "";
          let itemUnit = "ea";
          if (unitIdx !== -1) itemUnit = dData[i][unitIdx] || "ea";
          else {
            const key = matName.toString().trim().toLowerCase();
            if (masterMap[key]) itemUnit = masterMap[key];
          }
          items.push({ fixtureName: dData[i][fixIdx] || "", materialName: matName, unit: itemUnit, qty: parseFloat(dData[i][qtyIdx]) || 0, unitPrice: parseFloat(dData[i][priceIdx]) || 0, subtotal: parseFloat(dData[i][subtotalIdx]) || 0 });
        }
      }
    }
    
    const projectSheet = getSheetByNameRobust(ss, "Data_Projects");
    const pData = projectSheet.getDataRange().getValues();
    let projectMeta = {};
    if (pData.length > 1) {
      const headers = pData[0];
      const targetBudgetIdx = headers.indexOf("Target_Budget");
      const contingencyIdx = headers.indexOf("Contingency_Rate");
      const markupIdx = headers.indexOf("Markup_Rate");

      for (let i = 1; i < pData.length; i++) {
        if (pData[i][0] === projectId) {
          projectMeta = { 
            projectName: pData[i][3], 
            clientEmail: pData[i][4], 
            totalBudget: pData[i][5], 
            userNote: pData[i][6] || "",
            targetBudget: targetBudgetIdx !== -1 ? pData[i][targetBudgetIdx] : 0,
            contingencyRate: contingencyIdx !== -1 ? pData[i][contingencyIdx] : 0,
            markupRate: markupIdx !== -1 ? pData[i][markupIdx] : 0
          };
          break;
        }
      }
    }
    return { success: true, items: items, meta: projectMeta };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function sendRABEmailFromHistory(projectId, customEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const projDetails = getProjectDetails(projectId);
    if (!projDetails.success || !projDetails.meta.projectName) throw new Error("Metadata proyek tidak ditemukan!");
    const finalEmail = (customEmail || projDetails.meta.clientEmail || "").trim();
    if (!finalEmail) throw new Error("Email penerima kosong!");
    
    generateAndSendFiles(projectId, { projectName: projDetails.meta.projectName, clientEmail: finalEmail, totalBudget: projDetails.meta.totalBudget, items: projDetails.items });
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function generateAndSendFiles(projectId, payload) {
  const tempSS = SpreadsheetApp.create(`RAB_${payload.projectName}_${projectId}`);
  const tempSheet = tempSS.getSheets()[0]; tempSheet.setName("RAB Report");
  tempSheet.getRange("A1").setValue("RENCANA ANGGARAN BIAYA (RAB)").setFontSize(16).setFontWeight("bold");
  tempSheet.getRange("A2").setValue(`Project ID: ${projectId}`); tempSheet.getRange("A3").setValue(`Project Name: ${payload.projectName}`);
  const headers = ["Nama Item/Fixture", "Material", "Satuan", "QTY", "Harga Satuan", "Subtotal"];
  tempSheet.getRange(6, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f1f5f9");
  const reportRows = payload.items.map(item => [item.fixtureName, item.materialName, item.unit || "ea", item.qty, item.unitPrice, item.subtotal]);
  tempSheet.getRange(7, 1, reportRows.length, reportRows[0].length).setValues(reportRows);
  tempSheet.getRange(7 + reportRows.length, 5).setValue("TOTAL ANGGARAN:").setFontWeight("bold");
  tempSheet.getRange(7 + reportRows.length, 6).setValue(payload.totalBudget).setFontWeight("bold");
  tempSheet.getRange(7, 4, reportRows.length).setNumberFormat('#,##0.00');
  tempSheet.getRange(7, 5, reportRows.length + 1, 2).setNumberFormat('Rp #,##0');
  SpreadsheetApp.flush();
  MailApp.sendEmail({ to: payload.clientEmail, subject: `[RAB OUTBOUND] ${payload.projectName}`, body: `RAB Terlampir. Total: Rp ${payload.totalBudget.toLocaleString('id-ID')}`, attachments: [tempSS.getAs('application/pdf').setName(`RAB_${payload.projectName}.pdf`)] });
}

function optimizeRAB(currentItemsJson) {
  try {
    const prompt = `Diberikan daftar RAB:\n${JSON.stringify(currentItemsJson)}\nTemukan 2-3 item termahal, berikan alternatif lokal lebih murah, buat ulang daftar RAB baru dalam format JSON murni: { "explanation": "Penjelasan...", "optimizedRab": [ { "fixtureName": "Nama Bagian", "materialName": "Alternatif Material Hemat", "qty": 1.00, "unitPrice": 120000 } ] }`;
    
    const failoverResult = callAIChatFailover(prompt, "Anda adalah Ahli Quantity Surveyor (QS) senior di Indonesia.", null, true);
    if (!failoverResult.success) throw new Error(failoverResult.error);
    
    const parsedData = cleanAndParseJSON(failoverResult.text);
    return { success: true, data: parsedData, providerUsed: failoverResult.provider };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function extractJSONObjects(str) {
  const objects = []; let bracketStack = []; let inString = false; let escape = false; let startIdx = -1;
  for (let i = 0; i < str.length; i++) {
    const char = str[i];
    if (inString) {
      if (escape) escape = false; else if (char === '\\') escape = true; else if (char === '"') inString = false;
      continue;
    }
    if (char === '"') { inString = true; continue; }
    if (char === '{' || char === '[') {
      if (bracketStack.length === 0) startIdx = i;
      bracketStack.push(char);
    } else if (char === '}' || char === ']') {
      if (bracketStack.length > 0) {
        const last = bracketStack[bracketStack.length - 1];
        if ((char === '}' && last === '{') || (char === ']' && last === '[')) {
          bracketStack.pop();
          if (bracketStack.length === 0) {
            const candidate = str.substring(startIdx, i + 1);
            try { objects.push(JSON.parse(candidate)); } catch (e) {}
            startIdx = -1;
          }
        } else { bracketStack = []; startIdx = -1; }
      }
    }
  }
  return objects;
}

function cleanAndParseJSON(rawText) {
  let cleanText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();
  try { return JSON.parse(cleanText); } catch (e) {}
  let sanitized = ""; let inString = false; let escape = false;
  for (let i = 0; i < cleanText.length; i++) {
    let char = cleanText[i];
    if (char === '"' && !escape) inString = !inString;
    if (inString) {
      if (char === '\n') sanitized += '\\n'; else if (char === '\r') sanitized += '\\r'; else if (char === '\t') sanitized += '\\t'; else sanitized += char;
    } else sanitized += char;
    if (char === '\\' && !escape) escape = true; else escape = false;
  }
  const parsedObjects = extractJSONObjects(sanitized);
  if (parsedObjects.length === 0) throw new Error("Format output AI tidak valid!");
  let combinedRabItems = []; let combinedReply = ""; let explanationParts = [];
  parsedObjects.forEach(obj => {
    if (obj && typeof obj === 'object') {
      if (Array.isArray(obj.rabItems)) combinedRabItems = combinedRabItems.concat(obj.rabItems);
      else if (Array.isArray(obj)) combinedRabItems = combinedRabItems.concat(obj);
      else if (Array.isArray(obj.optimizedRab)) combinedRabItems = combinedRabItems.concat(obj.optimizedRab);
      else if (Array.isArray(obj.newRab)) combinedRabItems = combinedRabItems.concat(obj.newRab);
      else if (obj.fixtureName || obj.materialName) combinedRabItems.push(obj);
      if (obj.reply) { if (combinedReply) combinedReply += "<br><br>"; combinedReply += obj.reply; }
      if (obj.explanation) explanationParts.push(obj.explanation);
    }
  });
  const result = {};
  if (combinedRabItems.length > 0) { result.rabItems = combinedRabItems; result.optimizedRab = combinedRabItems; result.newRab = combinedRabItems; }
  if (combinedReply) result.reply = combinedReply;
  if (explanationParts.length > 0) result.explanation = explanationParts.join("<br><br>");
  parsedObjects.forEach(obj => {
    if (obj && typeof obj === 'object') {
      if (obj.price !== undefined) result.price = obj.price;
      if (obj.unit !== undefined) result.unit = obj.unit;
      if (obj.sourceUrl !== undefined) result.sourceUrl = obj.sourceUrl;
    }
  });
  return result;
}

function testCSERequest() {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GCSE_API_KEY');
    const cx = PropertiesService.getScriptProperties().getProperty('GCSE_CX');
    let logs = ["=== DIAGNOSTIK GOOGLE CSE ==="];
    if (!apiKey || !cx) return "ERROR: API KEY/CX Belum di-set!";
    const url = `https://www.googleapis.com/customsearch/v1?key=${apiKey}&cx=${cx}&q=test`;
    const res = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    logs.push("Status Code: " + res.getResponseCode());
    return logs.join("\n");
  } catch (e) { return "Error: " + e.toString(); }
}

// ==================================================
// SYSTEM VALIDASI, NORMALISASI, DAN PENYIMPANAN MASTER
// ==================================================

function getBestPrice(materialName) {
  try {
    const masterData = getMaterialData();
    const normalizedInput = materialName.toLowerCase().replace(/\s+/g, " ").trim();
    
    const foundInMaster = masterData.find(m => {
      if (!m.Item_Name) return false;
      return m.Item_Name.toLowerCase().replace(/\s+/g, " ").trim() === normalizedInput;
    });

    if (foundInMaster) {
      return {
        success: true,
        source: "MASTER",
        data: {
          actualPrice: parseFloat(foundInMaster.Unit_Price) || 0,
          sourceUrl: foundInMaster.Ref_URL || "#"
        }
      };
    }
    return researchMaterialGeminiSearch(materialName);
  } catch (e) {
    return { success: false, error: "Error di getBestPrice: " + e.toString() };
  }
}

function processResearchResult(materialName, aiResult) {
  try {
    const masterData = getMaterialData();
    const normalizedInput = materialName.toLowerCase().replace(/\s+/g, " ").trim();
    const masterItem = masterData.find(m => {
      if (!m.Item_Name) return false;
      return m.Item_Name.toLowerCase().replace(/\s+/g, " ").trim() === normalizedInput;
    });

    let finalPrice = parseFloat(aiResult.price) || 0;
    let aiUnit = aiResult.unit || "";

    if (masterItem) {
      const targetUnit = masterItem.Unit || "";
      const conversionFactor = masterItem.Conversion_Factor || 1;
      finalPrice = normalizePrice(finalPrice, aiUnit, targetUnit, conversionFactor);
    }

    return {
      success: true,
      source: "INTERNET",
      data: { actualPrice: finalPrice, sourceUrl: aiResult.sourceUrl || "#" }
    };
  } catch (e) {
    return { success: false, error: "Gagal normalisasi hasil riset: " + e.toString() };
  }
}

function normalizePrice(aiRawPrice, aiRawUnit, targetUnit, masterConversionFactor) {
  let price = parseFloat(aiRawPrice);
  let factor = parseFloat(masterConversionFactor);
  if (isNaN(price) || !factor || factor <= 0) return price;
  
  if (aiRawUnit && targetUnit && aiRawUnit.toLowerCase().trim() !== targetUnit.toLowerCase().trim()) {
     return price / factor;
  }
  return price;
}

function saveNewItemToMaster(itemData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = getSheetByNameRobust(ss, "Master_Harga");
    if (!masterSheet) throw new Error("Sheet Master_Harga tidak ditemukan!");

    const lastRow = masterSheet.getLastRow();
    let newId = "MAT001";
    if (lastRow > 1) {
      const lastIdStr = masterSheet.getRange(lastRow, 1).getValue().toString();
      const lastIdNum = parseInt(lastIdStr.replace("MAT", "")) || 0;
      newId = "MAT" + String(lastIdNum + 1).padStart(3, '0');
    }

    const rowToWrite = [
      newId,
      itemData.category || "Umum",
      itemData.itemName,
      itemData.specification || "-",
      itemData.unit || "ea",
      parseFloat(itemData.unitPrice) || 0,
      itemData.refUrl || "#",
      parseFloat(itemData.conversionFactor) || 1
    ];

    masterSheet.appendRow(rowToWrite);
    writeLog(itemData.currentUser || "Owner", "Owner", "ADD_MASTER", `Menambah ${itemData.itemName} dari Portal`);
    return { success: true, materialId: newId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function generateExcelRAB(payload) {
  try {
    const projectName = payload.projectName || "Proyek";
    const projectId = payload.projectId || "PRJ-TEMP";
    const items = payload.items || [];
    const contingencyRate = parseFloat(payload.contingencyRate) || 0;
    const markupRate = parseFloat(payload.markupRate) || 0;
    
    // Create a temporary spreadsheet in user's Drive
    const tempSS = SpreadsheetApp.create(`RAB_${projectName}_${projectId}`);
    const sheet = tempSS.getSheets()[0];
    sheet.setName("RAB Estimasi");
    
    // Turn on grid lines explicitly
    sheet.setHasGridlines(true);
    
    // Header Style
    sheet.getRange("A1").setValue("RENCANA ANGGARAN BIAYA (RAB)").setFontSize(16).setFontWeight("bold");
    sheet.getRange("A2").setValue(`Project ID: ${projectId}`).setFontSize(10).setFontStyle("italic");
    sheet.getRange("A3").setValue(`Nama Proyek: ${projectName}`).setFontSize(11).setFontWeight("bold");
    if (payload.clientEmail) {
      sheet.getRange("A4").setValue(`Klien: ${payload.clientEmail}`).setFontSize(10);
    }
    
    // Column Headers
    const headers = ["Nama Item/Fixture", "Material", "Satuan", "QTY", "Harga Satuan (Rp)", "Subtotal (Rp)"];
    sheet.getRange(6, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setBackground("#1e3a8a")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
      
    // Write Rows
    let startRow = 7;
    let dataRangeValues = [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      dataRangeValues.push([
        item.fixtureName,
        item.materialName,
        item.unit || "ea",
        parseFloat(item.qty) || 0,
        parseFloat(item.unitPrice) || 0,
        "" // Will be populated with active formula
      ]);
    }
    
    if (dataRangeValues.length > 0) {
      sheet.getRange(startRow, 1, dataRangeValues.length, dataRangeValues[0].length).setValues(dataRangeValues);
      
      // Inject Active Formulas for Subtotals: =D[Row]*E[Row]
      for (let i = 0; i < items.length; i++) {
        const curRow = startRow + i;
        sheet.getRange(curRow, 6).setFormula(`=D${curRow}*E${curRow}`);
      }
    }
    
    // Calculate total rows
    let lastDataRow = startRow + items.length - 1;
    let baseTotalRow = lastDataRow + 2;
    let contingencyRow = baseTotalRow + 1;
    let markupRow = contingencyRow + 1;
    let grandTotalRow = markupRow + 1;
    
    // Base Total
    sheet.getRange(baseTotalRow, 5).setValue("Estimasi Dasar (Subtotal):").setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(baseTotalRow, 6).setFormula(`=SUM(F${startRow}:F${lastDataRow})`).setFontWeight("bold");
    
    // Contingency
    sheet.getRange(contingencyRow, 5).setValue(`Contingency (${contingencyRate}%):`).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(contingencyRow, 6).setFormula(`=F${baseTotalRow}*${contingencyRate/100}`).setFontWeight("bold");
    
    // Markup
    sheet.getRange(markupRow, 5).setValue(`Contractor Markup (${markupRate}%):`).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(markupRow, 6).setFormula(`=(F${baseTotalRow}+F${contingencyRow})*${markupRate/100}`).setFontWeight("bold");
    
    // Grand Total
    sheet.getRange(grandTotalRow, 5).setValue("GRAND TOTAL ANGGARAN:").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("right").setBackground("#f1f5f9");
    sheet.getRange(grandTotalRow, 6).setFormula(`=F${baseTotalRow}+F${contingencyRow}+F${markupRow}`).setFontWeight("bold").setFontSize(12).setBackground("#f1f5f9");
    
    // Formats
    // QTY Format
    sheet.getRange(startRow, 4, items.length).setNumberFormat('#,##0.00');
    // Price Format
    sheet.getRange(startRow, 5, items.length, 2).setNumberFormat('Rp #,##0');
    // Summary Row Formats
    sheet.getRange(baseTotalRow, 6, 4, 1).setNumberFormat('Rp #,##0');
    
    // Set auto width for columns
    for (let c = 1; c <= headers.length; c++) {
      sheet.autoResizeColumn(c);
    }
    
    SpreadsheetApp.flush();
    
    // Get file in Drive to set permissions and get download link
    const file = DriveApp.getFileById(tempSS.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Generate the download link format for XLSX
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${tempSS.getId()}/export?format=xlsx`;
    
    return { success: true, downloadUrl: downloadUrl };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// =========================================================================
// HIGH-RESILIENCY MULTI-PROVIDER AI FAILOVER ENGINE (FREETIER MAXIMIZATION)
// urutan prioritas: Gemini -> Groq -> OpenRouter -> DeepSeek
// =========================================================================

function callAIChatFailover(prompt, systemInstruction, images, isJson, useSearchTool) {
  const props = PropertiesService.getScriptProperties();
  const keys = {
    gemini: props.getProperty('GEMINI_API_KEY'),
    groq: props.getProperty('GROQ_API_KEY'),
    openrouter: props.getProperty('OPENROUTER_API_KEY'),
    deepseek: props.getProperty('DEEPSEEK_API_KEY')
  };
  
  const providers = ['gemini', 'groq', 'openrouter', 'deepseek'];
  let lastError = "";
  
  for (let i = 0; i < providers.length; i++) {
    const provider = providers[i];
    const key = keys[provider];
    if (!key) {
      Logger.log(`Skipping provider ${provider} because API key is not configured.`);
      continue;
    }
    
    try {
      Logger.log(`Attempting AI call with provider: ${provider}...`);
      let result = null;
      
      if (provider === 'gemini') {
        result = callGeminiDirect(key, prompt, systemInstruction, images, isJson, useSearchTool);
      } else if (provider === 'groq') {
        result = callGroqDirect(key, prompt, systemInstruction, images, isJson);
      } else if (provider === 'openrouter') {
        result = callOpenRouterDirect(key, prompt, systemInstruction, images, isJson);
      } else if (provider === 'deepseek') {
        result = callDeepSeekDirect(key, prompt, systemInstruction, images, isJson);
      }
      
      if (result && result.success) {
        Logger.log(`AI call successfully answered by: ${provider}!`);
        return { success: true, text: result.text, provider: provider.toUpperCase() };
      } else {
        lastError = result ? result.error : "Unknown error";
        Logger.log(`Provider ${provider} failed: ${lastError}`);
      }
    } catch (e) {
      lastError = e.toString();
      Logger.log(`Exception during call to ${provider}: ${lastError}`);
    }
  }
  
  return { success: false, error: `Seluruh AI provider (Gemini, Groq, OpenRouter, DeepSeek) gagal merespon atau melebihi limit kuota. Error terakhir: ${lastError}` };
}

function callGeminiDirect(key, prompt, systemInstruction, images, isJson, useSearchTool) {
  let parts = [];
  if (systemInstruction) {
    parts.push({ text: "SYSTEM INSTRUCTION:\n" + systemInstruction + "\n\nUSER REQUEST:\n" + prompt });
  } else {
    parts.push({ text: prompt });
  }
  
  if (images && images.length > 0) {
    images.forEach(img => {
      const cleanData = img.base64Data.indexOf("base64,") !== -1 ? img.base64Data.split("base64,")[1] : img.base64Data;
      parts.push({
        inline_data: {
          mime_type: img.mimeType,
          data: cleanData
        }
      });
    });
  }
  
  const payload = { contents: [{ parts: parts }] };
  if (useSearchTool) {
    payload.tools = [{ google_search: {} }];
  }
  
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`;
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code !== 200) {
    return { success: false, error: `Gemini Status ${code}: ${responseText}` };
  }
  
  const json = JSON.parse(responseText);
  if (json.error) {
    return { success: false, error: json.error.message };
  }
  
  if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0]) {
    return { success: true, text: json.candidates[0].content.parts[0].text };
  }
  
  return { success: false, error: "Empty response from Gemini" };
}

function callGroqDirect(key, prompt, systemInstruction, images, isJson) {
  const hasImage = images && images.length > 0;
  const modelName = hasImage ? "llama-3.2-11b-vision-preview" : "llama-3.3-70b-versatile";
  
  let messages = [];
  if (systemInstruction) {
    messages.push({ role: "system", content: systemInstruction });
  }
  
  if (hasImage) {
    let contentParts = [{ type: "text", text: prompt }];
    images.forEach(img => {
      const cleanBase64 = img.base64Data.indexOf("base64,") !== -1 ? img.base64Data.split("base64,")[1] : img.base64Data;
      contentParts.push({
        type: "image_url",
        image_url: { url: `data:${img.mimeType};base64,${cleanBase64}` }
      });
    });
    messages.push({ role: "user", content: contentParts });
  } else {
    messages.push({ role: "user", content: prompt });
  }
  
  const payload = {
    model: modelName,
    messages: messages
  };
  if (isJson) {
    payload.response_format = { type: "json_object" };
  }
  
  const url = "https://api.groq.com/openai/v1/chat/completions";
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "Bearer " + key,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code !== 200) {
    return { success: false, error: `Groq Status ${code}: ${responseText}` };
  }
  
  const json = JSON.parse(responseText);
  if (json.error) {
    return { success: false, error: json.error.message };
  }
  
  if (json.choices && json.choices[0] && json.choices[0].message) {
    return { success: true, text: json.choices[0].message.content };
  }
  
  return { success: false, error: "Empty response from Groq" };
}

function callOpenRouterDirect(key, prompt, systemInstruction, images, isJson) {
  const hasImage = images && images.length > 0;
  const modelName = "google/gemini-2.5-flash:free";
  
  let messages = [];
  if (systemInstruction) {
    messages.push({ role: "system", content: systemInstruction });
  }
  
  if (hasImage) {
    let contentParts = [{ type: "text", text: prompt }];
    images.forEach(img => {
      const cleanBase64 = img.base64Data.indexOf("base64,") !== -1 ? img.base64Data.split("base64,")[1] : img.base64Data;
      contentParts.push({
        type: "image_url",
        image_url: { url: `data:${img.mimeType};base64,${cleanBase64}` }
      });
    });
    messages.push({ role: "user", content: contentParts });
  } else {
    messages.push({ role: "user", content: prompt });
  }
  
  const payload = {
    model: modelName,
    messages: messages
  };
  if (isJson) {
    payload.response_format = { type: "json_object" };
  }
  
  const url = "https://openrouter.ai/api/v1/chat/completions";
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "Bearer " + key,
      "Content-Type": "application/json",
      "HTTP-Referer": "https://stella-rab.vercel.app/",
      "X-Title": "Stella RAB Portal"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code !== 200) {
    return { success: false, error: `OpenRouter Status ${code}: ${responseText}` };
  }
  
  const json = JSON.parse(responseText);
  if (json.error) {
    return { success: false, error: json.error.message };
  }
  
  if (json.choices && json.choices[0] && json.choices[0].message) {
    return { success: true, text: json.choices[0].message.content };
  }
  
  return { success: false, error: "Empty response from OpenRouter" };
}

function callDeepSeekDirect(key, prompt, systemInstruction, images, isJson) {
  const hasImage = images && images.length > 0;
  let fullPrompt = prompt;
  if (hasImage) {
    fullPrompt = "[Notice: Image analysis requested but DeepSeek fallback is text-only. Processing text prompt.]\n" + prompt;
  }
  
  let messages = [];
  if (systemInstruction) {
    messages.push({ role: "system", content: systemInstruction });
  }
  messages.push({ role: "user", content: fullPrompt });
  
  const payload = {
    model: "deepseek-chat",
    messages: messages
  };
  if (isJson) {
    payload.response_format = { type: "json_object" };
  }
  
  const url = "https://api.deepseek.com/chat/completions";
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "Bearer " + key,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  const code = res.getResponseCode();
  const responseText = res.getContentText();
  
  if (code !== 200) {
    return { success: false, error: `DeepSeek Status ${code}: ${responseText}` };
  }
  
  const json = JSON.parse(responseText);
  if (json.error) {
    return { success: false, error: json.error.message };
  }
  
  if (json.choices && json.choices[0] && json.choices[0].message) {
    return { success: true, text: json.choices[0].message.content };
  }
  
  return { success: false, error: "Empty response from DeepSeek" };
}