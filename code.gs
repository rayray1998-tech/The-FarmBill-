function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Hundred Acre Wood Farm')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function loadPixelFarmState(userEmail) {
  return callPixelFarmRpc_('load_pixel_farm_state', {
    p_userid: normalizePixelFarmUserId_(userEmail)
  });
}

function savePixelFarmState(userEmail, pixelFarmStateData) {
  const normalizedUserId = normalizePixelFarmUserId_(userEmail);
  validatePixelFarmState_(pixelFarmStateData);
  return callPixelFarmRpc_('save_pixel_farm_state', {
    p_userid: normalizedUserId,
    p_pixel_farm_data: pixelFarmStateData
  });
}

function normalizePixelFarmUserId_(userId) {
  const normalizedUserId = String(userId || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizedUserId)) {
    throw new Error('A valid farm account email is required.');
  }
  return normalizedUserId;
}

function validatePixelFarmState_(state) {
  if (!state || typeof state !== 'object' || Array.isArray(state)) {
    throw new Error('Invalid pixel-farm save.');
  }
  if (state.schemaVersion !== 1 || !Number.isInteger(state.gridSize) ||
      state.gridSize < 12 || state.gridSize > 48 ||
      !Array.isArray(state.grid) || state.grid.length !== state.gridSize ||
      !Array.isArray(state.crops) || !Array.isArray(state.buildings) ||
      !Array.isArray(state.decorations)) {
    throw new Error('Pixel-farm save has an invalid shape.');
  }
  if (state.grid.some(function(row) {
    return !Array.isArray(row) || row.length !== state.gridSize;
  })) {
    throw new Error('Pixel-farm grid dimensions do not match.');
  }
  if (JSON.stringify(state).length > 250000) {
    throw new Error('Pixel-farm save exceeds the size limit.');
  }
}

function callPixelFarmRpc_(functionName, payload) {
  const properties = PropertiesService.getScriptProperties();
  const supabaseUrl = String(properties.getProperty('SUPABASE_URL') || '').replace(/\/$/, '');
  const supabaseKey = properties.getProperty('SUPABASE_ANON_KEY');
  if (!supabaseUrl || !supabaseKey) {
    throw new Error('Supabase Script Properties are not configured.');
  }

  const response = UrlFetchApp.fetch(supabaseUrl + '/rest/v1/rpc/' + functionName, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      apikey: supabaseKey,
      Authorization: 'Bearer ' + supabaseKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const status = response.getResponseCode();
  const responseText = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error('Supabase pixel-farm request failed (' + status + ').');
  }
  return responseText ? JSON.parse(responseText) : null;
}
