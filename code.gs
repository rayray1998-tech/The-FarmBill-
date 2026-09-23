/**
 * Serves the initial HTML page.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Ghibli Farm Sim - Final Phase')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Helper to get or create the main FarmData sheet.
 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('FarmData');
  if (!sheet) {
    sheet = ss.insertSheet('FarmData');
    sheet.appendRow(['UserID', 'FarmStateJSON', 'Level', 'Coins']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Helper to get or create the Market sheet for Flea Market listings.
 */
function getMarketSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Market');
  if (!sheet) {
    sheet = ss.insertSheet('Market');
    sheet.appendRow(['ListingID', 'Timestamp', 'SellerEmail', 'SellerName', 'Item', 'Price', 'Description', 'Status']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Fetches the active user's Google Account Email.
 */
function getInitialPlayerData() {
  try {
    let userEmail = Session.getActiveUser().getEmail();
    if (!userEmail || userEmail === "") {
      userEmail = "guest_" + Math.floor(Math.random() * 1000) + "@example.com";
    }
    return { email: userEmail };
  } catch (e) {
    return { email: "guest_" + Math.floor(Math.random() * 1000) + "@example.com" };
  }
}

/**
 * Fetches the active user's single-cell JSON state.
 */
function getFarmState(userId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) return data[i][1];
  }
  return null;
}

/**
 * Fetches a friend's farm state for read-only visiting.
 */
function getFriendFarmState(friendId) {
  return getFarmState(friendId);
}

/**
 * Fetches an array of all registered users (excluding the active player).
 */
function getNeighborList(currentUserId) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let neighbors = [];
  for (let i = 1; i < data.length; i++) {
    const dbUserId = data[i][0];
    if (dbUserId && dbUserId !== currentUserId) {
      neighbors.push(dbUserId);
    }
  }
  return neighbors;
}

/**
 * Updates a user's single-cell JSON state, or creates a new row.
 */
function saveFarmState(userId, stateJSON) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  
  let parsed = {};
  let level = 1;
  let coins = 0;
  try {
    parsed = JSON.parse(stateJSON);
    if (parsed.stats) {
      level = parsed.stats.level || 1;
      coins = parsed.stats.coins || 0;
    }
  } catch(e) {}

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      sheet.getRange(i + 1, 2).setValue(stateJSON);
      sheet.getRange(i + 1, 3).setValue(level);
      sheet.getRange(i + 1, 4).setValue(coins);
      return { success: true, message: "Saved" };
    }
  }
  sheet.appendRow([userId, stateJSON, level, coins]);
  return { success: true, message: "Created and Saved" };
}

/**
 * Reads the sheet, sorts by Level (Desc) then Coins (Desc), and returns the Top 10.
 */
function getLeaderboard() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let players = [];
  
  for (let i = 1; i < data.length; i++) {
    const email = data[i][0];
    if (!email) continue;
    
    let username = email.split('@')[0];
    let level = parseInt(data[i][2]) || 1;
    let coins = parseInt(data[i][3]) || 0;
    
    try {
      const parsed = JSON.parse(data[i][1]);
      if (parsed && parsed.farmName && parsed.farmName !== "My Farm" && parsed.farmName !== "Neighbor's Farm") {
        username = parsed.farmName;
      }
    } catch(e) {}
    
    players.push({ username: username, level: level, coins: coins });
  }
  
  players.sort((a, b) => {
    if (b.level !== a.level) return b.level - a.level;
    return b.coins - a.coins;
  });
  
  return players.slice(0, 10);
}

/**
 * Advanced Flea Market: Lists an item on the market
 */
function listMarketItem(playerEmail, sellerName, itemName, price, description) {
  const sheet = getMarketSheet();
  const id = "MKT_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
  const timestamp = new Date().toISOString();
  sheet.appendRow([id, timestamp, playerEmail, sellerName, itemName, price, description, "Active"]);
  return { success: true, id: id };
}

/**
 * Advanced Flea Market: Gets active listings
 */
function getMarketListings() {
  const sheet = getMarketSheet();
  const data = sheet.getDataRange().getValues();
  let listings = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === "Active") {
      listings.push({
        id: data[i][0],
        timestamp: data[i][1],
        sellerEmail: data[i][2],
        sellerName: data[i][3],
        item: data[i][4],
        price: Number(data[i][5]),
        description: data[i][6]
      });
    }
  }
  
  // Sort newest first
  listings.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return listings;
}

/**
 * Advanced Flea Market: Buy an item, transfer coins to seller in the background
 */
function buyMarketListing(buyerEmail, listingId) {
  const sheet = getMarketSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === listingId && data[i][7] === "Active") {
      if (data[i][2] === buyerEmail) {
        return { success: false, message: "Cannot buy your own listing." };
      }
      
      let itemStr = data[i][4];
      let price = Number(data[i][5]);
      let sellerEmail = data[i][2];
      
      // Mark as sold
      sheet.getRange(i + 1, 8).setValue("Sold");
      
      // Credit the seller's account directly in the FarmData JSON
      try {
        const farmSheet = getSheet();
        const farmData = farmSheet.getDataRange().getValues();
        for (let j = 1; j < farmData.length; j++) {
          if (farmData[j][0] === sellerEmail) {
            let sellerState = JSON.parse(farmData[j][1]);
            if (!sellerState.stats) sellerState.stats = { coins: 0 };
            sellerState.stats.coins += price;
            
            // Save updated state and coin column
            farmSheet.getRange(j + 1, 2).setValue(JSON.stringify(sellerState));
            farmSheet.getRange(j + 1, 4).setValue(sellerState.stats.coins);
            break;
          }
        }
      } catch (e) {
        // Failsafe: if seller not found or JSON breaks, the item is still bought
      }
      
      return { success: true, item: itemStr };
    }
  }
  return { success: false, message: "Listing not found or already sold." };
}
