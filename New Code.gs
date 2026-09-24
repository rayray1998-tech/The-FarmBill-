/**
 * Serves the main web application HTML page.
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
 * Helper to get or create the Masterlist sheet for authorized user logins.
 */
function getMasterlistSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Masterlist');
  if (!sheet) {
    sheet = ss.insertSheet('Masterlist');
    sheet.appendRow(['Email', 'Status', 'AddedDate']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Verifies if an email is present and authorized in the Masterlist tab.
 */
function verifyMasterlistEmail(email) {
  if (!email) return { allowed: false, message: "Email address is required." };
  
  const cleanEmail = email.trim().toLowerCase();
  const sheet = getMasterlistSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    let activeUser = Session.getActiveUser().getEmail();
    if (!activeUser || activeUser === "") {
      activeUser = cleanEmail;
    } else {
      activeUser = activeUser.trim().toLowerCase();
    }
    sheet.appendRow([activeUser, "Active", new Date().toISOString()]);
    if (cleanEmail !== activeUser) {
      sheet.appendRow([cleanEmail, "Active", new Date().toISOString()]);
    }
    return { allowed: true, email: cleanEmail };
  }
  
  for (let i = 1; i < data.length; i++) {
    const rowEmail = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    if (rowEmail === cleanEmail) {
      return { allowed: true, email: cleanEmail };
    }
  }
  
  return { allowed: false, message: "Email not authorized in Masterlist" };
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
  const targetId = userId ? userId.trim().toLowerCase() : "";
  for (let i = 1; i < data.length; i++) {
    const dbId = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    if (dbId === targetId) return data[i][1];
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
  const cleanCurrent = currentUserId ? currentUserId.trim().toLowerCase() : "";
  let neighbors = [];
  for (let i = 1; i < data.length; i++) {
    const dbUserId = data[i][0] ? data[i][0].toString().trim() : "";
    if (dbUserId && dbUserId.toLowerCase() !== cleanCurrent) {
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
  const cleanUserId = userId ? userId.trim().toLowerCase() : "";
  
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
    const dbId = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    if (dbId === cleanUserId) {
      sheet.getRange(i + 1, 2).setValue(stateJSON);
      sheet.getRange(i + 1, 3).setValue(level);
      sheet.getRange(i + 1, 4).setValue(coins);
      return { success: true, message: "Saved" };
    }
  }
  sheet.appendRow([cleanUserId, stateJSON, level, coins]);
  return { success: true, message: "Created and Saved" };
}

/**
 * Reads the sheet, sorts by Level (Desc), House Tier (Desc), then Coins (Desc), and returns the Top 10.
 */
function getLeaderboard() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let players = [];
  
  for (let i = 1; i < data.length; i++) {
    const email = data[i][0] ? data[i][0].toString().trim() : "";
    if (!email) continue;
    
    let username = email.split('@')[0];
    let level = parseInt(data[i][2]) || 1;
    let coins = parseInt(data[i][3]) || 0;
    let houseTier = 1;
    
    try {
      const parsed = JSON.parse(data[i][1]);
      if (parsed && parsed.farmName && parsed.farmName !== "My Farm" && parsed.farmName !== "Neighbor's Farm") {
        username = parsed.farmName;
      }
      if (parsed && parsed.houseTier) houseTier = parsed.houseTier;
      if (parsed && parsed.stats && parsed.stats.level) level = parsed.stats.level;
    } catch(e) {}
    
    players.push({ username: username, level: level, coins: coins, houseTier: houseTier });
  }
  
  players.sort((a, b) => {
    if (b.level !== a.level) return b.level - a.level;
    if (b.houseTier !== a.houseTier) return b.houseTier - a.houseTier;
    return b.coins - a.coins;
  });
  
  return players.slice(0, 10);
}

/**
 * Advanced Flea Market: Lists an item or animal on the market
 */
function listMarketItem(playerEmail, sellerName, itemName, price, description) {
  const sheet = getMarketSheet();
  const id = "MKT_" + Date.now() + "_" + Math.floor(Math.random() * 1000);
  const timestamp = new Date().toISOString();
  sheet.appendRow([id, timestamp, playerEmail.trim().toLowerCase(), sellerName, itemName, price, description, "Active"]);
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
  listings.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  return listings;
}

/**
 * Advanced Flea Market: Buy an item/animal
 */
function buyMarketListing(buyerEmail, listingId) {
  const sheet = getMarketSheet();
  const data = sheet.getDataRange().getValues();
  const cleanBuyer = buyerEmail ? buyerEmail.trim().toLowerCase() : "";
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === listingId && data[i][7] === "Active") {
      const sellerEmail = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
      if (sellerEmail === cleanBuyer) return { success: false, message: "Cannot buy your own listing." };
      
      let itemStr = data[i][4];
      let price = Number(data[i][5]);
      
      sheet.getRange(i + 1, 8).setValue("Sold");
      
      try {
        const farmSheet = getSheet();
        const farmData = farmSheet.getDataRange().getValues();
        for (let j = 1; j < farmData.length; j++) {
          const dbSeller = farmData[j][0] ? farmData[j][0].toString().trim().toLowerCase() : "";
          if (dbSeller === sellerEmail) {
            let sellerState = JSON.parse(farmData[j][1]);
            if (!sellerState.stats) sellerState.stats = { coins: 0 };
            sellerState.stats.coins += price;
            farmSheet.getRange(j + 1, 2).setValue(JSON.stringify(sellerState));
            farmSheet.getRange(j + 1, 4).setValue(sellerState.stats.coins);
            break;
          }
        }
      } catch (e) {}
      
      return { success: true, item: itemStr };
    }
  }
  return { success: false, message: "Listing not found or already sold." };
}

/**
 * Mailbox System: Leave a gift for a neighbor
 */
function leaveGift(targetUserId, senderName) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const cleanTarget = targetUserId ? targetUserId.trim().toLowerCase() : "";
  
  for (let i = 1; i < data.length; i++) {
    const dbId = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
    if (dbId === cleanTarget) {
      try {
        let state = JSON.parse(data[i][1]);
        if (!state.stats) state.stats = { coins: 0 };
        if (!state.stats.mailbox) state.stats.mailbox = [];
        
        if (state.stats.mailbox.length < 10) {
          state.stats.mailbox.push({ from: senderName, coins: 50, date: new Date().toISOString() });
          sheet.getRange(i + 1, 2).setValue(JSON.stringify(state));
          return { success: true, message: "Gift left in mailbox!" };
        } else {
          return { success: false, message: "Their mailbox is full!" };
        }
      } catch(e) {
        return { success: false, message: "Error processing gift." };
      }
    }
  }
  return { success: false, message: "Neighbor not found." };
}
