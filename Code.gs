function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Lead Generation')
      .addItem('Import Free Clinic Leads', 'fetchFreeLeads')
      .addToUi();
}

// Set up the GitHub pages URL for your demo website
var GITHUB_PAGES_URL = 'https://boostrnetwave.github.io/antigravity/';

function fetchFreeLeads() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // We use OpenStreetMap's free Overpass API instead of Google Maps. No API Key Needed!
  var queries = [
    { sheetName: "Dental Clinics", category: "Dental Clinic", overpass: 'nwr["amenity"="dentist"]' },
    { sheetName: "Medical Centres", category: "Medical Centre", overpass: 'nwr["amenity"="clinic"]' },
    { sheetName: "Cosmetic Clinics", category: "Cosmetic Clinic", overpass: 'nwr["healthcare"="clinic"]' },
    { sheetName: "Physiotherapy Clinics", category: "Physiotherapy Clinic", overpass: 'nwr["healthcare"="physiotherapist"]' }
  ];

  var errorOccurred = false;

  for (var i = 0; i < queries.length; i++) {
    var queryObj = queries[i];
    var sheet = ss.getSheetByName(queryObj.sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(queryObj.sheetName);
    } else {
      sheet.clear();
    }

    var headers = ["Business Name", "Address", "Phone Number", "Website", "Demo Website Link", "Send WhatsApp"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    // Query OpenStreetMap for Australia specifically (ISO3166-1="AU")
    // We request up to 100 results per category to avoid timeout limits
    var overpassQuery = '[out:json][timeout:25];area["ISO3166-1"="AU"][admin_level=2]->.searchArea;(' + queryObj.overpass + '(area.searchArea););out center 100;';
    var url = "https://overpass-api.de/api/interpreter?data=" + encodeURIComponent(overpassQuery);

    try {
      var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      var responseCode = response.getResponseCode();
      
      if (responseCode !== 200) {
        ui.alert("API Error", "The free leads API is currently busy or returned an error. Please try again in a few minutes.", ui.ButtonSet.OK);
        errorOccurred = true;
        break;
      }
      
      var json = JSON.parse(response.getContentText());
      var elements = json.elements;
      
      if (!elements || elements.length === 0) continue;

      var rows = [];
      
      for (var j = 0; j < elements.length; j++) {
        var el = elements[j];
        if (!el.tags || !el.tags.name) continue; // Skip businesses without a name
        
        var name = el.tags.name || "";
        var phone = el.tags.phone || el.tags["contact:phone"] || "";
        var website = el.tags.website || el.tags["contact:website"] || "";
        
        // Construct a basic address from tags if available
        var street = el.tags["addr:street"] || "";
        var housenumber = el.tags["addr:housenumber"] || "";
        var city = el.tags["addr:city"] || "";
        var postcode = el.tags["addr:postcode"] || "";
        
        var address = [housenumber, street, city, postcode].filter(Boolean).join(" ");
        if (!address) address = "Australia"; // fallback

        // Build Demo Link
        var demoLinkUrl = GITHUB_PAGES_URL + "?name=" + encodeURIComponent(name) + 
                          "&address=" + encodeURIComponent(address) + 
                          "&phone=" + encodeURIComponent(phone) + 
                          "&category=" + encodeURIComponent(queryObj.category);
                          
        var demoLinkFormula = '=HYPERLINK("' + demoLinkUrl + '", "View Demo Website")';

        // Build WhatsApp Link
        var cleanPhone = phone.replace(/[^0-9]/g, "");
        var greetingMessage = "Hello " + name + ", I noticed your clinic and we created a custom demo website for you! Check it out here: " + demoLinkUrl;
        var whatsappUrl = "https://wa.me/" + cleanPhone + "?text=" + encodeURIComponent(greetingMessage);
        
        var whatsappFormula = cleanPhone ? '=HYPERLINK("' + whatsappUrl + '", "Send WhatsApp")' : "No Phone Available";

        rows.push([name, address, phone, website, demoLinkFormula, whatsappFormula]);
      }

      // Write data to sheet
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }
      sheet.autoResizeColumns(1, headers.length);
      
    } catch (e) {
      ui.alert("Script Error", "An error occurred: " + e.toString(), ui.ButtonSet.OK);
      errorOccurred = true;
      break;
    }
    
    // Sleep briefly to be respectful to the free API server
    Utilities.sleep(1000); 
  }
  
  if (!errorOccurred) {
    ui.alert("Success", "Free leads imported successfully!", ui.ButtonSet.OK);
  }
}
