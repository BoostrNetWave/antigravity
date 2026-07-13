function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Lead Generation')
      .addItem('Import Clinic Leads', 'fetchAndImportLeads')
      .addToUi();
}

// NOTE: You will need a Google Maps API Key to fetch live data.
// Replace 'YOUR_API_KEY_HERE' with your actual API key.
var API_KEY = 'YOUR_API_KEY_HERE';

// Set up the GitHub pages URL (Update username if different)
var GITHUB_PAGES_URL = 'https://boostrnetwave.github.io/antigravity/';

function fetchAndImportLeads() {
  var ui = SpreadsheetApp.getUi();
  if (API_KEY === 'YOUR_API_KEY_HERE') {
    ui.alert("API Key Missing", "Please add your Google Maps API key in the Apps Script Code.gs file.", ui.ButtonSet.OK);
    return;
  }

  var queries = [
    { query: "Dental Clinics in Australia", sheetName: "Dental Clinics", category: "Dental Clinic" },
    { query: "Medical Centres in Australia", sheetName: "Medical Centres", category: "Medical Centre" },
    { query: "Cosmetic Clinics in Australia", sheetName: "Cosmetic Clinics", category: "Cosmetic Clinic" },
    { query: "Physiotherapy Clinics in Australia", sheetName: "Physiotherapy Clinics", category: "Physiotherapy Clinic" }
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var i = 0; i < queries.length; i++) {
    var queryObj = queries[i];
    var sheet = ss.getSheetByName(queryObj.sheetName);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(queryObj.sheetName);
    } else {
      sheet.clear(); // Clear existing data to refresh
    }

    // Set headers
    var headers = [
      "Business Name", 
      "Address", 
      "Rating", 
      "Phone Number", 
      "Website", 
      "Demo Website Link", 
      "Send WhatsApp"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    // Fetch data from Google Places API
    var url = "https://maps.googleapis.com/maps/api/place/textsearch/json?query=" + encodeURIComponent(queryObj.query) + "&key=" + API_KEY;
    
    try {
      var response = UrlFetchApp.fetch(url);
      var json = JSON.parse(response.getContentText());
      var results = json.results;
      
      if (!results || results.length === 0) continue;

      var rows = [];
      
      for (var j = 0; j < results.length; j++) {
        var place = results[j];
        
        // We need place details to get the phone number and website
        var detailsUrl = "https://maps.googleapis.com/maps/api/place/details/json?place_id=" + place.place_id + "&fields=name,formatted_address,rating,formatted_phone_number,website,photos&key=" + API_KEY;
        Utilities.sleep(100); // To prevent hitting rate limits
        
        var detailsResponse = UrlFetchApp.fetch(detailsUrl);
        var detailsJson = JSON.parse(detailsResponse.getContentText());
        var details = detailsJson.result;

        if (!details) continue;

        var name = details.name || "";
        var address = details.formatted_address || "";
        var rating = details.rating || "";
        var phone = details.formatted_phone_number || "";
        var website = details.website || "";
        
        // Build the Demo Website Link
        // We pass the name, address, phone, category, and rating to our dynamic GitHub Pages site
        var demoLinkUrl = GITHUB_PAGES_URL + "?name=" + encodeURIComponent(name) + 
                          "&address=" + encodeURIComponent(address) + 
                          "&phone=" + encodeURIComponent(phone) + 
                          "&category=" + encodeURIComponent(queryObj.category) +
                          "&rating=" + encodeURIComponent(rating);
                          
        // Optionally attach a photo reference if available
        if (details.photos && details.photos.length > 0) {
          var photoUrl = "https://maps.googleapis.com/maps/api/place/photo?maxwidth=800&photoreference=" + details.photos[0].photo_reference + "&key=" + API_KEY;
          demoLinkUrl += "&photo=" + encodeURIComponent(photoUrl);
        }

        var demoLinkFormula = '=HYPERLINK("' + demoLinkUrl + '", "View Demo Website")';

        // Build the WhatsApp Greeting Message
        var greetingMessage = "Hello " + name + ", I noticed your clinic and we created a custom demo website for you! Check it out here: " + demoLinkUrl;
        
        // Build the WhatsApp Link Formula
        // 1. Remove non-numeric characters from the phone number
        // 2. Add https://wa.me/NUMBER?text=ENCODED_MESSAGE
        var cleanPhone = phone.replace(/[^0-9]/g, "");
        var whatsappUrl = "https://wa.me/" + cleanPhone + "?text=" + encodeURIComponent(greetingMessage);
        
        var whatsappFormula = "";
        if (cleanPhone) {
          whatsappFormula = '=HYPERLINK("' + whatsappUrl + '", "Send WhatsApp")';
        } else {
          whatsappFormula = "No Phone Available";
        }

        rows.push([
          name,
          address,
          rating,
          phone,
          website,
          demoLinkFormula,
          whatsappFormula
        ]);
      }

      // Write data to sheet
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }
      
      // Auto-resize columns for better readability
      sheet.autoResizeColumns(1, headers.length);
      
    } catch (e) {
      Logger.log("Error fetching " + queryObj.query + ": " + e.toString());
    }
  }
  
  ui.alert("Success", "Leads imported successfully!", ui.ButtonSet.OK);
}
