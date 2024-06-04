function getYellowCardCollection() {
  // Credentials (Replace with your actual values)
  const credentialsSheetId = "1Ruu8fsh-Q0RV3FD-zdfBovCCz56cyQ1R-HF0QoTUYgc";
  const credentialsSheetName = "credentials";
  const dataSheetId = '1Ruu8fsh-Q0RV3FD-zdfBovCCz56cyQ1R-HF0QoTUYgc';
  const dataSheetName = 'networkId';

  // Fetch credentials from the sheet
  const credentialsSheet = SpreadsheetApp.openById(credentialsSheetId).getSheetByName(credentialsSheetName);
  const credentialsData = credentialsSheet.getRange("B1:B2").getValues();

  // Check if credentials are available
  if (credentialsData.length < 2 || !credentialsData[0][0] || !credentialsData[1][0]) {
    throw new Error("Missing or invalid API key or secret key in the credentials sheet.");
  }

  const secretKey = credentialsData[0][0];
  const apiKey = credentialsData[1][0];

  // Fetch the collection ID from the sheet
  const sheet = SpreadsheetApp.openById(dataSheetId).getSheetByName(dataSheetName);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const collectionIdIndex = headers.indexOf("id");
  const statusIndex = headers.indexOf("status");

  if (collectionIdIndex === -1 || statusIndex === -1) {
    throw new Error("Collection ID or Status column not found in the sheet.");
  }

  // Iterate over rows to update status
  for (let i = 1; i < values.length; i++) {
    const collectionId = values[i][collectionIdIndex];
    const timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");

    // Construct the Message to Sign
    const method = "GET";
    const path = "/business/collections/" + collectionId;
    const message = timestamp + path + method;

    // Calculate the HMAC Signature
    const signatureBytes = Utilities.computeHmacSha256Signature(message, secretKey);
    const signature = Utilities.base64Encode(signatureBytes);

    // Create Authorization Header
    const authorization = "YcHmacV1 " + apiKey + ":" + signature;

    // Make the Fetch Request
    const url = "https://sandbox.api.yellowcard.io" + path;
    const options = {
      "method": method,
      "headers": {
        "accept": "application/json",
        "Authorization": authorization,
        "X-YC-Timestamp": timestamp
      }
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const collectionData = JSON.parse(response.getContentText());
      const status = collectionData.status; // Assuming status is returned in the response

      // Update status in the sheet
      const rowToUpdate = i + 1; // Adjusting for 1-based indexing
      sheet.getRange(rowToUpdate, statusIndex + 1).setValue(status); // Adjusting for 1-based indexing

    } catch (err) {
      console.error("Error fetching collection details:", err);
    }
  }
}
