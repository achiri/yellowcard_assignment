function createYellowCardCollection() {
  const credentialsSheetId = "1Ruu8fsh-Q0RV3FD-zdfBovCCz56cyQ1R-HF0QoTUYgc";
  const credentialsSheetName = "credentials";
  const dataSheetId = '1Ruu8fsh-Q0RV3FD-zdfBovCCz56cyQ1R-HF0QoTUYgc';
  const dataSheetName = 'networkId';
  const formResponsesSheetId = "158zf0cYD0KWU3nRDDsORc-OLj1jfTpeQsyMywkT9hPU";
  const formResponsesSheetName = "Form Responses 1";

  // Fetch credentials from the sheet
  const credentialsSheet = SpreadsheetApp.openById(credentialsSheetId).getSheetByName(credentialsSheetName);
  const credentialsData = credentialsSheet.getRange("B1:B2").getValues();

  // Check if credentials are available
  if (credentialsData.length < 2 || !credentialsData[0][0] || !credentialsData[1][0]) {
    throw new Error("Missing or invalid API key or secret key in the credentials sheet.");
  }

  const secretKey = credentialsData[0][0];
  const apiKey = credentialsData[1][0];

  // Fetch the last row of the form responses sheet
  const formResponsesSheet = SpreadsheetApp.openById(formResponsesSheetId).getSheetByName(formResponsesSheetName);
  const lastRow = formResponsesSheet.getLastRow();
  const formData = formResponsesSheet.getRange(lastRow, 2, 1, 7).getValues()[0]; // Columns B to H

  const [
    name,         // Column B
    address,      // Column C
    dob,          // Column D
    email,        // Column E
    idNumber,     // Column F
    localAmount,  // Column G
    phone         // Column H
  ] = formData;

  // Ensure the localAmount is an integer
  const localAmountInt = parseInt(localAmount, 10);

  // 1. Get Current Timestamp (ISO8601)
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");

  // 2. Construct the Message to Sign
  const method = "POST";
  const path = "/business/collections";

  // Updated Body with data from the sheet
  const body = JSON.stringify({
    "channelId": "79da4d6e-1c42-4aac-ae7d-422730528f96",
    "sequenceId": Utilities.getUuid(),
    "localAmount": localAmountInt,
    "reason": "bill",
    "forceAccept": true,
    "recipient": {
      "name": String(name),
      "country": "CMR",
      "address": String(address),
      "dob": String(dob),
      "email": String(email),
      "idNumber": String(idNumber),
      "idType": "National ID",
      "phone": String(phone)
    },
    "source": {
      "accountType": "momo",
      "accountNumber": String(phone),
      "networkId": "cc2883ed-e431-444d-9264-8b7c1684b998"
    }
  });

  // Calculate Base64-encoded SHA256 hash of the body
  const bodyHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, body));

  // 3. Calculate the HMAC Signature 
  const message = timestamp + path + method + bodyHash;
  const signatureBytes = Utilities.computeHmacSha256Signature(message, secretKey);
  const signature = Utilities.base64Encode(signatureBytes);

  // 4. & 5. Create Authorization Header 
  const authorization = "YcHmacV1 " + apiKey + ":" + signature;

  // 6. Make the Fetch Request and Log Data
  const url = "https://sandbox.api.yellowcard.io" + path;
  const options = {
    "method": method,
    "headers": {
      "accept": "application/json",
      "Authorization": authorization,
      "X-YC-Timestamp": timestamp,
      'Content-Type': 'application/json'
    },
    "payload": body
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    const sheet = SpreadsheetApp.openById(dataSheetId).getSheetByName(dataSheetName);

    // Flatten the nested recipient and source objects
    const flatData = {
      ...data, 
      ...data.recipient, 
      ...data.source,
      ...data.customer
    };

    delete flatData.recipient;
    delete flatData.source;
    delete flatData.customer;

    // Prepare headers and values for logging
    const headers = Object.keys(flatData);
    const valuesToLog = [Object.values(flatData)];

    // Check if headers already exist
    const lastDataRow = sheet.getLastRow(); 
    if (lastDataRow === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]); 
    }

    // Append values to the next row 
    sheet.getRange(lastDataRow + 1, 1, valuesToLog.length, valuesToLog[0].length).setValues(valuesToLog);

  } catch (err) {
    console.error("Error creating collection or logging to sheet:", err.message); 
  }
}
