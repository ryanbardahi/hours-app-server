import express from "express";
import fetch from "node-fetch";
import cors from "cors";
import dotenv from "dotenv";
import { google } from "googleapis";

dotenv.config();

const app = express();
const PORT = 4000;

const corsOptions = {
  origin: ["http://localhost:3000", "https://hours-app-jade.vercel.app"],
  credentials: true,
  optionsSuccessStatus: 200,
};

app.use(cors(corsOptions));
app.use(express.json());

// Google Sheets Authentication
const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEYFILE),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });

// Helper function to apply formatting
const applyFormatting = async (sheets, spreadsheetId, sheetId, data) => {
  const requests = [];

  // 1. Format Column A as Text
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      cell: {
        userEnteredFormat: {
          numberFormat: {
            type: 'TEXT',
          },
        },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // 2. Format Cell A1
  requests.push({
    updateCells: {
      rows: [{
        values: [{
          userEnteredValue: { stringValue: 'DETAILED REPORT' },
          userEnteredFormat: {
            textFormat: {
              fontSize: 22,
              bold: true,
              foregroundColor: { red: 59/255, green: 143/255, blue: 194/255 }, // #3B8FC2
            },
          },
        }],
      }],
      fields: 'userEnteredValue,userEnteredFormat.textFormat',
      start: { sheetId: sheetId, rowIndex: 0, columnIndex: 0 },
    },
  });

  // 3. Format Cells A2 to A5
  const labels = [
    'Time frame',
    'Billable amount (hours)',
    'Total billable amount',
    'Total hours',
  ];

  const formattedRows = labels.map((label, index) => ({
    values: [{
      userEnteredValue: { stringValue: label },
      userEnteredFormat: {
        textFormat: {
          fontSize: 10,
          bold: false,
          // Default font color (black)
        },
      },
    }],
  }));

  requests.push({
    updateCells: {
      rows: formattedRows,
      fields: 'userEnteredValue,userEnteredFormat.textFormat',
      start: { sheetId: sheetId, rowIndex: 1, columnIndex: 0 }, // A2 starts at rowIndex 1
    },
  });

  // 5. Set Headers in Row 8
  const headers = [
    'USER', 'CLIENT', 'PROJECT', 'TASK', 'IS BILLABLE',
    'BILLABLE AMOUNT', 'START/FINISH TIME', 'TOTAL HOURS',
    'BILLABLE HOURS', 'DESCRIPTION'
  ];

  const headerValues = headers.map(header => ({
    userEnteredValue: { stringValue: header },
    userEnteredFormat: {
      textFormat: {
        fontSize: 10,
        bold: false,
      },
    },
  }));

  requests.push({
    updateCells: {
      rows: [{
        values: headerValues,
      }],
      fields: 'userEnteredValue,userEnteredFormat.textFormat',
      start: { sheetId: sheetId, rowIndex: 7, columnIndex: 0 }, // Row 8
    },
  });

  // 6. Set Column Widths
  // Column A
  requests.push({
    updateDimensionProperties: {
      range: {
        sheetId: sheetId,
        dimension: 'COLUMNS',
        startIndex: 0,
        endIndex: 1,
      },
      properties: {
        pixelSize: 250,
      },
      fields: 'pixelSize',
    },
  });

  // Columns B to I
  for (let col = 1; col <= 8; col++) {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId: sheetId,
          dimension: 'COLUMNS',
          startIndex: col,
          endIndex: col + 1,
        },
        properties: {
          pixelSize: 150,
        },
        fields: 'pixelSize',
      },
    });
  }

  // Column J
  requests.push({
    updateDimensionProperties: {
      range: {
        sheetId: sheetId,
        dimension: 'COLUMNS',
        startIndex: 9,
        endIndex: 10,
      },
      properties: {
        pixelSize: 500,
      },
      fields: 'pixelSize',
    },
  });

  // 7. Wrap Text for All Columns
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startRowIndex: 0,
        endRowIndex: 1000, // Adjust as needed
        startColumnIndex: 0,
        endColumnIndex: 10,
      },
      cell: {
        userEnteredFormat: {
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat.wrapStrategy',
    },
  });

  // 8. Write Date Range to B2
  requests.push({
    updateCells: {
      rows: [{
        values: [{
          userEnteredValue: { stringValue: data.dateRange },
          userEnteredFormat: {
            textFormat: {
              fontSize: 12,
              bold: false,
              foregroundColor: { red: 55/255, green: 93/255, blue: 117/255 }, // #375D75
            },
          },
        }],
      }],
      fields: 'userEnteredValue,userEnteredFormat.textFormat',
      start: { sheetId: sheetId, rowIndex: 1, columnIndex: 1 }, // B2
    },
  });

  // 9. Write Total Billable Amount to B3 and B4
  const formattedTotalBillable = `£${data.totalBillableAmount.toFixed(2)}`;

  for (let row = 2; row <= 3; row++) { // B3 and B4
    requests.push({
      updateCells: {
        rows: [{
          values: [{
            userEnteredValue: { stringValue: formattedTotalBillable },
            userEnteredFormat: {
              numberFormat: {
                type: 'CURRENCY',
                pattern: '£#,##0.00',
              },
              textFormat: {
                fontSize: 12,
                bold: false,
                foregroundColor: { red: 55/255, green: 93/255, blue: 117/255 }, // #375D75
              },
            },
          }],
        }],
        fields: 'userEnteredValue,userEnteredFormat.numberFormat,userEnteredFormat.textFormat',
        start: { sheetId: sheetId, rowIndex: row, columnIndex: 1 }, // B3 and B4
      },
    });
  }

  // 10. Write Total Labor Hours to B5
  requests.push({
    updateCells: {
      rows: [{
        values: [{
          userEnteredValue: { numberValue: data.totalLaborHours },
          userEnteredFormat: {
            numberFormat: {
              type: 'NUMBER',
              pattern: '0.00',
            },
            textFormat: {
              fontSize: 12,
              bold: false,
              foregroundColor: { red: 55/255, green: 93/255, blue: 117/255 }, // #375D75
            },
            horizontalAlignment: 'LEFT',
          },
        }],
      }],
      fields: 'userEnteredValue,userEnteredFormat.numberFormat,userEnteredFormat.textFormat',
      start: { sheetId: sheetId, rowIndex: 4, columnIndex: 1 }, // B5
    },
  });

  // Execute all requests
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: spreadsheetId,
    requestBody: {
      requests: requests,
    },
  });
};

app.post("/write-to-sheet", async (req, res) => {
  try {
    const { data, dateRange, totalBillableAmount, totalLaborHours } = req.body;

    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid data format. Expected 'data' to be an array of arrays." });
    }

    if (!dateRange || typeof dateRange !== 'string') {
      return res.status(400).json({ error: "Invalid or missing 'dateRange'. Expected a string." });
    }

    if (typeof totalBillableAmount !== 'number' || typeof totalLaborHours !== 'number') {
      return res.status(400).json({ error: "Invalid 'totalBillableAmount' or 'totalLaborHours'. Expected numbers." });
    }

    // Define the spreadsheet ID
    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Get existing sheets to check if "Detailed Report" exists
    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheet = sheetMetadata.data.sheets.find(
      (sheet) => sheet.properties.title === "Detailed Report"
    );

    let detailedReportSheetId;
    if (!sheet) {
      // Create the "Detailed Report" sheet if it doesn't exist
      const addSheetResponse = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: "Detailed Report",
                },
              },
            },
          ],
        },
      });
      detailedReportSheetId = addSheetResponse.data.replies[0].addSheet.properties.sheetId;
      console.log('"Detailed Report" sheet created.');
    } else {
      detailedReportSheetId = sheet.properties.sheetId;
      // Clear the existing "Detailed Report" sheet content except for formatting (optional)
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: "Detailed Report!A2:J", // Adjust range as needed to preserve formatting
      });
      console.log('"Detailed Report" sheet cleared.');
    }

    // Apply formatting and write date range and totals
    await applyFormatting(sheets, spreadsheetId, detailedReportSheetId, { dateRange, totalBillableAmount, totalLaborHours });

    console.log('Formatting applied successfully.');

    // 11. Write Data to the Sheet
    // Assuming data starts from row 9
    const dataRange = `Detailed Report!A9:J${8 + data.length}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: dataRange,
      valueInputOption: 'RAW',
      requestBody: {
        values: data,
      },
    });

    console.log('Data written successfully.');

    res.status(200).json({ message: "Detailed Report created and data written successfully!" });
  } catch (error) {
    console.error("Error in /write-to-sheet:", error);
    res.status(500).json({ error: "Failed to write to Google Sheet" });
  }
});

app.post("/login", async (req, res) => {
  try {
    const apiUrl = `${process.env.API_BASE_URL}/tokens/login`;

    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(req.body),
    });

    const data = await response.json();
    res.status(response.status).json(data);
  } catch (error) {
    res.status(500).json({ error: "Error connecting to the API." });
  }
});

app.get("/all-clients", async (req, res) => {
  try {
    const token = req.headers.authorization;

    if (!token || !token.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Unauthorized: Invalid token format" });
    }

    const apiUrl = `${process.env.API_BASE_URL}/Clients`;

    const response = await fetch(apiUrl, {
      method: "GET",
      headers: {
        "Authorization": token,
        "Accept": "application/json",
        "api-version": "1.0",
      },
    });

    if (!response.ok) {
      const errorMessage = await response.text();
      return res.status(response.status).json({ error: errorMessage });
    }

    const clients = await response.json();
    const activeClients = clients.filter((client) => !client.archived);

    res.status(200).json(activeClients);
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch clients. Please try again later." });
  }
});

app.get("/time-logs", async (req, res) => {
  try {
    const token = req.headers.authorization;

    if (!token || !token.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Unauthorized: Invalid token format" });
    }

    const { DateFrom, DateTo } = req.query;

    if (!DateFrom || !DateTo) {
      return res.status(400).json({ error: "Missing required query parameters" });
    }

    const apiUrl = `${process.env.API_BASE_URL}/Reports/activity?DateFrom=${DateFrom}&DateTo=${DateTo}`;

    const response = await fetch(apiUrl, {
      method: "GET",
      headers: {
        "Authorization": token,
        "Accept": "application/json",
        "api-version": "1.0",
      },
    });

    if (!response.ok) {
      const errorMessage = await response.text();
      return res.status(response.status).json({ error: errorMessage });
    }

    const logs = await response.json();
    res.status(200).json(logs);
  } catch (error) {
    res.status(500).json({ error: "Failed to fetch time logs. Please try again later." });
  }
});

app.put("/edit-log", async (req, res) => {
  try {
    const token = req.headers.authorization;

    if (!token || !token.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Unauthorized: Invalid token format" });
    }

    const apiUrl = `${process.env.API_BASE_URL}/Admin/editLogOnBehalf`;

    const response = await fetch(apiUrl, {
      method: "PUT",
      headers: {
        "Authorization": token,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(req.body), // Forward the body to the API
    });

    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json({ error: data.message || "Failed to update log" });
    }

    res.status(200).json(data);
  } catch (error) {
    res.status(500).json({ error: "Error updating the time log." });
  }
});


app.listen(PORT, () => {
  console.log(`Proxy server running on http://localhost:${PORT}`);
});
