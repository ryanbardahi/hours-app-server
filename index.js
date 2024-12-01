import express from "express";
import fetch from "node-fetch";
import cors from "cors";
import dotenv from "dotenv";
import { google } from "googleapis";

dotenv.config();

const app = express();
const PORT = 4000;

const corsOptions = {
  origin: ["http://localhost:3000", "https://hours-app-beta.vercel.app"],
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
const applyFormatting = async (sheets, spreadsheetId, sheetId, dateRowIndices, totalRowIndex) => {
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
            type: "TEXT",
          },
        },
      },
      fields: "userEnteredFormat.numberFormat",
    },
  });

  // 2. Format Header Cell (A1)
  requests.push({
    updateCells: {
      rows: [
        {
          values: [
            {
              userEnteredValue: { stringValue: "DETAILED REPORT" },
              userEnteredFormat: {
                textFormat: {
                  fontSize: 22,
                  bold: true,
                  foregroundColor: { red: 59 / 255, green: 143 / 255, blue: 194 / 255 }, // #3B8FC2
                },
              },
            },
          ],
        },
      ],
      fields: "userEnteredValue,userEnteredFormat.textFormat",
      start: { sheetId: sheetId, rowIndex: 0, columnIndex: 0 },
    },
  });

  // 3. Format Static Labels (A2-A5)
  const labels = [
    "Time frame",
    "Billable amount (hours)",
    "Total billable amount",
    "Total hours",
  ];

  const formattedRows = labels.map(label => ({
    values: [
      {
        userEnteredValue: { stringValue: label },
        userEnteredFormat: {
          textFormat: {
            fontSize: 10,
            bold: false,
          },
        },
      },
    ],
  }));

  requests.push({
    updateCells: {
      rows: formattedRows,
      fields: "userEnteredValue,userEnteredFormat.textFormat",
      start: { sheetId: sheetId, rowIndex: 1, columnIndex: 0 }, // A2 starts at rowIndex 1
    },
  });

  // 4. Format Headers in Row 8
  const headers = [
    "DATE", "USER", "CLIENT", "PROJECT", "TASK", "IS BILLABLE",
    "BILLABLE AMOUNT", "START/FINISH TIME", "TOTAL HOURS",
    "BILLABLE HOURS", "DESCRIPTION",
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
      rows: [{ values: headerValues }],
      fields: "userEnteredValue,userEnteredFormat.textFormat",
      start: { sheetId: sheetId, rowIndex: 7, columnIndex: 0 }, // Row 8
    },
  });

  // 5. Set Column Widths
  const columnWidths = [
    { index: 0, width: 280 }, // Column A
    { index: 9, width: 500 }, // Column J
  ];
  columnWidths.push(...Array.from({ length: 8 }, (_, i) => ({ index: i + 1, width: 150 }))); // Columns B-I

  columnWidths.forEach(({ index, width }) => {
    requests.push({
      updateDimensionProperties: {
        range: {
          sheetId: sheetId,
          dimension: "COLUMNS",
          startIndex: index,
          endIndex: index + 1,
        },
        properties: { pixelSize: width },
        fields: "pixelSize",
      },
    });
  });

  // 6. Wrap Text for All Columns
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
          wrapStrategy: "WRAP",
        },
      },
      fields: "userEnteredFormat.wrapStrategy",
    },
  });

  // 7. Apply Currency Format to Column F
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startRowIndex: 8,
        endRowIndex: 1000, // Adjust based on your dataset size
        startColumnIndex: 5,
        endColumnIndex: 6,
      },
      cell: {
        userEnteredFormat: {
          numberFormat: {
            type: "CURRENCY",
            pattern: "GBP#,##0.00",
          },
        },
      },
      fields: "userEnteredFormat.numberFormat",
    },
  });

  // 8. Format Date Rows
  dateRowIndices.forEach(index => {
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: index - 1, // Zero-based indexing
          endRowIndex: index,
          startColumnIndex: 0,
          endColumnIndex: 10,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 220 / 255, green: 234 / 255, blue: 250 / 255 }, // #dceefa
            textFormat: { bold: true },
          },
        },
        fields: "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold",
      },
    });

    // Format Columns F, H, I in Date Rows
    [5, 7, 8].forEach(colIndex => { // Columns F (5), H (7), I (8)
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: index - 1,
            endRowIndex: index,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: {
            userEnteredFormat: {
              textFormat: { bold: true },
              // Additional number formatting if needed
            },
          },
          fields: "userEnteredFormat.textFormat.bold",
        },
      });
    });
  });

  // 9. Format Total Row
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startRowIndex: totalRowIndex - 1, // Zero-based indexing
        endRowIndex: totalRowIndex,
        startColumnIndex: 0,
        endColumnIndex: 10,
      },
      cell: {
        userEnteredFormat: {
          textFormat: {
            bold: true,
            fontSize: 12,
          },
        },
      },
      fields: "userEnteredFormat.textFormat.bold,userEnteredFormat.textFormat.fontSize",
    },
  });

  // Format Columns F, H, I in Total Row
  [5, 7, 8].forEach(colIndex => { // Columns F (5), H (7), I (8)
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: totalRowIndex - 1,
          endRowIndex: totalRowIndex,
          startColumnIndex: colIndex,
          endColumnIndex: colIndex + 1,
        },
        cell: {
          userEnteredFormat: {
            textFormat: {
              bold: true,
              fontSize: 12,
            },
            // Ensure Column F is formatted as currency
            numberFormat: colIndex === 5 ? {
              type: "CURRENCY",
              pattern: "GBP#,##0.00",
            } : {
              type: "NUMBER",
              pattern: "0.00",
            },
          },
        },
        fields: "userEnteredFormat.textFormat,userEnteredFormat.numberFormat",
      },
    });
  });

  // Execute the batch update
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: spreadsheetId,
    requestBody: {
      requests: requests,
    },
  });
};

app.post("/write-to-sheet", async (req, res) => {
  try {
    const { data, dateRange, totalBillableAmount, totalLaborHours, totalBillableHours } = req.body;

    // Input Validation
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid data format. Expected 'data' to be an array of arrays." });
    }

    if (!dateRange || typeof dateRange !== 'string') {
      return res.status(400).json({ error: "Invalid or missing 'dateRange'. Expected a string." });
    }

    if (typeof totalBillableAmount !== 'number' || typeof totalLaborHours !== 'number' || typeof totalBillableHours !== 'number') {
      return res.status(400).json({ error: "Invalid 'totalBillableAmount', 'totalLaborHours', or 'totalBillableHours'. Expected numbers." });
    }

    // Spreadsheet Configuration
    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Check if "Detailed Report" sheet exists
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
      // Clear existing data from A9:J to preserve formatting
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: "Detailed Report!A9:J",
      });
      console.log('"Detailed Report" sheet cleared.');
    }

    // Process Data to Include Date Rows and Track Indices
    const processedData = [];
    let currentDate = null;
    const dateRowIndices = []; // To track where date rows are inserted
    let sheetRowIndex = 9; // Starting from row 9

    // Sort data by date to ensure proper grouping
    data.sort((a, b) => new Date(a[0]) - new Date(b[0]));

    data.forEach((log) => {
      const logDate = log[0]; // Assuming the first element is the date in 'YYYY-MM-DD' format

      if (logDate !== currentDate) {
        currentDate = logDate;
        const formattedDate = new Date(logDate).toLocaleDateString(undefined, {
          weekday: 'short',
          day: 'numeric',
          month: 'short',
          year: 'numeric',
        });
        // Insert Date Row
        const dateRow = [
          formattedDate, // Column A
          '', // B
          '', // C
          '', // D
          '', // E
          0,  // F (Billable Amount for the date)
          '', // G
          0,  // H (Total Hours for the date)
          0,  // I (Billable Hours for the date)
          '', // J
        ];
        processedData.push(dateRow);
        dateRowIndices.push(sheetRowIndex);
        sheetRowIndex += 1;
      }

      // Add Log Entry
      processedData.push(log);
      sheetRowIndex += 1;

      // Update the latest date row with accumulated values
      const dateRowIndex = dateRowIndices[dateRowIndices.length - 1] - 9; // Zero-based for processedData array
      // Column F (Billable Amount)
      processedData[dateRowIndex][5] += log[5];
      // Column H (Total Hours)
      processedData[dateRowIndex][7] += parseFloat(log[7]);
      // Column I (Billable Hours)
      processedData[dateRowIndex][8] += parseFloat(log[8]);
    });

    // Append Total Row
    const totalRow = [
      'TOTAL', // Column A
      '', // B
      '', // C
      '', // D
      '', // E
      totalBillableAmount, // F
      '', // G
      totalLaborHours, // H
      totalBillableHours, // I
      '', // J
    ];
    processedData.push(totalRow);
    const totalRowIndex = sheetRowIndex; // The next row after the last log

    // Write Processed Data to the Sheet
    const startRow = 9;
    const endRow = sheetRowIndex;
    const dataRange = `Detailed Report!A${startRow}:J${endRow}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: dataRange,
      valueInputOption: 'RAW',
      requestBody: {
        values: processedData,
      },
    });

    console.log('Data written successfully.');

    // Apply Formatting to Date Rows and Total Row
    await applyFormatting(sheets, spreadsheetId, detailedReportSheetId, dateRowIndices, totalRowIndex);

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