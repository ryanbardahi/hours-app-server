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

  const formattedRows = labels.map((label) => ({
    values: [{
      userEnteredValue: { stringValue: label },
      userEnteredFormat: {
        textFormat: {
          fontSize: 10,
          bold: false,
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

  // 4. Set Headers in Row 8
  const headers = [
    'DATE', 'USER', 'CLIENT', 'PROJECT', 'TASK', 'IS BILLABLE',
    'BILLABLE AMOUNT', 'START/FINISH TIME', 'TOTAL HOURS',
    'BILLABLE HOURS', 'DESCRIPTION'
  ];

  const headerValues = headers.map(header => ({
    userEnteredValue: { stringValue: header },
    userEnteredFormat: {
      textFormat: {
        fontSize: 10,
        bold: true,
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

  // 5. Set Column Widths
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
        pixelSize: 280,
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

  // 6. Wrap Text for All Columns
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startRowIndex: 0,
        endRowIndex: 1000, // Adjust as needed
        startColumnIndex: 0,
        endColumnIndex: 11, // Columns A-K
      },
      cell: {
        userEnteredFormat: {
          wrapStrategy: 'WRAP',
        },
      },
      fields: 'userEnteredFormat.wrapStrategy',
    },
  });

  // 7. Write Date Range to B2
  requests.push({
    updateCells: {
      rows: [{
        values: [{
          userEnteredValue: { stringValue: data.dateRange },
          userEnteredFormat: {
            textFormat: {
              fontSize: 12,
              bold: true,
              foregroundColor: { red: 55/255, green: 93/255, blue: 117/255 }, // #375D75
            },
          },
        }],
      }],
      fields: 'userEnteredValue,userEnteredFormat.textFormat',
      start: { sheetId: sheetId, rowIndex: 1, columnIndex: 1 }, // B2
    },
  });

  // 8. Write Total Billable Amount to B3 and B4
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
                bold: true,
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

  // 9. Write Total Labor Hours to B5
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
              bold: true,
              foregroundColor: { red: 55/255, green: 93/255, blue: 117/255 }, // #375D75
            },
            horizontalAlignment: 'LEFT',
          },
        }],
      }],
      fields: 'userEnteredValue,userEnteredFormat.numberFormat,userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment',
      start: { sheetId: sheetId, rowIndex: 4, columnIndex: 1 }, // B5
    },
  });

  // 10. Format Date Rows
  dateRowIndices.forEach((rowIndex) => {
    // Apply background color and bold text to entire row (A-J)
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: rowIndex - 1, // zero-based index
          endRowIndex: rowIndex,
          startColumnIndex: 0, // Column A
          endColumnIndex: 10, // Column J
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: {
              red: 220 / 255,
              green: 234 / 255,
              blue: 250 / 255,
            },
            textFormat: {
              bold: true,
            },
          },
        },
        fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold',
      },
    });

    // Specifically bold Columns A, F, H, I
    ['A', 'F', 'H', 'I'].forEach((col) => {
      const colIndex = col.charCodeAt(0) - 'A'.charCodeAt(0);
      requests.push({
        repeatCell: {
          range: {
            sheetId: sheetId,
            startRowIndex: rowIndex - 1,
            endRowIndex: rowIndex,
            startColumnIndex: colIndex,
            endColumnIndex: colIndex + 1,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 220 / 255,
                green: 234 / 255,
                blue: 250 / 255,
              },
              textFormat: {
                bold: true,
              },
            },
          },
          fields: 'userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.bold',
        },
      });
    });
  });

  // 11. Format Column F as Currency (GBP)
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startColumnIndex: 5, // Column F
        endColumnIndex: 6,
      },
      cell: {
        userEnteredFormat: {
          numberFormat: {
            type: 'CURRENCY',
            pattern: '£#,##0.00',
          },
        },
      },
      fields: 'userEnteredFormat.numberFormat',
    },
  });

  // 12. Format Total Row
  if (totalRowIndex) {
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: totalRowIndex - 1,
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
        fields: 'userEnteredFormat.textFormat',
      },
    });
  }

  // Execute all requests
  if (requests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: spreadsheetId,
      requestBody: {
        requests: requests,
      },
    });
  }
};

app.post("/write-to-sheet", async (req, res) => {
  try {
    const { data, dateRange, totalBillableAmount, totalLaborHours, totalBillableHours } = req.body;

    // Validate input
    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid data format. Expected 'data' to be an array of arrays." });
    }

    if (!dateRange || typeof dateRange !== 'string') {
      return res.status(400).json({ error: "Invalid or missing 'dateRange'. Expected a string." });
    }

    if (
      typeof totalBillableAmount !== 'number' ||
      typeof totalLaborHours !== 'number' ||
      typeof totalBillableHours !== 'number'
    ) {
      return res.status(400).json({
        error: "Invalid 'totalBillableAmount', 'totalLaborHours', or 'totalBillableHours'. Expected numbers.",
      });
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
      // Clear the existing "Detailed Report" sheet content except for formatting
      await sheets.spreadsheets.values.clear({
        spreadsheetId,
        range: "Detailed Report!A9:J", // Start clearing from row 9
      });
      console.log('"Detailed Report" sheet cleared.');
    }

    // Group data by date
    const groupedData = data.reduce((acc, row) => {
      const date = row[0]; // Assuming date is in Column A
      if (!acc[date]) {
        acc[date] = [];
      }
      acc[date].push(row);
      return acc;
    }, {});

    // Sort dates
    const sortedDates = Object.keys(groupedData).sort((a, b) => new Date(a) - new Date(b));

    // Prepare final data with date rows and log entries
    const finalData = [];
    const dateRowIndices = []; // To track which rows are date rows for formatting
    let currentRow = 9; // Starting from row 9

    sortedDates.forEach(date => {
      // Format the date as "Mon, 25 Nov, 2024"
      const dateObj = new Date(date);
      const formattedDate = dateObj.toLocaleDateString(undefined, {
        weekday: 'short',
        day: 'numeric',
        month: 'short',
        year: 'numeric',
      });

      // Calculate totals for this date
      const billableAmountForDate = groupedData[date].reduce((sum, row) => sum + (parseFloat(row[5]) || 0), 0);
      const totalHoursForDate = groupedData[date].reduce((sum, row) => sum + (parseFloat(row[7]) || 0), 0);
      const billableHoursForDate = groupedData[date].reduce((sum, row) => sum + (parseFloat(row[8]) || 0), 0);

      // Create date row
      const dateRow = [
        formattedDate, // Column A
        "", // B
        "", // C
        "", // D
        "", // E
        billableAmountForDate, // F
        "", // G
        totalHoursForDate, // H
        billableHoursForDate, // I
        "", // J
      ];

      finalData.push(dateRow);
      dateRowIndices.push(currentRow);
      currentRow++;

      // Add log entries under this date
      groupedData[date].forEach(log => {
        // Ensure each log entry has exactly 10 columns (A-J)
        const logEntry = [
          log[0], // DATE
          log[1], // USER
          log[2], // CLIENT
          log[3], // PROJECT
          log[4], // TASK
          log[5], // IS BILLABLE
          log[6], // BILLABLE AMOUNT
          log[7], // START/FINISH TIME
          log[8], // TOTAL HOURS
          log[9], // BILLABLE HOURS
          log[10] || "", // DESCRIPTION (if exists)
        ];
        finalData.push(logEntry);
        currentRow++;
      });
    });

    // Create total row
    const totalRow = [
      "TOTAL", // A
      "", // B
      "", // C
      "", // D
      "", // E
      totalBillableAmount, // F
      "", // G
      totalLaborHours, // H
      totalBillableHours, // I
      "", // J
    ];

    finalData.push(totalRow);
    const totalRowIndex = currentRow;
    currentRow++;

    // Write the final data to the sheet starting from row 9
    const dataRange = `Detailed Report!A9:J${8 + finalData.length}`; // A9:J{lastRow}

    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: dataRange,
      valueInputOption: 'RAW',
      requestBody: {
        values: finalData,
      },
    });

    console.log('Data written successfully.');

    // Apply formatting
    await applyFormatting(
      sheets,
      spreadsheetId,
      detailedReportSheetId,
      dateRowIndices,
      8 + finalData.length // Total row index (1-based)
    );

    console.log('Formatting applied successfully.');

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