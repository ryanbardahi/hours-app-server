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
      start: { sheetId: sheetId, rowIndex: 1, columnIndex: 0 },
    },
  });

  // 4. Format Headers in Row 8
  const headers = [
    "USER", "CLIENT", "PROJECT", "TASK", "IS BILLABLE",
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
      start: { sheetId: sheetId, rowIndex: 7, columnIndex: 0 },
    },
  });

  // 5. Format Date Rows
  dateRowIndices.forEach(index => {
    requests.push({
      repeatCell: {
        range: {
          sheetId: sheetId,
          startRowIndex: index,
          endRowIndex: index + 1,
          startColumnIndex: 0,
          endColumnIndex: 10,
        },
        cell: {
          userEnteredFormat: {
            backgroundColor: { red: 0.86, green: 0.93, blue: 0.98 }, // #dceefa
            textFormat: { bold: true },
          },
        },
        fields: "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
      },
    });
  });

  // 6. Format Total Row
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
      fields: "userEnteredFormat.textFormat",
    },
  });

  // 7. Set Column Widths
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

  // 8. Wrap Text for All Columns
  requests.push({
    repeatCell: {
      range: {
        sheetId: sheetId,
        startRowIndex: 0,
        endRowIndex: 1000,
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

  // 9. Apply Currency Format to Column F
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

  // Execute the batch update
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: { requests },
  });
};



app.post("/write-to-sheet", async (req, res) => {
  try {
    const { data, dateRange, totalBillableAmount, totalLaborHours } = req.body;

    if (!data || !Array.isArray(data)) {
      return res.status(400).json({ error: "Invalid data format. Expected 'data' to be an array of objects." });
    }

    if (!dateRange || typeof dateRange !== "string") {
      return res.status(400).json({ error: "Invalid or missing 'dateRange'. Expected a string." });
    }

    // Define the spreadsheet ID
    const spreadsheetId = process.env.SPREADSHEET_ID;

    // Get or create the "Detailed Report" sheet
    const sheetMetadata = await sheets.spreadsheets.get({ spreadsheetId });
    const sheet = sheetMetadata.data.sheets.find(sheet => sheet.properties.title === "Detailed Report");

    let detailedReportSheetId;
    if (!sheet) {
      const addSheetResponse = await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: "Detailed Report" } } }] },
      });
      detailedReportSheetId = addSheetResponse.data.replies[0].addSheet.properties.sheetId;
    } else {
      detailedReportSheetId = sheet.properties.sheetId;
      await sheets.spreadsheets.values.clear({ spreadsheetId, range: "Detailed Report!A9:J" });
    }

    // Organize data by date
    const groupedData = data.reduce((acc, log) => {
      const logDate = new Date(log.date).toLocaleDateString("en-GB", { weekday: "short", day: "numeric", month: "short", year: "numeric" });
      if (!acc[logDate]) acc[logDate] = [];
      acc[logDate].push(log);
      return acc;
    }, {});

    const rows = [];
    let grandTotalBillable = 0;
    let grandTotalHours = 0;
    let grandTotalBillableHours = 0;

    // Prepare rows for each date group
    Object.entries(groupedData).forEach(([date, logs]) => {
      const dateTotalBillable = logs.reduce((sum, log) => sum + (log.billableAmount || 0), 0);
      const dateTotalHours = logs.reduce((sum, log) => sum + (log.laborHours || 0), 0);
      const dateTotalBillableHours = logs.reduce((sum, log) => sum + (log.billableHours || 0), 0);

      rows.push([
        date, "", "", "", "",
        dateTotalBillable.toFixed(2), "", dateTotalHours.toFixed(2), dateTotalBillableHours.toFixed(2), "",
      ]);

      logs.forEach(log => {
        rows.push([
          log.userName || "N/A",
          log.clientName || "N/A",
          log.projectName || "N/A",
          log.taskName || "N/A",
          log.billable ? "Billable" : "Not Billable",
          (log.billableAmount || 0).toFixed(2),
          log.startEndTime || "-",
          (log.laborHours || 0).toFixed(2),
          (log.billableHours || 0).toFixed(2),
          log.note || "N/A",
        ]);
      });

      grandTotalBillable += dateTotalBillable;
      grandTotalHours += dateTotalHours;
      grandTotalBillableHours += dateTotalBillableHours;
    });

    // Add the total row at the end
    rows.push([
      "TOTAL", "", "", "", "",
      grandTotalBillable.toFixed(2), "", grandTotalHours.toFixed(2), grandTotalBillableHours.toFixed(2), "",
    ]);

    // Write data to sheet
    const dataRange = `Detailed Report!A9:J${8 + rows.length}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: dataRange,
      valueInputOption: "RAW",
      requestBody: { values: rows },
    });

    // Apply formatting
    await applyFormatting(sheets, spreadsheetId, detailedReportSheetId, Object.keys(groupedData).map((_, i) => 8 + i), rows.length + 8);

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