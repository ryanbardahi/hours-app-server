import express from "express";
import fetch from "node-fetch";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 4000;

const corsOptions = {
  origin: ["http://localhost:3000","https://hours-app-jade.vercel.app"],
  credentials: true,
  optionsSuccessStatus: 200
};

app.use(cors(corsOptions));
app.use(express.json());

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
    console.error("Error connecting to the API:", error.message);
    res.status(500).json({ error: "Error connecting to the API." });
  }
});

app.get("/all-clients", async (req, res) => {
  try {
    const token = req.headers.authorization;

    console.log("Authorization header received in backend:", token);

    if (!token || !token.startsWith("Bearer ")) {
      return res.status(401).json({ error: "Unauthorized: Invalid token format" });
    }

    const apiUrl = `${process.env.API_BASE_URL}/Clients`;

    console.log("Fetching clients from:", apiUrl);

    const response = await fetch(apiUrl, {
      method: "GET",
      headers: {
        "Authorization": token, // Pass the token to the upstream API
        "Accept": "application/json", // Required header
        "api-version": "1.0", // Required header
      },
    });

    console.log("Upstream API response status:", response.status);

    if (!response.ok) {
      const errorMessage = await response.text();
      console.error("Upstream API error:", errorMessage);
      return res.status(response.status).json({ error: errorMessage });
    }

    const clients = await response.json();
    res.status(200).json(clients);
  } catch (error) {
    console.error("Error fetching clients:", error.message);
    res.status(500).json({ error: "Failed to fetch clients. Please try again later." });
  }
});




// Start the server
app.listen(PORT, () => {
  console.log(`Proxy server running on http://localhost:${PORT}`);
});