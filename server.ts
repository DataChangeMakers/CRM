import express from "express";
import path from "path";
import { fileURLToPath } from "url";
import cors from "cors";
import dotenv from "dotenv";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// ProPublica Proxy
app.get("/api/propublica/:ein", async (req, res) => {
  try {
    const { ein } = req.params;
    const response = await fetch(`https://projects.propublica.org/nonprofits/api/v2/organizations/${ein}.json`);
    if (!response.ok) {
      return res.status(response.status).json({ error: "ProPublica API error" });
    }
    const data = await response.json();
    res.json(data);
  } catch (error) {
    console.error("ProPublica proxy error:", error);
    res.status(500).json({ error: "Internal server error" });
  }
});

// Diagnostics
app.get("/api/health", (req, res) => {
  res.json({ 
    status: "ok", 
    timestamp: new Date().toISOString(),
    env: {
      hasGeminiKey: !!process.env.GEMINI_API_KEY,
      nodeEnv: process.env.NODE_ENV
    }
  });
});

async function start() {
  const isProduction = process.env.NODE_ENV === "production" || process.env.NODE_ENV === "prod";

  if (!isProduction) {
    console.log("Starting in development mode with Vite middleware...");
    const { createServer: createViteServer } = await import("vite");
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    console.log("Starting in production mode...");
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*all", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running at http://0.0.0.0:${PORT} in ${isProduction ? 'production' : 'development'} mode`);
  });
}

start();
