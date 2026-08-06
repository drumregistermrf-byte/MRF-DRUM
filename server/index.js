import "dotenv/config";
import express from "express";
import cors from "cors";
import { MongoClient } from "mongodb";

const PORT = process.env.PORT || 4000;
const MONGODB_URI = process.env.MONGODB_URI;
const API_KEY = process.env.API_KEY;
const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

if (!MONGODB_URI) {
  throw new Error("MONGODB_URI env var is required");
}

const client = new MongoClient(MONGODB_URI);
await client.connect();
const db = client.db("mrf_drum_registry");
const storage = db.collection("storage");

const app = express();
app.use(express.json());
app.use(
  cors(
    ALLOWED_ORIGINS.length
      ? { origin: ALLOWED_ORIGINS }
      : undefined
  )
);

app.use((req, res, next) => {
  if (req.path === "/health") return next();
  if (!API_KEY) return next();
  if (req.get("x-api-key") !== API_KEY) {
    return res.status(401).json({ error: "unauthorized" });
  }
  next();
});

app.get("/health", (req, res) => res.json({ ok: true }));

app.get("/api/storage/:key", async (req, res) => {
  const doc = await storage.findOne({ _id: req.params.key });
  res.json({ value: doc ? doc.value : null });
});

app.put("/api/storage/:key", async (req, res) => {
  const { value } = req.body;
  if (typeof value !== "string") {
    return res.status(400).json({ error: "value must be a string" });
  }
  await storage.updateOne(
    { _id: req.params.key },
    { $set: { value, updatedAt: new Date() } },
    { upsert: true }
  );
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`MRF Drum Registry API listening on port ${PORT}`);
});
