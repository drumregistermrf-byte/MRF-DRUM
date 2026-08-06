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

const LOCK_TIMEOUT_MS = 2 * 60 * 1000;

const client = new MongoClient(MONGODB_URI);
await client.connect();
const db = client.db("mrf_drum_registry");
const storage = db.collection("storage");
const locks = db.collection("locks");

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

// Drum edit locks live in their own collection (not the generic storage blob)
// so acquiring one can be a single atomic Mongo operation instead of a
// read-modify-write on a shared JSON document, which would let two people
// clicking "edit" at the same moment both believe they got the lock.
app.get("/api/locks", async (req, res) => {
  const all = await locks.find({}).toArray();
  const out = {};
  for (const l of all) out[l._id] = { by: l.by, ts: l.ts };
  res.json({ locks: out });
});

app.post("/api/locks/:drumId/acquire", async (req, res) => {
  const { by } = req.body;
  if (!by) return res.status(400).json({ error: "by is required" });
  const drumId = req.params.drumId;
  const now = Date.now();
  const cutoff = now - LOCK_TIMEOUT_MS;

  try {
    const doc = await locks.findOneAndUpdate(
      { _id: drumId, $or: [{ ts: { $lt: cutoff } }, { by }] },
      { $set: { by, ts: now } },
      { upsert: true, returnDocument: "after" }
    );
    res.json({ ok: true, lock: { by: doc.by, ts: doc.ts } });
  } catch (err) {
    if (err.code === 11000) {
      const existing = await locks.findOne({ _id: drumId });
      return res.json({
        ok: false,
        lock: existing ? { by: existing.by, ts: existing.ts } : null,
      });
    }
    throw err;
  }
});

app.post("/api/locks/:drumId/release", async (req, res) => {
  const { by } = req.body;
  await locks.deleteOne({ _id: req.params.drumId, by });
  res.json({ ok: true });
});

app.listen(PORT, () => {
  console.log(`MRF Drum Registry API listening on port ${PORT}`);
});
