import { NextResponse } from "next/server";

// Simple file-based storage using Vercel's edge config or env
// We use a simple approach: store data as base64 in the response
// and read from cookies for per-user storage

const KV_BASE_URL = process.env.KV_REST_API_URL;
const KV_TOKEN = process.env.KV_REST_API_TOKEN;

async function kvGet(key) {
  if (!KV_BASE_URL || !KV_TOKEN) return null;
  try {
    const res = await fetch(`${KV_BASE_URL}/get/${key}`, {
      headers: { Authorization: `Bearer ${KV_TOKEN}` }
    });
    const data = await res.json();
    return data.result ? JSON.parse(data.result) : null;
  } catch { return null; }
}

async function kvSet(key, value) {
  if (!KV_BASE_URL || !KV_TOKEN) return false;
  try {
    await fetch(`${KV_BASE_URL}/set/${key}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${KV_TOKEN}`, "Content-Type": "application/json" },
      body: JSON.stringify(JSON.stringify(value))
    });
    return true;
  } catch { return false; }
}

export async function GET(request) {
  const user = request.cookies.get("auth_user")?.value || "default";
  const key = `kpi_data_${user.toLowerCase().replace(/\s/g, "_")}`;
  
  const data = await kvGet(key);
  return NextResponse.json({ success: true, data: data || {} });
}

export async function POST(request) {
  const user = request.cookies.get("auth_user")?.value || "default";
  const key = `kpi_data_${user.toLowerCase().replace(/\s/g, "_")}`;
  
  const body = await request.json();
  const saved = await kvSet(key, body);
  
  return NextResponse.json({ success: saved });
}
