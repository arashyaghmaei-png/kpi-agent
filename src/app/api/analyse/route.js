import { NextResponse } from "next/server";

const UPSTASH_URL = process.env.UPSTASH_REDIS_REST_URL;
const UPSTASH_TOKEN = process.env.UPSTASH_REDIS_REST_TOKEN;

async function redisGet(key) {
  if (!UPSTASH_URL || !UPSTASH_TOKEN) return null;
  try {
    const res = await fetch(`${UPSTASH_URL}/get/${key}`, {
      headers: { Authorization: `Bearer ${UPSTASH_TOKEN}` }
    });
    const data = await res.json();
    if (!data.result) return null;
    return JSON.parse(data.result);
  } catch { return null; }
}

async function redisSet(key, value) {
  if (!UPSTASH_URL || !UPSTASH_TOKEN) return false;
  try {
    await fetch(`${UPSTASH_URL}/set/${key}`, {
      method: "POST",
      headers: { 
        Authorization: `Bearer ${UPSTASH_TOKEN}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(JSON.stringify(value))
    });
    return true;
  } catch { return false; }
}

export async function GET(request) {
  const user = request.cookies.get("auth_user")?.value || "default";
  const key = `kpi_data_${user.toLowerCase().replace(/\s+/g, "_")}`;
  const data = await redisGet(key);
  return NextResponse.json({ success: true, data: data || {} });
}

export async function POST(request) {
  const user = request.cookies.get("auth_user")?.value || "default";
  const key = `kpi_data_${user.toLowerCase().replace(/\s+/g, "_")}`;
  const body = await request.json();
  const saved = await redisSet(key, body);
  return NextResponse.json({ success: saved });
}
