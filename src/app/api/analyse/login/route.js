import { NextResponse } from "next/server";

const USERS = [
  { username: process.env.USER1_NAME || "arash", password: process.env.USER1_PASS || "fibernc2024", displayName: "Arash" },
  { username: process.env.USER2_NAME || "leitstelle", password: process.env.USER2_PASS || "leitstelle123", displayName: "Leitstelle" },
  { username: process.env.USER3_NAME || "mitarbeiter", password: process.env.USER3_PASS || "kpi123", displayName: "Mitarbeiter" },
];

export async function POST(request) {
  try {
    const { username, password } = await request.json();
    const user = USERS.find(u => u.username.toLowerCase() === username.toLowerCase() && u.password === password);
    if (!user) return NextResponse.json({ error: "Benutzername oder Passwort falsch." }, { status: 401 });

    const token = Buffer.from(`${user.username}:${Date.now()}`).toString("base64");
    const response = NextResponse.json({ success: true, displayName: user.displayName });
    response.cookies.set("auth_token", token, { httpOnly: true, secure: true, sameSite: "lax", maxAge: 60*60*24*7, path: "/" });
    response.cookies.set("auth_user", user.displayName, { httpOnly: false, secure: true, sameSite: "lax", maxAge: 60*60*24*7, path: "/" });
    return response;
  } catch {
    return NextResponse.json({ error: "Server-Fehler." }, { status: 500 });
  }
}
