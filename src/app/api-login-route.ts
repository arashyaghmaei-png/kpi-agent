import { NextResponse } from "next/server";

const USERS = [
  { username: process.env.USER1_NAME || "arash", password: process.env.USER1_PASS || "fibernc2024", role: "admin", displayName: "Arash" },
  { username: process.env.USER2_NAME || "leitstelle", password: process.env.USER2_PASS || "leitstelle123", role: "user", displayName: "Leitstelle" },
  { username: process.env.USER3_NAME || "techniker", password: process.env.USER3_PASS || "tech123", role: "user", displayName: "Techniker" },
];

export async function POST(request) {
  try {
    const { username, password } = await request.json();

    const user = USERS.find(
      (u) => u.username.toLowerCase() === username.toLowerCase() && u.password === password
    );

    if (!user) {
      return NextResponse.json({ error: "Benutzername oder Passwort falsch." }, { status: 401 });
    }

    const token = Buffer.from(`${user.username}:${Date.now()}:${process.env.AUTH_SECRET || "fibernc_secret"}`).toString("base64");

    const response = NextResponse.json({ 
      success: true, 
      displayName: user.displayName,
      role: user.role
    });

    response.cookies.set("auth_token", token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      maxAge: 60 * 60 * 24 * 7, // 7 Tage
      path: "/",
    });

    response.cookies.set("auth_user", user.displayName, {
      httpOnly: false,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      maxAge: 60 * 60 * 24 * 7,
      path: "/",
    });

    return response;
  } catch {
    return NextResponse.json({ error: "Server-Fehler." }, { status: 500 });
  }
}
