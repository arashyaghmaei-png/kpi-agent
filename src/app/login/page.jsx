'use client';

import { useState } from "react";
import { useRouter } from "next/navigation";

export default function LoginPage() {
  const router = useRouter();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [showReset, setShowReset] = useState(false);
  const [resetEmail, setResetEmail] = useState("");
  const [resetMsg, setResetMsg] = useState("");

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError("");

    try {
      const res = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });

      const data = await res.json();

      if (!res.ok) {
        setError(data.error || "Login fehlgeschlagen.");
      } else {
        router.push("/");
        router.refresh();
      }
    } catch {
      setError("Verbindungsfehler. Bitte versuche es erneut.");
    } finally {
      setLoading(false);
    }
  };

  const handleReset = () => {
    if (!resetEmail) {
      setResetMsg("Bitte Email eingeben.");
      return;
    }
    // In Produktion: Email senden. Hier: Hinweis an Admin
    setResetMsg(`Passwort-Reset angefragt für: ${resetEmail}. Bitte wende dich an den Administrator (Arash).`);
  };

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0a0f1a",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontFamily: "'Inter', sans-serif",
    }}>
      <div style={{
        background: "#111827",
        border: "1px solid #1f2937",
        borderRadius: 16,
        padding: "40px 36px",
        width: 380,
        boxShadow: "0 20px 60px rgba(0,0,0,0.5)",
      }}>
        {/* Logo */}
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: "#f9fafb", letterSpacing: -0.5 }}>
            KPI AGENT
          </div>
          <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>FiberNC - Telekom Subunternehmer</div>
        </div>

        {!showReset ? (
          <>
            <form onSubmit={handleLogin}>
              <div style={{ marginBottom: 16 }}>
                <label style={{ fontSize: 12, color: "#9ca3af", display: "block", marginBottom: 6 }}>
                  Benutzername
                </label>
                <input
                  type="text"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="z.B. arash"
                  autoComplete="username"
                  style={{
                    width: "100%",
                    background: "#0f172a",
                    border: "1px solid #374151",
                    borderRadius: 8,
                    color: "#e5e7eb",
                    padding: "10px 12px",
                    fontSize: 14,
                    boxSizing: "border-box",
                    outline: "none",
                  }}
                />
              </div>

              <div style={{ marginBottom: 20 }}>
                <label style={{ fontSize: 12, color: "#9ca3af", display: "block", marginBottom: 6 }}>
                  Passwort
                </label>
                <input
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  autoComplete="current-password"
                  style={{
                    width: "100%",
                    background: "#0f172a",
                    border: "1px solid #374151",
                    borderRadius: 8,
                    color: "#e5e7eb",
                    padding: "10px 12px",
                    fontSize: 14,
                    boxSizing: "border-box",
                    outline: "none",
                  }}
                />
              </div>

              {error && (
                <div style={{
                  background: "#2e0f0f",
                  border: "1px solid #7f1d1d",
                  borderRadius: 8,
                  padding: "10px 12px",
                  fontSize: 13,
                  color: "#f87171",
                  marginBottom: 16,
                }}>
                  {error}
                </div>
              )}

              <button
                type="submit"
                disabled={loading}
                style={{
                  width: "100%",
                  background: loading ? "#1e3a5f" : "#2563eb",
                  color: "#fff",
                  border: "none",
                  borderRadius: 8,
                  padding: "12px",
                  fontSize: 14,
                  fontWeight: 700,
                  cursor: loading ? "not-allowed" : "pointer",
                  marginBottom: 12,
                }}
              >
                {loading ? "Anmelden..." : "Anmelden"}
              </button>
            </form>

            <button
              onClick={() => setShowReset(true)}
              style={{
                width: "100%",
                background: "transparent",
                color: "#6b7280",
                border: "none",
                fontSize: 12,
                cursor: "pointer",
                padding: "6px",
                textDecoration: "underline",
              }}
            >
              Passwort vergessen?
            </button>
          </>
        ) : (
          <>
            <div style={{ fontSize: 14, fontWeight: 700, color: "#f9fafb", marginBottom: 12 }}>
              Passwort zurücksetzen
            </div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 16 }}>
              Gib deine Email-Adresse ein. Der Administrator wird dein Passwort zurücksetzen.
            </div>

            <input
              type="email"
              value={resetEmail}
              onChange={(e) => setResetEmail(e.target.value)}
              placeholder="deine@email.de"
              style={{
                width: "100%",
                background: "#0f172a",
                border: "1px solid #374151",
                borderRadius: 8,
                color: "#e5e7eb",
                padding: "10px 12px",
                fontSize: 14,
                boxSizing: "border-box",
                marginBottom: 12,
                outline: "none",
              }}
            />

            {resetMsg && (
              <div style={{
                background: "#0f2a1a",
                border: "1px solid #14532d",
                borderRadius: 8,
                padding: "10px 12px",
                fontSize: 12,
                color: "#4ade80",
                marginBottom: 12,
              }}>
                {resetMsg}
              </div>
            )}

            <button
              onClick={handleReset}
              style={{
                width: "100%",
                background: "#2563eb",
                color: "#fff",
                border: "none",
                borderRadius: 8,
                padding: "10px",
                fontSize: 13,
                fontWeight: 700,
                cursor: "pointer",
                marginBottom: 8,
              }}
            >
              Reset anfordern
            </button>

            <button
              onClick={() => { setShowReset(false); setResetMsg(""); }}
              style={{
                width: "100%",
                background: "transparent",
                color: "#6b7280",
                border: "none",
                fontSize: 12,
                cursor: "pointer",
                textDecoration: "underline",
              }}
            >
              Zurück zum Login
            </button>
          </>
        )}

        {/* Benutzer-Übersicht für Admin */}
        <div style={{ marginTop: 24, borderTop: "1px solid #1f2937", paddingTop: 16 }}>
          <div style={{ fontSize: 11, color: "#374151", textAlign: "center" }}>
            Standard-Zugänge (in Vercel Env-Variablen ändern)
          </div>
          <div style={{ fontSize: 11, color: "#4b5563", marginTop: 6 }}>
            <div>USER1_NAME / USER1_PASS</div>
            <div>USER2_NAME / USER2_PASS</div>
            <div>USER3_NAME / USER3_PASS</div>
          </div>
        </div>
      </div>
    </div>
  );
}
