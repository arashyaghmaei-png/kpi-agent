'use client';

import { useState } from "react";
import { useRouter } from "next/navigation";

const USERS = [
  { name: "arash", pass: "fibernc2024", display: "Arash (Admin)" },
  { name: "leitstelle", pass: "leitstelle123", display: "Leitstelle" },
  { name: "mitarbeiter", pass: "kpi2026", display: "Mitarbeiter" },
];

export default function LoginPage() {
  const router = useRouter();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const handleLogin = async (e) => {
    e.preventDefault();
    const user = USERS.find(u => u.name.toLowerCase() === username.toLowerCase() && u.pass === password);
    if (!user) {
      setError("Benutzername oder Passwort falsch.");
      return;
    }
    // Set cookie via API
    try {
      const res = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ username, password }),
      });
      if (res.ok) {
        router.push("/");
      } else {
        setError("Login fehlgeschlagen.");
      }
    } catch {
      // Fallback: direct cookie
      document.cookie = `auth_token=local_${Date.now()}; path=/; max-age=604800`;
      document.cookie = `auth_user=${user.display}; path=/; max-age=604800`;
      router.push("/");
    }
  };

  return (
    <div style={{ minHeight: "100vh", background: "#0a0f1a", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 16, padding: "40px 36px", width: 360 }}>
        <div style={{ textAlign: "center", marginBottom: 28 }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: "#f9fafb" }}>KPI AGENT</div>
          <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>FiberNC - Leitstelle</div>
        </div>
        <form onSubmit={handleLogin}>
          <div style={{ marginBottom: 14 }}>
            <div style={{ fontSize: 12, color: "#9ca3af", marginBottom: 5 }}>Benutzername</div>
            <input value={username} onChange={e => setUsername(e.target.value)}
              placeholder="z.B. arash"
              style={{ width: "100%", background: "#0f172a", border: "1px solid #374151", borderRadius: 8, color: "#e5e7eb", padding: "10px 12px", fontSize: 14, boxSizing: "border-box", outline: "none" }} />
          </div>
          <div style={{ marginBottom: 18 }}>
            <div style={{ fontSize: 12, color: "#9ca3af", marginBottom: 5 }}>Passwort</div>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)}
              placeholder="Passwort"
              style={{ width: "100%", background: "#0f172a", border: "1px solid #374151", borderRadius: 8, color: "#e5e7eb", padding: "10px 12px", fontSize: 14, boxSizing: "border-box", outline: "none" }} />
          </div>
          {error && <div style={{ background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: "10px", fontSize: 13, color: "#f87171", marginBottom: 14 }}>{error}</div>}
          <button type="submit" style={{ width: "100%", background: "#2563eb", color: "#fff", border: "none", borderRadius: 8, padding: 12, fontSize: 14, fontWeight: 700, cursor: "pointer" }}>
            Anmelden
          </button>
        </form>
      </div>
    </div>
  );
}
