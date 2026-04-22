import React, { useState } from "react";
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./authConfig";
import { COLORS, FONTS, GRADIENT } from "./brand";
import QAForm from "./components/QAForm";
import Dashboard from "./components/Dashboard";
import Assignments from "./components/Assignments";

const msalInstance = new PublicClientApplication(msalConfig);

// ── Sign-In Page ────────────────────────────────────────────────────────────

function SignInPage() {
  const { instance } = useMsal();

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: COLORS.offWhite,
        fontFamily: FONTS.body,
      }}
    >
      <div
        style={{
          background: COLORS.white,
          borderRadius: 12,
          boxShadow: "0 4px 24px rgba(0,0,0,0.08)",
          padding: "48px 40px",
          textAlign: "center",
          maxWidth: 400,
        }}
      >
        {/* TNS logo-colored circle */}
        <div
          style={{
            width: 64,
            height: 64,
            borderRadius: "50%",
            background: "#FEF3E2",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            margin: "0 auto 20px",
          }}
        >
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none">
            <path d="M12 2L4 7v10l8 5 8-5V7l-8-5z" fill={COLORS.orange} />
            <path d="M12 2v20M4 7l16 10M20 7L4 17" stroke={COLORS.white} strokeWidth="1.5" />
          </svg>
        </div>
        <h2 style={{ margin: "0 0 4px", color: COLORS.orange, fontSize: 22, fontFamily: FONTS.heading }}>
          The Next Street
        </h2>
        <p style={{ margin: "0 0 4px", color: COLORS.gray, fontSize: 16, fontWeight: 600, fontFamily: FONTS.heading }}>
          Quality Assurance
        </p>
        <p style={{ color: COLORS.midGray, margin: "0 0 28px", fontSize: 14 }}>
          Sign in with your Microsoft work account to continue.
        </p>
        <button
          onClick={() => instance.loginPopup(loginRequest)}
          style={{
            width: "100%",
            padding: "12px",
            background: GRADIENT.orange,
            color: COLORS.white,
            border: "none",
            borderRadius: 8,
            fontSize: 15,
            fontWeight: 600,
            fontFamily: FONTS.heading,
            cursor: "pointer",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: 10,
          }}
        >
          Sign in with Microsoft
        </button>
        <p style={{ color: COLORS.lightGray, fontSize: 12, margin: "16px 0 0" }}>
          The Next Street {"\u00B7"} Customer Service Department
        </p>
      </div>
    </div>
  );
}

// ── Navigation Bar ──────────────────────────────────────────────────────────

function NavBar({ page, setPage }) {
  const { accounts, instance } = useMsal();

  return (
    <div
      style={{
        background: COLORS.charcoal,
        color: "rgba(255,255,255,0.7)",
        fontSize: 13,
        padding: "0 20px",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        fontFamily: FONTS.body,
      }}
    >
      {/* Left: brand + nav tabs */}
      <div style={{ display: "flex", alignItems: "center", gap: 0 }}>
        <span
          style={{
            fontFamily: FONTS.heading,
            fontWeight: 700,
            color: COLORS.orange,
            fontSize: 14,
            marginRight: 24,
            padding: "10px 0",
          }}
        >
          TNS Quality Assurance
        </span>
        {[
          { key: "dashboard", label: "Dashboard" },
          { key: "assignments", label: "Assignments" },
          { key: "form", label: "New Screening" },
        ].map((tab) => (
          <button
            key={tab.key}
            onClick={() => setPage(tab.key)}
            style={{
              background: "none",
              border: "none",
              borderBottom: page === tab.key ? `2px solid ${COLORS.orange}` : "2px solid transparent",
              color: page === tab.key ? COLORS.white : "rgba(255,255,255,0.5)",
              padding: "10px 16px",
              fontSize: 13,
              fontWeight: 600,
              fontFamily: FONTS.heading,
              cursor: "pointer",
              transition: "all 0.15s",
            }}
          >
            {tab.label}
          </button>
        ))}
      </div>

      {/* Right: user + sign out */}
      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <span style={{ fontSize: 12 }}>
          {accounts[0]?.name || accounts[0]?.username}
        </span>
        <button
          onClick={() => instance.logoutPopup()}
          style={{
            background: "none",
            border: `1px solid rgba(255,255,255,0.2)`,
            color: "rgba(255,255,255,0.6)",
            padding: "3px 10px",
            borderRadius: 4,
            cursor: "pointer",
            fontSize: 12,
            fontFamily: FONTS.body,
          }}
        >
          Sign out
        </button>
      </div>
    </div>
  );
}

// ── App Content ─────────────────────────────────────────────────────────────

function AppContent() {
  const [page, setPage] = useState("dashboard");
  // prefill holds data passed from Assignments → QAForm (and the assignment id to mark completed)
  const [prefill, setPrefill] = useState(null);

  function openScreeningFromAssignment(data) {
    setPrefill(data);
    setPage("form");
  }

  function handleTabChange(newPage) {
    // When manually switching to "form" tab, clear any prefill (user wants a fresh screening)
    if (newPage === "form" && page !== "form") setPrefill(null);
    setPage(newPage);
  }

  return (
    <>
      <AuthenticatedTemplate>
        <NavBar page={page} setPage={handleTabChange} />
        {page === "dashboard" ? (
          <Dashboard />
        ) : page === "assignments" ? (
          <Assignments onScreen={openScreeningFromAssignment} />
        ) : (
          <QAForm prefill={prefill} onDone={() => { setPrefill(null); setPage("assignments"); }} />
        )}
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <SignInPage />
      </UnauthenticatedTemplate>
    </>
  );
}

export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <AppContent />
    </MsalProvider>
  );
}
