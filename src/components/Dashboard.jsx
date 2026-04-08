import React, { useState, useEffect, useMemo } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../authConfig";
import { sharepointConfig } from "../authConfig";
import { COLORS, FONTS, GRADIENT } from "../brand";

// ── Styles ──────────────────────────────────────────────────────────────────

const s = {
  page: {
    minHeight: "100vh",
    background: COLORS.offWhite,
    padding: "24px 16px",
    fontFamily: FONTS.body,
  },
  card: {
    maxWidth: 1100,
    margin: "0 auto",
    background: COLORS.white,
    borderRadius: 12,
    boxShadow: "0 4px 24px rgba(0,0,0,0.08)",
    overflow: "hidden",
  },
  header: {
    background: GRADIENT.orange,
    padding: "28px 32px",
    color: COLORS.white,
  },
  headerTitle: { margin: 0, fontSize: 24, fontWeight: 700, fontFamily: FONTS.heading },
  headerSub: { margin: "6px 0 0", fontSize: 14, color: "rgba(255,255,255,0.8)" },
  body: { padding: "28px 32px" },

  // Summary cards
  statsRow: { display: "flex", gap: 16, marginBottom: 28, flexWrap: "wrap" },
  statCard: (borderColor) => ({
    flex: "1 1 180px",
    padding: "20px",
    borderRadius: 10,
    background: COLORS.white,
    border: `2px solid ${borderColor}`,
    textAlign: "center",
  }),
  statNum: (color) => ({
    fontSize: 32,
    fontWeight: 800,
    fontFamily: FONTS.heading,
    color,
    margin: 0,
  }),
  statLabel: {
    fontSize: 12,
    color: COLORS.midGray,
    marginTop: 4,
    textTransform: "uppercase",
    letterSpacing: 0.5,
    fontWeight: 600,
  },

  // Filters
  filterRow: {
    display: "flex",
    gap: 12,
    marginBottom: 20,
    alignItems: "center",
    flexWrap: "wrap",
  },
  filterSelect: {
    padding: "8px 12px",
    borderRadius: 8,
    border: `1.5px solid ${COLORS.lightGray}`,
    fontSize: 14,
    fontFamily: FONTS.body,
    color: COLORS.gray,
    background: COLORS.white,
    cursor: "pointer",
  },
  filterLabel: {
    fontSize: 13,
    fontWeight: 600,
    color: COLORS.gray,
    marginRight: 4,
  },

  // Table
  table: {
    width: "100%",
    borderCollapse: "collapse",
    fontSize: 14,
  },
  th: {
    textAlign: "left",
    padding: "10px 12px",
    background: COLORS.charcoal,
    color: COLORS.white,
    fontWeight: 600,
    fontFamily: FONTS.heading,
    fontSize: 12,
    textTransform: "uppercase",
    letterSpacing: 0.5,
  },
  td: {
    padding: "10px 12px",
    borderBottom: `1px solid ${COLORS.lightGray}`,
    color: COLORS.gray,
  },
  badge: (pass) => ({
    display: "inline-block",
    padding: "2px 10px",
    borderRadius: 12,
    fontSize: 12,
    fontWeight: 700,
    fontFamily: FONTS.heading,
    background: pass ? COLORS.passBg : COLORS.failBg,
    color: pass ? COLORS.green : COLORS.fail,
  }),
  scoreCell: (pct) => ({
    fontWeight: 700,
    fontFamily: FONTS.heading,
    color: pct >= 80 ? COLORS.green : pct >= 70 ? COLORS.orange : COLORS.fail,
  }),

  // Loading / empty
  center: {
    textAlign: "center",
    padding: "48px 20px",
    color: COLORS.midGray,
    fontSize: 15,
  },
  loadingDot: {
    display: "inline-block",
    width: 8,
    height: 8,
    borderRadius: "50%",
    background: COLORS.orange,
    margin: "0 3px",
    animation: "pulse 1.2s ease-in-out infinite",
  },
};

// ── Helper: fetch QA records via Microsoft Graph API ────────────────────────

const GRAPH_SITE = "allstardriver.sharepoint.com:/sites/ServiceExcellenceDepartment-ALL-CustomerServiceTeam:";

async function fetchQARecords(accessToken) {
  const { listName } = sharepointConfig;
  const endpoint =
    `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE}/lists/${listName}/items` +
    `?expand=fields($select=AgentName,AgentEmail,EvaluatorName,Channel,TotalScore,ScorePercent,PassFail,SubmissionDate,SuggestionsForImprovement)` +
    `&$top=500&$orderby=fields/SubmissionDate desc`;

  const response = await fetch(endpoint, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph API error ${response.status}: ${text}`);
  }

  const data = await response.json();
  return (data.value || []).map((item) => {
    const f = item.fields || {};
    return {
      id: item.id,
      agentName: f.AgentName || "",
      agentEmail: f.AgentEmail || "",
      evaluatorName: f.EvaluatorName || "",
      channel: f.Channel || "Phone",
      scorePercent: f.ScorePercent ?? f.TotalScore ?? 0,
      passFail: f.PassFail || ((f.ScorePercent ?? 0) >= 80 ? "Pass" : "Fail"),
      date: f.SubmissionDate ? new Date(f.SubmissionDate) : null,
      suggestions: f.SuggestionsForImprovement || "",
    };
  });
}

// ── Component ───────────────────────────────────────────────────────────────

export default function Dashboard() {
  const { instance, accounts } = useMsal();
  const [records, setRecords] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [channelFilter, setChannelFilter] = useState("All");
  const [agentFilter, setAgentFilter] = useState("All");

  useEffect(() => {
    async function load() {
      try {
        let tokenResponse;
        try {
          tokenResponse = await instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0],
          });
        } catch {
          tokenResponse = await instance.acquireTokenPopup(loginRequest);
        }
        const data = await fetchQARecords(tokenResponse.accessToken);
        setRecords(data);
      } catch (err) {
        setError(err.message);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [instance, accounts]);

  // Filtered records
  const filtered = useMemo(() => {
    return records.filter((r) => {
      if (channelFilter !== "All" && r.channel !== channelFilter) return false;
      if (agentFilter !== "All" && r.agentName !== agentFilter) return false;
      return true;
    });
  }, [records, channelFilter, agentFilter]);

  // Unique agents for dropdown
  const agents = useMemo(() => {
    const names = [...new Set(records.map((r) => r.agentName).filter(Boolean))];
    return names.sort();
  }, [records]);

  // Summary stats (based on filtered)
  const stats = useMemo(() => {
    if (filtered.length === 0) return { total: 0, avgScore: 0, passRate: 0, channels: {} };
    const avg = Math.round(filtered.reduce((sum, r) => sum + r.scorePercent, 0) / filtered.length);
    const passCount = filtered.filter((r) => r.passFail === "Pass").length;
    const passRate = Math.round((passCount / filtered.length) * 100);

    const channels = {};
    filtered.forEach((r) => {
      if (!channels[r.channel]) channels[r.channel] = 0;
      channels[r.channel]++;
    });

    return { total: filtered.length, avgScore: avg, passRate, channels };
  }, [filtered]);

  return (
    <div style={s.page}>
      <div style={s.card}>
        <div style={s.header}>
          <h1 style={s.headerTitle}>QA Dashboard</h1>
          <p style={s.headerSub}>Quality Assurance screening results and trends</p>
        </div>

        <div style={s.body}>
          {loading ? (
            <div style={s.center}>
              <div>
                <span style={{ ...s.loadingDot, animationDelay: "0s" }} />
                <span style={{ ...s.loadingDot, animationDelay: "0.2s" }} />
                <span style={{ ...s.loadingDot, animationDelay: "0.4s" }} />
              </div>
              <p style={{ marginTop: 12 }}>Loading screenings from SharePoint...</p>
              <style>{`@keyframes pulse { 0%,100% { opacity:.3 } 50% { opacity:1 } }`}</style>
            </div>
          ) : error ? (
            <div style={s.center}>
              <p style={{ color: COLORS.fail }}>{"\u26A0"} {error}</p>
              <p style={{ fontSize: 13 }}>Make sure you have access to the SharePoint list.</p>
            </div>
          ) : records.length === 0 ? (
            <div style={s.center}>
              <p style={{ fontSize: 18, fontFamily: FONTS.heading, color: COLORS.gray }}>No screenings yet</p>
              <p>Submit your first QA screening to see results here.</p>
            </div>
          ) : (
            <>
              {/* Summary Stats */}
              <div style={s.statsRow}>
                <div style={s.statCard(COLORS.orange)}>
                  <p style={s.statNum(COLORS.orange)}>{stats.total}</p>
                  <p style={s.statLabel}>Total Screenings</p>
                </div>
                <div style={s.statCard(COLORS.sky)}>
                  <p style={s.statNum(COLORS.sky)}>{stats.avgScore}%</p>
                  <p style={s.statLabel}>Average Score</p>
                </div>
                <div style={s.statCard(COLORS.green)}>
                  <p style={s.statNum(COLORS.green)}>{stats.passRate}%</p>
                  <p style={s.statLabel}>Pass Rate</p>
                </div>
                <div style={s.statCard(COLORS.gray)}>
                  <p style={s.statNum(COLORS.gray)}>{agents.length}</p>
                  <p style={s.statLabel}>Agents Reviewed</p>
                </div>
              </div>

              {/* Channel breakdown pills */}
              {Object.keys(stats.channels).length > 1 && (
                <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
                  {Object.entries(stats.channels).map(([ch, count]) => (
                    <span
                      key={ch}
                      style={{
                        padding: "4px 12px",
                        borderRadius: 12,
                        fontSize: 12,
                        fontWeight: 600,
                        background: "#FEF3E2",
                        color: COLORS.orange,
                        fontFamily: FONTS.heading,
                      }}
                    >
                      {ch}: {count}
                    </span>
                  ))}
                </div>
              )}

              {/* Filters */}
              <div style={s.filterRow}>
                <span style={s.filterLabel}>Filter:</span>
                <select
                  style={s.filterSelect}
                  value={channelFilter}
                  onChange={(e) => setChannelFilter(e.target.value)}
                >
                  <option value="All">All Channels</option>
                  <option value="Phone">Phone</option>
                  <option value="Chat">Chat</option>
                  <option value="Email">Email</option>
                  <option value="SMS">SMS</option>
                </select>
                <select
                  style={s.filterSelect}
                  value={agentFilter}
                  onChange={(e) => setAgentFilter(e.target.value)}
                >
                  <option value="All">All Agents</option>
                  {agents.map((name) => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
                <span style={{ fontSize: 13, color: COLORS.midGray }}>
                  Showing {filtered.length} of {records.length} screenings
                </span>
              </div>

              {/* Results Table */}
              <div style={{ overflowX: "auto" }}>
                <table style={s.table}>
                  <thead>
                    <tr>
                      <th style={s.th}>Date</th>
                      <th style={s.th}>Agent</th>
                      <th style={s.th}>Channel</th>
                      <th style={s.th}>Score</th>
                      <th style={s.th}>Result</th>
                      <th style={s.th}>Evaluator</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filtered.map((r, i) => (
                      <tr key={r.id || i} style={{ background: i % 2 === 0 ? COLORS.white : "#FAFAFA" }}>
                        <td style={s.td}>
                          {r.date ? r.date.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—"}
                        </td>
                        <td style={{ ...s.td, fontWeight: 600 }}>{r.agentName}</td>
                        <td style={s.td}>
                          <span style={{
                            padding: "2px 8px",
                            borderRadius: 8,
                            fontSize: 12,
                            fontWeight: 600,
                            background: "#FEF3E2",
                            color: COLORS.orange,
                          }}>
                            {r.channel}
                          </span>
                        </td>
                        <td style={{ ...s.td, ...s.scoreCell(r.scorePercent) }}>{r.scorePercent}%</td>
                        <td style={s.td}>
                          <span style={s.badge(r.passFail === "Pass")}>{r.passFail}</span>
                        </td>
                        <td style={s.td}>{r.evaluatorName}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {filtered.length === 0 && (
                <div style={{ ...s.center, padding: "24px 20px" }}>
                  No screenings match the selected filters.
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
