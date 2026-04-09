import React, { useState, useEffect, useMemo, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest, sharepointConfig } from "../authConfig";
import { COLORS, FONTS, GRADIENT } from "../brand";

const BACKEND_URL = "https://cxone-faq-bot-1.onrender.com";
const BOT_API_KEY = "tns-bot-secret-2024";
const GRAPH_SITE = "allstardriver.sharepoint.com:/sites/ServiceExcellenceDepartment-ALL-CustomerServiceTeam:";

// ── Styles ──────────────────────────────────────────────────────────────────

const s = {
  page: { minHeight: "100vh", background: COLORS.offWhite, padding: "24px 16px", fontFamily: FONTS.body },
  card: { maxWidth: 1100, margin: "0 auto", background: COLORS.white, borderRadius: 12, boxShadow: "0 4px 24px rgba(0,0,0,0.08)", overflow: "hidden" },
  header: { background: GRADIENT.orange, padding: "28px 32px", color: COLORS.white },
  headerTitle: { margin: 0, fontSize: 24, fontWeight: 700, fontFamily: FONTS.heading },
  headerSub: { margin: "6px 0 0", fontSize: 14, color: "rgba(255,255,255,0.8)" },
  body: { padding: "28px 32px" },

  topBar: { display: "flex", gap: 16, marginBottom: 24, alignItems: "center", flexWrap: "wrap" },
  generateBtn: (disabled) => ({
    padding: "10px 24px",
    background: disabled ? "#aaa" : GRADIENT.orange,
    color: COLORS.white,
    border: "none",
    borderRadius: 8,
    fontSize: 14,
    fontWeight: 700,
    fontFamily: FONTS.heading,
    cursor: disabled ? "not-allowed" : "pointer",
  }),
  dateInput: {
    padding: "8px 12px",
    borderRadius: 8,
    border: `1.5px solid ${COLORS.lightGray}`,
    fontSize: 14,
    fontFamily: FONTS.body,
  },
  label: { fontSize: 13, fontWeight: 600, color: COLORS.gray, marginBottom: 4 },

  statsRow: { display: "flex", gap: 16, marginBottom: 24, flexWrap: "wrap" },
  statCard: (color) => ({
    flex: "1 1 140px",
    padding: "16px",
    borderRadius: 10,
    background: COLORS.white,
    border: `2px solid ${color}`,
    textAlign: "center",
  }),
  statNum: (color) => ({ fontSize: 28, fontWeight: 800, fontFamily: FONTS.heading, color, margin: 0 }),
  statLabel: { fontSize: 11, color: COLORS.midGray, marginTop: 4, textTransform: "uppercase", letterSpacing: 0.5, fontWeight: 600 },

  // Evaluator section
  evalSection: { marginBottom: 24, border: `1px solid ${COLORS.lightGray}`, borderRadius: 10, overflow: "hidden" },
  evalHeader: {
    padding: "12px 16px",
    background: COLORS.charcoal,
    color: COLORS.white,
    fontFamily: FONTS.heading,
    fontWeight: 700,
    fontSize: 15,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    cursor: "pointer",
  },
  agentGroup: { padding: "16px", borderBottom: `1px solid ${COLORS.lightGray}` },
  agentName: { fontWeight: 700, color: COLORS.gray, fontSize: 15, fontFamily: FONTS.heading, marginBottom: 8 },

  // Interaction table
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: { textAlign: "left", padding: "6px 10px", background: "#F5F5F5", fontWeight: 600, fontSize: 11, textTransform: "uppercase", color: COLORS.midGray },
  td: { padding: "8px 10px", borderBottom: `1px solid ${COLORS.offWhite}`, color: COLORS.gray },
  contactId: { fontFamily: "monospace", fontSize: 13, fontWeight: 600, color: COLORS.sky },
  channelPill: (ch) => ({
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: 8,
    fontSize: 11,
    fontWeight: 600,
    background: "#FEF3E2",
    color: COLORS.orange,
  }),
  statusCheck: { width: 18, height: 18, cursor: "pointer", accentColor: COLORS.green },

  center: { textAlign: "center", padding: "48px 20px", color: COLORS.midGray, fontSize: 15 },
  errorBox: { background: COLORS.failBg, border: "1px solid #FFCDD2", borderRadius: 8, padding: "12px 16px", color: COLORS.fail, fontSize: 14, marginBottom: 16 },
  warningBox: { background: COLORS.warningBg, border: `1px solid ${COLORS.clementine}`, borderRadius: 8, padding: "12px 16px", fontSize: 13, color: "#795548", marginBottom: 16 },
};

// ── Helpers: SharePoint Graph API ───────────────────────────────────────────

const GRAPH_BASE = `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE}`;
const PREFER_HEADER = "HonorNonIndexedQueriesWarningMayFailRandomly";

async function graphFetch(token, path) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: PREFER_HEADER },
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph ${res.status}: ${err}`);
  }
  return res.json();
}

async function graphPost(token, path, body) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: PREFER_HEADER },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph POST ${res.status}: ${err}`);
  }
  return res.json();
}

async function graphPatch(token, path, body) {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method: "PATCH",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Prefer: PREFER_HEADER },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Graph PATCH ${res.status}: ${err}`);
  }
}

// Look up a SharePoint list ID by display name (more reliable than using name in URL)
const _listIdCache = {};
async function getListId(token, displayName) {
  if (_listIdCache[displayName]) return _listIdCache[displayName];
  // First try exact filter
  const data = await graphFetch(token, `/lists?$filter=displayName eq '${displayName}'`);
  const list = (data.value || [])[0];
  if (list) {
    _listIdCache[displayName] = list.id;
    return list.id;
  }
  // Not found — fetch all lists to help debug
  const allLists = await graphFetch(token, `/lists?$select=displayName,id&$top=100`);
  const names = (allLists.value || []).map((l) => l.displayName).join(", ");
  throw new Error(`List "${displayName}" not found. Available lists: ${names}`);
}

// Resolve SharePoint Person column LookupIds to user details
async function resolveUsers(token, lookupIds) {
  const userMap = {};
  for (const id of lookupIds) {
    if (!id || userMap[id]) continue;
    try {
      const user = await graphFetch(token,
        `/lists('User Information List')/items/${id}?$expand=fields($select=Title,EMail)`
      );
      userMap[id] = {
        name: user.fields?.Title || "",
        email: user.fields?.EMail || "",
      };
    } catch (e) {
      console.warn(`Could not resolve user lookupId ${id}:`, e.message);
      userMap[id] = { name: `User #${id}`, email: "" };
    }
  }
  return userMap;
}

async function fetchEvaluatorAgents(token) {
  const listId = await getListId(token, sharepointConfig.evaluatorAgentsListName);
  const data = await graphFetch(token,
    `/lists/${listId}/items?$expand=fields&$top=500`
  );
  const items = data.value || [];

  // If no items, return empty
  if (items.length === 0) return [];

  // Debug: inspect what field names SharePoint actually returns
  const sampleFields = items[0]?.fields || {};
  const fieldNames = Object.keys(sampleFields).filter((k) => !k.startsWith("@"));

  // Try to find Person column LookupId fields (could be Evaluator, Agent, or other names)
  const evalLookupKey = fieldNames.find((k) => /evaluator/i.test(k) && /lookupid/i.test(k));
  const agentLookupKey = fieldNames.find((k) => /agent/i.test(k) && /lookupid/i.test(k));

  // If no LookupId fields found, the columns might be text or named differently
  if (!evalLookupKey && !agentLookupKey) {
    // Try reading as plain text fields
    const evalNameKey = fieldNames.find((k) => /evaluator/i.test(k) && !/lookup/i.test(k));
    const agentNameKey = fieldNames.find((k) => /agent/i.test(k) && !/lookup/i.test(k) && !/cxone/i.test(k));

    if (evalNameKey || agentNameKey) {
      return items.map((item) => ({
        id: item.id,
        evaluatorName: item.fields?.[evalNameKey] || "",
        evaluatorEmail: "",
        agentName: item.fields?.[agentNameKey] || "",
        active: item.fields?.Active !== false,
      })).filter((r) => r.active && r.agentName);
    }

    // Nothing matched — throw helpful error with actual field names
    throw new Error(`Could not find Evaluator/Agent columns. Fields found: ${fieldNames.join(", ")}`);
  }

  // Person columns found — resolve LookupIds to names
  const lookupIds = new Set();
  items.forEach((item) => {
    if (evalLookupKey && item.fields?.[evalLookupKey]) lookupIds.add(item.fields[evalLookupKey]);
    if (agentLookupKey && item.fields?.[agentLookupKey]) lookupIds.add(item.fields[agentLookupKey]);
  });

  const userMap = await resolveUsers(token, lookupIds);

  return items
    .map((item) => {
      const evalId = evalLookupKey ? item.fields?.[evalLookupKey] : null;
      const agentId = agentLookupKey ? item.fields?.[agentLookupKey] : null;
      return {
        id: item.id,
        evaluatorName: userMap[evalId]?.name || "",
        evaluatorEmail: userMap[evalId]?.email || "",
        agentName: userMap[agentId]?.name || "",
        active: item.fields?.Active !== false,
      };
    })
    .filter((r) => r.active && r.agentName);
}

async function fetchAssignments(token, weekOf) {
  const listId = await getListId(token, sharepointConfig.assignmentsListName);
  const data = await graphFetch(token,
    `/lists/${listId}/items?$expand=fields&$top=500&$filter=fields/WeekOf eq '${weekOf}'`
  );
  return (data.value || []).map((item) => ({
    id: item.id,
    ...item.fields,
  }));
}

async function saveAssignment(token, fields) {
  const listId = await getListId(token, sharepointConfig.assignmentsListName);
  return graphPost(token,
    `/lists/${listId}/items`,
    { fields }
  );
}

async function updateStatus(token, itemId, status) {
  const listId = await getListId(token, sharepointConfig.assignmentsListName);
  return graphPatch(token,
    `/lists/${listId}/items/${itemId}/fields`,
    { Status: status }
  );
}

// ── Helper: get Monday of the week ─────────────────────────────────────────

function getMonday(date = new Date()) {
  const d = new Date(date);
  const day = d.getDay();
  d.setDate(d.getDate() - (day === 0 ? 6 : day - 1));
  return d.toISOString().split("T")[0];
}

function getNextMonday() {
  const d = new Date();
  d.setDate(d.getDate() + (8 - d.getDay()) % 7);
  return d.toISOString().split("T")[0];
}

// ── Component ───────────────────────────────────────────────────────────────

export default function Assignments() {
  const { instance, accounts } = useMsal();
  const [assignments, setAssignments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [generating, setGenerating] = useState(false);
  const [error, setError] = useState(null);
  const [genErrors, setGenErrors] = useState([]);
  const [dueDate, setDueDate] = useState(getNextMonday());
  const [weekOf, setWeekOf] = useState(getMonday());
  const [expandedEvals, setExpandedEvals] = useState({});

  // Get access token
  const getToken = useCallback(async () => {
    try {
      const res = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
      return res.accessToken;
    } catch {
      const res = await instance.acquireTokenPopup(loginRequest);
      return res.accessToken;
    }
  }, [instance, accounts]);

  // Load assignments for current week
  useEffect(() => {
    async function load() {
      setLoading(true);
      setError(null);
      try {
        const token = await getToken();
        const data = await fetchAssignments(token, weekOf);
        setAssignments(data);
      } catch (err) {
        setError(err.message);
      } finally {
        setLoading(false);
      }
    }
    load();
  }, [weekOf, getToken]);

  // Generate assignments
  async function handleGenerate() {
    setGenerating(true);
    setError(null);
    setGenErrors([]);

    try {
      const token = await getToken();

      // 1. Read evaluator-agent config from SharePoint
      // Debug: also fetch raw list data to diagnose issues
      const listId = await getListId(token, sharepointConfig.evaluatorAgentsListName);
      const rawData = await graphFetch(token, `/lists/${listId}/items?$expand=fields&$top=5`);
      const rawItems = rawData.value || [];
      const rawFieldNames = rawItems.length > 0 ? Object.keys(rawItems[0].fields || {}).filter(k => !k.startsWith("@")).join(", ") : "NO ITEMS";

      const config = await fetchEvaluatorAgents(token);
      if (config.length === 0) {
        setError(`No evaluator-agent mappings found. List has ${rawItems.length} items. Fields: ${rawFieldNames}`);
        setGenerating(false);
        return;
      }

      // 2. Call backend to generate from CXone
      const res = await fetch(`${BACKEND_URL}/api/assignments/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json", "X-API-Key": BOT_API_KEY },
        body: JSON.stringify({
          evaluator_agents: config.map((c) => ({
            evaluatorName: c.evaluatorName,
            evaluatorEmail: c.evaluatorEmail,
            agentName: c.agentName,
            agentCxoneId: c.agentCxoneId,
          })),
          screening_due_date: dueDate,
          interactions_per_agent: 5,
        }),
      });

      if (!res.ok) {
        const text = await res.text();
        throw new Error(`Backend error ${res.status}: ${text}`);
      }

      const result = await res.json();
      if (result.errors?.length) setGenErrors(result.errors);

      // 3. Save each interaction as a row in QA_Assignments
      const newWeekOf = result.weekOf;

      for (const assignment of result.assignments) {
        for (const interaction of assignment.interactions) {
          await saveAssignment(token, {
            WeekOf: newWeekOf,
            EvaluatorName: assignment.evaluatorName,
            AgentName: assignment.agentName,
            ContactId: String(interaction.contactId),
            Channel: interaction.channel,
            InteractionDate: interaction.interactionDate,
            Duration: Math.round(interaction.duration),
            SkillName: interaction.skillName || "",
            Status: "Pending",
            DueDate: dueDate,
          });
        }
      }

      // 4. Refresh the list
      setWeekOf(newWeekOf);
      const updated = await fetchAssignments(token, newWeekOf);
      setAssignments(updated);
    } catch (err) {
      setError(err.message);
    } finally {
      setGenerating(false);
    }
  }

  // Toggle status
  async function toggleStatus(item) {
    const newStatus = item.Status === "Completed" ? "Pending" : "Completed";
    try {
      const token = await getToken();
      await updateStatus(token, item.id, newStatus);
      setAssignments((prev) =>
        prev.map((a) => (a.id === item.id ? { ...a, Status: newStatus } : a))
      );
    } catch (err) {
      console.warn("Failed to update status:", err);
    }
  }

  // Group assignments by evaluator then agent
  const grouped = useMemo(() => {
    const map = {};
    assignments.forEach((a) => {
      const key = a.EvaluatorName || "Unassigned";
      if (!map[key]) map[key] = {};
      const agentKey = a.AgentName || "Unknown";
      if (!map[key][agentKey]) map[key][agentKey] = [];
      map[key][agentKey].push(a);
    });
    return map;
  }, [assignments]);

  const evaluators = Object.keys(grouped).sort();

  // Stats
  const stats = useMemo(() => {
    const total = assignments.length;
    const completed = assignments.filter((a) => a.Status === "Completed").length;
    const agents = new Set(assignments.map((a) => a.AgentName)).size;
    return { total, completed, pending: total - completed, agents };
  }, [assignments]);

  function toggleEval(name) {
    setExpandedEvals((prev) => ({ ...prev, [name]: !prev[name] }));
  }

  function formatDuration(sec) {
    if (!sec) return "-";
    const m = Math.floor(sec / 60);
    const sLeft = sec % 60;
    return `${m}m ${String(sLeft).padStart(2, "0")}s`;
  }

  function formatDate(dateStr) {
    if (!dateStr) return "-";
    try {
      return new Date(dateStr).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
    } catch { return dateStr; }
  }

  return (
    <div style={s.page}>
      <div style={s.card}>
        <div style={s.header}>
          <h1 style={s.headerTitle}>QA Assignments</h1>
          <p style={s.headerSub}>Weekly interaction assignments for quality screening</p>
        </div>

        <div style={s.body}>
          {/* Top controls */}
          <div style={s.topBar}>
            <div>
              <div style={s.label}>Screening Due Date</div>
              <input
                type="date"
                style={s.dateInput}
                value={dueDate}
                onChange={(e) => setDueDate(e.target.value)}
              />
            </div>
            <div>
              <div style={s.label}>Viewing Week Of</div>
              <input
                type="date"
                style={s.dateInput}
                value={weekOf}
                onChange={(e) => setWeekOf(e.target.value)}
              />
            </div>
            <div style={{ marginTop: 18 }}>
              <button
                style={s.generateBtn(generating)}
                onClick={handleGenerate}
                disabled={generating}
              >
                {generating ? "Generating..." : "Generate This Week's Assignments"}
              </button>
            </div>
          </div>

          {/* Errors / warnings */}
          {error && <div style={s.errorBox}>{"\u26A0"} {error}</div>}
          {genErrors.length > 0 && (
            <div style={s.warningBox}>
              <strong>Warnings:</strong>
              <ul style={{ margin: "8px 0 0", paddingLeft: 20 }}>
                {genErrors.map((e, i) => <li key={i}>{e}</li>)}
              </ul>
            </div>
          )}

          {loading ? (
            <div style={s.center}>
              <p>Loading assignments...</p>
              <style>{`@keyframes pulse { 0%,100% { opacity:.3 } 50% { opacity:1 } }`}</style>
            </div>
          ) : assignments.length === 0 ? (
            <div style={s.center}>
              <p style={{ fontSize: 18, fontFamily: FONTS.heading, color: COLORS.gray }}>
                No assignments for this week
              </p>
              <p>Click "Generate This Week's Assignments" to create them from CXone data.</p>
            </div>
          ) : (
            <>
              {/* Stats */}
              <div style={s.statsRow}>
                <div style={s.statCard(COLORS.orange)}>
                  <p style={s.statNum(COLORS.orange)}>{stats.total}</p>
                  <p style={s.statLabel}>Total</p>
                </div>
                <div style={s.statCard(COLORS.green)}>
                  <p style={s.statNum(COLORS.green)}>{stats.completed}</p>
                  <p style={s.statLabel}>Completed</p>
                </div>
                <div style={s.statCard(COLORS.sky)}>
                  <p style={s.statNum(COLORS.sky)}>{stats.pending}</p>
                  <p style={s.statLabel}>Pending</p>
                </div>
                <div style={s.statCard(COLORS.gray)}>
                  <p style={s.statNum(COLORS.gray)}>{stats.agents}</p>
                  <p style={s.statLabel}>Agents</p>
                </div>
              </div>

              {/* Assignments by evaluator */}
              {evaluators.map((evalName) => {
                const agents = grouped[evalName];
                const agentNames = Object.keys(agents).sort();
                const isOpen = expandedEvals[evalName] !== false; // Default open
                const evalTotal = agentNames.reduce((sum, a) => sum + agents[a].length, 0);
                const evalDone = agentNames.reduce(
                  (sum, a) => sum + agents[a].filter((x) => x.Status === "Completed").length, 0
                );

                return (
                  <div key={evalName} style={s.evalSection}>
                    <div style={s.evalHeader} onClick={() => toggleEval(evalName)}>
                      <span>{isOpen ? "\u25BC" : "\u25B6"} {evalName}</span>
                      <span style={{ fontSize: 12, fontWeight: 400, color: "rgba(255,255,255,0.7)" }}>
                        {evalDone}/{evalTotal} completed
                      </span>
                    </div>
                    {isOpen && agentNames.map((agentName) => (
                      <div key={agentName} style={s.agentGroup}>
                        <div style={s.agentName}>{agentName}</div>
                        <table style={s.table}>
                          <thead>
                            <tr>
                              <th style={s.th}>Contact ID</th>
                              <th style={s.th}>Channel</th>
                              <th style={s.th}>Date</th>
                              <th style={s.th}>Duration</th>
                              <th style={s.th}>Skill</th>
                              <th style={{ ...s.th, textAlign: "center" }}>Done</th>
                            </tr>
                          </thead>
                          <tbody>
                            {agents[agentName].map((item) => (
                              <tr key={item.id}>
                                <td style={{ ...s.td, ...s.contactId }}>{item.ContactId}</td>
                                <td style={s.td}><span style={s.channelPill(item.Channel)}>{item.Channel}</span></td>
                                <td style={s.td}>{formatDate(item.InteractionDate)}</td>
                                <td style={s.td}>{formatDuration(item.Duration)}</td>
                                <td style={s.td}>{item.SkillName || "-"}</td>
                                <td style={{ ...s.td, textAlign: "center" }}>
                                  <input
                                    type="checkbox"
                                    checked={item.Status === "Completed"}
                                    onChange={() => toggleStatus(item)}
                                    style={s.statusCheck}
                                  />
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ))}
                  </div>
                );
              })}
            </>
          )}
        </div>
      </div>
    </div>
  );
}
