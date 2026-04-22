import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../authConfig";
import { submitQARecord, sendScoreEmail, uploadAttachments, markAssignmentCompleted } from "../sharepointService";
import { QA_QUESTIONS_BY_CHANNEL, CHANNELS } from "../questions";
import { COLORS, FONTS, GRADIENT } from "../brand";

// ── Microsoft Graph people search ───────────────────────────────────────────

async function searchPeople(accessToken, query) {
  if (!query || query.length < 2) return [];
  const endpoint =
    `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${encodeURIComponent(query)}') or startswith(mail,'${encodeURIComponent(query)}')&$top=8&$select=displayName,mail,id`;
  try {
    const res = await fetch(endpoint, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!res.ok) return [];
    const data = await res.json();
    return (data.value || []).map((u) => ({
      name: u.displayName,
      email: u.mail || "",
      id: u.id,
    }));
  } catch {
    return [];
  }
}

// ── Type-ahead component ────────────────────────────────────────────────────

function UserTypeAhead({ value, email, onChange, onSelect, placeholder, label, accessToken }) {
  const [suggestions, setSuggestions] = useState([]);
  const [showDropdown, setShowDropdown] = useState(false);
  const [loading, setLoading] = useState(false);
  const debounceRef = useRef(null);
  const wrapperRef = useRef(null);

  const doSearch = useCallback(
    async (q) => {
      if (!accessToken || q.length < 2) {
        setSuggestions([]);
        return;
      }
      setLoading(true);
      const results = await searchPeople(accessToken, q);
      setSuggestions(results);
      setShowDropdown(results.length > 0);
      setLoading(false);
    },
    [accessToken]
  );

  function handleChange(e) {
    const val = e.target.value;
    onChange(val, "");
    clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => doSearch(val), 300);
  }

  function handleSelect(user) {
    onSelect(user.name, user.email);
    setShowDropdown(false);
    setSuggestions([]);
  }

  // Close dropdown when clicking outside
  useEffect(() => {
    function handleClick(e) {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target)) {
        setShowDropdown(false);
      }
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  return (
    <div ref={wrapperRef} style={{ position: "relative" }}>
      <label style={{ fontSize: 13, fontWeight: 600, color: COLORS.gray, marginBottom: 6, display: "block" }}>
        {label}
      </label>
      <input
        style={{
          padding: "10px 12px",
          border: `1.5px solid ${COLORS.lightGray}`,
          borderRadius: 8,
          fontSize: 14,
          outline: "none",
          fontFamily: FONTS.body,
          width: "100%",
          boxSizing: "border-box",
        }}
        value={value}
        onChange={handleChange}
        onFocus={() => suggestions.length > 0 && setShowDropdown(true)}
        placeholder={placeholder}
        autoComplete="off"
      />
      {email && (
        <div style={{ fontSize: 11, color: COLORS.midGray, marginTop: 3 }}>{email}</div>
      )}
      {showDropdown && (
        <div
          style={{
            position: "absolute",
            top: "100%",
            left: 0,
            right: 0,
            background: COLORS.white,
            border: `1.5px solid ${COLORS.lightGray}`,
            borderRadius: 8,
            boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
            zIndex: 100,
            maxHeight: 220,
            overflowY: "auto",
            marginTop: 4,
          }}
        >
          {loading && (
            <div style={{ padding: "10px 14px", fontSize: 13, color: COLORS.midGray }}>Searching...</div>
          )}
          {suggestions.map((user) => (
            <div
              key={user.id || user.email}
              onClick={() => handleSelect(user)}
              style={{
                padding: "10px 14px",
                cursor: "pointer",
                fontSize: 14,
                borderBottom: `1px solid ${COLORS.offWhite}`,
                transition: "background 0.1s",
              }}
              onMouseEnter={(e) => (e.currentTarget.style.background = "#FEF3E2")}
              onMouseLeave={(e) => (e.currentTarget.style.background = COLORS.white)}
            >
              <div style={{ fontWeight: 600, color: COLORS.gray }}>{user.name}</div>
              {user.email && (
                <div style={{ fontSize: 12, color: COLORS.midGray, marginTop: 2 }}>{user.email}</div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

const styles = {
  page: {
    minHeight: "100vh",
    background: COLORS.offWhite,
    padding: "24px 16px",
    fontFamily: FONTS.body,
  },
  card: {
    maxWidth: 860,
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

  row: { display: "flex", gap: 16, marginBottom: 16 },
  col: { flex: 1, display: "flex", flexDirection: "column" },
  label: { fontSize: 13, fontWeight: 600, color: COLORS.gray, marginBottom: 6 },
  input: {
    padding: "10px 12px",
    border: `1.5px solid ${COLORS.lightGray}`,
    borderRadius: 8,
    fontSize: 14,
    outline: "none",
    transition: "border-color 0.15s",
    fontFamily: FONTS.body,
  },

  sectionHeader: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    margin: "24px 0 12px",
  },
  sectionPill: {
    background: COLORS.orange,
    color: COLORS.white,
    fontSize: 11,
    fontWeight: 700,
    padding: "3px 10px",
    borderRadius: 12,
    letterSpacing: 0.5,
    textTransform: "uppercase",
    fontFamily: FONTS.heading,
  },
  sectionLine: { flex: 1, height: 1, background: COLORS.lightGray },

  questionRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: 12,
    padding: "10px 12px",
    borderRadius: 8,
    marginBottom: 6,
    transition: "background 0.1s",
  },
  questionNum: {
    width: 28,
    height: 28,
    minWidth: 28,
    borderRadius: "50%",
    background: "#FEF3E2",
    color: COLORS.orange,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: 12,
    fontWeight: 700,
    marginTop: 1,
    fontFamily: FONTS.heading,
  },
  questionText: { flex: 1, fontSize: 14, color: COLORS.gray, lineHeight: 1.5, paddingTop: 4 },
  toggleGroup: { display: "flex", gap: 6, marginTop: 4 },
  toggleBtn: (selected, variant) => ({
    padding: "5px 16px",
    borderRadius: 20,
    border: "1.5px solid",
    cursor: "pointer",
    fontSize: 13,
    fontWeight: 600,
    transition: "all 0.15s",
    fontFamily: FONTS.body,
    borderColor: variant === "Yes"
      ? (selected ? COLORS.green : "#ccc")
      : (selected ? COLORS.fail : "#ccc"),
    background: variant === "Yes"
      ? (selected ? COLORS.passBg : COLORS.white)
      : (selected ? COLORS.failBg : COLORS.white),
    color: variant === "Yes"
      ? (selected ? COLORS.green : "#999")
      : (selected ? COLORS.fail : "#999"),
  }),

  scoreBar: {
    margin: "24px 0",
    padding: "16px 20px",
    borderRadius: 10,
    border: "2px solid",
    display: "flex",
    alignItems: "center",
    gap: 20,
  },
  scoreNum: { fontSize: 36, fontWeight: 800, lineHeight: 1, fontFamily: FONTS.heading },
  scoreSub: { fontSize: 12, color: COLORS.midGray, marginTop: 2 },
  scoreBadge: {
    padding: "4px 14px",
    borderRadius: 20,
    fontSize: 13,
    fontWeight: 700,
    fontFamily: FONTS.heading,
  },
  progressTrack: {
    flex: 1,
    height: 10,
    background: "#eee",
    borderRadius: 5,
    overflow: "hidden",
  },

  textarea: {
    width: "100%",
    padding: "10px 12px",
    border: `1.5px solid ${COLORS.lightGray}`,
    borderRadius: 8,
    fontSize: 14,
    resize: "vertical",
    minHeight: 90,
    outline: "none",
    fontFamily: FONTS.body,
    boxSizing: "border-box",
  },

  submitBtn: (disabled) => ({
    display: "block",
    width: "100%",
    padding: "14px",
    marginTop: 24,
    background: disabled ? "#aaa" : GRADIENT.orange,
    color: COLORS.white,
    border: "none",
    borderRadius: 8,
    fontSize: 16,
    fontWeight: 700,
    fontFamily: FONTS.heading,
    cursor: disabled ? "not-allowed" : "pointer",
    transition: "opacity 0.2s",
  }),

  successBox: {
    textAlign: "center",
    padding: "48px 32px",
  },
  successCircle: {
    width: 72,
    height: 72,
    borderRadius: "50%",
    background: COLORS.passBg,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    margin: "0 auto 16px",
    fontSize: 36,
  },
  errorBox: {
    background: COLORS.failBg,
    border: "1px solid #FFCDD2",
    borderRadius: 8,
    padding: "12px 16px",
    color: COLORS.fail,
    fontSize: 14,
    marginTop: 16,
  },
};

function scoreColor(pct) {
  if (pct >= 90) return { border: COLORS.green, text: COLORS.green, bar: COLORS.green, bg: COLORS.passBg };
  if (pct >= 80) return { border: COLORS.green, text: COLORS.green, bar: "#66BB6A", bg: COLORS.passBg };
  if (pct >= 70) return { border: COLORS.orange, text: COLORS.orange, bar: COLORS.clementine, bg: COLORS.warningBg };
  return { border: COLORS.fail, text: COLORS.fail, bar: "#EF5350", bg: COLORS.failBg };
}

export default function QAForm({ prefill, onDone }) {
  const { instance, accounts } = useMsal();
  const [accessToken, setAccessToken] = useState(null);

  // Auto-fill evaluator from signed-in user and get access token
  const signedInName = accounts[0]?.name || accounts[0]?.username || "";

  useEffect(() => {
    async function getToken() {
      try {
        const tokenRes = await instance.acquireTokenSilent({
          ...loginRequest,
          account: accounts[0],
        });
        setAccessToken(tokenRes.accessToken);
      } catch {
        // Token will be acquired on submit if silent fails
      }
    }
    if (accounts[0]) getToken();
  }, [instance, accounts]);

  const [channel, setChannel] = useState(prefill?.channel || "Phone");
  const questions = QA_QUESTIONS_BY_CHANNEL[channel];
  const categories = useMemo(() => [...new Set(questions.map((q) => q.category))], [questions]);

  const initialAnswers = Object.fromEntries(questions.map((q) => [q.field, null]));
  const [answers, setAnswers] = useState(initialAnswers);
  const [agentName, setAgentName] = useState(prefill?.agentName || "");
  const [agentEmail, setAgentEmail] = useState(prefill?.agentEmail || "");
  const [evaluatorName, setEvaluatorName] = useState(signedInName);
  const [suggestions, setSuggestions] = useState("");
  const [attachments, setAttachments] = useState([]); // File[] from <input type="file">
  const [submitting, setSubmitting] = useState(false);
  const [submitted, setSubmitted] = useState(false);
  const [error, setError] = useState(null);

  // Pre-fill context info shown at top of form when coming from an assignment
  const contactId = prefill?.contactId || "";
  const skillName = prefill?.skillName || "";
  const assignmentId = prefill?.assignmentId || null;
  const interactionDate = prefill?.interactionDate || "";

  function handleChannelChange(newChannel) {
    setChannel(newChannel);
    setAnswers(Object.fromEntries(QA_QUESTIONS_BY_CHANNEL[newChannel].map((q) => [q.field, null])));
  }

  const { totalScore, scorePercent, passFail, answered } = useMemo(() => {
    const yesCount = questions.filter((q) => answers[q.field] === "Yes").length;
    const total = yesCount * 5;
    const pct = total;
    return {
      totalScore: total,
      scorePercent: pct,
      passFail: pct >= 80 ? "Pass" : "Fail",
      answered: questions.filter((q) => answers[q.field] !== null).length,
    };
  }, [answers, questions]);

  const allAnswered = answered === questions.length && agentName.trim() && agentEmail.trim() && evaluatorName.trim();
  const colors = scoreColor(scorePercent);

  async function doSubmission(accessToken) {
    // 1. Save the QA record and capture the newly-created item id (needed for attachments)
    const record = await submitQARecord(accessToken, {
      ...answers,
      AgentName: agentName.trim(),
      AgentEmail: agentEmail.trim(),
      EvaluatorName: evaluatorName.trim(),
      Channel: channel,
      ContactId: contactId,
      InteractionDate: interactionDate,
      SuggestionsForImprovement: suggestions.trim(),
      TotalScore: totalScore,
      ScorePercent: scorePercent,
      PassFail: passFail,
    });

    // 2. Upload attachments if any were chosen
    if (attachments.length > 0 && record?.Id) {
      try {
        await uploadAttachments(accessToken, record.Id, attachments);
      } catch (attErr) {
        console.warn("Attachment upload failed:", attErr.message);
      }
    }

    // 3. Send the score email (best-effort)
    try {
      await sendScoreEmail(accessToken, {
        agentName: agentName.trim(),
        agentEmail: agentEmail.trim(),
        evaluatorName: evaluatorName.trim(),
        channel,
        scorePercent,
        totalScore,
        passFail,
      });
    } catch (emailErr) {
      console.warn("Score email could not be sent:", emailErr.message);
    }

    // 4. If this screening came from an assignment, mark it completed (best-effort)
    if (assignmentId) {
      try {
        await markAssignmentCompleted(accessToken, assignmentId);
      } catch (markErr) {
        console.warn("Could not mark assignment completed:", markErr.message);
      }
    }
  }

  async function handleSubmit(e) {
    e.preventDefault();
    if (!allAnswered) return;
    setSubmitting(true);
    setError(null);

    try {
      const tokenResponse = await instance.acquireTokenSilent({
        ...loginRequest,
        account: accounts[0],
      });
      await doSubmission(tokenResponse.accessToken);
      setSubmitted(true);
    } catch (err) {
      if (err.name === "InteractionRequiredAuthError") {
        try {
          const tokenResponse = await instance.acquireTokenPopup(loginRequest);
          await doSubmission(tokenResponse.accessToken);
          setSubmitted(true);
        } catch (popupErr) {
          setError(popupErr.message);
        }
      } else {
        setError(err.message);
      }
    } finally {
      setSubmitting(false);
    }
  }

  function resetForm() {
    setChannel("Phone");
    setAnswers(Object.fromEntries(QA_QUESTIONS_BY_CHANNEL.Phone.map((q) => [q.field, null])));
    setAgentName("");
    setAgentEmail("");
    setEvaluatorName(signedInName);
    setSuggestions("");
    setAttachments([]);
    setSubmitted(false);
    setError(null);
  }

  if (submitted) {
    return (
      <div style={styles.page}>
        <div style={styles.card}>
          <div style={styles.header}>
            <h1 style={styles.headerTitle}>Support Quality Assurance</h1>
          </div>
          <div style={styles.successBox}>
            <div style={styles.successCircle}>{"\u2714"}</div>
            <h2 style={{ margin: "0 0 8px", color: COLORS.orange, fontFamily: FONTS.heading }}>Screening Submitted</h2>
            <p style={{ color: COLORS.gray, margin: "0 0 8px" }}>
              <strong>{agentName}</strong> {"\u00B7"} {channel} {"\u00B7"} evaluated by <strong>{evaluatorName}</strong>
            </p>
            <div
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 10,
                padding: "10px 24px",
                borderRadius: 10,
                background: colors.bg,
                border: `2px solid ${colors.border}`,
                margin: "12px 0 24px",
              }}
            >
              <span style={{ fontSize: 28, fontWeight: 800, color: colors.text, fontFamily: FONTS.heading }}>
                {scorePercent}%
              </span>
              <span
                style={{
                  ...styles.scoreBadge,
                  background: colors.border,
                  color: COLORS.white,
                }}
              >
                {passFail}
              </span>
            </div>
            <br />
            <div style={{ display: "flex", gap: 10, justifyContent: "center" }}>
              {assignmentId && (
                <button
                  onClick={() => onDone && onDone()}
                  style={{
                    padding: "10px 28px",
                    background: COLORS.green,
                    color: COLORS.white,
                    border: "none",
                    borderRadius: 8,
                    fontSize: 14,
                    fontWeight: 600,
                    fontFamily: FONTS.heading,
                    cursor: "pointer",
                  }}
                >
                  Back to Assignments
                </button>
              )}
              <button
                onClick={resetForm}
                style={{
                  padding: "10px 28px",
                  background: COLORS.orange,
                  color: COLORS.white,
                  border: "none",
                  borderRadius: 8,
                  fontSize: 14,
                  fontWeight: 600,
                  fontFamily: FONTS.heading,
                  cursor: "pointer",
                }}
              >
                Start New Screening
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      <div style={styles.card}>
        {/* Header */}
        <div style={styles.header}>
          <h1 style={styles.headerTitle}>Support Quality Assurance</h1>
          <p style={styles.headerSub}>
            {channel} {"\u00B7"} 20 criteria {"\u00B7"} 5 points each {"\u00B7"} 100 points max {"\u00B7"} Pass threshold: 80%
          </p>
        </div>

        <form onSubmit={handleSubmit} style={styles.body}>
          {/* Assignment context banner */}
          {contactId && (
            <div
              style={{
                background: "#FEF3E2",
                border: `1.5px solid ${COLORS.orange}`,
                borderRadius: 8,
                padding: "10px 14px",
                marginBottom: 16,
                fontSize: 13,
                color: COLORS.gray,
              }}
            >
              <strong style={{ color: COLORS.orange }}>Screening from Assignment</strong>
              <div style={{ marginTop: 4, fontSize: 12 }}>
                Contact ID: <span style={{ fontFamily: "monospace", fontWeight: 600 }}>{contactId}</span>
                {skillName ? <> {"\u00B7"} Skill: {skillName}</> : null}
              </div>
            </div>
          )}

          {/* Channel selector */}
          <div style={{ marginBottom: 16 }}>
            <label style={styles.label}>Channel *</label>
            <div style={{ display: "flex", gap: 8 }}>
              {CHANNELS.map((ch) => (
                <button
                  key={ch}
                  type="button"
                  onClick={() => handleChannelChange(ch)}
                  style={{
                    padding: "8px 20px",
                    borderRadius: 20,
                    border: "2px solid",
                    borderColor: channel === ch ? COLORS.orange : COLORS.lightGray,
                    background: channel === ch ? "#FEF3E2" : COLORS.white,
                    color: channel === ch ? COLORS.orange : COLORS.midGray,
                    fontSize: 14,
                    fontWeight: 600,
                    fontFamily: FONTS.heading,
                    cursor: "pointer",
                    transition: "all 0.15s",
                  }}
                >
                  {ch}
                </button>
              ))}
            </div>
          </div>

          {/* Agent / Evaluator */}
          <div style={styles.row}>
            <div style={{ ...styles.col, flex: 2 }}>
              <UserTypeAhead
                label="Agent Name *"
                value={agentName}
                email={agentEmail}
                onChange={(name, email) => { setAgentName(name); setAgentEmail(email); }}
                onSelect={(name, email) => { setAgentName(name); setAgentEmail(email); }}
                placeholder="Start typing agent name..."
                accessToken={accessToken}
              />
            </div>
            <div style={styles.col}>
              <label style={styles.label}>Evaluator</label>
              <input
                style={{ ...styles.input, background: "#F5F5F5", color: COLORS.midGray }}
                value={evaluatorName}
                readOnly
                tabIndex={-1}
              />
            </div>
          </div>

          {/* Live score bar */}
          {answered > 0 && (
            <div style={{ ...styles.scoreBar, borderColor: colors.border, background: colors.bg }}>
              <div>
                <div style={{ ...styles.scoreNum, color: colors.text }}>{scorePercent}%</div>
                <div style={styles.scoreSub}>
                  {totalScore} / 100 pts {"\u00B7"} {answered}/{questions.length} answered
                </div>
              </div>
              <div style={styles.progressTrack}>
                <div
                  style={{
                    height: "100%",
                    width: `${scorePercent}%`,
                    background: colors.bar,
                    borderRadius: 5,
                    transition: "width 0.3s",
                  }}
                />
              </div>
              <span
                style={{
                  ...styles.scoreBadge,
                  background: colors.border,
                  color: COLORS.white,
                }}
              >
                {passFail}
              </span>
            </div>
          )}

          {/* Questions grouped by category */}
          {categories.map((cat) => {
            const qs = questions.filter((q) => q.category === cat);
            return (
              <div key={cat}>
                <div style={styles.sectionHeader}>
                  <span style={styles.sectionPill}>{cat}</span>
                  <div style={styles.sectionLine} />
                </div>
                {qs.map((q, idx) => {
                  const globalIdx = questions.findIndex((x) => x.field === q.field);
                  const isEven = globalIdx % 2 === 0;
                  return (
                    <div
                      key={q.field}
                      style={{
                        ...styles.questionRow,
                        background: isEven ? "#FEF9F3" : COLORS.white,
                      }}
                    >
                      <div style={styles.questionNum}>{globalIdx + 1}</div>
                      <div style={{ ...styles.questionText }}>
                        {q.label}
                        <div style={styles.toggleGroup}>
                          {["Yes", "No"].map((opt) => (
                            <button
                              key={opt}
                              type="button"
                              style={styles.toggleBtn(answers[q.field] === opt, opt)}
                              onClick={() =>
                                setAnswers((prev) => ({ ...prev, [q.field]: opt }))
                              }
                            >
                              {opt}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}

          {/* Suggestions */}
          <div style={{ marginTop: 24 }}>
            <label style={styles.label}>Suggestions for Improvement</label>
            <textarea
              style={styles.textarea}
              value={suggestions}
              onChange={(e) => setSuggestions(e.target.value)}
              placeholder="Optional — specific feedback for the agent..."
            />
          </div>

          {/* Attachments */}
          <div style={{ marginTop: 20 }}>
            <label style={styles.label}>Attachments (call recording, transcript, etc.)</label>
            <input
              type="file"
              multiple
              onChange={(e) => setAttachments(Array.from(e.target.files || []))}
              style={{
                display: "block",
                width: "100%",
                padding: "10px 12px",
                border: `1.5px dashed ${COLORS.lightGray}`,
                borderRadius: 8,
                fontSize: 13,
                background: COLORS.offWhite,
                fontFamily: FONTS.body,
                cursor: "pointer",
                boxSizing: "border-box",
              }}
            />
            {attachments.length > 0 && (
              <div style={{ marginTop: 8, fontSize: 12, color: COLORS.midGray }}>
                {attachments.length} file{attachments.length > 1 ? "s" : ""} selected:{" "}
                {attachments.map((f) => f.name).join(", ")}
              </div>
            )}
          </div>

          {/* Error */}
          {error && <div style={styles.errorBox}>{"\u26A0\uFE0F"} {error}</div>}

          {/* Unanswered warning */}
          {answered < questions.length && answered > 0 && (
            <div
              style={{
                background: COLORS.warningBg,
                border: `1px solid ${COLORS.clementine}`,
                borderRadius: 8,
                padding: "10px 14px",
                fontSize: 13,
                color: "#795548",
                marginTop: 16,
              }}
            >
              {questions.length - answered} question
              {questions.length - answered > 1 ? "s" : ""} still need an answer before submitting.
            </div>
          )}

          <button type="submit" style={styles.submitBtn(!allAnswered || submitting)} disabled={!allAnswered || submitting}>
            {submitting ? "Submitting\u2026" : "Submit Screening"}
          </button>
        </form>
      </div>
    </div>
  );
}
