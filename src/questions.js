// ─────────────────────────────────────────────────────────────────────────────
// QA Screening Questions — per-channel
// Each object maps to a SharePoint Choice column (Yes / No).
// Chat and SMS share the same question set.
// ─────────────────────────────────────────────────────────────────────────────

const SHARED_QUESTIONS = [
  {
    field: "Q06",
    category: "Opening",
    label: "Greeted the customer professionally and stated their name and department",
  },
  {
    field: "Q07",
    category: "Opening",
    label: "Verified customer identity appropriately before proceeding",
  },
  {
    field: "Q08",
    category: "Needs Identification",
    label: "Effectively gathered information to understand the customer's needs",
  },
  {
    field: "Q09",
    category: "Needs Identification",
    label: "Asked clarifying questions to ensure full understanding of the issue",
  },
  {
    field: "Q10",
    category: "Communication",
    label: "Communicated clearly and used language the customer could understand",
  },
  {
    field: "Q11",
    category: "Communication",
    label: "Avoided jargon, acronyms, or technical language without explanation",
  },
  // Q12 — channel-specific (inserted below)
  // Q13 — channel-specific (inserted below)
  {
    field: "Q14",
    category: "Empathy",
    label: "Used empathetic language to acknowledge the customer's situation",
  },
  {
    field: "Q15",
    category: "Empathy",
    label: "Remained calm and patient throughout the interaction",
  },
  // Q16 — channel-specific (inserted below)
  // Q17 — channel-specific (inserted below)
  // Q18 — channel-specific (inserted below)
  {
    field: "Q19",
    category: "Customer Focus",
    label: "Took ownership and accountability for resolving the customer's issue",
  },
  {
    field: "Q20",
    category: "Problem Solving",
    label: "Provided an accurate and complete response or solution",
  },
  {
    field: "Q21",
    category: "Problem Solving",
    label: "Escalated appropriately when the issue exceeded their authority or knowledge",
  },
  {
    field: "Q22",
    category: "Compliance",
    label: "Followed required scripts, disclosures, or compliance language where applicable",
  },
  {
    field: "Q23",
    category: "Compliance",
    label: "Documented the interaction and any actions taken accurately",
  },
  // Q24 — channel-specific (inserted below)
  {
    field: "Q25",
    category: "Closing",
    label: "Closed the interaction professionally and offered further assistance",
  },
];

// ── Channel-specific questions ──────────────────────────────────────────────

const CHANNEL_SPECIFIC = {
  Phone: {
    Q12: { field: "Q12", category: "Building Rapport", label: "Established rapport and built a positive connection with the customer" },
    Q13: { field: "Q13", category: "Building Rapport", label: "Demonstrated active listening (acknowledged, summarized, confirmed)" },
    Q16: { field: "Q16", category: "Professionalism", label: "Maintained a professional and courteous tone throughout the call" },
    Q17: { field: "Q17", category: "Professionalism", label: "Avoided interrupting or speaking over the customer" },
    Q18: { field: "Q18", category: "Customer Focus", label: "Prioritized the customer's needs and kept the call focused on resolution" },
    Q24: { field: "Q24", category: "Closing", label: "Confirmed the customer's issue was resolved before ending the call" },
  },
  Chat: {
    Q12: { field: "Q12", category: "Building Rapport", label: "Responded in a timely manner and maintained appropriate response times" },
    Q13: { field: "Q13", category: "Building Rapport", label: "Acknowledged customer messages and confirmed understanding before proceeding" },
    Q16: { field: "Q16", category: "Professionalism", label: "Maintained a professional and courteous tone throughout the conversation" },
    Q17: { field: "Q17", category: "Professionalism", label: "Used correct grammar, spelling, and punctuation" },
    Q18: { field: "Q18", category: "Customer Focus", label: "Prioritized the customer's needs and kept the conversation focused on resolution" },
    Q24: { field: "Q24", category: "Closing", label: "Confirmed the customer's issue was resolved before ending the conversation" },
  },
  Email: {
    Q12: { field: "Q12", category: "Building Rapport", label: "Responded within an appropriate timeframe" },
    Q13: { field: "Q13", category: "Building Rapport", label: "Addressed all points and questions raised by the customer" },
    Q16: { field: "Q16", category: "Professionalism", label: "Maintained a professional and courteous tone throughout the email" },
    Q17: { field: "Q17", category: "Professionalism", label: "Used correct grammar, spelling, and punctuation" },
    Q18: { field: "Q18", category: "Customer Focus", label: "Prioritized the customer's needs and kept the email focused on resolution" },
    Q24: { field: "Q24", category: "Closing", label: "Confirmed the customer's issue was resolved or provided clear next steps" },
  },
};

// SMS uses the same questions as Chat
CHANNEL_SPECIFIC.SMS = CHANNEL_SPECIFIC.Chat;

// ── Build the full 20-question list per channel ─────────────────────────────

function buildQuestions(channel) {
  const specific = CHANNEL_SPECIFIC[channel];
  const all = [
    ...SHARED_QUESTIONS.slice(0, 6),   // Q06–Q11
    specific.Q12,
    specific.Q13,
    ...SHARED_QUESTIONS.slice(6, 8),   // Q14–Q15
    specific.Q16,
    specific.Q17,
    specific.Q18,
    ...SHARED_QUESTIONS.slice(8, 12),  // Q19–Q23
    specific.Q24,
    ...SHARED_QUESTIONS.slice(12),     // Q25
  ];
  return all;
}

export const CHANNELS = ["Phone", "Chat", "Email", "SMS"];

export const QA_QUESTIONS_BY_CHANNEL = Object.fromEntries(
  CHANNELS.map((ch) => [ch, buildQuestions(ch)])
);

// Backward-compatible default export (Phone)
export const QA_QUESTIONS = QA_QUESTIONS_BY_CHANNEL.Phone;
