import { sharepointConfig } from "./authConfig";

/**
 * Posts a new QA screening record to the SharePoint list.
 * @param {string} accessToken  - Bearer token from MSAL
 * @param {object} formData     - The form values to save
 */
export async function submitQARecord(accessToken, formData) {
  const { siteUrl, listName } = sharepointConfig;
  const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`;

  // Build the SharePoint item payload.
  // Choice field values are sent as plain strings ("Yes" / "No").
  const payload = {
    __metadata: { type: "SP.Data.QA_SupportPhonesListItem" },

    // Identification fields
    AgentName:    formData.AgentName,
    AgentEmail:   formData.AgentEmail,
    EvaluatorName: formData.EvaluatorName,
    Channel:      formData.Channel || "Phone",
    SubmissionDate: new Date().toISOString(),

    // 20 QA question fields (Choice: Yes / No)
    Q06: formData.Q06,
    Q07: formData.Q07,
    Q08: formData.Q08,
    Q09: formData.Q09,
    Q10: formData.Q10,
    Q11: formData.Q11,
    Q12: formData.Q12,
    Q13: formData.Q13,
    Q14: formData.Q14,
    Q15: formData.Q15,
    Q16: formData.Q16,
    Q17: formData.Q17,
    Q18: formData.Q18,
    Q19: formData.Q19,
    Q20: formData.Q20,
    Q21: formData.Q21,
    Q22: formData.Q22,
    Q23: formData.Q23,
    Q24: formData.Q24,
    Q25: formData.Q25,

    // Calculated score fields
    TotalScore:   formData.TotalScore,
    ScorePercent: formData.ScorePercent,
    PassFail:     formData.PassFail,

    // Open text
    SuggestionsForImprovement: formData.SuggestionsForImprovement || "",
  };

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`SharePoint error ${response.status}: ${errorText}`);
  }

  return await response.json();
}

/**
 * Sends the agent an email with their QA screening score via Microsoft Graph.
 * @param {string} accessToken - Bearer token from MSAL (needs Mail.Send scope)
 * @param {object} scoreData   - { agentName, agentEmail, evaluatorName, scorePercent, totalScore, passFail }
 */
export async function sendScoreEmail(accessToken, scoreData) {
  const { agentName, agentEmail, evaluatorName, channel = "Phone", scorePercent, totalScore, passFail } = scoreData;

  const passColor = passFail === "Pass" ? "#73BF45" : "#C62828";
  const passBg    = passFail === "Pass" ? "#EEF8E5" : "#FFEBEE";

  const htmlBody = `
    <div style="font-family:'Segoe UI',Arial,sans-serif;max-width:600px;margin:0 auto;">
      <div style="background:linear-gradient(135deg,#F58A21,#E07010);padding:24px 28px;border-radius:12px 12px 0 0;">
        <h2 style="color:#fff;margin:0;font-size:20px;">QA Screening Results</h2>
        <p style="color:rgba(255,255,255,0.8);margin:6px 0 0;font-size:13px;">The Next Street &middot; Customer Service</p>
      </div>
      <div style="background:#fff;padding:28px;border:1px solid #e8e8e8;border-top:none;border-radius:0 0 12px 12px;">
        <p style="color:#3B3B3B;font-size:15px;margin:0 0 16px;">
          Hi <strong>${agentName}</strong>,
        </p>
        <p style="color:#888;font-size:14px;margin:0 0 20px;">
          A QA screening for <strong style="color:#3B3B3B;">${channel}</strong> was completed for you by <strong style="color:#3B3B3B;">${evaluatorName}</strong>. Here are your results:
        </p>
        <div style="text-align:center;padding:20px;border-radius:10px;background:${passBg};border:2px solid ${passColor};margin:0 0 20px;">
          <div style="font-size:42px;font-weight:800;color:${passColor};">${scorePercent}%</div>
          <div style="font-size:13px;color:#888;margin:4px 0 10px;">${totalScore} / 100 points</div>
          <span style="display:inline-block;padding:5px 18px;border-radius:20px;background:${passColor};color:#fff;font-size:14px;font-weight:700;">
            ${passFail}
          </span>
        </div>
        <p style="color:#888;font-size:12px;margin:0;">
          If you have questions about this screening, please reach out to your supervisor.
        </p>
      </div>
    </div>
  `;

  const message = {
    message: {
      subject: `QA Screening Result (${channel}): ${passFail} (${scorePercent}%)`,
      body: {
        contentType: "HTML",
        content: htmlBody,
      },
      toRecipients: [
        {
          emailAddress: {
            address: agentEmail,
          },
        },
      ],
    },
    saveToSentItems: false,
  };

  const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(message),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Email send error ${response.status}: ${errorText}`);
  }
}
