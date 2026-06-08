/**
 * Applies organizational policy constraints to incoming transactions.
 */
function applyOrganizationPolicies(sender, subject, aiResponse) {
  const entities = aiResponse.extracted_entities || {};
  const numericValues = entities.NumericValues || [];
  const category = aiResponse.category || "";
  
  // Rule 1: POL-001 - High-Value Financial Ceiling Constraint
  if (category === "Sponsorship" || category === "Registration") {
    for (let val of numericValues) {
      const parsedNum = parseFloat(val.replace(/[^0-9.]/g, ''));
      if (!isNaN(parsedNum) && parsedNum > 1000) {
        return {
          ruleId: "POL-001-HIGH-VALUE",
          overriddenAction: "ESCALATE_TO_HUMAN",
          reason: `Auto-override: Financial threshold exceeded ($${parsedNum} is greater than the allowed $1,000 limit). Requires human clearance.`
        };
      }
    }
  }

  // Rule 2: POL-002 - Restricted Domain Blocklist
  const domain = sender.split('@').pop().toLowerCase();
  const blockedDomains = ["spamoffer.com", "phishlink.net", "malware-domain.xyz"];
  for (let blocked of blockedDomains) {
    if (domain.includes(blocked)) {
      return {
        ruleId: "POL-002-BLACKLIST",
        overriddenAction: "DB_UPDATE_ONLY",
        reason: `Auto-override: Domain matches restricted database blacklists. Ignored safely.`
      };
    }
  }

  // Rule 3: POL-003 - Fraud & Security Safeguards
  const lowerContent = (aiResponse.justification + " " + subject).toLowerCase();
  const sensitiveWords = ["payment failed", "fraud", "chargeback", "security threat", "gdpr", "legal actions", "unauthorized charge"];
  for (let word of sensitiveWords) {
    if (lowerContent.includes(word)) {
      return {
        ruleId: "POL-003-SECURITY-SAFEGUARD",
        overriddenAction: "ESCALATE_TO_HUMAN",
        reason: `Auto-override: Sensitive legal/fraud keyword '${word}' identified. Bypassed direct AI response.`
      };
    }
  }

  return {
    ruleId: "POL-DEFAULT-AI-ROUTING",
    overriddenAction: null,
    reason: "Passed all constraints. Utilizing standard AI suggested routing workflow."
  };
}