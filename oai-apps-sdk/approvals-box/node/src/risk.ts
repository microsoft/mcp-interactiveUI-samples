// ─────────────────────────────────────────────────────────────────────────────
// Risk computation engine
// Risk level is ALWAYS computed here — never by the model or widget.
// ─────────────────────────────────────────────────────────────────────────────

import type { ApprovalType, RiskLevel, ReviewMode, RiskAssessment, ApprovalField } from "./types.js";

interface RiskInput {
  id: string;
  type: ApprovalType;
  amount: number;
  currency: string;
  country: string;
  fields: ApprovalField[];
  attachment_ids: string[];
  due_at: string;
  summary: string;
}

// Amount thresholds in USD equivalents
const HIGH_AMOUNT_THRESHOLD: Partial<Record<ApprovalType, number>> = {
  capex: 100_000,
  purchase_order: 50_000,
  contract_sow: 75_000,
  marketing_budget: 50_000,
  travel_exception: 5_000,
  vendor_onboarding: 100_000,
  promo_pricing: 0, // no amount on promo pricing — risk driven by other factors
};

const MEDIUM_AMOUNT_THRESHOLD: Partial<Record<ApprovalType, number>> = {
  capex: 25_000,
  purchase_order: 10_000,
  contract_sow: 20_000,
  marketing_budget: 15_000,
  travel_exception: 2_000,
  vendor_onboarding: 25_000,
  promo_pricing: 0,
};

// Cross-border/international countries that add medium risk
const INTERNATIONAL_COUNTRIES = new Set([
  "MEX", "MX", "MEXICO",
  "GBR", "UK", "GB",
  "DEU", "DE", "GERMANY",
  "FRA", "FR", "FRANCE",
  "JPN", "JP", "JAPAN",
  "CHN", "CN", "CHINA",
  "IND", "IN", "INDIA",
  "BRA", "BR", "BRAZIL",
  "AUS", "AU", "AUSTRALIA",
]);

// Domestic countries (USA-based)
const DOMESTIC_COUNTRIES = new Set(["USA", "US", "UNITED STATES"]);

function getFieldValue(fields: ApprovalField[], key: string): string | undefined {
  const keyLower = key.toLowerCase();
  return fields.find((f) => {
    const words = f.label.toLowerCase().split(/[\s/_()\-]+/);
    return words.some((w) => w === keyLower);
  })?.value;
}

function hasAttachments(attachment_ids: string[]): boolean {
  return attachment_ids.length > 0;
}

function isDueWithin(due_at: string, hours: number): boolean {
  const now = Date.now();
  const due = new Date(due_at).getTime();
  return due > now && due - now < hours * 60 * 60 * 1000;
}

function isOverdue(due_at: string): boolean {
  return new Date(due_at).getTime() < Date.now();
}

// ─────────────────────────────────────────────────────────────────────────────
// Main risk computation
// ─────────────────────────────────────────────────────────────────────────────

export function computeRisk(input: RiskInput): RiskAssessment {
  const reasons: string[] = [];
  let score = 0; // 0-100; ≥70 = high, ≥35 = medium, <35 = low

  const normalizedCountry = input.country.toUpperCase().trim();
  const highThreshold = HIGH_AMOUNT_THRESHOLD[input.type] ?? 50_000;
  const mediumThreshold = MEDIUM_AMOUNT_THRESHOLD[input.type] ?? 10_000;

  // ── Amount scoring ──────────────────────────────────────────────────────
  if (input.amount > 0) {
    if (input.amount >= highThreshold && highThreshold > 0) {
      score += 50; // high amount alone → high risk
      reasons.push(`High-value ${input.type.replace(/_/g, " ")} request (${input.currency} ${input.amount.toLocaleString()})`);
    } else if (input.amount >= mediumThreshold && mediumThreshold > 0) {
      score += 20;
      reasons.push(`Moderate amount (${input.currency} ${input.amount.toLocaleString()})`);
    }
  }

  // ── Cross-border ────────────────────────────────────────────────────────
  if (INTERNATIONAL_COUNTRIES.has(normalizedCountry)) {
    score += 20;
    reasons.push("Cross-border / international request");
  } else if (!DOMESTIC_COUNTRIES.has(normalizedCountry) && normalizedCountry !== "") {
    score += 10;
    reasons.push("Non-standard region");
  }

  // ── Missing documentation ───────────────────────────────────────────────
  if (!hasAttachments(input.attachment_ids)) {
    score += 20;
    reasons.push("No supporting documents attached");
  }

  // ── Type-specific rules ──────────────────────────────────────────────────

  switch (input.type) {
    case "promo_pricing": {
      const discount = getFieldValue(input.fields, "discount");
      if (discount && parseFloat(discount) > 30) {
        score += 25;
        reasons.push(`High discount percentage (${discount}%)`);
      }
      const publisher = getFieldValue(input.fields, "publisher");
      const retailer = getFieldValue(input.fields, "retailer");
      if (publisher && retailer) {
        score += 5; // two-party deal — minor flag
      }
      break;
    }

    case "capex": {
      const justification = getFieldValue(input.fields, "justification") ?? "";
      if (justification.length < 50) {
        score += 15;
        reasons.push("Capital expenditure justification is insufficient");
      }
      const roi = getFieldValue(input.fields, "roi");
      if (!roi) {
        score += 10;
        reasons.push("No expected ROI provided for capex request");
      }
      break;
    }

    case "contract_sow": {
      // Long-duration contracts are higher risk
      const startDate = getFieldValue(input.fields, "start");
      const endDate = getFieldValue(input.fields, "end");
      if (startDate && endDate) {
        const durationDays =
          (new Date(endDate).getTime() - new Date(startDate).getTime()) / (86_400_000);
        if (durationDays > 365) {
          score += 25;
          reasons.push("Multi-year contract (>12 months)");
        } else if (durationDays > 180) {
          score += 10;
          reasons.push("Long-term contract (>6 months)");
        }
      }
      break;
    }

    case "travel_exception": {
      const exceptionReason = getFieldValue(input.fields, "exception");
      if (!exceptionReason || exceptionReason.length < 20) {
        score += 20;
        reasons.push("Policy exception reason not sufficiently explained");
      }
      if (INTERNATIONAL_COUNTRIES.has(normalizedCountry)) {
        score += 10; // already added cross-border but compound for travel
        reasons.push("International travel requires additional approval");
      }
      break;
    }

    case "marketing_budget": {
      const channel = getFieldValue(input.fields, "channel") ?? "";
      if (channel.toLowerCase().includes("sponsor") || channel.toLowerCase().includes("event")) {
        score += 15;
        reasons.push("Sponsorship or event spend requires additional review");
      }
      break;
    }

    case "vendor_onboarding": {
      const vendorType = getFieldValue(input.fields, "vendor type") ?? getFieldValue(input.fields, "type") ?? "";
      if (vendorType.toLowerCase().includes("offshore") || vendorType.toLowerCase().includes("contractor")) {
        score += 20;
        reasons.push("Offshore or contractor vendor requires compliance review");
      }
      break;
    }

    case "purchase_order":
      // Base amount rules apply; no extra type-specific rules
      break;
  }

  // ── Urgency does NOT change risk level (risk = policy/financial, not schedule) ──

  // ── Compute level ────────────────────────────────────────────────────────
  // ≥50: high (large amount, multi-factor), ≥15: medium, <15: low
  const risk_level: RiskLevel = score >= 50 ? "high" : score >= 15 ? "medium" : "low";

  const recommended_review_mode: ReviewMode =
    risk_level === "low" ? "inline" : "fullscreen";

  return {
    approval_id: input.id,
    risk_level,
    risk_reasons: reasons.length > 0 ? reasons : ["Standard request with no elevated risk factors"],
    recommended_review_mode,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Allowed actions — computed by backend, never by model or widget
// ─────────────────────────────────────────────────────────────────────────────

export function computeAllowedActions(
  status: string,
  actorId: string,
  assigneeIds: string[],
): string[] {
  // Resolved (non-pending) approvals allow view only
  if (status !== "pending") {
    return ["view", "comment"];
  }

  const isAssignee = assigneeIds.includes(actorId);

  if (isAssignee) {
    return ["view", "approve", "reject", "comment"];
  }

  // Non-assignees can still view and comment
  return ["view", "comment"];
}
