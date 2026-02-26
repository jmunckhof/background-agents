/**
 * Build Adaptive Card messages for completion notifications.
 */

import type { AgentResponse } from "../types";
import type { ManualPullRequestArtifactMetadata } from "@open-inspect/shared";

const TRUNCATE_LIMIT = 2000;
const FALLBACK_TEXT_LIMIT = 150;

/**
 * Build an Adaptive Card for completion message.
 */
export function buildCompletionCard(
  sessionId: string,
  response: AgentResponse,
  repoFullName: string,
  model: string,
  reasoningEffort: string | undefined,
  webAppUrl: string
): Record<string, unknown> {
  const body: unknown[] = [];

  // 1. Response text (truncated)
  const text = truncateText(response.textContent, TRUNCATE_LIMIT);
  body.push({
    type: "TextBlock",
    text: text || "_Agent completed._",
    wrap: true,
  });

  // 2. Artifacts (PRs, branches)
  if (response.artifacts.length > 0) {
    const artifactText = response.artifacts
      .map((a) => (a.url ? `- [${a.label}](${a.url})` : `- ${a.label}`))
      .join("\n");
    body.push({
      type: "TextBlock",
      text: `**Created:**\n${artifactText}`,
      wrap: true,
    });
  }

  // 3. Key tool actions
  const keyToolNames = ["Edit", "Write", "Bash"] as const;
  const keyTools = response.toolCalls
    .filter((t) => keyToolNames.includes(t.tool as (typeof keyToolNames)[number]))
    .slice(0, 5);
  if (keyTools.length > 0) {
    body.push({
      type: "TextBlock",
      text: keyTools.map((t) => t.summary).join(" | "),
      wrap: true,
      isSubtle: true,
      size: "Small",
    });
  }

  // 4. Status footer
  const statusIcon = response.success ? "\u2705" : "\u26a0\ufe0f";
  const status = response.success ? "Done" : "Completed with issues";
  const effortSuffix = reasoningEffort ? ` (${reasoningEffort})` : "";
  body.push({
    type: "TextBlock",
    text: `${statusIcon} ${status}  |  ${model}${effortSuffix}  |  ${repoFullName}`,
    wrap: true,
    isSubtle: true,
    size: "Small",
  });

  // 5. Action buttons
  const actions: unknown[] = [
    {
      type: "Action.OpenUrl",
      title: "View Session",
      url: `${webAppUrl}/session/${sessionId}`,
    },
  ];

  const hasPrArtifact = response.artifacts.some((artifact) => artifact.type === "pr");
  const manualCreatePrUrl = getManualCreatePrUrl(response.artifacts);
  if (!hasPrArtifact && manualCreatePrUrl) {
    actions.push({
      type: "Action.OpenUrl",
      title: "Create PR",
      url: manualCreatePrUrl,
    });
  }

  return {
    type: "AdaptiveCard",
    $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
    version: "1.4",
    body,
    actions,
  };
}

/**
 * Get truncated text for fallback display.
 */
export function getFallbackText(response: AgentResponse): string {
  return response.textContent.slice(0, FALLBACK_TEXT_LIMIT) || "Agent completed.";
}

function truncateText(text: string, maxLen: number): string {
  if (text.length <= maxLen) return text;
  const truncated = text.slice(0, maxLen);
  const lastPeriod = truncated.lastIndexOf(". ");
  if (lastPeriod > maxLen * 0.7) {
    return truncated.slice(0, lastPeriod + 1) + "\n\n_...truncated_";
  }
  return truncated + "...\n\n_...truncated_";
}

function getManualCreatePrUrl(artifacts: AgentResponse["artifacts"]): string | null {
  const manualBranchArtifact = artifacts.find((artifact) => {
    if (artifact.type !== "branch") return false;
    if (!artifact.metadata || typeof artifact.metadata !== "object") return false;
    const metadata = artifact.metadata as Partial<ManualPullRequestArtifactMetadata> &
      Record<string, unknown>;
    if (metadata.mode === "manual_pr") return true;
    return metadata.mode == null && typeof metadata.createPrUrl === "string";
  });

  if (!manualBranchArtifact) return null;

  const metadataUrl = manualBranchArtifact.metadata?.createPrUrl;
  if (typeof metadataUrl === "string" && metadataUrl.length > 0) return metadataUrl;

  return manualBranchArtifact.url || null;
}
