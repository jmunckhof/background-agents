import type { Env } from "../types";
import { generateInternalToken } from "./internal";
import { createLogger } from "../logger";

const log = createLogger("integration-config");

export interface ResolvedLinearConfig {
  model: string | null;
  reasoningEffort: string | null;
  allowUserPreferenceOverride: boolean;
  allowLabelModelOverride: boolean;
  emitToolProgressActivities: boolean;
  enabledRepos: string[] | null;
  updateIssueStatus: boolean;
}

const DEFAULT_CONFIG: ResolvedLinearConfig = {
  model: null,
  reasoningEffort: null,
  allowUserPreferenceOverride: true,
  allowLabelModelOverride: true,
  emitToolProgressActivities: true,
  enabledRepos: null,
  updateIssueStatus: false,
};

export async function getLinearConfig(env: Env, repo: string): Promise<ResolvedLinearConfig> {
  if (!env.INTERNAL_CALLBACK_SECRET) {
    log.warn("config.fallback_defaults", { repo, reason: "no_callback_secret" });
    return DEFAULT_CONFIG;
  }

  const [owner, name] = repo.split("/");
  if (!owner || !name) {
    log.warn("config.fallback_defaults", { repo, reason: "invalid_repo_format" });
    return DEFAULT_CONFIG;
  }

  const token = await generateInternalToken(env.INTERNAL_CALLBACK_SECRET);

  let response: Response;
  try {
    response = await env.CONTROL_PLANE.fetch(
      `https://internal/integration-settings/linear/resolved/${owner}/${name}`,
      { headers: { Authorization: `Bearer ${token}` } }
    );
  } catch (err) {
    log.error("config.fetch_failed", {
      repo,
      error: err instanceof Error ? err : new Error(String(err)),
    });
    return DEFAULT_CONFIG;
  }

  if (!response.ok) {
    log.warn("config.fetch_not_ok", { repo, status: response.status });
    return DEFAULT_CONFIG;
  }

  const data = (await response.json()) as { config: ResolvedLinearConfig | null };
  if (!data.config) {
    log.warn("config.null_config", { repo });
    return DEFAULT_CONFIG;
  }

  log.debug("config.resolved", {
    repo,
    update_issue_status: data.config.updateIssueStatus,
  });
  return data.config;
}
