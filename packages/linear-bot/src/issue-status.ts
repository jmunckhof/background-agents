/**
 * Issue workflow state transitions.
 *
 * Resolves target states by **position-based type resolution** (lowest position
 * within the target WorkflowStateType) per Linear's best practices.
 * Also handles setting the bot as issue delegate.
 */

import type { Env } from "./types";
import type { LinearApiClient, WorkflowState } from "./utils/linear-client";
import {
  fetchTeamWorkflowStates,
  updateIssueState,
  setIssueDelegate,
  fetchBotActorId,
} from "./utils/linear-client";
import { createLogger } from "./logger";

const log = createLogger("issue-status");

const STATES_CACHE_TTL_SECONDS = 900; // 15 minutes
const BOT_ACTOR_CACHE_KEY = "cache:bot-actor-id";

export type IssueTransition = "inProgress" | "completed" | "cancelled";

/** Maps our transition names to Linear's WorkflowStateType values. */
const TRANSITION_TO_STATE_TYPE: Record<IssueTransition, string> = {
  inProgress: "started",
  completed: "completed",
  cancelled: "canceled", // Linear uses American spelling
};

/**
 * State types that should block an "inProgress" transition.
 * Per Linear docs: don't move backwards once work is started/completed/canceled.
 */
const STARTED_GUARD_TYPES = new Set(["started", "completed", "canceled"]);

/**
 * Resolve the target state for a transition using position-based type resolution.
 * Returns the state with the lowest `position` for the target type.
 */
export function resolveLowestPositionState(
  states: WorkflowState[],
  targetType: string
): WorkflowState | null {
  const matching = states.filter((s) => s.type === targetType);
  if (matching.length === 0) return null;

  return matching.reduce((lowest, s) => {
    // Handle cached states that may lack position (backward compat)
    const pos = s.position ?? Infinity;
    const lowestPos = lowest.position ?? Infinity;
    return pos < lowestPos ? s : lowest;
  });
}

/** Fetch team workflow states, using KV cache to avoid repeated API calls. */
async function getTeamStates(
  env: Env,
  client: LinearApiClient,
  teamId: string
): Promise<WorkflowState[]> {
  const cacheKey = `states:${teamId}`;

  try {
    const cached = await env.LINEAR_KV.get(cacheKey, "json");
    if (cached && Array.isArray(cached)) {
      return cached as WorkflowState[];
    }
  } catch {
    /* cache miss */
  }

  const states = await fetchTeamWorkflowStates(client, teamId);
  if (states.length > 0) {
    await env.LINEAR_KV.put(cacheKey, JSON.stringify(states), {
      expirationTtl: STATES_CACHE_TTL_SECONDS,
    });
  }
  return states;
}

/**
 * Transition a Linear issue's workflow state.
 * Uses position-based type resolution (lowest position for target type).
 * Guards "inProgress" transitions: skips if issue is already started/completed/canceled.
 * No-op if the feature is disabled or the state cannot be resolved.
 */
export async function transitionIssueStatus(
  env: Env,
  client: LinearApiClient,
  issueId: string,
  teamId: string,
  transition: IssueTransition,
  updateIssueStatusEnabled: boolean,
  currentStateType: string | null,
  traceId?: string
): Promise<void> {
  if (!updateIssueStatusEnabled) return;

  // Guard: don't move to "started" if already in a forward state
  if (
    transition === "inProgress" &&
    currentStateType &&
    STARTED_GUARD_TYPES.has(currentStateType)
  ) {
    log.debug("issue_status.skipped", {
      trace_id: traceId,
      issue_id: issueId,
      transition,
      reason: "already_in_forward_state",
      current_state_type: currentStateType,
    });
    return;
  }

  const states = await getTeamStates(env, client, teamId);
  if (states.length === 0) {
    log.warn("issue_status.no_states", {
      trace_id: traceId,
      issue_id: issueId,
      team_id: teamId,
    });
    return;
  }

  const targetType = TRANSITION_TO_STATE_TYPE[transition];
  const targetState = resolveLowestPositionState(states, targetType);
  if (!targetState) {
    log.warn("issue_status.state_not_found", {
      trace_id: traceId,
      issue_id: issueId,
      team_id: teamId,
      transition,
      target_type: targetType,
    });
    return;
  }

  const success = await updateIssueState(client, issueId, targetState.id);
  log.info("issue_status.transition", {
    trace_id: traceId,
    issue_id: issueId,
    team_id: teamId,
    transition,
    target_state: targetState.name,
    state_id: targetState.id,
    success,
  });
}

/**
 * Set the bot as issue delegate if not already set.
 * Caches the bot's actor ID in KV to avoid repeated viewer queries.
 */
export async function setDelegateIfUnset(
  env: Env,
  client: LinearApiClient,
  issueId: string,
  currentDelegateId: string | null,
  traceId?: string
): Promise<void> {
  if (currentDelegateId) {
    log.debug("issue_status.delegate_already_set", {
      trace_id: traceId,
      issue_id: issueId,
      delegate_id: currentDelegateId,
    });
    return;
  }

  let botActorId = await env.LINEAR_KV.get(BOT_ACTOR_CACHE_KEY);
  if (!botActorId) {
    botActorId = await fetchBotActorId(client);
    if (!botActorId) {
      log.warn("issue_status.bot_actor_id_not_found", { trace_id: traceId, issue_id: issueId });
      return;
    }
    await env.LINEAR_KV.put(BOT_ACTOR_CACHE_KEY, botActorId, {
      expirationTtl: 86400, // 24 hours
    });
  }

  const success = await setIssueDelegate(client, issueId, botActorId);
  log.info("issue_status.delegate_set", {
    trace_id: traceId,
    issue_id: issueId,
    bot_actor_id: botActorId,
    success,
  });
}
