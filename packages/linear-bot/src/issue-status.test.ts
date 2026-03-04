import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  resolveLowestPositionState,
  transitionIssueStatus,
  setDelegateIfUnset,
} from "./issue-status";
import type { WorkflowState } from "./utils/linear-client";
import type { Env } from "./types";

// ─── Fake KVNamespace ────────────────────────────────────────────────────────

function createFakeKV(initial: Record<string, string> = {}) {
  const store = new Map(Object.entries(initial));

  const kv = {
    async get(key: string, type?: string) {
      const val = store.get(key) ?? null;
      if (val === null) return null;
      if (type === "json") return JSON.parse(val);
      return val;
    },
    async put(key: string, value: string, _options?: { expirationTtl?: number }) {
      store.set(key, value);
    },
    async delete(key: string) {
      store.delete(key);
    },
  };

  return { kv: kv as unknown as KVNamespace, store };
}

// ─── Test data ───────────────────────────────────────────────────────────────

const STATES: WorkflowState[] = [
  { id: "state-1", name: "Backlog", type: "backlog", position: 0 },
  { id: "state-2", name: "Todo", type: "unstarted", position: 1 },
  { id: "state-3", name: "In Progress", type: "started", position: 2 },
  { id: "state-4", name: "In Review", type: "started", position: 3 },
  { id: "state-5", name: "Done", type: "completed", position: 0 },
  { id: "state-6", name: "Deployed", type: "completed", position: 1 },
  { id: "state-7", name: "Cancelled", type: "canceled", position: 0 },
];

// ─── resolveLowestPositionState ─────────────────────────────────────────────

describe("resolveLowestPositionState", () => {
  it("returns the state with lowest position for the target type", () => {
    const result = resolveLowestPositionState(STATES, "started");
    expect(result).toEqual({ id: "state-3", name: "In Progress", type: "started", position: 2 });
  });

  it("returns lowest position for completed type", () => {
    const result = resolveLowestPositionState(STATES, "completed");
    expect(result).toEqual({ id: "state-5", name: "Done", type: "completed", position: 0 });
  });

  it("returns null when no states match the type", () => {
    const statesWithoutCanceled = STATES.filter((s) => s.type !== "canceled");
    expect(resolveLowestPositionState(statesWithoutCanceled, "canceled")).toBeNull();
  });

  it("returns null for empty states array", () => {
    expect(resolveLowestPositionState([], "started")).toBeNull();
  });

  it("handles states without position (backward compat)", () => {
    const legacyStates = [
      { id: "s1", name: "A", type: "started" },
      { id: "s2", name: "B", type: "started", position: 1 },
    ] as WorkflowState[];
    // s1 has no position → treated as Infinity, s2 wins
    const result = resolveLowestPositionState(legacyStates, "started");
    expect(result?.id).toBe("s2");
  });

  it("picks first when all positions equal", () => {
    const samePos = [
      { id: "s1", name: "A", type: "started", position: 5 },
      { id: "s2", name: "B", type: "started", position: 5 },
    ];
    const result = resolveLowestPositionState(samePos, "started");
    expect(result?.id).toBe("s1");
  });
});

// ─── transitionIssueStatus ───────────────────────────────────────────────────

// Mock the linear-client module
vi.mock("./utils/linear-client", () => ({
  fetchTeamWorkflowStates: vi.fn(async () => STATES),
  updateIssueState: vi.fn(async () => true),
  setIssueDelegate: vi.fn(async () => true),
  fetchBotActorId: vi.fn(async () => "bot-actor-123"),
}));

import {
  fetchTeamWorkflowStates,
  updateIssueState,
  setIssueDelegate,
  fetchBotActorId,
} from "./utils/linear-client";

const mockedFetchStates = vi.mocked(fetchTeamWorkflowStates);
const mockedUpdateState = vi.mocked(updateIssueState);
const mockedSetDelegate = vi.mocked(setIssueDelegate);
const mockedFetchBotActorId = vi.mocked(fetchBotActorId);

describe("transitionIssueStatus", () => {
  const client = { accessToken: "test-token" };
  const issueId = "issue-1";
  const teamId = "team-1";

  function makeEnv(kvInitial: Record<string, string> = {}): Env {
    const { kv } = createFakeKV(kvInitial);
    return { LINEAR_KV: kv } as unknown as Env;
  }

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("is a no-op when updateIssueStatus is false", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", false, null);
    expect(mockedFetchStates).not.toHaveBeenCalled();
    expect(mockedUpdateState).not.toHaveBeenCalled();
  });

  it("calls updateIssueState with lowest-position state for inProgress", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", true, null);
    expect(mockedFetchStates).toHaveBeenCalledWith(client, teamId);
    // "In Progress" has position 2, lowest among "started" type
    expect(mockedUpdateState).toHaveBeenCalledWith(client, issueId, "state-3");
  });

  it("calls updateIssueState with lowest-position state for completed", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "completed", true, null);
    // "Done" has position 0, lowest among "completed" type
    expect(mockedUpdateState).toHaveBeenCalledWith(client, issueId, "state-5");
  });

  it("calls updateIssueState with lowest-position state for cancelled", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "cancelled", true, null);
    expect(mockedUpdateState).toHaveBeenCalledWith(client, issueId, "state-7");
  });

  it("uses cached states from KV on second call", async () => {
    const env = makeEnv();

    await transitionIssueStatus(env, client, issueId, teamId, "inProgress", true, null);
    expect(mockedFetchStates).toHaveBeenCalledTimes(1);

    // Second call should use cached states
    await transitionIssueStatus(env, client, issueId, teamId, "cancelled", true, null);
    expect(mockedFetchStates).toHaveBeenCalledTimes(1);
  });

  it("does not call updateIssueState when no states returned", async () => {
    mockedFetchStates.mockResolvedValueOnce([]);
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", true, null);
    expect(mockedUpdateState).not.toHaveBeenCalled();
  });

  it("does not call updateIssueState when target type not found", async () => {
    mockedFetchStates.mockResolvedValueOnce([
      { id: "state-1", name: "Backlog", type: "backlog", position: 0 },
    ]);
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", true, null);
    expect(mockedUpdateState).not.toHaveBeenCalled();
  });

  // ─── State guard tests ───────────────────────────────────────────────────

  it("skips inProgress when current state is already started", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", true, "started");
    expect(mockedFetchStates).not.toHaveBeenCalled();
    expect(mockedUpdateState).not.toHaveBeenCalled();
  });

  it("skips inProgress when current state is completed", async () => {
    await transitionIssueStatus(
      makeEnv(),
      client,
      issueId,
      teamId,
      "inProgress",
      true,
      "completed"
    );
    expect(mockedUpdateState).not.toHaveBeenCalled();
  });

  it("skips inProgress when current state is canceled", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", true, "canceled");
    expect(mockedUpdateState).not.toHaveBeenCalled();
  });

  it("allows inProgress when current state is unstarted", async () => {
    await transitionIssueStatus(
      makeEnv(),
      client,
      issueId,
      teamId,
      "inProgress",
      true,
      "unstarted"
    );
    expect(mockedUpdateState).toHaveBeenCalledWith(client, issueId, "state-3");
  });

  it("allows inProgress when currentStateType is null (unknown)", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "inProgress", true, null);
    expect(mockedUpdateState).toHaveBeenCalledWith(client, issueId, "state-3");
  });

  it("always allows cancelled regardless of current state", async () => {
    await transitionIssueStatus(makeEnv(), client, issueId, teamId, "cancelled", true, "started");
    expect(mockedUpdateState).toHaveBeenCalledWith(client, issueId, "state-7");
  });
});

// ─── setDelegateIfUnset ─────────────────────────────────────────────────────

describe("setDelegateIfUnset", () => {
  const client = { accessToken: "test-token" };
  const issueId = "issue-1";

  function makeEnv(kvInitial: Record<string, string> = {}): Env {
    const { kv } = createFakeKV(kvInitial);
    return { LINEAR_KV: kv } as unknown as Env;
  }

  beforeEach(() => {
    vi.clearAllMocks();
  });

  it("sets delegate when not already set", async () => {
    await setDelegateIfUnset(makeEnv(), client, issueId, null);
    expect(mockedFetchBotActorId).toHaveBeenCalledWith(client);
    expect(mockedSetDelegate).toHaveBeenCalledWith(client, issueId, "bot-actor-123");
  });

  it("skips when delegate is already set", async () => {
    await setDelegateIfUnset(makeEnv(), client, issueId, "existing-delegate");
    expect(mockedFetchBotActorId).not.toHaveBeenCalled();
    expect(mockedSetDelegate).not.toHaveBeenCalled();
  });

  it("uses cached bot actor ID from KV", async () => {
    const env = makeEnv({ "cache:bot-actor-id": "cached-bot-id" });
    await setDelegateIfUnset(env, client, issueId, null);
    expect(mockedFetchBotActorId).not.toHaveBeenCalled();
    expect(mockedSetDelegate).toHaveBeenCalledWith(client, issueId, "cached-bot-id");
  });

  it("caches bot actor ID after first fetch", async () => {
    const env = makeEnv();
    await setDelegateIfUnset(env, client, issueId, null);
    expect(mockedFetchBotActorId).toHaveBeenCalledTimes(1);

    // Reset delegate mock to check second call
    mockedSetDelegate.mockClear();
    mockedFetchBotActorId.mockClear();

    // Remove delegate so it tries again
    await setDelegateIfUnset(env, client, "issue-2", null);
    expect(mockedFetchBotActorId).not.toHaveBeenCalled();
    expect(mockedSetDelegate).toHaveBeenCalledWith(client, "issue-2", "bot-actor-123");
  });

  it("does nothing when bot actor ID cannot be fetched", async () => {
    mockedFetchBotActorId.mockResolvedValueOnce(null);
    await setDelegateIfUnset(makeEnv(), client, issueId, null);
    expect(mockedSetDelegate).not.toHaveBeenCalled();
  });
});
