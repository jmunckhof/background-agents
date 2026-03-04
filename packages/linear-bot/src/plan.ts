/**
 * Agent plan step types and factory used by both the webhook handler and callbacks.
 */

export type PlanStepStatus = "pending" | "inProgress" | "completed" | "canceled";

export interface PlanStep {
  content: string;
  status: PlanStepStatus;
}

export function makePlan(
  stage: "start" | "repo_resolved" | "session_created" | "completed" | "failed"
): PlanStep[] {
  const steps = [
    "Analyze issue",
    "Resolve repository",
    "Create coding session",
    "Code changes",
    "Open PR",
  ];
  const statusMap: Record<string, PlanStepStatus[]> = {
    start: ["inProgress", "inProgress", "pending", "pending", "pending"],
    repo_resolved: ["completed", "completed", "inProgress", "pending", "pending"],
    session_created: ["completed", "completed", "completed", "inProgress", "pending"],
    completed: ["completed", "completed", "completed", "completed", "completed"],
    failed: ["completed", "completed", "completed", "completed", "canceled"],
  };
  const statuses = statusMap[stage];
  return steps.map((content, i) => ({ content, status: statuses[i] }));
}

/** Map TodoWrite status values to Linear PlanStepStatus. */
const TODO_STATUS_MAP: Record<string, PlanStepStatus> = {
  pending: "pending",
  in_progress: "inProgress",
  completed: "completed",
};

/**
 * Convert a TodoWrite tool call's args into Linear plan steps.
 * Returns null if the args don't contain a valid todos array.
 */
export function todosToplan(args: Record<string, unknown>): PlanStep[] | null {
  const todos = args.todos;
  if (!Array.isArray(todos) || todos.length === 0) return null;

  return todos.map((todo: { content?: string; status?: string }) => ({
    content: String(todo.content ?? ""),
    status: TODO_STATUS_MAP[String(todo.status ?? "pending")] ?? "pending",
  }));
}
