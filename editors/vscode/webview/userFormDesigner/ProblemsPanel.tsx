import { useRef, useState } from "preact/hooks";
import type { DesignerProblem, DesignerStrings } from "../../src/userFormEditor/protocol";

interface ProblemsPanelProps {
  problems: DesignerProblem[];
  version: number;
  strings: DesignerStrings;
  onOpenProblem: (version: number, index: number) => void;
}

type Severity = "error" | "warning" | "information" | "info" | "hint";

function severityOf(problem: DesignerProblem): Severity {
  const severity = (problem.severity ?? "warning").toLowerCase();
  return ["error", "warning", "information", "info", "hint"].includes(severity)
    ? (severity as Severity)
    : "warning";
}

function severityRank(severity: Severity): number {
  switch (severity) {
    case "error":
      return 4;
    case "warning":
      return 3;
    case "information":
    case "info":
      return 2;
    case "hint":
      return 1;
  }
}

function severityLabel(severity: Severity, strings: DesignerStrings): string {
  switch (severity) {
    case "error":
      return strings.problemSeverityError;
    case "information":
    case "info":
      return strings.problemSeverityInformation;
    case "hint":
      return strings.problemSeverityHint;
    default:
      return strings.problemSeverityWarning;
  }
}

function severityIcon(severity: Severity): string {
  switch (severity) {
    case "error":
      return "error";
    case "information":
    case "info":
      return "info";
    case "hint":
      return "lightbulb";
    default:
      return "warning";
  }
}

export function ProblemsPanel({ problems, version, strings, onOpenProblem }: ProblemsPanelProps) {
  const [open, setOpen] = useState(false);
  const toggle = useRef<HTMLButtonElement>(null);
  const close = () => {
    setOpen(false);
    toggle.current?.focus({ preventScroll: true });
  };
  const mostSevere = problems.reduce<DesignerProblem | undefined>(
    (current, problem) =>
      !current || severityRank(severityOf(problem)) > severityRank(severityOf(current))
        ? problem
        : current,
    undefined,
  );
  const topSeverity = mostSevere ? severityOf(mostSevere) : undefined;
  const topSeverityLabel = topSeverity ? severityLabel(topSeverity, strings) : strings.noProblems;

  return (
    <div
      class="designer-problems"
      onKeyDown={(event) => {
        if (!open || event.key !== "Escape" || event.defaultPrevented) return;
        event.preventDefault();
        event.stopPropagation();
        close();
      }}
    >
      <section
        id="designer-problems-panel"
        class="designer-problems-panel"
        aria-labelledby="designer-problems-heading"
        role="region"
        hidden={!open}
      >
        <div class="designer-problems-heading">
          <h2 id="designer-problems-heading">{strings.problems}</h2>
          <button type="button" aria-label={strings.closeProblems} onClick={close}>
            {strings.closeProblems}
          </button>
        </div>
        {problems.length === 0 ? (
          <p class="designer-problems-empty">{strings.noProblems}</p>
        ) : (
          <ul class="designer-problems-list">
            {problems.map((problem, index) => {
              const severity = severityOf(problem);
              const label = severityLabel(severity, strings);
              return (
                <li key={`${problem.code ?? "problem"}-${index}`}>
                  <button
                    type="button"
                    class={`designer-problem designer-problem-${severity}`}
                    aria-label={[label, problem.code, problem.field, problem.message]
                      .filter(Boolean)
                      .join(": ")}
                    onClick={() => onOpenProblem(version, index)}
                  >
                    <span class="designer-problem-severity">{label}</span>
                    {problem.code && <code>{problem.code}</code>}
                    {problem.field && <code class="designer-problem-field">{problem.field}</code>}
                    <span class="designer-problem-message">{problem.message}</span>
                  </button>
                </li>
              );
            })}
          </ul>
        )}
      </section>
      <button
        ref={toggle}
        id="designer-problems-toggle"
        class="designer-problems-toggle"
        type="button"
        aria-expanded={open}
        aria-controls="designer-problems-panel"
        aria-label={`${strings.problems}: ${problems.length}, ${topSeverityLabel}`}
        onClick={() => setOpen((current) => !current)}
      >
        {topSeverity && (
          <span
            class={`codicon codicon-${severityIcon(topSeverity)} designer-problems-severity-icon designer-problems-severity-${topSeverity}`}
            aria-hidden="true"
          />
        )}
        {strings.problems} {problems.length}
      </button>
    </div>
  );
}
