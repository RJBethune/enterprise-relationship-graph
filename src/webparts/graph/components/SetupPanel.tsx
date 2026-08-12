import * as React from 'react';
import styles from './GraphApp.module.scss';
import { IProvisioningPlan, describeAction, planSummary, isHealthy } from '../../../provisioning/planner';
import { IProvisioningStepResult } from '../../../services/sp/SpProvisioningService';

export interface ISetupPanelProps {
  plan: IProvisioningPlan;
  canEdit: boolean;
  running: boolean;
  progress: { done: number; total: number; label: string } | null;
  results: IProvisioningStepResult[] | null;
  onDeploy(): void;
  onRecheck(): void;
  onContinue(): void;
}

/**
 * The Deploy / Update Schema surface.
 *
 * It always shows the plan BEFORE running it. A button that silently changes a site's
 * schema is one nobody can be blamed for refusing to press; showing exactly what will
 * be created, and what the tool refuses to touch, is what makes it approvable by
 * someone who did not write it.
 */
export const SetupPanel: React.FunctionComponent<ISetupPanelProps> = (props) => {
  const { plan, canEdit, running, progress, results } = props;
  const healthy = isHealthy(plan);
  const pct = progress && progress.total > 0 ? Math.round((progress.done / progress.total) * 100) : 0;

  return (
    <div className={styles.panel}>
      <h2 className={styles.panelTitle}>
        {healthy ? 'SharePoint backend is ready' : 'Set up the SharePoint backend'}
      </h2>
      <p className={styles.panelText}>{planSummary(plan)}</p>

      {plan.conflicts.length > 0 && (
        <div className={styles.conflict}>
          <strong>Needs a person.</strong> These are deliberately not automated — the tool
          only ever adds, never retypes or deletes:
          <ul>
            {plan.conflicts.map((c, i) => (
              <li key={i}>{c.internal ? `${c.list}.${c.internal}: ` : `${c.list}: `}{c.reason}</li>
            ))}
          </ul>
        </div>
      )}

      {running && progress && (
        <>
          <div className={styles.progress}>
            <div className={styles.progressFill} style={{ width: `${pct}%` }} />
          </div>
          <p className={styles.panelText} aria-live="polite">
            {progress.done} of {progress.total} — {progress.label}
          </p>
        </>
      )}

      {!running && results && (
        <ul className={styles.actionList}>
          {results.map((r, i) => (
            <li key={i}>
              <span className={r.ok ? styles.stepOk : styles.stepFail} aria-hidden="true">
                {r.ok ? '✓' : '✕'}
              </span>
              <span>
                {r.label}
                {r.error ? ` — ${r.error}` : ''}
              </span>
            </li>
          ))}
        </ul>
      )}

      {!running && !results && plan.actions.length > 0 && (
        <ul className={styles.actionList}>
          {plan.actions.map((a, i) => <li key={i}>{describeAction(a)}</li>)}
        </ul>
      )}

      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        {!healthy && (
          <button
            type="button"
            className={`${styles.button} ${styles.primary}`}
            onClick={props.onDeploy}
            disabled={running || !canEdit}
            title={canEdit ? undefined : 'You need permission to manage lists on this site.'}
          >
            {results ? 'Run remaining steps' : 'Deploy list / update schema'}
          </button>
        )}
        <button type="button" className={styles.button} onClick={props.onRecheck} disabled={running}>
          Re-check
        </button>
        {healthy && (
          <button type="button" className={`${styles.button} ${styles.primary}`} onClick={props.onContinue}>
            Open the graph
          </button>
        )}
      </div>

      {!canEdit && !healthy && (
        <p className={styles.panelText} style={{ marginTop: 14 }}>
          Setup needs permission to manage lists on this site. Ask a site owner to open this
          page once and press the button — it only ever adds lists and columns.
        </p>
      )}
    </div>
  );
};
