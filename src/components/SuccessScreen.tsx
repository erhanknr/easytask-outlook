/**
 * SuccessScreen.tsx
 *
 * Shown after a task is successfully created from an email.
 * Displays the task title, project, and section, with options to
 * create another task or close the pane.
 */

import React from 'react';

interface Props {
  taskTitle:    string;
  projectName:  string;
  sectionName:  string;
  onCreateAnother: () => void;
}

const ORANGE = '#f97316';

const SuccessScreen: React.FC<Props> = ({
  taskTitle,
  projectName,
  sectionName,
  onCreateAnother,
}) => {
  return (
    <div style={s.page}>

      {/* Header */}
      <div style={s.header}>
        <div style={s.logo}>⚡ EasyTask</div>
      </div>

      {/* Success card */}
      <div style={s.card}>

        {/* Check icon */}
        <div style={s.iconWrap}>
          <div style={s.iconCircle}>
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="#fff" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="20 6 9 17 4 12" />
            </svg>
          </div>
        </div>

        <h2 style={s.heading}>Task Created!</h2>
        <p style={s.sub}>Your email has been turned into a task.</p>

        {/* Task summary */}
        <div style={s.summary}>
          <div style={s.summaryRow}>
            <span style={s.summaryIcon}>📋</span>
            <span style={s.summaryText} title={taskTitle}>
              {taskTitle.length > 50 ? taskTitle.slice(0, 50) + '…' : taskTitle}
            </span>
          </div>

          {projectName && (
            <div style={s.summaryRow}>
              <span style={s.summaryIcon}>📁</span>
              <span style={s.summaryText}>{projectName}</span>
            </div>
          )}

          {sectionName && (
            <div style={s.summaryRow}>
              <span style={s.summaryIcon}>📌</span>
              <span style={s.summaryText}>{sectionName}</span>
            </div>
          )}
        </div>

        {/* CTA */}
        <button onClick={onCreateAnother} style={s.btn}>
          ⚡ Create Another Task
        </button>

        <p style={s.hint}>
          View it at{' '}
          <a href="https://app.geteasytask.com" target="_blank" rel="noreferrer" style={s.link}>
            app.geteasytask.com
          </a>
        </p>
      </div>
    </div>
  );
};

// ─── Styles ───────────────────────────────────────────────────────────────────

const s: Record<string, React.CSSProperties> = {
  page: {
    fontFamily: 'system-ui, -apple-system, sans-serif',
    background: '#fff',
    minHeight: '100vh',
    padding: '0 0 32px',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '14px 16px 10px',
    borderBottom: '1px solid #e2e8f0',
    marginBottom: '24px',
  },
  logo: {
    fontSize: '16px',
    fontWeight: 900,
    color: ORANGE,
  },
  card: {
    margin: '0 16px',
    padding: '28px 20px 24px',
    borderRadius: '16px',
    border: '1px solid #e2e8f0',
    background: '#fff',
    boxShadow: '0 4px 24px rgba(0,0,0,0.05)',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
  },
  iconWrap: {
    marginBottom: '16px',
  },
  iconCircle: {
    width: '56px',
    height: '56px',
    borderRadius: '50%',
    background: 'linear-gradient(135deg, #22c55e, #16a34a)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    boxShadow: '0 4px 16px rgba(34,197,94,0.35)',
  },
  heading: {
    fontSize: '20px',
    fontWeight: 800,
    color: '#0f172a',
    margin: '0 0 6px',
  },
  sub: {
    fontSize: '13px',
    color: '#64748b',
    margin: '0 0 20px',
  },
  summary: {
    width: '100%',
    background: '#f8fafc',
    borderRadius: '10px',
    border: '1px solid #e2e8f0',
    padding: '12px 14px',
    marginBottom: '20px',
    display: 'flex',
    flexDirection: 'column',
    gap: '8px',
  },
  summaryRow: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    textAlign: 'left',
  },
  summaryIcon: {
    fontSize: '14px',
    flexShrink: 0,
  },
  summaryText: {
    fontSize: '12px',
    color: '#374151',
    fontWeight: 500,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap' as const,
  },
  btn: {
    width: '100%',
    padding: '12px',
    background: ORANGE,
    color: '#fff',
    border: 'none',
    borderRadius: '10px',
    fontSize: '14px',
    fontWeight: 700,
    cursor: 'pointer',
    boxShadow: '0 4px 14px rgba(249,115,22,0.3)',
    marginBottom: '14px',
  },
  hint: {
    fontSize: '11px',
    color: '#94a3b8',
    margin: 0,
  },
  link: {
    color: ORANGE,
    textDecoration: 'none',
    fontWeight: 600,
  },
};

export default SuccessScreen;
