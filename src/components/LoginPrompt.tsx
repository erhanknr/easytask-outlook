/**
 * LoginPrompt.tsx
 *
 * Shown when the user hasn't connected their EasyTask account yet.
 * Clicking "Connect Account" opens dialog.html as an Office.js popup.
 * When the user logs in, the dialog sends the Supabase session back here
 * via Office.context.ui.messageParent, and we apply it with supabase.auth.setSession().
 */

import React, { useState } from 'react';
import { supabase } from '../lib/supabase';

interface Props {
  onAuthenticated: () => void;
}

const ORANGE = '#f97316';
const BASE_URL = import.meta.env.VITE_ADDIN_BASE_URL as string;

const LoginPrompt: React.FC<Props> = ({ onAuthenticated }) => {
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState('');

  const openLoginDialog = () => {
    setLoading(true);
    setError('');

    const dialogUrl = `${BASE_URL}/dialog.html`;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 65, width: 35, displayInIframe: false },
      (asyncResult: Office.AsyncResult<Office.Dialog>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setError('Could not open login window. Please try again.');
          setLoading(false);
          return;
        }

        const dialog = asyncResult.value;

        // Listen for messages sent from dialog.tsx via messageParent()
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          async (args: any) => {
            dialog.close();

            try {
              const msg = JSON.parse(args.message);

              if (msg.type === 'SESSION' && msg.session) {
                // Apply the session to the task pane's Supabase client
                await supabase.auth.setSession(msg.session);
                onAuthenticated();
              } else {
                setError('Login was not completed. Please try again.');
              }
            } catch {
              setError('Something went wrong. Please try again.');
            }
            setLoading(false);
          }
        );

        // Handle dialog closed without completing login
        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          () => {
            setLoading(false);
          }
        );
      }
    );
  };

  return (
    <div style={styles.page}>
      {/* Logo */}
      <div style={styles.logoWrap}>
        <span style={styles.logo}>⚡</span>
        <span style={styles.logoText}>EasyTask</span>
      </div>

      <h2 style={styles.heading}>Turn emails into tasks</h2>
      <p style={styles.body}>
        Connect your EasyTask account to instantly create tasks from any email — with project, section, assignees, and due date.
      </p>

      {error && <div style={styles.error}>{error}</div>}

      <button
        onClick={openLoginDialog}
        disabled={loading}
        style={{ ...styles.btn, ...(loading ? styles.btnDisabled : {}) }}
      >
        {loading ? 'Opening sign-in…' : 'Connect EasyTask Account'}
      </button>

      <p style={styles.hint}>
        Use the same credentials as{' '}
        <span style={{ color: ORANGE }}>geteasytask.com</span>
      </p>
    </div>
  );
};

const styles: Record<string, React.CSSProperties> = {
  page: {
    padding: '32px 20px',
    fontFamily: 'system-ui, -apple-system, sans-serif',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    textAlign: 'center',
  },
  logoWrap: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    marginBottom: '20px',
  },
  logo: {
    fontSize: '28px',
  },
  logoText: {
    fontSize: '22px',
    fontWeight: 900,
    color: ORANGE,
    letterSpacing: '-0.5px',
  },
  heading: {
    fontSize: '17px',
    fontWeight: 800,
    color: '#0f172a',
    margin: '0 0 10px',
  },
  body: {
    fontSize: '13px',
    color: '#64748b',
    lineHeight: 1.6,
    margin: '0 0 24px',
    maxWidth: '260px',
  },
  error: {
    background: '#fff1f2',
    border: '1px solid #fecdd3',
    borderRadius: '8px',
    padding: '10px 12px',
    fontSize: '12px',
    color: '#e11d48',
    marginBottom: '14px',
    width: '100%',
    maxWidth: '260px',
    boxSizing: 'border-box' as const,
  },
  btn: {
    padding: '12px 20px',
    background: ORANGE,
    color: '#fff',
    border: 'none',
    borderRadius: '12px',
    fontSize: '14px',
    fontWeight: 700,
    cursor: 'pointer',
    width: '100%',
    maxWidth: '260px',
    boxShadow: '0 4px 14px rgba(249,115,22,0.3)',
  },
  btnDisabled: {
    opacity: 0.6,
    cursor: 'not-allowed',
  },
  hint: {
    fontSize: '11px',
    color: '#94a3b8',
    marginTop: '14px',
  },
};

export default LoginPrompt;
