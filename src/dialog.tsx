/**
 * dialog.tsx — EasyTask Login Dialog
 *
 * This page is opened as a popup by the task pane via:
 *   Office.context.ui.displayDialogAsync(dialogUrl, ...)
 *
 * It shows a Supabase email/password login form.
 * On success it sends the session back to the task pane via
 *   Office.context.ui.messageParent(JSON.stringify({ type: 'SESSION', session }))
 * then closes itself.
 */

import React, { useState, useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import { supabase } from './lib/supabase';

const ORANGE = '#f97316';

const styles: Record<string, React.CSSProperties> = {
  page: {
    minHeight: '100vh',
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    justifyContent: 'center',
    background: '#fff',
    padding: '24px',
    fontFamily: 'system-ui, -apple-system, sans-serif',
  },
  logo: {
    fontSize: '22px',
    fontWeight: 900,
    color: ORANGE,
    marginBottom: '4px',
    letterSpacing: '-0.5px',
  },
  subtitle: {
    fontSize: '13px',
    color: '#64748b',
    marginBottom: '28px',
  },
  card: {
    width: '100%',
    maxWidth: '320px',
    background: '#fff',
    borderRadius: '16px',
    border: '1px solid #e2e8f0',
    padding: '24px',
    boxShadow: '0 4px 24px rgba(0,0,0,0.06)',
  },
  label: {
    display: 'block',
    fontSize: '12px',
    fontWeight: 600,
    color: '#374151',
    marginBottom: '6px',
  },
  input: {
    width: '100%',
    padding: '10px 12px',
    borderRadius: '10px',
    border: '1.5px solid #e2e8f0',
    fontSize: '14px',
    outline: 'none',
    boxSizing: 'border-box',
    marginBottom: '14px',
    transition: 'border-color 0.15s',
  },
  btn: {
    width: '100%',
    padding: '11px',
    background: ORANGE,
    color: '#fff',
    border: 'none',
    borderRadius: '10px',
    fontSize: '14px',
    fontWeight: 700,
    cursor: 'pointer',
    marginTop: '4px',
  },
  btnDisabled: {
    opacity: 0.6,
    cursor: 'not-allowed',
  },
  error: {
    background: '#fff1f2',
    border: '1px solid #fecdd3',
    borderRadius: '8px',
    padding: '10px 12px',
    fontSize: '12px',
    color: '#e11d48',
    marginBottom: '14px',
  },
  success: {
    textAlign: 'center' as const,
    color: '#059669',
    fontSize: '14px',
    fontWeight: 600,
  },
};

const DialogApp: React.FC = () => {
  const [email, setEmail]       = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading]   = useState(false);
  const [error, setError]       = useState('');
  const [done, setDone]         = useState(false);
  const [officeReady, setOfficeReady] = useState(false);

  useEffect(() => {
    // Wait for Office.js to initialise before allowing login
    Office.onReady(() => setOfficeReady(true));
  }, []);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!officeReady) return;
    setLoading(true);
    setError('');

    const { data, error: authError } = await supabase.auth.signInWithPassword({ email, password });

    if (authError || !data.session) {
      setError(authError?.message || 'Login failed. Please check your credentials.');
      setLoading(false);
      return;
    }

    setDone(true);

    // Send session back to the task pane, then close the dialog
    Office.context.ui.messageParent(
      JSON.stringify({ type: 'SESSION', session: data.session })
    );
  };

  if (done) {
    return (
      <div style={styles.page}>
        <p style={styles.success}>✅ Signed in! Closing…</p>
      </div>
    );
  }

  return (
    <div style={styles.page}>
      <div style={styles.logo}>⚡ EasyTask</div>
      <div style={styles.subtitle}>Sign in to connect your account</div>

      <div style={styles.card}>
        <form onSubmit={handleLogin}>
          {error && <div style={styles.error}>{error}</div>}

          <label style={styles.label}>Email</label>
          <input
            style={styles.input}
            type="email"
            value={email}
            onChange={e => setEmail(e.target.value)}
            placeholder="you@example.com"
            required
            autoFocus
          />

          <label style={styles.label}>Password</label>
          <input
            style={styles.input}
            type="password"
            value={password}
            onChange={e => setPassword(e.target.value)}
            placeholder="••••••••"
            required
          />

          <button
            style={{ ...styles.btn, ...(loading ? styles.btnDisabled : {}) }}
            type="submit"
            disabled={loading || !officeReady}
          >
            {loading ? 'Signing in…' : 'Sign In to EasyTask'}
          </button>
        </form>
      </div>

      <p style={{ fontSize: '11px', color: '#94a3b8', marginTop: '16px', textAlign: 'center' }}>
        Use the same credentials as geteasytask.com
      </p>
    </div>
  );
};

// Wait for Office.js then mount React
Office.onReady(() => {
  const root = document.getElementById('dialog-root')!;
  createRoot(root).render(<DialogApp />);
});
