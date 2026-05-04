/**
 * App.tsx — Root component.
 * Checks Supabase session on mount.
 * Shows LoginPrompt if not authenticated, CreateTaskForm if authenticated.
 */

import React, { useEffect, useState } from 'react';
import { supabase } from '../lib/supabase';
import LoginPrompt from './LoginPrompt';
import CreateTaskForm from './CreateTaskForm';

type AuthState = 'loading' | 'unauthenticated' | 'authenticated';

const App: React.FC = () => {
  const [authState, setAuthState] = useState<AuthState>('loading');
  const [userId, setUserId]       = useState<string | null>(null);
  const [userEmail, setUserEmail] = useState<string>('');

  const checkSession = async () => {
    const { data: { session } } = await supabase.auth.getSession();
    if (session?.user) {
      setUserId(session.user.id);
      setUserEmail(session.user.email || '');
      setAuthState('authenticated');
    } else {
      setAuthState('unauthenticated');
    }
  };

  useEffect(() => {
    checkSession();
  }, []);

  const handleSignOut = async () => {
    await supabase.auth.signOut();
    setUserId(null);
    setUserEmail('');
    setAuthState('unauthenticated');
  };

  // ── Loading ──────────────────────────────────────────────────────────────
  if (authState === 'loading') {
    return (
      <div style={pageStyle}>
        <div style={logoStyle}>⚡ EasyTask</div>
        <p style={{ color: '#94a3b8', fontSize: '13px' }}>Loading…</p>
      </div>
    );
  }

  // ── Not authenticated ────────────────────────────────────────────────────
  if (authState === 'unauthenticated') {
    return <LoginPrompt onAuthenticated={checkSession} />;
  }

  // ── Authenticated ────────────────────────────────────────────────────────
  return (
    <CreateTaskForm
      userId={userId!}
      userEmail={userEmail}
      onSignOut={handleSignOut}
    />
  );
};

const pageStyle: React.CSSProperties = {
  minHeight: '100vh',
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'center',
  justifyContent: 'center',
  fontFamily: 'system-ui, -apple-system, sans-serif',
  background: '#fff',
};

const logoStyle: React.CSSProperties = {
  fontSize: '20px',
  fontWeight: 900,
  color: '#f97316',
  marginBottom: '12px',
};

export default App;
