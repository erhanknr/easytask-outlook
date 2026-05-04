/**
 * CreateTaskForm.tsx
 *
 * The main task creation form. On mount it:
 *  1. Reads the selected email from Office.context.mailbox.item
 *  2. Pre-fills Title (subject) and Description (body snippet)
 *  3. Loads the user's workspaces → projects → sections → members
 *
 * On submit it calls createTask() via useEasyTaskData and shows SuccessScreen.
 */

import React, { useState, useEffect } from 'react';
import { useEasyTaskData, Workspace, Project, Section, Member } from '../hooks/useEasyTaskData';
import SuccessScreen from './SuccessScreen';

interface Props {
  userId: string;
  userEmail: string;
  onSignOut: () => void;
}

const ORANGE = '#f97316';
const BORDER = '#e2e8f0';

const CreateTaskForm: React.FC<Props> = ({ userId, userEmail, onSignOut }) => {
  // ── Email data (read from Outlook) ───────────────────────────────────────
  const [emailSubject, setEmailSubject] = useState('');
  const [emailSender, setEmailSender]   = useState('');
  const [emailBody, setEmailBody]       = useState('');

  // ── Form state ───────────────────────────────────────────────────────────
  const [taskTitle, setTaskTitle]           = useState('');
  const [taskDescription, setTaskDescription] = useState('');
  const [dueDate, setDueDate]               = useState('');
  const [selectedWorkspace, setSelectedWorkspace] = useState('');
  const [selectedProject, setSelectedProject]     = useState('');
  const [selectedSection, setSelectedSection]     = useState('');
  const [selectedAssignees, setSelectedAssignees] = useState<string[]>([]);

  // ── UI state ─────────────────────────────────────────────────────────────
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState('');
  const [success, setSuccess]       = useState(false);
  const [createdTaskId, setCreatedTaskId] = useState('');

  const {
    workspaces, projects, sections, members, loading,
    fetchProjects, fetchSections, fetchMembers, createTask,
  } = useEasyTaskData(userId);

  // ── Read email from Outlook on mount ─────────────────────────────────────
  useEffect(() => {
    const item = Office.context.mailbox.item;
    if (!item) return;

    // Subject (synchronous in read mode)
    const subject = (item as any).subject || '';
    setEmailSubject(subject);
    setTaskTitle(subject);

    // Sender
    const from = (item as any).from;
    const sender = from
      ? `${from.displayName || ''} <${from.emailAddress || ''}>`.trim()
      : '';
    setEmailSender(sender);

    // Body (async — plain text, capped at 500 chars for description pre-fill)
    (item as any).body?.getAsync('text', (result: any) => {
      if (result.status === 'succeeded') {
        const snippet = (result.value || '').trim().slice(0, 400);
        setEmailBody(snippet);
        setTaskDescription(snippet);
      }
    });
  }, []);

  // ── Auto-select workspace when only one ──────────────────────────────────
  useEffect(() => {
    if (workspaces.length === 1 && !selectedWorkspace) {
      const ws = workspaces[0];
      setSelectedWorkspace(ws.id);
      fetchProjects(ws.id);
      fetchMembers(ws.id);
    }
  }, [workspaces, selectedWorkspace, fetchProjects, fetchMembers]);

  // ── Handlers ─────────────────────────────────────────────────────────────
  const handleWorkspaceChange = (id: string) => {
    setSelectedWorkspace(id);
    setSelectedProject('');
    setSelectedSection('');
    setSelectedAssignees([]);
    fetchProjects(id);
    fetchMembers(id);
  };

  const handleProjectChange = (id: string) => {
    setSelectedProject(id);
    setSelectedSection('');
    fetchSections(id);
  };

  const toggleAssignee = (id: string) => {
    setSelectedAssignees(prev =>
      prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]
    );
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedSection) { setSubmitError('Please select a section.'); return; }
    setSubmitting(true);
    setSubmitError('');

    const result = await createTask({
      title:        taskTitle,
      description:  taskDescription,
      sectionId:    selectedSection,
      projectId:    selectedProject,
      assigneeIds:  selectedAssignees,
      dueDate:      dueDate || null,
      createdBy:    userId,
      emailSubject,
      emailSender,
    });

    if (result.success) {
      setCreatedTaskId(result.taskId || '');
      setSuccess(true);
    } else {
      setSubmitError('Failed to create task. Please try again.');
    }
    setSubmitting(false);
  };

  const handleCreateAnother = () => {
    setSuccess(false);
    setCreatedTaskId('');
    setSubmitError('');
    setSelectedSection('');
    setSelectedAssignees([]);
    setDueDate('');
    setTaskTitle(emailSubject);
    setTaskDescription(emailBody);
  };

  // ── Success screen ───────────────────────────────────────────────────────
  if (success) {
    const project = projects.find(p => p.id === selectedProject);
    const section = sections.find(s => s.id === selectedSection);
    return (
      <SuccessScreen
        taskTitle={taskTitle}
        projectName={project?.title || ''}
        sectionName={section?.name || ''}
        onCreateAnother={handleCreateAnother}
      />
    );
  }

  // ── Form ─────────────────────────────────────────────────────────────────
  return (
    <div style={s.page}>

      {/* Header */}
      <div style={s.header}>
        <div style={s.logo}>⚡ EasyTask</div>
        <button onClick={onSignOut} style={s.signOut} title="Sign out">
          Sign out
        </button>
      </div>

      {/* Email source pill */}
      {emailSubject && (
        <div style={s.emailPill}>
          <span style={s.pillIcon}>📧</span>
          <span style={s.pillText} title={emailSubject}>
            {emailSubject.length > 45 ? emailSubject.slice(0, 45) + '…' : emailSubject}
          </span>
        </div>
      )}

      <form onSubmit={handleSubmit} style={s.form}>

        {/* Task Title */}
        <div style={s.field}>
          <label style={s.label}>Task Title *</label>
          <input
            style={s.input}
            value={taskTitle}
            onChange={e => setTaskTitle(e.target.value)}
            placeholder="Task title"
            required
          />
        </div>

        {/* Workspace (hidden if only one) */}
        {workspaces.length > 1 && (
          <div style={s.field}>
            <label style={s.label}>Workspace *</label>
            <select
              style={s.select}
              value={selectedWorkspace}
              onChange={e => handleWorkspaceChange(e.target.value)}
              required
            >
              <option value="">Select workspace…</option>
              {workspaces.map((ws: Workspace) => (
                <option key={ws.id} value={ws.id}>{ws.name}</option>
              ))}
            </select>
          </div>
        )}

        {/* Project */}
        <div style={s.field}>
          <label style={s.label}>Project *</label>
          <select
            style={s.select}
            value={selectedProject}
            onChange={e => handleProjectChange(e.target.value)}
            required
            disabled={!selectedWorkspace || loading}
          >
            <option value="">Select project…</option>
            {projects.map((p: Project) => (
              <option key={p.id} value={p.id}>{p.code} — {p.title}</option>
            ))}
          </select>
        </div>

        {/* Section */}
        <div style={s.field}>
          <label style={s.label}>Section *</label>
          <select
            style={s.select}
            value={selectedSection}
            onChange={e => setSelectedSection(e.target.value)}
            required
            disabled={!selectedProject || sections.length === 0}
          >
            <option value="">Select section…</option>
            {sections.map((sec: Section) => (
              <option key={sec.id} value={sec.id}>{sec.name}</option>
            ))}
          </select>
        </div>

        {/* Due Date */}
        <div style={s.field}>
          <label style={s.label}>Due Date</label>
          <input
            style={s.input}
            type="date"
            value={dueDate}
            onChange={e => setDueDate(e.target.value)}
            min={new Date().toISOString().split('T')[0]}
          />
        </div>

        {/* Assignees */}
        {members.length > 0 && (
          <div style={s.field}>
            <label style={s.label}>Assign to</label>
            <div style={s.assigneeList}>
              {members.map((m: Member) => {
                const isSelected = selectedAssignees.includes(m.id);
                const initials = m.name.split(' ').map((n: string) => n[0]).join('').slice(0, 2).toUpperCase();
                return (
                  <button
                    key={m.id}
                    type="button"
                    onClick={() => toggleAssignee(m.id)}
                    style={{
                      ...s.assigneeBtn,
                      background: isSelected ? '#fff7ed' : '#f8fafc',
                      border: isSelected ? `1.5px solid ${ORANGE}` : '1.5px solid #e2e8f0',
                      color: isSelected ? ORANGE : '#374151',
                    }}
                  >
                    <span style={{ ...s.avatar, background: isSelected ? ORANGE : '#cbd5e1' }}>
                      {initials}
                    </span>
                    <span style={s.memberName}>{m.name}</span>
                    {isSelected && <span style={s.tick}>✓</span>}
                  </button>
                );
              })}
            </div>
          </div>
        )}

        {/* Description */}
        <div style={s.field}>
          <label style={s.label}>Description</label>
          <textarea
            style={s.textarea}
            value={taskDescription}
            onChange={e => setTaskDescription(e.target.value)}
            placeholder="Add more details…"
            rows={4}
            maxLength={800}
          />
          <div style={s.charCount}>{taskDescription.length} / 800</div>
        </div>

        {submitError && <div style={s.error}>{submitError}</div>}

        <button
          type="submit"
          disabled={submitting || !selectedSection}
          style={{
            ...s.submitBtn,
            ...(!selectedSection || submitting ? s.submitBtnDisabled : {}),
          }}
        >
          {submitting ? 'Creating task…' : '⚡ Create Task'}
        </button>

        <p style={s.footerNote}>Signed in as {userEmail}</p>
      </form>
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
    borderBottom: `1px solid ${BORDER}`,
    marginBottom: '12px',
  },
  logo: {
    fontSize: '16px',
    fontWeight: 900,
    color: ORANGE,
  },
  signOut: {
    background: 'none',
    border: 'none',
    fontSize: '11px',
    color: '#94a3b8',
    cursor: 'pointer',
    padding: '2px 6px',
    borderRadius: '6px',
  },
  emailPill: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    margin: '0 16px 12px',
    padding: '7px 10px',
    background: '#f0f9ff',
    borderRadius: '8px',
    border: '1px solid #bae6fd',
  },
  pillIcon: { fontSize: '13px' },
  pillText: {
    fontSize: '11px',
    color: '#0369a1',
    fontWeight: 500,
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap' as const,
  },
  form: {
    padding: '0 16px',
  },
  field: {
    marginBottom: '14px',
  },
  label: {
    display: 'block',
    fontSize: '11px',
    fontWeight: 700,
    color: '#374151',
    textTransform: 'uppercase' as const,
    letterSpacing: '0.5px',
    marginBottom: '5px',
  },
  input: {
    width: '100%',
    padding: '9px 11px',
    borderRadius: '8px',
    border: `1.5px solid ${BORDER}`,
    fontSize: '13px',
    outline: 'none',
    boxSizing: 'border-box' as const,
    background: '#f8fafc',
  },
  select: {
    width: '100%',
    padding: '9px 11px',
    borderRadius: '8px',
    border: `1.5px solid ${BORDER}`,
    fontSize: '13px',
    outline: 'none',
    boxSizing: 'border-box' as const,
    background: '#f8fafc',
    cursor: 'pointer',
  },
  textarea: {
    width: '100%',
    padding: '9px 11px',
    borderRadius: '8px',
    border: `1.5px solid ${BORDER}`,
    fontSize: '12px',
    outline: 'none',
    boxSizing: 'border-box' as const,
    resize: 'vertical' as const,
    background: '#f8fafc',
    lineHeight: 1.5,
  },
  charCount: {
    fontSize: '10px',
    color: '#94a3b8',
    textAlign: 'right' as const,
    marginTop: '3px',
  },
  assigneeList: {
    display: 'flex',
    flexDirection: 'column' as const,
    gap: '5px',
    maxHeight: '150px',
    overflowY: 'auto' as const,
  },
  assigneeBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    padding: '7px 10px',
    borderRadius: '8px',
    cursor: 'pointer',
    textAlign: 'left' as const,
    width: '100%',
    transition: 'all 0.15s',
  },
  avatar: {
    width: '24px',
    height: '24px',
    borderRadius: '50%',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: '9px',
    fontWeight: 700,
    color: '#fff',
    flexShrink: 0,
  },
  memberName: {
    fontSize: '12px',
    fontWeight: 500,
    flex: 1,
  },
  tick: {
    fontSize: '11px',
    color: ORANGE,
    fontWeight: 700,
  },
  error: {
    background: '#fff1f2',
    border: '1px solid #fecdd3',
    borderRadius: '8px',
    padding: '9px 12px',
    fontSize: '12px',
    color: '#e11d48',
    marginBottom: '12px',
  },
  submitBtn: {
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
    marginTop: '4px',
  },
  submitBtnDisabled: {
    opacity: 0.5,
    cursor: 'not-allowed',
    boxShadow: 'none',
  },
  footerNote: {
    textAlign: 'center' as const,
    fontSize: '10px',
    color: '#94a3b8',
    marginTop: '10px',
  },
};

export default CreateTaskForm;
