import { useState, useEffect, useCallback } from 'react';
import { supabase } from '../lib/supabase';

// ─── Types ───────────────────────────────────────────────────────────────────

export interface Workspace {
  id: string;
  name: string;
}

export interface Project {
  id: string;
  title: string;
  code: string;
  workspace_id: string;
}

export interface Section {
  id: string;
  name: string;
  project_id: string;
}

export interface Member {
  id: string;
  name: string;
  email: string;
  role: string;
}

export interface CreateTaskPayload {
  title: string;
  description: string;
  sectionId: string;
  projectId: string;
  assigneeIds: string[];
  dueDate: string | null;
  createdBy: string;
  emailSubject: string;
  emailSender: string;
}

// ─── Hook ────────────────────────────────────────────────────────────────────

export const useEasyTaskData = (userId: string | null) => {
  const [workspaces, setWorkspaces]   = useState<Workspace[]>([]);
  const [projects, setProjects]       = useState<Project[]>([]);
  const [sections, setSections]       = useState<Section[]>([]);
  const [members, setMembers]         = useState<Member[]>([]);
  const [loading, setLoading]         = useState(false);
  const [error, setError]             = useState<string | null>(null);

  // ── Fetch workspaces the user belongs to ─────────────────────────────────
  useEffect(() => {
    if (!userId) return;

    const fetchWorkspaces = async () => {
      setLoading(true);
      setError(null);

      // workspace_members → workspaces join
      const { data, error } = await supabase
        .from('workspace_members')
        .select('workspace_id, workspaces(id, name)')
        .eq('user_id', userId) as any;

      if (error) {
        setError('Could not load workspaces.');
      } else {
        const ws: Workspace[] = (data || [])
          .map((row: any) => row.workspaces)
          .filter(Boolean);
        setWorkspaces(ws);
      }
      setLoading(false);
    };

    fetchWorkspaces();
  }, [userId]);

  // ── Fetch projects for a workspace ───────────────────────────────────────
  const fetchProjects = useCallback(async (workspaceId: string) => {
    setProjects([]);
    setSections([]);
    setMembers([]);

    const { data, error } = await supabase
      .from('projects')
      .select('id, title, code, workspace_id')
      .eq('workspace_id', workspaceId)
      .order('position', { ascending: true, nullsFirst: false });

    if (error) { setError('Could not load projects.'); return; }
    setProjects(data || []);
  }, []);

  // ── Fetch sections for a project ─────────────────────────────────────────
  const fetchSections = useCallback(async (projectId: string) => {
    setSections([]);

    const { data, error } = await supabase
      .from('sections')
      .select('id, name, project_id')
      .eq('project_id', projectId)
      .order('position', { ascending: true, nullsFirst: false });

    if (error) { setError('Could not load sections.'); return; }
    setSections(data || []);
  }, []);

  // ── Fetch members for a workspace ────────────────────────────────────────
  const fetchMembers = useCallback(async (workspaceId: string) => {
    setMembers([]);

    // workspace_members + profiles join
    const { data, error } = await supabase
      .from('workspace_members')
      .select('user_id, role, profiles(id, full_name, email)')
      .eq('workspace_id', workspaceId) as any;

    if (error) { setError('Could not load members.'); return; }

    const mems: Member[] = (data || []).map((row: any) => ({
      id:    row.user_id,
      name:  row.profiles?.full_name || row.profiles?.email?.split('@')[0] || 'Unknown',
      email: row.profiles?.email || '',
      role:  row.role,
    }));
    setMembers(mems);
  }, []);

  // ── Create a task ─────────────────────────────────────────────────────────
  const createTask = useCallback(async (payload: CreateTaskPayload): Promise<{ success: boolean; taskId?: string }> => {
    try {
      // 1. Insert task
      const { data: taskData, error: taskError } = await supabase
        .from('tasks')
        .insert({
          title:       payload.title,
          description: payload.description
            ? `${payload.description}\n\n---\n📧 From email: "${payload.emailSubject}" (${payload.emailSender})`
            : `📧 Created from email: "${payload.emailSubject}" (${payload.emailSender})`,
          section_id:  payload.sectionId,
          created_by:  payload.createdBy,
          status:      'todo',
          due_date:    payload.dueDate || null,
        })
        .select('id')
        .single();

      if (taskError) throw taskError;
      const taskId = taskData.id;

      // 2. Insert assignees (task_assignees table)
      if (payload.assigneeIds.length > 0) {
        const assigneeRows = payload.assigneeIds.map(uid => ({
          task_id: taskId,
          user_id: uid,
        }));
        await supabase.from('task_assignees' as any).insert(assigneeRows);
      }

      // 3. Log task creation in task_logs
      await supabase.from('task_logs' as any).insert({
        task_id: taskId,
        user_id: payload.createdBy,
        action:  'created',
      });

      return { success: true, taskId };
    } catch (err: any) {
      console.error('createTask error:', err);
      return { success: false };
    }
  }, []);

  return {
    workspaces,
    projects,
    sections,
    members,
    loading,
    error,
    fetchProjects,
    fetchSections,
    fetchMembers,
    createTask,
  };
};
