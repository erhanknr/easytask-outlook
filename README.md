# EasyTask Outlook Add-in

Turn emails into EasyTask tasks without leaving Microsoft Outlook.

---

## What it does

When you're reading an email in Outlook, click the **⚡ Create Task** button in the ribbon. A task pane slides open, pre-filled with:

- **Title** — the email subject
- **Description** — the first 400 characters of the email body, plus the sender and subject as a reference footnote

You then pick the project, section, assignees, and due date, and hit **Create Task**. The task lands in your EasyTask workspace immediately.

---

## Local development

### 1. Install dependencies

```bash
npm install
```

### 2. Add environment variables

Create a `.env` file in the project root (or copy `.env.example`):

```env
VITE_SUPABASE_URL=https://uduklpmxqjgwhfeljfwk.supabase.co
VITE_SUPABASE_ANON_KEY=your_anon_key_here
VITE_ADDIN_BASE_URL=https://localhost:3000
```

### 3. Start the dev server (HTTPS required)

```bash
npm run dev
```

The `@vitejs/plugin-basic-ssl` plugin automatically generates a self-signed certificate. Vite will serve on **https://localhost:3000**.

### 4. Trust the certificate in your browser

Open **https://localhost:3000** in your browser and accept the security warning. This step is required before Outlook will load the add-in — otherwise the task pane will show a blank white screen.

---

## Sideloading into Outlook (development)

### Outlook on the Web (recommended for testing)

1. Open [Outlook Web](https://outlook.office.com) and sign in.
2. Click the **Settings** gear → **View all Outlook settings**.
3. Go to **Mail → Customise actions → Manage add-ins** (or navigate directly to [https://outlook.office.com/mail/inclientstore](https://outlook.office.com/mail/inclientstore)).
4. Click **My add-ins** → **+ Add custom add-in** → **Add from file…**
5. Upload `manifest.xml` from this project.
6. Open any email — you should see **⚡ Create Task** in the message toolbar.

### Outlook Desktop (Windows)

1. Open Outlook.
2. Go to **File → Manage Add-ins** (opens Outlook Web add-ins page — same steps as above).

### Outlook Desktop (Mac)

1. Open Outlook.
2. Go to **Tools → Add-ins…**
3. Click the **+** button → select `manifest.xml`.

> **Note:** The dev server must be running at https://localhost:3000 whenever you use the add-in in development. If Outlook shows a loading spinner that never resolves, check that the dev server is up and the certificate has been trusted in your browser.

---

## Project structure

```
DevOutlook/
├── manifest.xml          # Office Add-in manifest (entry point for Outlook)
├── index.html            # Task pane HTML shell
├── dialog.html           # Auth dialog HTML shell
├── vite.config.ts        # Two entry points: taskpane + dialog, HTTPS enabled
├── tsconfig.json
├── .env                  # Supabase credentials + base URL
└── src/
    ├── lib/
    │   └── supabase.ts           # Supabase client (persistSession: true)
    ├── hooks/
    │   └── useEasyTaskData.ts    # Data hook: workspaces, projects, sections, members, createTask
    ├── taskpane.tsx              # Entry point — mounts <App /> on Office.onReady()
    ├── dialog.tsx                # Auth dialog — login form + messageParent()
    └── components/
        ├── App.tsx               # Auth state machine (loading / unauth / auth)
        ├── LoginPrompt.tsx       # "Connect EasyTask Account" — opens dialog via displayDialogAsync
        ├── CreateTaskForm.tsx    # Main form — reads email, creates task
        └── SuccessScreen.tsx     # Confirmation screen after task is created
```

---

## Deploying to production

### 1. Deploy the add-in files to Vercel

```bash
# From the DevOutlook directory
npx vercel --prod
```

Vercel will give you a URL like `https://easytask-outlook.vercel.app`.

### 2. Update URLs in `.env`

```env
VITE_ADDIN_BASE_URL=https://easytask-outlook.vercel.app
```

### 3. Update `manifest.xml`

Replace every occurrence of `https://localhost:3000` with your Vercel URL:

```xml
<!-- Task pane source -->
<SourceLocation DefaultValue="https://easytask-outlook.vercel.app/index.html" />

<!-- Dialog source (in LoginPrompt) -->
<!-- This is set via VITE_ADDIN_BASE_URL at build time -->

<!-- Icons -->
<bt:Image id="Icon16" DefaultValue="https://easytask-outlook.vercel.app/assets/icon-16.png" />
<bt:Image id="Icon32" DefaultValue="https://easytask-outlook.vercel.app/assets/icon-32.png" />
<bt:Image id="Icon80" DefaultValue="https://easytask-outlook.vercel.app/assets/icon-80.png" />

<!-- AppDomain -->
<AppDomain>https://easytask-outlook.vercel.app</AppDomain>
```

### 4. Rebuild and redeploy

```bash
npm run build
npx vercel --prod
```

---

## Distributing via Microsoft AppSource

To make the add-in available to all users in your Microsoft 365 organisation, or to publish it publicly on AppSource:

1. **Create a Partner Center account** at [partner.microsoft.com](https://partner.microsoft.com).
2. Go to **Marketplace offers → Office add-ins → New offer**.
3. Upload `manifest.xml` (pointing to your production Vercel URL).
4. Fill in the store listing (name, description, screenshots, privacy policy URL).
5. Submit for Microsoft review (typically 5–10 business days).

For organisation-only deployment (no AppSource review needed):

1. Go to the [Microsoft 365 admin center](https://admin.microsoft.com).
2. Navigate to **Settings → Integrated apps → Upload custom apps**.
3. Upload `manifest.xml`.
4. Assign to users or groups.

---

## Supabase requirements

The add-in uses the same Supabase project as the main EasyTask app. Make sure the following tables exist with the expected columns:

| Table | Required columns |
|---|---|
| `workspaces` | `id`, `name` |
| `workspace_members` | `workspace_id`, `user_id` |
| `projects` | `id`, `workspace_id`, `code`, `title`, `position` |
| `sections` | `id`, `project_id`, `name`, `position` |
| `profiles` | `id`, `full_name` |
| `tasks` | `id`, `title`, `description`, `section_id`, `project_id`, `due_date`, `created_by` |
| `task_assignees` | `task_id`, `user_id` |
| `task_logs` | `task_id`, `user_id`, `action`, `details` |

The Supabase anon key is safe to include in the client — Row Level Security on the tables handles access control.

---

## Icons

The manifest references icon files at `/assets/icon-16.png`, `/assets/icon-32.png`, and `/assets/icon-80.png`. Place these in the `public/assets/` folder before deploying. They should be the EasyTask logo (⚡) on a white or transparent background in the respective sizes.

---

## Troubleshooting

**Task pane shows a blank white screen**
→ The add-in can't load the dev server. Make sure `npm run dev` is running and you've opened https://localhost:3000 in a browser and accepted the certificate warning.

**"Could not open login window"**
→ Office failed to open the auth dialog. Check that `VITE_ADDIN_BASE_URL` is correct and the dialog.html page is accessible at `{BASE_URL}/dialog.html`.

**No workspaces/projects loading**
→ The logged-in user has no workspace memberships in `workspace_members`. Check the Supabase table and make sure the user's UUID is listed there.

**Tasks created but not appearing in the app**
→ Check that `section_id` and `project_id` foreign keys are correct. Also verify that Row Level Security policies allow insert on the `tasks` table for the authenticated user.
