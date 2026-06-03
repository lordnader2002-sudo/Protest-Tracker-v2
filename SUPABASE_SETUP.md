# Supabase setup — password lock + Tracked Products sync

The dashboard now supports a **single shared-password login** and **cloud sync**
of the Tracked Products list, with the underlying event data served from a
**private** Supabase bucket (so the raw data is no longer public). Follow these
one-time steps. Until you finish them the site keeps working in its old public
mode (no login), so nothing breaks in the meantime.

## 1. Create a project
1. Sign up at https://supabase.com (free, no card) and create a new project.
2. Go to **Settings → API** and copy three values:
   - **Project URL** (e.g. `https://abcdxyz.supabase.co`)
   - **anon public** key  → goes in the dashboard (safe to expose)
   - **service_role** key → a **secret**; only used by the GitHub Action

## 2. Create the table + private bucket
Open **SQL Editor**, paste this, and run it:

```sql
-- Tracked products (one row per marked event)
create table if not exists public.tracked_events (
  event_id   text primary key,
  snapshot   jsonb       not null default '{}'::jsonb,
  note       text        not null default '',
  streams    text[]      not null default '{}',
  checked_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.tracked_events enable row level security;
create policy "authenticated full access" on public.tracked_events
  for all to authenticated using (true) with check (true);

-- Private bucket that holds dashboard_data.json + the Excel report
insert into storage.buckets (id, name, public)
values ('dashboard', 'dashboard', false)
on conflict (id) do nothing;
create policy "authenticated read dashboard" on storage.objects
  for select to authenticated using (bucket_id = 'dashboard');
```

(The Action uploads with the **service_role** key, which bypasses RLS, so no
insert policy on storage is needed.)

## 3. Create the single shared login
1. **Authentication → Users → Add user.**
2. Email: `dashboard@a-becker.me` (or your choice), set a **strong password** —
   this password is the only gate, so make it long.
3. Tick **Auto Confirm User** (so it can log in immediately).
4. **Authentication → Sign In / Providers** (or Settings): turn **off**
   "Allow new users to sign up" so only this account exists.

## 4. Wire up the dashboard
In `dashboard_v2.html`, near the top, fill the **SUPABASE CONFIG** block:

```js
const SUPABASE_URL      = "https://abcdxyz.supabase.co";   // your Project URL
const SUPABASE_ANON_KEY = "eyJhbGciOi...";                 // anon public key
const SHARED_EMAIL      = "dashboard@a-becker.me";         // the user from step 3
```

If the email you used differs, update `SHARED_EMAIL` to match.

## 5. Add the GitHub secrets
Repo → **Settings → Secrets and variables → Actions → New repository secret**:
- `SUPABASE_URL` = your Project URL
- `SUPABASE_SERVICE_KEY` = the **service_role** key

Once both secrets exist, the next workflow run automatically: uploads the data
to the private bucket, stops publishing it publicly, and seeds new-event
detection from Supabase.

## 6. (Recommended) Make the repo private
For a complete lock, set the GitHub **repo to Private**. Event data is gitignored
on `main` and no longer written to `gh-pages`, so the private bucket becomes the
only home for event data — but a private repo closes any residual exposure.

---

### Notes
- The **anon key is meant to be public**; security comes from Auth + RLS, so the
  shared password strength is what matters.
- Free projects pause after ~1 week idle; the 3×/day workflow keeps it awake.
- The classic dashboard (`classic.html`) has been retired — it couldn't be
  secured and read the now-private data.
