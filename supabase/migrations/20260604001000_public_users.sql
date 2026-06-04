-- Public user profile mirror for Supabase Auth users.
-- Supabase Auth stores accounts in auth.users, which is visible in Authentication,
-- not as a public Table Editor table. This table lets the app query/profile users
-- from the public schema while keeping auth.users as the source of truth.

CREATE TABLE IF NOT EXISTS public.users (
  id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
  email TEXT UNIQUE,
  phone TEXT,
  display_name TEXT,
  avatar_url TEXT,
  provider TEXT,
  last_sign_in_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS users_email_idx ON public.users(email);
CREATE INDEX IF NOT EXISTS users_updated_at_idx ON public.users(updated_at DESC);

ALTER TABLE public.users ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "users can read own profile" ON public.users;
DROP POLICY IF EXISTS "users can insert own profile" ON public.users;
DROP POLICY IF EXISTS "users can update own profile" ON public.users;

CREATE POLICY "users can read own profile"
  ON public.users
  FOR SELECT
  USING (auth.uid() = id);

CREATE POLICY "users can insert own profile"
  ON public.users
  FOR INSERT
  WITH CHECK (auth.uid() = id);

CREATE POLICY "users can update own profile"
  ON public.users
  FOR UPDATE
  USING (auth.uid() = id)
  WITH CHECK (auth.uid() = id);
