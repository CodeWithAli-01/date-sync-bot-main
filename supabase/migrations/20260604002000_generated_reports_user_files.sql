-- Make generated report history available across local and production for the same
-- authenticated Supabase user, including the generated XLSX payload.

ALTER TABLE public.generated_reports
  ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE,
  ADD COLUMN IF NOT EXISTS mime_type TEXT NOT NULL DEFAULT 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  ADD COLUMN IF NOT EXISTS file_data TEXT;

CREATE INDEX IF NOT EXISTS generated_reports_user_created_at_idx
  ON public.generated_reports(user_id, created_at DESC);

DROP POLICY IF EXISTS "public read generated_reports" ON public.generated_reports;
DROP POLICY IF EXISTS "public write generated_reports" ON public.generated_reports;
DROP POLICY IF EXISTS "users can read own generated_reports" ON public.generated_reports;
DROP POLICY IF EXISTS "users can insert own generated_reports" ON public.generated_reports;
DROP POLICY IF EXISTS "users can update own generated_reports" ON public.generated_reports;
DROP POLICY IF EXISTS "users can delete own generated_reports" ON public.generated_reports;

CREATE POLICY "users can read own generated_reports"
  ON public.generated_reports
  FOR SELECT
  USING (auth.uid() = user_id);

CREATE POLICY "users can insert own generated_reports"
  ON public.generated_reports
  FOR INSERT
  WITH CHECK (auth.uid() = user_id);

CREATE POLICY "users can update own generated_reports"
  ON public.generated_reports
  FOR UPDATE
  USING (auth.uid() = user_id)
  WITH CHECK (auth.uid() = user_id);

CREATE POLICY "users can delete own generated_reports"
  ON public.generated_reports
  FOR DELETE
  USING (auth.uid() = user_id);
