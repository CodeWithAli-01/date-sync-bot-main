-- Stores generated report outputs that are also shown in the app's local History panel.
-- report_files remains as a compatibility fallback for file-hash dedup and older dashboards.

CREATE TABLE IF NOT EXISTS public.generated_reports (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  report_key TEXT NOT NULL UNIQUE,
  local_history_id TEXT,
  file_name TEXT NOT NULL,
  report_type TEXT NOT NULL DEFAULT 'Report',
  dates JSONB NOT NULL DEFAULT '[]'::jsonb,
  pdf_count INTEGER NOT NULL DEFAULT 0,
  total_employees INTEGER NOT NULL DEFAULT 0,
  matched_employees INTEGER NOT NULL DEFAULT 0,
  file_size INTEGER NOT NULL DEFAULT 0,
  file_hash TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS generated_reports_report_type_idx
  ON public.generated_reports(report_type);

CREATE INDEX IF NOT EXISTS generated_reports_created_at_idx
  ON public.generated_reports(created_at DESC);

CREATE INDEX IF NOT EXISTS generated_reports_file_hash_idx
  ON public.generated_reports(file_hash);

ALTER TABLE public.generated_reports ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "public read generated_reports" ON public.generated_reports;
DROP POLICY IF EXISTS "public write generated_reports" ON public.generated_reports;

CREATE POLICY "public read generated_reports"
  ON public.generated_reports
  FOR SELECT
  USING (true);

CREATE POLICY "public write generated_reports"
  ON public.generated_reports
  FOR ALL
  USING (true)
  WITH CHECK (true);
