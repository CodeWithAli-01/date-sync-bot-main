
-- Add employee_code to employees
ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS employee_code TEXT;
ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT now();
CREATE UNIQUE INDEX IF NOT EXISTS employees_name_key ON public.employees(name);
CREATE INDEX IF NOT EXISTS employees_code_idx ON public.employees(employee_code);

-- report_files table (file-hash dedup)
CREATE TABLE IF NOT EXISTS public.report_files (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name TEXT NOT NULL,
  file_hash TEXT NOT NULL UNIQUE,
  report_date DATE,
  upload_date TIMESTAMPTZ NOT NULL DEFAULT now(),
  processed_status TEXT NOT NULL DEFAULT 'pending'
);
CREATE INDEX IF NOT EXISTS report_files_hash_idx ON public.report_files(file_hash);
ALTER TABLE public.report_files ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "public read report_files" ON public.report_files;
DROP POLICY IF EXISTS "public write report_files" ON public.report_files;
CREATE POLICY "public read report_files" ON public.report_files FOR SELECT USING (true);
CREATE POLICY "public write report_files" ON public.report_files FOR ALL USING (true) WITH CHECK (true);

-- selfie_records: add employee_code + source_file_id, unique constraint
ALTER TABLE public.selfie_records ADD COLUMN IF NOT EXISTS employee_code TEXT;
ALTER TABLE public.selfie_records ADD COLUMN IF NOT EXISTS source_file_id UUID;
CREATE UNIQUE INDEX IF NOT EXISTS selfie_records_emp_date_idx
  ON public.selfie_records(employee_name, date);
CREATE INDEX IF NOT EXISTS selfie_records_code_idx ON public.selfie_records(employee_code);
CREATE INDEX IF NOT EXISTS selfie_records_date_idx ON public.selfie_records(date);

-- error_logs table
CREATE TABLE IF NOT EXISTS public.error_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name TEXT,
  error_message TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
ALTER TABLE public.error_logs ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "public read error_logs" ON public.error_logs;
DROP POLICY IF EXISTS "public write error_logs" ON public.error_logs;
CREATE POLICY "public read error_logs" ON public.error_logs FOR SELECT USING (true);
CREATE POLICY "public write error_logs" ON public.error_logs FOR ALL USING (true) WITH CHECK (true);

-- reports table: add unique on (date, file_name) for upsert
CREATE UNIQUE INDEX IF NOT EXISTS reports_date_file_idx ON public.reports(date, file_name);
