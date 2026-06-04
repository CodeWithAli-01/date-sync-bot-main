-- Scope report data to the authenticated user. This replaces early public
-- read/write policies that allowed users to see each other's reporting data.

CREATE TABLE IF NOT EXISTS public.employees (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  employee_code TEXT,
  region TEXT,
  city TEXT,
  designation TEXT,
  original_order INTEGER,
  status TEXT NOT NULL DEFAULT 'active',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.report_files (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name TEXT NOT NULL,
  file_hash TEXT NOT NULL,
  report_date DATE,
  upload_date TIMESTAMPTZ NOT NULL DEFAULT now(),
  uploaded_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  status TEXT NOT NULL DEFAULT 'pending',
  processed_status TEXT NOT NULL DEFAULT 'pending'
);

CREATE TABLE IF NOT EXISTS public.reports (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  date DATE NOT NULL,
  file_name TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.selfie_records (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_name TEXT NOT NULL,
  employee_code TEXT,
  date DATE NOT NULL,
  count INTEGER NOT NULL DEFAULT 0,
  source_file_id UUID,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.daily_records (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_code TEXT NOT NULL,
  report_date DATE NOT NULL,
  selfie_text TEXT NOT NULL DEFAULT '0 selfies with locations in grp',
  selfie_count INTEGER NOT NULL DEFAULT 0,
  calls_count INTEGER NOT NULL DEFAULT 0,
  total_count INTEGER NOT NULL DEFAULT 0,
  source_file_id UUID,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.error_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  file_name TEXT,
  error_message TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

ALTER TABLE public.employees ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.report_files ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.reports ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.selfie_records ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.daily_records ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.error_logs ENABLE ROW LEVEL SECURITY;

ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;
ALTER TABLE public.report_files ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;
ALTER TABLE public.reports ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;
ALTER TABLE public.selfie_records ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;
ALTER TABLE public.daily_records ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;
ALTER TABLE public.error_logs ADD COLUMN IF NOT EXISTS user_id UUID REFERENCES auth.users(id) ON DELETE CASCADE;

CREATE INDEX IF NOT EXISTS employees_user_id_idx ON public.employees(user_id);
CREATE INDEX IF NOT EXISTS report_files_user_id_idx ON public.report_files(user_id);
CREATE INDEX IF NOT EXISTS reports_user_id_idx ON public.reports(user_id);
CREATE INDEX IF NOT EXISTS selfie_records_user_id_idx ON public.selfie_records(user_id);
CREATE INDEX IF NOT EXISTS daily_records_user_id_idx ON public.daily_records(user_id);
CREATE INDEX IF NOT EXISTS error_logs_user_id_idx ON public.error_logs(user_id);

ALTER TABLE public.daily_records DROP CONSTRAINT IF EXISTS daily_records_employee_code_fkey;
ALTER TABLE public.selfie_records DROP CONSTRAINT IF EXISTS selfie_records_employee_code_fkey;
ALTER TABLE public.employees DROP CONSTRAINT IF EXISTS employees_name_key;
ALTER TABLE public.employees DROP CONSTRAINT IF EXISTS employees_employee_code_key;
ALTER TABLE public.employees DROP CONSTRAINT IF EXISTS employees_employee_code_unique;
ALTER TABLE public.report_files DROP CONSTRAINT IF EXISTS report_files_file_hash_key;
ALTER TABLE public.report_files DROP CONSTRAINT IF EXISTS report_files_file_name_unique;
ALTER TABLE public.reports DROP CONSTRAINT IF EXISTS reports_date_file_name_key;
ALTER TABLE public.reports DROP CONSTRAINT IF EXISTS reports_date_file_idx;
ALTER TABLE public.selfie_records DROP CONSTRAINT IF EXISTS selfie_records_employee_name_date_key;
ALTER TABLE public.selfie_records DROP CONSTRAINT IF EXISTS selfie_records_emp_date_idx;
ALTER TABLE public.selfie_records DROP CONSTRAINT IF EXISTS selfie_records_employee_code_date_key;
ALTER TABLE public.daily_records DROP CONSTRAINT IF EXISTS daily_records_employee_code_report_date_key;

DROP INDEX IF EXISTS employees_name_key;
DROP INDEX IF EXISTS employees_employee_code_key;
DROP INDEX IF EXISTS employees_employee_code_unique;
DROP INDEX IF EXISTS report_files_file_name_unique;
DROP INDEX IF EXISTS reports_date_file_idx;
DROP INDEX IF EXISTS selfie_records_emp_date_idx;
DROP INDEX IF EXISTS selfie_records_employee_code_date_key;

CREATE UNIQUE INDEX IF NOT EXISTS employees_user_employee_code_key
  ON public.employees(user_id, employee_code);

CREATE UNIQUE INDEX IF NOT EXISTS report_files_user_file_hash_key
  ON public.report_files(user_id, file_hash);

CREATE UNIQUE INDEX IF NOT EXISTS reports_user_date_file_key
  ON public.reports(user_id, date, file_name);

CREATE UNIQUE INDEX IF NOT EXISTS selfie_records_user_employee_code_date_key
  ON public.selfie_records(user_id, employee_code, date);

CREATE UNIQUE INDEX IF NOT EXISTS daily_records_user_employee_code_date_key
  ON public.daily_records(user_id, employee_code, report_date);

DROP POLICY IF EXISTS "public read employees" ON public.employees;
DROP POLICY IF EXISTS "public write employees" ON public.employees;
DROP POLICY IF EXISTS "users can read own employees" ON public.employees;
DROP POLICY IF EXISTS "users can insert own employees" ON public.employees;
DROP POLICY IF EXISTS "users can update own employees" ON public.employees;
DROP POLICY IF EXISTS "users can delete own employees" ON public.employees;

DROP POLICY IF EXISTS "public read report_files" ON public.report_files;
DROP POLICY IF EXISTS "public write report_files" ON public.report_files;
DROP POLICY IF EXISTS "users can read own report_files" ON public.report_files;
DROP POLICY IF EXISTS "users can insert own report_files" ON public.report_files;
DROP POLICY IF EXISTS "users can update own report_files" ON public.report_files;
DROP POLICY IF EXISTS "users can delete own report_files" ON public.report_files;

DROP POLICY IF EXISTS "public read reports" ON public.reports;
DROP POLICY IF EXISTS "public write reports" ON public.reports;
DROP POLICY IF EXISTS "users can read own reports" ON public.reports;
DROP POLICY IF EXISTS "users can insert own reports" ON public.reports;
DROP POLICY IF EXISTS "users can update own reports" ON public.reports;
DROP POLICY IF EXISTS "users can delete own reports" ON public.reports;

DROP POLICY IF EXISTS "public read selfie_records" ON public.selfie_records;
DROP POLICY IF EXISTS "public write selfie_records" ON public.selfie_records;
DROP POLICY IF EXISTS "users can read own selfie_records" ON public.selfie_records;
DROP POLICY IF EXISTS "users can insert own selfie_records" ON public.selfie_records;
DROP POLICY IF EXISTS "users can update own selfie_records" ON public.selfie_records;
DROP POLICY IF EXISTS "users can delete own selfie_records" ON public.selfie_records;

DROP POLICY IF EXISTS "public read daily_records" ON public.daily_records;
DROP POLICY IF EXISTS "public write daily_records" ON public.daily_records;
DROP POLICY IF EXISTS "users can read own daily_records" ON public.daily_records;
DROP POLICY IF EXISTS "users can insert own daily_records" ON public.daily_records;
DROP POLICY IF EXISTS "users can update own daily_records" ON public.daily_records;
DROP POLICY IF EXISTS "users can delete own daily_records" ON public.daily_records;

DROP POLICY IF EXISTS "public read error_logs" ON public.error_logs;
DROP POLICY IF EXISTS "public write error_logs" ON public.error_logs;
DROP POLICY IF EXISTS "users can read own error_logs" ON public.error_logs;
DROP POLICY IF EXISTS "users can insert own error_logs" ON public.error_logs;
DROP POLICY IF EXISTS "users can update own error_logs" ON public.error_logs;
DROP POLICY IF EXISTS "users can delete own error_logs" ON public.error_logs;

CREATE POLICY "users can read own employees" ON public.employees
  FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "users can insert own employees" ON public.employees
  FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can update own employees" ON public.employees
  FOR UPDATE USING (auth.uid() = user_id) WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can delete own employees" ON public.employees
  FOR DELETE USING (auth.uid() = user_id);

CREATE POLICY "users can read own report_files" ON public.report_files
  FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "users can insert own report_files" ON public.report_files
  FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can update own report_files" ON public.report_files
  FOR UPDATE USING (auth.uid() = user_id) WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can delete own report_files" ON public.report_files
  FOR DELETE USING (auth.uid() = user_id);

CREATE POLICY "users can read own reports" ON public.reports
  FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "users can insert own reports" ON public.reports
  FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can update own reports" ON public.reports
  FOR UPDATE USING (auth.uid() = user_id) WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can delete own reports" ON public.reports
  FOR DELETE USING (auth.uid() = user_id);

CREATE POLICY "users can read own selfie_records" ON public.selfie_records
  FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "users can insert own selfie_records" ON public.selfie_records
  FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can update own selfie_records" ON public.selfie_records
  FOR UPDATE USING (auth.uid() = user_id) WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can delete own selfie_records" ON public.selfie_records
  FOR DELETE USING (auth.uid() = user_id);

CREATE POLICY "users can read own daily_records" ON public.daily_records
  FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "users can insert own daily_records" ON public.daily_records
  FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can update own daily_records" ON public.daily_records
  FOR UPDATE USING (auth.uid() = user_id) WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can delete own daily_records" ON public.daily_records
  FOR DELETE USING (auth.uid() = user_id);

CREATE POLICY "users can read own error_logs" ON public.error_logs
  FOR SELECT USING (auth.uid() = user_id);
CREATE POLICY "users can insert own error_logs" ON public.error_logs
  FOR INSERT WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can update own error_logs" ON public.error_logs
  FOR UPDATE USING (auth.uid() = user_id) WITH CHECK (auth.uid() = user_id);
CREATE POLICY "users can delete own error_logs" ON public.error_logs
  FOR DELETE USING (auth.uid() = user_id);
