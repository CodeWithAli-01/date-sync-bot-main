-- Production schema for Pharma Selfie Reporting System.
-- This migration is additive and UPSERT-friendly. It does not delete old data.

ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS region TEXT;
ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS city TEXT;
ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS designation TEXT;
ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS original_order INTEGER;
ALTER TABLE public.employees ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ NOT NULL DEFAULT now();

CREATE UNIQUE INDEX IF NOT EXISTS employees_employee_code_unique
  ON public.employees(employee_code)
  WHERE employee_code IS NOT NULL;

CREATE INDEX IF NOT EXISTS employees_original_order_idx ON public.employees(original_order);

ALTER TABLE public.report_files ADD COLUMN IF NOT EXISTS status TEXT NOT NULL DEFAULT 'pending';
ALTER TABLE public.report_files ADD COLUMN IF NOT EXISTS uploaded_at TIMESTAMPTZ NOT NULL DEFAULT now();

UPDATE public.report_files
SET status = COALESCE(NULLIF(status, ''), processed_status, 'done')
WHERE processed_status IS NOT NULL;

UPDATE public.report_files
SET uploaded_at = COALESCE(uploaded_at, upload_date, now());

CREATE UNIQUE INDEX IF NOT EXISTS report_files_file_name_unique ON public.report_files(file_name);

CREATE TABLE IF NOT EXISTS public.daily_records (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  employee_code TEXT NOT NULL,
  report_date DATE NOT NULL,
  selfie_text TEXT NOT NULL DEFAULT '0 selfies with locations in grp',
  selfie_count INTEGER NOT NULL DEFAULT 0,
  calls_count INTEGER NOT NULL DEFAULT 0,
  source_file_id UUID REFERENCES public.report_files(id),
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(employee_code, report_date)
);

CREATE INDEX IF NOT EXISTS daily_records_employee_code_idx ON public.daily_records(employee_code);
CREATE INDEX IF NOT EXISTS daily_records_report_date_idx ON public.daily_records(report_date);

ALTER TABLE public.daily_records ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "public read daily_records" ON public.daily_records;
DROP POLICY IF EXISTS "public write daily_records" ON public.daily_records;
CREATE POLICY "public read daily_records" ON public.daily_records FOR SELECT USING (true);
CREATE POLICY "public write daily_records" ON public.daily_records FOR ALL USING (true) WITH CHECK (true);
