
CREATE TABLE public.employees (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  status TEXT NOT NULL DEFAULT 'active',
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE public.reports (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  date DATE NOT NULL,
  file_name TEXT NOT NULL,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(date, file_name)
);

CREATE TABLE public.selfie_records (
  id UUID NOT NULL DEFAULT gen_random_uuid() PRIMARY KEY,
  employee_name TEXT NOT NULL,
  date DATE NOT NULL,
  count INTEGER NOT NULL DEFAULT 0,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(employee_name, date)
);

CREATE INDEX idx_selfie_employee ON public.selfie_records(employee_name);
CREATE INDEX idx_selfie_date ON public.selfie_records(date);
CREATE INDEX idx_reports_date ON public.reports(date);

ALTER TABLE public.employees ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.reports ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.selfie_records ENABLE ROW LEVEL SECURITY;

CREATE POLICY "public read employees" ON public.employees FOR SELECT USING (true);
CREATE POLICY "public write employees" ON public.employees FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "public read reports" ON public.reports FOR SELECT USING (true);
CREATE POLICY "public write reports" ON public.reports FOR ALL USING (true) WITH CHECK (true);
CREATE POLICY "public read selfie_records" ON public.selfie_records FOR SELECT USING (true);
CREATE POLICY "public write selfie_records" ON public.selfie_records FOR ALL USING (true) WITH CHECK (true);
