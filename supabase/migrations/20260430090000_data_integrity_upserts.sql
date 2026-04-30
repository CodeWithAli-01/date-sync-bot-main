-- Data integrity hardening for rerunnable syncs.
-- Keep one selfie record per employee_code + date when a code is available.

WITH ranked AS (
  SELECT
    id,
    employee_code,
    date,
    count,
    max(count) OVER (PARTITION BY employee_code, date) AS max_count,
    row_number() OVER (
      PARTITION BY employee_code, date
      ORDER BY count DESC, created_at ASC, id ASC
    ) AS rn
  FROM public.selfie_records
  WHERE employee_code IS NOT NULL
)
UPDATE public.selfie_records s
SET count = ranked.max_count
FROM ranked
WHERE s.id = ranked.id
  AND ranked.rn = 1
  AND s.count <> ranked.max_count;

WITH ranked AS (
  SELECT
    id,
    row_number() OVER (
      PARTITION BY employee_code, date
      ORDER BY count DESC, created_at ASC, id ASC
    ) AS rn
  FROM public.selfie_records
  WHERE employee_code IS NOT NULL
)
DELETE FROM public.selfie_records s
USING ranked
WHERE s.id = ranked.id
  AND ranked.rn > 1;

CREATE UNIQUE INDEX IF NOT EXISTS selfie_records_employee_code_date_key
  ON public.selfie_records(employee_code, date)
  WHERE employee_code IS NOT NULL;

CREATE UNIQUE INDEX IF NOT EXISTS employees_employee_code_key
  ON public.employees(employee_code)
  WHERE employee_code IS NOT NULL;
