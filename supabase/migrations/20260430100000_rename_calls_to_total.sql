-- Correction: the paired PDF value is the report "Total" column, not Calls.
-- Keep old calls_count column untouched for compatibility; new writes use total_count.

ALTER TABLE public.daily_records ADD COLUMN IF NOT EXISTS total_count INTEGER NOT NULL DEFAULT 0;

UPDATE public.daily_records
SET total_count = calls_count
WHERE total_count = 0
  AND calls_count IS NOT NULL
  AND calls_count > 0;
