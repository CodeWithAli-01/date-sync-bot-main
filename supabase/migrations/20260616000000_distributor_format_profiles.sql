-- Persist Distributor Sales saved formats per authenticated user.

CREATE TABLE IF NOT EXISTS public.distributor_format_profiles (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  distributor_name TEXT NOT NULL,
  profile_name TEXT NOT NULL,
  source_sample_type TEXT NOT NULL DEFAULT 'Manual'
    CHECK (source_sample_type IN ('PDF', 'Excel', 'Screenshot', 'Manual')),
  source_sample_name TEXT,
  source_storage_path TEXT,
  source_mime_type TEXT,
  source_file_size BIGINT,
  uploaded_at TIMESTAMPTZ,
  numeric_order JSONB NOT NULL DEFAULT '[]'::jsonb,
  source_headers JSONB NOT NULL DEFAULT '[]'::jsonb,
  source_column_positions JSONB NOT NULL DEFAULT '{}'::jsonb,
  manual_column_mappings JSONB NOT NULL DEFAULT '{}'::jsonb,
  header_row_rule TEXT NOT NULL DEFAULT '',
  product_row_rule TEXT NOT NULL DEFAULT '',
  group_heading_rule TEXT NOT NULL DEFAULT '',
  group_total_rule TEXT NOT NULL DEFAULT '',
  date_extraction_rule TEXT NOT NULL DEFAULT '',
  distributor_name_extraction_rule TEXT,
  column_mapping_rules TEXT,
  product_code_extraction_rule TEXT,
  product_name_extraction_rule TEXT,
  multiline_product_name_rule TEXT,
  page_continuation_rule TEXT,
  active BOOLEAN NOT NULL DEFAULT true,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS distributor_format_profiles_user_name_key
  ON public.distributor_format_profiles(user_id, lower(distributor_name));

CREATE INDEX IF NOT EXISTS distributor_format_profiles_user_updated_idx
  ON public.distributor_format_profiles(user_id, updated_at DESC);

ALTER TABLE public.distributor_format_profiles ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "users can read own distributor format profiles"
  ON public.distributor_format_profiles;
DROP POLICY IF EXISTS "users can insert own distributor format profiles"
  ON public.distributor_format_profiles;
DROP POLICY IF EXISTS "users can update own distributor format profiles"
  ON public.distributor_format_profiles;
DROP POLICY IF EXISTS "users can delete own distributor format profiles"
  ON public.distributor_format_profiles;

CREATE POLICY "users can read own distributor format profiles"
  ON public.distributor_format_profiles
  FOR SELECT
  USING (auth.uid() = user_id);

CREATE POLICY "users can insert own distributor format profiles"
  ON public.distributor_format_profiles
  FOR INSERT
  WITH CHECK (auth.uid() = user_id);

CREATE POLICY "users can update own distributor format profiles"
  ON public.distributor_format_profiles
  FOR UPDATE
  USING (auth.uid() = user_id)
  WITH CHECK (auth.uid() = user_id);

CREATE POLICY "users can delete own distributor format profiles"
  ON public.distributor_format_profiles
  FOR DELETE
  USING (auth.uid() = user_id);

INSERT INTO storage.buckets (id, name, public)
VALUES ('distributor-format-samples', 'distributor-format-samples', false)
ON CONFLICT (id) DO NOTHING;

DROP POLICY IF EXISTS "users can read own distributor format samples"
  ON storage.objects;
DROP POLICY IF EXISTS "users can insert own distributor format samples"
  ON storage.objects;
DROP POLICY IF EXISTS "users can update own distributor format samples"
  ON storage.objects;
DROP POLICY IF EXISTS "users can delete own distributor format samples"
  ON storage.objects;

CREATE POLICY "users can read own distributor format samples"
  ON storage.objects
  FOR SELECT
  USING (
    bucket_id = 'distributor-format-samples'
    AND auth.uid()::text = (storage.foldername(name))[1]
  );

CREATE POLICY "users can insert own distributor format samples"
  ON storage.objects
  FOR INSERT
  WITH CHECK (
    bucket_id = 'distributor-format-samples'
    AND auth.uid()::text = (storage.foldername(name))[1]
  );

CREATE POLICY "users can update own distributor format samples"
  ON storage.objects
  FOR UPDATE
  USING (
    bucket_id = 'distributor-format-samples'
    AND auth.uid()::text = (storage.foldername(name))[1]
  )
  WITH CHECK (
    bucket_id = 'distributor-format-samples'
    AND auth.uid()::text = (storage.foldername(name))[1]
  );

CREATE POLICY "users can delete own distributor format samples"
  ON storage.objects
  FOR DELETE
  USING (
    bucket_id = 'distributor-format-samples'
    AND auth.uid()::text = (storage.foldername(name))[1]
  );
