-- Limit each authenticated user to two active app devices.
-- A device is released when the user signs out, or after 30 days without activity.

CREATE TABLE IF NOT EXISTS public.auth_device_sessions (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  device_id TEXT NOT NULL,
  user_agent TEXT,
  first_seen_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  last_seen_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  revoked_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  UNIQUE(user_id, device_id)
);

CREATE TABLE IF NOT EXISTS public.auth_login_alerts (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
  email TEXT,
  device_id TEXT,
  user_agent TEXT,
  attempted_at TIMESTAMPTZ NOT NULL DEFAULT now(),
  email_sent_at TIMESTAMPTZ,
  created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS auth_device_sessions_user_active_idx
  ON public.auth_device_sessions(user_id, revoked_at, last_seen_at DESC);

CREATE INDEX IF NOT EXISTS auth_login_alerts_user_attempted_idx
  ON public.auth_login_alerts(user_id, attempted_at DESC);

ALTER TABLE public.auth_device_sessions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.auth_login_alerts ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "users can read own device sessions" ON public.auth_device_sessions;
DROP POLICY IF EXISTS "users can read own login alerts" ON public.auth_login_alerts;

CREATE POLICY "users can read own device sessions"
  ON public.auth_device_sessions
  FOR SELECT
  USING (auth.uid() = user_id);

CREATE POLICY "users can read own login alerts"
  ON public.auth_login_alerts
  FOR SELECT
  USING (auth.uid() = user_id);

CREATE OR REPLACE FUNCTION public.claim_auth_device(
  p_device_id TEXT,
  p_user_agent TEXT DEFAULT NULL
)
RETURNS JSONB
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_user_id UUID := auth.uid();
  v_email TEXT;
  v_active_count INTEGER;
  v_max_devices INTEGER := 2;
BEGIN
  IF v_user_id IS NULL THEN
    RAISE EXCEPTION 'Not authenticated';
  END IF;

  IF p_device_id IS NULL OR length(trim(p_device_id)) < 16 THEN
    RAISE EXCEPTION 'Invalid device id';
  END IF;

  UPDATE public.auth_device_sessions
  SET
    last_seen_at = now(),
    updated_at = now(),
    revoked_at = NULL,
    user_agent = left(coalesce(p_user_agent, user_agent), 500)
  WHERE user_id = v_user_id
    AND device_id = p_device_id;

  IF FOUND THEN
    SELECT count(*) INTO v_active_count
    FROM public.auth_device_sessions
    WHERE user_id = v_user_id
      AND revoked_at IS NULL
      AND last_seen_at > now() - interval '30 days';

    RETURN jsonb_build_object(
      'allowed', true,
      'activeDeviceCount', v_active_count,
      'maxDevices', v_max_devices
    );
  END IF;

  SELECT count(*) INTO v_active_count
  FROM public.auth_device_sessions
  WHERE user_id = v_user_id
    AND revoked_at IS NULL
    AND last_seen_at > now() - interval '30 days';

  IF v_active_count >= v_max_devices THEN
    SELECT email INTO v_email FROM auth.users WHERE id = v_user_id;

    INSERT INTO public.auth_login_alerts (user_id, email, device_id, user_agent)
    VALUES (v_user_id, v_email, p_device_id, left(p_user_agent, 500));

    RETURN jsonb_build_object(
      'allowed', false,
      'activeDeviceCount', v_active_count,
      'maxDevices', v_max_devices
    );
  END IF;

  INSERT INTO public.auth_device_sessions (user_id, device_id, user_agent)
  VALUES (v_user_id, p_device_id, left(p_user_agent, 500));

  RETURN jsonb_build_object(
    'allowed', true,
    'activeDeviceCount', v_active_count + 1,
    'maxDevices', v_max_devices
  );
END;
$$;

CREATE OR REPLACE FUNCTION public.revoke_auth_device(p_device_id TEXT)
RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_user_id UUID := auth.uid();
BEGIN
  IF v_user_id IS NULL THEN
    RETURN;
  END IF;

  UPDATE public.auth_device_sessions
  SET revoked_at = now(), updated_at = now()
  WHERE user_id = v_user_id
    AND device_id = p_device_id
    AND revoked_at IS NULL;
END;
$$;

CREATE OR REPLACE FUNCTION public.mark_latest_login_alert_sent()
RETURNS VOID
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
  v_user_id UUID := auth.uid();
  v_alert_id UUID;
BEGIN
  IF v_user_id IS NULL THEN
    RETURN;
  END IF;

  SELECT id INTO v_alert_id
  FROM public.auth_login_alerts
  WHERE user_id = v_user_id
    AND email_sent_at IS NULL
  ORDER BY attempted_at DESC
  LIMIT 1;

  IF v_alert_id IS NOT NULL THEN
    UPDATE public.auth_login_alerts
    SET email_sent_at = now()
    WHERE id = v_alert_id;
  END IF;
END;
$$;

GRANT EXECUTE ON FUNCTION public.claim_auth_device(TEXT, TEXT) TO authenticated;
GRANT EXECUTE ON FUNCTION public.revoke_auth_device(TEXT) TO authenticated;
GRANT EXECUTE ON FUNCTION public.mark_latest_login_alert_sent() TO authenticated;
