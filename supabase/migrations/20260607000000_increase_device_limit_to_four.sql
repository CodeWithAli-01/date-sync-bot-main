-- Increase each authenticated user's active app device limit from 2 to 4.

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
  v_max_devices INTEGER := 4;
BEGIN
  IF v_user_id IS NULL THEN
    RAISE EXCEPTION 'Not authenticated';
  END IF;

  IF p_device_id IS NULL OR length(trim(p_device_id)) < 16 THEN
    RAISE EXCEPTION 'Invalid device id';
  END IF;

  -- Serialize claims for the same user so simultaneous auth events cannot race each other.
  PERFORM pg_advisory_xact_lock(hashtext(v_user_id::text));

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
  VALUES (v_user_id, p_device_id, left(p_user_agent, 500))
  ON CONFLICT (user_id, device_id) DO UPDATE
  SET
    last_seen_at = now(),
    updated_at = now(),
    revoked_at = NULL,
    user_agent = left(coalesce(excluded.user_agent, public.auth_device_sessions.user_agent), 500);

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
END;
$$;

GRANT EXECUTE ON FUNCTION public.claim_auth_device(TEXT, TEXT) TO authenticated;
