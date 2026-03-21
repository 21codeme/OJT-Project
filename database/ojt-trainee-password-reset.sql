-- Trainee forgot-password (OTP) — hindi ito Supabase Auth; gumagamit ng ojt_trainees.password_hash
-- I-run sa Supabase SQL Editor pagkatapos ng ojt-tables.sql (kailangan ang ojt_trainees).
-- Pagkatapos: NOTIFY pgrst, 'reload schema';

CREATE TABLE IF NOT EXISTS ojt_password_reset_codes (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    email TEXT NOT NULL,
    code TEXT NOT NULL,
    expires_at TIMESTAMPTZ NOT NULL,
    used BOOLEAN NOT NULL DEFAULT false,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_ojt_password_reset_email_created
    ON ojt_password_reset_codes (lower(email), created_at DESC);

ALTER TABLE ojt_password_reset_codes ENABLE ROW LEVEL SECURITY;

-- Walang direct na access mula sa client; RPC lang (SECURITY DEFINER).
DROP POLICY IF EXISTS "No direct access reset codes" ON ojt_password_reset_codes;
CREATE POLICY "No direct access reset codes" ON ojt_password_reset_codes
    FOR ALL USING (false) WITH CHECK (false);

CREATE OR REPLACE FUNCTION trainee_password_reset_request(p_email text)
RETURNS json
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
    v_email text := lower(trim(p_email));
    v_exists boolean;
    v_code text;
    v_recent int;
BEGIN
    IF v_email IS NULL OR v_email = '' OR v_email !~ '^[^@\s]+@[^@\s]+\.[^@\s]+$' THEN
        RETURN json_build_object('ok', false, 'error', 'invalid_email');
    END IF;

    SELECT EXISTS (
        SELECT 1 FROM ojt_trainees t WHERE lower(trim(t.email)) = v_email
    ) INTO v_exists;

    IF NOT v_exists THEN
        RETURN json_build_object('ok', true);
    END IF;

    SELECT count(*)::int INTO v_recent
    FROM ojt_password_reset_codes
    WHERE lower(email) = v_email AND created_at > now() - interval '1 hour';

    IF v_recent >= 5 THEN
        RETURN json_build_object('ok', false, 'error', 'too_many_requests');
    END IF;

    v_code := lpad((floor(random() * 1000000))::int::text, 6, '0');

    DELETE FROM ojt_password_reset_codes
    WHERE lower(email) = v_email AND used = false;

    INSERT INTO ojt_password_reset_codes (email, code, expires_at)
    VALUES (v_email, v_code, now() + interval '15 minutes');

    RETURN json_build_object('ok', true, 'code', v_code);
END;
$$;

CREATE OR REPLACE FUNCTION trainee_password_reset_complete(p_email text, p_code text, p_new_password text)
RETURNS json
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
DECLARE
    v_email text := lower(trim(p_email));
    v_code text := trim(p_code);
    v_row ojt_password_reset_codes%ROWTYPE;
    v_updated int;
BEGIN
    IF v_email = '' OR v_code = '' OR p_new_password IS NULL OR length(trim(p_new_password)) < 6 THEN
        RETURN json_build_object('ok', false, 'error', 'invalid_input');
    END IF;

    SELECT * INTO v_row
    FROM ojt_password_reset_codes
    WHERE lower(email) = v_email
      AND code = v_code
      AND used = false
      AND expires_at > now()
    ORDER BY created_at DESC
    LIMIT 1;

    IF NOT FOUND THEN
        RETURN json_build_object('ok', false, 'error', 'invalid_code');
    END IF;

    UPDATE ojt_trainees
    SET password_hash = trim(p_new_password),
        updated_at = now()
    WHERE lower(trim(email)) = v_email;

    GET DIAGNOSTICS v_updated = ROW_COUNT;

    IF v_updated = 0 THEN
        RETURN json_build_object('ok', false, 'error', 'account_missing');
    END IF;

    UPDATE ojt_password_reset_codes SET used = true WHERE id = v_row.id;

    RETURN json_build_object('ok', true);
END;
$$;

REVOKE ALL ON FUNCTION trainee_password_reset_request(text) FROM PUBLIC;
REVOKE ALL ON FUNCTION trainee_password_reset_complete(text, text, text) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION trainee_password_reset_request(text) TO anon, authenticated;
GRANT EXECUTE ON FUNCTION trainee_password_reset_complete(text, text, text) TO anon, authenticated;

NOTIFY pgrst, 'reload schema';
