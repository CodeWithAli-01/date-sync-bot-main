import { createClient } from "https://esm.sh/@supabase/supabase-js@2.105.1";

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, anon-key, content-type",
};

const appName = Deno.env.get("APP_NAME") || "Reporting Management";
const fromAddress = Deno.env.get("SECURITY_ALERT_FROM") || "Reporting Management <onboarding@resend.dev>";
const resendApiKey = Deno.env.get("RESEND_API_KEY");
const supabaseUrl = Deno.env.get("SUPABASE_URL");
const serviceRoleKey = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");

Deno.serve(async (request) => {
  if (request.method === "OPTIONS") {
    return new Response("ok", { headers: corsHeaders });
  }

  try {
    if (!supabaseUrl || !serviceRoleKey) {
      throw new Error("Supabase function environment is missing service credentials.");
    }

    const authHeader = request.headers.get("Authorization") || "";
    const token = authHeader.replace(/^Bearer\s+/i, "").trim();
    if (!token) {
      return json({ error: "Missing authorization token." }, 401);
    }

    const supabase = createClient(supabaseUrl, serviceRoleKey, {
      auth: { persistSession: false, autoRefreshToken: false },
    });
    const { data, error } = await supabase.auth.getUser(token);
    if (error || !data.user?.email) {
      return json({ error: "Invalid user session." }, 401);
    }

    const body = await request.json().catch(() => ({}));
    const userAgent = cleanText(body.userAgent, "Unknown device");
    const attemptedAt = new Date().toLocaleString("en-US", {
      dateStyle: "medium",
      timeStyle: "short",
      timeZone: "UTC",
    });

    if (!resendApiKey) {
      console.warn("[send-login-alert] RESEND_API_KEY is not configured.");
      return json({ sent: false, reason: "RESEND_API_KEY is not configured." });
    }

    const emailResult = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${resendApiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        from: fromAddress,
        to: data.user.email,
        subject: `${appName}: blocked login attempt`,
        text: [
          `A login attempt for your ${appName} account was blocked because your account is already active on 2 devices.`,
          "",
          `Attempt time: ${attemptedAt} UTC`,
          `Device: ${userAgent}`,
          "",
          "If this was you, sign out from another device and try again. If this was not you, change your password immediately.",
        ].join("\n"),
        html: `
          <div style="font-family:Arial,sans-serif;line-height:1.55;color:#111827">
            <h2 style="margin:0 0 12px">${appName}: blocked login attempt</h2>
            <p>A login attempt for your account was blocked because your account is already active on <strong>2 devices</strong>.</p>
            <p><strong>Attempt time:</strong> ${escapeHtml(attemptedAt)} UTC<br />
            <strong>Device:</strong> ${escapeHtml(userAgent)}</p>
            <p>If this was you, sign out from another device and try again. If this was not you, change your password immediately.</p>
          </div>
        `,
      }),
    });

    if (!emailResult.ok) {
      const errorText = await emailResult.text();
      console.error("[send-login-alert] Resend failed", errorText);
      return json({ sent: false, error: "Email provider failed." }, 502);
    }

    const { data: latestAlert } = await supabase
      .from("auth_login_alerts")
      .select("id")
      .eq("user_id", data.user.id)
      .is("email_sent_at", null)
      .order("attempted_at", { ascending: false })
      .limit(1)
      .maybeSingle();

    if (latestAlert?.id) {
      await supabase
        .from("auth_login_alerts")
        .update({ email_sent_at: new Date().toISOString() })
        .eq("id", latestAlert.id);
    }

    return json({ sent: true });
  } catch (error) {
    console.error("[send-login-alert] failed", error);
    return json({ error: error instanceof Error ? error.message : "Unknown error." }, 500);
  }
});

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, "Content-Type": "application/json" },
  });
}

function cleanText(value: unknown, fallback: string): string {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  return text ? text.slice(0, 500) : fallback;
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (char) => {
    const entities: Record<string, string> = {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#039;",
    };
    return entities[char];
  });
}
