import { supabase } from "@/integrations/supabase/client";
import type { Session } from "@supabase/supabase-js";

const DEVICE_ID_KEY = "reporting-management-device-id";
export const MAX_AUTH_DEVICES = 4;
const DEVICE_BLOCK_MESSAGE = `This account is already active on ${MAX_AUTH_DEVICES} devices. Please sign out from another device and try again.`;

let activeClaim:
  | {
      key: string;
      promise: Promise<DeviceLimitResult>;
    }
  | undefined;

export interface DeviceLimitResult {
  allowed: boolean;
  message?: string;
  activeDeviceCount?: number;
  maxDevices?: number;
}

export async function enforceDeviceLimit(session: Session): Promise<DeviceLimitResult> {
  if (typeof window === "undefined") return { allowed: true };

  const deviceId = getOrCreateDeviceId();
  const claimKey = `${session.user.id}:${deviceId}`;
  if (activeClaim?.key === claimKey) return activeClaim.promise;

  const userAgent = window.navigator.userAgent;
  const promise = claimDevice(session, deviceId, userAgent);
  activeClaim = { key: claimKey, promise };

  try {
    return await promise;
  } finally {
    if (activeClaim?.promise === promise) activeClaim = undefined;
  }
}

async function claimDevice(
  session: Session,
  deviceId: string,
  userAgent: string,
): Promise<DeviceLimitResult> {
  const { data, error } = await supabase.rpc("claim_auth_device", {
    p_device_id: deviceId,
    p_user_agent: userAgent,
  });

  if (error) throw error;

  const result = parseDeviceLimitResult(data);
  console.info("[Device Auth] Device claim checked", {
    deviceId: maskDeviceId(deviceId),
    allowed: result.allowed,
    activeDeviceCount: result.activeDeviceCount,
    maxDevices: result.maxDevices,
  });
  if (result.allowed) return result;

  await sendBlockedLoginAlert(session, userAgent).catch((alertError) => {
    console.warn("Blocked login alert email failed", alertError);
  });
  await supabase.auth.signOut();

  return {
    ...result,
    message: DEVICE_BLOCK_MESSAGE,
  };
}

export async function revokeCurrentDeviceSession(): Promise<void> {
  if (typeof window === "undefined") return;
  const deviceId = getCurrentDeviceId();
  if (!deviceId) return;

  const { error } = await supabase.rpc("revoke_auth_device", {
    p_device_id: deviceId,
  });
  if (error) throw error;
}

export function getCurrentDeviceId(): string | null {
  if (typeof window === "undefined") return null;
  return window.localStorage.getItem(DEVICE_ID_KEY);
}

function getOrCreateDeviceId(): string {
  const existing = window.localStorage.getItem(DEVICE_ID_KEY);
  if (existing && existing.length >= 16) return existing;

  const generated = crypto.randomUUID();
  window.localStorage.setItem(DEVICE_ID_KEY, generated);
  return generated;
}

async function sendBlockedLoginAlert(session: Session, userAgent: string): Promise<void> {
  const { error } = await supabase.functions.invoke("send-login-alert", {
    headers: {
      Authorization: `Bearer ${session.access_token}`,
    },
    body: {
      userAgent,
    },
  });
  if (error) throw error;
}

function parseDeviceLimitResult(value: unknown): DeviceLimitResult {
  if (!value || typeof value !== "object") return { allowed: true };
  const record = value as Record<string, unknown>;
  return {
    allowed: record.allowed !== false,
    activeDeviceCount: numberValue(record.activeDeviceCount),
    maxDevices: numberValue(record.maxDevices),
  };
}

function numberValue(value: unknown): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

function maskDeviceId(deviceId: string): string {
  return `${deviceId.slice(0, 8)}...${deviceId.slice(-4)}`;
}
