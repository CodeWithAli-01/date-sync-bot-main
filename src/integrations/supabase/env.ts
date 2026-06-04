const QUOTE_CHARS = /^[\s"'`“”‘’]+|[\s"'`“”‘’]+$/g;

export function cleanSupabaseEnv(value: string | undefined): string {
  return String(value ?? "")
    .replace(/^\uFEFF/, "")
    .replace(QUOTE_CHARS, "")
    .trim();
}

export function assertHeaderSafe(name: string, value: string) {
  for (const char of value) {
    const code = char.charCodeAt(0);
    if (code > 255) {
      throw new Error(
        `${name} contains an invalid character. Re-copy the value from Supabase and remove smart quotes or hidden characters.`,
      );
    }
  }
}
