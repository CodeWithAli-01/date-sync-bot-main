import { createServerFn } from "@tanstack/react-start";
import { z } from "zod";
import { requireSupabaseAuth } from "@/integrations/supabase/auth-middleware";

export const supportAiInput = z.object({
  messages: z.array(z.object({
    role: z.enum(["user", "assistant"]),
    content: z.string().trim().min(1).max(2000),
  })).min(1).max(12),
}).refine((input) => input.messages.at(-1)?.role === "user", "The last message must be from the customer.");

export const askSupportAi = createServerFn({ method: "POST" })
  .middleware([requireSupabaseAuth])
  .inputValidator(supportAiInput)
  .handler(async ({ data }) => {
    const { requestSupportAiReply } = await import("./support-ai.server");
    return requestSupportAiReply(data.messages, {
      apiKey: process.env.OPENAI_API_KEY,
      model: process.env.SUPPORT_AI_MODEL,
    });
  });
