import test from "node:test";
import assert from "node:assert/strict";
import { requestSupportAiReply } from "./support-ai.server.ts";

test("missing AI configuration does not call an external service", async () => {
  let called = false;
  const result = await requestSupportAiReply([{ role: "user", content: "hello" }], {
    fetchImpl: async () => { called = true; throw new Error("Unexpected request"); },
  });
  assert.equal(result.status, "not-configured");
  assert.equal(called, false);
});

test("AI receives conversation history and approved guides, and output text is read safely", async () => {
  const history = [
    { role: "user" as const, content: "kia hal ha" },
    { role: "assistant" as const, content: "Aap kaise hain?" },
    { role: "user" as const, content: "mn bi thk" },
  ];
  const result = await requestSupportAiReply(history, {
    apiKey: "test-key",
    fetchImpl: async (url, init) => {
      assert.equal(url, "https://api.openai.com/v1/responses");
      const body = JSON.parse(String(init?.body));
      assert.deepEqual(body.input, history);
      assert.match(body.instructions, /Roman Urdu/);
      assert.match(body.instructions, /Call Log Live/);
      assert.match(body.instructions, /Daily Report PDF files/);
      assert.match(body.instructions, /neutrophamra5@gmail.com/);
      assert.equal(body.store, false);
      return new Response(JSON.stringify({ output: [
        { type: "reasoning", content: [{ type: "output_text", text: "not a reply" }] },
        { type: "message", role: "assistant", content: [{ type: "output_text", text: "Yeh sun kar acha laga!" }] },
      ] }), { status: 200 });
    },
  });
  assert.deepEqual(result, { status: "ready", text: "Yeh sun kar acha laga!" });
});

test("provider failures and empty responses fall back without exposing provider details", async () => {
  for (const response of [new Response("secret provider error", { status: 429 }), new Response(JSON.stringify({ output: [] }))]) {
    assert.deepEqual(await requestSupportAiReply([{ role: "user", content: "hi" }], {
      apiKey: "test-key", fetchImpl: async () => response,
    }), { status: "unavailable" });
  }
});
