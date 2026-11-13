# Help & Support AI

The support chat now has an authenticated OpenAI connection. It sends the recent
conversation and the approved Daily/Monthly Report guides, so the assistant can
understand follow-up questions and respond naturally in the customer's language.
Report generation is independent of this connection.

## Configuration

Set these server environment variables:

```dotenv
OPENAI_API_KEY=your-private-api-key
SUPPORT_AI_MODEL=gpt-4.1-mini
```

Keep the key private. Do not prefix it with `VITE_`, put it in source code, or send
it in chat. AI requests use the configured account's API usage/billing.

The server also needs `SUPABASE_URL` and `SUPABASE_PUBLISHABLE_KEY` for the existing
authentication middleware. Only signed-in customers can call the AI function.

For local development, put the variables in the ignored `.env` file alongside
the existing Supabase variables. Stop the current development server, then run:

```powershell
node --env-file=.env node_modules/vite/bin/vite.js --host 127.0.0.1 --port 5173
```

For production, add the variables to the Vercel project's Production environment
and redeploy. No API key is configured by this code change.

## Verification

Open Help & Support while signed in. Try `kia hal ha`, then `mn bi thk`, or ask
the difference between Daily and Monthly Report. A successful API response shows
the **AI assistant** badge. Without a configured key, or if the provider is
unavailable, the chat continues using **Saved guides**; this is not a live AI reply.

The typing indicator stays visible for at least two seconds and continues while
the AI is responding. Selecting a report still only fills the composer; Send
submits it. No uploaded report files are sent to the model.

API shape: [OpenAI Responses API](https://developers.openai.com/api/reference/resources/responses/methods/create).
Model: [GPT-4.1 mini](https://developers.openai.com/api/docs/models/gpt-4.1-mini).
