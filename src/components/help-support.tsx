import { useEffect, useRef, useState, type FormEvent } from "react";
import { BarChart3, Bot, CalendarDays, FileSpreadsheet, Mail, MessageSquare, Phone, PieChart, Send, TableProperties } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Textarea } from "@/components/ui/textarea";
import type { SupportLanguage } from "@/lib/daily-report-help";
import { ADMIN_CONTACT, SUPPORT_COPY, planSupportReply, type SupportReply } from "@/lib/support-feedback";

const QUICK_MESSAGES = [
  { report: "Daily Report", icon: TableProperties },
  { report: "Monthly Report", icon: FileSpreadsheet },
  { report: "Doctor Coverage", icon: BarChart3 },
  { report: "Planned Summary", icon: CalendarDays },
  { report: "MREP Monthly Report", icon: PieChart },
  { report: "Avg. Per Day", icon: BarChart3 },
].map((item) => ({ ...item, message: `How can I Genrate the ${item.report}${/report$/i.test(item.report) ? "" : " rport"}?` }));

export function HelpSupport() {
  const [draft, setDraft] = useState("");
  const [messages, setMessages] = useState<(SupportReply & { role: "user" | "assistant" })[]>([]);
  const [feedbackLanguage, setFeedbackLanguage] = useState<SupportLanguage | null>(null);
  const [conversationLanguage, setConversationLanguage] = useState<SupportLanguage>("english");
  const [dailyReportContext, setDailyReportContext] = useState(false);
  const [feedbackPromptIndex, setFeedbackPromptIndex] = useState<number | null>(null);
  const [typingLanguage, setTypingLanguage] = useState<SupportLanguage | null>(null);
  const replyTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const conversationRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    const conversation = conversationRef.current;
    if (conversation && messages.length > 0) conversation.scrollTop = conversation.scrollHeight;
  }, [messages, typingLanguage]);

  useEffect(() => () => {
    if (replyTimer.current) clearTimeout(replyTimer.current);
  }, []);

  const sendMessage = (event?: FormEvent) => {
    event?.preventDefault();
    const message = draft.trim();
    if (!message || replyTimer.current) return;
    const plan = planSupportReply(message, feedbackLanguage, { language: conversationLanguage, dailyReport: dailyReportContext });
    setMessages((current) => [...current, { role: "user", text: message }]);
    setDraft("");
    inputRef.current?.focus();
    if (!plan.replies.length) return;
    setTypingLanguage(plan.language);
    replyTimer.current = setTimeout(() => {
      setMessages((current) => [...current, ...plan.replies.map((reply) => ({ ...reply, role: "assistant" as const }))]);
      const questionIndex = plan.replies.findIndex((reply) => reply.feedback);
      setFeedbackPromptIndex(questionIndex >= 0 ? messages.length + 1 + questionIndex : null);
      setFeedbackLanguage(plan.nextFeedback);
      setConversationLanguage(plan.language);
      setDailyReportContext(plan.dailyReport);
      setTypingLanguage(null);
      replyTimer.current = null;
    }, 2000);
  };

  return (
    <main className="mx-auto max-w-5xl">
      <Card className="flex h-[calc(100svh-12rem)] min-h-[24rem] flex-col overflow-hidden border-white/45 bg-card shadow-[var(--shadow-elegant)] lg:h-[calc(100svh-4rem)]">
        <header className="flex shrink-0 items-center justify-between gap-3 border-b border-border/60 bg-[var(--gradient-card)] px-4 py-4 sm:px-6 sm:py-5">
          <div className="flex min-w-0 items-center gap-3">
            <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-primary text-primary-foreground shadow-[var(--shadow-soft)]">
              <MessageSquare className="h-5 w-5" />
            </div>
            <div className="min-w-0">
              <h1 className="text-base font-semibold text-foreground sm:text-lg">Help &amp; Support</h1>
              <p className="mt-0.5 text-xs text-muted-foreground">Your reporting assistant</p>
            </div>
          </div>
          <span className="shrink-0 rounded-full border border-primary/15 bg-primary/5 px-3 py-1 text-xs font-medium text-primary">Chat preview</span>
        </header>

        <div ref={conversationRef} role="log" aria-label="Support conversation" aria-live="polite"
          className="min-h-0 flex-1 space-y-5 overflow-y-auto bg-gradient-to-b from-primary/[0.035] to-transparent p-4 sm:p-6">
          <div className="flex items-start gap-2.5 sm:gap-3">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-primary/10 text-primary"><Bot className="h-4 w-4" /></div>
            <div className="w-full max-w-sm space-y-4 rounded-2xl rounded-tl-sm border border-border/60 bg-card p-4 shadow-sm">
              <p className="text-sm font-medium leading-relaxed text-foreground">How can I help you?</p>
              <div className="flex flex-col gap-2" aria-label="Report help quick messages">
                {QUICK_MESSAGES.map(({ report, message, icon: Icon }) => (
                  <Button key={report} type="button" aria-pressed={draft === message}
                    className="min-h-10 w-full rounded-lg px-3 text-sm font-medium shadow-none"
                    onClick={() => {
                      setDraft(message);
                      inputRef.current?.focus();
                    }}>
                    <Icon className="h-4 w-4 shrink-0" />{report}
                  </Button>
                ))}
              </div>
            </div>
          </div>
          {messages.map((message, index) => (
            <div key={index} className={`flex ${message.role === "user" ? "justify-end" : "items-start gap-2.5 sm:gap-3"}`}>
              {message.role === "assistant" && <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-primary/10 text-primary"><Bot className="h-4 w-4" /></div>}
              <div dir="auto" className={`max-w-[85%] whitespace-pre-wrap break-words rounded-2xl px-4 py-3 text-sm leading-relaxed shadow-sm sm:max-w-[75%] ${message.role === "user" ? "rounded-tr-sm bg-primary text-primary-foreground" : "rounded-tl-sm border border-border/60 bg-card text-foreground"}`}>
                {message.text}
                {message.contact && <div className="mt-3 flex flex-col gap-2 border-t border-border/60 pt-3" dir="ltr">
                  <a href={ADMIN_CONTACT.phoneHref} className="flex items-center gap-2 text-primary hover:underline"><Phone className="h-4 w-4 shrink-0" />{ADMIN_CONTACT.phone}</a>
                  <a href={`mailto:${ADMIN_CONTACT.email}`} className="flex items-center gap-2 break-all text-primary hover:underline"><Mail className="h-4 w-4 shrink-0" />{ADMIN_CONTACT.email}</a>
                </div>}
                {message.feedback && <div className="mt-3 flex flex-wrap gap-2">
                  {[SUPPORT_COPY[message.feedback].yes, SUPPORT_COPY[message.feedback].no].map((label) => (
                    <Button key={label} type="button" size="sm" variant="outline" className="h-auto whitespace-normal py-2"
                      disabled={Boolean(typingLanguage) || feedbackPromptIndex !== index}
                      onClick={() => { setDraft(label); inputRef.current?.focus(); }}>{label}</Button>
                  ))}
                </div>}
              </div>
            </div>
          ))}
          {typingLanguage && <div className="flex items-start gap-3" role="status" aria-label={SUPPORT_COPY[typingLanguage].typing}>
            <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-xl bg-primary/10 text-primary"><Bot className="h-4 w-4" /></div>
            <div className="rounded-2xl rounded-tl-sm border border-border/60 bg-card px-4 py-3 shadow-sm">
              <div className="mb-1 flex gap-1" aria-hidden="true">{[0, 150, 300].map((delay) => <span key={delay} className="h-1.5 w-1.5 animate-bounce rounded-full bg-primary/65 motion-reduce:animate-none" style={{ animationDelay: `${delay}ms` }} />)}</div>
              <p dir="auto" className="text-xs text-muted-foreground">{SUPPORT_COPY[typingLanguage].typing}</p>
            </div>
          </div>}
        </div>

        <form onSubmit={sendMessage} className="shrink-0 border-t border-border/60 bg-card p-3 sm:p-5">
          <div className="flex items-end gap-2 rounded-2xl border border-border/80 bg-background/70 p-2 shadow-[var(--shadow-inset)] focus-within:border-primary/40 focus-within:ring-2 focus-within:ring-primary/10 sm:gap-3">
            <Textarea ref={inputRef} value={draft} onChange={(event) => setDraft(event.target.value)}
              dir="auto"
              aria-label="Message to support" placeholder="Type your message..." maxLength={2000} rows={2}
              className="max-h-36 min-h-12 min-w-0 resize-none border-0 bg-transparent px-2 py-2 text-sm shadow-none focus-visible:ring-0"
              onKeyDown={(event) => {
                if (event.key === "Enter" && !event.shiftKey && !event.nativeEvent.isComposing) {
                  event.preventDefault();
                  sendMessage();
                }
              }} />
            <Button type="submit" disabled={!draft.trim() || Boolean(typingLanguage)} className="h-11 shrink-0 gap-2 rounded-xl px-4 shadow-[var(--shadow-elegant)] sm:px-5" aria-label="Send message">
              <Send className="h-4 w-4" /><span>Send</span>
            </Button>
          </div>
        </form>
      </Card>
    </main>
  );
}
