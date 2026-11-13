import type * as React from "react";
import { Check, Upload, X } from "lucide-react";
import { Card } from "@/components/ui/card";
import { Button } from "@/components/ui/button";

export function UploadCard({
  step,
  title,
  subtitle,
  icon,
  ready = false,
  onClick,
  children,
  compact = false,
}: {
  step: number;
  title: string;
  subtitle: string;
  icon: React.ReactNode;
  accept: string;
  multiple?: boolean;
  ready?: boolean;
  onClick: () => void;
  children: React.ReactNode;
  compact?: boolean;
}) {
  return (
    <Card
      className={`flex flex-col border-white/45 bg-[var(--gradient-card)] shadow-[var(--shadow-soft)] transition-shadow hover:shadow-[var(--shadow-elegant)] ${
        compact ? "h-[22rem] p-4" : "min-h-[16rem] p-4 sm:p-5"
      }`}
    >
      <div className={compact ? "mb-3 space-y-3" : "mb-4 space-y-4"}>
        <div className="flex items-start justify-between gap-3">
          <div className="flex min-h-9 items-center gap-2 text-xs font-bold uppercase text-primary">
            Step {step}
            {ready && <Check className="h-3.5 w-3.5 text-success" />}
          </div>
          <Button
            variant="outline"
            size="sm"
            onClick={onClick}
            className="min-h-9 shrink-0 whitespace-nowrap px-4"
          >
            <Upload className="mr-1.5 h-3.5 w-3.5" />
            Upload
          </Button>
        </div>
        <div className="min-w-0">
          <h2
            className={`flex min-w-0 items-start gap-3 font-semibold leading-tight text-foreground ${
              compact ? "text-base sm:text-lg" : "text-base sm:text-lg 2xl:text-xl"
            }`}
          >
            <span
              className={`mt-0.5 flex shrink-0 items-center justify-center rounded-xl bg-card text-primary shadow-[var(--shadow-inset)] ${
                compact ? "h-8 w-8" : "h-9 w-9"
              }`}
            >
              {icon}
            </span>
            <span className="min-w-0 whitespace-normal break-words">{title}</span>
          </h2>
          <p
            className={`mt-2 max-w-[28rem] text-sm leading-relaxed text-muted-foreground ${
              compact ? "" : "sm:text-base"
            }`}
          >
            {subtitle}
          </p>
        </div>
      </div>
      <div
        role={ready ? undefined : "button"}
        tabIndex={ready ? undefined : 0}
        onClick={ready ? undefined : onClick}
        onKeyDown={(event) => {
          if (ready) return;
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            onClick();
          }
        }}
        className={`report-card-scroll mt-auto overflow-y-auto rounded-2xl border shadow-[var(--shadow-inset)] transition ${
          compact ? "h-[115px] max-h-[115px] p-5" : "min-h-28 max-h-52 p-3 sm:min-h-32"
        } ${
          ready ? "border-success/35 bg-success/5" : "border-white/40 bg-muted/35"
        } ${ready ? "" : "cursor-pointer hover:border-primary/35 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary"}`}
      >
        {children}
      </div>
    </Card>
  );
}


export function EmptyHint({ icon, label }: { icon: React.ReactNode; label: string }) {
  return (
    <div className="flex min-h-16 flex-col items-center justify-center gap-2 text-center text-sm text-muted-foreground">
      <span className="text-primary">{icon}</span>
      {label}
    </div>
  );
}


export function FilePill({
  name,
  onRemove,
  color,
}: {
  name: string;
  onRemove: () => void;
  color: "primary" | "accent";
}) {
  return (
    <div
      className={`flex min-w-0 items-center justify-between gap-2 rounded-xl border px-3 py-2 text-sm shadow-[var(--shadow-inset)] ${
        color === "primary"
          ? "border-primary/25 bg-primary/5 text-foreground"
          : "border-accent/40 bg-accent/30 text-accent-foreground"
      }`}
    >
      <span className="min-w-0 truncate">{name}</span>
      <button
        type="button"
        onClick={(event) => {
          event.stopPropagation();
          onRemove();
        }}
        className="rounded-lg p-0.5 text-muted-foreground transition-colors hover:bg-background hover:text-destructive"
        aria-label={`Remove ${name}`}
      >
        <X className="h-3.5 w-3.5" />
      </button>
    </div>
  );
}

