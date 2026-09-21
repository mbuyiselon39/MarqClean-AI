type LogoProps = {
  markOnly?: boolean;
  className?: string;
  label?: string;
};

function Mark({ className }: { className?: string }) {
  return (
    <svg
      className={className}
      width="32"
      height="32"
      viewBox="0 0 32 32"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden="true"
      focusable="false"
    >
      <path
        d="M20 5H11a6 6 0 0 0-6 6v10a6 6 0 0 0 6 6h10a6 6 0 0 0 6-6v-5"
        stroke="currentColor"
        strokeWidth="2.25"
        strokeLinecap="butt"
      />
      <rect x="15" y="11" width="8" height="8" rx="1.5" fill="rgb(var(--accent))" />
    </svg>
  );
}

export default function Logo({ markOnly = false, className = "", label = "MarqClean AI" }: LogoProps) {
  if (markOnly) {
    return (
      <span className={className} aria-label={label} role="img">
        <Mark className="block h-8 w-8" />
      </span>
    );
  }

  return (
    <span
      className={`mc-logo inline-flex min-h-5 items-center text-ink tracking-[-0.02em] ${className}`}
      aria-label={label}
      role="img"
    >
      <Mark className="block h-8 w-8 shrink-0" />
      <span className="ml-4 flex items-baseline whitespace-nowrap text-[15px] leading-5">
        <span className="font-medium text-ink">Marq</span>
        <span className="font-normal text-ink-2">Clean AI</span>
      </span>
    </span>
  );
}
