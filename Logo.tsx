type LogoProps = {
  markOnly?: boolean;
  className?: string;
  label?: string;
};

function Mark({ className }: { className?: string }) {
  return (
    <svg className={className} width="36" height="36" viewBox="0 0 36 36" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
      <rect width="36" height="36" rx="10" fill="rgb(var(--accent))" />
      <path d="M8 24.5V11.5L18 21.5L28 11.5V24.5" stroke="rgb(var(--accent-ink))" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" />
      <circle cx="8" cy="11.5" r="2" fill="rgb(var(--accent-ink))" />
      <circle cx="28" cy="11.5" r="2" fill="rgb(var(--accent-ink))" />
    </svg>
  );
}

export default function Logo({ markOnly = false, className = "", label = "MarqClean AI" }: LogoProps) {
  if (markOnly) return <span className={className} aria-label={label} role="img"><Mark className="block h-8 w-8" /></span>;
  return <span className={`mc-logo inline-flex items-center text-ink ${className}`} aria-label={label} role="img">
    <Mark className="block h-9 w-9 shrink-0" />
    <span className="ml-3 flex items-baseline whitespace-nowrap text-base leading-5"><span className="font-bold text-ink">marq</span><span className="font-medium text-ink-2">clean<span className="text-accent">.ai</span></span></span>
  </span>;
}
