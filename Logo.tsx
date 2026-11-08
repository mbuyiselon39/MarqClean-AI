type LogoProps = { markOnly?: boolean; className?: string; label?: string };

function Mark({ className }: { className?: string }) {
  return (
    <svg className={className} width="40" height="40" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
      <rect x="1" y="1" width="38" height="38" rx="11" fill="#173D32" />
      <path d="M9 27V13.5L19.5 23.2L30 13.5V27" stroke="#F5FBF6" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M19.5 23.2V29" stroke="#A9D96D" strokeWidth="2.6" strokeLinecap="round" />
      <circle cx="9" cy="12" r="2.1" fill="#A9D96D" />
      <circle cx="30" cy="12" r="2.1" fill="#A9D96D" />
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
