import React from "react";

interface MarqLogoProps {
  size?: "sm" | "md" | "lg" | "xl";
  showText?: boolean;
  showTagline?: boolean;
  className?: string;
  glow?: boolean;
}

export const MarqLogo: React.FC<MarqLogoProps> = ({
  size = "md",
  showText = true,
  showTagline = false,
  className = "",
  glow = true,
}) => {
  const sizeMap = {
    sm: { icon: 32, text: "text-lg", tag: "text-[9px]" },
    md: { icon: 42, text: "text-xl", tag: "text-[10px]" },
    lg: { icon: 54, text: "text-2xl", tag: "text-xs" },
    xl: { icon: 72, text: "text-4xl", tag: "text-sm" },
  };

  const currentSize = sizeMap[size];

  return (
    <div className={`inline-flex items-center gap-3 select-none ${className}`}>
      {/* 3D Crystal M Mark matching MarqC 2 LOGO 2 */}
      <div className="relative flex-shrink-0" style={{ width: currentSize.icon, height: currentSize.icon }}>
        {glow && (
          <div
            className="absolute inset-0 rounded-xl bg-cyan-400/25 blur-lg animate-pulse"
            style={{ transform: "scale(1.15)" }}
          />
        )}
        <svg
          viewBox="0 0 100 100"
          className="relative w-full h-full drop-shadow-[0_4px_12px_rgba(0,210,255,0.45)]"
          fill="none"
          xmlns="http://www.w3.org/2000/svg"
        >
          <defs>
            <linearGradient id="crystalGrad1" x1="0%" y1="0%" x2="100%" y2="100%">
              <stop offset="0%" stopColor="#00F0FF" />
              <stop offset="50%" stopColor="#0099FF" />
              <stop offset="100%" stopColor="#003DB8" />
            </linearGradient>
            <linearGradient id="crystalGrad2" x1="100%" y1="0%" x2="0%" y2="100%">
              <stop offset="0%" stopColor="#C4FAFF" />
              <stop offset="45%" stopColor="#00D2FF" />
              <stop offset="100%" stopColor="#005BFF" />
            </linearGradient>
            <linearGradient id="facetHighlight" x1="0%" y1="0%" x2="100%" y2="50%">
              <stop offset="0%" stopColor="#FFFFFF" stopOpacity="0.95" />
              <stop offset="100%" stopColor="#00F0FF" stopOpacity="0.6" />
            </linearGradient>
            <linearGradient id="innerGlow" x1="50%" y1="0%" x2="50%" y2="100%">
              <stop offset="0%" stopColor="#00F0FF" />
              <stop offset="100%" stopColor="#00388A" />
            </linearGradient>
          </defs>

          {/* Outer Crystal Facets */}
          <path
            d="M12 78 L12 26 L36 8 L50 24 L64 8 L88 26 L88 78 L72 64 L72 34 L50 50 L28 34 L28 64 Z"
            fill="url(#crystalGrad1)"
          />

          {/* Upper Facets */}
          <path d="M36 8 L50 24 L28 34 Z" fill="url(#facetHighlight)" opacity="0.9" />
          <path d="M64 8 L72 34 L50 24 Z" fill="url(#facetHighlight)" opacity="0.95" />
          
          {/* Side Reflective Crystal Planes */}
          <path d="M12 26 L36 8 L28 34 L12 48 Z" fill="url(#crystalGrad2)" opacity="0.9" />
          <path d="M88 26 L88 48 L72 34 L64 8 Z" fill="url(#crystalGrad2)" opacity="0.9" />
          
          {/* Inner V Planes */}
          <path d="M50 24 L50 50 L28 34 Z" fill="#004CBD" opacity="0.85" />
          <path d="M50 24 L72 34 L50 50 Z" fill="#0066E0" opacity="0.85" />

          {/* Bottom Crystal Prisms */}
          <path d="M12 48 L28 64 L12 78 Z" fill="#002A7A" />
          <path d="M88 48 L88 78 L72 64 Z" fill="#002A7A" />
          <path d="M28 64 L50 82 L50 50 Z" fill="url(#innerGlow)" opacity="0.95" />
          <path d="M72 64 L50 50 L50 82 Z" fill="url(#crystalGrad2)" opacity="0.95" />

          {/* Glowing Digital Circuit Nodes inside */}
          <circle cx="34" cy="48" r="2.2" fill="#FFFFFF" />
          <circle cx="66" cy="48" r="2.2" fill="#FFFFFF" />
          <circle cx="50" cy="64" r="2.8" fill="#FFFFFF" />
          <line x1="34" y1="48" x2="50" y2="64" stroke="#A7F3D0" strokeWidth="1.2" strokeDasharray="2 1" />
          <line x1="66" y1="48" x2="50" y2="64" stroke="#A7F3D0" strokeWidth="1.2" strokeDasharray="2 1" />
          <line x1="50" y1="64" x2="50" y2="78" stroke="#00F0FF" strokeWidth="1.5" />
        </svg>
      </div>

      {/* Typography */}
      {showText && (
        <div className="flex flex-col">
          <div className="flex items-center tracking-tight font-extrabold text-white">
            <span className={`${currentSize.text} tracking-wider font-sans font-black bg-gradient-to-r from-white via-slate-100 to-cyan-200 bg-clip-text text-transparent`}>
              MAR
            </span>
            <span className={`${currentSize.text} tracking-wider font-sans font-black text-cyan-400 drop-shadow-[0_0_10px_rgba(0,210,255,0.6)]`}>
              Q
            </span>
            <span className={`${currentSize.text} tracking-wider font-sans font-black bg-gradient-to-r from-slate-100 to-white bg-clip-text text-transparent`}>
              CLEAN
            </span>
            <span className={`ml-1.5 px-1.5 py-0.5 rounded bg-gradient-to-r from-cyan-500 to-blue-600 text-white font-mono text-[10px] font-bold tracking-widest uppercase shadow-sm`}>
              AI
            </span>
          </div>

          {showTagline && (
            <span className={`${currentSize.tag} font-mono uppercase tracking-[0.22em] text-cyan-400/90 font-medium`}>
              TRANSFORM CHAOS INTO INTELLIGENCE
            </span>
          )}
        </div>
      )}
    </div>
  );
};

export default MarqLogo;
