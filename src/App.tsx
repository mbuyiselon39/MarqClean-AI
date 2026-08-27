import { useState, useEffect } from "react";
import { Navbar, NavPageKey } from "./components/Navbar";
import { Footer } from "./components/Footer";
import { AuditModal } from "./components/AuditModal";
import { HomePage } from "./pages/HomePage";
import { ProductsPage } from "./pages/ProductsPage";
import { FeaturesPage } from "./pages/FeaturesPage";
import { FreeDataToolsPage } from "./pages/FreeDataToolsPage";
import { ExcelAcademyPage } from "./pages/ExcelAcademyPage";

export default function App() {
  const [currentPage, setCurrentPage] = useState<NavPageKey>("home");
  const [selectedSubTab, setSelectedSubTab] = useState<string | undefined>(undefined);
  const [isAuditModalOpen, setIsAuditModalOpen] = useState(false);
  const [auditEmail, setAuditEmail] = useState("");

  // Sync hash routing if present
  useEffect(() => {
    const handleHash = () => {
      const hash = window.location.hash.replace("#/", "").replace("#", "");
      if (hash === "products") {
        setCurrentPage("products");
      } else if (hash === "features") {
        setCurrentPage("features");
      } else if (hash === "free-data-tools") {
        setCurrentPage("free-data-tools");
      } else if (hash === "excel-academy") {
        setCurrentPage("excel-academy");
      } else if (hash === "contact" || hash === "about") {
        setCurrentPage(hash as NavPageKey);
      } else {
        setCurrentPage("home");
      }
    };

    handleHash();
    window.addEventListener("hashchange", handleHash);
    return () => window.removeEventListener("hashchange", handleHash);
  }, []);

  const handleNavigate = (page: NavPageKey, subTab?: string) => {
    setCurrentPage(page);
    setSelectedSubTab(subTab);
    window.location.hash = `#/${page}`;
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const handleRequestAudit = (email?: string) => {
    if (email) {
      setAuditEmail(email);
    }
    setIsAuditModalOpen(true);
  };

  return (
    <div className="min-h-screen bg-[#030611] font-sans text-slate-100 antialiased selection:bg-cyan-500/30 selection:text-cyan-200">
      {/* Universal Navigation Header */}
      <Navbar
        currentPage={currentPage}
        onNavigate={handleNavigate}
        onRequestAudit={() => handleRequestAudit()}
      />

      {/* Main Page View Router */}
      <main id="main-content" className="relative">
        {currentPage === "home" && (
          <HomePage
            onNavigate={handleNavigate}
            onRequestAudit={handleRequestAudit}
          />
        )}

        {currentPage === "products" && (
          <ProductsPage
            initialSubTab={selectedSubTab}
            onRequestAudit={() => handleRequestAudit()}
          />
        )}

        {currentPage === "features" && (
          <FeaturesPage
            initialSubTab={selectedSubTab}
            onRequestAudit={() => handleRequestAudit()}
          />
        )}

        {currentPage === "free-data-tools" && (
          <FreeDataToolsPage
            onRequestAudit={() => handleRequestAudit()}
          />
        )}

        {currentPage === "excel-academy" && (
          <ExcelAcademyPage
            onExit={() => handleNavigate("home")}
          />
        )}

        {(currentPage === "contact" || currentPage === "about") && (
          <div className="min-h-[70vh] py-20 px-4 sm:px-6 lg:px-8 max-w-4xl mx-auto">
            <div className="rounded-3xl border border-cyan-500/20 bg-slate-950/80 p-8 sm:p-12 text-center backdrop-blur-xl">
              <span className="text-xs font-bold uppercase tracking-widest text-cyan-400">
                Vertex Stream Group &bull; Corporate Information
              </span>
              <h1 className="mt-3 text-3xl font-extrabold text-white sm:text-4xl">
                {currentPage === "about" ? "About MarqClean AI" : "Contact Enterprise Support"}
              </h1>
              <p className="mx-auto mt-4 max-w-2xl text-base leading-relaxed text-slate-300 text-center">
                MarqClean AI provides high-precision automated data sanitation, reconciliation, and compliance infrastructure for financial analysts, auditors, and enterprise operations globally.
              </p>

              <div className="mt-8 grid gap-4 sm:grid-cols-2 text-left font-mono text-xs">
                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <span className="text-slate-400 uppercase text-[10px]">Headquarters</span>
                  <p className="mt-1 font-bold text-white">Johannesburg, Gauteng, South Africa</p>
                  <p className="text-slate-400 mt-1">Vertex Stream Group (Pty) Ltd</p>
                </div>

                <div className="rounded-2xl border border-slate-800 bg-slate-900/60 p-5">
                  <span className="text-slate-400 uppercase text-[10px]">Official Website &amp; Host</span>
                  <p className="mt-1 font-bold text-cyan-400">https://marqcleanai.vertexsg.co.za/</p>
                  <p className="text-slate-400 mt-1">support@vertexstreamtechnologies.com</p>
                </div>
              </div>

              <div className="mt-8 flex justify-center">
                <button
                  onClick={() => handleRequestAudit()}
                  className="rounded-2xl bg-gradient-to-r from-cyan-500 to-blue-600 px-6 py-3 text-xs font-bold uppercase text-slate-950 hover:brightness-110 transition shadow-lg shadow-cyan-500/25"
                >
                  Request Technical Audit
                </button>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Universal Footer */}
      <Footer
        onNavigate={handleNavigate}
        onRequestAudit={() => handleRequestAudit()}
      />

      {/* Interactive Audit Modal */}
      <AuditModal
        isOpen={isAuditModalOpen}
        onClose={() => setIsAuditModalOpen(false)}
        initialEmail={auditEmail}
      />
    </div>
  );
}
