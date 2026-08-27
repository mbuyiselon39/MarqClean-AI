import React from "react";
import AcademyHub from "../academy/AcademyHub";

interface ExcelAcademyPageProps {
  onExit: () => void;
}

export const ExcelAcademyPage: React.FC<ExcelAcademyPageProps> = ({ onExit }) => {
  return (
    <div className="min-h-screen bg-[#030611] py-8 px-4 sm:px-6 lg:px-8">
      <div className="mx-auto max-w-7xl">
        <AcademyHub onExit={onExit} />
      </div>
    </div>
  );
};

export default ExcelAcademyPage;
