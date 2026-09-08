import { AnimatePresence, motion } from "framer-motion";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import Papa from "papaparse";
import { lazy, Suspense, useEffect, useMemo, useRef, useState } from "react";

const ReconciliationHub = lazy(() => import("./ReconciliationHub"));
const DataToolbox = lazy(() => import("./DataToolbox"));
const ClientFunds = lazy(() => import("./ClientFunds"));
const AcademyHub = lazy(() => import("./AcademyHub"));
const HeroCarousel = lazy(() => import("./HeroCarousel"));