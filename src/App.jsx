import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutGrid, Users, Receipt, FileWarning, Landmark, Building2,
  ClipboardCheck, Search, ChevronRight, X, Check, AlertTriangle,
  Clock, TrendingUp, UserCircle2, Plus, Stamp, ChevronDown,
  Filter, ArrowUpRight, CircleDot, Loader2, RefreshCw, History,
  ChevronUp, CalendarDays, CalendarRange, Settings2, Trash2,
  Pencil, ChevronLeft, ShieldCheck, Home, LogOut, Mail, Lock, UserRound,
  Phone, Briefcase, UserCheck, Wallet, ShieldAlert, Menu, Bell, Clock3, ArrowLeft, ExternalLink,
  Eye, EyeOff, Copy, KeyRound, Download, MapPin, Contact, Scale,
  CheckCircle2, XCircle
} from "lucide-react";
import { supabase } from "./supabaseClient";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";
import { fetchTasks, createTask, updateTask, completeTask, deleteTask, subscribeTasks } from "./services/tasks";
import { logActivity, activityMessages } from "./services/activity";
import { bucketize as bucketizeDeadlines, BUCKET_LABELS as DEADLINE_BUCKET_LABELS } from "./services/deadlines";
import { TASK_STATUTS, TASK_STATUT_BY_CODE, TASK_PRIORITES, TASK_PRIORITE_BY_CODE, taskSortWeight, PILOTAGE_COLORS } from "./constants/pilotage";
import { detectAllAnomalies } from "./services/anomalies";

const T = {
  paper: "#F3F4F6", paperDeep: "#EEF2FF", ink: "#0F172A", inkSoft: "#475569", inkMuted: "#94A3B8",
  line: "#E2E8F0", card: "#FFFFFF", gold: "#D97706", goldSoft: "#FEF3C7",
  green: "#16A34A", greenSoft: "#DCFCE7", red: "#DC2626", redSoft: "#FEE2E2",
  amber: "#D97706", amberSoft: "#FEF3C7",
  /* accent sobre façon Kabineo (indigo) */
  navy: "#2563EB", navySoft: "#EFF6FF",
  /* sidebar sombre façon "slate/ardoise" : fond bleu-nuit très foncé, textes clairs */
  sidebarBg: "#0F172A", sidebarBg2: "#1E293B", sidebarInk: "#E2E8F0", sidebarInkMuted: "#94A3B8",
  sidebarActive: "rgba(99,102,241,0.22)", sidebarBorder: "#1E293B", sidebarAccent: "#60A5FA",
  serif: "'Plus Jakarta Sans', 'Inter', -apple-system, sans-serif", mono: "'JetBrains Mono', ui-monospace, monospace", sans: "'Inter', -apple-system, sans-serif",
  shadow: "0 1px 2px rgba(15,23,42,0.04), 0 8px 20px -6px rgba(15,23,42,0.08)",
  shadowSm: "0 1px 2px rgba(15,23,42,0.05), 0 1px 3px rgba(15,23,42,0.06)",
  shadowLg: "0 24px 48px -14px rgba(15,23,42,0.16)",
  radius: 16, radiusSm: 12, radiusLg: 20,
};

/* Regroupement des pièces du Dossier Permanent : source unique utilisée par la progression de l'accueil et les compteurs. */
const MISSION_GROUPS = [
  { title: "Identité & statuts", keys: ["KBIS", "Statuts", "CNI dirigeants", "CNI associés"] },
  { title: "Cadrage de la mission", keys: ["Notes entrée mission / Devizen", "Acceptation mission", "LM à jour"] },
  { title: "Conformité & suivi", keys: ["LAB / Kanta / Devizen à jour"] },
  { title: "Clôture du dossier", keys: ["Fiche client", "Bouclage"] },
];
const MISSION_ALL_KEYS = MISSION_GROUPS.flatMap((g) => g.keys);

/* ============================================================
   SEED DATA — extrait du fichier Excel de suivi du cabinet
   ============================================================ */
const RAW_SEED_CLIENTS = [{"nom":"A&D RESTOS","siren":"81276334","logiciel":"MYUNISOFT","collab":"Cheikh","tvaRegime":"CA12","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"AC INVEST","siren":"925320210","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"AD SOLUTION","siren":"942467515","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"AE BAT","siren":"931778112","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ALLO SOS MOTO","siren":"488698960","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":19,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ALPHA DIGITAL","siren":"920603560","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"AMS PLOMBERIE","siren":"917541906","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"APEL","siren":"326627247","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":19,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ARSA","siren":"954016481","logiciel":"QUADRA","collab":"Jacques","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ATELIER GOURMAND","siren":"84531968","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ATI INVEST","siren":"927855247","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"AU COIN DU PAIN","siren":"993062835","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BACKSTAGE BEAUTY GROUP","siren":"853273522","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BANGLA.COM","siren":"804359362","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BELARBI ABDALLAH","siren":"530122589","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BELARBI ABDELKAOUI","siren":"431443852","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BENGAL COIFFURE","siren":"830473351","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BHUVI BEAUTE","siren":"851770354","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BLAST","siren":"831893698","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BLUE SECURITY","siren":"929357168","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"BOUCHERIE SERAS","siren":"930591714","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"TRIM","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"CAFFE ITALIA","siren":"790500912","logiciel":"MYUNISOFT","collab":"Jacques","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"CELIA","siren":"923210215","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"CENTRAL'AUTO","siren":"891458457","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"CHARLOTTE FRANCISCO","siren":"819855727","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"CHATSN TRANSPORT BILAN 2025","siren":"789814399","logiciel":"QUADRA","collab":"Jacques","tvaRegime":"CA3","tvaExig":20,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"DAKAROIS KITCHEN","siren":"981110026","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"TRIM","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"DAVIDSEN","siren":"849091400","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"DESTOCK PIECES AUTO LE PERREUX","siren":"811969153","logiciel":"","collab":"Cheikh","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"DIALLO KAMION","siren":"809583669","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"DIAMOND SUSHI","siren":"890451271","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ECO NOISY TRANSPORT BILAN 2025","siren":"850096587","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ENERGIA","siren":"..","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ERATOMBE","siren":"849495742","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"EVE SECURITY","siren":"504487216","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":20,"tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"EXTERNALIS-CORPORATION","siren":"853414910","logiciel":"","collab":"Cheikh","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"FAIM DE SEMAINE","siren":"898044110","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"FBA BAT","siren":"538695313","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"FOURNIL JEAN XXIII","siren":"893192138","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"GALLERY BOUNAN CH","siren":"803649110","logiciel":"QUADRA","collab":"Jacques","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"GROUPE PNS","siren":"791864317","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"GTM","siren":"933355679","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"GUT HEALTH & WELLNESS","siren":"912855350","logiciel":"MYUNISOFT","collab":"Cheikh","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"HOLDING RAZA","siren":"939739967","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"I PRO BATIMENT","siren":"931787204","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"INIT SERVICES","siren":"941381568","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"KAILEY RENOVATION","siren":"842455537","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"KD","siren":"884824566","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"KHOUJABAT","siren":"894352939","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LA BONNE EPOQUE","siren":"833393598","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LE CENTRE MEDICAL DE VERDUN","siren":"897427761","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LE JUSTE PRIX","siren":"934480039","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LE PETIT MARCHE","siren":"837752062","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LEMNISCATE SOFTWARE","siren":"830376653","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"L'EPI D'OR","siren":"984628628","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LMG TRANSPORT","siren":"809937717","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"LO NAFI","siren":"904900065","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MAC CHICKEN","siren":"843437724","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"TRIM","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"Madame DIALLO KAMION","siren":"809583669","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MB LAUNDRY","siren":"940912835","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MED FOOD","siren":"951966795","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MEYO","siren":"983592122","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MFC DISTRIBUTION","siren":"989941216","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MIX TRAVAUX IDF","siren":"840772172","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MS TRANS (MTRANS SERVICES","siren":"893029538","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MTP HOLDING","siren":"978063147","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"MY GLOBAL","siren":"922568092","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"NETEN","siren":"511556193","logiciel":"QUADRA","collab":"Soli","tvaRegime":"TRIM","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"NEW STAR HOLDING INTERNATIONAL","siren":"907566871","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"NIA CONSEILS BILAN 2025","siren":"914568399","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"NICK SERVICES","siren":"512395823","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"NIRALI","siren":"819586017","logiciel":"QUADRA","collab":"Soli","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"NISHA ESTHETIQUE","siren":"880336128","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"O DELICE","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ON EST LA","siren":"929695120","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"OPTIMUS TECHNOLOGIES","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"OZONE HYGIENE ENVIRONNEMENT","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"PARIS CASH AND CARRY","siren":"803981943","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"PERI ALIMENTATION","siren":"797918489","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"PLACE COLETTE","siren":"928971233","logiciel":"MYUNISOFT","collab":"Cheikh","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"POWERFIT","siren":"83083017","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"RAMY 37","siren":"884284324","logiciel":"QUADRA","collab":"Jacques","tvaRegime":"CA3","tvaExig":24,"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"RED LIFE FRANCE","siren":"921174686","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"RED LIFE HOLDING","siren":"919332106","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"REPAIR MASTER","siren":"103229209","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"REPON SERGE AGRICULTURE","siren":"348050287","logiciel":"QUADRA","collab":"Cheikh","tvaRegime":"TRIM","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SABP","siren":"..","logiciel":"","collab":"Cheikh","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SAINT AMBROISE SAS","siren":"920863982","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SAISANTHU SOCIETE DE NETTOYAGE","siren":"940616766","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SCI KAMY TEAM","siren":"883778268","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SCI LES MANOS","siren":"499319796","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"sci p immo","siren":"825254378","logiciel":"QUADRA","collab":"Soli","tvaRegime":"TRIM","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SCI SHAANA","siren":"940587819","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SHAH JALAL 76","siren":"920826591","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SHIV-SAI","siren":"534331368","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":19,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SKYTECH","siren":"922167713","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SPEKCOM","siren":"788795631","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":20,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"STARTED FROM THE BOTTOM FACILITY SERVICES","siren":"832559884","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"STEEL PAINT","siren":"530937754","logiciel":"QUADRA","collab":"Cheikh","tvaRegime":"CA3","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"SUSHI KID","siren":"793336025","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"système automatique et securité","siren":"838588929","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"TIROUCHE Sofiane","siren":"812371276","logiciel":"QUADRA","collab":"Soli","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"tms concept","siren":"883337503","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"TRANS TP IDF","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"TRANSPORT FRET LOGISTIQUE","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"TSHI INVEST","siren":"927452532","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"WANDEE","siren":"833083546","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}},{"nom":"ZIANIDES","siren":"844162396","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"KBIS":false,"Statuts":false,"CNI dirigeants":false,"CNI associés":false,"Notes entrée mission / Devizen":false,"Acceptation mission":false,"LM à jour":false,"LAB / Kanta / Devizen à jour":false,"Bouclage":false,"Fiche client":false}}]
;

const PALETTE = ["#6366F1", "#10B981", "#F97316", "#EC4899", "#06B6D4", "#8B5CF6", "#F59E0B", "#14B8A6", "#F43F5E", "#3B82F6"];
const DEFAULT_TEAM = ["Cheikh", "Soli", "Emilie", "Jacques"].map((nom, i) => ({
  id: `seed-${i}`, nom, color: PALETTE[i % PALETTE.length],
}));

const MOIS_ORDER = ["Jan","Fév","Mar","Avr","Mai","Juin","Juil","Août","Sept","Oct","Nov","Déc"];
const MOIS_FULL = { Jan:"Janvier",Fév:"Février",Mar:"Mars",Avr:"Avril",Mai:"Mai",Juin:"Juin",Juil:"Juillet",Août:"Août",Sept:"Septembre",Oct:"Octobre",Nov:"Novembre",Déc:"Décembre" };
const REGIMES_TVA = ["CA3", "CA12", "FRANCHISE", "FEB", "TRIM"];
// FEB et TRIM sont d'anciens codes de régime présents dans des dossiers historiques
// (Franchise en base / Trimestriel) : on les garde sélectionnables pour que leur fiche
// ne perde pas leur régime affiché, sans les pousser comme nouveaux choix recommandés.
const REGIMES_TVA_LABELS = { FEB: "FEB (franchise en base)", TRIM: "TRIM (ancien code trimestriel)" };
/* ============================================================
   SECTEURS D'ACTIVITÉ — classification auto depuis le champ "Activité"
   ============================================================ */
const SECTEURS_ACTIVITE = [
  { id: "restauration", label: "Restauration / CHR", color: "#DC2626" },
  { id: "batiment", label: "Bâtiment / BTP", color: "#D97706" },
  { id: "beaute", label: "Beauté / Coiffure", color: "#DB2777" },
  { id: "sante", label: "Santé", color: "#16A34A" },
  { id: "commerce_detail", label: "Commerce de détail", color: "#2563EB" },
  { id: "commerce_gros", label: "Commerce de gros", color: "#0EA5E9" },
  { id: "informatique", label: "Informatique", color: "#4F46E5" },
  { id: "securite_nettoyage", label: "Sécurité / Nettoyage", color: "#64748B" },
  { id: "transport", label: "Transport", color: "#EA580C" },
  { id: "immobilier", label: "Immobilier / Holding / Patrimoine", color: "#7C3AED" },
  { id: "conseil", label: "Conseil", color: "#0891B2" },
  { id: "automobile", label: "Automobile", color: "#B45309" },
  { id: "autres", label: "Autres", color: "#94A3B8" },
];

// Mots-clés (issus des libellés officiels NAF/APE) associés à chaque secteur.
// Pour ajouter une activité : repérez le secteur concerné et complétez le tableau "keywords".
const ACTIVITE_KEYWORDS = [
  { secteurId: "restauration", keywords: [
    "restauration rapide", "restauration traditionnelle", "restaurant",
    "boulangerie", "patisserie", "traiteur", "debit de boissons", "cafe restaurant", "snack",
  ]},
  { secteurId: "batiment", keywords: [
    "maconnerie", "peinture", "vitrerie", "revetement des sols", "revetement des murs",
    "plomberie", "installation eau", "installation gaz", "installation electrique",
    "travaux de finition", "construction de maisons individuelles", "construction maisons",
    "gros oeuvre", "second oeuvre", "menuiserie", "couverture", "charpente", "terrassement", "btp",
  ]},
  { secteurId: "beaute", keywords: [
    "soins de beaute", "coiffure", "esthetique", "institut de beaute", "onglerie", "barbier",
  ]},
  { secteurId: "sante", keywords: [
    "medecin", "generaliste", "specialiste", "infirmier", "infirmiere", "sage-femme", "sage femme",
    "dentiste", "kinesitherapeute", "pharmacie", "professions paramedicales",
  ]},
  { secteurId: "commerce_detail", keywords: [
    "commerce de detail", "alimentation generale", "boucherie", "charcuterie",
    "commerce de detail de viandes", "telecommunication", "marches", "magasin non specialise",
    "cordonnerie", "supermarche", "epicerie",
  ]},
  { secteurId: "commerce_gros", keywords: [
    "commerce de gros", "grossiste", "cafe the epices", "vaisselle", "verrerie",
    "produits alimentaires non specialises en gros",
  ]},
  { secteurId: "informatique", keywords: [
    "programmation informatique", "conseil en systemes informatiques", "activites informatiques",
    "reparation d ordinateurs", "reparation ordinateurs", "edition de logiciels",
    "developpement informatique",
  ]},
  { secteurId: "securite_nettoyage", keywords: [
    "securite privee", "activites de securite privee", "nettoyage courant des batiments",
    "nettoyage industriel", "desinfection", "deratisation", "gardiennage", "surveillance",
  ]},
  { secteurId: "transport", keywords: [
    "transport routier de fret", "fret de proximite", "transport urbain", "transport de voyageurs",
    "transport routier", "messagerie", "demenagement",
  ]},
  { secteurId: "immobilier", keywords: [
    "societes holding", "holding", "gestion de fonds", "sieges sociaux",
    "location de terrains", "location immobiliere", "agence immobiliere",
    "administration d immeubles", "supports juridiques de gestion de patrimoine mobilier",
    "fonds de placement",
  ]},
  { secteurId: "conseil", keywords: [
    "relations publiques", "communication", "conseil pour les affaires", "conseil de gestion",
    "conseil en management", "ingenierie", "etudes techniques",
  ]},
  { secteurId: "automobile", keywords: [
    "entretien et reparation de vehicules automobiles legers", "entretien reparation vehicules",
    "garage automobile", "carrosserie", "reparation automobile",
  ]},
  { secteurId: "autres", keywords: [
    "blanchisserie", "teinturerie", "formation continue", "formation professionnelle",
    "fabrication de produits alimentaires", "reproduction de plantes", "pepiniere", "horticulture",
  ]},
];
/* ============================================================
   CHECKLIST FORME JURIDIQUE — points annuels à ne pas oublier
   ============================================================ */
const FORME_JURIDIQUE_CHECKLIST_ITEMS = {
  "SARL": [
    { id: "gerantRemuneration", label: "Rémunération gérant majoritaire déclarée (cotisations TNS)" },
    { id: "reserveLegale", label: "Dotation réserve légale (5%/an jusqu'à 10% du capital)" },
    { id: "approbationComptes", label: "Approbation des comptes en AG (dans les 6 mois de la clôture)" },
    { id: "depotGreffe", label: "Dépôt de la liasse au greffe" },
  ],
  "EURL": [
    { id: "gerantRemuneration", label: "Rémunération gérant majoritaire déclarée (cotisations TNS)" },
    { id: "reserveLegale", label: "Dotation réserve légale (5%/an jusqu'à 10% du capital)" },
    { id: "approbationComptes", label: "Approbation des comptes (dans les 6 mois de la clôture)" },
    { id: "depotGreffe", label: "Dépôt de la liasse au greffe" },
  ],
  "SAS": [
    { id: "presidentAssimile", label: "Président assimilé-salarié (pas de cotisations TNS)" },
    { id: "dividendesPS", label: "Prélèvements sociaux sur dividendes vérifiés" },
    { id: "commissaireComptes", label: "Nomination CAC vérifiée si seuils dépassés" },
    { id: "depotGreffe", label: "Dépôt de la liasse au greffe" },
  ],
  "SASU": [
    { id: "presidentAssimile", label: "Président assimilé-salarié (pas de cotisations TNS)" },
    { id: "dividendesPS", label: "Prélèvements sociaux sur dividendes vérifiés" },
    { id: "depotGreffe", label: "Dépôt de la liasse au greffe" },
  ],
  "SCI": [
    { id: "regimeIRIS", label: "Régime IR (par défaut) ou option IS documentée" },
    { id: "coherenceHonoraires", label: "Cohérence avec la convention d'honoraires (si option IS)" },
    { id: "approbationComptes", label: "Approbation des comptes en AG" },
  ],
  "SCM": [
    { id: "quotePartCharges", label: "Répartition des charges entre associés conforme au contrat" },
    { id: "pasBenefice", label: "Absence de bénéfice généré par la SCM vérifiée" },
    { id: "approbationComptes", label: "Approbation des comptes en AG" },
  ],
  "SELARL": [
    { id: "assuranceRCP", label: "Assurance RCP (responsabilité civile professionnelle) à jour" },
    { id: "ordreProfessionnel", label: "Inscription à l'ordre professionnel vérifiée" },
    { id: "capitalDetention", label: "Répartition du capital conforme aux règles de la profession" },
    { id: "depotGreffe", label: "Dépôt de la liasse au greffe" },
  ],
  "EI": [
    { id: "patrimoineAffecte", label: "Patrimoine professionnel affecté distinct du personnel (statut 2022)" },
    { id: "optionIS", label: "Option IS (si formulée) correctement documentée" },
  ],
  "SA": [
    { id: "commissaireComptes", label: "Commissaire aux comptes désigné" },
    { id: "reserveLegale", label: "Dotation réserve légale (5%/an jusqu'à 10% du capital)" },
    { id: "approbationComptes", label: "Approbation des comptes en AG (dans les 6 mois de la clôture)" },
    { id: "depotGreffe", label: "Dépôt de la liasse au greffe" },
  ],
  "SNC": [
    { id: "responsabiliteIndefinie", label: "Associés informés de la responsabilité indéfinie et solidaire" },
    { id: "approbationComptes", label: "Approbation des comptes en AG" },
  ],
  "Association": [
    { id: "agAnnuelle", label: "Assemblée générale annuelle tenue" },
    { id: "budgetVote", label: "Budget voté / comptes approuvés" },
  ],
};
// Complément spécifique Holding / Immobilier (affiché en plus si le secteur d'activité correspond)
const HOLDING_CHECKLIST_ITEMS = [
  { id: "integrationFiscale", label: "Intégration fiscale éventuelle vérifiée" },
  { id: "regimeMereFille", label: "Régime mère-fille applicable vérifié" },
  { id: "conventionsTresorerie", label: "Conventions de trésorerie / management fees documentées" },
];
/* ============================================================
   ACTUALITÉS & AIDES PAR SECTEUR — contenu pré-rempli, éditable
   (Expert / Chef de mission / Admin), stocké dans Supabase.
   À vérifier et mettre à jour régulièrement : les dispositifs évoluent.
   ============================================================ */
const SEED_AIDES_SECTEUR = {
  restauration: {
    aides: ["Aides à la rénovation énergétique des commerces (dispositifs locaux/FISAC)", "Aide à l'embauche en apprentissage restauration"],
    obligations: ["Licence de débit de boissons (si vente d'alcool)", "Permis d'exploitation (formation obligatoire)", "HACCP — formation hygiène alimentaire obligatoire", "Affichage des allergènes"],
  },
  batiment: {
    aides: ["Qualification RGE pour les aides à la rénovation énergétique", "MaPrimeRénov' (accessible via un pro RGE)", "Éco-prêt à taux zéro"],
    obligations: ["Assurance décennale obligatoire", "Garantie de parfait achèvement", "Carte professionnelle du bâtiment selon l'activité"],
  },
  beaute: {
    aides: ["Aides à la formation continue coiffure/esthétique (OPCO)"],
    obligations: ["Diplôme ou certification obligatoire (coiffure/esthétique)", "Normes d'hygiène spécifiques (stérilisation du matériel)"],
  },
  sante: {
    aides: [],
    obligations: ["Inscription à l'Ordre professionnel", "Assurance RCP obligatoire", "Numéro RPPS/ADELI à jour"],
  },
  commerce_detail: {
    aides: [],
    obligations: ["Registre du commerce à jour", "Affichage des prix", "Licence de vente d'alcool si applicable"],
  },
  commerce_gros: {
    aides: [],
    obligations: ["Traçabilité des produits alimentaires le cas échéant", "Respect des normes d'hygiène en gros"],
  },
  informatique: {
    aides: ["Crédit d'impôt recherche (CIR) / innovation (CII) pour l'édition de logiciels", "Aides France Num pour la transformation numérique"],
    obligations: ["Conformité RGPD si traitement de données personnelles", "Déclaration CNIL si applicable"],
  },
  securite_nettoyage: {
    aides: [],
    obligations: ["Carte professionnelle CNAPS obligatoire (sécurité privée)", "Autorisation d'exercice CNAPS pour l'entreprise", "Habilitations spécifiques nettoyage industriel/désinfection"],
  },
  transport: {
    aides: ["Suramortissement véhicules propres"],
    obligations: ["Capacité de transport (licence)", "Inscription au registre des transporteurs", "Carte de conducteur"],
  },
  immobilier: {
    aides: [],
    obligations: ["Carte professionnelle transaction/gestion immobilière si agence", "Garantie financière et assurance RCP", "Registre des mandats"],
  },
  conseil: {
    aides: ["Aides au conseil stratégique (Bpifrance Diag)"],
    obligations: ["Assurance RCP recommandée"],
  },
  automobile: {
    aides: ["Aides à la formation véhicules électriques/hybrides"],
    obligations: ["Agrément contrôle technique si applicable", "Gestion réglementée des déchets automobiles (huiles, pneus)"],
  },
  autres: { aides: [], obligations: [] },
};
/* ============================================================
   ACTUALITÉS EN DIRECT PAR SECTEUR — Google Actualités RSS
   + flux officiel Service-Public.fr (via proxy rss2json, gratuit sans clé)
   ============================================================ */
const SECTEUR_NEWS_QUERIES = {
  restauration: "aide dispositif restauration CHR France",
  batiment: "aide RGE rénovation bâtiment BTP France",
  beaute: "aide dispositif coiffure esthétique entreprise France",
  sante: "actualité professionnels de santé libéral France",
  commerce_detail: "aide commerce de détail commerçants France",
  commerce_gros: "actualité commerce de gros France",
  informatique: "aide crédit impôt innovation entreprises informatique France",
  securite_nettoyage: "actualité sécurité privée CNAPS nettoyage entreprises France",
  transport: "aide transport routier entreprises France",
  immobilier: "actualité holding patrimoine SCI fiscalité France",
  conseil: "actualité conseil entreprises aide Bpifrance France",
  automobile: "aide garage automobile entretien réparation France",
  autres: "actualité aide entreprise PME France",
};
const OFFICIAL_PRO_FEED = "https://www.service-public.fr/abonnements/rss/actu-actu-pro.rss";
function buildGoogleNewsRssUrl(query) {
  return `https://news.google.com/rss/search?q=${encodeURIComponent(query)}&hl=fr&gl=FR&ceid=FR:fr`;
}
// Proxy CORS gratuit, sans clé, sans compte — renvoie le XML brut du flux
function buildProxyUrl(rssUrl) {
  return `https://api.allorigins.win/raw?url=${encodeURIComponent(rssUrl)}`;
}
async function fetchRssFeed(rssUrl, count = 6) {
  const res = await fetch(buildProxyUrl(rssUrl));
  if (!res.ok) throw new Error(`Erreur réseau (${res.status})`);
  const xmlText = await res.text();
  const xml = new DOMParser().parseFromString(xmlText, "text/xml");
  if (xml.querySelector("parsererror")) throw new Error("Flux XML illisible");
  const channelTitle = xml.querySelector("channel > title")?.textContent || "";
  const nodes = Array.from(xml.querySelectorAll("item")).slice(0, count);
  if (nodes.length === 0) throw new Error("Flux vide");
  return nodes.map((node) => ({
    title: node.querySelector("title")?.textContent || "",
    link: node.querySelector("link")?.textContent || "",
    date: node.querySelector("pubDate")?.textContent || "",
    source: channelTitle,
  }));
}
// Récupère les actus d'un secteur : Google Actualités (ciblé) + Service-Public.fr (officiel, générique)
async function fetchSecteurNews(secteurId) {
  const query = SECTEUR_NEWS_QUERIES[secteurId] || SECTEUR_NEWS_QUERIES.autres;
  const results = await Promise.allSettled([
    fetchRssFeed(buildGoogleNewsRssUrl(query), 5),
    fetchRssFeed(OFFICIAL_PRO_FEED, 3),
  ]);
  const items = results.flatMap((r) => (r.status === "fulfilled" ? r.value : []));
  items.sort((a, b) => new Date(b.date) - new Date(a.date));
  if (items.length === 0 && results.every((r) => r.status === "rejected")) {
    throw new Error("Tous les flux sont indisponibles pour le moment.");
  }
  return items.slice(0, 6);
}

async function loadSecteurContentFromSupabase() {
  const { data, error } = await supabase.from("secteur_content").select("secteur_id, aides, obligations, updated_at, updated_by");
  if (error) { console.error("Erreur chargement contenu secteurs :", error.message); return null; }
  const map = {};
  (data || []).forEach((row) => {
    map[row.secteur_id] = { aides: row.aides || [], obligations: row.obligations || [], updatedAt: row.updated_at, updatedBy: row.updated_by };
  });
  return map;
}
async function upsertSecteurContentRemote(secteurId, patch, updatedBy) {
  const { error } = await supabase.from("secteur_content").upsert({
    secteur_id: secteurId, ...patch, updated_at: new Date().toISOString(), updated_by: updatedBy || null,
  });
  if (error) console.error("Erreur sauvegarde contenu secteur :", error.message);
}
function getFormeJuridiqueItems(client) {
  const base = FORME_JURIDIQUE_CHECKLIST_ITEMS[client.formeJuridique] || [];
  const extra = client.secteur === "immobilier" ? HOLDING_CHECKLIST_ITEMS : [];
  return [...base, ...extra];
}

function normalizeText(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

const NAF_SECTORS = [
  [/^56|^55/, "restauration"], [/^41|^42|^43/, "batiment"], [/^96\.02/, "beaute"], [/^86/, "sante"], [/^47/, "commerce_detail"], [/^46/, "commerce_gros"], [/^62|^63/, "informatique"], [/^80|^81/, "securite_nettoyage"], [/^49|^50|^51|^52|^53/, "transport"], [/^68|^64\.2/, "immobilier"], [/^69|^70|^71|^73|^74/, "conseil"], [/^45/, "automobile"]
];
function classifyNaf(code) { const c=String(code||"").replace(/\s/g,"").replace(/([0-9]{2})([A-Z])/,"$1.$2"); return (NAF_SECTORS.find(([r])=>r.test(c))||[])[1] || "autres"; }
function classifyActivite(activiteText) {
  const norm = normalizeText(activiteText);
  if (!norm) return "";
  for (const entry of ACTIVITE_KEYWORDS) {
    for (const kw of entry.keywords) {
      if (norm.includes(normalizeText(kw))) return entry.secteurId;
    }
  }
  return "autres";
}
const QUARTER_END_MONTHS = ["Mar", "Juin", "Sept", "Déc"]; // fins de trimestre civil, pour la périodicité CA3 trimestrielle
const TVA_PERIODICITES = ["mensuelle", "trimestrielle"];
const TVA_PERIODICITE_LABELS = { mensuelle: "Mensuelle", trimestrielle: "Trimestrielle" };

function currentMonthKey() { return MOIS_ORDER[new Date().getMonth()]; }
function previousMonthKey() { return MOIS_ORDER[(new Date().getMonth() + 11) % 12]; }
function todayISO() { return new Date().toISOString().slice(0, 10); }
function addYearISO(iso, years = 1) {
  if (!iso || typeof iso !== "string") return iso;
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return iso;
  return `${String(y + years).padStart(4, "0")}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}
function fmtFR(iso) {
  if (!iso || typeof iso !== "string") return "—";
  const [y, m, d] = iso.split("-");
  if (!y || !m || !d) return iso;
  return `${d}/${m}/${y}`;
}
function addMonthsISO(iso, months) {
  if (!iso || typeof iso !== "string") return null;
  const [y, m, d] = iso.split("-").map(Number);
  if (!y || !m || !d) return null;
  const dt = new Date(y, m - 1 + months, d);
  return dt.toISOString().slice(0, 10);
}
function fmtEUR(v) {
  const n = Number(v);
  if (v === "" || v == null || Number.isNaN(n)) return "—";
  return n.toLocaleString("fr-FR", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
}
// Échéance légale approximative de dépôt : clôture + 3 mois (déjà la règle utilisée pour les échéances fiscales)
function getBilanEcheance(dateCloture) {
  return dateCloture ? addMonthsISO(dateCloture, 3) : null;
}
function getBilanStatut(b, dateCloture) {
  const echeance = getBilanEcheance(dateCloture);
  const enRetard = echeance && !b.transmis && todayISO() > echeance;
  if (b.transmis) return { label: "Transmis", tone: "green" };
  if (enRetard) return { label: "En retard", tone: "red" };
  if (b.valideClient) return { label: "Validé client", tone: "purple" };
  if (b.revision === "terminee") return { label: "Révision terminée", tone: "purple" };
  if (b.revision === "en_cours") return { label: "En cours", tone: "amber" };
  return { label: "À faire", tone: "neutral" };
}
const BILAN_REVISION_STEPS = [
  { id: "a_faire", label: "À faire" },
  { id: "en_cours", label: "En cours" },
  { id: "terminee", label: "Terminée" },
];
function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}
function startOfWeek(d) {
  const x = new Date(d);
  const day = (x.getDay() + 6) % 7; // Monday = 0
  x.setDate(x.getDate() - day);
  x.setHours(0, 0, 0, 0);
  return x;
}

/* ============================================================
   MIGRATION — assure la compatibilité avec les anciennes données
   ============================================================ */
function migrateClients(list) {
  const year = new Date().getFullYear();
  return (list || []).map((c, i) => {
    const next = { ...c };
    if (!next.id) next.id = next.siren ? `siren-${next.siren}` : `c-${i}-${next.nom || "x"}`;
    if (!next.portefeuilleId) next.portefeuilleId = "axe"; // valeur par défaut pour les données d'origine (avant multi-cabinets)
   if (!next.statutDossier) next.statutDossier = "actif"; // Actif / Inactif
    if (next.tvaRegime === "CA3" && !next.tvaPeriodicite) next.tvaPeriodicite = "mensuelle"; // valeur par défaut = comportement historique
    if (!next.tvaControle) next.tvaControle = {}; // { [mois]: { commentaire, par, date } } — remarques du contrôle TVA (cf. écran TVA)
    if (!next.expert) next.expert = "";
    if (!next.chefMission) next.chefMission = "";
    if (!next.dateCloture) next.dateCloture = "";
    if (!next.formeJuridique) next.formeJuridique = "";
    if (!next.regimeFiscal) next.regimeFiscal = "";
    if (!next.capital) next.capital = "";
    if (!next.activite) next.activite = "";
    if (!next.contact) next.contact = { telephone: "", email: "", adresse: "", codePostal: "", ville: "", contactNom: "", contactFonction: "" };
    if (next.secteurManuel == null) next.secteurManuel = false;
    if (!next.secteur) next.secteur = classifyNaf(next.codeNaf);
    if (!next.formeJuridiqueHistory) next.formeJuridiqueHistory = {};
    if (!next.lienSharepoint) next.lienSharepoint = "";
    if (!next.revision) next.revision = {};
    if (!next.missionsExceptionnelles) next.missionsExceptionnelles = [];
    next.missionsExceptionnelles = next.missionsExceptionnelles.map((m) => (
      m.statut === "livree" ? { ...m, statut: "valide" } : m
    ));
    if (!next.corporate) {
      next.corporate = {
        kyc: { lab: false, mandat: false, choixPA: "", beneficiaireEffectif: false, beneficiaireNom: "" },
        kycExtra: [], notes: "",
      };
    }
    if (!next.ageAgoHistory) {
      next.ageAgoHistory = {};
      if (next.ageAgo && Object.keys(next.ageAgo).length) {
        next.ageAgoHistory[year] = {
          ago: !!next.ageAgo.ago, depose: false, deposePar: "",
          capitauxInf: !!next.ageAgo.capitauxInf, ageContinuite: !!next.ageAgo.ageContinuite,
        };
      }
    }
    if (!next.regimeHistory) next.regimeHistory = [];
    if (next.tvaExig == null) next.tvaExig = "";
    if (!next.honoraires) next.honoraires = { montant: "", historique: [] };
    if (!next.social) next.social = { concerne: false, effectif: "", cabinetPaie: "", periodicite: "Mensuelle", odMois: {} };
    if (!next.social) next.social = { concerne: false, effectif: "", cabinetPaie: "", periodicite: "Mensuelle", odMois: {} };
// nouveaux champs, à ajouter même si next.social existe déjà :
if (next.social.gestionnaireNom == null) next.social.gestionnaireNom = "";
if (next.social.gestionnaireEmail == null) next.social.gestionnaireEmail = "";
if (next.social.gestionnaireTel == null) next.social.gestionnaireTel = "";
if (next.social.conventionCollective == null) next.social.conventionCollective = "";
if (next.social.regimeDirigeant == null) next.social.regimeDirigeant = "";
    if (!next.demandesClient) next.demandesClient = [];
    if (!next.validationDossier) next.validationDossier = { collaborateur: false, chefMission: false, dateCollaborateur: "", dateChefMission: "", commentaire: "" };
    if (!next.rentabilite) next.rentabilite = { tempsPrevu: "", tempsReel: "", tarifHoraire: "", margeCible: "" };
    if (!next.documentsSuivi) next.documentsSuivi = { demandes: 0, recus: 0, controles: 0 };
    if (!next.notesCollab) next.notesCollab = [];
    if (!next.resiliation) {
  next.resiliation = {
    active: false, date: "", initiateur: "", motif: "", motifAutre: "",
    lettreEnvoyee: false, lettreDate: "", preavisRespecte: false,
    piecesRestituees: false, piecesRestitueesDate: "",
    confrereRepreneur: "", lettreConfraterniteEnvoyee: false, lettreConfraterniteRecue: false,
    honorairesSituation: "soldes", derniereCloture: "",
    historique: [],
  };
}
    return next;
  });
}

const MISSION_EXCEP_TYPES = ["Attestation", "Prévisionnel / situation intermédiaire", "Évaluation d'entreprise", "Dossier bancaire / levée de fonds", "Cession-transmission", "Formalité ponctuelle", "Expertise", "Autre"];
const MISSION_EXCEP_STATUTS = ["a_faire", "en_cours", "bloque", "valide"];
const MISSION_EXCEP_STATUT_LABELS = { a_faire: "À faire", en_cours: "En cours", bloque: "Bloqué", valide: "Validé" };
const MISSION_EXCEP_STATUT_TONE = { a_faire: "neutral", en_cours: "amber", livree: "green" };

function MissionsExceptionnellesTab({ client, team, onUpdate }) {
  const missions = client.missionsExceptionnelles || [];
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ type: MISSION_EXCEP_TYPES[0], dateDemande: todayISO(), dateLivraisonPrevue: "", statut: "a_faire", collaborateur: "", honoraires: "", lettreSignee: false, notes: "" });

  const addMission = () => {
    const entry = { id: `me-${Date.now()}`, ...form };
    onUpdate(client.id, { missionsExceptionnelles: [...missions, entry] });
    setForm({ type: MISSION_EXCEP_TYPES[0], dateDemande: todayISO(), dateLivraisonPrevue: "", statut: "a_faire", collaborateur: "", honoraires: "", lettreSignee: false, notes: "" });
    setShowForm(false);
  };
  const patchMission = (id, patch) => onUpdate(client.id, { missionsExceptionnelles: missions.map((m) => (m.id === id ? { ...m, ...patch } : m)) });
  const removeMission = (id) => onUpdate(client.id, { missionsExceptionnelles: missions.filter((m) => m.id !== id) });

  const sorted = [...missions].sort((a, b) => (a.dateDemande < b.dateDemande ? 1 : -1));

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
        <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: 0 }}>Missions exceptionnelles ({missions.length})</h4>
        <button onClick={() => setShowForm((s) => !s)} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 9, padding: "6px 12px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" }}>
          <Plus size={13} /> Nouvelle mission
        </button>
      </div>

      {showForm && (
        <div style={{ background: T.paper, borderRadius: 10, padding: 14, marginBottom: 16 }}>
          <FieldRow label="Type de mission"><SelectPill value={form.type} options={MISSION_EXCEP_TYPES} allowEmpty={false} onChange={(v) => setForm({ ...form, type: v })} /></FieldRow>
          <FieldRow label="Date de la demande"><input type="date" value={form.dateDemande} onChange={(e) => setForm({ ...form, dateDemande: e.target.value })} style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} /></FieldRow>
          <FieldRow label="Livraison prévue"><input type="date" value={form.dateLivraisonPrevue} onChange={(e) => setForm({ ...form, dateLivraisonPrevue: e.target.value })} style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} /></FieldRow>
          <FieldRow label="Collaborateur en charge"><SelectPill value={form.collaborateur} options={team.map((t) => t.nom)} onChange={(v) => setForm({ ...form, collaborateur: v })} /></FieldRow>
          <FieldRow label="Honoraires spécifiques"><TextInput defaultValue={form.honoraires} onCommit={(v) => setForm({ ...form, honoraires: v })} placeholder="ex. 800 € HT" width={160} /></FieldRow>
          <FieldRow label="Lettre de mission spécifique signée"><ToggleBtn on={form.lettreSignee} onClick={() => setForm({ ...form, lettreSignee: !form.lettreSignee })} /></FieldRow>
          <div style={{ padding: "10px 0" }}>
            <div style={{ fontSize: 12, color: T.inkMuted, marginBottom: 6 }}>Notes / livrable</div>
            <textarea value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} rows={2}
              style={{ width: "100%", padding: 8, borderRadius: 8, border: `1px solid ${T.line}`, fontSize: 12, background: T.card, resize: "vertical" }} />
          </div>
          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 8 }}>
            <button onClick={() => setShowForm(false)} style={{ padding: "7px 12px", borderRadius: 9, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 11.5 }}>Annuler</button>
            <button onClick={addMission} style={{ padding: "7px 14px", borderRadius: 9, border: "none", background: T.navy, color: "#fff", cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>Créer</button>
          </div>
        </div>
      )}

      {sorted.length === 0 ? <EmptyNote text="Aucune mission exceptionnelle pour ce dossier." /> : sorted.map((m) => (
        <div key={m.id} style={{ border: `1px solid ${T.line}`, borderRadius: 10, padding: "12px 14px", marginBottom: 10, background: T.card }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <span style={{ fontWeight: 700, fontSize: 12.5 }}>{m.type}</span>
            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
              <select value={m.statut} onChange={(e) => patchMission(m.id, { statut: e.target.value })}
                style={{ fontFamily: T.mono, fontSize: 11, fontWeight: 700, padding: "3px 6px", borderRadius: 7, border: `1px solid ${T.line}`, background: T.card }}>
                {MISSION_EXCEP_STATUTS.map((s) => <option key={s} value={s}>{MISSION_EXCEP_STATUT_LABELS[s]}</option>)}
              </select>
              <button onClick={() => removeMission(m.id)} style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted }}><Trash2 size={13} /></button>
            </div>
          </div>
          <div style={{ fontSize: 11.5, color: T.inkMuted, display: "flex", gap: 14, flexWrap: "wrap" }}>
            <span>Demande : {fmtFR(m.dateDemande)}</span>
            {m.dateLivraisonPrevue && <span>Livraison prévue : {fmtFR(m.dateLivraisonPrevue)}</span>}
            {m.collaborateur && <span>Collab. : {m.collaborateur}</span>}
            {m.honoraires && <span>Honoraires : {m.honoraires}</span>}
            <Stamped tone={m.lettreSignee ? "green" : "amber"} small>{m.lettreSignee ? "Lettre signée" : "Lettre à signer"}</Stamped>
          </div>
          {m.notes && <div style={{ fontSize: 12, color: T.inkSoft, marginTop: 8, whiteSpace: "pre-wrap" }}>{m.notes}</div>}
        </div>
      ))}
    </div>
  );
}
/* ============================================================
   IMPORT / EXPORT EXCEL — registre clients
   ============================================================ */
const EXCEL_COLUMNS = [
  { key: "nom", label: "Nom" },
  { key: "siren", label: "SIREN" },
  { key: "logiciel", label: "Logiciel" },
  { key: "lienSharepoint", label: "Lien SharePoint" },
  { key: "collab", label: "Collaborateur" },
  { key: "expert", label: "Expert" },
  { key: "chefMission", label: "Chef de mission" },
  { key: "formeJuridique", label: "Forme juridique" },
  { key: "capital", label: "Capital" },
  { key: "activite", label: "Activité" },
  { key: "secteur", label: "Secteur (auto)" },
  { key: "dateCloture", label: "Date clôture (AAAA-MM-JJ)" },
  { key: "tvaRegime", label: "Régime TVA" },
  { key: "tvaExig", label: "Jour exigibilité TVA" },
];
// En-têtes acceptés en entrée (tolère quelques variantes usuelles côté Excel)
const EXCEL_IMPORT_ALIASES = {
  nom: ["nom", "client", "dossier", "raison sociale"],
  siren: ["siren", "siret"],
  logiciel: ["logiciel"],
  lienSharepoint: ["lien sharepoint", "sharepoint", "lien", "url"],
  collab: ["collaborateur", "collab"],
  expert: ["expert"],
  chefMission: ["chef de mission", "chefmission"],
  formeJuridique: ["forme juridique", "formejuridique"],
  capital: ["capital", "capital social"],
  activite: ["activité", "activite"],
  secteur: ["secteur", "secteur d'activite", "secteur activite"],
  dateCloture: ["date clôture (aaaa-mm-jj)", "date cloture", "date de clôture", "date clôture", "datecloture"],
  tvaRegime: ["régime tva", "regime tva", "tvaregime"],
  tvaExig: ["jour exigibilité tva", "jour exigibilite tva", "tvaexig", "exigibilité"],
};
function normalizeHeader(h) { return String(h || "").trim().toLowerCase(); }
function buildHeaderMap(headers) {
  const map = {};
  headers.forEach((h) => {
    const norm = normalizeHeader(h);
    for (const key of Object.keys(EXCEL_IMPORT_ALIASES)) {
      if (EXCEL_IMPORT_ALIASES[key].includes(norm)) { map[h] = key; return; }
    }
  });
  return map;
}
function exportClientsToExcel(clients, filename = "registre-clients-axe-experts.xlsx") {
  const rows = clients.map((c) => {
    const row = {};
    EXCEL_COLUMNS.forEach(({ key, label }) => { row[label] = c[key] ?? ""; });
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(rows, { header: EXCEL_COLUMNS.map((c) => c.label) });
  ws["!cols"] = EXCEL_COLUMNS.map(() => ({ wch: 20 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Clients");
  XLSX.writeFile(wb, filename);
}
function exportAcomptesToExcel(clients, filename = "acomptes-is-cfe.xlsx") {
  const rows = clients.map((c) => ({
    "Dossier": c.nom || "",
    "SIREN": c.siren || "",
    "Montant N-1 (IS)": c.is?.montantN1 ?? "",
    "Concerné IS": Number(c.is?.montantN1) > 3000 ? "Oui" : "Non",
    "Acompte mars": c.is?.mars ? "Fait" : "",
    "Acompte juin (IS)": c.is?.juin ? "Fait" : "",
    "Acompte sept": c.is?.sept ? "Fait" : "",
    "Acompte déc (IS)": c.is?.dec ? "Fait" : "",
    "Montant N-1 (CFE)": c.cfe?.montantN1 ?? "",
    "Concerné CFE": Number(c.cfe?.montantN1) > 3000 ? "Oui" : "Non",
    "Acompte juin (CFE)": c.cfe?.juin ? "Fait" : "",
    "Solde déc (CFE)": c.cfe?.dec ? "Fait" : "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = Object.keys(rows[0] || {}).map(() => ({ wch: 18 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Acomptes");
  XLSX.writeFile(wb, filename);
}
// Vérifie qu'une chaîne est bien une date calendaire réelle au format AAAA-MM-JJ
// (rejette par ex. "2025-13-40" ou "31/12/2025", qui passeraient un simple test de format).
function isValidISODate(str) {
  if (typeof str !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(str)) return false;
  const [y, m, d] = str.split("-").map(Number);
  if (m < 1 || m > 12) return false;
  const dt = new Date(y, m - 1, d);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

// Lit un fichier .xlsx/.xls/.csv et retourne une liste d'objets clients partiels
// (uniquement les champs reconnus), prêts à être fusionnés avec le registre existant.
// Si une ou plusieurs lignes contiennent des valeurs non conformes (date invalide,
// SIREN invalide, etc.), l'import COMPLET est refusé (rejet de la Promise) avec un
// message détaillant les lignes fautives — plutôt que d'accepter des données
// corrompues qui feraient planter l'application plus tard, ailleurs dans l'écran.
function parseClientsExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Impossible de lire le fichier."));
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (!raw.length) return resolve([]);
        const headerMap = buildHeaderMap(Object.keys(raw[0]));
        const rowErrors = [];
        const rows = raw.map((r, i) => {
          const out = {};
          const excelRowNum = i + 2; // +1 pour l'index 0-based, +1 pour la ligne d'en-tête
          Object.keys(r).forEach((h) => {
            const key = headerMap[h];
            if (!key) return;
            let v = r[h];
            if (key === "tvaExig") v = v === "" ? "" : parseInt(v, 10) || "";
            if (typeof v === "string") v = v.trim();
            if (key === "siren" && v !== "") {
              v = String(v).trim();
              if (!/^\d{9}$/.test(v) && !/^\d{14}$/.test(v)) {
                rowErrors.push({ row: excelRowNum, nom: r.nom || r.Nom || "(sans nom)", field: "SIREN/SIRET", value: v, reason: "doit contenir 9 chiffres (SIREN) ou 14 chiffres (SIRET)" });
              }
            }
            if (key === "dateCloture" && v !== "") {
              if (v instanceof Date) {
                v = v.toISOString().slice(0, 10);
              } else if (typeof v === "number") {
                // numéro de série Excel -> date JS
                const d = XLSX.SSF.parse_date_code(v);
                v = `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
              } else if (typeof v === "string") {
                v = v.trim();
              }
              if (!isValidISODate(v)) {
                rowErrors.push({ row: excelRowNum, nom: out.nom || "(sans nom)", field: "Date clôture", value: v, reason: "format attendu AAAA-MM-JJ (ex. 2025-12-31)" });
              }
            }
            out[key] = v;
          });
          return out;
        }).filter((r) => r.nom || r.siren);

        if (rowErrors.length > 0) {
          const detail = rowErrors.slice(0, 8).map(
            (e) => `Ligne ${e.row} (${e.nom}) : "${e.field}" invalide ("${e.value}") — ${e.reason}`
          ).join("\n");
          const more = rowErrors.length > 8 ? `\n… et ${rowErrors.length - 8} autre(s) ligne(s) invalide(s).` : "";
          reject(new Error(
            `Import refusé : ${rowErrors.length} ligne(s) non conforme(s). Aucune donnée n'a été importée.\n${detail}${more}`
          ));
          return;
        }

        resolve(rows);
      } catch (err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}

/* ============================================================
   STATUT TVA EFFECTIF
   ============================================================ */
function effectiveTvaStatus(client, moisKey) {
  const manual = (client.tvaMois?.[moisKey] || "").toUpperCase();
  if (manual === "OK" || manual === "FAIT" || manual === "NA" || manual === "NON_VALIDE") return manual;

  // Régime CA12 : une seule déclaration annuelle, exigible en Mai N+1.
  // Les 11 autres mois de la grille ne sont pas concernés.
  if (client.tvaRegime === "CA12") {
    if (moisKey !== "Mai") return "";
    const exig = parseInt(client.tvaExig, 10) || 3;
    const now = new Date();
    const deadline = new Date(now.getFullYear(), 4, exig, 23, 59, 59); // 4 = Mai
    return deadline.getTime() < now.getTime() ? "RETARD" : "";
  }

 const exig = parseInt(client.tvaExig, 10);
  if (!exig) return "";
  // Régime CA3 en périodicité trimestrielle : seules les échéances de fin de
  // trimestre civil (Mars, Juin, Septembre, Décembre) sont concernées ; les
  // autres mois de la grille sont non applicables.
  if (client.tvaRegime === "CA3" && client.tvaPeriodicite === "trimestrielle" && !QUARTER_END_MONTHS.includes(moisKey)) {
    return "NA";
  }
  const monthIdx = MOIS_ORDER.indexOf(moisKey);
  const now = new Date();
  // Régime CA3 : la TVA du mois M est déclarée en M+1 (ex. la TVA de juillet
  // est exigible en août, et ne passe donc en retard qu'en août).
  const deadline = new Date(now.getFullYear(), monthIdx + 1, exig, 23, 59, 59);
  return deadline.getTime() < now.getTime() ? "RETARD" : "";
}
function tvaTone(status) {
  return status === "OK" ? "green" : status === "FAIT" ? "amber" : status === "NON_VALIDE" ? "purple" : status === "RETARD" ? "red" : "neutral";
}
// Libellé affiché pour chaque statut de cellule TVA — centralisé pour rester cohérent
// entre la grille TVA et la fiche client.
function tvaStatusLabel(status) {
  return status === "RETARD" ? "Retard"
    : status === "FAIT" ? "Fait"
    : status === "OK" ? "Validé"
    : status === "NON_VALIDE" ? "Non validé"
    : status === "NA" ? "N/A"
    : "·";
}

/* ============================================================
   ÉVÉNEMENTS FISCAUX — TVA / IS / CFE / Bilan / AGE-AGO
   ============================================================ */
function computeFiscalEvents(clients) {
  const events = [];
  const now = new Date();
  const year = now.getFullYear();

  clients.forEach((c) => {
   // TVA CA3 — la déclaration due ce mois-ci porte sur le mois précédent (M+1)
    if (c.tvaRegime === "CA3" && c.tvaExig) {
      const monthIdx = now.getMonth();
      const declaredMonthIdx = (monthIdx - 1 + 12) % 12;
      const declaredMonthKey = MOIS_ORDER[declaredMonthIdx];
      const isRelevantMonth = c.tvaPeriodicite !== "trimestrielle" || QUARTER_END_MONTHS.includes(declaredMonthKey);
      if (isRelevantMonth) {
        const status = effectiveTvaStatus(c, declaredMonthKey);
        if (status !== "OK" && status !== "NA") {
          events.push({
            id: `${c.id}-tva-${declaredMonthIdx}`, client: c, category: "TVA",
            label: `TVA ${MOIS_FULL[declaredMonthKey]}${c.tvaPeriodicite === "trimestrielle" ? " (trim.)" : ""}`,
            date: new Date(year, monthIdx, parseInt(c.tvaExig, 10) || 20),
            done: false, tone: tvaTone(status),
          });
        }
      }
    }
    // TVA CA12 — une seule échéance annuelle, en Mai N+1
    if (c.tvaRegime === "CA12") {
      const status = effectiveTvaStatus(c, "Mai");
      if (status !== "OK" && status !== "NA") {
        events.push({
          id: `${c.id}-tva-ca12`, client: c, category: "TVA",
          label: "TVA annuelle (CA12)",
          date: new Date(year, 4, parseInt(c.tvaExig, 10) || 3),
          done: false, tone: tvaTone(status),
        });
      }
    }
    // IS — acomptes (dates statutaires approximatives : 15 mars/juin/sept/déc)
    if (Number(c.is?.montantN1) > 3000) {
      [["mars", 2, "mars"], ["juin", 5, "juin"], ["sept", 8, "septembre"], ["dec", 11, "décembre"]].forEach(([key, m, label]) => {
        if (!c.is[key]) {
          events.push({
            id: `${c.id}-is-${key}`, client: c, category: "IS",
            label: `Acompte IS — ${label}`, date: new Date(year, m, 15), done: false, tone: "amber",
          });
        }
      });
    }
    // CFE — 15 juin / 15 déc
    if (Number(c.cfe?.montantN1) > 3000) {
      [["juin", 5, "juin (acompte)"], ["dec", 11, "décembre (solde)"]].forEach(([key, m, label]) => {
        if (!c.cfe[key]) {
          events.push({
            id: `${c.id}-cfe-${key}`, client: c, category: "CFE",
            label: `CFE — ${label}`, date: new Date(year, m, 15), done: false, tone: "amber",
          });
        }
      });
    }
    // Clôture d'exercice
    if (c.dateCloture && typeof c.dateCloture === "string") {
      const [cy, cm, cd] = c.dateCloture.split("-").map(Number);
      if (cy && cm && cd) {
        events.push({
          id: `${c.id}-cloture`, client: c, category: "Clôture",
          label: "Clôture d'exercice", date: new Date(cy, cm - 1, cd), done: false, tone: "neutral",
        });
        // Bilan (échéance approximative : clôture + 3 mois) si pas encore transmis
        if (!c.bilan?.transmis) {
          const echeance = addMonthsISO(c.dateCloture, 3);
          if (echeance) {
            const [by, bm, bd] = echeance.split("-").map(Number);
            events.push({
              id: `${c.id}-bilan`, client: c, category: "Bilan",
              label: "Dépôt du bilan", date: new Date(by, bm - 1, bd), done: false, tone: "red",
            });
          }
        }
      }
    }
    // AGE/AGO — approbation ~6 mois après clôture si non tenue
    if (c.dateCloture && typeof c.dateCloture === "string") {
      const latestYear = Object.keys(c.ageAgoHistory || {}).sort((a, b) => b - a)[0];
      const y = latestYear ? c.ageAgoHistory[latestYear] : null;
      if (y && !y.ago) {
        const echeance = addMonthsISO(c.dateCloture, 6);
        if (echeance) {
          const [ay, am, ad] = echeance.split("-").map(Number);
          events.push({
            id: `${c.id}-ago-${latestYear}`, client: c, category: "AGO",
            label: `Approbation des comptes ${latestYear}`, date: new Date(ay, am - 1, ad), done: false, tone: "amber",
          });
        }
      }
    }
  });
  return events;
}

// Alertes de proximité d'échéance (J-7 / J-3 / J-1), calculées à partir des mêmes
// événements fiscaux que le planning — pas de table ni de champ supplémentaire.
function computeEcheanceAlerts(clients) {
  const events = computeFiscalEvents(clients);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const thresholds = [7, 3, 1];
  const alerts = [];
  events.forEach((e) => {
    const d = new Date(e.date.getFullYear(), e.date.getMonth(), e.date.getDate());
    const diff = Math.round((d - today) / 86400000);
    if (thresholds.includes(diff)) {
      alerts.push({
        id: `alert-${e.id}-${diff}`,
        client_id: e.client.id,
        client_nom: e.client.nom,
        type: "echeance",
        message: `${e.client.nom} — ${e.label} dans ${diff} jour${diff > 1 ? "s" : ""}`,
        created_at: new Date().toISOString(),
        lu: false,
        isEcheance: true,
      });
    }
  });
  return alerts.sort((a, b) => a.message.localeCompare(b.message));
}
function taskBucket(date) {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const diffDays = Math.round((d - today) / 86400000);
  if (diffDays < 0) return "retard";
  if (diffDays === 0) return "aujourdhui";
  if (diffDays === 1) return "demain";
  const weekStart = startOfWeek(today);
  const weekEnd = new Date(weekStart); weekEnd.setDate(weekEnd.getDate() + 6);
  if (d >= weekStart && d <= weekEnd) return "semaine";
  if (d.getMonth() === today.getMonth() && d.getFullYear() === today.getFullYear()) return "mois";
  const q = Math.floor(today.getMonth() / 3);
  const dq = Math.floor(d.getMonth() / 3);
  if (dq === q && d.getFullYear() === today.getFullYear()) return "trimestre";
  return "plus-tard";
}
const BUCKET_LABELS = {
  retard: "En retard", aujourdhui: "Aujourd'hui", demain: "Demain",
  semaine: "Semaine", mois: "Mois", trimestre: "Trimestre", "plus-tard": "Plus tard",
};

/* ============================================================
   STORAGE HELPERS
   ============================================================ */
/* ---- Clients : Supabase (table "clients", colonnes id text + data jsonb + portefeuille_id) ---- */
async function loadClientsFromSupabase() {
  const { data, error } = await supabase.from("clients").select("id, data, portefeuille_id");
  if (error) { console.error("Erreur chargement clients :", error.message); return null; }
  if (!data) return null;
  return data.map((row) => ({ id: row.id, portefeuilleId: row.portefeuille_id, ...(row.data || {}) }));
}
async function insertClientRemote(client) {
  const { id, portefeuilleId, ...rest } = client;

  const { error } = await supabase
    .from("clients")
    .insert({
      id,
      data: rest,
      portefeuille_id: portefeuilleId || null
    });

  if (error) {
    console.error(
      "Erreur création client :",
      error.message
    );
  }
}
async function updateClientRemote(id, client) {
  const { id: _id, portefeuilleId, ...rest } = client;

  const { error } = await supabase
    .from("clients")
    .update({
      data: rest,
      portefeuille_id: portefeuilleId || null,
    })
    .eq("id", id);

  if (error) {
    console.error("Erreur mise à jour client :", error.message);
  }
}
async function deleteClientRemote(id) {
  const { error } = await supabase
    .from("clients")
    .delete()
    .eq("id", id);

  if (error) {
    throw error;
  }
}
/* ---- Équipe : Supabase (table "team"). Un compte = un collaborateur.
   Chaque inscription (email + mot de passe) crée automatiquement, côté base
   de données, une fiche "team" qui lui est liée (voir trigger handle_new_user
   dans supabase-init.sql). Le rôle (collaborateur / expert / chef_mission / admin)
   et le portefeuille (cabinet) déterminent ce que la personne peut voir et faire. ---- */
async function loadTeamFromSupabase() {
  const { data, error } = await supabase.from("team")
    .select("id, nom, color, email, telephone, cabinet_nom, role, statut, portefeuille_id, auth_user_id")
    .order("nom", { ascending: true });
  if (error) { console.error("Erreur chargement équipe :", error.message); return null; }
  return data;
}
async function insertTeamMemberRemote(member) {
  const { error } = await supabase.from("team").insert(member);
  if (error) console.error("Erreur création collaborateur :", error.message);
}
async function updateTeamMemberRemote(id, patch) {
  const { error } = await supabase.from("team").update(patch).eq("id", id);
  if (error) console.error("Erreur mise à jour collaborateur :", error.message);
}
async function deleteTeamMemberRemote(id) {
  const { error } = await supabase.from("team").delete().eq("id", id);
  if (error) console.error("Erreur suppression collaborateur :", error.message);
}

/* ---- Portefeuilles : les cabinets clients de l'outil (Axe Experts, KOF Experts, …) ---- */
async function loadPortefeuillesFromSupabase() {
  const { data, error } = await supabase.from("portefeuilles").select("id, nom, domaine").order("nom", { ascending: true });
  if (error) { console.error("Erreur chargement portefeuilles :", error.message); return null; }
  return data;
}
async function insertPortefeuilleRemote(p) {
  const { error } = await supabase.from("portefeuilles").insert(p);
  if (error) console.error("Erreur création portefeuille :", error.message);
  return error;
}
async function loadNotificationsFromSupabase(teamId) {
  if (!teamId) return [];
  const { data, error } = await supabase.from("notifications").select("*").eq("destinataire_id", teamId).order("created_at", { ascending: false }).limit(30);
  if (error) { console.error("Erreur chargement notifications :", error.message); return []; }
  return data || [];
}
async function insertNotificationRemote(n) {
  const { error } = await supabase.from("notifications").insert(n);
  if (error) console.error("Erreur création notification :", error.message);
}
async function markNotificationReadRemote(id) {
  const { error } = await supabase.from("notifications").update({ lu: true }).eq("id", id);
  if (error) console.error("Erreur notification :", error.message);
}

const ROLE_LABELS = { collaborateur: "Collaborateur", expert: "Expert", chef_mission: "Chef de mission", admin: "Admin" };

/* ============================================================
   APP
   ============================================================ */
function CabinetApp({ session, onLogout }) {
  const [clients, setClients] = useState(null);
  const [legalRequests, setLegalRequests] = useState(() => { try { return JSON.parse(localStorage.getItem("axe-legal-requests") || "[]"); } catch { return []; } });
  useEffect(() => { try { localStorage.setItem("axe-legal-requests", JSON.stringify(legalRequests)); } catch {} }, [legalRequests]);
  const [team, setTeam] = useState(null);
  const [portefeuilles, setPortefeuilles] = useState(null);
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("Tous");
  const [regimeFilter, setRegimeFilter] = useState("Tous");
  const [collabQuickFilter, setCollabQuickFilter] = useState(null); // filtre rapide "clic depuis Supervision d'équipe"
  const [dashboardFilter, setDashboardFilter] = useState(null);
  const [showAddClient, setShowAddClient] = useState(false);
  const [saveStatus, setSaveStatus] = useState("idle");
  const [openClientTabs, setOpenClientTabs] = useState([]); // [{id, label}]
  const [activeClientTab, setActiveClientTab] = useState(null); // id du dossier affiché en page pleine, ou null
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [tasksDb, setTasksDb] = useState([]); // tâches réelles (table "tasks"), indépendantes des échéances fiscales calculées
  const [secteurContent, setSecteurContent] = useState(null);
  const [notifications, setNotifications] = useState([]);
   const [viewHistory, setViewHistory] = useState([]);

  // Empêche le canal temps réel de "rejouer" nos propres écritures juste après qu'on les a envoyées
  const pendingLocalIds = useRef(new Set());
  const pendingLocalTeamIds = useRef(new Set());
  const pendingLocalPortefeuilleIds = useRef(new Set());

  useEffect(() => {
    (async () => {
      const [storedClients, storedTeam, storedPortefeuilles, storedSecteurContent] = await Promise.all([
        loadClientsFromSupabase(),
        loadTeamFromSupabase(),
        loadPortefeuillesFromSupabase(),
        loadSecteurContentFromSupabase(),
      ]);
      if (storedClients && storedClients.length) {
        setClients(migrateClients(storedClients));
      } else {
        // Table vide (premier lancement) : on part des données d'origine, à insérer une fois dans Supabase
        setClients(migrateClients(RAW_SEED_CLIENTS));
      }
      setTeam(storedTeam || []);
      setPortefeuilles(storedPortefeuilles || []);
      // Fusion : le contenu en base prime, la seed comble les secteurs jamais édités
      setSecteurContent({ ...SEED_AIDES_SECTEUR, ...(storedSecteurContent || {}) });
      setLoading(false);
    })();

    // Un seul canal WebSocket pour les deux tables ("clients" + "team") : plus
    // performant qu'ouvrir une connexion Realtime par table, et plus simple à
    // superviser/reconnecter.
    const channel = supabase
      .channel("cabinet-realtime")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "clients" },
        (payload) => {
          const id = payload.new?.id ?? payload.old?.id;
          if (id && pendingLocalIds.current.has(id)) {
            // C'est l'écho de notre propre écriture : déjà appliqué localement, on ignore
            pendingLocalIds.current.delete(id);
            return;
          }
          if (payload.eventType === "INSERT") {
            const incoming = { id: payload.new.id, portefeuilleId: payload.new.portefeuille_id, ...(payload.new.data || {}) };
            setClients((prev) => (prev.some((c) => c.id === incoming.id) ? prev : [...prev, incoming]));
          }
          if (payload.eventType === "UPDATE") {
            const incoming = { id: payload.new.id, portefeuilleId: payload.new.portefeuille_id, ...(payload.new.data || {}) };
            setClients((prev) => prev.map((c) => (c.id === incoming.id ? incoming : c)));
          }
          if (payload.eventType === "DELETE") {
            setClients((prev) => prev.filter((c) => c.id !== payload.old.id));
          }
        }
      )
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "team" },
        (payload) => {
          const id = payload.new?.id ?? payload.old?.id;
          if (id && pendingLocalTeamIds.current.has(id)) {
            pendingLocalTeamIds.current.delete(id);
            return;
          }
          if (payload.eventType === "INSERT") {
            const incoming = payload.new;
            setTeam((prev) => (prev.some((t) => t.id === incoming.id) ? prev : [...prev, incoming]));
          }
          if (payload.eventType === "UPDATE") {
            const incoming = payload.new;
            setTeam((prev) => prev.map((t) => (t.id === incoming.id ? incoming : t)));
          }
          if (payload.eventType === "DELETE") {
            setTeam((prev) => prev.filter((t) => t.id !== payload.old.id));
          }
        }
      )
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "portefeuilles" },
        (payload) => {
          const id = payload.new?.id ?? payload.old?.id;
          if (id && pendingLocalPortefeuilleIds.current.has(id)) {
            pendingLocalPortefeuilleIds.current.delete(id);
            return;
          }
          if (payload.eventType === "INSERT") {
            const incoming = payload.new;
            setPortefeuilles((prev) => (prev.some((p) => p.id === incoming.id) ? prev : [...prev, incoming]));
          }
          if (payload.eventType === "UPDATE") {
            const incoming = payload.new;
            setPortefeuilles((prev) => prev.map((p) => (p.id === incoming.id ? incoming : p)));
          }
          if (payload.eventType === "DELETE") {
            setPortefeuilles((prev) => prev.filter((p) => p.id !== payload.old.id));
          }
        }
      )
      .subscribe((status) => {
        // Reconnexion automatique + re-synchronisation si la connexion WebSocket tombe
        // (veille du poste, coupure réseau…) : on récupère alors ce qui a pu être manqué
        // pendant la coupure, pour rester à jour sans que l'utilisateur ait à recharger.
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          Promise.all([loadClientsFromSupabase(), loadTeamFromSupabase(), loadPortefeuillesFromSupabase()]).then(([c, t, p]) => {
            if (c) setClients(migrateClients(c));
            if (t) setTeam(t);
            if (p) setPortefeuilles(p);
          });
        }
      });

    return () => { supabase.removeChannel(channel); };
  }, []);

  // Tâches (table "tasks", indépendante des échéances fiscales calculées) : chargement
  // initial + rafraîchissement à chaque changement temps réel.
  useEffect(() => {
    let cancelled = false;
    const reload = () => fetchTasks().then((rows) => { if (!cancelled) setTasksDb(rows); });
    reload();
    const unsubscribe = subscribeTasks(reload);
    return () => { cancelled = true; unsubscribe(); };
  }, []);

  /* ---- Identité : plus de sélection manuelle ("qui consulte le dossier ?").
     Le compte connecté (session Supabase Auth) est automatiquement relié à SA
     fiche "team" via auth_user_id (voir trigger handle_new_user). "me" reste le
     nom affiché, utilisé partout ailleurs dans l'app pour filtrer "mes dossiers". ---- */
  const myRow = useMemo(() => {
    if (!team || !session?.user?.id) return null;
    return team.find((t) => t.auth_user_id === session.user.id) || null;
  }, [team, session]);
  const me = myRow?.nom || null;
  const myRole = myRow?.role || null;
  const myPortefeuilleId = myRow?.portefeuille_id || null;
  const isAdmin = myRole === "admin";
  const canManageTeam = isAdmin || myRole === "expert" || myRole === "chef_mission";
  const updateSecteurContent = useCallback((secteurId, patch) => {
    setSecteurContent((prev) => ({ ...prev, [secteurId]: { ...(prev?.[secteurId] || {}), ...patch } }));
    upsertSecteurContentRemote(secteurId, patch, me);
  }, [me]);
// Notifications (alertes TVA Fait → Chef de mission, confirmation → Collaborateur)
  useEffect(() => {
    if (!myRow?.id) return;
    let cancelled = false;
    const reload = () => loadNotificationsFromSupabase(myRow.id).then((rows) => { if (!cancelled) setNotifications(rows); });
    reload();
    const channel = supabase
      .channel(`notifs-${myRow.id}`)
      .on("postgres_changes", { event: "INSERT", schema: "public", table: "notifications", filter: `destinataire_id=eq.${myRow.id}` }, () => reload())
      .subscribe();
    return () => { cancelled = true; supabase.removeChannel(channel); };
  }, [myRow?.id]);

  const markNotificationRead = useCallback((id) => {
    setNotifications((prev) => prev.map((n) => (n.id === id ? { ...n, lu: true } : n)));
    if (!String(id).startsWith("alert-")) markNotificationReadRemote(id);
  }, []);
  // Pour les opérations qui touchent plusieurs clients d'un coup (renommage/suppression d'un collaborateur)
  const persistMany = useCallback((clientsToSave) => {
    setSaveStatus("saving");
    clientsToSave.forEach((c) => pendingLocalIds.current.add(c.id));
    Promise.all(clientsToSave.map((c) => updateClientRemote(c.id, c))).then(() => {
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 1200);
    });
  }, []);

  const updateClient = useCallback((id, patch) => {
    setClients((prev) => {
      const next = prev.map((c) => (c.id === id ? { ...c, ...patch } : c));
      const updated = next.find((c) => c.id === id);
      pendingLocalIds.current.add(id);
      setSaveStatus("saving");
      updateClientRemote(id, updated).then(() => {
        setSaveStatus("saved");
        logActivity({
          clientId: id,
          portefeuilleId: updated.portefeuilleId || null,
          type: "modification",
          message: `Dossier ${updated.nom || id} modifié par ${me || "utilisateur"}`,
          auteurId: myRow?.id || null,
        });
        setTimeout(() => setSaveStatus("idle"), 1200);
      });
      return next;
    });
  }, [me, myRow?.id]);

  const addClient = useCallback((newClient) => {
    setClients((prev) => [...prev, newClient]);
    pendingLocalIds.current.add(newClient.id);
    setSaveStatus("saving");
    insertClientRemote(newClient).then(() => {
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 1200);
    });
  }, []);
  const deleteClient = useCallback(async (id) => {
  const target = clients.find((c) => c.id === id);
  if (!target) return false;

  setSaveStatus("saving");

  try {
    await deleteClientRemote(id);

    setClients((prev) => prev.filter((c) => c.id !== id));
    setOpenClientTabs((prev) => prev.filter((t) => t.id !== id));

    if (activeClientTab === id) {
      setActiveClientTab(null);
    }

    logActivity({
      clientId: id,
      portefeuilleId: target.portefeuilleId || null,
      type: "suppression",
      message: `Dossier ${target.nom || id} supprimé par ${me || "utilisateur"}`,
      auteurId: myRow?.id || null,
    });

    setSaveStatus("saved");
    setTimeout(() => setSaveStatus("idle"), 1200);

    return true;
  } catch (error) {
    console.error("Erreur suppression client :", error.message);
    setSaveStatus("idle");

    alert(
      "Impossible de supprimer ce dossier. Vérifiez vos droits Supabase et réessayez."
    );

    return false;
  }
}, [clients, activeClientTab, me, myRow?.id]);

  // Import Excel/CSV : crée ou met à jour plusieurs fiches clients d'un coup.
  // Rapprochement par SIREN si renseigné, sinon par nom (insensible à la casse).
  const importClients = useCallback((rows) => {
    if (!rows || !rows.length) return { created: 0, updated: 0 };
    let created = 0, updated = 0;
    setClients((prev) => {
      const next = [...prev];
      const keyOf = (c) => (c.siren && String(c.siren).trim()) || (c.nom || "").trim().toLowerCase();
      const byKey = new Map(next.map((c) => [keyOf(c), c]));
      const toInsert = [];
      const toUpdate = [];
      rows.forEach((row) => {
        const key = (row.siren && String(row.siren).trim()) || (row.nom || "").trim().toLowerCase();
        if (!key) return;
        const existing = byKey.get(key);
        if (existing) {
          const merged = { ...existing, ...row, id: existing.id };
          const idx = next.findIndex((c) => c.id === existing.id);
          next[idx] = merged;
          byKey.set(key, merged);
          toUpdate.push(merged);
          updated += 1;
        } else {
          const id = row.siren ? `siren-${row.siren}` : `c-import-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
          const createdClient = migrateClients([{ ...row, id }])[0];
          next.push(createdClient);
          byKey.set(key, createdClient);
          toInsert.push(createdClient);
          created += 1;
        }
      });
      setSaveStatus("saving");
      Promise.all([
        ...toInsert.map((c) => { pendingLocalIds.current.add(c.id); return insertClientRemote(c); }),
        ...toUpdate.map((c) => { pendingLocalIds.current.add(c.id); return updateClientRemote(c.id, c); }),
      ]).then(() => { setSaveStatus("saved"); setTimeout(() => setSaveStatus("idle"), 1200); });
      return next;
    });
    return { created, updated };
  }, []);

  const renameTeamMember = useCallback((oldName, newName) => {
    if (!newName.trim() || newName === oldName) return;
    const member = team.find((t) => t.nom === oldName);
    if (member) {
      setTeam((prev) => prev.map((t) => (t.nom === oldName ? { ...t, nom: newName } : t)));
      pendingLocalTeamIds.current.add(member.id);
      updateTeamMemberRemote(member.id, { nom: newName });
    }
    setClients((prev) => {
      const next = prev.map((c) => ({
        ...c,
        collab: c.collab === oldName ? newName : c.collab,
        expert: c.expert === oldName ? newName : c.expert,
        chefMission: c.chefMission === oldName ? newName : c.chefMission,
      }));
      persistMany(next);
      return next;
    });
  }, [team, persistMany]);

  // Ajout manuel réservé à l'Admin (les collaborateurs rejoignent normalement en s'inscrivant eux-mêmes) :
  // utile pour un contact externe sans compte, ou pour dépanner.
  const addTeamMember = useCallback((nom, portefeuilleId, role) => {
    if (!nom.trim() || team.some((t) => t.nom === nom.trim())) return;
    const color = PALETTE[team.length % PALETTE.length];
    const member = { id: `t-${Date.now()}`, nom: nom.trim(), color, role: role || "collaborateur", statut: "actif", portefeuille_id: portefeuilleId || null };
    setTeam((prev) => [...prev, member]);
    pendingLocalTeamIds.current.add(member.id);
    insertTeamMemberRemote(member);
  }, [team]);

  const deleteTeamMember = useCallback((nom) => {
    const member = team.find((t) => t.nom === nom);
    setTeam((prev) => prev.filter((t) => t.nom !== nom));
    if (member) {
      pendingLocalTeamIds.current.add(member.id);
      deleteTeamMemberRemote(member.id);
    }
    setClients((prev) => {
      const next = prev.map((c) => ({
        ...c,
        collab: c.collab === nom ? "" : c.collab,
        expert: c.expert === nom ? "" : c.expert,
        chefMission: c.chefMission === nom ? "" : c.chefMission,
      }));
      persistMany(next);
      return next;
    });
  }, [team, persistMany]);

  // Modification générique d'une fiche équipe : rôle, portefeuille, validation d'une
  // demande en attente (statut -> actif). Réservé côté base aux Experts/Chefs de
  // mission (sur leur propre portefeuille) et à l'Admin (partout) — voir RLS "team_update".
  const updateTeamMember = useCallback((id, patch) => {
    setTeam((prev) => prev.map((t) => (t.id === id ? { ...t, ...patch } : t)));
    pendingLocalTeamIds.current.add(id);
    updateTeamMemberRemote(id, patch);
  }, []);

  // Création d'un nouveau portefeuille (nouveau cabinet client de l'outil) : réservé à l'Admin.
  const addPortefeuille = useCallback((nom, domaine) => {
    if (!nom.trim()) return null;
    const id = `pf-${Date.now()}`;
    const p = { id, nom: nom.trim(), domaine: domaine?.trim() ? domaine.trim().toLowerCase() : null };
    setPortefeuilles((prev) => [...prev, p]);
    pendingLocalPortefeuilleIds.current.add(id);
    insertPortefeuilleRemote(p);
    return id;
  }, []);

  const myClients = useMemo(() => {
    if (!clients || !me) return [];
    return clients.filter((c) => c.collab === me || c.expert === me || c.chefMission === me);
  }, [clients, me]);

  const myTasks = useMemo(() => {
    if (!myClients.length) return [];
    const events = computeFiscalEvents(myClients);
    const missionTasks = myClients.filter((c) => {
      const m = c.mission; if (!m) return false;
      const vals = Object.values(m); if (!vals.length) return false;
      return vals.filter(Boolean).length < vals.length;
    }).map((c) => ({ id: `${c.id}-mission`, client: c, category: "Accueil", label: "Dossier d'accueil incomplet", date: new Date(), tone: "amber" }));
    return [...events, ...missionTasks].map((t) => ({ ...t, bucket: taskBucket(t.date) }));
  }, [myClients]);

  const echeanceAlerts = useMemo(() => computeEcheanceAlerts(myClients), [myClients]);

  // Même échéances, mais mises en forme comme des "tâches" pour s'afficher dans la page Mes tâches
  const autoTasksForPage = useMemo(() => myTasks.map((t) => ({
    id: `auto-${t.id}`,
    isAuto: true,
    client_id: t.client.id,
    nom: t.label,
    commentaire: t.category,
    statut: "a_faire",
    priorite: t.tone === "red" ? "urgente" : t.tone === "amber" ? "haute" : "normale",
    date_echeance: t.date ? t.date.toISOString().slice(0, 10) : null,
    responsable_id: null,
  })), [myTasks]);

  // Tâches réelles (table "tasks") visibles : celles du portefeuille du dossier
  // consulté (l'Admin, sans portefeuille attitré, voit tout).
  const visibleTasksDb = useMemo(() => {
    if (!tasksDb) return [];
    if (isAdmin) return tasksDb;
    return tasksDb.filter((t) => !t.portefeuille_id || t.portefeuille_id === myPortefeuilleId);
  }, [tasksDb, isAdmin, myPortefeuilleId]);

  const handleCreateTask = useCallback(async (payload) => {
    const row = await createTask({ ...payload, portefeuille_id: payload.portefeuille_id || myPortefeuilleId || null, created_by: me });
    if (row) {
      logActivity({ clientId: row.client_id, portefeuilleId: row.portefeuille_id, type: "tache", message: activityMessages.tacheCreee(row.nom), auteurId: myRow?.id });
    }
    return row;
  }, [myPortefeuilleId, me, myRow]);

  const handleUpdateTask = useCallback(async (id, patch) => {
    const row = await updateTask(id, patch);
    if (row && patch.statut === "termine") {
      logActivity({ clientId: row.client_id, portefeuilleId: row.portefeuille_id, type: "tache", message: activityMessages.tacheTerminee(row.nom), auteurId: myRow?.id });
    }
    return row;
  }, [myRow]);

  const handleCompleteTask = useCallback(async (task) => {
    const row = await completeTask(task.id);
    if (row) {
      logActivity({ clientId: row.client_id, portefeuilleId: row.portefeuille_id, type: "tache", message: activityMessages.tacheTerminee(row.nom), auteurId: myRow?.id });
    }
    return row;
  }, [myRow]);

  if (loading || !team) {
    return (
      <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 14 }}>
        <Loader2 className="spin" size={28} color={T.navy} />
        <div style={{ fontFamily: T.mono, fontSize: 12, letterSpacing: "0.08em", color: T.inkMuted, textTransform: "uppercase" }}>Ouverture du registre…</div>
        <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }
  if (!myRow) {
    return (
      <AccountSyncScreen
        onRetry={() => { setLoading(true); loadTeamFromSupabase().then((t) => { setTeam(t || []); setLoading(false); }); }}
        onLogout={onLogout}
      />
    );
  }
  if (myRow.statut === "en_attente") {
    return <PendingScreen row={myRow} onLogout={onLogout} />;
  }

  const meColor = team.find((t) => t.nom === me)?.color || T.navy;
  const activeClient = activeClientTab ? clients.find((c) => c.id === activeClientTab) || null : null;

  // Ouvre un dossier dans un nouvel onglet façon MyUnisoft (page pleine, plus de tiroir latéral)
  const openClientTab = (id) => {
    const c = clients.find((x) => x.id === id);
    if (!c) return;
    setOpenClientTabs((prev) => (prev.some((t) => t.id === id) ? prev : [...prev, { id, label: c.nom }]));
    setActiveClientTab(id);
  };
  const closeClientTab = (id) => {
    setOpenClientTabs((prev) => {
      const closedIndex = prev.findIndex((t) => t.id === id);
      const next = prev.filter((t) => t.id !== id);
      if (activeClientTab === id) {
        // Bascule sur le voisin le plus proche : celui qui se trouve maintenant au
        // même index (l'onglet qui suivait), sinon le précédent, sinon l'accueil.
        const neighbor = next[closedIndex] || next[closedIndex - 1] || null;
        setActiveClientTab(neighbor ? neighbor.id : null);
      }
      return next;
    });
  };
  const goHome = () => setActiveClientTab(null);
  const navTo = (v) => {
    setViewHistory((h) => (v === view ? h : [...h, view]));
    setView(v);
    setActiveClientTab(null);
    if (v !== "clients") {
      setCollabQuickFilter(null);
      setDashboardFilter(null);
    } // les filtres rapides ne vivent que sur la vue "clients"
  };
  const goBack = () => {
    if (activeClientTab) { setActiveClientTab(null); return; } // dans une fiche client → retour à la liste
    setViewHistory((h) => {
      if (!h.length) return h;
      setView(h[h.length - 1]);
      return h.slice(0, -1);
    });
  };
  const canGoBack = !!activeClientTab || viewHistory.length > 0;
  // Équipe "visible" pour les listes déroulantes (assigner un collaborateur/expert/chef
  // de mission à un dossier) : uniquement les comptes actifs de mon portefeuille — l'Admin,
  // qui n'appartient à aucun portefeuille en particulier, voit tout le monde.
  const visibleTeam = (team || []).filter((t) => t.statut !== "en_attente" && (isAdmin || t.portefeuille_id === myPortefeuilleId));
  const myPortefeuille = (portefeuilles || []).find((p) => p.id === myPortefeuilleId) || null;

  return (
    <div style={S.appShell}>
      <GlobalStyle />
      <Sidebar view={view} setView={(v) => navTo(v)} me={me} meRole={myRole} mePortefeuille={myPortefeuille} team={team}
        onLogout={onLogout} counts={{ ...computeCounts(myClients), anomalies: detectAllAnomalies(myClients).length, tachesActives: visibleTasksDb.filter((t) => t.statut !== "termine").length + autoTasksForPage.length }}
        collapsed={sidebarCollapsed} setCollapsed={setSidebarCollapsed}
        mobileOpen={mobileMenuOpen} setMobileOpen={setMobileMenuOpen} />
      <div style={S.main}>
        <TopBar search={search} setSearch={setSearch} saveStatus={saveStatus} me={me} meColor={meColor}
          openTabs={openClientTabs} activeTab={activeClientTab} onHome={goHome} onBack={goBack} canGoBack={canGoBack}
          onSelectTab={(id) => setActiveClientTab(id)} onCloseTab={closeClientTab}
          onNav={navTo} onOpenClient={openClientTab} onNewClient={() => setShowAddClient(true)} clients={myClients}
          notifCount={myTasks.filter((t) => t.bucket === "retard").length || undefined}
          onOpenMobileMenu={() => setMobileMenuOpen(true)}
          notifications={[...echeanceAlerts, ...notifications]} onMarkNotificationRead={markNotificationRead} onOpenClient2={openClientTab}
 />
        <div className="px-3 py-3 md:px-7 md:py-6" style={{ ...S.content, padding: undefined }}>
          {activeClient ? (
            // key={activeClient.id} force le remontage complet du composant à chaque
            // changement d'onglet : les champs non-contrôlés (defaultValue) et l'état
            // interne (onglet secondaire "Infos / TVA / Bilan…") sont ainsi réinitialisés
            // avec les données du dossier sélectionné, au lieu de rester figés sur
            // l'ancien dossier affiché.
            <ClientEditorPage
  key={activeClient.id}
  client={activeClient}
  team={visibleTeam}
  me={me}
  meId={myRow?.id}
  portefeuilleId={myPortefeuilleId}
  onUpdate={updateClient}
  onDelete={deleteClient}
  onClose={() => closeClientTab(activeClient.id)}
  setView={navTo}
/>
          ) : (
            <>
              {view === "pilotage" && <PilotageView clients={myClients} tasks={[...visibleTasksDb, ...autoTasksForPage]} team={visibleTeam} me={me} onOpenClient={openClientTab} onView={navTo} />}
              {view === "dashboard" && (
                <Dashboard myClients={myClients} tasks={myTasks} me={me} meRole={myRole} team={visibleTeam}
                  onOpenClient={(id) => { navTo("clients"); openClientTab(id); }} setView={navTo}
                  onSuperviseClick={(collab) => { setCollabQuickFilter(collab); setDashboardFilter(null); navTo("clients"); }}
                  onDashboardFilter={(filter) => { setDashboardFilter(filter); setCollabQuickFilter(null); navTo("clients"); }} />
              )}
              {view === "clients" && (
                <ClientsRegistry clients={myClients} allClients={clients} search={search} setSearch={setSearch} roleFilter={roleFilter} setRoleFilter={setRoleFilter}
                  regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter} me={me} isAdmin={isAdmin}
                  collabQuickFilter={collabQuickFilter} setCollabQuickFilter={setCollabQuickFilter}
                  dashboardFilter={dashboardFilter} setDashboardFilter={setDashboardFilter}
                  selected={activeClientTab} setSelected={openClientTab} onAdd={() => setShowAddClient(true)}
                  onUpdate={updateClient}
onImport={importClients} onAddClient={addClient} />
              )}
              {view === "tva" && <TvaGrid clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter}
                me={me}
                onCycle={(id, mois, val) => {
                  const c = clients.find(x => x.id === id);
                  if (!c) return;
                  const previous = (c.tvaMois?.[mois] || "").toUpperCase();
                  updateClient(id, { tvaMois: { ...(c.tvaMois || {}), [mois]: val } });
                  // Collaborateur passe la cellule à "Fait" -> notification persistante au chef de mission.
                  if (val === "FAIT" && previous !== "FAIT") {
                    const dest = c.chefMission_id
                      ? team.find((t) => t.id === c.chefMission_id)
                      : (team.find((t) => t.nom === c.chefMission) || team.find((t) => t.role === "chef_mission" && (!c.portefeuilleId || t.portefeuille_id === c.portefeuilleId)));
                    if (dest && dest.id !== myRow?.id) insertNotificationRemote({
                      id: `n-${crypto?.randomUUID ? crypto.randomUUID() : Date.now()}`, destinataire_id: dest.id, expediteur_id: myRow?.id || null,
                      client_id: c.id, client_nom: c.nom, type: "tva_fait",
                      message: `${me} a préparé la TVA de ${mois} pour ${c.nom} — à vérifier.`,
                    });
                  }
                  // Chef de mission confirme (Fait -> OK) -> notifie le collaborateur en retour
                  if (val === "OK" && previous === "FAIT" && c.collab && c.collab !== me) {
                    const dest = team.find((t) => t.nom === c.collab);
                    if (dest) insertNotificationRemote({
                      id: `n-${Date.now()}`, destinataire_id: dest.id, expediteur_id: myRow?.id || null,
                      client_id: c.id, client_nom: c.nom, type: "tva_confirme",
                      message: `${me} a confirmé la TVA de ${mois} pour ${c.nom}.`,
                    });
                  }
                }}
                onReview={(id, mois, decision, commentaire) => {
                  const c = clients.find(x => x.id === id);
                  if (!c) return;
                  const patch = { tvaMois: { ...(c.tvaMois || {}), [mois]: decision } };
                  const nextControle = { ...(c.tvaControle || {}) };
                  if (decision === "NON_VALIDE") {
                    nextControle[mois] = { commentaire: commentaire || "", par: me, date: new Date().toISOString() };
                  } else {
                    // Contrôle validé : la remarque précédente n'a plus lieu d'être
                    delete nextControle[mois];
                  }
                  patch.tvaControle = nextControle;
                  updateClient(id, patch);
                  // Le chef de mission / expert qui contrôle notifie le collaborateur du dossier,
                  // qu'il soit invité à déclarer (validé) ou à corriger avant de déclarer (non validé).
                  if (c.collab && c.collab !== me) {
                    const dest = team.find((t) => t.nom === c.collab);
                    if (dest) insertNotificationRemote({
                      id: `n-${crypto?.randomUUID ? crypto.randomUUID() : Date.now()}`,
                      destinataire_id: dest.id, expediteur_id: myRow?.id || null,
                      client_id: c.id, client_nom: c.nom,
                      type: decision === "OK" ? "tva_confirme" : "tva_a_corriger",
                      message: decision === "OK"
                        ? `${me} a contrôlé et validé la TVA de ${mois} pour ${c.nom} — la déclaration peut être faite.`
                        : `${me} a contrôlé la TVA de ${mois} pour ${c.nom} : des modifications sont nécessaires avant la déclaration.${commentaire ? " " + commentaire : ""}`,
                    });
                  }
                }}
                onUpdate={updateClient} onOpenClient={openClientTab} />}
              {view === "bilans" && <BilansView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "acomptes" && <AcomptesView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "prestations-juridiques" && <LegalServicesView clients={myClients} requests={legalRequests} setRequests={setLegalRequests} />}
              {view === "age" && <AgeAgoView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "revision" && (
  <RevisionView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} setView={navTo} />
)}
              {view === "surveillance" && <SurveillanceView clients={myClients} search={search} me={me} onOpenClient={openClientTab} />}
              {view === "mission" && <MissionView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "regimes" && <RegimeChangeView clients={myClients} me={me} search={search} onUpdate={updateClient} />}
              {view === "honoraires" && <HonorairesView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} meId={myRow?.id} portefeuilleId={myPortefeuilleId} onUpdate={updateClient} />}
{view === "gestionnaire-paie" && <GestionnairePaieView clients={myClients} search={search} setSearch={setSearch} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "cotisations" && <CotisationsSocialesView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "social" && <CadreSocialView clients={myClients} search={search} setSearch={setSearch} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "fiscal" && <SuiviFiscalView clients={myClients} team={team} />}
              {view === "resiliation" && <ResiliationsView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} meId={myRow?.id} portefeuilleId={myPortefeuilleId} onUpdate={updateClient} />}
{view === "missionsExcep" && <MissionsExceptionnellesView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} team={team} />}
              {view === "reprise" && <ReprisesView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} meId={myRow?.id} portefeuilleId={myPortefeuilleId} onUpdate={updateClient} />}
              {view === "mes-taches" && (
                <TasksPage tasks={[...visibleTasksDb, ...autoTasksForPage]} clients={myClients} team={visibleTeam} me={me} myRow={myRow}
                  onCreate={handleCreateTask} onUpdate={handleUpdateTask} onComplete={handleCompleteTask}
                  onDelete={deleteTask} onOpenClient={openClientTab} />
              )}
            {view === "planning" && (
              <PlanningView
                tasks={[...visibleTasksDb, ...autoTasksForPage].filter((t) => !["TVA", "IS", "CFE", "Bilan", "Clôture", "AGO"].includes(t.commentaire))}
                clients={myClients} me={me}
                onUpdate={handleUpdateTask} onOpenClient={openClientTab} />
            )}
              {view === "equipe" && (
                <EquipeView team={team} portefeuilles={portefeuilles || []} clients={clients}
                  myRole={myRole} isAdmin={isAdmin} myPortefeuilleId={myPortefeuilleId}
                  canManageTeam={canManageTeam}
                  onAdd={addTeamMember} onRename={renameTeamMember} onDelete={deleteTeamMember}
                  onUpdateMember={updateTeamMember} onAddPortefeuille={addPortefeuille} />
              )}
              {view === "aides-secteur" && (
                <AidesSecteurView content={secteurContent || {}} canEdit={canManageTeam} onUpdate={updateSecteurContent} />
              )}
            </>
          )}
        </div>
      </div>

      {showAddClient && (
        <AddClientModal team={visibleTeam} me={me} portefeuilleId={myPortefeuilleId} onClose={() => setShowAddClient(false)}
          onCreate={(c) => { addClient(c); setShowAddClient(false); setOpenClientTabs((prev) => [...prev, { id: c.id, label: c.nom }]); setActiveClientTab(c.id); }} />
      )}
      <SaveToast status={saveStatus} />
    </div>
  );
}

/* ============================================================
   TOAST DE SAUVEGARDE — retour visuel clair après modification
   d'un champ (ex. Infos générales), en plus du petit indicateur
   discret dans la barre du haut.
   ============================================================ */
function SaveToast({ status }) {
  if (status !== "saved" && status !== "saving") return null;
  const saved = status === "saved";
  return (
    <div style={{
      position: "fixed", bottom: 20, right: 20, zIndex: 200,
      display: "flex", alignItems: "center", gap: 8,
      padding: "9px 16px", borderRadius: 12, fontSize: 12.5, fontWeight: 600,
      background: saved ? T.green : T.card, color: saved ? "#fff" : T.inkMuted,
      border: saved ? "none" : `1px solid ${T.line}`,
      boxShadow: T.shadowLg,
    }} className="reveal">
      {saved ? <Check size={14} /> : <Loader2 size={14} className="spin" />}
      {saved ? "Enregistré" : "Enregistrement…"}
    </div>
  );
}

/* ============================================================
   AUTHENTIFICATION — porte d'accès obligatoire (Supabase Auth)
   Tant qu'aucune session valide n'existe, seul <AuthPage/> est monté :
   il est donc impossible d'atteindre le tableau de bord sans se
   connecter. La session est écoutée en temps réel (connexion,
   déconnexion, expiration) via supabase.auth.onAuthStateChange.
   ============================================================ */
export default function App() {
  const [session, setSession] = useState(undefined); // undefined = vérification en cours, null = déconnecté
  const [recoveryMode, setRecoveryMode] = useState(false);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session ?? null));
    const { data: sub } = supabase.auth.onAuthStateChange((event, sess) => {
      if (event === "PASSWORD_RECOVERY") setRecoveryMode(true);
      setSession(sess);
    });
    return () => sub.subscription.unsubscribe();
  }, []);

  if (session === undefined) {
    return (
      <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 14, background: T.paper }}>
        <GlobalStyle />
        <Loader2 className="spin" size={26} color={T.navy} />
        <div style={{ fontFamily: T.mono, fontSize: 11, letterSpacing: "0.08em", color: T.inkMuted, textTransform: "uppercase" }}>Vérification de la session…</div>
        <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }
  if (recoveryMode) return <NewPasswordPage onDone={() => setRecoveryMode(false)} />;
  if (!session) return <AuthPage />;
  return <CabinetApp session={session} onLogout={() => supabase.auth.signOut()} />;
}

function NewPasswordPage({ onDone }) {
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const submit = async (e) => {
    e.preventDefault();
    setError(""); setLoading(true);
    const { error: err } = await supabase.auth.updateUser({ password });
    setLoading(false);
    if (err) { setError(err.message || "Une erreur est survenue."); return; }
    onDone();
  };
  return (
    <div style={{ minHeight: "100vh", width: "100%", display: "flex", alignItems: "center", justifyContent: "center", background: T.paper, fontFamily: T.sans, padding: 20 }}>
      <GlobalStyle />
      <div style={{ width: 380, maxWidth: "94vw", background: T.card, borderRadius: T.radiusLg, boxShadow: T.shadowLg, border: `1px solid ${T.line}`, padding: "32px 30px" }}>
        <div style={{ fontFamily: T.serif, fontWeight: 800, fontSize: 17, color: T.ink, marginBottom: 6 }}>Nouveau mot de passe</div>
        <p style={{ fontSize: 12.5, color: T.inkMuted, marginTop: 0, marginBottom: 20 }}>Choisissez un nouveau mot de passe pour votre compte.</p>
        <form onSubmit={submit} style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          <div style={{ position: "relative" }}>
            <Lock size={15} style={authIconStyle} />
            <input required type="password" minLength={6} autoComplete="new-password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" style={authInputStyle} />
          </div>
          {error && <div style={{ fontSize: 12.5, color: T.red, background: T.redSoft, padding: "10px 12px", borderRadius: 10 }}>{error}</div>}
          <button type="submit" disabled={loading} style={{ padding: "12px 0", borderRadius: 12, border: "none", background: T.navy, color: "#fff", fontSize: 13.5, fontWeight: 700, cursor: loading ? "default" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8 }}>
            {loading && <Loader2 size={15} className="spin" />} Valider
          </button>
        </form>
      </div>
    </div>
  );
}

/* ============================================================
   AUTH PAGE — connexion / inscription (style Kabineo)
   ============================================================ */
function AuthPage() {
  const [mode, setMode] = useState("login"); // login | signup
  const [fullName, setFullName] = useState("");
  const [email, setEmail] = useState("");
  const [telephone, setTelephone] = useState("");
  const [cabinetNom, setCabinetNom] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setInfo(""); setLoading(true);
    try {
      if (mode === "reset") {
        const { error: err } = await supabase.auth.resetPasswordForEmail(email, { redirectTo: window.location.origin });
        if (err) throw err;
        setInfo("Un email vous a été envoyé avec un lien pour réinitialiser votre mot de passe.");
      } else if (mode === "login") {
        const { error: err } = await supabase.auth.signInWithPassword({ email, password });
        if (err) throw err;
      } else {
        const { data, error: err } = await supabase.auth.signUp({
          email, password, options: { data: { full_name: fullName, telephone, cabinet_nom: cabinetNom } },
        });
        if (err) throw err;
        if (!data.session) {
          setInfo("Compte créé — vérifiez votre boîte mail pour confirmer votre adresse avant de vous connecter.");
        }
      }
    } catch (err) {
      const msg = err?.message === "Invalid login credentials" ? "Email ou mot de passe incorrect." : (err?.message || "Une erreur est survenue.");
      setError(msg);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: "100vh", width: "100%", display: "flex", alignItems: "center", justifyContent: "center",
      background: `radial-gradient(circle at 12% 8%, ${T.navySoft} 0%, ${T.paper} 42%), radial-gradient(circle at 90% 92%, #EEF2FF 0%, ${T.paper} 40%)`,
      fontFamily: T.sans, padding: 20,
    }}>
      <GlobalStyle />
      <motion.div initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.55, ease: EASE }}
        style={{ width: 420, maxWidth: "94vw", background: T.card, borderRadius: T.radiusLg, boxShadow: T.shadowLg, border: `1px solid ${T.line}`, padding: "38px 34px" }}>
        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.45, delay: 0.05, ease: EASE }}
          style={{ display: "flex", flexDirection: "column", alignItems: "center", marginBottom: 28 }}>
          <div style={{ width: 46, height: 46, borderRadius: 14, background: T.navy, display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 14, boxShadow: "0 8px 20px -6px rgba(79,70,229,0.45)" }}>
            <LayoutGrid size={21} color="#fff" strokeWidth={2.2} />
          </div>
          <div style={{ fontFamily: T.serif, fontWeight: 800, fontSize: 19, color: T.ink }}>AXE-EXPERTS</div>
          <div style={{ fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted, letterSpacing: "0.08em", textTransform: "uppercase", marginTop: 3 }}>Registre &amp; Pilotage</div>
        </motion.div>

        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.45, delay: 0.1, ease: EASE }}
          style={{ display: "flex", background: T.paper, borderRadius: 12, padding: 4, marginBottom: 24, gap: 4 }}>
          {[["login", "Se connecter"], ["signup", "S'inscrire"]].map(([id, label]) => (
            <button key={id} type="button" onClick={() => { setMode(id); setError(""); setInfo(""); }} style={{
              flex: 1, padding: "10px 0", border: "none", borderRadius: 9, cursor: "pointer", fontSize: 13, fontWeight: 700,
              background: mode === id ? T.card : "transparent", color: mode === id ? T.navy : T.inkMuted,
              boxShadow: mode === id ? T.shadowSm : "none", transition: "all .18s ease",
            }}>{label}</button>
          ))}
        </motion.div>

        <motion.form onSubmit={submit} initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.45, delay: 0.15, ease: EASE }}
          style={{ display: "flex", flexDirection: "column", gap: 15 }}>
          {mode === "signup" && (
            <div>
              <label style={authLabelStyle}>Nom et prénom</label>
              <div style={{ position: "relative" }}>
                <UserRound size={15} style={authIconStyle} />
                <input required value={fullName} onChange={(e) => setFullName(e.target.value)} placeholder="ex. Cheikh Diallo" style={authInputStyle} />
              </div>
            </div>
          )}
          <div>
            <label style={authLabelStyle}>Adresse email</label>
            <div style={{ position: "relative" }}>
              <Mail size={15} style={authIconStyle} />
              <input required type="email" autoComplete="email" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="vous@cabinet.fr" style={authInputStyle} />
            </div>
          </div>
          {mode === "signup" && (
            <div>
              <label style={authLabelStyle}>Téléphone</label>
              <div style={{ position: "relative" }}>
                <Phone size={15} style={authIconStyle} />
                <input required type="tel" value={telephone} onChange={(e) => setTelephone(e.target.value)} placeholder="06 12 34 56 78" style={authInputStyle} />
              </div>
            </div>
          )}
          {mode === "signup" && (
            <div>
              <label style={authLabelStyle}>Nom du cabinet</label>
              <div style={{ position: "relative" }}>
                <Briefcase size={15} style={authIconStyle} />
                <input value={cabinetNom} onChange={(e) => setCabinetNom(e.target.value)} placeholder="ex. Cabinet Dupont & Associés" style={authInputStyle} />
              </div>
              <div style={{ fontSize: 10.5, color: T.inkMuted, marginTop: 5, lineHeight: 1.5 }}>
                Si votre cabinet a déjà un accès (@axe-experts.com, @kof-experts.com), votre compte sera activé immédiatement.
                Sinon, cette information nous sert à vous recontacter pour activer votre accès.
              </div>
            </div>
          )}
         {mode !== "reset" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                <label style={{ ...authLabelStyle, marginBottom: 0 }}>Mot de passe</label>
                {mode === "login" && (
                  <button type="button" onClick={() => { setMode("reset"); setError(""); setInfo(""); }} style={{ ...authLinkStyle, fontSize: 11 }}>
                    Mot de passe oublié ?
                  </button>
                )}
              </div>
              <div style={{ position: "relative" }}>
                <Lock size={15} style={authIconStyle} />
                <input required type="password" minLength={6} autoComplete={mode === "login" ? "current-password" : "new-password"}
                  value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" style={authInputStyle} />
              </div>
            </div>
          )}

          {error && <div style={{ fontSize: 12.5, color: T.red, background: T.redSoft, padding: "10px 12px", borderRadius: 10 }}>{error}</div>}
          {info && <div style={{ fontSize: 12.5, color: T.green, background: T.greenSoft, padding: "10px 12px", borderRadius: 10 }}>{info}</div>}

          <button type="submit" disabled={loading} style={{
            marginTop: 4, padding: "13px 0", borderRadius: 12, border: "none", background: T.navy, color: "#fff",
            fontSize: 14, fontWeight: 700, cursor: loading ? "default" : "pointer", opacity: loading ? 0.75 : 1,
            display: "flex", alignItems: "center", justifyContent: "center", gap: 8, boxShadow: "0 10px 24px -8px rgba(79,70,229,0.5)",
          }}>
            {loading && <Loader2 size={15} className="spin" />}
            {mode === "login" ? "Se connecter" : mode === "reset" ? "Envoyer le lien" : "Créer mon compte"}
          </button>
        </motion.form>

        <div style={{ textAlign: "center", marginTop: 20, fontSize: 12, color: T.inkMuted }}>
          {mode === "login" && (
            <>Pas encore de compte ? <button type="button" onClick={() => { setMode("signup"); setError(""); setInfo(""); }} style={authLinkStyle}>Inscrivez-vous</button></>
          )}
          {mode === "signup" && (
            <>Déjà un compte ? <button type="button" onClick={() => { setMode("login"); setError(""); setInfo(""); }} style={authLinkStyle}>Connectez-vous</button></>
          )}
          {mode === "reset" && (
            <>Vous vous souvenez de votre mot de passe ? <button type="button" onClick={() => { setMode("login"); setError(""); setInfo(""); }} style={authLinkStyle}>Connectez-vous</button></>
          )}
        </div>
      </motion.div>
      <style>{`.spin{animation:spin .8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
const authLabelStyle = { display: "block", fontSize: 11.5, fontWeight: 600, color: T.inkSoft, marginBottom: 6 };
const authInputStyle = { width: "100%", padding: "10px 12px 10px 36px", borderRadius: 11, border: `1px solid ${T.line}`, fontSize: 13.5, background: T.paper, color: T.ink, fontFamily: T.sans };
const authIconStyle = { position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: T.inkMuted, pointerEvents: "none" };
const authLinkStyle = { background: "none", border: "none", padding: 0, color: T.navy, fontWeight: 700, cursor: "pointer", fontSize: 12 };

/* ============================================================
   DESIGN TOKENS — voir la définition de T tout en haut du fichier
   ============================================================ */
function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');
      * { box-sizing: border-box; } html { font-size: 12px; } body { margin: 0; background: ${T.paper}; font-size: 11.5px; } ::selection { background: ${T.navySoft}; }
      .hoverRow:hover { background: ${T.paperDeep} !important; } .clickable { cursor: pointer; }
      .topTab { display:flex; align-items:center; gap:6px; padding:7px 12px 7px 14px; font-size:12.5px; font-weight:600; border-radius:8px 8px 0 0; cursor:pointer; white-space:nowrap; }
      .topIconBtn { display:flex; align-items:center; justify-content:center; width:32px; height:32px; border-radius:9px; border:none; background:none; cursor:pointer; color:${T.inkMuted}; position:relative; flex-shrink:0; }
      .topIconBtn:hover { background:${T.paperDeep}; color:${T.ink}; }
      .scrollbar::-webkit-scrollbar { width: 8px; height: 8px; } .scrollbar::-webkit-scrollbar-thumb { background: ${T.line}; border-radius: 8px; }
      input, select, button, textarea { font-family: ${T.sans}; }
      button { transition: transform .15s ease, box-shadow .15s ease, background-color .15s ease, opacity .15s ease, border-color .15s ease; }
      button.clickable:hover, div.clickable:hover { transform: translateY(-1px); }
      button:focus-visible, input:focus-visible, select:focus-visible, [tabindex]:focus-visible { outline: 2px solid ${T.navy}; outline-offset: 2px; }
      .filterField { border: 1px solid ${T.line}; background: ${T.card}; transition: border-color .15s ease, box-shadow .15s ease; }
      .filterField:hover { border-color: ${T.navy}; }
      .filterField:focus, .filterField:focus-visible { border-color: ${T.navy}; box-shadow: 0 0 0 3px ${T.navySoft}; outline: none; }
      .sideGroupHeader { transition: color .15s ease; cursor: pointer; }
      .sideGroupHeader:hover { color: #fff !important; }
      .sideNavItem { transition: background-color .15s ease, color .15s ease; }
      .sideNavItem:hover { background: ${T.sidebarBg2} !important; color: #fff !important; }
      .statusToggle { transition: background-color .15s ease, color .15s ease, border-color .15s ease; cursor: pointer; }
      .statusToggle:hover { filter: brightness(0.96); }
      @keyframes fadeInUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
      .reveal { animation: fadeInUp .55s cubic-bezier(.16,.84,.44,1) both; }
      @media (prefers-reduced-motion: reduce) { * { animation: none !important; transition: none !important; } }
    `}</style>
  );
}
const S = {
  appShell: { display: "flex", height: "100vh", width: "100%", background: T.paper, fontFamily: T.sans, color: T.ink, overflow: "hidden" },
  main: { flex: 1, display: "flex", flexDirection: "column", minWidth: 0 },
  content: { flex: 1, overflowY: "auto", overflowX: "hidden", padding: "22px 28px 48px" },
};

/* ============================================================
   REVEAL — apparition progressive (fade + slide-up), avec cascade
   ============================================================ */
const EASE = [0.16, 0.84, 0.44, 1];
function Reveal({ children, index = 0, delay = 0, style, ...rest }) {
  const staggerIndex = Math.min(index, 10); // évite des délais trop longs sur les longues listes
  return (
    <motion.div
      initial={{ opacity: 0, y: 16 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true, margin: "-60px" }}
      transition={{ duration: 0.45, delay: delay + staggerIndex * 0.05, ease: EASE }}
      style={style}
      {...rest}
    >
      {children}
    </motion.div>
  );
}

/* ============================================================
   ACCOUNT SYNC SCREEN — cas rare où la fiche "team" (créée par le
   trigger à l'inscription) n'est pas encore visible côté client.
   ============================================================ */
function AccountSyncScreen({ onRetry, onLogout }) {
  return (
    <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 16, background: `radial-gradient(circle at 20% 15%, ${T.navySoft} 0%, ${T.paper} 45%), radial-gradient(circle at 85% 85%, #F1F5F9 0%, ${T.paper} 40%)` }}>
      <GlobalStyle />
      <Reveal style={{ textAlign: "center", maxWidth: 420, padding: 36, background: T.card, borderRadius: T.radiusLg, boxShadow: T.shadowLg, border: `1px solid ${T.line}` }}>
        <Loader2 size={26} color={T.navy} className="spin" style={{ marginBottom: 14 }} />
        <h1 style={{ fontFamily: T.serif, fontWeight: 700, fontSize: 17, margin: "0 0 8px", color: T.ink }}>Finalisation de votre compte…</h1>
        <p style={{ color: T.inkMuted, fontSize: 12.5, lineHeight: 1.6, marginBottom: 22 }}>
          Votre fiche collaborateur est en cours de création. Si ça persiste plus de quelques secondes, réessayez.
        </p>
        <div style={{ display: "flex", gap: 8, justifyContent: "center" }}>
          <button onClick={onRetry} style={{ padding: "9px 16px", borderRadius: 10, border: "none", background: T.navy, color: "#fff", cursor: "pointer", fontSize: 12.5, fontWeight: 700 }}>Réessayer</button>
          <button onClick={onLogout} style={{ padding: "9px 16px", borderRadius: 10, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 600, color: T.inkMuted }}>Déconnexion</button>
        </div>
        <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </Reveal>
    </div>
  );
}

/* ============================================================
   PENDING SCREEN — compte inscrit avec un email dont le domaine
   n'est rattaché à aucun portefeuille connu : en attente de
   validation par l'Admin (démarchage / devis / création manuelle).
   ============================================================ */
function PendingScreen({ row, onLogout }) {
  return (
    <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 16, background: `radial-gradient(circle at 20% 15%, ${T.amberSoft} 0%, ${T.paper} 45%), radial-gradient(circle at 85% 85%, #F1F5F9 0%, ${T.paper} 40%)` }}>
      <GlobalStyle />
      <Reveal style={{ textAlign: "center", maxWidth: 460, padding: 36, background: T.card, borderRadius: T.radiusLg, boxShadow: T.shadowLg, border: `1px solid ${T.line}` }}>
        <div style={{ width: 44, height: 44, borderRadius: 14, background: T.amber, margin: "0 auto 16px", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <Clock size={20} color="#fff" strokeWidth={2.2} />
        </div>
        <h1 style={{ fontFamily: T.serif, fontWeight: 700, fontSize: 18, margin: "0 0 8px", color: T.ink }}>Inscription en attente de validation</h1>
        <p style={{ color: T.inkMuted, fontSize: 12.5, lineHeight: 1.7, marginBottom: 10 }}>
          Merci {row.nom} — votre adresse <strong>{row.email}</strong> n'est pas encore rattachée à un cabinet sur cet outil.
        </p>
        <p style={{ color: T.inkMuted, fontSize: 12.5, lineHeight: 1.7, marginBottom: 22 }}>
          Notre équipe va prendre contact avec vous prochainement pour échanger sur votre cabinet et activer votre accès.
        </p>
        <button onClick={onLogout} style={{ padding: "9px 18px", borderRadius: 10, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 600, color: T.inkMuted }}>Déconnexion</button>
      </Reveal>
    </div>
  );
}

/* ============================================================
   SIDEBAR
   ============================================================ */
function Sidebar({ view, setView, me, meRole, mePortefeuille, team, onLogout, counts, collapsed, setCollapsed, mobileOpen, setMobileOpen }) {
    const GROUPS = [
    {
      id: "clients-accueil", label: "Gestion clients & accueil",
      items: [
        { id: "clients", label: "Registre clients", icon: Users, badge: counts.total },
        { id: "pilotage", label: "Pilotage cabinet", icon: LayoutGrid },
        { id: "mission", label: "Dossiers en accueil", icon: ClipboardCheck, badge: counts.missionIncomplete, badgeTone: "amber" },
        { id: "regimes", label: "Changements de régime", icon: RefreshCw },
        { id: "honoraires", label: "Honoraires", icon: Wallet },
      ],
    },
    {
  id: "fiscalite", label: "Fiscalité & comptabilité",
  items: [
    { id: "tva", label: "TVA (CA3 / CA12)", icon: Receipt, badge: counts.tvaAlert, badgeTone: "amber" },
    { id: "acomptes", label: "Impôts & cotisations", icon: Landmark },
    { id: "bilans", label: "Bilans", icon: FileWarning, badge: counts.bilanRetard, badgeTone: "red" },
    { id: "revision", label: "Révision comptable", icon: Search },
    { id: "surveillance", label: "À surveiller", icon: ShieldAlert, badge: counts.anomalies, badgeTone: "red" },
    { id: "fiscal", label: "Suivi fiscal", icon: CalendarDays },
  ],
},
    {
      id: "juridique", label: "Juridique",
      items: [
        { id: "prestations-juridiques", label: "Prestations juridiques", icon: Scale },
        { id: "age", label: "Assemblées (AGE / AGO)", icon: Building2, badge: counts.ageAlert, badgeTone: "amber" },
      ],
    },
    {
      id: "ressources", label: "Ressources sectorielles",
      items: [
        { id: "aides-secteur", label: "Actualités & Aides", icon: Briefcase },
      ],
    },
    {
  id: "evenements-client", label: "Événements client",
  items: [
    { id: "resiliation", label: "Résiliations", icon: FileWarning },
    { id: "missionsExcep", label: "Missions exceptionnelles", icon: Briefcase },
    { id: "reprise", label: "Reprises", icon: RefreshCw },
  ],
},
    {
      id: "social", label: "Social & paie",
      items: [
        { id: "gestionnaire-paie", label: "Gestionnaire de paie", icon: Contact },
        { id: "cotisations", label: "Cotisations sociales", icon: Landmark },
        { id: "social", label: "Suivi social (OD salaires)", icon: UserCheck },
      ],
    },
    {
      id: "organisation", label: "Organisation & équipe",
      items: [
        { id: "mes-taches", label: "Mes tâches", icon: ClipboardCheck, badge: counts.tachesActives, badgeTone: "amber" },
        { id: "planning", label: "Mon planning", icon: CalendarRange },
        { id: "equipe", label: "Équipe", icon: Settings2 },
      ],
    },
  ];
  const groupOf = (v) => GROUPS.find((g) => g.items.some((it) => it.id === v))?.id;
  const [openGroups, setOpenGroups] = useState(() => new Set([groupOf(view) || GROUPS[0].id]));
  const toggleGroup = (id) => setOpenGroups((prev) => {
    const next = new Set(prev);
    next.has(id) ? next.delete(id) : next.add(id);
    return next;
  });
  useEffect(() => {
    const g = groupOf(view);
    if (g) setOpenGroups((prev) => (prev.has(g) ? prev : new Set(prev).add(g)));
  }, [view]);

  const meColor = team.find((t) => t.nom === me)?.color || T.navy;

  const badgeToneCls = (tone) => tone === "red" ? "bg-badge-red-bg text-badge-red-text" : tone === "amber" ? "bg-badge-amber-bg text-badge-amber-text" : "bg-accent-soft text-accent-deep";

  // isMobile=true force toujours l'affichage déplié (le mode "réduit" n'a de sens que sur desktop)
  const SidebarInner = ({ isMobile }) => {
    const isCollapsed = isMobile ? false : collapsed;
    const NavButton = ({ it, indent }) => {
      const active = view === it.id; const Icon = it.icon;
      return (
        <button onClick={() => { setView(it.id); if (isMobile) setMobileOpen?.(false); }} title={it.label}
          className={`nav-item relative w-full ${isCollapsed ? "justify-center px-0" : `justify-start ${indent ? "pl-6" : ""}`} ${active ? "nav-item-active" : ""}`}>
          {active && !isCollapsed && <span className="absolute left-0 top-1.5 bottom-1.5 w-[3px] rounded-full bg-accent-deep" />}
          <Icon size={15} strokeWidth={2} className="shrink-0" />
          {!isCollapsed && <span className="flex-1 whitespace-nowrap overflow-hidden text-ellipsis text-left">{it.label}</span>}
          {!isCollapsed && !!it.badge && (
            <span className={`font-sans text-[10px] font-bold px-1.5 py-0.5 rounded-full ${badgeToneCls(it.badgeTone)}`}>{it.badge}</span>
          )}
        </button>
      );
    };
    return (
      <div className={`h-full flex-shrink-0 bg-white text-inksoft flex flex-col py-5 px-3 border-r border-line transition-[width] duration-200 ${isCollapsed ? "w-[76px]" : "w-[258px]"}`}>
        <div className={`flex items-center gap-2.5 pb-4 mb-1 border-b border-line ${isCollapsed ? "justify-center px-0" : "px-1.5"}`}>
          <div className="w-8 h-8 rounded-[10px] bg-accent flex items-center justify-center shrink-0 font-extrabold text-white text-sm">A</div>
          {!isCollapsed && (
            <div>
              <div className="text-[13px] font-extrabold tracking-tight text-ink">AXE-EXPERTS</div>
              <div className="font-mono text-[9.5px] text-inkmuted">Registre &amp; Pilotage</div>
            </div>
          )}
          {isMobile && (
            <button onClick={() => setMobileOpen?.(false)} className="ml-auto text-inkmuted hover:text-ink">
              <X size={18} />
            </button>
          )}
        </div>
        {!isMobile && (
          <button onClick={() => setCollapsed(!collapsed)} title={isCollapsed ? "Agrandir le menu" : "Réduire le menu"}
            className={`flex items-center gap-1.5 mb-3 bg-transparent border border-line rounded-lg text-inkmuted cursor-pointer hover:text-ink hover:border-accent transition-colors ${isCollapsed ? "justify-center w-full py-1.5" : "justify-end px-2.5 py-1"}`}>
            <ChevronLeft size={14} className={`transition-transform duration-200 ${isCollapsed ? "rotate-180" : ""}`} />
            {!isCollapsed && <span className="text-[11px] font-medium">Réduire</span>}
          </button>
        )}
        <nav className="flex flex-col gap-1 overflow-y-auto">
          <NavButton it={{ id: "dashboard", label: "Vue d'ensemble", icon: LayoutGrid }} />
          <div className="h-1.5" />
          {GROUPS.map((g) => {
            const isOpen = isCollapsed || openGroups.has(g.id);
            return (
              <div key={g.id}>
                {!isCollapsed && (
                  <button onClick={() => toggleGroup(g.id)}
                    className="flex items-center gap-1.5 w-full bg-transparent border-none px-3 pt-2.5 pb-1.5 text-[10px] font-bold tracking-wider uppercase text-inkmuted hover:text-ink transition-colors cursor-pointer">
                    <span className="flex-1 text-left">{g.label}</span>
                    <ChevronDown size={12} className={`transition-transform duration-150 ${isOpen ? "" : "-rotate-90"}`} />
                  </button>
                )}
                {isOpen && (
                  <div className="flex flex-col gap-0.5">
                    {g.items.map((it) => <NavButton key={it.id} it={it} indent={!isCollapsed} />)}
                  </div>
                )}
                {isCollapsed && <div className="h-1.5" />}
              </div>
            );
          })}
        </nav>
        <div className="mt-auto pt-3.5 border-t border-line flex flex-col gap-1.5">
          <div title={mePortefeuille ? `Portefeuille : ${mePortefeuille.nom}` : "Aucun portefeuille"}
            className={`flex items-center gap-2 w-full bg-app border border-line rounded-[11px] text-ink ${isCollapsed ? "justify-center p-1.5" : "justify-start px-2 py-2"}`}>
            <span className="w-6 h-6 rounded-full flex items-center justify-center font-bold text-[11px] text-white shrink-0" style={{ background: meColor }}>{me?.[0]}</span>
            {!isCollapsed && (
              <div className="text-left overflow-hidden">
                <div className="text-[11.5px] font-bold whitespace-nowrap overflow-hidden text-ellipsis">{me}</div>
                <div className="text-[10px] text-inkmuted whitespace-nowrap overflow-hidden text-ellipsis">
                  {ROLE_LABELS[meRole] || meRole}{mePortefeuille ? ` · ${mePortefeuille.nom}` : ""}
                </div>
              </div>
            )}
          </div>
          {onLogout && (
            <button onClick={onLogout} title="Déconnexion"
              className={`flex items-center gap-2 w-full bg-transparent border-none cursor-pointer rounded-[10px] text-red-500 text-[11.5px] font-semibold hover:bg-red-50 transition-colors ${isCollapsed ? "justify-center p-1.5" : "justify-start px-2 py-1.5"}`}>
              <LogOut size={14} />
              {!isCollapsed && <span>Déconnexion</span>}
            </button>
          )}
        </div>
      </div>
    );
  };

  return (
    <>
      {/* Desktop : sidebar fixe */}
      <div className="hidden md:block"><SidebarInner isMobile={false} /></div>
      {/* Mobile : tiroir (drawer) avec overlay */}
      {mobileOpen && (
        <div className="md:hidden fixed inset-0 z-50 flex">
          <div className="absolute inset-0 bg-black/40" onClick={() => setMobileOpen?.(false)} />
          <div className="relative z-10 h-full shadow-2xl"><SidebarInner isMobile={true} /></div>
        </div>
      )}
    </>
  );
}

/* ============================================================
   TOP BAR
   ============================================================ */
function TopBar({ search, setSearch, saveStatus, me, meColor, openTabs, activeTab, onHome, onBack, canGoBack, onSelectTab, onCloseTab, onNav, onOpenClient, onNewClient, clients, notifCount, onOpenMobileMenu, notifications, onMarkNotificationRead, onOpenClient2 }) {
  const [searchOpen, setSearchOpen] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [pickerQuery, setPickerQuery] = useState("");
  const [notifOpen, setNotifOpen] = useState(false);
  const pickerRef = useRef(null);
  const searchInputRef = useRef(null);

  useEffect(() => {
    const onKeyDown = (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        setSearchOpen(true);
        setTimeout(() => searchInputRef.current?.focus(), 0);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  const handleSearchKeyDown = (e) => {
    if (e.key === "Escape") { setSearch(""); setSearchOpen(false); e.currentTarget.blur(); return; }
    if (e.key !== "Enter") return;
    const q = search.trim().toLowerCase();
    if (!q) return;
    // APRÈS
const matches = clients.filter((c) => c.nom.toLowerCase().includes(q) || String(c.siren || "").includes(q));
    if (matches.length >= 1) {
      onOpenClient(matches[0].id);
      setSearch("");
      setSearchOpen(false);
    }
  };

  useEffect(() => {
    if (!pickerOpen) return;
    const onDocClick = (e) => { if (pickerRef.current && !pickerRef.current.contains(e.target)) setPickerOpen(false); };
    document.addEventListener("mousedown", onDocClick);
    return () => document.removeEventListener("mousedown", onDocClick);
  }, [pickerOpen]);

  const pickerResults = useMemo(() => {
    const q = pickerQuery.trim().toLowerCase();
   // APRÈS
const list = q ? clients.filter((c) => c.nom.toLowerCase().includes(q) || String(c.siren || "").includes(q)) : clients;
    return list.slice(0, 40);
  }, [clients, pickerQuery]);

  const toolIcons = [
    { key: "clients", icon: Users, title: "Registre clients", onClick: () => onNav("clients") },
    { key: "tva", icon: Receipt, title: "TVA — CA3/CA12", onClick: () => onNav("tva") },
    { key: "bilans", icon: FileWarning, title: "Bilans", onClick: () => onNav("bilans") },
    { key: "acomptes", icon: Landmark, title: "Acomptes IS / CFE", onClick: () => onNav("acomptes") },
    { key: "age", icon: Building2, title: "AGE / AGO", onClick: () => onNav("age") },
    { key: "mission", icon: ClipboardCheck, title: "Dossiers en accueil", onClick: () => onNav("mission") },
    { key: "fiscal", icon: CalendarDays, title: "Suivi fiscal", onClick: () => onNav("fiscal") },
    { key: "planning", icon: CalendarRange, title: "Mon planning", badge: notifCount, onClick: () => onNav("planning") },
    { key: "equipe", icon: Settings2, title: "Équipe", onClick: () => onNav("equipe") },
  ];

  const unread = (notifications || []).filter((n) => !n.lu).length + (notifCount || 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", flexShrink: 0, background: T.card, borderBottom: `1px solid ${T.line}` }}>
      <div style={{ display: "flex", alignItems: "center", padding: "0 10px", height: 46, gap: 2 }}>
        <button onClick={onOpenMobileMenu} title="Menu" className="md:hidden flex items-center justify-center w-8 h-8 rounded-lg text-inksoft hover:bg-app mr-1 shrink-0">
          <Menu size={18} strokeWidth={2} />
        </button>
        {canGoBack && (
          <button className="topIconBtn" onClick={onBack} title="Retour" style={{ marginRight: 2 }}>
            <ArrowLeft size={16} strokeWidth={2} />
          </button>
        )}
        <button className="topTab" onClick={onHome} title="Accueil" style={{
          background: !activeTab ? T.paperDeep : "transparent", color: !activeTab ? T.navy : T.inkMuted, border: "none", marginRight: 2,
        }}>
          <Home size={15} strokeWidth={2.2} />
        </button>
        <div className="scrollbar" style={{ display: "flex", alignItems: "center", overflowX: "auto", gap: 2, maxWidth: "44vw" }}>
          {openTabs.map((t) => (
            <div key={t.id} className="topTab" onClick={() => onSelectTab(t.id)} style={{
              background: activeTab === t.id ? T.paperDeep : "transparent", color: activeTab === t.id ? T.navy : T.inkSoft,
              borderBottom: activeTab === t.id ? `2px solid ${T.navy}` : "2px solid transparent",
            }}>
              <span style={{ maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.label}</span>
              <X size={12} onClick={(e) => { e.stopPropagation(); onCloseTab(t.id); }} style={{ opacity: 0.6 }} />
            </div>
          ))}
        </div>

        <div ref={pickerRef} style={{ position: "relative" }}>
          <button className="topIconBtn" title="Ouvrir un dossier existant" onClick={() => setPickerOpen((s) => !s)}><Plus size={15} /></button>
          {pickerOpen && (
            <div style={{ position: "absolute", top: 36, left: 0, width: 280, background: T.card, border: `1px solid ${T.line}`, borderRadius: 12, boxShadow: T.shadowLg, zIndex: 30, padding: 8 }}>
              <input autoFocus value={pickerQuery} onChange={(e) => setPickerQuery(e.target.value)} placeholder="Sélectionner un dossier…"
                style={{ width: "100%", padding: "7px 10px", borderRadius: 8, border: `1px solid ${T.line}`, fontSize: 12, marginBottom: 6 }} />
              <div className="scrollbar" style={{ maxHeight: 260, overflowY: "auto" }}>
                {pickerResults.map((c) => (
                  <div key={c.id} className="hoverRow clickable" onClick={() => { onOpenClient(c.id); setPickerOpen(false); setPickerQuery(""); }}
                    style={{ padding: "7px 8px", borderRadius: 7, fontSize: 12.5, display: "flex", justifyContent: "space-between", gap: 8 }}>
                    <span style={{ fontWeight: 600 }}>{c.nom}</span>
                    <span style={{ fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted }}>{c.siren}</span>
                  </div>
                ))}
                {pickerResults.length === 0 && <div style={{ padding: "10px 8px", fontSize: 12, color: T.inkMuted, fontStyle: "italic" }}>Aucun dossier trouvé.</div>}
              </div>
              <div style={{ borderTop: `1px solid ${T.line}`, marginTop: 6, paddingTop: 6 }}>
                <div className="hoverRow clickable" onClick={() => { setPickerOpen(false); setPickerQuery(""); onNewClient && onNewClient(); }}
                  style={{ padding: "7px 8px", borderRadius: 7, fontSize: 12.5, display: "flex", alignItems: "center", gap: 6, color: T.navy, fontWeight: 600 }}>
                  <Plus size={13} /> Nouveau client
                </div>
              </div>
            </div>
          )}
        </div>

        <div className="ml-auto flex items-center gap-0.5">
          <div style={{ display: "flex", alignItems: "center", gap: 4, fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted, marginRight: 8 }}>
            {saveStatus === "saving" && <><Loader2 size={12} className="spin" /> enreg.…</>}
            {saveStatus === "saved" && <><Check size={12} color={T.green} /> enregistré</>}
          </div>
          {searchOpen ? (
            <div className="relative flex-[0_1_260px] hidden sm:block mr-1">
              <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-inkmuted" />
              <input ref={searchInputRef} value={search} onChange={(e) => setSearch(e.target.value)} onKeyDown={handleSearchKeyDown}
                onBlur={() => !search && setSearchOpen(false)}
                placeholder="Rechercher un dossier, un SIREN…"
                className="input-field !rounded-full !py-1.5 !pl-8 !pr-16 !bg-app text-xs w-full" />
              <kbd className="absolute right-2 top-1/2 -translate-y-1/2 text-[9.5px] font-mono text-inkmuted bg-white border border-line rounded px-1.5 py-0.5">Échap</kbd>
            </div>
          ) : (
            <button onClick={() => { setSearchOpen(true); setTimeout(() => searchInputRef.current?.focus(), 0); }} title="Rechercher (⌘K)"
              className="hidden sm:flex items-center gap-2 rounded-full border border-line bg-app px-3 py-1.5 text-xs text-inkmuted hover:border-accent hover:text-accent transition-colors mr-1">
              <Search size={13} />
              <span className="hidden lg:inline">Rechercher…</span>
              <kbd className="font-mono text-[9.5px] bg-white border border-line rounded px-1.5 py-0.5 ml-1">⌘K</kbd>
            </button>
          )}
          <button className="sm:hidden topIconBtn" title="Rechercher" onClick={() => setSearchOpen((s) => !s)}><Search size={16} strokeWidth={1.9} /></button>

          <div className="relative hidden md:block">
            <button onClick={() => setNotifOpen((s) => !s)} title="Notifications" className="topIconBtn"><Bell size={16} strokeWidth={1.9} />
              {!!unread && <span className="absolute top-1 right-1 bg-badge-red-text text-white text-[9px] font-bold rounded-full px-[4px] leading-[13px]">{unread}</span>}
            </button>
            {notifOpen && (
              <div className="absolute right-0 top-9 w-80 card p-3 z-30 max-h-96 overflow-y-auto scrollbar">
                <div className="text-xs font-bold text-ink mb-2">Notifications</div>
                {(notifications || []).length === 0 && !notifCount && (
                  <div className="text-xs text-inkmuted italic px-2 py-1">Aucune notification pour le moment.</div>
                )}
                {!!notifCount && (
                  <div onClick={() => { onNav("planning"); setNotifOpen(false); }} className="hoverRow clickable text-xs text-inksoft rounded-lg p-2 cursor-pointer mb-1">
                    {notifCount} échéance{notifCount > 1 ? "s" : ""} en retard sur votre planning
                  </div>
                )}
                {(notifications || []).map((n) => (
                  <div key={n.id} onClick={() => { onMarkNotificationRead?.(n.id); if (n.client_id && onOpenClient2) onOpenClient2(n.client_id); setNotifOpen(false); }}
                    className="hoverRow clickable text-xs rounded-lg p-2 cursor-pointer flex items-start gap-2">
                    <span className={`w-1.5 h-1.5 rounded-full mt-1 shrink-0 ${n.lu ? "bg-line" : "bg-badge-red-text"}`} />
                    <div>
                      <div className={n.lu ? "text-inkmuted" : "text-ink font-medium"}>{n.message}</div>
                      <div className="text-[10px] text-inkmuted mt-0.5">{new Date(n.created_at).toLocaleString("fr-FR", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}</div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>

          {toolIcons.map((ic) => {
            const Icon = ic.icon;
            return (
              <button key={ic.key} className="hidden md:inline-flex topIconBtn" title={ic.title} onClick={ic.onClick}>
                <Icon size={16} strokeWidth={1.9} />
                {!!ic.badge && (
                  <span style={{ position: "absolute", top: 2, right: 2, background: T.amber, color: "#fff", fontSize: 9, fontWeight: 700, borderRadius: 999, padding: "0 4px", lineHeight: "13px" }}>{ic.badge}</span>
                )}
              </button>
            );
          })}
          <span className="w-7 h-7 rounded-full flex items-center justify-center font-bold text-[11px] text-white ml-1.5 shrink-0" style={{ background: meColor }}>{me?.[0]}</span>
        </div>
      </div>
      <style>{`.spin{animation:spin .8s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}

/* ============================================================
   HELPERS
   ============================================================ */
function isTvaLate(client) {
  if (!client.tvaRegime || client.tvaRegime === "FRANCHISE") return false;
  return effectiveTvaStatus(client, currentMonthKey()) === "RETARD";
}
function seuilEffectifAlert(effectif) {
  const n = parseInt(effectif, 10);
  if (!n) return null;
  if (n >= 50) return { label: "≥ 50 salariés", tone: "red" };
  if (n >= 45) return { label: "Approche 50", tone: "amber" };
  if (n >= 11) return { label: "≥ 11 salariés", tone: "amber" };
  if (n >= 9) return { label: "Approche 11", tone: "amber" };
  return null;
}
function missionCompletion(client) {
  const m = client.mission || {};
  const done = MISSION_ALL_KEYS.filter((k) => m[k]).length;
  const total = MISSION_ALL_KEYS.length;
  return { done, total, pct: total ? Math.round((done / total) * 100) : 100 };
}
function isBilanLate(client) {
  const b = client.bilan || {};
  if (b.transmis) return false;
  const echeance = getBilanEcheance(client.dateCloture);
  return !!(echeance && todayISO() > echeance);
}
function computeCounts(clients) {
  const total = clients.length;
  const tvaAlert = clients.filter(isTvaLate).length;
  const bilanRetard = clients.filter(isBilanLate).length;
  const missionIncomplete = clients.filter((c) => { const m = missionCompletion(c); return m && m.pct < 100; }).length;
  const ageAlert = clients.filter((c) => Object.values(c.ageAgoHistory || {}).some((y) => y.capitauxInf || y.ageContinuite)).length;
  return { total, tvaAlert, bilanRetard, missionIncomplete, ageAlert };
}
function filterByRole(clients, me, roleFilter) {
  if (roleFilter === "Collaborateur") return clients.filter((c) => c.collab === me);
  if (roleFilter === "Expert") return clients.filter((c) => c.expert === me);
  if (roleFilter === "Chef de mission") return clients.filter((c) => c.chefMission === me);
  return clients;
}
function filterClients(clients, search, roleFilter, me, regimeFilter, statutFilter = "actif") {
  let out = filterByRole(clients, me, roleFilter || "Tous");
  if (statutFilter === "actif") out = out.filter((c) => (c.statutDossier || "actif") === "actif");
  else if (statutFilter === "transfert") out = out.filter((c) => c.statutDossier === "transfert");
  else if (statutFilter === "inactif") out = out.filter((c) => c.statutDossier === "inactif");
  // statutFilter === "tous" -> pas de filtre supplémentaire
  if (regimeFilter && regimeFilter !== "Tous") out = out.filter((c) => c.tvaRegime === regimeFilter);
  // APRÈS
if (search.trim()) {
    const q = search.trim().toLowerCase();
    out = out.filter((c) => c.nom.toLowerCase().includes(q) || String(c.siren || "").includes(q));
}
  return out;
}
function Stamped({ tone = "green", children, small }) {
  const map = {
    green: "bg-badge-green-bg text-badge-green-text",
    red: "bg-badge-red-bg text-badge-red-text",
    amber: "bg-badge-amber-bg text-badge-amber-text",
    purple: "bg-badge-purple-bg text-badge-purple-text",
    neutral: "bg-badge-slate-bg text-badge-slate-text",
  };
  return (
    <span className={`badge-pill ${small ? "!text-[9.5px] !px-2 !py-0.5" : ""} ${map[tone] || map.neutral}`}>{children}</span>
  );
}
function RoleBadge({ role, name }) {
  if (!name) return null;
  return <span className="text-[11px] text-inkmuted">{role}: <strong className="text-inksoft">{name}</strong></span>;
}

/* ============================================================
   FILTER BAR — barre d'outils compacte sur une seule ligne
   (recherche + rôle + régime TVA + statut du dossier)
   ============================================================ */
const ROLE_FILTER_OPTIONS = [
  { value: "Tous", label: "Tous les rôles" },
  { value: "Collaborateur", label: "Collaborateur" },
  { value: "Expert", label: "Expert" },
  { value: "Chef de mission", label: "Chef de mission" },
];
const STATUT_FILTER_OPTIONS = [
  { value: "actif", label: "Actifs uniquement" },
  { value: "transfert", label: "En transfert" },
  { value: "inactif", label: "Inactifs" },
  { value: "tous", label: "Tous les dossiers" },
];
function FilterBar({ roleFilter, setRoleFilter, count, regimeFilter, setRegimeFilter, statutFilter, setStatutFilter, search, setSearch }) {
  const selectCls = "input-field !py-1.5 !w-auto text-xs md:text-[13px] font-medium cursor-pointer";
  return (
    <div className="flex items-center gap-2 flex-wrap mb-4 py-0.5">
      {setSearch && (
        <div className="relative flex-1 min-w-[160px] max-w-[260px]">
          <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-inkmuted" />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Rechercher un dossier…"
            className="input-field !py-1.5 !pl-8 text-xs md:text-[13px] w-full" />
        </div>
      )}
      <select value={roleFilter || "Tous"} onChange={(e) => setRoleFilter(e.target.value)} className={selectCls} title="Filtrer par mon rôle">
        {ROLE_FILTER_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.value === "Tous" ? o.label : `Mon rôle : ${o.label}`}</option>)}
      </select>
      {setRegimeFilter && (
        <select value={regimeFilter || "Tous"} onChange={(e) => setRegimeFilter(e.target.value)} className={`${selectCls} font-mono`} title="Filtrer par régime TVA">
          <option value="Tous">Régime TVA : Tous</option>
          {REGIMES_TVA.map((r) => <option key={r} value={r}>Régime TVA : {r}</option>)}
        </select>
      )}
      {setStatutFilter && (
        <select value={statutFilter || "actif"} onChange={(e) => setStatutFilter(e.target.value)} className={selectCls} title="Filtrer par statut du dossier">
          {STATUT_FILTER_OPTIONS.map((o) => <option key={o.value} value={o.value}>Statut : {o.label}</option>)}
        </select>
      )}
      <span className="ml-auto font-mono text-[11.5px] text-inkmuted whitespace-nowrap">{count} dossier(s)</span>
    </div>
  );
}

/* ============================================================
   DASHBOARD — ANALYTIQUES VISUELLES
   ============================================================ */

const DASHBOARD_CHART_COLORS = [
  "#2563EB",
  "#14B8A6",
  "#8B5CF6",
  "#F59E0B",
  "#F43F5E",
  "#06B6D4",
  "#64748B",
  "#22C55E",
];

function DonutDistribution({ title, items, total, icon: Icon = CircleDot, onItemClick }) {
  const safeItems = (items || []).filter((x) => Number(x.value) > 0);

  const sum = safeItems.reduce(
    (acc, x) => acc + Number(x.value),
    0
  );

  let cursor = 0;

  const stops = safeItems.map((item, i) => {
    const pct = (Number(item.value) / (sum || 1)) * 100;
    const start = cursor;

    cursor += pct;

    return `${item.color || DASHBOARD_CHART_COLORS[i % DASHBOARD_CHART_COLORS.length]} ${start.toFixed(2)}% ${cursor.toFixed(2)}%`;
  });

  const shownTotal = total ?? sum;

  return (
    <Panel
      title={title}
      right={<Icon size={15} color={T.inkMuted} />}
    >
      {!safeItems.length ? (
        <EmptyNote text="Aucune donnée renseignée pour le moment." />
      ) : (
        <div className="grid grid-cols-[118px_1fr] sm:grid-cols-[136px_1fr] gap-4 items-center">

          <div
            style={{
              position: "relative",
              width: 118,
              height: 118,
              margin: "0 auto",
            }}
          >
            <div
              aria-label={`${shownTotal} dossiers`}
              style={{
                width: "100%",
                height: "100%",
                borderRadius: "50%",
                background: `conic-gradient(${stops.join(", ")})`,
                boxShadow:
                  "inset 0 0 0 1px rgba(15,23,42,.04)",
              }}
            />

            <div
              style={{
                position: "absolute",
                inset: 19,
                borderRadius: "50%",
                background: T.card,
                border: `1px solid ${T.line}`,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
              }}
            >
              <strong
                style={{
                  fontFamily: T.serif,
                  fontSize: 20,
                  lineHeight: 1,
                  color: T.ink,
                }}
              >
                {shownTotal}
              </strong>

              <span
                style={{
                  fontSize: 9.5,
                  color: T.inkMuted,
                  marginTop: 5,
                }}
              >
                dossiers
              </span>
            </div>
          </div>

          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 8,
              minWidth: 0,
            }}
          >
            {safeItems.slice(0, 7).map((item, i) => {
              const pct =
                (Number(item.value) / (sum || 1)) * 100;

              const color =
                item.color ||
                DASHBOARD_CHART_COLORS[
                  i % DASHBOARD_CHART_COLORS.length
                ];

              return (
                <div
                  key={`${item.label}-${i}`}
                  onClick={() => onItemClick?.(item)}
                  className={onItemClick ? "clickable" : ""}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 7,
                    minWidth: 0,
                    cursor: onItemClick ? "pointer" : "default",
                    padding: "4px 5px",
                    borderRadius: 8,
                  }}
                >
                  <span
                    style={{
                      width: 8,
                      height: 8,
                      borderRadius: 99,
                      background: color,
                      flexShrink: 0,
                    }}
                  />

                  <span
                    style={{
                      flex: 1,
                      minWidth: 0,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                      fontSize: 10.5,
                      color: T.inkSoft,
                    }}
                  >
                    {item.label}
                  </span>

                  <strong
                    style={{
                      fontSize: 10.5,
                      color: T.ink,
                      fontFamily: T.mono,
                    }}
                  >
                    {pct.toFixed(pct >= 10 ? 0 : 1)}%
                  </strong>
                </div>
              );
            })}

            {safeItems.length > 7 && (
              <span
                style={{
                  fontSize: 9.5,
                  color: T.inkMuted,
                }}
              >
                + {safeItems.length - 7} autres catégories
              </span>
            )}
          </div>
        </div>
      )}
    </Panel>
  );
}

function HorizontalDistribution({
  title,
  items,
  icon: Icon = Users,
  onItemClick,
}) {
  const safe = (items || []).filter(
    (x) => Number(x.value) > 0
  );

  const max = Math.max(
    1,
    ...safe.map((x) => Number(x.value))
  );

  return (
    <Panel
      title={title}
      right={<Icon size={15} color={T.inkMuted} />}
    >
      {!safe.length ? (
        <EmptyNote text="Aucune donnée disponible." />
      ) : (
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 11,
          }}
        >
          {safe.slice(0, 8).map((item, i) => {
            const pct =
              (Number(item.value) / max) * 100;

            const color =
              item.color ||
              DASHBOARD_CHART_COLORS[
                i % DASHBOARD_CHART_COLORS.length
              ];

            return (
              <div
                key={`${item.label}-${i}`}
                onClick={() => onItemClick?.(item)}
                className={onItemClick ? "clickable" : ""}
                style={{ cursor: onItemClick ? "pointer" : "default", padding: "4px 5px", borderRadius: 8 }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    gap: 10,
                    marginBottom: 5,
                    fontSize: 10.5,
                  }}
                >
                  <span
                    style={{
                      color: T.inkSoft,
                      fontWeight: 600,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    }}
                  >
                    {item.label}
                  </span>

                  <strong
                    style={{
                      color: T.ink,
                      fontFamily: T.mono,
                    }}
                  >
                    {item.value}
                  </strong>
                </div>

                <div
                  style={{
                    height: 7,
                    borderRadius: 999,
                    background: T.paperDeep,
                    overflow: "hidden",
                  }}
                >
                  <div
                    style={{
                      width: `${pct}%`,
                      height: "100%",
                      borderRadius: 999,
                      background: color,
                      transition: "width .35s ease",
                    }}
                  />
                </div>
              </div>
            );
          })}
        </div>
      )}
    </Panel>
  );
}

function buildDistribution(
  clients,
  keyFn,
  labelFn,
  colorFn
) {
  const map = new Map();

  clients.forEach((client) => {
    const key =
      keyFn(client) || "non_renseigne";

    map.set(
      key,
      (map.get(key) || 0) + 1
    );
  });

  return [...map.entries()]
    .sort((a, b) => b[1] - a[1])
    .map(([key, value], i) => ({
      key,
      label: labelFn(key),
      value,
      color:
        colorFn?.(key, i) ||
        DASHBOARD_CHART_COLORS[
          i % DASHBOARD_CHART_COLORS.length
        ],
    }));
}

function inferLegalForm(client) {
  if (client.formeJuridique) {
    return client.formeJuridique;
  }

  const text =
    `${client.nom || ""} ${client.raisonSociale || ""}`
      .toUpperCase();

  const forms = [
    "SASU",
    "SARL",
    "EURL",
    "SAS",
    "SCI",
    "SELARL",
    "SNC",
    "SA",
  ];

  return (
    forms.find(
      (f) => new RegExp(`\\b${f}\\b`).test(text)
    ) || "Non renseignée"
  );
}
/* ============================================================
   DASHBOARD
   ============================================================ */

/* ============================================================
   PILOTAGE CABINET — vue chef de mission : risques, demandes,
   validations, charge et rentabilité.
   ============================================================ */
function PilotageView({ clients, tasks, team, me, onOpenClient, onView }) {
  const active = clients.filter((c) => c.statutDossier !== "inactif");
  const anomalies = useMemo(() => detectAllAnomalies(active), [active]);
  const requests = active.flatMap((c) => (c.demandesClient || []).map((r) => ({ ...r, clientId: c.id, clientNom: c.nom }))).filter((r) => r.statut !== "controle");
  const validations = active.filter((c) => c.validationDossier?.collaborateur && !c.validationDossier?.chefMission);
  const relances = requests.filter((r) => r.relanceLe && new Date(r.relanceLe) <= new Date());
  const workload = useMemo(() => team.map((m) => {
    const dossiers = active.filter((c) => c.collab === m.nom).length;
    const taches = (tasks || []).filter((t) => t.assignee_id === m.id || t.assignee === m.nom || t.collaborateur === m.nom).filter((t) => t.statut !== "termine").length;
    return { ...m, dossiers, taches };
  }).sort((a,b) => (b.taches + b.dossiers) - (a.taches + a.dossiers)), [team, active, tasks]);
  const rentabilite = active.filter((c) => Number(c.rentabilite?.tempsPrevu) > 0 && Number(c.rentabilite?.tempsReel) > Number(c.rentabilite?.tempsPrevu) * 1.25);
  const critical = anomalies.filter((a) => a.gravite === "haute");
  return <div>
    <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Pilotage cabinet</h1></Reveal>
    <p style={{ color: T.inkMuted, fontSize: 12.5, margin: "0 0 18px" }}>Le cockpit du chef de mission : ce qui nécessite une action, une relance ou une validation.</p>
    <MobileKpiSummary
      title="Pilotage cabinet"
      items={[
        { label: "Priorités", value: critical.length, tone: critical.length ? "red" : "green", onClick: () => onView("surveillance") },
        { label: "Pièces à relancer", value: relances.length, tone: relances.length ? "amber" : "green" },
        { label: "À valider CDM", value: validations.length, tone: validations.length ? "amber" : "green" },
        { label: "Dossiers à risque", value: rentabilite.length, tone: rentabilite.length ? "amber" : "green" },
        { label: "Dossiers actifs", value: active.length, tone: "neutral", onClick: () => onView("clients") },
      ]}
    />
    <div className="hidden md:grid grid-cols-2 md:grid-cols-5 gap-3" style={{ marginBottom: 18 }}>
      <KpiCard label="Priorités" value={critical.length} icon={ShieldAlert} tone={critical.length ? "red" : "green"} onClick={() => onView("surveillance")} />
      <KpiCard label="Pièces à relancer" value={relances.length} icon={Mail} tone={relances.length ? "amber" : "green"} />
      <KpiCard label="À valider CDM" value={validations.length} icon={Check} tone={validations.length ? "amber" : "green"} />
      <KpiCard label="Dossiers à risque" value={rentabilite.length} icon={Clock3} tone={rentabilite.length ? "amber" : "green"} />
      <KpiCard label="Dossiers actifs" value={active.length} icon={Users} tone="neutral" onClick={() => onView("clients")} />
    </div>
    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
      <Panel title="Actions prioritaires" right={<Stamped tone={critical.length ? "red" : "green"} small>{critical.length}</Stamped>}>
        {critical.slice(0,7).map(a => <div key={a.id} className="hoverRow clickable" onClick={() => onOpenClient(a.clientId)} style={{display:"flex",gap:9,alignItems:"center",padding:"9px 5px",borderBottom:`1px solid ${T.line}`}}><span style={{width:7,height:7,borderRadius:9,background:T.red}}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:11.5}}>{a.clientNom}</div><div style={{fontSize:10.5,color:T.inkMuted}}>{a.message}</div></div><Stamped tone="red" small>Prioritaire</Stamped></div>)}
        {!critical.length && <EmptyNote text="Aucune priorité critique." />}
      </Panel>
      <Panel title="Pièces et relances" right={<Stamped tone={relances.length ? "amber" : "green"} small>{relances.length}</Stamped>}>
        {relances.slice(0,7).map(r => <div key={`${r.clientId}-${r.id}`} className="hoverRow clickable" onClick={() => onOpenClient(r.clientId)} style={{display:"flex",gap:9,alignItems:"center",padding:"9px 5px",borderBottom:`1px solid ${T.line}`}}><Mail size={13} color={T.amber}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:11.5}}>{r.clientNom}</div><div style={{fontSize:10.5,color:T.inkMuted}}>{r.libelle || "Pièce demandée"}</div></div><Stamped tone="amber" small>Relancer</Stamped></div>)}
        {!relances.length && <EmptyNote text="Aucune relance à effectuer." />}
      </Panel>
      <Panel title="Validations chef de mission" right={<Stamped tone={validations.length ? "amber" : "green"} small>{validations.length}</Stamped>}>
        {validations.slice(0,7).map(c => <div key={c.id} className="hoverRow clickable" onClick={() => onOpenClient(c.id)} style={{display:"flex",gap:9,alignItems:"center",padding:"9px 5px",borderBottom:`1px solid ${T.line}`}}><Check size={13} color={T.amber}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:11.5}}>{c.nom}</div><div style={{fontSize:10.5,color:T.inkMuted}}>Le collaborateur a terminé sa partie</div></div><Stamped tone="amber" small>À valider</Stamped></div>)}
        {!validations.length && <EmptyNote text="Aucun dossier en attente de validation." />}
      </Panel>
      <Panel title="Charge de l'équipe">
        {workload.slice(0,8).map(m => <div key={m.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 4px",borderBottom:`1px solid ${T.line}`}}><div style={{width:30,height:30,borderRadius:9,background:m.color||T.navySoft,color:m.color?"#fff":T.navy,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:800}}>{(m.nom||"?").slice(0,2).toUpperCase()}</div><div style={{flex:1}}><div style={{fontWeight:700,fontSize:11.5}}>{m.nom}</div><div style={{fontSize:10.5,color:T.inkMuted}}>{m.dossiers} dossier{m.dossiers>1?"s":""} · {m.taches} tâche{m.taches>1?"s":""} active{m.taches>1?"s":""}</div></div><Stamped tone={m.taches>12?"red":m.taches>7?"amber":"green"} small>{m.taches>12?"Surchargé":m.taches>7?"À surveiller":"OK"}</Stamped></div>)}
        {!workload.length && <EmptyNote text="Aucune donnée d'équipe." />}
      </Panel>
    </div>
    <div style={{height:16}} />
    <Panel title="Dossiers dont le temps réel dépasse le prévu">
      {rentabilite.slice(0,8).map(c => <div key={c.id} className="hoverRow clickable" onClick={() => onOpenClient(c.id)} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 5px",borderBottom:`1px solid ${T.line}`}}><Clock3 size={13} color={T.amber}/><div style={{flex:1}}><div style={{fontWeight:700,fontSize:11.5}}>{c.nom}</div><div style={{fontSize:10.5,color:T.inkMuted}}>Prévu {c.rentabilite.tempsPrevu} h · Réel {c.rentabilite.tempsReel} h</div></div><Stamped tone="amber" small>Rentabilité</Stamped></div>)}
      {!rentabilite.length && <EmptyNote text="Aucun dossier ne dépasse actuellement le seuil de 25 %." />}
    </Panel>
  </div>;
}

/* ============================================================
   DEMANDES CLIENT / PIÈCES — suivi simple sans imposer de GED.
   ============================================================ */
function DemandesPiecesTab({ client, onUpdate }) {
  const [label, setLabel] = useState("");
  const [relanceLe, setRelanceLe] = useState("");
  const demandes = client.demandesClient || [];
  const add = () => {
    if (!label.trim()) return;
    const item = { id: uid(), libelle: label.trim(), demandeLe: todayISO(), relanceLe: relanceLe || "", statut: "demande", note: "" };
    onUpdate(client.id, { demandesClient: [...demandes, item] }); setLabel(""); setRelanceLe("");
  };
  const patch = (id, patch) => onUpdate(client.id, { demandesClient: demandes.map(d => d.id === id ? { ...d, ...patch } : d) });
  const remove = (id) => onUpdate(client.id, { demandesClient: demandes.filter(d => d.id !== id) });
  const counts = { demande: demandes.filter(d=>d.statut==="demande").length, recu: demandes.filter(d=>d.statut==="recu").length, controle: demandes.filter(d=>d.statut==="controle").length };
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><h4 style={{fontFamily:T.serif,fontSize:13,color:T.navy,margin:0}}>Demandes clients & pièces</h4><div style={{display:"flex",gap:5}}><Stamped tone={counts.demande?"amber":"green"} small>{counts.demande} demandée{counts.demande>1?"s":""}</Stamped><Stamped tone="green" small>{counts.recu} reçue{counts.recu>1?"s":""}</Stamped></div></div>
    <div style={{display:"flex",gap:7,flexWrap:"wrap",marginBottom:16}}><input value={label} onChange={e=>setLabel(e.target.value)} placeholder="Ex. Relevé bancaire de juillet" style={{padding:"6px 8px",borderRadius:8,border:`1px solid ${T.line}`,fontSize:11.5,width:250}}/><input type="date" value={relanceLe} onChange={e=>setRelanceLe(e.target.value)} style={{padding:"6px 8px",borderRadius:8,border:`1px solid ${T.line}`,fontSize:11.5}}/><button onClick={add} style={{border:"none",background:T.navy,color:"white",borderRadius:8,padding:"7px 11px",fontSize:11.5,fontWeight:700,cursor:"pointer"}}><Plus size={13} style={{verticalAlign:"-2px"}}/> Demander</button></div>
    {!demandes.length ? <EmptyNote text="Aucune pièce demandée. Ajoute ici les éléments attendus du client."/> : <div style={{display:"flex",flexDirection:"column",gap:7}}>{demandes.map(d=><div key={d.id} style={{border:`1px solid ${T.line}`,borderRadius:T.radiusSm,padding:"9px 10px",background:T.paper}}><div style={{display:"flex",alignItems:"center",gap:8}}><div style={{flex:1,fontWeight:700,fontSize:11.5}}>{d.libelle}</div><SelectPill value={d.statut} options={["demande","recu","controle"]} onChange={v=>patch(d.id,{statut:v})}/><button onClick={()=>remove(d.id)} style={{border:"none",background:"none",color:T.inkMuted,cursor:"pointer"}}><Trash2 size={13}/></button></div><div style={{fontSize:10.5,color:T.inkMuted,marginTop:5}}>Demandé le {fmtFR(d.demandeLe)}{d.relanceLe?` · Relance ${fmtFR(d.relanceLe)}`:""}</div></div>)}</div>}
  </div>;
}


function RentabiliteTab({ client, onUpdate }) {
  const r = client.rentabilite || { tempsPrevu:"", tempsReel:"", tarifHoraire:"", margeCible:"" };
  const patch = (p) => onUpdate(client.id, { rentabilite: { ...r, ...p } });
  const prev = Number(r.tempsPrevu), real = Number(r.tempsReel), rate = Number(r.tarifHoraire);
  const depassement = prev > 0 && real > prev ? ((real - prev) / prev) * 100 : 0;
  const ca = real > 0 && rate > 0 ? real * rate : 0;
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><h4 style={{fontFamily:T.serif,fontSize:13,color:T.navy,margin:0}}>Temps & rentabilité du dossier</h4><Stamped tone={depassement>25?"red":depassement>0?"amber":"green"} small>{depassement>0?`+${depassement.toFixed(0)} % vs prévu`:"Dans le prévu"}</Stamped></div>
    <FieldRow label="Temps prévu (h)"><TextInput defaultValue={r.tempsPrevu} onCommit={v=>patch({tempsPrevu:v})} placeholder="ex. 12" width={120}/></FieldRow>
    <FieldRow label="Temps réel (h)"><TextInput defaultValue={r.tempsReel} onCommit={v=>patch({tempsReel:v})} placeholder="ex. 14,5" width={120}/></FieldRow>
    <FieldRow label="Taux horaire (€)"><TextInput defaultValue={r.tarifHoraire} onCommit={v=>patch({tarifHoraire:v})} placeholder="ex. 85" width={120}/></FieldRow>
    <FieldRow label="Marge cible (%)"><TextInput defaultValue={r.margeCible} onCommit={v=>patch({margeCible:v})} placeholder="ex. 35" width={120}/></FieldRow>
    <div className="grid grid-cols-2 gap-2.5" style={{marginTop:16}}>
      <div style={{border:`1px solid ${T.line}`,borderRadius:T.radiusSm,padding:"10px 11px",background:T.paper}}><div style={{fontSize:9.5,color:T.inkMuted,textTransform:"uppercase",fontWeight:700}}>Heures</div><div style={{fontSize:17,fontWeight:800,fontFamily:T.mono,marginTop:3}}>{real||0} / {prev||0}</div></div>
      <div style={{border:`1px solid ${T.line}`,borderRadius:T.radiusSm,padding:"10px 11px",background:T.paper}}><div style={{fontSize:9.5,color:T.inkMuted,textTransform:"uppercase",fontWeight:700}}>Valorisation</div><div style={{fontSize:17,fontWeight:800,fontFamily:T.mono,marginTop:3}}>{ca?fmtEUR(ca):"—"}</div></div>
    </div>
  </div>;
}

function ValidationDossierTab({ client, onUpdate, me }) {
  const v = client.validationDossier || { collaborateur:false, chefMission:false, dateCollaborateur:"", dateChefMission:"", commentaire:"" };
  const complete = !!v.collaborateur && !!v.chefMission;
  return <div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><h4 style={{fontFamily:T.serif,fontSize:13,color:T.navy,margin:0}}>Validation de fin de dossier</h4><Stamped tone={complete?"green":"amber"} small>{complete?"Dossier validé":"Validation en attente"}</Stamped></div>
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 11px",border:`1px solid ${T.line}`,borderRadius:T.radiusSm}}><ToggleBtn on={!!v.collaborateur} onClick={()=>onUpdate(client.id,{validationDossier:{...v,collaborateur:!v.collaborateur,dateCollaborateur:!v.collaborateur?todayISO():v.dateCollaborateur}})}/><div><div style={{fontWeight:700,fontSize:11.5}}>Collaborateur — dossier terminé</div><div style={{fontSize:10.5,color:T.inkMuted}}>{v.dateCollaborateur?`Le ${fmtFR(v.dateCollaborateur)}`:"Non validé"}</div></div></div>
      <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 11px",border:`1px solid ${T.line}`,borderRadius:T.radiusSm}}><ToggleBtn on={!!v.chefMission} onClick={()=>onUpdate(client.id,{validationDossier:{...v,chefMission:!v.chefMission,dateChefMission:!v.chefMission?todayISO():v.dateChefMission}})}/><div><div style={{fontWeight:700,fontSize:11.5}}>Chef de mission — validation finale</div><div style={{fontSize:10.5,color:T.inkMuted}}>{v.dateChefMission?`Validé le ${fmtFR(v.dateChefMission)} par ${me||"le CDM"}`:"À valider"}</div></div></div>
      <FieldRow label="Commentaire"><TextInput defaultValue={v.commentaire} onCommit={x=>onUpdate(client.id,{validationDossier:{...v,commentaire:x}})} placeholder="Réserve ou remarque de clôture" width={360}/></FieldRow>
    </div>
  </div>;
}

function Dashboard({
  myClients,
  tasks,
  me,
  meRole,
  onOpenClient,
  setView,
  team,
  onSuperviseClick,
  onDashboardFilter
}) {
    const today = new Date();
  const dateStr = today.toLocaleDateString("fr-FR", {
    weekday: "long",
    day: "numeric",
    month: "long"
  });
  const counts = computeCounts(myClients);
  const anomalies = useMemo(() => detectAllAnomalies(myClients), [myClients]);
  const criticalAnomalies = anomalies.filter((a) => a.gravite === "haute");
  const importantAnomalies = anomalies.filter((a) => a.gravite !== "haute");
  const [taskFilter, setTaskFilter] = useState("Toutes");
  const buckets = ["retard", "aujourdhui", "demain", "semaine", "mois", "trimestre"];
  const filteredTasks = taskFilter === "Toutes" ? tasks.filter(t => buckets.includes(t.bucket)) : tasks.filter((t) => t.bucket === taskFilter);
  const sortedTasks = [...filteredTasks].sort((a, b) => a.date - b.date);

  const roleCounts = {
    Collaborateur: myClients.filter((c) => c.collab === me).length,
    Expert: myClients.filter((c) => c.expert === me).length,
    "Chef de mission": myClients.filter((c) => c.chefMission === me).length,
  };
  const maxRole = Math.max(1, ...Object.values(roleCounts));

  const prevKey = previousMonthKey();
  const nonRapprochesM1 = myClients.filter((c) => {
    const v = (c.revision?.banqueMois?.[prevKey] || "").toUpperCase();
    return v !== "FAIT" && v !== "NA";
  });

  // Widget Chef de mission : uniquement les dossiers où "me" est chef de mission, groupés par collaborateur.
  const superviseClients = myClients.filter((c) => c.chefMission === me);
  const byCollab = {};
  superviseClients.forEach((c) => {
    const key = c.collab || "Non assigné";
    if (!byCollab[key]) byCollab[key] = { total: 0, bilanRetard: 0, tvaAlert: 0 };
    byCollab[key].total += 1;
    if (isBilanLate(c)) byCollab[key].bilanRetard += 1;
    if (isTvaLate(c)) byCollab[key].tvaAlert += 1;
  });

  const sectorItems = useMemo(
  () =>
    buildDistribution(
      myClients,
      (c) =>
        c.secteur ||
        classifyActivite(c.activite),
      (key) =>
        SECTEURS_ACTIVITE.find(
          (s) => s.id === key
        )?.label || "Non classé",
      (key) =>
        SECTEURS_ACTIVITE.find(
          (s) => s.id === key
        )?.color
    ),
  [myClients]
);

const legalItems = useMemo(
  () =>
    buildDistribution(
      myClients,
      inferLegalForm,
      (key) => key
    ),
  [myClients]
);

const statusItems = useMemo(
  () =>
    buildDistribution(
      myClients,
      (c) =>
        c.statutDossier || "actif",
      (key) =>
        ({
          actif: "Actifs",
          transfert: "En transfert",
          inactif: "Inactifs",
        }[key] || key),
      (key) =>
        ({
          actif: T.green,
          transfert: T.amber,
          inactif: T.inkMuted,
        }[key])
    ),
  [myClients]
);

const tvaItems = useMemo(() => {
  const key = currentMonthKey();

  const relevant = myClients.filter(
    (c) =>
      c.tvaRegime &&
      c.tvaRegime !== "FRANCHISE"
  );

  const counts = {
    OK: 0,
    FAIT: 0,
    NON_VALIDE: 0,
    RETARD: 0,
    ATTENTE: 0,
    NA: 0,
  };

  relevant.forEach((c) => {
    const status =
      effectiveTvaStatus(c, key);

    if (status === "OK") counts.OK++;
    else if (status === "FAIT") counts.FAIT++;
    else if (status === "NON_VALIDE") counts.NON_VALIDE++;
    else if (status === "RETARD") counts.RETARD++;
    else if (status === "NA") counts.NA++;
    else counts.ATTENTE++;
  });

  return [
    {
      key: "OK",
      label: "Déclarées",
      value: counts.OK,
      color: T.green,
    },
    {
      key: "FAIT",
      label: "Préparées",
      value: counts.FAIT,
      color: T.amber,
    },
    {
      key: "NON_VALIDE",
      label: "À corriger",
      value: counts.NON_VALIDE,
      color: "#6D28D9",
    },
    {
      key: "RETARD",
      label: "En retard",
      value: counts.RETARD,
      color: T.red,
    },
    {
      key: "ATTENTE",
      label: "En attente",
      value: counts.ATTENTE,
      color: T.navy,
    },
    {
      key: "NA",
      label: "N/A",
      value: counts.NA,
      color: T.inkMuted,
    },
  ];
}, [myClients]);

const collaboratorItems = useMemo(
  () =>
    buildDistribution(
      myClients,
      (c) =>
        c.collab || "Non assigné",
      (key) => key,
      (_key, i) =>
        DASHBOARD_CHART_COLORS[
          i % DASHBOARD_CHART_COLORS.length
        ]
    ),
  [myClients]
);

  // Échéances à venir : regroupe les tâches par horizon, avec l'intitulé de la plus proche
  const upcomingGroups = [
    { key: "demain", label: "Demain" },
    { key: "semaine", label: "Cette semaine" },
    { key: "mois", label: "Ce mois-ci" },
    { key: "trimestre", label: "Ce trimestre" },
  ].map((g) => {
    const items = tasks.filter((t) => t.bucket === g.key).sort((a, b) => a.date - b.date);
    return { ...g, count: items.length, next: items[0] };
  });

  const taskToneStyle = (bucket) =>
    bucket === "retard" ? { color: T.red, bg: T.redSoft }
    : bucket === "aujourdhui" ? { color: T.amber, bg: T.amberSoft }
    : bucket === "demain" ? { color: T.navy, bg: T.navySoft }
    : { color: T.inkMuted, bg: T.paperDeep };
  const taskCategoryIcon = (category) =>
    category === "TVA" ? Receipt
    : category === "IS" ? Wallet
    : category === "CFE" ? Landmark
    : category === "Bilan" || category === "Clôture" ? FileWarning
    : category === "AGO" ? Building2
    : category === "Accueil" ? ClipboardCheck
    : CalendarDays;

  return (
    <div>
      <Reveal>
        <div style={{ marginBottom: 26 }}>
          <h1 style={{ fontFamily: T.serif, fontSize: 24, fontWeight: 800, color: T.ink, margin: 0 }}>Vue d'ensemble</h1>
          <div style={{ fontSize: 12.5, color: T.inkMuted, marginTop: 6 }}>
            Tableau de bord · Bonjour {me}, {dateStr}
          </div>
        </div>
      </Reveal>

      <MobileKpiSummary
        title="Vue d'ensemble"
        items={[
          { label: "Mes dossiers", value: counts.total, tone: "neutral", onClick: () => setView("clients") },
          { label: "TVA en retard ce mois", value: counts.tvaAlert, tone: counts.tvaAlert ? "amber" : "green", onClick: () => setView("tva") },
          { label: "Bilans en retard", value: counts.bilanRetard, tone: counts.bilanRetard ? "red" : "green", onClick: () => setView("bilans") },
          { label: "Accueils incomplets", value: counts.missionIncomplete, tone: counts.missionIncomplete ? "amber" : "green", onClick: () => setView("mission") },
        ]}
      />
      <div className="hidden md:grid grid-cols-2 md:grid-cols-4 gap-3 md:gap-4" style={{ marginBottom: 24 }}>
        <KpiCard index={0} label="Mes dossiers" value={counts.total} icon={Users} onClick={() => setView("clients")} linkLabel="Voir la liste" />
        <KpiCard index={1} label="TVA en retard ce mois" value={counts.tvaAlert} icon={Receipt} tone={counts.tvaAlert ? "amber" : "green"} onClick={() => setView("tva")} linkLabel="Voir les tâches" />
        <KpiCard index={2} label="Bilans en retard" value={counts.bilanRetard} icon={FileWarning} tone={counts.bilanRetard ? "red" : "green"} onClick={() => setView("bilans")} linkLabel="Voir les bilans" />
        <KpiCard index={3} label="Accueils incomplets" value={counts.missionIncomplete} icon={ClipboardCheck} tone={counts.missionIncomplete ? "amber" : "green"} onClick={() => setView("mission")} linkLabel="Voir les dossiers" />
      </div>
      <Reveal index={4}>
  <div style={{ marginBottom: 18 }}>

    <div
      style={{
        display: "flex",
        alignItems: "end",
        justifyContent: "space-between",
        gap: 12,
        marginBottom: 10,
      }}
    >
      <div>
        <h2
          style={{
            fontFamily: T.serif,
            fontSize: 15.5,
            fontWeight: 800,
            color: T.ink,
            margin: 0,
          }}
        >
          Répartition des dossiers
        </h2>

        <p
          style={{
            color: T.inkMuted,
            fontSize: 11,
            margin: "4px 0 0",
          }}
        >
          Une lecture immédiate du portefeuille par secteur,
          forme juridique et état d'avancement.
        </p>
      </div>

      <button
        onClick={() => setView("clients")}
        style={{
          background: T.navySoft,
          border: "none",
          borderRadius: 999,
          padding: "6px 10px",
          cursor: "pointer",
          color: T.navy,
          fontSize: 10.5,
          fontWeight: 700,
        }}
      >
        Voir le registre{" "}
        <ArrowUpRight
          size={11}
          style={{ verticalAlign: "middle" }}
        />
      </button>
    </div>

    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">

      <DonutDistribution
        title="Par secteur d'activité"
        items={sectorItems}
        total={myClients.length}
        icon={Briefcase}
        onItemClick={(item) => onDashboardFilter?.({ type: "secteur", value: item.key, label: item.label })}
      />

      <DonutDistribution
        title="Par formes juridiques"
        items={legalItems}
        total={myClients.length}
        icon={Landmark}
        onItemClick={(item) => onDashboardFilter?.({ type: "formeJuridique", value: item.key, label: item.label })}
      />

      <DonutDistribution
        title="Par statut du dossier"
        items={statusItems}
        total={myClients.length}
        icon={ShieldCheck}
        onItemClick={(item) => onDashboardFilter?.({ type: "statut", value: item.key, label: item.label })}
      />

      <DonutDistribution
        title="État de la TVA ce mois"
        items={tvaItems}
        total={tvaItems.reduce(
          (n, x) => n + x.value,
          0
        )}
        icon={Receipt}
        onItemClick={(item) => onDashboardFilter?.({ type: "tva", value: item.key, label: item.label })}
      />

    </div>

    <div
      className="grid grid-cols-1 md:grid-cols-2 gap-4"
      style={{ marginTop: 16 }}
    >

      <HorizontalDistribution
        title="Dossiers par collaborateur"
        items={collaboratorItems}
        icon={Users}
        onItemClick={(item) => onDashboardFilter?.({ type: "collaborateur", value: item.key, label: item.label })}
      />

      <Panel
        title="Indicateurs clés"
        right={
          <TrendingUp
            size={15}
            color={T.inkMuted}
          />
        }
      >
        <div className="grid grid-cols-2 gap-2.5">

          {[
            [
              ["Dossiers actifs", myClients.filter((c) => (c.statutDossier || "actif") === "actif"), T.green, { type: "statut", value: "actif", label: "Actifs" }],
              ["En transfert", myClients.filter((c) => c.statutDossier === "transfert"), T.amber, { type: "statut", value: "transfert", label: "En transfert" }],
              ["TVA en retard", myClients.filter(isTvaLate), T.red, { type: "tva", value: "RETARD", label: "TVA en retard" }],
              ["Mission < 100%", myClients.filter((c) => { const m = missionCompletion(c); return m && m.pct < 100; }), T.navy, { type: "missionIncomplete", value: "incomplete", label: "Mission < 100%" }],
            ].map(
            ([label, list, color, filter]) => (
              <div
                key={label}
                onClick={() => onDashboardFilter?.(filter)}
                className="clickable"
                style={{
                  padding: "12px 13px",
                  borderRadius: 12,
                  background: T.paper,
                  border: `1px solid ${T.line}`,
                  cursor: "pointer",
                }}
              >
                <div
                  style={{
                    width: 7,
                    height: 7,
                    borderRadius: 99,
                    background: color,
                    marginBottom: 9,
                  }}
                />

                <div
                  style={{
                    fontFamily: T.serif,
                    fontWeight: 800,
                    fontSize: 21,
                    color: T.ink,
                  }}
                >
                  {list.length}
                </div>

                <div
                  style={{
                    fontSize: 10.5,
                    color: T.inkMuted,
                    marginTop: 3,
                  }}
                >
                  {label}
                </div>
              </div>
            )
          )]}

        </div>
      </Panel>

    </div>
  </div>
</Reveal>

      <div style={{ marginBottom: 18 }}>
        <Panel index={4} title="À surveiller" right={
          <button onClick={() => setView("surveillance")} style={{ background: "none", border: "none", cursor: "pointer", color: T.navy, fontWeight: 600, fontSize: 11.5 }}>
            Voir tout <ArrowUpRight size={12} style={{ verticalAlign: "middle" }} />
          </button>
        }>
          {anomalies.length === 0 ? <EmptyNote text="Aucune anomalie détectée. Tous les contrôles connus sont au vert." /> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {anomalies.slice(0, 7).map((a) => (
                <div key={a.id} className="hoverRow clickable" onClick={() => onOpenClient(a.clientId)} style={{ display: "flex", alignItems: "center", gap: 9, padding: "8px 10px", borderRadius: T.radiusSm, border: `1px solid ${T.line}`, background: T.paper }}>
                  <span style={{ width: 7, height: 7, borderRadius: 99, background: a.gravite === "haute" ? T.red : T.amber, flexShrink: 0 }} />
                  <div style={{ flex: 1, minWidth: 0 }}><div style={{ fontWeight: 700, fontSize: 11.5 }}>{a.clientNom}</div><div style={{ fontSize: 10.5, color: T.inkMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.message}</div></div>
                  <Stamped tone={a.gravite === "haute" ? "red" : "amber"} small>{a.gravite === "haute" ? "Important" : "À vérifier"}</Stamped>
                </div>
              ))}
              {anomalies.length > 7 && <div style={{ fontSize: 10.5, color: T.inkMuted }}>{anomalies.length} anomalies au total · {criticalAnomalies.length} importantes · {importantAnomalies.length} à vérifier</div>}
            </div>
          )}
        </Panel>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-[1.3fr_1fr] gap-4 md:gap-[18px]">
        <Panel index={4} title="Tâches prioritaires" right={<Stamped tone="neutral" small>{sortedTasks.length}</Stamped>}>
          <div style={{ display: "flex", gap: 5, flexWrap: "wrap", marginBottom: 13 }}>
            {["Toutes", "retard", "aujourdhui", "demain", "semaine", "mois", "trimestre"].map((b) => (
              <button key={b} onClick={() => setTaskFilter(b)} style={{
                padding: "3.5px 9px", borderRadius: 999, fontSize: 10.5, fontWeight: 600,
                border: `1px solid ${taskFilter === b ? T.navy : T.line}`, background: taskFilter === b ? T.navySoft : T.card,
                color: taskFilter === b ? T.navy : T.inkSoft, cursor: "pointer",
              }}>{b === "Toutes" ? "Toutes" : BUCKET_LABELS[b]}</button>
            ))}
          </div>
          {sortedTasks.length === 0 ? <EmptyNote text="Rien à signaler sur cette période. Le registre est à jour." /> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
              {sortedTasks.slice(0, 5).map((t, i) => {
                const Icon = taskCategoryIcon(t.category);
                const { color, bg } = taskToneStyle(t.bucket);
                return (
                  <Reveal key={t.id} index={i} delay={0.1}>
                    <div className="hoverRow clickable" onClick={() => onOpenClient(t.client.id)} style={{ display: "flex", alignItems: "center", gap: 9, padding: "8px 10px", borderRadius: T.radiusSm, border: `1px solid ${T.line}`, background: T.paper }}>
                      <div style={{ width: 24, height: 24, borderRadius: 7, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", background: bg }}>
                        <Icon size={12} color={color} />
                      </div>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontWeight: 700, fontSize: 11.5, color: T.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.category} — {t.client.nom}</div>
                        <div style={{ fontSize: 10.5, color: T.inkMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{t.label}</div>
                      </div>
                      <Stamped tone={t.tone} small>{BUCKET_LABELS[t.bucket]}</Stamped>
                    </div>
                  </Reveal>
                );
              })}
            </div>
          )}
          {sortedTasks.length > 5 && (
            <button onClick={() => setView("mes-taches")} style={{ marginTop: 12, background: "none", border: "none", cursor: "pointer", color: T.navy, fontWeight: 600, fontSize: 11.5, display: "flex", alignItems: "center", gap: 4 }}>
              Voir toutes les tâches <ArrowUpRight size={12} />
            </button>
          )}
        </Panel>

        <Panel index={5} title="Dossiers non rapprochés" right={<Stamped tone={nonRapprochesM1.length ? "amber" : "green"} small>{nonRapprochesM1.length}</Stamped>}>
          {nonRapprochesM1.length === 0 ? <EmptyNote text="Tous les dossiers sont rapprochés sur le mois précédent." /> : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                <thead>
                  <tr style={{ textAlign: "left", color: T.inkMuted, fontSize: 9.5, textTransform: "uppercase", letterSpacing: "0.04em" }}>
                    <th style={{ padding: "0 6px 6px", fontWeight: 600, whiteSpace: "nowrap" }}>Dossier</th>
                    <th style={{ padding: "0 6px 6px", fontWeight: 600, whiteSpace: "nowrap" }}>Collab.</th>
                    <th style={{ padding: "0 6px 6px", fontWeight: 600, whiteSpace: "nowrap" }}>Période</th>
                    <th style={{ padding: "0 6px 6px", fontWeight: 600, whiteSpace: "nowrap" }}>Statut</th>
                  </tr>
                </thead>
                <tbody>
                  {nonRapprochesM1.slice(0, 6).map((c) => (
                    <tr key={c.id} className="hoverRow clickable" onClick={() => onOpenClient(c.id)} style={{ borderTop: `1px solid ${T.line}` }}>
                      <td style={{ padding: "6px 6px", fontWeight: 700, color: T.ink, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 130 }}>{c.nom}</td>
                      <td style={{ padding: "6px 6px", color: T.inkSoft, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 110 }}>{c.collab || "—"}</td>
                      <td style={{ padding: "6px 6px", color: T.inkSoft, whiteSpace: "nowrap" }}>{MOIS_FULL[prevKey] || prevKey}</td>
                      <td style={{ padding: "6px 6px", whiteSpace: "nowrap" }}><Stamped tone="amber" small>À rapprocher</Stamped></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
          {nonRapprochesM1.length > 6 && (
            <button onClick={() => setView("revision")} style={{ marginTop: 12, background: "none", border: "none", cursor: "pointer", color: T.navy, fontWeight: 600, fontSize: 11.5, display: "flex", alignItems: "center", gap: 4 }}>
              Voir les {nonRapprochesM1.length} dossiers <ArrowUpRight size={12} />
            </button>
          )}
        </Panel>
      </div>

      <div style={{ marginTop: 18 }}>
        <Panel index={6} title="Échéances à venir" right={
          <button onClick={() => setView("fiscal")} style={{ background: "none", border: "none", cursor: "pointer", color: T.navy, fontWeight: 600, fontSize: 11.5, display: "flex", alignItems: "center", gap: 4 }}>
            Voir le calendrier <ArrowUpRight size={12} />
          </button>
        }>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-2.5">
            {upcomingGroups.map((g) => (
              <div key={g.key} style={{ border: `1px solid ${T.line}`, borderRadius: T.radiusSm, padding: "9px 11px", background: T.paper }}>
                <div style={{ fontSize: 9.5, fontWeight: 700, color: T.inkMuted, textTransform: "uppercase", letterSpacing: "0.03em" }}>{g.label}</div>
                <div style={{ fontFamily: T.serif, fontSize: 18, fontWeight: 800, color: T.ink, margin: "4px 0 1px" }}>{g.count}</div>
                <div style={{ fontSize: 10, color: T.inkMuted, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                  {g.next ? g.next.category : "Rien de prévu"}
                </div>
              </div>
            ))}
          </div>
        </Panel>
      </div>

      <div style={{ marginTop: 18 }}>
        <Panel index={7} title="Mes dossiers par rôle">
          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {Object.entries(roleCounts).map(([role, n]) => (
              <div key={role}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 5 }}>
                  <span style={{ fontWeight: 600 }}>{role}</span>
                  <span style={{ fontFamily: T.mono, color: T.inkMuted }}>{n} dossiers</span>
                </div>
                <div style={{ height: 8, borderRadius: 4, background: T.paperDeep, overflow: "hidden" }}>
                  <div style={{ width: `${(n / maxRole) * 100}%`, height: "100%", background: T.navy, borderRadius: 4 }} />
                </div>
              </div>
            ))}
          </div>
        </Panel>
      </div>

      {(meRole === "chef_mission" || meRole === "admin") && superviseClients.length > 0 && (
        <div style={{ marginTop: 18 }}>
          <Panel index={8} title="Supervision d'équipe">
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              {Object.entries(byCollab).map(([collab, s]) => (
                <div key={collab} className="hoverRow clickable" onClick={() => onSuperviseClick && onSuperviseClick(collab)}
                  style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 14px", borderRadius: T.radiusSm, border: `1px solid ${T.line}`, background: T.paper }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700, fontSize: 12.5, color: T.ink }}>{collab}</div>
                    <div style={{ fontSize: 11, color: T.inkMuted }}>{s.total} dossier{s.total > 1 ? "s" : ""} sous supervision</div>
                  </div>
                  {s.bilanRetard > 0 && <Stamped tone="red" small>{s.bilanRetard} bilan{s.bilanRetard > 1 ? "s" : ""} retard</Stamped>}
                  {s.tvaAlert > 0 && <Stamped tone="amber" small>{s.tvaAlert} TVA</Stamped>}
                  {s.bilanRetard === 0 && s.tvaAlert === 0 && <Stamped tone="green" small>À jour</Stamped>}
                </div>
              ))}
            </div>
          </Panel>
        </div>
      )}
    </div>
  );
}

function MiniTvaOverview({ clients, setView }) {
  const key = currentMonthKey();
  const relevant = clients.filter((c) => c.tvaRegime === "CA3");
  const statuses = relevant.map((c) => effectiveTvaStatus(c, key));
  const ok = statuses.filter((s) => s === "OK").length;
  const fait = statuses.filter((s) => s === "FAIT").length;
  const na = statuses.filter((s) => s === "NA").length;
  const late = statuses.filter((s) => s === "RETARD").length;
  const enAttente = relevant.length - ok - fait - na - late;
  const total = relevant.length || 1;
  return (
    <div>
      <div style={{ display: "flex", height: 10, borderRadius: 5, overflow: "hidden", marginBottom: 10 }}>
        <div style={{ width: `${(ok / total) * 100}%`, background: T.green }} />
        <div style={{ width: `${(fait / total) * 100}%`, background: T.amber }} />
        <div style={{ width: `${(late / total) * 100}%`, background: T.red }} />
        <div style={{ width: `${(na / total) * 100}%`, background: T.line }} />
        <div style={{ width: `${(enAttente / total) * 100}%`, background: T.paperDeep }} />
      </div>
      <div style={{ display: "flex", gap: 16, fontSize: 12.5, flexWrap: "wrap" }}>
        <LegendDot color={T.green} label={`${ok} déclarées`} />
        <LegendDot color={T.amber} label={`${fait} préparées (à vérifier)`} />
        <LegendDot color={T.red} label={`${late} en retard`} />
        <LegendDot color={T.line} label={`${na} non applicable`} />
        <button onClick={() => setView("tva")} style={{ marginLeft: "auto", background: "none", border: "none", cursor: "pointer", color: T.navy, fontWeight: 600, fontSize: 12.5, display: "flex", alignItems: "center", gap: 4 }}>
          Voir le détail <ArrowUpRight size={13} />
        </button>
      </div>
    </div>
  );
}
function LegendDot({ color, label }) {
  return <span style={{ display: "flex", alignItems: "center", gap: 6, color: T.inkMuted }}><span style={{ width: 8, height: 8, borderRadius: "50%", background: color }} /> {label}</span>;
}
function KpiCard({ label, value, icon: Icon, tone, onClick, index = 0, linkLabel }) {
  const toneColor = tone === "red" ? T.red : tone === "amber" ? T.amber : tone === "green" ? T.green : T.navy;
  const toneSoft = tone === "red" ? T.redSoft : tone === "amber" ? T.amberSoft : tone === "green" ? T.greenSoft : T.navySoft;
  return (
    <Reveal index={index}>
      <div onClick={onClick} className={onClick ? "clickable" : ""} style={{ background: T.card, border: `1px solid ${T.line}`, borderRadius: T.radiusLg, boxShadow: T.shadowSm, padding: "22px 24px" }}>
        <div style={{ marginBottom: 14, width: 38, height: 38, borderRadius: 10, background: toneColor, display: "flex", alignItems: "center", justifyContent: "center" }}><Icon size={18} color="#FFFFFF" strokeWidth={2} /></div>
        <div style={{ fontFamily: T.serif, fontSize: 30, fontWeight: 800, color: T.ink, lineHeight: 1 }}>{value}</div>
        <div style={{ fontSize: 12.5, color: T.inkMuted, marginTop: 8, fontWeight: 500 }}>{label}</div>
        {onClick && (
          <div style={{ marginTop: 10, fontSize: 12, fontWeight: 600, color: T.navy, display: "flex", alignItems: "center", gap: 4 }}>
            {linkLabel || "Voir le détail"} <ArrowUpRight size={13} />
          </div>
        )}
      </div>
    </Reveal>
  );
}
/* ============================================================
   MOBILE KPI SUMMARY (variante "C") — remplace la grille de
   KpiCard sur mobile : une seule carte de synthèse avec des
   pastilles de couleur, le détail chiffré s'ouvrant dans une
   modale plein écran. La grille de KpiCard reste affichée telle
   quelle à partir de md: (desktop/tablette).
   Usage : <MobileKpiSummary title="…" items={[{label, value, tone, onClick}]} />
   ============================================================ */
function toneColors(tone) {
  if (tone === "red") return { dot: "#F87171", text: T.red };
  if (tone === "amber") return { dot: "#FBBF24", text: T.amber };
  if (tone === "green") return { dot: "#34D399", text: T.green };
  return { dot: "#93C5FD", text: T.navy };
}
function MobileKpiSummary({ title = "Vue d'ensemble", headlineItem, items = [] }) {
  const [open, setOpen] = useState(false);
  const headline = headlineItem || items[0];
  const rest = items.filter((it) => it !== headline);
  return (
    <div className="md:hidden" style={{ marginBottom: 18 }}>
      <div
        onClick={() => setOpen(true)}
        className="clickable"
        style={{
          background: `linear-gradient(135deg, ${T.sidebarBg2}, ${T.sidebarBg})`,
          borderRadius: T.radiusLg, padding: "18px 18px 16px", color: "#fff", boxShadow: T.shadow,
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
          <div>
            <div style={{ fontSize: 11.5, color: T.sidebarInkMuted, fontWeight: 600 }}>{headline?.label || title}</div>
            <div style={{ fontFamily: T.serif, fontSize: 28, fontWeight: 800, marginTop: 2 }}>{headline?.value ?? "—"}</div>
          </div>
          <button
            onClick={(e) => { e.stopPropagation(); setOpen(true); }}
            title="Voir tous les indicateurs"
            style={{ background: "rgba(255,255,255,0.12)", border: "none", color: "#fff", borderRadius: 9, width: 30, height: 30, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 }}
          >
            <ArrowUpRight size={15} strokeWidth={2.3} />
          </button>
        </div>
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
          {rest.map((it, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, color: T.sidebarInk }}>
              <span style={{ width: 8, height: 8, borderRadius: 8, background: toneColors(it.tone).dot, flexShrink: 0 }} />
              {it.value} {it.label.toLowerCase()}
            </div>
          ))}
        </div>
      </div>

      {open && (
        <>
          <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.5)", zIndex: 70 }} />
          <div style={{
            position: "fixed", left: 10, right: 10, top: "10%", bottom: "10%", background: T.card, borderRadius: T.radiusLg,
            padding: "16px 18px", boxShadow: T.shadowLg, zIndex: 71, overflowY: "auto",
          }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
              <h3 style={{ fontFamily: T.serif, fontSize: 16, fontWeight: 800, color: T.ink, margin: 0 }}>{title}</h3>
              <button onClick={() => setOpen(false)} style={{ width: 28, height: 28, borderRadius: 999, background: T.paper, border: "none", display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: T.inkSoft }}>
                <X size={14} />
              </button>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {items.map((it, i) => {
                const c = toneColors(it.tone);
                return (
                  <div
                    key={i}
                    onClick={it.onClick ? () => { setOpen(false); it.onClick(); } : undefined}
                    className={it.onClick ? "hoverRow clickable" : ""}
                    style={{ display: "flex", alignItems: "center", gap: 10, background: T.paper, border: `1px solid ${T.line}`, borderRadius: 14, padding: "12px 14px" }}
                  >
                    <span style={{ width: 8, height: 8, borderRadius: 8, background: c.dot, flexShrink: 0 }} />
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 13, fontWeight: 700, color: T.ink }}>{it.label}</div>
                      {it.sublabel && <div style={{ fontSize: 11, color: T.inkMuted, marginTop: 1 }}>{it.sublabel}</div>}
                    </div>
                    <div style={{ fontWeight: 800, color: c.text, fontSize: 15 }}>{it.value}</div>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}
    </div>
  );
}
function Panel({ title, children, right, index = 0 }) {
  return (
    <Reveal index={index}>
      <div className="card p-4 md:p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-sm md:text-[15px] font-bold text-ink m-0">{title}</h3>{right}
        </div>{children}
      </div>
    </Reveal>
  );
}
function EmptyNote({ text }) { return <div className="px-1 py-4 text-inkmuted text-xs italic">{text}</div>; }

/* ============================================================
   CLIENTS REGISTRY
   ============================================================ */
function ClientsRegistry({
  clients,
  allClients,
  search,
  setSearch,
  roleFilter,
  setRoleFilter,
  regimeFilter,
  setRegimeFilter,
  me,
  isAdmin,
  collabQuickFilter,
  setCollabQuickFilter,
  dashboardFilter,
  setDashboardFilter,
  selected,
  setSelected,
  onAdd,
  onUpdate,
  onImport,
  onAddClient,
}) {
  const [statutFilter, setStatutFilter] = useState("actif");
  const baseFiltered = useMemo(() => filterClients(clients, search, roleFilter, me, regimeFilter, dashboardFilter ? "tous" : statutFilter), [clients, search, roleFilter, me, regimeFilter, statutFilter, dashboardFilter]);

  const dashboardFiltered = useMemo(() => {
    if (!dashboardFilter) return baseFiltered;
    const { type, value } = dashboardFilter;
    return baseFiltered.filter((c) => {
      switch (type) {
        case "secteur":
          return (c.secteur || classifyActivite(c.activite)) === value;
        case "formeJuridique":
          return inferLegalForm(c) === value;
        case "statut":
          return (c.statutDossier || "actif") === value;
        case "tva": {
          if (!c.tvaRegime || c.tvaRegime === "FRANCHISE") return false;
          return effectiveTvaStatus(c, currentMonthKey()) === value;
        }
        case "collaborateur":
          return value === "Non assigné" ? !c.collab : c.collab === value;
        case "missionIncomplete": {
          const m = missionCompletion(c);
          return !!m && m.pct < 100;
        }
        default:
          return true;
      }
    });
  }, [baseFiltered, dashboardFilter]);

  const filtered = useMemo(() => {
    if (dashboardFilter) return dashboardFiltered;
    if (!collabQuickFilter) return baseFiltered;
    return baseFiltered.filter((c) => c.chefMission === me && (collabQuickFilter === "Non assigné" ? !c.collab : c.collab === collabQuickFilter));
  }, [baseFiltered, dashboardFiltered, dashboardFilter, collabQuickFilter, me]);
  const grouped = useMemo(() => {
    const g = {};
    [...filtered].sort((a, b) => a.nom.localeCompare(b.nom)).forEach((c) => {
      const letter = c.nom[0].toUpperCase(); g[letter] = g[letter] || []; g[letter].push(c);
    });
    return g;
  }, [filtered]);
  const [importBusy, setImportBusy] = useState(false);
  const [importMsg, setImportMsg] = useState(null);
  const fileInputRef = useRef(null);

  // Dossiers présents dans le référentiel d'origine (RAW_SEED_CLIENTS, embarqué dans le code)
  // mais absents de la base réelle (Supabase) — ex. un dossier ajouté au référentiel après le
  // tout premier chargement de l'application n'y est jamais inséré automatiquement.
  // Rapprochement par SIREN (ou par nom si le SIREN est vide/inconnu), comme pour l'import Excel.
  const [missingSeedOpen, setMissingSeedOpen] = useState(false);
  const [syncBusy, setSyncBusy] = useState(false);
  const missingSeedClients = useMemo(() => {
    if (!isAdmin) return [];
    const safeAll = Array.isArray(allClients) ? allClients.filter(Boolean) : [];
    const keyOf = (c) => (c.siren && String(c.siren).trim()) || (c.nom || "").trim().toLowerCase();
    const existingKeys = new Set(safeAll.map(keyOf).filter(Boolean));
    return RAW_SEED_CLIENTS.filter((c) => {
      const key = keyOf(c);
      return key && !existingKeys.has(key);
    });
  }, [allClients, isAdmin]);

  const handleSyncMissing = async () => {
    if (!missingSeedClients.length || !onAddClient) return;
    setSyncBusy(true);
    const normalized = migrateClients(missingSeedClients);
    for (const c of normalized) {
      onAddClient(c);
    }
    setSyncBusy(false);
    setMissingSeedOpen(false);
  };

  const handleFile = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ""; // permet de réimporter le même fichier
    if (!file) return;
    setImportBusy(true);
    setImportMsg(null);
    try {
      const rows = await parseClientsExcelFile(file);
      if (!rows.length) {
        setImportMsg({ tone: "amber", text: "Aucune ligne exploitable trouvée dans le fichier." });
      } else {
        const { created, updated } = onImport(rows);
        setImportMsg({ tone: "green", text: `${created} dossier(s) créé(s), ${updated} mis à jour.` });
      }
    } catch (err) {
      setImportMsg({ tone: "red", text: "Échec de l'import : " + err.message });
    } finally {
      setImportBusy(false);
      setTimeout(() => setImportMsg(null), 5000);
    }
  };

  let rowIndex = -1;
  return (
    <div>
      <Reveal>
        <div className="flex items-baseline justify-between mb-1.5 flex-wrap gap-2.5">
          <h1 className="text-base md:text-lg font-bold text-ink m-0">Registre clients</h1>
          <div className="flex items-center gap-2 flex-wrap">
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} className="hidden" />
            <button onClick={() => fileInputRef.current?.click()} disabled={importBusy} className="btn-secondary !py-2">
              {importBusy ? <Loader2 size={14} className="spin" /> : <ArrowUpRight size={14} className="-rotate-90" />}
              <span className="hidden sm:inline">Importer (Excel/CSV)</span>
            </button>
            <button onClick={() => exportClientsToExcel(filtered, `registre-clients-filtre-${todayISO()}.xlsx`)} className="btn-secondary !py-2">
              <ArrowUpRight size={14} className="rotate-90" /> <span className="hidden sm:inline">Exporter la liste (Excel)</span>
            </button>
            <button onClick={onAdd} className="btn-primary !py-2">
              <Plus size={15} /> <span className="hidden sm:inline">Nouveau client</span>
            </button>
          </div>
        </div>
        {collabQuickFilter && (
          <div className="flex items-center gap-2 mb-2">
            <span style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: T.navy, background: T.navySoft, padding: "4px 10px 4px 12px", borderRadius: 999 }}>
              Dossiers de {collabQuickFilter}
              <button onClick={() => setCollabQuickFilter && setCollabQuickFilter(null)} style={{ background: "none", border: "none", cursor: "pointer", color: T.navy, display: "flex", alignItems: "center", padding: 0 }}>
                <X size={13} />
              </button>
            </span>
          </div>
        )}
        {isAdmin && missingSeedClients.length > 0 && (
          <div style={{ marginBottom: 12, border: `1px solid #FCD34D`, background: T.amberSoft, borderRadius: 12, padding: "10px 14px" }}>
            <div className="flex items-center justify-between flex-wrap gap-2">
              <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: T.ink }}>
                <AlertTriangle size={15} color={T.amber} />
                <span>
                  <strong>{missingSeedClients.length}</strong> dossier{missingSeedClients.length > 1 ? "s" : ""} du référentiel initial {missingSeedClients.length > 1 ? "ne sont" : "n'est"} pas encore dans la base
                  {" "}({missingSeedClients.slice(0, 3).map((c) => c.nom).join(", ")}{missingSeedClients.length > 3 ? "…" : ""}).
                </span>
              </div>
              <div className="flex items-center gap-2">
                <button onClick={() => setMissingSeedOpen((o) => !o)} style={{ fontSize: 11.5, fontWeight: 600, color: T.navy, background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>
                  {missingSeedOpen ? "Masquer la liste" : "Voir la liste"}
                </button>
                <button onClick={handleSyncMissing} disabled={syncBusy} className="btn-secondary !py-1.5 !text-xs">
                  {syncBusy ? <Loader2 size={13} className="spin" /> : <Plus size={13} />}
                  Ajouter {missingSeedClients.length > 1 ? "ces dossiers" : "ce dossier"}
                </button>
              </div>
            </div>
            {missingSeedOpen && (
              <ul style={{ margin: "8px 0 0", padding: "0 0 0 18px", fontSize: 11.5, color: T.inkMuted, lineHeight: 1.7 }}>
                {missingSeedClients.map((c) => (
                  <li key={c.siren || c.nom}>{c.nom}{c.siren ? ` — SIREN ${c.siren}` : ""}</li>
                ))}
              </ul>
            )}
          </div>
        )}
        {dashboardFilter && (
          <div className="flex items-center gap-2 mb-2">
            <span style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, fontWeight: 700, color: T.navy, background: T.navySoft, padding: "4px 10px 4px 12px", borderRadius: 999 }}>
              Filtre dashboard : {dashboardFilter.label}
              <button onClick={() => setDashboardFilter && setDashboardFilter(null)} style={{ background: "none", border: "none", cursor: "pointer", color: T.navy, display: "flex", alignItems: "center", padding: 0 }}>
                <X size={13} />
              </button>
            </span>
          </div>
        )}
      </Reveal>
      {importMsg && (
          <div style={{ whiteSpace: "pre-line" }} className={`mt-2 text-xs font-semibold px-2.5 py-1.5 rounded-lg inline-block max-w-full ${importMsg.tone === "green" ? "bg-badge-green-bg text-badge-green-text" : importMsg.tone === "red" ? "bg-badge-red-bg text-badge-red-text" : "bg-badge-amber-bg text-badge-amber-text"}`}>{importMsg.text}</div>
        )}
      <p className="text-inkmuted text-xs mt-1.5 mb-5">Cliquez un dossier pour ouvrir sa fiche complète.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter}
        statutFilter={statutFilter} setStatutFilter={setStatutFilter} search={search} setSearch={setSearch} />

      {/* En-tête colonnes : visible à partir de md, masqué sur mobile (les dossiers s'affichent en cartes empilées) */}
      <div className="hidden md:grid gap-0" style={{ gridTemplateColumns: "1.8fr 0.9fr 1fr 0.9fr 0.7fr 0.8fr 1.2fr 92px", padding: "0 18px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: T.inkMuted, fontWeight: 600, marginBottom: 10 }}>
        <div>Dossier</div>
<div>SIREN</div>
<div>Rôles</div>
<div>Clôture</div>
<div>Régime</div>
<div>Logiciel</div>
<div>Statuts</div>
<div>Actions</div>
      </div>
      <div className="flex flex-col gap-2">
        {Object.keys(grouped).sort().map((letter) => (
          <div key={letter}>
            <div className="px-1.5 py-1 font-mono text-[10.5px] font-bold text-accent-deep tracking-widest">{letter}</div>
            <div className="flex flex-col gap-2">
              {grouped[letter].map((c) => {
                rowIndex += 1;
                const isInactif = c.statutDossier === "inactif";
                const isTransfert = c.statutDossier === "transfert";
                const alertBadge = isBilanLate(c) ? <Stamped tone="red" small>Bilan retard</Stamped>
                  : isTvaLate(c) ? <Stamped tone="amber" small>TVA</Stamped>
                  : null;
                const statutBadge = isInactif ? <Stamped tone="neutral" small>Inactif</Stamped>
                  : isTransfert ? <Stamped tone="amber" small>En transfert</Stamped>
                  : <Stamped tone="green" small>Actif</Stamped>;
                const roles = [c.collab === me && "Collaborateur", c.expert === me && "Expert", c.chefMission === me && "Chef de mission"].filter(Boolean);
                return (
                  <Reveal key={c.id} index={rowIndex}>
                    {/* Ligne tableau (md et +) */}
                    <div onClick={() => setSelected(c.id)}
                      className={`hoverRow clickable hidden md:grid items-center rounded-xl border px-4 py-3.5 text-xs ${selected === c.id ? "border-accent-deep bg-accent-soft" : "border-line bg-card shadow-xs"} ${isInactif ? "opacity-55" : ""}`}
                      style={{ gridTemplateColumns: "1.8fr 0.9fr 1fr 0.9fr 0.7fr 0.8fr 1.2fr 92px" }}>
                      <div className="font-semibold text-ink flex items-center gap-2">
                        {c.nom}
                      </div>
                      <div className="font-mono text-xs text-inkmuted">{c.siren}</div>
                      <div className="flex flex-col gap-0.5 text-[10.5px] text-inkmuted">
                        {roles.map((r) => <span key={r}>{r}</span>)}
                      </div>
                      <div className="font-mono text-[11.5px] text-inkmuted">
                        <input type="date" defaultValue={c.dateCloture || ""} onClick={(e) => e.stopPropagation()}
                          onChange={(e) => onUpdate(c.id, { dateCloture: e.target.value })}
                          className="border-none bg-transparent font-mono text-[11.5px] text-inkmuted w-[118px]" />
                      </div>
                      <div className="text-xs text-inksoft font-mono">{c.tvaRegime || "—"}</div>
                      <div className="text-xs text-inksoft font-mono">{c.logiciel || "—"}</div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {statutBadge}
                        {alertBadge}
                        {c.lienSharepoint && (
                          <a href={c.lienSharepoint} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()}
                            title="Ouvrir dans SharePoint" style={{ color: T.navy, display: "flex", alignItems: "center" }}>
                            <ExternalLink size={13} />
                          </a>
                        )}
                      </div>
                     <div
  className="flex items-center justify-end gap-1.5"
  onClick={(e) => e.stopPropagation()}
>
  <button
    title="Ouvrir"
    onClick={() => setSelected(c.id)}
    className="w-7 h-7 rounded-lg border border-line bg-white text-inkmuted hover:text-accent hover:border-accent inline-flex items-center justify-center"
  >
    <Eye size={13} />
  </button>

</div>
                    </div>
                    {/* Carte empilée (mobile) */}
                    <div onClick={() => setSelected(c.id)}
                      className={`hoverRow clickable md:hidden rounded-xl border p-3.5 flex flex-col gap-2 ${selected === c.id ? "border-accent-deep bg-accent-soft" : "border-line bg-card shadow-xs"} ${isInactif ? "opacity-55" : ""}`}>
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-semibold text-ink text-sm">{c.nom}</span>
                        <ChevronRight size={15} className="text-inkmuted shrink-0" />
                      </div>
                      <div className="flex items-center gap-2 flex-wrap text-[11px] text-inkmuted">
                        <span className="font-mono">{c.siren || "—"}</span>
                        {roles.length > 0 && <span>· {roles.join(", ")}</span>}
                        {c.tvaRegime && <span className="font-mono">· {c.tvaRegime}</span>}
                        {c.logiciel && <span className="font-mono">· {c.logiciel}</span>}
                        {c.lienSharepoint && (
                          <a href={c.lienSharepoint} target="_blank" rel="noopener noreferrer" onClick={(e) => e.stopPropagation()} style={{ color: T.navy, display: "flex", alignItems: "center" }}>
                            <ExternalLink size={12} />
                          </a>
                        )}
                      </div>
                      <div className="flex items-center justify-between gap-2">

  <div className="flex items-center gap-1.5 flex-wrap">
    {statutBadge}
    {alertBadge}
  </div>

</div>
                    </div>
                  </Reveal>
                );
              })}
            </div>
          </div>
        ))}
        {filtered.length === 0 && <EmptyNote text="Aucun dossier ne correspond à cette recherche." />}
      </div>
    </div>
  );
}

/* ============================================================
   CLIENT DRAWER
   ============================================================ */
/* ============================================================
   CLIENT EDITOR — page pleine avec onglets (façon MyUnisoft)
   Remplace l'ancien tiroir latéral : le dossier s'ouvre dans un
   onglet de la barre du haut, comme "AC INVEST" chez MyUnisoft.
   ============================================================ */
function ClientEditorPage({
  client,
  team,
  me,
  meId,
  portefeuilleId,
  onUpdate,
  onDelete,
  onClose,
  setView
}) {
  const [tab, setTab] = useState("infos");
  // Brouillon local : toutes les modifications restent ici tant qu'on n'a pas cliqué "Enregistrer".
  // Reset uniquement quand on change de dossier (changement de client.id), pas à chaque frappe.
  const [draft, setDraft] = useState(client);
  useEffect(() => { setDraft(client); }, [client.id]);
  const dirty = JSON.stringify(draft) !== JSON.stringify(client);
  const patchDraft = (_id, patch) => setDraft((d) => ({ ...d, ...patch }));
  // notesCollab est écrit en direct par NotesTab (temps réel, hors draft) : on ne le
  // laisse jamais dans le payload de sauvegarde, pour ne pas écraser une note ajoutée
  // entre-temps avec une version périmée du draft.
  const save = () => onUpdate(client.id, { ...draft, notesCollab: client.notesCollab });
  const discard = () => setDraft(client);
  const handleClose = () => {
    if (dirty && !confirm("Des modifications ne sont pas enregistrées. Fermer sans enregistrer ?")) return;
    onClose();
  };
  if (!client) return null;
  const tabs = [
     { id: "mission", label: "Accueil client" }, { id: "infos", label: "Infos générales" }, { id: "contact", label: "Fiche contact" }, { id: "corporate", label: "Corporate" }, { id: "tva", label: "TVA" },
    { id: "bilan", label: "Bilan" }, { id: "acomptes", label: "Acomptes" }, { id: "age", label: "AGE / AGO" },
    { id: "formeJuridique", label: "Forme juridique" },{ id: "revision", label: "Révision" },
    { id: "acces", label: "Accès & codes" },
    { id: "suivi", label: "Demandes & pièces" },
    { id: "rentabilite", label: "Temps & rentabilité" },
    { id: "validation", label: "Validation" },
    { id: "notes", label: "Notes" }, { id: "historique", label: "Historique" },
  ];
  return (
    <div>
      <Reveal>
        <div className="flex items-start justify-between gap-3 flex-wrap" style={{ marginBottom: 4 }}>
          <div className="min-w-0">
            <div style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{client.siren || "SIREN non renseigné"}</div>
            <input defaultValue={client.nom} onBlur={(e) => patchDraft(client.id, { nom: e.target.value || client.nom })}
              className="w-full sm:min-w-[260px] sm:w-auto"
              style={{ fontFamily: T.serif, fontSize: 16, fontWeight: 700, color: T.ink, border: "none", background: "transparent", padding: "2px 0", margin: "2px 0 6px" }} />
            <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap", fontSize: 12 }}>
              <RoleBadge role="Collab." name={client.collab} />
              <RoleBadge role="Expert" name={client.expert} />
              <RoleBadge role="Chef de mission" name={client.chefMission} />
              {client.tvaRegime && <span style={{ fontFamily: T.mono, fontSize: 11, color: T.navy, fontWeight: 700, background: T.navySoft, padding: "2px 9px", borderRadius: 999 }}>{client.tvaRegime}{client.tvaRegime === "CA3" && client.tvaPeriodicite ? ` · ${TVA_PERIODICITE_LABELS[client.tvaPeriodicite]}` : ""}</span>}
              {client.regimeFiscal && <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkSoft, fontWeight: 700, background: T.paperDeep, padding: "2px 9px", borderRadius: 999 }}>{client.regimeFiscal}</span>}
              {client.tvaExig && <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkSoft, fontWeight: 700, background: T.paper, padding: "2px 9px", borderRadius: 999 }}>Exig. {client.tvaExig}</span>}
              {client.statutDossier === "transfert" ? (
                <span title="Résiliation ou reprise en cours — se termine automatiquement depuis l'onglet concerné"
                  style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 700, padding: "3px 10px 3px 8px", borderRadius: 999, background: T.amberSoft, color: T.amber }}>
                  <span style={{ width: 7, height: 7, borderRadius: "50%", background: T.amber, flexShrink: 0 }} />
                  En transfert
                </span>
              ) : (
                <button
                  onClick={() => patchDraft(client.id, { statutDossier: draft.statutDossier === "inactif" ? "actif" : "inactif" })}
                  className="statusToggle"
                  title="Basculer le statut du dossier"
                  style={{
                    display: "flex", alignItems: "center", gap: 5, fontSize: 11, fontWeight: 700, padding: "3px 10px 3px 8px", borderRadius: 999, border: "none",
                    background: client.statutDossier === "inactif" ? T.paperDeep : T.greenSoft,
                    color: client.statutDossier === "inactif" ? T.inkMuted : T.green,
                  }}>
                  <span style={{ width: 7, height: 7, borderRadius: "50%", background: client.statutDossier === "inactif" ? T.inkMuted : T.green, flexShrink: 0 }} />
                  {client.statutDossier === "inactif" ? "Inactif" : "Actif"}
                </button>
              )}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
            {dirty && (
              <>
                <span style={{ fontSize: 11, color: T.amber, fontWeight: 600, display: "flex", alignItems: "center", gap: 4 }}>
                  <span style={{ width: 6, height: 6, borderRadius: "50%", background: T.amber }} /> Modifications non enregistrées
                </span>
                <button onClick={discard} style={{ background: "none", border: `1px solid ${T.line}`, borderRadius: 9, padding: "7px 12px", cursor: "pointer", color: T.inkMuted, fontSize: 12 }}>
                  Annuler
                </button>
                <button onClick={save} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, border: "none", borderRadius: 9, padding: "7px 14px", cursor: "pointer", color: "#fff", fontSize: 12, fontWeight: 700 }}>
                  <Check size={14} /> Enregistrer
                </button>
              </>
            )}
            {onDelete && (
  <button
    onClick={async () => {
      if (
        !confirm(
          `Supprimer définitivement le dossier « ${client.nom} » ?\n\nCette action est irréversible.`
        )
      ) {
        return;
      }

      const ok = await onDelete(client.id);

      if (ok) {
        onClose();
      }
    }}
    style={{
      display: "flex",
      alignItems: "center",
      gap: 6,
      background: T.redSoft,
      border: "1px solid #FECACA",
      borderRadius: 9,
      padding: "7px 12px",
      cursor: "pointer",
      color: T.red,
      fontSize: 12,
      fontWeight: 700,
    }}
  >
    <Trash2 size={14} />
    Supprimer
  </button>
)}
            <button onClick={handleClose} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${T.line}`, borderRadius: 9, padding: "7px 12px", cursor: "pointer", color: T.inkMuted, fontSize: 12 }}>
              <X size={14} /> Fermer l'onglet
            </button>
          </div>
        </div>
      </Reveal>
      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${T.line}`, marginTop: 16, marginBottom: 22, overflowX: "auto" }} className="scrollbar">
        {tabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "10px 14px", background: "none", border: "none", cursor: "pointer", fontSize: 12.5, fontWeight: 600, whiteSpace: "nowrap",
            color: tab === t.id ? T.navy : T.inkMuted, borderBottom: tab === t.id ? `2.5px solid ${T.navy}` : "2.5px solid transparent", marginBottom: -1,
          }}>{t.label}</button>
        ))}
      </div>
      <div style={{ maxWidth: 720, background: T.card, border: `1px solid ${T.line}`, borderRadius: T.radius, padding: "22px 24px", boxShadow: T.shadowSm }}>
        {tab === "infos" && <InfosTab client={draft} team={team} onUpdate={patchDraft} setView={setView} />}
        {tab === "contact" && <ContactTab client={draft} onUpdate={patchDraft} />}
        {tab === "corporate" && <CorporateTab client={draft} onUpdate={patchDraft} />}
        {tab === "tva" && <TvaTab client={draft} onUpdate={patchDraft} />}
        {tab === "bilan" && <BilanTab client={draft} onUpdate={patchDraft} />}
        {tab === "acomptes" && <AcomptesTab client={draft} onUpdate={patchDraft} />}
        {tab === "age" && <AgeAgoEditor client={draft} onUpdate={patchDraft} />}
        {tab === "formeJuridique" && <FormeJuridiqueEditor client={draft} onUpdate={patchDraft} />}
{tab === "revision" && <RevisionTab client={draft} onUpdate={patchDraft} setView={setView} />}
        {tab === "mission" && <MissionTab client={draft} onUpdate={patchDraft} />}
        {tab === "acces" && <AccesTab client={draft} onUpdate={patchDraft} />}
        {tab === "suivi" && <DemandesPiecesTab client={draft} onUpdate={patchDraft} />}
        {tab === "rentabilite" && <RentabiliteTab client={draft} onUpdate={patchDraft} />}
        {tab === "validation" && <ValidationDossierTab client={draft} onUpdate={patchDraft} me={me} />}
        {tab === "notes" && <NotesTab client={client} me={me} meId={meId} portefeuilleId={portefeuilleId} onUpdate={onUpdate} />}
        {tab === "historique" && <HistoriqueTab clientId={client.id} team={team} />}
      </div>
    </div>
  );
}
/* ============================================================
   ACCÈS & CODES — identifiants et mots de passe sensibles du
   dossier, classés par catégorie. Stocké dans client.acces =
   { banques: [], comptesFournisseurs: [], caisse: [], netEntreprise: [], autres: [] }
   chaque entrée : { id, libelle, identifiant, motDePasse, note }
   ============================================================ */
const ACCES_CATEGORIES = [
  { key: "banques", label: "Banque(s)", icon: Landmark, placeholder: "ex. BNP Paribas — compte courant" },
  { key: "comptesFournisseurs", label: "Comptes fournisseurs", icon: Briefcase, placeholder: "ex. Portail EDF Pro" },
  { key: "caisse", label: "Caisse / plateforme de facturation", icon: Wallet, placeholder: "ex. Logiciel de caisse Zettle" },
  { key: "netEntreprise", label: "Net-entreprise", icon: ShieldCheck, placeholder: "ex. Net-entreprise.fr" },
  { key: "autres", label: "Autres accès", icon: KeyRound, placeholder: "ex. Portail impots.gouv.fr" },
];
function uid() { return `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`; }

function AccesPasswordField({ value, onCommit }) {
  const [visible, setVisible] = useState(false);
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(value || "");
      setCopied(true);
      setTimeout(() => setCopied(false), 1400);
    } catch { /* clipboard indisponible — on ignore silencieusement */ }
  };
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
      <input
        type={visible ? "text" : "password"}
        defaultValue={value}
        onBlur={(e) => onCommit(e.target.value)}
        placeholder="Mot de passe"
        style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 8, border: `1px solid ${T.line}`, width: 140, background: T.card }}
      />
      <button type="button" onClick={() => setVisible((v) => !v)} title={visible ? "Masquer" : "Afficher"}
        style={{ background: "none", border: `1px solid ${T.line}`, borderRadius: 7, padding: 5, cursor: "pointer", color: T.inkMuted, display: "flex" }}>
        {visible ? <EyeOff size={13} /> : <Eye size={13} />}
      </button>
      <button type="button" onClick={copy} title="Copier le mot de passe"
        style={{ background: "none", border: `1px solid ${T.line}`, borderRadius: 7, padding: 5, cursor: "pointer", color: copied ? T.green : T.inkMuted, display: "flex" }}>
        <Copy size={13} />
      </button>
      {copied && <span style={{ fontSize: 10.5, color: T.green, fontWeight: 700 }}>Copié !</span>}
    </div>
  );
}

function AccesEntryRow({ entry, onChange, onRemove }) {
  const patch = (f) => onChange({ ...entry, ...f });
  return (
    <div style={{ border: `1px solid ${T.line}`, borderRadius: 11, padding: "10px 12px", marginBottom: 8, background: T.paper }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", marginBottom: 8 }}>
        <input defaultValue={entry.libelle} placeholder="Libellé (ex. BNP Paribas — compte courant)" onBlur={(e) => patch({ libelle: e.target.value })}
          style={{ flex: "1 1 200px", fontSize: 12.5, fontWeight: 600, padding: "6px 9px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />
        <button type="button" onClick={onRemove} title="Supprimer cet accès"
          style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted, display: "flex", flexShrink: 0 }}>
          <Trash2 size={14} />
        </button>
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 8, alignItems: "center" }}>
        <input defaultValue={entry.identifiant} placeholder="Identifiant" onBlur={(e) => patch({ identifiant: e.target.value })}
          style={{ fontSize: 12, padding: "5px 8px", borderRadius: 8, border: `1px solid ${T.line}`, width: 150, background: T.card }} />
        <AccesPasswordField value={entry.motDePasse} onCommit={(v) => patch({ motDePasse: v })} />
      </div>
      <textarea defaultValue={entry.note} placeholder="Note libre (URL, RIB, digicode…)" onBlur={(e) => patch({ note: e.target.value })}
        rows={2} style={{ width: "100%", marginTop: 8, fontSize: 12, padding: "6px 9px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card, resize: "vertical", fontFamily: T.sans }} />
    </div>
  );
}

function AccesCategoryPanel({ category, entries, onUpdate }) {
  const Icon = category.icon;
  const list = entries || [];
  const setList = (next) => onUpdate(next);
  const addEntry = () => setList([...list, { id: uid(), libelle: "", identifiant: "", motDePasse: "", note: "" }]);
  const updateEntry = (id, next) => setList(list.map((e) => (e.id === id ? next : e)));
  const removeEntry = (id) => {
    if (!confirm("Supprimer cet accès ?")) return;
    setList(list.filter((e) => e.id !== id));
  };
  return (
    <Panel title={`${category.label} (${list.length})`}>
      {list.length === 0 && <EmptyNote text="Aucun accès enregistré dans cette catégorie." />}
      {list.map((e) => <AccesEntryRow key={e.id} entry={e} onChange={(next) => updateEntry(e.id, next)} onRemove={() => removeEntry(e.id)} />)}
      <button type="button" onClick={addEntry} style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 4, background: "none", border: `1px dashed ${T.line}`, borderRadius: 9, padding: "7px 12px", fontSize: 12, color: T.navy, cursor: "pointer" }}>
        <Plus size={13} /> Ajouter {category.placeholder ? `(${category.placeholder})` : "un accès"}
      </button>
    </Panel>
  );
}

function AccesTab({ client, onUpdate }) {
  const acces = client.acces || {};
  const patchCategory = (key, list) => onUpdate(client.id, { acces: { ...acces, [key]: list } });
  return (
    <div>
      <div style={{ fontSize: 11.5, color: T.inkMuted, background: T.navySoft, padding: "8px 12px", borderRadius: 9, marginBottom: 16, display: "flex", alignItems: "center", gap: 8 }}>
        <KeyRound size={14} color={T.navy} style={{ flexShrink: 0 }} />
        Ces informations sont sensibles : elles ne sont visibles que par les collaborateurs ayant accès à ce dossier. Pensez à changer les mots de passe partagés régulièrement.
      </div>
      {ACCES_CATEGORIES.map((cat, i) => (
        <div key={cat.key}>
          <AccesCategoryPanel category={cat} entries={acces[cat.key]} onUpdate={(list) => patchCategory(cat.key, list)} />
          {i < ACCES_CATEGORIES.length - 1 && <div style={{ height: 14 }} />}
        </div>
      ))}
    </div>
  );
}

function SocialTab({ client, onUpdate }) {
  const s = client.social || {};
  const patch = (f) => onUpdate(client.id, { social: { ...s, ...f } });
  const alert = seuilEffectifAlert(s.effectif);
  return (
    <div>
      <FieldRow label="Concerné par le social">
        <ConcerneToggle on={!!s.concerne} onChange={(v) => patch({ concerne: v })} />
      </FieldRow>
      <FieldRow label="Effectif">
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <TextInput defaultValue={s.effectif} onCommit={(v) => patch({ effectif: v })} width={60} />
          {alert && <Stamped tone={alert.tone} small>{alert.label}</Stamped>}
        </div>
      </FieldRow>
      <FieldRow label="Cabinet de paie"><TextInput defaultValue={s.cabinetPaie} onCommit={(v) => patch({ cabinetPaie: v })} width={160} align="left" /></FieldRow>
      <FieldRow label="Contact gestionnaire — nom"><TextInput defaultValue={s.gestionnaireNom} onCommit={(v) => patch({ gestionnaireNom: v })} width={160} align="left" /></FieldRow>
      <FieldRow label="Contact gestionnaire — adresse"><TextInput defaultValue={s.gestionnaireAdresse} onCommit={(v) => patch({ gestionnaireAdresse: v })} width={220} align="left" /></FieldRow>
      <FieldRow label="Contact gestionnaire — email"><TextInput defaultValue={s.gestionnaireEmail} onCommit={(v) => patch({ gestionnaireEmail: v })} width={180} align="left" /></FieldRow>
      <FieldRow label="Contact gestionnaire — tél."><TextInput defaultValue={s.gestionnaireTel} onCommit={(v) => patch({ gestionnaireTel: v })} width={140} align="left" /></FieldRow>
      <FieldRow label="Convention collective"><TextInput defaultValue={s.conventionCollective} onCommit={(v) => patch({ conventionCollective: v })} width={180} align="left" /></FieldRow>
      <FieldRow label="Régime social du dirigeant">
        <SelectPill value={s.regimeDirigeant} options={["assimile_salarie", "tns"]} labels={{ assimile_salarie: "Assimilé salarié", tns: "TNS" }} onChange={(v) => patch({ regimeDirigeant: v })} />
      </FieldRow>
    </div>
  );
}

function FieldRow({ label, children }) {
  return <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderBottom: `1px solid ${T.line}`, gap: 10 }}>
    <span style={{ fontSize: 12, color: T.inkMuted }}>{label}</span><div>{children}</div>
  </div>;
}
function SelectPill({ value, options, onChange, allowEmpty = true, labels }) {
  return (
    <select value={value || ""} onChange={(e) => onChange(e.target.value)} style={{ fontFamily: T.mono, fontSize: 12, fontWeight: 600, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, color: T.ink }}>
      {allowEmpty && <option value="">—</option>}
      {options.map((o) => <option key={o} value={o}>{labels?.[o] || o}</option>)}
    </select>
  );
}
function TextInput({ defaultValue, onCommit, placeholder, width = 160, align = "right" }) {
  return <input defaultValue={defaultValue || ""} placeholder={placeholder} onBlur={(e) => onCommit(e.target.value)}
    style={{ fontFamily: T.sans, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, width, textAlign: align, background: T.card }} />;
}

function InfosTab({ client, team, onUpdate, setView }) {
  const teamNames = team.map((t) => t.nom);
  return (
    <div>
      <FieldRow label="SIREN"><TextInput defaultValue={client.siren} onCommit={(v) => onUpdate(client.id, { siren: v })} width={140} /></FieldRow>
      <FieldRow label="Honoraires actuels">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontFamily: T.mono, fontSize: 12.5, color: T.ink, fontWeight: 600 }}>{client.honoraires?.montant || "—"}</span>
          {setView && (
            <button onClick={() => setView("honoraires")} style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: `1px solid ${T.line}`, borderRadius: 8, padding: "3px 9px", fontSize: 11, color: T.navy, cursor: "pointer", fontWeight: 600 }}>
              <ExternalLink size={12} /> Modifier
            </button>
          )}
        </div>
      </FieldRow>
      <FieldRow label="Logiciel"><SelectPill value={client.logiciel} options={["MYUNISOFT", "QUADRA"]} onChange={(v) => onUpdate(client.id, { logiciel: v })} /></FieldRow>
      <FieldRow label="Lien SharePoint">
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <TextInput defaultValue={client.lienSharepoint} onCommit={(v) => onUpdate(client.id, { lienSharepoint: v })} placeholder="https://…sharepoint.com/…" width={200} align="left" />
          {client.lienSharepoint && (
            <a href={client.lienSharepoint} target="_blank" rel="noopener noreferrer" title="Ouvrir dans SharePoint"
              style={{ color: T.navy, display: "flex", alignItems: "center" }}>
              <ExternalLink size={15} />
            </a>
          )}
        </div>
      </FieldRow>
      <FieldRow label="Forme juridique"><SelectPill value={client.formeJuridique} options={["EI", "EURL", "SARL", "SAS", "SASU", "SCI", "SCM", "SELARL", "SA", "SNC", "Association"]} onChange={(v) => onUpdate(client.id, { formeJuridique: v })} /></FieldRow>
      <FieldRow label="Régime fiscal"><SelectPill value={client.regimeFiscal} options={["IS", "IR"]} labels={{ IS: "IS — Impôt sur les sociétés", IR: "IR — Impôt sur le revenu" }} onChange={(v) => onUpdate(client.id, { regimeFiscal: v })} /></FieldRow>
      <FieldRow label="Capital social"><TextInput defaultValue={client.capital} onCommit={(v) => onUpdate(client.id, { capital: v })} placeholder="ex. 5 000 €" width={140} /></FieldRow>
      <FieldRow label="Code NAF / APE"><TextInput defaultValue={client.codeNaf} onCommit={(v) => onUpdate(client.id, { codeNaf: v.toUpperCase(), secteur: classifyNaf(v) })} placeholder="ex. 56.10A" width={140} /></FieldRow>
      <FieldRow label="Activité"><TextInput defaultValue={client.activite} onCommit={(v) => onUpdate(client.id, { activite: v })} placeholder="Information descriptive" width={200} align="left" /></FieldRow>
      <FieldRow label="Secteur">
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{
            fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999,
            background: (SECTEURS_ACTIVITE.find((s) => s.id === client.secteur)?.color || T.inkMuted) + "22",
            color: SECTEURS_ACTIVITE.find((s) => s.id === client.secteur)?.color || T.inkMuted,
          }}>
            {SECTEURS_ACTIVITE.find((s) => s.id === client.secteur)?.label || "Non classé"}
          </span>
          <SelectPill
            value={client.secteur}
            options={SECTEURS_ACTIVITE.map((s) => s.id)}
            labels={Object.fromEntries(SECTEURS_ACTIVITE.map((s) => [s.id, s.label]))}
            onChange={(v) => onUpdate(client.id, { secteur: v, secteurManuel: true })}
          />
        </div>
      </FieldRow>
      <FieldRow label="Date de clôture d'exercice"><input type="date" defaultValue={client.dateCloture || ""} onChange={(e) => onUpdate(client.id, { dateCloture: e.target.value })} style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} /></FieldRow>
      <div style={{ height: 6 }} />
      <FieldRow label="Collaborateur"><SelectPill value={client.collab} options={teamNames} onChange={(v) => onUpdate(client.id, { collab: v })} /></FieldRow>
      <FieldRow label="Expert"><SelectPill value={client.expert} options={teamNames} onChange={(v) => onUpdate(client.id, { expert: v })} /></FieldRow>
      <FieldRow label="Chef de mission"><SelectPill value={client.chefMission} options={teamNames} onChange={(v) => {
        const chef = team.find((t) => t.nom === v);
        onUpdate(client.id, { chefMission: v, chefMission_id: chef?.id || "" });
      }} /></FieldRow>
      <FieldRow label="Régime TVA"><SelectPill value={client.tvaRegime} options={REGIMES_TVA} labels={REGIMES_TVA_LABELS} onChange={(v) => onUpdate(client.id, { tvaRegime: v })} /></FieldRow>
    </div>
  );
}

function ContactTab({ client, onUpdate }) {
  const contact = client.contact || {};
  const patch = (field, value) => onUpdate(client.id, { contact: { ...contact, [field]: value } });
  return (
    <div>
      <Panel title="Coordonnées du client">
        <FieldRow label="Nom du contact"><TextInput defaultValue={contact.contactNom} onCommit={(v) => patch("contactNom", v)} placeholder="Nom et prénom" width={220} align="left" /></FieldRow>
        <FieldRow label="Fonction"><TextInput defaultValue={contact.contactFonction} onCommit={(v) => patch("contactFonction", v)} placeholder="Dirigeant, comptable…" width={220} align="left" /></FieldRow>
        <FieldRow label="Téléphone"><TextInput defaultValue={contact.telephone} onCommit={(v) => patch("telephone", v)} placeholder="06 00 00 00 00" width={180} /></FieldRow>
        <FieldRow label="E-mail"><TextInput defaultValue={contact.email} onCommit={(v) => patch("email", v)} placeholder="contact@societe.fr" width={240} align="left" /></FieldRow>
        <FieldRow label="Adresse"><TextInput defaultValue={contact.adresse} onCommit={(v) => patch("adresse", v)} placeholder="Numéro et rue" width={280} align="left" /></FieldRow>
        <FieldRow label="Code postal"><TextInput defaultValue={contact.codePostal} onCommit={(v) => patch("codePostal", v)} placeholder="75000" width={100} /></FieldRow>
        <FieldRow label="Ville"><TextInput defaultValue={contact.ville} onCommit={(v) => patch("ville", v)} placeholder="Paris" width={180} align="left" /></FieldRow>
      </Panel>
    </div>
  );
}

const LAB_RISQUE_OPTIONS = ["Faible", "Moyen", "Élevé"];
function CorporateTab({ client, onUpdate }) {
  const corp = client.corporate || { kyc: {}, kycExtra: [], notes: "" };
  const kyc = corp.kyc || {};
  const patchKyc = (patch) => onUpdate(client.id, { corporate: { ...corp, kyc: { ...kyc, ...patch } } });
  const [newItem, setNewItem] = useState("");
  const addExtra = () => {
    if (!newItem.trim()) return;
    onUpdate(client.id, { corporate: { ...corp, kycExtra: [...(corp.kycExtra || []), { label: newItem.trim(), done: false }] } });
    setNewItem("");
  };
  const toggleExtra = (idx) => {
    const list = [...(corp.kycExtra || [])]; list[idx] = { ...list[idx], done: !list[idx].done };
    onUpdate(client.id, { corporate: { ...corp, kycExtra: list } });
  };
  const removeExtra = (idx) => {
    const list = (corp.kycExtra || []).filter((_, i) => i !== idx);
    onUpdate(client.id, { corporate: { ...corp, kycExtra: list } });
  };

  // Le LAB doit être révisé périodiquement (recommandé : tous les 12 mois)
  const labAJour = kyc.lab && kyc.labDate;
  const moisDepuisLab = kyc.labDate ? Math.floor((Date.now() - new Date(kyc.labDate).getTime()) / (1000 * 60 * 60 * 24 * 30)) : null;
  const labARevisor = labAJour && moisDepuisLab !== null && moisDepuisLab >= 12;

  const extraItems = corp.kycExtra || [];
  const extraDone = extraItems.filter((i) => i.done).length;
  const coreItems = [!!kyc.lab, !!kyc.mandat, !!kyc.beneficiaireEffectif];
  const coreDone = coreItems.filter(Boolean).length;
  const totalItems = coreItems.length + extraItems.length;
  const totalDone = coreDone + extraDone;

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
        <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: 0, display: "flex", alignItems: "center", gap: 6 }}>
          <ShieldCheck size={15} /> Base corporate / LAB
        </h4>
        <Stamped tone={labARevisor ? "red" : totalDone === totalItems ? "green" : "amber"} small>
          {labARevisor ? "LAB à réviser" : totalDone === totalItems ? "Dossier complet" : `${totalDone}/${totalItems} pièces`}
        </Stamped>
      </div>
      <div style={{ height: 6, borderRadius: 4, background: T.paperDeep, overflow: "hidden", marginBottom: 18 }}>
        <div style={{ width: `${totalItems ? (totalDone / totalItems) * 100 : 0}%`, height: "100%", background: T.navy }} />
      </div>

      <FieldRow label="Questionnaire LAB complété">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <ToggleBtn on={!!kyc.lab} onClick={() => patchKyc({ lab: !kyc.lab, labDate: !kyc.lab ? todayISO() : kyc.labDate })} />
          {kyc.lab && (
            <input type="date" value={kyc.labDate || ""} onChange={(e) => patchKyc({ labDate: e.target.value })}
              style={{ fontFamily: T.mono, fontSize: 12, padding: "4px 7px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />
          )}
        </div>
      </FieldRow>
      {labARevisor && (
        <div style={{ fontSize: 11.5, color: T.red, background: T.redSoft, padding: "7px 10px", borderRadius: 8, margin: "0 0 10px" }}>
          Questionnaire LAB daté de plus de 12 mois — une révision est recommandée.
        </div>
      )}
      <FieldRow label="Niveau de risque LAB">
        <SelectPill value={kyc.risqueLab} options={LAB_RISQUE_OPTIONS} onChange={(v) => patchKyc({ risqueLab: v })} />
      </FieldRow>
      <FieldRow label="Mandat signé">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <ToggleBtn on={!!kyc.mandat} onClick={() => patchKyc({ mandat: !kyc.mandat, mandatDate: !kyc.mandat ? todayISO() : kyc.mandatDate })} />
          {kyc.mandat && (
            <input type="date" value={kyc.mandatDate || ""} onChange={(e) => patchKyc({ mandatDate: e.target.value })}
              style={{ fontFamily: T.mono, fontSize: 12, padding: "4px 7px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />
          )}
        </div>
      </FieldRow>
      <FieldRow label="Choix du PA"><TextInput defaultValue={kyc.choixPA} onCommit={(v) => patchKyc({ choixPA: v })} placeholder="Prestataire agréé retenu" width={180} align="left" /></FieldRow>
      <FieldRow label="Bénéficiaire effectif identifié"><ToggleBtn on={!!kyc.beneficiaireEffectif} onClick={() => patchKyc({ beneficiaireEffectif: !kyc.beneficiaireEffectif })} /></FieldRow>
      <FieldRow label="Nom du bénéficiaire effectif"><TextInput defaultValue={kyc.beneficiaireNom} onCommit={(v) => patchKyc({ beneficiaireNom: v })} placeholder="Nom, prénom" width={180} align="left" /></FieldRow>

      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Autres pièces à suivre</h4>
      {(corp.kycExtra || []).map((it, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 4px", borderBottom: `1px solid ${T.line}` }}>
          <span onClick={() => toggleExtra(i)} className="clickable" style={{
            width: 19, height: 19, borderRadius: 5, border: `1.5px solid ${it.done ? T.green : T.line}`, background: it.done ? T.green : "transparent",
            display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
          }}>{it.done && <Check size={13} color="#fff" strokeWidth={3} />}</span>
          <span style={{ flex: 1, fontSize: 12, textDecoration: it.done ? "line-through" : "none", color: it.done ? T.inkMuted : T.ink }}>{it.label}</span>
          <button onClick={() => removeExtra(i)} style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted }}><Trash2 size={13} /></button>
        </div>
      ))}
      <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
        <input value={newItem} onChange={(e) => setNewItem(e.target.value)} placeholder="Ajouter un document à suivre…" style={{ ...inputStyle, flex: 1, padding: "7px 10px" }} onKeyDown={(e) => e.key === "Enter" && addExtra()} />
        <button onClick={addExtra} style={{ background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "7px 12px", cursor: "pointer" }}><Plus size={14} /></button>
      </div>

      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Notes</h4>
      <textarea defaultValue={corp.notes} onBlur={(e) => onUpdate(client.id, { corporate: { ...corp, notes: e.target.value } })}
        placeholder="Notes libres sur le dossier corporate…" rows={4}
        style={{ width: "100%", padding: 10, borderRadius: 10, border: `1px solid ${T.line}`, fontSize: 12, background: T.card, resize: "vertical" }} />
    </div>
  );
}

function TvaTab({ client, onUpdate }) {
  const currentMonth = client.tvaRegime === "CA12" ? "Mai" : currentMonthKey();
  const currentStatus = effectiveTvaStatus(client, currentMonth);
  const currentNote = client.tvaControle?.[currentMonth]?.commentaire || "";
  return (
    <div>
      <FieldRow label="Régime TVA"><SelectPill value={client.tvaRegime} options={REGIMES_TVA} labels={REGIMES_TVA_LABELS} onChange={(v) => onUpdate(client.id, { tvaRegime: v, tvaPeriodicite: v === "CA3" ? (client.tvaPeriodicite || "mensuelle") : client.tvaPeriodicite })} /></FieldRow>
      {client.tvaRegime === "CA3" && (
        <FieldRow label="Périodicité de déclaration">
          <SelectPill value={client.tvaPeriodicite || "mensuelle"} options={TVA_PERIODICITES} labels={TVA_PERIODICITE_LABELS} allowEmpty={false} onChange={(v) => onUpdate(client.id, { tvaPeriodicite: v })} />
        </FieldRow>
      )}
      {client.tvaRegime !== "CA12" && <FieldRow label="Jour limite de déclaration">
        <input type="number" min="1" max="31" defaultValue={client.tvaExig || ""} placeholder="ex. 19"
          onBlur={(e) => onUpdate(client.id, { tvaExig: e.target.value ? parseInt(e.target.value, 10) : "" })}
          style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, width: 60, textAlign: "right" }} />
      </FieldRow>}
      <FieldRow label="Statut courant">
        <Stamped tone={tvaTone(currentStatus)} small>{tvaStatusLabel(currentStatus)}</Stamped>
      </FieldRow>
      {currentStatus === "NON_VALIDE" && currentNote && (
        <FieldRow label="Remarques du contrôle">
          <div style={{ fontSize: 12, color: T.ink, background: T.redSoft || "#FEECEC", border: `1px solid ${T.line}`, borderRadius: 8, padding: "8px 10px", maxWidth: 320 }}>{currentNote}</div>
        </FieldRow>
      )}
      <div style={{ height: 14 }} />
      <Panel title="Paiement de TVA">
        <PaymentLine
          label={client.tvaRegime === "CA12" ? "TVA annuelle — CA12" : `TVA ${currentMonthKey()}`}
          amount={client.tvaPaiements?.[client.tvaRegime === "CA12" ? "CA12" : currentMonthKey()]?.montant ?? ""}
          status={client.tvaPaiements?.[client.tvaRegime === "CA12" ? "CA12" : currentMonthKey()]?.statut || "a_payer"}
          onAmountChange={(value) => { const k=client.tvaRegime === "CA12" ? "CA12" : currentMonthKey(); onUpdate(client.id, { tvaPaiements: { ...(client.tvaPaiements || {}), [k]: { ...(client.tvaPaiements?.[k] || {}), montant: value } } }) }}
          onStatusChange={(value) => { const k=client.tvaRegime === "CA12" ? "CA12" : currentMonthKey(); onUpdate(client.id, { tvaPaiements: { ...(client.tvaPaiements || {}), [k]: { ...(client.tvaPaiements?.[k] || {}), statut: value } } }) }}
        />
      </Panel>
      <div style={{ fontSize: 12, color: T.inkMuted, margin: "14px 0 0", lineHeight: 1.6 }}>
        {client.tvaRegime === "CA12"
          ? "Régime CA12 : une seule déclaration annuelle, exigible en Mai N+1."
          : client.tvaRegime === "CA3"
            ? client.tvaPeriodicite === "trimestrielle"
              ? "Régime CA3 trimestriel : une déclaration à la fin de chaque trimestre civil (Mars, Juin, Septembre, Décembre), exigible le mois suivant (M+1). Les autres mois sont non applicables."
              : "Régime CA3 mensuel : la TVA d'un mois donné est déclarée le mois suivant (M+1)."
            : "Sélectionnez un régime TVA pour activer le suivi des échéances."}
        {" "}Le suivi mois par mois (Fait / OK / N/A) se gère depuis l'écran <strong>TVA — CA3/CA12</strong>.
      </div>
    </div>
  );
}

function BilanTab({ client, onUpdate }) {
  const b = client.bilan || {};
  const patch = (fields) => onUpdate(client.id, { bilan: { ...b, ...fields } });
  const toggle = (field) => patch({ [field]: !b[field] });
  const toggleTransmission = () => {
    const next = !b.transmis;
    if (next && !b.transmis && client.dateCloture) {
      onUpdate(client.id, {
        bilan: { ...b, transmis: true, transmisDate: todayISO() },
        dateCloture: addYearISO(client.dateCloture, 1),
      });
      return;
    }
    patch({ transmis: false });
  };
  const echeance = getBilanEcheance(client.dateCloture);
  const statut = getBilanStatut(b, client.dateCloture);
  const enRetard = echeance && todayISO() > echeance && !b.transmis;

  return (
    <div>
      {/* 1) Cartouches de statut */}
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 20 }}>
        <div style={{ flex: "1 1 160px", border: `1px solid ${T.line}`, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 10.5, color: T.inkMuted, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>Statut</div>
          <Stamped tone={statut.tone}>{statut.label}</Stamped>
        </div>
        <div style={{ flex: "1 1 160px", border: `1px solid ${T.line}`, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 10.5, color: T.inkMuted, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>Échéance légale (clôture + 3 mois)</div>
          <div style={{ fontFamily: T.mono, fontSize: 13.5, fontWeight: 700, color: enRetard ? T.red : T.ink }}>{fmtFR(echeance)}</div>
        </div>
        <div style={{ flex: "1 1 160px", border: `1px solid ${T.line}`, borderRadius: 10, padding: "12px 14px" }}>
          <div style={{ fontSize: 10.5, color: T.inkMuted, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 6 }}>Date de clôture</div>
          <div style={{ fontFamily: T.mono, fontSize: 13.5, fontWeight: 700 }}>{fmtFR(client.dateCloture)}</div>
        </div>
      </div>

      {/* 2) Étapes d'avancement */}
      <Panel title="Étapes d'avancement">
        <FieldRow label="Révision comptable">
          <div style={{ display: "flex", gap: 6 }}>
            {BILAN_REVISION_STEPS.map((s) => (
              <button key={s.id} onClick={() => patch({ revision: s.id })} style={{
                fontSize: 11.5, fontWeight: 700, padding: "5px 11px", borderRadius: 999, cursor: "pointer",
                border: `1px solid ${(b.revision || "a_faire") === s.id ? T.navy : T.line}`,
                background: (b.revision || "a_faire") === s.id ? T.navy + "1A" : "transparent",
                color: (b.revision || "a_faire") === s.id ? T.navy : T.inkMuted,
              }}>{s.label}</button>
            ))}
          </div>
        </FieldRow>
        <FieldRow label="Révision de fin de bilan">
          {(() => {
            const checklist = b.finBilanChecklist || {};
            const items = [
              ["capitauxPropres", "Capitaux propres / capital social"],
              ["marges", "Marges et variations anormales"],
              ["cfe", "CFE et comptabilisation"],
              ["tvaCadrage", "Cadrage de TVA"],
              ["banque", "Rapprochements bancaires"],
              ["fournisseurs", "Comptes fournisseurs et soldes anciens"],
              ["clients", "Comptes clients et créances anciennes"],
              ["social", "Comptes sociaux, charges et dettes"],
              ["emprunts", "Emprunts et intérêts"],
              ["immobilisations", "Immobilisations et amortissements"],
              ["chargesProduits", "Charges / produits à rattacher"],
              ["comptesAttente", "Comptes d'attente et comptes divers"],
              ["impots", "IS/IR, CFE et autres impôts"],
              ["annexes", "Annexes et éléments de liasse"],
            ];
            const done = items.filter(([id]) => checklist[id]).length;
            return (
              <div style={{ width: "100%", border: `1px solid ${T.line}`, borderRadius: 10, padding: "9px 11px", background: T.paper }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, marginBottom: 7 }}>
                  <span style={{ fontSize: 11, color: T.inkMuted }}>{done}/{items.length} contrôles réalisés</span>
                  <Stamped tone={done === items.length ? "green" : "amber"} small>{done === items.length ? "Révision complète" : "À finaliser"}</Stamped>
                </div>
                {items.map(([id, label]) => (
                  <label key={id} style={{ display: "flex", alignItems: "center", gap: 8, padding: "5px 2px", borderBottom: `1px solid ${T.line}`, fontSize: 11.5, color: checklist[id] ? T.inkMuted : T.ink, textDecoration: checklist[id] ? "line-through" : "none", cursor: "pointer" }}>
                    <input type="checkbox" checked={!!checklist[id]} onChange={() => patch({ finBilanChecklist: { ...checklist, [id]: !checklist[id] } })} />
                    <span>{label}</span>
                  </label>
                ))}
              </div>
            );
          })()}
        </FieldRow>
        <FieldRow label="Projet de bilan validé par le client">
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <ToggleBtn on={!!b.valideClient} onClick={() => patch({ valideClient: !b.valideClient, valideClientDate: !b.valideClient ? todayISO() : b.valideClientDate })} />
            {b.valideClient && (
              <input type="date" value={b.valideClientDate || ""} onChange={(e) => patch({ valideClientDate: e.target.value })}
                style={{ fontFamily: T.mono, fontSize: 12, padding: "4px 7px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />
            )}
          </div>
        </FieldRow>
        <FieldRow label="Transmis (liasse télétransmise)">
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <ToggleBtn on={!!b.transmis} onClick={toggleTransmission} tone="green" />
            {b.transmis && (
              <input type="date" value={b.transmisDate || ""} onChange={(e) => patch({ transmisDate: e.target.value })}
                style={{ fontFamily: T.mono, fontSize: 12, padding: "4px 7px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />
            )}
          </div>
        </FieldRow>
      </Panel>

      <div style={{ height: 14 }} />

      {/* 3) Données financières clés */}
      <Panel title="Données financières clés">
        <FieldRow label="Chiffre d'affaires (CA)">
          <input type="number" defaultValue={b.ca ?? ""} onBlur={(e) => patch({ ca: e.target.value })} placeholder="0"
            style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, width: 140 }} />
        </FieldRow>
        <FieldRow label="Résultat net (bénéfice / perte)">
          <input type="number" defaultValue={b.resultat ?? ""} onBlur={(e) => patch({ resultat: e.target.value })} placeholder="0"
            style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, width: 140 }} />
        </FieldRow>
        <FieldRow label="Capital social">
          <input type="number" defaultValue={client.capitalSocial ?? ""} onBlur={(e) => onUpdate(client.id, { capitalSocial: e.target.value })} placeholder="0"
            style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, width: 140 }} />
        </FieldRow>
        <FieldRow label="Capitaux propres">
          <input type="number" defaultValue={b.capitauxPropres ?? ""} onBlur={(e) => patch({ capitauxPropres: e.target.value })} placeholder="0"
            style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, width: 140 }} />
        </FieldRow>
        <FieldRow label="Trésorerie finale">
          <input type="number" defaultValue={b.tresorerie ?? ""} onBlur={(e) => patch({ tresorerie: e.target.value })} placeholder="0"
            style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, width: 140 }} />
        </FieldRow>
        {(b.ca !== undefined && b.ca !== "") || (b.resultat !== undefined && b.resultat !== "") ? (
          <div style={{ display: "flex", gap: 18, marginTop: 10, paddingTop: 10, borderTop: `1px solid ${T.line}`, flexWrap: "wrap" }}>
            <div><span style={{ fontSize: 11, color: T.inkMuted }}>CA </span><strong style={{ fontFamily: T.mono }}>{fmtEUR(b.ca)}</strong></div>
            <div><span style={{ fontSize: 11, color: T.inkMuted }}>Résultat </span><strong style={{ fontFamily: T.mono, color: Number(b.resultat) < 0 ? T.red : T.green }}>{fmtEUR(b.resultat)}</strong></div>
            {b.capitauxPropres !== undefined && b.capitauxPropres !== "" && <div><span style={{ fontSize: 11, color: T.inkMuted }}>Capitaux propres </span><strong style={{ fontFamily: T.mono }}>{fmtEUR(b.capitauxPropres)}</strong></div>}
            {b.tresorerie !== undefined && b.tresorerie !== "" && <div><span style={{ fontSize: 11, color: T.inkMuted }}>Trésorerie </span><strong style={{ fontFamily: T.mono }}>{fmtEUR(b.tresorerie)}</strong></div>}
          </div>
        ) : null}
      </Panel>

      <div style={{ height: 14 }} />

      {/* 4) Legacy — ne pas supprimer : alimente le tableau de bord et les échéances fiscales */}
      <Panel title="Suivi du retard (utilisé par le tableau de bord)">
        <FieldRow label="Finalisé après échéance"><ToggleBtn on={!!b.finaliseApres} onClick={() => toggle("finaliseApres")} /></FieldRow>
        <FieldRow label="Non encore finalisé (en retard)"><ToggleBtn on={!!b.nonFinalise} onClick={() => toggle("nonFinalise")} tone="red" /></FieldRow>
        <FieldRow label="Courrier de retard signé et classé"><ToggleBtn on={!!b.courrier} onClick={() => toggle("courrier")} /></FieldRow>
      </Panel>
    </div>
  );
}
function ToggleBtn({ on, onClick, tone = "green" }) {
  return <button onClick={onClick} style={{ display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer", background: "none", padding: 0 }}>
    <Stamped tone={on ? tone : "neutral"} small>{on ? "Oui" : "Non"}</Stamped>
  </button>;
}

const PAYMENT_STATUS_OPTIONS = [
  { value: "a_payer", label: "À payer" },
  { value: "paye", label: "Payé" },
  { value: "partiellement_paye", label: "Partiellement payé" },
  { value: "non_concerne", label: "Non concerné" },
];

function PaymentStatus({ value, onChange }) {
  return (
    <select value={value || "a_payer"} onChange={(e) => onChange(e.target.value)}
      style={{ fontSize: 11.5, padding: "5px 8px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card, minWidth: 145 }}>
      {PAYMENT_STATUS_OPTIONS.map((o) => <option key={o.value} value={o.value}>{o.label}</option>)}
    </select>
  );
}

function PaymentLine({ label, amount, status, onAmountChange, onStatusChange }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "1.4fr 120px 165px", gap: 10, alignItems: "center", padding: "8px 0", borderBottom: `1px solid ${T.line}` }}>
      <div style={{ fontSize: 11.5, fontWeight: 600 }}>{label}</div>
      <input type="number" min="0" step="0.01" value={amount ?? ""} placeholder="Montant" onChange={(e) => onAmountChange(e.target.value)}
        style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card, width: "100%", textAlign: "right" }} />
      <PaymentStatus value={status} onChange={onStatusChange} />
    </div>
  );
}

function AcomptesTab({ client, onUpdate }) {
  const is = client.is || {}; const cfe = client.cfe || {};
  const toggleIs = (f) => onUpdate(client.id, { is: { ...is, [f]: !is[f] } });
  const toggleCfe = (f) => onUpdate(client.id, { cfe: { ...cfe, [f]: !cfe[f] } });
  const updatePayment = (tax, key, field, value) => {
    const source = tax === "is" ? is : cfe;
    onUpdate(client.id, { [tax]: { ...source, paiements: { ...(source.paiements || {}), [key]: { ...(source.paiements?.[key] || {}), [field]: value } } } });
  };
  const isConcerne = Number(is.montantN1) > 3000;
  const cfeConcerne = Number(cfe.montantN1) > 3000;
  const numInputStyle = { fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, width: 110, textAlign: "right" };
  return (
    <div>
      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "4px 0 8px" }}>Impôt sur les sociétés</h4>
      <FieldRow label="Montant IS N-1">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <input type="number" defaultValue={is.montantN1 ?? ""} placeholder="0"
            onBlur={(e) => {
              const v = e.target.value;
              onUpdate(client.id, { is: { ...is, montantN1: v, concerne: Number(v) > 3000 } });
            }} style={numInputStyle} />
          <Stamped tone={isConcerne ? "amber" : "neutral"} small>{isConcerne ? "Concerné (> 3000€)" : "Non concerné"}</Stamped>
        </div>
      </FieldRow>
      <Panel title="Paiements IS">
        {[
          ["mars", "Acompte IS — mars"],
          ["juin", "Acompte IS — juin"],
          ["sept", "Acompte IS — septembre"],
          ["dec", "Acompte IS — décembre"],
          ["solde", "Solde IS"],
        ].map(([key, label]) => (
          <PaymentLine key={key} label={label} amount={is.paiements?.[key]?.montant ?? (key === "solde" ? "" : (Number(is.montantN1) > 3000 ? (Number(is.montantN1)/4).toFixed(2) : ""))} status={is.paiements?.[key]?.statut || (is[key] ? "paye" : "a_payer")}
            onAmountChange={(v) => updatePayment("is", key, "montant", v)}
            onStatusChange={(v) => updatePayment("is", key, "statut", v)} />
        ))}
      </Panel>
      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>CFE</h4>
      <FieldRow label="Montant CFE N-1">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <input type="number" defaultValue={cfe.montantN1 ?? ""} placeholder="0"
            onBlur={(e) => {
              const v = e.target.value;
              onUpdate(client.id, { cfe: { ...cfe, montantN1: v, concerne: Number(v) > 3000 } });
            }} style={numInputStyle} />
          <Stamped tone={cfeConcerne ? "amber" : "neutral"} small>{cfeConcerne ? "Concerné (> 3000€)" : "Non concerné"}</Stamped>
        </div>
      </FieldRow>
      <FieldRow label="Acompte juin"><ToggleBtn on={!!cfe.juin} onClick={() => toggleCfe("juin")} /></FieldRow>
      <FieldRow label="Solde décembre"><ToggleBtn on={!!cfe.dec} onClick={() => toggleCfe("dec")} /></FieldRow>
      <div style={{ height: 12 }} />
      <Panel title="Paiements CFE">
        {[["juin", "Acompte CFE — juin"], ["dec", "Solde CFE — décembre"]].map(([key, label]) => (
          <PaymentLine key={key} label={label} amount={cfe.paiements?.[key]?.montant ?? (key === "juin" && Number(cfe.montantN1)>3000 ? (Number(cfe.montantN1)/2).toFixed(2) : "")} status={cfe.paiements?.[key]?.statut || (cfe[key] ? "paye" : "a_payer")}
            onAmountChange={(v) => updatePayment("cfe", key, "montant", v)}
            onStatusChange={(v) => updatePayment("cfe", key, "statut", v)} />
        ))}
      </Panel>
    </div>
  );
}

function MissionTab({ client, onUpdate }) {
  const m = client.mission || {};
  const groups = MISSION_GROUPS;
  const keys = MISSION_ALL_KEYS;
  const toggle = (k) => onUpdate(client.id, { mission: { ...m, [k]: !m[k] } });
  const done = keys.filter((k) => m[k]).length;
  const total = keys.length;

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 22 }}>
        <span style={{ fontFamily: T.serif, fontWeight: 700, fontSize: 13.5, color: T.ink }}>Progression de l'accueil</span>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <span style={{ fontFamily: T.mono, fontSize: 12, fontWeight: 600, color: T.inkMuted }}>{done}/{total}</span>
          <div style={{ width: 92, height: 6, borderRadius: 4, background: T.paperDeep, overflow: "hidden" }}>
            <div style={{ width: `${total ? (done / total) * 100 : 0}%`, height: "100%", background: T.navy, borderRadius: 4 }} />
          </div>
        </div>
      </div>

      {groups.map((group, gi) => {
        const groupKeys = group.keys;
        const groupDone = groupKeys.filter((k) => m[k]).length;
        const groupTotal = groupKeys.length;
        const isComplete = groupTotal > 0 && groupDone === groupTotal;
        const isStarted = groupDone > 0;
        const isLast = gi === groups.length - 1;
        return (
          <div key={group.title} style={{ display: "flex", gap: 16 }}>
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", flexShrink: 0 }}>
              <div style={{
                width: 30, height: 30, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center",
                background: isComplete ? T.green : isStarted ? T.navy : T.paperDeep,
                color: isComplete || isStarted ? "#fff" : T.inkMuted,
                fontSize: 12, fontWeight: 700, flexShrink: 0,
                boxShadow: `0 0 0 4px ${isComplete ? T.greenSoft : isStarted ? T.navySoft : "#fff"}`,
              }}>
                {isComplete ? <Check size={14} strokeWidth={3} /> : gi + 1}
              </div>
              {!isLast && <div style={{ width: 2, flex: 1, background: T.line, margin: "4px 0", minHeight: 26 }} />}
            </div>
            <div style={{ flex: 1, paddingBottom: isLast ? 4 : 22 }}>
              <div style={{ display: "flex", alignItems: "baseline", gap: 8, marginBottom: 9 }}>
                <span style={{ fontSize: 13, fontWeight: 700, color: T.ink }}>{group.title}</span>
                <span style={{ fontSize: 11, color: T.inkMuted, fontWeight: 600 }}>{groupDone}/{groupTotal}</span>
              </div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                {groupKeys.map((k) => {
                  const on = !!m[k];
                  return (
                    <span key={k} onClick={() => toggle(k)} className="clickable" style={{
                      display: "flex", alignItems: "center", gap: 6, padding: "6px 12px", borderRadius: 20,
                      fontSize: 12, fontWeight: 600, border: "1px solid transparent",
                      background: on ? T.greenSoft : T.paperDeep,
                      color: on ? "#0F7A50" : T.inkMuted,
                    }}>
                      <span style={{ width: 6, height: 6, borderRadius: "50%", background: on ? "#0F7A50" : T.inkMuted, flexShrink: 0 }} />
                      {k}
                    </span>
                  );
                })}
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
/* ============================================================
   NOTES COLLABORATIVES — journal par dossier, visible et
   alimenté par tous les collaborateurs. Append-only : on ajoute,
   on ne modifie/supprime pas l'historique.
   ============================================================ */
function NotesTab({ client, me, meId, portefeuilleId, onUpdate }) {
  const [texte, setTexte] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editTexte, setEditTexte] = useState("");
  const notes = client.notesCollab || [];
  const sorted = [...notes].sort((a, b) => (a.date < b.date ? 1 : -1));

  const addNote = () => {
    if (!texte.trim()) return;
    const entry = { id: `n-${Date.now()}`, texte: texte.trim(), auteur: me, date: new Date().toISOString() };
    onUpdate(client.id, { notesCollab: [...notes, entry] });
    logActivity({ clientId: client.id, portefeuilleId, type: "note", message: "Note ajoutée", auteurId: meId });
    setTexte("");
  };

  const startEdit = (n) => { setEditingId(n.id); setEditTexte(n.texte); };
  const cancelEdit = () => { setEditingId(null); setEditTexte(""); };
  const saveEdit = (id) => {
    if (!editTexte.trim()) return;
    onUpdate(client.id, { notesCollab: notes.map((n) => (n.id === id ? { ...n, texte: editTexte.trim() } : n)) });
    logActivity({ clientId: client.id, portefeuilleId, type: "note", message: "Note modifiée", auteurId: meId });
    cancelEdit();
  };
  const removeNote = (id) => {
    if (!confirm("Supprimer cette note ?")) return;
    onUpdate(client.id, { notesCollab: notes.filter((n) => n.id !== id) });
    logActivity({ clientId: client.id, portefeuilleId, type: "note", message: "Note supprimée", auteurId: meId });
  };

  return (
    <div>
      <div style={{ marginBottom: 16 }}>
        <textarea value={texte} onChange={(e) => setTexte(e.target.value)} rows={3}
          placeholder="Un besoin, une info à transmettre à l'équipe sur ce dossier…"
          style={{ width: "100%", padding: 10, borderRadius: 10, border: `1px solid ${T.line}`, fontSize: 12.5, background: T.card, resize: "vertical" }} />
        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 8 }}>
          <button onClick={addNote} disabled={!texte.trim()} style={{
            display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10,
            padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer", opacity: texte.trim() ? 1 : 0.6,
          }}>
            <Plus size={14} /> Ajouter la note
          </button>
        </div>
      </div>

      {sorted.length === 0 ? <EmptyNote text="Aucune note pour ce dossier pour l'instant." /> : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {sorted.map((n) => (
            <div key={n.id} style={{ border: `1px solid ${T.line}`, borderRadius: 10, padding: "10px 12px", background: T.paper }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 5 }}>
                <span style={{ fontWeight: 700, fontSize: 12 }}>{n.auteur}</span>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted }}>
                    {new Date(n.date).toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric" })}
                    {" · "}
                    {new Date(n.date).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}
                  </span>
                  {n.auteur === me && editingId !== n.id && (
                    <>
                      <button onClick={() => startEdit(n)} title="Modifier" style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted, display: "flex" }}>
                        <Pencil size={13} />
                      </button>
                      <button onClick={() => removeNote(n.id)} title="Supprimer" style={{ background: "none", border: "none", cursor: "pointer", color: T.red, display: "flex" }}>
                        <Trash2 size={13} />
                      </button>
                    </>
                  )}
                </div>
              </div>
              {editingId === n.id ? (
                <div>
                  <textarea value={editTexte} onChange={(e) => setEditTexte(e.target.value)} rows={3}
                    style={{ width: "100%", padding: 8, borderRadius: 8, border: `1px solid ${T.line}`, fontSize: 12.5, background: T.card, resize: "vertical" }} />
                  <div style={{ display: "flex", justifyContent: "flex-end", gap: 6, marginTop: 6 }}>
                    <button onClick={cancelEdit} style={{ background: "none", border: `1px solid ${T.line}`, borderRadius: 8, padding: "5px 10px", fontSize: 11.5, cursor: "pointer", color: T.inkMuted }}>Annuler</button>
                    <button onClick={() => saveEdit(n.id)} disabled={!editTexte.trim()} style={{ background: T.navy, color: "#fff", border: "none", borderRadius: 8, padding: "5px 10px", fontSize: 11.5, fontWeight: 600, cursor: "pointer", opacity: editTexte.trim() ? 1 : 0.6 }}>Enregistrer</button>
                  </div>
                </div>
              ) : (
                <div style={{ fontSize: 12.5, color: T.inkSoft, whiteSpace: "pre-wrap", lineHeight: 1.5 }}>{n.texte}</div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
const RESILIATION_MOTIFS = ["Impayés", "Cessation d'activité du client", "Désaccord", "Changement de cabinet", "Autre"];
/* ============================================================
   HISTORIQUE / AUDIT TRAIL — fil d'activité par dossier, lu
   directement depuis la table Supabase `activity_log` (déjà
   alimentée par logActivity sur les tâches, notes, résiliation,
   reprise, honoraires…).
   ============================================================ */
const ACTIVITY_TYPE_LABELS = {
  tache: "Tâche", note: "Note", resiliation: "Résiliation", reprise: "Reprise",
  honoraires: "Honoraires", mission: "Accueil", tva: "TVA", social: "Social",
};
function HistoriqueTab({ clientId, team }) {
  const [rows, setRows] = useState(null);
  useEffect(() => {
    let cancelled = false;
    setRows(null);
    supabase.from("activity_log").select("*").eq("client_id", clientId).order("created_at", { ascending: false }).limit(80)
      .then(({ data, error }) => {
        if (cancelled) return;
        if (error) { console.error("Erreur chargement historique :", error.message); setRows([]); return; }
        setRows(data || []);
      });
    return () => { cancelled = true; };
  }, [clientId]);

  const nameFor = (id) => team.find((t) => t.id === id)?.nom || "—";

  if (rows === null) return <div style={{ fontSize: 12.5, color: T.inkMuted, padding: "8px 0" }}>Chargement…</div>;
  if (rows.length === 0) return <EmptyNote text="Aucun événement enregistré pour ce dossier pour l'instant." />;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
      {rows.map((r) => (
        <div key={r.id} style={{ display: "flex", gap: 10, alignItems: "flex-start", padding: "9px 4px", borderBottom: `1px solid ${T.line}` }}>
          <span style={{ marginTop: 2 }}><History size={13} color={T.inkMuted} /></span>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 12.5, color: T.ink }}>{r.message}</div>
            <div style={{ fontSize: 10.5, color: T.inkMuted, fontFamily: T.mono, marginTop: 2 }}>
              {ACTIVITY_TYPE_LABELS[r.type] || r.type} · {nameFor(r.auteur_id)} · {new Date(r.created_at).toLocaleDateString("fr-FR", { day: "2-digit", month: "2-digit", year: "numeric" })}
              {" "}{new Date(r.created_at).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
const RESILIATION_INITIATEURS = ["Cabinet", "Client"];

function ResiliationTab({ client, me, meId, portefeuilleId, onUpdate }) {
  const r = client.resiliation || {};
  const patch = (fields) => onUpdate(client.id, { resiliation: { ...r, ...fields } });

  const activer = () => {
    const entry = { date: r.date || todayISO(), initiateur: r.initiateur, motif: r.motif === "Autre" ? r.motifAutre : r.motif, par: me };
    patch({ active: true, historique: [...(r.historique || []), entry] });
    // Statut intermédiaire : le dossier est en cours de sortie mais pas encore totalement clos côté cabinet.
    onUpdate(client.id, { statutDossier: "transfert" });
    logActivity({ clientId: client.id, portefeuilleId, type: "resiliation", message: `Résiliation démarrée (motif : ${entry.motif || "—"})`, auteurId: meId });
  };
  const annuler = () => {
    patch({ active: false });
    onUpdate(client.id, { statutDossier: "actif" });
    logActivity({ clientId: client.id, portefeuilleId, type: "resiliation", message: "Résiliation annulée", auteurId: meId });
  };
  const finaliser = () => {
    onUpdate(client.id, { statutDossier: "inactif" });
    logActivity({ clientId: client.id, portefeuilleId, type: "resiliation", message: "Sortie du dossier finalisée (Inactif)", auteurId: meId });
  };

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
        <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: 0 }}>Résiliation du dossier</h4>
        <Stamped tone={r.active ? "red" : "green"} small>{r.active ? "Dossier résilié" : "Dossier actif"}</Stamped>
      </div>

      {r.active && (
        <div style={{ fontSize: 11.5, color: T.red, background: T.redSoft, padding: "8px 12px", borderRadius: 9, marginBottom: 16 }}>
          Ce dossier est marqué comme résilié — le statut a été basculé sur « En transfert » en attendant la clôture complète.
          <button onClick={annuler} style={{ marginLeft: 10, background: "none", border: "none", color: T.navy, fontWeight: 700, cursor: "pointer", fontSize: 11.5 }}>Annuler la résiliation</button>
          {r.piecesRestituees && (
            <button onClick={finaliser} style={{ marginLeft: 10, background: "none", border: "none", color: T.red, fontWeight: 700, cursor: "pointer", fontSize: 11.5 }}>Finaliser la sortie (Inactif)</button>
          )}
        </div>
      )}

      <FieldRow label="Date de résiliation">
        <input type="date" value={r.date || ""} onChange={(e) => patch({ date: e.target.value })}
          style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} />
      </FieldRow>
      <FieldRow label="Initiateur"><SelectPill value={r.initiateur} options={RESILIATION_INITIATEURS} onChange={(v) => patch({ initiateur: v })} /></FieldRow>
      <FieldRow label="Motif"><SelectPill value={r.motif} options={RESILIATION_MOTIFS} onChange={(v) => patch({ motif: v })} /></FieldRow>
      {r.motif === "Autre" && <FieldRow label="Précisez"><TextInput defaultValue={r.motifAutre} onCommit={(v) => patch({ motifAutre: v })} width={200} align="left" /></FieldRow>}

      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Obligations légales & déontologiques</h4>
      <FieldRow label="Lettre de résiliation envoyée">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <ToggleBtn on={!!r.lettreEnvoyee} onClick={() => patch({ lettreEnvoyee: !r.lettreEnvoyee, lettreDate: !r.lettreEnvoyee ? todayISO() : r.lettreDate })} />
          {r.lettreEnvoyee && <input type="date" value={r.lettreDate || ""} onChange={(e) => patch({ lettreDate: e.target.value })}
            style={{ fontFamily: T.mono, fontSize: 12, padding: "4px 7px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />}
        </div>
      </FieldRow>
      <FieldRow label="Préavis contractuel respecté"><ToggleBtn on={!!r.preavisRespecte} onClick={() => patch({ preavisRespecte: !r.preavisRespecte })} /></FieldRow>
      <FieldRow label="Pièces comptables restituées au client">
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <ToggleBtn on={!!r.piecesRestituees} onClick={() => patch({ piecesRestituees: !r.piecesRestituees, piecesRestitueesDate: !r.piecesRestituees ? todayISO() : r.piecesRestitueesDate })} tone="green" />
          {r.piecesRestituees && <input type="date" value={r.piecesRestitueesDate || ""} onChange={(e) => patch({ piecesRestitueesDate: e.target.value })}
            style={{ fontFamily: T.mono, fontSize: 12, padding: "4px 7px", borderRadius: 8, border: `1px solid ${T.line}`, background: T.card }} />}
        </div>
      </FieldRow>

      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Confraternité</h4>
      <FieldRow label="Confrère repreneur"><TextInput defaultValue={r.confrereRepreneur} onCommit={(v) => patch({ confrereRepreneur: v })} placeholder="Nom du cabinet" width={200} align="left" /></FieldRow>
      <FieldRow label="Lettre de confraternité envoyée"><ToggleBtn on={!!r.lettreConfraterniteEnvoyee} onClick={() => patch({ lettreConfraterniteEnvoyee: !r.lettreConfraterniteEnvoyee })} /></FieldRow>
      <FieldRow label="Lettre de confraternité reçue"><ToggleBtn on={!!r.lettreConfraterniteRecue} onClick={() => patch({ lettreConfraterniteRecue: !r.lettreConfraterniteRecue })} /></FieldRow>

      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Situation financière</h4>
      <FieldRow label="Honoraires">
        <SelectPill value={r.honorairesSituation} options={["soldes", "restant_du"]} labels={{ soldes: "Soldés", restant_du: "Restant dû" }} onChange={(v) => patch({ honorairesSituation: v })} />
      </FieldRow>
      <FieldRow label="Dernière clôture traitée">
        <input type="date" value={r.derniereCloture || ""} onChange={(e) => patch({ derniereCloture: e.target.value })}
          style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} />
      </FieldRow>

      {!r.active && (
        <button onClick={activer} style={{ marginTop: 20, display: "flex", alignItems: "center", gap: 6, background: T.red, color: "#fff", border: "none", borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
          <AlertTriangle size={14} /> Confirmer la résiliation
        </button>
      )}

      {(r.historique || []).length > 0 && (
        <>
          <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Historique</h4>
          {r.historique.map((h, i) => (
            <div key={i} style={{ fontSize: 11.5, color: T.inkMuted, padding: "6px 0", borderBottom: `1px solid ${T.line}` }}>
              {fmtFR(h.date)} — {h.motif} · par {h.par}
            </div>
          ))}
        </>
      )}
    </div>
  );
}

/* ============================================================
   REPRISE — dossier repris à un confrère cédant. Sur le modèle
   de la Résiliation : statut, date, confrère, checklist des
   pièces reprises, historique.
   ============================================================ */
const REPRISE_PIECES = [
  "Bilan N-1", "FEC N-1", "Balance générale N-1", "Immobilisations / tableau d'amortissement",
  "Grand livre clients / fournisseurs", "Justificatifs bancaires", "Statuts à jour", "Attestation de non-opposition",
];

function RepriseTab({ client, me, meId, portefeuilleId, onUpdate }) {
  const r = client.reprise || {};
  const pieces = r.pieces || {};
  const patch = (fields) => onUpdate(client.id, { reprise: { ...r, ...fields } });
  const togglePiece = (k) => patch({ pieces: { ...pieces, [k]: !pieces[k] } });

  const activer = () => {
    const entry = { date: r.date || todayISO(), confrereCedant: r.confrereCedant, par: me };
    patch({ active: true, historique: [...(r.historique || []), entry] });
    // Statut intermédiaire : le dossier entre au cabinet mais n'est pas encore pleinement opérationnel.
    onUpdate(client.id, { statutDossier: "transfert" });
    logActivity({ clientId: client.id, portefeuilleId, type: "reprise", message: `Reprise démarrée (confrère cédant : ${r.confrereCedant || "—"})`, auteurId: meId });
  };
  const annuler = () => {
    patch({ active: false });
    onUpdate(client.id, { statutDossier: "actif" });
    logActivity({ clientId: client.id, portefeuilleId, type: "reprise", message: "Reprise annulée", auteurId: meId });
  };
  const finaliser = () => {
    onUpdate(client.id, { statutDossier: "actif" });
    logActivity({ clientId: client.id, portefeuilleId, type: "reprise", message: "Reprise finalisée (dossier actif)", auteurId: meId });
  };

  const doneCount = REPRISE_PIECES.filter((k) => pieces[k]).length;

  return (
    <div>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
        <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: 0 }}>Reprise du dossier</h4>
        <Stamped tone={r.active ? "amber" : "neutral"} small>{r.active ? "Reprise en cours" : "Aucune reprise en cours"}</Stamped>
      </div>

      {r.active && (
        <div style={{ fontSize: 11.5, color: T.amber, background: T.amberSoft, padding: "8px 12px", borderRadius: 9, marginBottom: 16 }}>
          Ce dossier est marqué en cours de reprise — statut « En transfert ».
          <button onClick={annuler} style={{ marginLeft: 10, background: "none", border: "none", color: T.navy, fontWeight: 700, cursor: "pointer", fontSize: 11.5 }}>Annuler la reprise</button>
          {doneCount === REPRISE_PIECES.length && (
            <button onClick={finaliser} style={{ marginLeft: 10, background: "none", border: "none", color: T.green, fontWeight: 700, cursor: "pointer", fontSize: 11.5 }}>Finaliser la reprise (Actif)</button>
          )}
        </div>
      )}

      <FieldRow label="Date de reprise">
        <input type="date" value={r.date || ""} onChange={(e) => patch({ date: e.target.value })}
          style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} />
      </FieldRow>
      <FieldRow label="Confrère cédant"><TextInput defaultValue={r.confrereCedant} onCommit={(v) => patch({ confrereCedant: v })} placeholder="Nom du cabinet cédant" width={200} align="left" /></FieldRow>
      <FieldRow label="Lettre de confraternité envoyée"><ToggleBtn on={!!r.lettreConfraterniteEnvoyee} onClick={() => patch({ lettreConfraterniteEnvoyee: !r.lettreConfraterniteEnvoyee })} /></FieldRow>
      <FieldRow label="Lettre de confraternité reçue"><ToggleBtn on={!!r.lettreConfraterniteRecue} onClick={() => patch({ lettreConfraterniteRecue: !r.lettreConfraterniteRecue })} /></FieldRow>

      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Suivi des pièces reprises</h4>
      <div style={{ marginBottom: 14 }}>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 6 }}>
          <span style={{ color: T.inkMuted }}>Progression</span><span style={{ fontFamily: T.mono, fontWeight: 600 }}>{doneCount}/{REPRISE_PIECES.length}</span>
        </div>
        <div style={{ height: 8, borderRadius: 4, background: T.paperDeep, overflow: "hidden" }}>
          <div style={{ width: `${(doneCount / REPRISE_PIECES.length) * 100}%`, height: "100%", background: T.navy }} />
        </div>
      </div>
      {REPRISE_PIECES.map((k) => (
        <div key={k} onClick={() => togglePiece(k)} className="clickable" style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 4px", borderBottom: `1px solid ${T.line}` }}>
          <span style={{ width: 19, height: 19, borderRadius: 5, border: `1.5px solid ${pieces[k] ? T.green : T.line}`, background: pieces[k] ? T.green : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            {pieces[k] && <Check size={13} color="#fff" strokeWidth={3} />}
          </span>
          <span style={{ fontSize: 12.5, color: pieces[k] ? T.inkMuted : T.ink, textDecoration: pieces[k] ? "line-through" : "none" }}>{k}</span>
        </div>
      ))}

      {!r.active && (
        <button onClick={activer} style={{ marginTop: 20, display: "flex", alignItems: "center", gap: 6, background: T.amber, color: "#fff", border: "none", borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
          <RefreshCw size={14} /> Démarrer le suivi de reprise
        </button>
      )}

      {(r.historique || []).length > 0 && (
        <>
          <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>Historique</h4>
          {r.historique.map((h, i) => (
            <div key={i} style={{ fontSize: 11.5, color: T.inkMuted, padding: "6px 0", borderBottom: `1px solid ${T.line}` }}>
              {fmtFR(h.date)} — confrère : {h.confrereCedant || "—"} · par {h.par}
            </div>
          ))}
        </>
      )}
    </div>
  );
}

function ReprisesView({ clients, search, roleFilter, setRoleFilter, me, meId, portefeuilleId, onUpdate }) {
  // statutFilter="tous" : sans ça, le filtre par défaut ("actif") masque les dossiers
  // dès qu'une reprise démarre et bascule le statut sur "transfert" — le dossier
  // disparaissait alors purement et simplement de cette liste. Voir aussi ResiliationsView.
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me, undefined, "tous"), [clients, search, roleFilter, me]);
  const [expanded, setExpanded] = useState(null);

  const enCours = filtered.filter((c) => c.reprise?.active);
  const autres = filtered.filter((c) => !c.reprise?.active);

  const renderRow = (c) => {
    const isOpen = expanded === c.id;
    const r = c.reprise || {};
    const pieces = r.pieces || {};
    const doneCount = REPRISE_PIECES.filter((k) => pieces[k]).length;
    const finalisee = r.active && c.statutDossier === "actif";
    const pretAFinaliser = r.active && !finalisee && doneCount === REPRISE_PIECES.length;
    return (
      <div key={c.id} style={{ borderBottom: `1px solid ${T.line}` }}>
        <div className="hoverRow clickable" onClick={() => setExpanded(isOpen ? null : c.id)}
          style={{ display: "flex", alignItems: "center", gap: 14, padding: "11px 4px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
          {finalisee && <Stamped tone="green" small>Terminée</Stamped>}
          {!finalisee && r.active && <Stamped tone={pretAFinaliser ? "green" : "amber"} small>{pretAFinaliser ? "Prêt à finaliser" : "En cours"}</Stamped>}
          {!r.active && <Stamped tone="neutral" small>—</Stamped>}
          {r.confrereCedant && <span style={{ fontSize: 11, color: T.inkMuted }}>{r.confrereCedant}</span>}
          {r.active && !finalisee && !isOpen && (
            <span style={{ fontSize: 11, fontWeight: 700, color: T.navy, display: "flex", alignItems: "center", gap: 3 }}>
              Reprendre <ArrowUpRight size={12} />
            </span>
          )}
          {isOpen ? <ChevronUp size={15} color={T.inkMuted} /> : <ChevronDown size={15} color={T.inkMuted} />}
        </div>
        {isOpen && <div style={{ padding: "0 4px 16px" }}><RepriseTab client={c} me={me} meId={meId} portefeuilleId={portefeuilleId} onUpdate={onUpdate} /></div>}
      </div>
    );
  };

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Reprises</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Dossiers repris à un confrère cédant : suivi des pièces et de la transition.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Reprises en cours (${enCours.length})`}>{enCours.length === 0 ? <EmptyNote text="Aucune reprise en cours." /> : enCours.map(renderRow)}</Panel>
      <div style={{ height: 16 }} />
      <Panel title="Tous les autres dossiers">{autres.length === 0 ? <EmptyNote text="Aucun autre dossier dans cette sélection." /> : autres.map(renderRow)}</Panel>
    </div>
  );
}

/* ============================================================
   TVA GRID VIEW
   ============================================================ */
function TvaGrid({ clients, search, roleFilter, setRoleFilter, me, onCycle, onReview, onUpdate, onOpenClient }) {
  const [collabFilter, setCollabFilter] = useState("Tous");
  const [regimeHeaderFilter, setRegimeHeaderFilter] = useState("Tous");
  const [exigHeaderFilter, setExigHeaderFilter] = useState("Tous");
  const [sortBy, setSortBy] = useState("nom"); // nom | retards
  // Cellule pour laquelle le petit menu "Contrôlé et validé / Contrôlé non validé" est ouvert
  const [reviewCell, setReviewCell] = useState(null); // { clientId, mois }
  // Panneau de saisie des éléments à modifier, ouvert quand on choisit "Contrôlé non validé"
  // (ou quand on rouvre une cellule déjà marquée "Non validé" pour consulter/modifier la remarque)
  const [correctionPanel, setCorrectionPanel] = useState(null); // { client, mois, initial }

  // La vue TVA doit rester robuste même pendant le chargement des dossiers.
  // On normalise systématiquement la source avant les filtres pour éviter qu'un
  // dossier incomplet ne fasse planter toute l'application à l'ouverture de l'onglet.
  const safeClients = Array.isArray(clients) ? clients.filter(Boolean) : [];
  const baseFiltered = useMemo(() =>
    filterClients(safeClients, search || "", roleFilter, me).filter((c) => !!c?.tvaRegime),
    [safeClients, search, roleFilter, me]
  );
  const collabOptions = useMemo(() =>
    Array.from(new Set(safeClients.map((c) => c?.collab).filter(Boolean))).sort(),
    [safeClients]
  );
  const regimeOptions = useMemo(() =>
    ["Tous", ...Array.from(new Set(baseFiltered.map((c) => c?.tvaRegime).filter(Boolean))).sort()],
    [baseFiltered]
  );
  const exigOptions = useMemo(() => {
    const values = baseFiltered
      .map((c) => c?.tvaExig)
      .filter((v) => v !== "" && v != null && Number.isFinite(Number(v)))
      .map((v) => Number(v));
    return ["Tous", ...Array.from(new Set(values)).sort((a, b) => a - b).map(String)];
  }, [baseFiltered]);
  const countRetards = (c) => MOIS_ORDER.filter((m) => effectiveTvaStatus(c, m) === "RETARD").length;
  const filtered = useMemo(() => {
    let out = collabFilter === "Tous" ? baseFiltered : baseFiltered.filter((c) => c?.collab === collabFilter);
    if (regimeHeaderFilter !== "Tous") out = out.filter((c) => c?.tvaRegime === regimeHeaderFilter);
    if (exigHeaderFilter !== "Tous") out = out.filter((c) => String(c?.tvaExig ?? "") === String(exigHeaderFilter));
    out = [...out].sort((a, b) => sortBy === "retards"
      ? countRetards(b) - countRetards(a)
      : String(a?.nom || "").localeCompare(String(b?.nom || ""))
    );
    return out;
  }, [baseFiltered, collabFilter, regimeHeaderFilter, exigHeaderFilter, sortBy]);
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 15, fontWeight: 600, color: T.ink, margin: "0 0 5px" }}>Échéances TVA</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 11, marginTop: 0, marginBottom: 16, lineHeight: 1.5 }}>
        Cliquez une cellule vide pour la passer à <Stamped tone="amber" small>Fait</Stamped> — cela notifie le chef de mission que le dossier est prêt à être contrôlé.
        {" "}Cliquez ensuite sur <Stamped tone="amber" small>Fait</Stamped> pour choisir <Stamped tone="green" small>Contrôlé et validé</Stamped> (le collaborateur peut déclarer) ou <Stamped tone="purple" small>Contrôlé non validé</Stamped> (des éléments sont à modifier avant la déclaration — un panneau s'ouvre pour préciser quoi).
        {" "}Le collaborateur est notifié dans les deux cas. Cliquez une cellule <Stamped tone="purple" small>Non validé</Stamped> pour revoir la remarque. Date limite dépassée sans saisie → <Stamped tone="red" small>Retard</Stamped> automatique.
        {" "}CA3 : déclaration du mois M exigible en M+1. CA12 : une seule déclaration, en Mai N+1.
      </p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <div className="flex items-center gap-2 flex-wrap mb-3">
        <select value={collabFilter} onChange={(e) => setCollabFilter(e.target.value)} className="input-field !py-1.5 !w-auto text-xs" title="Filtrer par collaborateur">
          <option value="Tous">Collaborateur : Tous</option>
          {collabOptions.map((c) => <option key={c} value={c}>Collaborateur : {c}</option>)}
        </select>
        <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="input-field !py-1.5 !w-auto text-xs" title="Trier">
          <option value="nom">Trier : Nom (A→Z)</option>
          <option value="retards">Trier : Nb de retards (décroissant)</option>
        </select>
      </div>
      <div className="scrollbar" style={{ overflowX: "auto", background: T.card, border: `1px solid ${T.line}`, borderRadius: 16, boxShadow: T.shadowSm }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11.5 }}>
          <thead><tr>
            <th style={thStyle}>Dossier</th>
            <th style={thStyle}>
              <select value={regimeHeaderFilter} onChange={(e) => setRegimeHeaderFilter(e.target.value)} title="Filtrer par régime TVA"
                style={{ border: "none", background: "transparent", font: "inherit", color: T.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", fontSize: 9.5, fontWeight: 600, cursor: "pointer", padding: 0 }}>
                {regimeOptions.map((r) => <option key={r} value={r}>{r === "Tous" ? "Régime" : r}</option>)}
              </select>
            </th>
            <th style={{ ...thStyle, textAlign: "center" }}>
              <select value={exigHeaderFilter} onChange={(e) => setExigHeaderFilter(e.target.value)} title="Filtrer par jour d'exigibilité"
                style={{ border: "none", background: "transparent", font: "inherit", color: T.inkMuted, textTransform: "uppercase", letterSpacing: "0.05em", fontSize: 9.5, fontWeight: 600, cursor: "pointer", padding: 0, textAlign: "center" }}>
                {exigOptions.map((d) => <option key={d} value={d}>{d === "Tous" ? "Exig." : `Exig. ${d}`}</option>)}
              </select>
            </th>
            {MOIS_ORDER.map((m) => <th key={m} style={{ ...thStyle, textAlign: "center" }}>{m}</th>)}
          </tr></thead>
          <tbody>
            // APRÈS
{filtered.map((c, rowIndex) => {
  const isCa12 = c.tvaRegime === "CA12";
  const isCa3Trim = c.tvaRegime === "CA3" && c.tvaPeriodicite === "trimestrielle";
  // Les dernières lignes du tableau n'ont pas assez de place en dessous pour
  // afficher le petit menu de contrôle : on l'ouvre alors vers le haut plutôt
  // que vers le bas, pour qu'il reste toujours entièrement visible/cliquable.
  const openUpward = rowIndex >= filtered.length - 3;
  return (
              <tr key={c.id} className="hoverRow">
                <td className={onOpenClient ? "clickable" : undefined} onClick={() => onOpenClient && onOpenClient(c.id)}
                  style={{ ...tdStyle, fontWeight: 600, whiteSpace: "nowrap", color: onOpenClient ? T.navy : T.ink }}>{c.nom}</td>
                <td style={{ ...tdStyle, fontFamily: T.mono, color: T.inkMuted }}>
                  {c.tvaRegime}{isCa3Trim && <span style={{ marginLeft: 4, fontSize: 9.5, color: T.navy, background: T.navySoft, padding: "1px 5px", borderRadius: 999 }}>Trim.</span>}
                </td>
                <td style={{ ...tdStyle, textAlign: "center" }}>
                  <input type="number" min="1" max="31" defaultValue={c.tvaExig || ""} placeholder="—"
                    onBlur={(e) => onUpdate(c.id, { tvaExig: e.target.value ? parseInt(e.target.value, 10) : "" })}
                    style={{ width: 42, textAlign: "center", fontFamily: T.mono, fontSize: 12, padding: "4px 2px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper }} />
                </td>
                {MOIS_ORDER.map((m) => {
                  if (isCa12 && m !== "Mai") {
                    return <td key={m} style={{ ...tdStyle, textAlign: "center", color: T.inkMuted, opacity: 0.45 }}>—</td>;
                  }
                  if (isCa3Trim && !QUARTER_END_MONTHS.includes(m)) {
                    return <td key={m} style={{ ...tdStyle, textAlign: "center", color: T.inkMuted, opacity: 0.45 }}>—</td>;
                  }
                  const manual = (c.tvaMois?.[m] || "").toUpperCase(); const display = effectiveTvaStatus(c, m); const tone = tvaTone(display);
                  const note = c.tvaControle?.[m]?.commentaire || "";
                  const isReviewOpen = reviewCell && reviewCell.clientId === c.id && reviewCell.mois === m;
                  return (
                    <td key={m} style={{ ...tdStyle, textAlign: "center", position: "relative" }}>
                      <button
                        className="clickable"
                        title={manual === "NON_VALIDE" && note ? note : undefined}
                        onClick={() => {
                          if (manual === "") { onCycle(c.id, m, "FAIT"); return; }
                          if (manual === "FAIT") { setReviewCell({ clientId: c.id, mois: m }); return; }
                          if (manual === "OK") { onCycle(c.id, m, "NA"); return; }
                          if (manual === "NON_VALIDE") { setCorrectionPanel({ client: c, mois: m, initial: note }); return; }
                          onCycle(c.id, m, ""); // NA -> vide
                        }}
                        style={{ background: "none", border: "none", padding: 0 }}
                      >
                        <Stamped tone={tone} small>{tvaStatusLabel(display)}</Stamped>
                      </button>
                      {isReviewOpen && (
                        <>
                          <div onClick={() => setReviewCell(null)} style={{ position: "fixed", inset: 0, zIndex: 45 }} />
                          // APRÈS
<div onClick={(e) => e.stopPropagation()} style={{
  position: "absolute", left: "50%", transform: "translateX(-50%)",
  ...(openUpward ? { bottom: "100%", marginBottom: 4 } : { top: "100%", marginTop: 4 }),
  zIndex: 46, width: 210, background: T.paper, border: `1px solid ${T.line}`, borderRadius: 10,
  boxShadow: "0 14px 32px rgba(15,23,42,0.2)", padding: 6, display: "flex", flexDirection: "column", gap: 3,
}}>
                            <div style={{ fontSize: 9.5, color: T.inkMuted, padding: "3px 6px", textAlign: "left", textTransform: "uppercase", letterSpacing: "0.04em" }}>
                              Contrôle {c.nom} — {m}
                            </div>
                            <button
                              onClick={() => { onReview(c.id, m, "OK"); setReviewCell(null); }}
                              style={{ display: "flex", alignItems: "center", gap: 7, textAlign: "left", fontSize: 12, fontWeight: 600, color: T.ink, background: "none", border: "none", borderRadius: 7, padding: "7px 8px", cursor: "pointer" }}
                              onMouseEnter={(e) => e.currentTarget.style.background = T.greenSoft}
                              onMouseLeave={(e) => e.currentTarget.style.background = "none"}
                            >
                              <CheckCircle2 size={14} color={T.green} /> Contrôlé et validé
                            </button>
                            <button
                              onClick={() => { setCorrectionPanel({ client: c, mois: m, initial: "" }); setReviewCell(null); }}
                              style={{ display: "flex", alignItems: "center", gap: 7, textAlign: "left", fontSize: 12, fontWeight: 600, color: T.ink, background: "none", border: "none", borderRadius: 7, padding: "7px 8px", cursor: "pointer" }}
                              onMouseEnter={(e) => e.currentTarget.style.background = "#EDE9FE"}
                              onMouseLeave={(e) => e.currentTarget.style.background = "none"}
                            >
                              <XCircle size={14} color="#6D28D9" /> Contrôlé non validé
                            </button>
                          </div>
                        </>
                      )}
                    </td>
                  );
                })}
              </tr>
            );})}
          </tbody>
        </table>
        {filtered.length === 0 && <EmptyNote text="Aucun dossier soumis à la TVA dans cette sélection." />}
      </div>
      {correctionPanel && (
        <TvaCorrectionPanel
          client={correctionPanel.client}
          mois={correctionPanel.mois}
          initial={correctionPanel.initial}
          onClose={() => setCorrectionPanel(null)}
          onSave={(commentaire) => { onReview(correctionPanel.client.id, correctionPanel.mois, "NON_VALIDE", commentaire); setCorrectionPanel(null); }}
          onMarkFixed={() => { onCycle(correctionPanel.client.id, correctionPanel.mois, "FAIT"); setCorrectionPanel(null); }}
        />
      )}
    </div>
  );
}

/* ============================================================
   PANNEAU DE CONTRÔLE TVA — éléments à modifier avant déclaration
   ============================================================ */
function TvaCorrectionPanel({ client, mois, initial, onClose, onSave, onMarkFixed }) {
  const [text, setText] = useState(initial || "");
  const isExisting = !!initial;
  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 60, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div onClick={onClose} style={{ position: "absolute", inset: 0, background: "rgba(28,37,65,0.4)" }} />
      <div className="scrollbar" style={{ position: "relative", background: T.paper, borderRadius: 14, padding: 24, width: 440, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
          <XCircle size={18} color="#6D28D9" />
          <h3 style={{ fontFamily: T.serif, fontSize: 15, fontWeight: 600, color: T.navy, margin: 0 }}>TVA {mois} — {client.nom}</h3>
        </div>
        <p style={{ fontSize: 11.5, color: T.inkMuted, margin: "4px 0 14px" }}>
          Contrôlé non validé : précisez ce que le collaborateur doit modifier avant de pouvoir déclarer.
        </p>
        <textarea
          autoFocus
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder="Ex. : facture n°4521 à recoder en 401, écart de 230 € sur le compte de TVA collectée…"
          rows={6}
          style={{ width: "100%", padding: 10, borderRadius: 10, border: `1px solid ${T.line}`, fontSize: 12.5, background: T.card, resize: "vertical", color: T.ink }}
        />
        <div style={{ display: "flex", gap: 8, marginTop: 18, justifyContent: "space-between", flexWrap: "wrap" }}>
          <div>
            {isExisting && (
              <button onClick={onMarkFixed} style={{ padding: "9px 12px", borderRadius: 10, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 11.5, color: T.navy, fontWeight: 600 }}>
                Corrigé → repasser en revue (Fait)
              </button>
            )}
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onClose} style={{ padding: "9px 14px", borderRadius: 10, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 12 }}>Annuler</button>
            <button onClick={() => onSave(text.trim())} style={{ padding: "9px 16px", borderRadius: 10, border: "none", background: "#6D28D9", color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
              Enregistrer — Non validé
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
const thStyle = { textAlign: "left", padding: "7px 10px", fontSize: 9.5, textTransform: "uppercase", letterSpacing: "0.05em", color: T.inkMuted, borderBottom: `1px solid ${T.line}`, position: "sticky", top: 0, background: T.card };
const tdStyle = { padding: "6px 10px", fontSize: 11.5, borderBottom: `1px solid ${T.line}` };

/* ============================================================
   BILANS VIEW
   ============================================================ */
function BilansView({ clients, search, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const late = filtered.filter(isBilanLate);
  const others = filtered.filter((c) => !isBilanLate(c) && c.bilan);
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Suivi des bilans</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Repérez les dossiers en retard et suivez le courrier de relance.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`En retard (${late.length})`}>{late.length === 0 ? <EmptyNote text="Aucun bilan en retard sur cette sélection." /> : <BilanTable clients={late} onUpdate={onUpdate} />}</Panel>
      <div style={{ height: 16 }} />
      <Panel title="Autres dossiers suivis">{others.length === 0 ? <EmptyNote text="Aucun autre dossier avec suivi de bilan renseigné." /> : <BilanTable clients={others} onUpdate={onUpdate} />}</Panel>
    </div>
  );
}
function BilanTable({ clients, onUpdate }) {
  return (
    <div>
      {clients.map((c) => {
        const b = c.bilan || {};
        return (
          <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 14, padding: "10px 4px", borderBottom: `1px solid ${T.line}`, flexWrap: "wrap" }}>
            <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
            <span style={{ fontSize: 11.5, color: T.inkMuted }}>Clôture: {fmtFR(c.dateCloture)}</span>
            <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: T.inkMuted }}><input type="checkbox" checked={!!b.finaliseApres} onChange={() => onUpdate(c.id, { bilan: { ...b, finaliseApres: !b.finaliseApres } })} /> finalisé après échéance</label>
            <Stamped tone={isBilanLate(c) ? "red" : "neutral"} small>{isBilanLate(c) ? "En retard" : "Dans les délais"}</Stamped>
            <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: T.inkMuted }}><input type="checkbox" checked={!!b.courrier} onChange={() => onUpdate(c.id, { bilan: { ...b, courrier: !b.courrier } })} /> courrier classé</label>
          </div>
        );
      })}
    </div>
  );
}

/* ============================================================
   ACOMPTES VIEW
   ============================================================ */
function AcomptesView({ clients, search, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const isConcerned = filtered.filter((c) => Number(c.is?.montantN1) > 3000);
const cfeConcerned = filtered.filter((c) => Number(c.cfe?.montantN1) > 3000);
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Acomptes IS &amp; CFE</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Dossiers dont l'impôt N-1 dépasse 3 000 €.</p>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 10 }}>
        <button onClick={() => exportAcomptesToExcel([...isConcerned, ...cfeConcerned.filter((c) => !isConcerned.includes(c))])} className="btn-secondary !py-2">
          Exporter la liste (Excel)
        </button>
      </div>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Acomptes IS (${isConcerned.length} dossiers concernés)`}>
        {isConcerned.length === 0 ? <EmptyNote text="Aucun dossier marqué concerné pour l'instant." /> : isConcerned.map((c) => (
          <AcompteRow key={c.id} client={c} fields={[["mars", "Mars"], ["juin", "Juin"], ["sept", "Sept"], ["dec", "Déc"], ["solde", "Solde"]]} field="is" onUpdate={onUpdate} />
        ))}
      </Panel>
      <div style={{ height: 16 }} />
      <Panel title={`Acomptes CFE (${cfeConcerned.length} dossiers concernés)`}>
        {cfeConcerned.length === 0 ? <EmptyNote text="Aucun dossier marqué concerné pour l'instant." /> : cfeConcerned.map((c) => (
          <AcompteRow key={c.id} client={c} fields={[["juin", "Juin"], ["dec", "Déc (solde)"]]} field="cfe" onUpdate={onUpdate} />
        ))}
      </Panel>
    </div>
  );
}
function AcompteRow({ client, fields, field, onUpdate }) {
  const obj = client[field] || {};
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 14, padding: "10px 4px", borderBottom: `1px solid ${T.line}`, flexWrap: "wrap" }}>
      <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{client.nom}</div>
      {fields.map(([k, label]) => (
        <label key={k} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: T.inkMuted }}>
          <input type="checkbox" checked={!!obj[k]} onChange={() => onUpdate(client.id, { [field]: { ...obj, [k]: !obj[k] } })} /> {label}
        </label>
      ))}
    </div>
  );
}
function ResiliationsView({ clients, search, roleFilter, setRoleFilter, me, meId, portefeuilleId, onUpdate }) {
  // statutFilter="tous" : corrige le bug où un dossier disparaissait de la liste dès le
  // démarrage de la résiliation. Dès qu'une résiliation démarrait, le dossier passait en
  // statut "transfert", mais cette vue ne regardait (via filterClients) que les dossiers
  // "actif" par défaut — le dossier disparaissait donc sans aucun moyen d'y revenir.
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me, undefined, "tous"), [clients, search, roleFilter, me]);
  const [expanded, setExpanded] = useState(null);

  const enCours = filtered.filter((c) => c.resiliation?.active);
  const autres = filtered.filter((c) => !c.resiliation?.active);

  const renderRow = (c) => {
    const isOpen = expanded === c.id;
    const r = c.resiliation || {};
    const finalisee = r.active && c.statutDossier === "inactif";
    const pretAFinaliser = r.active && !!r.piecesRestituees && c.statutDossier === "transfert";
    return (
      <div key={c.id} style={{ borderBottom: `1px solid ${T.line}` }}>
        <div className="hoverRow clickable" onClick={() => setExpanded(isOpen ? null : c.id)}
          style={{ display: "flex", alignItems: "center", gap: 14, padding: "11px 4px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
          {finalisee && <Stamped tone="neutral" small>Sortie finalisée</Stamped>}
          {!finalisee && r.active && <Stamped tone={pretAFinaliser ? "green" : "red"} small>{pretAFinaliser ? "Prêt à finaliser" : "En cours"}</Stamped>}
          {!r.active && <Stamped tone="neutral" small>Dossier actif</Stamped>}
          {r.motif && <span style={{ fontSize: 11, color: T.inkMuted }}>{r.motif}</span>}
          {r.active && !finalisee && !isOpen && (
            <span style={{ fontSize: 11, fontWeight: 700, color: T.navy, display: "flex", alignItems: "center", gap: 3 }}>
              Reprendre <ArrowUpRight size={12} />
            </span>
          )}
          {isOpen ? <ChevronUp size={15} color={T.inkMuted} /> : <ChevronDown size={15} color={T.inkMuted} />}
        </div>
        {isOpen && <div style={{ padding: "0 4px 16px" }}><ResiliationTab client={c} me={me} meId={meId} portefeuilleId={portefeuilleId} onUpdate={onUpdate} /></div>}
      </div>
    );
  };

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Résiliations</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Suivi des dossiers résiliés et des dossiers en cours de sortie.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Résiliés / en cours (${enCours.length})`}>{enCours.length === 0 ? <EmptyNote text="Aucun dossier résilié pour l'instant." /> : enCours.map(renderRow)}</Panel>
      <div style={{ height: 16 }} />
      <Panel title="Tous les autres dossiers">{autres.length === 0 ? <EmptyNote text="Aucun autre dossier dans cette sélection." /> : autres.map(renderRow)}</Panel>
    </div>
  );
}

function MissionsExceptionnellesView({ clients, search, roleFilter, setRoleFilter, me, onUpdate, team }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const [expanded, setExpanded] = useState(null);

  const avecMissions = filtered.filter((c) => (c.missionsExceptionnelles || []).length > 0);
  const sansMission = filtered.filter((c) => !(c.missionsExceptionnelles || []).length);

  const renderRow = (c) => {
    const isOpen = expanded === c.id;
    const nb = (c.missionsExceptionnelles || []).length;
    return (
      <div key={c.id} style={{ borderBottom: `1px solid ${T.line}` }}>
        <div className="hoverRow clickable" onClick={() => setExpanded(isOpen ? null : c.id)}
          style={{ display: "flex", alignItems: "center", gap: 14, padding: "11px 4px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
          <Stamped tone={nb > 0 ? "amber" : "neutral"} small>{nb} mission{nb > 1 ? "s" : ""}</Stamped>
          {isOpen ? <ChevronUp size={15} color={T.inkMuted} /> : <ChevronDown size={15} color={T.inkMuted} />}
        </div>
        {isOpen && <div style={{ padding: "0 4px 16px" }}><MissionsExceptionnellesTab client={c} team={team} onUpdate={onUpdate} /></div>}
      </div>
    );
  };

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Missions exceptionnelles</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Missions ponctuelles en dehors de la lettre de mission récurrente.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Dossiers avec mission(s) en cours (${avecMissions.length})`}>{avecMissions.length === 0 ? <EmptyNote text="Aucune mission exceptionnelle en cours." /> : avecMissions.map(renderRow)}</Panel>
      <div style={{ height: 16 }} />
      <Panel title="Tous les autres dossiers">{sansMission.length === 0 ? <EmptyNote text="Aucun autre dossier dans cette sélection." /> : sansMission.map(renderRow)}</Panel>
    </div>
  );
}

/* ============================================================
   AGE / AGO
   ============================================================ */
function AgeAgoEditor({ client, onUpdate }) {
  const history = client.ageAgoHistory || {};
  const years = Object.keys(history).sort((a, b) => b - a);
  const [newYear, setNewYear] = useState(String(new Date().getFullYear()));
  const addYear = () => { if (!newYear || history[newYear]) return; onUpdate(client.id, { ageAgoHistory: { ...history, [newYear]: { ago: false, depose: false, deposePar: "", capitauxInf: false, ageContinuite: false } } }); };
  const patchYear = (year, patch) => onUpdate(client.id, { ageAgoHistory: { ...history, [year]: { ...history[year], ...patch } } });
  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <input type="number" placeholder="Année" value={newYear} onChange={(e) => setNewYear(e.target.value)} style={{ width: 100, ...inputStyle, padding: "8px 10px" }} />
        <button onClick={addYear} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}><Plus size={14} /> Ajouter cet exercice</button>
      </div>
      {years.length === 0 ? <EmptyNote text="Aucun exercice suivi pour ce dossier." /> : years.map((year) => {
        const y = history[year];
        return (
          <div key={year} style={{ border: `1px solid ${T.line}`, borderRadius: 10, padding: "12px 14px", marginBottom: 10, background: T.card }}>
            <div style={{ fontFamily: T.serif, fontWeight: 600, fontSize: 13, color: T.navy, marginBottom: 8 }}>Exercice {year}</div>
            <FieldRow label="Assemblée tenue (AGO)"><ToggleBtn on={!!y.ago} onClick={() => patchYear(year, { ago: !y.ago })} /></FieldRow>
            <FieldRow label="Déposée au greffe"><ToggleBtn on={!!y.depose} onClick={() => patchYear(year, { depose: !y.depose })} /></FieldRow>
            <FieldRow label="Déposée par"><TextInput defaultValue={y.deposePar} onCommit={(v) => patchYear(year, { deposePar: v })} placeholder="ex. Soli" width={140} /></FieldRow>
            <FieldRow label="Capitaux propres < 1/2 capital social"><ToggleBtn on={!!y.capitauxInf} onClick={() => patchYear(year, { capitauxInf: !y.capitauxInf })} tone="red" /></FieldRow>
            <FieldRow label="AGE continuité d'exploitation requise"><ToggleBtn on={!!y.ageContinuite} onClick={() => patchYear(year, { ageContinuite: !y.ageContinuite })} tone="red" /></FieldRow>
          </div>
        );
      })}
    </div>
  );
}
function FormeJuridiqueEditor({ client, onUpdate }) {
  const history = client.formeJuridiqueHistory || {};
  const years = Object.keys(history).sort((a, b) => b - a);
  const [newYear, setNewYear] = useState(String(new Date().getFullYear()));
  const items = getFormeJuridiqueItems(client);
  const addYear = () => {
    if (!newYear || history[newYear]) return;
    const blank = Object.fromEntries(items.map((it) => [it.id, false]));
    onUpdate(client.id, { formeJuridiqueHistory: { ...history, [newYear]: { ...blank, notes: "" } } });
  };
  const patchYear = (year, patch) => onUpdate(client.id, { formeJuridiqueHistory: { ...history, [year]: { ...history[year], ...patch } } });

  if (!client.formeJuridique) {
    return <EmptyNote text="Renseignez d'abord la forme juridique dans l'onglet Infos générales." />;
  }
  if (items.length === 0) {
    return <EmptyNote text={`Aucune checklist définie pour "${client.formeJuridique}" pour le moment.`} />;
  }

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <input type="number" placeholder="Année" value={newYear} onChange={(e) => setNewYear(e.target.value)} style={{ width: 100, ...inputStyle, padding: "8px 10px" }} />
        <button onClick={addYear} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "8px 14px", fontSize: 12.5, fontWeight: 600, cursor: "pointer" }}><Plus size={14} /> Ajouter cet exercice</button>
      </div>
      {years.length === 0 ? <EmptyNote text="Aucun exercice suivi pour ce dossier." /> : years.map((year) => {
        const y = history[year] || {};
        return (
          <div key={year} style={{ border: `1px solid ${T.line}`, borderRadius: 10, padding: "12px 14px", marginBottom: 10, background: T.card }}>
            <div style={{ fontFamily: T.serif, fontWeight: 600, fontSize: 13, color: T.navy, marginBottom: 8 }}>
              Exercice {year} · <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{client.formeJuridique}</span>
            </div>
            {items.map((it) => (
              <FieldRow key={it.id} label={it.label}>
                <ToggleBtn on={!!y[it.id]} onClick={() => patchYear(year, { [it.id]: !y[it.id] })} />
              </FieldRow>
            ))}
            <FieldRow label="Notes"><TextInput defaultValue={y.notes} onCommit={(v) => patchYear(year, { notes: v })} placeholder="—" width={200} align="left" /></FieldRow>
          </div>
        );
      })}
    </div>
  );
}
function RevisionTab({ client, onUpdate, setView }) {
  const rev = client.revision || {};
  const patch = (f) => onUpdate(client.id, { revision: { ...rev, ...f } });
  const cycleMonth = (mois) => {
    const banqueMois = rev.banqueMois || {};
    patch({ banqueMois: { ...banqueMois, [mois]: bankCycle(banqueMois[mois]) } });
  };


  return (
    <div>
      <div style={{ height: 14 }} />
      <Panel title="Rapprochements bancaires">
        <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
          {MOIS_ORDER.map((m) => (
            <button key={m} className="clickable" onClick={() => cycleMonth(m)} style={{ background: "none", border: `1px solid ${T.line}`, borderRadius: 8, padding: "6px 4px", minWidth: 50, textAlign: "center" }}>
              <div style={{ fontSize: 10, color: T.inkMuted, marginBottom: 3 }}>{m}</div>
              <Stamped tone={bankTone(rev.banqueMois?.[m])} small>{bankLabel(rev.banqueMois?.[m])}</Stamped>
            </button>
          ))}
        </div>
      </Panel>

      <div style={{ height: 14 }} />

      <Panel title="OD de salaires">
        <p style={{ fontSize: 12, color: T.inkMuted, margin: "0 0 10px" }}>Le suivi mois par mois se fait depuis Social &amp; paie.</p>
        <button onClick={() => setView && setView("social")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${T.line}`, borderRadius: 9, padding: "7px 12px", fontSize: 12, color: T.navy, cursor: "pointer" }}>
          <ArrowUpRight size={13} /> Ouvrir le Suivi social (OD salaires)
        </button>
      </Panel>

      <div style={{ height: 14 }} />

      <Panel title="Révision des comptes de cotisations">
        <p style={{ fontSize: 12, color: T.inkMuted, margin: "0 0 10px" }}>
          La révision mensuelle (URSSAF, retraite, prévoyance{client.secteur === "batiment" ? ", PRO BTP, CIBTP" : ""}) se fait désormais depuis Social &amp; paie, sous forme de grille mensuelle.
        </p>
        <button onClick={() => setView && setView("cotisations")} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${T.line}`, borderRadius: 9, padding: "7px 12px", fontSize: 12, color: T.navy, cursor: "pointer" }}>
          <ArrowUpRight size={13} /> Ouvrir les Cotisations sociales
        </button>
      </Panel>


    </div>
  );
}
function AidesSecteurView({ content, canEdit, onUpdate }) {
  const [active, setActive] = useState(SECTEURS_ACTIVITE[0].id);
  const data = content[active] || { aides: [], obligations: [] };

  // Cache par secteur : { loading, error, items } — évite de re-fetcher à chaque clic d'onglet
  const [newsCache, setNewsCache] = useState({});
  useEffect(() => {
    if (newsCache[active]?.items?.length || newsCache[active]?.loading) return;
    let cancelled = false;
    setNewsCache((prev) => ({ ...prev, [active]: { loading: true, error: null, items: prev[active]?.items || [] } }));
    fetchSecteurNews(active)
      .then((items) => { if (!cancelled) setNewsCache((prev) => ({ ...prev, [active]: { loading: false, error: null, items } })); })
      .catch((err) => { if (!cancelled) setNewsCache((prev) => ({ ...prev, [active]: { loading: false, error: err.message, items: [] } })); });
    return () => { cancelled = true; };
  }, [active]); // eslint-disable-line react-hooks/exhaustive-deps
  const activeNews = newsCache[active] || { loading: false, error: null, items: [] };

  const addLine = (field) => onUpdate(active, { [field]: [...(data[field] || []), ""] });
  const editLine = (field, idx, v) => {
    const arr = [...(data[field] || [])]; arr[idx] = v; onUpdate(active, { [field]: arr });
  };
  const removeLine = (field, idx) => {
    const arr = (data[field] || []).filter((_, i) => i !== idx); onUpdate(active, { [field]: arr });
  };

  const renderList = (field, title, tone) => (
    <Panel title={title}>
      {(data[field] || []).length === 0 && !canEdit ? <EmptyNote text="Rien de renseigné pour ce secteur." /> : null}
      {(data[field] || []).map((line, idx) => (
        <div key={idx} style={{ display: "flex", alignItems: "center", gap: 8, padding: "7px 4px", borderBottom: `1px solid ${T.line}` }}>
          <CircleDot size={7} color={tone} style={{ flexShrink: 0 }} />
          {canEdit ? (
            <input defaultValue={line} onBlur={(e) => editLine(field, idx, e.target.value)}
              style={{ flex: 1, fontSize: 12.5, border: "none", background: "transparent", padding: "2px 0" }} />
          ) : (
            <span style={{ flex: 1, fontSize: 12.5, color: T.ink }}>{line}</span>
          )}
          {canEdit && (
            <button onClick={() => removeLine(field, idx)} style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted }}>
              <Trash2 size={13} />
            </button>
          )}
        </div>
      ))}
      {canEdit && (
        <button onClick={() => addLine(field)} style={{ display: "flex", alignItems: "center", gap: 6, marginTop: 10, background: "none", border: `1px dashed ${T.line}`, borderRadius: 9, padding: "7px 12px", fontSize: 12, color: T.navy, cursor: "pointer" }}>
          <Plus size={13} /> Ajouter une ligne
        </button>
      )}
    </Panel>
  );

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Actualités & Aides par secteur</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>
        Aides, dispositifs et obligations réglementaires propres à chaque secteur d'activité.
        {canEdit ? " Modifiable directement ici." : " Lecture seule — seuls Expert, Chef de mission et Admin peuvent l'éditer."}
      </p>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 18 }}>
        {SECTEURS_ACTIVITE.map((s) => (
          <button key={s.id} onClick={() => setActive(s.id)} style={{
            fontSize: 11.5, fontWeight: 700, padding: "5px 11px", borderRadius: 999, cursor: "pointer",
            border: `1px solid ${active === s.id ? s.color : T.line}`,
            background: active === s.id ? s.color + "1A" : "transparent",
            color: active === s.id ? s.color : T.inkMuted,
          }}>{s.label}</button>
        ))}
      </div>
      {renderList("aides", "Aides & dispositifs", T.navy)}
      <div style={{ height: 14 }} />
      {renderList("obligations", "Obligations réglementaires", T.gold)}
      <div style={{ height: 14 }} />
      <Panel title="Actualités en direct (Google Actualités + Service-Public.fr)">
        {activeNews.loading && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 4px", color: T.inkMuted, fontSize: 12.5 }}>
            <Loader2 size={14} className="spin" /> Chargement des actualités…
          </div>
        )}
        {!activeNews.loading && activeNews.error && (
          <EmptyNote text={`Flux indisponible pour le moment (${activeNews.error})`} />
        )}
        {!activeNews.loading && !activeNews.error && activeNews.items.length === 0 && (
          <EmptyNote text="Aucune actualité récente trouvée pour ce secteur." />
        )}
        {!activeNews.loading && activeNews.items.map((it, idx) => (
          <a key={idx} href={it.link} target="_blank" rel="noopener noreferrer"
            style={{ display: "block", padding: "8px 4px", borderBottom: `1px solid ${T.line}`, textDecoration: "none" }}>
            <div style={{ fontSize: 12.5, color: T.ink, fontWeight: 600 }}>{it.title}</div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginTop: 2 }}>
              {it.source}{it.date ? ` · ${new Date(it.date).toLocaleDateString("fr-FR")}` : ""}
            </div>
          </a>
        ))}
      </Panel>
      {content[active]?.updatedAt && (
        <div style={{ fontSize: 11, color: T.inkMuted, marginTop: 14 }}>
          Dernière mise à jour : {new Date(content[active].updatedAt).toLocaleDateString("fr-FR")}{content[active].updatedBy ? ` par ${content[active].updatedBy}` : ""}
        </div>
      )}
    </div>
  );
}
function LegalServicesView({ clients, requests, setRequests }) {
  const [form, setForm] = useState({ clientId:"", nom:"", prestation:"Création de société", notes:"", statut:"demande" });
  const add = () => { if (!form.nom.trim()) return; setRequests([{ id: `jur-${Date.now()}`, ...form, createdAt: new Date().toISOString() }, ...requests]); setForm({ clientId:"", nom:"", prestation:"Création de société", notes:"", statut:"demande" }); };
  const update = (id, patch) => setRequests(requests.map(r=>r.id===id?{...r,...patch}:r));
  const remove = id => setRequests(requests.filter(r=>r.id!==id));
  const statuses=[['demande','Demande'],['production','En cours de production'],['termine','Terminé']];
  return <div><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}><div><h2 style={{margin:0,fontFamily:T.serif,color:T.navy}}>Prestations juridiques</h2><div style={{fontSize:12,color:T.inkMuted,marginTop:5}}>Commandes adressées au pôle juridique — indépendantes ou liées à un dossier existant.</div></div></div>
    <Panel title="Nouvelle demande"><div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:10}}><input value={form.nom} onChange={e=>setForm({...form,nom:e.target.value})} placeholder="Nom du client ou du projet" style={{padding:9,border:`1px solid ${T.line}`,borderRadius:9}}/><select value={form.clientId} onChange={e=>setForm({...form,clientId:e.target.value})} style={{padding:9,border:`1px solid ${T.line}`,borderRadius:9}}><option value="">Sans dossier existant</option>{clients.map(c=><option key={c.id} value={c.id}>{c.nom}</option>)}</select><select value={form.prestation} onChange={e=>setForm({...form,prestation:e.target.value})} style={{padding:9,border:`1px solid ${T.line}`,borderRadius:9}}>{["Création de société","Modification statutaire","Modification des éléments de la société","Montage holding","Transformation","Dissolution / liquidation","Autre prestation juridique"].map(x=><option key={x}>{x}</option>)}</select><button onClick={add} style={{background:T.navy,color:"white",border:0,borderRadius:9,fontWeight:700,cursor:"pointer"}}>+ Créer la demande</button></div><textarea value={form.notes} onChange={e=>setForm({...form,notes:e.target.value})} placeholder="Précisions / pièces / consignes pour le pôle juridique…" style={{marginTop:10,width:"100%",minHeight:70,padding:9,border:`1px solid ${T.line}`,borderRadius:9}}/></Panel>
    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:14,marginTop:16}}>{statuses.map(([key,label])=><div key={key} style={{background:T.card,border:`1px solid ${T.line}`,borderRadius:14,padding:12}}><div style={{fontWeight:800,color:T.navy,marginBottom:10}}>{label} <span style={{color:T.inkMuted,fontSize:11}}>{requests.filter(r=>r.statut===key).length}</span></div>{requests.filter(r=>r.statut===key).map(r=><div key={r.id} style={{border:`1px solid ${T.line}`,borderRadius:11,padding:10,marginBottom:9,boxShadow:T.shadowSm}}><div style={{fontWeight:800,fontSize:13}}>{r.nom}</div><div style={{fontSize:11,color:T.inkMuted,margin:"3px 0 8px"}}>{r.prestation}{r.clientId?" · Dossier lié":" · Hors dossier"}</div>{r.notes&&<div style={{fontSize:11,color:T.inkSoft,marginBottom:8}}>{r.notes}</div>}<div style={{display:"flex",gap:5,alignItems:"center"}}>{statuses.map(([sk,sl])=><button key={sk} onClick={()=>update(r.id,{statut:sk})} style={{fontSize:10,padding:"4px 6px",borderRadius:7,border:`1px solid ${T.line}`,background:r.statut===sk?T.navySoft:T.card,color:r.statut===sk?T.navy:T.inkMuted,cursor:"pointer"}}>{sk==='demande'?'Demande':sk==='production'?'Production':'Terminé'}</button>)}<button onClick={()=>remove(r.id)} style={{marginLeft:"auto",border:0,background:"none",color:T.red,cursor:"pointer"}}><Trash2 size={14}/></button></div></div>)}</div>)}</div></div>;
}

function AgeAgoView({ clients, search, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const [expanded, setExpanded] = useState(null);
  const withAlert = filtered.filter((c) => Object.values(c.ageAgoHistory || {}).some((y) => y.capitauxInf || y.ageContinuite));
  const rest = filtered.filter((c) => !withAlert.includes(c));
  const renderRow = (c) => {
    const h = c.ageAgoHistory || {}; const latestYear = Object.keys(h).sort((a, b) => b - a)[0]; const latest = latestYear ? h[latestYear] : null;
    const isOpen = expanded === c.id;
    return (
      <div key={c.id} style={{ borderBottom: `1px solid ${T.line}` }}>
        <div className="hoverRow clickable" onClick={() => setExpanded(isOpen ? null : c.id)} style={{ display: "flex", alignItems: "center", gap: 14, padding: "11px 4px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
          {latest ? (
            <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
              <Stamped tone={latest.ago ? "green" : "neutral"} small>{latestYear} · {latest.ago ? "AGO tenue" : "AGO à tenir"}</Stamped>
              <Stamped tone={latest.depose ? "green" : "amber"} small>{latest.depose ? "Déposée" : "Non déposée"}</Stamped>
              {latest.capitauxInf && <Stamped tone="red" small>Capitaux &lt; 1/2</Stamped>}
              {latest.ageContinuite && <Stamped tone="red" small>AGE continuité</Stamped>}
            </div>
          ) : <Stamped tone="neutral" small>Aucun exercice suivi</Stamped>}
          {isOpen ? <ChevronUp size={15} color={T.inkMuted} /> : <ChevronDown size={15} color={T.inkMuted} />}
        </div>
        {isOpen && <div style={{ padding: "0 4px 16px" }}><AgeAgoEditor client={c} onUpdate={onUpdate} /></div>}
      </div>
    );
  };
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>AGE / AGO</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Suivi par exercice : assemblée tenue, dépôt au greffe et par qui, situations à surveiller.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Dossiers signalés (${withAlert.length})`}>{withAlert.length === 0 ? <EmptyNote text="Aucun dossier signalé pour le moment." /> : withAlert.map(renderRow)}</Panel>
      <div style={{ height: 16 }} />
      <Panel title="Tous les dossiers">{rest.length === 0 ? <EmptyNote text="Aucun autre dossier dans cette sélection." /> : rest.map(renderRow)}</Panel>
    </div>
  );
}
function RevisionView({ clients, search, roleFilter, setRoleFilter, me, onUpdate, setView }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const [expanded, setExpanded] = useState(null);

  const revisionStatus = (c) => {
    const rev = c.revision || {};
    const bankDone = MOIS_ORDER.every((m) => (rev.banqueMois?.[m] || "") !== "");
    if (bankDone) return "complete";
    if (rev.banqueMois && Object.keys(rev.banqueMois).length) return "encours";
    return "nondemarre";
  };

  const late = filtered.filter((c) => revisionStatus(c) === "nondemarre");
  const encours = filtered.filter((c) => revisionStatus(c) === "encours");
  const complete = filtered.filter((c) => revisionStatus(c) === "complete");

  const renderRow = (c) => {
    const isOpen = expanded === c.id;
    return (
      <div key={c.id} style={{ borderBottom: `1px solid ${T.line}` }}>
        <div className="hoverRow clickable" onClick={() => setExpanded(isOpen ? null : c.id)}
          style={{ display: "flex", alignItems: "center", gap: 14, padding: "11px 4px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
          <Stamped tone={revisionStatus(c) === "complete" ? "green" : revisionStatus(c) === "encours" ? "amber" : "neutral"} small>
            {revisionStatus(c) === "complete" ? "Terminée" : revisionStatus(c) === "encours" ? "En cours" : "Non démarrée"}
          </Stamped>
          {isOpen ? <ChevronUp size={15} color={T.inkMuted} /> : <ChevronDown size={15} color={T.inkMuted} />}
        </div>
        {isOpen && <div style={{ padding: "0 4px 16px" }}><RevisionTab client={c} onUpdate={onUpdate} setView={setView} /></div>}
      </div>
    );
  };

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Révision comptable</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>
        Rapprochements bancaires, dossier par dossier. Le suivi des OD de salaires et la révision des comptes de cotisations se font désormais depuis le menu « Social &amp; paie ».
      </p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Non démarrée (${late.length})`}>{late.length === 0 ? <EmptyNote text="Aucun dossier dans cette situation." /> : late.map(renderRow)}</Panel>
      <div style={{ height: 16 }} />
      <Panel title={`En cours (${encours.length})`}>{encours.length === 0 ? <EmptyNote text="Aucun dossier dans cette situation." /> : encours.map(renderRow)}</Panel>
      <div style={{ height: 16 }} />
      <Panel title={`Terminée (${complete.length})`}>{complete.length === 0 ? <EmptyNote text="Aucun dossier dans cette situation." /> : complete.map(renderRow)}</Panel>
    </div>
  );
}

/* ============================================================
   À SURVEILLER — vue portefeuille / anomalies
   ============================================================ */
function SurveillanceView({ clients, search, me, onOpenClient }) {
  const anomalies = useMemo(() => detectAllAnomalies(clients), [clients]);
  const filtered = anomalies.filter((a) => {
    const q = (search || "").trim().toLowerCase();
    return !q || a.clientNom.toLowerCase().includes(q) || a.message.toLowerCase().includes(q);
  });
  const critical = filtered.filter((a) => a.gravite === "haute");
  const others = filtered.filter((a) => a.gravite !== "haute");
  const Row = ({ a }) => <div className="hoverRow clickable" onClick={() => onOpenClient(a.clientId)} style={{ display: "flex", alignItems: "center", gap: 10, padding: "11px 5px", borderBottom: `1px solid ${T.line}` }}>
    <span style={{ width: 8, height: 8, borderRadius: 99, background: a.gravite === "haute" ? T.red : T.amber, flexShrink: 0 }} />
    <div style={{ flex: 1, minWidth: 0 }}><div style={{ fontWeight: 700, fontSize: 12 }}>{a.clientNom}</div><div style={{ color: T.inkMuted, fontSize: 11 }}>{a.message}</div></div>
    <Stamped tone={a.gravite === "haute" ? "red" : "amber"} small>{a.gravite === "haute" ? "Important" : "À vérifier"}</Stamped>
  </div>;
  return <div>
    <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>À surveiller</h1></Reveal>
    <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Les contrôles automatiques qui méritent l’attention du chef de mission.</p>
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3" style={{ marginBottom: 18 }}>
      <KpiCard label="Anomalies" value={filtered.length} icon={ShieldAlert} tone={filtered.length ? "amber" : "green"} />
      <KpiCard label="Importantes" value={critical.length} icon={AlertTriangle} tone={critical.length ? "red" : "green"} />
      <KpiCard label="À vérifier" value={others.length} icon={Search} tone={others.length ? "amber" : "green"} />
      <KpiCard label="Dossiers concernés" value={new Set(filtered.map((a) => a.clientId)).size} icon={Users} tone="neutral" />
    </div>
    <Panel title={`Priorité haute (${critical.length})`}>{critical.length ? critical.map((a) => <Row key={a.id} a={a} />) : <EmptyNote text="Aucune anomalie importante." />}</Panel>
    <div style={{ height: 16 }} />
    <Panel title={`À vérifier (${others.length})`}>{others.length ? others.map((a) => <Row key={a.id} a={a} />) : <EmptyNote text="Aucun contrôle en attente." />}</Panel>
  </div>;
}

/* ============================================================
   MISSION VIEW
   ============================================================ */
function MissionView({ clients, search, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const incomplete = filtered.filter((c) => (missionCompletion(c)?.pct ?? 100) < 100).sort((a, b) => (missionCompletion(a).pct - missionCompletion(b).pct));
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Dossiers en accueil</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Suivi du processus d'acceptation de mission et de reprise de dossier.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />
      <Panel title={`Accueils incomplets (${incomplete.length})`}>
        {incomplete.length === 0 ? <EmptyNote text="Tous les dossiers d'accueil sont complets." /> : incomplete.map((c) => {
          const comp = missionCompletion(c);
          return (
            <div key={c.id} style={{ padding: "12px 4px", borderBottom: `1px solid ${T.line}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                <span style={{ fontWeight: 600, fontSize: 12.5 }}>{c.nom}</span><span style={{ fontFamily: T.mono, fontSize: 11.5, color: T.inkMuted }}>{comp.done}/{comp.total}</span>
              </div>
              <div style={{ height: 6, borderRadius: 3, background: T.paperDeep, overflow: "hidden", marginBottom: 8 }}><div style={{ width: `${comp.pct}%`, height: "100%", background: comp.pct === 100 ? T.green : T.navy }} /></div>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {MISSION_ALL_KEYS.map((k) => {
                  const v = !!(c.mission || {})[k];
                  return (
                    <button key={k} onClick={() => onUpdate(c.id, { mission: { ...c.mission, [k]: !v } })} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
                      <Stamped tone={v ? "green" : "neutral"} small>{k}</Stamped>
                    </button>
                  );
                })}
              </div>
            </div>
          );
        })}
      </Panel>
    </div>
  );
}

/* ============================================================
   CHANGEMENTS DE RÉGIME TVA
   ============================================================ */
function RegimeChangeView({ clients, me, search, onUpdate }) {
  const sorted = useMemo(() => [...clients].sort((a, b) => a.nom.localeCompare(b.nom)), [clients]);
  const [clientId, setClientId] = useState(sorted[0]?.id || "");
  const [nouveau, setNouveau] = useState("CA3");
  const [dateChangement, setDateChangement] = useState(todayISO());
  const [motif, setMotif] = useState("Dépassement de seuil de chiffre d'affaires");
  const [motifAutre, setMotifAutre] = useState("");

  const filteredList = useMemo(() => {
    if (!search.trim()) return sorted;
    const q = search.trim().toLowerCase();
    return sorted.filter((c) => c.nom.toLowerCase().includes(q));
  }, [sorted, search]);

  const client = clients.find((c) => c.id === clientId);

  const submit = () => {
    if (!client) return;
    const finalMotif = motif === "Autre" ? (motifAutre.trim() || "Autre") : motif;
    const entry = { date: dateChangement || todayISO(), ancien: client.tvaRegime || "—", nouveau, motif: finalMotif, par: me };
    onUpdate(client.id, { tvaRegime: nouveau, regimeHistory: [...(client.regimeHistory || []), entry] });
    setMotifAutre("");
  };

  const allHistory = useMemo(() => {
    const rows = [];
    clients.forEach((c) => (c.regimeHistory || []).forEach((h) => rows.push({ ...h, client: c.nom })));
    return rows.sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [clients]);

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Changements de régime TVA</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Enregistrez un passage réel normal / réel simplifié / franchise, avec traçabilité complète.</p>

      <Panel title="Enregistrer un changement">
        <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr", gap: 10, marginBottom: 12 }}>
          <select value={clientId} onChange={(e) => setClientId(e.target.value)} style={inputStyle}>
            {filteredList.map((c) => <option key={c.id} value={c.id}>{c.nom}</option>)}
          </select>
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: T.inkMuted, padding: "0 4px" }}>
            Régime actuel : <Stamped tone="neutral" small>{client?.tvaRegime || "—"}</Stamped>
          </div>
          <select value={nouveau} onChange={(e) => setNouveau(e.target.value)} style={inputStyle}>{REGIMES_TVA.map((r) => <option key={r} value={r}>Nouveau : {r}</option>)}</select>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 10 }}>
          <div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Date du changement</div>
            <input type="date" value={dateChangement} onChange={(e) => setDateChangement(e.target.value)} style={{ ...inputStyle, width: "100%" }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Motif</div>
            <select value={motif} onChange={(e) => setMotif(e.target.value)} style={{ ...inputStyle, width: "100%" }}>
              <option>Dépassement de seuil de chiffre d'affaires</option><option>Option du client</option>
              <option>Création / reprise de dossier</option><option>Régularisation administrative</option><option>Autre</option>
            </select>
          </div>
        </div>
        {motif === "Autre" && <input placeholder="Précisez le motif" value={motifAutre} onChange={(e) => setMotifAutre(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 14 }} />}
        <button onClick={submit} disabled={!client} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "10px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
          <RefreshCw size={14} /> Enregistrer le changement
        </button>
      </Panel>

      <div style={{ height: 16 }} />
      <Panel title={`Historique des changements (${allHistory.length})`} right={<History size={16} color={T.inkMuted} />}>
        {allHistory.length === 0 ? <EmptyNote text="Aucun changement de régime enregistré pour l'instant." /> : (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1.6fr 0.9fr 0.9fr 1.6fr 0.9fr 0.9fr", padding: "6px 4px", fontSize: 10.5, textTransform: "uppercase", letterSpacing: "0.05em", color: T.inkMuted, fontWeight: 600, borderBottom: `1px solid ${T.line}` }}>
              <div>Dossier</div><div>Ancien</div><div>Nouveau</div><div>Motif</div><div>Date</div><div>Par</div>
            </div>
            {allHistory.map((h, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "1.6fr 0.9fr 0.9fr 1.6fr 0.9fr 0.9fr", padding: "9px 4px", fontSize: 12.5, borderBottom: `1px solid ${T.line}`, alignItems: "center" }}>
                <div style={{ fontWeight: 600 }}>{h.client}</div>
                <div style={{ fontFamily: T.mono, color: T.inkMuted }}>{h.ancien}</div>
                <div style={{ fontFamily: T.mono, color: T.green, fontWeight: 600 }}>{h.nouveau}</div>
                <div style={{ color: T.inkSoft }}>{h.motif}</div>
                <div style={{ fontFamily: T.mono, fontSize: 11.5, color: T.inkMuted }}>{fmtFR(h.date)}</div>
                <div style={{ fontSize: 12 }}>{h.par}</div>
              </div>
            ))}
          </div>
        )}
      </Panel>
    </div>
  );
}
/* ============================================================
   HONORAIRES — montant courant + historique des changements,
   avec rappel "lettre de mission signée" et lien SharePoint.
   ============================================================ */
function HonorairesView({ clients, search, roleFilter, setRoleFilter, me, meId, portefeuilleId, onUpdate }) {
  const sorted = useMemo(() => [...clients].sort((a, b) => a.nom.localeCompare(b.nom)), [clients]);
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const [clientId, setClientId] = useState(sorted[0]?.id || "");
  const [nouveauMontant, setNouveauMontant] = useState("");
  const [dateChangement, setDateChangement] = useState(todayISO());
  const [motif, setMotif] = useState("Revalorisation annuelle");
  const [motifAutre, setMotifAutre] = useState("");
  const [lettreSignee, setLettreSignee] = useState(false);
  const [sharepointUrl, setSharepointUrl] = useState("");

  const client = clients.find((c) => c.id === clientId);

  const submit = () => {
    if (!client || !nouveauMontant.trim()) return;
    const finalMotif = motif === "Autre" ? (motifAutre.trim() || "Autre") : motif;
    const entry = {
      date: dateChangement || todayISO(), ancien: client.honoraires?.montant || "—", nouveau: nouveauMontant.trim(),
      motif: finalMotif, lettreSignee, sharepointUrl: lettreSignee ? sharepointUrl.trim() : "", par: me,
    };
    onUpdate(client.id, { honoraires: { montant: nouveauMontant.trim(), historique: [...(client.honoraires?.historique || []), entry] } });
    logActivity({ clientId: client.id, portefeuilleId, type: "honoraires", message: `Honoraires : ${entry.ancien} → ${entry.nouveau} (${finalMotif})`, auteurId: meId });
    setNouveauMontant(""); setMotifAutre(""); setLettreSignee(false); setSharepointUrl("");
  };

  const allHistory = useMemo(() => {
    const rows = [];
    clients.forEach((c) => (c.honoraires?.historique || []).forEach((h) => rows.push({ ...h, client: c.nom })));
    return rows.sort((a, b) => (a.date < b.date ? 1 : -1));
  }, [clients]);

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Honoraires</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Montant courant par dossier et historique des changements. Un changement d'honoraires implique de vérifier la lettre de mission.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />

      <Panel title="Enregistrer un changement">
        <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr", gap: 10, marginBottom: 10 }}>
          <select value={clientId} onChange={(e) => setClientId(e.target.value)} style={inputStyle}>
            {sorted.map((c) => <option key={c.id} value={c.id}>{c.nom}</option>)}
          </select>
          <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: T.inkMuted, padding: "0 4px" }}>
            Montant actuel : <strong style={{ color: T.ink }}>{client?.honoraires?.montant || "—"}</strong>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
          <input placeholder="Nouveau montant (ex. 2 400 € HT/an)" value={nouveauMontant} onChange={(e) => setNouveauMontant(e.target.value)} style={inputStyle} />
          <input type="date" value={dateChangement} onChange={(e) => setDateChangement(e.target.value)} style={inputStyle} />
          <select value={motif} onChange={(e) => setMotif(e.target.value)} style={inputStyle}>
            <option>Revalorisation annuelle</option><option>Extension de mission</option>
            <option>Renégociation à la baisse</option><option>Nouveau dossier</option><option>Autre</option>
          </select>
        </div>
        {motif === "Autre" && <input placeholder="Précisez le motif" value={motifAutre} onChange={(e) => setMotifAutre(e.target.value)} style={{ ...inputStyle, width: "100%", marginBottom: 10 }} />}

        <div style={{ background: T.paper, borderRadius: 10, padding: "10px 12px", marginBottom: 12 }}>
          <FieldRow label={<span style={{ display: "flex", alignItems: "center", gap: 6 }}><Stamp size={14} /> Lettre de mission signée reçue ?</span>}>
            <ToggleBtn on={lettreSignee} onClick={() => setLettreSignee(!lettreSignee)} />
          </FieldRow>
          {lettreSignee && (
            <div style={{ marginTop: 8 }}>
              <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Lien du document déposé sur SharePoint</div>
              <input placeholder="https://…sharepoint.com/…" value={sharepointUrl} onChange={(e) => setSharepointUrl(e.target.value)} style={{ ...inputStyle, width: "100%" }} />
            </div>
          )}
        </div>

        <button onClick={submit} disabled={!client || !nouveauMontant.trim()} style={{
          display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10,
          padding: "10px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer", opacity: (!client || !nouveauMontant.trim()) ? 0.6 : 1,
        }}>
          <RefreshCw size={14} /> Enregistrer le changement
        </button>
      </Panel>

      <div style={{ height: 16 }} />
      <Panel title={`Historique des changements (${allHistory.length})`} right={<History size={16} color={T.inkMuted} />}>
        {allHistory.length === 0 ? <EmptyNote text="Aucun changement d'honoraires enregistré pour l'instant." /> : (
          <div>
            <div style={{ display: "grid", gridTemplateColumns: "1.4fr 0.9fr 0.9fr 1.2fr 1fr 0.8fr 0.8fr", padding: "6px 4px", fontSize: 10, textTransform: "uppercase", letterSpacing: "0.05em", color: T.inkMuted, fontWeight: 600, borderBottom: `1px solid ${T.line}` }}>
              <div>Dossier</div><div>Ancien</div><div>Nouveau</div><div>Motif</div><div>Date</div><div>Lettre</div><div>Par</div>
            </div>
            {allHistory.map((h, i) => (
              <div key={i} style={{ display: "grid", gridTemplateColumns: "1.4fr 0.9fr 0.9fr 1.2fr 1fr 0.8fr 0.8fr", padding: "9px 4px", fontSize: 12, borderBottom: `1px solid ${T.line}`, alignItems: "center" }}>
                <div style={{ fontWeight: 600 }}>{h.client}</div>
                <div style={{ color: T.inkMuted, fontSize: 11.5 }}>{h.ancien}</div>
                <div style={{ color: T.green, fontWeight: 600, fontSize: 11.5 }}>{h.nouveau}</div>
                <div style={{ color: T.inkSoft, fontSize: 11.5 }}>{h.motif}</div>
                <div style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{fmtFR(h.date)}</div>
                <div>
                  {h.lettreSignee
                    ? (h.sharepointUrl ? <a href={h.sharepointUrl} target="_blank" rel="noreferrer"><Stamped tone="green" small>Signée ↗</Stamped></a> : <Stamped tone="green" small>Signée</Stamped>)
                    : <Stamped tone="amber" small>À signer</Stamped>}
                </div>
                <div style={{ fontSize: 11.5 }}>{h.par}</div>
              </div>
            ))}
          </div>
        )}
      </Panel>
    </div>
  );
}

/* ============================================================
   CADRE SOCIAL — le cabinet ne fait pas la paie, mais suit la
   réception et la comptabilisation des OD de salaire mois par mois.
   ============================================================ */
function odCycle(val) {
  const v = (val || "").toUpperCase();
  return v === "" ? "RECU" : v === "RECU" ? "COMPTA" : v === "COMPTA" ? "NA" : "";
}
function odTone(val) {
  const v = (val || "").toUpperCase();
  return v === "COMPTA" ? "green" : v === "RECU" ? "amber" : v === "NA" ? "neutral" : "neutral";
}
function odLabel(val) {
  const v = (val || "").toUpperCase();
  return v === "COMPTA" ? "Compta" : v === "RECU" ? "Reçu" : v === "NA" ? "N/A" : "·";
}
function bankCycle(val) {
  const v = (val || "").toUpperCase();
  return v === "" ? "FAIT" : v === "FAIT" ? "NA" : "";
}
function bankTone(val) {
  const v = (val || "").toUpperCase();
  return v === "FAIT" ? "green" : v === "NA" ? "neutral" : "neutral";
}
function bankLabel(val) {
  const v = (val || "").toUpperCase();
  return v === "FAIT" ? "Fait" : v === "NA" ? "N/A" : "·";
}
function ConcerneToggle({ on, onChange, small }) {
  const size = small ? { fontSize: 9.5, padding: "2px 8px" } : { fontSize: 11, padding: "4px 11px" };
  return (
    <div style={{ display: "inline-flex", borderRadius: 999, border: `1px solid ${T.line}`, overflow: "hidden" }}>
      <button onClick={() => onChange(true)} style={{ ...size, fontWeight: 700, border: "none", cursor: "pointer", background: on ? T.navy : "transparent", color: on ? "#fff" : T.inkMuted }}>Concerné</button>
      <button onClick={() => onChange(false)} style={{ ...size, fontWeight: 700, border: "none", cursor: "pointer", background: !on ? T.paperDeep : "transparent", color: !on ? T.inkSoft : T.inkMuted }}>Non concerné</button>
    </div>
  );
}
/* ============================================================
   GESTIONNAIRE DE PAIE — coordonnées du gestionnaire externe
   par dossier (nom, adresse, téléphone, e-mail).
   ============================================================ */
function GestionnairePaieView({ clients, search, setSearch, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const concernes = filtered.filter((c) => c.social?.concerne);
  const autres = filtered.filter((c) => !c.social?.concerne);
  const patchSocial = (c, patch) => onUpdate(c.id, { social: { ...(c.social || {}), ...patch } });
  const cellInput = { fontSize: 11.5, padding: "3px 6px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper };
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Gestionnaire de paie</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18, lineHeight: 1.6 }}>
        Coordonnées du gestionnaire de paie externe (cabinet de paie), dossier par dossier — utile pour contacter directement l'interlocuteur en cas de question sur les bulletins ou les OD de salaires.
      </p>
      <FilterBar search={search} setSearch={setSearch} roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={concernes.length} />
      <Panel title={`Dossiers concernés (${concernes.length})`}>
        {concernes.length === 0 ? <EmptyNote text="Aucun dossier marqué « concerné » pour l'instant — depuis Cadre social ou Cotisations sociales." /> : (
          <div className="scrollbar" style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11.5 }}>
              <thead>
                <tr>
                  <th style={thStyle}>Dossier</th><th style={thStyle}>Cabinet de paie</th><th style={thStyle}>Gestionnaire</th>
                  <th style={thStyle}>Adresse</th><th style={thStyle}>Téléphone</th><th style={thStyle}>E-mail</th>
                </tr>
              </thead>
              <tbody>
                {concernes.map((c) => (
                  <tr key={c.id} className="hoverRow">
                    <td style={{ ...tdStyle, fontWeight: 600, whiteSpace: "nowrap" }}>{c.nom}</td>
                    <td style={tdStyle}>
                      <input defaultValue={c.social?.cabinetPaie || ""} placeholder="ex. Silae, ADP…" onBlur={(e) => patchSocial(c, { cabinetPaie: e.target.value })} style={{ ...cellInput, width: 110 }} />
                    </td>
                    <td style={tdStyle}>
                      <input defaultValue={c.social?.gestionnaireNom || ""} placeholder="Nom du gestionnaire" onBlur={(e) => patchSocial(c, { gestionnaireNom: e.target.value })} style={{ ...cellInput, width: 130 }} />
                    </td>
                    <td style={tdStyle}>
                      <input defaultValue={c.social?.gestionnaireAdresse || ""} placeholder="Adresse" onBlur={(e) => patchSocial(c, { gestionnaireAdresse: e.target.value })} style={{ ...cellInput, width: 170 }} />
                    </td>
                    <td style={tdStyle}>
                      <input defaultValue={c.social?.gestionnaireTel || ""} placeholder="Téléphone" onBlur={(e) => patchSocial(c, { gestionnaireTel: e.target.value })} style={{ ...cellInput, width: 110 }} />
                    </td>
                    <td style={tdStyle}>
                      <input defaultValue={c.social?.gestionnaireEmail || ""} placeholder="E-mail" onBlur={(e) => patchSocial(c, { gestionnaireEmail: e.target.value })} style={{ ...cellInput, width: 170 }} />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
      <div style={{ height: 16 }} />
      <Panel title={`Autres dossiers (${autres.length})`}>
        {autres.length === 0 ? <EmptyNote text="Tous les dossiers de cette sélection sont marqués concernés." /> : autres.map((c) => (
          <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "9px 4px", borderBottom: `1px solid ${T.line}` }}>
            <span style={{ flex: 1, fontWeight: 600, fontSize: 12.5 }}>{c.nom}</span>
            <ConcerneToggle on={!!c.social?.concerne} onChange={(v) => patchSocial(c, { concerne: v })} />
          </div>
        ))}
      </Panel>
    </div>
  );
}

/* ============================================================
   COTISATIONS SOCIALES — grille mensuelle cliquable (comme les
   rapprochements bancaires) pour URSSAF / retraite / prévoyance
   (+ PRO BTP / CIBTP si secteur BTP). Historique conservé mois
   par mois dans client.revision.cotisMois.
   ============================================================ */
const COTISATION_TYPES = [
  { key: "urssaf", label: "URSSAF" },
  { key: "retraite", label: "Caisse de retraite" },
  { key: "prevoyance", label: "Prévoyance" },
];
const COTISATION_TYPES_BTP = [
  { key: "proBtp", label: "PRO BTP" },
  { key: "ciBtp", label: "CIBTP (caisse congés payés)" },
];
function cotisationTypesFor(client) {
  return client.secteur === "batiment" ? [...COTISATION_TYPES, ...COTISATION_TYPES_BTP] : COTISATION_TYPES;
}
function CotisationMonthlyGrid({ client, onUpdate }) {
  const rev = client.revision || {};
  const cotisMois = rev.cotisMois || {};
  const types = cotisationTypesFor(client);
  const cycleCell = (typeKey, mois) => {
    const monthsObj = cotisMois[typeKey] || {};
    onUpdate(client.id, { revision: { ...rev, cotisMois: { ...cotisMois, [typeKey]: { ...monthsObj, [mois]: bankCycle(monthsObj[mois]) } } } });
  };
  return (
    <div>
      {types.map((t) => (
        <div key={t.key} style={{ marginBottom: 14 }}>
          <div style={{ fontSize: 11.5, fontWeight: 700, color: T.inkSoft, marginBottom: 6 }}>{t.label}</div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {MOIS_ORDER.map((m) => (
              <button key={m} className="clickable" onClick={() => cycleCell(t.key, m)} style={{ background: "none", border: `1px solid ${T.line}`, borderRadius: 8, padding: "6px 4px", minWidth: 50, textAlign: "center" }}>
                <div style={{ fontSize: 10, color: T.inkMuted, marginBottom: 3 }}>{m}</div>
                <Stamped tone={bankTone(cotisMois[t.key]?.[m])} small>{bankLabel(cotisMois[t.key]?.[m])}</Stamped>
              </button>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
function CotisationsSocialesView({ clients, search, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const concernes = filtered.filter((c) => c.social?.concerne);
  const [expanded, setExpanded] = useState(null);
  const completion = (c) => {
    const rev = c.revision || {}; const cotisMois = rev.cotisMois || {};
    const types = cotisationTypesFor(c); const m = currentMonthKey();
    const done = types.filter((t) => (cotisMois[t.key]?.[m] || "") !== "").length;
    return { done, total: types.length };
  };
  const renderRow = (c) => {
    const isOpen = expanded === c.id;
    const comp = completion(c);
    return (
      <div key={c.id} style={{ borderBottom: `1px solid ${T.line}` }}>
        <div className="hoverRow clickable" onClick={() => setExpanded(isOpen ? null : c.id)}
          style={{ display: "flex", alignItems: "center", gap: 14, padding: "11px 4px", flexWrap: "wrap" }}>
          <div style={{ flex: 1, fontWeight: 600, fontSize: 12.5, minWidth: 140 }}>{c.nom}</div>
          <Stamped tone={comp.done === comp.total ? "green" : comp.done > 0 ? "amber" : "neutral"} small>{comp.done}/{comp.total} ce mois-ci</Stamped>
          {isOpen ? <ChevronUp size={15} color={T.inkMuted} /> : <ChevronDown size={15} color={T.inkMuted} />}
        </div>
        {isOpen && <div style={{ padding: "0 4px 16px" }}><CotisationMonthlyGrid client={c} onUpdate={onUpdate} /></div>}
      </div>
    );
  };
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Cotisations sociales</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18, lineHeight: 1.6 }}>
        Révision des comptes de cotisations — URSSAF, caisse de retraite, prévoyance (et PRO BTP / CIBTP pour les dossiers du bâtiment) — mois par mois, dossier par dossier.
        {" "}Cliquez une cellule : vide → <Stamped tone="green" small>Fait</Stamped> → <Stamped tone="neutral" small>N/A</Stamped>. L'historique est conservé mois par mois.
      </p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={concernes.length} />
      <Panel title={`Dossiers concernés (${concernes.length})`}>
        {concernes.length === 0 ? <EmptyNote text="Aucun dossier marqué « concerné par le social » pour l'instant." /> : concernes.map(renderRow)}
      </Panel>
    </div>
  );
}
function CadreSocialView({ clients, search, setSearch, roleFilter, setRoleFilter, me, onUpdate }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me), [clients, search, roleFilter, me]);
  const concernes = filtered.filter((c) => c.social?.concerne);
  const autres = filtered.filter((c) => !c.social?.concerne);

  const patchSocial = (c, patch) => onUpdate(c.id, { social: { ...(c.social || {}), ...patch } });
  const cycleMonth = (c, mois) => {
    const odMois = c.social?.odMois || {};
    patchSocial(c, { odMois: { ...odMois, [mois]: odCycle(odMois[mois]) } });
  };

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Cadre social</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18, lineHeight: 1.6 }}>
        Le cabinet n'établit pas les bulletins de paie. Ce suivi concerne uniquement la réception et la comptabilisation des OD de salaire transmises par le cabinet de paie externe.
        {" "}Cliquez une cellule : vide → <Stamped tone="amber" small>Reçu</Stamped> → <Stamped tone="green" small>Compta</Stamped> → <Stamped tone="neutral" small>N/A</Stamped>.
      </p>
      <FilterBar search={search} setSearch={setSearch} roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} />

      <Panel title={`Dossiers concernés (${concernes.length})`}>
        {concernes.length === 0 ? <EmptyNote text="Aucun dossier marqué « concerné » pour l'instant." /> : (
          <div className="scrollbar" style={{ overflowX: "auto" }}>
            <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11.5 }}>
            <thead>
  <tr>
    <th style={thStyle}>Dossier</th><th style={thStyle}>Effectif</th><th style={thStyle}>Cabinet de paie</th>
    <th style={thStyle}>Convention collective</th><th style={thStyle}>Régime dirigeant</th><th style={thStyle}>Seuil</th>
    {MOIS_ORDER.map((m) => <th key={m} style={{ ...thStyle, textAlign: "center" }}>{m}</th>)}
  </tr>
</thead>
<tbody>
  {concernes.map((c) => (
    <tr key={c.id} className="hoverRow">
      <td style={{ ...tdStyle, fontWeight: 600, whiteSpace: "nowrap" }}>{c.nom}</td>
      <td style={tdStyle}>
        <input defaultValue={c.social?.effectif || ""} placeholder="—" onBlur={(e) => patchSocial(c, { effectif: e.target.value })}
          style={{ width: 44, textAlign: "center", fontFamily: T.mono, fontSize: 11.5, padding: "3px 2px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper }} />
      </td>
      <td style={tdStyle}>
        <input defaultValue={c.social?.cabinetPaie || ""} placeholder="ex. Silae, ADP…" onBlur={(e) => patchSocial(c, { cabinetPaie: e.target.value })}
          style={{ width: 120, fontSize: 11.5, padding: "3px 6px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper }} />
      </td>
      <td style={tdStyle}>
        <input defaultValue={c.social?.conventionCollective || ""} placeholder="ex. Syntec" onBlur={(e) => patchSocial(c, { conventionCollective: e.target.value })}
          style={{ width: 100, fontSize: 11.5, padding: "3px 6px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper }} />
      </td>
      <td style={tdStyle}>
        <select defaultValue={c.social?.regimeDirigeant || ""} onChange={(e) => patchSocial(c, { regimeDirigeant: e.target.value })}
          style={{ fontFamily: T.mono, fontSize: 11, padding: "3px 4px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper }}>
          <option value="">—</option><option value="assimile_salarie">Assimilé salarié</option><option value="tns">TNS</option>
        </select>
      </td>
      <td style={tdStyle}>
        {(() => { const s = seuilEffectifAlert(c.social?.effectif); return s ? <Stamped tone={s.tone} small>{s.label}</Stamped> : <span style={{ color: T.inkMuted }}>—</span>; })()}
      </td>
      {MOIS_ORDER.map((m) => (
        <td key={m} style={{ ...tdStyle, textAlign: "center" }}>
          <button className="clickable" onClick={() => cycleMonth(c, m)} style={{ background: "none", border: "none", padding: 0 }}>
            <Stamped tone={odTone(c.social?.odMois?.[m])} small>{odLabel(c.social?.odMois?.[m])}</Stamped>
          </button>
        </td>
      ))}
    </tr>
  ))}
</tbody>
            </table>
          </div>
        )}
      </Panel>

      <div style={{ height: 16 }} />
      <Panel title={`Autres dossiers (${autres.length})`}>
        {autres.length === 0 ? <EmptyNote text="Tous les dossiers de cette sélection sont marqués concernés." /> : autres.map((c) => (
          <div key={c.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "9px 4px", borderBottom: `1px solid ${T.line}` }}>
            <span style={{ flex: 1, fontWeight: 600, fontSize: 12.5 }}>{c.nom}</span>
            <ConcerneToggle on={!!c.social?.concerne} onChange={(v) => patchSocial(c, { concerne: v })} />
          </div>
        ))}
      </Panel>
    </div>
  );
}
/* ============================================================
   SUIVI FISCAL — calendrier / agenda
   ============================================================ */
function SuiviFiscalView({ clients, team }) {
  const [monthOffset, setMonthOffset] = useState(0);
  const [clientFilter, setClientFilter] = useState("Tous");
  const [collabFilter, setCollabFilter] = useState("Tous");
  const [categoryFilter, setCategoryFilter] = useState("Toutes");

  const now = new Date();
  const viewDate = new Date(now.getFullYear(), now.getMonth() + monthOffset, 1);
  const monthLabel = viewDate.toLocaleDateString("fr-FR", { month: "long", year: "numeric" });

  const allEvents = useMemo(() => computeFiscalEvents(clients), [clients]);
  const collabOptions = useMemo(() => Array.from(new Set(clients.map((c) => c.collab).filter(Boolean))), [clients]);
  const categories = ["TVA", "IS", "CFE", "Bilan", "Clôture", "AGO"];

  const monthEvents = allEvents.filter((e) =>
    e.date.getFullYear() === viewDate.getFullYear() && e.date.getMonth() === viewDate.getMonth() &&
    (clientFilter === "Tous" || e.client.nom === clientFilter) &&
    (collabFilter === "Tous" || e.client.collab === collabFilter) &&
    (categoryFilter === "Toutes" || e.category === categoryFilter)
  ).sort((a, b) => a.date - b.date);

  const byDay = {};
  monthEvents.forEach((e) => { const d = e.date.getDate(); byDay[d] = byDay[d] || []; byDay[d].push(e); });

  const firstWeekday = (new Date(viewDate.getFullYear(), viewDate.getMonth(), 1).getDay() + 6) % 7;
  const daysInMonth = new Date(viewDate.getFullYear(), viewDate.getMonth() + 1, 0).getDate();
  const cells = [...Array(firstWeekday).fill(null), ...Array.from({ length: daysInMonth }, (_, i) => i + 1)];

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Suivi fiscal</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Calendrier de l'ensemble des échéances fiscales de mes dossiers.</p>

      <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <select value={clientFilter} onChange={(e) => setClientFilter(e.target.value)} style={inputStyle}>
          <option value="Tous">Tous les clients</option>
          {clients.map((c) => <option key={c.id} value={c.nom}>{c.nom}</option>)}
        </select>
        <select value={collabFilter} onChange={(e) => setCollabFilter(e.target.value)} style={inputStyle}>
          <option value="Tous">Tous les collaborateurs</option>
          {collabOptions.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>
        <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
          {["Toutes", ...categories].map((cat) => (
            <button key={cat} onClick={() => setCategoryFilter(cat)} style={{
              padding: "6px 11px", borderRadius: 20, fontSize: 11.5, fontWeight: 600,
              border: `1px solid ${categoryFilter === cat ? T.navy : T.line}`, background: categoryFilter === cat ? T.navy : T.card,
              color: categoryFilter === cat ? "#fff" : T.inkSoft, cursor: "pointer",
            }}>{cat}</button>
          ))}
        </div>
        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 8 }}>
          <button onClick={() => setMonthOffset((m) => m - 1)} style={navBtnStyle}><ChevronLeft size={15} /></button>
          <span style={{ fontFamily: T.serif, fontWeight: 600, fontSize: 13, color: T.navy, textTransform: "capitalize", minWidth: 140, textAlign: "center" }}>{monthLabel}</span>
          <button onClick={() => setMonthOffset((m) => m + 1)} style={navBtnStyle}><ChevronRight size={15} /></button>
          {monthOffset !== 0 && <button onClick={() => setMonthOffset(0)} style={{ ...navBtnStyle, fontSize: 11, width: "auto", padding: "0 10px" }}>Aujourd'hui</button>}
        </div>
      </div>

      <div style={{ background: T.card, border: `1px solid ${T.line}`, borderRadius: 16, boxShadow: T.shadowSm, padding: 14, marginBottom: 16 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4, marginBottom: 4 }}>
          {["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"].map((d) => <div key={d} style={{ fontSize: 10.5, color: T.inkMuted, textAlign: "center", fontWeight: 600 }}>{d}</div>)}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 4 }}>
          {cells.map((day, i) => {
            if (!day) return <div key={i} />;
            const evts = byDay[day] || [];
            const isToday = sameDay(new Date(viewDate.getFullYear(), viewDate.getMonth(), day), now);
            return (
              <div key={i} style={{
                minHeight: 66, border: `1px solid ${isToday ? T.navy : T.line}`, borderRadius: 9, padding: 5,
                background: isToday ? T.amberSoft : T.paper,
              }}>
                <div style={{ fontFamily: T.mono, fontSize: 10.5, color: isToday ? T.amber : T.inkMuted, fontWeight: isToday ? 700 : 500, marginBottom: 3 }}>{day}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                  {evts.slice(0, 3).map((e) => <div key={e.id} title={`${e.client.nom} — ${e.label}`} style={{ fontSize: 9.5, padding: "1px 4px", borderRadius: 3, background: T.card, color: T.inkSoft, border: `1px solid ${T.line}`, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{e.client.nom}</div>)}
                  {evts.length > 3 && <div style={{ fontSize: 9, color: T.inkMuted }}>+{evts.length - 3}</div>}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <Panel title={`Échéances du mois (${monthEvents.length})`}>
        {monthEvents.length === 0 ? <EmptyNote text="Aucune échéance sur cette période avec ces filtres." /> : (
          <div>
            {monthEvents.map((e) => (
              <div key={e.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 4px", borderBottom: `1px solid ${T.line}` }}>
                <span style={{ fontFamily: T.mono, fontSize: 12, color: T.inkMuted, width: 34 }}>{String(e.date.getDate()).padStart(2, "0")}</span>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600, fontSize: 12 }}>{e.client.nom}</div>
                  <div style={{ fontSize: 11.5, color: T.inkMuted }}>{e.label} {e.client.collab ? `· ${e.client.collab}` : ""}</div>
                </div>
                <Stamped tone={e.tone} small>{e.category}</Stamped>
              </div>
            ))}
          </div>
        )}
      </Panel>
    </div>
  );
}
const navBtnStyle = { width: 30, height: 30, borderRadius: 10, border: `1px solid ${T.line}`, background: T.card, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: T.inkSoft };


/* ============================================================
   PLANNING (personnel)
   ============================================================ */
/* ============================================================
   MES TÂCHES — page dédiée au système de tâches (table "tasks")
   Aujourd'hui / En retard / Cette semaine / À venir, avec filtres
   collaborateur / client / statut / priorité.
   ============================================================ */
const TASK_STATUT_TONE = { a_faire: "neutral", en_cours: "purple", termine: "green", bloque: "red" };
const TASK_PRIORITE_TONE = { faible: "neutral", normale: "neutral", haute: "amber", urgente: "red" };

function TasksPage({ tasks, clients, team, me, myRow, onCreate, onUpdate, onComplete, onDelete, onOpenClient }) {
  const [filterResponsable, setFilterResponsable] = useState("Tous");
  const [filterClient, setFilterClient] = useState("Tous");
  const [filterStatut, setFilterStatut] = useState("Toutes");
  const [filterPriorite, setFilterPriorite] = useState("Toutes");
  const [showForm, setShowForm] = useState(false);

  const clientById = useMemo(() => Object.fromEntries(clients.map((c) => [c.id, c])), [clients]);
  const memberById = useMemo(() => Object.fromEntries((team || []).map((t) => [t.id, t])), [team]);

  const filtered = useMemo(() => {
    return (tasks || []).filter((t) => {
      if (filterResponsable !== "Tous" && t.responsable_id !== filterResponsable) return false;
      if (filterClient !== "Tous" && t.client_id !== filterClient) return false;
      if (filterStatut !== "Toutes" && t.statut !== filterStatut) return false;
      if (filterPriorite !== "Toutes" && t.priorite !== filterPriorite) return false;
      return true;
    });
  }, [tasks, filterResponsable, filterClient, filterStatut, filterPriorite]);

  const buckets = useMemo(() => {
    const getDate = (t) => {
      if (!t.date_echeance) return null;
      const [y, m, d] = t.date_echeance.split("-").map(Number);
      return new Date(y, m - 1, d);
    };
    const b = bucketizeDeadlines(filtered.filter((t) => t.statut !== "termine"), getDate);
    const sansEcheance = filtered.filter((t) => t.statut !== "termine" && !t.date_echeance);
    b.avenir = [...new Set([...b.avenir, ...sansEcheance])];
    Object.keys(b).forEach((k) => { b[k] = b[k].sort((x, y) => taskSortWeight(x) - taskSortWeight(y)); });
    return b;
  }, [filtered]);

  const nbTermineesFiltrees = filtered.filter((t) => t.statut === "termine").length;
  const selectCls = "input-field !py-1.5 !w-auto text-xs md:text-[13px] font-medium cursor-pointer";

  return (
    <div>
      <Reveal>
        <div className="flex items-start justify-between mb-4">
          <div>
            <h1 className="font-bold text-[17px] text-ink m-0">Mes tâches</h1>
            <p className="text-inkmuted text-xs mt-1 mb-0">
              {filtered.length - nbTermineesFiltrees} tâche(s) active(s), {nbTermineesFiltrees} terminée(s) sur la sélection.
            </p>
          </div>
          <button onClick={() => setShowForm((v) => !v)} className="btn-primary">
            <Plus size={14} /> Nouvelle tâche
          </button>
        </div>
      </Reveal>

      {showForm && (
        <NewTaskForm clients={clients} team={team} onCancel={() => setShowForm(false)}
          onSubmit={async (payload) => { await onCreate(payload); setShowForm(false); }} />
      )}

      <div className="flex items-center gap-2 flex-wrap mb-4">
        <select value={filterResponsable} onChange={(e) => setFilterResponsable(e.target.value)} className={selectCls}>
          <option value="Tous">Collaborateur : Tous</option>
          {(team || []).map((t) => <option key={t.id} value={t.id}>Collaborateur : {t.nom}</option>)}
        </select>
        <select value={filterClient} onChange={(e) => setFilterClient(e.target.value)} className={selectCls}>
          <option value="Tous">Client : Tous</option>
          {clients.map((c) => <option key={c.id} value={c.id}>Client : {c.nom}</option>)}
        </select>
        <select value={filterStatut} onChange={(e) => setFilterStatut(e.target.value)} className={selectCls}>
          <option value="Toutes">Statut : Tous</option>
          {TASK_STATUTS.map((s) => <option key={s.code} value={s.code}>Statut : {s.label}</option>)}
        </select>
        <select value={filterPriorite} onChange={(e) => setFilterPriorite(e.target.value)} className={selectCls}>
          <option value="Toutes">Priorité : Toutes</option>
          {TASK_PRIORITES.map((p) => <option key={p.code} value={p.code}>Priorité : {p.label}</option>)}
        </select>
      </div>

      {["retard", "aujourdhui", "semaine", "avenir"].map((bucketKey) => (
        <div key={bucketKey} className="mb-5">
          <Panel title={`${DEADLINE_BUCKET_LABELS[bucketKey]} (${buckets[bucketKey]?.length || 0})`}>
            {!buckets[bucketKey]?.length ? <EmptyNote text="Rien ici." /> : (
              <div className="flex flex-col gap-2">
                {buckets[bucketKey].map((t, i) => (
                  <TaskRow key={t.id} task={t} index={i} client={clientById[t.client_id]} responsable={memberById[t.responsable_id]}
                    onOpenClient={onOpenClient} onUpdate={onUpdate} onComplete={onComplete} onDelete={onDelete} />
                ))}
              </div>
            )}
          </Panel>
        </div>
      ))}
    </div>
  );
}

function TaskRow({ task, index, client, responsable, onOpenClient, onUpdate, onComplete, onDelete }) {
  return (
    <Reveal index={index} delay={0.05}>
      <div className="flex flex-wrap sm:flex-nowrap items-center gap-2.5 sm:gap-3 px-3 sm:px-3.5 py-3 rounded-xl border border-line bg-white">
        {task.isAuto ? (
          <div className="w-5 h-5 rounded-full border-[1.5px] border-line flex items-center justify-center shrink-0" title="Échéance calculée automatiquement">
            <Clock3 size={11} className="text-inkmuted" />
          </div>
        ) : (
          <button onClick={() => onComplete(task)} title="Marquer terminé"
            className="w-5 h-5 rounded-full border-[1.5px] border-badge-green-text flex items-center justify-center shrink-0 hover:bg-badge-green-bg transition-colors">
            <Check size={12} className="text-badge-green-text" />
          </button>
        )}
        <div className="flex-1 min-w-[140px] sm:min-w-0">
          <div className={`font-semibold text-xs text-ink inline-block ${client ? "cursor-pointer hover:text-accent" : ""}`}
            onClick={() => client && onOpenClient(client.id)}>
            {client ? client.nom : "Dossier non lié"}
          </div>
          <div className="text-[11.5px] text-inkmuted">{task.nom}{task.commentaire ? ` — ${task.commentaire}` : ""}</div>
        </div>
        {!task.isAuto && (
          <button onClick={() => onDelete(task.id)} title="Supprimer" className="text-inkmuted hover:text-badge-red-text transition-colors order-2 sm:order-none">
            <Trash2 size={13} />
          </button>
        )}
        <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto pl-[30px] sm:pl-0 order-3 sm:order-none">
          {responsable && <RoleBadge role="Resp." name={responsable.nom} />}
          {task.isAuto ? (
            <Stamped tone="neutral" small>Auto</Stamped>
          ) : (
            <select value={task.statut} onChange={(e) => onUpdate(task.id, { statut: e.target.value })}
              className="input-field !w-auto !py-1 !px-2 text-[10.5px] font-bold cursor-pointer">
              {TASK_STATUTS.map((s) => <option key={s.code} value={s.code}>{s.label}</option>)}
            </select>
          )}
          <Stamped tone={TASK_PRIORITE_TONE[task.priorite]} small>{TASK_PRIORITE_BY_CODE[task.priorite]?.label}</Stamped>
          {task.date_echeance && <span className="font-mono text-[10.5px] text-inkmuted whitespace-nowrap">{fmtFR(task.date_echeance)}</span>}
        </div>
      </div>
    </Reveal>
  );
}

function NewTaskForm({ clients, team, onCancel, onSubmit }) {
  const [nom, setNom] = useState("");
  const [clientId, setClientId] = useState(clients[0]?.id || "");
  const [responsableId, setResponsableId] = useState("");
  const [priorite, setPriorite] = useState("normale");
  const [dateEcheance, setDateEcheance] = useState("");
  const [commentaire, setCommentaire] = useState("");
  const [saving, setSaving] = useState(false);

  const submit = async () => {
    if (!nom.trim()) return;
    setSaving(true);
    await onSubmit({
      nom: nom.trim(), client_id: clientId || null, responsable_id: responsableId || null,
      priorite, statut: "a_faire", date_echeance: dateEcheance || null, commentaire: commentaire.trim() || null,
    });
    setSaving(false);
  };

  return (
    <Panel title="Nouvelle tâche">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
        <FieldRow label="Nom de la tâche">
          <input value={nom} onChange={(e) => setNom(e.target.value)} placeholder="Ex : Relancer client pour justificatifs"
            className="input-field !w-56" />
        </FieldRow>
        <FieldRow label="Client">
          <select value={clientId} onChange={(e) => setClientId(e.target.value)} className="input-field !w-auto">
            {clients.map((c) => <option key={c.id} value={c.id}>{c.nom}</option>)}
          </select>
        </FieldRow>
        <FieldRow label="Responsable">
          <select value={responsableId} onChange={(e) => setResponsableId(e.target.value)} className="input-field !w-auto">
            <option value="">— Non assigné —</option>
            {(team || []).map((t) => <option key={t.id} value={t.id}>{t.nom}</option>)}
          </select>
        </FieldRow>
        <FieldRow label="Priorité">
          <select value={priorite} onChange={(e) => setPriorite(e.target.value)} className="input-field !w-auto">
            {TASK_PRIORITES.map((p) => <option key={p.code} value={p.code}>{p.label}</option>)}
          </select>
        </FieldRow>
        <FieldRow label="Échéance">
          <input type="date" value={dateEcheance} onChange={(e) => setDateEcheance(e.target.value)} className="input-field !w-auto" />
        </FieldRow>
      </div>
      <FieldRow label="Commentaire">
        <textarea value={commentaire} onChange={(e) => setCommentaire(e.target.value)} rows={2} className="input-field !w-full resize-y" />
      </FieldRow>
      <div className="flex gap-2 mt-4 justify-end">
        <button onClick={onCancel} className="btn-secondary">Annuler</button>
        <button onClick={submit} disabled={saving || !nom.trim()} className="btn-primary disabled:opacity-60">
          {saving ? "Création…" : "Créer la tâche"}
        </button>
      </div>
    </Panel>
  );
}

/* ============================================================
   PLANNING — calendrier semaine avec drag & drop réel.
   Branché sur la vraie table "tasks" (pas les échéances fiscales
   calculées) : responsable_id, client_id, date_echeance,
   heure_debut, duree_minutes, statut, priorite, commentaire.
   ============================================================ */
const PLANNING_HOURS = Array.from({ length: 11 }, (_, i) => 8 + i); // 8h -> 18h
const PLANNING_SLOT_H = 48; // hauteur en px d'un créneau d'1h

function planningBucket(task, weekStart) {
  if (!task.date_echeance) return "sans-date";
  const [y, m, d] = task.date_echeance.split("-").map(Number);
  const date = new Date(y, m - 1, d);
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const diff = Math.round((date - today) / 86400000);
  if (diff < 0) return "retard";
  if (diff === 0) return "aujourdhui";
  const weekEnd = new Date(weekStart); weekEnd.setDate(weekEnd.getDate() + 6);
  if (date >= weekStart && date <= weekEnd) return "semaine";
  return "plus-tard";
}
/* ---- Export .ics (Outlook / calendrier) ----
   Une synchronisation live à deux sens nécessiterait un serveur avec l'API
   Microsoft Graph + OAuth, impossible depuis une appli 100% front-end. On génère
   donc ici un fichier .ics standard téléchargeable, importable manuellement dans
   Outlook (Fichier > Ouvrir & exporter > Importer/Exporter, ou double-clic direct). */
function icsEscape(str) {
  return String(str || "").replace(/\\/g, "\\\\").replace(/;/g, "\\;").replace(/,/g, "\\,").replace(/\r?\n/g, "\\n");
}
function icsDateTime(dateISO, timeHHMM) {
  const [y, m, d] = dateISO.split("-").map(Number);
  const [h, min] = (timeHHMM || "09:00").split(":").map(Number);
  return `${String(y).padStart(4, "0")}${String(m).padStart(2, "0")}${String(d).padStart(2, "0")}T${String(h).padStart(2, "0")}${String(min).padStart(2, "0")}00`;
}
function icsDateTimePlusMinutes(dateISO, timeHHMM, minutes) {
  const [y, m, d] = dateISO.split("-").map(Number);
  const [h, min] = (timeHHMM || "09:00").split(":").map(Number);
  const dt = new Date(y, m - 1, d, h, min);
  dt.setMinutes(dt.getMinutes() + (minutes || 60));
  return `${dt.getFullYear()}${String(dt.getMonth() + 1).padStart(2, "0")}${String(dt.getDate()).padStart(2, "0")}T${String(dt.getHours()).padStart(2, "0")}${String(dt.getMinutes()).padStart(2, "0")}00`;
}
function buildPlanningICS(tasks, clientById) {
  const now = new Date();
  const stamp = `${now.getUTCFullYear()}${String(now.getUTCMonth() + 1).padStart(2, "0")}${String(now.getUTCDate()).padStart(2, "0")}T${String(now.getUTCHours()).padStart(2, "0")}${String(now.getUTCMinutes()).padStart(2, "0")}${String(now.getUTCSeconds()).padStart(2, "0")}Z`;
  const lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//AXE-EXPERTS//Planning//FR", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"];
  tasks.filter((t) => t.date_echeance && t.heure_debut).forEach((t) => {
    const client = clientById[t.client_id];
    const summary = client ? `${client.nom} — ${t.nom}` : t.nom;
    lines.push("BEGIN:VEVENT");
    lines.push(`UID:${t.id}@axe-experts.planning`);
    lines.push(`DTSTAMP:${stamp}`);
    lines.push(`DTSTART:${icsDateTime(t.date_echeance, t.heure_debut)}`);
    lines.push(`DTEND:${icsDateTimePlusMinutes(t.date_echeance, t.heure_debut, t.duree_minutes || 60)}`);
    lines.push(`SUMMARY:${icsEscape(summary)}`);
    if (t.commentaire) lines.push(`DESCRIPTION:${icsEscape(t.commentaire)}`);
    lines.push("END:VEVENT");
  });
  lines.push("END:VCALENDAR");
  return lines.join("\r\n");
}
function exportPlanningToICS(tasks, clientById) {
  const ics = buildPlanningICS(tasks, clientById);
  const blob = new Blob([ics], { type: "text/calendar;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `planning-axe-experts-${todayISO()}.ics`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
const PLANNING_FILTERS = [
  { id: "toutes", label: "Toutes" },
  { id: "retard", label: "En retard" },
  { id: "aujourdhui", label: "Aujourd'hui" },
  { id: "semaine", label: "Cette semaine" },
];

function PlanningTaskCard({ task, client, draggable = true, onOpenClient }) {
  const tone = TASK_PRIORITE_TONE[task.priorite] || "neutral";
  return (
    <div
      draggable={draggable}
      onDragStart={(e) => draggable && e.dataTransfer.setData("text/plain", JSON.stringify({ type: "task", id: task.id }))}
      onClick={() => task.isAuto && client && onOpenClient && onOpenClient(client.id)}
      style={{
        background: T.card, border: `1px solid ${T.line}`, borderLeft: `4px solid ${tone === "red" ? T.red : tone === "amber" ? T.amber : T.navy}`,
        borderRadius: 9, padding: "9px 10px", marginBottom: 7, cursor: task.isAuto ? "pointer" : (draggable ? "grab" : "default"), boxShadow: T.shadowSm,
      }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: T.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {client ? client.nom : "Sans dossier"}
        </span>
        {task.isAuto ? <Stamped tone="neutral" small>Auto</Stamped> : <Stamped tone={tone} small>{TASK_PRIORITE_BY_CODE[task.priorite]?.label}</Stamped>}
      </div>
      <div style={{ fontSize: 11, color: T.inkMuted, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{task.nom}</div>
    </div>
  );
}

function PlanningView({ tasks, clients, me, onUpdate, onOpenClient }) {
  const clickTimeoutRef = useRef(null);
  const [weekOffset, setWeekOffset] = useState(0);
  const [filter, setFilter] = useState("toutes");
  const [dragOverCell, setDragOverCell] = useState(null);

  const weekStart = useMemo(() => { const s = startOfWeek(new Date()); s.setDate(s.getDate() + weekOffset * 7); return s; }, [weekOffset]);
  const weekDays = useMemo(() => Array.from({ length: 5 }, (_, i) => { const d = new Date(weekStart); d.setDate(d.getDate() + i); return d; }), [weekStart]);
  const clientById = useMemo(() => Object.fromEntries(clients.map((c) => [c.id, c])), [clients]);

  const unscheduled = useMemo(() => {
    let list = tasks.filter((t) => t.statut !== "termine" && !t.heure_debut);
    if (filter !== "toutes") list = list.filter((t) => planningBucket(t, weekStart) === filter);
    return list.sort((a, b) => (a.date_echeance || "9999").localeCompare(b.date_echeance || "9999"));
  }, [tasks, filter, weekStart]);

  const scheduled = useMemo(() => tasks.filter((t) => t.statut !== "termine" && t.heure_debut && t.date_echeance), [tasks]);
  const isoOf = (d) => d.toISOString().slice(0, 10);

  const handleDrop = (dayIso, hour) => (e) => {
    e.preventDefault();
    setDragOverCell(null);
    let data;
    try { data = JSON.parse(e.dataTransfer.getData("text/plain")); } catch { return; }
    if (!data?.id) return;
    onUpdate(data.id, { date_echeance: dayIso, heure_debut: `${String(hour).padStart(2, "0")}:00` });
  };

  const unschedule = (task) => onUpdate(task.id, { heure_debut: null });

  return (
    <div>
      <Reveal>
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", flexWrap: "wrap", gap: 8, marginBottom: 4 }}>
          <div>
            <h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: 0 }}>Mon planning</h1>
            <p style={{ color: T.inkMuted, fontSize: 12.5, margin: "4px 0 0" }}>Glissez une tâche sur un créneau pour la planifier, {me}.</p>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button onClick={() => setWeekOffset((w) => w - 1)} style={navBtnStyle}><ChevronLeft size={15} /></button>
            <span style={{ fontFamily: T.serif, fontWeight: 600, fontSize: 12.5, color: T.navy, minWidth: 150, textAlign: "center" }}>
              Semaine du {fmtFR(isoOf(weekStart))}
            </span>
            <button onClick={() => setWeekOffset((w) => w + 1)} style={navBtnStyle}><ChevronRight size={15} /></button>
            {weekOffset !== 0 && <button onClick={() => setWeekOffset(0)} style={{ ...navBtnStyle, width: "auto", padding: "0 10px", fontSize: 11 }}>Aujourd'hui</button>}
            <button onClick={() => exportPlanningToICS(scheduled, clientById)} title="Télécharge un fichier .ics avec toutes les tâches planifiées, à importer dans Outlook"
              style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${T.line}`, borderRadius: 9, padding: "7px 12px", fontSize: 11.5, fontWeight: 700, color: T.navy, cursor: "pointer" }}>
              <Download size={13} /> Exporter vers Outlook
            </button>
          </div>
        </div>
      </Reveal>

      <div style={{ display: "grid", gridTemplateColumns: "240px 1fr", gap: 14, marginTop: 16, alignItems: "start" }}>
        {/* ---------- À planifier ---------- */}
        <div style={{ background: T.card, border: `1px solid ${T.line}`, borderRadius: T.radius, padding: 12, boxShadow: T.shadowSm }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: T.ink, marginBottom: 8 }}>
            À planifier <span style={{ color: T.inkMuted, fontWeight: 500 }}>({unscheduled.length})</span>
          </div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 5, marginBottom: 10 }}>
            {PLANNING_FILTERS.map((f) => (
              <button key={f.id} onClick={() => setFilter(f.id)} style={{
                fontSize: 10.5, padding: "4px 9px", borderRadius: 999, fontWeight: 700, cursor: "pointer",
                border: `1px solid ${filter === f.id ? T.navy : T.line}`,
                background: filter === f.id ? T.navySoft : T.card, color: filter === f.id ? T.navy : T.inkMuted,
              }}>{f.label}</button>
            ))}
          </div>
          <div className="scrollbar" style={{ maxHeight: 560, overflowY: "auto" }}>
            {unscheduled.length === 0 ? <EmptyNote text="Tout est planifié." /> : unscheduled.map((t) => (
              <PlanningTaskCard key={t.id} task={t} client={clientById[t.client_id]} onOpenClient={onOpenClient} />
            ))}
          </div>
        </div>

        {/* ---------- Grille semaine ---------- */}
        <div className="scrollbar" style={{ background: T.card, border: `1px solid ${T.line}`, borderRadius: T.radius, overflow: "auto", boxShadow: T.shadowSm }}>
          <div style={{ display: "grid", gridTemplateColumns: "52px repeat(5, minmax(140px, 1fr))", minWidth: 760 }}>
            <div style={{ borderBottom: `1px solid ${T.line}`, borderRight: `1px solid ${T.line}` }} />
            {weekDays.map((d) => {
              const isToday = sameDay(d, new Date());
              return (
                <div key={d.toISOString()} style={{
                  padding: "9px 8px", textAlign: "center", borderBottom: `1px solid ${T.line}`, borderRight: `1px solid ${T.line}`,
                  background: isToday ? T.navySoft : T.paper,
                }}>
                  <div style={{ fontSize: 10, textTransform: "uppercase", color: T.inkMuted, fontWeight: 700 }}>{d.toLocaleDateString("fr-FR", { weekday: "short" })}</div>
                  <div style={{ fontFamily: T.serif, fontWeight: 700, fontSize: 13, color: isToday ? T.navy : T.ink }}>{d.getDate()}</div>
                </div>
              );
            })}

            {PLANNING_HOURS.map((h) => (
              <React.Fragment key={h}>
                <div style={{ fontSize: 10.5, color: T.inkMuted, textAlign: "right", paddingRight: 6, paddingTop: 3, borderRight: `1px solid ${T.line}`, borderBottom: `1px solid ${T.line}`, height: PLANNING_SLOT_H }}>{h}h</div>
                {weekDays.map((d) => {
                  const dayIso = isoOf(d);
                  const cellKey = `${dayIso}-${h}`;
                  const items = scheduled.filter((t) => t.date_echeance === dayIso && parseInt((t.heure_debut || "").slice(0, 2), 10) === h);
                  const isOver = dragOverCell === cellKey;
                  return (
                    <div key={cellKey}
                      onDragOver={(e) => { e.preventDefault(); setDragOverCell(cellKey); }}
                      onDragLeave={() => setDragOverCell((c) => (c === cellKey ? null : c))}
                      onDrop={handleDrop(dayIso, h)}
                      style={{
                        borderRight: `1px solid ${T.line}`, borderBottom: `1px solid ${T.line}`, minHeight: PLANNING_SLOT_H,
                        background: isOver ? T.navySoft : "transparent", padding: 3,
                      }}>
                      {items.map((t) => {
                        const dureeH = Math.max(1, Math.round((t.duree_minutes || 60) / 60));
                        const client = clientById[t.client_id];
                        const tone = TASK_PRIORITE_TONE[t.priorite] || "neutral";
                        const bg = tone === "red" ? T.redSoft : tone === "amber" ? T.amberSoft : T.navySoft;
                        const fg = tone === "red" ? T.red : tone === "amber" ? T.amber : T.navy;
                        return (
                          <div key={t.id}
                            draggable
                            onDragStart={(e) => e.dataTransfer.setData("text/plain", JSON.stringify({ type: "task", id: t.id }))}
                            onClick={() => {
  if (clickTimeoutRef.current) return;
  clickTimeoutRef.current = setTimeout(() => {
    clickTimeoutRef.current = null;
    client && onOpenClient(client.id);
  }, 250);
}}
onDoubleClick={(e) => {
  e.stopPropagation();
  clearTimeout(clickTimeoutRef.current);
  clickTimeoutRef.current = null;
  unschedule(t);
}}
                            title="Double-clic pour déplanifier"
                            style={{
                              background: bg, border: `1px solid ${fg}33`, borderRadius: 7, padding: "4px 6px", cursor: "grab",
                              height: `calc(${dureeH * PLANNING_SLOT_H}px - 6px)`, overflow: "hidden",
                            }}>
                            <div style={{ fontSize: 10.5, fontWeight: 700, color: fg, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{client ? client.nom : t.nom}</div>
                            <div style={{ fontSize: 9.5, color: fg, opacity: 0.85, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.nom}</div>
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </React.Fragment>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   ÉQUIPE — rôles, portefeuilles (cabinets), demandes en attente.
   L'affectation d'un portefeuille / rôle est réservée aux comptes
   Expert, Chef de mission (sur leur propre portefeuille) et Admin
   (partout) — appliqué à la fois ici (UI) et côté base (RLS).
   ============================================================ */
function EquipeView({ team, portefeuilles, clients, myRole, isAdmin, myPortefeuilleId, canManageTeam, onAdd, onRename, onDelete, onUpdateMember, onAddPortefeuille }) {
  const [newName, setNewName] = useState("");
  const [newRole, setNewRole] = useState("collaborateur");
  const [newPortefeuille, setNewPortefeuille] = useState(myPortefeuilleId || "");
  const [editing, setEditing] = useState(null);
  const [editValue, setEditValue] = useState("");
  const [pfNom, setPfNom] = useState("");
  const [pfDomaine, setPfDomaine] = useState("");
  const [validating, setValidating] = useState(null); // id de la demande en cours de validation
  const [validatePortefeuille, setValidatePortefeuille] = useState("");
  const [validateNewPf, setValidateNewPf] = useState(false);
  const [validateNewPfNom, setValidateNewPfNom] = useState("");
  const [validateNewPfDomaine, setValidateNewPfDomaine] = useState("");
  const [validateRole, setValidateRole] = useState("collaborateur");

  const countFor = (nom) => clients.filter((c) => c.collab === nom || c.expert === nom || c.chefMission === nom).length;
  const portefeuilleName = (id) => portefeuilles.find((p) => p.id === id)?.nom || "—";

  const pending = team.filter((t) => t.statut === "en_attente");
  const activeTeam = team.filter((t) => t.statut !== "en_attente");
  // Regroupement par portefeuille (utile surtout pour l'Admin, qui voit tous les cabinets)
  const groups = useMemo(() => {
    const byId = new Map();
    activeTeam.forEach((t) => {
      const key = t.portefeuille_id || "—";
      if (!byId.has(key)) byId.set(key, []);
      byId.get(key).push(t);
    });
    return Array.from(byId.entries());
  }, [activeTeam]);

  const startValidation = (row) => {
    setValidating(row.id);
    setValidatePortefeuille(portefeuilles[0]?.id || "");
    setValidateNewPf(false);
    setValidateNewPfNom(row.cabinet_nom || "");
    setValidateNewPfDomaine("");
    setValidateRole("collaborateur");
  };

  const confirmValidation = (row) => {
    if (validateNewPf) {
      if (!validateNewPfNom.trim()) return;
      const id = onAddPortefeuille(validateNewPfNom, validateNewPfDomaine);
      onUpdateMember(row.id, { portefeuille_id: id, role: validateRole, statut: "actif" });
    } else {
      if (!validatePortefeuille) return;
      onUpdateMember(row.id, { portefeuille_id: validatePortefeuille, role: validateRole, statut: "actif" });
    }
    setValidating(null);
  };

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Équipe</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>
        {isAdmin ? "Gérez les portefeuilles (cabinets), les rôles et les demandes d'accès." : "Rôles et affectations de votre cabinet."}
      </p>

      {!canManageTeam && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: T.inkMuted, background: T.paper, border: `1px solid ${T.line}`, borderRadius: 12, padding: "10px 14px", marginBottom: 16 }}>
          <ShieldAlert size={15} />
          Seuls les Experts et Chefs de mission peuvent modifier les rôles et affecter des portefeuilles.
        </div>
      )}

      {isAdmin && pending.length > 0 && (
        <>
          <Panel title={`Demandes en attente (${pending.length})`}>
            {pending.map((row) => (
              <div key={row.id} style={{ padding: "12px 4px", borderBottom: `1px solid ${T.line}` }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <span style={{ width: 26, height: 26, borderRadius: "50%", background: row.color, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: T.serif, fontWeight: 600, fontSize: 12, color: "#fff", flexShrink: 0 }}>{row.nom?.[0]}</span>
                  <div style={{ flex: 1, minWidth: 160 }}>
                    <div style={{ fontWeight: 700, fontSize: 12.5 }}>{row.nom}</div>
                    <div style={{ fontSize: 11, color: T.inkMuted }}>{row.email}{row.telephone ? ` · ${row.telephone}` : ""}{row.cabinet_nom ? ` · ${row.cabinet_nom}` : ""}</div>
                  </div>
                  {validating !== row.id && (
                    <button onClick={() => startValidation(row)} style={{ background: T.navy, color: "#fff", border: "none", borderRadius: 9, padding: "7px 14px", fontSize: 11.5, fontWeight: 700, cursor: "pointer" }}>Valider l'accès</button>
                  )}
                </div>
                {validating === row.id && (
                  <div style={{ marginTop: 10, padding: 12, background: T.paper, borderRadius: 12, display: "grid", gap: 8 }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                      <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11.5 }}>
                        <input type="radio" checked={!validateNewPf} onChange={() => setValidateNewPf(false)} /> Portefeuille existant
                      </label>
                      <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11.5 }}>
                        <input type="radio" checked={validateNewPf} onChange={() => setValidateNewPf(true)} /> Nouveau portefeuille
                      </label>
                    </div>
                    {!validateNewPf ? (
                      <select value={validatePortefeuille} onChange={(e) => setValidatePortefeuille(e.target.value)} style={inputStyle}>
                        {portefeuilles.map((p) => <option key={p.id} value={p.id}>{p.nom}{p.domaine ? ` (${p.domaine})` : ""}</option>)}
                      </select>
                    ) : (
                      <div style={{ display: "flex", gap: 8 }}>
                        <input placeholder="Nom du cabinet" value={validateNewPfNom} onChange={(e) => setValidateNewPfNom(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
                        <input placeholder="Domaine email (ex. cabinet.fr)" value={validateNewPfDomaine} onChange={(e) => setValidateNewPfDomaine(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
                      </div>
                    )}
                    <select value={validateRole} onChange={(e) => setValidateRole(e.target.value)} style={inputStyle}>
                      <option value="collaborateur">Collaborateur</option>
                      <option value="expert">Expert</option>
                      <option value="chef_mission">Chef de mission</option>
                    </select>
                    <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                      <button onClick={() => setValidating(null)} style={{ padding: "7px 12px", borderRadius: 9, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 11.5 }}>Annuler</button>
                      <button onClick={() => confirmValidation(row)} style={{ padding: "7px 14px", borderRadius: 9, border: "none", background: T.green, color: "#fff", cursor: "pointer", fontSize: 11.5, fontWeight: 700 }}>Confirmer</button>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </Panel>
          <div style={{ height: 16 }} />
        </>
      )}

      {isAdmin && (
        <>
          <Panel title="Portefeuilles (cabinets)">
            {portefeuilles.map((p) => (
              <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "10px 4px", borderBottom: `1px solid ${T.line}` }}>
                <Wallet size={15} color={T.navy} style={{ flexShrink: 0 }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 700, fontSize: 12.5 }}>{p.nom}</div>
                  {p.domaine && <div style={{ fontSize: 11, color: T.inkMuted }}>@{p.domaine}</div>}
                </div>
                <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{activeTeam.filter((t) => t.portefeuille_id === p.id).length} membre(s)</span>
              </div>
            ))}
            <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
              <input placeholder="Nom du nouveau cabinet" value={pfNom} onChange={(e) => setPfNom(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
              <input placeholder="Domaine email (ex. cabinet.fr)" value={pfDomaine} onChange={(e) => setPfDomaine(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
              <button onClick={() => { if (pfNom.trim()) { onAddPortefeuille(pfNom, pfDomaine); setPfNom(""); setPfDomaine(""); } }} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer", whiteSpace: "nowrap" }}>
                <Plus size={15} /> Créer
              </button>
            </div>
          </Panel>
          <div style={{ height: 16 }} />
        </>
      )}

      {isAdmin && (
        <>
          <Panel title="Ajouter un membre manuellement">
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Nom du collaborateur" style={{ ...inputStyle, flex: 1, minWidth: 160 }} />
              <select value={newRole} onChange={(e) => setNewRole(e.target.value)} style={inputStyle}>
                <option value="collaborateur">Collaborateur</option>
                <option value="expert">Expert</option>
                <option value="chef_mission">Chef de mission</option>
              </select>
              <select value={newPortefeuille} onChange={(e) => setNewPortefeuille(e.target.value)} style={inputStyle}>
                <option value="">Aucun portefeuille</option>
                {portefeuilles.map((p) => <option key={p.id} value={p.id}>{p.nom}</option>)}
              </select>
              <button onClick={() => { onAdd(newName, newPortefeuille, newRole); setNewName(""); }} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
                <Plus size={15} /> Ajouter
              </button>
            </div>
            <div style={{ fontSize: 10.5, color: T.inkMuted, marginTop: 8 }}>Réservé aux entrées sans compte (contact externe, dépannage) — les collaborateurs rejoignent normalement en s'inscrivant eux-mêmes.</div>
          </Panel>
          <div style={{ height: 16 }} />
        </>
      )}

      {groups.map(([pfId, members]) => (
        <React.Fragment key={pfId}>
          <Panel title={isAdmin ? `${portefeuilleName(pfId)} (${members.length})` : `Membres de l'équipe (${members.length})`}>
            {members.map((t) => (
              <div key={t.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 4px", borderBottom: `1px solid ${T.line}`, flexWrap: "wrap" }}>
                <span style={{ width: 26, height: 26, borderRadius: "50%", background: t.color, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: T.serif, fontWeight: 600, fontSize: 12, color: "#fff", flexShrink: 0 }}>{t.nom?.[0]}</span>
                {editing === t.id ? (
                  <input value={editValue} onChange={(e) => setEditValue(e.target.value)} autoFocus
                    onBlur={() => { onRename(t.nom, editValue); setEditing(null); }}
                    onKeyDown={(e) => { if (e.key === "Enter") { onRename(t.nom, editValue); setEditing(null); } }}
                    style={{ ...inputStyle, flex: 1, minWidth: 120 }} />
                ) : (
                  <div style={{ flex: 1, minWidth: 120 }}>
                    <span style={{ fontWeight: 600, fontSize: 12.5 }}>{t.nom}</span>
                    {t.email && <div style={{ fontSize: 10.5, color: T.inkMuted }}>{t.email}</div>}
                  </div>
                )}
                {canManageTeam ? (
                  <select value={t.role || "collaborateur"} onChange={(e) => onUpdateMember(t.id, { role: e.target.value })} style={{ ...inputStyle, padding: "6px 10px", fontSize: 11.5 }}>
                    <option value="collaborateur">Collaborateur</option>
                    <option value="expert">Expert</option>
                    <option value="chef_mission">Chef de mission</option>
                    {t.role === "admin" && <option value="admin">Admin</option>}
                  </select>
                ) : (
                  <span style={{ fontSize: 11, fontWeight: 600, color: T.navy, background: T.navySoft, padding: "3px 9px", borderRadius: 999 }}>{ROLE_LABELS[t.role] || t.role}</span>
                )}
                {isAdmin && (
                  <select value={t.portefeuille_id || ""} onChange={(e) => onUpdateMember(t.id, { portefeuille_id: e.target.value || null })} style={{ ...inputStyle, padding: "6px 10px", fontSize: 11.5 }}>
                    <option value="">Aucun portefeuille</option>
                    {portefeuilles.map((p) => <option key={p.id} value={p.id}>{p.nom}</option>)}
                  </select>
                )}
                <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{countFor(t.nom)} dossier(s)</span>
                {canManageTeam && (
                  <>
                    <button onClick={() => { setEditing(t.id); setEditValue(t.nom); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted }}><Pencil size={14} /></button>
                    <button onClick={() => { if (confirm(`Supprimer ${t.nom} de l'équipe ? Ses dossiers seront désassignés.`)) onDelete(t.nom); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.red }}><Trash2 size={14} /></button>
                  </>
                )}
              </div>
            ))}
          </Panel>
          <div style={{ height: 16 }} />
        </React.Fragment>
      ))}
    </div>
  );
}

/* ============================================================
   ADD CLIENT MODAL
   ============================================================ */
function AddClientModal({ team, me, portefeuilleId, onClose, onCreate }) {
  const [nom, setNom] = useState("");
  const [siren, setSiren] = useState("");
  const [logiciel, setLogiciel] = useState("MYUNISOFT");
  const [dateCloture, setDateCloture] = useState(`${new Date().getFullYear()}-12-31`);
  const [collab, setCollab] = useState(me);
  const [expert, setExpert] = useState("");
  const [chefMission, setChefMission] = useState("");
  const teamNames = team.map((t) => t.nom);

  const submit = () => {
    if (!nom.trim()) return;
    onCreate({
      id: `c-${Date.now()}`, portefeuilleId, statutDossier: "actif", nom: nom.trim(), siren: siren.trim(), logiciel, dateCloture,
      collab, expert, chefMission, formeJuridique: "", capital: "", activite: "",
      tvaRegime: "", tvaExig: "", tvaMois: {}, regimeHistory: [], ageAgoHistory: {}, formeJuridiqueHistory: {},
      corporate: { kyc: { lab: false, mandat: false, choixPA: "", beneficiaireEffectif: false, beneficiaireNom: "" }, kycExtra: [], notes: "" },
      mission: { "KBIS": false, "Statuts": false, "CNI dirigeants": false, "CNI associés": false, "Notes entrée mission / Devizen": false, "Acceptation mission": false, "LM à jour": false, "LAB / Kanta / Devizen à jour": false, "Bouclage": false, "Fiche client": false },
    });
  };

  return (
    <div style={{ position: "fixed", inset: 0, zIndex: 60, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div onClick={onClose} style={{ position: "absolute", inset: 0, background: "rgba(28,37,65,0.4)" }} />
      <div className="scrollbar" style={{ position: "relative", background: T.paper, borderRadius: 14, padding: 26, width: 420, maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(0,0,0,0.25)" }}>
        <h3 style={{ fontFamily: T.serif, fontSize: 15, fontWeight: 600, color: T.navy, margin: "0 0 16px" }}>Nouveau dossier</h3>
        <div style={{ display: "grid", gap: 10 }}>
          <input autoFocus placeholder="Nom du client" value={nom} onChange={(e) => setNom(e.target.value)} style={inputStyle} />
          <input placeholder="SIREN" value={siren} onChange={(e) => setSiren(e.target.value)} style={inputStyle} />
          <select value={logiciel} onChange={(e) => setLogiciel(e.target.value)} style={inputStyle}>
            <option value="MYUNISOFT">MYUNISOFT</option><option value="QUADRA">QUADRA</option>
          </select>
          <div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Date de clôture d'exercice</div>
            <input type="date" value={dateCloture} onChange={(e) => setDateCloture(e.target.value)} style={{ ...inputStyle, width: "100%" }} />
          </div>
          <div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Collaborateur</div>
            <select value={collab} onChange={(e) => setCollab(e.target.value)} style={{ ...inputStyle, width: "100%" }}>
              <option value="">—</option>{teamNames.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Expert</div>
            <select value={expert} onChange={(e) => setExpert(e.target.value)} style={{ ...inputStyle, width: "100%" }}>
              <option value="">—</option>{teamNames.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <div style={{ fontSize: 11, color: T.inkMuted, marginBottom: 4 }}>Chef de mission</div>
            <select value={chefMission} onChange={(e) => setChefMission(e.target.value)} style={{ ...inputStyle, width: "100%" }}>
              <option value="">—</option>{teamNames.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 20, justifyContent: "flex-end" }}>
          <button onClick={onClose} style={{ padding: "9px 14px", borderRadius: 10, border: `1px solid ${T.line}`, background: "none", cursor: "pointer", fontSize: 12 }}>Annuler</button>
          <button onClick={submit} style={{ padding: "9px 16px", borderRadius: 10, border: "none", background: T.navy, color: "#fff", cursor: "pointer", fontSize: 12, fontWeight: 600 }}>Créer le dossier</button>
        </div>
      </div>
    </div>
  );
}
const inputStyle = { padding: "10px 12px", borderRadius: 10, border: `1px solid ${T.line}`, fontSize: 12.5, background: T.card, color: T.ink };
