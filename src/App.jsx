import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutGrid, Users, Receipt, FileWarning, Landmark, Building2,
  ClipboardCheck, Search, ChevronRight, X, Check, AlertTriangle,
  Clock, TrendingUp, UserCircle2, Plus, Stamp, ChevronDown,
  Filter, ArrowUpRight, CircleDot, Loader2, RefreshCw, History,
  ChevronUp, CalendarDays, CalendarRange, Settings2, Trash2,
  Pencil, ChevronLeft, ShieldCheck, Home, LogOut, Mail, Lock, UserRound,
  Phone, Briefcase, UserCheck, Wallet, ShieldAlert, Menu, Bell
} from "lucide-react";
import { supabase } from "./supabaseClient";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";
import { fetchTasks, createTask, updateTask, completeTask, deleteTask, subscribeTasks } from "./services/tasks";
import { logActivity, activityMessages } from "./services/activity";
import { bucketize as bucketizeDeadlines, BUCKET_LABELS as DEADLINE_BUCKET_LABELS } from "./services/deadlines";
import { TASK_STATUTS, TASK_STATUT_BY_CODE, TASK_PRIORITES, TASK_PRIORITE_BY_CODE, taskSortWeight, PILOTAGE_COLORS } from "./constants/pilotage";

const T = {
  paper: "#F3F4F6", paperDeep: "#EEF2FF", ink: "#0F172A", inkSoft: "#475569", inkMuted: "#94A3B8",
  line: "#E2E8F0", card: "#FFFFFF", gold: "#D97706", goldSoft: "#FEF3C7",
  green: "#16A34A", greenSoft: "#DCFCE7", red: "#DC2626", redSoft: "#FEE2E2",
  amber: "#D97706", amberSoft: "#FEF3C7",
  /* accent sobre façon Kabineo (indigo) */
  navy: "#4F46E5", navySoft: "#EEF2FF",
  /* sidebar sombre façon "slate/ardoise" : fond bleu-nuit très foncé, textes clairs */
  sidebarBg: "#0F172A", sidebarBg2: "#1E293B", sidebarInk: "#E2E8F0", sidebarInkMuted: "#94A3B8",
  sidebarActive: "rgba(99,102,241,0.22)", sidebarBorder: "#1E293B", sidebarAccent: "#818CF8",
  serif: "'Plus Jakarta Sans', 'Inter', -apple-system, sans-serif", mono: "'JetBrains Mono', ui-monospace, monospace", sans: "'Inter', -apple-system, sans-serif",
  shadow: "0 1px 2px rgba(15,23,42,0.04), 0 8px 20px -6px rgba(15,23,42,0.08)",
  shadowSm: "0 1px 2px rgba(15,23,42,0.05), 0 1px 3px rgba(15,23,42,0.06)",
  shadowLg: "0 24px 48px -14px rgba(15,23,42,0.16)",
  radius: 16, radiusSm: 12, radiusLg: 20,
};

/* ============================================================
   SEED DATA — extrait du fichier Excel de suivi du cabinet
   ============================================================ */
const RAW_SEED_CLIENTS = [{"nom":"A&D RESTOS","siren":"81276334","logiciel":"MYUNISOFT","collab":"Cheikh","tvaRegime":"CA12","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"AC INVEST","siren":"925320210","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"AD SOLUTION","siren":"942467515","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"AE BAT","siren":"931778112","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ALLO SOS MOTO","siren":"488698960","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":19,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ALPHA DIGITAL","siren":"920603560","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"AMS PLOMBERIE","siren":"917541906","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"APEL","siren":"326627247","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":19,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ARSA","siren":"954016481","logiciel":"QUADRA","collab":"Jacques","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ATELIER GOURMAND","siren":"84531968","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ATI INVEST","siren":"927855247","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"AU COIN DU PAIN","siren":"993062835","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BACKSTAGE BEAUTY GROUP","siren":"853273522","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BANGLA.COM","siren":"804359362","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BELARBI ABDALLAH","siren":"530122589","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BELARBI ABDELKAOUI","siren":"431443852","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BENGAL COIFFURE","siren":"830473351","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BHUVI BEAUTE","siren":"851770354","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BLAST","siren":"831893698","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BLUE SECURITY","siren":"929357168","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"BOUCHERIE SERAS","siren":"930591714","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"TRIM","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"CAFFE ITALIA","siren":"790500912","logiciel":"MYUNISOFT","collab":"Jacques","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"CELIA","siren":"923210215","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"CENTRAL'AUTO","siren":"891458457","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"CHARLOTTE FRANCISCO","siren":"819855727","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"CHATSN TRANSPORT BILAN 2025","siren":"789814399","logiciel":"QUADRA","collab":"Jacques","tvaRegime":"CA3","tvaExig":20,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"DAKAROIS KITCHEN","siren":"981110026","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"TRIM","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"DAVIDSEN","siren":"849091400","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"DESTOCK PIECES AUTO LE PERREUX","siren":"811969153","logiciel":"","collab":"Cheikh","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"DIALLO KAMION","siren":"809583669","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"DIAMOND SUSHI","siren":"890451271","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ECO NOISY TRANSPORT BILAN 2025","siren":"850096587","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ENERGIA","siren":"..","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ERATOMBE","siren":"849495742","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"EVE SECURITY","siren":"504487216","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":20,"tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"EXTERNALIS-CORPORATION","siren":"853414910","logiciel":"","collab":"Cheikh","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"FAIM DE SEMAINE","siren":"898044110","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"FBA BAT","siren":"538695313","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"FOURNIL JEAN XXIII","siren":"893192138","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"GALLERY BOUNAN CH","siren":"803649110","logiciel":"QUADRA","collab":"Jacques","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"GROUPE PNS","siren":"791864317","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"GTM","siren":"933355679","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"GUT HEALTH & WELLNESS","siren":"912855350","logiciel":"MYUNISOFT","collab":"Cheikh","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"HOLDING RAZA","siren":"939739967","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"I PRO BATIMENT","siren":"931787204","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"INIT SERVICES","siren":"941381568","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"KAILEY RENOVATION","siren":"842455537","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"KD","siren":"884824566","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"KHOUJABAT","siren":"894352939","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LA BONNE EPOQUE","siren":"833393598","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LE CENTRE MEDICAL DE VERDUN","siren":"897427761","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LE JUSTE PRIX","siren":"934480039","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LE PETIT MARCHE","siren":"837752062","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LEMNISCATE SOFTWARE","siren":"830376653","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"L'EPI D'OR","siren":"984628628","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LMG TRANSPORT","siren":"809937717","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"LO NAFI","siren":"904900065","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MAC CHICKEN","siren":"843437724","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"TRIM","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"Madame DIALLO KAMION","siren":"809583669","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MB LAUNDRY","siren":"940912835","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MED FOOD","siren":"951966795","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MEYO","siren":"983592122","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MFC DISTRIBUTION","siren":"989941216","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MIX TRAVAUX IDF","siren":"840772172","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MS TRANS (MTRANS SERVICES","siren":"893029538","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MTP HOLDING","siren":"978063147","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"MY GLOBAL","siren":"922568092","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"NETEN","siren":"511556193","logiciel":"QUADRA","collab":"Soli","tvaRegime":"TRIM","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"NEW STAR HOLDING INTERNATIONAL","siren":"907566871","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"NIA CONSEILS BILAN 2025","siren":"914568399","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"NICK SERVICES","siren":"512395823","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"NIRALI","siren":"819586017","logiciel":"QUADRA","collab":"Soli","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"NISHA ESTHETIQUE","siren":"880336128","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"O DELICE","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ON EST LA","siren":"929695120","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"OPTIMUS TECHNOLOGIES","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"OZONE HYGIENE ENVIRONNEMENT","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"PARIS CASH AND CARRY","siren":"803981943","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"PERI ALIMENTATION","siren":"797918489","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"PLACE COLETTE","siren":"928971233","logiciel":"MYUNISOFT","collab":"Cheikh","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"POWERFIT","siren":"83083017","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"RAMY 37","siren":"884284324","logiciel":"QUADRA","collab":"Jacques","tvaRegime":"CA3","tvaExig":24,"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"RED LIFE FRANCE","siren":"921174686","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"RED LIFE HOLDING","siren":"919332106","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"REPAIR MASTER","siren":"103229209","logiciel":"MYUNISOFT","collab":"Soli","tvaRegime":"CA3","tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"REPON SERGE AGRICULTURE","siren":"348050287","logiciel":"QUADRA","collab":"Cheikh","tvaRegime":"TRIM","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SABP","siren":"..","logiciel":"","collab":"Cheikh","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SAINT AMBROISE SAS","siren":"920863982","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SAISANTHU SOCIETE DE NETTOYAGE","siren":"940616766","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SCI KAMY TEAM","siren":"883778268","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SCI LES MANOS","siren":"499319796","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"sci p immo","siren":"825254378","logiciel":"QUADRA","collab":"Soli","tvaRegime":"TRIM","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SCI SHAANA","siren":"940587819","logiciel":"QUADRA","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SHAH JALAL 76","siren":"920826591","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SHIV-SAI","siren":"534331368","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":19,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SKYTECH","siren":"922167713","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SPEKCOM","siren":"788795631","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":20,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"STARTED FROM THE BOTTOM FACILITY SERVICES","siren":"832559884","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"STEEL PAINT","siren":"530937754","logiciel":"QUADRA","collab":"Cheikh","tvaRegime":"CA3","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"SUSHI KID","siren":"793336025","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":21,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"système automatique et securité","siren":"838588929","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"TIROUCHE Sofiane","siren":"812371276","logiciel":"QUADRA","collab":"Soli","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"tms concept","siren":"883337503","logiciel":"QUADRA","collab":"Soli","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"TRANS TP IDF","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"TRANSPORT FRET LOGISTIQUE","siren":"..","logiciel":"","collab":"Cheikh","tvaRegime":"CA12","tvaMois":{"Jan":"NA","Fév":"NA","Mar":"NA","Avr":"NA","Mai":"NA","Juin":"NA","Juil":"NA","Août":"NA","Sept":"NA","Oct":"NA","Nov":"NA","Déc":"NA"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"TSHI INVEST","siren":"927452532","logiciel":"MYUNISOFT","collab":"Jacques","tvaRegime":"CA12","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"WANDEE","siren":"833083546","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"CA3","tvaExig":24,"tvaMois":{"Jan":"OK","Fév":"OK","Mar":"OK","Avr":"OK","Mai":"OK","Juin":"OK","Juil":"FAIT"},"mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}},{"nom":"ZIANIDES","siren":"844162396","logiciel":"MYUNISOFT","collab":"Emilie","tvaRegime":"FEB","mission":{"Acceptation mission":false,"Lettre reprise":false,"LAB":false,"Fiche client":false,"Lettre mission":false,"Mandat":false,"Attestation bilan":false}}]
;

const PALETTE = ["#6366F1", "#10B981", "#F97316", "#EC4899", "#06B6D4", "#8B5CF6", "#F59E0B", "#14B8A6", "#F43F5E", "#3B82F6"];
const DEFAULT_TEAM = ["Cheikh", "Soli", "Emilie", "Jacques"].map((nom, i) => ({
  id: `seed-${i}`, nom, color: PALETTE[i % PALETTE.length],
}));

const MOIS_ORDER = ["Jan","Fév","Mar","Avr","Mai","Juin","Juil","Août","Sept","Oct","Nov","Déc"];
const MOIS_FULL = { Jan:"Janvier",Fév:"Février",Mar:"Mars",Avr:"Avril",Mai:"Mai",Juin:"Juin",Juil:"Juillet",Août:"Août",Sept:"Septembre",Oct:"Octobre",Nov:"Novembre",Déc:"Décembre" };
const REGIMES_TVA = ["CA3", "CA12", "FRANCHISE"];

function currentMonthKey() { return MOIS_ORDER[new Date().getMonth()]; }
function todayISO() { return new Date().toISOString().slice(0, 10); }
function fmtFR(iso) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  if (!y || !m || !d) return iso;
  return `${d}/${m}/${y}`;
}
function addMonthsISO(iso, months) {
  if (!iso) return null;
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1 + months, d);
  return dt.toISOString().slice(0, 10);
}
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
    if (!next.expert) next.expert = "";
    if (!next.chefMission) next.chefMission = "";
    if (!next.dateCloture) next.dateCloture = "";
    if (!next.formeJuridique) next.formeJuridique = "";
    if (!next.capital) next.capital = "";
    if (!next.activite) next.activite = "";
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
    return next;
  });
}

/* ============================================================
   IMPORT / EXPORT EXCEL — registre clients
   ============================================================ */
const EXCEL_COLUMNS = [
  { key: "nom", label: "Nom" },
  { key: "siren", label: "SIREN" },
  { key: "logiciel", label: "Logiciel" },
  { key: "collab", label: "Collaborateur" },
  { key: "expert", label: "Expert" },
  { key: "chefMission", label: "Chef de mission" },
  { key: "formeJuridique", label: "Forme juridique" },
  { key: "capital", label: "Capital" },
  { key: "activite", label: "Activité" },
  { key: "dateCloture", label: "Date clôture (AAAA-MM-JJ)" },
  { key: "tvaRegime", label: "Régime TVA" },
  { key: "tvaExig", label: "Jour exigibilité TVA" },
];
// En-têtes acceptés en entrée (tolère quelques variantes usuelles côté Excel)
const EXCEL_IMPORT_ALIASES = {
  nom: ["nom", "client", "dossier", "raison sociale"],
  siren: ["siren", "siret"],
  logiciel: ["logiciel"],
  collab: ["collaborateur", "collab"],
  expert: ["expert"],
  chefMission: ["chef de mission", "chefmission"],
  formeJuridique: ["forme juridique", "formejuridique"],
  capital: ["capital", "capital social"],
  activite: ["activité", "activite"],
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
// Lit un fichier .xlsx/.xls/.csv et retourne une liste d'objets clients partiels
// (uniquement les champs reconnus), prêts à être fusionnés avec le registre existant.
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
        const rows = raw.map((r) => {
          const out = {};
          Object.keys(r).forEach((h) => {
            const key = headerMap[h];
            if (!key) return;
            let v = r[h];
            if (key === "tvaExig") v = v === "" ? "" : parseInt(v, 10) || "";
            if (typeof v === "string") v = v.trim();
            out[key] = v;
          });
          return out;
        }).filter((r) => r.nom || r.siren);
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
  if (manual === "OK" || manual === "FAIT" || manual === "NA") return manual;

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
  const monthIdx = MOIS_ORDER.indexOf(moisKey);
  const now = new Date();
  // Régime CA3 : la TVA du mois M est déclarée en M+1 (ex. la TVA de juillet
  // est exigible en août, et ne passe donc en retard qu'en août).
  const deadline = new Date(now.getFullYear(), monthIdx + 1, exig, 23, 59, 59);
  return deadline.getTime() < now.getTime() ? "RETARD" : "";
}
function tvaTone(status) {
  return status === "OK" ? "green" : status === "FAIT" ? "amber" : status === "RETARD" ? "red" : "neutral";
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
      const status = effectiveTvaStatus(c, MOIS_ORDER[declaredMonthIdx]);
      if (status !== "OK" && status !== "NA") {
        events.push({
          id: `${c.id}-tva-${declaredMonthIdx}`, client: c, category: "TVA",
          label: `TVA ${MOIS_FULL[MOIS_ORDER[declaredMonthIdx]]}`,
          date: new Date(year, monthIdx, parseInt(c.tvaExig, 10) || 20),
          done: false, tone: tvaTone(status),
        });
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
    if (c.is?.concerne) {
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
    if (c.cfe?.concerne) {
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
    if (c.dateCloture) {
      const [cy, cm, cd] = c.dateCloture.split("-").map(Number);
      events.push({
        id: `${c.id}-cloture`, client: c, category: "Clôture",
        label: "Clôture d'exercice", date: new Date(cy, cm - 1, cd), done: false, tone: "neutral",
      });
      // Bilan (échéance approximative : clôture + 3 mois) si non finalisé
      if (c.bilan?.nonFinalise) {
        const echeance = addMonthsISO(c.dateCloture, 3);
        const [by, bm, bd] = echeance.split("-").map(Number);
        events.push({
          id: `${c.id}-bilan`, client: c, category: "Bilan",
          label: "Dépôt du bilan", date: new Date(by, bm - 1, bd), done: false, tone: "red",
        });
      }
    }
    // AGE/AGO — approbation ~6 mois après clôture si non tenue
    if (c.dateCloture) {
      const latestYear = Object.keys(c.ageAgoHistory || {}).sort((a, b) => b - a)[0];
      const y = latestYear ? c.ageAgoHistory[latestYear] : null;
      if (y && !y.ago) {
        const echeance = addMonthsISO(c.dateCloture, 6);
        const [ay, am, ad] = echeance.split("-").map(Number);
        events.push({
          id: `${c.id}-ago-${latestYear}`, client: c, category: "AGO",
          label: `Approbation des comptes ${latestYear}`, date: new Date(ay, am - 1, ad), done: false, tone: "amber",
        });
      }
    }
  });
  return events;
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
  const { error } = await supabase.from("clients").insert({ id, data: rest, portefeuille_id: portefeuilleId || null });
  if (error) console.error("Erreur création client :", error.message);
}
async function updateClientRemote(id, fullClient) {
  const { id: _drop, portefeuilleId: _drop2, ...rest } = fullClient;
  const { error } = await supabase.from("clients").update({ data: rest }).eq("id", id);
  if (error) console.error("Erreur sauvegarde client :", error.message);
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

const ROLE_LABELS = { collaborateur: "Collaborateur", expert: "Expert", chef_mission: "Chef de mission", admin: "Admin" };

/* ============================================================
   APP
   ============================================================ */
function CabinetApp({ session, onLogout }) {
  const [clients, setClients] = useState(null);
  const [team, setTeam] = useState(null);
  const [portefeuilles, setPortefeuilles] = useState(null);
  const [loading, setLoading] = useState(true);
  const [view, setView] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("Tous");
  const [regimeFilter, setRegimeFilter] = useState("Tous");
  const [showAddClient, setShowAddClient] = useState(false);
  const [saveStatus, setSaveStatus] = useState("idle");
  const [openClientTabs, setOpenClientTabs] = useState([]); // [{id, label}]
  const [activeClientTab, setActiveClientTab] = useState(null); // id du dossier affiché en page pleine, ou null
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [tasksDb, setTasksDb] = useState([]); // tâches réelles (table "tasks"), indépendantes des échéances fiscales

  // Empêche le canal temps réel de "rejouer" nos propres écritures juste après qu'on les a envoyées
  const pendingLocalIds = useRef(new Set());
  const pendingLocalTeamIds = useRef(new Set());
  const pendingLocalPortefeuilleIds = useRef(new Set());

  useEffect(() => {
    (async () => {
      const [storedClients, storedTeam, storedPortefeuilles] = await Promise.all([
        loadClientsFromSupabase(),
        loadTeamFromSupabase(),
        loadPortefeuillesFromSupabase(),
      ]);
      if (storedClients && storedClients.length) {
        setClients(migrateClients(storedClients));
      } else {
        // Table vide (premier lancement) : on part des données d'origine, à insérer une fois dans Supabase
        setClients(migrateClients(RAW_SEED_CLIENTS));
      }
      setTeam(storedTeam || []);
      setPortefeuilles(storedPortefeuilles || []);
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
        setTimeout(() => setSaveStatus("idle"), 1200);
      });
      return next;
    });
  }, []);

  const addClient = useCallback((newClient) => {
    setClients((prev) => [...prev, newClient]);
    pendingLocalIds.current.add(newClient.id);
    setSaveStatus("saving");
    insertClientRemote(newClient).then(() => {
      setSaveStatus("saved");
      setTimeout(() => setSaveStatus("idle"), 1200);
    });
  }, []);

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
  const navTo = (v) => { setView(v); setActiveClientTab(null); };

  // Équipe "visible" pour les listes déroulantes (assigner un collaborateur/expert/chef
  // de mission à un dossier) : uniquement les comptes actifs de mon portefeuille — l'Admin,
  // qui n'appartient à aucun portefeuille en particulier, voit tout le monde.
  const visibleTeam = (team || []).filter((t) => t.statut !== "en_attente" && (isAdmin || t.portefeuille_id === myPortefeuilleId));
  const myPortefeuille = (portefeuilles || []).find((p) => p.id === myPortefeuilleId) || null;

  return (
    <div style={S.appShell}>
      <GlobalStyle />
      <Sidebar view={view} setView={(v) => navTo(v)} me={me} meRole={myRole} mePortefeuille={myPortefeuille} team={team}
        onLogout={onLogout} counts={{ ...computeCounts(myClients), tachesActives: visibleTasksDb.filter((t) => t.statut !== "termine").length }}
        collapsed={sidebarCollapsed} setCollapsed={setSidebarCollapsed}
        mobileOpen={mobileMenuOpen} setMobileOpen={setMobileMenuOpen} />
      <div style={S.main}>
        <TopBar search={search} setSearch={setSearch} saveStatus={saveStatus} me={me} meColor={meColor}
          openTabs={openClientTabs} activeTab={activeClientTab} onHome={goHome}
          onSelectTab={(id) => setActiveClientTab(id)} onCloseTab={closeClientTab}
          onNav={navTo} onOpenClient={openClientTab} onNewClient={() => setShowAddClient(true)} clients={myClients}
          notifCount={myTasks.filter((t) => t.bucket === "retard").length || undefined}
          onOpenMobileMenu={() => setMobileMenuOpen(true)} />
        <div className="px-3 py-3 md:px-7 md:py-6" style={{ ...S.content, padding: undefined }}>
          {activeClient ? (
            // key={activeClient.id} force le remontage complet du composant à chaque
            // changement d'onglet : les champs non-contrôlés (defaultValue) et l'état
            // interne (onglet secondaire "Infos / TVA / Bilan…") sont ainsi réinitialisés
            // avec les données du dossier sélectionné, au lieu de rester figés sur
            // l'ancien dossier affiché.
            <ClientEditorPage key={activeClient.id} client={activeClient} team={visibleTeam} onUpdate={updateClient}
              onClose={() => closeClientTab(activeClient.id)} />
          ) : (
            <>
              {view === "dashboard" && (
                <Dashboard myClients={myClients} tasks={myTasks} me={me}
                  onOpenClient={(id) => { navTo("clients"); openClientTab(id); }} setView={navTo} />
              )}
              {view === "clients" && (
                <ClientsRegistry clients={myClients} allClients={clients} search={search} setSearch={setSearch} roleFilter={roleFilter} setRoleFilter={setRoleFilter}
                  regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter} me={me}
                  selected={activeClientTab} setSelected={openClientTab} onAdd={() => setShowAddClient(true)}
                  onUpdate={updateClient} onImport={importClients} />
              )}
              {view === "tva" && <TvaGrid clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter}
                regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter} me={me}
                onCycle={(id, mois, val) => { const c = clients.find(x => x.id === id); updateClient(id, { tvaMois: { ...(c.tvaMois || {}), [mois]: val } }); }}
                onUpdate={updateClient} onOpenClient={openClientTab} />}
              {view === "bilans" && <BilansView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "acomptes" && <AcomptesView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "age" && <AgeAgoView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "mission" && <MissionView clients={myClients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter} me={me} onUpdate={updateClient} />}
              {view === "regimes" && <RegimeChangeView clients={myClients} me={me} search={search} onUpdate={updateClient} />}
              {view === "fiscal" && <SuiviFiscalView clients={myClients} team={team} />}
              {view === "mes-taches" && (
                <TasksPage tasks={visibleTasksDb} clients={myClients} team={visibleTeam} me={me} myRow={myRow}
                  onCreate={handleCreateTask} onUpdate={handleUpdateTask} onComplete={handleCompleteTask}
                  onDelete={deleteTask} onOpenClient={openClientTab} />
              )}
            {view === "planning" && (
  <PlanningView tasks={visibleTasksDb} clients={myClients} me={me}
    onUpdate={handleUpdateTask} onOpenClient={openClientTab} />
)}
              {view === "equipe" && (
                <EquipeView team={team} portefeuilles={portefeuilles || []} clients={clients}
                  myRole={myRole} isAdmin={isAdmin} myPortefeuilleId={myPortefeuilleId}
                  canManageTeam={canManageTeam}
                  onAdd={addTeamMember} onRename={renameTeamMember} onDelete={deleteTeamMember}
                  onUpdateMember={updateTeamMember} onAddPortefeuille={addPortefeuille} />
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

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session ?? null));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, sess) => setSession(sess));
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
  if (!session) return <AuthPage />;
  return <CabinetApp session={session} onLogout={() => supabase.auth.signOut()} />;
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
      if (mode === "login") {
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
          <div>
            <label style={authLabelStyle}>Mot de passe</label>
            <div style={{ position: "relative" }}>
              <Lock size={15} style={authIconStyle} />
              <input required type="password" minLength={6} autoComplete={mode === "login" ? "current-password" : "new-password"}
                value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" style={authInputStyle} />
            </div>
          </div>

          {error && <div style={{ fontSize: 12.5, color: T.red, background: T.redSoft, padding: "10px 12px", borderRadius: 10 }}>{error}</div>}
          {info && <div style={{ fontSize: 12.5, color: T.green, background: T.greenSoft, padding: "10px 12px", borderRadius: 10 }}>{info}</div>}

          <button type="submit" disabled={loading} style={{
            marginTop: 4, padding: "13px 0", borderRadius: 12, border: "none", background: T.navy, color: "#fff",
            fontSize: 14, fontWeight: 700, cursor: loading ? "default" : "pointer", opacity: loading ? 0.75 : 1,
            display: "flex", alignItems: "center", justifyContent: "center", gap: 8, boxShadow: "0 10px 24px -8px rgba(79,70,229,0.5)",
          }}>
            {loading && <Loader2 size={15} className="spin" />}
            {mode === "login" ? "Se connecter" : "Créer mon compte"}
          </button>
        </motion.form>

        <div style={{ textAlign: "center", marginTop: 20, fontSize: 12, color: T.inkMuted }}>
          {mode === "login" ? (
            <>Pas encore de compte ? <button type="button" onClick={() => { setMode("signup"); setError(""); setInfo(""); }} style={authLinkStyle}>Inscrivez-vous</button></>
          ) : (
            <>Déjà un compte ? <button type="button" onClick={() => { setMode("login"); setError(""); setInfo(""); }} style={authLinkStyle}>Connectez-vous</button></>
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
        { id: "mission", label: "Dossiers en accueil", icon: ClipboardCheck, badge: counts.missionIncomplete, badgeTone: "amber" },
        { id: "regimes", label: "Changements de régime", icon: RefreshCw },
      ],
    },
    {
      id: "fiscalite", label: "Fiscalité & comptabilité",
      items: [
        { id: "tva", label: "TVA (CA3 / CA12)", icon: Receipt, badge: counts.tvaAlert, badgeTone: "amber" },
        { id: "acomptes", label: "Impôts & cotisations", icon: Landmark },
        { id: "bilans", label: "Bilans", icon: FileWarning, badge: counts.bilanRetard, badgeTone: "red" },
        { id: "fiscal", label: "Suivi fiscal", icon: CalendarDays },
      ],
    },
    {
      id: "juridique", label: "Juridique",
      items: [
        { id: "age", label: "Assemblées (AGE / AGO)", icon: Building2, badge: counts.ageAlert, badgeTone: "amber" },
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
function TopBar({ search, setSearch, saveStatus, me, meColor, openTabs, activeTab, onHome, onSelectTab, onCloseTab, onNav, onOpenClient, onNewClient, clients, notifCount, onOpenMobileMenu }) {
  const [searchOpen, setSearchOpen] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [pickerQuery, setPickerQuery] = useState("");
  const [notifOpen, setNotifOpen] = useState(false);
  const pickerRef = useRef(null);
  const searchInputRef = useRef(null);

  // Raccourci clavier Cmd+K / Ctrl+K : ouvre la recherche rapide depuis n'importe où
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

  // Recherche rapide : Entrée dans le champ de recherche ouvre directement le
  // dossier trouvé (par nom ou SIREN) s'il n'y en a qu'un seul de correspondant.
  const handleSearchKeyDown = (e) => {
    if (e.key === "Escape") { setSearch(""); setSearchOpen(false); e.currentTarget.blur(); return; }
    if (e.key !== "Enter") return;
    const q = search.trim().toLowerCase();
    if (!q) return;
    const matches = clients.filter((c) => c.nom.toLowerCase().includes(q) || (c.siren || "").includes(q));
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
    const list = q ? clients.filter((c) => c.nom.toLowerCase().includes(q) || (c.siren || "").includes(q)) : clients;
    return list.slice(0, 40);
  }, [clients, pickerQuery]);

  // Chaque icône reprend exactement celle utilisée pour le même élément dans le menu latéral
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
  return (
    <div style={{ display: "flex", flexDirection: "column", flexShrink: 0, background: T.card, borderBottom: `1px solid ${T.line}` }}>
      <div style={{ display: "flex", alignItems: "center", padding: "0 10px", height: 46, gap: 2 }}>
        <button onClick={onOpenMobileMenu} title="Menu" className="md:hidden flex items-center justify-center w-8 h-8 rounded-lg text-inksoft hover:bg-app mr-1 shrink-0">
          <Menu size={18} strokeWidth={2} />
        </button>
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

        {searchOpen && (
          <div className="relative ml-2.5 flex-[0_1_280px] hidden sm:block">
            <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-inkmuted" />
            <input ref={searchInputRef} autoFocus value={search} onChange={(e) => setSearch(e.target.value)} onKeyDown={handleSearchKeyDown}
              onBlur={() => !search && setSearchOpen(false)}
              placeholder="Rechercher un dossier, un SIREN…"
              className="input-field !rounded-full !py-1.5 !pl-8 !pr-16 !bg-app text-xs w-full" />
            <kbd className="absolute right-2 top-1/2 -translate-y-1/2 text-[9.5px] font-mono text-inkmuted bg-white border border-line rounded px-1.5 py-0.5">Échap</kbd>
          </div>
        )}

        <div className="ml-auto flex items-center gap-0.5">
          <div style={{ display: "flex", alignItems: "center", gap: 4, fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted, marginRight: 8 }}>
            {saveStatus === "saving" && <><Loader2 size={12} className="spin" /> enreg.…</>}
            {saveStatus === "saved" && <><Check size={12} color={T.green} /> enregistré</>}
          </div>
          {!searchOpen && (
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
              {!!notifCount && <span className="absolute top-1 right-1 bg-badge-red-text text-white text-[9px] font-bold rounded-full px-[4px] leading-[13px]">{notifCount}</span>}
            </button>
            {notifOpen && (
              <div className="absolute right-0 top-9 w-72 card p-3 z-30">
                <div className="text-xs font-bold text-ink mb-2">Notifications</div>
                {notifCount ? (
                  <div onClick={() => { onNav("planning"); setNotifOpen(false); }} className="hoverRow clickable text-xs text-inksoft rounded-lg p-2 cursor-pointer">
                    {notifCount} échéance{notifCount > 1 ? "s" : ""} en retard sur votre planning
                  </div>
                ) : (
                  <div className="text-xs text-inkmuted italic px-2 py-1">Aucune notification pour le moment.</div>
                )}
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
function missionCompletion(client) {
  const m = client.mission; if (!m) return null;
  const vals = Object.values(m); if (!vals.length) return null;
  const done = vals.filter(Boolean).length;
  return { done, total: vals.length, pct: Math.round((done / vals.length) * 100) };
}
function isBilanLate(client) { return !!(client.bilan && client.bilan.nonFinalise); }
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
  if (statutFilter === "actif") out = out.filter((c) => (c.statutDossier || "actif") !== "inactif");
  else if (statutFilter === "inactif") out = out.filter((c) => c.statutDossier === "inactif");
  // statutFilter === "tous" -> pas de filtre supplémentaire
  if (regimeFilter && regimeFilter !== "Tous") out = out.filter((c) => c.tvaRegime === regimeFilter);
  if (search.trim()) {
    const q = search.trim().toLowerCase();
    out = out.filter((c) => c.nom.toLowerCase().includes(q) || (c.siren || "").includes(q));
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
   DASHBOARD
   ============================================================ */
function Dashboard({ myClients, tasks, me, onOpenClient, setView }) {
  const counts = computeCounts(myClients);
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

  const today = new Date();
  const dateStr = today.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" });

  return (
    <div>
      <Reveal>
        <div style={{ marginBottom: 30 }}>
          <div style={{ fontFamily: T.mono, fontSize: 11, letterSpacing: "0.1em", color: T.inkMuted, textTransform: "uppercase" }}>{dateStr}</div>
          <h1 style={{ fontFamily: T.serif, fontSize: 19, fontWeight: 700, color: T.ink, margin: "6px 0 0" }}>Bonjour {me}</h1>
        </div>
      </Reveal>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 md:gap-4" style={{ marginBottom: 24 }}>
        <KpiCard index={0} label="Mes dossiers" value={counts.total} icon={Users} onClick={() => setView("clients")} />
        <KpiCard index={1} label="TVA en retard ce mois" value={counts.tvaAlert} icon={Receipt} tone={counts.tvaAlert ? "amber" : "green"} onClick={() => setView("tva")} />
        <KpiCard index={2} label="Bilans en retard" value={counts.bilanRetard} icon={FileWarning} tone={counts.bilanRetard ? "red" : "green"} onClick={() => setView("bilans")} />
        <KpiCard index={3} label="Accueils incomplets" value={counts.missionIncomplete} icon={ClipboardCheck} tone={counts.missionIncomplete ? "amber" : "green"} onClick={() => setView("mission")} />
      </div>

      <div className="grid grid-cols-1 md:grid-cols-[1.3fr_1fr] gap-4 md:gap-[18px]">
        <Panel index={4} title="Mes tâches">
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
            {["Toutes", "retard", "aujourdhui", "demain", "semaine", "mois", "trimestre"].map((b) => (
              <button key={b} onClick={() => setTaskFilter(b)} style={{
                padding: "6px 12px", borderRadius: 999, fontSize: 11.5, fontWeight: 600,
                border: `1px solid ${taskFilter === b ? T.navy : T.line}`, background: taskFilter === b ? T.navySoft : T.card,
                color: taskFilter === b ? T.navy : T.inkSoft, cursor: "pointer",
              }}>{b === "Toutes" ? "Toutes" : BUCKET_LABELS[b]}</button>
            ))}
          </div>
          {sortedTasks.length === 0 ? <EmptyNote text="Rien à signaler sur cette période. Le registre est à jour." /> : (
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {sortedTasks.slice(0, 10).map((t, i) => (
                <Reveal key={t.id} index={i} delay={0.1}>
                  <div className="hoverRow clickable" onClick={() => onOpenClient(t.client.id)} style={{ display: "flex", alignItems: "center", gap: 12, padding: "13px 14px", borderRadius: T.radiusSm, border: `1px solid ${T.line}`, background: T.paper }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 12.5, color: T.ink }}>{t.client.nom}</div>
                      <div style={{ fontSize: 11.5, color: T.inkMuted }}>{t.label}</div>
                    </div>
                    <span style={{ fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted }}>{BUCKET_LABELS[t.bucket]}</span>
                    <Stamped tone={t.tone} small>{t.category}</Stamped>
                    <ChevronRight size={15} color={T.inkMuted} />
                  </div>
                </Reveal>
              ))}
            </div>
          )}
        </Panel>

        <Panel index={5} title="Mes dossiers par rôle">
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

      <div style={{ marginTop: 18 }}>
        <Panel index={6} title="Échéance TVA du mois — vue d'ensemble">
          <MiniTvaOverview clients={myClients} setView={setView} />
        </Panel>
      </div>
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
function KpiCard({ label, value, icon: Icon, tone, onClick, index = 0 }) {
  const toneColor = tone === "red" ? T.red : tone === "amber" ? T.amber : tone === "green" ? T.green : T.navy;
  const toneSoft = tone === "red" ? T.redSoft : tone === "amber" ? T.amberSoft : tone === "green" ? T.greenSoft : T.navySoft;
  return (
    <Reveal index={index}>
      <div onClick={onClick} className={onClick ? "clickable" : ""} style={{ background: T.card, border: `1px solid ${T.line}`, borderRadius: T.radiusLg, boxShadow: T.shadowSm, padding: "22px 24px" }}>
        <div style={{ marginBottom: 14, width: 38, height: 38, borderRadius: 12, background: toneSoft, display: "flex", alignItems: "center", justifyContent: "center" }}><Icon size={17} color={toneColor} /></div>
        <div style={{ fontFamily: T.serif, fontSize: 18, fontWeight: 700, color: T.ink, lineHeight: 1 }}>{value}</div>
        <div style={{ fontSize: 12.5, color: T.inkMuted, marginTop: 8, fontWeight: 500 }}>{label}</div>
      </div>
    </Reveal>
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
function ClientsRegistry({ clients, allClients, search, setSearch, roleFilter, setRoleFilter, regimeFilter, setRegimeFilter, me, selected, setSelected, onAdd, onUpdate, onImport }) {
  const [statutFilter, setStatutFilter] = useState("actif");
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me, regimeFilter, statutFilter), [clients, search, roleFilter, me, regimeFilter, statutFilter]);
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
            <button onClick={() => exportClientsToExcel(allClients || clients)} className="btn-secondary !py-2">
              <ArrowUpRight size={14} className="rotate-90" /> <span className="hidden sm:inline">Exporter la liste (Excel)</span>
            </button>
            <button onClick={onAdd} className="btn-primary !py-2">
              <Plus size={15} /> <span className="hidden sm:inline">Nouveau client</span>
            </button>
          </div>
        </div>
        {importMsg && (
          <div className={`mt-2 text-xs font-semibold px-2.5 py-1.5 rounded-lg inline-block ${importMsg.tone === "green" ? "bg-badge-green-bg text-badge-green-text" : importMsg.tone === "red" ? "bg-badge-red-bg text-badge-red-text" : "bg-badge-amber-bg text-badge-amber-text"}`}>{importMsg.text}</div>
        )}
      </Reveal>
      <p className="text-inkmuted text-xs mt-1.5 mb-5">Cliquez un dossier pour ouvrir sa fiche complète.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter}
        statutFilter={statutFilter} setStatutFilter={setStatutFilter} search={search} setSearch={setSearch} />

      {/* En-tête colonnes : visible à partir de md, masqué sur mobile (les dossiers s'affichent en cartes empilées) */}
      <div className="hidden md:grid gap-0" style={{ gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1.3fr 40px", padding: "0 18px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: T.inkMuted, fontWeight: 600, marginBottom: 10 }}>
        <div>Dossier</div><div>SIREN</div><div>Rôles</div><div>Clôture</div><div>Régime</div><div>Statuts</div><div />
      </div>
      <div className="flex flex-col gap-2">
        {Object.keys(grouped).sort().map((letter) => (
          <div key={letter}>
            <div className="px-1.5 py-1 font-mono text-[10.5px] font-bold text-accent-deep tracking-widest">{letter}</div>
            <div className="flex flex-col gap-2">
              {grouped[letter].map((c) => {
                rowIndex += 1;
                const statusBadge = isBilanLate(c) ? <Stamped tone="red" small>Bilan retard</Stamped>
                  : isTvaLate(c) ? <Stamped tone="amber" small>TVA</Stamped>
                  : <Stamped tone="green" small>À jour</Stamped>;
                const roles = [c.collab === me && "Collaborateur", c.expert === me && "Expert", c.chefMission === me && "Chef de mission"].filter(Boolean);
                return (
                  <Reveal key={c.id} index={rowIndex}>
                    {/* Ligne tableau (md et +) */}
                    <div onClick={() => setSelected(c.id)}
                      className={`hoverRow clickable hidden md:grid items-center rounded-xl border px-4 py-3.5 text-xs ${selected === c.id ? "border-accent-deep bg-accent-soft" : "border-line bg-card shadow-xs"} ${c.statutDossier === "inactif" ? "opacity-55" : ""}`}
                      style={{ gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1.3fr 40px" }}>
                      <div className="font-semibold text-ink flex items-center gap-2">
                        {c.nom}
                        {c.statutDossier === "inactif" && <Stamped tone="neutral" small>Inactif</Stamped>}
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
                      <div className="flex gap-1.5 flex-wrap">{statusBadge}</div>
                      <ChevronRight size={15} className="text-inkmuted" />
                    </div>
                    {/* Carte empilée (mobile) */}
                    <div onClick={() => setSelected(c.id)}
                      className={`hoverRow clickable md:hidden rounded-xl border p-3.5 flex flex-col gap-2 ${selected === c.id ? "border-accent-deep bg-accent-soft" : "border-line bg-card shadow-xs"} ${c.statutDossier === "inactif" ? "opacity-55" : ""}`}>
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-semibold text-ink text-sm">{c.nom}</span>
                        <ChevronRight size={15} className="text-inkmuted shrink-0" />
                      </div>
                      <div className="flex items-center gap-2 flex-wrap text-[11px] text-inkmuted">
                        <span className="font-mono">{c.siren || "—"}</span>
                        {roles.length > 0 && <span>· {roles.join(", ")}</span>}
                        {c.tvaRegime && <span className="font-mono">· {c.tvaRegime}</span>}
                      </div>
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {statusBadge}
                        {c.statutDossier === "inactif" && <Stamped tone="neutral" small>Inactif</Stamped>}
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
function ClientEditorPage({ client, team, onUpdate, onClose }) {
  const [tab, setTab] = useState("infos");
  if (!client) return null;
  const tabs = [
    { id: "infos", label: "Infos générales" }, { id: "corporate", label: "Corporate" }, { id: "tva", label: "TVA" },
    { id: "bilan", label: "Bilan" }, { id: "acomptes", label: "Acomptes" }, { id: "age", label: "AGE / AGO" }, { id: "mission", label: "Accueil" },
  ];
  return (
    <div>
      <Reveal>
        <div className="flex items-start justify-between gap-3 flex-wrap" style={{ marginBottom: 4 }}>
          <div className="min-w-0">
            <div style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{client.siren || "SIREN non renseigné"}</div>
            <input defaultValue={client.nom} onBlur={(e) => onUpdate(client.id, { nom: e.target.value || client.nom })}
              className="w-full sm:min-w-[260px] sm:w-auto"
              style={{ fontFamily: T.serif, fontSize: 16, fontWeight: 700, color: T.ink, border: "none", background: "transparent", padding: "2px 0", margin: "2px 0 6px" }} />
            <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap", fontSize: 12 }}>
              <RoleBadge role="Collab." name={client.collab} />
              <RoleBadge role="Expert" name={client.expert} />
              <RoleBadge role="Chef de mission" name={client.chefMission} />
              {client.tvaRegime && <span style={{ fontFamily: T.mono, fontSize: 11, color: T.navy, fontWeight: 700, background: T.navySoft, padding: "2px 9px", borderRadius: 999 }}>{client.tvaRegime}</span>}
              <button
                onClick={() => onUpdate(client.id, { statutDossier: client.statutDossier === "inactif" ? "actif" : "inactif" })}
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
            </div>
          </div>
          <button onClick={onClose} style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: `1px solid ${T.line}`, borderRadius: 9, padding: "7px 12px", cursor: "pointer", color: T.inkMuted, fontSize: 12 }}>
            <X size={14} /> Fermer l'onglet
          </button>
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
        {tab === "infos" && <InfosTab client={client} team={team} onUpdate={onUpdate} />}
        {tab === "corporate" && <CorporateTab client={client} onUpdate={onUpdate} />}
        {tab === "tva" && <TvaTab client={client} onUpdate={onUpdate} />}
        {tab === "bilan" && <BilanTab client={client} onUpdate={onUpdate} />}
        {tab === "acomptes" && <AcomptesTab client={client} onUpdate={onUpdate} />}
        {tab === "age" && <AgeAgoEditor client={client} onUpdate={onUpdate} />}
        {tab === "mission" && <MissionTab client={client} onUpdate={onUpdate} />}
      </div>
    </div>
  );
}

function FieldRow({ label, children }) {
  return <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderBottom: `1px solid ${T.line}`, gap: 10 }}>
    <span style={{ fontSize: 12, color: T.inkMuted }}>{label}</span><div>{children}</div>
  </div>;
}
function SelectPill({ value, options, onChange, allowEmpty = true }) {
  return (
    <select value={value || ""} onChange={(e) => onChange(e.target.value)} style={{ fontFamily: T.mono, fontSize: 12, fontWeight: 600, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card, color: T.ink }}>
      {allowEmpty && <option value="">—</option>}
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}
function TextInput({ defaultValue, onCommit, placeholder, width = 160, align = "right" }) {
  return <input defaultValue={defaultValue || ""} placeholder={placeholder} onBlur={(e) => onCommit(e.target.value)}
    style={{ fontFamily: T.sans, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, width, textAlign: align, background: T.card }} />;
}

function InfosTab({ client, team, onUpdate }) {
  const teamNames = team.map((t) => t.nom);
  return (
    <div>
      <FieldRow label="SIREN"><TextInput defaultValue={client.siren} onCommit={(v) => onUpdate(client.id, { siren: v })} width={140} /></FieldRow>
      <FieldRow label="Logiciel"><SelectPill value={client.logiciel} options={["MYUNISOFT", "QUADRA"]} onChange={(v) => onUpdate(client.id, { logiciel: v })} /></FieldRow>
      <FieldRow label="Forme juridique"><SelectPill value={client.formeJuridique} options={["SARL", "EURL", "SAS", "SASU", "SCI", "SA", "SNC", "EI", "Association"]} onChange={(v) => onUpdate(client.id, { formeJuridique: v })} /></FieldRow>
      <FieldRow label="Capital social"><TextInput defaultValue={client.capital} onCommit={(v) => onUpdate(client.id, { capital: v })} placeholder="ex. 5 000 €" width={140} /></FieldRow>
      <FieldRow label="Activité"><TextInput defaultValue={client.activite} onCommit={(v) => onUpdate(client.id, { activite: v })} placeholder="ex. Restauration" width={200} align="left" /></FieldRow>
      <FieldRow label="Date de clôture d'exercice"><input type="date" defaultValue={client.dateCloture || ""} onChange={(e) => onUpdate(client.id, { dateCloture: e.target.value })} style={{ fontFamily: T.mono, fontSize: 12.5, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, background: T.card }} /></FieldRow>
      <div style={{ height: 6 }} />
      <FieldRow label="Collaborateur"><SelectPill value={client.collab} options={teamNames} onChange={(v) => onUpdate(client.id, { collab: v })} /></FieldRow>
      <FieldRow label="Expert"><SelectPill value={client.expert} options={teamNames} onChange={(v) => onUpdate(client.id, { expert: v })} /></FieldRow>
      <FieldRow label="Chef de mission"><SelectPill value={client.chefMission} options={teamNames} onChange={(v) => onUpdate(client.id, { chefMission: v })} /></FieldRow>
      <FieldRow label="Régime TVA"><SelectPill value={client.tvaRegime} options={REGIMES_TVA} onChange={(v) => onUpdate(client.id, { tvaRegime: v })} /></FieldRow>
    </div>
  );
}

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
  return (
    <div>
      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "4px 0 8px", display: "flex", alignItems: "center", gap: 6 }}>
        <ShieldCheck size={15} /> Base corporate / LAB
      </h4>
      <FieldRow label="Questionnaire LAB complété"><ToggleBtn on={!!kyc.lab} onClick={() => patchKyc({ lab: !kyc.lab })} /></FieldRow>
      <FieldRow label="Mandat signé"><ToggleBtn on={!!kyc.mandat} onClick={() => patchKyc({ mandat: !kyc.mandat })} /></FieldRow>
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
  const currentStatus = client.tvaRegime === "CA12" ? effectiveTvaStatus(client, "Mai") : effectiveTvaStatus(client, currentMonthKey());
  return (
    <div>
      <FieldRow label="Régime TVA"><SelectPill value={client.tvaRegime} options={REGIMES_TVA} onChange={(v) => onUpdate(client.id, { tvaRegime: v })} /></FieldRow>
      <FieldRow label="Jour limite de déclaration">
        <input type="number" min="1" max="31" defaultValue={client.tvaExig || ""} placeholder="ex. 19"
          onBlur={(e) => onUpdate(client.id, { tvaExig: e.target.value ? parseInt(e.target.value, 10) : "" })}
          style={{ fontFamily: T.mono, fontSize: 12, padding: "5px 8px", borderRadius: 9, border: `1px solid ${T.line}`, width: 60, textAlign: "right" }} />
      </FieldRow>
      <FieldRow label="Statut courant">
        <Stamped tone={tvaTone(currentStatus)} small>{currentStatus === "RETARD" ? "Retard" : currentStatus === "FAIT" ? "Fait" : currentStatus === "OK" ? "OK" : currentStatus === "NA" ? "N/A" : "—"}</Stamped>
      </FieldRow>
      <div style={{ fontSize: 12, color: T.inkMuted, margin: "14px 0 0", lineHeight: 1.6 }}>
        {client.tvaRegime === "CA12"
          ? "Régime CA12 : une seule déclaration annuelle, exigible en Mai N+1."
          : client.tvaRegime === "CA3"
            ? "Régime CA3 : la TVA d'un mois donné est déclarée le mois suivant (M+1)."
            : "Sélectionnez un régime TVA pour activer le suivi des échéances."}
        {" "}Le suivi mois par mois (Fait / OK / N/A) se gère depuis l'écran <strong>TVA — CA3/CA12</strong>.
      </div>
    </div>
  );
}

function BilanTab({ client, onUpdate }) {
  const b = client.bilan || {};
  const toggle = (field) => onUpdate(client.id, { bilan: { ...b, [field]: !b[field] } });
  return (
    <div>
      <FieldRow label="Finalisé après échéance"><ToggleBtn on={!!b.finaliseApres} onClick={() => toggle("finaliseApres")} /></FieldRow>
      <FieldRow label="Non encore finalisé (en retard)"><ToggleBtn on={!!b.nonFinalise} onClick={() => toggle("nonFinalise")} tone="red" /></FieldRow>
      <FieldRow label="Courrier de retard signé et classé"><ToggleBtn on={!!b.courrier} onClick={() => toggle("courrier")} /></FieldRow>
    </div>
  );
}
function ToggleBtn({ on, onClick, tone = "green" }) {
  return <button onClick={onClick} style={{ display: "flex", alignItems: "center", gap: 6, border: "none", cursor: "pointer", background: "none", padding: 0 }}>
    <Stamped tone={on ? tone : "neutral"} small>{on ? "Oui" : "Non"}</Stamped>
  </button>;
}

function AcomptesTab({ client, onUpdate }) {
  const is = client.is || {}; const cfe = client.cfe || {};
  const toggleIs = (f) => onUpdate(client.id, { is: { ...is, [f]: !is[f] } });
  const toggleCfe = (f) => onUpdate(client.id, { cfe: { ...cfe, [f]: !cfe[f] } });
  return (
    <div>
      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "4px 0 8px" }}>Impôt sur les sociétés</h4>
      <FieldRow label="Concerné (IS N-1 > 3000€)"><ToggleBtn on={!!is.concerne} onClick={() => toggleIs("concerne")} /></FieldRow>
      <FieldRow label="Acompte mars"><ToggleBtn on={!!is.mars} onClick={() => toggleIs("mars")} /></FieldRow>
      <FieldRow label="Acompte juin"><ToggleBtn on={!!is.juin} onClick={() => toggleIs("juin")} /></FieldRow>
      <FieldRow label="Acompte septembre"><ToggleBtn on={!!is.sept} onClick={() => toggleIs("sept")} /></FieldRow>
      <FieldRow label="Acompte décembre"><ToggleBtn on={!!is.dec} onClick={() => toggleIs("dec")} /></FieldRow>
      <FieldRow label="Solde IS"><ToggleBtn on={!!is.solde} onClick={() => toggleIs("solde")} /></FieldRow>
      <h4 style={{ fontFamily: T.serif, fontSize: 13, color: T.navy, margin: "20px 0 8px" }}>CFE</h4>
      <FieldRow label="Concerné (CFE N-1 > 3000€)"><ToggleBtn on={!!cfe.concerne} onClick={() => toggleCfe("concerne")} /></FieldRow>
      <FieldRow label="Acompte juin"><ToggleBtn on={!!cfe.juin} onClick={() => toggleCfe("juin")} /></FieldRow>
      <FieldRow label="Solde décembre"><ToggleBtn on={!!cfe.dec} onClick={() => toggleCfe("dec")} /></FieldRow>
    </div>
  );
}

function MissionTab({ client, onUpdate }) {
  const m = client.mission || {};
  const keys = Object.keys(m).length ? Object.keys(m) : ["Acceptation mission", "Lettre reprise", "LAB", "Fiche client", "Lettre mission", "Mandat", "Attestation bilan"];
  const toggle = (k) => onUpdate(client.id, { mission: { ...m, [k]: !m[k] } });
  const done = keys.filter((k) => m[k]).length;
  return (
    <div>
      <div style={{ marginBottom: 14 }}>
        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 6 }}>
          <span style={{ color: T.inkMuted }}>Progression de l'accueil</span><span style={{ fontFamily: T.mono, fontWeight: 600 }}>{done}/{keys.length}</span>
        </div>
        <div style={{ height: 8, borderRadius: 4, background: T.paperDeep, overflow: "hidden" }}><div style={{ width: `${(done / keys.length) * 100}%`, height: "100%", background: T.navy }} /></div>
      </div>
      {keys.map((k) => (
        <div key={k} onClick={() => toggle(k)} className="clickable" style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 4px", borderBottom: `1px solid ${T.line}` }}>
          <span style={{ width: 19, height: 19, borderRadius: 5, border: `1.5px solid ${m[k] ? T.green : T.line}`, background: m[k] ? T.green : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            {m[k] && <Check size={13} color="#fff" strokeWidth={3} />}
          </span>
          <span style={{ fontSize: 12.5, color: m[k] ? T.inkMuted : T.ink, textDecoration: m[k] ? "line-through" : "none" }}>{k}</span>
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   TVA GRID VIEW
   ============================================================ */
function TvaGrid({ clients, search, roleFilter, setRoleFilter, regimeFilter, setRegimeFilter, me, onCycle, onUpdate, onOpenClient }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me, regimeFilter).filter(c => c.tvaRegime), [clients, search, roleFilter, me, regimeFilter]);
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 15, fontWeight: 600, color: T.ink, margin: "0 0 5px" }}>Échéances TVA</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 11, marginTop: 0, marginBottom: 16, lineHeight: 1.5 }}>
        Cliquez une cellule : vide → <Stamped tone="amber" small>Fait</Stamped> → <Stamped tone="green" small>OK</Stamped> → <Stamped tone="neutral" small>N/A</Stamped>. Date limite dépassée sans saisie → <Stamped tone="red" small>Retard</Stamped> automatique.
        {" "}CA3 : déclaration du mois M exigible en M+1. CA12 : une seule déclaration, en Mai N+1.
      </p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter} />
      <div className="scrollbar" style={{ overflowX: "auto", background: T.card, border: `1px solid ${T.line}`, borderRadius: 16, boxShadow: T.shadowSm }}>
        <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 11.5 }}>
          <thead><tr><th style={thStyle}>Dossier</th><th style={thStyle}>Régime</th><th style={{ ...thStyle, textAlign: "center" }}>Éxig.</th>{MOIS_ORDER.map((m) => <th key={m} style={{ ...thStyle, textAlign: "center" }}>{m}</th>)}</tr></thead>
          <tbody>
            {filtered.map((c) => {
              const isCa12 = c.tvaRegime === "CA12";
              return (
              <tr key={c.id} className="hoverRow">
                <td className={onOpenClient ? "clickable" : undefined} onClick={() => onOpenClient && onOpenClient(c.id)}
                  style={{ ...tdStyle, fontWeight: 600, whiteSpace: "nowrap", color: onOpenClient ? T.navy : T.ink }}>{c.nom}</td>
                <td style={{ ...tdStyle, fontFamily: T.mono, color: T.inkMuted }}>{c.tvaRegime}</td>
                <td style={{ ...tdStyle, textAlign: "center" }}>
                  <input type="number" min="1" max="31" defaultValue={c.tvaExig || ""} placeholder="—"
                    onBlur={(e) => onUpdate(c.id, { tvaExig: e.target.value ? parseInt(e.target.value, 10) : "" })}
                    style={{ width: 42, textAlign: "center", fontFamily: T.mono, fontSize: 12, padding: "4px 2px", borderRadius: 5, border: `1px solid ${T.line}`, background: T.paper }} />
                </td>
                {MOIS_ORDER.map((m) => {
                  if (isCa12 && m !== "Mai") {
                    return <td key={m} style={{ ...tdStyle, textAlign: "center", color: T.inkMuted, opacity: 0.45 }}>—</td>;
                  }
                  const manual = (c.tvaMois?.[m] || "").toUpperCase(); const display = effectiveTvaStatus(c, m); const tone = tvaTone(display);
                  return (
                    <td key={m} style={{ ...tdStyle, textAlign: "center" }}>
                      <button className="clickable" onClick={() => onCycle(c.id, m, manual === "" ? "FAIT" : manual === "FAIT" ? "OK" : manual === "OK" ? "NA" : "")} style={{ background: "none", border: "none", padding: 0 }}>
                        <Stamped tone={tone} small>{display === "RETARD" ? "Retard" : display === "FAIT" ? "Fait" : display === "OK" ? "OK" : display === "NA" ? "N/A" : "·"}</Stamped>
                      </button>
                    </td>
                  );
                })}
              </tr>
            );})}
          </tbody>
        </table>
        {filtered.length === 0 && <EmptyNote text="Aucun dossier soumis à la TVA dans cette sélection." />}
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
            <label style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: T.red }}><input type="checkbox" checked={!!b.nonFinalise} onChange={() => onUpdate(c.id, { bilan: { ...b, nonFinalise: !b.nonFinalise } })} /> en retard</label>
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
  const isConcerned = filtered.filter((c) => c.is?.concerne);
  const cfeConcerned = filtered.filter((c) => c.cfe?.concerne);
  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Acomptes IS &amp; CFE</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Dossiers dont l'impôt N-1 dépasse 3 000 €.</p>
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
                {Object.entries(c.mission || {}).map(([k, v]) => (
                  <button key={k} onClick={() => onUpdate(c.id, { mission: { ...c.mission, [k]: !v } })} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
                    <Stamped tone={v ? "green" : "neutral"} small>{k}</Stamped>
                  </button>
                ))}
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
        <button onClick={() => onComplete(task)} title="Marquer terminé"
          className="w-5 h-5 rounded-full border-[1.5px] border-badge-green-text flex items-center justify-center shrink-0 hover:bg-badge-green-bg transition-colors">
          <Check size={12} className="text-badge-green-text" />
        </button>
        <div className="flex-1 min-w-[140px] sm:min-w-0">
          <div className={`font-semibold text-xs text-ink inline-block ${client ? "cursor-pointer hover:text-accent" : ""}`}
            onClick={() => client && onOpenClient(client.id)}>
            {client ? client.nom : "Dossier non lié"}
          </div>
          <div className="text-[11.5px] text-inkmuted">{task.nom}{task.commentaire ? ` — ${task.commentaire}` : ""}</div>
        </div>
        <button onClick={() => onDelete(task.id)} title="Supprimer" className="text-inkmuted hover:text-badge-red-text transition-colors order-2 sm:order-none">
          <Trash2 size={13} />
        </button>
        <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto pl-[30px] sm:pl-0 order-3 sm:order-none">
          {responsable && <RoleBadge role="Resp." name={responsable.nom} />}
          <select value={task.statut} onChange={(e) => onUpdate(task.id, { statut: e.target.value })}
            className="input-field !w-auto !py-1 !px-2 text-[10.5px] font-bold cursor-pointer">
            {TASK_STATUTS.map((s) => <option key={s.code} value={s.code}>{s.label}</option>)}
          </select>
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
const PLANNING_FILTERS = [
  { id: "toutes", label: "Toutes" },
  { id: "retard", label: "En retard" },
  { id: "aujourdhui", label: "Aujourd'hui" },
  { id: "semaine", label: "Cette semaine" },
];

function PlanningTaskCard({ task, client, draggable = true }) {
  const tone = TASK_PRIORITE_TONE[task.priorite] || "neutral";
  return (
    <div
      draggable={draggable}
      onDragStart={(e) => e.dataTransfer.setData("text/plain", JSON.stringify({ type: "task", id: task.id }))}
      style={{
        background: T.card, border: `1px solid ${T.line}`, borderLeft: `4px solid ${tone === "red" ? T.red : tone === "amber" ? T.amber : T.navy}`,
        borderRadius: 9, padding: "9px 10px", marginBottom: 7, cursor: draggable ? "grab" : "default", boxShadow: T.shadowSm,
      }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6 }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: T.ink, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
          {client ? client.nom : "Sans dossier"}
        </span>
        <Stamped tone={tone} small>{TASK_PRIORITE_BY_CODE[task.priorite]?.label}</Stamped>
      </div>
      <div style={{ fontSize: 11, color: T.inkMuted, marginTop: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{task.nom}</div>
    </div>
  );
}

function PlanningView({ tasks, clients, me, onUpdate, onOpenClient }) {
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
              <PlanningTaskCard key={t.id} task={t} client={clientById[t.client_id]} />
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
                            onClick={() => client && onOpenClient(client.id)}
                            onDoubleClick={(e) => { e.stopPropagation(); unschedule(t); }}
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
      tvaRegime: "", tvaExig: "", tvaMois: {}, regimeHistory: [], ageAgoHistory: {},
      corporate: { kyc: { lab: false, mandat: false, choixPA: "", beneficiaireEffectif: false, beneficiaireNom: "" }, kycExtra: [], notes: "" },
      mission: { "Acceptation mission": false, "Lettre reprise": false, "LAB": false, "Fiche client": false, "Lettre mission": false, "Mandat": false, "Attestation bilan": false },
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
