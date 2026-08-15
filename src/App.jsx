import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  LayoutGrid, Users, Receipt, FileWarning, Landmark, Building2,
  ClipboardCheck, Search, ChevronRight, X, Check, AlertTriangle,
  Clock, TrendingUp, UserCircle2, Plus, Stamp, ChevronDown,
  Filter, ArrowUpRight, CircleDot, Loader2, RefreshCw, History,
  ChevronUp, CalendarDays, CalendarRange, Settings2, Trash2,
  Pencil, ChevronLeft, ShieldCheck, Home, LogOut, Mail, Lock, UserRound
} from "lucide-react";
import { supabase } from "./supabaseClient";
import { motion } from "framer-motion";
import * as XLSX from "xlsx";

const T = {
  paper: "#F3F4F6", paperDeep: "#EEF2FF", ink: "#0F172A", inkSoft: "#475569", inkMuted: "#94A3B8",
  line: "#E2E8F0", card: "#FFFFFF", gold: "#D97706", goldSoft: "#FEF3C7",
  green: "#16A34A", greenSoft: "#DCFCE7", red: "#DC2626", redSoft: "#FEE2E2",
  amber: "#D97706", amberSoft: "#FEF3C7",
  /* accent sobre façon Kabineo (indigo) */
  navy: "#4F46E5", navySoft: "#EEF2FF",
  /* sidebar claire façon Kabineo : blanc pur + bordure très légère */
  sidebarBg: "#FFFFFF", sidebarBg2: "#F1F5F9", sidebarInk: "#334155", sidebarInkMuted: "#94A3B8", sidebarActive: "#EEF2FF",
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
/* ---- Clients : Supabase (table "clients", colonnes id text + data jsonb) ---- */
async function loadClientsFromSupabase() {
  const { data, error } = await supabase.from("clients").select("id, data");
  if (error) { console.error("Erreur chargement clients :", error.message); return null; }
  if (!data) return null;
  return data.map((row) => ({ id: row.id, ...(row.data || {}) }));
}
async function insertClientRemote(client) {
  const { id, ...rest } = client;
  const { error } = await supabase.from("clients").insert({ id, data: rest });
  if (error) console.error("Erreur création client :", error.message);
}
async function updateClientRemote(id, fullClient) {
  const { id: _drop, ...rest } = fullClient;
  const { error } = await supabase.from("clients").update({ data: rest }).eq("id", id);
  if (error) console.error("Erreur sauvegarde client :", error.message);
}

/* ---- Équipe : Supabase (table "team", colonnes id text + nom text + color text) ----
   Avant : localStorage (team-v1), donc invisible pour les autres postes.
   Le roster (liste des collaborateurs et leurs couleurs) est maintenant une donnée
   partagée comme les clients, avec le même mécanisme de synchronisation temps réel. */
async function loadTeamFromSupabase() {
  const { data, error } = await supabase.from("team").select("id, nom, color").order("nom", { ascending: true });
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

/* ---- Identité locale : quel collaborateur utilise CE poste. Volontairement gardée
   en localStorage : ce n'est pas une donnée métier à partager, juste une préférence
   d'appareil (comme "se souvenir de moi"), donc pas de table Supabase nécessaire. ---- */
async function loadMe() {
  try { return localStorage.getItem("me-v1") || null; } catch (e) {}
  return null;
}
async function saveMe(name) {
  try { localStorage.setItem("me-v1", name); } catch (e) {}
}

/* ============================================================
   APP
   ============================================================ */
function CabinetApp({ session, onLogout }) {
  const [clients, setClients] = useState(null);
  const [team, setTeam] = useState(null);
  const [loading, setLoading] = useState(true);
  const [me, setMe] = useState(null);
  const [view, setView] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("Tous");
  const [regimeFilter, setRegimeFilter] = useState("Tous");
  const [showAddClient, setShowAddClient] = useState(false);
  const [saveStatus, setSaveStatus] = useState("idle");
  const [openClientTabs, setOpenClientTabs] = useState([]); // [{id, label}]
  const [activeClientTab, setActiveClientTab] = useState(null); // id du dossier affiché en page pleine, ou null
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

  // Empêche le canal temps réel de "rejouer" nos propres écritures juste après qu'on les a envoyées
  const pendingLocalIds = useRef(new Set());
  const pendingLocalTeamIds = useRef(new Set());

  useEffect(() => {
    (async () => {
      const [storedClients, storedTeam, meName] = await Promise.all([
        loadClientsFromSupabase(),
        loadTeamFromSupabase(),
        loadMe(),
      ]);
      if (storedClients && storedClients.length) {
        setClients(migrateClients(storedClients));
      } else {
        // Table vide (premier lancement) : on part des données d'origine, à insérer une fois dans Supabase
        setClients(migrateClients(RAW_SEED_CLIENTS));
      }
      if (storedTeam && storedTeam.length) {
        setTeam(storedTeam);
      } else {
        // Table "team" vide (premier lancement) : on insère le roster par défaut dans Supabase
        // une bonne fois pour toutes, pour que tous les postes démarrent avec la même liste.
        setTeam(DEFAULT_TEAM);
        Promise.all(DEFAULT_TEAM.map((m) => insertTeamMemberRemote(m)));
      }
      setMe(meName);
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
            const incoming = { id: payload.new.id, ...(payload.new.data || {}) };
            setClients((prev) => (prev.some((c) => c.id === incoming.id) ? prev : [...prev, incoming]));
          }
          if (payload.eventType === "UPDATE") {
            const incoming = { id: payload.new.id, ...(payload.new.data || {}) };
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
      .subscribe((status) => {
        // Reconnexion automatique + re-synchronisation si la connexion WebSocket tombe
        // (veille du poste, coupure réseau…) : on récupère alors ce qui a pu être manqué
        // pendant la coupure, pour rester à jour sans que l'utilisateur ait à recharger.
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          Promise.all([loadClientsFromSupabase(), loadTeamFromSupabase()]).then(([c, t]) => {
            if (c) setClients(migrateClients(c));
            if (t) setTeam(t);
          });
        }
      });

    return () => { supabase.removeChannel(channel); };
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
    if (me === oldName) { setMe(newName); saveMe(newName); }
  }, [team, persistMany, me]);

  const addTeamMember = useCallback((nom) => {
    if (!nom.trim() || team.some((t) => t.nom === nom.trim())) return;
    const color = PALETTE[team.length % PALETTE.length];
    const member = { id: `t-${Date.now()}`, nom: nom.trim(), color };
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

  if (loading || !team) {
    return (
      <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", flexDirection: "column", gap: 14 }}>
        <Loader2 className="spin" size={28} color={T.navy} />
        <div style={{ fontFamily: T.mono, fontSize: 12, letterSpacing: "0.08em", color: T.inkMuted, textTransform: "uppercase" }}>Ouverture du registre…</div>
        <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }
  if (!me) return <WelcomeGate team={team} onPick={(n) => { setMe(n); saveMe(n); }} />;

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

  return (
    <div style={S.appShell}>
      <GlobalStyle />
      <Sidebar view={view} setView={(v) => navTo(v)} me={me} team={team}
        onSwitchMe={() => setMe(null)} onLogout={onLogout} counts={computeCounts(myClients)}
        collapsed={sidebarCollapsed} setCollapsed={setSidebarCollapsed} />
      <div style={S.main}>
        <TopBar search={search} setSearch={setSearch} saveStatus={saveStatus} me={me} meColor={meColor}
          openTabs={openClientTabs} activeTab={activeClientTab} onHome={goHome}
          onSelectTab={(id) => setActiveClientTab(id)} onCloseTab={closeClientTab}
          onNav={navTo} onOpenClient={openClientTab} onNewClient={() => setShowAddClient(true)} clients={myClients}
          notifCount={myTasks.filter((t) => t.bucket === "retard").length || undefined} />
        <div style={S.content}>
          {activeClient ? (
            // key={activeClient.id} force le remontage complet du composant à chaque
            // changement d'onglet : les champs non-contrôlés (defaultValue) et l'état
            // interne (onglet secondaire "Infos / TVA / Bilan…") sont ainsi réinitialisés
            // avec les données du dossier sélectionné, au lieu de rester figés sur
            // l'ancien dossier affiché.
            <ClientEditorPage key={activeClient.id} client={activeClient} team={team} onUpdate={updateClient}
              onClose={() => closeClientTab(activeClient.id)} />
          ) : (
            <>
              {view === "dashboard" && (
                <Dashboard myClients={myClients} tasks={myTasks} me={me}
                  onOpenClient={(id) => { navTo("clients"); openClientTab(id); }} setView={navTo} />
              )}
              {view === "clients" && (
                <ClientsRegistry clients={myClients} allClients={clients} search={search} roleFilter={roleFilter} setRoleFilter={setRoleFilter}
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
              {view === "planning" && <PlanningView tasks={myTasks} me={me} />}
              {view === "equipe" && <EquipeView team={team} clients={clients} onAdd={addTeamMember} onRename={renameTeamMember} onDelete={deleteTeamMember} />}
            </>
          )}
        </div>
      </div>

      {showAddClient && (
        <AddClientModal team={team} me={me} onClose={() => setShowAddClient(false)}
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
          email, password, options: { data: { full_name: fullName } },
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
      @keyframes fadeInUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
      .reveal { animation: fadeInUp .55s cubic-bezier(.16,.84,.44,1) both; }
      @media (prefers-reduced-motion: reduce) { * { animation: none !important; transition: none !important; } }
    `}</style>
  );
}
const S = {
  appShell: { display: "flex", height: "100vh", width: "100%", background: T.paper, fontFamily: T.sans, color: T.ink, overflow: "hidden" },
  main: { flex: 1, display: "flex", flexDirection: "column", minWidth: 0 },
  content: { flex: 1, overflowY: "auto", padding: "22px 28px 48px" },
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
   WELCOME GATE
   ============================================================ */
function WelcomeGate({ team, onPick }) {
  return (
    <div style={{ ...S.appShell, alignItems: "center", justifyContent: "center", flexDirection: "column", background: `radial-gradient(circle at 20% 15%, ${T.navySoft} 0%, ${T.paper} 45%), radial-gradient(circle at 85% 85%, #F1F5F9 0%, ${T.paper} 40%)` }}>
      <GlobalStyle />
      <Reveal style={{ textAlign: "center", maxWidth: 440, padding: 36, background: T.card, borderRadius: T.radiusLg, boxShadow: T.shadowLg, border: `1px solid ${T.line}` }}>
        <div style={{ width: 44, height: 44, borderRadius: 14, background: T.navy, margin: "0 auto 16px", display: "flex", alignItems: "center", justifyContent: "center" }}>
          <LayoutGrid size={20} color="#fff" strokeWidth={2.2} />
        </div>
        <div style={{ fontFamily: T.mono, fontSize: 11, letterSpacing: "0.18em", color: T.navy, textTransform: "uppercase", marginBottom: 10, fontWeight: 600 }}>AXE-EXPERTS</div>
        <h1 style={{ fontFamily: T.serif, fontWeight: 700, fontSize: 20, margin: "0 0 8px", color: T.ink }}>Qui consulte le dossier ?</h1>
        <p style={{ color: T.inkMuted, fontSize: 12.5, lineHeight: 1.6, marginBottom: 26 }}>
          Chacun ne voit que ses propres dossiers (en tant que collaborateur, expert ou chef de mission).
        </p>
        <div style={{ display: "grid", gap: 10 }}>
          {team.map((t, i) => (
            <Reveal key={t.id} index={i} delay={0.15}>
              <button onClick={() => onPick(t.nom)} className="clickable" style={{
                display: "flex", alignItems: "center", gap: 12, padding: "14px 18px", borderRadius: 14,
                border: `1.5px solid ${T.line}`, background: T.paper, cursor: "pointer", fontSize: 13, fontWeight: 600, color: T.ink, textAlign: "left", width: "100%",
              }}
              onMouseEnter={(e) => { e.currentTarget.style.borderColor = t.color; e.currentTarget.style.background = T.card; }}
              onMouseLeave={(e) => { e.currentTarget.style.borderColor = T.line; e.currentTarget.style.background = T.paper; }}
            >
              <span style={{ width: 30, height: 30, borderRadius: "50%", background: t.color, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontFamily: T.serif, fontWeight: 700, fontSize: 12.5 }}>{t.nom?.[0]}</span>
              {t.nom}
              <ChevronRight size={16} style={{ marginLeft: "auto", color: T.inkMuted }} />
              </button>
            </Reveal>
          ))}
        </div>
      </Reveal>
    </div>
  );
}

/* ============================================================
   SIDEBAR
   ============================================================ */
function Sidebar({ view, setView, me, team, onSwitchMe, onLogout, counts, collapsed, setCollapsed }) {
  const items = [
    { id: "dashboard", label: "Vue d'ensemble", icon: LayoutGrid },
    { id: "clients", label: "Registre clients", icon: Users, badge: counts.total },
    { id: "tva", label: "TVA — CA3/CA12", icon: Receipt, badge: counts.tvaAlert, badgeTone: "amber" },
    { id: "bilans", label: "Bilans", icon: FileWarning, badge: counts.bilanRetard, badgeTone: "red" },
    { id: "acomptes", label: "Acomptes IS / CFE", icon: Landmark },
    { id: "age", label: "AGE / AGO", icon: Building2, badge: counts.ageAlert, badgeTone: "amber" },
    { id: "mission", label: "Dossiers en accueil", icon: ClipboardCheck, badge: counts.missionIncomplete, badgeTone: "amber" },
    { id: "regimes", label: "Changements de régime", icon: RefreshCw },
    { id: "fiscal", label: "Suivi fiscal", icon: CalendarDays },
    { id: "planning", label: "Mon planning", icon: CalendarRange },
    { id: "equipe", label: "Équipe", icon: Settings2 },
  ];
  const meColor = team.find((t) => t.nom === me)?.color || T.navy;
  const W = collapsed ? 76 : 244;
  return (
    <div style={{ width: W, flexShrink: 0, background: T.sidebarBg, color: T.sidebarInk, display: "flex", flexDirection: "column", padding: "22px 12px", borderRight: `1px solid ${T.line}`, transition: "width .18s ease" }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "0 6px 20px", marginBottom: 12, borderBottom: `1px solid ${T.line}`, justifyContent: collapsed ? "center" : "flex-start" }}>
        <div style={{ width: 34, height: 34, borderRadius: 10, background: T.navy, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, fontFamily: T.serif, fontWeight: 800, color: "#fff", fontSize: 14 }}>A</div>
        {!collapsed && (
          <div>
            <div style={{ fontFamily: T.serif, fontSize: 13, fontWeight: 800, letterSpacing: "0.01em", color: T.ink }}>AXE-EXPERTS</div>
            <div style={{ fontFamily: T.mono, fontSize: 9.5, color: T.inkMuted }}>Registre &amp; Pilotage</div>
          </div>
        )}
        <button onClick={() => setCollapsed(!collapsed)} title="Réduire / agrandir" style={{ marginLeft: "auto", background: "none", border: "none", color: T.inkMuted, cursor: "pointer", display: collapsed ? "none" : "flex" }}>
          <ChevronLeft size={14} />
        </button>
      </div>
      <nav className="scrollbar" style={{ display: "flex", flexDirection: "column", gap: 2, overflowY: "auto" }}>
        {items.map((it, i) => {
          const active = view === it.id; const Icon = it.icon;
          return (
            <Reveal key={it.id} index={Math.min(i, 6)} transition={{ duration: 0.35, delay: i * 0.035, ease: EASE }}>
              <button onClick={() => setView(it.id)} title={it.label} style={{
                display: "flex", alignItems: "center", gap: 10, padding: collapsed ? "10px 0" : "9px 11px", borderRadius: 10, border: "none",
                background: active ? T.sidebarActive : "transparent", color: active ? T.navy : T.inkSoft,
                cursor: "pointer", fontSize: 12.5, fontWeight: active ? 700 : 500, textAlign: "left", width: "100%",
                justifyContent: collapsed ? "center" : "flex-start",
              }}>
                <Icon size={16} strokeWidth={2} style={{ flexShrink: 0 }} />
                {!collapsed && <span style={{ flex: 1 }}>{it.label}</span>}
                {!collapsed && !!it.badge && (
                  <span style={{
                    fontFamily: T.sans, fontSize: 10, fontWeight: 700, padding: "1px 7px", borderRadius: 999,
                    background: it.badgeTone === "red" ? T.redSoft : it.badgeTone === "amber" ? T.amberSoft : T.navySoft,
                    color: it.badgeTone === "red" ? T.red : it.badgeTone === "amber" ? T.amber : T.navy,
                  }}>{it.badge}</span>
                )}
              </button>
            </Reveal>
          );
        })}
      </nav>
      <div style={{ marginTop: "auto", paddingTop: 14, borderTop: `1px solid ${T.line}`, display: "flex", flexDirection: "column", gap: 6 }}>
        <button onClick={onSwitchMe} title="Changer de session" style={{ display: "flex", alignItems: "center", gap: 9, width: "100%", background: T.paper, border: `1px solid ${T.line}`, cursor: "pointer", padding: collapsed ? "7px" : "8px 9px", borderRadius: 11, color: T.ink, justifyContent: collapsed ? "center" : "flex-start" }}>
          <span style={{ width: 24, height: 24, borderRadius: "50%", background: meColor, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: T.serif, fontWeight: 700, fontSize: 11, color: "#fff", flexShrink: 0 }}>{me?.[0]}</span>
          {!collapsed && (
            <div style={{ textAlign: "left" }}>
              <div style={{ fontSize: 11.5, fontWeight: 700 }}>{me}</div>
              <div style={{ fontSize: 10, color: T.inkMuted }}>Changer de session</div>
            </div>
          )}
        </button>
        {onLogout && (
          <button onClick={onLogout} title="Déconnexion" style={{ display: "flex", alignItems: "center", gap: 9, width: "100%", background: "none", border: "none", cursor: "pointer", padding: collapsed ? "7px" : "7px 9px", borderRadius: 10, color: T.red, fontSize: 11.5, fontWeight: 600, justifyContent: collapsed ? "center" : "flex-start" }}>
            <LogOut size={14} />
            {!collapsed && <span>Déconnexion</span>}
          </button>
        )}
      </div>
    </div>
  );
}

/* ============================================================
   TOP BAR
   ============================================================ */
function TopBar({ search, setSearch, saveStatus, me, meColor, openTabs, activeTab, onHome, onSelectTab, onCloseTab, onNav, onOpenClient, onNewClient, clients, notifCount }) {
  const [searchOpen, setSearchOpen] = useState(false);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [pickerQuery, setPickerQuery] = useState("");
  const pickerRef = useRef(null);

  // Recherche rapide : Entrée dans le champ de recherche ouvre directement le
  // dossier trouvé (par nom ou SIREN) s'il n'y en a qu'un seul de correspondant.
  const handleSearchKeyDown = (e) => {
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
          <div style={{ position: "relative", marginLeft: 10, flex: "0 1 260px" }}>
            <Search size={13} style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: T.inkMuted }} />
            <input autoFocus value={search} onChange={(e) => setSearch(e.target.value)} onKeyDown={handleSearchKeyDown}
              placeholder="Rechercher un dossier, un SIREN… (Entrée pour ouvrir)"
              style={{ width: "100%", padding: "7px 10px 7px 28px", borderRadius: 999, border: `1px solid ${T.line}`, background: T.paper, fontSize: 12, color: T.ink }} />
          </div>
        )}

        <div style={{ marginLeft: "auto", display: "flex", alignItems: "center", gap: 1 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 4, fontFamily: T.mono, fontSize: 10.5, color: T.inkMuted, marginRight: 8 }}>
            {saveStatus === "saving" && <><Loader2 size={12} className="spin" /> enreg.…</>}
            {saveStatus === "saved" && <><Check size={12} color={T.green} /> enregistré</>}
          </div>
          <button className="topIconBtn" title="Rechercher" onClick={() => setSearchOpen((s) => !s)}><Search size={16} strokeWidth={1.9} /></button>
          {toolIcons.map((ic) => {
            const Icon = ic.icon;
            return (
              <button key={ic.key} className="topIconBtn" title={ic.title} onClick={ic.onClick}>
                <Icon size={16} strokeWidth={1.9} />
                {!!ic.badge && (
                  <span style={{ position: "absolute", top: 2, right: 2, background: T.amber, color: "#fff", fontSize: 9, fontWeight: 700, borderRadius: 999, padding: "0 4px", lineHeight: "13px" }}>{ic.badge}</span>
                )}
              </button>
            );
          })}
          <span style={{ width: 26, height: 26, borderRadius: "50%", background: meColor, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: T.serif, fontWeight: 700, fontSize: 11, color: "#fff", marginLeft: 6 }}>{me?.[0]}</span>
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
function filterClients(clients, search, roleFilter, me, regimeFilter) {
  let out = filterByRole(clients, me, roleFilter || "Tous");
  if (regimeFilter && regimeFilter !== "Tous") out = out.filter((c) => c.tvaRegime === regimeFilter);
  if (search.trim()) {
    const q = search.trim().toLowerCase();
    out = out.filter((c) => c.nom.toLowerCase().includes(q) || (c.siren || "").includes(q));
  }
  return out;
}
function Stamped({ tone = "green", children, small }) {
  const map = {
    green: { bg: T.greenSoft, fg: T.green }, red: { bg: T.redSoft, fg: T.red },
    amber: { bg: T.amberSoft, fg: T.amber }, neutral: { bg: T.paperDeep, fg: T.inkMuted },
  };
  const c = map[tone] || map.neutral;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5, fontFamily: T.sans, fontSize: small ? 9.5 : 11, fontWeight: 600,
      letterSpacing: "0.01em", padding: small ? "2px 7px" : "4px 12px", borderRadius: 999,
      background: c.bg, color: c.fg,
    }}>{children}</span>
  );
}
function RoleBadge({ role, name }) {
  if (!name) return null;
  return <span style={{ fontSize: 11, color: T.inkMuted }}>{role}: <strong style={{ color: T.inkSoft }}>{name}</strong></span>;
}

/* ============================================================
   FILTER BAR (rôle + régime)
   ============================================================ */
function FilterBar({ roleFilter, setRoleFilter, count, regimeFilter, setRegimeFilter }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
        <Filter size={14} color={T.inkMuted} />
        {["Tous", "Collaborateur", "Expert", "Chef de mission"].map((r) => (
          <button key={r} onClick={() => setRoleFilter(r)} style={{
            padding: "6px 12px", borderRadius: 20, fontSize: 12.5, fontWeight: 600,
            border: `1px solid ${roleFilter === r ? T.navy : T.line}`, background: roleFilter === r ? T.navy : T.card,
            color: roleFilter === r ? "#fff" : T.inkSoft, cursor: "pointer",
          }}>{r === "Tous" ? "Tous mes dossiers" : `Je suis ${r.toLowerCase()}`}</button>
        ))}
        <span style={{ marginLeft: "auto", fontFamily: T.mono, fontSize: 11.5, color: T.inkMuted }}>{count} dossier(s)</span>
      </div>
      {setRegimeFilter && (
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <Receipt size={14} color={T.inkMuted} />
          {["Tous", ...REGIMES_TVA].map((r) => (
            <button key={r} onClick={() => setRegimeFilter(r)} style={{
              padding: "6px 12px", borderRadius: 999, fontSize: 12, fontWeight: 600,
              border: `1px solid ${regimeFilter === r ? T.navy : T.line}`, background: regimeFilter === r ? T.navySoft : T.card,
              color: regimeFilter === r ? T.navy : T.inkSoft, cursor: "pointer", fontFamily: T.mono,
            }}>{r}</button>
          ))}
        </div>
      )}
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

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 16, marginBottom: 24 }}>
        <KpiCard index={0} label="Mes dossiers" value={counts.total} icon={Users} onClick={() => setView("clients")} />
        <KpiCard index={1} label="TVA en retard ce mois" value={counts.tvaAlert} icon={Receipt} tone={counts.tvaAlert ? "amber" : "green"} onClick={() => setView("tva")} />
        <KpiCard index={2} label="Bilans en retard" value={counts.bilanRetard} icon={FileWarning} tone={counts.bilanRetard ? "red" : "green"} onClick={() => setView("bilans")} />
        <KpiCard index={3} label="Accueils incomplets" value={counts.missionIncomplete} icon={ClipboardCheck} tone={counts.missionIncomplete ? "amber" : "green"} onClick={() => setView("mission")} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 18 }}>
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
      <div style={{ background: T.card, border: `1px solid ${T.line}`, borderRadius: T.radiusLg, boxShadow: T.shadowSm, padding: "22px 24px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          <h3 style={{ fontFamily: T.serif, fontSize: 13.5, fontWeight: 700, margin: 0, color: T.ink }}>{title}</h3>{right}
        </div>{children}
      </div>
    </Reveal>
  );
}
function EmptyNote({ text }) { return <div style={{ padding: "18px 4px", color: T.inkMuted, fontSize: 12, fontStyle: "italic" }}>{text}</div>; }

/* ============================================================
   CLIENTS REGISTRY
   ============================================================ */
function ClientsRegistry({ clients, allClients, search, roleFilter, setRoleFilter, regimeFilter, setRegimeFilter, me, selected, setSelected, onAdd, onUpdate, onImport }) {
  const filtered = useMemo(() => filterClients(clients, search, roleFilter, me, regimeFilter), [clients, search, roleFilter, me, regimeFilter]);
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
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 6, flexWrap: "wrap", gap: 10 }}>
          <h1 style={{ fontFamily: T.serif, fontSize: 16, fontWeight: 700, color: T.ink, margin: 0 }}>Registre clients</h1>
          <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: "none" }} />
            <button onClick={() => fileInputRef.current?.click()} disabled={importBusy}
              style={{ display: "flex", alignItems: "center", gap: 6, background: T.card, color: T.inkSoft, border: `1px solid ${T.line}`, borderRadius: 10, padding: "9px 14px", fontSize: 12, fontWeight: 600, cursor: importBusy ? "default" : "pointer" }}>
              {importBusy ? <Loader2 size={14} className="spin" /> : <ArrowUpRight size={14} style={{ transform: "rotate(-90deg)" }} />}
              Importer (Excel/CSV)
            </button>
            <button onClick={() => exportClientsToExcel(allClients || clients)}
              style={{ display: "flex", alignItems: "center", gap: 6, background: T.card, color: T.inkSoft, border: `1px solid ${T.line}`, borderRadius: 10, padding: "9px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              <ArrowUpRight size={14} style={{ transform: "rotate(90deg)" }} /> Exporter la liste (Excel)
            </button>
            <button onClick={onAdd} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              <Plus size={15} /> Nouveau client
            </button>
          </div>
        </div>
        {importMsg && (
          <div style={{
            marginTop: 8, fontSize: 12, fontWeight: 600, padding: "6px 10px", borderRadius: 8, display: "inline-block",
            color: importMsg.tone === "green" ? T.green : importMsg.tone === "red" ? T.red : T.amber,
            background: importMsg.tone === "green" ? T.greenSoft : importMsg.tone === "red" ? T.redSoft : T.amberSoft,
          }}>{importMsg.text}</div>
        )}
      </Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12, marginTop: 6, marginBottom: 20 }}>Cliquez un dossier pour ouvrir sa fiche complète.</p>
      <FilterBar roleFilter={roleFilter} setRoleFilter={setRoleFilter} count={filtered.length} regimeFilter={regimeFilter} setRegimeFilter={setRegimeFilter} />

      <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1.3fr 40px", padding: "0 18px", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", color: T.inkMuted, fontWeight: 600, marginBottom: 10 }}>
        <div>Dossier</div><div>SIREN</div><div>Rôles</div><div>Clôture</div><div>Régime</div><div>Statuts</div><div />
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {Object.keys(grouped).sort().map((letter) => (
          <div key={letter}>
            <div style={{ padding: "4px 6px", fontFamily: T.mono, fontSize: 10.5, fontWeight: 700, color: T.navy, letterSpacing: "0.1em" }}>{letter}</div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {grouped[letter].map((c) => {
                rowIndex += 1;
                return (
                  <Reveal key={c.id} index={rowIndex}>
                    <div className="hoverRow clickable" onClick={() => setSelected(c.id)} style={{
                      display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1.3fr 40px", padding: "14px 18px", alignItems: "center", fontSize: 12,
                      borderRadius: T.radiusSm, border: `1px solid ${selected === c.id ? T.navy : T.line}`,
                      background: selected === c.id ? T.navySoft : T.card, boxShadow: T.shadowSm,
                    }}>
                      <div style={{ fontWeight: 600, color: T.ink }}>{c.nom}</div>
                      <div style={{ fontFamily: T.mono, fontSize: 12, color: T.inkMuted }}>{c.siren}</div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 1, fontSize: 10.5, color: T.inkMuted }}>
                        {c.collab === me && <span>Collaborateur</span>}
                        {c.expert === me && <span>Expert</span>}
                        {c.chefMission === me && <span>Chef de mission</span>}
                      </div>
                      <div style={{ fontFamily: T.mono, fontSize: 11.5, color: T.inkMuted }}>
                        <input type="date" defaultValue={c.dateCloture || ""} onClick={(e) => e.stopPropagation()}
                          onChange={(e) => onUpdate(c.id, { dateCloture: e.target.value })}
                          style={{ border: "none", background: "transparent", fontFamily: T.mono, fontSize: 11.5, color: T.inkMuted, width: 118 }} />
                      </div>
                      <div style={{ fontSize: 12, color: T.inkSoft, fontFamily: T.mono }}>{c.tvaRegime || "—"}</div>
                      <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                        {isBilanLate(c) && <Stamped tone="red" small>Bilan retard</Stamped>}
                        {isTvaLate(c) && <Stamped tone="amber" small>TVA</Stamped>}
                        {!isBilanLate(c) && !isTvaLate(c) && <Stamped tone="green" small>À jour</Stamped>}
                      </div>
                      <ChevronRight size={15} color={T.inkMuted} />
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
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 4 }}>
          <div>
            <div style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{client.siren || "SIREN non renseigné"}</div>
            <input defaultValue={client.nom} onBlur={(e) => onUpdate(client.id, { nom: e.target.value || client.nom })}
              style={{ fontFamily: T.serif, fontSize: 16, fontWeight: 700, color: T.ink, border: "none", background: "transparent", padding: "2px 0", margin: "2px 0 6px", minWidth: 260 }} />
            <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap", fontSize: 12 }}>
              <RoleBadge role="Collab." name={client.collab} />
              <RoleBadge role="Expert" name={client.expert} />
              <RoleBadge role="Chef de mission" name={client.chefMission} />
              {client.tvaRegime && <span style={{ fontFamily: T.mono, fontSize: 11, color: T.navy, fontWeight: 700, background: T.navySoft, padding: "2px 9px", borderRadius: 999 }}>{client.tvaRegime}</span>}
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
function PlanningView({ tasks, me }) {
  const [mode, setMode] = useState("semaine"); // jour | semaine | mois | année
  const now = new Date();

  const grouped = useMemo(() => {
    if (mode === "jour") {
      return { [now.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" })]: tasks.filter((t) => sameDay(t.date, now)) };
    }
    if (mode === "semaine") {
      const start = startOfWeek(now);
      const days = Array.from({ length: 7 }, (_, i) => { const d = new Date(start); d.setDate(d.getDate() + i); return d; });
      const g = {};
      days.forEach((d) => { g[d.toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "short" })] = tasks.filter((t) => sameDay(t.date, d)); });
      return g;
    }
    if (mode === "mois") {
      const g = {};
      tasks.filter((t) => t.date.getMonth() === now.getMonth() && t.date.getFullYear() === now.getFullYear())
        .forEach((t) => { const k = `Semaine du ${fmtFR(startOfWeek(t.date).toISOString().slice(0, 10))}`; g[k] = g[k] || []; g[k].push(t); });
      return g;
    }
    // année
    const g = {};
    tasks.filter((t) => t.date.getFullYear() === now.getFullYear()).forEach((t) => {
      const k = t.date.toLocaleDateString("fr-FR", { month: "long" }); g[k] = g[k] || []; g[k].push(t);
    });
    return g;
  }, [mode, tasks, now]);

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Mon planning</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Vos tâches et échéances, {me}.</p>

      <div style={{ display: "flex", gap: 6, marginBottom: 18 }}>
        {[["jour", "Jour"], ["semaine", "Semaine"], ["mois", "Mois"], ["annee", "Année"]].map(([id, label]) => (
          <button key={id} onClick={() => setMode(id === "annee" ? "année" : id)} style={{
            padding: "7px 14px", borderRadius: 20, fontSize: 12.5, fontWeight: 600,
            border: `1px solid ${mode === (id === "annee" ? "année" : id) ? T.navy : T.line}`,
            background: mode === (id === "annee" ? "année" : id) ? T.navy : T.card,
            color: mode === (id === "annee" ? "année" : id) ? "#fff" : T.inkSoft, cursor: "pointer",
          }}>{label}</button>
        ))}
      </div>

      {Object.entries(grouped).map(([label, items]) => (
        <Panel key={label} title={`${label.charAt(0).toUpperCase() + label.slice(1)} (${items.length})`}>
          {items.length === 0 ? <EmptyNote text="Rien de prévu." /> : (
            <div>
              {[...items].sort((a, b) => a.date - b.date).map((t) => (
                <div key={t.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "9px 4px", borderBottom: `1px solid ${T.line}` }}>
                  <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted, width: 60 }}>{fmtFR(t.date.toISOString().slice(0, 10))}</span>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 600, fontSize: 12 }}>{t.client.nom}</div>
                    <div style={{ fontSize: 11.5, color: T.inkMuted }}>{t.label}</div>
                  </div>
                  <Stamped tone={t.tone} small>{t.category}</Stamped>
                </div>
              ))}
            </div>
          )}
        </Panel>
      ))}
      <div style={{ height: 8 }} />
    </div>
  );
}

/* ============================================================
   ÉQUIPE
   ============================================================ */
function EquipeView({ team, clients, onAdd, onRename, onDelete }) {
  const [newName, setNewName] = useState("");
  const [editing, setEditing] = useState(null);
  const [editValue, setEditValue] = useState("");

  const countFor = (nom) => clients.filter((c) => c.collab === nom || c.expert === nom || c.chefMission === nom).length;

  return (
    <div>
      <Reveal><h1 style={{ fontFamily: T.serif, fontSize: 17, fontWeight: 700, color: T.ink, margin: "0 0 6px" }}>Équipe</h1></Reveal>
      <p style={{ color: T.inkMuted, fontSize: 12.5, marginTop: 0, marginBottom: 18 }}>Gérez les collaborateurs, experts et chefs de mission du cabinet.</p>

      <Panel title="Ajouter un membre">
        <div style={{ display: "flex", gap: 8 }}>
          <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Nom du collaborateur" style={{ ...inputStyle, flex: 1 }} onKeyDown={(e) => e.key === "Enter" && (onAdd(newName), setNewName(""))} />
          <button onClick={() => { onAdd(newName); setNewName(""); }} style={{ display: "flex", alignItems: "center", gap: 6, background: T.navy, color: "#fff", border: "none", borderRadius: 10, padding: "9px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
            <Plus size={15} /> Ajouter
          </button>
        </div>
      </Panel>

      <div style={{ height: 16 }} />
      <Panel title={`Membres de l'équipe (${team.length})`}>
        {team.map((t) => (
          <div key={t.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "11px 4px", borderBottom: `1px solid ${T.line}` }}>
            <span style={{ width: 26, height: 26, borderRadius: "50%", background: t.color, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: T.serif, fontWeight: 600, fontSize: 12, color: "#fff", flexShrink: 0 }}>{t.nom[0]}</span>
            {editing === t.id ? (
              <input value={editValue} onChange={(e) => setEditValue(e.target.value)} autoFocus
                onBlur={() => { onRename(t.nom, editValue); setEditing(null); }}
                onKeyDown={(e) => { if (e.key === "Enter") { onRename(t.nom, editValue); setEditing(null); } }}
                style={{ ...inputStyle, flex: 1 }} />
            ) : (
              <span style={{ flex: 1, fontWeight: 600, fontSize: 12.5 }}>{t.nom}</span>
            )}
            <span style={{ fontFamily: T.mono, fontSize: 11, color: T.inkMuted }}>{countFor(t.nom)} dossier(s)</span>
            <button onClick={() => { setEditing(t.id); setEditValue(t.nom); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.inkMuted }}><Pencil size={14} /></button>
            <button onClick={() => { if (confirm(`Supprimer ${t.nom} de l'équipe ? Ses dossiers seront désassignés.`)) onDelete(t.nom); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.red }}><Trash2 size={14} /></button>
          </div>
        ))}
      </Panel>
    </div>
  );
}

/* ============================================================
   ADD CLIENT MODAL
   ============================================================ */
function AddClientModal({ team, me, onClose, onCreate }) {
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
      id: `c-${Date.now()}`, nom: nom.trim(), siren: siren.trim(), logiciel, dateCloture,
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
