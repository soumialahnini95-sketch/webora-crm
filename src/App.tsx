import { FormEvent, useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

type Status =
  | "À contacter"
  | "Contact initial"
  | "En discussion"
  | "Démo planifiée"
  | "C'est OK"
  | "Perdu";

type Note = {
  id: string;
  text: string;
  date: string;
};

type Prospect = {
  id: string;
  company: string;
  city: string;
  industry: string;
  contact: string;
  phone: string;
  email: string;
  status: Status;
  followUp: string;
  amount: number;
  notes: Note[];
  createdAt: string;
};

type IconName =
  | "dashboard"
  | "contacts"
  | "pipeline"
  | "settings"
  | "search"
  | "plus"
  | "table"
  | "kanban"
  | "phone"
  | "mail"
  | "edit"
  | "users"
  | "message"
  | "trophy"
  | "chart"
  | "moon"
  | "sun"
  | "close"
  | "calendar"
  | "chevron"
  | "more"
  | "trash"
  | "sparkles";

type User = {
  name: string;
  email: string;
};

const statuses: Status[] = [
  "À contacter",
  "Contact initial",
  "En discussion",
  "Démo planifiée",
  "C'est OK",
  "Perdu",
];

const statusStyle: Record<Status, string> = {
  "À contacter": "status-slate",
  "Contact initial": "status-blue",
  "En discussion": "status-amber",
  "Démo planifiée": "status-violet",
  "C'est OK": "status-green",
  Perdu: "status-red",
};

const initialProspects: Prospect[] = [
  {
    id: "1",
    company: "Atelier Noé",
    city: "Casablanca",
    industry: "Architecture",
    contact: "Clara Martin",
    phone: "+212 6 12 34 56 78",
    email: "clara@ateliernoe.fr",
    status: "En discussion",
    followUp: "2026-03-20",
    amount: 18000,
    notes: [
      {
        id: "n1",
        text: "Intéressée par l'offre Équipe. Envoyer le détail des intégrations.",
        date: "2026-03-12T10:30:00",
      },
    ],
    createdAt: "2026-03-03",
  },
  {
    id: "2",
    company: "Maison Junot",
    city: "Rabat",
    industry: "Immobilier",
    contact: "Hugo Bernard",
    phone: "+212 6 28 45 17 90",
    email: "hugo@maisonjunot.fr",
    status: "Démo planifiée",
    followUp: "2026-03-18",
    amount: 32000,
    notes: [
      {
        id: "n2",
        text: "Démo confirmée avec la direction commerciale.",
        date: "2026-03-13T14:15:00",
      },
    ],
    createdAt: "2026-03-01",
  },
  {
    id: "3",
    company: "Café Miro",
    city: "Marrakech",
    industry: "Restauration",
    contact: "Sarah Petit",
    phone: "+212 7 14 22 09 61",
    email: "sarah@cafemiro.fr",
    status: "C'est OK",
    followUp: "2026-03-26",
    amount: 12400,
    notes: [
      {
        id: "n3",
        text: "Contrat annuel signé. Onboarding prévu la semaine prochaine.",
        date: "2026-03-10T09:00:00",
      },
    ],
    createdAt: "2026-02-18",
  },
  {
    id: "4",
    company: "Studio Kanso",
    city: "Casablanca",
    industry: "Design",
    contact: "Léa Robert",
    phone: "+212 6 90 31 72 44",
    email: "lea@studiokanso.fr",
    status: "Contact initial",
    followUp: "2026-03-19",
    amount: 8500,
    notes: [],
    createdAt: "2026-03-08",
  },
  {
    id: "5",
    company: "Novalys Conseil",
    city: "Tanger",
    industry: "Conseil",
    contact: "Marc Leroy",
    phone: "+212 7 55 19 86 02",
    email: "marc@novalys.fr",
    status: "À contacter",
    followUp: "2026-03-21",
    amount: 24000,
    notes: [],
    createdAt: "2026-03-14",
  },
  {
    id: "6",
    company: "Les Jardins",
    city: "Agadir",
    industry: "Paysagisme",
    contact: "Inès Moreau",
    phone: "+212 6 44 08 92 17",
    email: "ines@lesjardins.fr",
    status: "Perdu",
    followUp: "",
    amount: 9600,
    notes: [
      {
        id: "n4",
        text: "Budget reporté au second semestre.",
        date: "2026-03-08T16:45:00",
      },
    ],
    createdAt: "2026-02-22",
  },
  {
    id: "7",
    company: "Orbe Finance",
    city: "Rabat",
    industry: "Finance",
    contact: "Thomas Roux",
    phone: "+212 7 81 20 43 66",
    email: "thomas@orbe.finance",
    status: "C'est OK",
    followUp: "2026-04-02",
    amount: 44500,
    notes: [],
    createdAt: "2026-02-12",
  },
  {
    id: "8",
    company: "Plume Éditions",
    city: "Casablanca",
    industry: "Édition",
    contact: "Eva Fournier",
    phone: "+212 6 18 76 33 50",
    email: "eva@plume-editions.fr",
    status: "En discussion",
    followUp: "2026-03-22",
    amount: 15200,
    notes: [],
    createdAt: "2026-03-05",
  },
];

const emptyProspect: Omit<Prospect, "id" | "createdAt"> = {
  company: "",
  city: "",
  industry: "",
  contact: "",
  phone: "",
  email: "",
  status: "À contacter",
  followUp: "",
  amount: 0,
  notes: [],
};

function Icon({ name, size = 18 }: { name: IconName; size?: number }) {
  const paths: Record<IconName, React.ReactNode> = {
    dashboard: (
      <>
        <rect x="3" y="3" width="7" height="7" rx="2" />
        <rect x="14" y="3" width="7" height="7" rx="2" />
        <rect x="3" y="14" width="7" height="7" rx="2" />
        <rect x="14" y="14" width="7" height="7" rx="2" />
      </>
    ),
    contacts: (
      <>
        <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
        <circle cx="9" cy="7" r="4" />
        <path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" />
      </>
    ),
    pipeline: (
      <>
        <path d="M3 6h18M6 6v12M12 6v12M18 6v12" />
        <circle cx="6" cy="10" r="2" />
        <circle cx="12" cy="15" r="2" />
        <circle cx="18" cy="11" r="2" />
      </>
    ),
    settings: (
      <>
        <circle cx="12" cy="12" r="3" />
        <path d="M19.4 15a1.7 1.7 0 0 0 .34 1.88l.06.06-2.83 2.83-.06-.06a1.7 1.7 0 0 0-1.88-.34 1.7 1.7 0 0 0-1.03 1.56V21h-4v-.09A1.7 1.7 0 0 0 9 19.37a1.7 1.7 0 0 0-1.88.34l-.06.06-2.83-2.83.06-.06A1.7 1.7 0 0 0 4.63 15 1.7 1.7 0 0 0 3.09 14H3v-4h.09A1.7 1.7 0 0 0 4.63 9a1.7 1.7 0 0 0-.34-1.88l-.06-.06 2.83-2.83.06.06A1.7 1.7 0 0 0 9 4.63h.02A1.7 1.7 0 0 0 10 3.09V3h4v.09A1.7 1.7 0 0 0 15 4.63a1.7 1.7 0 0 0 1.88-.34l.06-.06 2.83 2.83-.06.06A1.7 1.7 0 0 0 19.37 9v.02A1.7 1.7 0 0 0 20.91 10H21v4h-.09A1.7 1.7 0 0 0 19.4 15Z" />
      </>
    ),
    search: <><circle cx="11" cy="11" r="7" /><path d="m20 20-4-4" /></>,
    plus: <path d="M12 5v14M5 12h14" />,
    table: <><rect x="3" y="4" width="18" height="16" rx="2" /><path d="M3 10h18M9 4v16" /></>,
    kanban: <><rect x="3" y="4" width="5" height="16" rx="1" /><rect x="10" y="4" width="5" height="10" rx="1" /><rect x="17" y="4" width="4" height="13" rx="1" /></>,
    phone: <path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6A19.79 19.79 0 0 1 2.12 4.18 2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.12.9.33 1.78.62 2.63a2 2 0 0 1-.45 2.11L8 9.73a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.85.29 1.73.5 2.63.62A2 2 0 0 1 22 16.92Z" />,
    mail: <><rect x="3" y="5" width="18" height="14" rx="2" /><path d="m3 7 9 6 9-6" /></>,
    edit: <><path d="M12 20h9" /><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L8 18l-4 1 1-4Z" /></>,
    users: <><path d="M16 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" /><circle cx="8.5" cy="7" r="4" /><path d="M20 8v6M23 11h-6" /></>,
    message: <><path d="M21 15a4 4 0 0 1-4 4H8l-5 3V7a4 4 0 0 1 4-4h10a4 4 0 0 1 4 4Z" /><path d="M8 9h8M8 13h5" /></>,
    trophy: <><path d="M8 21h8M12 17v4M7 4h10v5a5 5 0 0 1-10 0Z" /><path d="M7 6H4v2a4 4 0 0 0 4 4M17 6h3v2a4 4 0 0 1-4 4" /></>,
    chart: <><path d="M3 3v18h18" /><path d="m7 16 4-5 4 3 5-7" /></>,
    moon: <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79Z" />,
    sun: <><circle cx="12" cy="12" r="4" /><path d="M12 2v2M12 20v2M4.93 4.93l1.42 1.42M17.66 17.66l1.41 1.41M2 12h2M20 12h2M4.93 19.07l1.42-1.42M17.66 6.34l1.41-1.41" /></>,
    close: <path d="M18 6 6 18M6 6l12 12" />,
    calendar: <><rect x="3" y="5" width="18" height="16" rx="2" /><path d="M16 3v4M8 3v4M3 11h18" /></>,
    chevron: <path d="m9 18 6-6-6-6" />,
    more: <><circle cx="5" cy="12" r="1" fill="currentColor" /><circle cx="12" cy="12" r="1" fill="currentColor" /><circle cx="19" cy="12" r="1" fill="currentColor" /></>,
    trash: <><path d="M3 6h18M8 6V4h8v2M19 6l-1 15H6L5 6M10 11v6M14 11v6" /></>,
    sparkles: <><path d="m12 3-1.2 3.2L8 7.5l2.8 1.3L12 12l1.2-3.2L16 7.5l-2.8-1.3Z" /><path d="m19 14-.8 2.2L16 17l2.2.8L19 20l.8-2.2L22 17l-2.2-.8ZM5 13l-.8 2.2L2 16l2.2.8L5 19l.8-2.2L8 16l-2.2-.8Z" /></>,
  };

  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      {paths[name]}
    </svg>
  );
}

function StatusBadge({ status }: { status: Status }) {
  return <span className={`status-badge ${statusStyle[status]}`}><span className="status-dot" />{status}</span>;
}

function Avatar({ name, small = false }: { name: string; small?: boolean }) {
  const initials = name.split(" ").map((word) => word[0]).slice(0, 2).join("");
  const colors = ["avatar-blue", "avatar-violet", "avatar-green", "avatar-amber", "avatar-rose"];
  const index = name.split("").reduce((sum, char) => sum + char.charCodeAt(0), 0) % colors.length;
  return <span className={`avatar ${colors[index]} ${small ? "avatar-small" : ""}`}>{initials}</span>;
}

function formatCurrency(value: number) {
  return `${new Intl.NumberFormat("fr-MA", { maximumFractionDigits: 0 }).format(value)} DH`;
}

function formatDate(value: string) {
  if (!value) return "Non planifiée";
  return new Intl.DateTimeFormat("fr-MA", { day: "numeric", month: "short", year: "numeric" }).format(new Date(`${value}T12:00:00`));
}

function MetricCard({ label, value, detail, icon, tone }: { label: string; value: string; detail: string; icon: IconName; tone: string }) {
  return (
    <article className="metric-card">
      <div className="metric-top">
        <span className={`metric-icon ${tone}`}><Icon name={icon} size={20} /></span>
        <span className="metric-trend">↗ 12%</span>
      </div>
      <p className="metric-label">{label}</p>
      <p className="metric-value">{value}</p>
      <p className="metric-detail">{detail}</p>
    </article>
  );
}

export default function App() {
  const [user, setUser] = useState<User | null>(() => {
    try {
      const saved = localStorage.getItem("webora-user");
      return saved ? JSON.parse(saved) : null;
    } catch {
      return null;
    }
  });
  const [prospects, setProspects] = useState<Prospect[]>([]);
  const [view, setView] = useState<"table" | "kanban">("table");
  const [activePage, setActivePage] = useState<"dashboard" | "contacts" | "pipeline" | "admin">("dashboard");
  const [query, setQuery] = useState("");
  const [city, setCity] = useState("Toutes les villes");
  const [industry, setIndustry] = useState("Tous les secteurs");
  const [modalOpen, setModalOpen] = useState(false);
  const [editing, setEditing] = useState<Prospect | null>(null);
  const [draft, setDraft] = useState<Omit<Prospect, "id" | "createdAt">>(emptyProspect);
  const [newNote, setNewNote] = useState("");
  const [dark, setDark] = useState(() => localStorage.getItem("webora-theme") === "dark");
  const [toast, setToast] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (!user) return;
    try {
      const saved = localStorage.getItem(`webora-contacts-${user.email}`);
      setProspects(saved ? JSON.parse(saved) : initialProspects);
    } catch {
      setProspects(initialProspects);
    }
  }, [user]);

  useEffect(() => {
    if (user && prospects.length >= 0) {
      localStorage.setItem(`webora-contacts-${user.email}`, JSON.stringify(prospects));
    }
  }, [prospects, user]);

  useEffect(() => {
    document.documentElement.classList.toggle("dark", dark);
    localStorage.setItem("webora-theme", dark ? "dark" : "light");
  }, [dark]);

  useEffect(() => {
    if (!toast) return;
    const timeout = window.setTimeout(() => setToast(""), 2600);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  const cities = useMemo(() => [...new Set(prospects.map((p) => p.city))].sort(), [prospects]);
  const industries = useMemo(() => [...new Set(prospects.map((p) => p.industry))].sort(), [prospects]);
  const filtered = useMemo(() => {
    const normalized = query.toLowerCase().trim();
    return prospects.filter((p) => {
      const matchesQuery = !normalized || [p.company, p.contact, p.email, p.city, p.industry].join(" ").toLowerCase().includes(normalized);
      return matchesQuery && (city === "Toutes les villes" || p.city === city) && (industry === "Tous les secteurs" || p.industry === industry);
    });
  }, [prospects, query, city, industry]);

  const won = prospects.filter((p) => p.status === "C'est OK");
  const contacted = prospects.filter((p) => p.status !== "À contacter").length;
  const conversion = prospects.length ? Math.round((won.length / prospects.length) * 100) : 0;
  const pipelineValue = prospects.filter((p) => p.status !== "Perdu").reduce((sum, p) => sum + p.amount, 0);

  function openAdd() {
    setEditing(null);
    setDraft({ ...emptyProspect, notes: [] });
    setNewNote("");
    setModalOpen(true);
  }

  function openEdit(prospect: Prospect) {
    setEditing(prospect);
    setDraft({ ...prospect, notes: [...prospect.notes] });
    setNewNote("");
    setModalOpen(true);
  }

  function saveProspect(event: FormEvent) {
    event.preventDefault();
    const notes = newNote.trim()
      ? [{ id: crypto.randomUUID(), text: newNote.trim(), date: new Date().toISOString() }, ...draft.notes]
      : draft.notes;

    if (editing) {
      setProspects((current) => current.map((p) => p.id === editing.id ? { ...p, ...draft, notes } : p));
      setToast("Prospect mis à jour");
    } else {
      setProspects((current) => [{ ...draft, notes, id: crypto.randomUUID(), createdAt: new Date().toISOString().slice(0, 10) }, ...current]);
      setToast("Nouveau prospect ajouté");
    }
    setModalOpen(false);
  }

  function changeStatus(id: string, status: Status) {
    setProspects((current) => current.map((p) => p.id === id ? { ...p, status } : p));
    setToast(`Statut mis à jour : ${status}`);
  }

  function deleteProspect(prospect: Prospect) {
    if (!window.confirm(`Supprimer définitivement le contact « ${prospect.company} » ?`)) return;
    setProspects((current) => current.filter((item) => item.id !== prospect.id));
    if (editing?.id === prospect.id) setModalOpen(false);
    setToast("Contact supprimé");
  }

  function login(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const data = new FormData(event.currentTarget);
    const nextUser = {
      name: String(data.get("name") || "").trim(),
      email: String(data.get("email") || "").trim().toLowerCase(),
    };
    localStorage.setItem("webora-user", JSON.stringify(nextUser));
    setUser(nextUser);
  }

  function logout() {
    localStorage.removeItem("webora-user");
    setUser(null);
    setProspects([]);
  }

  function exportContacts() {
    const rows = prospects.map((p) => ({
      Entreprise: p.company,
      Ville: p.city,
      Secteur: p.industry,
      Contact: p.contact,
      Téléphone: p.phone,
      Email: p.email,
      Statut: p.status,
      "Prochaine relance": p.followUp,
      "Montant (DH)": p.amount,
      Notes: p.notes.map((note) => note.text).join(" | "),
    }));
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet["!cols"] = [{ wch: 25 }, { wch: 16 }, { wch: 20 }, { wch: 22 }, { wch: 18 }, { wch: 28 }, { wch: 20 }, { wch: 18 }, { wch: 14 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(workbook, worksheet, "Contacts");
    XLSX.writeFile(workbook, `webora-contacts-${new Date().toISOString().slice(0, 10)}.xlsx`);
    setToast("Fichier Excel exporté");
  }

  async function importContacts(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const sourceRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(firstSheet, { defval: "" });
    if (!sourceRows.length) {
      setToast("Le fichier Excel est vide");
      event.target.value = "";
      return;
    }
    const value = (row: Record<string, unknown>, names: string[]) => {
      const key = Object.keys(row).find((item) => names.includes(item.toLowerCase().trim()));
      return key ? String(row[key]).trim() : "";
    };
    const imported = sourceRows.map((row) => {
      const company = value(row, ["entreprise", "société", "societe", "company"]);
      const city = value(row, ["ville", "city"]);
      const industry = value(row, ["secteur", "secteur d'activité", "activite", "activité"]);
      const contact = value(row, ["contact", "nom", "nom du contact"]);
      const phone = value(row, ["téléphone", "telephone", "tel", "numéro", "numero", "phone"]);
      const email = value(row, ["email", "e-mail", "mail"]);
      const rawStatus = value(row, ["statut", "status"]);
      const followUp = value(row, ["prochaine relance", "relance", "date de relance"]);
      const amount = value(row, ["montant", "montant (dh)", "valeur", "deal"]);
      const note = value(row, ["notes", "note", "commentaire"]);
      const status = statuses.includes(rawStatus as Status) ? rawStatus as Status : "À contacter";
      return {
        id: crypto.randomUUID(),
        company: company || "Entreprise sans nom",
        city: city || "Non renseignée",
        industry: industry || "Non renseigné",
        contact: contact || "Contact à identifier",
        phone: phone || "",
        email: email || "",
        status,
        followUp: followUp || "",
        amount: Number(String(amount || "0").replace(/\s/g, "").replace(",", ".")) || 0,
        notes: note ? [{ id: crypto.randomUUID(), text: note, date: new Date().toISOString() }] : [],
        createdAt: new Date().toISOString().slice(0, 10),
      };
    });
    setProspects((current) => [...imported, ...current]);
    event.target.value = "";
    setToast(`${imported.length} contact${imported.length > 1 ? "s" : ""} importé${imported.length > 1 ? "s" : ""}`);
  }

  function downloadTemplate() {
    const worksheet = XLSX.utils.aoa_to_sheet([
      ["Entreprise", "Ville", "Secteur", "Contact", "Téléphone", "Email", "Statut", "Prochaine relance", "Montant (DH)", "Notes"],
      ["Exemple SARL", "Casablanca", "Conseil", "Sara Alaoui", "+212 6 00 00 00 00", "sara@exemple.ma", "À contacter", "2026-04-15", 10000, "Premier appel à effectuer"],
    ]);
    worksheet["!cols"] = [{ wch: 25 }, { wch: 16 }, { wch: 20 }, { wch: 22 }, { wch: 18 }, { wch: 28 }, { wch: 20 }, { wch: 18 }, { wch: 14 }, { wch: 40 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Contacts à importer");
    XLSX.writeFile(workbook, "modele-import-webora.xlsx");
    setToast("Modèle Excel téléchargé");
  }

  if (!user) {
    return (
      <div className="login-page">
        <div className="login-visual">
          <div className="login-brand"><span className="brand-mark"><Icon name="sparkles" size={19} /></span><span>webora</span></div>
          <div className="login-message">
            <span className="login-pill">PROSPECTION, SIMPLIFIÉE</span>
            <h1>Chaque contact.<br />Chaque appel.<br /><em>Au bon moment.</em></h1>
            <p>Centralisez votre fichier de prospection, traitez chaque numéro et transformez plus d'opportunités en clients.</p>
          </div>
          <div className="login-preview">
            <div><span className="preview-dot" /><span className="preview-dot" /><span className="preview-dot" /></div>
            <div className="preview-row"><Avatar name="Maison Junot" /><span><strong>Maison Junot</strong><small>À appeler aujourd'hui</small></span><StatusBadge status="Démo planifiée" /></div>
            <div className="preview-row"><Avatar name="Studio Kanso" /><span><strong>Studio Kanso</strong><small>Relance planifiée</small></span><StatusBadge status="Contact initial" /></div>
          </div>
        </div>
        <div className="login-panel">
          <button className="theme-login" onClick={() => setDark(!dark)} aria-label="Changer de thème"><Icon name={dark ? "sun" : "moon"} /></button>
          <form className="login-form" onSubmit={login}>
            <div className="login-mobile-brand"><span className="brand-mark"><Icon name="sparkles" size={18} /></span><strong>webora</strong></div>
            <p className="eyebrow">BIENVENUE SUR WEBORA</p>
            <h2>Connectez-vous à votre espace</h2>
            <p>Retrouvez votre fichier de contacts et reprenez votre prospection là où vous l'avez laissée.</p>
            <label><span>Votre nom</span><input name="name" required placeholder="Ex. Alex Morgan" autoComplete="name" /></label>
            <label><span>Adresse email</span><input name="email" type="email" required placeholder="vous@entreprise.fr" autoComplete="email" /></label>
            <label><span>Mot de passe</span><input name="password" type="password" required minLength={4} placeholder="••••••••" autoComplete="current-password" /></label>
            <div className="login-options"><label><input type="checkbox" defaultChecked />Se souvenir de moi</label><button type="button">Mot de passe oublié ?</button></div>
            <button className="login-submit" type="submit">Accéder à mon espace <Icon name="chevron" size={16} /></button>
            <p className="login-disclaimer">En continuant, vous acceptez les conditions d'utilisation de Webora.</p>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <span className="brand-mark"><Icon name="sparkles" size={19} /></span>
          <span>webora</span>
        </div>
        <div className="workspace">
          <Avatar name={user.name} small />
          <div><span>Mon espace</span><strong>{user.name}</strong></div>
          <Icon name="chevron" size={14} />
        </div>
        <nav className="nav">
          <p>MENU</p>
          <button className={`nav-item ${activePage === "dashboard" ? "active" : ""}`} onClick={() => setActivePage("dashboard")}><Icon name="dashboard" />Vue d'ensemble</button>
          <button className={`nav-item ${activePage === "contacts" ? "active" : ""}`} onClick={() => { setActivePage("contacts"); setView("table"); }}><Icon name="contacts" />Fichier contacts<span className="nav-count">{prospects.length}</span></button>
          <button className={`nav-item ${activePage === "pipeline" ? "active" : ""}`} onClick={() => { setActivePage("pipeline"); setView("kanban"); }}><Icon name="pipeline" />Pipeline</button>
          <p className="nav-section">ESPACE</p>
          <button className={`nav-item ${activePage === "admin" ? "active" : ""}`} onClick={() => setActivePage("admin")}><Icon name="settings" />Administration</button>
        </nav>
        <div className="sidebar-bottom">
          <div className="upgrade">
            <span className="upgrade-icon"><Icon name="sparkles" size={17} /></span>
            <strong>Passez au niveau supérieur</strong>
            <p>Débloquez les automatisations et les rapports avancés.</p>
            <button>Découvrir Pro</button>
          </div>
          <div className="profile">
            <Avatar name={user.name} />
            <div><strong>{user.name}</strong><span>{user.email}</span></div>
            <button className="logout-button" onClick={logout} title="Se déconnecter"><Icon name="chevron" /></button>
          </div>
        </div>
      </aside>

      <main className="main">
        <header className="topbar">
          <div className="mobile-brand"><span className="brand-mark"><Icon name="sparkles" size={17} /></span><strong>webora</strong></div>
          <div className="global-search"><Icon name="search" size={17} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Rechercher partout…" /><kbd>⌘ K</kbd></div>
          <div className="top-actions">
            <button className="icon-button" onClick={() => setDark(!dark)} aria-label="Changer de thème"><Icon name={dark ? "sun" : "moon"} /></button>
            <button className="add-button" onClick={openAdd}><Icon name="plus" size={17} />Nouveau prospect</button>
            <details className="mobile-account">
              <summary aria-label="Ouvrir le menu du compte"><Avatar name={user.name} small /></summary>
              <div className="mobile-account-menu">
                <div><strong>{user.name}</strong><span>{user.email}</span></div>
                <button onClick={logout}>Se déconnecter</button>
              </div>
            </details>
          </div>
        </header>

        <div className="content">
          {activePage === "dashboard" && (
            <>
          <section className="page-heading">
            <div>
              <p className="eyebrow">VUE D'ENSEMBLE</p>
              <h1>Bonjour {user.name.split(" ")[0]}, voici votre activité.</h1>
              <p>Suivez vos opportunités et transformez vos prospects en clients.</p>
            </div>
            <div className="period"><Icon name="calendar" size={16} />30 derniers jours<Icon name="chevron" size={13} /></div>
          </section>

          <section className="metrics-grid">
            <MetricCard label="Total prospects" value={String(prospects.length)} detail="+3 cette semaine" icon="users" tone="tone-blue" />
            <MetricCard label="Prospects contactés" value={String(contacted)} detail={`${Math.round((contacted / Math.max(prospects.length, 1)) * 100)}% du portefeuille`} icon="message" tone="tone-violet" />
            <MetricCard label="Deals gagnés" value={String(won.length)} detail={formatCurrency(won.reduce((sum, p) => sum + p.amount, 0))} icon="trophy" tone="tone-green" />
            <MetricCard label="Taux de conversion" value={`${conversion}%`} detail="+4,2% vs. mois dernier" icon="chart" tone="tone-amber" />
          </section>

          <section className="pipeline-summary">
            <div>
              <span>Valeur du pipeline</span>
              <strong>{formatCurrency(pipelineValue)}</strong>
            </div>
            <div className="pipeline-bar">
              {statuses.slice(0, 5).map((status) => {
                const count = prospects.filter((p) => p.status === status).length;
                return <span key={status} className={statusStyle[status]} style={{ flex: Math.max(count, 0.4) }} title={`${status} : ${count}`} />;
              })}
            </div>
            <div className="pipeline-legend">
              {statuses.slice(0, 5).map((status) => <span key={status}><i className={statusStyle[status]} />{status}</span>)}
            </div>
          </section>
            </>
          )}

          {(activePage === "contacts" || activePage === "pipeline") && (
            <>
          <section className="page-heading">
            <div>
              <p className="eyebrow">{activePage === "pipeline" ? "SUIVI COMMERCIAL" : "BASE DE PROSPECTION"}</p>
              <h1>{activePage === "pipeline" ? "Pipeline de conversion" : "Votre fichier de contacts"}</h1>
              <p>{activePage === "pipeline" ? "Faites avancer chaque prospect jusqu'à la conversion." : "Retrouvez, filtrez et traitez tous les numéros à contacter."}</p>
            </div>
            <button className="add-button" onClick={openAdd}><Icon name="plus" size={16} />Ajouter un contact</button>
          </section>
          <section className="prospects-section">
            <div className="section-heading">
              <div><h2>Fichier contacts</h2><span>{filtered.length} contact{filtered.length > 1 ? "s" : ""}</span></div>
              <div className="section-actions">
                <input ref={fileInputRef} className="hidden-file" type="file" accept=".xlsx,.xls,.csv" onChange={importContacts} />
                <button className="file-action" onClick={() => fileInputRef.current?.click()}>Importer Excel</button>
                <button className="file-action" onClick={exportContacts}>Exporter Excel</button>
                <div className="view-toggle">
                  <button className={view === "table" ? "active" : ""} onClick={() => setView("table")}><Icon name="table" size={16} />Tableau</button>
                  <button className={view === "kanban" ? "active" : ""} onClick={() => setView("kanban")}><Icon name="kanban" size={16} />Kanban</button>
                </div>
              </div>
            </div>

            <div className="filters">
              <label className="filter-search"><Icon name="search" size={17} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Rechercher une entreprise, un contact…" /></label>
              <select value={city} onChange={(e) => setCity(e.target.value)}>
                <option>Toutes les villes</option>
                {cities.map((item) => <option key={item}>{item}</option>)}
              </select>
              <select value={industry} onChange={(e) => setIndustry(e.target.value)}>
                <option>Tous les secteurs</option>
                {industries.map((item) => <option key={item}>{item}</option>)}
              </select>
              {(query || city !== "Toutes les villes" || industry !== "Tous les secteurs") && (
                <button className="reset-button" onClick={() => { setQuery(""); setCity("Toutes les villes"); setIndustry("Tous les secteurs"); }}>Réinitialiser</button>
              )}
            </div>

            {view === "table" ? (
              <div className="table-wrap">
                <table>
                  <thead><tr><th>Entreprise</th><th>Contact</th><th>Statut</th><th>Prochaine relance</th><th>Valeur</th><th>Actions</th></tr></thead>
                  <tbody>
                    {filtered.map((prospect) => (
                      <tr key={prospect.id}>
                        <td><div className="company-cell"><Avatar name={prospect.company} /><div><strong>{prospect.company}</strong><span>{prospect.city} · {prospect.industry}</span></div></div></td>
                        <td><div className="contact-cell"><strong>{prospect.contact}</strong><span>{prospect.email}</span></div></td>
                        <td><select className={`status-select ${statusStyle[prospect.status]}`} value={prospect.status} onChange={(e) => changeStatus(prospect.id, e.target.value as Status)}>{statuses.map((s) => <option key={s}>{s}</option>)}</select></td>
                        <td><span className="date-cell"><Icon name="calendar" size={15} />{formatDate(prospect.followUp)}</span></td>
                        <td><strong className="amount">{formatCurrency(prospect.amount)}</strong></td>
                        <td><div className="row-actions"><a href={`tel:${prospect.phone}`} title="Appeler"><Icon name="phone" size={16} /></a><a href={`mailto:${prospect.email}`} title="Envoyer un email"><Icon name="mail" size={16} /></a><button onClick={() => openEdit(prospect)} title="Modifier"><Icon name="edit" size={16} /></button><button className="delete-action" onClick={() => deleteProspect(prospect)} title="Supprimer"><Icon name="trash" size={16} /></button></div></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {!filtered.length && <div className="empty-state"><span><Icon name="search" /></span><h3>Aucun prospect trouvé</h3><p>Essayez de modifier vos filtres de recherche.</p></div>}
              </div>
            ) : (
              <div className="kanban-board">
                {statuses.map((status) => {
                  const cards = filtered.filter((p) => p.status === status);
                  return (
                    <div className="kanban-column" key={status}>
                      <div className="kanban-title"><StatusBadge status={status} /><span>{cards.length}</span></div>
                      <div className="kanban-list">
                        {cards.map((prospect) => (
                          <article className="kanban-card" key={prospect.id}>
                            <div className="kanban-card-top"><span>{prospect.industry}</span><button onClick={() => openEdit(prospect)}><Icon name="more" size={17} /></button></div>
                            <h3>{prospect.company}</h3>
                            <p><Avatar name={prospect.contact} small />{prospect.contact}</p>
                            <div className="kanban-date"><Icon name="calendar" size={14} />{formatDate(prospect.followUp)}</div>
                            <div className="kanban-footer"><strong>{formatCurrency(prospect.amount)}</strong><select value={prospect.status} onChange={(e) => changeStatus(prospect.id, e.target.value as Status)} aria-label="Changer le statut">{statuses.map((s) => <option key={s}>{s}</option>)}</select></div>
                          </article>
                        ))}
                        {!cards.length && <div className="kanban-empty">Déposez une opportunité ici</div>}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </section>
            </>
          )}

          {activePage === "admin" && (
            <section className="admin-page">
              <div className="page-heading">
                <div>
                  <p className="eyebrow">ESPACE ADMINISTRATEUR</p>
                  <h1>Administration des contacts</h1>
                  <p>Ajoutez en une fois tous les numéros que votre équipe devra contacter.</p>
                </div>
                <span className="admin-badge"><Icon name="settings" size={15} />Administrateur</span>
              </div>

              <div className="admin-stats">
                <article><span className="metric-icon tone-blue"><Icon name="contacts" /></span><div><strong>{prospects.length}</strong><p>Contacts enregistrés</p></div></article>
                <article><span className="metric-icon tone-amber"><Icon name="phone" /></span><div><strong>{prospects.filter((p) => p.status === "À contacter").length}</strong><p>Numéros à traiter</p></div></article>
                <article><span className="metric-icon tone-green"><Icon name="trophy" /></span><div><strong>{contacted}</strong><p>Contacts déjà traités</p></div></article>
              </div>

              <div className="admin-grid">
                <article className="import-card">
                  <span className="import-icon"><Icon name="table" size={27} /></span>
                  <p className="eyebrow">IMPORT EN MASSE</p>
                  <h2>Importer votre fichier Excel</h2>
                  <p>Chargez un fichier `.xlsx`, `.xls` ou `.csv`. Chaque ligne deviendra un contact avec le statut « À contacter » si aucun statut n'est indiqué.</p>
                  <input ref={fileInputRef} className="hidden-file" type="file" accept=".xlsx,.xls,.csv" onChange={importContacts} />
                  <button className="admin-import" onClick={() => fileInputRef.current?.click()}><Icon name="plus" />Choisir un fichier Excel</button>
                  <button className="template-button" onClick={downloadTemplate}>Télécharger le modèle Webora</button>
                </article>

                <article className="format-card">
                  <div className="format-heading"><span><Icon name="table" /></span><div><h2>Colonnes reconnues</h2><p>La première ligne du fichier doit contenir les titres.</p></div></div>
                  <div className="column-list">
                    {[
                      ["Téléphone", "Obligatoire pour appeler"],
                      ["Entreprise", "Nom de la société"],
                      ["Contact", "Prénom et nom"],
                      ["Ville", "Localisation"],
                      ["Secteur", "Activité"],
                      ["Email", "Adresse électronique"],
                      ["Statut", "À contacter par défaut"],
                      ["Notes", "Informations utiles"],
                    ].map(([label, detail], index) => (
                      <div key={label}><span>{index + 1}</span><strong>{label}</strong><small>{detail}</small></div>
                    ))}
                  </div>
                </article>
              </div>

              <div className="admin-tip">
                <span><Icon name="sparkles" /></span>
                <div><strong>Conseil Webora</strong><p>Commencez par télécharger notre modèle Excel. Remplissez une ligne par numéro puis importez-le ici. Les contacts seront immédiatement visibles dans « Fichier contacts » et « Pipeline ».</p></div>
              </div>
            </section>
          )}
        </div>
      </main>

      <nav className="mobile-bottom-nav" aria-label="Navigation principale">
        <button className={activePage === "dashboard" ? "active" : ""} onClick={() => setActivePage("dashboard")}>
          <Icon name="dashboard" size={19} />
          <span>Accueil</span>
        </button>
        <button className={activePage === "contacts" ? "active" : ""} onClick={() => { setActivePage("contacts"); setView("table"); }}>
          <Icon name="contacts" size={19} />
          <span>Contacts</span>
        </button>
        <button className={activePage === "pipeline" ? "active" : ""} onClick={() => { setActivePage("pipeline"); setView("kanban"); }}>
          <Icon name="pipeline" size={19} />
          <span>Pipeline</span>
        </button>
        <button className={activePage === "admin" ? "active" : ""} onClick={() => setActivePage("admin")}>
          <Icon name="settings" size={19} />
          <span>Admin</span>
        </button>
      </nav>

      {modalOpen && (
        <div className="modal-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) setModalOpen(false); }}>
          <div className="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title">
            <div className="modal-header">
              <div><span className="modal-icon"><Icon name={editing ? "edit" : "plus"} /></span><div><h2 id="modal-title">{editing ? "Modifier le prospect" : "Nouveau prospect"}</h2><p>{editing ? "Mettez à jour les informations et le suivi." : "Ajoutez une nouvelle opportunité à votre pipeline."}</p></div></div>
              <button className="modal-close" onClick={() => setModalOpen(false)} aria-label="Fermer"><Icon name="close" /></button>
            </div>
            <form onSubmit={saveProspect}>
              <div className="modal-body">
                <div className="form-section">
                  <h3>Informations générales</h3>
                  <div className="form-grid">
                    <label><span>Nom de l'entreprise *</span><input required value={draft.company} onChange={(e) => setDraft({ ...draft, company: e.target.value })} placeholder="Ex. Acme Studio" /></label>
                    <label><span>Secteur d'activité *</span><input required value={draft.industry} onChange={(e) => setDraft({ ...draft, industry: e.target.value })} placeholder="Ex. Conseil" /></label>
                    <label><span>Ville *</span><input required value={draft.city} onChange={(e) => setDraft({ ...draft, city: e.target.value })} placeholder="Ex. Casablanca" /></label>
                    <label><span>Nom du contact *</span><input required value={draft.contact} onChange={(e) => setDraft({ ...draft, contact: e.target.value })} placeholder="Prénom Nom" /></label>
                    <label><span>Email *</span><input required type="email" value={draft.email} onChange={(e) => setDraft({ ...draft, email: e.target.value })} placeholder="contact@entreprise.fr" /></label>
                    <label><span>Téléphone</span><input value={draft.phone} onChange={(e) => setDraft({ ...draft, phone: e.target.value })} placeholder="+212 6 00 00 00 00" /></label>
                  </div>
                </div>
                <div className="form-section">
                  <h3>Suivi commercial</h3>
                  <div className="form-grid form-grid-three">
                    <label><span>Statut actuel</span><select value={draft.status} onChange={(e) => setDraft({ ...draft, status: e.target.value as Status })}>{statuses.map((s) => <option key={s}>{s}</option>)}</select></label>
                    <label><span>Prochaine relance</span><input type="date" value={draft.followUp} onChange={(e) => setDraft({ ...draft, followUp: e.target.value })} /></label>
                    <label><span>Montant estimé (DH)</span><input type="number" min="0" value={draft.amount || ""} onChange={(e) => setDraft({ ...draft, amount: Number(e.target.value) })} placeholder="0" /></label>
                  </div>
                  <label className="notes-field"><span>Ajouter une note</span><textarea value={newNote} onChange={(e) => setNewNote(e.target.value)} placeholder="Compte-rendu d'appel, besoins, prochaines étapes…" rows={3} /></label>
                </div>
                {editing && draft.notes.length > 0 && (
                  <div className="form-section history">
                    <h3>Historique des notes <span>{draft.notes.length}</span></h3>
                    {draft.notes.map((note) => (
                      <div className="history-item" key={note.id}>
                        <span className="history-line" />
                        <Avatar name="Alex Morgan" small />
                        <div><p>{note.text}</p><time>{new Intl.DateTimeFormat("fr-MA", { dateStyle: "medium", timeStyle: "short" }).format(new Date(note.date))}</time></div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div className="modal-footer">{editing && <button type="button" className="danger-button" onClick={() => deleteProspect(editing)}><Icon name="trash" size={15} />Supprimer</button>}<button type="button" className="secondary-button" onClick={() => setModalOpen(false)}>Annuler</button><button type="submit" className="add-button">{editing ? "Enregistrer les modifications" : "Ajouter le prospect"}</button></div>
            </form>
          </div>
        </div>
      )}

      {toast && <div className="toast"><span>✓</span>{toast}</div>}
    </div>
  );
}
