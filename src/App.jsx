import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Building2, 
  FileText, 
  Users, 
  AlertTriangle, 
  CheckCircle2, 
  XCircle, 
  Search, 
  ExternalLink, 
  LogOut, 
  Plus, 
  ShieldCheck, 
  FileSpreadsheet, 
  Trash2, 
  Link as LinkIcon,
  HardHat,
  Bell,
  ArrowLeft,
  ChevronRight,
  FileSignature,
  Check,
  X,
  UserCheck,
  RefreshCw
} from 'lucide-react';
import * as XLSX from 'xlsx';

// ==========================================
// CONFIGURACIÓN DE GOOGLE SHEETS / DRIVE
// ==========================================
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxNaUJqxU9M_cik1AqlSVQw7lfizQziZo3qbNggh1z6ydmemTe-jLLlpxYx4nuO19U/exec";

// --- DATOS INICIALES ---

const INITIAL_USERS = [
  { id: '1', email: 'neuralprl', code: 'Neuralprl@', name: 'Superadministrador', role: 'superadmin', assignedCentres: ['ALL'], company: 'Neural PRL' },
  { id: '2', email: 'director.madrid@neural.es', code: 'Pass1234@', name: 'Carlos (Director Madrid)', role: 'corporativo', assignedCentres: ['c1'], company: 'Neural SRL' },
  { id: '3', email: 'prevencion@contratasvalencia.com', code: 'Externa123@', name: 'Mantenimientos Levante SL', role: 'externo', assignedCentres: ['c2'], company: 'Mantenimientos Levante SL' }
];

const INITIAL_GENERAL_DOCS = [
  { id: 'gd1', title: 'Procedimiento General de Evacuación v2', category: 'Procedimientos', link: 'https://sharepoint.com/doc1', date: '2026-01-15' },
  { id: 'gd2', title: 'Protocolo de Actuación Accidentes Laborales', category: 'Protocolos', link: 'https://sharepoint.com/doc2', date: '2026-02-01' },
  { id: 'gd3', title: 'Plantilla de Inspección de Equipos de Protección', category: 'Plantillas', link: 'https://sharepoint.com/doc3', date: '2026-02-10' },
];

const INITIAL_CENTRES = [
  {
    id: 'c1',
    name: 'Centro Neural Madrid - Castellana',
    zone: 'Madrid Norte',
    users: ['neuralprl', 'director.madrid@neural.es'],
    docs: {
      evaluacion_riesgos: [{ id: 'd1', name: 'Evaluacion_Riesgos_2026_Madrid.pdf', link: 'https://sharepoint.com/eval-madrid.pdf' }],
      informacion_riesgos: [{ id: 'd2', name: 'Info_Riesgos_Puestos_Madrid.pdf', link: 'https://sharepoint.com/info-madrid.pdf' }],
      medidas_emergencia: []
    }
  },
  {
    id: 'c2',
    name: 'Centro Neural Valencia - Mestalla',
    zone: 'Comunidad Valenciana',
    users: ['neuralprl', 'prevencion@contratasvalencia.com'],
    docs: {
      evaluacion_riesgos: [{ id: 'd3', name: 'Evaluacion_Riesgos_Mestalla_v1.pdf', link: 'https://sharepoint.com/eval-valencia.pdf' }],
      informacion_riesgos: [],
      medidas_emergencia: [{ id: 'd4', name: 'Plan_Emergencia_Valencia_2026.pdf', link: 'https://sharepoint.com/emerg-valencia.pdf' }]
    }
  }
];

const INITIAL_CAE_RECORDS = [
  {
    id: 'cae_1',
    centreId: 'c2',
    companyName: 'Mantenimientos Levante SL',
    userEmail: 'prevencion@contratasvalencia.com',
    companyDocs: { prl: true, er: true, sp: false },
    updatedAt: '2026-02-20',
    workers: [
      { id: 'w1', name: 'Juan Pérez Gómez', dni: '12345678A', signedAt: null, signatureImage: null, checks: { epis: true, inf: true, for: true, vs: true } },
      { id: 'w2', name: 'María López Sanchis', dni: '87654321B', signedAt: null, signatureImage: null, checks: { epis: true, inf: true, for: false, vs: true } },
      { id: 'w3', name: 'Carlos Ruiz Delgado', dni: '45678912C', signedAt: null, signatureImage: null, checks: { epis: false, inf: false, for: false, vs: false } }
    ]
  },
  {
    id: 'cae_2',
    centreId: 'c1',
    companyName: 'Construcciones e Instalaciones Norte SA',
    userEmail: 'obras@nortesa.com',
    companyDocs: { prl: true, er: true, sp: true },
    updatedAt: '2026-02-22',
    workers: [
      { id: 'w4', name: 'Antonio García Vidal', dni: '11223344D', signedAt: null, signatureImage: null, checks: { epis: true, inf: true, for: true, vs: true } }
    ]
  }
];

const extractFileNameFromUrl = (url) => {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    const pathname = parsed.pathname;
    const filename = pathname.split('/').pop();
    if (filename && filename.length > 0) return decodeURIComponent(filename);
  } catch (e) {
    const parts = url.split('/');
    const last = parts.pop() || parts.pop();
    if (last) return decodeURIComponent(last.split('?')[0]);
  }
  return 'Documento SharePoint';
};

const isWorkerFullyApproved = (worker) => {
  if (!worker) return false;
  return !!worker.signedAt || (worker.checks && worker.checks.epis && worker.checks.inf && worker.checks.for && worker.checks.vs);
};

const isCompanyFullyApproved = (record) => {
  if (!record) return false;
  const companyOk = record.companyDocs?.prl && record.companyDocs?.er && record.companyDocs?.sp;
  const workers = record.workers || [];
  const allWorkersOk = workers.length > 0 && workers.every(isWorkerFullyApproved);
  return companyOk && allWorkersOk;
};

// ==========================================
// COMPONENTE: MODAL DE FIRMA INTERACTIVA
// ==========================================
function SignatureModal({ worker, companyName, onClose, onSaveSignature }) {
  const canvasRef = useRef(null);
  const [isDrawing, setIsDrawing] = useState(false);
  const [hasSignature, setHasSignature] = useState(false);

  const startDrawing = (e) => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    const x = (e.touches ? e.touches[0].clientX : e.clientX) - rect.left;
    const y = (e.touches ? e.touches[0].clientY : e.clientY) - rect.top;

    ctx.beginPath();
    ctx.moveTo(x, y);
    setIsDrawing(true);
  };

  const draw = (e) => {
    if (!isDrawing) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    const rect = canvas.getBoundingClientRect();
    const x = (e.touches ? e.touches[0].clientX : e.clientX) - rect.left;
    const y = (e.touches ? e.touches[0].clientY : e.clientY) - rect.top;

    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.strokeStyle = '#0f172a';
    ctx.lineTo(x, y);
    ctx.stroke();
    setHasSignature(true);
  };

  const stopDrawing = () => {
    setIsDrawing(false);
  };

  const clearCanvas = () => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasSignature(false);
  };

  const handleConfirm = () => {
    if (!hasSignature) {
      alert('Por favor, realiza la firma antes de confirmar.');
      return;
    }
    const canvas = canvasRef.current;
    const signatureDataUrl = canvas.toDataURL('image/png');
    onSaveSignature(worker.id, signatureDataUrl);
    onClose();
  };

  return (
    <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-xl shadow-2xl max-w-2xl w-full p-6 space-y-4 max-h-[90vh] flex flex-col border border-slate-200">
        <div className="flex justify-between items-center border-b pb-3">
          <div>
            <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
              <FileSignature className="w-5 h-5 text-blue-600" />
              Firma de Documento de PRL y Recepción de EPIs
            </h3>
            <p className="text-xs text-slate-500">Trabajador: <strong>{worker.name}</strong> ({worker.dni})</p>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 p-1">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto bg-slate-50 p-4 rounded-lg border border-slate-200 text-xs text-slate-700 space-y-3 leading-relaxed font-sans">
          <div className="text-center font-bold text-sm text-slate-900 uppercase border-b pb-2">
            DECLARACIÓN JURADA Y RECEPCIÓN DE INFORMACIÓN / FORMACIÓN / EPIs
          </div>
          <p>
            Mediante la firma del presente documento, el/la trabajador/a <strong>{worker.name}</strong>, con DNI <strong>{worker.dni}</strong>, empleado/a de la empresa <strong>{companyName}</strong>, declara formalmente:
          </p>
          <ul className="list-disc pl-5 space-y-1 text-slate-600">
            <li>Haber recibido la <strong>Información sobre los Riesgos Laborales (INF)</strong> específicos de su puesto de trabajo.</li>
            <li>Haber recibido la <strong>Formación en PRL (FOR)</strong> acorde a la normativa vigente.</li>
            <li>Haber hecho uso del derecho o reconocimiento de la <strong>Vigilancia de la Salud (VS)</strong> de forma apta.</li>
            <li>Haber recibido los <strong>Equipos de Protección Individual (EPIs)</strong> obligatorios para su actividad.</li>
          </ul>
          <p className="text-[11px] text-slate-500 italic pt-2">
            Fecha de emisión y firma electrónica: {new Date().toLocaleDateString('es-ES')} - {new Date().toLocaleTimeString('es-ES')}
          </p>
        </div>

        <div className="space-y-2">
          <div className="flex justify-between items-center">
            <label className="text-xs font-bold text-slate-700">Firme dentro del recuadro:</label>
            <button type="button" onClick={clearCanvas} className="text-[11px] text-rose-600 hover:underline font-semibold">
              Borrar firma
            </button>
          </div>

          <div className="border-2 border-dashed border-slate-300 rounded-lg bg-white overflow-hidden touch-none">
            <canvas 
              ref={canvasRef}
              width={550}
              height={140}
              className="w-full cursor-crosshair"
              onMouseDown={startDrawing}
              onMouseMove={draw}
              onMouseUp={stopDrawing}
              onMouseLeave={stopDrawing}
              onTouchStart={startDrawing}
              onTouchMove={draw}
              onTouchEnd={stopDrawing}
            />
          </div>
        </div>

        <div className="flex justify-end space-x-3 pt-2 border-t">
          <button onClick={onClose} className="px-4 py-2 text-xs font-semibold text-slate-600 hover:bg-slate-100 rounded-lg transition">
            Cancelar
          </button>
          <button onClick={handleConfirm} className="px-5 py-2 text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg shadow-sm transition flex items-center gap-1.5">
            <Check className="w-4 h-4" />
            Guardar Firma y Validar
          </button>
        </div>
      </div>
    </div>
  );
}

// ==========================================
// COMPONENTE PRINCIPAL (APP)
// ==========================================
export default function App() {
  // Autenticación
  const [currentUser, setCurrentUser] = useState(null);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');

  // Estado Global
  const [users, setUsers] = useState(INITIAL_USERS);
  const [centres, setCentres] = useState(INITIAL_CENTRES);
  const [generalDocs, setGeneralDocs] = useState(INITIAL_GENERAL_DOCS);
  const [caeRecords, setCaeRecords] = useState(INITIAL_CAE_RECORDS);
  const [syncLoading, setSyncLoading] = useState(false);

  // Navegación y Filtros
  const [activeTab, setActiveTab] = useState('cae');
  const [selectedCentreId, setSelectedCentreId] = useState(null);
  const [selectedCaeCompanyId, setSelectedCaeCompanyId] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [caeSearchTerm, setCaeSearchTerm] = useState('');

  // Modales y Formularios
  const [isAddingProcedure, setIsAddingProcedure] = useState(false);
  const [newProcTitle, setNewProcTitle] = useState('');
  const [newProcCategory, setNewProcCategory] = useState('Procedimientos');
  const [newProcLink, setNewProcLink] = useState('');

  const [isAddingWorker, setIsAddingWorker] = useState(false);
  const [newWorkerName, setNewWorkerName] = useState('');
  const [newWorkerDni, setNewWorkerDni] = useState('');

  const [signingWorker, setSigningWorker] = useState(null);

  const [editDocModal, setEditDocModal] = useState({ open: false, centreId: null, categoryKey: null, categoryLabel: '' });
  const [newLinkUrl, setNewLinkUrl] = useState('');
  const [newLinkName, setNewLinkName] = useState('');

  // Sincronización con Google Sheets (Google Drive)
  const syncWithGoogleSheets = async () => {
    setSyncLoading(true);
    try {
      const response = await fetch(GOOGLE_SCRIPT_URL);
      const data = await response.json();
      if (Array.isArray(data) && data.length > 0) {
        setCaeRecords(prev => {
          return prev.map(record => {
            const remoteComp = data.find(c => c.id === record.id || c.companyName === record.companyName);
            if (remoteComp && remoteComp.workers) {
              return { ...record, workers: remoteComp.workers };
            }
            return record;
          });
        });
      }
    } catch (error) {
      console.error("Error sincronizando con Google Sheets:", error);
    } finally {
      setSyncLoading(false);
    }
  };

  useEffect(() => {
    syncWithGoogleSheets();
  }, []);

  // Manejo de Login
  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');
    const user = users.find(
      u => u.email.trim().toLowerCase() === loginEmail.trim().toLowerCase() && 
           u.code.trim() === loginCode.trim()
    );

    if (user) {
      setCurrentUser(user);
      setActiveTab('cae');
    } else {
      setLoginError('Usuario o contraseña incorrectos.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setLoginEmail('');
    setLoginCode('');
  };

  // Centros accesibles
  const accessibleCentres = useMemo(() => {
    if (!currentUser) return [];
    if (currentUser.role === 'superadmin') return centres;
    return centres.filter(c => c.users.includes(currentUser.email));
  }, [currentUser, centres]);

  const selectedCentreRecord = useMemo(() => {
    if (!selectedCentreId) return null;
    return centres.find(c => c.id === selectedCentreId);
  }, [centres, selectedCentreId]);

  const selectedCaeRecord = useMemo(() => {
    if (!selectedCaeCompanyId) return null;
    return caeRecords.find(r => r.id === selectedCaeCompanyId);
  }, [caeRecords, selectedCaeCompanyId]);

  const filteredCaeRecords = useMemo(() => {
    if (!caeSearchTerm.trim()) return caeRecords;
    return caeRecords.filter(r => 
      r.companyName.toLowerCase().includes(caeSearchTerm.toLowerCase()) ||
      r.userEmail.toLowerCase().includes(caeSearchTerm.toLowerCase())
    );
  }, [caeRecords, caeSearchTerm]);

  const filteredCentres = useMemo(() => {
    if (!searchTerm.trim()) return accessibleCentres;
    return accessibleCentres.filter(c => 
      c.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      c.zone.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [accessibleCentres, searchTerm]);

  const handleAddProcedure = (e) => {
    e.preventDefault();
    if (!newProcTitle.trim() || !newProcLink.trim()) return;

    const newProc = {
      id: `gd_${Date.now()}`,
      title: newProcTitle.trim(),
      category: newProcCategory,
      link: newProcLink.trim(),
      date: new Date().toISOString().split('T')[0]
    };

    setGeneralDocs(prev => [newProc, ...prev]);
    setNewProcTitle('');
    setNewProcLink('');
    setIsAddingProcedure(false);
  };

  const handleToggleCompanyDoc = (recordId, docType) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        return {
          ...r,
          companyDocs: {
            ...r.companyDocs,
            [docType]: !r.companyDocs?.[docType]
          }
        };
      }
      return r;
    }));
  };

  const handleAddWorker = (recordId) => {
    if (!newWorkerName.trim() || !newWorkerDni.trim()) return;

    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        const newWorker = {
          id: `w_${Date.now()}`,
          name: newWorkerName.trim(),
          dni: newWorkerDni.trim(),
          signedAt: null,
          signatureImage: null,
          checks: { epis: false, inf: false, for: false, vs: false }
        };
        return { ...r, workers: [...(r.workers || []), newWorker] };
      }
      return r;
    }));

    setNewWorkerName('');
    setNewWorkerDni('');
    setIsAddingWorker(false);
  };

  const handleDeleteWorker = (recordId, workerId) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        return { ...r, workers: (r.workers || []).filter(w => w.id !== workerId) };
      }
      return r;
    }));
  };

  // Guardar Firma e Integrar con Google Sheets / Drive
  const handleSaveWorkerSignature = async (workerId, signatureDataUrl) => {
    const signedAtString = new Date().toLocaleString('es-ES');

    // 1. Guardar localmente
    setCaeRecords(prev => prev.map(r => {
      if (r.id === selectedCaeCompanyId) {
        const updatedWorkers = (r.workers || []).map(w => {
          if (w.id === workerId) {
            return {
              ...w,
              signedAt: signedAtString,
              signatureImage: signatureDataUrl,
              checks: { epis: true, inf: true, for: true, vs: true }
            };
          }
          return w;
        });
        return { ...r, workers: updatedWorkers };
      }
      return r;
    }));

    // 2. Transmitir a Google Apps Script
    try {
      await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          workerId: workerId,
          signatureImage: signatureDataUrl,
          signedAt: signedAtString
        })
      });
    } catch (error) {
      console.error("Error guardando en Google Sheets:", error);
    }
  };

  // Carga Masiva Excel
  const handleFileUploadCentres = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const rows = data.slice(1);
      const newCentres = [...centres];
      const newUsers = [...users];

      rows.forEach((row, idx) => {
        if (!row[0]) return;

        const centreName = row[0].toString().trim();
        const zone = row[1] ? row[1].toString().trim() : 'General';
        const assignedUserEmails = [];

        const userCols = [
          { emailIdx: 2, codeIdx: 3 },
          { emailIdx: 4, codeIdx: 5 },
          { emailIdx: 6, codeIdx: 7 },
          { emailIdx: 8, codeIdx: 9 },
          { emailIdx: 10, codeIdx: 11 }
        ];

        userCols.forEach(({ emailIdx, codeIdx }) => {
          if (row[emailIdx] && row[codeIdx]) {
            const email = row[emailIdx].toString().trim();
            const code = row[codeIdx].toString().trim();
            assignedUserEmails.push(email);

            const exists = newUsers.some(u => u.email.toLowerCase() === email.toLowerCase());
            if (!exists) {
              newUsers.push({
                id: `u_${Date.now()}_${Math.random()}`,
                email: email,
                code: code,
                name: email.includes('@') ? email.split('@')[0] : email,
                role: 'corporativo',
                assignedCentres: []
              });
            }
          }
        });

        const existingIndex = newCentres.findIndex(c => c.name.toLowerCase() === centreName.toLowerCase());
        const centreObj = {
          id: existingIndex >= 0 ? newCentres[existingIndex].id : `c_${Date.now()}_${idx}`,
          name: centreName,
          zone: zone,
          users: assignedUserEmails,
          docs: existingIndex >= 0 ? newCentres[existingIndex].docs : {
            evaluacion_riesgos: [],
            informacion_riesgos: [],
            medidas_emergencia: []
          }
        };

        if (existingIndex >= 0) {
          newCentres[existingIndex] = centreObj;
        } else {
          newCentres.push(centreObj);
        }
      });

      setCentres(newCentres);
      setUsers(newUsers);
      alert('¡Centros y usuarios importados correctamente!');
    };
    reader.readAsBinaryString(file);
  };

  const handleUrlChange = (url) => {
    setNewLinkUrl(url);
    if (url.trim() && !newLinkName) {
      setNewLinkName(extractFileNameFromUrl(url));
    }
  };

  const handleAddLink = (e) => {
    e.preventDefault();
    if (!newLinkUrl.trim() || !editDocModal.centreId || !editDocModal.categoryKey) return;

    const fileName = newLinkName.trim() || extractFileNameFromUrl(newLinkUrl);
    const newDocObj = {
      id: `doc_${Date.now()}`,
      name: fileName,
      link: newLinkUrl.trim()
    };

    setCentres(prevCentres => prevCentres.map(c => {
      if (c.id === editDocModal.centreId) {
        const currentList = Array.isArray(c.docs[editDocModal.categoryKey]) 
          ? c.docs[editDocModal.categoryKey] 
          : [];
        return {
          ...c,
          docs: {
            ...c.docs,
            [editDocModal.categoryKey]: [...currentList, newDocObj]
          }
        };
      }
      return c;
    }));

    setNewLinkUrl('');
    setNewLinkName('');
  };

  const handleDeleteLink = (docId) => {
    setCentres(prevCentres => prevCentres.map(c => {
      if (c.id === editDocModal.centreId) {
        const currentList = Array.isArray(c.docs[editDocModal.categoryKey]) 
          ? c.docs[editDocModal.categoryKey] 
          : [];
        return {
          ...c,
          docs: {
            ...c.docs,
            [editDocModal.categoryKey]: currentList.filter(d => d.id !== docId)
          }
        };
      }
      return c;
    }));
  };

  // --- VISTA DE LOGIN ---
  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md w-full space-y-6">
          <div className="text-center space-y-2">
            <div className="bg-blue-600 text-white w-12 h-12 rounded-lg flex items-center justify-center mx-auto shadow-lg">
              <ShieldCheck className="w-8 h-8" />
            </div>
            <h1 className="text-2xl font-bold text-slate-800">Plataforma PRL & CAE</h1>
            <p className="text-sm text-slate-500">Gestión Documental y Coordinación Empresarial</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Usuario / Correo</label>
              <input 
                type="text" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm"
                placeholder="neuralprl"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Contraseña</label>
              <input 
                type="password" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm"
                placeholder="••••••••"
                value={loginCode}
                onChange={(e) => setLoginCode(e.target.value)}
              />
            </div>

            {loginError && (
              <div className="p-3 bg-red-50 border border-red-200 text-red-600 text-xs rounded-lg flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <button 
              type="submit" 
              className="w-full py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition shadow-md text-sm"
            >
              Iniciar Sesión
            </button>
          </form>

          <div className="text-xs text-center text-slate-400 border-t pt-4">
            Credenciales SA: <code className="bg-slate-100 px-1 rounded text-slate-700">neuralprl</code> / <code className="bg-slate-100 px-1 rounded text-slate-700">Neuralprl@</code>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans">
      {/* Header */}
      <header className="bg-slate-800 text-white shadow-md sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <ShieldCheck className="w-7 h-7 text-blue-400" />
            <div>
              <h1 className="font-bold text-lg leading-tight">Neural PRL</h1>
              <span className="text-xs text-slate-400">Coordinación de Actividades Empresariales</span>
            </div>
          </div>

          <div className="flex items-center space-x-4">
            {currentUser.role === 'superadmin' && (
              <button className="p-2 text-slate-300 hover:text-white rounded-lg relative">
                <Bell className="w-5 h-5" />
              </button>
            )}

            <div className="text-right hidden sm:block">
              <p className="text-sm font-medium">{currentUser.name}</p>
              <div className="flex items-center justify-end gap-1">
                <span className={`text-[10px] px-2 py-0.5 rounded font-semibold uppercase ${
                  currentUser.role === 'superadmin' ? 'bg-purple-900 text-purple-200' :
                  currentUser.role === 'corporativo' ? 'bg-blue-900 text-blue-200' : 'bg-emerald-900 text-emerald-200'
                }`}>
                  {currentUser.role === 'superadmin' ? 'SUPERADMIN (SA)' :
                   currentUser.role === 'corporativo' ? 'Usuario Corporativo (UC)' : 'Empresa Externa (UX)'}
                </span>
              </div>
            </div>

            <button 
              onClick={handleLogout} 
              className="p-2 text-slate-300 hover:text-white hover:bg-slate-700 rounded-lg transition"
              title="Cerrar Sesión"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 py-6 flex-1 w-full flex flex-col md:flex-row gap-6">
        
        {/* NAVEGACIÓN LATERAL */}
        <aside className="w-full md:w-64 bg-white rounded-xl border border-slate-200 shadow-sm p-4 shrink-0 h-fit space-y-6">
          
          <div className="space-y-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2 block">
              PROCEDIMIENTOS GENERALES
            </span>
            
            <button 
              onClick={() => setActiveTab('general')}
              className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'general' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
            >
              <FileText className="w-4 h-4 text-blue-500" />
              <span>Doc. General PRL</span>
            </button>
          </div>

          <div className="pt-4 border-t border-slate-100 space-y-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2 block">
              MÓDULOS ESPECIALES
            </span>
            
            <button 
              onClick={() => {
                setActiveTab('cae');
                setSelectedCaeCompanyId(null);
              }}
              className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center justify-between transition ${activeTab === 'cae' ? 'bg-amber-500 text-white font-bold shadow-md' : 'text-slate-700 hover:bg-slate-50'}`}
            >
              <div className="flex items-center space-x-2">
                <HardHat className={`w-4 h-4 ${activeTab === 'cae' ? 'text-white' : 'text-amber-500'}`} />
                <span>Módulo CAE- Empresas</span>
              </div>
            </button>

            <button 
              onClick={() => {
                setActiveTab('centres');
                setSelectedCentreId(null);
              }}
              className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center justify-between transition ${activeTab === 'centres' ? 'bg-blue-600 text-white font-bold shadow-md' : 'text-slate-700 hover:bg-slate-50'}`}
            >
              <div className="flex items-center space-x-2">
                <Building2 className={`w-4 h-4 ${activeTab === 'centres' ? 'text-white' : 'text-blue-500'}`} />
                <span>Módulo Centros de Trabajo</span>
              </div>
            </button>
          </div>

          {currentUser.role === 'superadmin' && (
            <div className="pt-4 border-t border-slate-100 space-y-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2 block">Administración SA</span>
              
              <button 
                onClick={() => setActiveTab('excel')}
                className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'excel' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>Carga Masiva Excel</span>
              </button>

              <button 
                onClick={() => setActiveTab('users')}
                className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'users' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
              >
                <Users className="w-4 h-4" />
                <span>Gestión Perfiles</span>
              </button>
            </div>
          )}
        </aside>

        {/* CONTENIDO PRINCIPAL */}
        <main className="flex-1">
          
          {/* TAB: MÓDULO CAE CON GOOGLE DRIVE */}
          {activeTab === 'cae' && (
            <div className="space-y-6">
              {!selectedCaeRecord ? (
                <div className="space-y-6">
                  <div className="bg-amber-600 text-white p-6 rounded-xl shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                    <div>
                      <div className="flex items-center space-x-2">
                        <HardHat className="w-6 h-6 text-amber-200" />
                        <h2 className="text-xl font-bold">Módulo CAE - Empresas y Trabajadores</h2>
                      </div>
                      <p className="text-xs text-amber-100 mt-1">
                        Sincronización en tiempo real con Google Sheets y Google Drive.
                      </p>
                    </div>

                    <button 
                      onClick={syncWithGoogleSheets}
                      disabled={syncLoading}
                      className="px-3 py-2 bg-amber-700 hover:bg-amber-800 text-white text-xs font-bold rounded-lg flex items-center gap-2 shadow transition"
                    >
                      <RefreshCw className={`w-4 h-4 ${syncLoading ? 'animate-spin' : ''}`} />
                      Sincronizar Google Drive
                    </button>
                  </div>

                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex justify-between items-center gap-4">
                    <div className="relative w-full sm:w-80">
                      <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                      <input 
                        type="text"
                        placeholder="Buscar empresa o email..."
                        className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-amber-500"
                        value={caeSearchTerm}
                        onChange={(e) => setCaeSearchTerm(e.target.value)}
                      />
                    </div>
                  </div>

                  <div className="flex flex-col space-y-3">
                    {filteredCaeRecords.map((record) => {
                      const approved = isCompanyFullyApproved(record);
                      return (
                        <div 
                          key={record.id}
                          onClick={() => setSelectedCaeCompanyId(record.id)}
                          className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:shadow-md hover:border-amber-300 transition cursor-pointer flex items-center justify-between group"
                        >
                          <div className="space-y-1">
                            <div className="flex items-center space-x-3">
                              <h3 className="text-base font-bold text-slate-800 group-hover:text-amber-600 transition">
                                {record.companyName}
                              </h3>
                              {approved ? (
                                <span className="px-2 py-0.5 bg-emerald-100 text-emerald-700 text-[10px] font-bold rounded-full flex items-center gap-1">
                                  <CheckCircle2 className="w-3 h-3" /> Apto CAE
                                </span>
                              ) : (
                                <span className="px-2 py-0.5 bg-rose-100 text-rose-700 text-[10px] font-bold rounded-full flex items-center gap-1">
                                  <XCircle className="w-3 h-3" /> Pendiente
                                </span>
                              )}
                            </div>
                            <p className="text-xs text-slate-400">Email: {record.userEmail}</p>
                          </div>

                          <div className="flex items-center space-x-3">
                            <span className="text-xs font-semibold text-slate-500">
                              {record.workers?.length || 0} Trabajadores
                            </span>
                            <ChevronRight className="w-5 h-5 text-slate-400 group-hover:text-amber-500 group-hover:translate-x-1 transition" />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              ) : (
                /* DETALLE DE EMPRESA SELECCIONADA EN CAE */
                <div className="space-y-6">
                  <div className="flex items-center justify-between">
                    <button 
                      onClick={() => setSelectedCaeCompanyId(null)}
                      className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-600 hover:text-slate-900 bg-white px-3 py-2 rounded-lg border shadow-sm transition"
                    >
                      <ArrowLeft className="w-4 h-4" />
                      <span>Volver al listado de empresas</span>
                    </button>

                    <button 
                      onClick={syncWithGoogleSheets}
                      disabled={syncLoading}
                      className="px-3 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-semibold rounded-lg flex items-center gap-2 transition"
                    >
                      <RefreshCw className={`w-4 h-4 ${syncLoading ? 'animate-spin' : ''}`} />
                      Actualizar
                    </button>
                  </div>

                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-6">
                    <div className="flex justify-between items-start border-b pb-4">
                      <div>
                        <h2 className="text-xl font-bold text-slate-800">{selectedCaeRecord.companyName}</h2>
                        <p className="text-xs text-slate-500">Email asignado: {selectedCaeRecord.userEmail}</p>
                      </div>

                      {isCompanyFullyApproved(selectedCaeRecord) ? (
                        <span className="px-3 py-1 bg-emerald-100 text-emerald-800 font-bold text-xs rounded-full flex items-center gap-1">
                          <CheckCircle2 className="w-4 h-4" /> Empresa Homologada
                        </span>
                      ) : (
                        <span className="px-3 py-1 bg-rose-100 text-rose-800 font-bold text-xs rounded-full flex items-center gap-1">
                          <AlertTriangle className="w-4 h-4" /> Documentación Incompleta
                        </span>
                      )}
                    </div>

                    {/* DOCUMENTOS DE LA EMPRESA */}
                    <div className="space-y-3">
                      <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Documentación Empresarial</h3>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        {[
                          { id: 'prl', label: 'Plan de Prevención (PRL)' },
                          { id: 'er', label: 'Evaluación de Riesgos' },
                          { id: 'sp', label: 'Servicio de Prevención' }
                        ].map(doc => {
                          const isOk = selectedCaeRecord.companyDocs?.[doc.id];
                          return (
                            <button 
                              key={doc.id}
                              onClick={() => handleToggleCompanyDoc(selectedCaeRecord.id, doc.id)}
                              className={`p-3 rounded-lg border text-left flex items-center justify-between transition ${
                                isOk ? 'bg-emerald-50 border-emerald-300 text-emerald-900' : 'bg-slate-50 border-slate-200 text-slate-700'
                              }`}
                            >
                              <span className="text-xs font-semibold">{doc.label}</span>
                              {isOk ? <CheckCircle2 className="w-4 h-4 text-emerald-600" /> : <XCircle className="w-4 h-4 text-slate-400" />}
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* TRABAJADORES Y FIRMAS */}
                    <div className="space-y-4 pt-4 border-t">
                      <div className="flex justify-between items-center">
                        <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">Trabajadores y Firmas Digitales</h3>
                        <button 
                          onClick={() => setIsAddingWorker(!isAddingWorker)}
                          className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold rounded-lg shadow transition flex items-center gap-1"
                        >
                          <Plus className="w-4 h-4" />
                          <span>Añadir Trabajador</span>
                        </button>
                      </div>

                      {isAddingWorker && (
                        <div className="p-4 bg-amber-50 border border-amber-200 rounded-lg flex flex-col sm:flex-row gap-3 items-center">
                          <input 
                            type="text" 
                            placeholder="Nombre completo" 
                            className="w-full px-3 py-1.5 border rounded-lg text-xs"
                            value={newWorkerName}
                            onChange={(e) => setNewWorkerName(e.target.value)}
                          />
                          <input 
                            type="text" 
                            placeholder="DNI / NIE" 
                            className="w-full sm:w-48 px-3 py-1.5 border rounded-lg text-xs"
                            value={newWorkerDni}
                            onChange={(e) => setNewWorkerDni(e.target.value)}
                          />
                          <div className="flex gap-2 w-full sm:w-auto">
                            <button 
                              onClick={() => handleAddWorker(selectedCaeRecord.id)}
                              className="px-4 py-1.5 bg-amber-600 text-white text-xs font-bold rounded-lg"
                            >
                              Guardar
                            </button>
                            <button 
                              onClick={() => setIsAddingWorker(false)}
                              className="px-3 py-1.5 bg-slate-200 text-slate-600 text-xs font-medium rounded-lg"
                            >
                              Cancelar
                            </button>
                          </div>
                        </div>
                      )}

                      <div className="space-y-3">
                        {selectedCaeRecord.workers?.map(worker => {
                          const isSigned = !!worker.signedAt;
                          return (
                            <div key={worker.id} className={`p-4 rounded-xl border flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 transition ${
                              isSigned ? 'bg-emerald-50/70 border-emerald-300' : 'bg-white border-slate-200 shadow-sm'
                            }`}>
                              <div>
                                <p className="font-bold text-slate-800 text-sm flex items-center gap-2">
                                  {worker.name}
                                  {isSigned && (
                                    <span className="px-2 py-0.5 bg-emerald-600 text-white text-[10px] font-bold rounded-full flex items-center gap-1">
                                      <UserCheck className="w-3 h-3" /> Firmado
                                    </span>
                                  )}
                                </p>
                                <p className="text-xs text-slate-400 font-mono">DNI: {worker.dni}</p>
                                {isSigned && (
                                  <p className="text-[10px] text-emerald-700 mt-1">
                                    Firmado el: {worker.signedAt}
                                  </p>
                                )}
                              </div>

                              <div className="flex items-center space-x-3 w-full sm:w-auto justify-end">
                                {isSigned ? (
                                  <div className="flex items-center gap-2">
                                    <img 
                                      src={worker.signatureImage} 
                                      alt="Firma" 
                                      className="h-8 border border-slate-300 rounded bg-white px-1"
                                    />
                                    <button 
                                      onClick={() => setSigningWorker(worker)}
                                      className="text-xs font-semibold text-blue-600 hover:underline px-2 py-1"
                                    >
                                      Re-firmar
                                    </button>
                                  </div>
                                ) : (
                                  <button 
                                    onClick={() => setSigningWorker(worker)}
                                    className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg shadow-sm transition flex items-center gap-1.5"
                                  >
                                    <FileSignature className="w-4 h-4" />
                                    FIRMAR
                                  </button>
                                )}

                                <button 
                                  onClick={() => handleDeleteWorker(selectedCaeRecord.id, worker.id)}
                                  className="p-1.5 text-slate-400 hover:text-rose-600 rounded-lg transition"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: PROCEDIMIENTOS GENERALES */}
          {activeTab === 'general' && (
            <div className="space-y-6">
              <div className="flex justify-between items-center">
                <div>
                  <h2 className="text-xl font-bold text-slate-800">Procedimientos Generales de PRL</h2>
                  <p className="text-sm text-slate-500">Consulta y gestión en vertical de todos los procedimientos y normas generales.</p>
                </div>

                {currentUser.role === 'superadmin' && (
                  <button 
                    onClick={() => setIsAddingProcedure(!isAddingProcedure)}
                    className="inline-flex items-center gap-1.5 px-3 py-2 bg-blue-600 text-white hover:bg-blue-700 rounded-lg text-xs font-bold transition shadow-sm"
                  >
                    <Plus className="w-4 h-4" />
                    <span>Añadir procedimiento</span>
                  </button>
                )}
              </div>

              {isAddingProcedure && (
                <form onSubmit={handleAddProcedure} className="p-4 bg-white border border-blue-200 rounded-xl space-y-3 shadow-sm">
                  <span className="text-xs font-bold text-slate-800 block">Nuevo Procedimiento General</span>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <input 
                      type="text" 
                      required
                      placeholder="Título del procedimiento"
                      className="px-3 py-2 border text-xs rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                      value={newProcTitle}
                      onChange={(e) => setNewProcTitle(e.target.value)}
                    />
                    <select 
                      className="px-3 py-2 border text-xs rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                      value={newProcCategory}
                      onChange={(e) => setNewProcCategory(e.target.value)}
                    >
                      <option value="Procedimientos">Procedimientos</option>
                      <option value="Protocolos">Protocolos</option>
                      <option value="Plantillas">Plantillas</option>
                      <option value="Instrucciones Técnicas">Instrucciones Técnicas</option>
                    </select>
                  </div>
                  <input 
                    type="url" 
                    required
                    placeholder="URL de SharePoint / Archivo (https://...)"
                    className="w-full px-3 py-2 border text-xs rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                    value={newProcLink}
                    onChange={(e) => setNewProcLink(e.target.value)}
                  />
                  <div className="flex justify-end gap-2 pt-1">
                    <button 
                      type="button"
                      onClick={() => setIsAddingProcedure(false)}
                      className="px-3 py-1.5 text-xs text-slate-500 hover:bg-slate-100 rounded-lg font-medium"
                    >
                      Cancelar
                    </button>
                    <button 
                      type="submit"
                      className="px-4 py-1.5 text-xs bg-blue-600 text-white font-bold rounded-lg hover:bg-blue-700 shadow"
                    >
                      Guardar
                    </button>
                  </div>
                </form>
              )}

              <div className="flex flex-col space-y-3">
                {generalDocs.map((doc) => (
                  <div 
                    key={doc.id}
                    className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex items-center justify-between hover:shadow-md transition"
                  >
                    <div className="flex items-center space-x-3">
                      <div className="p-2.5 bg-blue-50 text-blue-600 rounded-lg">
                        <FileText className="w-5 h-5" />
                      </div>
                      <div>
                        <h3 className="text-sm font-bold text-slate-800">{doc.title}</h3>
                        <div className="flex items-center space-x-2 mt-0.5">
                          <span className="px-2 py-0.5 bg-slate-100 text-slate-600 text-[10px] font-semibold rounded">
                            {doc.category}
                          </span>
                          <span className="text-[11px] text-slate-400">Publicado: {doc.date}</span>
                        </div>
                      </div>
                    </div>

                    <a 
                      href={doc.link} 
                      target="_blank" 
                      rel="noreferrer"
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 hover:bg-blue-50 text-slate-700 hover:text-blue-600 rounded-lg text-xs font-semibold transition"
                    >
                      <span>Abrir</span>
                      <ExternalLink className="w-3.5 h-3.5" />
                    </a>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB: MÓDULO CENTROS DE TRABAJO */}
          {activeTab === 'centres' && (
            <div className="space-y-6">
              {!selectedCentreRecord ? (
                <div className="space-y-6">
                  <div className="bg-slate-800 text-white p-6 rounded-xl shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                    <div>
                      <div className="flex items-center space-x-2">
                        <Building2 className="w-6 h-6 text-blue-400" />
                        <h2 className="text-xl font-bold">Módulo Centros de Trabajo</h2>
                      </div>
                      <p className="text-xs text-slate-300 mt-1">
                        Selecciona un centro de trabajo para inspeccionar o editar su documentación de Prevención.
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex justify-between items-center gap-4">
                    <div className="relative w-full sm:w-80">
                      <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                      <input 
                        type="text"
                        placeholder="Buscar centro o zona..."
                        className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                      />
                    </div>
                    <span className="text-xs text-slate-500 font-medium hidden sm:inline">
                      Total: {filteredCentres.length} centro(s)
                    </span>
                  </div>

                  <div className="flex flex-col space-y-3">
                    {filteredCentres.length === 0 ? (
                      <div className="bg-white p-8 rounded-xl border border-slate-200 text-center text-slate-400 text-sm italic">
                        No se han encontrado centros de trabajo.
                      </div>
                    ) : (
                      filteredCentres.map((centre) => (
                        <div 
                          key={centre.id}
                          onClick={() => setSelectedCentreId(centre.id)}
                          className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:shadow-md hover:border-blue-300 transition cursor-pointer flex items-center justify-between group"
                        >
                          <div className="space-y-1">
                            <h3 className="text-base font-bold text-slate-800 group-hover:text-blue-600 transition">
                              {centre.name}
                            </h3>
                            <p className="text-xs text-slate-400">Zona: {centre.zone}</p>
                          </div>
                          <ChevronRight className="w-5 h-5 text-slate-400 group-hover:text-blue-600 group-hover:translate-x-1 transition" />
                        </div>
                      ))
                    )}
                  </div>
                </div>
              ) : (
                /* DETALLE DEL CENTRO SELECCIONADO */
                <div className="space-y-6">
                  <button 
                    onClick={() => setSelectedCentreId(null)}
                    className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-600 hover:text-slate-900 bg-white px-3 py-2 rounded-lg border shadow-sm transition"
                  >
                    <ArrowLeft className="w-4 h-4" />
                    <span>Volver a listado de centros</span>
                  </button>

                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-6">
                    <div>
                      <h2 className="text-xl font-bold text-slate-800">{selectedCentreRecord.name}</h2>
                      <p className="text-xs text-slate-500">Zona geográfica: {selectedCentreRecord.zone}</p>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-4 border-t">
                      {[
                        { key: 'evaluacion_riesgos', label: 'Evaluación de Riesgos' },
                        { key: 'informacion_riesgos', label: 'Información de Riesgos' },
                        { key: 'medidas_emergencia', label: 'Medidas de Emergencia' }
                      ].map((cat) => {
                        const docList = Array.isArray(selectedCentreRecord.docs?.[cat.key]) ? selectedCentreRecord.docs[cat.key] : [];
                        return (
                          <div key={cat.key} className="p-4 bg-slate-50 border rounded-xl space-y-3">
                            <div className="flex justify-between items-center">
                              <h4 className="text-xs font-bold text-slate-700">{cat.label}</h4>
                              <button 
                                onClick={() => setEditDocModal({ open: true, centreId: selectedCentreRecord.id, categoryKey: cat.key, categoryLabel: cat.label })}
                                className="p-1 text-blue-600 hover:bg-blue-100 rounded transition"
                              >
                                <Plus className="w-4 h-4" />
                              </button>
                            </div>

                            <div className="space-y-2">
                              {docList.length === 0 ? (
                                <p className="text-[11px] text-slate-400 italic">Sin documentos</p>
                              ) : (
                                docList.map(doc => (
                                  <div key={doc.id} className="p-2 bg-white rounded border flex justify-between items-center text-xs">
                                    <span className="truncate max-w-[140px] font-medium">{doc.name}</span>
                                    <a href={doc.link} target="_blank" rel="noreferrer" className="text-blue-600 hover:underline text-[11px]">Abrir</a>
                                  </div>
                                ))
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: CARGA MASIVA EXCEL */}
          {activeTab === 'excel' && currentUser.role === 'superadmin' && (
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
              <h2 className="text-lg font-bold text-slate-800">Carga Masiva de Centros y Usuarios</h2>
              <p className="text-xs text-slate-500">Selecciona un archivo Excel (.xlsx) para cargar o actualizar de forma masiva los centros de trabajo.</p>

              <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 text-center space-y-3">
                <FileSpreadsheet className="w-10 h-10 text-emerald-600 mx-auto" />
                <input 
                  type="file" 
                  accept=".xlsx, .xls"
                  onChange={handleFileUploadCentres}
                  className="block w-full text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100 cursor-pointer"
                />
              </div>
            </div>
          )}

          {/* TAB: GESTIÓN DE PERFILES */}
          {activeTab === 'users' && currentUser.role === 'superadmin' && (
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
              <h2 className="text-lg font-bold text-slate-800">Gestión de Perfiles y Usuarios</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 text-slate-600 uppercase font-bold border-b">
                    <tr>
                      <th className="p-3">Nombre</th>
                      <th className="p-3">Email</th>
                      <th className="p-3">Rol</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {users.map(u => (
                      <tr key={u.id} className="hover:bg-slate-50">
                        <td className="p-3 font-semibold">{u.name}</td>
                        <td className="p-3 text-slate-500">{u.email}</td>
                        <td className="p-3">
                          <span className="px-2 py-0.5 bg-slate-100 rounded text-[10px] font-bold uppercase">{u.role}</span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

        </main>
      </div>

      {/* MODAL DE FIRMA INTERACTIVA */}
      {signingWorker && selectedCaeRecord && (
        <SignatureModal 
          worker={signingWorker}
          companyName={selectedCaeRecord.companyName}
          onClose={() => setSigningWorker(null)}
          onSaveSignature={handleSaveWorkerSignature}
        />
      )}

      {/* MODAL EDITAR / AÑADIR ENLACE DE DOCUMENTO EN CENTRO */}
      {editDocModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl p-6 max-w-md w-full space-y-4 shadow-xl border border-slate-200">
            <div className="flex justify-between items-center border-b pb-3">
              <h3 className="font-bold text-sm text-slate-800">
                Añadir Documento: {editDocModal.categoryLabel}
              </h3>
              <button 
                onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })} 
                className="text-slate-400 hover:text-slate-600"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleAddLink} className="space-y-3">
              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">URL de SharePoint / Web</label>
                <input 
                  type="url" 
                  required
                  placeholder="https://..."
                  className="w-full px-3 py-2 border rounded-lg text-xs focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  value={newLinkUrl}
                  onChange={(e) => handleUrlChange(e.target.value)}
                />
              </div>

              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Nombre visible del documento</label>
                <input 
                  type="text" 
                  placeholder="Nombre del archivo"
                  className="w-full px-3 py-2 border rounded-lg text-xs focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  value={newLinkName}
                  onChange={(e) => setNewLinkName(e.target.value)}
                />
              </div>

              <div className="flex justify-end gap-2 pt-2 border-t">
                <button 
                  type="button"
                  onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })}
                  className="px-3 py-1.5 text-xs text-slate-600 hover:bg-slate-100 rounded-lg"
                >
                  Cancelar
                </button>
                <button 
                  type="submit"
                  className="px-4 py-1.5 text-xs bg-blue-600 text-white font-bold rounded-lg hover:bg-blue-700 shadow"
                >
                  Añadir Enlace
                </button>
              </div>
            </form>

            <div className="pt-2 border-t space-y-2">
              <span className="text-[10px] font-bold text-slate-400 uppercase">Documentos Existentes:</span>
              <div className="max-h-36 overflow-y-auto space-y-1">
                {Array.isArray(centres.find(c => c.id === editDocModal.centreId)?.docs?.[editDocModal.categoryKey]) && 
                 centres.find(c => c.id === editDocModal.centreId).docs[editDocModal.categoryKey].map(doc => (
                  <div key={doc.id} className="flex justify-between items-center p-2 bg-slate-50 rounded text-xs">
                    <span className="truncate max-w-[200px] font-medium">{doc.name}</span>
                    <button 
                      onClick={() => handleDeleteLink(doc.id)} 
                      className="text-rose-600 hover:text-rose-800"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
