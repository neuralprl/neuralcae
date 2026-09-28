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

  // Coordenadas ajustadas a la escala real del canvas (evita firmas desplazadas)
  const getPos = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const point = e.touches ? e.touches[0] : e;
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
      x: (point.clientX - rect.left) * scaleX,
      y: (point.clientY - rect.top) * scaleY
    };
  };

  const startDrawing = (e) => {
    const ctx = canvasRef.current.getContext('2d');
    const { x, y } = getPos(e);
    ctx.beginPath();
    ctx.moveTo(x, y);
    setIsDrawing(true);
  };

  const draw = (e) => {
    if (!isDrawing) return;
    const ctx = canvasRef.current.getContext('2d');
    const { x, y } = getPos(e);

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
  // Solo se permite guardar en el servidor cuando la carga inicial ha tenido éxito
  const [stateLoaded, setStateLoaded] = useState(false);

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

  // ==========================================
  // SINCRONIZACIÓN CON GOOGLE (Drive / Sheets)
  // ==========================================

  // Carga el estado completo guardado en el servidor
  const syncWithGoogleSheets = async () => {
    setSyncLoading(true);
    try {
      const response = await fetch(`${GOOGLE_SCRIPT_URL}?action=state`);
      const data = await response.json();

      if (data.status === 'error') {
        throw new Error(data.message);
      }

      if (data.state) {
        const s = data.state;
        if (s.users) setUsers(s.users);
        if (s.centres) setCentres(s.centres);
        if (s.generalDocs) setGeneralDocs(s.generalDocs);
        if (s.caeRecords) setCaeRecords(s.caeRecords);
      }

      setStateLoaded(true);
    } catch (error) {
      console.error('Error cargando datos del servidor:', error);
      alert('No se pudieron cargar los datos del servidor. El guardado automático queda desactivado para no sobrescribir datos existentes.');
    } finally {
      setSyncLoading(false);
    }
  };

  useEffect(() => {
    syncWithGoogleSheets();
  }, []);

  // Guardado automático de todo el estado (con pequeña espera para agrupar cambios)
  useEffect(() => {
    if (!stateLoaded) return;

    const timer = setTimeout(() => {
      fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'saveState',
          state: { users, centres, generalDocs, caeRecords }
        })
      }).catch(err => console.error('Error guardando estado:', err));
    }, 1000);

    return () => clearTimeout(timer);
  }, [stateLoaded, users, centres, generalDocs, caeRecords]);

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

  // Guardar Firma e Integrar con Google Sheets / Drive (Persistencia Real)
  const handleSaveWorkerSignature = async (workerId, signatureDataUrl) => {
    const signedAtString = new Date().toLocaleString('es-ES');
    const currentCompany = caeRecords.find(r => r.id === selectedCaeCompanyId);
    const currentWorker = currentCompany?.workers?.find(w => w.id === workerId);

    setSyncLoading(true);

    try {
      const response = await fetch(GOOGLE_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          companyId: currentCompany?.id || '',
          companyName: currentCompany?.companyName || '',
          workerId: workerId,
          workerName: currentWorker?.name || '',
          workerDni: currentWorker?.dni || '',
          signatureImage: signatureDataUrl,
          signedAt: signedAtString
        })
      });

      const result = await response.json();

      // Si Drive falla, NO se marca como firmado
      if (result.status !== 'success' || !result.signatureUrl) {
        throw new Error(result.message || 'Respuesta no válida del servidor');
      }

      const savedUrl = result.signatureUrl;

      // Actualizar estado local (se guarda después automáticamente en el servidor)
      setCaeRecords(prev => prev.map(r => {
        if (r.id === selectedCaeCompanyId) {
          const updatedWorkers = (r.workers || []).map(w => {
            if (w.id === workerId) {
              return {
                ...w,
                signedAt: signedAtString,
                signatureImage: savedUrl,
                checks: { epis: true, inf: true, for: true, vs: true }
              };
            }
            return w;
          });
          return { ...r, workers: updatedWorkers };
        }
        return r;
      }));

    } catch (error) {
      console.error("Error guardando en Google Sheets:", error);
      alert("Error al enviar la firma a Google Drive/Sheets: " + error.message);
    } finally {
      setSyncLoading(false);
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
              disabled={syncLoading}
              className="w-full py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition shadow-md text-sm disabled:opacity-60"
            >
              {syncLoading ? 'Cargando datos...' : 'Iniciar Sesión'}
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
                        Sincronización con Google Sheets y Google Drive.
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

                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex justify-between items-center">
                    <div className="relative flex-1 max-w-md">
                      <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                      <input 
                        type="text" 
                        placeholder="Buscar empresa contratista..."
                        className="w-full pl-9 pr-4 py-2 border rounded-lg text-xs"
                        value={caeSearchTerm}
                        onChange={e => setCaeSearchTerm(e.target.value)}
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 gap-4">
                    {filteredCaeRecords.map(record => {
                      const isApproved = isCompanyFullyApproved(record);
                      return (
                        <div 
                          key={record.id}
                          onClick={() => setSelectedCaeCompanyId(record.id)}
                          className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:border-amber-400 transition cursor-pointer flex items-center justify-between"
                        >
                          <div className="space-y-1">
                            <div className="flex items-center gap-2">
                              <h3 className="text-base font-bold text-slate-800">{record.companyName}</h3>
                              {isApproved ? (
                                <span className="px-2 py-0.5 bg-emerald-100 text-emerald-700 text-[10px] font-bold rounded-full">Apta</span>
                              ) : (
                                <span className="px-2 py-0.5 bg-amber-100 text-amber-700 text-[10px] font-bold rounded-full">Pendiente</span>
                              )}
                            </div>
                            <p className="text-xs text-slate-400">{record.userEmail}</p>
                          </div>
                          <ChevronRight className="w-5 h-5 text-slate-400" />
                        </div>
                      );
                    })}
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  <button onClick={() => setSelectedCaeCompanyId(null)} className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-600 bg-white px-3 py-2 rounded-lg border">
                    <ArrowLeft className="w-4 h-4" /> Volver a Empresas
                  </button>

                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-6">
                    <div className="flex justify-between items-start border-b pb-4">
                      <div>
                        <h2 className="text-xl font-bold text-slate-800">{selectedCaeRecord.companyName}</h2>
                        <p className="text-xs text-slate-500">{selectedCaeRecord.userEmail}</p>
                      </div>
                    </div>

                    <div className="space-y-3">
                      <h3 className="text-xs font-bold text-slate-400 uppercase">Documentación Empresarial</h3>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                        {[
                          { key: 'prl', label: 'Plan de Prevención' },
                          { key: 'er', label: 'Evaluación de Riesgos' },
                          { key: 'sp', label: 'Servicio de Prevención' }
                        ].map(doc => (
                          <div 
                            key={doc.key}
                            onClick={() => handleToggleCompanyDoc(selectedCaeRecord.id, doc.key)}
                            className={`p-3 rounded-lg border text-xs font-bold cursor-pointer flex items-center justify-between ${selectedCaeRecord.companyDocs?.[doc.key] ? 'bg-emerald-50 border-emerald-200 text-emerald-800' : 'bg-slate-50 border-slate-200 text-slate-500'}`}
                          >
                            <span>{doc.label}</span>
                            {selectedCaeRecord.companyDocs?.[doc.key] ? <Check className="w-4 h-4 text-emerald-600" /> : <X className="w-4 h-4 text-slate-400" />}
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className="space-y-4 pt-4 border-t">
                      <div className="flex justify-between items-center">
                        <h3 className="text-xs font-bold text-slate-400 uppercase">Trabajadores Adscritos</h3>
                        <button onClick={() => setIsAddingWorker(!isAddingWorker)} className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold rounded-lg flex items-center gap-1">
                          <Plus className="w-4 h-4" /> Añadir Trabajador
                        </button>
                      </div>

                      {isAddingWorker && (
                        <div className="p-4 bg-amber-50 border border-amber-200 rounded-lg flex flex-col sm:flex-row gap-3">
                          <input type="text" placeholder="Nombre completo" className="px-3 py-1.5 border rounded-lg text-xs w-full" value={newWorkerName} onChange={e => setNewWorkerName(e.target.value)} />
                          <input type="text" placeholder="DNI" className="px-3 py-1.5 border rounded-lg text-xs w-full" value={newWorkerDni} onChange={e => setNewWorkerDni(e.target.value)} />
                          <button onClick={() => handleAddWorker(selectedCaeRecord.id)} className="px-4 py-1.5 bg-amber-600 text-white text-xs font-bold rounded-lg shrink-0">Guardar</button>
                        </div>
                      )}

                      <div className="space-y-3">
                        {selectedCaeRecord.workers?.map(worker => {
                          const isSigned = !!worker.signedAt;
                          return (
                            <div key={worker.id} className="p-4 rounded-xl border border-slate-200 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 bg-white shadow-sm">
                              <div>
                                <div className="flex items-center gap-2">
                                  <p className="font-bold text-slate-800 text-sm">{worker.name}</p>
                                  {isSigned && <span className="px-2 py-0.5 bg-emerald-600 text-white text-[10px] font-bold rounded-full">Firmado e Integrado</span>}
                                </div>
                                <p className="text-xs text-slate-400 font-mono">DNI: {worker.dni}</p>
                              </div>

                              <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                                {isSigned ? (
                                  <a href={worker.signatureImage} target="_blank" rel="noopener noreferrer" className="flex items-center gap-2">
                                    <img src={worker.signatureImage} alt="Firma" className="h-9 border rounded bg-white px-1 shadow-sm" />
                                    <ExternalLink className="w-4 h-4 text-blue-600" />
                                  </a>
                                ) : (
                                  <button onClick={() => setSigningWorker(worker)} className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-lg shadow-sm flex items-center gap-1.5">
                                    <FileSignature className="w-4 h-4" /> FIRMAR AHORA
                                  </button>
                                )}
                                <button onClick={() => handleDeleteWorker(selectedCaeRecord.id, worker.id)} className="p-1.5 text-slate-400 hover:text-rose-600 rounded">
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

          {/* TAB: MÓDULO CENTROS DE TRABAJO */}
          {activeTab === 'centres' && (
            <div className="space-y-6">
              {!selectedCentreRecord ? (
                <div className="space-y-6">
                  <div className="bg-blue-600 text-white p-6 rounded-xl shadow-sm flex justify-between items-center">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2">
                        <Building2 className="w-6 h-6" /> Centros de Trabajo
                      </h2>
                      <p className="text-xs text-blue-100 mt-1">Gestión de Evaluaciones de Riesgo y Planes por Centro</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    {filteredCentres.map(centre => (
                      <div 
                        key={centre.id}
                        onClick={() => setSelectedCentreId(centre.id)}
                        className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:border-blue-400 transition cursor-pointer flex justify-between items-center"
                      >
                        <div>
                          <h3 className="font-bold text-slate-800 text-base">{centre.name}</h3>
                          <p className="text-xs text-slate-400">{centre.zone}</p>
                        </div>
                        <ChevronRight className="w-5 h-5 text-slate-400" />
                      </div>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  <button onClick={() => setSelectedCentreId(null)} className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-600 bg-white px-3 py-2 rounded-lg border">
                    <ArrowLeft className="w-4 h-4" /> Volver a Centros
                  </button>

                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-6">
                    <div>
                      <h2 className="text-xl font-bold text-slate-800">{selectedCentreRecord.name}</h2>
                      <p className="text-xs text-slate-400">{selectedCentreRecord.zone}</p>
                    </div>

                    <div className="space-y-4">
                      {[
                        { key: 'evaluacion_riesgos', label: 'Evaluación de Riesgos' },
                        { key: 'informacion_riesgos', label: 'Información de Riesgos del Centro' },
                        { key: 'medidas_emergencia', label: 'Medidas de Emergencia y Evacuación' }
                      ].map(cat => (
                        <div key={cat.key} className="p-4 rounded-xl border bg-slate-50 space-y-2">
                          <div className="flex justify-between items-center">
                            <h4 className="text-xs font-bold text-slate-700">{cat.label}</h4>
                            <button 
                              onClick={() => setEditDocModal({ open: true, centreId: selectedCentreRecord.id, categoryKey: cat.key, categoryLabel: cat.label })}
                              className="text-xs text-blue-600 font-bold hover:underline flex items-center gap-1"
                            >
                              <Plus className="w-3 h-3" /> Añadir Enlace
                            </button>
                          </div>

                          <div className="space-y-1.5">
                            {selectedCentreRecord.docs?.[cat.key]?.length > 0 ? (
                              selectedCentreRecord.docs[cat.key].map(doc => (
                                <div key={doc.id} className="bg-white p-2.5 rounded-lg border flex justify-between items-center text-xs">
                                  <span className="font-semibold text-slate-700">{doc.name}</span>
                                  <a href={doc.link} target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline flex items-center gap-1 font-bold">
                                    Abrir <ExternalLink className="w-3 h-3" />
                                  </a>
                                </div>
                              ))
                            ) : (
                              <p className="text-[11px] text-slate-400 italic">No hay documentos registrados.</p>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: DOCUMENTACIÓN GENERAL PRL */}
          {activeTab === 'general' && (
            <div className="space-y-6">
              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
                <div className="flex justify-between items-center border-b pb-4">
                  <div>
                    <h2 className="text-lg font-bold text-slate-800">Documentación General de PRL</h2>
                    <p className="text-xs text-slate-500">Procedimientos, protocolos y plantillas transversales</p>
                  </div>
                  {currentUser.role === 'superadmin' && (
                    <button onClick={() => setIsAddingProcedure(!isAddingProcedure)} className="px-3 py-2 bg-blue-600 text-white text-xs font-bold rounded-lg flex items-center gap-1">
                      <Plus className="w-4 h-4" /> Nuevo Procedimiento
                    </button>
                  )}
                </div>

                {isAddingProcedure && (
                  <form onSubmit={handleAddProcedure} className="p-4 bg-slate-50 border rounded-xl space-y-3">
                    <input type="text" placeholder="Título del procedimiento" required className="w-full px-3 py-2 border rounded-lg text-xs" value={newProcTitle} onChange={e => setNewProcTitle(e.target.value)} />
                    <input type="url" placeholder="Enlace de SharePoint (https://...)" required className="w-full px-3 py-2 border rounded-lg text-xs" value={newProcLink} onChange={e => setNewProcLink(e.target.value)} />
                    <div className="flex justify-end gap-2">
                      <button type="submit" className="px-4 py-2 bg-blue-600 text-white text-xs font-bold rounded-lg">Guardar</button>
                    </div>
                  </form>
                )}

                <div className="space-y-2">
                  {generalDocs.map(doc => (
                    <div key={doc.id} className="p-3 bg-slate-50 rounded-lg border flex justify-between items-center text-xs">
                      <div>
                        <p className="font-bold text-slate-800">{doc.title}</p>
                        <p className="text-[10px] text-slate-400">{doc.category} • {doc.date}</p>
                      </div>
                      <a href={doc.link} target="_blank" rel="noopener noreferrer" className="px-3 py-1.5 bg-white border border-slate-300 rounded text-blue-600 font-bold flex items-center gap-1">
                        Ver <ExternalLink className="w-3 h-3" />
                      </a>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* TAB: CARGA MASIVA EXCEL (Solo Superadmin) */}
          {activeTab === 'excel' && currentUser.role === 'superadmin' && (
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
              <h2 className="text-lg font-bold text-slate-800">Carga Masiva de Centros y Usuarios</h2>
              <p className="text-xs text-slate-500">Sube un archivo Excel para crear centros y asignar accesos automáticamente.</p>

              <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 text-center bg-slate-50">
                <FileSpreadsheet className="w-10 h-10 text-emerald-600 mx-auto mb-2" />
                <label className="cursor-pointer bg-emerald-600 text-white px-4 py-2 rounded-lg text-xs font-bold inline-block hover:bg-emerald-700">
                  Seleccionar archivo Excel (.xlsx)
                  <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUploadCentres} />
                </label>
              </div>
            </div>
          )}

          {/* TAB: GESTIÓN DE PERFILES (Solo Superadmin) */}
          {activeTab === 'users' && currentUser.role === 'superadmin' && (
            <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-4">
              <h2 className="text-lg font-bold text-slate-800">Gestión de Perfiles y Usuarios</h2>
              <div className="space-y-2">
                {users.map(u => (
                  <div key={u.id} className="p-3 bg-slate-50 rounded-lg border flex justify-between items-center text-xs">
                    <div>
                      <p className="font-bold text-slate-800">{u.name} ({u.email})</p>
                      <p className="text-[10px] text-slate-400">Rol: {u.role} | Empresa: {u.company}</p>
                    </div>
                    <span className="font-mono bg-slate-200 px-2 py-0.5 rounded text-[10px]">{u.code}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

        </main>
      </div>

      {/* MODAL EDITAR ENLACE DOCUMENTAL */}
      {editDocModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl max-w-md w-full p-6 space-y-4">
            <h3 className="font-bold text-slate-800 text-sm">Añadir Enlace en {editDocModal.categoryLabel}</h3>
            <form onSubmit={handleAddLink} className="space-y-3">
              <input type="url" placeholder="https://sharepoint.com/archivo.pdf" required className="w-full px-3 py-2 border rounded-lg text-xs" value={newLinkUrl} onChange={e => handleUrlChange(e.target.value)} />
              <input type="text" placeholder="Nombre visible del documento" className="w-full px-3 py-2 border rounded-lg text-xs" value={newLinkName} onChange={e => setNewLinkName(e.target.value)} />
              <div className="flex justify-end gap-2 pt-2">
                <button type="button" onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })} className="px-3 py-1.5 text-xs text-slate-600">Cancelar</button>
                <button type="submit" className="px-4 py-1.5 bg-blue-600 text-white text-xs font-bold rounded-lg">Guardar</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL DE FIRMA */}
      {signingWorker && selectedCaeRecord && (
        <SignatureModal 
          worker={signingWorker}
          companyName={selectedCaeRecord.companyName}
          onClose={() => setSigningWorker(null)}
          onSaveSignature={handleSaveWorkerSignature}
        />
      )}
    </div>
  );
}

