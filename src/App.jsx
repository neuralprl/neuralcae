import React, { useState, useMemo } from 'react';
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
  ChevronRight
} from 'lucide-react';
import * as XLSX from 'xlsx';

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
    companyDocs: {
      prl: true,
      er: true,
      sp: false
    },
    updatedAt: '2026-02-20',
    workers: [
      { id: 'w1', name: 'Juan Pérez Gómez', dni: '12345678A', checks: { epis: true, inf: true, for: true, vs: true } },
      { id: 'w2', name: 'María López Sanchis', dni: '87654321B', checks: { epis: true, inf: true, for: false, vs: true } },
      { id: 'w3', name: 'Carlos Ruiz Delgado', dni: '45678912C', checks: { epis: false, inf: false, for: false, vs: false } }
    ]
  },
  {
    id: 'cae_2',
    centreId: 'c1',
    companyName: 'Construcciones e Instalaciones Norte SA',
    userEmail: 'obras@nortesa.com',
    companyDocs: {
      prl: true,
      er: true,
      sp: true
    },
    updatedAt: '2026-02-22',
    workers: [
      { id: 'w4', name: 'Antonio García Vidal', dni: '11223344D', checks: { epis: true, inf: true, for: true, vs: true } }
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

// Helper para determinar si un trabajador tiene todos sus requerimientos OK
const isWorkerFullyApproved = (worker) => {
  if (!worker || !worker.checks) return false;
  return worker.checks.epis && worker.checks.inf && worker.checks.for && worker.checks.vs;
};

// Helper para determinar si una empresa CAE está totalmente aprobada (empresa OK + todos los trabajadores OK)
const isCompanyFullyApproved = (record) => {
  if (!record) return false;
  const companyOk = record.companyDocs?.prl && record.companyDocs?.er && record.companyDocs?.sp;
  const workers = record.workers || [];
  const allWorkersOk = workers.length > 0 && workers.every(isWorkerFullyApproved);
  return companyOk && allWorkersOk;
};

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

  // Navegación y Filtros
  const [activeTab, setActiveTab] = useState('cae');
  const [selectedCentreId, setSelectedCentreId] = useState(null);
  const [selectedCaeCompanyId, setSelectedCaeCompanyId] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [caeSearchTerm, setCaeSearchTerm] = useState('');

  // Estado para añadir nuevos procedimientos
  const [isAddingProcedure, setIsAddingProcedure] = useState(false);
  const [newProcTitle, setNewProcTitle] = useState('');
  const [newProcCategory, setNewProcCategory] = useState('Procedimientos');
  const [newProcLink, setNewProcLink] = useState('');

  // Estado para añadir trabajadores
  const [isAddingWorker, setIsAddingWorker] = useState(false);
  const [newWorkerName, setNewWorkerName] = useState('');
  const [newWorkerDni, setNewWorkerDni] = useState('');

  // Modal para SharePoint
  const [editDocModal, setEditDocModal] = useState({ open: false, centreId: null, categoryKey: null, categoryLabel: '' });
  const [newLinkUrl, setNewLinkUrl] = useState('');
  const [newLinkName, setNewLinkName] = useState('');

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

  // Registro CAE seleccionado
  const selectedCaeRecord = useMemo(() => {
    if (!selectedCaeCompanyId) return null;
    return caeRecords.find(r => r.id === selectedCaeCompanyId);
  }, [caeRecords, selectedCaeCompanyId]);

  // Búsqueda en CAE
  const filteredCaeRecords = useMemo(() => {
    if (!caeSearchTerm.trim()) return caeRecords;
    return caeRecords.filter(r => 
      r.companyName.toLowerCase().includes(caeSearchTerm.toLowerCase()) ||
      r.userEmail.toLowerCase().includes(caeSearchTerm.toLowerCase())
    );
  }, [caeRecords, caeSearchTerm]);

  // Búsqueda en Centros
  const filteredCentres = useMemo(() => {
    if (!searchTerm.trim()) return accessibleCentres;
    return accessibleCentres.filter(c => 
      c.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      c.zone.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [accessibleCentres, searchTerm]);

  // Añadir procedimiento
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

  // --- CONTROLES CAE: EMPRESA Y TRABAJADORES ---

  const handleToggleCompanyDoc = (recordId, docType) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        const updatedCompanyDocs = {
          ...r.companyDocs,
          [docType]: !r.companyDocs?.[docType]
        };
        return { ...r, companyDocs: updatedCompanyDocs };
      }
      return r;
    }));
  };

  const handleToggleWorkerCheck = (recordId, workerId, checkType) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        const updatedWorkers = (r.workers || []).map(w => {
          if (w.id === workerId) {
            const updatedChecks = {
              ...w.checks,
              [checkType]: !w.checks?.[checkType]
            };
            return { ...w, checks: updatedChecks };
          }
          return w;
        });
        return { ...r, workers: updatedWorkers };
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

  // Enlaces SharePoint
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
          
          {/* SECCIÓN 1: PROCEDIMIENTOS GENERALES */}
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

          {/* SECCIÓN 2: MÓDULOS ESPECIALES */}
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

          {/* SECCIÓN 3: ADMINISTRACIÓN SA */}
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

              {/* FORMULARIO NUEVO PROCEDIMIENTO */}
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

              {/* LISTA VERTICAL */}
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
                      filteredCentres.map((centre) => {
                        const categoriesKeys = ['evaluacion_riesgos', 'informacion_riesgos', 'medidas_emergencia'];
                        const pendingCount = categoriesKeys.filter(k => {
                          const list = Array.isArray(centre.docs[k]) ? centre.docs[k] : [];
                          return list.length === 0;
                        }).length;

                        return (
                          <div 
                            key={centre.id}
                            onClick={() => setSelectedCentreId(centre.id)}
                            className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:shadow-md hover:border-blue-300 transition cursor-pointer flex items-center justify-between group"
                          >
                            <div className="space-y-1">
                              <div className="flex items-center space-x-3">
                                <h3 className="text-base font-bold text-slate-800 group-hover:text-blue-600 transition">
                                  {centre.name}
                                </h3>
                                <span className="text-xs font-semibold px-2.5 py-0.5 bg-slate-100 text-slate-600 rounded">
                                  {centre.zone}
                                </span>
                              </div>
                              <p className="text-xs text-slate-400">
                                Usuarios con acceso: {centre.users.join(', ')}
                              </p>
                              <div className="pt-1">
                                {pendingCount === 0 ? (
                                  <span className="text-[11px] font-bold text-emerald-600 inline-flex items-center gap-1">
                                    <CheckCircle2 className="w-3.5 h-3.5" /> Toda la documentación cargada (OK)
                                  </span>
                                ) : (
                                  <span className="text-[11px] font-bold text-amber-600 inline-flex items-center gap-1">
                                    <AlertTriangle className="w-3.5 h-3.5" /> {pendingCount} Categoría(s) pendientes
                                  </span>
                                )}
                              </div>
                            </div>

                            <div className="flex items-center space-x-2 text-slate-400 group-hover:text-blue-600">
                              <span className="text-xs font-medium hidden sm:inline">Ver / Editar Centro</span>
                              <ChevronRight className="w-5 h-5" />
                            </div>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  <div className="flex justify-between items-center">
                    <button 
                      onClick={() => setSelectedCentreId(null)}
                      className="inline-flex items-center space-x-2 px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-lg text-xs font-bold transition shadow-sm"
                    >
                      <ArrowLeft className="w-4 h-4" />
                      <span>← Volver al listado de centros</span>
                    </button>

                    <span className="text-xs text-slate-400 font-mono">
                      ID Centro: {selectedCentreRecord.id}
                    </span>
                  </div>

                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-2">
                    <span className="text-xs font-semibold px-2.5 py-1 bg-slate-100 text-slate-600 rounded">
                      {selectedCentreRecord.zone}
                    </span>
                    <h2 className="text-2xl font-bold text-slate-800">{selectedCentreRecord.name}</h2>
                    <p className="text-xs text-slate-400">
                      Usuarios con acceso: {selectedCentreRecord.users.join(', ')}
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    {[
                      { key: 'evaluacion_riesgos', label: 'Evaluación de Riesgos' },
                      { key: 'informacion_riesgos', label: 'Información de Riesgos' },
                      { key: 'medidas_emergencia', label: 'Medidas de Emergencia' }
                    ].map(({ key, label }) => {
                      const docList = Array.isArray(selectedCentreRecord.docs[key]) ? selectedCentreRecord.docs[key] : [];
                      const isPresent = docList.length > 0;

                      return (
                        <div key={key} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-between space-y-4">
                          <div>
                            <div className="flex justify-between items-center mb-3">
                              <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Categoría</span>
                              {isPresent ? (
                                <span className="text-xs font-bold px-2.5 py-1 bg-emerald-100 text-emerald-700 rounded-full flex items-center gap-1">
                                  <CheckCircle2 className="w-3 h-3" /> {docList.length} Archivo(s)
                                </span>
                              ) : (
                                <span className="text-xs font-bold px-2.5 py-1 bg-rose-100 text-rose-700 rounded-full flex items-center gap-1">
                                  <XCircle className="w-3 h-3" /> Pendiente
                                </span>
                              )}
                            </div>

                            <h3 className="text-base font-bold text-slate-800 mb-3">{label}</h3>

                            {isPresent ? (
                              <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                                {docList.map(doc => (
                                  <a 
                                    key={doc.id}
                                    href={doc.link} 
                                    target="_blank" 
                                    rel="noreferrer"
                                    className="p-2.5 bg-slate-50 hover:bg-blue-50 border border-slate-200 hover:border-blue-200 rounded-lg flex items-center justify-between text-xs transition group"
                                  >
                                    <div className="flex items-center space-x-2 truncate pr-2">
                                      <FileText className="w-4 h-4 text-blue-600 shrink-0" />
                                      <span className="font-medium text-slate-700 group-hover:text-blue-700 truncate">{doc.name}</span>
                                    </div>
                                    <ExternalLink className="w-3.5 h-3.5 text-slate-400 group-hover:text-blue-600 shrink-0" />
                                  </a>
                                ))}
                              </div>
                            ) : (
                              <p className="text-xs text-rose-500 italic bg-rose-50 p-3 rounded border border-rose-100">
                                Sin documentos vinculados.
                              </p>
                            )}
                          </div>

                          <div className="space-y-2 pt-2 border-t border-slate-100">
                            {currentUser.role === 'superadmin' && (
                              <button 
                                onClick={() => {
                                  setEditDocModal({
                                    open: true,
                                    centreId: selectedCentreRecord.id,
                                    categoryKey: key,
                                    categoryLabel: label
                                  });
                                  setNewLinkUrl('');
                                  setNewLinkName('');
                                }}
                                className="w-full py-2 px-3 bg-slate-100 text-slate-700 font-medium text-xs rounded-lg hover:bg-slate-200 transition flex items-center justify-center gap-1.5"
                              >
                                <Plus className="w-3.5 h-3.5" />
                                <span>Gestionar / Añadir Enlaces</span>
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

            </div>
          )}

          {/* TAB: MÓDULO CAE */}
          {activeTab === 'cae' && (
            <div className="space-y-6">
              
              {/* VISTA 1: LISTA VERTICAL DE EMPRESAS */}
              {!selectedCaeRecord ? (
                <div className="space-y-6">
                  <div className="bg-slate-800 text-white p-6 rounded-xl shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                    <div>
                      <div className="flex items-center space-x-2">
                        <HardHat className="w-6 h-6 text-amber-400" />
                        <h2 className="text-xl font-bold">Módulo CAE- Empresas</h2>
                      </div>
                      <p className="text-xs text-slate-300 mt-1">
                        Selecciona una empresa para gestionar la documentación y la validación de trabajadores.
                      </p>
                    </div>
                  </div>

                  <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex justify-between items-center gap-4">
                    <div className="relative w-full sm:w-80">
                      <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                      <input 
                        type="text"
                        placeholder="Buscar empresa por nombre o email..."
                        className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                        value={caeSearchTerm}
                        onChange={(e) => setCaeSearchTerm(e.target.value)}
                      />
                    </div>
                    <span className="text-xs text-slate-500 font-medium hidden sm:inline">
                      Total: {filteredCaeRecords.length} empresa(s)
                    </span>
                  </div>

                  <div className="flex flex-col space-y-3">
                    {filteredCaeRecords.length === 0 ? (
                      <div className="bg-white p-8 rounded-xl border border-slate-200 text-center text-slate-400 text-sm italic">
                        No se han encontrado empresas registradas en el Módulo CAE.
                      </div>
                    ) : (
                      filteredCaeRecords.map((rec) => {
                        const centre = centres.find(c => c.id === rec.centreId);
                        const isApproved = isCompanyFullyApproved(rec);
                        const workerList = rec.workers || [];
                        const approvedWorkers = workerList.filter(isWorkerFullyApproved).length;

                        return (
                          <div 
                            key={rec.id}
                            onClick={() => {
                              setSelectedCaeCompanyId(rec.id);
                              setIsAddingWorker(false);
                            }}
                            className={`p-5 rounded-xl border transition cursor-pointer flex items-center justify-between group ${
                              isApproved 
                                ? 'bg-emerald-50/60 border-emerald-300 hover:border-emerald-500' 
                                : 'bg-white border-slate-200 hover:border-blue-300 shadow-sm hover:shadow-md'
                            }`}
                          >
                            <div className="space-y-1">
                              <div className="flex items-center space-x-3">
                                <h3 className={`text-base font-bold transition ${isApproved ? 'text-emerald-950' : 'text-slate-800 group-hover:text-blue-600'}`}>
                                  {rec.companyName}
                                </h3>
                                {isApproved ? (
                                  <span className="px-2.5 py-0.5 bg-emerald-600 text-white font-bold rounded-full text-[10px] shadow-sm">
                                    Acceso Libre (OK)
                                  </span>
                                ) : (
                                  <span className="px-2.5 py-0.5 bg-amber-100 text-amber-800 font-bold rounded-full text-[10px]">
                                    Pendiente
                                  </span>
                                )}
                              </div>
                              <p className="text-xs text-slate-500">
                                Centro: <span className="font-semibold text-slate-700">{centre?.name || 'No asignado'}</span> | Contacto: {rec.userEmail}
                              </p>
                              <div className="flex items-center space-x-3 pt-1 text-[11px] text-slate-500">
                                <span>Trabajadores autorizados: <strong className="text-emerald-700">{approvedWorkers}/{workerList.length}</strong></span>
                                <span>•</span>
                                <span>Documentación Empresa: <strong className={rec.companyDocs?.prl && rec.companyDocs?.er && rec.companyDocs?.sp ? 'text-emerald-700' : 'text-amber-700'}>
                                  {[rec.companyDocs?.prl, rec.companyDocs?.er, rec.companyDocs?.sp].filter(Boolean).length}/3 OK
                                </strong></span>
                              </div>
                            </div>

                            <div className="flex items-center space-x-2 text-slate-400 group-hover:text-blue-600">
                              <span className="text-xs font-medium hidden sm:inline">Gestionar Empresa</span>
                              <ChevronRight className="w-5 h-5" />
                            </div>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              ) : (
                /* VISTA 2: DESPLEGADO Y EDICIÓN DE EMPRESA SELECCIONADA */
                <div className="space-y-6">
                  <div className="flex justify-between items-center">
                    <button 
                      onClick={() => setSelectedCaeCompanyId(null)}
                      className="inline-flex items-center space-x-2 px-3 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-lg text-xs font-bold transition shadow-sm"
                    >
                      <ArrowLeft className="w-4 h-4" />
                      <span>← Volver al listado de empresas</span>
                    </button>

                    <span className="text-xs text-slate-400 font-mono">
                      ID CAE: {selectedCaeRecord.id}
                    </span>
                  </div>

                  {/* PANEL DE LA EMPRESA */}
                  {(() => {
                    const isCompanyOk = isCompanyFullyApproved(selectedCaeRecord);

                    return (
                      <div className={`rounded-xl border shadow-sm p-6 space-y-6 transition ${
                        isCompanyOk ? 'bg-emerald-50/50 border-emerald-300' : 'bg-white border-slate-200'
                      }`}>
                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-slate-200/60 pb-4">
                          <div>
                            <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Módulo CAE - Detalle de Empresa</span>
                            <h2 className="text-2xl font-bold text-slate-800">{selectedCaeRecord.companyName}</h2>
                            <p className="text-xs text-slate-500">{selectedCaeRecord.userEmail}</p>
                          </div>

                          <div>
                            {isCompanyOk ? (
                              <div className="px-4 py-2 bg-emerald-600 text-white rounded-lg font-bold text-xs shadow flex items-center gap-2">
                                <CheckCircle2 className="w-4 h-4" />
                                <span>EMPRESA OK - ACCESO LIBRE</span>
                              </div>
                            ) : (
                              <div className="px-4 py-2 bg-amber-100 text-amber-900 border border-amber-300 rounded-lg font-bold text-xs flex items-center gap-2">
                                <AlertTriangle className="w-4 h-4 text-amber-600" />
                                <span>DOCUMENTACIÓN PENDIENTE</span>
                              </div>
                            )}
                          </div>
                        </div>

                        {/* TRES BOTONES DE LA EMPRESA: PRL, ER, SP */}
                        <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-2">
                          <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block">
                            Documentación de Empresa (3 Requerimientos)
                          </span>
                          
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                            {/* BOTÓN PRL */}
                            <button
                              onClick={() => handleToggleCompanyDoc(selectedCaeRecord.id, 'prl')}
                              className={`py-2.5 px-3 rounded-lg border text-xs font-bold transition flex items-center justify-between ${
                                selectedCaeRecord.companyDocs?.prl
                                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                              }`}
                            >
                              <span>Certificado PRL</span>
                              <span className="text-[10px] uppercase font-mono px-1.5 py-0.5 rounded bg-black/10">
                                {selectedCaeRecord.companyDocs?.prl ? 'OK' : 'Pendiente'}
                              </span>
                            </button>

                            {/* BOTÓN ER */}
                            <button
                              onClick={() => handleToggleCompanyDoc(selectedCaeRecord.id, 'er')}
                              className={`py-2.5 px-3 rounded-lg border text-xs font-bold transition flex items-center justify-between ${
                                selectedCaeRecord.companyDocs?.er
                                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                              }`}
                            >
                              <span>Evaluación Riesgos (ER)</span>
                              <span className="text-[10px] uppercase font-mono px-1.5 py-0.5 rounded bg-black/10">
                                {selectedCaeRecord.companyDocs?.er ? 'OK' : 'Pendiente'}
                              </span>
                            </button>

                            {/* BOTÓN SP */}
                            <button
                              onClick={() => handleToggleCompanyDoc(selectedCaeRecord.id, 'sp')}
                              className={`py-2.5 px-3 rounded-lg border text-xs font-bold transition flex items-center justify-between ${
                                selectedCaeRecord.companyDocs?.sp
                                  ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-100'
                              }`}
                            >
                              <span>Servicio Prevención (SP)</span>
                              <span className="text-[10px] uppercase font-mono px-1.5 py-0.5 rounded bg-black/10">
                                {selectedCaeRecord.companyDocs?.sp ? 'OK' : 'Pendiente'}
                              </span>
                            </button>
                          </div>
                        </div>

                        {/* SECCIÓN DE TRABAJADORES */}
                        <div className="space-y-4 pt-2">
                          <div className="flex justify-between items-center">
                            <h3 className="text-base font-bold text-slate-800">
                              Gestión de Trabajadores
                            </h3>

                            <button 
                              onClick={() => setIsAddingWorker(!isAddingWorker)}
                              className="inline-flex items-center gap-1 px-3 py-1.5 bg-blue-600 text-white hover:bg-blue-700 rounded-lg text-xs font-bold transition shadow-sm"
                            >
                              <Plus className="w-4 h-4" />
                              <span>Añadir trabajador</span>
                            </button>
                          </div>

                          {/* FORMULARIO TRABAJADOR */}
                          {isAddingWorker && (
                            <div className="p-4 bg-white border border-blue-200 rounded-xl space-y-3 max-w-md shadow-sm">
                              <span className="text-xs font-bold text-slate-700 block">Nuevo Trabajador para {selectedCaeRecord.companyName}</span>
                              <input 
                                type="text"
                                placeholder="Nombre y Apellidos"
                                className="w-full px-3 py-2 border text-xs rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                                value={newWorkerName}
                                onChange={(e) => setNewWorkerName(e.target.value)}
                              />
                              <input 
                                type="text"
                                placeholder="DNI / NIE"
                                className="w-full px-3 py-2 border text-xs rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
                                value={newWorkerDni}
                                onChange={(e) => setNewWorkerDni(e.target.value)}
                              />
                              <div className="flex justify-end gap-2 pt-1">
                                <button 
                                  onClick={() => setIsAddingWorker(false)}
                                  className="px-3 py-1.5 text-xs text-slate-500 hover:bg-slate-100 rounded-lg font-medium"
                                >
                                  Cancelar
                                </button>
                                <button 
                                  onClick={() => handleAddWorker(selectedCaeRecord.id)}
                                  className="px-4 py-1.5 text-xs bg-blue-600 text-white font-bold rounded-lg hover:bg-blue-700 shadow"
                                >
                                  Guardar Trabajador
                                </button>
                              </div>
                            </div>
                          )}

                          {/* LISTADO DE TRABAJADORES: AUTORIZADOS ARRIBA, PENDIENTES ABAJO */}
                          {(() => {
                            const allWorkers = selectedCaeRecord.workers || [];
                            const approvedWorkers = allWorkers.filter(isWorkerFullyApproved);
                            const pendingWorkers = allWorkers.filter(w => !isWorkerFullyApproved(w));

                            if (allWorkers.length === 0) {
                              return (
                                <p className="text-xs text-slate-400 italic bg-slate-50 p-4 rounded-lg border border-dashed text-center">
                                  No hay trabajadores registrados en esta empresa.
                                </p>
                              );
                            }

                            return (
                              <div className="flex flex-col space-y-6">
                                
                                {/* 1. TRABAJADORES AUTORIZADOS (ARRIBA - VERDES) */}
                                <div className="space-y-2">
                                  <div className="flex items-center space-x-2">
                                    <span className="w-2.5 h-2.5 rounded-full bg-emerald-500"></span>
                                    <span className="text-xs font-bold text-emerald-800 uppercase tracking-wider">
                                      Trabajadores autorizados ({approvedWorkers.length})
                                    </span>
                                  </div>

                                  {approvedWorkers.length === 0 ? (
                                    <p className="text-xs text-slate-400 italic bg-slate-50/50 p-3 rounded-lg border border-slate-200">
                                      Sin trabajadores totalmente autorizados.
                                    </p>
                                  ) : (
                                    <div className="flex flex-col space-y-2">
                                      {approvedWorkers.map(worker => (
                                        <div 
                                          key={worker.id}
                                          className="p-3.5 bg-emerald-100/80 border border-emerald-300 rounded-xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-sm"
                                        >
                                          <div>
                                            <p className="font-bold text-emerald-950 text-sm">{worker.name}</p>
                                            <p className="text-[11px] text-emerald-800 font-mono">DNI: {worker.dni}</p>
                                          </div>

                                          {/* CUATRO BOTONES: EPIs, INF, FOR, VS */}
                                          <div className="flex items-center gap-1.5 flex-wrap">
                                            {[
                                              { key: 'epis', label: 'EPIs' },
                                              { key: 'inf', label: 'INF' },
                                              { key: 'for', label: 'FOR' },
                                              { key: 'vs', label: 'VS' }
                                            ].map(({ key, label }) => {
                                              const checked = worker.checks?.[key];
                                              return (
                                                <button
                                                  key={key}
                                                  onClick={() => handleToggleWorkerCheck(selectedCaeRecord.id, worker.id, key)}
                                                  className={`px-2.5 py-1 rounded text-xs font-bold transition border ${
                                                    checked 
                                                      ? 'bg-emerald-700 text-white border-emerald-700 shadow-sm' 
                                                      : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'
                                                  }`}
                                                >
                                                  {label} {checked ? '✓' : ''}
                                                </button>
                                              );
                                            })}

                                            <button 
                                              onClick={() => handleDeleteWorker(selectedCaeRecord.id, worker.id)}
                                              className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition ml-1"
                                              title="Eliminar trabajador"
                                            >
                                              <Trash2 className="w-4 h-4" />
                                            </button>
                                          </div>
                                        </div>
                                      ))}
                                    </div>
                                  )}
                                </div>

                                {/* 2. TRABAJADORES PENDIENTES DE OK (ABAJO) */}
                                <div className="space-y-2 pt-2 border-t border-slate-200">
                                  <div className="flex items-center space-x-2">
                                    <span className="w-2.5 h-2.5 rounded-full bg-amber-500"></span>
                                    <span className="text-xs font-bold text-amber-800 uppercase tracking-wider">
                                      Trabajadores pendientes de ok ({pendingWorkers.length})
                                    </span>
                                  </div>

                                  {pendingWorkers.length === 0 ? (
                                    <p className="text-xs text-slate-400 italic bg-slate-50/50 p-3 rounded-lg border border-slate-200">
                                      Todos los trabajadores están autorizados.
                                    </p>
                                  ) : (
                                    <div className="flex flex-col space-y-2">
                                      {pendingWorkers.map(worker => (
                                        <div 
                                          key={worker.id}
                                          className="p-3.5 bg-white border border-slate-200 rounded-xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-sm hover:border-slate-300"
                                        >
                                          <div>
                                            <p className="font-bold text-slate-800 text-sm">{worker.name}</p>
                                            <p className="text-[11px] text-slate-400 font-mono">DNI: {worker.dni}</p>
                                          </div>

                                          {/* CUATRO BOTONES: EPIs, INF, FOR, VS */}
                                          <div className="flex items-center gap-1.5 flex-wrap">
                                            {[
                                              { key: 'epis', label: 'EPIs' },
                                              { key: 'inf', label: 'INF' },
                                              { key: 'for', label: 'FOR' },
                                              { key: 'vs', label: 'VS' }
                                            ].map(({ key, label }) => {
                                              const checked = worker.checks?.[key];
                                              return (
                                                <button
                                                  key={key}
                                                  onClick={() => handleToggleWorkerCheck(selectedCaeRecord.id, worker.id, key)}
                                                  className={`px-2.5 py-1 rounded text-xs font-bold transition border ${
                                                    checked 
                                                      ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm' 
                                                      : 'bg-slate-100 text-slate-600 border-slate-300 hover:bg-slate-200'
                                                  }`}
                                                >
                                                  {label} {checked ? '✓' : ''}
                                                </button>
                                              );
                                            })}

                                            <button 
                                              onClick={() => handleDeleteWorker(selectedCaeRecord.id, worker.id)}
                                              className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition ml-1"
                                              title="Eliminar trabajador"
                                            >
                                              <Trash2 className="w-4 h-4" />
                                            </button>
                                          </div>
                                        </div>
                                      ))}
                                    </div>
                                  )}
                                </div>

                              </div>
                            );
                          })()}
                        </div>
                      </div>
                    );
                  })()}
                </div>
              )}

            </div>
          )}

          {/* TAB: CARGA MASIVA */}
          {activeTab === 'excel' && currentUser.role === 'superadmin' && (
            <div className="bg-white p-8 rounded-xl border border-slate-200 shadow-sm space-y-6">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Carga Masiva de Centros y Usuarios</h2>
                <p className="text-sm text-slate-500">Sube tu hoja de cálculo (.xlsx) para actualizar centros y accesos.</p>
              </div>

              <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 text-center hover:border-blue-500 transition cursor-pointer bg-slate-50">
                <input 
                  type="file" 
                  accept=".xlsx, .xls"
                  onChange={handleFileUploadCentres}
                  className="hidden" 
                  id="excel-upload" 
                />
                <label htmlFor="excel-upload" className="cursor-pointer space-y-2 block">
                  <FileSpreadsheet className="w-10 h-10 text-emerald-600 mx-auto" />
                  <span className="block text-sm font-medium text-slate-700">Seleccionar archivo Excel</span>
                </label>
              </div>
            </div>
          )}

          {/* TAB: GESTIÓN DE PERFILES */}
          {activeTab === 'users' && currentUser.role === 'superadmin' && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Usuarios y Roles Registrados</h2>
                <p className="text-sm text-slate-500">Listado de credenciales y perfiles.</p>
              </div>

              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <table className="w-full text-left text-sm text-slate-600">
                  <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                    <tr>
                      <th className="px-6 py-3">Nombre / Empresa</th>
                      <th className="px-6 py-3">Usuario / Correo</th>
                      <th className="px-6 py-3">Contraseña</th>
                      <th className="px-6 py-3">Rol</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {users.map((u) => (
                      <tr key={u.id} className="hover:bg-slate-50">
                        <td className="px-6 py-4 font-medium text-slate-800">{u.name}</td>
                        <td className="px-6 py-4">{u.email}</td>
                        <td className="px-6 py-4 font-mono text-xs text-slate-600">{u.code}</td>
                        <td className="px-6 py-4">
                          <span className={`px-2 py-1 text-xs font-semibold rounded ${
                            u.role === 'superadmin' ? 'bg-purple-100 text-purple-700' :
                            u.role === 'corporativo' ? 'bg-blue-100 text-blue-700' : 'bg-emerald-100 text-emerald-700'
                          }`}>
                            {u.role === 'superadmin' ? 'Superadmin (SA)' :
                             u.role === 'corporativo' ? 'Usuario Corporativo (UC)' : 'Empresa Externa (UX)'}
                          </span>
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

      {/* MODAL SHAREPOINT */}
      {editDocModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl shadow-xl max-w-lg w-full p-6 space-y-6 max-h-[90vh] flex flex-col">
            <div className="flex justify-between items-center border-b pb-3">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Gestionar Enlaces SharePoint</h3>
                <p className="text-xs text-slate-500">{editDocModal.categoryLabel}</p>
              </div>
              <button 
                onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })}
                className="text-slate-400 hover:text-slate-600 text-xl font-bold"
              >
                &times;
              </button>
            </div>

            <form onSubmit={handleAddLink} className="space-y-3 bg-slate-50 p-4 rounded-lg border border-slate-200">
              <span className="text-xs font-bold text-slate-700 uppercase block">Añadir nuevo enlace</span>
              
              <div>
                <label className="block text-xs text-slate-500 mb-1">URL de SharePoint</label>
                <div className="relative">
                  <LinkIcon className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
                  <input 
                    type="url" 
                    required
                    className="w-full pl-9 pr-3 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="https://empresa.sharepoint.com/..."
                    value={newLinkUrl}
                    onChange={(e) => handleUrlChange(e.target.value)}
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs text-slate-500 mb-1">Nombre del Archivo</label>
                <input 
                  type="text" 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                  placeholder="Ej. Evaluacion_Mestalla_2026.pdf"
                  value={newLinkName}
                  onChange={(e) => setNewLinkName(e.target.value)}
                />
              </div>

              <button 
                type="submit"
                className="w-full py-2 bg-blue-600 text-white font-medium text-xs rounded-lg hover:bg-blue-700 transition shadow flex items-center justify-center gap-1"
              >
                <Plus className="w-4 h-4" />
                Guardar Enlace
              </button>
            </form>

            <div className="flex-1 overflow-y-auto space-y-2">
              <span className="text-xs font-bold text-slate-700 uppercase block">Enlaces Guardados</span>
              
              {(() => {
                const centre = centres.find(c => c.id === editDocModal.centreId);
                const docList = centre && Array.isArray(centre.docs[editDocModal.categoryKey]) 
                  ? centre.docs[editDocModal.categoryKey] 
                  : [];

                if (docList.length === 0) {
                  return (
                    <p className="text-xs text-slate-400 text-center py-4 italic border border-dashed rounded-lg">
                      No hay enlaces guardados.
                    </p>
                  );
                }

                return docList.map(doc => (
                  <div key={doc.id} className="flex items-center justify-between p-3 bg-white border border-slate-200 rounded-lg text-xs shadow-sm">
                    <div className="flex items-center space-x-2 truncate pr-2">
                      <FileText className="w-4 h-4 text-blue-600 shrink-0" />
                      <div className="truncate">
                        <p className="font-semibold text-slate-800 truncate">{doc.name}</p>
                        <a href={doc.link} target="_blank" rel="noreferrer" className="text-slate-400 hover:text-blue-600 truncate block text-[10px]">
                          {doc.link}
                        </a>
                      </div>
                    </div>

                    <button 
                      onClick={() => handleDeleteLink(doc.id)}
                      className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition shrink-0"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ));
              })()}
            </div>

            <div className="flex justify-end pt-3 border-t">
              <button 
                onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })}
                className="px-4 py-2 text-xs font-medium bg-slate-800 text-white rounded-lg hover:bg-slate-700 transition"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
