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
  Check,
  CheckSquare,
  Square,
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
    docsRead: false,
    docsSent: false,
    status: 'pendiente',
    updatedAt: '2026-02-20',
    workers: [
      { id: 'w1', name: 'Juan Pérez Gómez', dni: '12345678A', approved: false },
      { id: 'w2', name: 'María López Sanchis', dni: '87654321B', approved: true },
      { id: 'w3', name: 'Carlos Ruiz Delgado', dni: '45678912C', approved: false }
    ]
  },
  {
    id: 'cae_2',
    centreId: 'c1',
    companyName: 'Construcciones e Instalaciones Norte SA',
    userEmail: 'obras@nortesa.com',
    docsRead: true,
    docsSent: true,
    status: 'pendiente',
    updatedAt: '2026-02-22',
    workers: [
      { id: 'w4', name: 'Antonio García Vidal', dni: '11223344D', approved: false }
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
      if (user.role === 'externo') {
        setActiveTab('cae');
      } else {
        setActiveTab('cae');
      }
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

  // Añadir un nuevo procedimiento a Procedimientos Generales
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

  // --- ACCIONES DE CAE Y TRABAJADORES ---

  const handleAddWorker = (recordId) => {
    if (!newWorkerName.trim() || !newWorkerDni.trim()) return;

    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        const newWorker = {
          id: `w_${Date.now()}`,
          name: newWorkerName.trim(),
          dni: newWorkerDni.trim(),
          approved: false
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

  const handleToggleWorkerApproval = (recordId, workerId) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        const updatedWorkers = (r.workers || []).map(w => 
          w.id === workerId ? { ...w, approved: !w.approved } : w
        );
        return { ...r, workers: updatedWorkers };
      }
      return r;
    }));
  };

  const handleToggleSaDocRead = (recordId) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        return { ...r, docsRead: !r.docsRead };
      }
      return r;
    }));
  };

  const handleToggleSaDocSent = (recordId) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        return { ...r, docsSent: !r.docsSent };
      }
      return r;
    }));
  };

  const handleToggleApproveCae = (recordId) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        const isApproved = r.status === 'completado';
        if (isApproved) {
          return { ...r, status: 'pendiente' };
        } else {
          return { 
            ...r, 
            docsRead: true,
            docsSent: true,
            status: 'completado',
            workers: (r.workers || []).map(w => ({ ...w, approved: true }))
          };
        }
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
        
        {/* NAVEGACIÓN LATERAL REORGANIZADA */}
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

          {/* SECCIÓN 2: MÓDULOS ESPECIALES (A MEDIA PÁGINA) */}
          <div className="pt-4 border-t border-slate-100 space-y-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2 block">
              MÓDULOS ESPECIALES
            </span>
            
            {/* MÓDULO CAE - EMPRESAS */}
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

            {/* MÓDULO CENTROS DE TRABAJO */}
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

              {/* FORMULARIO PARA AÑADIR PROCEDIMIENTO */}
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

              {/* LISTA VERTICAL DE PROCEDIMIENTOS */}
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
              
              {/* VISTA 1: LISTA VERTICAL DE CENTROS */}
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

                  {/* BÚSQUEDA Y FILTRO */}
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

                  {/* LISTADO VERTICAL DE CENTROS */}
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
                                    <AlertTriangle className="w-3.5 h-3.5" /> {pendingCount} Categoría(s) pendientes de documentación
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
                /* VISTA 2: DESPLEGADO Y EDICIÓN DEL CENTRO SELECCIONADO */
                <div className="space-y-6">
                  {/* BOTÓN PARA VOLVER ATRÁS */}
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

                  {/* CABECERA DEL CENTRO */}
                  <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-2">
                    <span className="text-xs font-semibold px-2.5 py-1 bg-slate-100 text-slate-600 rounded">
                      {selectedCentreRecord.zone}
                    </span>
                    <h2 className="text-2xl font-bold text-slate-800">{selectedCentreRecord.name}</h2>
                    <p className="text-xs text-slate-400">
                      Usuarios con acceso: {selectedCentreRecord.users.join(', ')}
                    </p>
                  </div>

                  {/* CATEGORÍAS DOCUMENTALES */}
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
                                Sin documentos vinculados en esta categoría.
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
                        Selecciona una empresa para gestionar su situación documental y la validación de trabajadores.
                      </p>
                    </div>
                  </div>

                  {/* BÚSQUEDA Y FILTRO */}
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

                  {/* LISTADO VERTICAL DE EMPRESAS */}
                  <div className="flex flex-col space-y-3">
                    {filteredCaeRecords.length === 0 ? (
                      <div className="bg-white p-8 rounded-xl border border-slate-200 text-center text-slate-400 text-sm italic">
                        No se han encontrado empresas registradas en el Módulo CAE.
                      </div>
                    ) : (
                      filteredCaeRecords.map((rec) => {
                        const centre = centres.find(c => c.id === rec.centreId);
                        const isCompleted = rec.status === 'completado';
                        const workerList = rec.workers || [];
                        const approvedWorkers = workerList.filter(w => w.approved).length;

                        return (
                          <div 
                            key={rec.id}
                            onClick={() => {
                              setSelectedCaeCompanyId(rec.id);
                              setIsAddingWorker(false);
                            }}
                            className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:shadow-md hover:border-blue-300 transition cursor-pointer flex items-center justify-between group"
                          >
                            <div className="space-y-1">
                              <div className="flex items-center space-x-3">
                                <h3 className="text-base font-bold text-slate-800 group-hover:text-blue-600 transition">
                                  {rec.companyName}
                                </h3>
                                {isCompleted ? (
                                  <span className="px-2.5 py-0.5 bg-emerald-100 text-emerald-800 font-bold rounded-full text-[10px]">
                                    Acceso Libre
                                  </span>
                                ) : (
                                  <span className="px-2.5 py-0.5 bg-slate-100 text-slate-600 font-bold rounded-full text-[10px]">
                                    Pendiente
                                  </span>
                                )}
                              </div>
                              <p className="text-xs text-slate-500">
                                Centro: <span className="font-semibold text-slate-700">{centre?.name || 'No asignado'}</span> | Contacto: {rec.userEmail}
                              </p>
                              <div className="flex items-center space-x-3 pt-1 text-[11px] text-slate-400">
                                <span>Trabajadores: <strong className="text-slate-700">{approvedWorkers}/{workerList.length} Aprobados</strong></span>
                                <span>•</span>
                                <span>Docs Lectura: <strong className={rec.docsRead ? 'text-emerald-600' : 'text-slate-600'}>{rec.docsRead ? 'OK' : 'Pendiente'}</strong></span>
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
                  {/* BOTÓN PARA VOLVER ATRÁS */}
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

                  {/* PANEL PRINCIPAL DE LA EMPRESA */}
                  <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-6 space-y-6">
                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-slate-100 pb-4">
                      <div>
                        <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">Detalle CAE Empresa</span>
                        <h2 className="text-2xl font-bold text-slate-800">{selectedCaeRecord.companyName}</h2>
                        <p className="text-xs text-slate-500">{selectedCaeRecord.userEmail}</p>
                      </div>

                      <button 
                        onClick={() => handleToggleApproveCae(selectedCaeRecord.id)}
                        className={`py-2.5 px-4 rounded-lg font-bold text-xs shadow-md transition flex items-center gap-2 text-white ${
                          selectedCaeRecord.status === 'completado'
                            ? 'bg-emerald-600 hover:bg-emerald-700' 
                            : 'bg-rose-600 hover:bg-rose-700 active:scale-95'
                        }`}
                      >
                        <Check className="w-4 h-4" />
                        <span>{selectedCaeRecord.status === 'completado' ? 'Validado y Libres (OK)' : 'Validar y Liberar Acceso (OK)'}</span>
                      </button>
                    </div>

                    {/* CONTROLES DE VERIFICACIÓN RÁPIDA DE LA EMPRESA */}
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-4 bg-slate-50 p-4 rounded-xl border border-slate-200">
                      <div>
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1">Centro Asignado</span>
                        <p className="text-xs font-bold text-slate-800">
                          {centres.find(c => c.id === selectedCaeRecord.centreId)?.name || 'Sin asignación'}
                        </p>
                      </div>

                      <div>
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1">Lectura PRL (Click SA)</span>
                        <button 
                          onClick={() => handleToggleSaDocRead(selectedCaeRecord.id)}
                          className={`px-3 py-1.5 rounded border text-xs font-bold inline-flex items-center gap-1.5 transition ${
                            selectedCaeRecord.docsRead 
                              ? 'bg-emerald-100 text-emerald-800 border-emerald-300' 
                              : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'
                          }`}
                        >
                          {selectedCaeRecord.docsRead ? <CheckSquare className="w-4 h-4 text-emerald-600" /> : <Square className="w-4 h-4" />}
                          <span>{selectedCaeRecord.docsRead ? 'Confirmado (OK)' : 'Pendiente'}</span>
                        </button>
                      </div>

                      <div>
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider block mb-1">Documentación Recibida</span>
                        <button 
                          onClick={() => handleToggleSaDocSent(selectedCaeRecord.id)}
                          className={`px-3 py-1.5 rounded border text-xs font-bold inline-flex items-center gap-1.5 transition ${
                            selectedCaeRecord.docsSent 
                              ? 'bg-emerald-100 text-emerald-800 border-emerald-300' 
                              : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-100'
                          }`}
                        >
                          {selectedCaeRecord.docsSent ? <CheckSquare className="w-4 h-4 text-emerald-600" /> : <Square className="w-4 h-4" />}
                          <span>{selectedCaeRecord.docsSent ? 'Recibido (OK)' : 'Pendiente'}</span>
                        </button>
                      </div>
                    </div>

                    {/* SECCIÓN DE TRABAJADORES LISTADOS EN VERTICAL */}
                    <div className="space-y-4 pt-2">
                      <div className="flex justify-between items-center">
                        <h3 className="text-base font-bold text-slate-800">
                          Lista de Trabajadores Autorizados
                        </h3>

                        <button 
                          onClick={() => setIsAddingWorker(!isAddingWorker)}
                          className="inline-flex items-center gap-1 px-3 py-1.5 bg-blue-50 text-blue-700 hover:bg-blue-100 rounded-lg border border-blue-200 text-xs font-bold transition"
                        >
                          <Plus className="w-4 h-4 text-blue-600" />
                          <span>Añadir trabajador</span>
                        </button>
                      </div>

                      {/* FORMULARIO PARA AÑADIR TRABAJADOR */}
                      {isAddingWorker && (
                        <div className="p-4 bg-slate-50 border border-blue-200 rounded-xl space-y-3 max-w-md">
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
                              className="px-3 py-1.5 text-xs text-slate-500 hover:bg-slate-200 rounded-lg font-medium"
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

                      {/* TRABAJADORES REORDENADOS: PENDIENTES ARRIBA, APROBADOS (OK) EN VERDE ABAJO */}
                      {(() => {
                        const allWorkers = selectedCaeRecord.workers || [];
                        const pendingWorkers = allWorkers.filter(w => !w.approved);
                        const approvedWorkers = allWorkers.filter(w => w.approved);

                        if (allWorkers.length === 0) {
                          return (
                            <p className="text-xs text-slate-400 italic bg-slate-50 p-4 rounded-lg border border-dashed text-center">
                              No hay trabajadores registrados en esta empresa.
                            </p>
                          );
                        }

                        return (
                          <div className="flex flex-col space-y-3">
                            
                            {/* BLOQUE TRABAJADORES PENDIENTES (EN VERTICAL) */}
                            {pendingWorkers.length > 0 && (
                              <div className="space-y-2">
                                <span className="text-[10px] font-bold text-amber-600 uppercase tracking-wider block">
                                  Pendientes de OK ({pendingWorkers.length})
                                </span>
                                <div className="flex flex-col space-y-2">
                                  {pendingWorkers.map(worker => (
                                    <div 
                                      key={worker.id}
                                      className="p-3 bg-white border border-slate-200 rounded-xl flex items-center justify-between text-xs shadow-sm"
                                    >
                                      <div>
                                        <p className="font-bold text-slate-800">{worker.name}</p>
                                        <p className="text-[11px] text-slate-400 font-mono">DNI: {worker.dni}</p>
                                      </div>

                                      <div className="flex items-center space-x-2 shrink-0">
                                        <button 
                                          onClick={() => handleToggleWorkerApproval(selectedCaeRecord.id, worker.id)}
                                          className="px-3 py-1.5 rounded-lg border text-xs font-bold flex items-center gap-1.5 transition bg-slate-100 text-slate-600 border-slate-300 hover:bg-emerald-50 hover:text-emerald-700 hover:border-emerald-300"
                                        >
                                          <Square className="w-4 h-4" />
                                          <span>Validar (OK)</span>
                                        </button>

                                        <button 
                                          onClick={() => handleDeleteWorker(selectedCaeRecord.id, worker.id)}
                                          className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition"
                                          title="Eliminar trabajador"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              </div>
                            )}

                            {/* BLOQUE TRABAJADORES EN VERDE / APROBADOS ABAJO (EN VERTICAL) */}
                            {approvedWorkers.length > 0 && (
                              <div className="space-y-2 pt-3 border-t border-slate-100">
                                <span className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider block">
                                  Trabajadores Aprobados / OK ({approvedWorkers.length})
                                </span>
                                <div className="flex flex-col space-y-2">
                                  {approvedWorkers.map(worker => (
                                    <div 
                                      key={worker.id}
                                      className="p-3 bg-emerald-50/70 border border-emerald-200 rounded-xl flex items-center justify-between text-xs shadow-sm"
                                    >
                                      <div>
                                        <p className="font-bold text-emerald-950">{worker.name}</p>
                                        <p className="text-[11px] text-emerald-700 font-mono">DNI: {worker.dni}</p>
                                      </div>

                                      <div className="flex items-center space-x-2 shrink-0">
                                        <button 
                                          onClick={() => handleToggleWorkerApproval(selectedCaeRecord.id, worker.id)}
                                          className="px-3 py-1.5 rounded-lg border text-xs font-bold flex items-center gap-1.5 transition bg-emerald-600 text-white border-emerald-600 hover:bg-emerald-700 shadow-sm"
                                        >
                                          <CheckSquare className="w-4 h-4" />
                                          <span>Aprobado (OK)</span>
                                        </button>

                                        <button 
                                          onClick={() => handleDeleteWorker(selectedCaeRecord.id, worker.id)}
                                          className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition"
                                          title="Eliminar trabajador"
                                        >
                                          <Trash2 className="w-4 h-4" />
                                        </button>
                                      </div>
                                    </div>
                                  ))}
                                </div>
                              </div>
                            )}

                          </div>
                        );
                      })()}
                    </div>
                  </div>
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
