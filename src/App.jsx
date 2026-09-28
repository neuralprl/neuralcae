import React, { useState, useEffect } from 'react';
import { 
  Building2, 
  FileText, 
  AlertTriangle, 
  ExternalLink, 
  LogOut, 
  ShieldCheck, 
  HardHat, 
  ArrowLeft, 
  ChevronRight, 
  Upload, 
  Download, 
  CheckCircle2, 
  RefreshCw 
} from 'lucide-react';

const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxNaUJqxU9M_cik1AqlSVQw7lfizQziZo3qbNggh1z6ydmemTe-jLLlpxYx4nuO19U/exec";

const INITIAL_USERS = [
  { email: 'neuralprl', code: 'Neuralprl@', name: 'Superadministrador', role: 'superadmin' }
];

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');

  const [caeRecords, setCaeRecords] = useState([]);
  const [syncLoading, setSyncLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('cae');
  const [selectedCaeCompanyId, setSelectedCaeCompanyId] = useState(null);

  // Sincronizar empresas, credenciales y documentos desde Google Sheets
  const syncWithGoogleSheets = async () => {
    setSyncLoading(true);
    try {
      const response = await fetch(GOOGLE_SCRIPT_URL);
      const remoteData = await response.json();

      if (Array.isArray(remoteData)) {
        setCaeRecords(remoteData);
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

  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');

    // 1. Comprobar si es Superadmin
    const admin = INITIAL_USERS.find(
      u => u.email.toLowerCase() === loginEmail.trim().toLowerCase() && u.code === loginCode.trim()
    );

    if (admin) {
      setCurrentUser({ ...admin, name: 'Superadministrador', role: 'superadmin' });
      setActiveTab('cae');
      return;
    }

    // 2. Comprobar si es una empresa contratista registrada en Google Sheets
    const empresaMatch = caeRecords.find(
      r => r.userEmail.toLowerCase() === loginEmail.trim().toLowerCase() && r.password === loginCode.trim()
    );

    if (empresaMatch) {
      setCurrentUser({
        email: empresaMatch.userEmail,
        name: empresaMatch.companyName,
        role: 'externo',
        companyId: empresaMatch.companyId
      });
      setActiveTab('cae');
      setSelectedCaeCompanyId(empresaMatch.companyId);
    } else {
      setLoginError('Usuario o contraseña incorrectos.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setLoginEmail('');
    setLoginCode('');
    setSelectedCaeCompanyId(null);
  };

  const handleDownloadTemplate = () => {
    const templateContent = "CERTIFICADO DE CUMPLIMIENTO DE PRL Y ENTREGA DE EPIs\n\nYo, en representación de la empresa contratista, certifico que nuestros trabajadores cumplen con la normativa de Prevención de Riesgos Laborales.\n\nFirma y Sello:";
    const blob = new Blob([templateContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Plantilla_Certificado_PRL.txt';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleFileUpload = (companyId, docType, file) => {
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (evt) => {
      const base64Data = evt.target.result;
      const currentComp = caeRecords.find(r => r.companyId === companyId);

      setSyncLoading(true);
      try {
        const response = await fetch(GOOGLE_SCRIPT_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({
            companyId: companyId,
            companyName: currentComp?.companyName || 'Empresa',
            userEmail: currentComp?.userEmail || '',
            docType: docType,
            fileData: base64Data
          })
        });

        const result = await response.json();
        const fileUrl = result.fileUrl || base64Data;

        setCaeRecords(prev => prev.map(r => {
          if (r.companyId === companyId) {
            return {
              ...r,
              companyDocs: {
                ...r.companyDocs,
                [docType]: fileUrl
              }
            };
          }
          return r;
        }));

        alert('¡Documento guardado con éxito en su carpeta de Google Drive!');
      } catch (err) {
        console.error("Error al subir archivo:", err);
        alert("Error al guardar el documento.");
      } finally {
        setSyncLoading(false);
      }
    };
    reader.readAsDataURL(file);
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md w-full space-y-6">
          <div className="text-center space-y-2">
            <div className="bg-blue-600 text-white w-12 h-12 rounded-lg flex items-center justify-center mx-auto shadow-lg">
              <ShieldCheck className="w-8 h-8" />
            </div>
            <h1 className="text-2xl font-bold text-slate-800">Plataforma PRL & CAE</h1>
            <p className="text-sm text-slate-500">Acceso a Contratas y Sincronización Cloud</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Correo Electrónico (Usuario)</label>
              <input 
                type="text" 
                required
                className="w-full px-4 py-2 border rounded-lg text-sm"
                placeholder="ejemplo@contrata.com"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Contraseña</label>
              <input 
                type="password" 
                required
                className="w-full px-4 py-2 border rounded-lg text-sm"
                placeholder="••••••••"
                value={loginCode}
                onChange={(e) => setLoginCode(e.target.value)}
              />
            </div>

            {loginError && (
              <div className="p-3 bg-red-50 text-red-600 text-xs rounded-lg flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <button type="submit" className="w-full py-3 bg-blue-600 text-white font-medium rounded-lg text-sm shadow">
              Iniciar Sesión
            </button>
          </form>
        </div>
      </div>
    );
  }

  const myCompanyRecord = currentUser.role === 'externo' 
    ? caeRecords.find(r => r.companyId === currentUser.companyId)
    : null;

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans">
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
            <div className="text-right hidden sm:block">
              <p className="text-sm font-medium">{currentUser.name}</p>
              <span className="text-[10px] px-2 py-0.5 bg-blue-900 text-blue-200 rounded font-semibold uppercase">
                {currentUser.role}
              </span>
            </div>
            <button onClick={handleLogout} className="p-2 text-slate-300 hover:text-white rounded-lg">
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 py-6 flex-1 w-full flex flex-col md:flex-row gap-6">
        <aside className="w-full md:w-64 bg-white rounded-xl border p-4 h-fit space-y-2">
          <button 
            onClick={() => { 
              setActiveTab('cae'); 
              if (currentUser.role === 'externo' && myCompanyRecord) setSelectedCaeCompanyId(myCompanyRecord.companyId);
              else setSelectedCaeCompanyId(null);
            }}
            className="w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2 bg-amber-500 text-white font-bold"
          >
            <HardHat className="w-4 h-4" />
            <span>Módulo CAE - Empresas</span>
          </button>
        </aside>

        <main className="flex-1">
          {activeTab === 'cae' && (
            <div className="space-y-6">
              {currentUser.role !== 'externo' && !selectedCaeCompanyId ? (
                <div className="space-y-6">
                  <div className="bg-amber-600 text-white p-6 rounded-xl flex items-center justify-between">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2">
                        <HardHat className="w-6 h-6" /> Módulo CAE - Empresas Contratistas
                      </h2>
                      <p className="text-xs text-amber-100 mt-1">Sincronizado con la pestaña "Contratas" de Google Sheets y carpetas Drive.</p>
                    </div>
                    <button 
                      onClick={syncWithGoogleSheets}
                      disabled={syncLoading}
                      className="px-3 py-2 bg-amber-700 hover:bg-amber-800 text-white text-xs font-bold rounded-lg flex items-center gap-2"
                    >
                      <RefreshCw className={`w-4 h-4 ${syncLoading ? 'animate-spin' : ''}`} />
                      Sincronizar
                    </button>
                  </div>

                  <div className="grid grid-cols-1 gap-4">
                    {caeRecords.map(record => (
                      <div 
                        key={record.companyId}
                        onClick={() => setSelectedCaeCompanyId(record.companyId)}
                        className="bg-white p-5 rounded-xl border hover:border-amber-400 transition cursor-pointer flex items-center justify-between"
                      >
                        <div>
                          <h3 className="text-base font-bold text-slate-800">{record.companyName}</h3>
                          <p className="text-xs text-slate-400">{record.userEmail} • <span className="font-mono text-slate-600">Password: {record.password}</span></p>
                        </div>
                        <ChevronRight className="w-5 h-5 text-slate-400" />
                      </div>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  {currentUser.role !== 'externo' && (
                    <button onClick={() => setSelectedCaeCompanyId(null)} className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-600 bg-white px-3 py-2 rounded-lg border">
                      <ArrowLeft className="w-4 h-4" /> Volver a Empresas
                    </button>
                  )}

                  {(() => {
                    const comp = caeRecords.find(r => r.companyId === (currentUser.role === 'externo' ? myCompanyRecord?.companyId : selectedCaeCompanyId));
                    if (!comp) return <p>Cargando información de la empresa...</p>;

                    return (
                      <div className="bg-white p-6 rounded-xl border shadow-sm space-y-6">
                        <div className="border-b pb-4 flex justify-between items-center">
                          <div>
                            <h2 className="text-xl font-bold text-slate-800">{comp.companyName}</h2>
                            <p className="text-xs text-slate-500">{comp.userEmail}</p>
                          </div>
                          {currentUser.role === 'externo' && (
                            <button 
                              onClick={handleDownloadTemplate}
                              className="px-4 py-2 bg-blue-50 text-blue-700 border border-blue-200 hover:bg-blue-100 rounded-lg text-xs font-bold flex items-center gap-2 shadow-sm transition"
                            >
                              <Download className="w-4 h-4" /> Descargar Plantilla Certificado PRL
                            </button>
                          )}
                        </div>

                        <div className="space-y-4">
                          <h3 className="text-xs font-bold text-slate-400 uppercase tracking-wider">
                            Documentación Obligatoria (3 Ítems)
                          </h3>

                          <div className="grid grid-cols-1 gap-4">
                            {[
                              { key: 'prl', label: '1. Certificado / Plan de Prevención (PRL)' },
                              { key: 'er', label: '2. Evaluación de Riesgos Específica' },
                              { key: 'sp', label: '3. Acreditación Servicio de Prevención' }
                            ].map(item => {
                              const docUrl = comp.companyDocs?.[item.key];
                              const isUploaded = !!docUrl;

                              return (
                                <div key={item.key} className="p-4 rounded-xl border flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 bg-slate-50">
                                  <div>
                                    <p className="font-bold text-slate-800 text-sm flex items-center gap-2">
                                      {item.label}
                                      {isUploaded && <span className="px-2 py-0.5 bg-emerald-600 text-white text-[10px] font-bold rounded-full flex items-center gap-1"><CheckCircle2 className="w-3 h-3"/> Subido</span>}
                                    </p>
                                    <p className="text-xs text-slate-400">
                                      {isUploaded ? 'Documento guardado en carpeta de Google Drive' : 'Pendiente de subida'}
                                    </p>
                                  </div>

                                  <div className="flex items-center gap-2 w-full sm:w-auto justify-end">
                                    {isUploaded && (
                                      <a href={docUrl} target="_blank" rel="noopener noreferrer" className="px-3 py-2 bg-white border rounded-lg text-xs font-bold text-blue-600 flex items-center gap-1">
                                        Ver <ExternalLink className="w-3 h-3" />
                                      </a>
                                    )}

                                    {currentUser.role === 'externo' && (
                                      <label className="cursor-pointer px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-lg shadow-sm flex items-center gap-1.5 transition">
                                        <Upload className="w-4 h-4" /> {isUploaded ? 'Actualizar Archivo' : 'Subir Archivo'}
                                        <input 
                                          type="file" 
                                          accept=".pdf,.png,.jpg,.jpeg,.txt" 
                                          className="hidden" 
                                          onChange={e => handleFileUpload(comp.companyId, item.key, e.target.files[0])}
                                        />
                                      </label>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    );
                  })()}
                </div>
              )}
            </div>
          )}
        </main>
      </div>
    </div>
  );
}
