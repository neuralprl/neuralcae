import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Building2, 
  FileText, 
  Users, 
  AlertTriangle, 
  Search, 
  ExternalLink, 
  LogOut, 
  Plus, 
  ShieldCheck, 
  FileSpreadsheet, 
  Trash2, 
  HardHat, 
  Bell, 
  ArrowLeft, 
  ChevronRight, 
  FileSignature, 
  Check, 
  X, 
  RefreshCw 
} from 'lucide-react';
import * as XLSX from 'xlsx';

const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxNaUJqxU9M_cik1AqlSVQw7lfizQziZo3qbNggh1z6ydmemTe-jLLlpxYx4nuO19U/exec";
                           
const INITIAL_USERS = [
  { id: '1', email: 'neuralprl', code: 'Neuralprl@', name: 'Superadministrador', role: 'superadmin', assignedCentres: ['ALL'], company: 'Neural PRL' },
  { id: '2', email: 'director.madrid@neural.es', code: 'Pass1234@', name: 'Carlos (Director Madrid)', role: 'corporativo', assignedCentres: ['c1'], company: 'Neural SRL' },
  { id: '3', email: 'prevencion@contratasvalencia.com', code: 'Externa123@', name: 'Mantenimientos Levante SL', role: 'externo', assignedCentres: ['c2'], company: 'Mantenimientos Levante SL' }
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

  const stopDrawing = () => setIsDrawing(false);

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
            Mediante la firma del presente documento, el/la trabajador/a <strong>{worker.name}</strong>, con DNI <strong>{worker.dni}</strong>, empleado/a de la empresa <strong>{companyName}</strong>, declara formalmente haber recibido la información y formación correspondientes.
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

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');

  const [users, setUsers] = useState(INITIAL_USERS);
  const [caeRecords, setCaeRecords] = useState(INITIAL_CAE_RECORDS);
  const [syncLoading, setSyncLoading] = useState(false);

  const [activeTab, setActiveTab] = useState('cae');
  const [selectedCaeCompanyId, setSelectedCaeCompanyId] = useState(null);
  const [caeSearchTerm, setCaeSearchTerm] = useState('');

  const [isAddingWorker, setIsAddingWorker] = useState(false);
  const [newWorkerName, setNewWorkerName] = useState('');
  const [newWorkerDni, setNewWorkerDni] = useState('');
  const [signingWorker, setSigningWorker] = useState(null);

  // Cargar firmas guardadas en Google Sheets
  const syncWithGoogleSheets = async () => {
    setSyncLoading(true);
    try {
      const response = await fetch(GOOGLE_SCRIPT_URL);
      const remoteData = await response.json();

      if (Array.isArray(remoteData) && remoteData.length > 0) {
        setCaeRecords(prevRecords => {
          return prevRecords.map(company => {
            const updatedWorkers = company.workers.map(worker => {
              const remoteSign = remoteData.find(
                r => r.workerDni === worker.dni || r.workerId === worker.id
              );
              if (remoteSign && remoteSign.signatureImage) {
                return {
                  ...worker,
                  signedAt: remoteSign.signedAt,
                  signatureImage: remoteSign.signatureImage,
                  checks: { epis: true, inf: true, for: true, vs: true }
                };
              }
              return worker;
            });

            remoteData.forEach(r => {
              if (r.companyName === company.companyName) {
                const exists = updatedWorkers.some(w => w.dni === r.workerDni || w.id === r.workerId);
                if (!exists && r.workerName) {
                  updatedWorkers.push({
                    id: r.workerId || `w_${Date.now()}`,
                    name: r.workerName,
                    dni: r.workerDni,
                    signedAt: r.signedAt,
                    signatureImage: r.signatureImage,
                    checks: { epis: true, inf: true, for: true, vs: true }
                  });
                }
              }
            });

            return { ...company, workers: updatedWorkers };
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

  // Guardar Firma en Google Sheets
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
      console.log(text);
      const savedUrl = result.signatureUrl || signatureDataUrl;

      // Actualizar estado local
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
      alert("Error al enviar la firma a Google Drive/Sheets. Comprueba la conexión.");
    } finally {
      setSyncLoading(false);
    }
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
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Usuario</label>
              <input 
                type="text" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg text-sm"
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
                className="w-full px-4 py-2 border border-slate-300 rounded-lg text-sm"
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

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans">
      <header className="bg-slate-800 text-white shadow-md sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <ShieldCheck className="w-7 h-7 text-blue-400" />
            <div>
              <h1 className="font-bold text-lg leading-tight">Neural PRL</h1>
              <span className="text-xs text-slate-400">Coordinación Empresarial</span>
            </div>
          </div>

          <button onClick={handleLogout} className="p-2 text-slate-300 hover:text-white rounded-lg">
            <LogOut className="w-5 h-5" />
          </button>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 py-6 flex-1 w-full flex flex-col md:flex-row gap-6">
        <aside className="w-full md:w-64 bg-white rounded-xl border p-4 h-fit space-y-4">
          <button 
            onClick={() => { setActiveTab('cae'); setSelectedCaeCompanyId(null); }}
            className="w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2 bg-amber-500 text-white font-bold"
          >
            <HardHat className="w-4 h-4" />
            <span>Módulo CAE - Empresas</span>
          </button>
        </aside>

        <main className="flex-1">
          {activeTab === 'cae' && (
            <div className="space-y-6">
              {!selectedCaeRecord ? (
                <div className="space-y-6">
                  <div className="bg-amber-600 text-white p-6 rounded-xl flex items-center justify-between">
                    <div>
                      <h2 className="text-xl font-bold flex items-center gap-2">
                        <HardHat className="w-6 h-6" /> Módulo CAE - Empresas
                      </h2>
                      <p className="text-xs text-amber-100 mt-1">Conectado a Google Sheets</p>
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

                  <div className="flex flex-col space-y-3">
                    {filteredCaeRecords.map(record => (
                      <div 
                        key={record.id}
                        onClick={() => setSelectedCaeCompanyId(record.id)}
                        className="bg-white p-5 rounded-xl border hover:border-amber-400 transition cursor-pointer flex items-center justify-between"
                      >
                        <div>
                          <h3 className="text-base font-bold text-slate-800">{record.companyName}</h3>
                          <p className="text-xs text-slate-400">{record.userEmail}</p>
                        </div>
                        <ChevronRight className="w-5 h-5 text-slate-400" />
                      </div>
                    ))}
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  <button onClick={() => setSelectedCaeCompanyId(null)} className="inline-flex items-center gap-1.5 text-xs font-bold text-slate-600 bg-white px-3 py-2 rounded-lg border">
                    <ArrowLeft className="w-4 h-4" /> Volver
                  </button>

                  <div className="bg-white p-6 rounded-xl border space-y-6">
                    <div className="border-b pb-4">
                      <h2 className="text-xl font-bold text-slate-800">{selectedCaeRecord.companyName}</h2>
                      <p className="text-xs text-slate-500">{selectedCaeRecord.userEmail}</p>
                    </div>

                    <div className="space-y-4">
                      <div className="flex justify-between items-center">
                        <h3 className="text-xs font-bold text-slate-400 uppercase">Trabajadores</h3>
                        <button onClick={() => setIsAddingWorker(!isAddingWorker)} className="px-3 py-1.5 bg-amber-500 text-white text-xs font-bold rounded-lg">
                          + Añadir
                        </button>
                      </div>

                      {isAddingWorker && (
                        <div className="p-4 bg-amber-50 border rounded-lg flex gap-3">
                          <input type="text" placeholder="Nombre completo" className="px-3 py-1.5 border rounded-lg text-xs w-full" value={newWorkerName} onChange={e => setNewWorkerName(e.target.value)} />
                          <input type="text" placeholder="DNI" className="px-3 py-1.5 border rounded-lg text-xs w-full" value={newWorkerDni} onChange={e => setNewWorkerDni(e.target.value)} />
                          <button onClick={() => handleAddWorker(selectedCaeRecord.id)} className="px-4 py-1.5 bg-amber-600 text-white text-xs font-bold rounded-lg">Guardar</button>
                        </div>
                      )}

                      <div className="space-y-3">
                        {selectedCaeRecord.workers?.map(worker => {
                          const isSigned = !!worker.signedAt;
                          return (
                            <div key={worker.id} className="p-4 rounded-xl border flex items-center justify-between bg-white">
                              <div>
                                <p className="font-bold text-slate-800 text-sm flex items-center gap-2">
                                  {worker.name}
                                  {isSigned && <span className="px-2 py-0.5 bg-emerald-600 text-white text-[10px] font-bold rounded-full">Firmado</span>}
                                </p>
                                <p className="text-xs text-slate-400 font-mono">DNI: {worker.dni}</p>
                              </div>

                              <div className="flex items-center gap-2">
                                {isSigned ? (
                                  <a href={worker.signatureImage} target="_blank" rel="noopener noreferrer">
                                    <img src={worker.signatureImage} alt="Firma" className="h-8 border rounded bg-white px-1" />
                                  </a>
                                ) : (
                                  <button onClick={() => setSigningWorker(worker)} className="px-4 py-2 bg-blue-600 text-white font-bold text-xs rounded-lg">
                                    FIRMAR
                                  </button>
                                )}
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
        </main>
      </div>

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
