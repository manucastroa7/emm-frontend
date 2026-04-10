import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import toast, { Toaster } from 'react-hot-toast';
import {
  FileSpreadsheet,
  Download,
  Upload,
  User,
  LogOut,
  FileText,
  Info,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  Trash2,
  Users,
  Search,
  UserPlus,
  RefreshCw,
  X,
  School,
  LayoutDashboard,
  FileDown,
  Clock,
  Edit
} from 'lucide-react';
import { AnimatePresence, motion } from 'framer-motion';
import { type StudentData as BaseStudentData, getSubjectsByLicencia } from './services/pdfService';
import logo from './assets/logo.png';
import logoHorizontal from './assets/logo_horizontal.png';

type UserRole = 'admin' | 'editor' | 'viewer' | 'student';

interface UserProfile {
  id: string;
  email?: string;
  documento?: string;
  role: UserRole;
  name: string;
  licencia?: string;
  permissions?: Record<string, string>;
}

interface StudentData extends BaseStudentData {
  email?: string;
  apellido?: string;
  nacionalidad?: string;
  carrera_licencia?: string;
  comision?: string;
  situacion?: string;
  estado_analitico?: 'borrador' | 'emitido';
  diploma_emitido?: boolean;
  fecha_emision?: string;
  fecha_fin_cursada?: string;
  historial?: any[];
}

class ErrorBoundary extends React.Component<{ children: React.ReactNode }, { hasError: boolean }> {
  constructor(props: any) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() { return { hasError: true }; }
  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen flex items-center justify-center bg-red-50 p-4">
          <div className="bg-white p-8 rounded-xl shadow-lg border border-red-200 text-center">
            <h1 className="text-2xl font-bold text-red-600 mb-4">Algo salió mal</h1>
            <p className="text-slate-600 mb-6">Hubo un error al cargar la aplicación.</p>
            <button onClick={() => window.location.reload()} className="bg-red-600 text-white px-6 py-2 rounded-lg">
              Recargar página
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

function AppContent() {
  const API_URL = import.meta.env.VITE_API_URL || 'https://analiticos-backend-production.up.railway.app';

  const [activeTab, setActiveTab] = useState<'dashboard' | 'alumnos' | 'usuarios' | 'horas-campo'>('dashboard');
  const [user, setUser] = useState<UserProfile | null>(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [students, setStudents] = useState<StudentData[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [completenessFilter, setCompletenessFilter] = useState<'all' | 'completos' | 'incompletos'>('all');
  const [licenciaFilter, setLicenciaFilter] = useState('all');
  const [comisionFilter, setComisionFilter] = useState('all');
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(25);
  const [backendStatus, setBackendStatus] = useState<'checking' | 'ok' | 'error'>('checking');
  const [importConfig, setImportConfig] = useState<{ isOpen: boolean; mode: 'db' | 'zip' | null }>({ isOpen: false, mode: null });
  const [diplomaModal, setDiplomaModal] = useState<{isOpen: boolean, student: StudentData | null}>({ isOpen: false, student: null });

  // Gestión de usuarios
  const [appUsers, setAppUsers] = useState<any[]>([]);
  const [editingUser, setEditingUser] = useState<any | null>(null);
  const [newUserNombre, setNewUserNombre] = useState('');
  const [newUserEmail, setNewUserEmail] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserRole, setNewUserRole] = useState<'admin' | 'editor' | 'viewer'>('editor');
  const [newUserPermissions, setNewUserPermissions] = useState<Record<string, string>>({
    'analiticos': 'none',
    'horas-campo': 'none'
  });

  // Alta manual de alumno
  const [newStudentModal, setNewStudentModal] = useState(false);
  const [newStuLicencia, setNewStuLicencia] = useState<'CB' | 'A' | 'PRO' | 'TD1' | 'TD2' | 'ACTUALIZACION'>('CB');
  const [newStuNombre, setNewStuNombre] = useState('');
  const [newStuApellido, setNewStuApellido] = useState('');
  const [newStuDni, setNewStuDni] = useState('');
  const [newStuNacionalidad, setNewStuNacionalidad] = useState('ARGENTINA');
  const [newStuFechaEmision, setNewStuFechaEmision] = useState<string>(new Date().toISOString().split('T')[0]);
  const [newStuFechaFin, setNewStuFechaFin] = useState<string>('');

  // Horas de Campo
  const [entregas, setEntregas] = useState<any[]>([]);

  const [uploadingDoc, setUploadingDoc] = useState<string | null>(null);
  const [fieldHourSubTab, setFieldHourSubTab] = useState<'summary' | 'general' | 'descarga' | 'tp' | 'entrega' | 'dashboard'>('summary');
  const [fieldHourStats, setFieldHourStats] = useState<{
    totalStudents: number;
    pendingReviews: number;
    completedStudents: number;
    recentActivity: any[];
  } | null>(null);

  // Modal de confirmación personalizado
  const [confirmModal, setConfirmModal] = useState<{ open: boolean; title: string; message: string; onConfirm: (val: string) => void; type?: 'danger' | 'warning' | 'info'; withInput?: boolean; inputLabel?: string; inputPlaceholder?: string }>({ open: false, title: '', message: '', onConfirm: () => { } });
  const [confirmInput, setConfirmInput] = useState('');
  const [selectedStudents, setSelectedStudents] = useState<string[]>([]);

  const [selectedStudent, setSelectedStudent] = useState<StudentData | null>(null);
  const [editingDatos, setEditingDatos] = useState(false);
  const [editDni, setEditDni] = useState('');
  const [editNombre, setEditNombre] = useState('');
  const [editApellido, setEditApellido] = useState('');
  const [editNacionalidad, setEditNacionalidad] = useState('');
  const [editEmail, setEditEmail] = useState('');
  const [showAddNota, setShowAddNota] = useState(false);
  const [newMateria, setNewMateria] = useState('');
  const [newNota, setNewNota] = useState('');

  const isAnaliticoCompleto = (student: StudentData) => {
    if (!student) return false;
    if ((student.licencia || '').toUpperCase() === 'ACTUALIZACION') return true;
    const required = getSubjectsByLicencia(student.licencia || '');
    if (required.length === 0) return false;

    const normalize = (s: string) => s.toUpperCase().replace(/,/g, '').replace(/\s+/g, ' ').trim();

    return required.every(sub => {
      const normSub = normalize(sub);
      return student.notas?.some(n => normalize(n.materia) === normSub && n.nota > 0);
    });
  };

  // Mapa de notas pendientes de guardar (materia -> valor editado)
  const [pendingNotas, setPendingNotas] = useState<Record<string, string>>({});
  const [savingNota, setSavingNota] = useState<string | null>(null);

  useEffect(() => {
    const handleEsc = (e: KeyboardEvent) => {
      if (e.key === 'Escape') closeStudentModal();
    };
    window.addEventListener('keydown', handleEsc);
    return () => window.removeEventListener('keydown', handleEsc);
  }, []);

  useEffect(() => {
    fetch(`${API_URL}/api/health`)
      .then(res => res.json())
      .then(data => {
        if (data.status === 'ok') setBackendStatus('ok');
        else setBackendStatus('error');
      })
      .catch(() => setBackendStatus('error'));

    // Restaurar sesión si existe
    const saved = localStorage.getItem('mm-user');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        setUser(parsed);
        
        const hasHorasCampo = parsed.role === 'admin' || (parsed.permissions && parsed.permissions['horas-campo'] && parsed.permissions['horas-campo'] !== 'none');
        const hasAnaliticos = parsed.role === 'admin' || (parsed.permissions && parsed.permissions['analiticos'] && parsed.permissions['analiticos'] !== 'none');

        if (parsed.role === 'student' || (hasHorasCampo && !hasAnaliticos)) {
          setActiveTab('horas-campo');
          if (parsed.role === 'student' || (parsed.permissions && parsed.permissions['horas-campo'] === 'admin')) {
            setFieldHourSubTab(parsed.role === 'student' ? 'summary' : 'dashboard');
          } else {
            setFieldHourSubTab('summary');
          }
        } else if (parsed.role !== 'admin' && hasAnaliticos) {
          setActiveTab('alumnos');
        } else if (parsed.role === 'admin') {
          // Keep default or set to dashboard
          setActiveTab('dashboard');
        }
        fetchStudents();
      } catch (err) {
        console.error('Error al restaurar sesión:', err);
        localStorage.removeItem('mm-user');
      }
    }
  }, []);

  useEffect(() => {
    if (activeTab === 'horas-campo') {
      fetchEntregas();
      fetchTemplates();
      const isFieldHourAdmin = user?.role === 'admin' || user?.permissions?.['horas-campo'] === 'admin';
      if (isFieldHourAdmin) {
        fetchFieldHourStats();
      }
    }
  }, [activeTab, user]);

  const fetchFieldHourStats = async () => {
    try {
      console.log('Fetching Horas de Campo stats...');
      const res = await fetch(`${API_URL}/api/horas-campo/admin/stats`);
      if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
      const data = await res.json();
      console.log('Stats received:', data);
      if (data.success) {
        setFieldHourStats(data.data);
      } else {
        console.error('API Error:', data.error);
        toast.error('No se pudieron cargar las estadísticas');
      }
    } catch (err) {
      console.error('Fetch error:', err);
      toast.error('Error de conexión con el servidor (Estadísticas)');
    }
  };


  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, statusFilter, licenciaFilter, comisionFilter]);

  // Actualizar detalle mientras el modal esté abierto (incluye cambios de otros usuarios)
  useEffect(() => {
    if (!selectedStudent?.id) return;
    const id = selectedStudent.id;
    loadStudentDetail(id);
    const interval = setInterval(() => loadStudentDetail(id), 5000);
    return () => clearInterval(interval);
  }, [selectedStudent?.id]);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    try {
      const res = await fetch(`${API_URL}/api/auth/login`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ email: email.trim(), password })
      });
      const data = await res.json();
      if (!res.ok) {
        setError(data.error || 'Credenciales inválidas');
        return;
      }
      const userData: UserProfile = { 
        id: data.id, 
        email: data.email, 
        documento: data.documento,
        role: data.role, 
        name: data.nombre,
        licencia: data.licencia,
        permissions: data.permissions || {}
      };
      setUser(userData);
      localStorage.setItem('mm-user', JSON.stringify(userData));
      
      const hasHorasCampo = data.role === 'admin' || (data.permissions && data.permissions['horas-campo'] && data.permissions['horas-campo'] !== 'none');
      const hasAnaliticos = data.role === 'admin' || (data.permissions && data.permissions['analiticos'] && data.permissions['analiticos'] !== 'none');

      if (data.role === 'student' || (hasHorasCampo && !hasAnaliticos)) {
        setActiveTab('horas-campo');
        if (data.role === 'student' || (data.permissions && data.permissions['horas-campo'] === 'admin')) {
          setFieldHourSubTab(data.role === 'student' ? 'summary' : 'dashboard');
          if (data.role !== 'student') fetchFieldHourStats();
        } else {
          setFieldHourSubTab('summary');
        }
      } else if (data.role !== 'admin' && hasAnaliticos) {
        setActiveTab('alumnos');
      } else if (data.role === 'admin') {
        setActiveTab('dashboard');
      } else {
        // Fallback for users with no permissions
        setActiveTab('dashboard');
      }
      setError('');
    } catch (err) {
      setError('Error de conexión con el servidor');
    }
  };

  const fetchStudents = async () => {
    try {
      const res = await fetch(`${API_URL}/api/students`);
      const data = await res.json();
      if (data.data) {
        setStudents(data.data.map((s: any) => ({
          ...s,
          fecha: new Date().toLocaleDateString('es-AR')
        })));
      }
    } catch (err) {
      console.error(err);
    }
  };

  const loadStudentDetail = async (id: string) => {
    try {
      const res = await fetch(`${API_URL}/api/students/${id}`);
      const data = await res.json();
      if (data && data.data) {
        setSelectedStudent(data.data);
      }
    } catch (err) {
      console.warn('No se pudo cargar el detalle del alumno');
    }
  };

  const getUserQuery = () => `user=${encodeURIComponent(user?.email || user?.documento || 'Sistema')}&nombre=${encodeURIComponent(user?.name || '')}`;

  const fetchEntregas = async () => {
    if (!user) return;
    try {
      const url = user.role === 'student' 
        ? `${API_URL}/api/horas-campo/mis-entregas?studentId=${user.id}`
        : `${API_URL}/api/horas-campo/admin/todas`;
      const res = await fetch(url);
      const data = await res.json();
      setEntregas(data.data || []);
    } catch (err) {
      console.warn('Error fetching entregas');
    }
  };

  const fetchTemplates = async () => {
    try {
      const res = await fetch(`${API_URL}/api/horas-campo/templates`);
      const data = await res.json();
      setTemplates(data.data || []);
    } catch (err) {
      console.warn('Error fetching templates');
    }
  };

  const handleUploadFieldHour = async (file: File, docType: string) => {
    if (!user) return;
    setUploadingDoc(docType);
    const formData = new FormData();
    formData.append('file', file);
    formData.append('studentId', user.id);
    formData.append('documentType', docType);

    try {
      const res = await fetch(`${API_URL}/api/horas-campo/upload`, {
        method: 'POST',
        body: formData
      });
      if (res.ok) {
        toast.success('Archivo subido correctamente');
        fetchEntregas();
      } else {
        const d = await res.json();
        toast.error(d.error || 'Error al subir');
      }
    } catch (err) {
      toast.error('Error de conexión');
    } finally {
      setUploadingDoc(null);
    }
  };

  const handleUpdateEntregaEstado = async (id: string, status: string, observations?: string) => {
    try {
      const res = await fetch(`${API_URL}/api/horas-campo/${id}/estado`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status, observations, user: user?.email })
      });
      if (res.ok) {
        toast.success(`Estado actualizado a ${status}`);
        fetchEntregas();
      }
    } catch (err) {
      toast.error('Error al actualizar estado');
    }
  };

  const fetchAppUsers = async () => {
    try {
      const res = await fetch(`${API_URL}/api/users`);
      const data = await res.json();
      setAppUsers(Array.isArray(data) ? data : []);
    } catch (err) {
      console.warn('No se pudieron cargar los usuarios');
    }
  };

  const handleCreateUser = async () => {
    if (!newUserEmail || !newUserNombre || (!editingUser && !newUserPassword)) {
      toast.error('Completar todos los campos');
      return;
    }

    const payload: any = { 
      email: newUserEmail, 
      nombre: newUserNombre, 
      role: newUserRole,
      permissions: newUserPermissions
    };
    if (newUserPassword) payload.password = newUserPassword;

    const url = editingUser ? `${API_URL}/api/users/${editingUser.id}` : `${API_URL}/api/users`;
    const method = editingUser ? 'PUT' : 'POST';

    const res = await fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (!res.ok) { 
      toast.error(data.error || 'Error al procesar solicitud');
      return; 
    }
    
    toast.success(editingUser ? 'Usuario actualizado' : 'Usuario creado');
    setEditingUser(null);
    setNewUserEmail('');
    setNewUserNombre('');
    setNewUserPassword('');
    setNewUserRole('editor');
    setNewUserPermissions({ 'analiticos': 'none', 'horas-campo': 'none' });
    await fetchAppUsers();
  };

  const handleDeleteUser = async (id: string) => {
    if (!confirm('\u00bfEliminar este usuario?')) return;
    await fetch(`${API_URL}/api/users/${id}`, { method: 'DELETE' });
    await fetchAppUsers();
  };

  const closeStudentModal = () => {
    setSelectedStudent(null);
    setEditingDatos(false);
    setShowAddNota(false);
    setEditDni('');
    setEditNombre('');
    setEditApellido('');
    setEditNacionalidad('');
    setEditEmail('');
    setPendingNotas({});
  };

  const handleLogout = () => {
    setUser(null);
    localStorage.removeItem('mm-user');
    setEmail('');
    setPassword('');
    setStudents([]);
    closeStudentModal();
    setAppUsers([]);
    setActiveTab('dashboard');
  };

  const handleCreateStudent = async (e: React.FormEvent) => {
    e.preventDefault();
    if (user?.role === 'viewer') { toast.error('Solo lectura: no puedes crear alumnos.'); return; }
    if (!newStuNombre || !newStuApellido || !newStuDni) { toast.error('Completa nombre, apellido y documento'); return; }
    setIsUploading(true);
    try {
      const payload = {
        nombre: newStuNombre,
        apellido: newStuApellido,
        documento: newStuDni,
        nacionalidad: newStuNacionalidad,
        licencia: newStuLicencia,
        fecha_emision: newStuFechaEmision,
        fecha_fin_cursada: newStuLicencia === 'ACTUALIZACION' ? undefined : newStuFechaFin
      };
      const res = await fetch(`${API_URL}/api/students?${getUserQuery()}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'Error creando alumno');
      toast.success('Alumno creado correctamente');
      setNewStudentModal(false);
      setNewStuNombre('');
      setNewStuApellido('');
      setNewStuDni('');
      setNewStuNacionalidad('ARGENTINA');
      setNewStuFechaFin('');
      setNewStuFechaEmision(new Date().toISOString().split('T')[0]);
      await fetchStudents();
    } catch (err: any) {
      toast.error(err.message || 'Error creando alumno');
    } finally {
      setIsUploading(false);
    }
  };

  const handleModalSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const form = new FormData(e.currentTarget);
    const file = form.get('excelFile') as File;
    if (!file || file.size === 0) { toast.error('Selecciona un archivo excel'); return; }

    setIsUploading(true);

    const submitData = new FormData();
    submitData.append(importConfig.mode === 'db' ? 'file' : 'excelFile', file);
    submitData.append('licencia', form.get('licencia') as string);
    submitData.append('fecha_fin_cursada', form.get('fecha_fin_cursada') as string);
    submitData.append('fecha_emision', form.get('fecha_emision') as string);

    try {
      if (importConfig.mode === 'db') {
        const res = await fetch(`${API_URL}/api/process-excel?${getUserQuery()}`, { method: 'POST', body: submitData });
        if (!res.ok) {
          const err = await res.json();
          throw new Error(err.error || 'Error al importar en servidor');
        }
        const result = await res.json();
        toast.success(`Se vincularon notas a ${result.data.matchCount} alumnos. ${result.data.noMatchCount} IDs sin registro.`);
        await fetchStudents();
      } else {
        const res = await fetch(`${API_URL}/api/generate-certificates`, { method: 'POST', body: submitData });
        if (!res.ok) {
          const err = await res.json();
          throw new Error(err.error || 'Error en generación de PDFs masiva');
        }
        const data = await res.json();
        if (data.success && data.downloadUrl) {
          window.location.href = `${API_URL}${data.downloadUrl}`;
          setTimeout(() => toast.success('Certificados generados. La descarga comenzará enseguida.'), 500);
        } else throw new Error('No se recibió la URL de descarga.');
      }
      setImportConfig({ isOpen: false, mode: null });
    } catch (err: any) {
      console.error(err);
      toast.error(err.message || 'Error de conexión al procesar');
    } finally {
      setIsUploading(false);
    }
  };

  const handleQuinttosUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsUploading(true);
    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch(`${API_URL}/api/quinttos?${getUserQuery()}`, {
        method: 'POST',
        body: formData,
      });
      const result = await response.json();

      if (result.success) {
        toast.success(result.message);
        await fetchStudents();
      } else {
        toast.error(result.error || 'Error al procesar padrón QUINTTOS');
      }
    } catch (err) {
      console.error(err);
      alert('Error de conexión con el servidor');
    } finally {
      setIsUploading(false);
    }
  };

  const handleDelete = async (id?: string) => {
    if (!id) return;
    setConfirmModal({
      open: true,
      title: 'Eliminar Alumno',
      message: '¿Estás seguro de que deseas eliminar este registro? Esta acción no se puede deshacer.',
      type: 'danger',
      onConfirm: async (_val: string) => {
        try {
          await fetch(`${API_URL}/api/students/${id}?${getUserQuery()}`, { method: 'DELETE' });
          setStudents(prev => prev.filter(s => s.id !== id));
          if (selectedStudent?.id === id) closeStudentModal();
        } catch (err) {
          toast.error('Error al eliminar');
        }
        setConfirmModal(prev => ({ ...prev, open: false }));
      }
    });
  };

  const handleDeleteNotas = async (id?: string) => {
    if (!id) return;
    setConfirmModal({
      open: true,
      title: 'Eliminar Notas',
      message: '¿Estás seguro de que deseas eliminar las notas de este alumno? Volverá a borrador.',
      type: 'warning',
      onConfirm: async (_val: string) => {
        try {
          await fetch(`${API_URL}/api/students/${id}/notas?${getUserQuery()}`, { method: 'DELETE' });
          await fetchStudents();
          if (selectedStudent?.id === id) {
            setSelectedStudent(prev => prev ? ({ ...prev, notas: [], estado_analitico: 'borrador', promedio: 0 }) : null);
          }
          toast.success('Notas eliminadas correctamente.');
        } catch (err) {
          toast.error('Error al eliminar notas');
        }
        setConfirmModal(prev => ({ ...prev, open: false }));
      }
    });
  };

  const toggleSelectStudent = (id: string) => {
    setSelectedStudents(prev => 
      prev.includes(id) ? prev.filter(i => i !== id) : [...prev, id]
    );
  };

  const toggleSelectAll = () => {
    const currentIds = filteredStudents
      .slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage)
      .map(s => s.id)
      .filter((id): id is string => !!id);
    
    if (currentIds.every(id => selectedStudents.includes(id))) {
      setSelectedStudents(prev => prev.filter(id => !currentIds.includes(id)));
    } else {
      setSelectedStudents(prev => [...new Set([...prev, ...currentIds])]);
    }
  };

  const handleGenerateDiploma = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!diplomaModal.student) return;
    if (user?.role === 'viewer') { toast.error('Solo lectura: no puedes emitir diplomas.'); return; }
    const isAct = (diplomaModal.student.licencia || '').toUpperCase() === 'ACTUALIZACION';
    if (!isAct && !isAnaliticoCompleto(diplomaModal.student)) { toast.error('Analítico incompleto: faltan notas obligatorias.'); return; }

    const formData = new FormData(e.currentTarget);
    const data = {
      nacionalidad: formData.get('nacionalidad'),
      fecha_emision: formData.get('fecha_emision'),
      nombre: diplomaModal.student.nombre,
      apellido: diplomaModal.student.apellido
    };

    try {
      setIsUploading(true);
      const res = await fetch(`${API_URL}/api/students/${diplomaModal.student.id}/diploma`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      if (!res.ok) throw new Error('Error al generar diploma');

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Diploma_${diplomaModal.student.apellido}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      
      setDiplomaModal({ isOpen: false, student: null });
      // Refresh student data to show new nationality/date
      fetchStudents();
    } catch (err) {
      console.error(err);
      alert('Error generando el diploma');
    } finally {
      setIsUploading(false);
    }
  };

  const handleBulkDelete = async () => {
    if (selectedStudents.length === 0) return;
    setConfirmModal({
      open: true,
      title: 'Eliminar Seleccionados',
      message: `¿Estás seguro de que deseas eliminar los ${selectedStudents.length} alumnos seleccionados? Esta acción es permanente.`,
      type: 'danger',
      onConfirm: async () => {
        setIsUploading(true);
        try {
          const res = await fetch(`${API_URL}/api/students/bulk?${getUserQuery()}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: selectedStudents })
          });
          const data = await res.json();
          if (res.ok) {
            toast.success(data.message || 'Alumnos eliminados correctamente.');
            setSelectedStudents([]);
            fetchStudents();
          } else {
            toast.error(data.error || 'Error al eliminar alumnos.');
          }
        } catch (err) {
          toast.error('Error de conexión.');
        } finally {
          setIsUploading(false);
          setConfirmModal(prev => ({ ...prev, open: false }));
        }
      }
    });
  };

  const handleResetDatabase = async () => {
    setConfirmInput('');
    setConfirmModal({
      open: true,
      title: 'Reiniciar Base de Datos',
      message: '¡ATENCIÓN! Esta acción eliminará a TODOS los alumnos, sus notas y su historial permanentemente. Para confirmar, escribí "BORRAR TODO" en el campo de abajo:',
      type: 'danger',
      withInput: true,
      inputLabel: 'Confirmación',
      inputPlaceholder: 'BORRAR TODO',
      onConfirm: async (val: string) => {
        if (val.trim() !== 'BORRAR TODO') {
          toast.error('La palabra de confirmación no coincide.');
          return;
        }
        setIsUploading(true);
        try {
          const res = await fetch(`${API_URL}/api/database/reset?${getUserQuery()}`, { method: 'DELETE' });
          const data = await res.json();
          if (res.ok) {
            toast.success(data.message || 'Base de datos reiniciada.');
            fetchStudents();
            closeStudentModal();
          } else {
            toast.error(data.error || 'Error al reiniciar');
          }
        } catch (err) {
          toast.error('Error de conexión');
        } finally {
          setIsUploading(false);
          setConfirmModal(prev => ({ ...prev, open: false }));
        }
      }
    });
  };


  const handleToggleEstado = async (id?: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    if (!id) return;
    const nuevoEstado = selectedStudent?.estado_analitico === 'emitido' ? 'borrador' : 'emitido';

    // VALIDACION: Solo si pasa a emitido
    if (nuevoEstado === 'emitido' && selectedStudent) {
      if (!isAnaliticoCompleto(selectedStudent)) {
        toast.error('⚠️ Analítico Incompleto: Faltan materias obligatorias.');
      }
    }

    const executeToggle = async (motivoAdicional = '') => {
      try {
        const res = await fetch(`${API_URL}/api/students/${id}/estado?${getUserQuery()}`, {
          method: 'PUT',
          body: JSON.stringify({ estado: nuevoEstado, motivo: motivoAdicional }),
          headers: { 'Content-Type': 'application/json' }
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.error);

        await fetchStudents();
        if (selectedStudent?.id === id) {
          setSelectedStudent(prev => prev ? ({ ...prev, estado_analitico: data.estado }) : null);
          loadStudentDetail(id);
        }
        toast.success(nuevoEstado === 'emitido' ? '¡Analítico Emitido!' : 'Revertido a Borrador');
      } catch (err: any) {
        toast.error(err.message || 'Error al cambiar estado');
      }
    };

    if (nuevoEstado === 'borrador') {
      setConfirmInput('');
      setConfirmModal({
        open: true,
        title: 'Revertir a Borrador',
        message: 'Ingresá el motivo por el cual se revierte el estado:',
        type: 'warning',
        withInput: true,
        inputLabel: 'Motivo',
        inputPlaceholder: 'Ej: Error en el apellido, nota incorrecta, etc.',
        onConfirm: async (val: string) => {
          await executeToggle(val.trim());
          setConfirmModal(prev => ({ ...prev, open: false }));
        }
      });
      return;
    }

    // Si es a 'emitido', ejecutar directo
    await executeToggle();
  };

  const handleSaveFechas = async () => {
    if (!selectedStudent || !selectedStudent.id) return;
    try {
      const res = await fetch(`${API_URL}/api/students/${selectedStudent.id}/dates?${getUserQuery()}`, {
        method: 'PUT',
        body: JSON.stringify({ fecha_emision: selectedStudent.fecha_emision, fecha_fin_cursada: selectedStudent.fecha_fin_cursada }),
        headers: { 'Content-Type': 'application/json' }
      });
      if (res.ok) {
        toast.success('Fechas guardadas correctamente');
        fetchStudents();
      }
    } catch (err) {
      console.error(err);
      toast.error('Error al guardar fechas');
    }
  };

  const startEditDatos = () => {
    if (!selectedStudent) return;
    setEditDni(selectedStudent.dni || '');
    setEditNombre(selectedStudent.nombre);
    setEditApellido(selectedStudent.apellido || '');
    setEditNacionalidad(selectedStudent.nacionalidad || '');
    setEditEmail(selectedStudent.email || '');
    setEditingDatos(true);
  };

  const saveDatos = async () => {
    if (!selectedStudent?.id) return;
    try {
      await fetch(`${API_URL}/api/students/${selectedStudent.id}/datos?${getUserQuery()}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          documento: editDni,
          nombre: editNombre,
          apellido: editApellido,
          nacionalidad: editNacionalidad,
          email: editEmail
        })
      });
      await fetchStudents();
      setEditingDatos(false);
      setSelectedStudent(prev => prev ? ({
        ...prev,
        dni: editDni,
        nombre: `${editNombre} ${editApellido}`.trim(),
        apellido: editApellido,
        nacionalidad: editNacionalidad,
        email: editEmail
      }) : null);
      toast.success('Datos personales guardados correctamente.');
    } catch (err) {
      toast.error('Error al guardar datos personales');
    }
  };


  const saveNotaManual = async () => {
    if (!selectedStudent?.id || !newMateria.trim() || !newNota) return;
    try {
      const res = await fetch(`${API_URL}/api/students/${selectedStudent.id}/nota?${getUserQuery()}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ asignatura: newMateria.trim(), nota: parseFloat(newNota) })
      });
      const data = await res.json();
      if (!res.ok) { alert(data.error || 'Error'); return; }
      await fetchStudents();
      const updated = await fetch(`${API_URL}/api/students`).then(r => r.json());
      const s = updated.data?.find((x: any) => x.id === selectedStudent.id);
      if (s) setSelectedStudent(s);
      setNewMateria('');
      setNewNota('');
      setShowAddNota(false);
    } catch (err) {
      toast.error('Error al guardar la nota');
    }
  };

  // Solo descarga el PDF sin cambiar el estado (Vista Previa)
  const downloadPDF = (student: StudentData) => {
    if (!student.id) { alert("El alumno no tiene ID registrado."); return; }
    if (user?.role === 'viewer') { toast.error('Solo lectura: no puedes generar PDFs.'); return; }
    if (!isAnaliticoCompleto(student)) { toast.error('Analítico incompleto: faltan notas obligatorias.'); return; }
    window.open(`${API_URL}/api/students/${student.id}/certificate`, '_blank');
  };

  /* DESACTIVADO TEMPORALMENTE
  const downloadDiploma = (student: StudentData) => {
    if (!student.id) { alert("El alumno no tiene ID registrado."); return; }
    window.open(`${API_URL}/api/students/${student.id}/diploma`, '_blank');
  };
  */

  // Descarga el PDF y marca automáticamente como Emitido
  const downloadPDFAndEmit = async (student: StudentData) => {
    if (!student.id) { alert("El alumno no tiene ID registrado."); return; }
    if (user?.role === 'viewer') { toast.error('Solo lectura: no puedes emitir.'); return; }
    if (!isAnaliticoCompleto(student)) {
      toast.error('Analítico incompleto: faltan notas obligatorias.');
      return;
    }

    // 1. Descargar Analítico
    downloadPDF(student);
    // setTimeout(() => downloadDiploma(student), 800); // DESACTIVADO TEMPORALMENTE

    // 2. Marcar como Emitido si todavía no lo estaba
    if (student.estado_analitico !== 'emitido') {
      if (!isAnaliticoCompleto(student)) {
        toast.error('⚠️ Generado: Faltan notas obligatorias.');
      }
      try {
        const res = await fetch(
          `${API_URL}/api/students/${student.id}/estado?${getUserQuery()}`,
          { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ estado: 'emitido', motivo: 'Generado desde el sistema' }) }
        );
        const data = await res.json();
        await fetchStudents();
        setSelectedStudent(prev => prev ? ({ ...prev, estado_analitico: data.estado }) : null);
        if (student.id) loadStudentDetail(student.id);
      } catch (err) { console.error('No se pudo marcar como Emitido:', err); }
    }
  };

  const exportToExcel = () => {
    const dataToExport = filteredStudents.map(student => ({
      ID: student.dni,
      Nombre: student.nombre,
      Email: student.email || '',
      'Licencia/Carrera': student.licencia || '',
      Comision: student.comision || '',
      'Situacion Academica': student.situacion || '',
      'Estado Analitico': student.estado_analitico === 'emitido' ? 'Emitido' : 'Borrador',
      Promedio: student.notas && student.notas.length > 0 ? student.promedio.toFixed(2) : 'Sin notas',
      'Materias Aprobadas': student.notas?.length || 0,
      'Fecha Emision': student.fecha_emision || '',
      'Fecha Fin Cursada': student.fecha_fin_cursada || '',
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Alumnos");
    XLSX.writeFile(workbook, "Listado_Alumnos.xlsx");
  };

  const uniqueStatuses = Array.from(new Set(students.map(s => s.situacion).filter(Boolean)));
  const uniqueLicencias = Array.from(new Set(students.map(s => s.licencia).filter(Boolean)));
  const uniqueComisiones = Array.from(new Set(students.map(s => s.comision).filter(Boolean)));

  const formatLicencia = (lic?: string) => {
    if (!lic) return 'S/LI';
    const upper = lic.toUpperCase();
    if (upper.includes('TRAYECTORIA DESTACADA') && upper.includes('1')) return 'TD 1';
    if (upper.includes('TRAYECTORIA DESTACADA') && upper.includes('2')) return 'TD 2';
    return lic;
  };



  const formatUserDisplay = (email?: string, nombre?: string) => {
    if (nombre && nombre.trim()) return nombre;
    if (!email) return 'Sin usuario';
    if (email.includes('@')) return email.split('@')[0];
    return email;
  };

  const filteredStudents = students.filter(s => {
    const term = searchTerm.toLowerCase();
    const matchesSearch =
      s.nombre.toLowerCase().includes(term) ||
      (s.apellido || '').toLowerCase().includes(term) ||
      s.dni.includes(searchTerm);
    const matchesStatus = statusFilter === 'all' || s.situacion === statusFilter;
    const matchesLicencia = licenciaFilter === 'all' || (s.licencia || '').toUpperCase() === licenciaFilter.toUpperCase();
    const matchesComision = comisionFilter === 'all' || (s.comision || '').toUpperCase() === comisionFilter.toUpperCase();
    const completo = isAnaliticoCompleto(s);
    const matchesCompleteness =
      completenessFilter === 'all' ||
      (completenessFilter === 'completos' && completo) ||
      (completenessFilter === 'incompletos' && !completo);
    return matchesSearch && matchesStatus && matchesLicencia && matchesComision && matchesCompleteness;
  });

  if (!user) {
    return (
      <div
        className="min-h-screen flex items-center justify-center p-6"
        style={{
          backgroundImage: `
            radial-gradient(circle at 20% 20%, rgba(15, 90, 92, 0.28), transparent 26%),
            radial-gradient(circle at 80% 10%, rgba(12, 64, 78, 0.38), transparent 22%),
            radial-gradient(110% 120% at 50% 0%, rgba(4, 18, 24, 0.9), rgba(4, 12, 18, 0.95)),
            linear-gradient(135deg, #031219 0%, #071f2a 60%, #031017 100%)`,
          backgroundColor: '#041218'
        }}
      >
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="w-full max-w-xl bg-[#041218cc] backdrop-blur-xl border border-white/5 shadow-[0_40px_120px_-40px_rgba(0,0,0,0.9)] rounded-[28px] overflow-hidden"
        >
          <div className="flex flex-col items-center gap-3 px-10 pt-10 pb-6 text-center">
            <img src={logo} alt="Escuela Maradona Menotti" className="h-24 w-auto object-contain drop-shadow-[0_16px_34px_rgba(0,45,43,0.45)]" />
            <div className="px-4 py-1.5 text-[11px] font-black uppercase tracking-[0.18em] text-white bg-[#00968f]/70 border border-[#0ffff4]/50 rounded-full shadow-sm drop-shadow">
              Acceso seguro
            </div>
          </div>

          <form onSubmit={handleLogin} className="space-y-4 px-10 pb-10 text-white">
            <div className="space-y-1">
              <label className="block text-sm font-semibold text-white/90">Usuario / Documento</label>
              <input
                type="text"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                className="w-full px-4 py-3 rounded-xl border border-[#0ffff4]/20 bg-white/88 text-slate-800 placeholder:text-slate-500 focus:ring-2 focus:ring-[#00968f] focus:border-transparent outline-none transition-all shadow-[0_10px_35px_-30px_rgba(0,0,0,0.35)]"
                placeholder="DNI o email institucional"
                required
              />
            </div>
            <div className="space-y-1">
              <label className="block text-sm font-semibold text-white/90">Contraseña</label>
              <div className="relative">
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full px-4 py-3 rounded-xl border border-[#0ffff4]/20 bg-white/88 text-slate-800 placeholder:text-slate-500 focus:ring-2 focus:ring-[#00968f] focus:border-transparent outline-none transition-all shadow-[0_10px_35px_-30px_rgba(0,0,0,0.35)]"
                  placeholder="Ingresa tu contraseña"
                  required
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(prev => !prev)}
                  className="absolute inset-y-0 right-3 my-1 px-3 text-xs font-bold text-[#002d2b] bg-[#0ffff4]/30 hover:bg-[#0ffff4]/50 rounded-lg transition-colors"
                >
                  {showPassword ? 'Ocultar' : 'Mostrar'}
                </button>
              </div>
            </div>
            {error && <p className="text-red-500 text-xs mt-1 font-semibold">{error}</p>}
            <button
              type="submit"
              disabled={backendStatus === 'checking'}
              className="w-full bg-gradient-to-r from-[#0ffff4] to-[#00968f] disabled:from-slate-400 disabled:to-slate-500 text-[#002d2b] font-extrabold py-3.5 rounded-xl transition-all shadow-[0_16px_50px_-22px_rgba(0,255,244,0.35)] hover:-translate-y-0.5 active:scale-[0.98]"
            >
              {backendStatus === 'checking' ? 'Iniciando...' : 'Ingresar al Sistema'}
            </button>
          </form>

          <div className="px-10 pb-10">
            <div className="rounded-2xl bg-[#0ffff4]/12 border border-[#0ffff4]/30 px-4 py-3 text-center">
              <p className="text-xs text-white/90 font-medium">Acceso restringido a personal autorizado y alumnos matriculados.</p>
            </div>
          </div>
        </motion.div>
      </div>
    );
  }

  const renderDashboard = () => {
    const totalAlumnos = students.length;
    const conAnalitico = students.filter(s => s.notas && s.notas.length > 0).length;
    const sinAnalitico = totalAlumnos - conAnalitico;

    // Agrupar por Carrera
    const porCarrera = students.reduce((acc, curr) => {
      const c = curr.licencia || 'Sin Carrera Especificada';
      acc[c] = (acc[c] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    // Emitidos por licencia
    const emitidosPorLicencia = students
      .filter(s => s.estado_analitico === 'emitido')
      .reduce((acc, curr) => {
        const c = curr.licencia || 'Sin licencia';
        acc[c] = (acc[c] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

    // Analíticos completos por licencia
    const completosPorLicencia = students
      .filter(s => isAnaliticoCompleto(s))
      .reduce((acc, curr) => {
        const c = curr.licencia || 'Sin licencia';
        acc[c] = (acc[c] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

    // Últimos 5 días emitidos
    const today = new Date();
    const recentEmitidos = students
      .filter(s => s.estado_analitico === 'emitido' && s.fecha_emision)
      .filter(s => {
        const f = new Date(s.fecha_emision as string);
        const diff = (today.getTime() - f.getTime()) / (1000 * 60 * 60 * 24);
        return diff <= 5;
      })
      .sort((a, b) => new Date(b.fecha_emision || '').getTime() - new Date(a.fecha_emision || '').getTime());

    return (
      <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
        <h2 className="text-3xl font-bold text-slate-900 drop-shadow-sm">Dashboard General</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm hover:shadow-md transition-shadow">
            <h3 className="text-slate-500 text-sm font-medium mb-3 uppercase tracking-wider">Total Alumnos</h3>
            <p className="text-4xl font-extrabold text-[#002d2b]">{totalAlumnos}</p>
          </div>
          <div className="bg-gradient-to-br from-emerald-50 to-emerald-100 p-6 rounded-2xl border border-emerald-200 shadow-sm hover:shadow-md transition-shadow">
            <h3 className="text-emerald-700 text-sm font-medium mb-3 uppercase tracking-wider">Con Analítico</h3>
            <p className="text-4xl font-extrabold text-emerald-800">{conAnalitico}</p>
          </div>
          <div className="bg-gradient-to-br from-amber-50 to-amber-100 p-6 rounded-2xl border border-amber-200 shadow-sm hover:shadow-md transition-shadow">
            <h3 className="text-amber-700 text-sm font-medium mb-3 uppercase tracking-wider">Esperando Notas</h3>
            <p className="text-4xl font-extrabold text-amber-800">{sinAnalitico}</p>
          </div>
        </div>

        {/* Vista colapsada manual eliminada para favorecer navegación por sidebar */}


        <div className="pt-6 border-t border-slate-200">
          <h3 className="text-xl font-bold text-slate-800 mb-6 drop-shadow-sm">Distribucion por Licencia / Carrera</h3>
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
            {Object.entries(porCarrera).map(([carr, count]) => (
              <div key={carr} className="bg-white p-5 rounded-2xl border border-slate-200 shadow-sm flex items-center justify-between hover:bg-slate-50 transition-colors">
                <span className="font-semibold text-slate-700 line-clamp-1" title={carr}>{carr}</span>
                <span className="bg-[#0ffff4]/25 text-[#002d2b] px-4 py-1.5 rounded-full text-sm font-black shadow-inner border border-[#0ffff4]/50">{count}</span>
              </div>
            ))}
            {Object.keys(porCarrera).length === 0 && (
              <p className="text-slate-500 italic">Cargue el padron QUINTTOS para ver estadisticas.</p>
            )}
          </div>
        </div>

        <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
            <h3 className="text-lg font-bold text-slate-800 mb-4">Emitidos por Licencia</h3>
            <div className="space-y-3">
              {Object.entries(emitidosPorLicencia).map(([lic, count]) => (
                <div key={lic} className="flex items-center justify-between px-3 py-2 rounded-lg border border-slate-100 bg-slate-50/60">
                  <span className="font-semibold text-slate-700">{lic}</span>
                  <span className="text-sm font-black text-[#002d2b] bg-[#0ffff4]/20 border border-[#0ffff4]/40 rounded-full px-3 py-1">{count}</span>
                </div>
              ))}
              {Object.keys(emitidosPorLicencia).length === 0 && (
                <p className="text-slate-500 text-sm">Aún no hay analíticos emitidos.</p>
              )}
            </div>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
            <h3 className="text-lg font-bold text-slate-800 mb-4">Emitidos últimos 5 días</h3>
            <div className="space-y-3">
              {recentEmitidos.slice(0, 8).map((s, idx) => (
                <div key={`${s.id}-${idx}`} className="flex items-center justify-between px-3 py-2 rounded-lg border border-slate-100 bg-[#0ffff4]/10">
                  <div>
                    <p className="font-bold text-slate-800 text-sm">{s.nombre}</p>
                    <p className="text-xs text-slate-500">Licencia {s.licencia || 'S/LI'} · ID {s.dni}</p>
                  </div>
                  <span className="text-[11px] font-black uppercase tracking-wider text-[#00968f]">{s.fecha_emision}</span>
                </div>
              ))}
              {recentEmitidos.length === 0 && (
                <p className="text-slate-500 text-sm">Sin emisiones en los últimos 5 días.</p>
              )}
            </div>
          </div>

          <div className="bg-white p-6 rounded-2xl border border-slate-200 shadow-sm">
            <h3 className="text-lg font-bold text-slate-800 mb-4">Analíticos completos por Licencia</h3>
            <div className="space-y-3">
              {Object.entries(completosPorLicencia).map(([lic, count]) => (
                <div key={lic} className="flex items-center justify-between px-3 py-2 rounded-lg border border-slate-100 bg-slate-50/60">
                  <span className="font-semibold text-slate-700">{lic}</span>
                  <span className="text-sm font-black text-[#00968f] bg-[#0ffff4]/15 border border-[#0ffff4]/40 rounded-full px-3 py-1">{count}</span>
                </div>
              ))}
              {Object.keys(completosPorLicencia).length === 0 && (
                <p className="text-slate-500 text-sm">Sin analíticos completos aún.</p>
              )}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderAlumnos = () => (
    <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
      <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 mb-8 flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-slate-900">Padron de Alumnos y Analiticos</h2>
          <p className="text-slate-500 text-sm">Sincronice el padron e importe notas para generar certificados.</p>
        </div>

        <div className="flex items-center justify-end gap-3 flex-wrap">
          {user.role !== 'viewer' && (
            <button
              onClick={() => setNewStudentModal(true)}
              disabled={isUploading}
              className={`flex items-center gap-2 px-4 py-3 bg-[#0f766e] hover:bg-[#0b5f59] text-white font-medium rounded-xl transition-all shadow hover:-translate-y-0.5 active:scale-[0.98] ${isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}
            >
              <User className="w-4 h-4" />
              <span>Nuevo Alumno</span>
            </button>
          )}

          {user.role === 'admin' && (
            <>
              <button
                onClick={handleResetDatabase}
                disabled={isUploading}
                className={`flex items-center gap-2 px-4 py-3 bg-red-600 hover:bg-red-700 text-white font-medium rounded-xl transition-all shadow ${isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}
                title="Borrar TODOS los alumnos y notas (Reset Total)"
              >
                <Trash2 className="w-4 h-4" />
                <span>Reiniciar BD</span>
              </button>
              {selectedStudents.length > 0 && (
                <button
                  onClick={handleBulkDelete}
                  disabled={isUploading}
                  className={`flex items-center gap-2 px-4 py-3 bg-rose-600 hover:bg-rose-700 text-white font-medium rounded-xl transition-all shadow ${isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}
                  title="Eliminar alumnos seleccionados"
                >
                  <Trash2 className="w-4 h-4" />
                  <span>Eliminar {selectedStudents.length}</span>
                </button>
              )}
              <label className={`flex items-center gap-2 px-4 py-3 bg-[#002d2b] hover:bg-[#00968f] text-white font-medium rounded-xl cursor-pointer transition-all shadow hover:-translate-y-0.5 active:scale-[0.98] ${isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}>
                <Upload className="w-4 h-4" />
                <span>Sincronizar QUINTTOS</span>
                <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleQuinttosUpload} disabled={isUploading} />
              </label>
              <button onClick={() => setImportConfig({ isOpen: true, mode: 'db' })} disabled={isUploading} className={`flex items-center gap-2 px-4 py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-medium rounded-xl transition-all shadow ${isUploading ? 'opacity-50 cursor-not-allowed' : ''}`}>
                <Upload className="w-4 h-4" />
                <span>Importar Notas</span>
              </button>
            </>
          )}
        </div>
      </div>

      <div className="space-y-6">
        <div className="grid grid-cols-1 xl:grid-cols-5 gap-3 mb-3 items-stretch">
          <div className="relative shadow-sm rounded-xl col-span-1 xl:col-span-2 min-w-0">
            <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
            <input
              type="text"
              placeholder="Buscar por nombre, apellido o ID..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full h-14 pl-12 pr-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-[#00968f] outline-none transition-all bg-white font-medium"
            />
          </div>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="min-w-0 h-14 px-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-[#00968f] outline-none transition-all bg-white font-medium text-slate-700 cursor-pointer shadow-sm"
          >
            <option value="all">Todas las situaciones</option>
            {uniqueStatuses.map(st => (
              <option key={st as string} value={st as string}>{st as string}</option>
            ))}
          </select>
          <select
            value={licenciaFilter}
            onChange={(e) => setLicenciaFilter(e.target.value)}
            className="min-w-0 h-14 px-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-[#00968f] outline-none transition-all bg-white font-medium text-slate-700 cursor-pointer shadow-sm"
          >
            <option value="all">Todas las licencias</option>
            {uniqueLicencias.map(l => (
              <option key={l as string} value={l as string}>{l as string}</option>
            ))}
          </select>
          <select
            value={comisionFilter}
            onChange={(e) => setComisionFilter(e.target.value)}
            className="min-w-0 h-14 px-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-[#00968f] outline-none transition-all bg-white font-medium text-slate-700 cursor-pointer shadow-sm"
          >
            <option value="all">Todas las comisiones</option>
            {uniqueComisiones.map(c => (
              <option key={c as string} value={c as string}>{c as string}</option>
            ))}
          </select>
          <div className="flex gap-2 items-center">
            <select
              value={completenessFilter}
              onChange={(e) => setCompletenessFilter(e.target.value as 'all' | 'completos' | 'incompletos')}
              className="min-w-0 h-14 px-4 rounded-xl border border-slate-200 focus:ring-2 focus:ring-[#00968f] outline-none transition-all bg-white font-medium text-slate-700 cursor-pointer shadow-sm flex-1"
            >
              <option value="all">Completos e incompletos</option>
              <option value="completos">Solo completos</option>
              <option value="incompletos">Solo incompletos</option>
            </select>
            <button
              onClick={exportToExcel}
              className="h-14 px-4 bg-[#002d2b] hover:bg-[#00968f] text-white rounded-xl font-bold transition-all shadow-sm whitespace-nowrap hover:-translate-y-0.5 active:scale-[0.98]"
              title="Exportar listado actual a Excel"
            >
              <FileSpreadsheet className="w-5 h-5 inline-block mr-2 align-middle" />
              <span className="align-middle">Exportar</span>
            </button>
          </div>
        </div>

        {/* LIST VIEW INSTED OF GRID */}
        <div className="flex flex-col space-y-3">
          {filteredStudents.length > 0 && (
            <div className="flex items-center gap-4 px-4 py-2 bg-slate-50 rounded-xl border border-slate-200 mb-2">
              <input 
                type="checkbox" 
                className="w-5 h-5 rounded border-slate-300 text-[#002d2b] focus:ring-[#00968f] cursor-pointer"
                checked={filteredStudents.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage).every(s => s.id && selectedStudents.includes(s.id))}
                onChange={toggleSelectAll}
              />
              <span className="text-sm font-bold text-slate-600">Seleccionar todos en esta página</span>
            </div>
          )}
          <AnimatePresence mode="popLayout">
            {filteredStudents.length > 0 ? (
              filteredStudents.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage).map((student, idx) => (
                <motion.div
                  key={student.dni + idx}
                  initial={{ opacity: 0, y: 10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  onClick={() => { setSelectedStudent(student); if (student.id) loadStudentDetail(student.id); }}
                  className={`p-4 rounded-2xl border transition-all flex flex-col md:flex-row md:items-center justify-between gap-4 cursor-pointer group min-h-[110px] ${
                    selectedStudents.includes(student.id || '') 
                      ? 'bg-[#0ffff4]/15 border-[#00968f] shadow-md ring-1 ring-[#00968f]' 
                      : student.situacion === 'DUPLICADO'
                        ? 'bg-rose-50 border-rose-300 hover:border-rose-400 hover:shadow-md'
                        : 'bg-white border-slate-200 shadow-sm hover:shadow-md hover:border-[#00968f33]'
                  }`}
                >
                  {/* Left: Checkbox, Icon, Details */}
                  <div className="flex items-center gap-4">
                    <input 
                      type="checkbox" 
                      className="w-5 h-5 rounded border-slate-300 text-[#002d2b] focus:ring-[#00968f] cursor-pointer"
                      checked={selectedStudents.includes(student.id || '')}
                      onClick={(e) => e.stopPropagation()}
                      onChange={() => toggleSelectStudent(student.id || '')}
                    />
                    <div className={`p-3 rounded-xl hidden sm:flex shrink-0 transition-colors ${
                      selectedStudents.includes(student.id || '')
                        ? 'bg-[#0ffff4]/30 text-[#002d2b]'
                        : student.situacion === 'DUPLICADO' 
                          ? 'bg-rose-100 group-hover:bg-rose-200' 
                          : 'bg-[#0ffff4]/10 group-hover:bg-[#0ffff4]/20'
                    }`}>
                      <User className={`w-5 h-5 ${student.situacion === 'DUPLICADO' ? 'text-rose-700' : 'text-[#002d2b]'}`} />
                    </div>
                    <div className="flex flex-col gap-1">
                      <h3 className="text-lg font-bold text-slate-900 leading-tight">
                        {student.nombre} {student.apellido && student.apellido !== 'Sin Apellido' && !student.nombre.includes(student.apellido) ? student.apellido : ''} {student.situacion === 'DUPLICADO' && <span className="text-rose-600 text-sm ml-2">(DUPLICADO)</span>}
                      </h3>
                      <div className="flex flex-wrap items-center gap-3 text-slate-600 text-xs font-semibold">
                        <span className="flex items-center gap-1 text-slate-500">ID: {student.dni}</span>
                        <span className="hidden sm:inline text-slate-300">⬢</span>
                        <span className="flex items-center gap-1 text-slate-500">MATRÍCULA: {student.dni}</span>
                      </div>
                      <div className="flex flex-wrap items-center gap-3 text-slate-600 text-xs font-semibold">
                        <span className="bg-slate-100 text-slate-700 px-2 py-0.5 rounded uppercase tracking-wider border border-slate-200">
                          LICENCIA {formatLicencia(student.licencia)}
                        </span>
                        <span className="text-[#002d2b] bg-[#0ffff4]/15 px-2 py-0.5 rounded uppercase tracking-wider border border-[#0ffff4]/40">
                          COMISIÓN {student.comision || '-'}
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="flex items-center justify-between w-full md:w-auto gap-4 border-t md:border-0 pt-3 md:pt-0 border-slate-100">
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider border ${student.estado_analitico === 'emitido' ? 'bg-[#00968f] text-white border-[#00968f]' : 'bg-slate-100 text-slate-500 border-slate-200'}`}>
                      {student.estado_analitico === 'emitido' ? 'Emitido' : 'Borrador'}
                    </span>
                    <span className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider border ${isAnaliticoCompleto(student) ? 'bg-emerald-100 text-emerald-700 border-emerald-200' : 'bg-rose-100 text-rose-700 border-rose-200'}`}>
                      {isAnaliticoCompleto(student) ? 'Completo' : 'Incompleto'}
                    </span>
                    {(!student.notas || student.notas.length === 0) && (
                      <span className="text-amber-700 font-bold text-xs bg-amber-50 border border-amber-100 px-3 py-1.5 rounded-lg shadow-sm whitespace-nowrap">Esperando Notas</span>
                    )}

                    {/* Quick actions row */}
                    <div className="flex items-center gap-1 border-l pl-3 border-slate-100" onClick={(e) => e.stopPropagation()}>
                      {user.role === 'admin' && (
                        <button
                          onClick={() => handleDelete(student.id as any)}
                          className="p-2 text-red-400 hover:bg-red-50 hover:text-red-700 rounded-lg transition-colors"
                          title="Eliminar Alumno"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                      {student.notas && student.notas.length > 0 && (
                        <button
                          onClick={() => downloadPDF(student)}
                          className="p-2 text-[#00968f] hover:bg-[#0ffff4]/15 hover:text-[#002d2b] rounded-lg transition-colors"
                          title="Descargar Analítico"
                        >
                          <Download className="w-5 h-5" />
                        </button>
                      )}
                      {student.notas && student.notas.length > 0 && (
                        <button
                          onClick={() => setDiplomaModal({ isOpen: true, student })}
                          className="p-2 text-indigo-600 hover:bg-indigo-50 hover:text-indigo-800 rounded-lg transition-colors"
                          title="Generar Diploma"
                        >
                          <School className="w-5 h-5" />
                        </button>
                      )}
                    </div>
                  </div>
                </motion.div>
              ))
            ) : (
              <div className="w-full py-24 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 border-dashed">
                <FileSpreadsheet className="w-16 h-16 text-slate-200 mb-4" />
                <p className="text-slate-500 font-medium">No se encontraron alumnos para mostrar.</p>
              </div>
            )}
          </AnimatePresence>
          {/* Paginator Controls */}
          {filteredStudents.length > 0 && (
            <div className="flex flex-col sm:flex-row items-center justify-between gap-4 py-6 mt-4 border-t border-slate-200 text-sm font-medium text-slate-600 bg-white/50 backdrop-blur-sm rounded-2xl px-6">
              <div className="flex items-center gap-2">
                <span>Mostrar:</span>
                <select
                  value={itemsPerPage}
                  onChange={(e) => { setItemsPerPage(Number(e.target.value)); setCurrentPage(1); }}
                  className="bg-white border border-slate-200 rounded-lg px-2 py-1.5 focus:ring-2 focus:ring-[#00968f] outline-none cursor-pointer"
                >
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={500}>500</option>
                </select>
                <span>por página</span>
              </div>

              <div className="flex items-center gap-1.5 bg-white border border-slate-200 p-1 rounded-xl shadow-sm">
                <button
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                  className="px-4 py-2 rounded-lg hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-bold text-slate-700"
                >
                  Anterior
                </button>
                <div className="px-4 py-2 bg-[#0ffff4]/15 text-[#002d2b] rounded-lg font-black text-sm border border-[#0ffff4]/40">
                  {currentPage} de {Math.max(1, Math.ceil(filteredStudents.length / itemsPerPage))}
                </div>
                <button
                  onClick={() => setCurrentPage(p => Math.min(Math.ceil(filteredStudents.length / itemsPerPage), p + 1))}
                  disabled={currentPage >= Math.ceil(filteredStudents.length / itemsPerPage)}
                  className="px-4 py-2 rounded-lg hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors font-bold text-slate-700"
                >
                  Siguiente
                </button>
              </div>
              <div className="text-slate-500 whitespace-nowrap font-bold bg-white px-4 py-2 rounded-xl shadow-sm border border-slate-100">
                Total: {filteredStudents.length} alumnos
              </div>
            </div>
          )}
      </div>
    </div>
  </div>
);

  const renderAnaliticoModals = () => (
    <>

      {/* MODAL PARA DIPLOMA */}
      {diplomaModal.isOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[70] flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl w-full max-w-md overflow-hidden animate-in fade-in zoom-in duration-200">
            <div className="p-6 border-b border-slate-100 bg-indigo-50/50">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-indigo-100 text-indigo-700 rounded-lg">
                    <School className="w-5 h-5" />
                  </div>
                  <h2 className="text-xl font-bold text-slate-800">Generar Diploma</h2>
                </div>
                <button onClick={() => setDiplomaModal({ isOpen: false, student: null })} className="text-slate-400 hover:text-slate-600 transition-colors bg-white p-1 rounded-full border border-slate-100 shadow-sm">
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>
            
            <form onSubmit={handleGenerateDiploma} className="p-6 space-y-5">
              <p className="text-sm text-slate-500">
                Confirmá los datos para el diploma de <span className="font-bold text-slate-800">{diplomaModal.student?.nombre} {diplomaModal.student?.apellido}</span>.
              </p>
              
              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">Nacionalidad</label>
                  <input 
                    required 
                    name="nacionalidad" 
                    defaultValue={diplomaModal.student?.nacionalidad || 'ARGENTINA'} 
                    className="w-full border border-slate-200 rounded-xl p-3 focus:ring-2 ring-indigo-100 outline-none font-medium" 
                    placeholder="Ej: ARGENTINA"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5">Fecha de Emisión</label>
                  <input 
                    required 
                    type="date" 
                    name="fecha_emision" 
                    defaultValue={new Date().toISOString().split('T')[0]} 
                    className="w-full border border-slate-200 rounded-xl p-3 focus:ring-2 ring-indigo-100 outline-none font-medium" 
                  />
                </div>
              </div>

              <div className="pt-4 flex gap-3">
                <button type="button" onClick={() => setDiplomaModal({ isOpen: false, student: null })} className="flex-1 py-3 bg-white border border-slate-200 rounded-xl font-bold text-slate-600 hover:bg-slate-50 transition-colors">Cancelar</button>
                <button type="submit" disabled={isUploading} className="flex-1 py-3 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold transition-all shadow-lg shadow-indigo-200 flex items-center justify-center gap-2">
                  {isUploading ? 'Generando...' : <><Download className="w-4 h-4" />Descargar</>}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
      <AnimatePresence>
        {selectedStudent && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center p-4 z-50 shadow-2xl"
            onClick={closeStudentModal}
          >
            <motion.div
              initial={{ scale: 0.95, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh] shadow-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              {/* Modal Header */}
              <div className="p-6 md:px-8 border-b border-slate-100 flex justify-between items-start bg-slate-50/50">
                <div className="flex-1">
                  {!editingDatos ? (
                    <>
                      <h2 className="text-2xl font-black text-slate-900">{selectedStudent.nombre}</h2>
                      <div className="flex flex-wrap items-center gap-3 mt-2 text-slate-500 text-sm font-medium">
                        <span className="bg-white border border-slate-200 px-3 py-1 rounded shadow-sm">ID: {selectedStudent.dni}</span>
                        {selectedStudent.email && (
                          <span className="bg-white border border-slate-200 px-3 py-1 rounded shadow-sm">Email: {selectedStudent.email}</span>
                        )}
                        {(user?.role === 'admin' || user?.role === 'editor') && (
                          <button
                            onClick={startEditDatos}
                            className="text-[#00968f] hover:text-[#002d2b] text-xs font-bold border border-[#0ffff4]/40 bg-[#0ffff4]/15 px-2 py-0.5 rounded transition-colors"
                          >
                            Editar Datos
                          </button>
                        )}
                      </div>
                    </>
                  ) : (
                    <div className="space-y-3">
                      <p className="text-xs font-bold text-slate-500 uppercase tracking-widest">Editando Datos Personales</p>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                        <div>
                          <label className="block text-xs text-slate-500 mb-1">ID</label>
                          <input
                            value={editDni}
                            onChange={e => setEditDni(e.target.value)}
                            className="w-full px-2 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                            placeholder="Nro. documento"
                          />
                        </div>
                        <div>
                          <label className="block text-xs text-slate-500 mb-1">Nombre</label>
                          <input
                            value={editNombre}
                            onChange={e => setEditNombre(e.target.value)}
                            className="w-full px-2 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                            placeholder="Nombre(s)"
                          />
                        </div>
                        <div>
                          <label className="block text-xs text-slate-500 mb-1">Apellido</label>
                          <input
                            value={editApellido}
                            onChange={e => setEditApellido(e.target.value)}
                            className="w-full px-2 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                            placeholder="Apellido"
                          />
                        </div>
                        <div>
                          <label className="block text-xs text-slate-500 mb-1">Nacionalidad</label>
                          <input
                            value={editNacionalidad}
                            onChange={e => setEditNacionalidad(e.target.value)}
                            className="w-full px-2 py-1.5 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                            placeholder="Ej: Argentina"
                          />
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <button onClick={saveDatos} className="bg-[#002d2b] hover:bg-[#00968f] text-white px-4 py-1.5 rounded-lg text-sm font-bold transition-colors">Guardar</button>
                        <button onClick={() => setEditingDatos(false)} className="bg-slate-200 hover:bg-slate-300 text-slate-700 px-4 py-1.5 rounded-lg text-sm font-bold transition-colors">Cancelar</button>
                      </div>
                    </div>
                  )}
                </div>
                <button
                  onClick={closeStudentModal}
                  className="p-2 bg-slate-200 hover:bg-slate-300 text-slate-600 rounded-full transition-colors shrink-0 ml-4"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Modal Body */}
              <div className="p-6 md:p-8 overflow-y-auto custom-scrollbar">
                {selectedStudent.estado_analitico === 'emitido' && !isAnaliticoCompleto(selectedStudent) && (
                  <div className="mb-6 p-4 bg-amber-50 border border-amber-200 rounded-2xl flex items-center gap-3 text-amber-800 animate-pulse shadow-sm">
                    <AlertTriangle className="w-6 h-6 shrink-0" />
                    <div>
                      <p className="font-bold text-sm text-amber-900 leading-tight">⚠️ Analítico Generado: Faltan Notas</p>
                      <p className="text-xs text-amber-700 mt-0.5">Marcado como emitido pero con materias obligatorias pendientes.</p>
                    </div>
                  </div>
                )}
                <div className="grid grid-cols-2 gap-4 mb-8">
                  <div className="bg-[#0ffff4]/12 p-4 border border-[#0ffff4]/35 rounded-2xl">
                    <p className="text-[11px] font-black text-[#00968f] uppercase tracking-widest">Licencia / Carrera</p>
                    <p className="font-bold text-slate-800 mt-1 text-base">{selectedStudent.licencia || 'NO ASIGNADA'}</p>
                  </div>
                  <div className="bg-emerald-50 p-4 border border-emerald-100 rounded-2xl">
                    <p className="text-[11px] font-black text-emerald-600 uppercase tracking-widest">Nacionalidad</p>
                    <p className="font-bold text-slate-800 mt-1 text-base">{selectedStudent.nacionalidad || 'SIN ESPECIFICAR'}</p>
                  </div>
                  <div className="bg-emerald-50 p-4 border border-emerald-100 rounded-2xl">
                    <p className="text-[11px] font-black text-emerald-600 uppercase tracking-widest">Comisión</p>
                    <p className="font-bold text-slate-800 mt-1 text-base">{selectedStudent.comision || 'SIN COMISIÓN'}</p>
                  </div>
                  {selectedStudent.situacion && (
                    <div className="col-span-2 bg-indigo-50 p-4 border border-indigo-100 rounded-2xl flex items-center justify-between">
                      <div>
                        <p className="text-[11px] font-black text-indigo-500 uppercase tracking-widest">Situación Académica</p>
                        <p className="font-bold text-slate-800 mt-1 text-base">{selectedStudent.situacion}</p>
                      </div>
                      <div className="bg-indigo-100 w-10 h-10 rounded-full flex items-center justify-center">
                        <FileText className="w-5 h-5 text-indigo-600" />
                      </div>
                    </div>
                  )}
                </div>

                <div className="flex items-center justify-between mb-4">
                  <h3 className="text-lg font-bold text-slate-900">Materias Aprobadas ({selectedStudent.notas?.length || 0})</h3>
                  {(user?.role === 'admin' || user?.role === 'editor') && (
                    <button
                      onClick={() => setShowAddNota(prev => !prev)}
                      className="text-xs font-bold bg-[#002d2b] hover:bg-[#00968f] text-white px-3 py-1.5 rounded-lg transition-colors"
                    >
                      + Agregar Nota
                    </button>
                  )}
                </div>

                {showAddNota && (
                  <div className="mb-4 p-4 bg-[#0ffff4]/12 border border-[#0ffff4]/35 rounded-xl space-y-3">
                    <p className="text-xs font-bold text-[#00968f] uppercase tracking-widest">Nueva Materia / Nota</p>
                    <div className="flex gap-2 items-end">
                      <div className="flex-1">
                        <label className="block text-xs text-slate-500 mb-1">Materia</label>
                        <input
                          value={newMateria}
                          onChange={e => setNewMateria(e.target.value)}
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                          placeholder="Ej: TÉCNICA TÁCTICA I"
                        />
                      </div>
                      <div className="w-24">
                        <label className="block text-xs text-slate-500 mb-1">Nota (0-10)</label>
                        <input
                          type="number"
                          step="0.5"
                          min="0"
                          max="10"
                          value={newNota}
                          onChange={e => setNewNota(e.target.value)}
                          className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                          placeholder="8.5"
                        />
                      </div>
                      <button onClick={saveNotaManual} className="bg-[#002d2b] hover:bg-[#00968f] text-white px-4 py-2 rounded-lg text-sm font-bold transition-colors shrink-0">
                        Guardar
                      </button>
                    </div>
                  </div>
                )}

                {(() => {
                  const stripAccents = (s: string) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[,.:;]/g, '').replace(/\s+/g, ' ').toUpperCase().trim();
                  const planMaterias = getSubjectsByLicencia(selectedStudent.licencia || '');
                  const notasMap = Object.fromEntries(
                    (selectedStudent.notas || []).map(n => [stripAccents(n.materia), n.nota])
                  );

                  return planMaterias.length > 0 ? (
                    <div className="space-y-1 border border-slate-100 rounded-xl overflow-hidden shadow-sm">
                      {planMaterias.map((materia, i) => {
                        const notaActual = notasMap[stripAccents(materia)] ?? 0;
                        const tiene = notaActual > 0;
                        const pending = pendingNotas[materia];
                        const hasPending = pending !== undefined && parseFloat(pending) !== notaActual;
                        const isSaving = savingNota === materia;

                        const guardarNota = async (mat: string, val: string) => {
                          const num = parseFloat(val);
                          if (isNaN(num) || num < 0 || num > 10) {
                            toast.error(`Nota inválida: "${val}". Debe ser un número entre 0 y 10.`);
                            return;
                          }
                          if (num === notaActual) return;
                          setSavingNota(mat);
                          try {
                            const url = `${API_URL}/api/students/${selectedStudent.id}/nota?${getUserQuery()}`;
                            console.log('[guardarNota] POST', url, { asignatura: mat, nota: num });
                            const res = await fetch(url, {
                              method: 'POST',
                              headers: { 'Content-Type': 'application/json' },
                              body: JSON.stringify({ asignatura: mat, nota: num })
                            });
                            const data = await res.json();
                            console.log('[guardarNota] respuesta:', data);
                            if (!res.ok) {
                              toast.error(`Error al guardar: ${data.error || res.status}`);
                              setSavingNota(null);
                              return;
                            }
                            await fetchStudents();
                            const updated = await fetch(`${API_URL}/api/students`).then(r => r.json());
                            const s = updated.data?.find((x: any) => x.id === selectedStudent.id);
                            if (s) setSelectedStudent(s);
                            setPendingNotas(prev => { const nx = { ...prev }; delete nx[mat]; return nx; });
                          } catch (err: any) {
                            toast.error(`Error de red: ${err.message}`);
                          }
                          setSavingNota(null);
                        };

                        return (
                          <div key={i} className={`flex justify-between items-center px-4 py-2.5 border-b border-slate-50 last:border-0 transition-colors ${tiene && !hasPending ? 'bg-white' : hasPending ? 'bg-yellow-50' : 'bg-amber-50'}`}>
                            <span className="font-medium text-slate-700 text-sm flex-1 pr-4">{materia}</span>
                            <div className="flex items-center gap-2 shrink-0">
                              {(user?.role === 'admin' || user?.role === 'editor') ? (
                                <>
                                  <input
                                    type="number"
                                    step="0.5"
                                    min="0"
                                    max="10"
                                    value={pending !== undefined ? pending : (notaActual || '')}
                                    onChange={e => setPendingNotas(prev => ({ ...prev, [materia]: e.target.value }))}
                                    onKeyDown={e => { if (e.key === 'Enter') guardarNota(materia, pending ?? String(notaActual)); }}
                                    className={`w-16 text-center px-2 py-1 rounded-lg text-sm font-black border outline-none focus:ring-2 focus:ring-[#00968f] transition-colors ${hasPending ? 'bg-yellow-100 border-yellow-400 text-yellow-900' : tiene ? 'bg-[#0ffff4]/15 text-[#002d2b] border-[#0ffff4]/40' : 'bg-red-50 text-red-500 border-red-200'}`}
                                    placeholder="0"
                                  />
                                  {hasPending && (
                                    <button
                                      onClick={() => guardarNota(materia, pending!)}
                                      disabled={isSaving}
                                      className="text-xs bg-[#002d2b] hover:bg-[#00968f] text-white px-2 py-1 rounded-lg font-bold transition-colors disabled:opacity-50"
                                    >
                                      {isSaving ? '...' : 'OK'}
                                    </button>
                                  )}
                                </>
                              ) : (
                                <span className={`w-10 text-center py-1 rounded font-black text-sm ${tiene ? 'bg-[#0ffff4]/20 text-[#002d2b]' : 'bg-red-50 text-red-400'}`}>{notaActual || '-'}</span>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  ) : (
                    // Fallback: si la licencia no tiene plan definido, mostrar las notas cargadas
                    selectedStudent.notas && selectedStudent.notas.length > 0 ? (
                      <div className="space-y-2 border border-slate-100 rounded-xl overflow-hidden shadow-sm">
                        {selectedStudent.notas.map((nota, i) => (
                          <div key={i} className="flex justify-between items-center p-3.5 bg-white border-b border-slate-50 hover:bg-slate-50 transition-colors last:border-0">
                            <span className="font-semibold text-slate-700 text-sm">{nota.materia}</span>
                            <span className="bg-[#0ffff4]/20 text-[#002d2b] w-10 text-center py-1 rounded font-black">{nota.nota}</span>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="bg-amber-50 border border-amber-100 p-6 rounded-2xl text-center">
                        <p className="text-amber-700 font-semibold mb-1">Licencia no reconocida o sin plan definido.</p>
                        <p className="text-amber-600/70 text-sm">Verificá que el alumno tenga una licencia CB, A o PRO asignada.</p>
                      </div>
                    )
                  );
                })()}

                {/* Fechas Importantes */}
                <div className="mt-8 border-t border-slate-100 pt-6">
                  <h3 className="text-lg font-bold text-slate-900 mb-4">Fechas Importantes</h3>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">Fecha Fin de Cursada</label>
                      <input
                        type="date"
                        value={selectedStudent.fecha_fin_cursada || ''}
                        onChange={(e) => setSelectedStudent({ ...selectedStudent, fecha_fin_cursada: e.target.value })}
                        className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-[#00968f] outline-none transition-all"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">Fecha de Emisión</label>
                      <input
                        type="date"
                        value={selectedStudent.fecha_emision || ''}
                        onChange={(e) => setSelectedStudent({ ...selectedStudent, fecha_emision: e.target.value })}
                        className="w-full px-3 py-2 border border-slate-200 rounded-lg focus:ring-2 focus:ring-[#00968f] outline-none transition-all"
                      />
                    </div>
                  </div>
                  <div className="mt-4 flex justify-end">
                    <button onClick={handleSaveFechas} className="bg-slate-800 hover:bg-slate-900 text-white px-4 py-2 rounded-lg text-sm font-bold transition-colors">
                      Guardar Fechas
                    </button>
                  </div>
                </div>

                {/* Historial de Actividad */}
                {selectedStudent.historial && selectedStudent.historial.length > 0 && (
                  <div className="mt-8 border-t border-slate-100 pt-6">
                    <h3 className="text-lg font-bold text-slate-900 mb-4">Historial de Actividad</h3>
                    <div className="space-y-3">
                      {selectedStudent.historial.slice().reverse().map((log: any, i: number) => (
                        <div key={i} className="bg-slate-50 rounded-lg p-3 border border-slate-100 text-sm">
                          <div className="flex justify-between items-center mb-1">
                            <div className="flex flex-col">
                              <span className="font-bold text-slate-700">{formatUserDisplay(log.usuario, log.nombre)}</span>
                              {log.usuario && <span className="text-[11px] text-slate-400">{log.usuario}</span>}
                            </div>
                            <span className="text-xs text-slate-400 font-medium">{new Date(log.fecha).toLocaleString()}</span>
                          </div>
                          <p className="text-slate-600">{log.accion}</p>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>

                                          {/* Modal Footer */}
              <div className="p-6 border-t border-slate-100 bg-slate-50 flex justify-between items-center shrink-0 flex-wrap gap-4">
                <div className="flex flex-wrap items-center gap-3 w-full justify-between">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className={`px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider ${selectedStudent.estado_analitico === 'emitido' ? 'bg-green-100 text-green-700' : 'bg-slate-200 text-slate-700'}`}>
                      Analítico {selectedStudent.estado_analitico === 'emitido' ? 'Emitido' : 'Borrador'}
                    </span>
                    <span className={`px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider ${selectedStudent.diploma_emitido ? 'bg-green-100 text-green-700' : 'bg-slate-200 text-slate-700'}`}>
                      Diploma {selectedStudent.diploma_emitido ? 'Emitido' : 'Pendiente'}
                    </span>
                    {!isAnaliticoCompleto(selectedStudent) && (
                      <span className="px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider bg-amber-100 text-amber-700">
                        Incompleto (faltan notas)
                      </span>
                    )}
                    {user.role === 'viewer' && (
                      <span className="px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider bg-slate-200 text-slate-700">
                        Solo lectura
                      </span>
                    )}
                  </div>

                  <div className="flex items-center gap-2 w-full justify-end flex-wrap mt-2">
                    {user.role === 'admin' && selectedStudent.estado_analitico === 'emitido' && (
                      <button
                        onClick={() => handleToggleEstado(selectedStudent.id as any)}
                        className="px-3.5 py-2 rounded-lg text-sm font-bold flex gap-2 items-center transition-all shadow-sm border bg-amber-50 hover:bg-amber-100 text-amber-800 border-amber-300"
                        title="Revertir a Borrador (requiere justificación)"
                      >
                        Desmarcar Emitido
                      </button>
                    )}

                    {user.role === 'admin' && selectedStudent.notas && selectedStudent.notas.length > 0 && (
                      <button
                        onClick={() => handleDeleteNotas(selectedStudent.id as any)}
                        className="p-2 text-red-500 hover:bg-red-100 rounded-lg transition-colors border border-transparent hover:border-red-200"
                        title="Eliminar solo las notas (Quedar en Borrador)"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    )}

                    {(() => {
                      const disabled = !selectedStudent || !isAnaliticoCompleto(selectedStudent) || user.role === 'viewer';
                      const disabledClasses = disabled ? ' opacity-50 cursor-not-allowed' : '';
                      return (
                        <>
                          <button
                            onClick={() => !disabled && downloadPDF(selectedStudent)}
                            disabled={disabled}
                            className={`px-4 py-2 rounded-lg text-sm font-bold flex gap-2 items-center transition-all border border-[#0ffff4]/40 bg-[#0ffff4]/10 hover:bg-[#0ffff4]/20 text-[#002d2b]${disabledClasses}`}
                            title="Abre el PDF para verlo, sin marcarlo como Emitido"
                          >
                            <Download className="w-4 h-4" /> Vista Previa
                          </button>
                          <button
                            onClick={() => !disabled && setDiplomaModal({ isOpen: true, student: selectedStudent })}
                            disabled={disabled}
                            className={`flex items-center justify-center gap-2 px-4 py-2 border-2 border-indigo-600 text-indigo-700 hover:bg-indigo-50 rounded-lg text-sm font-bold transition-all${disabledClasses}`}
                          >
                            <School className="w-5 h-5" />
                            Generar Diploma
                          </button>
                          <button
                            onClick={() => !disabled && downloadPDFAndEmit(selectedStudent)}
                            disabled={disabled}
                            className={`flex items-center justify-center gap-2 px-4 py-2 bg-[#002d2b] hover:bg-[#00968f] text-white rounded-lg text-sm font-bold transition-all shadow-md shadow-[#002d2b]/20${disabledClasses}`}
                            title="Genera el PDF y marca el analítico como Emitido"
                          >
                            <Download className="w-5 h-5" />
                            Generar Analítico PDF
                          </button>
                        </>
                      );
                    })()}
                  </div>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </>
  );



  const renderFieldHourDashboard = () => {
    if (!fieldHourStats) return (
      <div className="flex flex-col items-center justify-center p-12 text-slate-400">
        <RefreshCw className="w-8 h-8 animate-spin mb-4" />
        <p className="font-bold">Cargando estadísticas...</p>
      </div>
    );

    const stats = [
      { 
        title: 'Total Alumnos', 
        value: fieldHourStats.totalStudents, 
        icon: Users, 
        color: 'bg-indigo-50 text-indigo-600',
        desc: 'Alumnos en el padrón'
      },
      { 
        title: 'Pendientes de Revisión', 
        value: fieldHourStats.pendingReviews, 
        icon: Clock, 
        color: 'bg-amber-50 text-amber-600',
        desc: 'Documentos por aprobar'
      },
      { 
        title: 'Trámites Completados', 
        value: fieldHourStats.completedStudents, 
        icon: CheckCircle2, 
        color: 'bg-emerald-50 text-emerald-600',
        desc: '4/4 documentos aprobados'
      }
    ];

    return (
      <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-10">
        <header>
          <h2 className="text-3xl font-extrabold text-slate-900 tracking-tight">Dashboard Inteligente</h2>
          <p className="text-slate-500 font-medium">Resumen general de gestión de Horas de Campo.</p>
        </header>

        {/* Stats Grid */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {stats.map((s, i) => (
            <div key={i} className="bg-white p-6 rounded-[28px] border border-slate-200 shadow-sm hover:shadow-md transition-all flex items-center gap-5">
              <div className={`w-14 h-14 rounded-2xl flex items-center justify-center ${s.color}`}>
                <s.icon className="w-7 h-7" />
              </div>
              <div>
                <p className="text-[11px] font-black uppercase text-slate-400 tracking-widest">{s.title}</p>
                <h3 className="text-2xl font-black text-slate-800">{s.value}</h3>
                <p className="text-[10px] text-slate-400 font-medium mt-0.5">{s.desc}</p>
              </div>
            </div>
          ))}
        </div>

        {/* Recent Activity */}
        <div className="space-y-6">
          <h3 className="text-xl font-bold text-slate-800">Actividad Reciente</h3>
          <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-hidden">
            <div className="divide-y divide-slate-50">
              {fieldHourStats.recentActivity.map((e, i) => (
                <div key={i} className="p-5 flex items-center justify-between hover:bg-slate-50 transition-colors">
                  <div className="flex items-center gap-4">
                    <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center text-slate-400">
                      <FileText className="w-5 h-5" />
                    </div>
                    <div>
                      <p className="text-sm font-bold text-slate-800">{e.documentType}</p>
                      <p className="text-[11px] text-slate-500 font-medium">
                        Entregado por <span className="text-slate-700 font-bold">{e.student ? `${e.student.nombre} ${e.student.apellido}` : 'Alumno desconocido'}</span>
                      </p>
                    </div>
                  </div>
                  <div className="text-right">
                    <span className={`px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-wider ${
                      e.status === 'Aprobado' ? 'bg-emerald-50 text-emerald-600' : 
                      e.status === 'Reentregar' ? 'bg-rose-50 text-rose-600' : 
                      'bg-amber-50 text-amber-600'
                    }`}>
                    </span>
                    <p className="text-[9px] text-slate-400 mt-1">{new Date(e.uploadDate).toLocaleString()}</p>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderHorasCampo = () => {
    const isFieldHourAdmin = user?.role === 'admin' || user?.permissions?.['horas-campo'] === 'admin';
    const getSuffix = (lic: string) => {
      const l = lic.toUpperCase();
      if (l.includes('PRO')) return 'PRO';
      if (l.includes('CB')) return 'CB';
      if (l.includes('A')) return 'A';
      if (l.includes('TD1')) return 'TD1';
      if (l.includes('TD2')) return 'TD2';
      return '';
    };

    // Si el admin está en el subtab dashboard, lo mostramos
    if (isFieldHourAdmin && fieldHourSubTab === 'dashboard') {
      return renderFieldHourDashboard();
    }

    const userBaseLic = user?.licencia ? getSuffix(user.licencia) : '';
      
      const licenses = [
        { id: 'CB', name: 'Licencia CB', color: 'slate' },
        { id: 'A', name: 'Licencia A', color: 'emerald' },
        { id: 'PRO', name: 'Licencia PRO', color: 'amber' },
        { id: 'TD1', name: 'TD1', color: 'indigo' },
        { id: 'TD2', name: 'TD2', color: 'rose' }
      ];

      const docTypes = (suffix: string) => [
        `PLANILLA DE ASISTENCIA A CLUBES ${suffix}`,
        `CARTA DE HORAS PRESENCIALES EN CAMPO/PEDAGÓGICAS ${suffix}`,
        `CARTA DE HORAS PROFESIONALIZANTES ${suffix}`,
        "TRABAJO PRÁCTICO"
      ];

      const templatesToDownload = [
        { name: 'PLANILLA DE ASISTENCIA A CLUBES', suffix: userBaseLic },
        { name: 'CARTA DE HORAS PRESENCIALES EN CAMPO/PEDAGÓGICAS', suffix: userBaseLic },
        { name: 'CARTA DE HORAS PROFESIONALIZANTES', suffix: userBaseLic }
      ];

      if (user?.role === 'student' || !isFieldHourAdmin) {
        return (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-8 pb-12">
          <header className="flex flex-col gap-6">
            <div className="flex items-center justify-between">
              <div className="flex flex-col gap-1">
                <h2 className="text-3xl font-extrabold text-slate-900 tracking-tight">Horas de Campo</h2>
                <div className="flex items-center gap-2">
                  <span className="px-3 py-1 bg-[#0ffff4]/20 text-[#002d2b] rounded-full text-[10px] font-black uppercase tracking-widest border border-[#0ffff4]/40">
                    ID Alumno: {user.documento || user.id}
                  </span>
                  {user.licencia && (
                    <span className="px-3 py-1 bg-emerald-50 text-emerald-700 rounded-full text-[10px] font-black uppercase tracking-widest border border-emerald-100">
                      {user.licencia}
                    </span>
                  )}
                </div>
              </div>
            </div>
          </header>

          {fieldHourSubTab === 'summary' && (
            <div className="animate-in fade-in slide-in-from-bottom-4 duration-700 space-y-10">
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                {docTypes(userBaseLic).map((type, i) => {
                  const entrega = entregas.find(e => e.documentType === type);
                  const isAproved = entrega?.status === 'Aprobado';
                  return (
                    <div key={i} className={`p-5 rounded-[24px] border ${isAproved ? 'bg-emerald-50 border-emerald-100' : 'bg-white border-slate-100'} shadow-sm flex flex-col items-center text-center gap-3`}>
                      <div className={`w-10 h-10 rounded-xl flex items-center justify-center ${isAproved ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-50 text-slate-300'}`}>
                        {isAproved ? <CheckCircle2 className="w-5 h-5" /> : <Clock className="w-5 h-5" />}
                      </div>
                      <p className="text-[10px] font-bold text-slate-800 uppercase leading-tight">{type}</p>
                      <span className={`text-[9px] font-black uppercase tracking-wider ${isAproved ? 'text-emerald-600' : 'text-slate-400'}`}>
                        {isAproved ? 'Completado' : 'Pendiente'}
                      </span>
                    </div>
                  );
                })}
              </div>

              <div className="bg-[#002d2b] rounded-[32px] p-8 text-white relative overflow-hidden shadow-xl border border-white/5">
                <div className="relative z-10 flex flex-col md:flex-row items-center justify-between gap-8">
                  <div className="space-y-4 text-center md:text-left">
                    <h3 className="text-2xl font-black">Tu Progreso General</h3>
                    <p className="text-white/60 text-sm max-w-sm">Para completar tus horas de campo debes entregar las 4 planillas obligatorias firmadas y selladas.</p>
                    <div className="flex items-center gap-4 justify-center md:justify-start">
                      <div className="px-5 py-2 bg-white/10 rounded-2xl border border-white/10 text-xl font-black">
                        {entregas.filter(e => e.status === 'Aprobado').length} / 4
                      </div>
                      <span className="text-xs font-bold text-[#0ffff4] uppercase tracking-widest">Documentos Aprobados</span>
                    </div>
                  </div>
                  <div className="w-40 h-40 relative flex items-center justify-center">
                    <svg className="w-full h-full transform -rotate-90">
                      <circle cx="80" cy="80" r="70" fill="transparent" stroke="currentColor" strokeWidth="12" className="text-white/10" />
                      <circle cx="80" cy="80" r="70" fill="transparent" stroke="currentColor" strokeWidth="12" strokeDasharray="440" strokeDashoffset={440 - (440 * (entregas.filter(e => e.status === 'Aprobado').length / 4))} className="text-[#0ffff4] transition-all duration-1000" strokeLinecap="round" />
                    </svg>
                    <div className="absolute inset-0 flex flex-col items-center justify-center leading-none">
                      <span className="text-3xl font-black">{Math.round((entregas.filter(e => e.status === 'Aprobado').length / 4) * 100)}%</span>
                      <span className="text-[9px] font-black uppercase tracking-widest text-[#0ffff4]">Total</span>
                    </div>
                  </div>
                </div>
                {/* Decoration */}
                <div className="absolute top-0 right-0 -mr-20 -mt-20 w-64 h-64 bg-[#0ffff4]/5 rounded-full blur-3xl"></div>
              </div>
            </div>
          )}

          {fieldHourSubTab === 'general' && (
            <div className="animate-in fade-in zoom-in-95 duration-500 py-12 flex flex-col items-center justify-center text-center bg-white rounded-[32px] border border-slate-100 shadow-sm">
              <div className="w-20 h-20 bg-indigo-50 text-indigo-500 rounded-3xl flex items-center justify-center mb-6">
                <Info className="w-10 h-10" />
              </div>
              <h3 className="text-2xl font-black text-slate-800 mb-2">Información General</h3>
              <p className="text-slate-400 font-medium max-w-sm">Aquí encontrarás guías detalladas, noticias y avisos importantes sobre las horas de campo.</p>
              <div className="mt-8 px-6 py-2 bg-indigo-50 text-indigo-600 rounded-full text-[10px] font-black uppercase tracking-[0.2em] animate-pulse">
                Próximamente más información
              </div>
            </div>
          )}

          {fieldHourSubTab === 'descarga' && (
            <div className="space-y-6 animate-in fade-in slide-in-from-top-4 duration-500">
              <div className="flex items-center gap-3">
                <h3 className="text-xl font-bold text-slate-800">Descarga de Planillas</h3>
                {userBaseLic && <span className="bg-emerald-100 text-emerald-700 text-[10px] font-black px-2 py-0.5 rounded-full">LICENCIA {userBaseLic}</span>}
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {templatesToDownload.map((tpl, i) => (
                  <div key={i} className="bg-white p-6 rounded-[24px] border border-slate-200 shadow-sm hover:shadow-md transition-all flex flex-col gap-4">
                    <div className="flex items-start gap-4">
                      <div className="p-3 bg-slate-50 rounded-2xl text-slate-400">
                        <FileDown className="w-6 h-6" />
                      </div>
                      <span className="font-bold text-slate-700 text-sm leading-tight">{tpl.name}</span>
                    </div>
                    <a 
                      href={`${API_URL}/api/horas-campo/download-template/${encodeURIComponent(`${tpl.name} ${tpl.suffix}`)}.pdf`}
                      target="_blank"
                      rel="noreferrer"
                      className="w-full py-3 bg-[#0ffff4]/10 hover:bg-[#0ffff4]/20 text-[#002d2b] rounded-xl text-sm font-bold text-center transition-all border border-[#0ffff4]/30"
                    >
                      Descargar PDF
                    </a>
                  </div>
                ))}
              </div>
            </div>
          )}

          {fieldHourSubTab === 'tp' && (
            <div className="animate-in fade-in slide-in-from-top-4 duration-500 space-y-6">
              <h3 className="text-xl font-bold text-slate-800">Trabajo Práctico</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="bg-white rounded-[24px] border border-slate-200 shadow-sm p-6 hover:shadow-md transition-all flex flex-col gap-6 relative group">
                  <div className="flex items-start justify-between">
                    <div className="flex flex-col gap-1">
                      <h3 className="font-bold text-slate-800 leading-snug pr-8">PLANTILLA DE TRABAJO PRÁCTICO</h3>
                      <span className="text-[10px] text-slate-400 font-medium italic">Descarga la plantilla para realizar el trabajo.</span>
                    </div>
                    <div className="p-3 bg-[#0ffff4]/10 rounded-2xl">
                      <FileDown className="w-6 h-6 text-[#00968f]" />
                    </div>
                  </div>
                  <a 
                    href={`${API_URL}/api/horas-campo/download-template/${encodeURIComponent('TRABAJO PRÁCTICO')}.pdf`}
                    target="_blank"
                    rel="noreferrer"
                    className="w-full py-3 bg-[#0ffff4]/10 hover:bg-[#0ffff4]/20 text-[#002d2b] rounded-xl text-sm font-bold text-center transition-all border border-[#0ffff4]/30"
                  >
                    Descargar Plantilla TP
                  </a>
                </div>
              </div>
            </div>
          )}

          {fieldHourSubTab === 'entrega' && (
            <div className="space-y-10 animate-in fade-in slide-in-from-top-4 duration-500">
              {/* License Cards Grid - 5 columns */}
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4">
                {licenses.map(lic => {
                  const isEnabled = lic.id === userBaseLic;
                  return (
                    <div 
                      key={lic.id} 
                      className={`relative p-5 rounded-[24px] border-2 transition-all overflow-hidden ${
                        isEnabled 
                          ? 'border-[#00968f] bg-[#0ffff4]/5 shadow-lg ring-1 ring-[#00968f]/20' 
                          : 'border-slate-100 bg-slate-50 opacity-60'
                      }`}
                    >
                      {!isEnabled && (
                        <div className="absolute inset-0 bg-slate-200/20 backdrop-blur-[1px] flex items-center justify-center p-4">
                          <span className="bg-slate-800/80 text-white text-[8px] font-bold px-2 py-1 rounded-full uppercase tracking-widest">Inhabilitado</span>
                        </div>
                      )}
                      <div className="flex flex-col gap-3 relative z-10">
                        <div className={`w-10 h-10 rounded-2xl flex items-center justify-center ${isEnabled ? 'bg-[#0ffff4] text-[#002d2b]' : 'bg-slate-200 text-slate-400'}`}>
                          <School className="w-5 h-5" />
                        </div>
                        <div>
                          <h3 className={`font-black text-sm ${isEnabled ? 'text-[#002d2b]' : 'text-slate-400'}`}>{lic.name}</h3>
                          <p className={`text-[9px] font-bold uppercase tracking-widest ${isEnabled ? 'text-[#00968f]' : 'text-slate-400'}`}>
                            {isEnabled ? 'Tu Licencia' : 'No corresponde'}
                          </p>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Documents Area */}
              <div className="animate-in fade-in slide-in-from-top-4 duration-700">
                <div className="flex items-center gap-3 mb-6">
                  <div className="h-px bg-slate-200 flex-1"></div>
                  <h3 className="text-sm font-black text-slate-400 uppercase tracking-widest">Documentos Obligatorios</h3>
                  <div className="h-px bg-slate-200 flex-1"></div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {docTypes(userBaseLic).map((type) => {
                    const entrega = entregas.find(e => e.documentType === type);
                    const isUploading = uploadingDoc === type;

                    return (
                      <div key={type} className="bg-white rounded-[24px] border border-slate-200 shadow-sm p-6 hover:shadow-md transition-all flex flex-col justify-between gap-6 overflow-hidden relative group">
                        <div className="flex items-start justify-between">
                          <div className="flex flex-col gap-1">
                            <h3 className="font-bold text-slate-800 leading-snug pr-8">{type}</h3>
                            {entrega ? (
                              <div className="flex items-center gap-2 mt-1">
                                <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider ${
                                  entrega.status === 'Aprobado' ? 'bg-emerald-100 text-emerald-700 border border-emerald-200' : 
                                  entrega.status === 'Reentregar' ? 'bg-rose-100 text-rose-700 border border-rose-200' : 
                                  'bg-amber-100 text-amber-700 border border-amber-200'
                                }`}>
                                  {entrega.status}
                                </span>
                                <span className="text-[10px] text-slate-400 font-medium">
                                  {new Date(entrega.uploadDate || entrega.createdAt).toLocaleDateString()}
                                </span>
                              </div>
                            ) : (
                              <span className="text-[10px] bg-slate-100 text-slate-500 px-3 py-1 rounded-full font-black uppercase tracking-wider border border-slate-200 mt-1">Pendiente de entrega</span>
                            )}
                          </div>
                          <div className="p-3 bg-slate-50 rounded-2xl group-hover:bg-[#0ffff4]/10 transition-colors">
                            <FileText className="w-6 h-6 text-slate-400 group-hover:text-[#00968f] transition-colors" />
                          </div>
                        </div>

                        {entrega?.observations && (
                          <div className="bg-rose-50 border border-rose-100 rounded-xl p-3 text-xs text-rose-800">
                            <strong>Observaciones:</strong> {entrega.observations}
                          </div>
                        )}

                        <div className="flex items-center gap-3 mt-2">
                          {entrega?.status === 'Aprobado' ? (
                            <div className="w-full py-3 bg-emerald-50 text-emerald-700 rounded-xl text-sm font-bold flex items-center justify-center gap-2 border border-emerald-100">
                              <CheckCircle2 className="w-5 h-5" />
                              Documento Aprobado y Bloqueado
                            </div>
                          ) : (
                            <label className={`w-full relative flex items-center justify-center gap-2 py-3 ${isUploading ? 'bg-slate-400' : 'bg-[#002d2b] hover:bg-[#00968f]'} text-white rounded-xl text-sm font-bold transition-all shadow-sm cursor-pointer`}>
                              {isUploading ? (
                                <RefreshCw className="w-4 h-4 animate-spin" />
                              ) : (
                                <Upload className="w-4 h-4" />
                              )}
                              <span>{entrega ? 'Reemplazar Entrega' : 'Subir Entrega'}</span>
                              <input 
                                type="file" 
                                className="hidden" 
                                accept=".pdf,image/*" 
                                onChange={(e) => {
                                  const file = e.target.files?.[0];
                                  if (file) handleUploadFieldHour(file, type);
                                }}
                                disabled={isUploading}
                              />
                            </label>
                          )}
                        </div>
                        
                        {entrega?.serverPath && (
                          <a 
                            href={`${API_URL}/api/horas-campo/download/${entrega.id}`}
                            target="_blank"
                            rel="noreferrer"
                            className="text-[10px] text-center text-slate-400 hover:text-[#00968f] underline font-medium"
                          >
                            Ver mi entrega actual
                          </a>
                        )}
                      </div>
                    );
                  })}
                  </div>
                </div>
              </div>
            )}
        </div>
      );
    } // Closes if (user.role === 'student' || admin restricted)

    // VISTA ADMIN (Solo para el rol admin o permiso específico)
    if (!isFieldHourAdmin) {
      return (
        <div className="flex flex-col items-center justify-center p-12 text-center">
          <div className="bg-amber-50 text-amber-700 p-6 rounded-2xl border border-amber-100 max-w-md">
            <AlertTriangle className="w-12 h-12 mx-auto mb-4 opacity-50" />
            <h3 className="text-lg font-bold mb-2">Acceso Restringido</h3>
            <p className="text-sm">Esta sección de administración está reservada para el departamento de Horas de Campo.</p>
          </div>
        </div>
      );
    }

    return (
      <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-8">
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h2 className="text-3xl font-extrabold text-slate-900 tracking-tight">Gestión de Horas de Campo</h2>
            <p className="text-slate-500 font-medium">Revisión y aprobación de documentos entregados por alumnos.</p>
          </div>
          <button 
            onClick={() => setNewStudentModal(true)}
            className="flex items-center gap-2 px-5 py-3 bg-[#002d2b] hover:bg-[#00968f] text-white rounded-xl text-sm font-bold transition-all shadow-lg hover:-translate-y-0.5 active:scale-95"
          >
            <UserPlus className="w-5 h-5" />
            + Nuevo Alumno
          </button>
        </header>

        <div className="bg-white rounded-[28px] border border-slate-200 shadow-sm overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-slate-50/50 border-b border-slate-100">
                  <th className="px-6 py-4 text-[11px] font-black uppercase tracking-widest text-slate-400">Alumno</th>
                  <th className="px-6 py-4 text-[11px] font-black uppercase tracking-widest text-slate-400">Documento</th>
                  <th className="px-6 py-4 text-[11px] font-black uppercase tracking-widest text-slate-400">Fecha</th>
                  <th className="px-6 py-4 text-[11px] font-black uppercase tracking-widest text-slate-400">Estado</th>
                  <th className="px-6 py-4 text-[11px] font-black uppercase tracking-widest text-slate-400 text-right">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-50">
                {entregas.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="px-6 py-12 text-center text-slate-400 italic">No hay entregas registradas todavía.</td>
                  </tr>
                ) : (
                  entregas.map((e) => (
                    <tr key={e.id} className="hover:bg-slate-50/80 transition-colors group">
                      <td className="px-6 py-4">
                        <div className="flex flex-col">
                          <span className="font-bold text-slate-800">{e.student ? `${e.student.nombre} ${e.student.apellido}` : 'Alumno'}</span>
                          <div className="flex items-center gap-2">
                            <span className="text-[10px] text-slate-400 font-medium">PK: {e.student?.documento || e.student?.id}</span>
                            {/* Alerta inteligente de completitud */}
                            {entregas.filter(ent => ent.studentId === e.studentId && ent.status === 'Aprobado').length >= 4 && (
                              <span className="bg-emerald-500 text-white text-[8px] font-black px-1.5 py-0.5 rounded-full shadow-sm animate-pulse">
                                COMPLETO
                              </span>
                            )}
                          </div>
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span className="text-sm font-semibold text-slate-600 line-clamp-1" title={e.documentType}>{e.documentType}</span>
                      </td>
                      <td className="px-6 py-4">
                        <span className="text-xs text-slate-500">{new Date(e.uploadDate).toLocaleDateString()}</span>
                      </td>
                      <td className="px-6 py-4">
                        <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-wider border ${
                          e.status === 'Aprobado' ? 'bg-emerald-50 text-emerald-700 border-emerald-200' : 
                          e.status === 'Reentregar' ? 'bg-rose-50 text-rose-700 border-rose-200' : 
                          'bg-amber-50 text-amber-700 border-amber-200'
                        }`}>
                          {e.status}
                        </span>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                          <a 
                            href={`${API_URL}/api/horas-campo/download/${e.id}`}
                            target="_blank"
                            rel="noreferrer"
                            className="p-2 bg-slate-100 hover:bg-[#0ffff4]/20 text-slate-600 hover:text-[#00968f] rounded-lg transition-all border border-slate-200"
                            title="Descargar archivo"
                          >
                            <Download className="w-4 h-4" />
                          </a>
                          <button 
                            onClick={() => handleUpdateEntregaEstado(e.id, 'Aprobado')}
                            className="p-2 bg-emerald-50 hover:bg-emerald-100 text-emerald-600 rounded-lg transition-all border border-emerald-200"
                            title="Aprobar"
                          >
                            <CheckCircle2 className="w-4 h-4" />
                          </button>
                          <button 
                            onClick={() => {
                              setConfirmInput('');
                              setConfirmModal({
                                open: true,
                                title: 'Solicitar Nueva Entrega',
                                message: 'Ingresá las observaciones para el alumno:',
                                type: 'warning',
                                withInput: true,
                                onConfirm: (val) => {
                                  handleUpdateEntregaEstado(e.id, 'Reentregar', val);
                                  setConfirmModal(p => ({...p, open: false}));
                                }
                              });
                            }}
                            className="p-2 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg transition-all border border-rose-200"
                            title="Marcar para reentregar"
                          >
                            <XCircle className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    );
  };

  const renderUsuarios = () => {
    const handleEditClick = (u: any) => {
      setEditingUser(u);
      setNewUserNombre(u.nombre);
      setNewUserEmail(u.email);
      setNewUserPassword(''); // No mostramos el password
      setNewUserRole(u.role);
      setNewUserPermissions(u.permissions || { 'analiticos': 'none', 'horas-campo': 'none' });
    };

    const togglePermission = (module: string) => {
      setNewUserPermissions(prev => {
        const current = prev[module] || 'none';
        const next = current === 'none' ? 'editor' : 'none';
        return { ...prev, [module]: next };
      });
    };

    const updateModuleRole = (module: string, role: string) => {
      setNewUserPermissions(prev => ({ ...prev, [module]: role }));
    };

    return (
      <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-8">
        <h2 className="text-3xl font-bold text-slate-900">Gestión de Usuarios</h2>

        {/* Formulario crear/editar */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6">
          <h3 className="text-lg font-bold text-slate-800 mb-4">{editingUser ? `Editando Usuario: ${editingUser.nombre}` : 'Crear nuevo usuario'}</h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">Nombre</label>
              <input
                value={newUserNombre}
                onChange={e => setNewUserNombre(e.target.value)}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                placeholder="Ej: Maria Gonzalez"
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">Email</label>
              <input
                type="email"
                value={newUserEmail}
                onChange={e => setNewUserEmail(e.target.value)}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                placeholder="editor@ejemplo.com"
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">Contrasena {editingUser && '(dejar vacío para no cambiar)'}</label>
              <input
                type="text"
                value={newUserPassword}
                onChange={e => setNewUserPassword(e.target.value)}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none"
                placeholder={editingUser ? "********" : "Contrasena temporal"}
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-1.5">Rol Global</label>
              <select
                value={newUserRole}
                onChange={e => setNewUserRole(e.target.value as any)}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm focus:ring-2 focus:ring-[#00968f] outline-none bg-white"
              >
                <option value="admin">Superadmin (Acceso Total)</option>
                <option value="editor">Editor (Personalizado)</option>
                <option value="viewer">Viewer (Solo Lectura)</option>
              </select>
            </div>
          </div>

          {newUserRole !== 'admin' && (
            <div className="space-y-4 mb-6">
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest">Permisos por Módulo</label>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {/* Modulo Analíticos */}
                <div className={`p-4 rounded-xl border transition-all ${newUserPermissions['analiticos'] !== 'none' ? 'border-[#0ffff4] bg-[#0ffff4]/5' : 'border-slate-100 bg-slate-50'}`}>
                  <div className="flex items-center justify-between mb-3">
                    <div className="flex items-center gap-2">
                      <Users className="w-5 h-5 text-slate-400" />
                      <span className="font-bold text-slate-700">Módulo Analíticos</span>
                    </div>
                    <button 
                      onClick={() => togglePermission('analiticos')}
                      className={`px-3 py-1 rounded-full text-[10px] font-black uppercase transition-all ${newUserPermissions['analiticos'] !== 'none' ? 'bg-[#002d2b] text-white' : 'bg-slate-200 text-slate-500'}`}
                    >
                      {newUserPermissions['analiticos'] !== 'none' ? 'Habilitado' : 'Deshabilitado'}
                    </button>
                  </div>
                  {newUserPermissions['analiticos'] !== 'none' && (
                    <select
                      value={newUserPermissions['analiticos']}
                      onChange={e => updateModuleRole('analiticos', e.target.value)}
                      className="w-full px-3 py-1.5 border border-slate-200 rounded-lg text-xs bg-white outline-none focus:ring-2 focus:ring-[#00968f]"
                    >
                      <option value="viewer">Solo Lectura (Viewer)</option>
                      <option value="editor">Edición (Editor)</option>
                    </select>
                  )}
                </div>

                {/* Modulo Horas de Campo */}
                <div className={`p-4 rounded-xl border transition-all ${newUserPermissions['horas-campo'] !== 'none' ? 'border-[#0ffff4] bg-[#0ffff4]/5' : 'border-slate-100 bg-slate-50'}`}>
                  <div className="flex items-center justify-between mb-3">
                    <div className="flex items-center gap-2">
                      <LayoutDashboard className="w-5 h-5 text-slate-400" />
                      <span className="font-bold text-slate-700">Horas de Campo</span>
                    </div>
                    <button 
                      onClick={() => togglePermission('horas-campo')}
                      className={`px-3 py-1 rounded-full text-[10px] font-black uppercase transition-all ${newUserPermissions['horas-campo'] !== 'none' ? 'bg-[#002d2b] text-white' : 'bg-slate-200 text-slate-500'}`}
                    >
                      {newUserPermissions['horas-campo'] !== 'none' ? 'Habilitado' : 'Deshabilitado'}
                    </button>
                  </div>
                  {newUserPermissions['horas-campo'] !== 'none' && (
                    <select
                      value={newUserPermissions['horas-campo']}
                      onChange={e => updateModuleRole('horas-campo', e.target.value)}
                      className="w-full px-3 py-1.5 border border-slate-200 rounded-lg text-xs bg-white outline-none focus:ring-2 focus:ring-[#00968f]"
                    >
                      <option value="viewer">Solo Alumno (Summary)</option>
                      <option value="admin">Administrador del Módulo</option>
                    </select>
                  )}
                </div>
              </div>
            </div>
          )}

          <div className="flex justify-end gap-3 mt-4">
            {editingUser && (
              <button
                onClick={() => {
                  setEditingUser(null);
                  setNewUserNombre('');
                  setNewUserEmail('');
                  setNewUserPassword('');
                  setNewUserRole('editor');
                  setNewUserPermissions({ 'analiticos': 'none', 'horas-campo': 'none' });
                }}
                className="px-6 py-2 rounded-lg text-sm font-bold text-slate-500 hover:bg-slate-100 transition-all"
              >
                Cancelar Edición
              </button>
            )}
            <button
              onClick={handleCreateUser}
              className="bg-[#002d2b] hover:bg-[#00968f] text-white px-6 py-2 rounded-lg text-sm font-bold transition-all shadow hover:-translate-y-0.5 active:scale-[0.98]"
            >
              {editingUser ? 'Actualizar Usuario' : 'Crear Usuario'}
            </button>
          </div>
        </div>

        {/* Lista de usuarios */}
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="p-5 border-b border-slate-100">
            <h3 className="text-lg font-bold text-slate-800">Usuarios del Sistema ({appUsers.length})</h3>
          </div>
          {appUsers.length === 0 ? (
            <div className="p-8 text-center text-slate-400">
              <p>No hay usuarios registrados todavía.</p>
            </div>
          ) : (
            <div className="divide-y divide-slate-50">
              {appUsers.map(u => (
                <div key={u.id} className="flex items-center justify-between px-6 py-5 hover:bg-slate-50 transition-colors">
                  <div className="flex items-center gap-4">
                    <div className={`p-2 rounded-xl ${u.role === 'admin' ? 'bg-[#0ffff4]/20' : 'bg-emerald-100'}`}>
                      <User className={`w-5 h-5 ${u.role === 'admin' ? 'text-[#002d2b]' : 'text-emerald-700'}`} />
                    </div>
                    <div>
                      <p className="font-semibold text-slate-900">{u.nombre}</p>
                      <p className="text-sm text-slate-500">{u.email}</p>
                    </div>
                  </div>
                  
                  <div className="flex items-center gap-6">
                    {/* Resumen de permisos */}
                    <div className="hidden md:flex items-center gap-2">
                      {u.role === 'admin' ? (
                        <span className="px-2 py-0.5 bg-slate-900 text-white text-[8px] font-black rounded uppercase tracking-widest">Superadmin</span>
                      ) : (
                        <>
                          <span className={`px-2 py-0.5 rounded text-[8px] font-black uppercase tracking-widest border ${u.permissions?.['analiticos'] !== 'none' ? 'bg-[#0ffff4]/10 text-[#002d2b] border-[#0ffff4]/30' : 'bg-slate-100 text-slate-300 border-slate-200'}`}>
                            Analíticos: {u.permissions?.['analiticos'] || 'none'}
                          </span>
                          <span className={`px-2 py-0.5 rounded text-[8px] font-black uppercase tracking-widest border ${u.permissions?.['horas-campo'] !== 'none' ? 'bg-[#0ffff4]/10 text-[#002d2b] border-[#0ffff4]/30' : 'bg-slate-100 text-slate-300 border-slate-200'}`}>
                            Horas Campo: {u.permissions?.['horas-campo'] || 'none'}
                          </span>
                        </>
                      )}
                    </div>

                    <div className="flex items-center gap-2">
                      <button
                        onClick={() => handleEditClick(u)}
                        className="p-2 text-slate-400 hover:bg-slate-100 hover:text-[#00968f] rounded-lg transition-colors"
                        title="Editar usuario"
                      >
                        <Edit className="w-4 h-4" />
                      </button>
                      {u.role !== 'admin' && (
                        <button
                          onClick={() => handleDeleteUser(u.id)}
                          className="p-2 text-red-300 hover:bg-red-50 hover:text-red-500 rounded-lg transition-colors"
                          title="Eliminar usuario"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="h-screen font-sans flex overflow-hidden text-slate-900">
      {/* SIDEBAR */}
      <aside className="w-72 bg-[#002d2b] flex flex-col shadow-2xl relative z-20 shrink-0 text-white">
        <div className="p-6 flex items-center border-b border-white/10">
          <img src={logoHorizontal} alt="Escuela Maradona Menotti" className="h-10 w-auto object-contain" />
        </div>

        <nav className="flex-1 px-4 py-6 space-y-8 overflow-y-auto custom-scrollbar">
          {/* SECCIÓN DASHBOARD */}
          {(user.role === 'admin' || user.role === 'editor') && user.email !== 'horasdecampo@maradonamenotti.com.ar' && (
            <div className="space-y-2">
              <p className="px-4 text-[10px] font-black uppercase tracking-[0.2em] text-white/30 mb-2">General</p>
              <button
                onClick={() => setActiveTab('dashboard')}
                className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === 'dashboard' ? 'bg-[#0ffff4]/10 text-[#0ffff4] font-bold shadow-sm border border-[#0ffff4]/30' : 'text-white/60 hover:bg-white/5 hover:text-white font-medium'}`}
              >
                <LayoutDashboard className="w-5 h-5" />
                Dashboard General
              </button>
            </div>
          )}

          {/* SECCIÓN ANALÍTICOS */}
          {user.role !== 'student' && user.email !== 'horasdecampo@maradonamenotti.com.ar' && (user.role === 'admin' || (user.permissions?.['analiticos'] && user.permissions['analiticos'] !== 'none')) && (
            <div className="space-y-2">
              <p className="px-4 text-[10px] font-black uppercase tracking-[0.2em] text-white/30 mb-2">Padrón de Alumnos</p>
              <button
                onClick={() => setActiveTab('alumnos')}
                className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === 'alumnos' ? 'bg-[#0ffff4]/10 text-[#0ffff4] font-bold shadow-sm border border-[#0ffff4]/30' : 'text-white/60 hover:bg-white/5 hover:text-white font-medium'}`}
              >
                <Users className="w-5 h-5" />
                Listado de Alumnos
              </button>
            </div>
          )}

          {/* SECCIÓN HORAS DE CAMPO */}
          {(user.role === 'student' || user.email === 'horasdecampo@maradonamenotti.com.ar' || (user.role === 'admin' && user.email !== 'titulos@maradonamenotti.com.ar')) && (
            <div className="space-y-2">
              <p className="px-4 text-[10px] font-black uppercase tracking-[0.2em] text-white/30 mb-2">Horas de Campo</p>
              
              {/* Para ADMINS de Horas de Campo o General */}
              {(user.role === 'admin' || user.permissions?.['horas-campo'] === 'admin') ? (
                <div className="space-y-1">
                  <button
                    onClick={() => { setActiveTab('horas-campo'); setFieldHourSubTab('dashboard'); }}
                    className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === 'horas-campo' && fieldHourSubTab === 'dashboard' ? 'bg-[#0ffff4]/10 text-[#0ffff4] font-bold shadow-sm border border-[#0ffff4]/30' : 'text-white/60 hover:bg-white/5 hover:text-white font-medium'}`}
                  >
                    <LayoutDashboard className="w-5 h-5" />
                    Dashboard Inteligente
                  </button>
                  <button
                    onClick={() => { setActiveTab('horas-campo'); setFieldHourSubTab('entrega'); }}
                    className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === 'horas-campo' && fieldHourSubTab === 'entrega' ? 'bg-[#0ffff4]/10 text-[#0ffff4] font-bold shadow-sm border border-[#0ffff4]/30' : 'text-white/60 hover:bg-white/5 hover:text-white font-medium'}`}
                  >
                    <FileText className="w-5 h-5" />
                    Gestión de Entregas
                  </button>
                </div>
              ) : (
                /* Para ALUMNOS */
                <div className="space-y-1">
                  <button
                    onClick={() => { setActiveTab('horas-campo'); setFieldHourSubTab('summary'); }}
                    className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === 'horas-campo' && fieldHourSubTab === 'summary' ? 'bg-[#0ffff4]/10 text-[#0ffff4] font-bold shadow-sm border border-[#0ffff4]/30' : 'text-white/60 hover:bg-white/5 hover:text-white font-medium'}`}
                  >
                    <LayoutDashboard className="w-5 h-5" />
                    Horas de Campo
                  </button>
                  {activeTab === 'horas-campo' && (
                    <div className="pl-4 mt-1 space-y-1">
                      <button 
                        onClick={() => setFieldHourSubTab('general')}
                        className={`w-full flex items-center gap-2.5 px-4 py-2 rounded-lg text-xs transition-all ${fieldHourSubTab === 'general' ? 'text-[#0ffff4] font-bold bg-white/5' : 'text-white/40 hover:text-white hover:bg-white/5 font-medium'}`}
                      >
                        <Info className="w-3.5 h-3.5" />
                        Información General
                      </button>
                      <button 
                        onClick={() => setFieldHourSubTab('descarga')}
                        className={`w-full flex items-center gap-2.5 px-4 py-2 rounded-lg text-xs transition-all ${fieldHourSubTab === 'descarga' ? 'text-[#0ffff4] font-bold bg-white/5' : 'text-white/40 hover:text-white hover:bg-white/5 font-medium'}`}
                      >
                        <Download className="w-3.5 h-3.5" />
                        Descargar Planillas
                      </button>
                      <button 
                        onClick={() => setFieldHourSubTab('tp')}
                        className={`w-full flex items-center gap-2.5 px-4 py-2 rounded-lg text-xs transition-all ${fieldHourSubTab === 'tp' ? 'text-[#0ffff4] font-bold bg-white/5' : 'text-white/40 hover:text-white hover:bg-white/5 font-medium'}`}
                      >
                        <FileText className="w-3.5 h-3.5" />
                        Trabajo Práctico
                      </button>
                      <button 
                        onClick={() => setFieldHourSubTab('entrega')}
                        className={`w-full flex items-center gap-2.5 px-4 py-2 rounded-lg text-xs transition-all ${fieldHourSubTab === 'entrega' ? 'text-[#0ffff4] font-bold bg-white/5' : 'text-white/40 hover:text-white hover:bg-white/5 font-medium'}`}
                      >
                        <Upload className="w-3.5 h-3.5" />
                        Entrega de Planillas
                      </button>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {/* SECCIÓN Gestión Usuarios - Al final */}
          {user.role === 'admin' && user.email !== 'horasdecampo@maradonamenotti.com.ar' && (
            <div className="space-y-2">
              <p className="px-4 text-[10px] font-black uppercase tracking-[0.2em] text-white/30 mb-2">Acceso y Seguridad</p>
              <button
                onClick={() => { setActiveTab('usuarios'); fetchAppUsers(); }}
                className={`w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all ${activeTab === 'usuarios' ? 'bg-[#0ffff4]/10 text-[#0ffff4] font-bold shadow-sm border border-[#0ffff4]/30' : 'text-white/60 hover:bg-white/5 hover:text-white font-medium'}`}
              >
                <UserPlus className="w-5 h-5" />
                Gestión de Usuarios
              </button>
            </div>
          )}

        </nav>

        <div className="p-5 bg-white/5 mt-auto border-t border-white/10">
          <div className="flex flex-col gap-4">
            <div className="flex items-center gap-3 px-2">
              <div className="bg-white/10 p-2.5 rounded-xl shadow-sm ring-2 ring-[#0ffff4]/25">
                <User className="w-5 h-5 text-[#0ffff4]" />
              </div>
              <div className="flex flex-col overflow-hidden">
                <span className="text-sm font-semibold text-white truncate" title={user.name}>{user.name}</span>
                <span className="text-[11px] font-black uppercase text-[#0ffff4] tracking-wider mt-0.5">{user.role}</span>
              </div>
            </div>
            <button
              onClick={handleLogout}
              className="flex items-center justify-center gap-2 w-full py-3 bg-white/10 hover:bg-white/15 text-white rounded-xl transition-all border border-white/15"
            >
              <LogOut className="w-4 h-4" />
              <span className="text-sm font-bold">Cerrar Sesión</span>
            </button>
          </div>
        </div>
      </aside>

      {/* MAIN CONTENT HEADER + SCROLLABLE AREA */}
      <div className="flex-1 flex flex-col h-full bg-transparent min-w-0">
        <header className="h-20 bg-white/90 backdrop-blur-xl border-b border-[#00968f26] px-8 flex items-center sticky top-0 z-10 shadow-sm">
          <h1 className="text-[#002d2b] text-2xl font-black tracking-tight">
            {activeTab === 'dashboard' ? 'Módulo Analíticos: Panel de Control' : 
             activeTab === 'usuarios' ? 'Gestión de Usuarios' : 
             activeTab === 'horas-campo' ? `Horas de Campo: ${fieldHourSubTab === 'dashboard' ? 'Dashboard Inteligente' : 'Gestión de Entregas'}` : 'Padrón de Alumnos'}
          </h1>
        </header>

        <main className="flex-1 overflow-y-auto p-8 custom-scrollbar">
          <div className="max-w-7xl mx-auto pb-12">
            {activeTab === 'dashboard' ? renderDashboard() : 
             activeTab === 'usuarios' ? renderUsuarios() : 
             activeTab === 'horas-campo' ? renderHorasCampo() : renderAlumnos()}
          </div>
        </main>
      </div>

      {renderAnaliticoModals()}

      {/* TOASTER para notificaciones */}
      <Toaster position="bottom-right" toastOptions={{ duration: 4000, style: { fontWeight: 'bold' } }} />

      {/* MODAL DE CONFIRMACIÓN CUSTOM */}
      <AnimatePresence>
        {confirmModal.open && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 10 }}
              className="bg-white rounded-2xl shadow-2xl max-w-md w-full overflow-hidden border border-slate-200 flex flex-col"
            >
              <div className={`p-5 flex items-start gap-4 border-b border-slate-100 ${confirmModal.type === 'danger' ? 'bg-red-50' : confirmModal.type === 'warning' ? 'bg-amber-50' : 'bg-[#0ffff4]/10'}`}>
                <div className={`shrink-0 p-3 rounded-xl ${confirmModal.type === 'danger' ? 'bg-red-100 text-red-600' : confirmModal.type === 'warning' ? 'bg-amber-100 text-amber-600' : 'bg-[#00968f]/10 text-[#00968f]'}`}>
                  {confirmModal.type === 'danger' ? <Trash2 className="w-6 h-6" /> : confirmModal.type === 'warning' ? <AlertTriangle className="w-6 h-6" /> : <CheckCircle2 className="w-6 h-6" />}
                </div>
                <div>
                  <h3 className="text-lg font-bold text-slate-800">{confirmModal.title}</h3>
                  <p className="text-slate-600 text-sm mt-1">{confirmModal.message}</p>
                </div>
              </div>

              {confirmModal.withInput && (
                <div className="p-5 border-b border-slate-100">
                  <label className="block text-sm font-bold text-slate-700 mb-2">{confirmModal.inputLabel || 'Motivo'}</label>
                  <input
                    type="text"
                    value={confirmInput}
                    onChange={(e) => setConfirmInput(e.target.value)}
                    placeholder={confirmModal.inputPlaceholder}
                    className="w-full px-4 py-2.5 rounded-lg border border-slate-300 focus:ring-2 focus:ring-[#00968f] outline-none transition-all shadow-sm"
                    autoFocus
                  />
                </div>
              )}

              <div className="p-4 bg-slate-50 flex justify-end gap-3">
                <button
                  onClick={() => setConfirmModal(prev => ({ ...prev, open: false }))}
                  className="px-5 py-2.5 rounded-lg text-sm font-bold text-slate-600 hover:bg-slate-200 transition-colors"
                >
                  Cancelar
                </button>
                <button
                  onClick={() => {
                    if (confirmModal.withInput && !confirmInput.trim()) {
                      toast.error('Debe completar este campo.');
                      return;
                    }
                    confirmModal.onConfirm(confirmInput);
                  }}
                  className={`px-5 py-2.5 rounded-lg text-sm font-bold text-white transition-colors shadow-sm ${confirmModal.type === 'danger' ? 'bg-red-600 hover:bg-red-700' : confirmModal.type === 'warning' ? 'bg-amber-600 hover:bg-amber-700' : 'bg-[#002d2b] hover:bg-[#00968f]'}`}
                >
                  Confirmar
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {importConfig.isOpen && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between p-6 border-b border-slate-100 bg-slate-50/50">
              <h2 className="text-xl font-bold text-slate-800">
                {importConfig.mode === 'db' ? 'Importar Notas' : 'Generar ZIP Masivo'}
              </h2>
              <button type="button" onClick={() => setImportConfig({ isOpen: false, mode: null })} className="text-slate-400 hover:text-slate-600 transition-colors bg-white p-1 rounded-full shadow-sm border border-slate-100">
                <X className="w-5 h-5" />
              </button>
            </div>
            <form onSubmit={handleModalSubmit} className="p-6 space-y-5">
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5">Archivo de Notas Excel (.xlsx)</label>
                <input required type="file" accept=".xlsx, .xls" name="excelFile" className="w-full border border-slate-200 rounded-xl p-2 focus:ring-2 outline-none focus:ring-[#0ffff4]/50" />
              </div>
              <div>
                <label className="block text-sm font-semibold text-slate-700 mb-1.5">Licencia Correspondiente</label>
                <select required name="licencia" className="w-full border border-slate-200 rounded-xl p-2.5 bg-white focus:ring-2 outline-none focus:ring-[#0ffff4]/50">
                  <option value="CB">Licencia CB</option>
                  <option value="A">Licencia A</option>
                  <option value="B">Licencia B</option>
                  <option value="PRO">Licencia PRO</option>
                  <option value="TD1">TD1 (Trayectoria I)</option>
                  <option value="TD2">TD2 (Trayectoria II)</option>
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1 opacity-90">Fecha Fin Cursada</label>
                  <input type="date" name="fecha_fin_cursada" className="w-full text-sm border border-slate-200 rounded-xl p-2" />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1 opacity-90">Fecha Emisión</label>
                  <input type="date" name="fecha_emision" className="w-full text-sm border border-slate-200 rounded-xl p-2" />
                </div>
              </div>
              <div className="pt-5 border-t border-slate-100 flex gap-3">
                <button type="button" onClick={() => setImportConfig({ isOpen: false, mode: null })} className="flex-1 py-3 bg-white border border-slate-200 rounded-xl font-bold text-slate-600 hover:bg-slate-50 transition-colors">Cancelar</button>
                <button type="submit" disabled={isUploading} className="flex-1 py-3 bg-[#002d2b] hover:bg-[#00968f] text-white rounded-xl font-bold transition-all shadow shadow-[#002d2b]/30 hover:-translate-y-0.5 active:scale-[0.98]">
                  {isUploading ? 'Procesando...' : (importConfig.mode === 'db' ? 'Subir Notas' : 'Crear ZIP')}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {newStudentModal && (
        <div className="fixed inset-0 z-[120] flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg overflow-hidden animate-in zoom-in-95 duration-200">
            <div className="flex items-center justify-between p-6 border-b border-slate-100 bg-slate-50/50">
              <h2 className="text-xl font-bold text-slate-800">Nuevo Alumno</h2>
              <button type="button" onClick={() => setNewStudentModal(false)} className="text-slate-400 hover:text-slate-600 transition-colors bg-white p-1 rounded-full shadow-sm border border-slate-100">
                <X className="w-5 h-5" />
              </button>
            </div>
            <form onSubmit={handleCreateStudent} className="p-6 space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Nombre</label>
                  <input required value={newStuNombre} onChange={e => setNewStuNombre(e.target.value)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm" />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Apellido</label>
                  <input required value={newStuApellido} onChange={e => setNewStuApellido(e.target.value)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm" />
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Documento</label>
                  <input required value={newStuDni} onChange={e => setNewStuDni(e.target.value)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm" />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Nacionalidad</label>
                  <input value={newStuNacionalidad} onChange={e => setNewStuNacionalidad(e.target.value)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm uppercase" />
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Carrera / Licencia</label>
                  <select value={newStuLicencia} onChange={e => setNewStuLicencia(e.target.value as any)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm bg-white">
                    <option value="CB">Licencia CB</option>
                    <option value="A">Licencia A</option>
                    <option value="PRO">Licencia PRO</option>
                    <option value="TD1">TD1</option>
                    <option value="TD2">TD2</option>
                    <option value="ACTUALIZACION">Curso de Actualización</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-slate-600 mb-1">Fecha de Emisión</label>
                  <input type="date" value={newStuFechaEmision} onChange={e => setNewStuFechaEmision(e.target.value)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm" />
                </div>
              </div>
              {newStuLicencia !== 'ACTUALIZACION' && (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-xs font-semibold text-slate-600 mb-1">Fecha de Graduación / Fin de cursada</label>
                    <input type="date" value={newStuFechaFin} onChange={e => setNewStuFechaFin(e.target.value)} className="w-full border border-slate-200 rounded-xl p-2.5 text-sm" />
                  </div>
                  <div className="flex items-center text-xs text-slate-500 bg-slate-50 border border-slate-200 rounded-xl px-3">
                    Al guardar se crean todas las materias de la licencia con nota 0.
                  </div>
                </div>
              )}
              {newStuLicencia === 'ACTUALIZACION' && (
                <p className="text-xs text-slate-500 bg-amber-50 border border-amber-200 rounded-xl px-3 py-2">
                  Solo se piden los datos que requiere el certificado del Curso de Actualización.
                </p>
              )}
              <div className="pt-4 border-t border-slate-100 flex gap-3">
                <button type="button" onClick={() => setNewStudentModal(false)} className="flex-1 py-3 bg-white border border-slate-200 rounded-xl font-bold text-slate-600 hover:bg-slate-50 transition-colors">Cancelar</button>
                <button type="submit" disabled={isUploading} className="flex-1 py-3 bg-[#002d2b] hover:bg-[#00968f] text-white rounded-xl font-bold transition-all shadow hover:-translate-y-0.5 active:scale-[0.98]">
                  {isUploading ? 'Guardando...' : 'Crear alumno'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppContent />
    </ErrorBoundary>
  );
}














