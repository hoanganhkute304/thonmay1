import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  collection, 
  addDoc, 
  setDoc,
  query, 
  where, 
  orderBy, 
  onSnapshot, 
  doc,
  deleteDoc,
  getDocFromServer,
  getDocs,
  limit
} from 'firebase/firestore';
import { onAuthStateChanged, User, signInAnonymously } from 'firebase/auth';
import { db, auth, logout } from './firebase';
import { toPng, toBlob } from 'html-to-image';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { 
  Calculator, 
  History, 
  Save, 
  LogOut, 
  LogIn, 
  TrendingUp, 
  TrendingDown, 
  AlertCircle,
  FileSpreadsheet,
  Calendar as CalendarIcon,
  DollarSign,
  Plus,
  Trash2,
  Copy,
  Check,
  User as UserIcon,
  Wallet,
  ArrowRightLeft,
  PieChart,
  Building2,
  Settings,
  ShieldCheck,
  ChevronRight,
  PlusCircle,
  MapPin,
  Phone,
  Image as ImageIcon,
  Edit2,
  Download
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

// --- Constants ---

// --- Types ---
interface SalaryAdvance {
  employee: string;
  method: 'salary_cash' | 'salary_bank' | 'cash' | 'bank' | 'debt' | 'fine';
  amount: number;
}

interface Employee {
  id: string;
  name: string;
  branchId: string;
}

interface ExpenseCategory {
  id: string;
  name: string;
  branchId: string;
}

interface ExpenseItem {
  description: string;
  amount: number;
  category: string;
}

interface RevenueRecord {
  id?: string;
  date: string;
  duDau: number;
  revenue: number;
  cash: number;
  bank: number;
  reserve: number;
  spentCash: number;
  spentBank: number;
  spentCashItems: ExpenseItem[];
  spentBankItems: ExpenseItem[];
  salaryAdvances: SalaryAdvance[];
  totalSalaryAdvances: number;
  totalBook: number;
  totalBookExpected: number;
  totalExpenses: number;
  totalActualAssets: number;
  recoveredMoney: number;
  diff: number;
  netProfit: number;
  branchId: string;
  uid: string;
  createdAt: string;
}

interface Branch {
  id: string;
  username?: string;
  password?: string;
  role: 'admin' | 'branch';
  displayName: string;
  address?: string;
  phone?: string;
  createdAt: string;
}

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: any;
}

// --- Helper Functions ---
function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

const formatCurrency = (val: number) => {
  return new Intl.NumberFormat('vi-VN').format(val);
};

const parseCurrency = (val: string) => {
  return Number(val.replace(/[^0-9]/g, '')) || 0;
};

// --- Components ---
const CurrencyInput = ({ 
  value, 
  onChange, 
  placeholder, 
  className = "" 
}: { 
  value: string, 
  onChange: (val: string) => void, 
  placeholder?: string,
  className?: string
}) => {
  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const rawValue = e.target.value.replace(/[^0-9]/g, '');
    if (rawValue === '') {
      onChange('');
      return;
    }
    const formattedValue = new Intl.NumberFormat('vi-VN').format(Number(rawValue));
    onChange(formattedValue);
  };

  return (
    <div className="relative w-full">
      <input
        type="text"
        inputMode="numeric"
        value={value}
        onChange={handleChange}
        placeholder={placeholder}
        className={`w-full bg-[#111] border border-white/5 p-3 rounded-lg text-[#00ff88] focus:outline-none focus:ring-2 focus:ring-[#00ff88]/20 transition-all placeholder:text-gray-700 font-mono text-sm ${className}`}
      />
      <span className="absolute right-3 top-1/2 -translate-y-1/2 text-[9px] text-gray-600 font-mono pointer-events-none">VND</span>
    </div>
  );
};

export default function App() {
  const [user, setUser] = useState<User | null>(null);
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [isSessionSynced, setIsSessionSynced] = useState(false);
  const [currentBranch, setCurrentBranch] = useState<Branch | null>(null);
  const [records, setRecords] = useState<RevenueRecord[]>([]);
  const [branches, setBranches] = useState<Branch[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [view, setView] = useState<'calc' | 'history' | 'summary' | 'admin' | 'admin-revenue' | 'settings'>('calc');
  const [copySuccess, setCopySuccess] = useState(false);
  const [saveSuccess, setSaveSuccess] = useState(false);
  const [isExporting, setIsExporting] = useState<string | null>(null);
  const [recordToDelete, setRecordToDelete] = useState<string | null>(null);
  const recordRefs = useRef<{ [key: string]: HTMLDivElement | null }>({});
  const [summaryMonth, setSummaryMonth] = useState(new Date().toISOString().slice(0, 7)); // YYYY-MM
  const [summaryPeriod, setSummaryPeriod] = useState<'day' | 'month' | 'year'>('month');
  const [summaryDate, setSummaryDate] = useState(new Date().toISOString().split('T')[0]);
  const [summaryYear, setSummaryYear] = useState(new Date().getFullYear().toString());
  const [summaryBranchId, setSummaryBranchId] = useState<string>('all');
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [expenseCategories, setExpenseCategories] = useState<ExpenseCategory[]>([]);

  // Login State
  const [loginBranch, setLoginBranch] = useState('');
  const [isLoggingIn, setIsLoggingIn] = useState(false);

  // Admin: Create Branch State
  const [newBranchId, setNewBranchId] = useState('');
  const [newBranchUser, setNewBranchUser] = useState('');
  const [newBranchPass, setNewBranchPass] = useState('');
  const [newBranchName, setNewBranchName] = useState('');
  const [newBranchAddr, setNewBranchAddr] = useState('');
  const [newBranchPhone, setNewBranchPhone] = useState('');
  const [editingBranchId, setEditingBranchId] = useState<string | null>(null);

  // Employee Management State
  const [newEmployeeName, setNewEmployeeName] = useState('');
  const [newEmployeeBranch, setNewEmployeeBranch] = useState('all');
  const [editingEmployeeId, setEditingEmployeeId] = useState<string | null>(null);

  // Expense Category Management State
  const [newCategoryName, setNewCategoryName] = useState('');
  const [newCategoryBranch, setNewCategoryBranch] = useState('all');
  const [editingCategoryId, setEditingCategoryId] = useState<string | null>(null);

  // Form State
  const [workingDate, setWorkingDate] = useState(new Date().toISOString().split('T')[0]);
  const [duDau, setDuDau] = useState<string>('');
  const [revenue, setRevenue] = useState<string>('');
  const [cash, setCash] = useState<string>('');
  const [bank, setBank] = useState<string>('');
  const [reserve, setReserve] = useState<string>('');
  const [spentCashItems, setSpentCashItems] = useState<ExpenseItem[]>([]);
  const [spentBankItems, setSpentBankItems] = useState<ExpenseItem[]>([]);
  const [salaryAdvances, setSalaryAdvances] = useState<SalaryAdvance[]>([]);

  // --- Auth Setup ---
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      setUser(currentUser);
      setIsAuthReady(true);
      
      // Check if this is the default admin via Google
      if (currentUser && currentUser.email === 'kute123kuto123@gmail.com' && currentUser.emailVerified) {
        const adminBranch: Branch = {
          id: 'admin',
          username: 'admin',
          password: 'admin',
          role: 'admin',
          displayName: 'Hệ Thống Admin',
          createdAt: new Date().toISOString()
        };
        setCurrentBranch(adminBranch);
        localStorage.setItem('currentBranch', JSON.stringify(adminBranch));
        
        // Sync session info
        try {
          await setDoc(doc(db, 'users', currentUser.uid), {
            branchId: 'admin',
            role: 'admin',
            updatedAt: new Date().toISOString()
          });
          setIsSessionSynced(true);
        } catch (err) {
          console.error("Admin session sync failed:", err);
          setIsSessionSynced(true);
        }
        return;
      }
      
      // Load branch from local storage if exists
      const savedBranch = localStorage.getItem('currentBranch');
      if (savedBranch && currentUser) {
        try {
          const branchData = JSON.parse(savedBranch);
          
          // Safety check: If ID is missing, we might have a corrupted session from a previous bug
          if (!branchData.id) {
            console.warn("Corrupted branch data found in local storage, clearing session.");
            localStorage.removeItem('currentBranch');
            setCurrentBranch(null);
            setIsSessionSynced(false);
            return;
          }

          setCurrentBranch(branchData);
          
          // Re-sync session info to Firestore just in case
          await setDoc(doc(db, 'users', currentUser.uid), {
            branchId: branchData.id,
            role: branchData.role,
            updatedAt: new Date().toISOString()
          });
          setIsSessionSynced(true);
        } catch (err) {
          console.error("Session sync failed:", err);
          // Even if it fails, we might still have the doc from before
          setIsSessionSynced(true); 
        }
      } else {
        setIsSessionSynced(false);
      }
    });
    return () => unsubscribe();
  }, []);

  // --- Bootstrap Initial Branches ---
  useEffect(() => {
    if (isAuthReady) {
      const bootstrap = async () => {
        const branchesRef = collection(db, 'branches');
        const q = query(branchesRef, limit(1));
        const snapshot = await getDocs(q);
        
        if (snapshot.empty) {
          console.log("Bootstrapping initial branches...");
          const initialBranches = [
            {
              id: 'admin',
              username: 'admin',
              password: 'admin',
              role: 'admin',
              displayName: 'Hệ Thống Admin',
              createdAt: new Date().toISOString()
            },
            {
              id: 'thonmay',
              username: '0967123396',
              password: 'maoivitchet304',
              role: 'branch',
              displayName: 'Thôn Mây',
              createdAt: new Date().toISOString()
            }
          ];
          
          for (const b of initialBranches) {
            await setDoc(doc(db, 'branches', b.id), b);
          }
        }
      };
      bootstrap();
    }
  }, [isAuthReady]);

  // --- Fetch Records ---
  useEffect(() => {
    if (!isAuthReady || !user || !currentBranch || !isSessionSynced) {
      setRecords([]);
      setLoading(false);
      return;
    }

    const path = 'revenueRecords';
    let q;
    
    if (currentBranch.role === 'admin') {
      // Admin sees everything
      q = query(
        collection(db, path),
        orderBy('date', 'desc')
      );
    } else {
      // Branches see only their own
      q = query(
        collection(db, path),
        where('branchId', '==', currentBranch.id),
        orderBy('date', 'desc')
      );
    }

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const fetchedRecords = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as RevenueRecord[];
      setRecords(fetchedRecords);
      setLoading(false);
    }, (err) => {
      handleFirestoreError(err, OperationType.LIST, path);
    });

    return () => unsubscribe();
  }, [isAuthReady, user, currentBranch, isSessionSynced]);

  // --- Fetch Employees ---
  useEffect(() => {
    if (!isAuthReady || !user || !currentBranch || !isSessionSynced) {
      setEmployees([]);
      return;
    }

    const path = 'employees';
    let q;
    if (currentBranch.role === 'admin') {
      q = query(collection(db, path));
    } else {
      q = query(collection(db, path), where('branchId', 'in', [currentBranch.id, 'all']));
    }

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const fetchedEmployees = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as Employee[];
      setEmployees(fetchedEmployees);
    }, (err) => {
      handleFirestoreError(err, OperationType.LIST, path);
    });

    return () => unsubscribe();
  }, [isAuthReady, user, currentBranch, isSessionSynced]);

  // --- Fetch Expense Categories ---
  useEffect(() => {
    if (!isAuthReady || !user || !currentBranch || !isSessionSynced) {
      setExpenseCategories([]);
      return;
    }

    const path = 'expenseCategories';
    let q;
    if (currentBranch.role === 'admin') {
      q = query(collection(db, path));
    } else {
      q = query(collection(db, path), where('branchId', 'in', [currentBranch.id, 'all']));
    }

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const fetchedCategories = snapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as ExpenseCategory[];
      setExpenseCategories(fetchedCategories);
    }, (err) => {
      handleFirestoreError(err, OperationType.LIST, path);
    });

    return () => unsubscribe();
  }, [isAuthReady, user, currentBranch, isSessionSynced]);

  // --- Seed Initial Data ---
  useEffect(() => {
    const seedData = async () => {
      if (!isAuthReady || !user || !currentBranch || currentBranch.role !== 'admin' || !isSessionSynced) return;

      // Seed Categories
      const categoriesPath = 'expenseCategories';
      const categoriesSnap = await getDocs(collection(db, categoriesPath));
      if (categoriesSnap.empty) {
        const initialCategories = [
          'Trái cây', 'Mứt', 'Siro', 'Nước ngọt', 'Kem', 
          'Thế giới xanh', 'Đồ đông lạnh', 'Vật liệu', 'Khác'
        ];
        for (const cat of initialCategories) {
          await addDoc(collection(db, categoriesPath), {
            name: cat,
            branchId: 'all',
            createdAt: new Date().toISOString()
          });
        }
      }

      // Seed Employees
      const employeesPath = 'employees';
      const employeesSnap = await getDocs(collection(db, employeesPath));
      if (employeesSnap.empty) {
        const initialEmployees = [
          'Nam', 'Quyết', 'Tú', 'Trang', 'Hân', 'Thư', 'Khánh', 
          'Linh', 'Đạt', 'Minh', 'Quân', 'Công', 'Tùng', 'Anh', 
          'Hải', 'Phước', 'Thoa'
        ];
        for (const emp of initialEmployees) {
          await addDoc(collection(db, employeesPath), {
            name: emp,
            branchId: 'all',
            createdAt: new Date().toISOString()
          });
        }
      }
    };

    seedData();
  }, [isAuthReady, user, currentBranch, isSessionSynced]);

  // --- Fetch All Branches (Admin Only) ---
  useEffect(() => {
    if (currentBranch?.role === 'admin') {
      const unsubscribe = onSnapshot(collection(db, 'branches'), (snapshot) => {
        const fetchedBranches = snapshot.docs.map(doc => ({
          id: doc.id,
          ...doc.data()
        })) as Branch[];
        setBranches(fetchedBranches);
      });
      return () => unsubscribe();
    }
  }, [currentBranch]);

  // --- Login Handler ---
  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsLoggingIn(true);
    setError(null);

    try {
      // 1. Check branch in Firestore
      const branchId = loginBranch.toLowerCase().trim();
      const branchDoc = await getDocFromServer(doc(db, 'branches', branchId));
      
      if (!branchDoc.exists()) {
        throw new Error('Chi nhánh không tồn tại.');
      }

      const branchData = { id: branchDoc.id, ...branchDoc.data() } as Branch;
      
      // 2. Sign in anonymously to Firebase to get a UID for Firestore rules
      let userCredential = auth.currentUser;
      if (!userCredential) {
        const result = await signInAnonymously(auth);
        userCredential = result.user;
      }

      if (userCredential) {
        // 3. Save session info to Firestore so rules can verify branch/role
        await setDoc(doc(db, 'users', userCredential.uid), {
          branchId: branchData.id,
          role: branchData.role,
          updatedAt: new Date().toISOString()
        });
        setIsSessionSynced(true);
      }

      // 4. Set current branch
      setCurrentBranch(branchData);
      localStorage.setItem('currentBranch', JSON.stringify(branchData));
      
      // 5. Reset login form
      setLoginBranch('');
      
      if (branchData.role === 'admin') {
        setView('admin');
      } else {
        setView('calc');
      }
    } catch (err: any) {
      setError(err.message || 'Đã xảy ra lỗi khi truy cập.');
    } finally {
      setIsLoggingIn(false);
    }
  };

  const handleLogout = () => {
    setCurrentBranch(null);
    setIsSessionSynced(false);
    localStorage.removeItem('currentBranch');
    logout();
    setView('calc');
  };

  const createBranch = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentBranch) return;
    
    setSaving(true);
    try {
      const branchId = editingBranchId || newBranchId.toLowerCase().replace(/\s+/g, '');
      const existingBranch = branches.find(b => b.id === branchId);
      
      const newBranch: Branch = {
        id: branchId,
        username: newBranchUser || (existingBranch?.username || ''),
        password: newBranchPass || (existingBranch?.password || ''),
        role: existingBranch?.role || 'branch',
        displayName: newBranchName,
        address: newBranchAddr,
        phone: newBranchPhone,
        createdAt: existingBranch?.createdAt || new Date().toISOString()
      };
      
      await setDoc(doc(db, 'branches', branchId), newBranch);
      
      if (currentBranch.role === 'admin') {
        setNewBranchId('');
        setNewBranchUser('');
        setNewBranchPass('');
        setNewBranchName('');
        setNewBranchAddr('');
        setNewBranchPhone('');
        setEditingBranchId(null);
      }
      
      alert('Đã cập nhật thông tin thành công!');
    } catch (err) {
      alert('Lỗi khi lưu thông tin.');
    } finally {
      setSaving(false);
    }
  };

  const deleteBranch = async (id: string) => {
    if (!currentBranch || currentBranch.role !== 'admin') return;
    if (id === currentBranch.id) {
      alert('Không thể xóa chi nhánh đang đăng nhập.');
      return;
    }
    if (!confirm(`Bạn có chắc chắn muốn xóa chi nhánh ${id}?`)) return;

    try {
      await deleteDoc(doc(db, 'branches', id));
      alert('Đã xóa chi nhánh.');
    } catch (err) {
      alert('Lỗi khi xóa chi nhánh.');
    }
  };

  const startEditBranch = (branch: Branch) => {
    setEditingBranchId(branch.id);
    setNewBranchId(branch.id);
    setNewBranchName(branch.displayName);
    setNewBranchUser(branch.username);
    setNewBranchPass(branch.password);
    setNewBranchAddr(branch.address || '');
    setNewBranchPhone(branch.phone || '');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const createEmployee = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newEmployeeName || !currentBranch) return;
    setSaving(true);
    try {
      const empData = {
        name: newEmployeeName,
        branchId: currentBranch.role === 'admin' ? newEmployeeBranch : currentBranch.id,
        createdAt: new Date().toISOString()
      };
      if (editingEmployeeId) {
        await setDoc(doc(db, 'employees', editingEmployeeId), empData, { merge: true });
      } else {
        await addDoc(collection(db, 'employees'), empData);
      }
      setNewEmployeeName('');
      setNewEmployeeBranch(currentBranch.role === 'admin' ? 'all' : currentBranch.id);
      setEditingEmployeeId(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, 'employees');
    } finally {
      setSaving(false);
    }
  };

  const deleteEmployee = async (id: string) => {
    if (!confirm('Xác nhận xóa nhân viên này?')) return;
    try {
      await deleteDoc(doc(db, 'employees', id));
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, 'employees');
    }
  };

  const startEditEmployee = (emp: Employee) => {
    setNewEmployeeName(emp.name);
    setNewEmployeeBranch(emp.branchId);
    setEditingEmployeeId(emp.id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const createCategory = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCategoryName || !currentBranch) return;
    setSaving(true);
    try {
      const catData = {
        name: newCategoryName,
        branchId: currentBranch.role === 'admin' ? newCategoryBranch : currentBranch.id,
        createdAt: new Date().toISOString()
      };
      if (editingCategoryId) {
        await setDoc(doc(db, 'expenseCategories', editingCategoryId), catData, { merge: true });
      } else {
        await addDoc(collection(db, 'expenseCategories'), catData);
      }
      setNewCategoryName('');
      setNewCategoryBranch(currentBranch.role === 'admin' ? 'all' : currentBranch.id);
      setEditingCategoryId(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, 'expenseCategories');
    } finally {
      setSaving(false);
    }
  };

  const deleteCategory = async (id: string) => {
    if (!confirm('Xác nhận xóa phân loại này?')) return;
    try {
      await deleteDoc(doc(db, 'expenseCategories', id));
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, 'expenseCategories');
    }
  };

  const startEditCategory = (cat: ExpenseCategory) => {
    setNewCategoryName(cat.name);
    setNewCategoryBranch(cat.branchId);
    setEditingCategoryId(cat.id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  useEffect(() => {
    if (view === 'admin' && currentBranch) {
      if (currentBranch.role !== 'admin') {
        startEditBranch(currentBranch);
        setNewEmployeeBranch(currentBranch.id);
      } else {
        setNewEmployeeBranch('all');
      }
    }
  }, [view, currentBranch]);

  // --- Auto-load record when date changes ---
  useEffect(() => {
    if (isAuthReady && user && records.length > 0) {
      const existingRecord = records.find(r => r.date === workingDate);
      if (existingRecord) {
        setDuDau(existingRecord.duDau.toLocaleString('vi-VN'));
        setRevenue(existingRecord.revenue.toLocaleString('vi-VN'));
        setCash(existingRecord.cash.toLocaleString('vi-VN'));
        setBank(existingRecord.bank.toLocaleString('vi-VN'));
        setReserve(existingRecord.reserve.toLocaleString('vi-VN'));
        setSpentCashItems(existingRecord.spentCashItems || []);
        setSpentBankItems(existingRecord.spentBankItems || []);
        setSalaryAdvances(existingRecord.salaryAdvances || []);
      } else {
        // Only clear if we are moving to a date that truly has no data
        // and we were previously on a date that DID have data
        setDuDau('');
        setRevenue('');
        setCash('');
        setBank('');
        setReserve('');
        setSpentCashItems([]);
        setSpentBankItems([]);
        setSalaryAdvances([]);
      }
    }
  }, [workingDate, records, isAuthReady, user]);

  // --- Calculations ---
  const totals = useMemo(() => {
    const nDuDau = parseCurrency(duDau);
    const nRevenue = parseCurrency(revenue);
    const nCash = parseCurrency(cash);
    const nBank = parseCurrency(bank);
    const nReserve = parseCurrency(reserve);
    
    const nSpentCash = spentCashItems.reduce((sum, item) => sum + item.amount, 0);
    const nSpentBank = spentBankItems.reduce((sum, item) => sum + item.amount, 0);
    
    // Salaries and Advances from Section D
    const nSalaryCash = salaryAdvances.filter(a => a.method === 'salary_cash').reduce((sum, item) => sum + item.amount, 0);
    const nSalaryBank = salaryAdvances.filter(a => a.method === 'salary_bank').reduce((sum, item) => sum + item.amount, 0);
    const nAdvCash = salaryAdvances.filter(a => a.method === 'cash').reduce((sum, item) => sum + item.amount, 0);
    const nAdvBank = salaryAdvances.filter(a => a.method === 'bank').reduce((sum, item) => sum + item.amount, 0);
    const nAdvDebt = salaryAdvances.filter(a => a.method === 'debt').reduce((sum, item) => sum + item.amount, 0);
    const nAdvFine = salaryAdvances.filter(a => a.method === 'fine').reduce((sum, item) => sum + item.amount, 0);
    
    const totalSalaryAdvances = nAdvCash + nAdvBank + nAdvDebt + nAdvFine;
    
    const totalBook = nDuDau + nRevenue;
    const totalExpenses = nSpentCash + nSpentBank + totalSalaryAdvances + nSalaryCash + nSalaryBank;
    
    // Reconciliation logic:
    // Actual Assets (Cash + Bank) + Cash Out (Spent Cash + Adv Cash + Salary Cash)
    // should equal Book (Initial + Revenue)
    // Bank Out (Spent Bank + Adv Bank + Salary Bank) are from a separate account, so they don't affect this reconciliation.
    const totalActualAssets = nCash + nBank + nSpentCash + nAdvCash + nSalaryCash;
    const totalBookExpected = totalBook;
    
    const diff = totalActualAssets - totalBookExpected;
    const netProfit = nRevenue - totalExpenses;
    const recoveredMoney = nCash - nReserve;

    return { 
      totalBook, 
      totalExpenses, 
      totalActualAssets, 
      totalBookExpected,
      diff, 
      netProfit, 
      recoveredMoney,
      totalSalaryAdvances, 
      nSpentCash: nSpentCash + nSalaryCash, 
      nSpentBank: nSpentBank + nSalaryBank,
      nAdvCash,
      nAdvBank,
      nSalaryCash,
      nSalaryBank
    };
  }, [duDau, revenue, cash, bank, reserve, spentCashItems, spentBankItems, salaryAdvances]);

  // --- Handlers ---
  const addExpenseItem = (type: 'cash' | 'bank') => {
    const defaultCat = expenseCategories.length > 0 ? expenseCategories[0].name : "Khác";
    if (type === 'cash') {
      setSpentCashItems([...spentCashItems, { description: '', amount: 0, category: defaultCat }]);
    } else {
      setSpentBankItems([...spentBankItems, { description: '', amount: 0, category: defaultCat }]);
    }
  };

  const updateExpenseItem = (type: 'cash' | 'bank', index: number, field: keyof ExpenseItem, value: any) => {
    if (type === 'cash') {
      const newItems = [...spentCashItems];
      newItems[index] = { ...newItems[index], [field]: value };
      setSpentCashItems(newItems);
    } else {
      const newItems = [...spentBankItems];
      newItems[index] = { ...newItems[index], [field]: value };
      setSpentBankItems(newItems);
    }
  };

  const removeExpenseItem = (type: 'cash' | 'bank', index: number) => {
    if (type === 'cash') {
      setSpentCashItems(spentCashItems.filter((_, i) => i !== index));
    } else {
      setSpentBankItems(spentBankItems.filter((_, i) => i !== index));
    }
  };
  const addSalaryAdvance = () => {
    const defaultEmp = employees.length > 0 ? employees[0].name : "Nhân viên";
    setSalaryAdvances([...salaryAdvances, { employee: defaultEmp, method: 'salary_cash', amount: 0 }]);
  };

  const updateSalaryAdvance = (index: number, field: keyof SalaryAdvance, value: any) => {
    const newAdvances = [...salaryAdvances];
    newAdvances[index] = { ...newAdvances[index], [field]: value };
    setSalaryAdvances(newAdvances);
  };

  const removeSalaryAdvance = (index: number) => {
    setSalaryAdvances(salaryAdvances.filter((_, i) => i !== index));
  };

  const handleEditRecord = (record: RevenueRecord) => {
    setWorkingDate(record.date);
    setDuDau(record.duDau.toLocaleString('vi-VN'));
    setRevenue(record.revenue.toLocaleString('vi-VN'));
    setCash(record.cash.toLocaleString('vi-VN'));
    setBank(record.bank.toLocaleString('vi-VN'));
    setReserve(record.reserve.toLocaleString('vi-VN'));
    setSpentCashItems(record.spentCashItems || []);
    setSpentBankItems(record.spentBankItems || []);
    setSalaryAdvances(record.salaryAdvances || []);
    setView('calc');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const confirmDeleteRecord = async () => {
    if (!recordToDelete) return;
    try {
      await deleteDoc(doc(db, 'revenueRecords', recordToDelete));
      setRecordToDelete(null);
    } catch (err) {
      console.error("Delete failed:", err);
      alert("Không thể xóa báo cáo. Vui lòng thử lại.");
    }
  };

  const exportAsImage = async (recordId: string) => {
    const element = recordRefs.current[recordId];
    if (!element) return;

    setIsExporting(recordId);
    try {
      // Small delay to ensure UI updates if needed
      await new Promise(resolve => setTimeout(resolve, 100));
      
      const dataUrl = await toPng(element, {
        backgroundColor: '#1a1a1a',
        style: {
          borderRadius: '0',
        }
      });
      
      const link = document.createElement('a');
      link.download = `bao-cao-${recordId.slice(0, 5)}.png`;
      link.href = dataUrl;
      link.click();
    } catch (err) {
      console.error('Export failed', err);
    } finally {
      setIsExporting(null);
    }
  };

  const copyAsImage = async (recordId: string) => {
    const element = recordRefs.current[recordId];
    if (!element) return;

    setIsExporting(recordId);
    try {
      await new Promise(resolve => setTimeout(resolve, 100));
      
      const blob = await toBlob(element, {
        backgroundColor: '#1a1a1a',
        style: {
          borderRadius: '0',
        }
      });
      
      if (blob) {
        await navigator.clipboard.write([
          new ClipboardItem({
            [blob.type]: blob
          })
        ]);
        alert('Đã sao chép ảnh vào bộ nhớ tạm!');
      }
    } catch (err) {
      console.error('Copy failed', err);
      alert('Không thể sao chép ảnh. Vui lòng thử lại.');
    } finally {
      setIsExporting(null);
    }
  };

  const exportToExcel = async (filteredRecords: RevenueRecord[], periodLabel: string) => {
    const workbook = new ExcelJS.Workbook();
    
    // ==========================================
    // SHEET 1: TỔNG HỢP
    // ==========================================
    const wsSummary = workbook.addWorksheet('Tổng hợp');

    // Title
    wsSummary.mergeCells('A1:E1');
    const summaryTitle = wsSummary.getCell('A1');
    summaryTitle.value = `BÁO CÁO TỔNG HỢP - ${periodLabel.toUpperCase()}`;
    summaryTitle.font = { name: 'Arial', size: 18, bold: true, color: { argb: 'FFFFFFFF' } };
    summaryTitle.alignment = { vertical: 'middle', horizontal: 'center' };
    summaryTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };

    wsSummary.mergeCells('A2:E2');
    const summaryInfo = wsSummary.getCell('A2');
    summaryInfo.value = `Ngày xuất: ${new Date().toLocaleString('vi-VN')} | Chi nhánh: ${summaryBranchId === 'all' ? 'Tất cả' : branches.find(b => b.id === summaryBranchId)?.displayName}`;
    summaryInfo.font = { italic: true, size: 10 };
    summaryInfo.alignment = { horizontal: 'right' };

    // Overall Stats Table
    const totals = filteredRecords.reduce((acc, r) => ({
      duDau: acc.duDau + r.duDau,
      revenue: acc.revenue + r.revenue,
      spentCash: acc.spentCash + r.spentCash,
      spentBank: acc.spentBank + r.spentBank,
      salary: acc.salary + r.totalSalaryAdvances,
      salaryDebt: acc.salaryDebt + (r.salaryAdvances?.filter(a => a.method === 'debt').reduce((sum, a) => sum + a.amount, 0) || 0),
      salaryFine: acc.salaryFine + (r.salaryAdvances?.filter(a => a.method === 'fine').reduce((sum, a) => sum + a.amount, 0) || 0),
      profit: acc.profit + r.netProfit,
      recoveredMoney: acc.recoveredMoney + (r.recoveredMoney || 0),
      actual: acc.actual + r.totalActualAssets,
      book: acc.book + r.totalBookExpected,
      diff: acc.diff + r.diff
    }), { duDau: 0, revenue: 0, spentCash: 0, spentBank: 0, salary: 0, salaryDebt: 0, salaryFine: 0, profit: 0, recoveredMoney: 0, actual: 0, book: 0, diff: 0 });

    wsSummary.addRow([]);
    const statsHeader = wsSummary.addRow(['CHỈ SỐ TỔNG QUAN', 'GIÁ TRỊ']);
    statsHeader.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2D2D' } };
    });

    const statsData = [
      ['Tổng Doanh Thu', totals.revenue],
      ['Tổng Lợi Nhuận', totals.profit],
      ['Tổng Tiền Thu Hồi', totals.recoveredMoney],
      ['Tổng Chi Tiêu', totals.spentCash + totals.spentBank + totals.salary],
      ['   - Chi Tiền Mặt', totals.spentCash],
      ['   - Chi Chuyển Khoản', totals.spentBank],
      ['   - Ứng Lương', totals.salary - totals.salaryDebt - totals.salaryFine],
      ['   - Nợ Lương', totals.salaryDebt],
      ['   - Phạt', totals.salaryFine],
      ['Tổng Chênh Lệch Đối Soát', totals.diff],
      ['Số Ngày Ghi Chép', filteredRecords.length]
    ];

    statsData.forEach((item, idx) => {
      const row = wsSummary.addRow(item);
      row.getCell(2).numFmt = '#,##0';
      if (idx === 0) row.getCell(2).font = { bold: true, color: { argb: 'FF00B050' }, size: 12 };
      if (idx === 1) row.getCell(2).font = { bold: true, size: 12 };
      if (idx === 2) row.getCell(2).font = { bold: true, color: { argb: 'FF00B050' } };
      if (idx === 9) {
        const val = Number(row.getCell(2).value);
        row.getCell(2).font = { color: { argb: val >= 0 ? 'FF00B050' : 'FFFF0000' }, bold: true };
      }
    });

    // Branch Breakdown (if Admin & All Branches)
    if (summaryBranchId === 'all' && currentBranch?.role === 'admin') {
      wsSummary.addRow([]);
      wsSummary.addRow([]);
      const branchHeader = wsSummary.addRow(['CHI NHÁNH', 'DOANH THU', 'LỢI NHUẬN', 'CHI PHÍ', 'SỐ NGÀY']);
      branchHeader.eachCell(cell => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2D2D' } };
        cell.alignment = { horizontal: 'center' };
      });

      branches.filter(b => b.role !== 'admin').forEach(b => {
        const bRecords = filteredRecords.filter(r => r.branchId === b.id);
        if (bRecords.length === 0) return;
        const bRevenue = bRecords.reduce((sum, r) => sum + r.revenue, 0);
        const bProfit = bRecords.reduce((sum, r) => sum + r.netProfit, 0);
        const bExpenses = bRecords.reduce((sum, r) => sum + (r.spentCash + r.spentBank + r.totalSalaryAdvances), 0);
        
        const row = wsSummary.addRow([b.displayName, bRevenue, bProfit, bExpenses, bRecords.length]);
        row.getCell(2).numFmt = '#,##0';
        row.getCell(3).numFmt = '#,##0';
        row.getCell(4).numFmt = '#,##0';
        row.alignment = { horizontal: 'center' };
        row.getCell(1).alignment = { horizontal: 'left' };
      });
    }

    wsSummary.getColumn(1).width = 30;
    wsSummary.getColumn(2).width = 20;
    wsSummary.getColumn(3).width = 20;
    wsSummary.getColumn(4).width = 20;
    wsSummary.getColumn(5).width = 15;

    // ==========================================
    // SHEET 2: CHI TIẾT
    // ==========================================
    const wsDetails = workbook.addWorksheet('Chi tiết');

    // Title & Info
    wsDetails.mergeCells('A1:N1');
    const detailsTitle = wsDetails.getCell('A1');
    detailsTitle.value = `CHI TIẾT GIAO DỊCH - ${periodLabel.toUpperCase()}`;
    detailsTitle.font = { name: 'Arial', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
    detailsTitle.alignment = { vertical: 'middle', horizontal: 'center' };
    detailsTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2D2D' } };

    wsDetails.mergeCells('A2:N2');
    const detailsInfo = wsDetails.getCell('A2');
    detailsInfo.value = `Ngày xuất: ${new Date().toLocaleString('vi-VN')}`;
    detailsInfo.font = { italic: true, size: 10 };
    detailsInfo.alignment = { horizontal: 'right' };

    // Headers
    const headers = [
      'Ngày', 'Chi nhánh', 'Dư đầu', 'Doanh thu', 'Tiền thu hồi', 'Chi TM', 'Chi tiết Chi TM', 'Chi CK', 'Chi tiết Chi CK', 'Ứng lương', 'Chi tiết Ứng', 'Lợi nhuận', 'Thực có', 'Sổ sách', 'Chênh lệch'
    ];
    const headerRow = wsDetails.addRow(headers);
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF444444' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // Data Rows
    filteredRecords.forEach((r, index) => {
      const branch = branches.find(b => b.id === r.branchId);
      const cashDetails = r.spentCashItems?.map(item => `${item.description} / ${item.category} / ${formatCurrency(item.amount)}`).join('\n') || '';
      const bankDetails = r.spentBankItems?.map(item => `${item.description} / ${item.category} / ${formatCurrency(item.amount)}`).join('\n') || '';
      const salaryDetails = r.salaryAdvances?.map(item => {
        const methodLabels: Record<string, string> = {
          salary_cash: 'Lương TM',
          salary_bank: 'Lương CK',
          cash: 'Ứng TM',
          bank: 'Ứng CK',
          debt: 'Nợ',
          fine: 'Phạt'
        };
        return `${item.employee} (${methodLabels[item.method] || item.method}): ${formatCurrency(item.amount)}`;
      }).join('\n') || '';

      const rowData = [
        r.date,
        branch?.displayName || r.branchId,
        r.duDau,
        r.revenue,
        r.recoveredMoney || (r.cash - r.reserve),
        r.spentCash,
        cashDetails,
        r.spentBank,
        bankDetails,
        r.totalSalaryAdvances,
        salaryDetails,
        r.netProfit,
        r.totalActualAssets,
        r.totalBookExpected,
        r.diff
      ];
      const row = wsDetails.addRow(rowData);
      
      if (index % 2 === 1) {
        row.eachCell((cell) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF9F9F9' } };
        });
      }

      row.eachCell((cell, colNumber) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
          right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
        };
        
        const numericCols = [3, 4, 5, 6, 8, 10, 12, 13, 14, 15];
        if (numericCols.includes(colNumber)) {
          cell.numFmt = '#,##0';
          cell.alignment = { horizontal: 'right', vertical: 'top' };
        } else if ([7, 9, 11].includes(colNumber)) {
          cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
          cell.font = { size: 9 };
        } else {
          cell.alignment = { horizontal: 'center', vertical: 'top' };
        }

        if (colNumber === 12 || colNumber === 15) {
          const val = Number(cell.value);
          if (val > 0) cell.font = { color: { argb: 'FF00B050' }, bold: true };
          else if (val < 0) cell.font = { color: { argb: 'FFFF0000' }, bold: true };
        }
      });
    });

    // Totals Row on Details Sheet
    const totalRow = wsDetails.addRow([
      'TỔNG CỘNG', '', totals.duDau, totals.revenue, totals.spentCash, '', totals.spentBank, '', totals.salary, '', totals.profit, totals.actual, totals.book, totals.diff
    ]);
    totalRow.eachCell((cell, colNumber) => {
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
      cell.border = { top: { style: 'medium' } };
      const numericCols = [3, 4, 5, 7, 9, 11, 12, 13, 14];
      if (numericCols.includes(colNumber)) {
        cell.numFmt = '#,##0';
        cell.alignment = { horizontal: 'right' };
      }
    });
    wsDetails.mergeCells(`A${totalRow.number}:B${totalRow.number}`);

    // Column Widths for Details
    wsDetails.columns.forEach((column, i) => {
      const colIndex = i + 1;
      if ([6, 8, 10].includes(colIndex)) {
        column.width = 35;
      } else {
        let maxLength = 0;
        column.eachCell!({ includeEmpty: true }, (cell) => {
          const columnLength = cell.value ? cell.value.toString().length : 10;
          if (columnLength > maxLength) maxLength = columnLength;
        });
        column.width = maxLength < 12 ? 12 : (maxLength > 25 ? 25 : maxLength + 2);
      }
    });

    // ==========================================
    // SHEET 3: PHÂN TÍCH CHI PHÍ
    // ==========================================
    const wsExpenses = workbook.addWorksheet('Phân tích Chi phí');
    wsExpenses.addRow(['BÁO CÁO CHI TIẾT CHI PHÍ THEO HẠNG MỤC']).font = { bold: true, size: 14 };
    wsExpenses.addRow([`Giai đoạn: ${periodLabel}`]);
    wsExpenses.addRow([]);

    const expenseHeader = wsExpenses.addRow(['Hạng mục', 'Số tiền', 'Tỷ lệ (%)']);
    expenseHeader.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2D2D' } };
    });

    const categoryTotals: { [key: string]: number } = {};
    let totalAllExpenses = 0;

    filteredRecords.forEach(r => {
      [...(r.spentCashItems || []), ...(r.spentBankItems || [])].forEach(item => {
        categoryTotals[item.category] = (categoryTotals[item.category] || 0) + item.amount;
        totalAllExpenses += item.amount;
      });
    });

    Object.entries(categoryTotals)
      .sort((a, b) => b[1] - a[1])
      .forEach(([cat, amount]) => {
        const percentage = totalAllExpenses > 0 ? (amount / totalAllExpenses) * 100 : 0;
        const row = wsExpenses.addRow([cat, amount, percentage / 100]);
        row.getCell(2).numFmt = '#,##0';
        row.getCell(3).numFmt = '0.00%';
      });

    wsExpenses.addRow([]);
    const expTotalRow = wsExpenses.addRow(['TỔNG CỘNG', totalAllExpenses, 1]);
    expTotalRow.font = { bold: true };
    expTotalRow.getCell(2).numFmt = '#,##0';
    expTotalRow.getCell(3).numFmt = '0.00%';

    wsExpenses.getColumn(1).width = 30;
    wsExpenses.getColumn(2).width = 20;
    wsExpenses.getColumn(3).width = 15;

    // ==========================================
    // SHEET 4: NHÂN SỰ & LƯƠNG
    // ==========================================
    const wsPersonnel = workbook.addWorksheet('Nhân sự & Lương');
    wsPersonnel.addRow(['BÁO CÁO CHI TIẾT ỨNG LƯƠNG NHÂN VIÊN']).font = { bold: true, size: 14 };
    wsPersonnel.addRow([`Giai đoạn: ${periodLabel}`]);
    wsPersonnel.addRow([]);

    const personnelHeader = wsPersonnel.addRow(['Nhân viên', 'Tổng ứng TM', 'Tổng ứng CK', 'Nợ lương', 'Phạt', 'Tổng cộng']);
    personnelHeader.eachCell(cell => {
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2D2D' } };
    });

    const employeeStats: { [key: string]: { cash: number, bank: number, debt: number, fine: number } } = {};
    filteredRecords.forEach(r => {
      (r.salaryAdvances || []).forEach(adv => {
        if (!employeeStats[adv.employee]) employeeStats[adv.employee] = { cash: 0, bank: 0, debt: 0, fine: 0 };
        if (adv.method === 'cash' || adv.method === 'salary_cash') employeeStats[adv.employee].cash += adv.amount;
        else if (adv.method === 'bank' || adv.method === 'salary_bank') employeeStats[adv.employee].bank += adv.amount;
        else if (adv.method === 'debt') employeeStats[adv.employee].debt += adv.amount;
        else if (adv.method === 'fine') employeeStats[adv.employee].fine += adv.amount;
      });
    });

    Object.entries(employeeStats)
      .sort((a, b) => (b[1].cash + b[1].bank + b[1].debt + b[1].fine) - (a[1].cash + a[1].bank + a[1].debt + a[1].fine))
      .forEach(([name, stats]) => {
        const row = wsPersonnel.addRow([name, stats.cash, stats.bank, stats.debt, stats.fine, stats.cash + stats.bank + stats.debt + stats.fine]);
        row.getCell(2).numFmt = '#,##0';
        row.getCell(3).numFmt = '#,##0';
        row.getCell(4).numFmt = '#,##0';
        row.getCell(5).numFmt = '#,##0';
        row.getCell(6).numFmt = '#,##0';
      });

    wsPersonnel.getColumn(1).width = 30;
    wsPersonnel.getColumn(2).width = 15;
    wsPersonnel.getColumn(3).width = 15;
    wsPersonnel.getColumn(4).width = 15;
    wsPersonnel.getColumn(5).width = 15;
    wsPersonnel.getColumn(6).width = 15;

    // --- Save File ---
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    saveAs(blob, `Bao_cao_${periodLabel.replace(/\//g, '-')}_${new Date().getTime()}.xlsx`);
  };

  const generateReportText = () => {
    const cashDetails = spentCashItems.length > 0 
      ? spentCashItems.map(item => `     + ${item.description}: ${formatCurrency(item.amount)}`).join('\n')
      : '';
    const bankDetails = spentBankItems.length > 0 
      ? spentBankItems.map(item => `     + ${item.description}: ${formatCurrency(item.amount)}`).join('\n')
      : '';
    const salaryDetails = salaryAdvances.length > 0
      ? salaryAdvances.map(adv => {
          const methodLabels: Record<string, string> = {
            salary_cash: 'Lương TM',
            salary_bank: 'Lương CK',
            cash: 'Ứng TM',
            bank: 'Ứng CK',
            debt: 'Nợ',
            fine: 'Phạt'
          };
          return `     + ${adv.employee}: ${methodLabels[adv.method]}: ${formatCurrency(adv.amount)}`;
        }).join('\n')
      : '';

    return `
<b>📊 BÁO CÁO THU CHI - ${new Date(workingDate).toLocaleDateString('vi-VN')}</b>
---------------------------
💰 DOANH THU: ${formatCurrency(parseCurrency(revenue))}
💸 TIỀN THU HỒI: ${formatCurrency(totals.recoveredMoney)}
📉 TỔNG CHI: ${formatCurrency(totals.totalExpenses)}
   - Chi TM: ${formatCurrency(totals.nSpentCash)}
${cashDetails}
   - Chi CK: ${formatCurrency(totals.nSpentBank)}
${bankDetails}
   - Ứng lương: ${formatCurrency(totals.totalSalaryAdvances)}
${salaryDetails}
✨ LỢI NHUẬN: ${formatCurrency(totals.netProfit)}
---------------------------
🔍 ĐỐI SOÁT:
- Sổ sách (Dư + Thu - Chi CK): ${formatCurrency(totals.totalBookExpected)}
- Thực có (Tài sản + Chi TM): ${formatCurrency(totals.totalActualAssets)}
⚖️ CHÊNH LỆCH: <b>${totals.diff >= 0 ? '+' : ''}${formatCurrency(totals.diff)}</b>
${totals.diff >= 0 ? '✅ KHỚP/THỪA' : '❌ THIẾU TIỀN'}
    `.trim().replace(/\n\n+/g, '\n');
  };

  const handleSave = async () => {
    if (!user) return;
    setSaving(true);
    setError(null);

    const path = 'revenueRecords';
    try {
      const reportText = generateReportText();
      const newRecord: Omit<RevenueRecord, 'id'> = {
        date: workingDate,
        duDau: parseCurrency(duDau),
        revenue: parseCurrency(revenue),
        cash: parseCurrency(cash),
        bank: parseCurrency(bank),
        reserve: parseCurrency(reserve),
        spentCash: totals.nSpentCash,
        spentBank: totals.nSpentBank,
        spentCashItems: spentCashItems,
        spentBankItems: spentBankItems,
        salaryAdvances: salaryAdvances,
        totalSalaryAdvances: totals.totalSalaryAdvances,
        totalBook: totals.totalBook,
        totalBookExpected: totals.totalBookExpected,
        totalExpenses: totals.totalExpenses,
        totalActualAssets: totals.totalActualAssets,
        recoveredMoney: totals.recoveredMoney,
        diff: totals.diff,
        netProfit: totals.netProfit,
        branchId: currentBranch!.id,
        uid: user!.uid,
        createdAt: new Date().toISOString()
      };

      const existingRecord = records.find(r => r.date === workingDate);
      if (existingRecord?.id) {
        await setDoc(doc(db, path, existingRecord.id), newRecord);
        setSaveSuccess(true);
      } else {
        await addDoc(collection(db, path), newRecord);
        setSaveSuccess(true);
      }
      
      // Clear form
      setDuDau('');
      setRevenue('');
      setCash('');
      setBank('');
      setReserve('');
      setSpentCashItems([]);
      setSpentBankItems([]);
      setSalaryAdvances([]);
      
      setTimeout(() => setSaveSuccess(false), 3000);
      setView('history');
    } catch (err) {
      setError('Không thể lưu dữ liệu. Vui lòng thử lại.');
      const existingRecord = records.find(r => r.date === workingDate);
      handleFirestoreError(err, existingRecord?.id ? OperationType.UPDATE : OperationType.CREATE, path);
    } finally {
      setSaving(false);
    }
  };

  const copyReport = () => {
    const report = generateReportText().replace(/<[^>]*>/g, ''); // Strip HTML tags for clipboard

    navigator.clipboard.writeText(report).then(() => {
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    });
  };

  if (!isAuthReady) {
    return (
      <div className="min-h-screen bg-[#1a1a1a] flex items-center justify-center">
        <div className="animate-spin rounded-full h-12 w-12 border-t-2 border-b-2 border-[#00ff88]"></div>
      </div>
    );
  }

  if (!currentBranch) {
    return (
      <div className="min-h-screen bg-[#111] flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="w-full max-w-md bg-[#1a1a1a] p-8 rounded-3xl border border-white/5 shadow-2xl"
        >
          <div className="flex flex-col items-center mb-8">
            <div className="bg-[#00ff88] p-4 rounded-2xl mb-4 shadow-lg shadow-[#00ff88]/20">
              <Building2 className="w-8 h-8 text-black" />
            </div>
            <h1 className="text-2xl font-bold text-white tracking-tight">Hệ Thống Thôn Mây</h1>
            <p className="text-gray-500 text-sm mt-1">Nhập tên chi nhánh để tiếp tục</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-6">
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-gray-500 uppercase ml-1 tracking-widest">Tên chi nhánh</label>
              <input 
                type="text"
                value={loginBranch}
                onChange={(e) => setLoginBranch(e.target.value)}
                placeholder="vd: thonmay"
                required
                className="w-full bg-black/40 border border-white/5 p-5 rounded-2xl text-white text-lg font-bold focus:outline-none focus:border-[#00ff88] focus:ring-4 focus:ring-[#00ff88]/10 transition-all placeholder:text-gray-700"
              />
              <p className="text-[10px] text-gray-600 ml-1 italic">* Nhập mã chi nhánh đã được cấp</p>
            </div>

            {error && (
              <div className="bg-red-500/10 border border-red-500/20 p-4 rounded-xl flex items-center gap-3 text-red-400 text-sm">
                <AlertCircle className="w-5 h-5 flex-shrink-0" />
                {error}
              </div>
            )}

            <button 
              type="submit"
              disabled={isLoggingIn}
              className="w-full bg-[#00ff88] text-black font-black uppercase tracking-widest py-5 rounded-2xl shadow-lg shadow-[#00ff88]/20 hover:scale-[1.02] active:scale-[0.98] transition-all disabled:opacity-50 flex items-center justify-center gap-3"
            >
              {isLoggingIn ? (
                <div className="animate-spin rounded-full h-5 w-5 border-t-2 border-b-2 border-black"></div>
              ) : (
                <>
                  <ArrowRightLeft className="w-5 h-5" />
                  Truy cập dữ liệu
                </>
              )}
            </button>

            <div className="relative my-8">
              <div className="absolute inset-0 flex items-center">
                <div className="w-full border-t border-white/5"></div>
              </div>
              <div className="relative flex justify-center text-[10px] uppercase font-bold">
                <span className="bg-[#1a1a1a] px-4 text-gray-600 tracking-widest">Quản trị viên</span>
              </div>
            </div>

            <button 
              type="button"
              onClick={async () => {
                try {
                  const { signInWithGoogle } = await import('./firebase');
                  await signInWithGoogle();
                } catch (err: any) {
                  setError('Lỗi khi đăng nhập bằng Google.');
                }
              }}
              className="w-full bg-white/5 text-white font-bold py-4 rounded-2xl border border-white/10 hover:bg-white/10 transition-all flex items-center justify-center gap-2 group"
            >
              <ShieldCheck className="w-5 h-5 text-[#00ff88] group-hover:scale-110 transition-transform" />
              Đăng nhập Admin (Google)
            </button>
          </form>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#111] text-white font-sans selection:bg-[#00ff88]/30">
      {/* Header */}
      <header className="sticky top-0 z-50 bg-[#1a1a1a]/80 backdrop-blur-md border-b border-white/5 px-4 py-4">
        <div className="max-w-xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-[#00ff88] p-1.5 rounded-lg">
              <Calculator className="w-5 h-5 text-black" />
            </div>
            <h1 className="font-bold text-lg tracking-tight uppercase">{currentBranch.displayName}</h1>
          </div>
          <div className="flex items-center gap-3">
            <button 
              onClick={() => setView('admin')}
              className={`p-2 rounded-full transition-colors ${view === 'admin' ? 'bg-[#00ff88] text-black' : 'hover:bg-white/5 text-gray-400'}`}
            >
              <Settings className="w-5 h-5" />
            </button>
            <button 
              onClick={() => {
                if (view === 'calc') setView('history');
                else if (view === 'history') setView('summary');
                else setView('calc');
              }}
              className="p-2 hover:bg-white/5 rounded-full transition-colors flex items-center gap-2"
            >
              {view === 'calc' && <History className="w-5 h-5 text-gray-400" />}
              {view === 'history' && <TrendingUp className="w-5 h-5 text-[#00ff88]" />}
              {view === 'summary' && <Calculator className="w-5 h-5 text-gray-400" />}
            </button>
            <button 
              onClick={handleLogout}
              className="p-2 hover:bg-white/5 rounded-full transition-colors"
            >
              <LogOut className="w-5 h-5 text-gray-400" />
            </button>
            <div className="w-8 h-8 rounded-full bg-white/5 border border-white/10 flex items-center justify-center text-[10px] font-bold text-[#00ff88]">
              {currentBranch.id.slice(0, 2).toUpperCase()}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-xl mx-auto p-4 pb-32">
        <AnimatePresence mode="wait">
          {view === 'admin' ? (
            <motion.div
              key="admin"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="space-y-6"
            >
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <button 
                    onClick={() => setView('calc')}
                    className="p-2 hover:bg-white/10 rounded-full text-gray-400 transition-colors"
                  >
                    <ChevronRight className="w-6 h-6 rotate-180" />
                  </button>
                  <h2 className="text-xl font-bold flex items-center gap-2">
                    <Settings className="w-6 h-6 text-[#00ff88]" />
                    Cài đặt {currentBranch.role === 'admin' ? 'hệ thống' : 'chi nhánh'}
                  </h2>
                </div>
                {currentBranch.role === 'admin' && (
                  <button 
                    onClick={() => {
                      setSummaryBranchId('all');
                      setView('summary');
                    }}
                    className="text-xs font-bold text-[#00ff88] bg-[#00ff88]/10 px-4 py-2 rounded-xl"
                  >
                    Xem doanh số tổng
                  </button>
                )}
              </div>

              {/* Branch Settings Form */}
              <div className="bg-[#2d2d2d] p-6 rounded-3xl border border-white/5 shadow-xl">
                <h3 className="text-sm font-bold text-gray-400 uppercase tracking-widest mb-6 flex items-center gap-2">
                  <Building2 className="w-4 h-4" /> {currentBranch.role === 'admin' ? (editingBranchId ? 'Cập nhật chi nhánh' : 'Tạo chi nhánh mới') : 'Thông tin chi nhánh'}
                </h3>
                <form onSubmit={createBranch} className="space-y-4">
                  <div className="grid grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-500 uppercase ml-1">Mã chi nhánh (ID)</label>
                      <input 
                        type="text"
                        value={newBranchId}
                        onChange={(e) => setNewBranchId(e.target.value)}
                        placeholder="vd: thonmay2"
                        required
                        disabled={true}
                        className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88] disabled:opacity-50"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-500 uppercase ml-1">Tên hiển thị</label>
                      <input 
                        type="text"
                        value={newBranchName}
                        onChange={(e) => setNewBranchName(e.target.value)}
                        placeholder="vd: Thôn Mây 2"
                        required
                        className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                      />
                    </div>
                  </div>
                  {currentBranch.role === 'admin' && (
                    <div className="grid grid-cols-2 gap-4">
                      <div className="space-y-1">
                        <label className="text-[10px] text-gray-500 uppercase ml-1">Tài khoản</label>
                        <input 
                          type="text"
                          value={newBranchUser}
                          onChange={(e) => setNewBranchUser(e.target.value)}
                          placeholder="Số điện thoại"
                          className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                        />
                      </div>
                      <div className="space-y-1">
                        <label className="text-[10px] text-gray-500 uppercase ml-1">Mật khẩu</label>
                        <input 
                          type="text"
                          value={newBranchPass}
                          onChange={(e) => setNewBranchPass(e.target.value)}
                          placeholder="Mật khẩu"
                          className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                        />
                      </div>
                    </div>
                  )}
                  <div className="space-y-1">
                    <label className="text-[10px] text-gray-500 uppercase ml-1">Địa chỉ (Nâng cao)</label>
                    <input 
                      type="text"
                      value={newBranchAddr}
                      onChange={(e) => setNewBranchAddr(e.target.value)}
                      placeholder="Địa chỉ chi nhánh"
                      className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-[10px] text-gray-500 uppercase ml-1">Số điện thoại (Nâng cao)</label>
                    <input 
                      type="text"
                      value={newBranchPhone}
                      onChange={(e) => setNewBranchPhone(e.target.value)}
                      placeholder="Số điện thoại liên hệ"
                      className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                    />
                  </div>
                  <div className="flex gap-2">
                    {editingBranchId && (
                      <button 
                        type="button"
                        onClick={() => {
                          setEditingBranchId(null);
                          setNewBranchId('');
                          setNewBranchName('');
                          setNewBranchUser('');
                          setNewBranchPass('');
                          setNewBranchAddr('');
                          setNewBranchPhone('');
                        }}
                        className="flex-1 bg-white/5 text-white font-bold py-3 rounded-xl border border-white/10"
                      >
                        Hủy
                      </button>
                    )}
                    <button 
                      type="submit"
                      disabled={saving}
                      className="flex-[2] bg-[#00ff88] text-black font-bold py-3 rounded-xl shadow-lg shadow-[#00ff88]/10 disabled:opacity-50"
                    >
                      {saving ? 'Đang lưu...' : (editingBranchId ? 'Cập nhật' : 'Xác nhận tạo chi nhánh')}
                    </button>
                  </div>
                </form>
              </div>

              {/* Branches List (Admin Only) */}
              {currentBranch.role === 'admin' && (
                <div className="space-y-4">
                  <h3 className="text-sm font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                    <Building2 className="w-4 h-4" /> Danh sách chi nhánh ({branches.length})
                  </h3>
                  <div className="grid grid-cols-1 gap-3">
                    {branches.map(b => (
                      <div key={b.id} className="bg-[#2d2d2d] p-4 rounded-2xl border border-white/5 flex items-center justify-between">
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-white/5 flex items-center justify-center text-[#00ff88] font-bold">
                            {b.id.slice(0, 2).toUpperCase()}
                          </div>
                          <div>
                            <div className="font-bold text-sm">{b.displayName}</div>
                            <div className="text-[10px] text-gray-500 font-mono">ID: {b.id} | User: {b.username}</div>
                          </div>
                        </div>
                        <div className="flex items-center gap-2">
                          <div className="flex flex-col items-end gap-1">
                            <span className={`text-[8px] px-2 py-0.5 rounded-full font-bold uppercase ${b.role === 'admin' ? 'bg-purple-500/20 text-purple-400' : 'bg-[#00ff88]/20 text-[#00ff88]'}`}>
                              {b.role}
                            </span>
                            {b.phone && <div className="text-[10px] text-gray-500 flex items-center gap-1"><Phone className="w-2 h-2" /> {b.phone}</div>}
                          </div>
                          <div className="flex gap-1">
                            <button 
                              onClick={() => startEditBranch(b)}
                              className="p-2 hover:bg-white/10 rounded-lg text-blue-400 transition-colors"
                            >
                              <Edit2 className="w-4 h-4" />
                            </button>
                            {b.role !== 'admin' && (
                              <button 
                                onClick={() => deleteBranch(b.id)}
                                className="p-2 hover:bg-white/10 rounded-lg text-red-400 transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Employee Management Section */}
              <div className="bg-[#2d2d2d] p-6 rounded-3xl border border-white/5 shadow-xl">
                <h2 className="text-sm font-bold text-[#00ff88] uppercase tracking-widest mb-6 flex items-center gap-2">
                  D. Quản lý nhân viên
                </h2>
                <h3 className="text-[10px] font-bold text-gray-500 uppercase tracking-widest mb-4 flex items-center gap-2">
                  <UserIcon className="w-4 h-4" /> {editingEmployeeId ? 'Cập nhật nhân viên' : 'Thêm nhân viên mới'}
                </h3>
                <form onSubmit={createEmployee} className="space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-500 uppercase ml-1">Tên nhân viên</label>
                      <input 
                        type="text"
                        value={newEmployeeName}
                        onChange={(e) => setNewEmployeeName(e.target.value)}
                        placeholder="vd: Nguyễn Văn A"
                        required
                        className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-500 uppercase ml-1">Chi nhánh</label>
                      <select
                        value={currentBranch.role === 'admin' ? newEmployeeBranch : currentBranch.id}
                        onChange={(e) => setNewEmployeeBranch(e.target.value)}
                        disabled={currentBranch.role !== 'admin'}
                        className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88] disabled:opacity-50"
                      >
                        {currentBranch.role === 'admin' && <option value="all">Tất cả chi nhánh</option>}
                        {branches.map(b => (
                          <option key={b.id} value={b.id}>{b.displayName}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    {editingEmployeeId && (
                      <button 
                        type="button"
                        onClick={() => {
                          setEditingEmployeeId(null);
                          setNewEmployeeName('');
                          setNewEmployeeBranch('all');
                        }}
                        className="flex-1 bg-white/5 text-white font-bold py-3 rounded-xl border border-white/10"
                      >
                        Hủy
                      </button>
                    )}
                    <button 
                      type="submit"
                      disabled={saving}
                      className="flex-[2] bg-[#00ff88] text-black font-bold py-3 rounded-xl shadow-lg shadow-[#00ff88]/10 disabled:opacity-50"
                    >
                      {saving ? 'Đang lưu...' : (editingEmployeeId ? 'Cập nhật' : 'Thêm nhân viên')}
                    </button>
                  </div>
                </form>
              </div>

              {/* Expense Category Management Section */}
              <div className="bg-[#2d2d2d] p-6 rounded-3xl border border-white/5 shadow-xl">
                <h3 className="text-sm font-bold text-gray-400 uppercase tracking-widest mb-6 flex items-center gap-2">
                  <Settings className="w-4 h-4" /> {editingCategoryId ? 'Cập nhật phân loại' : 'Thêm phân loại chi mới'}
                </h3>
                <form onSubmit={createCategory} className="space-y-4">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-500 uppercase ml-1">Tên phân loại</label>
                      <input 
                        type="text"
                        value={newCategoryName}
                        onChange={(e) => setNewCategoryName(e.target.value)}
                        placeholder="vd: Trái cây"
                        required
                        className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88]"
                      />
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-500 uppercase ml-1">Chi nhánh</label>
                      <select
                        value={currentBranch.role === 'admin' ? newCategoryBranch : currentBranch.id}
                        onChange={(e) => setNewCategoryBranch(e.target.value)}
                        disabled={currentBranch.role !== 'admin'}
                        className="w-full bg-black/20 border border-white/5 p-3 rounded-xl text-sm focus:outline-none focus:border-[#00ff88] disabled:opacity-50"
                      >
                        {currentBranch.role === 'admin' && <option value="all">Tất cả chi nhánh</option>}
                        {branches.map(b => (
                          <option key={b.id} value={b.id}>{b.displayName}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <div className="flex gap-2">
                    {editingCategoryId && (
                      <button 
                        type="button"
                        onClick={() => {
                          setEditingCategoryId(null);
                          setNewCategoryName('');
                          setNewCategoryBranch('all');
                        }}
                        className="flex-1 bg-white/5 text-white font-bold py-3 rounded-xl border border-white/10"
                      >
                        Hủy
                      </button>
                    )}
                    <button 
                      type="submit"
                      disabled={saving}
                      className="flex-[2] bg-[#00ff88] text-black font-bold py-3 rounded-xl shadow-lg shadow-[#00ff88]/10 disabled:opacity-50"
                    >
                      {saving ? 'Đang lưu...' : (editingCategoryId ? 'Cập nhật' : 'Thêm phân loại')}
                    </button>
                  </div>
                </form>
              </div>

              {/* Expense Categories List */}
              <div className="space-y-4">
                <h3 className="text-sm font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <Settings className="w-4 h-4" /> Danh sách phân loại ({expenseCategories.length})
                </h3>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {expenseCategories.map(cat => (
                    <div key={cat.id} className="bg-[#2d2d2d] p-3 rounded-2xl border border-white/5 flex items-center justify-between">
                      <div className="flex flex-col">
                        <div className="font-bold text-xs">{cat.name}</div>
                        <div className="text-[8px] text-gray-500">
                          {cat.branchId === 'all' ? 'Tất cả' : (branches.find(b => b.id === cat.branchId)?.displayName || cat.branchId)}
                        </div>
                      </div>
                      <div className="flex gap-1">
                        <button 
                          onClick={() => startEditCategory(cat)}
                          className="p-1 hover:bg-white/10 rounded-lg text-blue-400 transition-colors"
                        >
                          <Edit2 className="w-3 h-3" />
                        </button>
                        <button 
                          onClick={() => deleteCategory(cat.id)}
                          className="p-1 hover:bg-white/10 rounded-lg text-red-400 transition-colors"
                        >
                          <Trash2 className="w-3 h-3" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Employees List */}
              <div className="space-y-4">
                <h3 className="text-sm font-bold text-gray-400 uppercase tracking-widest flex items-center gap-2">
                  <UserIcon className="w-4 h-4" /> Danh sách nhân viên ({employees.length})
                </h3>
                <div className="grid grid-cols-1 gap-3">
                  {employees.map(emp => (
                    <div key={emp.id} className="bg-[#2d2d2d] p-4 rounded-2xl border border-white/5 flex items-center justify-between">
                      <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-xl bg-white/5 flex items-center justify-center text-[#00ff88] font-bold">
                          {emp.name.slice(0, 1).toUpperCase()}
                        </div>
                        <div>
                          <div className="font-bold text-sm">{emp.name}</div>
                          <div className="text-[10px] text-gray-500">
                            Chi nhánh: {emp.branchId === 'all' ? 'Tất cả' : (branches.find(b => b.id === emp.branchId)?.displayName || emp.branchId)}
                          </div>
                        </div>
                      </div>
                      <div className="flex gap-1">
                        <button 
                          onClick={() => startEditEmployee(emp)}
                          className="p-2 hover:bg-white/10 rounded-lg text-blue-400 transition-colors"
                        >
                          <Edit2 className="w-4 h-4" />
                        </button>
                        <button 
                          onClick={() => deleteEmployee(emp.id)}
                          className="p-2 hover:bg-white/10 rounded-lg text-red-400 transition-colors"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </motion.div>
          ) : view === 'summary' ? (
            <motion.div
              key="summary"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-6"
            >
              <div className="space-y-4 mb-6">
                <div className="flex items-center justify-between">
                  <h2 className="text-xl font-bold flex items-center gap-2">
                    <TrendingUp className="w-6 h-6 text-[#00ff88]" />
                    Tổng hợp doanh thu
                  </h2>
                  <div className="flex items-center gap-2">
                    <button
                      onClick={() => {
                        let label = '';
                        if (summaryPeriod === 'day') label = summaryDate;
                        else if (summaryPeriod === 'month') label = summaryMonth;
                        else label = summaryYear;
                        
                        let filtered = records;
                        if (summaryPeriod === 'day') filtered = records.filter(r => r.date === summaryDate);
                        else if (summaryPeriod === 'month') filtered = records.filter(r => r.date.startsWith(summaryMonth));
                        else filtered = records.filter(r => r.date.startsWith(summaryYear));
                        
                        if (summaryBranchId !== 'all') filtered = filtered.filter(r => r.branchId === summaryBranchId);
                        
                        exportToExcel(filtered, label);
                      }}
                      className="p-2 bg-[#00ff88]/10 text-[#00ff88] rounded-xl hover:bg-[#00ff88]/20 transition-all flex items-center gap-2 text-xs font-bold"
                    >
                      <Download className="w-4 h-4" />
                      Xuất Excel
                    </button>
                    <div className="flex gap-1 bg-white/5 p-1 rounded-xl">
                      {(['day', 'month', 'year'] as const).map(p => (
                        <button
                          key={p}
                          onClick={() => setSummaryPeriod(p)}
                          className={`px-3 py-1 rounded-lg text-[10px] font-bold uppercase transition-all ${summaryPeriod === p ? 'bg-[#00ff88] text-black' : 'text-gray-500 hover:text-white'}`}
                        >
                          {p === 'day' ? 'Ngày' : p === 'month' ? 'Tháng' : 'Năm'}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-2">
                  {summaryPeriod === 'day' && (
                    <input 
                      type="date" 
                      value={summaryDate}
                      onChange={(e) => setSummaryDate(e.target.value)}
                      onClick={(e) => {
                        if (typeof e.currentTarget.showPicker === 'function') {
                          try {
                            e.currentTarget.showPicker();
                          } catch (err) {
                            console.warn("showPicker failed", err);
                          }
                        }
                      }}
                      className="bg-[#2d2d2d] border border-white/10 rounded-xl px-3 py-2 text-sm focus:outline-none focus:border-[#00ff88] w-full cursor-pointer"
                    />
                  )}
                  {summaryPeriod === 'month' && (
                    <input 
                      type="month" 
                      value={summaryMonth}
                      onChange={(e) => setSummaryMonth(e.target.value)}
                      onClick={(e) => {
                        if (typeof e.currentTarget.showPicker === 'function') {
                          try {
                            e.currentTarget.showPicker();
                          } catch (err) {
                            console.warn("showPicker failed", err);
                          }
                        }
                      }}
                      className="bg-[#2d2d2d] border border-white/10 rounded-xl px-3 py-2 text-sm focus:outline-none focus:border-[#00ff88] w-full cursor-pointer"
                    />
                  )}
                  {summaryPeriod === 'year' && (
                    <select
                      value={summaryYear}
                      onChange={(e) => setSummaryYear(e.target.value)}
                      className="bg-[#2d2d2d] border border-white/10 rounded-xl px-3 py-2 text-sm focus:outline-none focus:border-[#00ff88] w-full"
                    >
                      {[2024, 2025, 2026, 2027].map(y => (
                        <option key={y} value={y.toString()}>{y}</option>
                      ))}
                    </select>
                  )}
                  
                  {currentBranch.role === 'admin' && (
                    <select
                      value={summaryBranchId}
                      onChange={(e) => setSummaryBranchId(e.target.value)}
                      className="bg-[#2d2d2d] border border-white/10 rounded-xl px-3 py-2 text-sm focus:outline-none focus:border-[#00ff88] w-full"
                    >
                      <option value="all">Tất cả chi nhánh</option>
                      {branches.map(b => (
                        <option key={b.id} value={b.id}>{b.displayName}</option>
                      ))}
                    </select>
                  )}
                </div>
              </div>

              {(() => {
                let filteredRecords = records;
                
                if (summaryPeriod === 'day') {
                  filteredRecords = records.filter(r => r.date === summaryDate);
                } else if (summaryPeriod === 'month') {
                  filteredRecords = records.filter(r => r.date.startsWith(summaryMonth));
                } else {
                  filteredRecords = records.filter(r => r.date.startsWith(summaryYear));
                }

                if (summaryBranchId !== 'all') {
                  filteredRecords = filteredRecords.filter(r => r.branchId === summaryBranchId);
                }

                const periodTotals = filteredRecords.reduce((acc, r) => ({
                  revenue: acc.revenue + r.revenue,
                  expenses: acc.expenses + r.totalExpenses,
                  profit: acc.profit + r.netProfit,
                  spentCash: acc.spentCash + (r.spentCash || 0),
                  spentBank: acc.spentBank + (r.spentBank || 0),
                  salary: acc.salary + (r.totalSalaryAdvances ?? r.salaryAdvances?.reduce((sum, adv) => sum + adv.amount, 0) ?? 0),
                  salaryDebt: acc.salaryDebt + (r.salaryAdvances?.filter(a => a.method === 'debt').reduce((sum, a) => sum + a.amount, 0) || 0),
                  salaryFine: acc.salaryFine + (r.salaryAdvances?.filter(a => a.method === 'fine').reduce((sum, a) => sum + a.amount, 0) || 0),
                  recoveredMoney: acc.recoveredMoney + (r.recoveredMoney || (r.cash - r.reserve)),
                  diff: acc.diff + (r.diff || 0)
                }), { revenue: 0, expenses: 0, profit: 0, recoveredMoney: 0, spentCash: 0, spentBank: 0, salary: 0, salaryDebt: 0, salaryFine: 0, diff: 0 });

                // Calculate Category Breakdown
                const categoryStats: { [key: string]: number } = {};
                filteredRecords.forEach(r => {
                  [...(r.spentCashItems || []), ...(r.spentBankItems || [])].forEach(item => {
                    categoryStats[item.category] = (categoryStats[item.category] || 0) + item.amount;
                  });
                });
                const sortedCategories = Object.entries(categoryStats).sort((a, b) => b[1] - a[1]);
                const totalCategorySpent = Object.values(categoryStats).reduce((sum, val) => sum + val, 0);

                // Calculate Employee Breakdown
                const employeeStats: { [key: string]: { total: number, cash: number, bank: number, debt: number, fine: number } } = {};
                filteredRecords.forEach(r => {
                  (r.salaryAdvances || []).forEach(adv => {
                    if (!employeeStats[adv.employee]) employeeStats[adv.employee] = { total: 0, cash: 0, bank: 0, debt: 0, fine: 0 };
                    employeeStats[adv.employee].total += adv.amount;
                    if (adv.method === 'cash' || adv.method === 'salary_cash') employeeStats[adv.employee].cash += adv.amount;
                    else if (adv.method === 'bank' || adv.method === 'salary_bank') employeeStats[adv.employee].bank += adv.amount;
                    else if (adv.method === 'debt') employeeStats[adv.employee].debt += adv.amount;
                    else if (adv.method === 'fine') employeeStats[adv.employee].fine += adv.amount;
                  });
                });
                const sortedEmployees = Object.entries(employeeStats).sort((a, b) => b[1].total - a[1].total);

                if (filteredRecords.length === 0) {
                  return (
                    <div className="text-center py-20 text-gray-500">
                      Không có dữ liệu cho giai đoạn này.
                    </div>
                  );
                }

                return (
                  <div className="space-y-4">
                    <div className="grid grid-cols-2 gap-4">
                      <div className="bg-[#2d2d2d] p-4 rounded-2xl border border-white/5">
                        <span className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Tổng Doanh Thu</span>
                        <span className="text-2xl font-mono text-[#00ff88]">{formatCurrency(periodTotals.revenue)}</span>
                      </div>
                      <div className="bg-[#2d2d2d] p-4 rounded-2xl border border-white/5">
                        <span className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Tổng Lợi Nhuận</span>
                        <span className="text-2xl font-mono text-white">{formatCurrency(periodTotals.profit)}</span>
                      </div>
                      <div className="bg-[#2d2d2d] p-4 rounded-2xl border border-white/5">
                        <span className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Tổng Tiền Thu Hồi</span>
                        <span className="text-2xl font-mono text-[#00ff88]">{formatCurrency(periodTotals.recoveredMoney)}</span>
                      </div>
                    </div>

                    {summaryBranchId === 'all' && currentBranch.role === 'admin' && (
                      <div className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5 space-y-4">
                        <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest border-b border-white/5 pb-2">Doanh thu theo chi nhánh</h3>
                        <div className="space-y-3">
                          {branches.filter(b => b.role !== 'admin').map(b => {
                            const branchRecords = filteredRecords.filter(r => r.branchId === b.id);
                            const branchRevenue = branchRecords.reduce((sum, r) => sum + r.revenue, 0);
                            const branchProfit = branchRecords.reduce((sum, r) => sum + r.netProfit, 0);
                            if (branchRecords.length === 0) return null;
                            return (
                              <div key={b.id} className="flex justify-between items-center">
                                <div>
                                  <div className="text-sm font-bold">{b.displayName}</div>
                                  <div className="text-[10px] text-gray-500">{branchRecords.length} ngày có dữ liệu</div>
                                </div>
                                <div className="text-right">
                                  <div className="text-sm font-mono text-[#00ff88]">{formatCurrency(branchRevenue)}</div>
                                  <div className="text-[10px] font-mono text-gray-400">LN: {formatCurrency(branchProfit)}</div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5 space-y-4">
                        <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest border-b border-white/5 pb-2">Chi tiết chi phí</h3>
                        <div className="space-y-3">
                          <div className="flex justify-between items-center">
                            <span className="text-gray-400 text-sm">Tổng chi tiêu</span>
                            <span className="font-mono text-red-400">{formatCurrency(periodTotals.expenses)}</span>
                          </div>
                          <div className="pl-4 space-y-2 border-l border-white/10">
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-500">Chi tiền mặt</span>
                              <span className="text-gray-300">{formatCurrency(periodTotals.spentCash)}</span>
                            </div>
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-500">Chi chuyển khoản</span>
                              <span className="text-gray-300">{formatCurrency(periodTotals.spentBank)}</span>
                            </div>
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-500">Ứng lương</span>
                              <span className="text-gray-300">{formatCurrency(periodTotals.salary - periodTotals.salaryDebt - periodTotals.salaryFine)}</span>
                            </div>
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-500">Nợ lương</span>
                              <span className="text-gray-300">{formatCurrency(periodTotals.salaryDebt)}</span>
                            </div>
                            <div className="flex justify-between text-xs">
                              <span className="text-gray-500">Phạt</span>
                              <span className="text-gray-300">{formatCurrency(periodTotals.salaryFine)}</span>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5 space-y-4">
                        <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest border-b border-white/5 pb-2">Phân bổ chi phí</h3>
                        <div className="space-y-3">
                          {sortedCategories.slice(0, 5).map(([cat, amount]) => (
                            <div key={cat} className="space-y-1">
                              <div className="flex justify-between text-[10px]">
                                <span className="text-gray-400">{cat}</span>
                                <span className="text-gray-200 font-mono">{formatCurrency(amount)}</span>
                              </div>
                              <div className="h-1 bg-white/5 rounded-full overflow-hidden">
                                <div 
                                  className="h-full bg-[#00ff88]" 
                                  style={{ width: `${(amount / totalCategorySpent) * 100}%` }}
                                />
                              </div>
                            </div>
                          ))}
                          {sortedCategories.length > 5 && (
                            <div className="text-center text-[10px] text-gray-500">
                              + {sortedCategories.length - 5} hạng mục khác
                            </div>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5 space-y-4">
                      <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest border-b border-white/5 pb-2">Chi tiết ứng lương nhân viên</h3>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-4">
                        {sortedEmployees.map(([name, stats]) => (
                          <div key={name} className="space-y-1 border-b border-white/5 pb-2">
                            <div className="flex justify-between items-center">
                              <span className="text-sm font-bold text-gray-200">{name}</span>
                              <span className="text-sm font-mono text-[#00ff88]">{formatCurrency(stats.total)}</span>
                            </div>
                            <div className="flex flex-wrap gap-x-3 gap-y-1">
                              {(stats.cash > 0 || stats.bank > 0) && (
                                <span className="text-[9px] text-gray-500">Ứng: {formatCurrency(stats.cash + stats.bank)}</span>
                              )}
                              {stats.debt > 0 && (
                                <span className="text-[9px] text-red-400/70">Nợ: {formatCurrency(stats.debt)}</span>
                              )}
                              {stats.fine > 0 && (
                                <span className="text-[9px] text-orange-400/70">Phạt: {formatCurrency(stats.fine)}</span>
                              )}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>

                    {summaryPeriod !== 'day' && (
                      <div className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5 space-y-4">
                        <h3 className="text-xs font-bold text-gray-400 uppercase tracking-widest border-b border-white/5 pb-2">Diễn biến doanh thu</h3>
                        <div className="h-32 flex items-end gap-1 px-2">
                          {filteredRecords.slice(-31).map((r, i) => {
                            const maxRevenue = Math.max(...filteredRecords.map(rec => rec.revenue));
                            const height = (r.revenue / maxRevenue) * 100;
                            return (
                              <div 
                                key={i} 
                                className="flex-1 bg-[#00ff88]/20 hover:bg-[#00ff88] transition-all rounded-t-sm group relative"
                                style={{ height: `${height}%` }}
                              >
                                <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 bg-black text-white text-[8px] px-1 py-0.5 rounded opacity-0 group-hover:opacity-100 whitespace-nowrap z-10">
                                  {r.date.slice(-2)}: {formatCurrency(r.revenue)}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                        <div className="flex justify-between text-[8px] text-gray-500 px-2">
                          <span>{filteredRecords[0]?.date}</span>
                          <span>{filteredRecords[filteredRecords.length - 1]?.date}</span>
                        </div>
                      </div>
                    )}

                    <div className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5">
                      <div className="flex justify-between items-center">
                        <div>
                          <span className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Tổng chênh lệch đối soát</span>
                          <span className={`text-lg font-mono ${periodTotals.diff >= 0 ? 'text-[#00ff88]' : 'text-red-400'}`}>
                            {periodTotals.diff >= 0 ? '+' : ''}{formatCurrency(periodTotals.diff)}
                          </span>
                        </div>
                        <div className="text-right">
                          <span className="text-[10px] font-bold text-gray-500 uppercase block mb-1">Số ngày ghi chép</span>
                          <span className="text-lg font-mono text-white">{filteredRecords.length} ngày</span>
                        </div>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </motion.div>
          ) : view === 'calc' ? (
            <motion.div
              key="calc"
              initial={{ opacity: 0, x: -20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: 20 }}
              className="space-y-6"
            >
              <div className="bg-[#2d2d2d] p-6 rounded-2xl shadow-xl border border-white/5">
                <h2 className="text-[#00ff88] text-center text-xl font-bold mb-6 uppercase tracking-widest">Quản Lý Thu Chi</h2>
                
                <div className="space-y-8">
                  {/* Section A+B: General Info & Inventory (Compact) */}
                  <div className="space-y-4">
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest flex items-center gap-2">
                      <Wallet className="w-3 h-3" /> Thông tin chung & Kiểm kê
                    </label>
                    <div className="grid grid-cols-2 gap-3">
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Ngày làm việc</span>
                        <input 
                          type="date" 
                          value={workingDate}
                          onChange={(e) => setWorkingDate(e.target.value)}
                          onClick={(e) => {
                          if (typeof e.currentTarget.showPicker === 'function') {
                            try {
                              e.currentTarget.showPicker();
                            } catch (err) {
                              console.warn("showPicker failed", err);
                            }
                          }
                        }}
                          className="w-full bg-[#111] border border-white/5 p-3 rounded-lg text-[#00ff88] text-sm focus:outline-none focus:ring-2 focus:ring-[#00ff88]/20 transition-all cursor-pointer"
                        />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Dư đầu ngày</span>
                        <CurrencyInput value={duDau} onChange={setDuDau} placeholder="0" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Doanh thu (Sổ)</span>
                        <CurrencyInput value={revenue} onChange={setRevenue} placeholder="0" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Tiền mặt hiện có</span>
                        <CurrencyInput value={cash} onChange={setCash} placeholder="0" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Tiền chuyển khoản</span>
                        <CurrencyInput value={bank} onChange={setBank} placeholder="0" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Tiền chừa két</span>
                        <CurrencyInput value={reserve} onChange={setReserve} placeholder="0" />
                      </div>
                      <div className="space-y-1">
                        <span className="text-[9px] text-gray-600 uppercase ml-1">Tiền thu hồi</span>
                        <div className="w-full bg-black/30 border border-white/5 p-3 rounded-lg text-[#00ff88] text-sm font-mono">
                          {formatCurrency(totals.recoveredMoney)}
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Section C: Expenses */}
                  <div className="space-y-4">
                    <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest flex items-center gap-2">
                      <TrendingDown className="w-3 h-3" /> C. Chi phí phát sinh (Ngoài lương)
                    </label>
                    
                    {/* Cash Expenses */}
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] text-gray-600 uppercase ml-1">Chi tiền mặt ({formatCurrency(totals.nSpentCash)})</span>
                      </div>
                      <div className="space-y-2">
                        {spentCashItems.map((item, index) => (
                          <motion.div initial={{ opacity: 0, y: 5 }} animate={{ opacity: 1, y: 0 }} key={index} className="flex flex-col gap-2 bg-black/10 p-3 rounded-2xl border border-white/5">
                            <div className="flex gap-2">
                              <input 
                                placeholder="Dừa, kem..."
                                value={item.description}
                                onChange={(e) => updateExpenseItem('cash', index, 'description', e.target.value)}
                                className="flex-1 bg-[#111] border border-white/5 p-3 rounded-xl text-xs text-white focus:outline-none"
                              />
                              <div className="w-32">
                                <CurrencyInput 
                                  value={!item.amount ? '' : item.amount.toLocaleString('vi-VN')} 
                                  onChange={(val) => updateExpenseItem('cash', index, 'amount', parseCurrency(val))} 
                                  className="p-3 text-xs"
                                />
                              </div>
                              <button onClick={() => removeExpenseItem('cash', index)} className="p-2 text-red-500/30 hover:text-red-500"><Trash2 className="w-4 h-4" /></button>
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="text-[9px] text-gray-500 uppercase font-bold">Phân loại:</span>
                              <select
                                value={item.category}
                                onChange={(e) => updateExpenseItem('cash', index, 'category', e.target.value)}
                                className="flex-1 bg-[#111] border border-white/5 p-2 rounded-lg text-[10px] text-gray-400 focus:outline-none"
                              >
                                {expenseCategories.map(cat => (
                                  <option key={cat.id} value={cat.name}>{cat.name}</option>
                                ))}
                                {expenseCategories.length === 0 && <option value="Khác">Khác</option>}
                              </select>
                            </div>
                          </motion.div>
                        ))}
                        <button 
                          onClick={() => addExpenseItem('cash')}
                          className="w-full py-3 border-2 border-dashed border-white/5 rounded-2xl text-[10px] font-bold text-gray-500 uppercase hover:border-[#00ff88]/20 hover:text-[#00ff88] transition-all flex items-center justify-center gap-2"
                        >
                          <Plus className="w-3 h-3" /> Thêm chi TM
                        </button>
                      </div>
                    </div>

                    {/* Bank Expenses */}
                    <div className="space-y-3">
                      <div className="flex items-center justify-between">
                        <span className="text-[10px] text-gray-600 uppercase ml-1">Chi chuyển khoản ({formatCurrency(totals.nSpentBank)})</span>
                      </div>
                      <div className="space-y-2">
                        {spentBankItems.map((item, index) => (
                          <motion.div initial={{ opacity: 0, y: 5 }} animate={{ opacity: 1, y: 0 }} key={index} className="flex flex-col gap-2 bg-black/10 p-3 rounded-2xl border border-white/5">
                            <div className="flex gap-2">
                              <input 
                                placeholder="Tiền điện, nước..."
                                value={item.description}
                                onChange={(e) => updateExpenseItem('bank', index, 'description', e.target.value)}
                                className="flex-1 bg-[#111] border border-white/5 p-3 rounded-xl text-xs text-white focus:outline-none"
                              />
                              <div className="w-32">
                                <CurrencyInput 
                                  value={!item.amount ? '' : item.amount.toLocaleString('vi-VN')} 
                                  onChange={(val) => updateExpenseItem('bank', index, 'amount', parseCurrency(val))} 
                                  className="p-3 text-xs"
                                />
                              </div>
                              <button onClick={() => removeExpenseItem('bank', index)} className="p-2 text-red-500/30 hover:text-red-500"><Trash2 className="w-4 h-4" /></button>
                            </div>
                            <div className="flex items-center gap-2">
                              <span className="text-[9px] text-gray-500 uppercase font-bold">Phân loại:</span>
                              <select
                                value={item.category}
                                onChange={(e) => updateExpenseItem('bank', index, 'category', e.target.value)}
                                className="flex-1 bg-[#111] border border-white/5 p-2 rounded-lg text-[10px] text-gray-400 focus:outline-none"
                              >
                                {expenseCategories.map(cat => (
                                  <option key={cat.id} value={cat.name}>{cat.name}</option>
                                ))}
                                {expenseCategories.length === 0 && <option value="Khác">Khác</option>}
                              </select>
                            </div>
                          </motion.div>
                        ))}
                        <button 
                          onClick={() => addExpenseItem('bank')}
                          className="w-full py-3 border-2 border-dashed border-white/5 rounded-2xl text-[10px] font-bold text-gray-500 uppercase hover:border-[#00ff88]/20 hover:text-[#00ff88] transition-all flex items-center justify-center gap-2"
                        >
                          <Plus className="w-3 h-3" /> Thêm chi CK
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Section D: Employee Management */}
                  <div className="space-y-4">
                    <div className="flex items-center justify-between">
                      <label className="text-[10px] font-bold text-gray-500 uppercase tracking-widest flex items-center gap-2">
                        <UserIcon className="w-3 h-3" /> D. Quản lý nhân viên
                      </label>
                    </div>
                    
                    <div className="space-y-3">
                      {salaryAdvances.map((advance, index) => (
                        <motion.div 
                          initial={{ opacity: 0, scale: 0.95 }}
                          animate={{ opacity: 1, scale: 1 }}
                          key={index} 
                          className="bg-[#111] p-3 rounded-xl border border-white/5 space-y-3"
                        >
                          <div className="flex items-center gap-2">
                            <select 
                              value={advance.employee}
                              onChange={(e) => updateSalaryAdvance(index, 'employee', e.target.value)}
                              className="flex-1 bg-[#222] border border-white/5 p-2 rounded-lg text-sm text-white focus:outline-none"
                            >
                              {employees.length > 0 ? (
                                employees.map(emp => <option key={emp.id} value={emp.name}>{emp.name}</option>)
                              ) : (
                                <option value="Nhân viên">Nhân viên (Chưa có DS)</option>
                              )}
                            </select>
                            <select 
                              value={advance.method}
                              onChange={(e) => updateSalaryAdvance(index, 'method', e.target.value as any)}
                              className="w-24 bg-[#222] border border-white/5 p-2 rounded-lg text-[10px] text-white focus:outline-none"
                            >
                              <option value="salary_cash">Lương TM</option>
                              <option value="salary_bank">Lương CK</option>
                              <option value="cash">Ứng TM</option>
                              <option value="bank">Ứng CK</option>
                              <option value="debt">Nợ</option>
                              <option value="fine">Phạt</option>
                            </select>
                            <button 
                              onClick={() => removeSalaryAdvance(index)}
                              className="p-2 text-red-500/50 hover:text-red-500 transition-colors"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                          <div className="space-y-1">
                            <CurrencyInput 
                              value={!advance.amount ? '' : advance.amount.toLocaleString('vi-VN')} 
                              onChange={(val) => updateSalaryAdvance(index, 'amount', parseCurrency(val))} 
                              placeholder="Số tiền"
                              className="p-2 text-sm"
                            />
                          </div>
                        </motion.div>
                      ))}
                      {salaryAdvances.length === 0 && (
                        <p className="text-center text-gray-600 text-[10px] italic py-4">Chưa có khoản ứng lương nào</p>
                      )}
                      <button 
                        onClick={addSalaryAdvance}
                        className="w-full py-3 border-2 border-dashed border-white/5 rounded-2xl text-[10px] font-bold text-gray-500 uppercase hover:border-[#00ff88]/20 hover:text-[#00ff88] transition-all flex items-center justify-center gap-2"
                      >
                        <Plus className="w-3 h-3" /> Thêm ứng lương / Phạt
                      </button>
                    </div>
                  </div>
                </div>

                {/* Summary Dashboard */}
                <div className="mt-12 space-y-4">
                  <div className="grid grid-cols-2 gap-4">
                    <div className="bg-black/20 p-4 rounded-2xl border border-white/5">
                      <span className="text-[8px] font-bold text-gray-500 uppercase tracking-widest block mb-1 text-center">Doanh thu</span>
                      <div className="text-lg font-black font-mono text-[#00ff88] text-center">
                        {formatCurrency(parseCurrency(revenue))}
                      </div>
                    </div>
                    <div className="bg-black/20 p-4 rounded-2xl border border-white/5">
                      <span className="text-[8px] font-bold text-gray-500 uppercase tracking-widest block mb-1 text-center">Lợi nhuận ròng</span>
                      <div className={`text-lg font-black font-mono text-center ${totals.netProfit >= 0 ? 'text-[#00ff88]' : 'text-red-500'}`}>
                        {formatCurrency(totals.netProfit)}
                      </div>
                    </div>
                  </div>

                  <div className="grid grid-cols-3 gap-2">
                    <div className="bg-black/20 p-3 rounded-xl border border-white/5 text-center">
                      <span className="text-[7px] font-bold text-gray-500 uppercase tracking-widest block mb-1">Chi TM</span>
                      <div className="text-xs font-bold font-mono text-gray-300">
                        {formatCurrency(totals.nSpentCash)}
                      </div>
                    </div>
                    <div className="bg-black/20 p-3 rounded-xl border border-white/5 text-center">
                      <span className="text-[7px] font-bold text-gray-500 uppercase tracking-widest block mb-1">Chi CK</span>
                      <div className="text-xs font-bold font-mono text-gray-300">
                        {formatCurrency(totals.nSpentBank)}
                      </div>
                    </div>
                    <div className="bg-black/20 p-3 rounded-xl border border-white/5 text-center">
                      <span className="text-[7px] font-bold text-gray-500 uppercase tracking-widest block mb-1">Tổng chi</span>
                      <div className="text-xs font-bold font-mono text-gray-300">
                        {formatCurrency(totals.totalExpenses)}
                      </div>
                    </div>
                  </div>

                  <div className={`p-6 rounded-2xl text-center border-2 transition-all ${totals.diff >= 0 ? 'bg-[#00ff88]/5 border-[#00ff88]/40 text-[#00ff88]' : 'bg-red-500/5 border-red-500/40 text-red-500'}`}>
                    <span className="text-[10px] font-bold uppercase tracking-widest block mb-1">Chênh lệch đối soát</span>
                    <div className="text-4xl font-black font-mono">
                      {totals.diff > 0 ? '+' : ''}{formatCurrency(totals.diff)}
                    </div>
                    <p className="text-[10px] mt-2 font-bold uppercase tracking-widest opacity-60">
                      {totals.diff >= 0 ? '✅ Khớp / Thừa tiền' : '❌ Thiếu hụt tiền'}
                    </p>
                  </div>

                  <div className="flex justify-between px-2 text-[9px] font-bold text-gray-600 uppercase tracking-widest">
                    <div className="flex flex-col items-start">
                      <span>Thực có (Tài sản + Chi TM)</span>
                      <span className="text-gray-400 text-xs font-mono">{formatCurrency(totals.totalActualAssets)}</span>
                    </div>
                    <div className="flex flex-col items-end">
                      <span>Sổ sách (Dư + Thu - Chi CK)</span>
                      <span className="text-gray-400 text-xs font-mono">{formatCurrency(totals.totalBookExpected)}</span>
                    </div>
                  </div>
                </div>

                {error && (
                  <div className="mt-6 p-3 bg-red-500/10 border border-red-500/20 rounded-xl flex items-center gap-2 text-red-500 text-xs">
                    <AlertCircle className="w-4 h-4" />
                    {error}
                  </div>
                )}

                <div className="grid grid-cols-2 gap-4 mt-8">
                  <button
                    onClick={copyReport}
                    className="bg-white/5 text-white font-bold py-4 rounded-xl flex items-center justify-center gap-2 hover:bg-white/10 transition-all border border-white/5"
                  >
                    {copySuccess ? <Check className="w-5 h-5 text-[#00ff88]" /> : <Copy className="w-5 h-5" />}
                    {copySuccess ? 'Đã Copy' : 'Copy Báo Cáo'}
                  </button>
                  <button
                    onClick={handleSave}
                    disabled={saving}
                    className="bg-[#00ff88] text-black font-bold py-4 rounded-xl flex items-center justify-center gap-2 hover:bg-[#00cc6e] transition-all disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-[#00ff88]/20"
                  >
                    {saving ? (
                      <div className="w-5 h-5 border-2 border-black/30 border-t-black rounded-full animate-spin" />
                    ) : (
                      <>
                        <Save className="w-5 h-5" />
                        Lưu Báo Cáo
                      </>
                    )}
                  </button>
                </div>
              </div>
            </motion.div>
          ) : (
            <motion.div
              key="history"
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              className="space-y-4"
            >
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xl font-bold flex items-center gap-2">
                  <History className="w-5 h-5 text-[#00ff88]" />
                  Lịch sử báo cáo
                </h2>
                <span className="text-[10px] font-bold text-gray-500 uppercase bg-white/5 px-2 py-1 rounded-md">
                  {records.length} ngày
                </span>
              </div>

              <AnimatePresence>
                {saveSuccess && (
                  <motion.div
                    initial={{ opacity: 0, y: -20 }}
                    animate={{ opacity: 1, y: 0 }}
                    exit={{ opacity: 0, y: -20 }}
                    className="bg-[#00ff88]/10 border border-[#00ff88]/20 p-4 rounded-xl flex items-center gap-3 mb-4"
                  >
                    <div className="bg-[#00ff88] rounded-full p-1">
                      <Check className="w-3 h-3 text-black" />
                    </div>
                    <span className="text-[#00ff88] text-sm font-bold">Báo cáo đã được lưu thành công!</span>
                  </motion.div>
                )}
              </AnimatePresence>

              {loading ? (
                <div className="flex flex-col items-center justify-center py-20 gap-4">
                  <div className="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-[#00ff88]"></div>
                  <p className="text-gray-500 text-sm">Đang tải lịch sử...</p>
                </div>
              ) : records.length === 0 ? (
                <div className="bg-[#2d2d2d] p-12 rounded-2xl border border-white/5 text-center">
                  <div className="bg-white/5 w-16 h-16 rounded-full flex items-center justify-center mx-auto mb-4">
                    <AlertCircle className="w-8 h-8 text-gray-600" />
                  </div>
                  <p className="text-gray-400">Chưa có báo cáo nào.</p>
                  <button 
                    onClick={() => setView('calc')}
                    className="mt-4 text-[#00ff88] text-sm font-bold hover:underline"
                  >
                    Tạo báo cáo đầu tiên
                  </button>
                </div>
              ) : (
                <div className="space-y-4">
                  {records.map((record) => (
                    <motion.div
                      layout
                      key={record.id}
                      ref={el => recordRefs.current[record.id!] = el}
                      className="bg-[#2d2d2d] p-5 rounded-2xl border border-white/5 hover:border-white/10 transition-all group relative overflow-hidden"
                    >
                      <div className="flex items-start justify-between mb-4">
                        <div className="flex items-center gap-3">
                          <div className={`w-12 h-12 rounded-2xl flex items-center justify-center ${record.diff >= 0 ? 'bg-[#00ff88]/10 text-[#00ff88]' : 'bg-red-500/10 text-red-500'}`}>
                            {record.diff >= 0 ? <TrendingUp className="w-6 h-6" /> : <TrendingDown className="w-6 h-6" />}
                          </div>
                          <div>
                            <div className="text-base font-bold flex items-center gap-2">
                              <CalendarIcon className="w-4 h-4 text-gray-500" />
                              {new Date(record.date).toLocaleDateString('vi-VN', { 
                                weekday: 'long', 
                                day: '2-digit', 
                                month: '2-digit',
                                year: 'numeric'
                              })}
                            </div>
                            <div className="text-[10px] text-gray-500 font-mono uppercase tracking-widest mt-0.5">
                              Lưu lúc: {new Date(record.createdAt).toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })}
                            </div>
                          </div>
                        </div>
                        <div className={`text-right ${record.diff >= 0 ? 'text-[#00ff88]' : 'text-red-500'}`}>
                          <div className="text-xl font-black font-mono">
                            {record.diff > 0 ? '+' : ''}{formatCurrency(record.diff)}
                          </div>
                          <div className="text-[10px] font-bold uppercase opacity-60">Chênh lệch</div>
                        </div>
                      </div>
                      
                      <div className="bg-black/20 rounded-2xl border border-white/10 mt-4 overflow-hidden">
                        <div className="grid grid-cols-3">
                          <div className="p-3 border-b border-r border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Dư đầu</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.duDau)}</span>
                          </div>
                          <div className="p-3 border-b border-r border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Doanh thu</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.revenue)}</span>
                          </div>
                          <div className="p-3 border-b border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Tiền mặt</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.cash)}</span>
                          </div>
                          <div className="p-3 border-b border-r border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Chuyển khoản</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.bank)}</span>
                          </div>
                          <div className="p-3 border-b border-r border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Chừa két</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.reserve)}</span>
                          </div>
                          <div className="p-3 border-b border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Sổ sách</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.totalBookExpected)}</span>
                          </div>
                          <div className="p-3 border-r border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-gray-600 uppercase block">Thực có</span>
                            <span className="text-[13px] font-mono text-gray-400">{formatCurrency(record.totalActualAssets)}</span>
                          </div>
                          <div className="p-3 border-r border-white/10 space-y-1">
                            <span className="text-[9px] font-bold text-[#00ff88] uppercase block">Lợi nhuận</span>
                            <span className="text-[13px] font-mono text-[#00ff88]">{formatCurrency(record.netProfit)}</span>
                          </div>
                          <div className="p-3 space-y-1">
                            <span className="text-[9px] font-bold text-[#00ff88] uppercase block">Thu hồi</span>
                            <span className="text-[13px] font-mono text-[#00ff88]">{formatCurrency(record.recoveredMoney || (record.cash - record.reserve))}</span>
                          </div>
                        </div>
                      </div>

                      {/* Detailed Expenses in History */}
                      <div className="space-y-4 mt-4">
                        <div className="space-y-2">
                          <span className="text-[10px] font-bold text-white uppercase tracking-widest block">Chi tiết chi TM ({formatCurrency(record.spentCash)})</span>
                          <div className="flex flex-wrap gap-2">
                            {record.spentCashItems && record.spentCashItems.length > 0 ? (
                              record.spentCashItems.map((item, i) => (
                                <div key={i} className="bg-white/5 px-2 py-1 rounded-md text-[10px] text-gray-400 border border-white/5">
                                  <span className="text-gray-200">{item.description}</span>: {formatCurrency(item.amount)} <span className="text-[8px] opacity-60">({item.category})</span>
                                </div>
                              ))
                            ) : (
                              <span className="text-[10px] text-gray-600 italic">Không có chi tiết</span>
                            )}
                          </div>
                        </div>
                        <div className="space-y-2">
                          <span className="text-[10px] font-bold text-white uppercase tracking-widest block">Chi tiết chi CK ({formatCurrency(record.spentBank)})</span>
                          <div className="flex flex-wrap gap-2">
                            {record.spentBankItems && record.spentBankItems.length > 0 ? (
                              record.spentBankItems.map((item, i) => (
                                <div key={i} className="bg-white/5 px-2 py-1 rounded-md text-[10px] text-gray-400 border border-white/5">
                                  <span className="text-gray-200">{item.description}</span>: {formatCurrency(item.amount)} <span className="text-[8px] opacity-60">({item.category})</span>
                                </div>
                              ))
                            ) : (
                              <span className="text-[10px] text-gray-600 italic">Không có chi tiết</span>
                            )}
                          </div>
                        </div>
                      </div>

                      {record.salaryAdvances && record.salaryAdvances.length > 0 && (
                        <div className="mt-4 space-y-2">
                          <span className="text-[10px] font-bold text-white uppercase tracking-widest">Chi tiết ứng lương & Phạt:</span>
                          <div className="flex flex-wrap gap-2">
                            {record.salaryAdvances.map((adv, i) => {
                              const methodLabels: Record<string, string> = {
                                salary_cash: 'Lương TM',
                                salary_bank: 'Lương CK',
                                cash: 'Ứng TM',
                                bank: 'Ứng CK',
                                debt: 'Nợ',
                                fine: 'Phạt'
                              };
                              return (
                                <div key={i} className="bg-white/5 px-2 py-1 rounded-md text-[10px] text-gray-400 border border-white/5">
                                  <span className="text-gray-200">{adv.employee}</span>: {formatCurrency(adv.amount)} <span className="text-[8px] opacity-60">({methodLabels[adv.method] || adv.method})</span>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}

                      {/* Action Buttons at Bottom */}
                      <div className="mt-6 pt-4 border-t border-white/5 flex flex-wrap gap-2">
                        <button 
                          onClick={() => handleEditRecord(record)}
                          className="flex-1 bg-white/5 text-white px-3 py-2.5 rounded-xl text-[10px] font-bold flex items-center justify-center gap-2 hover:bg-white/10 transition-all border border-white/5"
                        >
                          <Edit2 className="w-3 h-3" /> Sửa
                        </button>
                        <button 
                          onClick={() => copyAsImage(record.id!)}
                          className="flex-1 bg-yellow-500/10 text-yellow-400 px-3 py-2.5 rounded-xl text-[10px] font-bold flex items-center justify-center gap-2 hover:bg-yellow-500/20 transition-all border border-yellow-500/10"
                        >
                          <Copy className="w-3 h-3" /> Copy ảnh
                        </button>
                        <button 
                          onClick={() => exportAsImage(record.id!)}
                          className="flex-1 bg-[#00ff88]/10 text-[#00ff88] px-3 py-2.5 rounded-xl text-[10px] font-bold flex items-center justify-center gap-2 hover:bg-[#00ff88]/20 transition-all border border-[#00ff88]/10"
                        >
                          <ImageIcon className="w-3 h-3" /> Lưu ảnh
                        </button>
                        <button 
                          onClick={() => exportToExcel([record], `Báo cáo ngày ${record.date}`)}
                          className="flex-1 bg-blue-500/10 text-blue-400 px-3 py-2.5 rounded-xl text-[10px] font-bold flex items-center justify-center gap-2 hover:bg-blue-500/20 transition-all border border-blue-500/10"
                        >
                          <FileSpreadsheet className="w-3 h-3" /> Excel
                        </button>
                        <button 
                          onClick={() => setRecordToDelete(record.id!)}
                          className="flex-1 bg-red-500/10 text-red-400 px-3 py-2.5 rounded-xl text-[10px] font-bold flex items-center justify-center gap-2 hover:bg-red-500/20 transition-all border border-red-500/10"
                        >
                          <Trash2 className="w-3 h-3" /> Xóa
                        </button>
                      </div>
                    </motion.div>
                  ))}
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Delete Confirmation Modal */}
      <AnimatePresence>
        {recordToDelete && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-[#2d2d2d] w-full max-w-sm rounded-3xl p-6 border border-white/10 shadow-2xl"
            >
              <div className="w-16 h-16 bg-red-500/10 rounded-2xl flex items-center justify-center text-red-500 mx-auto mb-4">
                <AlertCircle className="w-8 h-8" />
              </div>
              <h3 className="text-xl font-bold text-white text-center mb-2">Xác nhận xóa?</h3>
              <p className="text-gray-400 text-center text-sm mb-6">
                Bạn có chắc chắn muốn xóa báo cáo này không? Hành động này không thể hoàn tác.
              </p>
              <div className="flex gap-3">
                <button 
                  onClick={() => setRecordToDelete(null)}
                  className="flex-1 py-3 rounded-2xl bg-white/5 text-gray-400 font-bold text-sm hover:bg-white/10 transition-all"
                >
                  Hủy
                </button>
                <button 
                  onClick={confirmDeleteRecord}
                  className="flex-1 py-3 rounded-2xl bg-red-500 text-white font-bold text-sm hover:bg-red-600 transition-all shadow-lg shadow-red-500/20"
                >
                  Xóa ngay
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Bottom Nav */}
      <nav className="fixed bottom-8 left-1/2 -translate-x-1/2 bg-[#2d2d2d]/90 backdrop-blur-xl border border-white/10 rounded-3xl px-3 py-2 shadow-2xl flex items-center gap-2 z-50">
        <button 
          onClick={() => setView('calc')}
          className={`flex items-center gap-2 px-6 py-3 rounded-2xl transition-all ${view === 'calc' ? 'bg-[#00ff88] text-black font-bold shadow-lg shadow-[#00ff88]/20' : 'text-gray-400 hover:text-white'}`}
        >
          <Calculator className="w-5 h-5" />
          <span className="text-sm">Tính toán</span>
        </button>
        <button 
          onClick={() => setView('history')}
          className={`flex items-center gap-2 px-6 py-3 rounded-2xl transition-all ${view === 'history' ? 'bg-[#00ff88] text-black font-bold shadow-lg shadow-[#00ff88]/20' : 'text-gray-400 hover:text-white'}`}
        >
          <History className="w-5 h-5" />
          <span className="text-sm">Lịch sử</span>
        </button>
        <button 
          onClick={() => setView('summary')}
          className={`flex items-center gap-2 px-6 py-3 rounded-2xl transition-all ${view === 'summary' ? 'bg-[#00ff88] text-black font-bold shadow-lg shadow-[#00ff88]/20' : 'text-gray-400 hover:text-white'}`}
        >
          <TrendingUp className="w-5 h-5" />
          <span className="text-sm">Tổng hợp</span>
        </button>
      </nav>
    </div>
  );
}
