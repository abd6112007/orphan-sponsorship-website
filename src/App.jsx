import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'

const STORAGE_KEY = 'orphans_v2'

const CYCLES = {
  monthly:     'شهرية',
  quarterly:   'كل 3 شهور',
  semi_annual: 'كل 6 شهور',
  annual:      'سنوية',
}

const CYCLE_BADGE = {
  monthly:     'badge badge-green',
  quarterly:   'badge badge-blue',
  semi_annual: 'badge badge-amber',
  annual:      'badge badge-purple',
}

const INIT_DATA = [
  { id: 1, name: 'أحمد محمد علي حسن',       sponsor: 'عبد الرحمن الخطيب',    income: 1200, adminPct: 120, disbursement: 800,  cycle: 'monthly'     },
  { id: 2, name: 'فاطمة يوسف إبراهيم ناصر', sponsor: 'شركة الخير للتبرعات', income: 1500, adminPct: 150, disbursement: 600,  cycle: 'quarterly'   },
  { id: 3, name: 'عمر سامي محمود جاد',       sponsor: 'أسرة الرحمة',          income: 900,  adminPct: 90,  disbursement: 700,  cycle: 'monthly'     },
  { id: 4, name: 'زينب عادل كمال طه',        sponsor: 'صندوق الأيتام',        income: 2000, adminPct: 200, disbursement: 800,  cycle: 'semi_annual' },
  { id: 5, name: 'يوسف رامي نبيل سعد',       sponsor: 'جمعية الأمل',          income: 1100, adminPct: 110, disbursement: 500,  cycle: 'annual'      },
]

const EMPTY_FORM = { name: '', sponsor: '', income: '', adminPct: '', disbursement: '', cycle: 'monthly' }

const calcRemaining = (r) =>
  (Number(r.income) || 0) - (Number(r.adminPct) || 0) - (Number(r.disbursement) || 0)

export default function App() {
  const [data, setData] = useState(() => {
    try { const s = localStorage.getItem(STORAGE_KEY); return s ? JSON.parse(s) : INIT_DATA }
    catch { return INIT_DATA }
  })
  const [form, setForm]               = useState(EMPTY_FORM)
  const [editId, setEditId]           = useState(null)
  const [showForm, setShowForm]       = useState(false)
  const [search, setSearch]           = useState('')
  const [cycleFilter, setCycleFilter] = useState('')
  const [deleteId, setDeleteId]       = useState(null)
  const [toast, setToast]             = useState(null)

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data))
  }, [data])

  const showToast = (msg, type = 'success') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 2500)
  }

  const openAdd = () => { setEditId(null); setForm(EMPTY_FORM); setShowForm(true) }

  const openEdit = (r) => {
    setEditId(r.id)
    setForm({ name: r.name, sponsor: r.sponsor, income: r.income, adminPct: r.adminPct, disbursement: r.disbursement, cycle: r.cycle })
    setShowForm(true)
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  const closeForm = () => { setShowForm(false); setForm(EMPTY_FORM); setEditId(null) }

  const handleChange = (e) => setForm(f => ({ ...f, [e.target.name]: e.target.value }))

  const save = () => {
    if (!form.name.trim() || !form.sponsor.trim()) { showToast('الاسم الرباعي واسم الكافل مطلوبان', 'error'); return }
    const rec = {
      name: form.name.trim(), sponsor: form.sponsor.trim(),
      income: Number(form.income) || 0,
      adminPct: Number(form.adminPct) || 0,
      disbursement: Number(form.disbursement) || 0,
      cycle: form.cycle,
    }
    if (editId !== null) {
      setData(d => d.map(x => x.id === editId ? { ...rec, id: editId } : x))
      showToast('✅ تم تعديل السجل بنجاح')
    } else {
      setData(d => [...d, { ...rec, id: Date.now() }])
      showToast('✅ تم إضافة اليتيم بنجاح')
    }
    closeForm()
  }

  const doDelete = () => {
    setData(d => d.filter(x => x.id !== deleteId))
    setDeleteId(null)
    showToast('تم حذف السجل', 'error')
  }

  const exportExcel = () => {
    const rows = data.map(r => ({
      'الاسم الرباعي':         r.name,
      'اسم الكافل':            r.sponsor,
      'الوارد (₪)':            r.income,
      'النسبة الإدارية (₪)':  r.adminPct,
      'الصادر لليتيم (₪)':    r.disbursement,
      'المتبقي (₪)':           calcRemaining(r),
      'دورة الصرف':            CYCLES[r.cycle],
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [25, 20, 13, 18, 16, 13, 15].map(w => ({ wch: w }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'كفالات الأيتام')
    const d = new Date()
const today = `${d.getDate()}-${d.getMonth()+1}-${d.getFullYear()}`
XLSX.writeFile(wb, `كفالات_الأيتام_${today}.xlsx`)
    showToast('✅ تم تصدير Excel بنجاح')
  }

  const filtered = data.filter(r =>
    (!search || r.name.includes(search) || r.sponsor.includes(search)) &&
    (!cycleFilter || r.cycle === cycleFilter)
  )

  const totalIncome = data.reduce((s, r) => s + r.income, 0)
  const totalAdmin  = data.reduce((s, r) => s + r.adminPct, 0)
  const totalDisb   = data.reduce((s, r) => s + r.disbursement, 0)
  const totalRemain = data.reduce((s, r) => s + calcRemaining(r), 0)
  const formRem     = (Number(form.income)||0) - (Number(form.adminPct)||0) - (Number(form.disbursement)||0)

  const filtTotalIncome = filtered.reduce((s,r)=>s+r.income,0)
  const filtTotalAdmin  = filtered.reduce((s,r)=>s+r.adminPct,0)
  const filtTotalDisb   = filtered.reduce((s,r)=>s+r.disbursement,0)
  const filtTotalRem    = filtered.reduce((s,r)=>s+calcRemaining(r),0)

  return (
    <div className="app" dir="rtl">

      {/* ── Header ── */}
      <header>
        <div>
          <h1>نظام إدارة كفالات الأيتام</h1>
          <p>إدارة بيانات الأيتام والكفلاء والصرفيات المالية</p>
        </div>
        <div className="header-actions">
          <button className="btn" onClick={exportExcel}>⬇ تصدير Excel</button>
          <button className="btn btn-primary" onClick={openAdd}>+ إضافة يتيم</button>
        </div>
      </header>

      <main>

        {/* ── Stats ── */}
        <div className="stats-grid">
          <div className="stat-card">
            <div className="stat-label">عدد الأيتام</div>
            <div className="stat-value">{data.length}</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">إجمالي الوارد</div>
            <div className="stat-value">{totalIncome.toLocaleString('ar')} ₪</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">إجمالي النسبة الإدارية</div>
            <div className="stat-value">{totalAdmin.toLocaleString('ar')} ₪</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">إجمالي الصادر لليتيم</div>
            <div className="stat-value">{totalDisb.toLocaleString('ar')} ₪</div>
          </div>
          <div className="stat-card">
            <div className="stat-label">إجمالي المتبقي</div>
            <div className={`stat-value ${totalRemain >= 0 ? 'green' : 'red'}`}>
              {totalRemain.toLocaleString('ar')} ₪
            </div>
          </div>
        </div>

        {/* ── Form ── */}
        {showForm && (
          <div className="form-card">
            <div className="form-header">
              <h2>{editId ? '✏️ تعديل بيانات يتيم' : '➕ إضافة يتيم جديد'}</h2>
              <button onClick={closeForm}>✕</button>
            </div>
            <div className="form-body">
              <div className="form-grid">
                <div className="form-group span2">
                  <label>الاسم الرباعي *</label>
                  <input name="name" value={form.name} onChange={handleChange} placeholder="مثال: محمد أحمد علي خالد" />
                </div>
                <div className="form-group span2">
                  <label>اسم الكافل *</label>
                  <input name="sponsor" value={form.sponsor} onChange={handleChange} placeholder="اسم الكافل أو الجهة الكافلة" />
                </div>
                <div className="form-group">
                  <label>الوارد المالي (₪)</label>
                  <input type="number" min="0" name="income" value={form.income} onChange={handleChange} placeholder="0" />
                </div>
                <div className="form-group">
                  <label>النسبة الإدارية (₪)</label>
                  <input type="number" min="0" name="adminPct" value={form.adminPct} onChange={handleChange} placeholder="0" />
                </div>
                <div className="form-group">
                  <label>الصادر لليتيم (₪)</label>
                  <input type="number" min="0" name="disbursement" value={form.disbursement} onChange={handleChange} placeholder="0" />
                </div>
                <div className="form-group">
                  <label>المتبقي — محسوب تلقائياً</label>
                  <div className={`remaining-preview ${formRem >= 0 ? 'pos' : 'neg'}`}>
                    {formRem.toLocaleString('ar')} ₪
                  </div>
                </div>
                <div className="form-group span2">
                  <label>دورة الصرفية</label>
                  <select name="cycle" value={form.cycle} onChange={handleChange}>
                    {Object.entries(CYCLES).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                  </select>
                </div>
              </div>
              <div className="form-actions">
                <button className="btn" onClick={closeForm}>إلغاء</button>
                <button className="btn btn-primary" onClick={save}>
                  {editId ? 'حفظ التعديلات' : 'إضافة اليتيم'}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Filters ── */}
        <div className="filters">
          <input className="search-input" placeholder="🔍 ابحث باسم اليتيم أو الكافل..."
            value={search} onChange={e => setSearch(e.target.value)} />
          <select className="filter-select" value={cycleFilter} onChange={e => setCycleFilter(e.target.value)}>
            <option value="">كل الدورات</option>
            {Object.entries(CYCLES).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <span className="records-count">{filtered.length} سجل</span>
        </div>

        {/* ── Table ── */}
        <div className="table-card">
          <div className="table-scroll">
            <table>
              <thead>
                <tr>
                  {['#','الاسم الرباعي','اسم الكافل','الوارد (₪)','النسبة الإدارية (₪)','الصادر لليتيم (₪)','المتبقي (₪)','دورة الصرف','إجراءات'].map(h => (
                    <th key={h}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 ? (
                  <tr className="empty-row"><td colSpan={9}>لا توجد سجلات مطابقة</td></tr>
                ) : filtered.map((r, i) => {
                  const rem = calcRemaining(r)
                  return (
                    <tr key={r.id}>
                      <td className="muted">{i + 1}</td>
                      <td style={{ fontWeight: 500 }}>{r.name}</td>
                      <td className="muted">{r.sponsor}</td>
                      <td className="num">{r.income.toLocaleString('ar')}</td>
                      <td className="num orange">{r.adminPct.toLocaleString('ar')}</td>
                      <td className="num blue">{r.disbursement.toLocaleString('ar')}</td>
                      <td className={`num ${rem >= 0 ? 'green' : 'red'}`}>{rem.toLocaleString('ar')}</td>
                      <td><span className={CYCLE_BADGE[r.cycle]}>{CYCLES[r.cycle]}</span></td>
                      <td>
                        <div className="row-actions">
                          <button className="btn btn-sm" onClick={() => openEdit(r)}>تعديل</button>
                          <button className="btn btn-sm btn-danger" onClick={() => setDeleteId(r.id)}>حذف</button>
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
              {filtered.length > 0 && (
                <tfoot>
                  <tr>
                    <td colSpan={3}>الإجمالي ({filtered.length} سجل)</td>
                    <td className="num">{filtTotalIncome.toLocaleString('ar')}</td>
                    <td className="num orange">{filtTotalAdmin.toLocaleString('ar')}</td>
                    <td className="num blue">{filtTotalDisb.toLocaleString('ar')}</td>
                    <td className={`num ${filtTotalRem >= 0 ? 'green' : 'red'}`}>{filtTotalRem.toLocaleString('ar')}</td>
                    <td colSpan={2}></td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>

        {/* ── Info ── */}
        <div className="info-box">
          <strong>💾 إدارة localStorage</strong><br/>
          • <strong>مسح البيانات:</strong> افتح Console (F12) واكتب: <code>localStorage.removeItem('orphans_v2')</code> ثم اضغط F5<br/>
          • <strong>تصدير JSON:</strong> في Console اكتب: <code>copy(localStorage.getItem('orphans_v2'))</code><br/>
          • <strong>تصدير Excel:</strong> استخدم زر "تصدير Excel" في الأعلى
        </div>

      </main>

      {/* ── Delete Modal ── */}
      {deleteId && (
        <div className="modal-overlay" onClick={e => { if (e.target === e.currentTarget) setDeleteId(null) }}>
          <div className="modal-box">
            <div className="modal-icon">🗑️</div>
            <div className="modal-title">تأكيد الحذف</div>
            <div className="modal-desc">هل أنت متأكد من حذف هذا السجل؟<br/>لا يمكن التراجع عن هذه العملية.</div>
            <div className="modal-actions">
              <button className="btn" onClick={() => setDeleteId(null)}>إلغاء</button>
              <button className="btn btn-danger" style={{ background: '#ef4444', color: '#fff', borderColor: '#ef4444' }} onClick={doDelete}>
                نعم، احذف
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Toast ── */}
      {toast && (
        <div className={`toast ${toast.type}`}>{toast.msg}</div>
      )}

    </div>
  )
}