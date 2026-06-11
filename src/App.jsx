import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'

const STORAGE_KEY = 'orphans_v3'

const CYCLES = {
  monthly:     'شهرية',
  quarterly:   'كل 3 شهور',
  semi_annual: 'كل 6 شهور',
  annual:      'سنوية',
}

const CURRENCIES = {
  ILS: '₪ شيكل',
  USD: '$ دولار',
  JOD: 'د.أ دينار',
}

const CURRENCY_SYMBOL = {
  ILS: '₪',
  USD: '$',
  JOD: 'د.أ',
}

const CYCLE_BADGE = {
  monthly:     'badge badge-green',
  quarterly:   'badge badge-blue',
  semi_annual: 'badge badge-amber',
  annual:      'badge badge-purple',
}

const INIT_DATA = [
  { id: 1, name: 'أحمد محمد علي حسن',       sponsor: 'عبد الرحمن الخطيب',    income: 1200, adminPct: 120, disbursement: 800, cycle: 'monthly',     currency: 'ILS', log: [] },
  { id: 2, name: 'فاطمة يوسف إبراهيم ناصر', sponsor: 'شركة الخير للتبرعات', income: 1500, adminPct: 150, disbursement: 600, cycle: 'quarterly',   currency: 'USD', log: [] },
  { id: 3, name: 'عمر سامي محمود جاد',       sponsor: 'أسرة الرحمة',          income: 900,  adminPct: 90,  disbursement: 700, cycle: 'monthly',     currency: 'ILS', log: [] },
  { id: 4, name: 'زينب عادل كمال طه',        sponsor: 'صندوق الأيتام',        income: 2000, adminPct: 200, disbursement: 800, cycle: 'semi_annual', currency: 'JOD', log: [] },
  { id: 5, name: 'يوسف رامي نبيل سعد',       sponsor: 'جمعية الأمل',          income: 1100, adminPct: 110, disbursement: 500, cycle: 'annual',      currency: 'ILS', log: [] },
]

const EMPTY_FORM = { name: '', sponsor: '', income: '', adminPct: '', disbursement: '', cycle: 'monthly', currency: 'ILS' }

const calcRemaining = (r) =>
  (Number(r.income) || 0) - (Number(r.adminPct) || 0) - (Number(r.disbursement) || 0)

const now = () => new Date().toLocaleString('ar-EG')

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
  const [logOrphan, setLogOrphan]     = useState(null) // يتيم نعرض سجله

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
    setForm({ name: r.name, sponsor: r.sponsor, income: r.income, adminPct: r.adminPct, disbursement: r.disbursement, cycle: r.cycle, currency: r.currency || 'ILS' })
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
      currency: form.currency,
    }
    if (editId !== null) {
      // بناء سجل التعديل
      const old = data.find(x => x.id === editId)
      const changes = []
      if (old.income !== rec.income)       changes.push(`الوارد: ${old.income} ← ${rec.income}`)
      if (old.adminPct !== rec.adminPct)   changes.push(`النسبة الإدارية: ${old.adminPct} ← ${rec.adminPct}`)
      if (old.disbursement !== rec.disbursement) changes.push(`الصادر لليتيم: ${old.disbursement} ← ${rec.disbursement}`)
      if (old.currency !== rec.currency)   changes.push(`العملة: ${CURRENCIES[old.currency]} ← ${CURRENCIES[rec.currency]}`)
      const logEntry = changes.length > 0 ? {
        date: now(),
        type: 'تعديل',
        details: changes.join(' | '),
        remaining: calcRemaining(rec),
        currency: rec.currency,
      } : null
      setData(d => d.map(x => x.id === editId ? {
        ...rec, id: editId,
        log: logEntry ? [...(old.log || []), logEntry] : (old.log || [])
      } : x))
      showToast('✅ تم تعديل السجل بنجاح')
    } else {
      const logEntry = {
        date: now(),
        type: 'إضافة',
        details: `تم إضافة اليتيم — وارد: ${rec.income} | نسبة إدارية: ${rec.adminPct} | صادر: ${rec.disbursement}`,
        remaining: calcRemaining(rec),
        currency: rec.currency,
      }
      setData(d => [...d, { ...rec, id: Date.now(), log: [logEntry] }])
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
    const d = new Date()
    const today = `${d.getDate()}-${d.getMonth()+1}-${d.getFullYear()}`
    const rows = data.map(r => ({
      'الاسم الرباعي':        r.name,
      'اسم الكافل':           r.sponsor,
      'العملة':               CURRENCIES[r.currency || 'ILS'],
      'الوارد':               r.income,
      'النسبة الإدارية':      r.adminPct,
      'الصادر لليتيم':        r.disbursement,
      'المتبقي':              calcRemaining(r),
      'دورة الصرف':           CYCLES[r.cycle],
    }))
    const ws = XLSX.utils.json_to_sheet(rows)
    ws['!cols'] = [25,20,12,12,16,14,12,15].map(w => ({ wch: w }))
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'كفالات الأيتام')
    XLSX.writeFile(wb, `كفالات_الأيتام_${today}.xlsx`)
    showToast('✅ تم تصدير Excel بنجاح')
  }

  const filtered = data.filter(r =>
    (!search || r.name.includes(search) || r.sponsor.includes(search)) &&
    (!cycleFilter || r.cycle === cycleFilter)
  )

  const formRem = (Number(form.income)||0) - (Number(form.adminPct)||0) - (Number(form.disbursement)||0)

  // إحصائيات مجمّعة (نعرضها بدون جمع عملات مختلفة)
  const totalCount  = data.length
  const ilsData     = data.filter(r => (r.currency||'ILS') === 'ILS')
  const usdData     = data.filter(r => r.currency === 'USD')
  const jodData     = data.filter(r => r.currency === 'JOD')

  const sumByCurrency = (arr, field) => arr.reduce((s,r) => s + (Number(r[field])||0), 0)
  const remByCurrency = (arr) => arr.reduce((s,r) => s + calcRemaining(r), 0)

  const filtTotalIncome = filtered.reduce((s,r)=>s+r.income,0)
  const filtTotalAdmin  = filtered.reduce((s,r)=>s+r.adminPct,0)
  const filtTotalDisb   = filtered.reduce((s,r)=>s+r.disbursement,0)
  const filtTotalRem    = filtered.reduce((s,r)=>s+calcRemaining(r),0)

  return (
    <div className="app" dir="rtl">
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

        {/* ── إحصائيات مقسّمة بالعملة ── */}
        <div className="stats-grid">
          <div className="stat-card">
            <div className="stat-label">عدد الأيتام</div>
            <div className="stat-value">{totalCount}</div>
          </div>
          {[['ILS','₪'],['USD','$'],['JOD','د.أ']].map(([cur, sym]) => {
            const arr = data.filter(r => (r.currency||'ILS') === cur)
            if (!arr.length) return null
            const rem = remByCurrency(arr)
            return (
              <div key={cur} className="stat-card">
                <div className="stat-label">متبقي {CURRENCIES[cur].split(' ')[1]} ({arr.length} يتيم)</div>
                <div className={`stat-value ${rem >= 0 ? 'green' : 'red'}`}>{sym} {rem.toLocaleString('ar')}</div>
              </div>
            )
          })}
        </div>

        {/* ── نموذج الإضافة/التعديل ── */}
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
                {/* العملة */}
                <div className="form-group span2">
                  <label>نوع العملة</label>
                  <div style={{display:'flex',gap:10,flexWrap:'wrap'}}>
                    {Object.entries(CURRENCIES).map(([val, label]) => (
                      <label key={val} style={{display:'flex',alignItems:'center',gap:6,cursor:'pointer',padding:'8px 16px',border:`2px solid ${form.currency===val?'#059669':'#e2e8f0'}`,borderRadius:9,background:form.currency===val?'#ecfdf5':'#fff',fontWeight:form.currency===val?700:400,fontSize:14,transition:'all 0.15s'}}>
                        <input type="radio" name="currency" value={val} checked={form.currency===val} onChange={handleChange} style={{display:'none'}} />
                        {label}
                      </label>
                    ))}
                  </div>
                </div>
                <div className="form-group">
                  <label>الوارد المالي ({CURRENCY_SYMBOL[form.currency]})</label>
                  <input type="number" min="0" name="income" value={form.income} onChange={handleChange} placeholder="0" />
                </div>
                <div className="form-group">
                  <label>النسبة الإدارية ({CURRENCY_SYMBOL[form.currency]})</label>
                  <input type="number" min="0" name="adminPct" value={form.adminPct} onChange={handleChange} placeholder="0" />
                </div>
                <div className="form-group">
                  <label>الصادر لليتيم ({CURRENCY_SYMBOL[form.currency]})</label>
                  <input type="number" min="0" name="disbursement" value={form.disbursement} onChange={handleChange} placeholder="0" />
                </div>
                <div className="form-group">
                  <label>المتبقي — محسوب تلقائياً</label>
                  <div className={`remaining-preview ${formRem >= 0 ? 'pos' : 'neg'}`}>
                    {CURRENCY_SYMBOL[form.currency]} {formRem.toLocaleString('ar')}
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
                <button className="btn btn-primary" onClick={save}>{editId ? 'حفظ التعديلات' : 'إضافة اليتيم'}</button>
              </div>
            </div>
          </div>
        )}

        {/* ── فلاتر ── */}
        <div className="filters">
          <input className="search-input" placeholder="🔍 ابحث باسم اليتيم أو الكافل..."
            value={search} onChange={e => setSearch(e.target.value)} />
          <select className="filter-select" value={cycleFilter} onChange={e => setCycleFilter(e.target.value)}>
            <option value="">كل الدورات</option>
            {Object.entries(CYCLES).map(([v, l]) => <option key={v} value={v}>{l}</option>)}
          </select>
          <span className="records-count">{filtered.length} سجل</span>
        </div>

        {/* ── الجدول ── */}
        <div className="table-card">
          <div className="table-scroll">
            <table>
              <thead>
                <tr>
                  {['#','الاسم الرباعي','اسم الكافل','العملة','الوارد','النسبة الإدارية','الصادر لليتيم','المتبقي','دورة الصرف','إجراءات'].map(h => (
                    <th key={h}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 ? (
                  <tr className="empty-row"><td colSpan={10}>لا توجد سجلات مطابقة</td></tr>
                ) : filtered.map((r, i) => {
                  const rem = calcRemaining(r)
                  const sym = CURRENCY_SYMBOL[r.currency || 'ILS']
                  return (
                    <tr key={r.id}>
                      <td className="muted">{i + 1}</td>
                      <td style={{fontWeight:500}}>{r.name}</td>
                      <td className="muted">{r.sponsor}</td>
                      <td>
                        <span className="badge badge-blue">{CURRENCIES[r.currency||'ILS']}</span>
                      </td>
                      <td className="num">{sym} {r.income.toLocaleString('ar')}</td>
                      <td className="num orange">{sym} {r.adminPct.toLocaleString('ar')}</td>
                      <td className="num blue">{sym} {r.disbursement.toLocaleString('ar')}</td>
                      <td className={`num ${rem >= 0 ? 'green' : 'red'}`}>{sym} {rem.toLocaleString('ar')}</td>
                      <td><span className={CYCLE_BADGE[r.cycle]}>{CYCLES[r.cycle]}</span></td>
                      <td>
                        <div className="row-actions">
                          <button className="btn btn-sm" onClick={() => openEdit(r)}>تعديل</button>
                          <button className="btn btn-sm" style={{color:'#7c3aed',borderColor:'#ddd6fe',background:'#f5f3ff'}} onClick={() => setLogOrphan(r)}>السجل</button>
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
                    <td colSpan={4}>الإجمالي ({filtered.length} سجل) *</td>
                    <td className="num">{filtTotalIncome.toLocaleString('ar')}</td>
                    <td className="num orange">{filtTotalAdmin.toLocaleString('ar')}</td>
                    <td className="num blue">{filtTotalDisb.toLocaleString('ar')}</td>
                    <td className={`num ${filtTotalRem >= 0 ? 'green' : 'red'}`}>{filtTotalRem.toLocaleString('ar')}</td>
                    <td colSpan={2} style={{fontSize:10,color:'#94a3b8'}}>* الإجمالي يجمع كل العملات</td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>

        <div className="info-box">
          <strong>💾 إدارة localStorage</strong><br/>
          • <strong>مسح البيانات:</strong> افتح Console (F12) واكتب: <code>localStorage.removeItem('orphans_v3')</code> ثم اضغط F5<br/>
          • <strong>تصدير Excel:</strong> استخدم زر "تصدير Excel" في الأعلى
        </div>

      </main>

      {/* ── مودال سجل اليتيم ── */}
      {logOrphan && (
        <div className="modal-overlay" onClick={e => { if(e.target===e.currentTarget) setLogOrphan(null) }}>
          <div style={{background:'#fff',borderRadius:18,padding:28,width:'100%',maxWidth:580,maxHeight:'80vh',overflowY:'auto',direction:'rtl'}}>
            <div style={{display:'flex',justifyContent:'space-between',alignItems:'center',marginBottom:16}}>
              <div>
                <div style={{fontSize:16,fontWeight:700,color:'#0f172a'}}>📋 سجل التعديلات</div>
                <div style={{fontSize:13,color:'#64748b',marginTop:2}}>{logOrphan.name}</div>
              </div>
              <button className="btn btn-sm" onClick={() => setLogOrphan(null)}>✕ إغلاق</button>
            </div>
            {(!logOrphan.log || logOrphan.log.length === 0) ? (
              <div style={{textAlign:'center',padding:'2rem',color:'#94a3b8'}}>لا توجد سجلات بعد</div>
            ) : (
              <div style={{display:'flex',flexDirection:'column',gap:10}}>
                {[...logOrphan.log].reverse().map((entry, i) => (
                  <div key={i} style={{background: entry.type==='إضافة'?'#ecfdf5':entry.type==='تعديل'?'#eff6ff':'#fff5f5', border:`1px solid ${entry.type==='إضافة'?'#a7f3d0':entry.type==='تعديل'?'#bfdbfe':'#fecaca'}`, borderRadius:10, padding:'10px 14px'}}>
                    <div style={{display:'flex',justifyContent:'space-between',marginBottom:4}}>
                      <span style={{fontSize:12,fontWeight:700,color: entry.type==='إضافة'?'#065f46':entry.type==='تعديل'?'#1e40af':'#991b1b'}}>
                        {entry.type==='إضافة'?'➕':entry.type==='تعديل'?'✏️':'🗑️'} {entry.type}
                      </span>
                      <span style={{fontSize:11,color:'#94a3b8'}}>{entry.date}</span>
                    </div>
                    <div style={{fontSize:13,color:'#475569'}}>{entry.details}</div>
                    <div style={{fontSize:12,marginTop:4,fontWeight:600,color: entry.remaining>=0?'#059669':'#ef4444'}}>
                      المتبقي بعد العملية: {CURRENCY_SYMBOL[entry.currency||'ILS']} {entry.remaining?.toLocaleString('ar')}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* ── مودال الحذف ── */}
      {deleteId && (
        <div className="modal-overlay" onClick={e => { if(e.target===e.currentTarget) setDeleteId(null) }}>
          <div className="modal-box">
            <div className="modal-icon">🗑️</div>
            <div className="modal-title">تأكيد الحذف</div>
            <div className="modal-desc">هل أنت متأكد من حذف هذا السجل؟<br/>لا يمكن التراجع عن هذه العملية.</div>
            <div className="modal-actions">
              <button className="btn" onClick={() => setDeleteId(null)}>إلغاء</button>
              <button className="btn btn-danger" style={{background:'#ef4444',color:'#fff',borderColor:'#ef4444'}} onClick={doDelete}>نعم، احذف</button>
            </div>
          </div>
        </div>
      )}

      {toast && <div className={`toast ${toast.type}`}>{toast.msg}</div>}
    </div>
  )
}
