const { useState, useCallback, useMemo } = React;
const TODAY = new Date().toISOString().split('T')[0];
const FACTURAS_RAW = __FJ__;

const fmtM = n => {
  if (!n) return '$0';
  if (n >= 1e9) return '$' + (n/1e9).toFixed(2) + 'B';
  if (n >= 1e6) return '$' + (n/1e6).toFixed(1) + 'M';
  if (n >= 1e3) return '$' + Math.round(n/1e3).toLocaleString('es-CL') + 'K';
  return '$' + Math.round(n).toLocaleString('es-CL');
};
const fmtFull = n => '$' + Math.round(n||0).toLocaleString('es-CL');
const fmtDate = d => {
  if (!d) return '—';
  try { return new Date(d + 'T12:00:00').toLocaleDateString('es-CL', {day:'2-digit',month:'short',year:'numeric'}); } catch(e) { return d; }
};
const diasAtraso = fecha => {
  if (!fecha) return 0;
  return Math.floor((new Date(TODAY) - new Date(fecha + 'T12:00:00')) / 86400000);
};
const eN = e => (e||'').toLowerCase().trim();

function EstadoBadge({estado}) {
  const e = eN(estado);
  const map = {'pagada':['Pagada','badge-success'],'pagado':['Pagada','badge-success'],'impaga':['Impaga','badge-warn'],'pendiente':['Pendiente','badge-warn'],'vencida':['Vencida','badge-danger'],'anulada':['Anulada','badge-muted']};
  const [label,cls] = map[e]||[estado||'—','badge-muted'];
  return <span className={`badge ${cls}`}>{label}</span>;
}
function SiiBadge({sii}) {
  const v=(sii||'').toLowerCase().trim();
  if(!v||v==='—') return null;
  const ok=v==='sí'||v==='si'||v==='yes'||v==='1'||v==='true';
  return <span className={`badge ${ok?'badge-success':'badge-danger'}`}>{ok?'✓ SII':'✗ SII'}</span>;
}
function AtrasoChip({fecha, estado}) {
  const e=eN(estado);
  if(e==='pagada'||e==='pagado'||e==='anulada') return null;
  const d=diasAtraso(fecha);
  if(d<=0) return <span className="atraso-chip" style={{background:'#edfaf4',color:'#1a9e6a'}}>Vigente</span>;
  if(d<=30) return <span className="atraso-chip" style={{background:'#fff7ed',color:'#d97b10'}}>{d}d atraso</span>;
  return <span className="atraso-chip" style={{background:'#fff0f0',color:'#d93b3b'}}>{d}d atraso</span>;
}
function LogoIcon() {
  return (<svg viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M8 14C8 11 10 9 13 9H20L14 17H8V14Z" fill="#5fb3f0"/>
    <path d="M8 17H14L8 25V17Z" fill="#4a9de0"/>
    <path d="M14 17L20 9H27C30 9 32 11 32 14V17L26 25H14L20 17H26V14H20L14 17Z" fill="#3a6fd8"/>
    <path d="M14 25H26L20 33H13C10 33 8 31 8 28V25L14 25Z" fill="#5fb3f0"/>
    <path d="M26 25L32 17V28C32 31 30 33 27 33H20L26 25Z" fill="#3a6fd8"/>
  </svg>);
}
function Toast({msg, onDone}) {
  React.useEffect(() => { const t=setTimeout(onDone,3500); return () => clearTimeout(t); }, []);
  const ok = msg.startsWith('✅');
  return <div className="toast" style={{background: ok ? 'var(--success)' : 'var(--danger)'}}>{msg}</div>;
}
function ModalPago({factura, onClose, onConfirm, saving}) {
  const [fechaPago, setFechaPago] = useState(TODAY);
  const [montoPagado, setMontoPagado] = useState(factura.total||0);
  const handleConfirm = () => onConfirm({fechaPago, montoPagado: parseFloat(montoPagado)||0});
  return (
    <div className="overlay" onClick={e => e.target===e.currentTarget && onClose()}>
      <div className="modal">
        <p className="modal-title">💰 Confirmar Pago</p>
        <div className="kv-grid">
          <div className="kv"><span className="kv-key">Cliente</span><span className="kv-val">{factura.razon}</span></div>
          <div className="kv"><span className="kv-key">N° Factura</span><span className="kv-val">{factura.nFactura}</span></div>
          <div className="kv"><span className="kv-key">Monto factura</span><span className="kv-val">{fmtFull(factura.total)}</span></div>
          <div className="kv"><span className="kv-key">Fecha emisión</span><span className="kv-val">{fmtDate(factura.fecha)}</span></div>
        </div>
        <div className="form-row">
          <div className="form-group">
            <label className="form-label">📅 Fecha de pago</label>
            <input type="date" className="form-input" value={fechaPago} onChange={e => setFechaPago(e.target.value)} max={TODAY}/>
          </div>
          <div className="form-group">
            <label className="form-label">💵 Monto pagado</label>
            <div className="input-prefix-wrap">
              <span className="input-prefix">$</span>
              <input type="number" className="form-input with-prefix" value={montoPagado} onChange={e => setMontoPagado(e.target.value)} min="0" step="1"/>
            </div>
            {parseFloat(montoPagado) < factura.total && parseFloat(montoPagado) > 0 &&
              <span className="form-hint warn">⚠️ Pago parcial: falta {fmtFull(factura.total - parseFloat(montoPagado))}</span>
            }
          </div>
        </div>
        <div className="modal-actions">
          <button className="btn btn-ghost" onClick={onClose} disabled={saving}>Cancelar</button>
          <button className="btn btn-success" onClick={handleConfirm} disabled={saving || !fechaPago || !montoPagado}>
            {saving ? '⏳ Guardando...' : '✓ Confirmar Pago'}
          </button>
        </div>
      </div>
    </div>
  );
}
function FichaCliente({cliente, facturas, onClose, onPagar}) {
  const facs = facturas.filter(f => f.rut === cliente.rut);
  const activas = facs.filter(f => eN(f.estado) !== 'anulada');
  const pagadas = activas.filter(f => ['pagada','pagado'].includes(eN(f.estado)));
  const impagas = activas.filter(f => !['pagada','pagado','anulada'].includes(eN(f.estado)));
  const total = activas.reduce((s,f) => s+(f.total||0),0);
  const pagado = pagadas.reduce((s,f) => s+(f.total||0),0);
  const pendiente = impagas.reduce((s,f) => s+(f.total||0),0);
  const pct = total > 0 ? Math.round(pagado/total*100) : 0;
  return (
    <div className="overlay" onClick={e => e.target===e.currentTarget && onClose()}>
      <div className="modal modal-lg">
        <div className="ficha-header">
          <div>
            <p className="modal-title" style={{marginBottom:4}}>🏢 {cliente.razon}</p>
            <p style={{fontSize:'.8rem',color:'var(--muted)'}}>RUT: {cliente.rut} · {facs.length} facturas en total</p>
          </div>
          <button className="btn btn-ghost btn-sm" onClick={onClose}>✕ Cerrar</button>
        </div>
        <div className="ficha-stats">
          <div className="ficha-stat"><span className="ficha-stat-val">{fmtM(total)}</span><span className="ficha-stat-lbl">Total facturado</span></div>
          <div className="ficha-stat"><span className="ficha-stat-val" style={{color:'var(--success)'}}>{fmtM(pagado)}</span><span className="ficha-stat-lbl">Pagado</span></div>
          <div className="ficha-stat"><span className="ficha-stat-val" style={{color:'var(--danger)'}}>{fmtM(pendiente)}</span><span className="ficha-stat-lbl">Pendiente</span></div>
          <div className="ficha-stat">
            <span className="ficha-stat-val">{pct}%</span><span className="ficha-stat-lbl">Recuperación</span>
            <div className="progress-bar" style={{marginTop:4}}><div className="progress-fill" style={{width:pct+'%'}}/></div>
          </div>
        </div>
        <div style={{overflowX:'auto'}}>
          <table><thead><tr><th>Año</th><th>N° Factura</th><th>Fecha</th><th>Neto</th><th>Total c/IVA</th><th>Estado</th><th>F. Pago</th><th>OC</th><th></th></tr></thead>
          <tbody>{facs.sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||'')).map(f => {
            const esI = !['pagada','pagado','anulada'].includes(eN(f.estado));
            const venc = esI && diasAtraso(f.fecha) > (parseInt(f.plazo)||30);
            return (
              <tr key={f.nFactura+'-'+f.row+'-'+f.year} style={{background:venc?'#fff8f8':esI?'#fffdf5':''}}>
                <td style={{fontWeight:600,color:'var(--muted)',fontSize:'.78rem'}}>{f.year}</td>
                <td style={{fontWeight:700,color:'var(--navy)'}}>{f.nFactura}</td>
                <td style={{fontSize:'.78rem'}}>{fmtDate(f.fecha)}</td>
                <td style={{fontSize:'.82rem'}}>{fmtFull(f.neto)}</td>
                <td style={{fontWeight:700,color:venc?'var(--danger)':esI?'var(--warn)':'var(--success)'}}>{fmtFull(f.total)}</td>
                <td><EstadoBadge estado={f.estado}/></td>
                <td style={{fontSize:'.78rem'}}>{fmtDate(f.fechaPago)}</td>
                <td style={{fontSize:'.76rem',color:'var(--muted)'}}>{f.oc||'—'}</td>
                <td>{esI && <button className="btn btn-success btn-sm" onClick={() => { onClose(); onPagar(f); }}>✓ Pagar</button>}</td>
              </tr>
            );
          })}</tbody></table>
        </div>
      </div>
    </div>
  );
}
async function marcarPagadaAPI(factura, fechaPago, montoPagado) {
  const r = await fetch('/api/pagar', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({year:factura.year,row:factura.row,nFactura:factura.nFactura,fechaPago,montoPagado})});
  if(!r.ok) throw new Error(await r.text());
  return r.json();
}
function Dashboard({facturas, setView, onPagar}) {
  const activas=facturas.filter(f=>eN(f.estado)!=='anulada');
  const pagadas=activas.filter(f=>['pagada','pagado'].includes(eN(f.estado)));
  const impagas=activas.filter(f=>!['pagada','pagado','anulada'].includes(eN(f.estado)));
  const vencidas=impagas.filter(f=>f.fecha&&diasAtraso(f.fecha)>(parseInt(f.plazo)||30));
  const totalF=activas.reduce((s,f)=>s+(f.total||0),0);
  const totalP=pagadas.reduce((s,f)=>s+(f.total||0),0);
  const totalPend=impagas.reduce((s,f)=>s+(f.total||0),0);
  const recup=totalF>0?Math.round(totalP/totalF*100):0;
  const porCliente={};
  impagas.forEach(f=>{if(!porCliente[f.razon])porCliente[f.razon]={razon:f.razon,rut:f.rut,pendiente:0,nImpagas:0};porCliente[f.razon].pendiente+=f.total||0;porCliente[f.razon].nImpagas++;});
  const top5=Object.values(porCliente).sort((a,b)=>b.pendiente-a.pendiente).slice(0,5);
  return (<>
    <div className="page-header"><p className="page-title">Dashboard</p><p className="page-sub">Trayectoria EST · {activas.length} facturas · Actualizado __TODAY__</p></div>
    <div className="stats-grid">{[{label:'Total Facturado',val:fmtM(totalF),sub:fmtFull(totalF),icon:'📋',bg:'#eef4fd'},{label:'Por Cobrar',val:fmtM(totalPend),sub:fmtFull(totalPend),icon:'⏳',bg:'#fff7ed'},{label:'Recuperado',val:fmtM(totalP),sub:recup+'% del total',icon:'✅',bg:'#edfaf4'},{label:'Vencidas',val:vencidas.length,sub:fmtM(vencidas.reduce((s,f)=>s+(f.total||0),0))+' en riesgo',icon:'⚠️',bg:'#fff0f0'}].map(c=>(
      <div key={c.label} className="stat-card"><div className="stat-card-top"><div className="stat-label">{c.label}</div><div className="stat-icon" style={{background:c.bg}}>{c.icon}</div></div><div className="stat-val">{c.val}</div><div className="stat-sub">{c.sub}</div></div>
    ))}</div>
    <div className="two-col">
      <div className="table-wrap"><div className="table-header"><span className="table-title">Top 5 — Pendiente</span></div>
        <table><thead><tr><th>#</th><th>Cliente</th><th>Pendiente</th><th>Impagas</th></tr></thead>
        <tbody>{top5.map((c,i)=>(<tr key={c.razon}><td style={{fontWeight:700,color:'var(--muted)'}}>{i+1}</td><td><div style={{fontWeight:600,color:'var(--navy)'}}>{c.razon}</div><div style={{fontSize:'.74rem',color:'var(--muted)'}}>{c.rut}</div></td><td style={{fontWeight:700,color:'var(--danger)'}}>{fmtM(c.pendiente)}</td><td><span style={{fontWeight:700,color:'var(--warn)'}}>{c.nImpagas}</span></td></tr>))}</tbody>
      </table></div>
      <div className="table-wrap"><div className="table-header"><span className="table-title">Vencidas Recientes</span><button className="btn btn-ghost btn-sm" onClick={()=>setView('facturas')}>Ver todas →</button></div>
        <table><thead><tr><th>N° Factura</th><th>Cliente</th><th>Total</th><th>Atraso</th><th></th></tr></thead>
        <tbody>{vencidas.slice(0,6).map(f=>(<tr key={f.nFactura+f.row}><td style={{fontWeight:700}}>{f.nFactura}</td><td style={{fontSize:'.8rem'}}>{f.razon}</td><td style={{fontWeight:700,color:'var(--danger)'}}>{fmtM(f.total)}</td><td><AtrasoChip fecha={f.fecha} estado={f.estado}/></td><td><button className="btn btn-success btn-sm" onClick={()=>onPagar(f)}>✓ Pagar</button></td></tr>))}
        {vencidas.length===0&&<tr><td colSpan="5"><div className="empty-state" style={{padding:20}}>Sin vencidas 🎉</div></td></tr>}</tbody>
      </table></div>
    </div>
  </>);
}
function FacturasView({facturas, onPagar}) {
  const [search,setSearch]=useState('');
  const [efE,setEfE]=useState('todas');
  const [efY,setEfY]=useState('todos');
  const years=[...new Set(facturas.map(f=>f.year))].sort();
  const fil=useMemo(()=>facturas.filter(f=>{
    if(!(f.razon+f.rut+f.nFactura+(f.oc||'')).toLowerCase().includes(search.toLowerCase())) return false;
    if(efY!=='todos'&&f.year!==parseInt(efY)) return false;
    const e=eN(f.estado);
    if(efE==='todas') return true;
    if(efE==='impaga') return !['pagada','pagado','anulada'].includes(e);
    if(efE==='pagada') return ['pagada','pagado'].includes(e);
    if(efE==='anulada') return e==='anulada';
    if(efE==='vencida') return !['pagada','pagado','anulada'].includes(e)&&diasAtraso(f.fecha)>(parseInt(f.plazo)||30);
    return true;
  }).sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||'')),[facturas,search,efE,efY]);
  const cnt={todas:facturas.length,impaga:facturas.filter(f=>!['pagada','pagado','anulada'].includes(eN(f.estado))).length,pagada:facturas.filter(f=>['pagada','pagado'].includes(eN(f.estado))).length,vencida:facturas.filter(f=>!['pagada','pagado','anulada'].includes(eN(f.estado))&&diasAtraso(f.fecha)>(parseInt(f.plazo)||30)).length,anulada:facturas.filter(f=>eN(f.estado)==='anulada').length};
  return (<>
    <div className="page-header"><p className="page-title">Facturas</p><p className="page-sub">Todas las facturas desde SharePoint</p></div>
    <div className="filter-bar">{['todas','impaga','pagada','vencida','anulada'].map(e=>(<button key={e} className={`btn btn-sm ${efE===e?'btn-primary':'btn-ghost'}`} onClick={()=>setEfE(e)}>{e.charAt(0).toUpperCase()+e.slice(1)} ({cnt[e]||0})</button>))}
      <select className="select-bar" value={efY} onChange={e=>setEfY(e.target.value)}><option value="todos">Todos los años</option>{years.map(y=><option key={y} value={y}>{y}</option>)}</select>
    </div>
    <div className="table-wrap">
      <div className="table-header"><span className="table-title">{fil.length} facturas</span><input className="search-bar" placeholder="🔍 Buscar..." value={search} onChange={e=>setSearch(e.target.value)}/></div>
      <table><thead><tr><th>Año</th><th>N° Factura</th><th>Cliente</th><th>RUT</th><th>Fecha</th><th>Neto</th><th>Total c/IVA</th><th>Estado</th><th>Plazo</th><th>F. Pago</th><th>OC</th><th>SII</th><th></th></tr></thead>
      <tbody>{fil.map(f=>{const esI=!['pagada','pagado','anulada'].includes(eN(f.estado));const venc=esI&&diasAtraso(f.fecha)>(parseInt(f.plazo)||30);return(<tr key={f.nFactura+'-'+f.row+'-'+f.year} style={{background:venc?'#fff8f8':esI?'#fffdf5':''}}><td style={{fontWeight:600,color:'var(--muted)',fontSize:'.78rem'}}>{f.year}</td><td style={{fontWeight:700,color:'var(--navy)'}}>{f.nFactura}</td><td><div style={{fontWeight:600,fontSize:'.82rem'}}>{f.razon}</div></td><td style={{color:'var(--muted)',fontSize:'.78rem'}}>{f.rut}</td><td style={{fontSize:'.78rem',color:'var(--text-2)'}}>{fmtDate(f.fecha)}</td><td style={{fontSize:'.82rem'}}>{fmtFull(f.neto)}</td><td style={{fontWeight:700,color:venc?'var(--danger)':esI?'var(--warn)':'var(--success)'}}>{fmtFull(f.total)}</td><td><EstadoBadge estado={f.estado}/></td><td style={{fontSize:'.78rem',color:'var(--muted)'}}>{f.plazo?f.plazo+' días':'—'}</td><td style={{fontSize:'.78rem'}}>{fmtDate(f.fechaPago)}</td><td style={{fontSize:'.76rem',color:'var(--muted)'}}>{f.oc||'—'}</td><td><SiiBadge sii={f.sii}/></td><td>{esI&&<button className="btn btn-success btn-sm" onClick={()=>onPagar(f)}>✓ Pagar</button>}</td></tr>);})
      }{fil.length===0&&<tr><td colSpan="13"><div className="empty-state">Sin facturas.</div></td></tr>}</tbody></table>
    </div>
  </>);
}
function ClientesView({facturas, onPagar}) {
  const [search,setSearch]=useState('');
  const [fichaCliente,setFichaCliente]=useState(null);
  const clientes=useMemo(()=>{const map={};facturas.filter(f=>eN(f.estado)!=='anulada').forEach(f=>{if(!map[f.rut])map[f.rut]={razon:f.razon,rut:f.rut,total:0,pendiente:0,pagado:0,nFacturas:0,nImpagas:0};map[f.rut].nFacturas++;map[f.rut].total+=f.total||0;if(['pagada','pagado'].includes(eN(f.estado)))map[f.rut].pagado+=f.total||0;else{map[f.rut].pendiente+=f.total||0;map[f.rut].nImpagas++;}});return Object.values(map).sort((a,b)=>b.pendiente-a.pendiente);},[facturas]);
  const fil=clientes.filter(c=>c.razon.toLowerCase().includes(search.toLowerCase())||c.rut.includes(search));
  return (<>
    <div className="page-header"><p className="page-title">Clientes</p><p className="page-sub">{clientes.length} clientes activos</p></div>
    <div className="table-wrap">
      <div className="table-header"><span className="table-title">{fil.length} clientes</span><input className="search-bar" placeholder="🔍 Buscar..." value={search} onChange={e=>setSearch(e.target.value)}/></div>
      <table><thead><tr><th>Cliente</th><th>RUT</th><th>Facturas</th><th>Total</th><th>Pendiente</th><th>Recuperación</th><th></th></tr></thead>
      <tbody>{fil.map(c=>{const pct=c.total>0?Math.round(c.pagado/c.total*100):0;return(<tr key={c.rut} style={{cursor:'pointer'}} onClick={()=>setFichaCliente(c)}><td style={{fontWeight:600,color:'var(--navy)'}}>{c.razon}</td><td style={{color:'var(--muted)',fontSize:'.78rem'}}>{c.rut}</td><td><span style={{fontWeight:700,color:'var(--warn)'}}>{c.nImpagas}</span><span style={{color:'var(--muted)',fontSize:'.78rem'}}> imp./{c.nFacturas}</span></td><td style={{fontSize:'.82rem'}}>{fmtM(c.total)}</td><td style={{fontWeight:700,color:'var(--danger)'}}>{fmtM(c.pendiente)}</td><td style={{width:130}}><div style={{display:'flex',justifyContent:'space-between',fontSize:'.72rem',color:'var(--muted)',marginBottom:3}}><span>{pct}%</span></div><div className="progress-bar"><div className="progress-fill" style={{width:pct+'%'}}/></div></td><td><button className="btn btn-ghost btn-sm" onClick={e=>{e.stopPropagation();setFichaCliente(c);}}>Ver ficha →</button></td></tr>);})
      }</tbody></table>
    </div>
    {fichaCliente && <FichaCliente cliente={fichaCliente} facturas={facturas} onClose={()=>setFichaCliente(null)} onPagar={f=>{setFichaCliente(null);onPagar(f);}}/> }
  </>);
}
function App() {
  const [facturas,setFacturas]=useState(FACTURAS_RAW);
  const [view,setView]=useState('dashboard');
  const [modal,setModal]=useState(null);
  const [saving,setSaving]=useState(false);
  const [toast,setToast]=useState(null);
  const impagas=facturas.filter(f=>!['pagada','pagado','anulada'].includes(eN(f.estado)));
  const handleConfirmar = async ({fechaPago, montoPagado}) => {
    setSaving(true);
    try {
      await marcarPagadaAPI(modal, fechaPago, montoPagado);
      setFacturas(prev=>prev.map(f=>f.row===modal.row&&f.year===modal.year?{...f,estado:'Pagada',fechaPago}:f));
      setToast('✅ Factura '+modal.nFactura+' pagada');
      setModal(null);
    } catch(err) { setToast('❌ Error: '+err.message); }
    finally { setSaving(false); }
  };
  return (
    <div className="crm-root">
      <nav className="sidebar">
        <div className="sidebar-logo"><div className="logo-mark"><div className="logo-icon"><LogoIcon/></div><div className="logo-text">Trayectoria EST<span>CRM Facturación</span></div></div></div>
        <div className="nav-section">Menú</div>
        {[{id:'dashboard',icon:'📊',label:'Dashboard'},{id:'facturas',icon:'🧾',label:'Facturas',badge:impagas.length>0?impagas.length:null},{id:'clientes',icon:'🏢',label:'Clientes'}].map(n=>(
          <button key={n.id} className={`nav-btn ${view===n.id?'active':''}`} onClick={()=>setView(n.id)}>
            <span className="nav-icon">{n.icon}</span>{n.label}{n.badge&&<span className="nav-badge">{n.badge}</span>}
          </button>
        ))}
        <div className="sync-info">🔄 Sincronizado desde SharePoint<br/>Actualizado __TODAY__ a las 08:00</div>
      </nav>
      <main className="main">
        {view==='dashboard'&&<Dashboard facturas={facturas} setView={setView} onPagar={setModal}/>}
        {view==='facturas'&&<FacturasView facturas={facturas} onPagar={setModal}/>}
        {view==='clientes'&&<ClientesView facturas={facturas} onPagar={setModal}/>}
      </main>
      {modal&&<ModalPago factura={modal} onClose={()=>setModal(null)} onConfirm={handleConfirmar} saving={saving}/>}
      {toast&&<Toast msg={toast} onDone={()=>setToast(null)}/>}
    </div>
  );
}
ReactDOM.createRoot(document.getElementById('root')).render(<App/>);
