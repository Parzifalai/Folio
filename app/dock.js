// Пенал с инструментами можно перетащить к любому краю окна: снизу, сверху, слева, справа.
// У боковых краёв он становится вертикальным; переход анимирован — кнопки «перетекают».
const DOCK_EDGES = ['bottom', 'top', 'left', 'right'];
const DOCK_EASE = 'cubic-bezier(.3, .7, .2, 1)';
// Размер пенала: тянуть можно за любой свободный край или угол. Кнопки сами переносятся
// в ряды/столбцы и уменьшаются или растут (--ds), чтобы занять заданную рамку.
const DOCK_MIN = 40, DOCK_RZ = 6, DOCK_RC = 16, DS_MIN = 0.6, DS_MAX = 1.3;
const DOCK_CURSOR = { t: 'ns-resize', b: 'ns-resize', l: 'ew-resize', r: 'ew-resize', tl: 'nwse-resize', br: 'nwse-resize', tr: 'nesw-resize', bl: 'nesw-resize' };

function makeDockMovable(root, dock, ctx, kind) {
  const key = 'folio.dockEdge.' + kind;
  const grip = el('div', 'dock-grip');
  grip.dataset.tip = 'Перетащите к любому краю · двойной щелчок — вниз · размер — за край панели';
  grip.innerHTML = '<i></i><i></i><i></i><i></i><i></i><i></i>';
  dock.prepend(grip);
  const hideBtn = [...dock.children].find((b) => (b.dataset.tip || '').startsWith('Скрыть'));
  if (hideBtn) hideBtn.classList.add('dock-hide');
  const zone = el('div', 'dock-zone');
  root.append(zone);

  let edge = 'bottom';
  try { const v = localStorage.getItem(key); if (DOCK_EDGES.includes(v)) edge = v; } catch {}

  // Свой размер у горизонтального и у вертикального пенала; null — по содержимому, как раньше.
  const sizeKey = (o) => 'folio.dockSize.' + kind + '.' + o;
  const loadSize = (o) => {
    try { const z = JSON.parse(localStorage.getItem(sizeKey(o)) || 'null'); if (z && z.w > 0 && z.h > 0) return z; } catch {}
    return null;
  };
  const sizes = { h: loadSize('h'), v: loadSize('v') };
  const orient = () => (dock.classList.contains('vertical') ? 'v' : 'h');

  function applySize() {
    const z = sizes[orient()], st = dock.style;
    if (!z) {
      dock.classList.remove('sized', 'scroll');
      st.width = st.height = '';
      st.removeProperty('--ds');
      return;
    }
    const W = root.clientWidth, H = root.clientHeight;
    const w = W ? clamp(z.w, DOCK_MIN, W - 24) : z.w, h = H ? clamp(z.h, DOCK_MIN, H - 24) : z.h;
    dock.classList.add('sized');
    dock.classList.remove('scroll');
    st.width = w + 'px';
    st.height = h + 'px';
    // Самый крупный масштаб кнопок, при котором всё помещается в рамку.
    const fits = (k) => {
      st.setProperty('--ds', k);
      return dock.scrollWidth <= dock.clientWidth + 1 && dock.scrollHeight <= dock.clientHeight + 1;
    };
    if (fits(DS_MAX)) return;
    if (!fits(DS_MIN)) { dock.classList.add('scroll'); return; }
    let lo = DS_MIN, hi = DS_MAX;
    for (let i = 0; i < 7; i++) { const m = (lo + hi) / 2; if (fits(m)) lo = m; else hi = m; }
    st.setProperty('--ds', lo.toFixed(3));
  }

  function setEdge(e) {
    edge = e;
    dock.classList.remove('floating');
    for (const x of DOCK_EDGES) {
      dock.classList.toggle('edge-' + x, x === e);
      root.classList.toggle('dock-at-' + x, x === e);
    }
    dock.classList.toggle('vertical', e === 'left' || e === 'right');
    applySize();
  }

  function setFloat(cx, cy) {
    for (const x of DOCK_EDGES) dock.classList.remove('edge-' + x);
    dock.classList.remove('vertical');
    dock.classList.add('floating');
    dock.style.setProperty('--dock-x', cx + 'px');
    dock.style.setProperty('--dock-y', cy + 'px');
    applySize();
  }

  // Панель свойств ставим рядом с пеналом, со стороны документа.
  function placeCtx() {
    if (!ctx) return;
    const s = ctx.style;
    s.left = s.right = s.top = s.bottom = s.transform = '';
    if (dock.classList.contains('floating')) return;
    const rr = root.getBoundingClientRect(), d = dock.getBoundingClientRect();
    const g = 10;
    if (edge === 'bottom') s.bottom = rr.bottom - d.top + g + 'px';
    else if (edge === 'top') { s.top = d.bottom - rr.top + g + 'px'; s.bottom = 'auto'; }
    else if (edge === 'left') { s.left = d.right - rr.left + g + 'px'; s.top = '50%'; s.bottom = 'auto'; s.transform = 'translateY(-50%)'; }
    else { s.right = rr.right - d.left + g + 'px'; s.left = 'auto'; s.top = '50%'; s.bottom = 'auto'; s.transform = 'translateY(-50%)'; }
  }

  // Плавная смена формы: фон перетекает из старой рамки в новую, кнопки едут на свои места.
  function morph(change, dur = 420, withBg = true) {
    const rr = root.getBoundingClientRect();
    const kids = [...dock.children];
    dock.classList.add('morphing');
    const d0 = dock.getBoundingClientRect();
    const k0 = kids.map((k) => k.getBoundingClientRect());
    change();
    const d1 = dock.getBoundingClientRect();
    const k1 = kids.map((k) => k.getBoundingClientRect());
    const done = () => { dock.classList.remove('morphing'); placeCtx(); };
    if (matchMedia('(prefers-reduced-motion: reduce)').matches || !d0.width) { done(); return; }
    if (!withBg) {
      // Пенал уже едет за курсором — анимируем только кнопки, фон остаётся на нём самом.
      dock.classList.remove('morphing');
      kids.forEach((k, i) => {
        const dx = k0[i].left - k1[i].left - (d0.left - d1.left), dy = k0[i].top - k1[i].top - (d0.top - d1.top);
        if (Math.abs(dx) + Math.abs(dy) < 0.5) return;
        k.animate([{ transform: `translate(${dx}px, ${dy}px)` }, { transform: 'none' }], { duration: dur, easing: DOCK_EASE });
      });
      placeCtx();
      return;
    }
    const bg = el('div', 'dock-morph');
    root.append(bg);
    const box = (r) => ({ left: r.left - rr.left + 'px', top: r.top - rr.top + 'px', width: r.width + 'px', height: r.height + 'px' });
    const anim = bg.animate([box(d0), box(d1)], { duration: dur, easing: DOCK_EASE, fill: 'forwards' });
    kids.forEach((k, i) => {
      const dx = k0[i].left - k1[i].left, dy = k0[i].top - k1[i].top;
      if (Math.abs(dx) + Math.abs(dy) < 0.5) return;
      k.animate([{ transform: `translate(${dx}px, ${dy}px)` }, { transform: 'none' }], { duration: dur, easing: DOCK_EASE, delay: Math.min(i * 14, 160) });
    });
    anim.onfinish = () => { bg.remove(); done(); };
  }

  const nearest = (x, y) => {
    const W = root.clientWidth, H = root.clientHeight;
    const d = { left: x, right: W - x, top: y, bottom: H - y };
    return DOCK_EDGES.reduce((a, b) => (d[b] < d[a] ? b : a));
  };

  const showZone = (e) => {
    for (const x of DOCK_EDGES) zone.classList.toggle('edge-' + x, x === e);
    zone.classList.toggle('on', !!e);
  };

  function moveTo(e, save = true) {
    if (save) { try { localStorage.setItem(key, e); } catch {} }
    morph(() => setEdge(e));
  }

  grip.addEventListener('dblclick', () => moveTo('bottom'));
  grip.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    e.stopPropagation();
    grip.setPointerCapture(e.pointerId);
    const rr = root.getBoundingClientRect();
    const sx = e.clientX, sy = e.clientY;
    let cx = 0, cy = 0, ox = 0, oy = 0, tx = 0, ty = 0, raf = 0, live = true, started = false;
    let target = edge;
    // Перетаскивание начинается, только когда мышь сдвинулась: простой и двойной щелчок не трогают пенал.
    const begin = () => {
      started = true;
      const d0 = dock.getBoundingClientRect();
      cx = d0.left + d0.width / 2 - rr.left;
      cy = d0.top + d0.height / 2 - rr.top;
      // Подхватили — пенал «висит» под курсором. Снизу/сверху форма та же, поэтому без анимации;
      // с бокового края кнопки быстро перестраиваются в ряд.
      if (dock.classList.contains('vertical')) morph(() => setFloat(cx, cy), 240, false);
      else {
        dock.classList.add('morphing');
        setFloat(cx, cy);
        void dock.offsetWidth;
        dock.classList.remove('morphing');
      }
      const gr = grip.getBoundingClientRect(), dr = dock.getBoundingClientRect();
      ox = gr.left + gr.width / 2 - (dr.left + dr.width / 2);
      oy = gr.top + gr.height / 2 - (dr.top + dr.height / 2);
      dock.classList.add('dragging');
      document.body.classList.add('grabbing');
      target = null; // подсветка края появится с первым движением
      raf = requestAnimationFrame(step);
    };
    const step = () => {
      cx += (tx - cx) * 0.32;
      cy += (ty - cy) * 0.32;
      dock.style.setProperty('--dock-x', cx + 'px');
      dock.style.setProperty('--dock-y', cy + 'px');
      if (live) raf = requestAnimationFrame(step);
    };
    // Пенал целиком остаётся в окне (вертикальный — высокий, поэтому считаем по его размеру).
    const keepInside = (px, py) => {
      const hw = dock.offsetWidth / 2 + 8, hh = dock.offsetHeight / 2 + 8;
      tx = clamp(px - ox, hw, rr.width - hw);
      ty = clamp(py - oy, hh, rr.height - hh);
    };
    const move = (ev) => {
      if (!started) {
        if (Math.hypot(ev.clientX - sx, ev.clientY - sy) < 5) return;
        begin();
      }
      const px = ev.clientX - rr.left, py = ev.clientY - rr.top;
      keepInside(px, py);
      const t = nearest(px, py);
      if (t !== target) {
        target = t;
        showZone(t);
        // Прямо в руке пенал принимает форму будущего места: у боковых краёв — столбик.
        const wantV = t === 'left' || t === 'right';
        if (wantV !== dock.classList.contains('vertical')) {
          morph(() => { dock.classList.toggle('vertical', wantV); applySize(); }, 260, false);
          // Ручка теперь в другом месте пенала — держим её под курсором.
          // (по раскладке, а не по экрану: кнопки ещё летят на новые места)
          ox = grip.offsetLeft + grip.offsetWidth / 2 - dock.offsetWidth / 2;
          oy = grip.offsetTop + grip.offsetHeight / 2 - dock.offsetHeight / 2;
          keepInside(px, py);
        }
      }
    };
    const up = () => {
      live = false;
      cancelAnimationFrame(raf);
      grip.removeEventListener('pointermove', move);
      grip.removeEventListener('pointerup', up);
      grip.removeEventListener('pointercancel', up);
      if (!started) return;
      dock.classList.remove('dragging');
      document.body.classList.remove('grabbing');
      showZone(null);
      dock.style.setProperty('--dock-x', cx + 'px');
      dock.style.setProperty('--dock-y', cy + 'px');
      moveTo(target);
    };
    grip.addEventListener('pointermove', move);
    grip.addEventListener('pointerup', up);
    grip.addEventListener('pointercancel', up);
  });

  // ---------- Размер за край ----------
  // Какие края под курсором: все, кроме прижатого к краю окна.
  function rzSides(e) {
    if (dock.classList.contains('floating') || document.body.classList.contains('dock-hidden')) return '';
    const r = dock.getBoundingClientRect();
    const x = e.clientX - r.left, y = e.clientY - r.top;
    // У скруглённого угла хватаем шире: угол — вдоль края на DOCK_RC от него.
    const near = (d, lim) => d < lim;
    const vt = near(y, DOCK_RZ), vb = near(r.height - y, DOCK_RZ), hl = near(x, DOCK_RZ), hr = near(r.width - x, DOCK_RZ);
    const ct = near(y, DOCK_RC), cb = near(r.height - y, DOCK_RC), cl = near(x, DOCK_RC), cr = near(r.width - x, DOCK_RC);
    const anyH = hl || hr || ((vt || vb) && (cl || cr)), anyV = vt || vb || ((hl || hr) && (ct || cb));
    let s = '';
    if (anyV && (ct || vt) && y < r.height / 2 && edge !== 'top') s += 't';
    else if (anyV && (cb || vb) && edge !== 'bottom') s += 'b';
    if (anyH && (cl || hl) && x < r.width / 2 && edge !== 'left') s += 'l';
    else if (anyH && (cr || hr) && edge !== 'right') s += 'r';
    return s;
  }
  const saveSize = (o) => {
    try { if (sizes[o]) localStorage.setItem(sizeKey(o), JSON.stringify(sizes[o])); else localStorage.removeItem(sizeKey(o)); } catch {}
  };
  function resetSize() {
    const o = orient();
    if (!sizes[o]) return;
    sizes[o] = null;
    saveSize(o);
    morph(applySize, 320);
  }

  let lastRz = 0;
  dock.addEventListener('pointermove', (e) => {
    if (e.buttons) return;
    const s = rzSides(e);
    if (s) { dock.dataset.rz = s; dock.style.setProperty('--rz', DOCK_CURSOR[s]); }
    else delete dock.dataset.rz;
  });
  dock.addEventListener('pointerleave', () => delete dock.dataset.rz);
  // Ловим на погружении, чтобы кнопка у края не нажималась, когда тянут размер.
  dock.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    const s = rzSides(e);
    if (!s) return;
    e.preventDefault();
    e.stopPropagation();
    // Двойной щелчок по краю — снова размер по содержимому (detail у pointerdown всегда 0, считаем сами).
    const now = performance.now();
    if (now - lastRz < 400) { lastRz = 0; resetSize(); return; }
    lastRz = now;
    dock.setPointerCapture(e.pointerId);
    const o = orient(), r0 = dock.getBoundingClientRect();
    // Вдоль оси, по которой пенал стоит по центру, он растёт в обе стороны сразу.
    const fx = o === 'v' ? 1 : 2, fy = o === 'v' ? 2 : 1;
    const sx = e.clientX, sy = e.clientY;
    document.body.style.setProperty('--rz', DOCK_CURSOR[s]);
    document.body.classList.add('dock-resizing');
    let raf = 0, lx = sx, ly = sy;
    const frame = () => {
      raf = 0;
      let w = r0.width, h = r0.height;
      const dx = lx - sx, dy = ly - sy;
      if (s.includes('r')) w += dx * fx; else if (s.includes('l')) w -= dx * fx;
      if (s.includes('b')) h += dy * fy; else if (s.includes('t')) h -= dy * fy;
      sizes[o] = {
        w: Math.round(clamp(w, DOCK_MIN, root.clientWidth - 24)),
        h: Math.round(clamp(h, DOCK_MIN, root.clientHeight - 24)),
      };
      applySize();
      placeCtx();
    };
    const move = (ev) => {
      lx = ev.clientX;
      ly = ev.clientY;
      if (!raf) raf = requestAnimationFrame(frame);
    };
    const up = () => {
      if (raf) { cancelAnimationFrame(raf); frame(); }
      dock.removeEventListener('pointermove', move);
      dock.removeEventListener('pointerup', up);
      dock.removeEventListener('pointercancel', up);
      document.body.classList.remove('dock-resizing');
      saveSize(o);
      // Щелчок после перетаскивания края не должен нажать кнопку под курсором.
      const eat = (ce) => { ce.stopPropagation(); ce.preventDefault(); };
      dock.addEventListener('click', eat, { capture: true, once: true });
      setTimeout(() => dock.removeEventListener('click', eat, { capture: true }), 0);
    };
    dock.addEventListener('pointermove', move);
    dock.addEventListener('pointerup', up);
    dock.addEventListener('pointercancel', up);
  }, true);

  setEdge(edge);
  new ResizeObserver(() => { applySize(); placeCtx(); }).observe(root);
  requestAnimationFrame(placeCtx);
  return { placeCtx, applySize, resetSize, edge: () => edge };
}
