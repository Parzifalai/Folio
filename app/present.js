// Показ документа на весь экран — по одному листу, как презентация.
async function startShow(bytes, start, onClose) {
  const pdf = await pdfjsLib.getDocument({ data: bytes, isEvalSupported: false }).promise;
  const n = pdf.numPages;
  let i = clamp(start || 0, 0, n - 1);
  let ended = false;

  const box = el('div', 'slideshow');
  const stage = el('div', 'show-stage');
  const count = el('div', 'show-count');
  const x = el('button', 'show-x', ICONS.close);
  x.dataset.tip = 'Выйти · Esc';
  const prevBtn = el('button', 'show-nav prev', ICONS.left);
  const nextBtn = el('button', 'show-nav next', ICONS.right);
  const hint = el('div', 'show-hint', '← → листать · Esc — выход');
  box.append(stage, prevBtn, nextBtn, count, x, hint);
  document.body.append(box);

  const wasFull = document.body.classList.contains('full');
  window.folio.setFullscreen(true);

  // Листы рисуем под размер экрана и держим под рукой соседние.
  const cache = new Map();
  const dpr = window.devicePixelRatio || 1;
  const render = (k) => {
    if (k < 0 || k >= n) return null;
    if (!cache.has(k)) {
      cache.set(k, (async () => {
        const page = await pdf.getPage(k + 1);
        const base = page.getViewport({ scale: 1 });
        const s = Math.min(box.clientWidth / base.width, box.clientHeight / base.height);
        const vp = page.getViewport({ scale: s * dpr });
        const c = document.createElement('canvas');
        c.width = Math.floor(vp.width);
        c.height = Math.floor(vp.height);
        c.style.width = vp.width / dpr + 'px';
        c.style.height = vp.height / dpr + 'px';
        await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
        return c;
      })());
    }
    return cache.get(k);
  };

  let token = 0;
  const show = async (k) => {
    ended = false;
    box.classList.remove('black');
    i = clamp(k, 0, n - 1);
    const my = ++token;
    count.textContent = `${i + 1} / ${n}`;
    const c = await render(i);
    if (my !== token) return;
    stage.replaceChildren(c);
    c.classList.remove('in');
    void c.offsetWidth;
    c.classList.add('in');
    render(i + 1);
    render(i - 1);
    for (const key of [...cache.keys()]) if (Math.abs(key - i) > 2) cache.delete(key);
  };

  const showEnd = () => {
    ended = true;
    ++token;
    stage.replaceChildren(el('div', 'show-end', 'Конец показа — щёлкните или нажмите любую клавишу, чтобы выйти'));
    count.textContent = '';
  };

  const next = () => (ended ? close() : i < n - 1 ? show(i + 1) : showEnd());
  const prev = () => (ended ? show(n - 1) : show(i - 1));

  // Курсор и подписи прячутся, пока мышь не двигается.
  let idleTimer;
  const wake = () => {
    box.classList.remove('idle');
    clearTimeout(idleTimer);
    idleTimer = setTimeout(() => box.classList.add('idle'), 2000);
  };

  const onKey = (e) => {
    e.preventDefault();
    e.stopPropagation();
    const k = e.key;
    if (k === 'Escape') return close();
    if (ended && !['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(k)) return close();
    if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n', 'т'].includes(k)) next();
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'з'].includes(k)) prev();
    else if (k === 'Home') show(0);
    else if (k === 'End') show(n - 1);
    else if (k === 'b' || k === 'и' || k === '.') box.classList.toggle('black');
  };
  let wheelAt = 0;
  const onWheel = (e) => {
    e.preventDefault();
    const now = Date.now();
    if (now - wheelAt < 280 || Math.abs(e.deltaY) < 4) return;
    wheelAt = now;
    e.deltaY > 0 ? next() : prev();
  };

  box.addEventListener('click', (e) => {
    if (e.target.closest('.show-x')) return close();
    if (e.target.closest('.show-nav.prev')) return prev();
    next();
  });
  box.addEventListener('contextmenu', (e) => { e.preventDefault(); prev(); });
  box.addEventListener('mousemove', wake);
  box.addEventListener('wheel', onWheel, { passive: false });
  document.addEventListener('keydown', onKey, true);

  // Окно ушло в полный экран не сразу — когда размер установится, перерисовываем.
  let lastSize = '';
  const ro = new ResizeObserver(() => {
    const size = `${box.clientWidth}x${box.clientHeight}`;
    if (size === lastSize) return;
    lastSize = size;
    cache.clear();
    if (!ended) show(i);
  });
  ro.observe(box);

  function close() {
    ro.disconnect();
    clearTimeout(idleTimer);
    document.removeEventListener('keydown', onKey, true);
    box.remove();
    pdf.destroy();
    if (!wasFull) window.folio.setFullscreen(false);
    onClose && onClose();
  }

  wake();
}
