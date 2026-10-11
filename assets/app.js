(() => {
  'use strict';

  const cfg = window.SITE_CONFIG || {};
  const API = (cfg.API_URL || '').trim();
  const DEMO = !API;
  const TG = (cfg.TELEGRAM || '').replace(/^@/, '');

  const $ = (sel) => document.querySelector(sel);
  const els = {
    tz: $('#tzLabel'),
    state: $('#slotsState'),
    ui: $('#slotsUi'),
    days: $('#days'),
    times: $('#times'),
    form: $('#bookingForm'),
    picked: $('#picked'),
    pickedValue: $('#pickedValue'),
    submit: $('#submitBtn'),
    note: $('#formNote'),
    success: $('#success'),
    successText: $('#successText'),
    again: $('#againBtn'),
  };

  const state = { slots: [], byDate: new Map(), date: null, slot: null };

  // ---------- Мелочи ----------
  const fmtWeekday = new Intl.DateTimeFormat('ru-RU', { weekday: 'short' });
  const fmtMonthShort = new Intl.DateTimeFormat('ru-RU', { month: 'short' });
  const fmtLong = new Intl.DateTimeFormat('ru-RU', { weekday: 'long', day: 'numeric', month: 'long' });

  const parseDate = (iso) => { const [y, m, d] = iso.split('-').map(Number); return new Date(y, m - 1, d); };
  const toIso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const cap = (s) => s.charAt(0).toUpperCase() + s.slice(1);
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const escapeHtml = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function slotLabel(slot) {
    return `${cap(fmtLong.format(parseDate(slot.date)))}, ${slot.time}`;
  }

  // ---------- Данные ----------
  function demoSlots() {
    const times = ['10:00', '12:00', '15:30', '17:00', '18:30', '20:00'];
    const out = [];
    const today = new Date();
    for (let i = 1; i <= 16; i++) {
      const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() + i);
      if (d.getDay() === 0) continue; // воскресенье — выходной
      times.forEach((t, j) => {
        if ((i * 7 + j * 3) % 5 < 2) return; // часть окон «уже занята»
        out.push({ id: `${toIso(d)} ${t}`, date: toIso(d), time: t, duration: 60 });
      });
    }
    return out;
  }

  async function fetchSlots() {
    if (DEMO) { await sleep(500); return demoSlots(); }
    const res = await fetch(`${API}?action=slots&t=${Date.now()}`, { cache: 'no-store' });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || 'bad response');
    return data.slots;
  }

  async function sendBooking(payload) {
    if (DEMO) { await sleep(900); return { ok: true }; }
    // text/plain — чтобы браузер не делал preflight-запрос, который Apps Script не понимает
    const res = await fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    return res.json();
  }

  // ---------- Отрисовка слотов ----------
  function setSlots(slots) {
    state.slots = slots.slice().sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
    state.byDate = new Map();
    for (const s of state.slots) {
      if (!state.byDate.has(s.date)) state.byDate.set(s.date, []);
      state.byDate.get(s.date).push(s);
    }
    if (!state.byDate.has(state.date)) state.date = state.byDate.keys().next().value || null;
    if (state.slot && !state.slots.some((s) => s.id === state.slot.id)) selectSlot(null);
    render();
  }

  function render() {
    if (!state.slots.length) {
      els.ui.hidden = true;
      els.state.hidden = false;
      els.state.innerHTML = `<div class="msg"><b>Свободных окон пока нет</b>
        Новое расписание скоро появится. А пока можно написать мне напрямую.
        <br>${TG ? `<a class="btn btn--ghost" href="https://t.me/${encodeURIComponent(TG)}" target="_blank" rel="noopener">Написать в Telegram</a>` : ''}</div>`;
      return;
    }
    els.state.hidden = true;
    els.ui.hidden = false;
    renderDays();
    renderTimes();
  }

  function renderDays() {
    els.days.innerHTML = '';
    for (const [iso, list] of state.byDate) {
      const d = parseDate(iso);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'day';
      btn.setAttribute('role', 'tab');
      btn.setAttribute('aria-selected', String(iso === state.date));
      btn.setAttribute('aria-label', `${fmtLong.format(d)}, свободно окон: ${list.length}`);
      btn.innerHTML = `
        <span class="day__wd">${fmtWeekday.format(d).replace('.', '')}</span>
        <span class="day__d">${d.getDate()}</span>
        <span class="day__m">${fmtMonthShort.format(d).replace('.', '')}</span>
        <span class="day__c">${list.length}</span>`;
      btn.addEventListener('click', () => { state.date = iso; renderDays(); renderTimes(); });
      els.days.appendChild(btn);
    }
  }

  function renderTimes() {
    els.times.innerHTML = '';
    for (const s of state.byDate.get(state.date) || []) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'time';
      btn.dataset.id = s.id;
      btn.setAttribute('aria-pressed', String(state.slot?.id === s.id));
      btn.innerHTML = `${s.time}${s.duration ? `<small>${s.duration} мин</small>` : ''}`;
      btn.addEventListener('click', () => selectSlot(s));
      els.times.appendChild(btn);
    }
  }

  function selectSlot(slot) {
    state.slot = slot;
    els.picked.dataset.empty = String(!slot);
    els.pickedValue.textContent = slot ? slotLabel(slot) : 'пока не выбрано';
    els.times.querySelectorAll('.time').forEach((b) => {
      b.setAttribute('aria-pressed', String(!!slot && b.dataset.id === slot.id));
    });
    // На телефоне после выбора времени плавно ведём к форме
    if (slot && window.matchMedia('(max-width: 960px)').matches) {
      els.picked.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }

  async function loadSlots() {
    try {
      setSlots(await fetchSlots());
    } catch (err) {
      console.error(err);
      els.ui.hidden = true;
      els.state.hidden = false;
      els.state.innerHTML = `<div class="msg"><b>Не получилось загрузить расписание</b>
        Проверьте интернет и попробуйте ещё раз.<br>
        <button type="button" class="btn btn--ghost" id="retryBtn">Обновить</button></div>`;
      $('#retryBtn').addEventListener('click', () => {
        els.state.innerHTML = '<div class="msg">Загружаю…</div>';
        loadSlots();
      });
    }
  }

  // ---------- Форма ----------
  function setError(name, text) {
    const err = els.form.querySelector(`[data-for="${name}"]`);
    err.textContent = text || '';
    err.closest('.field').classList.toggle('is-invalid', !!text);
  }

  function validate(data) {
    let ok = true;
    if (data.name.length < 2) { setError('name', 'Напишите, как к вам обращаться'); ok = false; } else setError('name');
    if (!/^[A-Za-z0-9_]{4,32}$/.test(data.telegram)) {
      setError('telegram', 'Ник — латиница, цифры и «_», от 4 символов. Например: anna_k');
      ok = false;
    } else setError('telegram');
    if (!state.slot) {
      ok = false;
      els.picked.classList.remove('is-shake');
      void els.picked.offsetWidth;
      els.picked.classList.add('is-shake');
      note('Сначала выберите день и время урока', true);
    }
    return ok;
  }

  function note(text, isError) {
    els.note.textContent = text;
    els.note.classList.toggle('is-error', !!isError);
  }
  const defaultNote = els.note.textContent;

  els.form.telegram.addEventListener('input', (e) => {
    // Убираем @ и ссылку t.me/, если человек вставил их
    const v = e.target.value.replace(/^\s*(https?:\/\/)?(t\.me\/)?@?/i, '').replace(/\s/g, '');
    if (v !== e.target.value) e.target.value = v;
  });
  ['name', 'telegram'].forEach((n) => els.form[n].addEventListener('input', () => setError(n)));

  els.form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const data = {
      name: els.form.name.value.trim(),
      telegram: els.form.telegram.value.trim().replace(/^@/, ''),
      comment: els.form.comment.value.trim(),
      website: els.form.website.value,
    };
    if (!validate(data)) return;
    note(defaultNote);

    const slot = state.slot;
    els.submit.disabled = true;
    els.submit.classList.add('is-loading');
    els.submit.querySelector('.btn__text').textContent = 'Отправляю…';

    try {
      const res = await sendBooking({ ...data, slotId: slot.id });
      if (res.ok) {
        els.form.hidden = true;
        els.success.hidden = false;
        els.successText.innerHTML = `Вы записаны на <b>${escapeHtml(slotLabel(slot))}</b>. Скоро напишу вам в Telegram на <b>@${escapeHtml(data.telegram)}</b>.`;
        state.slots = state.slots.filter((s) => s.id !== slot.id);
        selectSlot(null);
        setSlots(state.slots);
        els.form.comment.value = '';
      } else if (res.error === 'taken') {
        note('Ой, это время только что заняли. Выберите, пожалуйста, другое.', true);
        selectSlot(null);
        await loadSlots();
      } else if (res.error === 'telegram') {
        setError('telegram', 'Проверьте ник в Telegram');
      } else {
        throw new Error(res.error || 'unknown');
      }
    } catch (err) {
      console.error(err);
      note(`Не получилось отправить. Попробуйте ещё раз${TG ? ` или напишите мне в Telegram: @${TG}` : ''}.`, true);
    } finally {
      els.submit.disabled = false;
      els.submit.classList.remove('is-loading');
      els.submit.querySelector('.btn__text').textContent = 'Записаться на урок';
    }
  });

  els.again.addEventListener('click', () => {
    els.success.hidden = true;
    els.form.hidden = false;
    document.getElementById('booking').scrollIntoView({ behavior: 'smooth' });
  });

  // ---------- Старт ----------
  els.tz.textContent = cfg.TIMEZONE_LABEL || '';
  els.tz.hidden = !cfg.TIMEZONE_LABEL;
  document.querySelectorAll('[data-telegram-link]').forEach((a) => {
    if (TG) { a.href = `https://t.me/${encodeURIComponent(TG)}`; a.target = '_blank'; a.rel = 'noopener'; }
    else a.hidden = true;
  });
  if (DEMO) {
    const banner = document.createElement('div');
    banner.className = 'demo-banner';
    banner.textContent = 'Демо-режим: окна пример, заявки никуда не отправляются. Подключите таблицу в config.js.';
    els.form.prepend(banner);
  }

  const reviews = Array.isArray(cfg.REVIEWS) ? cfg.REVIEWS.filter((r) => r && r.text) : [];
  if (reviews.length) {
    $('#reviewsList').innerHTML = reviews.map((r) => `
      <figure class="review">
        ${r.result ? `<span class="review__result">${escapeHtml(r.result)}</span>` : ''}
        <blockquote>${escapeHtml(r.text)}</blockquote>
        <figcaption>${escapeHtml(r.author || '')}${r.about ? `<span>${escapeHtml(r.about)}</span>` : ''}</figcaption>
      </figure>`).join('');
    $('#reviews').hidden = false;
  }

  // Плавающая кнопка «Записаться» на телефоне прячется, когда форма записи уже на экране
  const sticky = $('#stickyCta');
  if (sticky && 'IntersectionObserver' in window) {
    const seen = new Set();
    const io = new IntersectionObserver((entries) => {
      entries.forEach((e) => (e.isIntersecting ? seen.add(e.target) : seen.delete(e.target)));
      sticky.classList.toggle('is-hidden', seen.size > 0);
    });
    ['.hero__actions', '#booking'].forEach((sel) => io.observe($(sel)));
  }

  const topbar = document.querySelector('.topbar');
  const onScroll = () => topbar.classList.toggle('is-scrolled', window.scrollY > 8);
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();

  loadSlots();
})();
