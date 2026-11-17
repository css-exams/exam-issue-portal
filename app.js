(() => {
  'use strict';
  const $ = s => document.querySelector(s);
  const form = $('#reportForm');
  const search = $('#courseSearch');
  const results = $('#courseResults');
  const chosen = $('#selectedCourse');
  const courseId = $('#courseId');
  const status = $('#status');
  const submitBtn = $('#submitBtn');
  const attachment = $('#attachment');
  const cfg = window.PORTAL_CONFIG || {};
  let activeRequest = null;
  let busy = false;

  const normalize = s => String(s || '').toLowerCase().normalize('NFKC').replace(/[\u064b-\u065f\u0670\u0640]/g, '').replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/\s+/g, ' ').trim();
  const courses = Array.isArray(window.COURSES) ? window.COURSES : [];
  const setStatus = (message, kind) => { status.hidden = false; status.className = `status ${kind}`; status.textContent = message; };
  const hideResults = () => { results.hidden = true; search.setAttribute('aria-expanded', 'false'); };
  const selectCourse = course => {
    courseId.value = String(course.id);
    search.value = course.label;
    chosen.textContent = `تم اختيار: ${course.label}`;
    chosen.hidden = false;
    hideResults();
  };
  const renderCourses = () => {
    const q = normalize(search.value);
    const filtered = q ? courses.filter(c => normalize(c.label).includes(q)) : courses;
    results.replaceChildren();
    for (const course of filtered) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'course-result';
      btn.setAttribute('role', 'option');
      btn.textContent = course.label;
      btn.addEventListener('click', () => selectCourse(course));
      results.appendChild(btn);
    }
    if (!filtered.length) { const empty = document.createElement('div'); empty.className='course-result'; empty.textContent='لا توجد مقررات مطابقة'; results.appendChild(empty); }
    results.hidden = false;
    search.setAttribute('aria-expanded', 'true');
  };
  search.addEventListener('input', () => { courseId.value=''; chosen.hidden=true; renderCourses(); });
  search.addEventListener('focus', renderCourses);
  search.addEventListener('keydown', e => {
    if (e.key === 'Escape') hideResults();
    if (e.key === 'Enter' && !results.hidden) {
      e.preventDefault(); const b = results.querySelector('button'); if (b) b.click();
    }
  });
  document.addEventListener('click', e => { if (!e.target.closest('.course-picker')) hideResults(); });
  attachment.addEventListener('change', () => {
    const f = attachment.files?.[0];
    $('#fileName').textContent = f ? `${f.name} (${(f.size/1024/1024).toFixed(2)} MB)` : 'اختيار ملف';
  });
  if (!cfg.endpoint) $('#setupBanner').hidden = false;

  const toBase64 = file => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
  const setBusy = flag => {
    busy = flag; submitBtn.disabled = flag;
    submitBtn.firstChild.textContent = flag ? 'جارٍ إرسال البلاغ...' : 'إرسال البلاغ ';
  };
  window.addEventListener('message', event => {
    if (!busy || !activeRequest) return;
    const data = event.data;
    // Match the unpredictable per-submission nonce before accepting any response.
    if (!data || data.kind !== 'exam-issue-result' || data.requestId !== activeRequest) return;
    // Google Apps Script HTMLService may emit an opaque ("null") origin from its sandbox.
    // Accept only Google-owned origins or that opaque sandbox origin, with the nonce above.
    let allowedOrigin = event.origin === 'null';
    if (!allowedOrigin) {
      try {
        const host = new URL(event.origin).hostname;
        allowedOrigin = host === 'script.google.com' || host === 'script.googleusercontent.com' || host.endsWith('.googleusercontent.com');
      } catch (_) { return; }
    }
    if (!allowedOrigin) return;
    if (data.ok) {
      setStatus(`تم استلام البلاغ بنجاح. الرقم المرجعي: ${data.reference}. احتفظ بهذا الرقم للمتابعة.`, 'success');
      form.reset(); courseId.value=''; chosen.hidden=true; search.value=''; $('#fileName').textContent='اختيار ملف';
    } else setStatus(data.message || 'تعذر حفظ البلاغ. يرجى المحاولة مرة أخرى.', 'error');
    activeRequest = null; setBusy(false);
    status.scrollIntoView({behavior:'smooth', block:'center'});
  });

  form.addEventListener('submit', async e => {
    e.preventDefault();
    if (busy) return;
    status.hidden = true;
    if (!form.reportValidity()) return;
    if (!courseId.value) { setStatus('يرجى اختيار المقرر من النتائج الظاهرة بعد البحث.', 'error'); search.focus(); return; }
    if (!cfg.endpoint || !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec(?:\?.*)?$/.test(cfg.endpoint)) {
      setStatus('النموذج في وضع المعاينة. يجب إضافة رابط Google Apps Script الصحيح في config.js قبل استقبال البلاغات.', 'error'); return;
    }
    const file = attachment.files?.[0];
    const allowed = ['image/jpeg','image/png','image/webp','application/pdf'];
    if (file && (!allowed.includes(file.type) || file.size > (cfg.maxFileBytes || 2097152))) {
      setStatus('نوع الملف غير مدعوم أو حجمه أكبر من 2 ميجابايت.', 'error'); return;
    }
    setBusy(true);
    try {
      const requestId = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
      activeRequest = requestId;
      const payload = Object.fromEntries(new FormData(form).entries());
      delete payload.attachment;
      payload.courseId = courseId.value;
      payload.requestId = requestId;
      payload.fileName = file?.name || '';
      payload.fileMime = file?.type || '';
      payload.fileBase64 = file ? await toBase64(file) : '';
      const target = $('#submissionFrame');
      target.name = 'submissionFrame';
      const post = document.createElement('form');
      post.action = cfg.endpoint;
      post.method = 'POST';
      post.target = target.name;
      post.style.display = 'none';
      Object.entries(payload).forEach(([key, value]) => {
        const input = document.createElement('input'); input.type='hidden'; input.name=key; input.value=String(value); post.appendChild(input);
      });
      document.body.appendChild(post);
      post.submit(); post.remove();
      setStatus('جارٍ إرسال البلاغ. لا تغلق الصفحة قبل ظهور رقم البلاغ.', 'success');
      // Never resend automatically: the server may already have saved the request.
      setTimeout(() => {
        if (activeRequest !== requestId) return;
        activeRequest = null;
        setBusy(false);
        setStatus('لم يصل تأكيد الاستلام إلى الصفحة. ربما تم حفظ البلاغ بالفعل؛ يُرجى مراجعة القسم قبل إعادة الإرسال لتجنب تكرار البلاغ.', 'error');
      }, 30000);
    } catch (err) { activeRequest = null; setBusy(false); setStatus('تعذر تجهيز البلاغ للإرسال. يرجى إعادة المحاولة.', 'error'); }
  });
})();
