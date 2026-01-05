(function () {
  const { createClient } = supabase;

  // ✅ GANTI INI
  const SUPABASE_URL = "https://unhacmkhjawhoizdctdk.supabase.co";
  const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InVuaGFjbWtoamF3aG9pemRjdGRrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU4MjczOTMsImV4cCI6MjA4MTQwMzM5M30.oKIm1s9gwotCeZVvS28vOCLddhIN9lopjG-YeaULMtk";

  const db = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

  const els = {
    profil: document.getElementById("f_profil"),
    indikator: document.getElementById("f_indikator"),
    pic: document.getElementById("f_pic"),
    program: document.getElementById("f_program"),
    penilaian: document.getElementById("f_penilaian"),
    q: document.getElementById("q"),
    btnApply: document.getElementById("btn_apply"),
    btnReset: document.getElementById("btn_reset"),
    tbody: document.getElementById("tbody"),
    count: document.getElementById("count"),
    status: document.getElementById("status"),
  };

  function setStatus(msg, loading = false) {
    els.status.textContent = msg;
    document.body.classList.toggle("loading", loading);
  }

  function setOptions(selectEl, items) {
    const currentValue = selectEl.value;
    const keep = selectEl.options[0]; // "Semua"
    selectEl.innerHTML = "";
    selectEl.appendChild(keep);
    
    let valueFound = false;

    (items || []).forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      selectEl.appendChild(opt);
      if (v === currentValue) {
        valueFound = true;
      }
    });

    if (valueFound) {
      selectEl.value = currentValue;
    } else {
      selectEl.value = ""; 
    }
  }

  function readFilters() {
    return {
      profil: els.profil.value.trim(),
      indikator: els.indikator.value.trim(),
      pic: els.pic.value.trim(),
      program: els.program.value.trim(),
      penilaian: els.penilaian.value.trim(),
      q: els.q.value.trim().toLowerCase(),
    };
  }
  
  async function loadFilterOptions(currentFilters) {
    setStatus("Memuat opsi filter relevan...", true);

    const { profil, indikator, pic, program, penilaian } = currentFilters;

    // Asumsi RPC Supabase 'get_program_pontren_options' sudah dibuat
    
    const { data, error } = await db
      .rpc('get_program_pontren_options', { 
        profil_filter: profil || null, 
        indikator_filter: indikator || null, 
        pic_filter: pic || null, 
        program_filter: program || null,
        penilaian_filter: penilaian || null 
      })
      .single(); 

    if (error) {
      console.warn("Gagal memanggil RPC get_program_pontren_options. Mengambil opsi default.");
      
      const defaultOptions = await db
        .from("program_pontren_filter_options")
        .select("*")
        .single();
        
      if (defaultOptions.error) {
          console.error(defaultOptions.error);
          setStatus("Gagal memuat opsi filter. Cek console.", false);
          return;
      }
      data = defaultOptions.data;
    }

    setOptions(els.profil, data.profil);
    setOptions(els.indikator, data.indikator);
    setOptions(els.pic, data.pic);
    setOptions(els.program, data.program);
    setOptions(els.penilaian, data.penilaian);

    setStatus("Siap.", false);
  }
  
  function renderRows(rows) {
    if (!rows || rows.length === 0) {
      els.tbody.innerHTML =
        `<tr><td colspan="5" class="muted">Tidak ada data yang cocok.</td></tr>`;
      els.count.textContent = "0";
      return;
    }

    const html = rows
      .map((r) => {
        const safe = (x) =>
          (x ?? "").toString().replace(/</g, "&lt;").replace(/>/g, "&gt;");
          
        // Profil dan Definisi digabung
        const profil_definisi = `<b>${safe(r.profil)}</b><br><span class="muted-small">(${safe(r.definisi)})</span>`;
        
        return `
        <tr>
          <td>${profil_definisi}</td> 
          <td>${safe(r.indikator)}</td>
          <td>${safe(r.pic)}</td>
          <td>${safe(r.program)}</td>
          <td>${safe(r.penilaian || "")}</td>
        </tr>`;
      })
      .join("");

    els.tbody.innerHTML = html;
    els.count.textContent = String(rows.length);
  }

  async function fetchData() {
    const f = readFilters();
    setStatus("Mengambil data...", true);

    // 1. Ambil Opsi Filter yang Relevan
    await loadFilterOptions(f); 
    
    // 2. Lakukan Query Data Utama
    let q = db
      .from("program_pontren")
      .select("id, profil, definisi, indikator, pic, program, penilaian")
      .order("profil", { ascending: true })
      .order("indikator", { ascending: true }); 

    if (f.profil) q = q.eq("profil", f.profil);
    if (f.indikator) q = q.eq("indikator", f.indikator);
    if (f.pic) q = q.eq("pic", f.pic);
    if (f.program) q = q.eq("program", f.program);
    if (f.penilaian) q = q.eq("penilaian", f.penilaian);
    
    // Server-Side Search (ilike)
    if (f.q) {
        const searchQuery = `%${f.q}%`;
        q = q.or(
            `profil.ilike.${searchQuery},definisi.ilike.${searchQuery},indikator.ilike.${searchQuery},pic.ilike.${searchQuery},program.ilike.${searchQuery},penilaian.ilike.${searchQuery}`
        );
    }

    const { data, error } = await q;

    if (error) {
      console.error(error);
      setStatus("Gagal ambil data. Cek console.", false);
      renderRows([]);
      return;
    }

    renderRows(data || []);
    setStatus("Selesai.", false);
  }

  function resetFilters() {
    els.profil.value = "";
    els.indikator.value = ""; 
    els.pic.value = "";
    els.program.value = "";
    els.penilaian.value = "";
    els.q.value = "";
  }

  // Events
  els.btnApply.addEventListener("click", fetchData);
  els.q.addEventListener("keyup", (e) => {
    if (e.key === 'Enter') {
        fetchData();
    }
  });
  els.btnReset.addEventListener("click", async () => {
    resetFilters();
    await fetchData(); 
  });

  [els.profil, els.indikator, els.pic, els.program, els.penilaian].forEach((el) =>
    el.addEventListener("change", fetchData)
  );

  // Init
  (async function init() {
    await loadFilterOptions(readFilters()); 
    await fetchData();
  })();
})();
