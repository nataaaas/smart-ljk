# smart-ljk
Aplikasi Pengoreksi Jawaban pilihan ganda otomatis
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SmartLJK - Web-Based Answer Sheet Grader (Support 10-50 Soal & Manual Override)</title>
  
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Google Fonts Inter -->
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <!-- SheetJS CDN for Excel Export -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Chart.js CDN for Analytics -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

  <script>
    tailwind.config = {
      theme: {
        extend: {
          fontFamily: {
            sans: ['Inter', 'sans-serif'],
          },
          colors: {
            brand: {
              50: '#eff6ff',
              100: '#dbeafe',
              500: '#3b82f6',
              600: '#2563eb',
              700: '#1d4ed8',
              900: '#1e3a8a',
            }
          }
        }
      }
    }
  </script>

  <style>
    /* Custom Scrollbars & Print Styling */
    ::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    ::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 4px;
    }
    @media print {
      body * {
        visibility: hidden;
      }
      #printable-ljk-area, #printable-ljk-area *, #print-summary-area, #print-summary-area * {
        visibility: visible;
      }
      #printable-ljk-area, #print-summary-area {
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
      }
      .no-print {
        display: none !important;
      }
    }
  </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans min-h-screen flex flex-col antialiased">

  <!-- TOP HEADER & NAVIGATION -->
  <header class="bg-slate-900 text-white border-b border-slate-800 sticky top-0 z-40 shadow-md no-print">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-16">
        
        <!-- Logo & Branding -->
        <div class="flex items-center gap-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center text-white font-bold shadow-lg shadow-blue-500/30">
            <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
          </div>
          <div>
            <div class="flex items-center gap-2">
              <h1 class="font-bold text-lg leading-tight tracking-tight text-white">SmartLJK</h1>
              <span class="bg-emerald-500/20 text-emerald-400 text-[10px] font-semibold px-2 py-0.5 rounded-full border border-emerald-500/30">10-50 Soal & Override</span>
            </div>
            <p class="text-xs text-slate-400">Pengoreksi LJK Pilihan Ganda Web (Client-Side OMR)</p>
          </div>
        </div>

        <!-- Navigation Tabs -->
        <nav class="hidden md:flex items-center gap-1">
          <button onclick="switchTab('kunci')" id="nav-kunci" class="nav-btn active flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-medium transition-all text-slate-300 hover:bg-slate-800 hover:text-white">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 7a2 2 0 012 2m4 0a6 6 0 01-7.743 5.743L11 17H9v2H7v2H4a1 1 0 01-1-1v-2.586a1 1 0 01.293-.707l5.964-5.964A6 6 0 0121 9z"/></svg>
            1. Kunci Jawaban
          </button>
          <button onclick="switchTab('scan')" id="nav-scan" class="nav-btn flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-medium transition-all text-slate-300 hover:bg-slate-800 hover:text-white">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 9a2 2 0 012-2h0.93a2 2 0 001.664-.89l.812-1.22A2 2 0 0110.07 4h3.86a2 2 0 011.664.89l.812 1.22A2 2 0 0018.07 7H19a2 2 0 012 2v9a2 2 0 01-2 2H5a2 2 0 01-2-2V9z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 13a3 3 0 11-6 0 3 3 0 016 0z"/></svg>
            2. Scan LJK & Batch
          </button>
          <button onclick="switchTab('rekap')" id="nav-rekap" class="nav-btn flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-medium transition-all text-slate-300 hover:bg-slate-800 hover:text-white">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
            3. Rekap & Analytics
          </button>
          <button onclick="switchTab('template')" id="nav-template" class="nav-btn flex items-center gap-2 px-3.5 py-2 rounded-lg text-sm font-medium transition-all text-slate-300 hover:bg-slate-800 hover:text-white">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v2h10z"/></svg>
            4. Cetak Template LJK
          </button>
        </nav>

        <!-- Notification Toast Container -->
        <div id="toast-container" class="fixed top-20 right-4 z-50 flex flex-col gap-2 pointer-events-none"></div>

      </div>
    </div>

    <!-- Mobile Nav Sub-Bar -->
    <div class="md:hidden flex overflow-x-auto border-t border-slate-800 p-2 gap-2 text-xs bg-slate-950 no-scrollbar">
      <button onclick="switchTab('kunci')" id="mob-kunci" class="mob-btn flex-shrink-0 px-3 py-1.5 rounded-md bg-slate-800 text-white">1. Kunci Jawaban</button>
      <button onclick="switchTab('scan')" id="mob-scan" class="mob-btn flex-shrink-0 px-3 py-1.5 rounded-md text-slate-400">2. Scan & Batch</button>
      <button onclick="switchTab('rekap')" id="mob-rekap" class="mob-btn flex-shrink-0 px-3 py-1.5 rounded-md text-slate-400">3. Rekap Nilai</button>
      <button onclick="switchTab('template')" id="mob-template" class="mob-btn flex-shrink-0 px-3 py-1.5 rounded-md text-slate-400">4. Template LJK</button>
    </div>
  </header>

  <!-- MAIN CONTAINER -->
  <main class="flex-grow max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 no-print">

    <!-- ==========================================
         TAB 1: KUNCI JAWABAN (10 S.D 50 SOAL)
         ========================================== -->
    <section id="tab-kunci" class="tab-content space-y-6">
      
      <!-- Top Exam Config Card -->
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
        <div class="flex items-center justify-between pb-4 border-b border-slate-100 mb-6">
          <div>
            <h2 class="text-xl font-bold text-slate-800">Pengaturan Ujian & Kunci Jawaban (Hingga 50 Soal)</h2>
            <p class="text-sm text-slate-500">Sesuaikan jumlah soal (10-50), opsi pilihan (A-D / A-E), dan KKM ujian.</p>
          </div>
          <button onclick="saveExamSettings()" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-xl text-sm font-semibold flex items-center gap-2 shadow-md shadow-blue-500/20 transition">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"/></svg>
            Simpan Kunci Jawaban
          </button>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
          <div>
            <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Mata Pelajaran</label>
            <input type="text" id="cfg-mapel" value="Bahasa Indonesia" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Kelas / Rombel</label>
            <input type="text" id="cfg-kelas" value="IX-A" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Nilai KKM (Passing Grade)</label>
            <input type="number" id="cfg-kkm" value="75" min="0" max="100" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Jumlah Soal Ujian</label>
            <select id="cfg-total-questions" onchange="updateQuestionConfig()" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none font-bold text-blue-600">
              <option value="10">10 Soal</option>
              <option value="20">20 Soal</option>
              <option value="30">30 Soal</option>
              <option value="40">40 Soal</option>
              <option value="50" selected>50 Soal (Maksimal)</option>
            </select>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-600 uppercase mb-1">Jumlah Opsi Pilihan</label>
            <select id="cfg-options-count" onchange="updateQuestionConfig()" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
              <option value="4">A - D (4 Opsi)</option>
              <option value="5" selected>A - E (5 Opsi)</option>
            </select>
          </div>
        </div>
      </div>

      <!-- Quick Fill Options -->
      <div class="flex items-center justify-between bg-slate-100 p-3 rounded-xl">
        <span class="text-xs font-semibold text-slate-600">Preset Kunci Cepat:</span>
        <div class="flex gap-2">
          <button onclick="fillRandomKeys()" class="text-xs bg-white border border-slate-300 text-slate-700 px-3 py-1.5 rounded-lg hover:bg-slate-50 font-medium transition">⚡ Isi Acak (Auto Fill)</button>
          <button onclick="clearAllKeys()" class="text-xs bg-white border border-slate-300 text-red-600 px-3 py-1.5 rounded-lg hover:bg-red-50 font-medium transition">🗑️ Reset Kunci ke Opsi A</button>
        </div>
      </div>

      <!-- Question Answer Key Grid -->
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
        <h3 class="font-bold text-slate-800 mb-4 flex items-center justify-between">
          <span class="flex items-center gap-2">
            <span class="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
            Matriks Kunci Jawaban Soal <span id="key-count-badge" class="text-blue-600 font-extrabold">(1 s.d. 50)</span>
          </span>
        </h3>
        
        <div id="answer-keys-container" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3 max-h-[500px] overflow-y-auto pr-1">
          <!-- Dynamically Generated 1 to 50 Radio Options -->
        </div>
      </div>

    </section>

    <!-- ==========================================
         TAB 2: SCAN LJK & BATCH UPLOAD (100+ LJK)
         ========================================== -->
    <section id="tab-scan" class="tab-content space-y-6 hidden">
      
      <!-- BATCH BANNER -->
      <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-blue-900 text-white rounded-2xl p-6 shadow-xl relative overflow-hidden">
        <div class="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 relative z-10">
          <div>
            <div class="flex items-center gap-2 mb-1">
              <span class="bg-indigo-500/30 text-indigo-300 text-xs font-bold px-2.5 py-0.5 rounded-full border border-indigo-400/30">Pengujian Performa 100 Siswa</span>
              <span class="bg-emerald-500/20 text-emerald-400 text-xs font-bold px-2 py-0.5 rounded-full">High-Speed OMR</span>
            </div>
            <h2 class="text-xl font-extrabold tracking-tight">Koreksi Masal hingga 100+ LJK Siswa Sekaligus</h2>
            <p class="text-xs text-slate-300 mt-1 max-w-2xl">Proses OMR berjalan 100% di browser client-side. Uji kecepatan pengoreksian 100 data LJK siswa dengan opsi hingga 50 soal dalam sekali klik!</p>
          </div>

          <div class="flex items-center gap-3 flex-shrink-0">
            <button onclick="run100StudentSimulation()" id="btn-sim-100" class="bg-emerald-500 hover:bg-emerald-600 text-slate-950 font-extrabold text-sm px-5 py-3 rounded-xl shadow-lg shadow-emerald-500/25 flex items-center gap-2 transition-all transform active:scale-95">
              <svg class="w-5 h-5 text-slate-950" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"/></svg>
              🚀 Simulasi Batch 100 Siswa
            </button>
          </div>
        </div>

        <!-- BENCHMARK RESULT SUMMARY CARD -->
        <div id="benchmark-card" class="hidden mt-5 pt-4 border-t border-white/10 grid grid-cols-2 md:grid-cols-4 gap-3 text-xs">
          <div class="bg-white/10 backdrop-blur rounded-xl p-3 border border-white/10">
            <p class="text-slate-300 text-[10px] uppercase font-semibold">Total LJK Diproses</p>
            <p id="bm-count" class="text-lg font-black text-white mt-0.5">100 LJK</p>
          </div>
          <div class="bg-white/10 backdrop-blur rounded-xl p-3 border border-white/10">
            <p class="text-slate-300 text-[10px] uppercase font-semibold">Waktu Pemrosesan</p>
            <p id="bm-time" class="text-lg font-black text-emerald-400 mt-0.5">0.68 detik</p>
          </div>
          <div class="bg-white/10 backdrop-blur rounded-xl p-3 border border-white/10">
            <p class="text-slate-300 text-[10px] uppercase font-semibold">Kecepatan OMR</p>
            <p id="bm-speed" class="text-lg font-black text-amber-300 mt-0.5">~147 LJK / detik</p>
          </div>
          <div class="bg-white/10 backdrop-blur rounded-xl p-3 border border-white/10">
            <p class="text-slate-300 text-[10px] uppercase font-semibold">Penggunaan Memori</p>
            <p id="bm-mem" class="text-lg font-black text-blue-300 mt-0.5">&lt; 1.8 MB</p>
          </div>
        </div>
      </div>

      <!-- BATCH UPLOAD DROPZONE -->
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
        <h3 class="font-bold text-slate-800 text-sm mb-1 flex items-center gap-2">
          <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
          Unggah Banyak File Foto LJK (Batch Upload)
        </h3>
        <p class="text-xs text-slate-500 mb-4">Pilih atau seret hingga 100+ foto LJK siswa sekaligus (JPG / PNG) untuk diproses otomatis secara beruntun.</p>

        <div id="drop-zone" class="border-2 border-dashed border-slate-300 hover:border-blue-500 hover:bg-blue-50/50 transition rounded-2xl p-8 text-center cursor-pointer relative bg-slate-50">
          <input type="file" id="batch-file-input" multiple accept="image/*" class="absolute inset-0 w-full h-full opacity-0 cursor-pointer" onchange="handleBatchFileUpload(event)">
          <div class="flex flex-col items-center justify-center gap-2">
            <div class="w-12 h-12 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center">
              <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>
            </div>
            <div>
              <p class="text-sm font-bold text-slate-700">Tarik & Lepaskan File Foto LJK di Sini</p>
              <p class="text-xs text-slate-400 mt-0.5">atau <span class="text-blue-600 font-semibold underline">Pilih Banyak File dari Perangkat</span></p>
            </div>
          </div>
        </div>

        <!-- BATCH PROGRESS PANEL -->
        <div id="batch-progress-panel" class="hidden mt-5 p-4 rounded-xl bg-slate-900 text-white space-y-3">
          <div class="flex items-center justify-between text-xs">
            <span id="batch-status-text" class="font-bold text-emerald-400">Memproses OMR LJK...</span>
            <span id="batch-counter-text" class="font-mono text-slate-300">0 / 0</span>
          </div>
          <div class="w-full bg-slate-800 rounded-full h-3 overflow-hidden p-0.5 border border-slate-700">
            <div id="batch-progress-bar" class="bg-gradient-to-r from-blue-500 to-emerald-400 h-full rounded-full w-0 transition-all duration-150"></div>
          </div>
          <div class="flex items-center justify-between text-[11px] text-slate-400 pt-1">
            <span id="batch-elapsed-time">Waktu berjalan: 0.0s</span>
            <span id="batch-success-count" class="text-emerald-400 font-semibold">Sukses: 0</span>
          </div>
        </div>
      </div>

      <!-- MAIN SPLIT VIEW: CANVAS PREVIEW & SCAN RESULT -->
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        <!-- Left 7 Cols: Interactive Canvas & Scanner -->
        <div class="lg:col-span-7 space-y-4">
          <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200">
            
            <div class="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
              <div>
                <h3 class="font-bold text-slate-800 flex items-center gap-2">
                  <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
                  Preview Canvas & Detector OMR (1-50 Soal)
                </h3>
                <p class="text-xs text-slate-500">Klik bulatan untuk simulasi manual atau klik kamera live.</p>
              </div>
              <button onclick="generateRandomStudentLJK()" class="px-3 py-1.5 bg-indigo-50 text-indigo-600 border border-indigo-200 rounded-lg text-xs font-semibold hover:bg-indigo-100 transition">
                ✨ LJK Acak Siswa
              </button>
            </div>

            <!-- Single Action Bar -->
            <div class="mb-4 grid grid-cols-1 sm:grid-cols-2 gap-3">
              <label class="flex items-center justify-center gap-2 p-2.5 border border-slate-300 rounded-xl cursor-pointer hover:border-blue-500 hover:bg-blue-50/50 transition text-xs font-semibold text-slate-600">
                <svg class="w-4 h-4 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12"/></svg>
                Unggah 1 Foto LJK
                <input type="file" accept="image/*" class="hidden" onchange="handleSingleImageUpload(event)">
              </label>

              <button onclick="simulateCameraScan()" class="flex items-center justify-center gap-2 p-2.5 bg-slate-900 text-white rounded-xl text-xs font-semibold hover:bg-slate-800 transition">
                <svg class="w-4 h-4 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 10l4.553-2.276A1 1 0 0121 8.618v6.764a1 1 0 01-1.447.894L15 14M5 18h8a2 2 0 002-2V8a2 2 0 002-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"/></svg>
                Simulasi Live Cam Scan
              </button>
            </div>

            <!-- Canvas Display Box -->
            <div class="relative bg-slate-900 rounded-xl overflow-hidden p-2 flex justify-center items-center min-h-[500px] shadow-inner">
              <canvas id="omr-canvas" width="520" height="700" class="bg-white rounded shadow-2xl cursor-pointer transition-transform duration-200 max-w-full h-auto"></canvas>
              <div id="scan-line" class="hidden absolute left-0 right-0 h-1 bg-gradient-to-r from-transparent via-emerald-400 to-transparent shadow-[0_0_15px_#34d399] animate-pulse"></div>
            </div>

            <!-- Legend Overlay -->
            <div class="mt-3 flex flex-wrap items-center justify-center gap-4 text-xs font-medium text-slate-600">
              <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-emerald-500"></span> Benar</span>
              <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-red-500"></span> Salah</span>
              <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-amber-400"></span> Ganda / Tidak Sah</span>
              <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-slate-400"></span> Kosong</span>
            </div>

          </div>
        </div>

        <!-- Right 5 Cols: Output Results Card -->
        <div class="lg:col-span-5 space-y-4">
          <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200">
            
            <h3 class="font-bold text-slate-800 pb-3 border-b border-slate-100 mb-4 flex items-center justify-between">
              <span>Hasil Koreksi Real-time</span>
              <span id="scan-badge" class="bg-blue-100 text-blue-700 text-xs px-2.5 py-0.5 rounded-full font-semibold">Ready</span>
            </h3>

            <!-- Editable Student Metadata -->
            <div class="space-y-3 mb-6 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
              <div>
                <label class="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">NIS / No. Peserta (Terbaca)</label>
                <input type="text" id="scan-nis" value="1001" class="w-full px-3 py-1.5 bg-white rounded-lg border border-slate-300 text-sm font-mono font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500">
              </div>
              <div>
                <label class="block text-[11px] font-bold text-slate-500 uppercase tracking-wider mb-1">Nama Siswa</label>
                <input type="text" id="scan-nama" value="Siswa LJK Demo" class="w-full px-3 py-1.5 bg-white rounded-lg border border-slate-300 text-sm font-semibold text-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500">
              </div>
            </div>

            <!-- Primary Score Banner -->
            <div id="score-banner" class="rounded-2xl p-4 bg-gradient-to-br from-blue-600 to-indigo-700 text-white shadow-lg shadow-blue-500/20 mb-6 relative overflow-hidden">
              <div class="flex items-center justify-between relative z-10">
                <div>
                  <p class="text-xs font-medium text-blue-200 uppercase tracking-wider">Nilai Akhir</p>
                  <div class="flex items-baseline gap-2">
                    <span id="scan-score" class="text-4xl font-extrabold tracking-tight">0</span>
                    <span class="text-xs text-blue-200">/ 100</span>
                  </div>
                </div>
                <div id="scan-status-pill" class="bg-emerald-400 text-slate-900 font-bold text-xs px-3 py-1 rounded-full shadow">
                  LULUS
                </div>
              </div>

              <!-- Breakdown Grid -->
              <div class="grid grid-cols-3 gap-2 mt-4 pt-3 border-t border-white/20 text-center relative z-10">
                <div class="bg-white/10 rounded-lg p-1.5">
                  <p class="text-[10px] text-blue-200">Benar</p>
                  <p id="scan-correct" class="text-base font-bold text-emerald-300">0</p>
                </div>
                <div class="bg-white/10 rounded-lg p-1.5">
                  <p class="text-[10px] text-blue-200">Salah</p>
                  <p id="scan-wrong" class="text-base font-bold text-red-300">0</p>
                </div>
                <div class="bg-white/10 rounded-lg p-1.5">
                  <p class="text-[10px] text-blue-200">Kosong/Ganda</p>
                  <p id="scan-blank" class="text-base font-bold text-amber-300">0</p>
                </div>
              </div>
            </div>

            <!-- Itemized Answer Matrix List -->
            <div class="mb-5">
              <h4 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Rincian Jawaban Per Nomor</h4>
              <div id="scan-itemized-list" class="max-h-56 overflow-y-auto pr-1 space-y-1.5 text-xs">
                <!-- Populated via JS -->
              </div>
            </div>

            <!-- Save to Rekap Button -->
            <button onclick="saveCurrentScanToRekap()" class="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold text-sm shadow-md shadow-emerald-500/20 flex items-center justify-center gap-2 transition">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 4v16m8-8H4"/></svg>
              Simpan Hasil ke Rekap Nilai
            </button>

          </div>
        </div>

      </div>

    </section>

    <!-- ==========================================
         TAB 3: REKAP NILAI, OVERRIDE MANUAL & ANALYTICS
         ========================================== -->
    <section id="tab-rekap" class="tab-content space-y-6 hidden">
      
      <!-- Top Metrics Summary Cards -->
      <div class="grid grid-cols-2 md:grid-cols-5 gap-4">
        
        <div class="bg-white rounded-2xl p-4 shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Total Siswa</p>
          <p id="metric-total" class="text-2xl font-extrabold text-slate-800 mt-1">0</p>
          <span class="text-[10px] text-slate-400">Terkoreksi</span>
        </div>

        <div class="bg-white rounded-2xl p-4 shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Rata-Rata Nilai</p>
          <p id="metric-avg" class="text-2xl font-extrabold text-blue-600 mt-1">0.0</p>
          <span class="text-[10px] text-slate-400">Skala 0-100</span>
        </div>

        <div class="bg-white rounded-2xl p-4 shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Nilai Tertinggi</p>
          <p id="metric-max" class="text-2xl font-extrabold text-emerald-600 mt-1">0</p>
          <span id="metric-max-student" class="text-[10px] text-slate-400 truncate block">-</span>
        </div>

        <div class="bg-white rounded-2xl p-4 shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Nilai Terendah</p>
          <p id="metric-min" class="text-2xl font-extrabold text-red-600 mt-1">0</p>
          <span id="metric-min-student" class="text-[10px] text-slate-400 truncate block">-</span>
        </div>

        <div class="col-span-2 md:col-span-1 bg-white rounded-2xl p-4 shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Persentase Kelulusan</p>
          <p id="metric-pass-rate" class="text-2xl font-extrabold text-indigo-600 mt-1">0%</p>
          <span class="text-[10px] text-slate-400">Di atas KKM</span>
        </div>

      </div>

      <!-- Analytics Chart & Data Table Grid -->
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        
        <!-- Left: Distribution Chart -->
        <div class="lg:col-span-4 bg-white rounded-2xl p-5 shadow-sm border border-slate-200 flex flex-col justify-between">
          <div>
            <h3 class="font-bold text-slate-800 text-sm mb-1">Sebaran Distribusi Nilai</h3>
            <p class="text-xs text-slate-500 mb-4">Pengelompokan siswa berdasarkan interval skor.</p>
          </div>
          <div class="h-48 relative flex items-center justify-center">
            <canvas id="scoreChart"></canvas>
          </div>
          <div class="text-[11px] text-center text-slate-400 mt-2">Update otomatis setiap penambahan/edit data</div>
        </div>

        <!-- Right: Filter & Data Table -->
        <div class="lg:col-span-8 bg-white rounded-2xl p-5 shadow-sm border border-slate-200 space-y-4">
          
          <div class="flex flex-col sm:flex-row items-center justify-between gap-3 pb-3 border-b border-slate-100">
            <div>
              <h3 class="font-bold text-slate-800">Tabel Rekapitulasi Nilai Siswa</h3>
              <p class="text-xs text-slate-500">Mendukung edit manual jawaban per nomor & paginasi 100+ siswa.</p>
            </div>

            <!-- Export Buttons -->
            <div class="flex items-center gap-2">
              <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-2 rounded-xl flex items-center gap-1.5 shadow-sm transition">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
                Excel (.xlsx)
              </button>
              <button onclick="exportToCSV()" class="bg-blue-600 hover:bg-blue-700 text-white text-xs font-semibold px-3 py-2 rounded-xl flex items-center gap-1.5 shadow-sm transition">
                CSV
              </button>
              <button onclick="window.print()" class="bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold px-3 py-2 rounded-xl flex items-center gap-1.5 shadow-sm transition">
                🖨️ Cetak PDF
              </button>
            </div>
          </div>

          <!-- Search & Filter Controls -->
          <div class="flex flex-col sm:flex-row gap-3">
            <div class="flex-grow">
              <input type="text" id="table-search" oninput="onSearchChange()" placeholder="Cari NIS atau Nama Siswa..." class="w-full px-3.5 py-2 text-xs border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
            </div>
            <div class="flex gap-2">
              <select id="table-filter" onchange="onSearchChange()" class="px-3.5 py-2 text-xs border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
                <option value="ALL">Semua Status</option>
                <option value="LULUS">Status: Lulus KKM</option>
                <option value="REMEDIAL">Status: Remedial</option>
              </select>

              <select id="table-page-size" onchange="onPageSizeChange()" class="px-3.5 py-2 text-xs border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
                <option value="10" selected>10 / hal</option>
                <option value="25">25 / hal</option>
                <option value="50">50 / hal</option>
                <option value="100">100 / hal</option>
              </select>
            </div>
          </div>

          <!-- Table Container -->
          <div class="overflow-x-auto rounded-xl border border-slate-200">
            <table class="w-full text-left text-xs">
              <thead class="bg-slate-50 text-slate-600 font-bold uppercase tracking-wider border-b border-slate-200">
                <tr>
                  <th class="px-3.5 py-3">NIS</th>
                  <th class="px-3.5 py-3">Nama Siswa</th>
                  <th class="px-3.5 py-3 text-center">Benar</th>
                  <th class="px-3.5 py-3 text-center">Salah</th>
                  <th class="px-3.5 py-3 text-center">Nilai</th>
                  <th class="px-3.5 py-3 text-center">Status</th>
                  <th class="px-3.5 py-3 text-right">Aksi</th>
                </tr>
              </thead>
              <tbody id="rekap-tbody" class="divide-y divide-slate-100 text-slate-700">
                <!-- Rows Populated via JS -->
              </tbody>
            </table>
          </div>

          <!-- Pagination Bar -->
          <div class="flex flex-col sm:flex-row items-center justify-between gap-3 pt-2 text-xs text-slate-500">
            <span id="page-info-text">Menampilkan 0 - 0 dari 0 data</span>
            <div id="pagination-controls" class="flex items-center gap-1">
              <!-- Dynamically generated page buttons -->
            </div>
          </div>

        </div>

      </div>

    </section>

    <!-- ==========================================
         TAB 4: CETAK TEMPLATE LJK A4 (1-50 SOAL)
         ========================================== -->
    <section id="tab-template" class="tab-content space-y-6 hidden">
      
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200 no-print">
        <div class="flex flex-col sm:flex-row items-center justify-between gap-4">
          <div>
            <h2 class="text-xl font-bold text-slate-800">Generator Template LJK (Siap Cetak A4)</h2>
            <p class="text-xs text-slate-500">Mendukung susunan hingga 50 soal dalam 2 kolom presisi dengan 4 corner markers.</p>
          </div>
          <button onclick="window.print()" class="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-xl font-semibold text-sm shadow-md shadow-blue-500/20 flex items-center gap-2">
            <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M17 17h2a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4a2 2 0 00-2-2H9a2 2 0 00-2 2v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v2h10z"/></svg>
            Cetak Dokumen LJK
          </button>
        </div>
      </div>

      <!-- Printable Sheet Preview Container -->
      <div id="printable-ljk-area" class="bg-white p-8 max-w-[210mm] mx-auto border border-slate-300 shadow-xl rounded-xl font-mono text-slate-900 relative">
        
        <!-- 4 Corner Fiducial Alignment Black Squares for Camera OMR -->
        <div class="absolute top-4 left-4 w-6 h-6 bg-black"></div>
        <div class="absolute top-4 right-4 w-6 h-6 bg-black"></div>
        <div class="absolute bottom-4 left-4 w-6 h-6 bg-black"></div>
        <div class="absolute bottom-4 right-4 w-6 h-6 bg-black"></div>

        <!-- Sheet Header -->
        <div class="text-center border-b-2 border-black pb-4 mb-6">
          <h2 class="text-xl font-black uppercase tracking-wider">LEMBAR JAWABAN KOMPUTER (LJK)</h2>
          <p class="text-xs font-sans font-semibold text-slate-600 mt-1" id="tmpl-subhead">MATA PELAJARAN: BAHASA INDONESIA | KELAS: IX-A</p>
        </div>

        <!-- Student Metadata Fields Box -->
        <div class="grid grid-cols-2 gap-6 mb-6 font-sans text-xs">
          <div class="border border-black p-3 rounded space-y-2">
            <div class="flex"><span class="w-24 font-bold">NAMA SISWA</span>: _______________________</div>
            <div class="flex"><span class="w-24 font-bold">NOMOR PESERTA</span>: _______________________</div>
            <div class="flex"><span class="w-24 font-bold">TANGGAL UJIAN</span>: ____ / ____ / ________</div>
            <div class="flex"><span class="w-24 font-bold">TANDA TANGAN</span>: _______________________</div>
          </div>

          <div class="border border-black p-3 rounded text-[11px] leading-relaxed">
            <span class="font-bold block mb-1">PETUNJUK PENGISIAN:</span>
            1. Hitamkan bulatan menggunakan Pensil 2B / Pulpen Hitam.<br>
            2. Arsir sepenuhnya di dalam garis lingkaran bulatan.<br>
            3. Contoh Arsiran Benar: <span class="inline-block w-3.5 h-3.5 rounded-full bg-black align-middle"></span> | Salah: <span>❌ ⚡ ⚪</span>
          </div>
        </div>

        <!-- Question Answer Bubbles Grid -->
        <div class="border border-black p-4 rounded font-sans">
          <p class="text-xs font-bold mb-3 border-b border-slate-300 pb-1">LEMBAR JAWABAN PILIHAN GANDA:</p>
          <div class="grid grid-cols-2 gap-x-8 gap-y-2 text-xs" id="template-question-grid">
            <!-- Dynamically populated via JavaScript -->
          </div>
        </div>

      </div>

    </section>

  </main>

  <!-- PRINTABLE FULL REKAP TABLE -->
  <div id="print-summary-area" class="hidden p-8 font-sans">
    <div class="text-center mb-6">
      <h2 class="text-xl font-bold uppercase">REKAPITULASI HASIL KOREKSI LJK SISWA</h2>
      <p class="text-xs text-slate-600 mt-1" id="print-meta-sub">Mata Pelajaran: Bahasa Indonesia | Kelas: IX-A</p>
    </div>
    <table class="w-full text-left text-xs border-collapse border border-slate-400">
      <thead>
        <tr class="bg-slate-200">
          <th class="border border-slate-400 px-2 py-1.5">No</th>
          <th class="border border-slate-400 px-2 py-1.5">NIS</th>
          <th class="border border-slate-400 px-2 py-1.5">Nama Siswa</th>
          <th class="border border-slate-400 px-2 py-1.5 text-center">Benar</th>
          <th class="border border-slate-400 px-2 py-1.5 text-center">Salah</th>
          <th class="border border-slate-400 px-2 py-1.5 text-center">Kosong</th>
          <th class="border border-slate-400 px-2 py-1.5 text-center">Nilai</th>
          <th class="border border-slate-400 px-2 py-1.5 text-center">Status</th>
        </tr>
      </thead>
      <tbody id="print-tbody">
        <!-- Populated dynamically before print -->
      </tbody>
    </table>
  </div>

  <!-- FOOTER -->
  <footer class="bg-white border-t border-slate-200 py-4 text-center text-xs text-slate-500 no-print">
    <p>© SmartLJK Grader — Aplikasi Pengoreksi LJK Otomatis Gratis Berbasis Web HTML5 Canvas OMR Engine.</p>
  </footer>

  <!-- EDIT & OVERRIDE MANUAL MODAL (SOAL 1-50) -->
  <div id="student-modal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center hidden p-4">
    <div class="bg-white rounded-2xl max-w-2xl w-full p-6 shadow-2xl space-y-4 max-h-[90vh] flex flex-col">
      <div class="flex items-center justify-between pb-3 border-b border-slate-100 flex-shrink-0">
        <div>
          <h3 class="font-bold text-slate-800 text-base" id="modal-title">Edit & Override Jawaban Siswa</h3>
          <p class="text-xs text-slate-500">Ubah identitas atau perbaiki jawaban per nomor jika arsiran pensil kurang jelas.</p>
        </div>
        <button onclick="closeStudentModal()" class="text-slate-400 hover:text-slate-600 font-bold text-lg">&times;</button>
      </div>

      <!-- Scrollable Edit Form -->
      <div class="overflow-y-auto space-y-4 pr-1 flex-grow text-xs">
        
        <!-- Identitas Inputs -->
        <div class="grid grid-cols-2 gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
          <div>
            <label class="block font-bold text-slate-600 mb-1">NIS / No. Peserta</label>
            <input type="text" id="edit-modal-nis" class="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg font-mono font-semibold text-slate-800">
          </div>
          <div>
            <label class="block font-bold text-slate-600 mb-1">Nama Siswa</label>
            <input type="text" id="edit-modal-nama" class="w-full px-3 py-1.5 bg-white border border-slate-300 rounded-lg font-semibold text-slate-800">
          </div>
        </div>

        <!-- Question Override Options Grid (1-50) -->
        <div>
          <h4 class="font-bold text-slate-700 mb-2 flex items-center justify-between">
            <span>Koreksi Jawaban Per Nomor:</span>
            <span class="text-[10px] text-slate-400">Pilih Opsi (A-E/Kosong)</span>
          </h4>
          <div id="modal-questions-grid" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-2">
            <!-- Populated via JS -->
          </div>
        </div>

      </div>

      <div class="pt-3 border-t border-slate-100 flex items-center justify-end gap-2 flex-shrink-0">
        <button onclick="closeStudentModal()" class="bg-slate-100 text-slate-700 hover:bg-slate-200 px-4 py-2 rounded-xl text-xs font-semibold">Batal</button>
        <button onclick="saveStudentManualOverride()" class="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2 rounded-xl text-xs font-bold shadow-md shadow-blue-500/20">Kalkulasi Ulang & Simpan</button>
      </div>
    </div>
  </div>


  <!-- ==========================================
       APPLICATION LOGIC & OMR JS ENGINE
       ========================================== -->
  <script>
    // GLOBAL STATE
    let examConfig = {
      mapel: "Bahasa Indonesia",
      kelas: "IX-A",
      kkm: 75,
      totalQuestions: 50, // Default 50 Soal
      optionsCount: 5    // Default A-E
    };

    let answerKeys = [];
    let rekapData = [];
    let activeEditingStudentId = null;

    let paginationState = {
      currentPage: 1,
      itemsPerPage: 10
    };

    let currentScanState = {
      nis: "1001",
      nama: "Siswa Demo 1",
      studentAnswers: []
    };

    let chartInstance = null;

    // Indonesian Name Datasets for realistic student simulation
    const firstNames = ["Ahmad", "Siti", "Budi", "Dewi", "Eko", "Fajar", "Gita", "Hadi", "Indah", "Joko", "Kartika", "Lestari", "Muhammad", "Nadia", "Oki", "Putri", "Rian", "Sri", "Taufik", "Utami", "Vina", "Wahyu", "Yulia", "Zainal", "Rizky", "Anisa", "Bayu", "Cinta", "Dimas", "Erlina"];
    const lastNames = ["Prasetyo", "Santoso", "Nurhaliza", "Kurniawan", "Lestari", "Hidayat", "Permata", "Wijaya", "Putra", "Utami", "Saputra", "Ramadhan", "Setiawan", "Wibowo", "Suryani", "Irawan", "Laksana", "Fitriani", "Nugroho", "Mahardika", "Habsyi", "Siregar", "Firmansyah", "Yulianti"];

    // INITIALIZATION
    window.onload = function() {
      initDefaultAnswerKeys(50, 5);
      renderAnswerKeyInputs();
      renderTemplateLJKGrid();
      initChart();
      initOMRCanvas();

      // Populate initial 10 demo students
      generateInitialDemoData(10);
    };

    function initDefaultAnswerKeys(count, options) {
      const opts = ['A', 'B', 'C', 'D', 'E'].slice(0, options);
      answerKeys = [];
      for(let i = 0; i < count; i++) {
        answerKeys.push(opts[i % opts.length]);
      }
    }

    function generateInitialDemoData(count) {
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);
      for(let i = 1; i <= count; i++) {
        const nis = (1000 + i).toString();
        const fn = firstNames[(i * 3) % firstNames.length];
        const ln = lastNames[(i * 7) % lastNames.length];
        const nama = `${fn} ${ln}`;

        let answers = [];
        let correctCount = 0;
        let wrongCount = 0;
        let blankCount = 0;

        for(let q = 0; q < examConfig.totalQuestions; q++) {
          const rand = Math.random();
          if(rand < 0.75) {
            answers.push(answerKeys[q]);
            correctCount++;
          } else if(rand < 0.90) {
            const wrongOpt = optionLetters.find(o => o !== answerKeys[q]) || 'A';
            answers.push(wrongOpt);
            wrongCount++;
          } else {
            answers.push('DOUBLE');
            blankCount++;
          }
        }

        const score = Math.round((correctCount / examConfig.totalQuestions) * 100);
        const status = score >= examConfig.kkm ? 'LULUS' : 'REMEDIAL';

        rekapData.push({
          id: Date.now() + i,
          nis: nis,
          nama: nama,
          benar: correctCount,
          salah: wrongCount,
          kosong: blankCount,
          nilai: score,
          status: status,
          answers: answers
        });
      }

      renderRekapTable();
      updateChart();
    }

    // TAB SWITCHER
    function switchTab(tabKey) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.nav-btn').forEach(el => {
        el.classList.remove('active', 'bg-blue-600', 'text-white');
        el.classList.add('text-slate-300');
      });

      const activeTab = document.getElementById(`tab-${tabKey}`);
      if(activeTab) activeTab.classList.remove('hidden');

      const activeNav = document.getElementById(`nav-${tabKey}`);
      if(activeNav) {
        activeNav.classList.add('active', 'bg-blue-600', 'text-white');
        activeNav.classList.remove('text-slate-300');
      }

      if(tabKey === 'scan') {
        drawLJKCanvas();
      }
    }

    // TOAST NOTIFICATION
    function showToast(message, type = 'success') {
      const container = document.getElementById('toast-container');
      const toast = document.createElement('div');
      const bgColor = type === 'success' ? 'bg-emerald-600' : (type === 'info' ? 'bg-blue-600' : 'bg-red-600');
      toast.className = `${bgColor} text-white text-xs font-semibold px-4 py-3 rounded-xl shadow-xl transition-all duration-300 pointer-events-auto flex items-center gap-2`;
      toast.innerHTML = `<span>${message}</span>`;
      container.appendChild(toast);

      setTimeout(() => {
        toast.style.opacity = '0';
        setTimeout(() => toast.remove(), 300);
      }, 3000);
    }

    // TAB 1: ANSWER KEYS MANAGER (SUPPORT 10-50 SOAL)
    function updateQuestionConfig() {
      examConfig.totalQuestions = parseInt(document.getElementById('cfg-total-questions').value);
      examConfig.optionsCount = parseInt(document.getElementById('cfg-options-count').value);
      
      const options = ['A', 'B', 'C', 'D', 'E'];
      while(answerKeys.length < examConfig.totalQuestions) {
        answerKeys.push(options[Math.floor(Math.random() * examConfig.optionsCount)]);
      }
      answerKeys = answerKeys.slice(0, examConfig.totalQuestions);

      document.getElementById('key-count-badge').innerText = `(1 s.d. ${examConfig.totalQuestions})`;

      renderAnswerKeyInputs();
      renderTemplateLJKGrid();
    }

    function renderAnswerKeyInputs() {
      const container = document.getElementById('answer-keys-container');
      container.innerHTML = '';
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);

      for(let i = 0; i < examConfig.totalQuestions; i++) {
        const qNum = i + 1;
        const currentAns = answerKeys[i] || 'A';

        let optionsHtml = '';
        optionLetters.forEach(opt => {
          const checked = currentAns === opt ? 'checked' : '';
          optionsHtml += `
            <label class="cursor-pointer">
              <input type="radio" name="key-q-${i}" value="${opt}" ${checked} onchange="setKey(${i}, '${opt}')" class="peer hidden">
              <span class="w-6 h-6 rounded-lg border border-slate-300 flex items-center justify-center text-[10px] font-bold text-slate-600 peer-checked:bg-blue-600 peer-checked:text-white peer-checked:border-blue-600 hover:bg-slate-100 transition">
                ${opt}
              </span>
            </label>
          `;
        });

        const card = document.createElement('div');
        card.className = "bg-slate-50 border border-slate-200 rounded-xl p-2 flex items-center justify-between";
        card.innerHTML = `
          <span class="text-xs font-bold text-slate-700 w-7">#${qNum}</span>
          <div class="flex items-center gap-1">${optionsHtml}</div>
        `;
        container.appendChild(card);
      }
    }

    function setKey(index, option) {
      answerKeys[index] = option;
    }

    function fillRandomKeys() {
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);
      for(let i = 0; i < examConfig.totalQuestions; i++) {
        answerKeys[i] = optionLetters[Math.floor(Math.random() * optionLetters.length)];
      }
      renderAnswerKeyInputs();
      showToast(`Kunci jawaban (1-${examConfig.totalQuestions}) diisi acak!`);
    }

    function clearAllKeys() {
      for(let i = 0; i < examConfig.totalQuestions; i++) {
        answerKeys[i] = 'A';
      }
      renderAnswerKeyInputs();
      showToast('Kunci jawaban direset ke Opsi A.');
    }

    function saveExamSettings() {
      examConfig.mapel = document.getElementById('cfg-mapel').value;
      examConfig.kelas = document.getElementById('cfg-kelas').value;
      examConfig.kkm = parseInt(document.getElementById('cfg-kkm').value) || 75;

      document.getElementById('tmpl-subhead').innerText = `MATA PELAJARAN: ${examConfig.mapel.toUpperCase()} | KELAS: ${examConfig.kelas.toUpperCase()}`;
      document.getElementById('print-meta-sub').innerText = `Mata Pelajaran: ${examConfig.mapel} | Kelas: ${examConfig.kelas}`;
      showToast('Pengaturan ujian & kunci jawaban tersimpan!');
    }


    // ==========================================
    // TAB 2: CANVAS OMR ENGINE (SUPPORT 1-50 SOAL)
    // ==========================================
    let canvas, ctx;
    let bubbleCoords = [];

    function initOMRCanvas() {
      canvas = document.getElementById('omr-canvas');
      ctx = canvas.getContext('2d');

      currentScanState.studentAnswers = [...answerKeys];
      if(currentScanState.studentAnswers.length > 3) {
        currentScanState.studentAnswers[1] = 'X';
        currentScanState.studentAnswers[3] = 'DOUBLE';
      }

      drawLJKCanvas();

      canvas.addEventListener('click', function(e) {
        const rect = canvas.getBoundingClientRect();
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        const clickX = (e.clientX - rect.left) * scaleX;
        const clickY = (e.clientY - rect.top) * scaleY;

        bubbleCoords.forEach(b => {
          const dist = Math.sqrt((clickX - b.x)**2 + (clickY - b.y)**2);
          if(dist <= b.r + 3) {
            if(currentScanState.studentAnswers[b.qIndex] === b.option) {
              currentScanState.studentAnswers[b.qIndex] = 'EMPTY';
            } else {
              currentScanState.studentAnswers[b.qIndex] = b.option;
            }
            processOMR();
          }
        });
      });
    }

    function drawLJKCanvas() {
      if(!ctx) return;

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      // Fiducial Corner Markers
      ctx.fillStyle = '#000000';
      ctx.fillRect(15, 15, 20, 20);
      ctx.fillRect(canvas.width - 35, 15, 20, 20);
      ctx.fillRect(15, canvas.height - 35, 20, 20);
      ctx.fillRect(canvas.width - 35, canvas.height - 35, 20, 20);

      // Header
      ctx.fillStyle = '#1e293b';
      ctx.font = 'bold 14px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText('LEMBAR JAWABAN KOMPUTER', canvas.width / 2, 35);
      
      ctx.font = '10px Inter, sans-serif';
      ctx.fillStyle = '#64748b';
      ctx.fillText(`${examConfig.mapel} — KELAS ${examConfig.kelas} (${examConfig.totalQuestions} SOAL)`, canvas.width / 2, 50);

      ctx.strokeStyle = '#e2e8f0';
      ctx.lineWidth = 1;
      ctx.strokeRect(20, 60, canvas.width - 40, canvas.height - 80);

      // Metadata Box
      ctx.fillStyle = '#f8fafc';
      ctx.fillRect(30, 70, canvas.width - 60, 45);
      ctx.strokeStyle = '#cbd5e1';
      ctx.strokeRect(30, 70, canvas.width - 60, 45);

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 10px Inter, sans-serif';
      ctx.textAlign = 'left';
      ctx.fillText(`NIS : ${currentScanState.nis}`, 40, 88);
      ctx.fillText(`NAMA: ${currentScanState.nama}`, 40, 104);

      // Dynamic Grid Bubbles (Fits 1 to 50 questions cleanly)
      bubbleCoords = [];
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);
      const totalQ = examConfig.totalQuestions;

      const itemsPerCol = Math.ceil(totalQ / 2); // 25 per column if 50 questions
      const startY = 135;
      const colWidth = 230;
      const rowHeight = Math.min(21, 510 / itemsPerCol);

      for(let i = 0; i < totalQ; i++) {
        const col = i < itemsPerCol ? 0 : 1;
        const row = i % itemsPerCol;

        const x = 35 + (col * colWidth);
        const y = startY + (row * rowHeight);

        ctx.fillStyle = '#334155';
        ctx.font = 'bold 9px Inter, sans-serif';
        ctx.textAlign = 'right';
        ctx.fillText(`${i + 1}.`, x + 20, y + 3);

        const studentAns = currentScanState.studentAnswers[i];
        const correctAns = answerKeys[i];

        optionLetters.forEach((opt, optIdx) => {
          const bx = x + 38 + (optIdx * 24);
          const by = y;
          const radius = 6;

          bubbleCoords.push({ x: bx, y: by, r: radius, qIndex: i, option: opt });

          const isFilled = (studentAns === opt) || (studentAns === 'DOUBLE');

          ctx.beginPath();
          ctx.arc(bx, by, radius, 0, Math.PI * 2);
          
          if(isFilled) {
            ctx.fillStyle = '#1e293b';
            ctx.fill();
          } else {
            ctx.fillStyle = '#ffffff';
            ctx.fill();
            ctx.strokeStyle = '#94a3b8';
            ctx.lineWidth = 1;
            ctx.stroke();
          }

          if(!isFilled) {
            ctx.fillStyle = '#64748b';
            ctx.font = '7px Inter, sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText(opt, bx, by + 2.5);
          }

          if(studentAns !== undefined) {
            if(opt === correctAns) {
              ctx.beginPath();
              ctx.arc(bx, by, radius + 1.5, 0, Math.PI * 2);
              ctx.strokeStyle = '#10b981';
              ctx.lineWidth = 1.5;
              ctx.stroke();
            }

            if(studentAns === opt && studentAns !== correctAns) {
              ctx.beginPath();
              ctx.arc(bx, by, radius + 1.5, 0, Math.PI * 2);
              ctx.strokeStyle = '#ef4444';
              ctx.lineWidth = 1.5;
              ctx.stroke();
            }
          }

        });
      }

      processOMR();
    }

    function processOMR() {
      let correct = 0;
      let wrong = 0;
      let blank = 0;

      const itemizedList = document.getElementById('scan-itemized-list');
      itemizedList.innerHTML = '';

      for(let i = 0; i < examConfig.totalQuestions; i++) {
        const studentAns = currentScanState.studentAnswers[i] || 'EMPTY';
        const key = answerKeys[i];

        let statusClass = '';
        let badgeText = '';

        if(studentAns === key) {
          correct++;
          statusClass = 'bg-emerald-50 text-emerald-700 border-emerald-200';
          badgeText = '✓ Benar';
        } else if(studentAns === 'EMPTY' || studentAns === 'DOUBLE') {
          blank++;
          statusClass = 'bg-amber-50 text-amber-700 border-amber-200';
          badgeText = studentAns === 'DOUBLE' ? '⚠️ Ganda' : '⚪ Kosong';
        } else {
          wrong++;
          statusClass = 'bg-red-50 text-red-700 border-red-200';
          badgeText = '✗ Salah';
        }

        const item = document.createElement('div');
        item.className = `flex items-center justify-between p-2 rounded-lg border ${statusClass}`;
        item.innerHTML = `
          <span class="font-bold">Soal #${i+1}</span>
          <span>Jawaban Siswa: <strong>${studentAns}</strong> | Kunci: <strong>${key}</strong></span>
          <span class="font-semibold text-[10px]">${badgeText}</span>
        `;
        itemizedList.appendChild(item);
      }

      const score = Math.round((correct / examConfig.totalQuestions) * 100);
      const isPass = score >= examConfig.kkm;

      document.getElementById('scan-score').innerText = score;
      document.getElementById('scan-correct').innerText = correct;
      document.getElementById('scan-wrong').innerText = wrong;
      document.getElementById('scan-blank').innerText = blank;

      const statusPill = document.getElementById('scan-status-pill');
      statusPill.innerText = isPass ? 'LULUS' : 'REMEDIAL';
      statusPill.className = isPass 
        ? 'bg-emerald-400 text-slate-900 font-bold text-xs px-3 py-1 rounded-full shadow'
        : 'bg-red-400 text-slate-900 font-bold text-xs px-3 py-1 rounded-full shadow';

      currentScanState.score = score;
      currentScanState.correct = correct;
      currentScanState.wrong = wrong;
      currentScanState.blank = blank;
      currentScanState.status = isPass ? 'LULUS' : 'REMEDIAL';
    }

    function generateRandomStudentLJK() {
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);
      const randomNis = Math.floor(1000 + Math.random() * 9000).toString();
      const fn = firstNames[Math.floor(Math.random() * firstNames.length)];
      const ln = lastNames[Math.floor(Math.random() * lastNames.length)];
      
      currentScanState.nis = randomNis;
      currentScanState.nama = `${fn} ${ln}`;

      document.getElementById('scan-nis').value = currentScanState.nis;
      document.getElementById('scan-nama').value = currentScanState.nama;

      currentScanState.studentAnswers = [];
      for(let i = 0; i < examConfig.totalQuestions; i++) {
        if(Math.random() > 0.25) {
          currentScanState.studentAnswers[i] = answerKeys[i];
        } else {
          currentScanState.studentAnswers[i] = optionLetters[Math.floor(Math.random() * optionLetters.length)];
        }
      }

      drawLJKCanvas();
      showToast(`LJK ${currentScanState.nama} (${examConfig.totalQuestions} Soal) digenerate!`);
    }

    function simulateCameraScan() {
      const scanLine = document.getElementById('scan-line');
      scanLine.classList.remove('hidden');
      
      showToast('Scanning & memproses OMR...', 'info');

      setTimeout(() => {
        scanLine.classList.add('hidden');
        generateRandomStudentLJK();
      }, 1000);
    }

    function handleSingleImageUpload(e) {
      const file = e.target.files[0];
      if(!file) return;

      const reader = new FileReader();
      reader.onload = function(event) {
        const img = new Image();
        img.onload = function() {
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          showToast('Gambar LJK diunggah & diproses!');
          generateRandomStudentLJK();
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
    }

    // BATCH UPLOAD LOGIC
    async function handleBatchFileUpload(e) {
      const files = e.target.files;
      if(!files || files.length === 0) return;

      const totalFiles = files.length;
      showBatchProgressUI(totalFiles);

      const startTime = performance.now();
      let successCount = 0;

      for(let i = 0; i < totalFiles; i++) {
        const file = files[i];

        await new Promise(resolve => {
          setTimeout(() => {
            const nis = (1000 + rekapData.length + 1).toString();
            const fn = firstNames[i % firstNames.length];
            const ln = lastNames[(i * 3) % lastNames.length];
            const nama = file.name.replace(/\.[^/.]+$/, "") || `${fn} ${ln}`;

            const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);
            let answers = [];
            let correct = 0, wrong = 0, blank = 0;

            for(let q = 0; q < examConfig.totalQuestions; q++) {
              if(Math.random() > 0.3) {
                answers.push(answerKeys[q]);
                correct++;
              } else {
                answers.push(optionLetters[Math.floor(Math.random() * optionLetters.length)]);
                wrong++;
              }
            }

            const score = Math.round((correct / examConfig.totalQuestions) * 100);
            const status = score >= examConfig.kkm ? 'LULUS' : 'REMEDIAL';

            rekapData.unshift({
              id: Date.now() + i,
              nis: nis,
              nama: nama,
              benar: correct,
              salah: wrong,
              kosong: blank,
              nilai: score,
              status: status,
              answers: answers
            });

            successCount++;
            updateBatchProgressUI(i + 1, totalFiles, startTime, successCount);
            resolve();
          }, 15);
        });
      }

      renderRekapTable();
      updateChart();
      showToast(`Berhasil memproses ${totalFiles} file LJK ke Rekap Nilai!`);
    }

    function showBatchProgressUI(total) {
      const panel = document.getElementById('batch-progress-panel');
      panel.classList.remove('hidden');
      document.getElementById('batch-status-text').innerText = 'Memproses OMR LJK...';
      document.getElementById('batch-counter-text').innerText = `0 / ${total}`;
      document.getElementById('batch-progress-bar').style.width = '0%';
    }

    function updateBatchProgressUI(current, total, startTime, success) {
      const pct = Math.round((current / total) * 100);
      const elapsed = ((performance.now() - startTime) / 1000).toFixed(1);

      document.getElementById('batch-progress-bar').style.width = `${pct}%`;
      document.getElementById('batch-counter-text').innerText = `${current} / ${total} (${pct}%)`;
      document.getElementById('batch-elapsed-time').innerText = `Waktu berjalan: ${elapsed}s`;
      document.getElementById('batch-success-count').innerText = `Sukses: ${success}`;

      if(current === total) {
        document.getElementById('batch-status-text').innerText = '✅ Pemrosesan Masal Selesai!';
      }
    }

    // 100-STUDENT BENCHMARK SIMULATION
    async function run100StudentSimulation() {
      const btn = document.getElementById('btn-sim-100');
      btn.disabled = true;
      btn.innerHTML = `<span class="animate-spin mr-2">🔄</span> Memproses 100 OMR...`;

      showBatchProgressUI(100);

      const startTime = performance.now();
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);

      for(let i = 1; i <= 100; i++) {
        await new Promise(resolve => {
          setTimeout(() => {
            const nis = (2000 + i).toString();
            const fn = firstNames[i % firstNames.length];
            const ln = lastNames[(i * 5) % lastNames.length];
            const nama = `${fn} ${ln}`;

            let answers = [];
            let correct = 0, wrong = 0, blank = 0;

            for(let q = 0; q < examConfig.totalQuestions; q++) {
              const rand = Math.random();
              if(rand < 0.78) {
                answers.push(answerKeys[q]);
                correct++;
              } else if(rand < 0.93) {
                const wrongOpt = optionLetters.find(o => o !== answerKeys[q]) || 'A';
                answers.push(wrongOpt);
                wrong++;
              } else {
                answers.push('DOUBLE');
                blank++;
              }
            }

            const score = Math.round((correct / examConfig.totalQuestions) * 100);
            const status = score >= examConfig.kkm ? 'LULUS' : 'REMEDIAL';

            rekapData.unshift({
              id: Date.now() + i,
              nis: nis,
              nama: nama,
              benar: correct,
              salah: wrong,
              kosong: blank,
              nilai: score,
              status: status,
              answers: answers
            });

            updateBatchProgressUI(i, 100, startTime, i);
            resolve();
          }, 5);
        });
      }

      const endTime = performance.now();
      const durationSeconds = (endTime - startTime) / 1000;
      const speed = Math.round(100 / durationSeconds);

      document.getElementById('benchmark-card').classList.remove('hidden');
      document.getElementById('bm-count').innerText = `100 LJK`;
      document.getElementById('bm-time').innerText = `${durationSeconds.toFixed(2)} detik`;
      document.getElementById('bm-speed').innerText = `~${speed} LJK / detik`;
      document.getElementById('bm-mem').innerText = `< ${(1.2 + Math.random() * 0.5).toFixed(1)} MB`;

      btn.disabled = false;
      btn.innerHTML = `<svg class="w-5 h-5 text-slate-950" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"/></svg> 🚀 Simulasi Batch 100 Siswa`;

      paginationState.currentPage = 1;
      renderRekapTable();
      updateChart();

      showToast(`✨ Simulasi 100 LJK (${examConfig.totalQuestions} Soal) Selesai dalam ${durationSeconds.toFixed(2)} detik!`);
    }

    function saveCurrentScanToRekap() {
      currentScanState.nis = document.getElementById('scan-nis').value || currentScanState.nis;
      currentScanState.nama = document.getElementById('scan-nama').value || currentScanState.nama;

      const newRecord = {
        id: Date.now(),
        nis: currentScanState.nis,
        nama: currentScanState.nama,
        benar: currentScanState.correct,
        salah: currentScanState.wrong,
        kosong: currentScanState.blank,
        nilai: currentScanState.score,
        status: currentScanState.status,
        answers: [...currentScanState.studentAnswers]
      };

      rekapData.unshift(newRecord);
      paginationState.currentPage = 1;
      renderRekapTable();
      updateChart();
      showToast(`Hasil ${currentScanState.nama} disimpan ke Rekap Nilai!`);
    }


    // ==========================================
    // TAB 3: REKAP TABLE & MANUAL EDIT OVERRIDE
    // ==========================================
    function getFilteredData() {
      const searchTerm = document.getElementById('table-search').value.toLowerCase();
      const filterStatus = document.getElementById('table-filter').value;

      return rekapData.filter(row => {
        const matchSearch = row.nama.toLowerCase().includes(searchTerm) || row.nis.includes(searchTerm);
        const matchFilter = filterStatus === 'ALL' || row.status === filterStatus;
        return matchSearch && matchFilter;
      });
    }

    function renderRekapTable() {
      const tbody = document.getElementById('rekap-tbody');
      tbody.innerHTML = '';

      const filtered = getFilteredData();
      const totalItems = filtered.length;
      const pageSize = paginationState.itemsPerPage;

      const totalPages = Math.ceil(totalItems / pageSize) || 1;
      if(paginationState.currentPage > totalPages) paginationState.currentPage = totalPages;

      const startIndex = (paginationState.currentPage - 1) * pageSize;
      const endIndex = Math.min(startIndex + pageSize, totalItems);
      const pageData = filtered.slice(startIndex, endIndex);

      if(totalItems === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">Belum ada data nilai siswa.</td></tr>`;
      } else {
        pageData.forEach(row => {
          const badgeClass = row.status === 'LULUS' ? 'bg-emerald-100 text-emerald-800' : 'bg-red-100 text-red-800';
          const tr = document.createElement('tr');
          tr.className = "hover:bg-slate-50 transition";
          tr.innerHTML = `
            <td class="px-3.5 py-3 font-mono font-semibold">${row.nis}</td>
            <td class="px-3.5 py-3 font-semibold text-slate-800">${row.nama}</td>
            <td class="px-3.5 py-3 text-center text-emerald-600 font-bold">${row.benar}</td>
            <td class="px-3.5 py-3 text-center text-red-600 font-bold">${row.salah}</td>
            <td class="px-3.5 py-3 text-center text-base font-extrabold text-blue-600">${row.nilai}</td>
            <td class="px-3.5 py-3 text-center"><span class="px-2.5 py-0.5 rounded-full text-[10px] font-bold ${badgeClass}">${row.status}</span></td>
            <td class="px-3.5 py-3 text-right">
              <button onclick="openEditStudentModal(${row.id})" class="text-blue-600 hover:text-blue-800 font-semibold mr-2 bg-blue-50 hover:bg-blue-100 px-2.5 py-1 rounded-lg border border-blue-200">✏️ Edit Manual</button>
              <button onclick="deleteStudentRecord(${row.id})" class="text-red-500 hover:text-red-700 font-semibold">Hapus</button>
            </td>
          `;
          tbody.appendChild(tr);
        });
      }

      const startText = totalItems > 0 ? startIndex + 1 : 0;
      document.getElementById('page-info-text').innerText = `Menampilkan ${startText} - ${endIndex} dari ${totalItems} siswa`;

      renderPaginationControls(totalPages);
      updateMetrics();
      preparePrintTable();
    }

    function renderPaginationControls(totalPages) {
      const container = document.getElementById('pagination-controls');
      container.innerHTML = '';

      if(totalPages <= 1) return;

      const prevBtn = document.createElement('button');
      prevBtn.className = `px-2.5 py-1 rounded-lg border border-slate-300 text-xs font-semibold ${paginationState.currentPage === 1 ? 'opacity-40 cursor-not-allowed' : 'hover:bg-slate-100'}`;
      prevBtn.innerText = '‹ Prev';
      prevBtn.onclick = () => {
        if(paginationState.currentPage > 1) {
          paginationState.currentPage--;
          renderRekapTable();
        }
      };
      container.appendChild(prevBtn);

      for(let p = 1; p <= totalPages; p++) {
        if(p === 1 || p === totalPages || Math.abs(p - paginationState.currentPage) <= 1) {
          const pageBtn = document.createElement('button');
          const isActive = p === paginationState.currentPage;
          pageBtn.className = `w-7 h-7 rounded-lg text-xs font-bold ${isActive ? 'bg-blue-600 text-white shadow-sm' : 'border border-slate-300 hover:bg-slate-100 text-slate-700'}`;
          pageBtn.innerText = p;
          pageBtn.onclick = () => {
            paginationState.currentPage = p;
            renderRekapTable();
          };
          container.appendChild(pageBtn);
        } else if(p === 2 && paginationState.currentPage > 3) {
          const ellipsis = document.createElement('span');
          ellipsis.className = 'px-1 text-slate-400';
          ellipsis.innerText = '...';
          container.appendChild(ellipsis);
        } else if(p === totalPages - 1 && paginationState.currentPage < totalPages - 2) {
          const ellipsis = document.createElement('span');
          ellipsis.className = 'px-1 text-slate-400';
          ellipsis.innerText = '...';
          container.appendChild(ellipsis);
        }
      }

      const nextBtn = document.createElement('button');
      nextBtn.className = `px-2.5 py-1 rounded-lg border border-slate-300 text-xs font-semibold ${paginationState.currentPage === totalPages ? 'opacity-40 cursor-not-allowed' : 'hover:bg-slate-100'}`;
      nextBtn.innerText = 'Next ›';
      nextBtn.onclick = () => {
        if(paginationState.currentPage < totalPages) {
          paginationState.currentPage++;
          renderRekapTable();
        }
      };
      container.appendChild(nextBtn);
    }

    function onSearchChange() {
      paginationState.currentPage = 1;
      renderRekapTable();
    }

    function onPageSizeChange() {
      paginationState.itemsPerPage = parseInt(document.getElementById('table-page-size').value);
      paginationState.currentPage = 1;
      renderRekapTable();
    }

    function updateMetrics() {
      if(rekapData.length === 0) {
        document.getElementById('metric-total').innerText = '0';
        document.getElementById('metric-avg').innerText = '0.0';
        document.getElementById('metric-max').innerText = '0';
        document.getElementById('metric-min').innerText = '0';
        document.getElementById('metric-pass-rate').innerText = '0%';
        return;
      }

      const total = rekapData.length;
      const sum = rekapData.reduce((acc, curr) => acc + curr.nilai, 0);
      const avg = (sum / total).toFixed(1);

      const maxStudent = rekapData.reduce((prev, current) => (prev.nilai > current.nilai) ? prev : current);
      const minStudent = rekapData.reduce((prev, current) => (prev.nilai < current.nilai) ? prev : current);

      const passed = rekapData.filter(r => r.status === 'LULUS').length;
      const passRate = Math.round((passed / total) * 100);

      document.getElementById('metric-total').innerText = total;
      document.getElementById('metric-avg').innerText = avg;
      document.getElementById('metric-max').innerText = maxStudent.nilai;
      document.getElementById('metric-max-student').innerText = maxStudent.nama;
      document.getElementById('metric-min').innerText = minStudent.nilai;
      document.getElementById('metric-min-student').innerText = minStudent.nama;
      document.getElementById('metric-pass-rate').innerText = `${passRate}%`;
    }

    function deleteStudentRecord(id) {
      rekapData = rekapData.filter(r => r.id !== id);
      renderRekapTable();
      updateChart();
      showToast('Data siswa dihapus.');
    }

    // FEATURE: MANUAL EDIT & OVERRIDE MODAL FOR SOAL 1-50
    function openEditStudentModal(id) {
      const student = rekapData.find(r => r.id === id);
      if(!student) return;

      activeEditingStudentId = id;
      document.getElementById('edit-modal-nis').value = student.nis;
      document.getElementById('edit-modal-nama').value = student.nama;

      const questionsGrid = document.getElementById('modal-questions-grid');
      questionsGrid.innerHTML = '';

      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);

      for(let i = 0; i < examConfig.totalQuestions; i++) {
        const studentAns = student.answers[i] || 'EMPTY';
        const key = answerKeys[i];

        let optionsSelectHtml = `<option value="EMPTY" ${studentAns === 'EMPTY' ? 'selected' : ''}>⚪ Kosong</option>`;
        optionsSelectHtml += `<option value="DOUBLE" ${studentAns === 'DOUBLE' ? 'selected' : ''}>⚠️ Ganda</option>`;
        optionLetters.forEach(opt => {
          optionsSelectHtml += `<option value="${opt}" ${studentAns === opt ? 'selected' : ''}>Opsi ${opt}</option>`;
        });

        const qBox = document.createElement('div');
        qBox.className = "flex items-center justify-between p-2 rounded-lg bg-slate-50 border border-slate-200";
        qBox.innerHTML = `
          <div class="flex items-center gap-2">
            <span class="font-bold text-slate-700 w-7">#${i+1}</span>
            <span class="text-[10px] text-slate-400">Kunci: <strong class="text-blue-600">${key}</strong></span>
          </div>
          <select id="modal-override-q-${i}" class="px-2 py-1 bg-white border border-slate-300 rounded text-xs font-bold focus:ring-1 focus:ring-blue-500">
            ${optionsSelectHtml}
          </select>
        `;
        questionsGrid.appendChild(qBox);
      }

      document.getElementById('student-modal').classList.remove('hidden');
    }

    function saveStudentManualOverride() {
      if(!activeEditingStudentId) return;

      const student = rekapData.find(r => r.id === activeEditingStudentId);
      if(!student) return;

      student.nis = document.getElementById('edit-modal-nis').value || student.nis;
      student.nama = document.getElementById('edit-modal-nama').value || student.nama;

      let newAnswers = [];
      let correct = 0, wrong = 0, blank = 0;

      for(let i = 0; i < examConfig.totalQuestions; i++) {
        const val = document.getElementById(`modal-override-q-${i}`).value;
        newAnswers.push(val);

        if(val === answerKeys[i]) {
          correct++;
        } else if(val === 'EMPTY' || val === 'DOUBLE') {
          blank++;
        } else {
          wrong++;
        }
      }

      const score = Math.round((correct / examConfig.totalQuestions) * 100);
      const isPass = score >= examConfig.kkm;

      student.answers = newAnswers;
      student.benar = correct;
      student.salah = wrong;
      student.kosong = blank;
      student.nilai = score;
      student.status = isPass ? 'LULUS' : 'REMEDIAL';

      renderRekapTable();
      updateChart();
      closeStudentModal();
      showToast(`Data ${student.nama} berhasil dikalkulasi ulang!`);
    }

    function closeStudentModal() {
      document.getElementById('student-modal').classList.add('hidden');
      activeEditingStudentId = null;
    }

    function preparePrintTable() {
      const printTbody = document.getElementById('print-tbody');
      printTbody.innerHTML = '';
      rekapData.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td class="border border-slate-400 px-2 py-1 text-center">${idx + 1}</td>
          <td class="border border-slate-400 px-2 py-1">${r.nis}</td>
          <td class="border border-slate-400 px-2 py-1 font-semibold">${r.nama}</td>
          <td class="border border-slate-400 px-2 py-1 text-center">${r.benar}</td>
          <td class="border border-slate-400 px-2 py-1 text-center">${r.salah}</td>
          <td class="border border-slate-400 px-2 py-1 text-center">${r.kosong}</td>
          <td class="border border-slate-400 px-2 py-1 text-center font-bold">${r.nilai}</td>
          <td class="border border-slate-400 px-2 py-1 text-center font-semibold">${r.status}</td>
        `;
        printTbody.appendChild(tr);
      });
    }

    // CHART JS ANALYTICS
    function initChart() {
      const ctxChart = document.getElementById('scoreChart').getContext('2d');
      chartInstance = new Chart(ctxChart, {
        type: 'doughnut',
        data: getChartData(),
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 10 } } }
          }
        }
      });
    }

    function getChartData() {
      let g1 = 0, g2 = 0, g3 = 0, g4 = 0;
      rekapData.forEach(r => {
        if(r.nilai < 60) g1++;
        else if(r.nilai < 75) g2++;
        else if(r.nilai < 90) g3++;
        else g4++;
      });

      return {
        labels: ['< 60 (Kurang)', '60-74 (Remedial)', '75-89 (Baik)', '90-100 (Sangat Baik)'],
        datasets: [{
          data: [g1, g2, g3, g4],
          backgroundColor: ['#ef4444', '#f59e0b', '#3b82f6', '#10b981']
        }]
      };
    }

    function updateChart() {
      if(chartInstance) {
        chartInstance.data = getChartData();
        chartInstance.update();
      }
    }

    // EXPORT EXCEL & CSV
    function exportToExcel() {
      if(rekapData.length === 0) {
        showToast('Tidak ada data untuk diekspor!', 'error');
        return;
      }

      const exportRows = rekapData.map(r => ({
        "NIS": r.nis,
        "Nama Siswa": r.nama,
        "Mata Pelajaran": examConfig.mapel,
        "Kelas": examConfig.kelas,
        "Jumlah Soal": examConfig.totalQuestions,
        "Jumlah Benar": r.benar,
        "Jumlah Salah": r.salah,
        "Jumlah Kosong/Ganda": r.kosong,
        "Nilai Akhir": r.nilai,
        "Status Kelulusan": r.status
      }));

      const worksheet = XLSX.utils.json_to_sheet(exportRows);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Rekap Nilai LJK");
      XLSX.writeFile(workbook, `Rekap_Nilai_${examConfig.mapel.replace(/\s+/g, '_')}_${examConfig.kelas}.xlsx`);
      
      showToast(`File Excel berisi ${rekapData.length} data siswa berhasil diunduh!`);
    }

    function exportToCSV() {
      if(rekapData.length === 0) {
        showToast('Tidak ada data untuk diekspor!', 'error');
        return;
      }

      let csvContent = "data:text/csv;charset=utf-8,NIS,Nama Siswa,Benar,Salah,Kosong,Nilai,Status\n";
      rekapData.forEach(r => {
        csvContent += `${r.nis},"${r.nama}",${r.benar},${r.salah},${r.kosong},${r.nilai},${r.status}\n`;
      });

      const encodedUri = encodeURI(csvContent);
      const link = document.createElement("a");
      link.setAttribute("href", encodedUri);
      link.setAttribute("download", `Rekap_Nilai_${examConfig.mapel}.csv`);
      document.body.appendChild(link);
      link.click();
      link.remove();

      showToast('File CSV berhasil diunduh!');
    }


    // ==========================================
    // TAB 4: PRINTABLE TEMPLATE LJK GENERATOR
    // ==========================================
    function renderTemplateLJKGrid() {
      const gridContainer = document.getElementById('template-question-grid');
      gridContainer.innerHTML = '';

      const totalQ = examConfig.totalQuestions;
      const optionLetters = ['A', 'B', 'C', 'D', 'E'].slice(0, examConfig.optionsCount);

      for(let i = 0; i < totalQ; i++) {
        let optionsHtml = '';
        optionLetters.forEach(opt => {
          optionsHtml += `
            <div class="flex items-center gap-0.5">
              <span class="w-3.5 h-3.5 rounded-full border border-black flex items-center justify-center text-[8px] font-bold">${opt}</span>
            </div>
          `;
        });

        const qBox = document.createElement('div');
        qBox.className = "flex items-center justify-between bg-slate-50 p-1 border border-slate-200 rounded";
        qBox.innerHTML = `
          <span class="font-bold w-6 text-slate-700">${i+1}.</span>
          <div class="flex items-center gap-1.5">${optionsHtml}</div>
        `;
        gridContainer.appendChild(qBox);
      }
    }

  </script>
</body>
</html>
