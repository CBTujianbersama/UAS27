<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="theme-color" content="#0f172a">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <title>CBT-PRO | Sistem Ujian Berbasis Komputer</title>
    
    <!-- Tailwind CSS CDN (Render Dinamis & Konsisten di GitHub Pages) -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'sans-serif'],
                    }
                }
            }
        };
    </script>

    <!-- Google Fonts: Inter -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    
    <!-- Lucide Icons CDN -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- SheetJS (XLSX) CDN untuk Import/Export Excel -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
    <!-- Mammoth.js CDN untuk membaca file Word (.docx) beserta Gambar -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"></script>
    <!-- PDF.js CDN untuk membaca file PDF (.pdf) -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>

    <script>
        if (typeof pdfjsLib !== 'undefined') {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
    </script>

    <style>
        /* Reset & Stabilizer Layout agar tidak saling timpa di HP maupun Komputer (GitHub Pages Safe) */
        *, *::before, *::after {
            box-sizing: border-box;
        }

        html, body {
            margin: 0;
            padding: 0;
            width: 100%;
            min-height: 100vh;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background-color: #f8fafc;
            color: #0f172a;
            overflow-x: hidden;
            -webkit-font-smoothing: antialiased;
        }

        img, svg, video, canvas {
            max-width: 100%;
            display: block;
        }

        table {
            border-collapse: collapse;
            width: 100%;
        }

        .hidden {
            display: none !important;
        }

        /* Mencegah seleksi teks saat siswa mengerjakan ujian */
        .no-select {
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
        }

        /* Animasi Background Bergerak Portal */
        @keyframes gradientAurora {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }

        @keyframes floatOrbSlow {
            0%, 100% { transform: translate(0px, 0px) scale(1); }
            33% { transform: translate(28px, -22px) scale(1.06); }
            66% { transform: translate(-20px, 18px) scale(0.96); }
        }

        @keyframes floatOrbReverse {
            0%, 100% { transform: translate(0px, 0px) scale(1); }
            50% { transform: translate(-32px, -25px) scale(1.08); }
        }

        @keyframes subtleGridMove {
            0% { background-position: 0 0; }
            100% { background-position: 48px 48px; }
        }

        .portal-animated-hero {
            background: linear-gradient(-45deg, #0f172a, #172554, #1e1b4b, #0c4a6e, #0f172a);
            background-size: 300% 300%;
            animation: gradientAurora 14s ease infinite;
        }

        .portal-ambient-bg {
            background-color: #f8fafc;
            background-image: 
                radial-gradient(at 15% 15%, rgba(59, 130, 246, 0.08) 0px, transparent 50%),
                radial-gradient(at 85% 25%, rgba(99, 102, 241, 0.08) 0px, transparent 50%),
                radial-gradient(at 50% 85%, rgba(16, 185, 129, 0.06) 0px, transparent 50%);
        }

        .portal-grid-overlay {
            background-image: radial-gradient(rgba(255, 255, 255, 0.12) 1px, transparent 1px);
            background-size: 24px 24px;
            animation: subtleGridMove 12s linear infinite;
        }

        .orb-float-1 {
            animation: floatOrbSlow 11s ease-in-out infinite;
        }

        .orb-float-2 {
            animation: floatOrbReverse 14s ease-in-out infinite;
        }

        /* Animasi latar belakang layar terkunci anti-curang */
        @keyframes lockPulse {
            0%, 100% { background-color: rgba(127, 29, 29, 0.98); }
            50% { background-color: rgba(185, 28, 28, 0.98); }
        }
        .lock-overlay-bg {
            animation: lockPulse 2s infinite;
        }

        /* Stabilizer Sticky Header */
        #mainHeader {
            position: sticky;
            top: 0;
            z-index: 40;
            background-color: #ffffff;
        }

        /* Aturan Cetak (Print A4 2 Kolom) Khusus Kartu Login Peserta */
        @media print {
            body * {
                visibility: hidden;
            }
            #printCardsContainer, #printCardsContainer * {
                visibility: visible;
            }
            #printCardsContainer {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                display: grid !important;
                grid-template-columns: repeat(2, 1fr) !important;
                gap: 12px !important;
                padding: 8px !important;
                background: #ffffff !important;
            }
            .no-print {
                display: none !important;
            }
            .login-card-item {
                break-inside: avoid;
                page-break-inside: avoid;
                border: 1.5px solid #0f172a !important;
                box-shadow: none !important;
            }
        }

        /* Scrollbar halus */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #94a3b8;
        }

        /* Anti-Overlap & Hosting Layout Stabilizer (HP & Desktop Safe) */
        main > section {
            position: relative;
            width: 100%;
            clear: both;
        }

        .grid > *, [class*="grid-cols-"] > * {
            min-width: 0;
        }

        /* Pastikan semua modal benar-benar tersembunyi secara default sebelum dibuka */
        #confirmModal.hidden,
        #antiCheatLockOverlay.hidden,
        #loginModal.hidden,
        #studentTokenEntryModal.hidden,
        #editStudentModal.hidden,
        #editQuestionModal.hidden,
        #installPlatformModal.hidden,
        #examFinishedModal.hidden,
        #newsReaderModal.hidden,
        #imageLightboxModal.hidden,
        #gasCodeModal.hidden {
            display: none !important;
            visibility: hidden !important;
            pointer-events: none !important;
        }

        /* Saat modal aktif */
        #confirmModal.flex,
        #antiCheatLockOverlay.flex,
        #loginModal.flex,
        #studentTokenEntryModal.flex,
        #editStudentModal.flex,
        #editQuestionModal.flex,
        #installPlatformModal.flex,
        #examFinishedModal.flex,
        #newsReaderModal.flex,
        #imageLightboxModal.flex,
        #gasCodeModal.flex {
            display: flex !important;
            visibility: visible !important;
            pointer-events: auto !important;
        }

        /* Stabilisasi Sidebar & Tabel Responsif agar tidak bertabrakan */
        #view-admin aside {
            position: relative;
            z-index: 20;
        }
        @media (min-width: 1024px) {
            #view-admin aside {
                position: sticky;
                top: 5rem;
                max-height: calc(100vh - 6rem);
                overflow-y: auto;
            }
        }

        .break-anywhere {
            overflow-wrap: anywhere;
            word-break: break-word;
        }
    </style>
</head>
<body class="min-h-screen flex flex-col portal-ambient-bg">

    <!-- ===================================================================== -->
    <!-- 1. SISTEM TOAST & DIALOG KONFIRMASI                                   -->
    <!-- ===================================================================== -->
    <div id="toastContainer" class="fixed top-4 right-4 z-[9999] flex flex-col gap-2 max-w-sm w-full pointer-events-none px-4 sm:px-0"></div>

    <div id="confirmModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[9990] items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center gap-3 mb-3">
                <div class="w-10 h-10 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                    <i data-lucide="alert-triangle" class="w-5 h-5"></i>
                </div>
                <h3 id="confirmTitle" class="text-lg font-bold text-slate-900">Konfirmasi</h3>
            </div>
            <p id="confirmMessage" class="text-sm text-slate-600 mb-6 leading-relaxed">Apakah Anda yakin?</p>
            <div class="flex justify-end gap-2.5">
                <button id="confirmCancelBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 bg-slate-100 hover:bg-slate-200 transition">Batal</button>
                <button id="confirmOkBtn" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-white bg-blue-600 hover:bg-blue-700 transition shadow-sm">Ya, Lanjutkan</button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 2. OVERLAY LAYAR TERKUNCI ANTI-CURANG + SIRINE ALARM OTOMATIS         -->
    <!-- ===================================================================== -->
    <div id="antiCheatLockOverlay" class="fixed inset-0 lock-overlay-bg z-[99999] flex-col items-center justify-center p-4 sm:p-6 text-white text-center select-none hidden">
        <div class="max-w-lg w-full bg-red-950/95 border-2 border-red-400 rounded-3xl p-6 sm:p-8 shadow-2xl">
            <div class="w-20 h-20 rounded-full bg-red-600/30 border-2 border-red-400 flex items-center justify-center mx-auto mb-4 animate-bounce">
                <i data-lucide="siren" class="w-10 h-10 text-amber-300"></i>
            </div>
            <div class="flex items-center justify-center gap-2 mb-3">
                <span class="inline-block px-3 py-1 rounded-full bg-red-600 text-white text-xs font-extrabold uppercase tracking-wider">
                    SIRINE ALARM AKTIF • PELANGGARAN TERDETEKSI
                </span>
            </div>
            <h1 class="text-2xl sm:text-3xl font-extrabold tracking-tight mb-2">LAYAR UJIAN TERKUNCI!</h1>
            <p id="lockReasonText" class="text-red-200 text-xs sm:text-sm mb-5 leading-relaxed">Layar ujian kehilangan fokus! Dilarang membuka jendela lain, split-screen, Developer Tools, atau berpindah aplikasi.</p>

            <div class="bg-red-900/70 border border-red-700 rounded-2xl p-4 mb-5 text-left text-xs space-y-2">
                <div class="flex justify-between gap-2">
                    <span class="text-red-300">Nama Peserta:</span>
                    <span id="lockStudentName" class="font-bold text-white text-right">-</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300">NISN / Kelas:</span>
                    <span id="lockStudentMeta" class="font-semibold text-white text-right">-</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300">Mata Pelajaran Ujian:</span>
                    <span id="lockExamSubject" class="font-semibold text-amber-200 text-right">-</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300">Total Pelanggaran Tercatat:</span>
                    <span id="lockViolationBadge" class="font-extrabold text-amber-300">1 Kali</span>
                </div>
                <div class="flex justify-between gap-2">
                    <span class="text-red-300">Status Alarm &amp; Log:</span>
                    <span class="font-semibold text-emerald-300 text-right">Sirine Menyala • Terkirim ke Pengawas &amp; Admin</span>
                </div>
            </div>

            <div class="space-y-3">
                <label class="block text-xs font-bold uppercase tracking-wider text-red-200">
                    Masukkan Kode PIN Pengawas untuk Mematikan Sirine &amp; Membuka Kunci
                </label>
                <input type="password" id="unlockPinField" placeholder="Masukkan Kode PIN Pengawas" class="w-full px-4 py-3.5 rounded-xl bg-white text-slate-900 font-extrabold text-center text-lg tracking-widest focus:outline-none focus:ring-4 focus:ring-amber-400" onkeypress="if(event.key === 'Enter') verifyPinUnlock()">
                <button onclick="verifyPinUnlock()" class="w-full py-3.5 px-6 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold text-sm transition shadow-lg flex items-center justify-center gap-2">
                    <i data-lucide="unlock" class="w-4 h-4"></i> MATIKAN SIRINE &amp; BUKA KUNCI
                </button>
            </div>
            <p class="mt-4 text-[11px] text-red-300/90">
                Sirine peringatan akan terus berbunyi sampai Pengawas Ruangan memverifikasi PIN atau membuka kunci jarak jauh.
            </p>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 3. HEADER / NAVBAR UTAMA                                              -->
    <!-- ===================================================================== -->
    <header id="mainHeader" class="bg-white/95 backdrop-blur-md border-b border-slate-200 sticky top-0 z-40 shadow-sm no-print">
        <div class="max-w-7xl mx-auto px-3 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between min-h-[4rem] py-2 gap-2 sm:gap-3">
                <div class="flex items-center gap-2.5 sm:gap-3 cursor-pointer min-w-0 flex-1" onclick="handleHeaderBrandClick()">
                    <div class="w-9 h-9 sm:w-10 sm:h-10 rounded-xl bg-gradient-to-br from-blue-600 to-indigo-700 text-white flex items-center justify-center shadow-md shrink-0">
                        <i data-lucide="shield-check" class="w-5 h-5 sm:w-6 sm:h-6"></i>
                    </div>
                    <div class="min-w-0 flex-1">
                        <div class="flex items-center gap-2">
                            <h1 id="headerSchoolTitle" class="font-extrabold text-xs sm:text-base text-slate-900 truncate">CBT-PRO | Sistem Ujian Berbasis Komputer</h1>
                            <span id="cloudSyncBadge" onclick="event.stopPropagation(); syncFromCloud(true); showToast('Menyinkronkan data dengan Google Apps Script...', 'info');" title="Klik untuk Sinkronisasi Manual" class="hidden md:inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold shrink-0 cursor-pointer hover:bg-emerald-200 transition">
                                <span class="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping"></span> LIVE SYNC
                            </span>
                        </div>
                        <p id="headerSchoolSubtitle" class="text-[10px] sm:text-[11px] text-slate-500 truncate">Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027</p>
                    </div>
                </div>

                <!-- Tombol Aksi Header -->
                <div class="flex items-center gap-1.5 sm:gap-2 shrink-0">
                    <button onclick="toggleAppFullscreen()" title="Mode Layar Penuh Aplikasi" class="p-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition hidden sm:flex items-center justify-center">
                        <i data-lucide="maximize" class="w-4 h-4"></i>
                    </button>

                    <div id="headerGuestButtons" class="flex items-center gap-1.5 sm:gap-2">
                        <button onclick="openLoginModal('siswa')" class="px-2.5 sm:px-3 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-[11px] sm:text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                            <i data-lucide="graduation-cap" class="w-3.5 h-3.5 sm:w-4 sm:h-4"></i>
                            <span>Siswa</span>
                        </button>
                        <button onclick="openLoginModal('pengawas')" class="px-2.5 sm:px-3 py-2 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-[11px] sm:text-xs font-bold transition flex items-center gap-1.5">
                            <i data-lucide="user-check" class="w-3.5 h-3.5 sm:w-4 sm:h-4"></i>
                            <span class="hidden sm:inline">Pengawas</span>
                        </button>
                        <button onclick="openLoginModal('admin')" class="px-2.5 sm:px-3 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-[11px] sm:text-xs font-bold transition flex items-center gap-1.5">
                            <i data-lucide="settings" class="w-3.5 h-3.5 sm:w-4 sm:h-4"></i>
                            <span class="hidden sm:inline">Admin</span>
                        </button>
                    </div>

                    <div id="headerAuthControls" class="items-center gap-1.5 sm:gap-2 hidden">
                        <span id="headerUserBadge" class="hidden md:inline-block px-3 py-1.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold max-w-[160px] truncate">Administrator</span>
                        <button id="headerBackDashBtn" onclick="returnToRoleDashboard()" class="px-2.5 sm:px-3 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-[11px] sm:text-xs font-bold transition">
                            Dashboard
                        </button>
                        <button onclick="navigateTo('portal')" class="px-2.5 sm:px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] sm:text-xs font-bold transition">
                            Portal
                        </button>
                        <button onclick="logoutSession()" class="px-2.5 sm:px-3 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-[11px] sm:text-xs font-bold transition flex items-center gap-1">
                            <i data-lucide="log-out" class="w-3.5 h-3.5"></i> <span class="hidden sm:inline">Keluar</span>
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </header>

    <!-- ===================================================================== -->
    <!-- 4. KONTEN UTAMA (PORTAL, DASHBOARD SISWA, PENGAWAS, ADMIN, UJIAN)     -->
    <!-- ===================================================================== -->
    <main class="flex-grow">

        <!-- ================================================================= -->
        <!-- VIEW A: HALAMAN PORTAL UTAMA (ANIMATED BG + BERITA PENDIDIKAN)    -->
        <!-- ================================================================= -->
        <section id="view-portal" class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8 space-y-10">
            <!-- HERO BANNER DENGAN BACKGROUND BERGERAK -->
            <div id="portalHeroBanner" class="portal-animated-hero rounded-3xl p-6 sm:p-10 text-white shadow-xl relative overflow-hidden">
                <!-- Elemen Animasi Latar Bergerak (Non-Interactive) -->
                <div id="portalAnimatedOrbs" class="inset-0 absolute pointer-events-none overflow-hidden">
                    <div class="portal-grid-overlay absolute inset-0 opacity-60"></div>
                    <div class="orb-float-1 absolute -top-24 -right-24 w-80 h-80 bg-blue-500/25 rounded-full blur-3xl"></div>
                    <div class="orb-float-2 absolute -bottom-24 -left-20 w-72 h-72 bg-indigo-500/25 rounded-full blur-3xl"></div>
                    <div class="orb-float-1 absolute top-1/2 left-1/2 w-60 h-60 bg-cyan-400/15 rounded-full blur-3xl"></div>
                </div>

                <div class="relative z-10 grid grid-cols-1 lg:grid-cols-12 gap-8 items-center">
                    <div class="lg:col-span-7 space-y-5 min-w-0">
                        <div class="flex flex-wrap items-center gap-2">
                            <div class="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-blue-500/20 border border-blue-400/30 text-blue-200 text-xs font-semibold">
                                <i data-lucide="shield-alert" class="w-4 h-4 text-amber-300 shrink-0"></i>
                                <span>Sistem Ujian Multi-Tipe Soal • Single-App Lock • Sirine Anti-Curang</span>
                            </div>
                            <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-white/10 border border-white/15 text-emerald-300 text-[11px] font-bold">
                                <i data-lucide="code-2" class="w-3.5 h-3.5"></i> Development : Jovano F Kakalang, S.Ak.
                            </span>
                        </div>

                        <h2 class="text-2xl sm:text-4xl font-extrabold tracking-tight leading-tight">
                            Selamat Datang di Portal <span class="text-blue-400" id="heroExamTitle">Ujian Akhir Semester (UAS) Berbasis Komputer</span>
                        </h2>
                        <p class="text-slate-300 text-sm sm:text-base leading-relaxed">
                            Silakan login ke <strong>Dashboard Siswa</strong> menggunakan <strong>NISN / Username</strong> dan <strong>Password</strong> Kartu Login untuk melihat profil, jadwal penugasan mata pelajaran, serta memasukkan <strong>Token Ujian</strong> saat soal sudah siap dikerjakan.
                        </p>

                        <!-- Tombol Login 3 Peran -->
                        <div class="flex flex-wrap gap-2.5 sm:gap-3 pt-1">
                            <button onclick="openLoginModal('siswa')" class="px-5 py-3 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-bold text-xs sm:text-sm transition flex items-center gap-2 shadow-lg shadow-blue-950/60">
                                <i data-lucide="graduation-cap" class="w-4 h-4"></i> Login Dashboard Siswa
                            </button>
                            <button onclick="openLoginModal('pengawas')" class="px-4 py-3 rounded-xl bg-amber-500 hover:bg-amber-400 text-slate-950 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="eye" class="w-4 h-4"></i> Portal Pengawas
                            </button>
                            <button onclick="openLoginModal('admin')" class="px-4 py-3 rounded-xl bg-white/15 hover:bg-white/25 text-white border border-white/20 font-bold text-xs sm:text-sm transition flex items-center gap-2">
                                <i data-lucide="sliders" class="w-4 h-4"></i> Administrator
                            </button>
                        </div>

                        <!-- Tombol Salin Link Akses -->
                        <div class="pt-3 border-t border-white/10 flex flex-wrap items-center gap-2 text-xs">
                            <span class="text-slate-400 font-medium mr-1">Bagikan Link Akses:</span>
                            <button onclick="copyPortalLink('siswa')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-blue-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Portal Siswa
                            </button>
                            <button onclick="copyPortalLink('pengawas')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-amber-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Pengawas
                            </button>
                            <button onclick="copyPortalLink('admin')" class="px-3 py-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-indigo-200 font-semibold transition flex items-center gap-1.5">
                                <i data-lucide="link" class="w-3.5 h-3.5"></i> Link Administrator
                            </button>
                        </div>
                    </div>

                    <!-- Kartu Motivasi & Statistik Live -->
                    <div class="lg:col-span-5 min-w-0">
                        <div class="bg-white/10 backdrop-blur-md border border-white/20 rounded-3xl p-5 sm:p-6 space-y-4">
                            <div class="flex items-center justify-between gap-2">
                                <span class="inline-flex items-center gap-1.5 text-xs font-bold uppercase tracking-wider text-amber-300">
                                    <i data-lucide="sparkles" class="w-4 h-4 shrink-0"></i> Motivasi Belajar &amp; Kejujuran
                                </span>
                                <button onclick="nextMotivationQuote()" class="px-2.5 py-1 rounded-lg bg-white/10 hover:bg-white/20 text-xs font-semibold text-white transition flex items-center gap-1 shrink-0">
                                    <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i> Acak
                                </button>
                            </div>
                            <blockquote class="space-y-3">
                                <p id="motivationQuoteText" class="text-sm sm:text-base font-medium italic text-white leading-relaxed">
                                    "Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."
                                </p>
                                <footer id="motivationQuoteAuthor" class="text-xs font-bold text-blue-300">
                                    — Pesan Akademik Sekolah
                                </footer>
                            </blockquote>
                            <div class="pt-3 border-t border-white/10 grid grid-cols-3 gap-2 text-center">
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSoal" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Paket Ujian</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatSiswa" class="text-lg font-extrabold text-white">0</div>
                                    <div class="text-[10px] text-slate-400">Peserta Aktif</div>
                                </div>
                                <div class="bg-slate-900/50 rounded-xl p-2.5">
                                    <div id="portalStatDurasi" class="text-lg font-extrabold text-amber-300">0</div>
                                    <div class="text-[10px] text-slate-400">Terjadwal</div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <!-- BAGIAN BERITA TERKINI PENDIDIKAN & INFORMASI AKADEMIK -->
            <div class="space-y-5">
                <div class="flex flex-col md:flex-row md:items-end justify-between gap-3">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-indigo-600 flex items-center gap-1.5">
                            <i data-lucide="newspaper" class="w-3.5 h-3.5"></i> Wawasan &amp; Informasi Terkini
                        </span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Berita Terkini Pendidikan &amp; Pengumuman Akademik</h2>
                    </div>
                    <div class="flex flex-wrap gap-1.5" id="newsCategoryFilters">
                        <button onclick="filterPortalNews('ALL')" data-cat="ALL" class="news-cat-btn px-3 py-1.5 rounded-xl bg-indigo-600 text-white text-xs font-bold transition">Semua</button>
                        <button onclick="filterPortalNews('Pendidikan Nasional')" data-cat="Pendidikan Nasional" class="news-cat-btn px-3 py-1.5 rounded-xl bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs font-semibold transition">Pendidikan Nasional</button>
                        <button onclick="filterPortalNews('Tips Ujian & Literasi')" data-cat="Tips Ujian & Literasi" class="news-cat-btn px-3 py-1.5 rounded-xl bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs font-semibold transition">Tips Ujian &amp; Literasi</button>
                        <button onclick="filterPortalNews('Pengumuman Sekolah')" data-cat="Pengumuman Sekolah" class="news-cat-btn px-3 py-1.5 rounded-xl bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs font-semibold transition">Pengumuman Sekolah</button>
                    </div>
                </div>

                <div id="portalNewsGrid" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                    <!-- Dirender via JS -->
                </div>
            </div>

            <!-- BAGIAN PANDUAN SISTEM UJIAN CBT (SISWA, PENGAWAS, ADMIN) -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Tata Tertib &amp; Alur Pelaksanaan</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Alur Sistem Ujian CBT Terpadu</h2>
                    </div>
                    <p class="text-xs text-slate-500">Mendukung Soal Pilihan Ganda, Menjodohkan, Isian Singkat, Uraian &amp; Gambar JPG/PNG</p>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <!-- Panduan Siswa -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                                <i data-lucide="user-check" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">1. Alur Peserta / Siswa</h3>
                                <span class="text-xs text-blue-600 font-semibold">Dashboard &amp; Pengerjaan Multi-Tipe</span>
                            </div>
                        </div>
                        <ol class="space-y-2.5 text-xs text-slate-600 list-decimal pl-4 leading-relaxed">
                            <li><strong>Login ke Dashboard Siswa:</strong> Masuk dengan NISN/Username &amp; Password Kartu Login.</li>
                            <li><strong>Cek Profil &amp; Jadwal Penugasan:</strong> Lihat mata pelajaran yang sudah di-assign oleh Admin atau masih menunggu jadwal.</li>
                            <li><strong>Masukkan Token saat Siap:</strong> Input <strong>Token Aktif</strong> dari Pengawas untuk memulai soal Pilihan Ganda, Menjodohkan, Isian, maupun Uraian.</li>
                            <li><strong>Single-App Lock &amp; Sirine:</strong> Dilarang membuka aplikasi/tab lain atau Developer Tools saat ujian berlangsung.</li>
                        </ol>
                    </div>

                    <!-- Panduan Pengawas -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-amber-100 text-amber-600 flex items-center justify-center shrink-0">
                                <i data-lucide="shield" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">2. Alur Pengawas Ruang</h3>
                                <span class="text-xs text-amber-600 font-semibold">3 Menu Kendali Pengawasan</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Pantau Keluar/Masuk &amp; Token:</strong> Lihat log kehadiran siswa secara live beserta Token Ujian yang sinkron otomatis dengan Admin.</li>
                            <li><strong>Monitor Pelanggaran &amp; Sirine:</strong> Deteksi siswa yang keluar aplikasi, matikan alarm &amp; buka kunci layar jarak jauh.</li>
                            <li><strong>Input &amp; Upload Soal:</strong> Input soal berbagai tipe (PG, Menjodohkan, Isian, Uraian, Foto JPG/PNG) atau upload Word (.docx) / PDF (.pdf).</li>
                        </ul>
                    </div>

                    <!-- Panduan Administrator -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                        <div class="flex items-center gap-3">
                            <div class="w-11 h-11 rounded-2xl bg-indigo-100 text-indigo-600 flex items-center justify-center shrink-0">
                                <i data-lucide="cpu" class="w-6 h-6"></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">3. Alur Administrator</h3>
                                <span class="text-xs text-indigo-600 font-semibold">Penjadwalan, Cloud Sync &amp; Kontrol</span>
                            </div>
                        </div>
                        <ul class="space-y-2.5 text-xs text-slate-600 list-disc pl-4 leading-relaxed">
                            <li><strong>Bank Soal Multi-Format:</strong> Kelola paket soal per mata pelajaran lewat upload Word/PDF atau input manual dengan lampiran foto JPG/PNG.</li>
                            <li><strong>Assign Jadwal Mata Pelajaran:</strong> Tentukan tanggal/jam jadwal &amp; kelas target di menu Assignment.</li>
                            <li><strong>Pengaturan Semua Dashboard:</strong> Atur tampilan &amp; hak akses Dashboard Siswa maupun Pengawas langsung dari Setelan Keseluruhan.</li>
                        </ul>
                    </div>
                </div>
            </div>

            <!-- BAGIAN GALERI FOTO KEGIATAN UJIAN -->
            <div class="space-y-5">
                <div class="flex flex-col sm:flex-row sm:items-end justify-between gap-2">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-emerald-600">Dokumentasi &amp; Suasana Ujian</span>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900 mt-1">Galeri Kegiatan Ujian Berbasis Komputer</h2>
                    </div>
                    <p class="text-xs text-slate-500">Dokumentasi kegiatan dapat dikelola melalui menu Setelan Administrator</p>
                </div>
                <div id="portalGalleryGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6"></div>
            </div>

            <!-- BAGIAN DOWNLOAD & INSTALASI APLIKASI -->
            <div class="bg-white rounded-3xl border border-slate-200 p-6 sm:p-8 shadow-sm space-y-6">
                <div class="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-5 border-b border-slate-100">
                    <div>
                        <span class="text-xs font-bold uppercase tracking-wider text-blue-600">Aplikasi Multi-Platform &amp; GitHub Pages Ready</span>
                        <h2 class="text-xl font-extrabold text-slate-900 mt-1">Download &amp; Pasang Aplikasi di HP Android, iPhone, maupun Komputer</h2>
                        <p class="text-xs sm:text-sm text-slate-500 mt-1">Pasang sebagai aplikasi mandiri layar penuh untuk kenyamanan dan keamanan ujian maksimal.</p>
                    </div>
                    <button onclick="downloadStandaloneApp()" class="px-4 py-3 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-2 self-start md:self-center shrink-0 shadow-sm">
                        <i data-lucide="download" class="w-4 h-4"></i> Unduh File Aplikasi (index.html)
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div class="rounded-2xl border border-emerald-200 bg-emerald-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">HP / Tablet Android</h3>
                                    <span class="text-[11px] font-semibold text-emerald-700">Android 8.0+ (Chrome / WebApp)</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang aplikasi CBT langsung ke layar utama Android tanpa bilah alamat browser sehingga berjalan seperti aplikasi tunggal terkunci.
                            </p>
                        </div>
                        <button onclick="openInstallModal('android')" class="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="download-cloud" class="w-4 h-4"></i> Install / Panduan Android
                        </button>
                    </div>

                    <div class="rounded-2xl border border-slate-200 bg-slate-50/70 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-slate-900 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="tablet-smartphone" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">iPhone &amp; iPad (iOS)</h3>
                                    <span class="text-[11px] font-semibold text-slate-600">iOS Safari WebClip / PWA</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Gunakan fitur <em>"Tambahkan ke Layar Utama"</em> di Safari iPhone/iPad untuk menjalankan CBT dalam mode layar penuh yang stabil.
                            </p>
                        </div>
                        <button onclick="openInstallModal('ios')" class="w-full py-2.5 px-4 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="smartphone" class="w-4 h-4"></i> Install / Panduan iPhone (iOS)
                        </button>
                    </div>

                    <div class="rounded-2xl border border-blue-200 bg-blue-50/40 p-5 flex flex-col justify-between space-y-4">
                        <div class="space-y-3">
                            <div class="flex items-center gap-3">
                                <div class="w-11 h-11 rounded-2xl bg-blue-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                                    <i data-lucide="monitor" class="w-6 h-6"></i>
                                </div>
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm">Komputer / Laptop (PC &amp; Mac)</h3>
                                    <span class="text-[11px] font-semibold text-blue-700">Windows / macOS / Chromebook</span>
                                </div>
                            </div>
                            <p class="text-xs text-slate-600 leading-relaxed">
                                Pasang sebagai aplikasi desktop mandiri lewat Chrome/Edge atau jalankan langsung melalui hosting GitHub Pages sekolah.
                            </p>
                        </div>
                        <button onclick="openInstallModal('desktop')" class="w-full py-2.5 px-4 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center justify-center gap-2">
                            <i data-lucide="laptop" class="w-4 h-4"></i> Install / Download Komputer
                        </button>
                    </div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW B: DASHBOARD SISWA (MENU TETAP + TERKENDALI SETELAN ADMIN)   -->
        <!-- ================================================================= -->
        <section id="view-siswa-dashboard" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
            <!-- Banner Pengumuman Khusus Siswa (Diatur Admin) -->
            <div id="stuCustomNoticeBanner" class="hidden bg-indigo-50 border border-indigo-200 rounded-2xl p-4 flex items-start gap-3 text-xs text-indigo-950">
                <i data-lucide="megaphone" class="w-5 h-5 text-indigo-600 shrink-0 mt-0.5"></i>
                <div>
                    <span class="font-extrabold uppercase tracking-wider text-[10px] text-indigo-600 block">Pengumuman Administrator untuk Peserta</span>
                    <p id="stuCustomNoticeText" class="font-medium mt-0.5 leading-relaxed"></p>
                </div>
            </div>

            <!-- Header Sapaan Siswa -->
            <div class="bg-gradient-to-r from-slate-900 via-blue-950 to-indigo-950 rounded-3xl p-5 sm:p-6 text-white shadow-lg flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div class="flex items-center gap-3 sm:gap-4 min-w-0">
                    <div class="w-12 h-12 sm:w-14 sm:h-14 rounded-2xl bg-blue-600/30 border border-blue-400/40 flex items-center justify-center font-extrabold text-xl text-blue-200 shrink-0">
                        <i data-lucide="user-check" class="w-6 h-6 sm:w-7 sm:h-7"></i>
                    </div>
                    <div class="min-w-0">
                        <span class="px-2.5 py-0.5 rounded-full bg-emerald-500/20 border border-emerald-400/30 text-emerald-300 text-[10px] font-bold uppercase tracking-wider">
                            Akun Peserta Terverifikasi
                        </span>
                        <h2 id="stuDashWelcomeName" class="text-lg sm:text-2xl font-extrabold mt-1 truncate">Selamat Datang, Peserta</h2>
                        <p id="stuDashWelcomeSub" class="text-xs text-blue-200 mt-0.5 truncate">Silakan pilih mata pelajaran yang sudah di-assign untuk memulai ujian.</p>
                    </div>
                </div>
                <div class="flex items-center gap-2 self-start md:self-center shrink-0">
                    <span class="px-3 py-1.5 rounded-xl bg-white/10 border border-white/15 text-xs font-semibold flex items-center gap-1.5">
                        <span class="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span> Sinkron Jadwal Live
                    </span>
                </div>
            </div>

            <!-- Navigasi Tab Dashboard Siswa (Menu Asli Dipertahankan) -->
            <div class="flex flex-wrap gap-2 bg-white p-2 rounded-2xl border border-slate-200 shadow-sm">
                <button onclick="switchStudentDashTab('ujian')" id="stuTabBtn-ujian" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-blue-600 text-white transition flex items-center gap-2">
                    <i data-lucide="book-open-check" class="w-4 h-4"></i> Penugasan &amp; Jadwal Ujian
                </button>
                <button onclick="switchStudentDashTab('profil')" id="stuTabBtn-profil" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="id-card" class="w-4 h-4"></i> Profil &amp; Kartu Login Saya
                </button>
                <button onclick="switchStudentDashTab('panduan')" id="stuTabBtn-panduan" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="shield-alert" class="w-4 h-4"></i> Panduan Menjawab &amp; Tata Tertib
                </button>
            </div>

            <!-- TAB 1 SISWA: PENUGASAN / UJIAN MATA PELAJARAN TERJADWAL -->
            <div id="stuDashTab-ujian" class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-8 space-y-4 min-w-0">
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-extrabold text-slate-900 text-base">Daftar Ujian &amp; Penugasan Mata Pelajaran</h3>
                            <p class="text-xs text-slate-500">Pilih ujian yang berstatus <strong>Aktif / Siap Dikerjakan</strong>, lalu masukkan Token Ujian saat Anda siap.</p>
                        </div>
                        <span id="stuAssignedCountBadge" class="px-3 py-1 rounded-full bg-blue-50 text-blue-700 text-xs font-bold self-start shrink-0">0 Paket</span>
                    </div>

                    <div id="studentAssignmentsList" class="space-y-4"></div>
                </div>

                <!-- Ringkasan Profil Kartu Login & Pengingat Tata Tertib -->
                <div class="lg:col-span-4 space-y-4 min-w-0">
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between border-b border-slate-100 pb-2.5">
                            <span class="text-xs font-extrabold uppercase tracking-wider text-blue-600">Ringkasan Kartu Peserta</span>
                            <button id="stuSideProfileLink" onclick="switchStudentDashTab('profil')" class="text-[11px] font-bold text-slate-500 hover:text-blue-600">Lihat Kartu →</button>
                        </div>
                        <div class="space-y-2 text-xs">
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500">Nama Lengkap:</span>
                                <span id="stuSideName" class="font-bold text-slate-900 text-right">-</span>
                            </div>
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500">NISN / Username:</span>
                                <span id="stuSideNisn" class="font-mono font-bold text-slate-800 text-right">-</span>
                            </div>
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500">Kelas:</span>
                                <span id="stuSideClass" class="font-bold text-indigo-700 text-right">-</span>
                            </div>
                            <div class="flex justify-between gap-2">
                                <span class="text-slate-500">Ruangan Ujian:</span>
                                <span id="stuSideRoom" class="font-bold text-slate-900 text-right">-</span>
                            </div>
                        </div>
                    </div>

                    <div class="bg-amber-50 border border-amber-200 rounded-2xl p-5 space-y-2.5 text-xs text-amber-950">
                        <div class="font-extrabold text-amber-900 flex items-center gap-1.5 text-sm">
                            <i data-lucide="siren" class="w-4 h-4 text-red-600 shrink-0"></i> Peringatan Keamanan Ruang Ujian
                        </div>
                        <p class="leading-relaxed text-amber-900">
                            Setelah Anda menekan tombol <strong>Mulai Ujian</strong> dan memasukkan Token:
                        </p>
                        <ul class="list-disc pl-4 space-y-1 text-amber-900">
                            <li>Aplikasi akan mengunci mode layar penuh tunggal (Single-App Lock).</li>
                            <li>Dilarang berpindah tab, membuka WhatsApp/Google, Inspect Element, atau membagi layar.</li>
                            <li>Jika keluar sistem, <strong>layar akan terkunci otomatis &amp; sirine alarm berbunyi keras</strong>.</li>
                        </ul>
                        <button id="stuSideGuideLink" onclick="switchStudentDashTab('panduan')" class="mt-2 inline-flex items-center gap-1 text-xs font-bold text-red-700 hover:underline">
                            Baca Selengkapnya Tata Tertib &amp; Panduan →
                        </button>
                    </div>
                </div>
            </div>

            <!-- TAB 2 SISWA: PROFIL LENGKAP SESUAI KARTU LOGIN -->
            <div id="stuDashTab-profil" class="hidden grid grid-cols-1 lg:grid-cols-12 gap-6">
                <div class="lg:col-span-6 bg-white rounded-3xl min-w-0 border-2 border-slate-300 p-6 shadow-sm space-y-4">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-200">
                        <div class="flex items-center gap-3 min-w-0">
                            <div class="w-11 h-11 rounded-xl bg-slate-900 text-white flex items-center justify-center font-extrabold text-sm shrink-0">CBT</div>
                            <div class="min-w-0">
                                <div class="text-xs font-extrabold uppercase tracking-wider text-slate-900">KARTU IDENTITAS PESERTA UJIAN</div>
                                <div id="stuCardSchoolName" class="text-[11px] text-slate-500 truncate">CBT-PRO | Sistem Ujian Berbasis Komputer</div>
                            </div>
                        </div>
                        <span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[10px] font-bold shrink-0">AKTIF</span>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs pt-1">
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Nama Lengkap Peserta</span>
                            <span id="stuProfName" class="text-sm font-extrabold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">NISN</span>
                            <span id="stuProfNisn" class="text-sm font-mono font-extrabold text-blue-700 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Kelas / Jurusan</span>
                            <span id="stuProfClass" class="text-sm font-bold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200">
                            <span class="text-[10px] uppercase font-bold text-slate-400 block">Ruangan Ujian</span>
                            <span id="stuProfRoom" class="text-sm font-bold text-indigo-700 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                            <span class="text-[10px] uppercase font-bold text-blue-500 block">Username Kartu Login</span>
                            <span id="stuProfUser" class="text-sm font-mono font-extrabold text-slate-900 mt-0.5 block">-</span>
                        </div>
                        <div class="p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                            <span class="text-[10px] uppercase font-bold text-blue-500 block">Status Sesi Login</span>
                            <span id="stuProfLoginTime" class="text-xs font-bold text-emerald-700 mt-0.5 block">Terhubung Real-Time</span>
                        </div>
                    </div>
                </div>

                <div class="lg:col-span-6 bg-white rounded-3xl min-w-0 border border-slate-200 p-6 shadow-sm space-y-4">
                    <h3 class="font-extrabold text-slate-900 text-base flex items-center gap-2">
                        <i data-lucide="history" class="w-4 h-4 text-blue-600"></i> Riwayat Pengerjaan Ujian Saya
                    </h3>
                    <div id="stuMyHistoryList" class="space-y-2.5 text-xs"></div>
                </div>
            </div>

            <!-- TAB 3 SISWA: PANDUAN MENJAWAB SOAL & TATA TERTIB RUANG UJIAN -->
            <div id="stuDashTab-panduan" class="hidden grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="bg-white rounded-3xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div class="flex items-center gap-3 pb-3 border-b border-slate-100">
                        <div class="w-10 h-10 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                            <i data-lucide="help-circle" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h3 class="font-extrabold text-slate-900 text-base">Panduan Cara Menjawab Soal CBT</h3>
                            <span class="text-xs text-slate-500">Langkah teknis pengerjaan berbagai tipe soal</span>
                        </div>
                    </div>
                    <ol class="space-y-3 text-xs sm:text-sm text-slate-700 list-decimal pl-5 leading-relaxed">
                        <li><strong>Pilih Mata Pelajaran Terjadwal:</strong> Pada tab <em>Penugasan &amp; Jadwal Ujian</em>, pastikan status ujian sudah <strong>Siap Dikerjakan (Assigned)</strong>, lalu klik tombol <strong>Mulai Ujian</strong>.</li>
                        <li><strong>Masukkan Token Ujian:</strong> Ketikkan kode <strong>Token Ujian</strong> yang diberikan oleh Pengawas Ruangan.</li>
                        <li><strong>Cara Menjawab 4 Tipe Soal:</strong>
                            <ul class="list-disc pl-4 mt-1 space-y-1 text-xs text-slate-600">
                                <li><strong>Pilihan Ganda:</strong> Klik opsi A, B, C, D, atau E.</li>
                                <li><strong>Menjodohkan:</strong> Pilih pasangan jawaban yang sesuai pada dropdown di sebelah kanan setiap pernyataan.</li>
                                <li><strong>Isian Singkat:</strong> Ketik jawaban pendek pada kotak yang tersedia.</li>
                                <li><strong>Uraian:</strong> Tuliskan penjelasan jawaban secara lengkap pada kolom uraian.</li>
                            </ul>
                        </li>
                        <li><strong>Zoom Gambar Soal:</strong> Jika soal memiliki foto/diagram, klik gambar untuk memperbesar tampilan.</li>
                        <li><strong>Mengakhiri Ujian:</strong> Setelah seluruh soal terjawab dan diperiksa, klik tombol <strong>Selesai Ujian</strong>.</li>
                    </ol>
                </div>

                <div class="bg-white rounded-3xl border border-red-200 p-6 shadow-sm space-y-4">
                    <div class="flex items-center gap-3 pb-3 border-b border-red-100">
                        <div class="w-10 h-10 rounded-xl bg-red-100 text-red-600 flex items-center justify-center shrink-0">
                            <i data-lucide="shield-alert" class="w-5 h-5"></i>
                        </div>
                        <div>
                            <h3 class="font-extrabold text-red-900 text-base">Tata Tertib Peserta di Ruangan Ujian</h3>
                            <span class="text-xs text-red-600 font-semibold">Wajib Dipatuhi Seluruh Peserta Ujian</span>
                        </div>
                    </div>
                    <ul class="space-y-3 text-xs sm:text-sm text-slate-700 list-disc pl-5 leading-relaxed">
                        <li>Peserta wajib hadir di ruangan ujian minimal 15 menit sebelum jadwal dimulai dengan membawa <strong>Kartu Login Peserta</strong>.</li>
                        <li>Peserta dilarang keras membuka tab browser lain, aplikasi pesan, kalkulator eksternal, Developer Tools, atau melakukan <em>split-screen</em> selama ujian berlangsung.</li>
                        <li>Peserta dilarang keluar dari mode layar penuh (<em>Fullscreen</em>), menekan tombol kembali (<em>Back</em>), atau mengambil tangkapan layar (<em>Screenshot</em>).</li>
                        <li><strong>Sanksi Sistem Otomatis:</strong> Setiap upaya keluar dari halaman ujian akan langsung <strong>mengunci layar ujian, membunyikan sirine alarm peringatan</strong>, serta mencatat pelanggaran secara real-time di layar Pengawas dan Administrator.</li>
                        <li>Pembukaan kunci layar hanya dapat dilakukan oleh <strong>Pengawas Ruangan</strong> setelah memeriksa perangkat peserta.</li>
                    </ul>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW C: DASHBOARD ADMINISTRATOR (7 MENU SIDEBAR TETAP UTUH)       -->
        <!-- ================================================================= -->
        <section id="view-admin" class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 hidden">
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                
                <!-- SIDEBAR NAVIGASI (7 MENU ASLI TIDAK DIUBAH) -->
                <aside class="lg:col-span-3 bg-white min-w-0 rounded-2xl border border-slate-200 p-4 shadow-sm space-y-2 lg:sticky lg:top-20 no-print">
                    <div class="px-3 py-2 border-b border-slate-100 mb-2">
                        <span class="text-[10px] font-extrabold uppercase tracking-wider text-indigo-600">Panel Kendali Utama</span>
                        <h2 class="font-extrabold text-slate-900 text-base">Menu Administrator</h2>
                    </div>

                    <nav class="flex lg:flex-col overflow-x-auto lg:overflow-visible gap-1.5 pb-1 lg:pb-0">
                        <button onclick="switchAdminMenu('monitoring')" id="adminNav-monitoring" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap">
                            <i data-lucide="activity" class="w-4 h-4 shrink-0"></i>
                            <span>1. Monitoring Siswa</span>
                        </button>
                        <button onclick="switchAdminMenu('soal')" id="adminNav-soal" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-text" class="w-4 h-4 shrink-0"></i>
                            <span>2. Bank Soal (Per Mapel)</span>
                        </button>
                        <button onclick="switchAdminMenu('assign')" id="adminNav-assign" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="calendar-check" class="w-4 h-4 shrink-0"></i>
                            <span>3. Assignment &amp; Jadwal</span>
                        </button>
                        <button onclick="switchAdminMenu('logs')" id="adminNav-logs" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="scroll-text" class="w-4 h-4 shrink-0"></i>
                            <span>4. Log Masuk/Keluar &amp; Pelanggaran</span>
                        </button>
                        <button onclick="switchAdminMenu('nilai')" id="adminNav-nilai" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="file-spreadsheet" class="w-4 h-4 shrink-0"></i>
                            <span>5. Data &amp; Download Nilai</span>
                        </button>
                        <button onclick="switchAdminMenu('kartu')" id="adminNav-kartu" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="id-card" class="w-4 h-4 shrink-0"></i>
                            <span>6. Akun Siswa &amp; Kartu</span>
                        </button>
                        <button onclick="switchAdminMenu('setelan')" id="adminNav-setelan" class="admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap">
                            <i data-lucide="settings-2" class="w-4 h-4 shrink-0"></i>
                            <span>7. Setelan Keseluruhan</span>
                        </button>
                    </nav>

                    <!-- Info Token Aktif di Sidebar -->
                    <div class="hidden lg:block mt-4 pt-4 border-t border-slate-100">
                        <div class="bg-slate-900 text-white rounded-xl p-3.5">
                            <div class="flex items-center justify-between text-[10px] text-slate-400 uppercase font-bold">
                                <span>Token Ujian Aktif</span>
                                <span class="text-emerald-400">SINKRON</span>
                            </div>
                            <div class="flex items-center justify-between mt-1">
                                <span id="sidebarTokenDisplay" class="font-mono font-extrabold text-xl text-amber-400 tracking-widest">CBT26X</span>
                                <button onclick="copyTextToClipboard(document.getElementById('sidebarTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white" title="Salin Token">
                                    <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                </aside>

                <!-- KONTEN KANAN ADMINISTRATOR -->
                <div class="lg:col-span-9 space-y-6 min-w-0">

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 1: MONITORING SISWA (REAL-TIME)                  -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-monitoring" class="space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Monitoring Peserta Ujian Real-Time</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Pantau status siswa di Dashboard maupun ruang ujian, matikan sirine/buka kunci layar, atau reset sesi.</p>
                            </div>
                            <span class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start">
                                <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Sinkronisasi Cloud
                            </span>
                        </div>

                        <div class="grid grid-cols-2 sm:grid-cols-4 gap-4">
                            <div class="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm">
                                <span class="text-xs text-slate-500 font-medium">Total Peserta</span>
                                <div id="monStatTotal" class="text-2xl font-extrabold text-slate-900 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-blue-200 p-4 shadow-sm">
                                <span class="text-xs text-blue-600 font-semibold">Sedang Ujian / Online</span>
                                <div id="monStatActive" class="text-2xl font-extrabold text-blue-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-emerald-200 p-4 shadow-sm">
                                <span class="text-xs text-emerald-600 font-semibold">Selesai Ujian</span>
                                <div id="monStatDone" class="text-2xl font-extrabold text-emerald-700 mt-1">0</div>
                            </div>
                            <div class="bg-white rounded-2xl border border-red-200 p-4 shadow-sm">
                                <span class="text-xs text-red-600 font-semibold">Terkunci / Sirine</span>
                                <div id="monStatLocked" class="text-2xl font-extrabold text-red-600 mt-1">0</div>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="monSearchInput" oninput="renderAdminMonitoring()" placeholder="Cari nama atau NISN siswa..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-64 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <select id="monRoomFilter" onchange="renderAdminMonitoring()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white"></select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Peserta &amp; NISN</th>
                                            <th class="py-3 px-4 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-4 font-bold">Mapel Aktif</th>
                                            <th class="py-3 px-4 font-bold">Status Live</th>
                                            <th class="py-3 px-4 font-bold">Progres</th>
                                            <th class="py-3 px-4 font-bold">Pelanggaran</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Cepat</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminMonitoringTbody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 2: BANK SOAL MULTI-TIPE & FOTO JPG/PNG           -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-soal" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Bank Soal &amp; Paket Mata Pelajaran</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Mendukung tipe soal <strong>Pilihan Ganda, Menjodohkan, Isian Singkat, Uraian</strong> serta <strong>Foto JPG/PNG</strong> via Upload Word/PDF atau Input Manual.</p>
                            </div>
                        </div>

                        <!-- Kelola / Pilih Paket Mata Pelajaran -->
                        <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-4">
                            <div class="flex flex-col md:flex-row md:items-end justify-between gap-4">
                                <div class="flex-1 min-w-0">
                                    <label class="block text-xs font-extrabold uppercase tracking-wider text-indigo-700 mb-1.5">Pilih Paket Soal Mata Pelajaran yang Dikelola:</label>
                                    <select id="adminPackageSelector" onchange="selectAdminQuestionPackage(this.value)" class="w-full px-4 py-2.5 rounded-xl border-2 border-indigo-300 bg-indigo-50/40 text-sm font-bold text-slate-900"></select>
                                </div>
                                <form onsubmit="createNewSubjectPackage(event)" class="flex flex-wrap sm:flex-nowrap gap-2">
                                    <input type="text" id="newPkgSubjectName" required placeholder="Nama Mapel Baru (Misal: Matematika)" class="px-3.5 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                    <button type="submit" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold whitespace-nowrap transition">
                                        + Buat Paket Mapel
                                    </button>
                                </form>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6">
                            <!-- Kolom Input Soal -->
                            <div class="xl:col-span-5 space-y-6 min-w-0">
                                <!-- Upload Word (.docx) & PDF (.pdf) -->
                                <div class="bg-white rounded-2xl border border-indigo-200 p-5 shadow-sm space-y-3">
                                    <div class="flex items-center justify-between gap-2">
                                        <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                            <i data-lucide="file-up" class="w-4 h-4 text-indigo-600 shrink-0"></i> Upload Word (.docx) / PDF (.pdf)
                                        </h3>
                                        <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-indigo-600 hover:underline shrink-0">
                                            Unduh Contoh Format
                                        </button>
                                    </div>
                                    <div class="text-[11px] text-slate-600 leading-relaxed bg-slate-50 p-2.5 rounded-xl border border-slate-200 space-y-1">
                                        <div class="font-bold text-slate-800">Format Otomatis Word / PDF (Termasuk Gambar di .DOCX):</div>
                                        <div>• <code>1. [PG] Pertanyaan... A. ... B. ... Kunci: B</code></div>
                                        <div>• <code>2. [MENJODOHKAN] Soal... 1) Kiri = Kanan</code></div>
                                        <div>• <code>3. [ISIAN] Pertanyaan... Kunci: Jawaban1, Jawaban2</code></div>
                                        <div>• <code>4. [URAIAN] Pertanyaan... Kunci: kata kunci1, kata kunci2</code></div>
                                    </div>
                                    <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-indigo-300 rounded-2xl bg-indigo-50/40 hover:bg-indigo-50 cursor-pointer transition p-4 text-center">
                                        <i data-lucide="upload-cloud" class="w-7 h-7 text-indigo-600 mb-1"></i>
                                        <span class="text-xs font-bold text-indigo-900">Klik Pilih File .DOCX, .PDF, atau .TXT</span>
                                        <span class="text-[10px] text-slate-500 mt-0.5">Ekstraksi otomatis PG, Menjodohkan, Isian, Uraian &amp; Gambar</span>
                                        <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event, 'admin')" class="hidden">
                                    </label>
                                    <div id="uploadParseStatus" class="hidden text-xs p-2.5 rounded-xl bg-blue-50 text-blue-700 font-medium"></div>
                                </div>

                                <!-- Form Input Soal Manual Multi-Tipe + Foto JPG/PNG -->
                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Soal Manual ke Paket Ini</h3>
                                    <form onsubmit="handleAddManualQuestion(event, 'admin')" class="space-y-3">
                                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Tipe Soal *</label>
                                                <select id="adminQType" onchange="toggleQuestionTypeFields('admin')" class="w-full px-3 py-2 rounded-xl border border-indigo-300 bg-indigo-50/40 text-xs font-bold text-indigo-900">
                                                    <option value="pg">Pilihan Ganda (A-E)</option>
                                                    <option value="menjodohkan">Menjodohkan (Pasangan)</option>
                                                    <option value="isian">Isian Singkat</option>
                                                    <option value="uraian">Uraian / Esai</option>
                                                </select>
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin *</label>
                                                <input type="number" id="adminQPoints" value="10" min="1" max="100" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                            </div>
                                        </div>

                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan / Instruksi *</label>
                                            <textarea id="adminQText" required rows="2" placeholder="Tuliskan pertanyaan soal..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none"></textarea>
                                        </div>

                                        <!-- Upload Foto Soal JPG / PNG -->
                                        <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                                            <div class="flex items-center justify-between">
                                                <label class="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                                                    <i data-lucide="image-plus" class="w-3.5 h-3.5 text-indigo-600"></i> Lampirkan Foto Soal (JPG / PNG)
                                                </label>
                                                <button type="button" id="adminQImgClearBtn" onclick="clearQuestionImageInput('admin')" class="hidden text-[11px] font-bold text-red-600 hover:underline">Hapus Foto</button>
                                            </div>
                                            <input type="file" id="adminQImageFile" accept="image/jpeg,image/png" onchange="handleQuestionImageSelect(event, 'admin')" class="block w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-indigo-100 file:text-indigo-700 hover:file:bg-indigo-200">
                                            <input type="hidden" id="adminQImageData">
                                            <div id="adminQImagePreviewBox" class="hidden pt-1">
                                                <img id="adminQImagePreview" src="" alt="Pratinjau Foto Soal" class="max-h-36 rounded-lg border border-slate-300 object-contain mx-auto">
                                            </div>
                                        </div>

                                        <!-- Sub-Form: PILIHAN GANDA -->
                                        <div id="adminFields-pg" class="space-y-2">
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">A.</span>
                                                <input type="text" id="adminQOptA" placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">B.</span>
                                                <input type="text" id="adminQOptB" placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">C.</span>
                                                <input type="text" id="adminQOptC" placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">D.</span>
                                                <input type="text" id="adminQOptD" placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="flex items-center gap-2">
                                                <span class="w-6 text-xs font-bold text-slate-500">E.</span>
                                                <input type="text" id="adminQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                            </div>
                                            <div class="pt-1">
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban Pilihan Ganda</label>
                                                <select id="adminQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                                                    <option value="A">Opsi A</option>
                                                    <option value="B">Opsi B</option>
                                                    <option value="C">Opsi C</option>
                                                    <option value="D">Opsi D</option>
                                                    <option value="E">Opsi E</option>
                                                </select>
                                            </div>
                                        </div>

                                        <!-- Sub-Form: MENJODOHKAN -->
                                        <div id="adminFields-menjodohkan" class="hidden space-y-2 p-3 rounded-xl bg-amber-50/60 border border-amber-200">
                                            <div class="flex items-center justify-between">
                                                <span class="text-xs font-bold text-amber-900">Pasangan Pernyataan (Kiri = Kanan)</span>
                                                <button type="button" onclick="addMatchingPairRow('adminMatchingRows')" class="px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-bold">+ Pasangan</button>
                                            </div>
                                            <div id="adminMatchingRows" class="space-y-2"></div>
                                        </div>

                                        <!-- Sub-Form: ISIAN SINGKAT -->
                                        <div id="adminFields-isian" class="hidden space-y-1.5 p-3 rounded-xl bg-emerald-50/60 border border-emerald-200">
                                            <label class="block text-xs font-bold text-emerald-900">Kunci Jawaban Isian Singkat *</label>
                                            <input type="text" id="adminQShortAnswer" placeholder="Contoh: 17 Agustus 1945, 17-08-1945" class="w-full px-3 py-2 rounded-lg border border-emerald-300 text-xs bg-white">
                                            <p class="text-[10px] text-emerald-700">Pisahkan dengan tanda koma (,) jika ada lebih dari satu variasi jawaban benar. Tidak peka huruf besar/kecil.</p>
                                        </div>

                                        <!-- Sub-Form: URAIAN -->
                                        <div id="adminFields-uraian" class="hidden space-y-1.5 p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                                            <label class="block text-xs font-bold text-blue-900">Pedoman Jawaban / Kata Kunci Penilaian Uraian</label>
                                            <textarea id="adminQEssayRubric" rows="2" placeholder="Contoh: kejujuran, integritas, mandiri, disiplin" class="w-full px-3 py-2 rounded-lg border border-blue-300 text-xs bg-white"></textarea>
                                            <p class="text-[10px] text-blue-700">Pisahkan kata kunci utama dengan koma untuk penilaian otomatis, serta dapat ditinjau pada Rekap Nilai.</p>
                                        </div>

                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Soal ke Paket
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <!-- Daftar & Edit Soal -->
                            <div class="xl:col-span-7 bg-white min-w-0 rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col">
                                <div class="flex flex-wrap items-center justify-between gap-2 mb-4 pb-3 border-b border-slate-100">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Soal Paket Terpilih (<span id="adminBankSoalCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Soal yang sudah siap di sini dapat langsung di-assign di Menu 3</p>
                                    </div>
                                    <button onclick="confirmClearPackageQuestions()" class="px-3 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                        Kosongkan Soal Paket
                                    </button>
                                </div>
                                <div id="adminSoalListContainer" class="space-y-3 overflow-y-auto max-h-[760px] pr-1"></div>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 3: ASSIGNMENT MATA PELAJARAN, JADWAL & TOKEN     -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-assign" class="space-y-6 hidden">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm">
                            <h2 class="text-xl font-extrabold text-slate-900">Assignment Mata Pelajaran, Jadwal Ujian &amp; Token</h2>
                            <p class="text-xs text-slate-500 mt-0.5">Pilih soal mata pelajaran yang telah disiapkan di Bank Soal, atur jadwal pelaksanaan &amp; status Assignment (Aktif / Menunggu Assign), serta kelola Token sinkron.</p>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
                            <!-- Manajemen Token Ujian -->
                            <div class="lg:col-span-5 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-5 self-start">
                                <div>
                                    <span class="px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">Sinkron Otomatis ke Pengawas &amp; Siswa</span>
                                    <h3 class="font-bold text-slate-900 text-base mt-2">Token Akses Ujian Aktif</h3>
                                    <p class="text-xs text-slate-500 mt-0.5">Siswa memasukkan token ini di Dashboard Siswa saat menekan tombol Mulai Ujian pada mapel yang sudah di-assign.</p>
                                </div>

                                <div class="bg-slate-900 text-white p-5 rounded-2xl flex items-center justify-between">
                                    <div>
                                        <span class="text-[10px] uppercase text-slate-400 font-bold block">Token Saat Ini</span>
                                        <span id="assignTokenDisplay" class="text-3xl font-mono font-extrabold text-amber-400 tracking-widest">CBT26X</span>
                                    </div>
                                    <button onclick="copyTextToClipboard(document.getElementById('assignTokenDisplay').innerText, 'Token berhasil disalin!')" class="p-3 rounded-xl bg-white/10 hover:bg-white/20 text-white transition" title="Salin Token">
                                        <i data-lucide="copy" class="w-5 h-5"></i>
                                    </button>
                                </div>

                                <div class="space-y-2 pt-2 border-t border-slate-100">
                                    <label class="block text-xs font-bold text-slate-700">Ubah Token Secara Manual</label>
                                    <div class="flex gap-2">
                                        <input type="text" id="customTokenInput" maxlength="10" placeholder="Misal: UAS2026" class="flex-1 px-3.5 py-2.5 rounded-xl border border-slate-300 font-mono font-bold uppercase text-sm focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                        <button onclick="saveManualToken()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                            Simpan
                                        </button>
                                    </div>
                                </div>

                                <button onclick="generateRandomToken()" class="w-full py-3 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition flex items-center justify-center gap-2 shadow-sm">
                                    <i data-lucide="refresh-cw" class="w-4 h-4"></i> Acak / Generate Token Baru Otomatis
                                </button>
                            </div>

                            <!-- Form Penugasan (Assign) Paket Soal Mata Pelajaran -->
                            <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="calendar-plus" class="w-4 h-4 text-indigo-600"></i> Jadwalkan &amp; Assign Paket Soal ke Siswa
                                </h3>
                                <form onsubmit="savePackageAssignment(event)" class="space-y-4">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Pilih Paket Soal dari Bank Soal *</label>
                                        <select id="assignPackageSelect" onchange="loadSelectedPackageToAssignForm()" required class="w-full px-3.5 py-2.5 rounded-xl border-2 border-indigo-200 bg-indigo-50/30 text-sm font-bold text-slate-900"></select>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Judul Sesi Ujian *</label>
                                            <input type="text" id="assignExamTitle" required placeholder="Misal: UAS Ganjil 2026" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Mata Pelajaran *</label>
                                            <input type="text" id="assignSubject" required placeholder="Misal: Akuntansi Dasar" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Tanggal Jadwal *</label>
                                            <input type="date" id="assignScheduleDate" required class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Jam Mulai - Selesai *</label>
                                            <input type="text" id="assignScheduleTime" required placeholder="08:00 - 09:30 WITA" class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm font-semibold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Durasi (Menit) *</label>
                                            <input type="number" id="assignDuration" min="1" max="360" required value="45" class="w-full px-3 py-2.5 rounded-xl border border-slate-300 text-sm font-bold">
                                        </div>
                                    </div>

                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Target Kelas / Ruangan</label>
                                            <select id="assignTargetGroup" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-semibold"></select>
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Status Penugasan (Assign)</label>
                                            <select id="assignStatusSelect" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 text-sm bg-white font-bold text-emerald-700">
                                                <option value="ASSIGNED">DI-ASSIGN (Siap Dikerjakan Siswa)</option>
                                                <option value="WAITING">TERJADWAL (Menunggu Assign)</option>
                                            </select>
                                        </div>
                                    </div>

                                    <div class="p-4 rounded-2xl bg-slate-50 border border-slate-200 grid grid-cols-1 sm:grid-cols-2 gap-3">
                                        <label class="flex items-center gap-2.5 cursor-pointer">
                                            <input type="checkbox" id="assignShuffleQuestions" class="w-4 h-4 rounded text-indigo-600">
                                            <span class="text-xs font-bold text-slate-800">Acak Urutan Soal</span>
                                        </label>
                                        <label class="flex items-center gap-2.5 cursor-pointer">
                                            <input type="checkbox" id="assignShuffleOptions" class="w-4 h-4 rounded text-indigo-600">
                                            <span class="text-xs font-bold text-slate-800">Acak Pilihan Jawaban</span>
                                        </label>
                                    </div>

                                    <div class="flex justify-end pt-1">
                                        <button type="submit" class="px-6 py-3 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs sm:text-sm font-bold transition shadow-sm flex items-center gap-2">
                                            <i data-lucide="check-circle" class="w-4 h-4"></i> Simpan &amp; Sinkronkan Penugasan
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>

                        <!-- Tabel Daftar Penugasan Mata Pelajaran -->
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">Daftar Jadwal &amp; Status Assignment Mata Pelajaran</h3>
                                <p class="text-xs text-slate-500">Ubah status antara <strong>Siap Dikerjakan (Assigned)</strong> atau <strong>Menunggu Assign</strong> kapan saja sesuai jadwal.</p>
                            </div>
                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-4 font-bold">Mata Pelajaran &amp; Paket</th>
                                            <th class="py-3 px-4 font-bold">Jadwal Pelaksanaan</th>
                                            <th class="py-3 px-4 font-bold">Target</th>
                                            <th class="py-3 px-4 font-bold">Jumlah Soal</th>
                                            <th class="py-3 px-4 font-bold">Status</th>
                                            <th class="py-3 px-4 font-bold text-right">Aksi Assignment</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminAssignmentsTableBody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 4: LOG MASUK/KELUAR SISWA & PENGAWAS             -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-logs" class="space-y-6 hidden">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Audit Log Aktivitas Masuk/Keluar &amp; Pelanggaran</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Rekam jejak real-time seluruh aktivitas login, logout, mulai ujian, selesai ujian, dan pelanggaran keluar sistem oleh Siswa &amp; Pengawas.</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-2">
                                <button onclick="exportActivityLogsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="download" class="w-3.5 h-3.5"></i> Export Log (.XLSX)
                                </button>
                                <button onclick="confirmClearActivityLogs()" class="px-3.5 py-2.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                    Bersihkan Log
                                </button>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row gap-3 justify-between">
                                <input type="text" id="logSearchInput" oninput="renderAdminLogsTable()" placeholder="Cari nama pengguna, NISN, atau aktivitas..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full sm:w-72 focus:ring-2 focus:ring-indigo-500 focus:outline-none">
                                <div class="flex flex-wrap gap-2">
                                    <select id="logRoleFilter" onchange="renderAdminLogsTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                        <option value="">Semua Peran (Siswa &amp; Pengawas)</option>
                                        <option value="siswa">Khusus Siswa</option>
                                        <option value="pengawas">Khusus Pengawas</option>
                                        <option value="admin">Administrator</option>
                                    </select>
                                    <select id="logTypeFilter" onchange="renderAdminLogsTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white">
                                        <option value="">Semua Jenis Log</option>
                                        <option value="LOGIN">Login / Masuk Sistem</option>
                                        <option value="LOGOUT">Logout / Keluar Sistem</option>
                                        <option value="EXAM_START">Mulai Ujian</option>
                                        <option value="EXAM_FINISH">Selesai Ujian</option>
                                        <option value="VIOLATION">Pelanggaran / Sirine</option>
                                        <option value="UNLOCK">Buka Kunci Layar</option>
                                    </select>
                                </div>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[540px] overflow-y-auto">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                            <th class="py-3 px-4 font-bold">Waktu</th>
                                            <th class="py-3 px-4 font-bold">Peran</th>
                                            <th class="py-3 px-4 font-bold">Nama Pengguna &amp; ID</th>
                                            <th class="py-3 px-4 font-bold">Kategori</th>
                                            <th class="py-3 px-4 font-bold">Keterangan Aktivitas</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminLogsTableBody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 5: DATA SISWA & DOWNLOAD NILAI (.XLSX)           -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-nilai" class="space-y-6 hidden">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Data Peserta &amp; Rekap Nilai Ujian</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Nilai masuk secara real-time untuk seluruh tipe soal (PG, Menjodohkan, Isian, Uraian) dan dapat diunduh dalam format Excel (.xlsx).</p>
                            </div>
                            <button onclick="downloadExcelResults()" class="px-5 py-3 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2 shadow-sm self-start">
                                <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Download Nilai Excel (.XLSX)
                            </button>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                                <input type="text" id="nilaiSearchInput" oninput="renderAdminNilaiTable()" placeholder="Cari Nama atau NISN..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                <select id="nilaiClassFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white"></select>
                                <select id="nilaiRoomFilter" onchange="renderAdminNilaiTable()" class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm bg-white"></select>
                            </div>

                            <div class="overflow-x-auto border border-slate-200 rounded-xl">
                                <table class="w-full text-left border-collapse text-xs sm:text-sm">
                                    <thead>
                                        <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                            <th class="py-3 px-3 font-bold">No</th>
                                            <th class="py-3 px-3 font-bold">Nama Peserta</th>
                                            <th class="py-3 px-3 font-bold">NISN</th>
                                            <th class="py-3 px-3 font-bold">Kelas / Ruang</th>
                                            <th class="py-3 px-3 font-bold">Mata Pelajaran</th>
                                            <th class="py-3 px-3 font-bold">Status</th>
                                            <th class="py-3 px-3 font-bold">Benar</th>
                                            <th class="py-3 px-3 font-bold">Salah</th>
                                            <th class="py-3 px-3 font-bold">Nilai Akhir</th>
                                            <th class="py-3 px-3 font-bold">Pelanggaran</th>
                                        </tr>
                                    </thead>
                                    <tbody id="adminNilaiTbody" class="divide-y divide-slate-100"></tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 6: AKUN SISWA SIAP DIGUNAKAN & KARTU LOGIN        -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-kartu" class="space-y-6 hidden">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4 no-print">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Manajemen Akun Siswa &amp; Kartu Login Resmi</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Kelola akun siswa siap digunakan berbasis NISN, edit/reset password, import Excel, dan cetak Kartu Login A4.</p>
                            </div>
                            <div class="flex flex-wrap items-center gap-2">
                                <button onclick="exportStudentAccountsExcel()" class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Export Akun (.XLSX)
                                </button>
                                <button onclick="triggerPrintLoginCards()" class="px-4 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition flex items-center gap-1.5 shadow-sm">
                                    <i data-lucide="printer" class="w-4 h-4"></i> Cetak Kartu Login (A4)
                                </button>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 xl:grid-cols-12 gap-6 no-print">
                            <div class="xl:col-span-4 space-y-5 min-w-0">
                                <div class="bg-white rounded-2xl border border-emerald-200 p-5 shadow-sm space-y-3">
                                    <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                        <i data-lucide="file-spreadsheet" class="w-4 h-4 text-emerald-600"></i> Import Data Siswa via Excel
                                    </h3>
                                    <p class="text-[11px] text-slate-500 leading-relaxed">
                                        Format kolom: <strong>Nama, NISN, Kelas, Ruangan</strong>. Username otomatis menggunakan NISN dan Password dibuat otomatis.
                                    </p>
                                    <button onclick="downloadStudentExcelTemplate()" class="w-full py-2.5 px-4 rounded-xl bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 text-xs font-bold transition flex items-center justify-center gap-2">
                                        <i data-lucide="download" class="w-4 h-4"></i> Download Template Excel (.XLSX)
                                    </button>
                                    <label class="flex flex-col items-center justify-center w-full py-3.5 px-3 border-2 border-dashed border-emerald-300 rounded-xl bg-emerald-50/30 hover:bg-emerald-50 cursor-pointer transition text-center">
                                        <span class="text-xs font-bold text-emerald-900"><i data-lucide="upload" class="w-3.5 h-3.5 inline mr-1"></i> Upload File Excel (.XLSX / .CSV)</span>
                                        <input type="file" accept=".xlsx,.xls,.csv" onchange="handleStudentExcelUpload(event)" class="hidden">
                                    </label>
                                </div>

                                <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                                    <h3 class="font-bold text-slate-900 text-sm mb-3">Tambah Akun Siswa Manual</h3>
                                    <form onsubmit="handleAddManualStudent(event)" class="space-y-3">
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">Nama Lengkap *</label>
                                            <input type="text" id="kStuName" required placeholder="Contoh: Nadia Putri" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-semibold text-slate-700 mb-1">NISN *</label>
                                            <input type="text" id="kStuNisn" required placeholder="Contoh: 0081234504" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-mono">
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Kelas *</label>
                                                <input type="text" id="kStuClass" required placeholder="XII IPA 1" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Ruangan *</label>
                                                <input type="text" id="kStuRoom" required placeholder="Ruang 01" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm">
                                            </div>
                                        </div>
                                        <div class="grid grid-cols-2 gap-2.5">
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Username (Opsional)</label>
                                                <input type="text" id="kStuUser" placeholder="Otomatis NISN" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                            <div>
                                                <label class="block text-xs font-semibold text-slate-700 mb-1">Password (Opsional)</label>
                                                <input type="text" id="kStuPass" placeholder="Otomatis Acak" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono">
                                            </div>
                                        </div>
                                        <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm">
                                            + Simpan Akun &amp; Buat Kartu
                                        </button>
                                    </form>
                                </div>
                            </div>

                            <div class="xl:col-span-8 bg-white min-w-0 rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-col space-y-4">
                                <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                                    <div>
                                        <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Akun Siswa Aktif (<span id="studentAccountCount">0</span>)</h3>
                                        <p class="text-[11px] text-slate-500">Gunakan NISN/Username dan Password di bawah untuk login ke Dashboard Siswa</p>
                                    </div>
                                    <div class="flex flex-wrap items-center gap-2">
                                        <button onclick="toggleShowStudentPasswords()" id="btnTogglePassVis" class="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5"></button>
                                        <button onclick="confirmClearAllStudents()" class="px-3 py-1.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition">
                                            Kosongkan
                                        </button>
                                    </div>
                                </div>

                                <input type="text" id="accountSearchInput" oninput="renderAdminStudentAccountTable()" placeholder="Cari berdasarkan Nama, NISN, Kelas, atau Ruangan..." class="px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm w-full focus:ring-2 focus:ring-indigo-500 focus:outline-none">

                                <div class="overflow-x-auto border border-slate-200 rounded-xl max-h-[380px] overflow-y-auto">
                                    <table class="w-full text-left border-collapse text-xs">
                                        <thead>
                                            <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600 sticky top-0">
                                                <th class="py-2.5 px-3 font-bold">Nama &amp; NISN</th>
                                                <th class="py-2.5 px-3 font-bold">Kelas / Ruang</th>
                                                <th class="py-2.5 px-3 font-bold">Username</th>
                                                <th class="py-2.5 px-3 font-bold">Password</th>
                                                <th class="py-2.5 px-3 font-bold text-right">Kelola Akun</th>
                                            </tr>
                                        </thead>
                                        <tbody id="adminStudentAccountTbody" class="divide-y divide-slate-100"></tbody>
                                    </table>
                                </div>
                            </div>
                        </div>

                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-wrap items-center justify-between gap-2 no-print">
                                <div>
                                    <h3 class="font-bold text-slate-900 text-sm sm:text-base">Pratinjau Kartu Login Peserta Siap Cetak (<span id="cardTotalCount">0</span> Kartu)</h3>
                                    <p class="text-[11px] text-slate-500">Format cetak otomatis 2 kolom per halaman A4</p>
                                </div>
                                <button onclick="triggerPrintLoginCards()" class="px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="printer" class="w-3.5 h-3.5"></i> Cetak Sekarang
                                </button>
                            </div>
                            <div id="printCardsContainer" class="grid grid-cols-1 sm:grid-cols-2 gap-4"></div>
                        </div>
                    </div>

                    <!-- ----------------------------------------------------- -->
                    <!-- MENU 7: SETELAN KESELURUHAN + PENGATURAN SEMUA DASHBOARD -->
                    <!-- ----------------------------------------------------- -->
                    <div id="adminSection-setelan" class="hidden space-y-6">
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div>
                                <h2 class="text-xl font-extrabold text-slate-900">Setelan Keseluruhan, Pengaturan Semua Dashboard &amp; Cloud Sync</h2>
                                <p class="text-xs text-slate-500 mt-0.5">Atur identitas sekolah, pengaturan Dashboard Siswa &amp; Dashboard Pengawas, keamanan single-app lockdown, berita portal, dan sinkronisasi Google Apps Script.</p>
                            </div>
                            <span class="px-3 py-1.5 rounded-xl bg-slate-900 text-white text-[11px] font-bold self-start shrink-0">
                                Development : Jovano F Kakalang, S.Ak.
                            </span>
                        </div>

                        <!-- PANEL BARU: PENGATURAN SEMUA DASHBOARD (SISWA & PENGAWAS) -->
                        <div class="bg-white rounded-2xl border-2 border-indigo-200 p-6 shadow-sm space-y-5">
                            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-100 pb-3">
                                <div>
                                    <span class="px-2.5 py-0.5 rounded-full bg-indigo-100 text-indigo-700 text-[10px] font-extrabold uppercase tracking-wider">Kontrol Terpusat</span>
                                    <h3 class="font-extrabold text-slate-900 text-base mt-1 flex items-center gap-2">
                                        <i data-lucide="layout-dashboard" class="w-4 h-4 text-indigo-600"></i> Pengaturan Dashboard Siswa &amp; Dashboard Pengawas
                                    </h3>
                                </div>
                                <button onclick="saveAllDashboardSettings()" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition shadow-sm self-start sm:self-center">
                                    Simpan Pengaturan Dashboard
                                </button>
                            </div>

                            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                                <!-- Pengaturan Dashboard Siswa -->
                                <div class="p-4 rounded-2xl bg-blue-50/40 border border-blue-200 space-y-3">
                                    <div class="font-extrabold text-blue-900 text-sm flex items-center gap-2">
                                        <i data-lucide="graduation-cap" class="w-4 h-4 text-blue-600"></i> Pengaturan Dashboard Siswa
                                    </div>
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Banner Pengumuman di Dashboard Siswa (Opsional)</label>
                                        <input type="text" id="setStuBannerMsg" placeholder="Misal: Harap periksa kembali identitas kartu Anda sebelum memulai ujian." class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs bg-white">
                                    </div>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700 pt-1">
                                        <span>Tampilkan Tab "Profil &amp; Kartu Login Saya"</span>
                                        <input type="checkbox" id="setStuShowProfileTab" class="w-4 h-4 rounded text-blue-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Tampilkan Tab "Panduan &amp; Tata Tertib"</span>
                                        <input type="checkbox" id="setStuShowGuideTab" class="w-4 h-4 rounded text-blue-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Izinkan Kontrol Ukuran Teks Soal (A- / A / A+)</span>
                                        <input type="checkbox" id="setStuAllowFontResize" class="w-4 h-4 rounded text-blue-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Tampilkan Skor Akhir &amp; Riwayat Nilai ke Siswa</span>
                                        <input type="checkbox" id="setShowScoreToStudent" class="w-4 h-4 rounded text-blue-600">
                                    </label>
                                </div>

                                <!-- Pengaturan Dashboard Pengawas -->
                                <div class="p-4 rounded-2xl bg-amber-50/40 border border-amber-200 space-y-3">
                                    <div class="font-extrabold text-amber-900 text-sm flex items-center gap-2">
                                        <i data-lucide="user-check" class="w-4 h-4 text-amber-600"></i> Pengaturan Dashboard Pengawas
                                    </div>
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Banner Instruksi di Dashboard Pengawas (Opsional)</label>
                                        <input type="text" id="setProcBannerMsg" placeholder="Misal: Pastikan seluruh peserta memasuki ruang 15 menit sebelum ujian." class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs bg-white">
                                    </div>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700 pt-1">
                                        <span>Tampilkan Token Ujian Aktif di Dashboard Pengawas</span>
                                        <input type="checkbox" id="setProcShowToken" class="w-4 h-4 rounded text-amber-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Izinkan Pengawas Input &amp; Upload Soal (Menu 3)</span>
                                        <input type="checkbox" id="setProcAllowQuestionInput" class="w-4 h-4 rounded text-amber-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Izinkan Pengawas Buka Kunci &amp; Matikan Sirine Jarak Jauh</span>
                                        <input type="checkbox" id="setProcAllowRemoteUnlock" class="w-4 h-4 rounded text-amber-600">
                                    </label>
                                </div>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                            <!-- Identitas Sekolah & Akun Admin -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="school" class="w-4 h-4 text-indigo-600"></i> Identitas Sekolah &amp; Akun Administrator
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Nama Sekolah / Lembaga</label>
                                        <input type="text" id="setSchoolName" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div>
                                        <label class="block text-xs font-bold text-slate-700 mb-1">Subjudul / Alamat &amp; Tahun Pelajaran</label>
                                        <input type="text" id="setSchoolSubtitle" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                                    </div>
                                    <div class="pt-2 border-t border-slate-100 grid grid-cols-2 gap-2.5">
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Username Admin</label>
                                            <input type="text" id="setAdminUsername" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                        <div>
                                            <label class="block text-xs font-bold text-slate-700 mb-1">Password Admin</label>
                                            <input type="text" id="setAdminPassword" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                                        </div>
                                    </div>
                                    <button onclick="saveSchoolIdentity()" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold transition">
                                        Simpan Identitas &amp; Akun Admin
                                    </button>
                                </div>
                            </div>

                            <!-- Setelan Keamanan, Single-App Lock & Sirine -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="shield-alert" class="w-4 h-4 text-red-600"></i> Keamanan Single-App Lock, Sirine &amp; Portal
                                </h3>
                                <div class="space-y-3">
                                    <div>
                                        <label class="block text-xs font-bold text-red-700 mb-1">Kode PIN Pembuka Kunci Layar (Default: 123)</label>
                                        <input type="text" id="setUnlockPin" class="w-full px-3.5 py-2 rounded-xl border border-red-300 bg-red-50/40 font-mono font-extrabold text-red-700 text-sm">
                                    </div>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700 pt-1">
                                        <span>Kunci Satu Aplikasi (Blokir Klik Kanan, F12 &amp; Inspect Element)</span>
                                        <input type="checkbox" id="setSingleAppLock" class="w-4 h-4 rounded text-red-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Wajib Mode Layar Penuh (Fullscreen) Saat Ujian</span>
                                        <input type="checkbox" id="setRequireFullscreen" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Kunci Otomatis Jika Pindah Tab / Aplikasi</span>
                                        <input type="checkbox" id="setDetectTabSwitch" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Bunyikan Sirine Alarm Keras Saat Pelanggaran</span>
                                        <input type="checkbox" id="setEnableSiren" class="w-4 h-4 rounded text-red-600">
                                    </label>
                                    <label class="flex items-center justify-between text-xs font-semibold text-slate-700">
                                        <span>Aktifkan Background Bergerak pada Portal</span>
                                        <input type="checkbox" id="setPortalAnimatedBg" class="w-4 h-4 rounded text-indigo-600">
                                    </label>
                                    <div class="flex flex-wrap gap-2 pt-1">
                                        <button onclick="saveSecuritySettings()" class="px-5 py-2.5 rounded-xl bg-slate-900 hover:bg-slate-800 text-white text-xs font-bold transition">
                                            Simpan Setelan Keamanan
                                        </button>
                                        <button type="button" onclick="testSirenSound()" class="px-4 py-2.5 rounded-xl bg-red-50 hover:bg-red-100 text-red-700 border border-red-200 text-xs font-bold transition flex items-center gap-1.5">
                                            <i data-lucide="volume-2" class="w-3.5 h-3.5"></i> Tes Suara Sirine (2 Detik)
                                        </button>
                                    </div>
                                </div>
                            </div>

                            <!-- Manajemen Akun Pengawas Ruangan -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="user-cog" class="w-4 h-4 text-amber-600"></i> Manajemen Akun Pengawas Ruangan
                                </h3>
                                <form onsubmit="handleAddProctor(event)" class="grid grid-cols-1 sm:grid-cols-2 gap-2.5 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <input type="text" id="procName" required placeholder="Nama Pengawas" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procRoom" required placeholder="Ruang Tugas (Misal: Ruang 01 / Semua)" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <input type="text" id="procUser" required placeholder="Username Login" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                    <div class="flex gap-2">
                                        <input type="text" id="procPass" required placeholder="Password" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-mono">
                                        <button type="submit" class="px-3 py-2 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold shrink-0">
                                            + Tambah
                                        </button>
                                    </div>
                                </form>
                                <div id="proctorListContainer" class="space-y-2 max-h-52 overflow-y-auto"></div>
                            </div>

                            <!-- Kelola Berita Portal & Galeri -->
                            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="newspaper" class="w-4 h-4 text-blue-600"></i> Kelola Berita Pendidikan &amp; Galeri Portal
                                </h3>
                                <form onsubmit="handleAddPortalNews(event)" class="space-y-2 bg-blue-50/40 p-3.5 rounded-xl border border-blue-200">
                                    <div class="text-xs font-bold text-blue-900">Tambah Berita Terkini / Pengumuman Portal</div>
                                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                                        <input type="text" id="newsTitleInput" required placeholder="Judul Berita Pendidikan..." class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                        <select id="newsCategoryInput" class="px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white font-semibold">
                                            <option value="Pendidikan Nasional">Pendidikan Nasional</option>
                                            <option value="Tips Ujian & Literasi">Tips Ujian &amp; Literasi</option>
                                            <option value="Pengumuman Sekolah">Pengumuman Sekolah</option>
                                        </select>
                                    </div>
                                    <textarea id="newsContentInput" required rows="2" placeholder="Isi lengkap berita atau pengumuman..." class="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white"></textarea>
                                    <button type="submit" class="px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold">
                                        + Publikasikan Berita ke Portal
                                    </button>
                                </form>

                                <form onsubmit="handleAddGalleryItem(event)" class="space-y-2 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
                                    <div class="text-xs font-bold text-slate-800">Tambah Foto Galeri Kegiatan</div>
                                    <input type="text" id="galTitle" required placeholder="Judul Foto Kegiatan (Misal: Pelaksanaan Ujian Sesi 1)" class="w-full px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                    <div class="flex gap-2">
                                        <input type="url" id="galUrl" placeholder="URL Gambar Eksternal (Kosongkan untuk Ilustrasi SVG)" class="flex-1 px-3 py-2 rounded-lg border border-slate-300 text-xs bg-white">
                                        <button type="submit" class="px-3.5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold shrink-0">
                                            + Foto
                                        </button>
                                    </div>
                                </form>
                            </div>
                        </div>

                        <!-- Koneksi Google Apps Script & Backup JSON -->
                        <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4">
                            <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                                <h3 class="font-bold text-slate-900 text-base flex items-center gap-2">
                                    <i data-lucide="cloud" class="w-4 h-4 text-emerald-600"></i> Integrasi Google Apps Script Cloud &amp; Cadangan Database
                                </h3>
                                <button onclick="openGasCodeModal()" class="px-3.5 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-bold transition self-start">
                                    Lihat Kode Apps Script (Code.gs)
                                </button>
                            </div>
                            <div class="grid grid-cols-1 md:grid-cols-12 gap-3 items-center">
                                <div class="md:col-span-9">
                                    <label class="block text-[11px] font-bold text-slate-500 mb-1">Endpoint Web App Google Apps Script Aktif:</label>
                                    <input type="text" id="setGasUrlDisplay" readonly class="w-full px-3.5 py-2 rounded-xl border border-emerald-200 bg-emerald-50/40 text-xs font-mono text-emerald-900">
                                </div>
                                <div class="md:col-span-3 md:pt-5">
                                    <button onclick="syncFromCloud(true); showToast('Sinkronisasi cloud berhasil dijalankan!', 'success');" class="w-full py-2.5 px-4 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold transition flex items-center justify-center gap-1.5">
                                        <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i> Sinkronkan Cloud
                                    </button>
                                </div>
                            </div>

                            <div class="pt-3 border-t border-slate-100 flex flex-wrap gap-2">
                                <button onclick="exportDatabaseJSON()" class="px-3.5 py-2 rounded-xl bg-blue-50 hover:bg-blue-100 text-blue-700 text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="download" class="w-3.5 h-3.5"></i> Backup JSON
                                </button>
                                <label class="px-3.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer">
                                    <i data-lucide="upload" class="w-3.5 h-3.5"></i> Restore JSON
                                    <input type="file" accept=".json" onchange="importDatabaseJSON(event)" class="hidden">
                                </label>
                                <button onclick="confirmCleanStartEmpty()" class="px-3.5 py-2 rounded-xl bg-amber-50 hover:bg-amber-100 text-amber-800 border border-amber-200 text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="eraser" class="w-3.5 h-3.5"></i> Mulai Kosong
                                </button>
                                <button onclick="confirmFactoryReset()" class="px-3.5 py-2 rounded-xl bg-red-50 hover:bg-red-100 text-red-600 text-xs font-bold transition flex items-center gap-1.5">
                                    <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i> Reset Bawaan
                                </button>
                            </div>
                        </div>
                    </div>

                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW D: DASHBOARD PENGAWAS RUANG (3 MENU TETAP UTUH + SINKRON)    -->
        <!-- ================================================================= -->
        <section id="view-pengawas" class="hidden max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 sm:py-8 space-y-6">
            <!-- Banner Instruksi Khusus Pengawas (Diatur Admin) -->
            <div id="procCustomNoticeBanner" class="hidden bg-amber-50 border border-amber-200 rounded-2xl p-4 flex items-start gap-3 text-xs text-amber-950">
                <i data-lucide="megaphone" class="w-5 h-5 text-amber-600 shrink-0 mt-0.5"></i>
                <div>
                    <span class="font-extrabold uppercase tracking-wider text-[10px] text-amber-700 block">Instruksi Administrator untuk Pengawas</span>
                    <p id="procCustomNoticeText" class="font-medium mt-0.5 leading-relaxed"></p>
                </div>
            </div>

            <!-- Header Pengawas -->
            <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
                <div>
                    <div class="flex flex-wrap items-center gap-2">
                        <span class="px-2.5 py-0.5 rounded-md bg-amber-100 text-amber-800 text-xs font-bold">DASHBOARD PENGAWAS RUANG</span>
                        <span id="proctorAssignedRoomBadge" class="px-2.5 py-0.5 rounded-md bg-slate-100 text-slate-700 text-xs font-semibold">Semua Ruangan</span>
                    </div>
                    <h2 id="proctorDashboardName" class="text-2xl font-extrabold text-slate-900 mt-1">Pengawas Ujian</h2>
                    <p class="text-xs text-slate-500 mt-0.5">Terhubung real-time dengan Administrator &amp; Peserta Ujian di ruangan.</p>
                </div>

                <!-- Kartu Token Ruang Aktif -->
                <div id="proctorHeaderTokenBox" class="bg-slate-900 text-white px-5 py-3.5 rounded-2xl flex items-center gap-4 self-start md:self-center">
                    <div>
                        <span class="text-[10px] uppercase text-emerald-400 font-bold block">Token Sinkron Admin</span>
                        <span id="proctorTokenDisplay" class="text-2xl font-mono font-extrabold text-amber-400 tracking-widest">------</span>
                    </div>
                    <button onclick="copyTextToClipboard(document.getElementById('proctorTokenDisplay').innerText, 'Token berhasil disalin!')" class="px-3 py-2 rounded-xl bg-white/10 hover:bg-white/20 text-xs font-bold transition flex items-center gap-1.5">
                        <i data-lucide="copy" class="w-4 h-4"></i> Salin
                    </button>
                </div>
            </div>

            <!-- 3 MENU UTAMA PENGAWAS (TIDAK DIUBAH) -->
            <div class="flex flex-wrap gap-2 bg-white p-2 rounded-2xl border border-slate-200 shadow-sm">
                <button onclick="switchProctorMenu('kehadiran')" id="procMenuBtn-kehadiran" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-amber-500 text-white transition flex items-center gap-2">
                    <i data-lucide="arrow-left-right" class="w-4 h-4"></i> 1. Pantau Siswa Keluar/Masuk &amp; Token
                </button>
                <button onclick="switchProctorMenu('pelanggaran')" id="procMenuBtn-pelanggaran" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="siren" class="w-4 h-4"></i> 2. Pelanggaran &amp; Buka Kunci Sirine
                    <span id="proctorViolationCountPill" class="px-2 py-0.5 rounded-full bg-red-100 text-red-700 text-[10px] font-extrabold">0</span>
                </button>
                <button onclick="switchProctorMenu('soal')" id="procMenuBtn-soal" class="px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2">
                    <i data-lucide="file-plus-2" class="w-4 h-4"></i> 3. Input &amp; Upload Soal (Word / PDF)
                </button>
            </div>

            <!-- MENU 1 PENGAWAS: PANTAU SISWA KELUAR MASUK & TOKEN SINKRON -->
            <div id="proctorSection-kehadiran" class="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-8 bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-4 min-w-0">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div>
                            <h3 class="font-bold text-slate-900 text-base">Status Kehadiran &amp; Pengerjaan Siswa Real-Time</h3>
                            <p class="text-xs text-slate-500">Memantau siswa yang masuk dashboard, sedang mengerjakan soal, atau sudah selesai.</p>
                        </div>
                        <span class="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-bold self-start">
                            <span class="w-2 h-2 rounded-full bg-emerald-500 animate-ping"></span> Live Sync
                        </span>
                    </div>

                    <div id="proctorStudentCards" class="space-y-3"></div>
                </div>

                <!-- Panel Token & Log Keluar Masuk Siswa Ruang -->
                <div class="lg:col-span-4 space-y-4 min-w-0">
                    <div id="proctorSideTokenCard" class="bg-slate-900 text-white rounded-2xl p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <span class="text-xs font-bold uppercase tracking-wider text-amber-300">Token Ujian Aktif</span>
                            <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 text-[10px] font-bold">SINKRON ADMIN</span>
                        </div>
                        <div class="text-3xl font-mono font-extrabold tracking-widest text-white" id="proctorCardTokenBig">------</div>
                        <p class="text-[11px] text-slate-300 leading-relaxed">
                            Berikan token ini kepada peserta di ruangan saat mereka siap menekan tombol <strong>Mulai Ujian</strong> pada mata pelajaran yang telah di-assign.
                        </p>
                        <div class="pt-2 border-t border-white/10 flex justify-between text-xs">
                            <span class="text-slate-400">PIN Buka Kunci Layar:</span>
                            <span id="proctorPinDisplay" class="font-mono font-extrabold text-amber-400">123</span>
                        </div>
                    </div>

                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3">
                        <h4 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-blue-600"></i> Log Keluar / Masuk Siswa Terbaru
                        </h4>
                        <div id="proctorInOutLogList" class="space-y-2 max-h-72 overflow-y-auto text-xs"></div>
                    </div>
                </div>
            </div>

            <!-- MENU 2 PENGAWAS: PELANGGARAN & KENDALI SIRINE -->
            <div id="proctorSection-pelanggaran" class="hidden space-y-6">
                <div class="bg-white rounded-2xl border border-red-200 p-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                        <h3 class="text-lg font-extrabold text-slate-900 flex items-center gap-2">
                            <i data-lucide="siren" class="w-5 h-5 text-red-600"></i> Monitoring Pelanggaran &amp; Layar Terkunci (Sirine)
                        </h3>
                        <p class="text-xs text-slate-500 mt-0.5">Daftar peserta yang terdeteksi keluar dari halaman ujian, berpindah tab, atau membuka aplikasi lain.</p>
                    </div>
                    <div class="bg-red-50 border border-red-200 px-4 py-2 rounded-xl text-xs text-red-900">
                        PIN Pembuka Kunci Manual: <strong id="proctorPinDisplay2" class="font-mono text-sm font-extrabold">123</strong>
                    </div>
                </div>

                <!-- Siswa Sedang Terkunci Saat Ini -->
                <div id="proctorActiveLockedBox" class="space-y-3"></div>

                <!-- Riwayat Log Pelanggaran -->
                <div class="bg-white rounded-2xl border border-slate-200 p-6 shadow-sm space-y-3">
                    <h4 class="font-bold text-slate-900 text-sm">Riwayat Catatan Pelanggaran Peserta</h4>
                    <div class="overflow-x-auto border border-slate-200 rounded-xl">
                        <table class="w-full text-left border-collapse text-xs sm:text-sm">
                            <thead>
                                <tr class="bg-slate-50 border-b border-slate-200 text-[11px] uppercase text-slate-600">
                                    <th class="py-3 px-4 font-bold">Waktu</th>
                                    <th class="py-3 px-4 font-bold">Peserta &amp; NISN</th>
                                    <th class="py-3 px-4 font-bold">Keterangan Pelanggaran</th>
                                </tr>
                            </thead>
                            <tbody id="proctorViolationLogsTbody" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- MENU 3 PENGAWAS: INPUT & UPLOAD SOAL MULTI-TIPE + FOTO JPG/PNG -->
            <div id="proctorSection-soal" class="hidden grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
                <div class="lg:col-span-5 space-y-5 min-w-0">
                    <!-- Upload Word / PDF Pengawas -->
                    <div class="bg-white rounded-2xl border border-amber-200 p-5 shadow-sm space-y-3">
                        <div class="flex items-center justify-between">
                            <h3 class="font-bold text-slate-900 text-sm flex items-center gap-2">
                                <i data-lucide="file-up" class="w-4 h-4 text-amber-600"></i> Upload File Word (.docx) / PDF (.pdf)
                            </h3>
                            <button onclick="downloadSampleQuestionFormat()" class="text-[11px] font-bold text-amber-700 hover:underline">
                                Unduh Contoh Format
                            </button>
                        </div>

                        <div>
                            <label class="block text-xs font-bold text-slate-700 mb-1">Pilih Paket Mata Pelajaran Tujuan:</label>
                            <select id="proctorPackageSelector" onchange="selectProctorQuestionPackage(this.value)" class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs sm:text-sm font-bold bg-white"></select>
                        </div>

                        <label class="flex flex-col items-center justify-center w-full h-28 border-2 border-dashed border-amber-300 rounded-2xl bg-amber-50/40 hover:bg-amber-50 cursor-pointer transition p-4 text-center">
                            <i data-lucide="upload-cloud" class="w-7 h-7 text-amber-600 mb-1"></i>
                            <span class="text-xs font-bold text-amber-900">Klik Pilih File .DOCX, .PDF, atau .TXT</span>
                            <span class="text-[10px] text-slate-500 mt-0.5">Langsung tersinkron ke Bank Soal Administrator</span>
                            <input type="file" accept=".docx,.pdf,.txt,.csv" onchange="handleDocumentQuestionUpload(event, 'pengawas')" class="hidden">
                        </label>
                    </div>

                    <!-- Form Input Soal Manual oleh Pengawas -->
                    <div class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
                        <h3 class="font-bold text-slate-900 text-sm mb-3">Input Soal Manual (Sinkron ke Bank Soal)</h3>
                        <form onsubmit="handleAddManualQuestion(event, 'pengawas')" class="space-y-3">
                            <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                                <div>
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Tipe Soal *</label>
                                    <select id="proctorQType" onchange="toggleQuestionTypeFields('proctor')" class="w-full px-3 py-2 rounded-xl border border-amber-300 bg-amber-50/40 text-xs font-bold text-amber-900">
                                        <option value="pg">Pilihan Ganda (A-E)</option>
                                        <option value="menjodohkan">Menjodohkan (Pasangan)</option>
                                        <option value="isian">Isian Singkat</option>
                                        <option value="uraian">Uraian / Esai</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Bobot Poin *</label>
                                    <input type="number" id="proctorQPoints" value="10" min="1" max="100" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                                </div>
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-700 mb-1">Teks Pertanyaan *</label>
                                <textarea id="proctorQText" required rows="2" placeholder="Tuliskan pertanyaan soal..." class="w-full p-2.5 rounded-xl border border-slate-300 text-xs sm:text-sm focus:ring-2 focus:ring-amber-500 focus:outline-none"></textarea>
                            </div>

                            <!-- Foto Soal JPG/PNG Pengawas -->
                            <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                                <div class="flex items-center justify-between">
                                    <label class="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                                        <i data-lucide="image-plus" class="w-3.5 h-3.5 text-amber-600"></i> Lampirkan Foto Soal (JPG / PNG)
                                    </label>
                                    <button type="button" id="proctorQImgClearBtn" onclick="clearQuestionImageInput('proctor')" class="hidden text-[11px] font-bold text-red-600 hover:underline">Hapus Foto</button>
                                </div>
                                <input type="file" id="proctorQImageFile" accept="image/jpeg,image/png" onchange="handleQuestionImageSelect(event, 'proctor')" class="block w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-amber-100 file:text-amber-800 hover:file:bg-amber-200">
                                <input type="hidden" id="proctorQImageData">
                                <div id="proctorQImagePreviewBox" class="hidden pt-1">
                                    <img id="proctorQImagePreview" src="" alt="Pratinjau Foto Soal" class="max-h-36 rounded-lg border border-slate-300 object-contain mx-auto">
                                </div>
                            </div>

                            <div id="proctorFields-pg" class="space-y-2">
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">A.</span>
                                    <input type="text" id="proctorQOptA" placeholder="Opsi A" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">B.</span>
                                    <input type="text" id="proctorQOptB" placeholder="Opsi B" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">C.</span>
                                    <input type="text" id="proctorQOptC" placeholder="Opsi C" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">D.</span>
                                    <input type="text" id="proctorQOptD" placeholder="Opsi D" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="flex items-center gap-2">
                                    <span class="w-6 text-xs font-bold text-slate-500">E.</span>
                                    <input type="text" id="proctorQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs">
                                </div>
                                <div class="pt-1">
                                    <label class="block text-xs font-semibold text-slate-700 mb-1">Kunci Jawaban</label>
                                    <select id="proctorQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-amber-800">
                                        <option value="A">Opsi A</option>
                                        <option value="B">Opsi B</option>
                                        <option value="C">Opsi C</option>
                                        <option value="D">Opsi D</option>
                                        <option value="E">Opsi E</option>
                                    </select>
                                </div>
                            </div>

                            <div id="proctorFields-menjodohkan" class="hidden space-y-2 p-3 rounded-xl bg-amber-50/60 border border-amber-200">
                                <div class="flex items-center justify-between">
                                    <span class="text-xs font-bold text-amber-900">Pasangan Pernyataan (Kiri = Kanan)</span>
                                    <button type="button" onclick="addMatchingPairRow('proctorMatchingRows')" class="px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-bold">+ Pasangan</button>
                                </div>
                                <div id="proctorMatchingRows" class="space-y-2"></div>
                            </div>

                            <div id="proctorFields-isian" class="hidden space-y-1.5 p-3 rounded-xl bg-emerald-50/60 border border-emerald-200">
                                <label class="block text-xs font-bold text-emerald-900">Kunci Jawaban Isian Singkat *</label>
                                <input type="text" id="proctorQShortAnswer" placeholder="Contoh: 17 Agustus 1945" class="w-full px-3 py-2 rounded-lg border border-emerald-300 text-xs bg-white">
                            </div>

                            <div id="proctorFields-uraian" class="hidden space-y-1.5 p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                                <label class="block text-xs font-bold text-blue-900">Pedoman / Kata Kunci Uraian</label>
                                <textarea id="proctorQEssayRubric" rows="2" placeholder="Kata kunci penilaian uraian..." class="w-full px-3 py-2 rounded-lg border border-blue-300 text-xs bg-white"></textarea>
                            </div>

                            <button type="submit" class="w-full py-2.5 px-4 rounded-xl bg-amber-500 hover:bg-amber-600 text-white text-xs font-bold transition shadow-sm">
                                + Simpan Soal ke Paket Terpilih
                            </button>
                        </form>
                    </div>
                </div>

                <div class="lg:col-span-7 bg-white rounded-2xl border border-slate-200 p-5 shadow-sm space-y-3 min-w-0">
                    <div class="flex items-center justify-between pb-3 border-b border-slate-100">
                        <div>
                            <h3 class="font-bold text-slate-900 text-sm sm:text-base">Daftar Soal Paket Terpilih (<span id="proctorPkgSoalCount">0</span> Soal)</h3>
                            <p class="text-[11px] text-slate-500">Tersinkronisasi langsung dengan Bank Soal Administrator</p>
                        </div>
                    </div>
                    <div id="proctorSoalListContainer" class="space-y-3 max-h-[700px] overflow-y-auto pr-1"></div>
                </div>
            </div>
        </section>

        <!-- ================================================================= -->
        <!-- VIEW E: HALAMAN UJIAN SISWA (4 TIPE SOAL + FOTO + LOCKDOWN)       -->
        <!-- ================================================================= -->
        <section id="view-siswa" class="min-h-screen bg-slate-100 flex flex-col no-select hidden">
            <div class="bg-slate-900 text-white px-4 sm:px-6 py-3.5 shadow-md sticky top-0 z-30">
                <div class="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
                    <div class="flex items-center gap-3 min-w-0">
                        <div class="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center font-extrabold shrink-0">
                            <i data-lucide="shield-check" class="w-5 h-5"></i>
                        </div>
                        <div class="min-w-0">
                            <h2 id="studentExamTitle" class="font-bold text-sm sm:text-base leading-tight truncate">Ujian Berbasis Komputer</h2>
                            <p class="text-xs text-slate-300 truncate">
                                <strong id="studentExamName" class="text-white">-</strong> • NISN: <span id="studentExamNisn">-</span> • <span id="studentExamRoom">-</span>
                            </p>
                        </div>
                    </div>

                    <div class="flex items-center gap-2 sm:gap-3 shrink-0">
                        <div id="examFontResizeControls" class="hidden sm:flex items-center bg-slate-800 border border-slate-700 rounded-xl p-1">
                            <button onclick="changeExamFontSize('sm')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perkecil Teks">A-</button>
                            <button onclick="changeExamFontSize('base')" class="px-2 py-1 rounded-lg text-xs font-bold bg-slate-700" title="Ukuran Normal">A</button>
                            <button onclick="changeExamFontSize('lg')" class="px-2 py-1 rounded-lg text-xs font-bold hover:bg-slate-700" title="Perbesar Teks">A+</button>
                        </div>

                        <div class="bg-slate-800 border border-slate-700 px-3.5 py-1.5 rounded-xl flex items-center gap-2">
                            <i data-lucide="clock" class="w-4 h-4 text-amber-400"></i>
                            <span id="studentTimerDisplay" class="font-mono font-extrabold text-base sm:text-lg text-amber-400">00:00</span>
                        </div>

                        <button onclick="confirmSubmitExam()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs sm:text-sm font-bold transition shadow">
                            Selesai Ujian
                        </button>
                    </div>
                </div>
            </div>

            <div class="max-w-7xl w-full mx-auto px-4 sm:px-6 py-6 grid grid-cols-1 lg:grid-cols-12 gap-6 flex-grow">
                <div class="lg:col-span-8 bg-white rounded-3xl border border-slate-200 p-5 sm:p-8 shadow-sm flex flex-col justify-between min-w-0">
                    <div>
                        <div class="flex flex-wrap items-center justify-between gap-2 pb-4 mb-5 border-b border-slate-100">
                            <div class="flex flex-wrap items-center gap-2">
                                <span class="px-3.5 py-1.5 rounded-xl bg-blue-50 text-blue-700 font-extrabold text-xs sm:text-sm">
                                    SOAL NO. <span id="examCurrentNum">1</span> DARI <span id="examTotalNum">1</span>
                                </span>
                                <span id="examQuestionTypeBadge" class="px-3 py-1.5 rounded-xl bg-indigo-50 text-indigo-700 font-bold text-xs">
                                    Pilihan Ganda
                                </span>
                            </div>
                            <button id="doubtToggleBtn" onclick="toggleDoubtQuestion()" class="px-3.5 py-1.5 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-bold transition flex items-center gap-1.5">
                                <i data-lucide="flag" class="w-3.5 h-3.5"></i> <span>Tandai Ragu-Ragu</span>
                            </button>
                        </div>

                        <!-- Gambar Soal Jika Ada (JPG/PNG) -->
                        <div id="examQuestionImageBox" class="hidden mb-5 bg-slate-50 border border-slate-200 rounded-2xl p-3 text-center">
                            <img id="examQuestionImage" src="" alt="Gambar Soal" onclick="openImageLightbox(this.src)" class="max-h-72 mx-auto rounded-xl object-contain cursor-zoom-in">
                            <span class="text-[11px] text-slate-400 mt-1.5 inline-flex items-center gap-1">
                                <i data-lucide="zoom-in" class="w-3 h-3"></i> Klik gambar untuk memperbesar
                            </span>
                        </div>

                        <div id="examQuestionBody" class="font-semibold text-slate-900 leading-relaxed mb-6 text-base break-anywhere"></div>
                        <div id="examAnswerWorkspace" class="space-y-3"></div>
                    </div>

                    <div class="flex items-center justify-between pt-8 mt-8 border-t border-slate-100">
                        <button id="btnPrevQuestion" onclick="goPrevQuestion()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs sm:text-sm font-bold transition flex items-center gap-2 disabled:opacity-40">
                            <i data-lucide="chevron-left" class="w-4 h-4"></i> Sebelumnya
                        </button>
                        <button id="btnNextQuestion" onclick="goNextQuestion()" class="px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2">
                            Selanjutnya <i data-lucide="chevron-right" class="w-4 h-4"></i>
                        </button>
                    </div>
                </div>

                <div class="lg:col-span-4 space-y-4 min-w-0">
                    <div class="bg-white rounded-3xl border border-slate-200 p-6 shadow-sm">
                        <h3 class="font-bold text-slate-900 text-sm mb-3">Navigasi Nomor Soal</h3>
                        <div id="examNumberGrid" class="grid grid-cols-5 gap-2"></div>
                        <div class="mt-5 pt-4 border-t border-slate-100 grid grid-cols-3 gap-2 text-[11px] text-slate-600">
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-blue-600 inline-block"></span> Dijawab</span>
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-amber-400 inline-block"></span> Ragu</span>
                            <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded bg-slate-100 border border-slate-300 inline-block"></span> Kosong</span>
                        </div>
                    </div>

                    <div class="bg-red-50 border border-red-200 rounded-2xl p-4 text-xs text-red-900 space-y-1">
                        <div class="font-bold flex items-center gap-1.5 text-red-700">
                            <i data-lucide="siren" class="w-4 h-4"></i> Proteksi Layar &amp; Sirine Aktif
                        </div>
                        <p class="leading-relaxed">
                            Dilarang berpindah tab, membuka aplikasi lain, Inspect Element, atau keluar halaman ujian. Pelanggaran akan <strong>mengunci layar &amp; membunyikan sirine alarm otomatis</strong>.
                        </p>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- ===================================================================== -->
    <!-- 5. MODAL LOGIN MULTI-PERAN (SISWA KE DASHBOARD, PENGAWAS, ADMIN)      -->
    <!-- ===================================================================== -->
    <div id="loginModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl border border-slate-100 relative">
            <button onclick="closeLoginModal()" class="absolute top-5 right-5 w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition">
                <i data-lucide="x" class="w-4 h-4"></i>
            </button>

            <div class="mb-6">
                <span id="loginModalBadge" class="px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold uppercase">Peserta Ujian</span>
                <h2 id="loginModalTitle" class="text-2xl font-extrabold text-slate-900 mt-2">Login Dashboard Siswa</h2>
                <p id="loginModalDesc" class="text-xs text-slate-500 mt-1">Masukkan NISN/Username dan Password sesuai Kartu Login Peserta.</p>
            </div>

            <form onsubmit="processLogin(event)" class="space-y-4">
                <input type="hidden" id="loginRoleField" value="siswa">

                <div>
                    <label id="loginUsernameLabel" class="block text-xs font-bold text-slate-700 mb-1">NISN / Username Peserta</label>
                    <input type="text" id="loginUsernameField" required placeholder="Masukkan NISN atau Username..." class="w-full px-4 py-3 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Password</label>
                    <input type="password" id="loginPasswordField" required placeholder="Masukkan password..." class="w-full px-4 py-3 rounded-xl border border-slate-300 text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none">
                </div>

                <button type="submit" id="loginSubmitButton" class="w-full py-3.5 px-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-lg mt-2">
                    Masuk ke Dashboard
                </button>
            </form>

            <div class="mt-5 pt-3 border-t border-slate-100 text-center text-[11px] text-slate-400">
                CBT-PRO Security System • Development : <strong class="text-slate-600">Jovano F Kakalang, S.Ak.</strong>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 6. MODAL INPUT TOKEN SAAT SISWA SUDAH SIAP MEMULAI UJIAN              -->
    <!-- ===================================================================== -->
    <div id="studentTokenEntryModal" class="fixed inset-0 bg-slate-900/70 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl border border-slate-200 relative">
            <button onclick="closeStudentTokenModal()" class="absolute top-5 right-5 w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center transition">
                <i data-lucide="x" class="w-4 h-4"></i>
            </button>

            <div class="space-y-4">
                <div class="w-12 h-12 rounded-2xl bg-blue-100 text-blue-600 flex items-center justify-center">
                    <i data-lucide="key-round" class="w-6 h-6"></i>
                </div>
                <div>
                    <span class="text-[11px] font-extrabold uppercase tracking-wider text-blue-600">Konfirmasi Mulai Ujian</span>
                    <h3 id="tokenModalExamTitle" class="text-xl font-extrabold text-slate-900 mt-0.5">Mata Pelajaran</h3>
                    <p id="tokenModalExamMeta" class="text-xs text-slate-500 mt-1">Durasi: 45 Menit • Jumlah Soal: 5 Butir</p>
                </div>

                <div class="p-3.5 rounded-2xl bg-amber-50 border border-amber-200 text-xs text-amber-900 leading-relaxed">
                    <strong>Perhatian:</strong> Setelah token dikonfirmasi, waktu hitung mundur berjalan dan sistem akan mengaktifkan <strong>Mode Layar Penuh + Single-App Lock + Sirine Anti-Curang</strong>.
                </div>

                <form onsubmit="confirmStartExamWithToken(event)" class="space-y-4">
                    <input type="hidden" id="tokenModalPackageId">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1.5">Masukkan Token Ujian dari Pengawas *</label>
                        <input type="text" id="studentExamTokenInput" required maxlength="10" placeholder="CONTOH: CBT26X" class="w-full px-4 py-3.5 rounded-xl border-2 border-blue-400 bg-blue-50/40 font-mono font-extrabold text-center text-lg uppercase tracking-widest text-blue-950 focus:outline-none focus:ring-4 focus:ring-blue-500/30">
                    </div>
                    <div class="flex gap-2">
                        <button type="button" onclick="closeStudentTokenModal()" class="flex-1 py-3 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold transition">
                            Batal
                        </button>
                        <button type="submit" class="flex-1 py-3 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition shadow-md">
                            Mulai Kerjakan Soal →
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 7. MODAL EDIT AKUN & RESET PASSWORD SISWA                             -->
    <!-- ===================================================================== -->
    <div id="editStudentModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-200">
            <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                <h3 class="text-lg font-extrabold text-slate-900">Edit Akun &amp; Password Siswa</h3>
                <button onclick="closeEditStudentModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <form onsubmit="saveEditedStudent(event)" class="space-y-3">
                <input type="hidden" id="editStuId">
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Nama Lengkap *</label>
                    <input type="text" id="editStuName" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">NISN *</label>
                    <input type="text" id="editStuNisn" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm font-mono">
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Kelas *</label>
                        <input type="text" id="editStuClass" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Ruangan *</label>
                        <input type="text" id="editStuRoom" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-sm">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3 pt-1">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Username Login *</label>
                        <input type="text" id="editStuUser" required class="w-full px-3.5 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold">
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Password *</label>
                        <div class="flex gap-1.5">
                            <input type="text" id="editStuPass" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-mono font-bold text-blue-700">
                            <button type="button" onclick="generateRandomStudentPassInModal()" class="px-2.5 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-700 text-xs font-bold" title="Acak Password Baru">
                                <i data-lucide="refresh-cw" class="w-3.5 h-3.5"></i>
                            </button>
                        </div>
                    </div>
                </div>
                <div class="flex justify-end gap-2 pt-3">
                    <button type="button" onclick="closeEditStudentModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Batal</button>
                    <button type="submit" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold">Simpan Perubahan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 8. MODAL EDIT SOAL (MENDUKUNG 4 TIPE SOAL + FOTO JPG/PNG)             -->
    <!-- ===================================================================== -->
    <div id="editQuestionModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 hidden items-center justify-center p-4">
        <div class="bg-white rounded-3xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto">
            <div class="flex items-center justify-between pb-3 mb-4 border-b border-slate-100">
                <h3 class="text-lg font-extrabold text-slate-900">Edit Soal Ujian</h3>
                <button onclick="closeEditQuestionModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <form onsubmit="saveEditedQuestion(event)" class="space-y-3">
                <input type="hidden" id="editQPkgId">
                <input type="hidden" id="editQId">
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Tipe Soal</label>
                        <select id="editQType" onchange="toggleQuestionTypeFields('edit')" class="w-full px-3 py-2 rounded-xl border border-indigo-300 bg-indigo-50/40 text-xs font-bold text-indigo-900">
                            <option value="pg">Pilihan Ganda (A-E)</option>
                            <option value="menjodohkan">Menjodohkan</option>
                            <option value="isian">Isian Singkat</option>
                            <option value="uraian">Uraian / Esai</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-1">Bobot Poin</label>
                        <input type="number" id="editQPoints" min="1" max="100" required class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold">
                    </div>
                </div>
                <div>
                    <label class="block text-xs font-bold text-slate-700 mb-1">Pertanyaan</label>
                    <textarea id="editQText" required rows="3" class="w-full p-2.5 rounded-xl border border-slate-300 text-sm"></textarea>
                </div>

                <!-- Foto Soal JPG / PNG di Modal Edit -->
                <div class="p-3 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                    <div class="flex items-center justify-between">
                        <label class="text-xs font-bold text-slate-700">Foto Soal (JPG / PNG)</label>
                        <button type="button" id="editQImgClearBtn" onclick="clearQuestionImageInput('edit')" class="hidden text-[11px] font-bold text-red-600 hover:underline">Hapus Foto</button>
                    </div>
                    <input type="file" id="editQImageFile" accept="image/jpeg,image/png" onchange="handleQuestionImageSelect(event, 'edit')" class="block w-full text-xs text-slate-500 file:mr-2 file:py-1.5 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-bold file:bg-indigo-100 file:text-indigo-700">
                    <input type="hidden" id="editQImageData">
                    <div id="editQImagePreviewBox" class="hidden pt-1">
                        <img id="editQImagePreview" src="" alt="Preview" class="max-h-36 rounded-lg border border-slate-300 object-contain mx-auto">
                    </div>
                </div>

                <div id="editFields-pg" class="space-y-2">
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">A.</span>
                        <input type="text" id="editQOptA" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">B.</span>
                        <input type="text" id="editQOptB" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">C.</span>
                        <input type="text" id="editQOptC" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">D.</span>
                        <input type="text" id="editQOptD" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="flex items-center gap-2">
                        <span class="w-6 text-xs font-bold">E.</span>
                        <input type="text" id="editQOptE" placeholder="Opsi E (Opsional)" class="flex-1 px-3 py-2 rounded-xl border border-slate-300 text-xs">
                    </div>
                    <div class="pt-1">
                        <label class="block text-xs font-bold text-slate-700 mb-1">Kunci Jawaban</label>
                        <select id="editQAnswer" class="w-full px-3 py-2 rounded-xl border border-slate-300 text-xs font-bold text-indigo-700">
                            <option value="A">Opsi A</option>
                            <option value="B">Opsi B</option>
                            <option value="C">Opsi C</option>
                            <option value="D">Opsi D</option>
                            <option value="E">Opsi E</option>
                        </select>
                    </div>
                </div>

                <div id="editFields-menjodohkan" class="hidden space-y-2 p-3 rounded-xl bg-amber-50/60 border border-amber-200">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-amber-900">Pasangan Pernyataan (Kiri = Kanan)</span>
                        <button type="button" onclick="addMatchingPairRow('editMatchingRows')" class="px-2.5 py-1 rounded-lg bg-amber-500 hover:bg-amber-600 text-white text-[11px] font-bold">+ Pasangan</button>
                    </div>
                    <div id="editMatchingRows" class="space-y-2"></div>
                </div>

                <div id="editFields-isian" class="hidden space-y-1.5 p-3 rounded-xl bg-emerald-50/60 border border-emerald-200">
                    <label class="block text-xs font-bold text-emerald-900">Kunci Jawaban Isian Singkat</label>
                    <input type="text" id="editQShortAnswer" class="w-full px-3 py-2 rounded-lg border border-emerald-300 text-xs bg-white">
                </div>

                <div id="editFields-uraian" class="hidden space-y-1.5 p-3 rounded-xl bg-blue-50/60 border border-blue-200">
                    <label class="block text-xs font-bold text-blue-900">Pedoman / Kata Kunci Uraian</label>
                    <textarea id="editQEssayRubric" rows="2" class="w-full px-3 py-2 rounded-lg border border-blue-300 text-xs bg-white"></textarea>
                </div>

                <div class="flex justify-end gap-2 pt-3">
                    <button type="button" onclick="closeEditQuestionModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Batal</button>
                    <button type="submit" class="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-bold">Simpan Perubahan</button>
                </div>
            </form>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 9. MODAL PANDUAN INSTALASI PERANGKAT (DEFAULT HIDDEN!)                -->
    <!-- ===================================================================== -->
    <div id="installPlatformModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 shadow-2xl border border-slate-100">
            <div class="flex items-center justify-between mb-3">
                <h3 id="installPlatformTitle" class="text-lg font-extrabold text-slate-900">Instalasi Aplikasi</h3>
                <button onclick="closeInstallModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <div id="installPlatformContent" class="text-xs sm:text-sm text-slate-600 space-y-3 leading-relaxed"></div>
            <div class="mt-6 flex justify-between items-center gap-2">
                <button onclick="downloadStandaloneApp()" class="px-4 py-2.5 rounded-xl bg-slate-100 hover:bg-slate-200 text-slate-800 text-xs font-bold">
                    Unduh File index.html
                </button>
                <button onclick="closeInstallModal()" class="px-5 py-2.5 rounded-xl bg-blue-600 text-white text-xs font-bold">
                    Tutup
                </button>
            </div>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 10. MODAL HASIL SELESAI UJIAN SISWA                                   -->
    <!-- ===================================================================== -->
    <div id="examFinishedModal" class="fixed inset-0 bg-slate-900/75 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-md w-full p-6 sm:p-8 shadow-2xl text-center space-y-4">
            <div class="w-16 h-16 rounded-full bg-emerald-100 text-emerald-600 flex items-center justify-center mx-auto">
                <i data-lucide="check-circle-2" class="w-9 h-9"></i>
            </div>
            <h2 class="text-2xl font-extrabold text-slate-900">Ujian Telah Selesai!</h2>
            <p class="text-xs sm:text-sm text-slate-600 leading-relaxed">
                Terima kasih telah mengerjakan ujian dengan jujur. Seluruh jawaban Anda telah dikirim secara real-time ke server Administrator.
            </p>
            <div id="studentScoreSummaryBox" class="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-1">
                <span class="text-xs text-slate-500 block">Nilai Akhir Anda</span>
                <span id="studentFinalScoreText" class="text-3xl font-extrabold text-blue-600">0</span>
                <span id="studentFinalDetailText" class="text-xs text-slate-600 block">Benar 0 • Salah 0</span>
            </div>
            <button onclick="closeExamFinishedModal()" class="w-full py-3 px-6 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm transition">
                Kembali ke Dashboard Siswa
            </button>
        </div>
    </div>

    <!-- ===================================================================== -->
    <!-- 11. MODAL PEMBACA ARTIKEL BERITA PENDIDIKAN & LIGHTBOX GAMBAR         -->
    <!-- ===================================================================== -->
    <div id="newsReaderModal" class="fixed inset-0 bg-slate-900/65 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-2xl w-full p-6 sm:p-8 shadow-2xl border border-slate-200 max-h-[88vh] overflow-y-auto space-y-4">
            <div class="flex items-start justify-between gap-3 border-b border-slate-100 pb-3">
                <div>
                    <span id="newsModalCategory" class="px-2.5 py-1 rounded-full bg-indigo-100 text-indigo-700 text-[11px] font-bold">Pendidikan</span>
                    <h3 id="newsModalTitle" class="text-lg sm:text-xl font-extrabold text-slate-900 mt-2 leading-snug">Judul Berita</h3>
                    <p id="newsModalMeta" class="text-xs text-slate-400 mt-1">Tanggal • Penulis</p>
                </div>
                <button onclick="closeNewsReaderModal()" class="w-8 h-8 rounded-full bg-slate-100 hover:bg-slate-200 text-slate-600 flex items-center justify-center shrink-0">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <div id="newsModalBody" class="text-xs sm:text-sm text-slate-700 leading-relaxed space-y-3 whitespace-pre-line"></div>
            <div class="pt-3 border-t border-slate-100 flex justify-end">
                <button onclick="closeNewsReaderModal()" class="px-5 py-2.5 rounded-xl bg-slate-900 text-white text-xs font-bold">Tutup Artikel</button>
            </div>
        </div>
    </div>

    <div id="imageLightboxModal" onclick="closeImageLightbox()" class="fixed inset-0 bg-slate-950/85 backdrop-blur-sm z-[9995] items-center justify-center p-4 hidden">
        <div class="max-w-3xl w-full text-center space-y-3" onclick="event.stopPropagation()">
            <img id="lightboxImgTarget" src="" alt="Perbesar Gambar Soal" class="max-h-[80vh] mx-auto rounded-2xl border-2 border-white/20 shadow-2xl object-contain bg-white">
            <button onclick="closeImageLightbox()" class="px-5 py-2 rounded-xl bg-white text-slate-900 text-xs font-extrabold shadow">Tutup Pratinjau</button>
        </div>
    </div>

    <div id="gasCodeModal" class="fixed inset-0 bg-slate-900/65 backdrop-blur-sm z-50 items-center justify-center p-4 hidden">
        <div class="bg-white rounded-3xl max-w-2xl w-full p-6 shadow-2xl border border-slate-200 max-h-[88vh] overflow-y-auto space-y-4">
            <div class="flex items-center justify-between border-b border-slate-100 pb-3">
                <h3 class="text-base sm:text-lg font-extrabold text-slate-900">Kode Google Apps Script (Code.gs)</h3>
                <button onclick="closeGasCodeModal()" class="w-8 h-8 rounded-full bg-slate-100 text-slate-600 flex items-center justify-center">
                    <i data-lucide="x" class="w-4 h-4"></i>
                </button>
            </div>
            <p class="text-xs text-slate-600 leading-relaxed">
                Aplikasi ini sudah terhubung langsung dengan Web App Google Apps Script Anda. Jika ingin memastikan script di Google Spreadsheet memiliki fungsi penyimpanan kunci-nilai real-time (`doGet` &amp; `doPost`), gunakan kode berikut:
            </p>
            <pre id="gasCodeSnippet" class="bg-slate-900 text-emerald-300 p-4 rounded-2xl text-[11px] font-mono overflow-x-auto leading-relaxed">function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CBT_DB') || ss.insertSheet('CBT_DB');
  var rows = sheet.getDataRange().getValues();
  var data = {};
  for (var i = 1; i &lt; rows.length; i++) {
    if (rows[i][0]) {
      try { data[rows[i][0]] = JSON.parse(rows[i][1]); } catch (err) {}
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('CBT_DB') || ss.insertSheet('CBT_DB');
  if (sheet.getLastRow() === 0) sheet.appendRow(['KEY', 'VALUE_JSON', 'UPDATED_AT']);
  var body = JSON.parse(e.postData.contents);
  var rows = sheet.getDataRange().getValues();
  var foundRow = -1;
  for (var i = 1; i &lt; rows.length; i++) {
    if (rows[i][0] === body.key) { foundRow = i + 1; break; }
  }
  var jsonStr = JSON.stringify(body.value);
  var now = new Date().toISOString();
  if (foundRow &gt; 0) {
    sheet.getRange(foundRow, 2, 1, 2).setValues([[jsonStr, now]]);
  } else {
    sheet.appendRow([body.key, jsonStr, now]);
  }
  return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}</pre>
            <div class="flex justify-end gap-2">
                <button onclick="copyTextToClipboard(document.getElementById('gasCodeSnippet').innerText, 'Kode Code.gs berhasil disalin!')" class="px-4 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold">Salin Kode Code.gs</button>
                <button onclick="closeGasCodeModal()" class="px-4 py-2 rounded-xl bg-slate-100 text-slate-700 text-xs font-bold">Tutup</button>
            </div>
        </div>
    </div>

    <!-- Footer -->
    <footer id="mainFooter" class="bg-white border-t border-slate-200 py-6 mt-auto no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-500 text-center sm:text-left">
            <div>
                <strong id="footerSchoolName" class="text-slate-800">CBT-PRO | Sistem Ujian Berbasis Komputer</strong> © 2026 • <span class="text-indigo-700 font-bold">Development : Jovano F Kakalang, S.Ak.</span>
            </div>
            <div class="flex flex-wrap items-center justify-center gap-3">
                <span>Single-App Lockdown &amp; Sirine</span>
                <span>•</span>
                <span>GitHub Pages &amp; Cloud Sync</span>
            </div>
        </div>
    </footer>

    <!-- ===================================================================== -->
    <!-- 12. JAVASCRIPT ENGINE LENGKAP (CBT-PRO v3.0)                          -->
    <!-- ===================================================================== -->
    <script>
        // =================================================================
        // A. KONSTANTA DATABASE LOKAL & CLOUD APPS SCRIPT
        // =================================================================
        const STORAGE_KEYS = {
            SETTINGS: 'cbt_prod_settings_v3',
            PACKAGES: 'cbt_prod_packages_v3',
            STUDENTS: 'cbt_prod_students_v3',
            PROCTORS: 'cbt_prod_proctors_v3',
            LIVE_STATUS: 'cbt_prod_live_status_v3',
            LOGS: 'cbt_prod_activity_logs_v3',
            GALLERY: 'cbt_prod_gallery_v3',
            NEWS: 'cbt_prod_news_v3',
            ACTIVE_SESSION: 'cbt_prod_active_session_v3',
            UI_STATE: 'cbt_prod_ui_state_v3',
            ACTIVE_EXAM: 'cbt_prod_active_exam_v3'
        };

        const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxG14rQ28hWsNgQ8JzrGKmECZ-lgqfUJCgCmhr2BI3tYUduez8WprCw-MTiAXJzMoGztQ/exec";

        const DEFAULT_SETTINGS = {
            schoolName: 'CBT-PRO | Sistem Ujian Berbasis Komputer',
            schoolSubtitle: 'Portal Evaluasi Akademik Terpadu • Tahun Pelajaran 2026/2027',
            adminUsername: 'admin',
            adminPassword: 'admin123',
            examTitle: 'Ujian Akhir Semester (UAS) Berbasis Komputer',
            activeToken: 'CBT26X',
            unlockPin: '123',
            singleAppLock: true,
            requireFullscreen: true,
            detectTabSwitch: true,
            enableSiren: true,
            portalAnimatedBg: true,
            showScoreToStudent: true,
            // Pengaturan Dashboard Siswa
            stuBannerMsg: '',
            stuShowProfileTab: true,
            stuShowGuideTab: true,
            stuAllowFontResize: true,
            // Pengaturan Dashboard Pengawas
            procBannerMsg: '',
            procShowToken: true,
            procAllowQuestionInput: true,
            procAllowRemoteUnlock: true
        };

        const DEFAULT_NEWS = [
            {
                id: 'news_1',
                category: 'Pendidikan Nasional',
                title: 'Transformasi Asesmen Digital 2026: Evaluasi Berbasis Penalaran, Literasi & Numerasi',
                date: '23 November 2026',
                author: 'Redaksi Pendidikan CBT-PRO',
                excerpt: 'Evaluasi pembelajaran modern kini tidak hanya mengandalkan pilihan ganda, tetapi memadukan soal menjodohkan, isian singkat, dan uraian analitis.',
                content: 'Transformasi digital di dunia pendidikan Indonesia pada tahun pelajaran 2026/2027 menekankan pengukuran kompetensi literasi membaca, numerasi, dan penalaran kritis peserta didik secara menyeluruh.\n\nMelalui pemanfaatan sistem Computer-Based Test (CBT) modern, guru dan satuan pendidikan dapat menyusun variasi soal yang lebih kaya, mulai dari Pilihan Ganda, Menjodohkan konsep, Isian Singkat, hingga Uraian analitis yang dilengkapi stimulus gambar beresolusi tinggi.\n\nDengan dukungan sinkronisasi cloud real-time, proses koreksi dan pemantauan integritas ujian menjadi jauh lebih transparan, cepat, dan ramah lingkungan tanpa penggunaan kertas berlebih.'
            },
            {
                id: 'news_2',
                category: 'Tips Ujian & Literasi',
                title: 'Strategi Efektif Menjawab Soal Menjodohkan, Isian Singkat & Uraian di Sistem CBT',
                date: '21 November 2026',
                author: 'Tim Kurikulum & Evaluasi',
                excerpt: 'Simak langkah praktis mengatur manajemen waktu dan ketelitian saat mengerjakan variasi tipe soal dalam ujian berbasis komputer.',
                content: 'Menghadapi ujian berbasis komputer dengan beragam tipe soal membutuhkan ketenangan dan strategi manajemen waktu yang baik:\n\n1. Soal Pilihan Ganda: Baca pokok soal dengan saksama sebelum melihat opsi jawaban, lalu eliminasi pilihan yang kurang tepat.\n2. Soal Menjodohkan: Kerjakan pasangan pernyataan yang paling Anda yakini terlebih dahulu agar pilihan yang tersisa lebih mudah dicocokkan.\n3. Soal Isian Singkat: Perhatikan ejaan kata dan satuan angka yang diminta secara tepat dan ringkas.\n4. Soal Uraian: Susun jawaban secara runtut, memuat poin-poin inti pembahasan, serta gunakan bahasa Indonesia yang jelas.\n\nGunakan fitur Tandai Ragu-Ragu pada navigasi soal bila ingin meninjau ulang jawaban sebelum menekan tombol Selesai Ujian.'
            },
            {
                id: 'news_3',
                category: 'Pengumuman Sekolah',
                title: 'Kesiapan Infrastruktur & Tata Tertib Ruang Ujian Akhir Semester (UAS) 2026/2027',
                date: '20 November 2026',
                author: 'Panitia Ujian Sekolah',
                excerpt: 'Seluruh peserta ujian wajib memastikan kartu login aktif dan mematuhi aturan Single-App Lockdown selama berada di ruang ujian.',
                content: 'Diberitahukan kepada seluruh peserta didik dan pengawas ruangan bahwa sistem ujian CBT-PRO telah terhubung dengan server cloud sekolah.\n\nSetiap peserta wajib hadir 15 menit sebelum jadwal sesi dimulai dengan membawa Kartu Login resmi yang tercantum NISN dan Password.\n\nDemi menjunjung tinggi kejujuran akademik, aplikasi dilengkapi sistem Single-App Lockdown dan sirine otomatis yang akan aktif apabila perangkat terdeteksi keluar dari halaman ujian atau membuka aplikasi lain.'
            }
        ];

        const DEFAULT_PACKAGES = [
            {
                id: 'pkg_001',
                examTitle: 'Ujian Akhir Semester (UAS) 2026',
                subjectName: 'Kompetensi Akademik & Informatika',
                scheduleDate: '2026-11-23',
                scheduleTime: '08:00 - 09:30 WITA',
                durationMinutes: 45,
                targetGroup: 'ALL',
                status: 'ASSIGNED',
                shuffleQuestions: true,
                shuffleOptions: true,
                questions: [
                    {
                        id: 'q_init_1',
                        type: 'pg',
                        question: 'Komponen perangkat keras komputer yang berfungsi sebagai unit pemroses utama seluruh instruksi logika dan aritmatika adalah...',
                        image: '',
                        options: { A: 'RAM (Random Access Memory)', B: 'CPU (Central Processing Unit)', C: 'Solid State Drive (SSD)', D: 'Power Supply Unit', E: 'Network Interface Card' },
                        answer: 'B',
                        points: 25
                    },
                    {
                        id: 'q_init_2',
                        type: 'menjodohkan',
                        question: 'Jodohkan istilah jaringan komputer dan keamanan digital berikut dengan fungsinya yang tepat!',
                        image: '',
                        pairs: [
                            { left: 'HTTPS', right: 'Protokol web terenkripsi SSL/TLS' },
                            { left: 'Firewall', right: 'Penyaring lalu lintas jaringan masuk & keluar' },
                            { left: 'Cloud Sync', right: 'Sinkronisasi data secara daring real-time' }
                        ],
                        points: 25
                    },
                    {
                        id: 'q_init_3',
                        type: 'isian',
                        question: 'Teks Proklamasi Kemerdekaan Republik Indonesia dibacakan oleh Ir. Soekarno pada tanggal ... (Tuliskan tanggal, bulan, dan tahun lengkap)',
                        image: '',
                        shortAnswer: '17 Agustus 1945, 17-08-1945',
                        points: 25
                    },
                    {
                        id: 'q_init_4',
                        type: 'uraian',
                        question: 'Jelaskan mengapa sikap jujur dan mandiri sangat penting dijunjung tinggi oleh siswa saat mengikuti ujian berbasis komputer!',
                        image: '',
                        essayRubric: 'jujur, mandiri, integritas, kemampuan, percaya diri',
                        points: 25
                    }
                ]
            },
            {
                id: 'pkg_002',
                examTitle: 'Ujian Akhir Semester (UAS) 2026',
                subjectName: 'Akuntansi Keuangan & Perpajakan',
                scheduleDate: '2026-11-24',
                scheduleTime: '10:00 - 11:30 WITA',
                durationMinutes: 60,
                targetGroup: 'ALL',
                status: 'WAITING',
                shuffleQuestions: true,
                shuffleOptions: true,
                questions: [
                    {
                        id: 'q_akt_1',
                        type: 'pg',
                        question: 'Persamaan dasar akuntansi yang benar dalam pencatatan transaksi keuangan perusahaan adalah...',
                        image: '',
                        options: {
                            A: 'Aset = Liabilitas + Ekuitas',
                            B: 'Aset + Liabilitas = Ekuitas',
                            C: 'Liabilitas = Aset + Ekuitas',
                            D: 'Ekuitas = Aset + Pendapatan',
                            E: 'Aset = Pendapatan - Beban'
                        },
                        answer: 'A',
                        points: 50
                    },
                    {
                        id: 'q_akt_2',
                        type: 'isian',
                        question: 'Bukti transaksi yang digunakan untuk mencatat pembelian barang dagang secara kredit dari pemasok disebut...',
                        image: '',
                        shortAnswer: 'Faktur, Faktur Pembelian',
                        points: 50
                    }
                ]
            }
        ];

        const DEFAULT_STUDENTS = [
            { id: 'stu_001', name: 'Ahmad Fauzi', nisn: '0081234501', className: 'XII IPA 1', room: 'Ruang 01', username: '0081234501', password: '84921' },
            { id: 'stu_002', name: 'Siti Nurhaliza', nisn: '0081234502', className: 'XII IPA 1', room: 'Ruang 01', username: '0081234502', password: '37194' },
            { id: 'stu_003', name: 'Budi Santoso', nisn: '0081234503', className: 'XII IPS 2', room: 'Ruang 02', username: '0081234503', password: '62508' }
        ];

        const DEFAULT_PROCTORS = [
            { id: 'proc_1', name: 'Drs. Hendra Wijaya', room: 'Semua Ruangan', username: 'pengawas', password: 'pengawas123' },
            { id: 'proc_2', name: 'Rina Marlina, S.Pd', room: 'Ruang 01', username: 'pengawas_r1', password: '74125' }
        ];

        const DEFAULT_GALLERY = [
            { id: 'gal_1', title: 'Pelaksanaan Ujian Berbasis Komputer di Laboratorium 1', svgType: 'lab', url: '' },
            { id: 'gal_2', title: 'Pengawasan Ruang Ujian & Verifikasi Token Peserta', svgType: 'proctor', url: '' },
            { id: 'gal_3', title: 'Fokus & Kejujuran Akademik Siswa Saat Mengerjakan Soal', svgType: 'student', url: '' }
        ];

        const MOTIVATION_QUOTES = [
            {
                text: '"Prestasi sejati lahir dari kerja keras dan kejujuran. Nilai tinggi tanpa kejujuran hanyalah angka, sedangkan kejujuran adalah bekal kesuksesan masa depanmu."',
                author: '— Pesan Akademik Sekolah'
            },
            {
                text: '"Jangan takut pada soal yang sulit; percayalah pada proses belajar yang sudah kamu lalui setiap hari. Kerjakan dengan tenang dan teliti!"',
                author: '— Motivasi Belajar CBT-PRO'
            },
            {
                text: '"Pendidikan adalah paspor menuju masa depan, karena hari esok dimiliki oleh mereka yang mempersiapkannya hari ini."',
                author: '— Malcolm X'
            },
            {
                text: '"Ujian bukan sekadar mengukur seberapa banyak yang kamu hafal, tetapi melatih integritas, fokus, dan ketangguhan dirimu."',
                author: '— Dewan Guru & Pengawas'
            }
        ];

        let currentQuoteIdx = 0;
        let showStudentPasswords = false;
        let selectedAdminPkgId = 'pkg_001';
        let selectedProctorPkgId = 'pkg_001';
        let selectedNewsCategory = 'ALL';
        window.__allowPrint = false;

        const SYNC_KEYS = [
            STORAGE_KEYS.SETTINGS,
            STORAGE_KEYS.PACKAGES,
            STORAGE_KEYS.STUDENTS,
            STORAGE_KEYS.PROCTORS,
            STORAGE_KEYS.LIVE_STATUS,
            STORAGE_KEYS.LOGS,
            STORAGE_KEYS.GALLERY,
            STORAGE_KEYS.NEWS
        ];

        const getDB = (key, fallback) => {
            try {
                const raw = localStorage.getItem(key);
                return raw ? JSON.parse(raw) : fallback;
            } catch (e) {
                return fallback;
            }
        };

        const setDB = (key, value) => {
            localStorage.setItem(key, JSON.stringify(value));
            if (SYNC_KEYS.includes(key) && GAS_WEB_APP_URL) {
                fetch(GAS_WEB_APP_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({ key: key, value: value })
                }).catch(err => console.warn('Gagal sinkronisasi ke cloud:', err));
            }
        };

        let isSyncingCloud = false;
        async function syncFromCloud(forceRender = false) {
            if (!GAS_WEB_APP_URL || isSyncingCloud) return false;
            isSyncingCloud = true;
            try {
                const res = await fetch(GAS_WEB_APP_URL + "?t=" + Date.now());
                const json = await res.json();
                if (json && json.status === "success" && json.data) {
                    let hasChanges = false;
                    SYNC_KEYS.forEach(k => {
                        if (json.data[k] !== undefined) {
                            const remoteStr = JSON.stringify(json.data[k]);
                            const localStr = localStorage.getItem(k);
                            if (remoteStr !== localStr) {
                                localStorage.setItem(k, remoteStr);
                                hasChanges = true;
                            }
                        }
                    });
                    if (hasChanges || forceRender) {
                        if (examSession.isRunning && examSession.isLocked && activeSession?.user) {
                            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                            if (liveMap[activeSession.user.id]?.remoteUnlocked) {
                                unlockExamScreen();
                            }
                        }
                        syncAllViews();
                    }
                    return hasChanges;
                }
            } catch (err) {
                console.warn("Mode lokal aktif / sinkronisasi tertunda:", err);
            } finally {
                isSyncingCloud = false;
            }
            return false;
        }

        function persistSessionState(extraUi = {}) {
            if (activeSession) {
                localStorage.setItem(STORAGE_KEYS.ACTIVE_SESSION, JSON.stringify(activeSession));
                try { sessionStorage.setItem(STORAGE_KEYS.ACTIVE_SESSION, JSON.stringify(activeSession)); } catch (e) {}
            } else {
                localStorage.removeItem(STORAGE_KEYS.ACTIVE_SESSION);
                try { sessionStorage.removeItem(STORAGE_KEYS.ACTIVE_SESSION); } catch (e) {}
            }
            const prevUi = getDB(STORAGE_KEYS.UI_STATE, {});
            const nextUi = {
                ...prevUi,
                adminMenu: currentAdminMenu,
                proctorMenu: currentProctorMenu,
                studentDashTab: currentStudentDashTab,
                ...extraUi
            };
            localStorage.setItem(STORAGE_KEYS.UI_STATE, JSON.stringify(nextUi));
            try { sessionStorage.setItem(STORAGE_KEYS.UI_STATE, JSON.stringify(nextUi)); } catch (e) {}
        }

        function persistActiveExamState() {
            if (examSession && examSession.isRunning && activeSession?.role === 'siswa') {
                const payload = JSON.stringify({
                    studentId: activeSession.user.id,
                    packageId: examSession.packageId,
                    packageSubject: examSession.packageSubject,
                    packageTitle: examSession.packageTitle,
                    questions: examSession.questions,
                    answers: examSession.answers,
                    doubts: examSession.doubts,
                    currentIdx: examSession.currentIdx,
                    remainingSeconds: examSession.remainingSeconds,
                    violations: examSession.violations,
                    isLocked: examSession.isLocked,
                    savedAt: Date.now()
                });
                localStorage.setItem(STORAGE_KEYS.ACTIVE_EXAM, payload);
                try { sessionStorage.setItem(STORAGE_KEYS.ACTIVE_EXAM, payload); } catch (e) {}
            } else {
                localStorage.removeItem(STORAGE_KEYS.ACTIVE_EXAM);
                try { sessionStorage.removeItem(STORAGE_KEYS.ACTIVE_EXAM); } catch (e) {}
            }
        }

        function ensureLocalDatabaseDefaults() {
            if (!localStorage.getItem(STORAGE_KEYS.SETTINGS)) localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(DEFAULT_SETTINGS));
            if (!localStorage.getItem(STORAGE_KEYS.PACKAGES)) localStorage.setItem(STORAGE_KEYS.PACKAGES, JSON.stringify(DEFAULT_PACKAGES));
            if (!localStorage.getItem(STORAGE_KEYS.STUDENTS)) localStorage.setItem(STORAGE_KEYS.STUDENTS, JSON.stringify(DEFAULT_STUDENTS));
            if (!localStorage.getItem(STORAGE_KEYS.PROCTORS)) localStorage.setItem(STORAGE_KEYS.PROCTORS, JSON.stringify(DEFAULT_PROCTORS));
            if (!localStorage.getItem(STORAGE_KEYS.LIVE_STATUS)) localStorage.setItem(STORAGE_KEYS.LIVE_STATUS, JSON.stringify({}));
            if (!localStorage.getItem(STORAGE_KEYS.LOGS)) localStorage.setItem(STORAGE_KEYS.LOGS, JSON.stringify([]));
            if (!localStorage.getItem(STORAGE_KEYS.GALLERY)) localStorage.setItem(STORAGE_KEYS.GALLERY, JSON.stringify(DEFAULT_GALLERY));
            if (!localStorage.getItem(STORAGE_KEYS.NEWS)) localStorage.setItem(STORAGE_KEYS.NEWS, JSON.stringify(DEFAULT_NEWS));

            const pkgs = getDB(STORAGE_KEYS.PACKAGES, []);
            if (pkgs.length > 0) {
                if (!pkgs.some(p => p.id === selectedAdminPkgId)) selectedAdminPkgId = pkgs[0].id;
                if (!pkgs.some(p => p.id === selectedProctorPkgId)) selectedProctorPkgId = pkgs[0].id;
            }
        }

        // =================================================================
        // B. SISTEM LOG AKTIVITAS MASUK / KELUAR & PELANGGARAN REAL-TIME
        // =================================================================
        function recordSystemLog(role, userName, userIdentifier, roomOrClass, category, detail) {
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            logs.unshift({
                id: 'log_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
                timestamp: new Date().toLocaleString('id-ID'),
                role,
                userName,
                userIdentifier,
                roomOrClass,
                category,
                detail
            });
            if (logs.length > 300) logs.length = 300;
            setDB(STORAGE_KEYS.LOGS, logs);
        }

        // =================================================================
        // C. ENGINE SUARA SIRINE ALARM OTOMATIS (WEB AUDIO API)
        // =================================================================
        let sirenAudioCtx = null;
        let sirenOscillator = null;
        let sirenGain = null;
        let sirenModInterval = null;
        let isSirenPlaying = false;

        function startSirenAlarm() {
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            if (settings.enableSiren === false) return;
            if (isSirenPlaying) return;

            try {
                const AudioCtx = window.AudioContext || window.webkitAudioContext;
                if (!AudioCtx) return;
                sirenAudioCtx = new AudioCtx();
                if (sirenAudioCtx.state === 'suspended') {
                    sirenAudioCtx.resume();
                }

                sirenOscillator = sirenAudioCtx.createOscillator();
                sirenGain = sirenAudioCtx.createGain();

                sirenOscillator.type = 'sawtooth';
                sirenOscillator.frequency.setValueAtTime(680, sirenAudioCtx.currentTime);
                sirenGain.gain.setValueAtTime(0.32, sirenAudioCtx.currentTime);

                sirenOscillator.connect(sirenGain);
                sirenGain.connect(sirenAudioCtx.destination);
                sirenOscillator.start();
                isSirenPlaying = true;

                let high = false;
                sirenModInterval = setInterval(() => {
                    if (!sirenOscillator || !sirenAudioCtx) return;
                    const now = sirenAudioCtx.currentTime;
                    high = !high;
                    sirenOscillator.frequency.cancelScheduledValues(now);
                    sirenOscillator.frequency.linearRampToValueAtTime(high ? 1150 : 620, now + 0.35);
                }, 380);
            } catch (e) {
                console.warn('Web Audio API siren blocked:', e);
            }
        }

        function stopSirenAlarm() {
            isSirenPlaying = false;
            if (sirenModInterval) {
                clearInterval(sirenModInterval);
                sirenModInterval = null;
            }
            if (sirenOscillator) {
                try { sirenOscillator.stop(); } catch (e) {}
                try { sirenOscillator.disconnect(); } catch (e) {}
                sirenOscillator = null;
            }
            if (sirenAudioCtx) {
                try { sirenAudioCtx.close(); } catch (e) {}
                sirenAudioCtx = null;
            }
        }

        function testSirenSound() {
            startSirenAlarm();
            showToast('Menguji suara sirine peringatan selama 2 detik...', 'warning');
            setTimeout(() => {
                if (!examSession.isLocked) {
                    stopSirenAlarm();
                }
            }, 2000);
        }

        // =================================================================
        // D. SINGLE-APP LOCKDOWN & ANTI-DEVELOPER PROTECTION
        // =================================================================
        function initSingleAppSecurityLock() {
            // 1. Blokir klik kanan (Context Menu)
            document.addEventListener('contextmenu', function (e) {
                const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
                if (settings.singleAppLock !== false || examSession.isRunning) {
                    e.preventDefault();
                    if (examSession.isRunning) {
                        showToast('Klik kanan dinonaktifkan selama ujian berlangsung!', 'warning');
                    }
                }
            });

            // 2. Blokir Shortcut Pengembang (F12, Ctrl+Shift+I/J/C, Ctrl+U, Ctrl+S, Ctrl+P, dll.)
            document.addEventListener('keydown', function (e) {
                const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
                const lockEnabled = settings.singleAppLock !== false || examSession.isRunning;
                if (!lockEnabled) return;

                const key = (e.key || '').toLowerCase();
                const ctrlOrMeta = e.ctrlKey || e.metaKey;

                const isDevShortcut =
                    e.key === 'F12' ||
                    (ctrlOrMeta && e.shiftKey && ['i', 'j', 'c', 'k'].includes(key)) ||
                    (ctrlOrMeta && ['u', 's'].includes(key));

                if (isDevShortcut) {
                    e.preventDefault();
                    e.stopPropagation();
                    if (examSession.isRunning) {
                        triggerAntiCheatViolation('Peserta mencoba membuka Developer Tools / Inspect Element!');
                    } else {
                        showToast('Mode Kunci Satu Aplikasi aktif: Akses Developer Tools / Source dikunci.', 'warning');
                    }
                    return false;
                }

                // Blokir Ctrl+P kecuali saat menekan tombol Cetak Kartu Resmi
                if (ctrlOrMeta && key === 'p' && !window.__allowPrint) {
                    e.preventDefault();
                    showToast('Gunakan tombol Cetak Kartu Resmi di menu Akun Siswa.', 'warning');
                    return false;
                }

                // Blokir Copy/Paste/Cut/NewTab saat Ujian
                if (examSession.isRunning) {
                    if ((ctrlOrMeta && ['c', 'v', 'x', 'a', 'f', 't', 'n', 'w'].includes(key)) || e.key === 'PrintScreen') {
                        e.preventDefault();
                        if (e.key === 'PrintScreen') {
                            triggerAntiCheatViolation('Peserta mencoba mengambil tangkapan layar (Screenshot)!');
                        } else {
                            showToast('Fitur salin/tempel/tab baru dikunci selama ujian!', 'warning');
                        }
                        return false;
                    }
                }
            }, true);

            // 3. Kunci Riwayat Navigasi (Cegah tombol Back keluar aplikasi)
            window.history.pushState(null, '', window.location.href);
            window.addEventListener('popstate', function () {
                window.history.pushState(null, '', window.location.href);
                if (examSession.isRunning) {
                    triggerAntiCheatViolation('Peserta menekan tombol Kembali (Back) untuk keluar ujian!');
                } else {
                    showToast('Aplikasi berjalan dalam mode Single-App. Gunakan menu navigasi internal.', 'info');
                }
            });

            // 4. Deteksi DevTools Undocked / Split Screen Threshold saat Ujian
            setInterval(() => {
                if (!examSession.isRunning || examSession.isLocked) return;
                const widthDiff = window.outerWidth - window.innerWidth;
                const heightDiff = window.outerHeight - window.innerHeight;
                if (widthDiff > 220 || heightDiff > 220) {
                    triggerAntiCheatViolation('Terdeteksi pembukaan Developer Tools atau Split-Screen!');
                }
            }, 2500);
        }

        function toggleAppFullscreen() {
            if (!document.fullscreenElement) {
                document.documentElement.requestFullscreen().catch(() => {
                    showToast('Browser membatasi mode layar penuh otomatis.', 'info');
                });
            } else {
                if (!examSession.isRunning) {
                    document.exitFullscreen().catch(() => {});
                } else {
                    showToast('Tidak dapat keluar dari layar penuh saat ujian sedang berlangsung!', 'warning');
                }
            }
        }

        function triggerPrintLoginCards() {
            window.__allowPrint = true;
            window.print();
            setTimeout(() => { window.__allowPrint = false; }, 1500);
        }

        // =================================================================
        // E. STATE SESI PENGGUNA & MESIN UJIAN
        // =================================================================
        let activeSession = null;
        let currentAdminMenu = 'monitoring';
        let currentProctorMenu = 'kehadiran';
        let currentStudentDashTab = 'ujian';

        let examSession = {
            isRunning: false,
            isLocked: false,
            packageId: null,
            packageSubject: '',
            packageTitle: '',
            questions: [],
            answers: {},
            doubts: {},
            currentIdx: 0,
            remainingSeconds: 0,
            timerId: null,
            violations: 0
        };

        // =================================================================
        // F. NAVIGASI SPA & RENDER UI GLOBAL
        // =================================================================
        function navigateTo(view) {
            persistSessionState({ currentView: view });
            ['portal', 'siswa-dashboard', 'admin', 'pengawas', 'siswa'].forEach(v => {
                const el = document.getElementById(`view-${v}`);
                if (el) el.classList.add('hidden');
            });
            const target = document.getElementById(`view-${view}`);
            if (target) target.classList.remove('hidden');

            const header = document.getElementById('mainHeader');
            const footer = document.getElementById('mainFooter');
            const guestBtns = document.getElementById('headerGuestButtons');
            const authControls = document.getElementById('headerAuthControls');
            const backDashBtn = document.getElementById('headerBackDashBtn');

            if (view === 'siswa') {
                header.classList.add('hidden');
                footer.classList.add('hidden');
            } else {
                header.classList.remove('hidden');
                footer.classList.remove('hidden');
                if (activeSession) {
                    guestBtns.classList.add('hidden');
                    authControls.classList.remove('hidden');
                    authControls.classList.add('flex');

                    if (activeSession.role === 'admin') {
                        document.getElementById('headerUserBadge').innerText = 'Administrator';
                    } else if (activeSession.role === 'pengawas') {
                        document.getElementById('headerUserBadge').innerText = `Pengawas: ${activeSession.user.name}`;
                    } else {
                        document.getElementById('headerUserBadge').innerText = `Siswa: ${activeSession.user.name}`;
                    }

                    if (view === 'portal') {
                        backDashBtn.classList.remove('hidden');
                    } else {
                        backDashBtn.classList.add('hidden');
                    }
                } else {
                    guestBtns.classList.remove('hidden');
                    authControls.classList.add('hidden');
                    authControls.classList.remove('flex');
                }
            }

            syncAllViews();
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function handleHeaderBrandClick() {
            if (examSession.isRunning) return;
            navigateTo('portal');
        }

        function returnToRoleDashboard() {
            if (!activeSession) return navigateTo('portal');
            if (activeSession.role === 'admin') navigateTo('admin');
            else if (activeSession.role === 'pengawas') navigateTo('pengawas');
            else navigateTo('siswa-dashboard');
        }

        function syncAllViews() {
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const students = getDB(STORAGE_KEYS.STUDENTS, []);

            // Header & Footer Identitas Sekolah
            document.getElementById('headerSchoolTitle').innerText = settings.schoolName;
            document.getElementById('headerSchoolSubtitle').innerText = settings.schoolSubtitle;
            document.getElementById('footerSchoolName').innerText = settings.schoolName;

            // Kontrol Animasi Background Portal
            const heroEl = document.getElementById('portalHeroBanner');
            const orbsEl = document.getElementById('portalAnimatedOrbs');
            if (heroEl && orbsEl) {
                if (settings.portalAnimatedBg !== false) {
                    heroEl.classList.add('portal-animated-hero');
                    heroEl.classList.remove('bg-slate-900');
                    orbsEl.classList.remove('hidden');
                } else {
                    heroEl.classList.remove('portal-animated-hero');
                    heroEl.classList.add('bg-slate-900');
                    orbsEl.classList.add('hidden');
                }
            }

            // Portal Stats, News & Gallery
            const assignedCount = packages.filter(p => p.status === 'ASSIGNED').length;
            document.getElementById('heroExamTitle').innerText = settings.examTitle;
            document.getElementById('portalStatSoal').innerText = packages.length;
            document.getElementById('portalStatSiswa').innerText = students.length;
            document.getElementById('portalStatDurasi').innerText = assignedCount;
            renderPortalNews();
            renderPortalGallery();

            // Student Dashboard View (Terapkan Pengaturan Dashboard Siswa)
            if (activeSession?.role === 'siswa' && !examSession.isRunning) {
                applyStudentDashboardSettings(settings);
                renderStudentDashboard();
            }

            // Admin Views
            if (activeSession?.role === 'admin') {
                document.getElementById('sidebarTokenDisplay').innerText = settings.activeToken;
                document.getElementById('assignTokenDisplay').innerText = settings.activeToken;
                populateFilterDropdowns();
                renderAdminMonitoring();
                renderAdminPackageControls();
                renderAdminQuestionList();
                renderAdminAssignmentsTable();
                renderAdminLogsTable();
                renderAdminNilaiTable();
                renderAdminStudentAccountTable();
                renderAdminLoginCards();
                renderAdminSettingsForms();
            }

            // Pengawas Views (Terapkan Pengaturan Dashboard Pengawas)
            if (activeSession?.role === 'pengawas') {
                applyProctorDashboardSettings(settings);
                document.getElementById('proctorDashboardName').innerText = activeSession.user.name;
                document.getElementById('proctorAssignedRoomBadge').innerText = `Tugas: ${activeSession.user.room}`;
                document.getElementById('proctorTokenDisplay').innerText = settings.activeToken;
                document.getElementById('proctorCardTokenBig').innerText = settings.activeToken;
                document.getElementById('proctorPinDisplay').innerText = settings.unlockPin;
                document.getElementById('proctorPinDisplay2').innerText = settings.unlockPin;
                renderProctorMonitoring();
                renderProctorViolationsView();
                renderProctorSoalView();
            }

            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function applyStudentDashboardSettings(settings) {
            const banner = document.getElementById('stuCustomNoticeBanner');
            const bannerText = document.getElementById('stuCustomNoticeText');
            if (settings.stuBannerMsg && settings.stuBannerMsg.trim()) {
                banner.classList.remove('hidden');
                bannerText.innerText = settings.stuBannerMsg.trim();
            } else {
                banner.classList.add('hidden');
            }

            const profTabBtn = document.getElementById('stuTabBtn-profil');
            const sideProfLink = document.getElementById('stuSideProfileLink');
            if (settings.stuShowProfileTab === false) {
                profTabBtn.classList.add('hidden');
                if (sideProfLink) sideProfLink.classList.add('hidden');
                if (currentStudentDashTab === 'profil') switchStudentDashTab('ujian');
            } else {
                profTabBtn.classList.remove('hidden');
                if (sideProfLink) sideProfLink.classList.remove('hidden');
            }

            const guideTabBtn = document.getElementById('stuTabBtn-panduan');
            const sideGuideLink = document.getElementById('stuSideGuideLink');
            if (settings.stuShowGuideTab === false) {
                guideTabBtn.classList.add('hidden');
                if (sideGuideLink) sideGuideLink.classList.add('hidden');
                if (currentStudentDashTab === 'panduan') switchStudentDashTab('ujian');
            } else {
                guideTabBtn.classList.remove('hidden');
                if (sideGuideLink) sideGuideLink.classList.remove('hidden');
            }
        }

        function applyProctorDashboardSettings(settings) {
            const banner = document.getElementById('procCustomNoticeBanner');
            const bannerText = document.getElementById('procCustomNoticeText');
            if (settings.procBannerMsg && settings.procBannerMsg.trim()) {
                banner.classList.remove('hidden');
                bannerText.innerText = settings.procBannerMsg.trim();
            } else {
                banner.classList.add('hidden');
            }

            const hToken = document.getElementById('proctorHeaderTokenBox');
            const sToken = document.getElementById('proctorSideTokenCard');
            if (settings.procShowToken === false) {
                hToken.classList.add('hidden');
                sToken.classList.add('hidden');
            } else {
                hToken.classList.remove('hidden');
                sToken.classList.remove('hidden');
            }

            const soalBtn = document.getElementById('procMenuBtn-soal');
            if (settings.procAllowQuestionInput === false) {
                soalBtn.classList.add('hidden');
                if (currentProctorMenu === 'soal') switchProctorMenu('kehadiran');
            } else {
                soalBtn.classList.remove('hidden');
            }
        }

        // =================================================================
        // G. TOAST, KONFIRMASI & CLIPBOARD
        // =================================================================
        function showToast(message, type = 'info') {
            const container = document.getElementById('toastContainer');
            const el = document.createElement('div');
            const palette = {
                success: 'bg-emerald-600 text-white border-emerald-500',
                error: 'bg-red-600 text-white border-red-500',
                warning: 'bg-amber-400 text-slate-950 border-amber-300',
                info: 'bg-slate-900 text-white border-slate-700'
            };
            el.className = `pointer-events-auto flex items-center gap-2.5 px-4 py-3 rounded-2xl shadow-xl border text-xs sm:text-sm font-semibold transition-all duration-300 ${palette[type] || palette.info}`;
            el.innerHTML = `<span class="flex-1">${message}</span>`;
            container.appendChild(el);
            setTimeout(() => {
                el.classList.add('opacity-0', '-translate-y-2');
                setTimeout(() => el.remove(), 300);
            }, 3500);
        }

        function showConfirmDialog(title, message, onYes) {
            const modal = document.getElementById('confirmModal');
            document.getElementById('confirmTitle').innerText = title;
            document.getElementById('confirmMessage').innerText = message;
            modal.classList.remove('hidden');
            modal.classList.add('flex');

            const okBtn = document.getElementById('confirmOkBtn');
            const cancelBtn = document.getElementById('confirmCancelBtn');

            const close = () => {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                okBtn.onclick = null;
                cancelBtn.onclick = null;
            };
            cancelBtn.onclick = close;
            okBtn.onclick = () => {
                close();
                onYes();
            };
        }

        function copyTextToClipboard(text, successMessage = 'Berhasil disalin!') {
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text).then(() => {
                    showToast(successMessage, 'success');
                }).catch(() => fallbackCopyText(text, successMessage));
            } else {
                fallbackCopyText(text, successMessage);
            }
        }

        function fallbackCopyText(text, successMessage) {
            const ta = document.createElement('textarea');
            ta.value = text;
            document.body.appendChild(ta);
            ta.select();
            try {
                document.execCommand('copy');
                showToast(successMessage, 'success');
            } catch (e) {
                showToast('Gagal menyalin teks.', 'error');
            }
            document.body.removeChild(ta);
        }

        function copyPortalLink(role) {
            const currentCleanUrl = window.location.href.split('?')[0].split('#')[0];
            const fullLink = `${currentCleanUrl}?role=${role}`;
            const label = role === 'admin' ? 'Link Administrator' : role === 'pengawas' ? 'Link Portal Pengawas' : 'Link Portal Siswa';
            copyTextToClipboard(fullLink, `${label} disalin! Siap dibagikan.`);
        }

        // =================================================================
        // H. PORTAL UTAMA: BERITA PENDIDIKAN, MOTIVASI, GALERI & INSTALASI
        // =================================================================
        function filterPortalNews(cat) {
            selectedNewsCategory = cat;
            document.querySelectorAll('.news-cat-btn').forEach(btn => {
                if (btn.getAttribute('data-cat') === cat) {
                    btn.className = 'news-cat-btn px-3 py-1.5 rounded-xl bg-indigo-600 text-white text-xs font-bold transition';
                } else {
                    btn.className = 'news-cat-btn px-3 py-1.5 rounded-xl bg-white border border-slate-200 text-slate-600 hover:bg-slate-100 text-xs font-semibold transition';
                }
            });
            renderPortalNews();
        }

        function renderPortalNews() {
            const newsList = getDB(STORAGE_KEYS.NEWS, DEFAULT_NEWS);
            const grid = document.getElementById('portalNewsGrid');
            if (!grid) return;

            const filtered = selectedNewsCategory === 'ALL'
                ? newsList
                : newsList.filter(n => n.category === selectedNewsCategory);

            if (filtered.length === 0) {
                grid.innerHTML = `<div class="col-span-full bg-white rounded-2xl border border-slate-200 p-8 text-center text-xs text-slate-400">Belum ada berita pada kategori ini.</div>`;
                return;
            }

            grid.innerHTML = filtered.map(item => {
                const badgeColor = item.category === 'Pengumuman Sekolah'
                    ? 'bg-amber-100 text-amber-800'
                    : item.category === 'Tips Ujian & Literasi'
                        ? 'bg-emerald-100 text-emerald-700'
                        : 'bg-indigo-100 text-indigo-700';

                return `
                    <article class="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm hover:shadow-md transition flex flex-col justify-between space-y-4">
                        <div class="space-y-2.5">
                            <div class="flex items-center justify-between gap-2">
                                <span class="px-2.5 py-0.5 rounded-full text-[10px] font-extrabold uppercase ${badgeColor}">${item.category}</span>
                                <span class="text-[11px] text-slate-400">${item.date}</span>
                            </div>
                            <h3 class="font-extrabold text-slate-900 text-sm sm:text-base leading-snug line-clamp-2">${item.title}</h3>
                            <p class="text-xs text-slate-600 leading-relaxed">${item.excerpt}</p>
                        </div>
                        <div class="pt-3 border-t border-slate-100 flex items-center justify-between gap-2">
                            <span class="text-[11px] text-slate-400 truncate">${item.author}</span>
                            <button onclick="openNewsReaderModal('${item.id}')" class="text-xs font-bold text-indigo-600 hover:text-indigo-800 hover:underline shrink-0">
                                Baca Selengkapnya →
                            </button>
                        </div>
                    </article>
                `;
            }).join('');
        }

        function openNewsReaderModal(newsId) {
            const newsList = getDB(STORAGE_KEYS.NEWS, DEFAULT_NEWS);
            const item = newsList.find(n => n.id === newsId);
            if (!item) return;
            document.getElementById('newsModalCategory').innerText = item.category;
            document.getElementById('newsModalTitle').innerText = item.title;
            document.getElementById('newsModalMeta').innerText = `${item.date} • ${item.author}`;
            document.getElementById('newsModalBody').innerText = item.content;
            const modal = document.getElementById('newsReaderModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeNewsReaderModal() {
            const modal = document.getElementById('newsReaderModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function openImageLightbox(src) {
            if (!src) return;
            document.getElementById('lightboxImgTarget').src = src;
            const modal = document.getElementById('imageLightboxModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeImageLightbox() {
            const modal = document.getElementById('imageLightboxModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function openGasCodeModal() {
            const modal = document.getElementById('gasCodeModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeGasCodeModal() {
            const modal = document.getElementById('gasCodeModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function nextMotivationQuote() {
            currentQuoteIdx = (currentQuoteIdx + 1) % MOTIVATION_QUOTES.length;
            const q = MOTIVATION_QUOTES[currentQuoteIdx];
            document.getElementById('motivationQuoteText').innerText = q.text;
            document.getElementById('motivationQuoteAuthor').innerText = q.author;
        }

        function getGalleryIllustrationSVG(type) {
            if (type === 'proctor') {
                return `
                <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-amber-50 to-orange-100">
                    <rect x="30" y="30" width="340" height="160" rx="16" fill="#ffffff" stroke="#fde68a" stroke-width="2"/>
                    <circle cx="110" cy="95" r="28" fill="#f59e0b"/>
                    <path d="M65 165 C65 135, 155 135, 155 165" fill="#d97706"/>
                    <rect x="180" y="65" width="160" height="90" rx="10" fill="#1e293b"/>
                    <rect x="192" y="78" width="80" height="10" rx="4" fill="#fbbf24"/>
                    <rect x="192" y="96" width="130" height="8" rx="4" fill="#64748b"/>
                    <rect x="192" y="112" width="110" height="8" rx="4" fill="#38bdf8"/>
                    <circle cx="315" cy="132" r="12" fill="#10b981"/>
                </svg>`;
            }
            if (type === 'student') {
                return `
                <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-emerald-50 to-teal-100">
                    <rect x="30" y="30" width="340" height="160" rx="16" fill="#ffffff" stroke="#a7f3d0" stroke-width="2"/>
                    <rect x="85" y="65" width="230" height="105" rx="12" fill="#0f172a"/>
                    <rect x="98" y="78" width="204" height="75" rx="6" fill="#eff6ff"/>
                    <rect x="112" y="90" width="140" height="10" rx="4" fill="#2563eb"/>
                    <rect x="112" y="108" width="170" height="8" rx="4" fill="#94a3b8"/>
                    <rect x="112" y="124" width="120" height="8" rx="4" fill="#10b981"/>
                    <rect x="65" y="170" width="270" height="10" rx="5" fill="#64748b"/>
                </svg>`;
            }
            return `
            <svg viewBox="0 0 400 220" class="w-full h-44 object-cover bg-gradient-to-br from-blue-50 to-indigo-100">
                <rect x="25" y="25" width="350" height="170" rx="16" fill="#ffffff" stroke="#bfdbfe" stroke-width="2"/>
                <rect x="55" y="60" width="120" height="80" rx="8" fill="#1e3a8a"/>
                <rect x="63" y="68" width="104" height="60" rx="4" fill="#dbeafe"/>
                <rect x="95" y="140" width="40" height="15" fill="#64748b"/>
                <rect x="75" y="155" width="80" height="6" rx="3" fill="#475569"/>
                <rect x="225" y="60" width="120" height="80" rx="8" fill="#1e3a8a"/>
                <rect x="233" y="68" width="104" height="60" rx="4" fill="#dbeafe"/>
                <rect x="265" y="140" width="40" height="15" fill="#64748b"/>
                <rect x="245" y="155" width="80" height="6" rx="3" fill="#475569"/>
                <circle cx="200" cy="45" r="10" fill="#10b981"/>
            </svg>`;
        }

        function renderPortalGallery() {
            const gallery = getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
            const grid = document.getElementById('portalGalleryGrid');
            if (!grid) return;

            grid.innerHTML = gallery.map(item => {
                const mediaHtml = item.url
                    ? `<img src="${item.url}" alt="${item.title}" class="w-full h-44 object-cover bg-slate-100" onerror="this.outerHTML=getGalleryIllustrationSVG('lab')">`
                    : getGalleryIllustrationSVG(item.svgType || 'lab');

                return `
                    <div class="bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm hover:shadow-md transition flex flex-col">
                        ${mediaHtml}
                        <div class="p-4 flex-1 flex flex-col justify-between">
                            <p class="text-xs sm:text-sm font-bold text-slate-800 leading-snug">${item.title}</p>
                            <span class="text-[11px] text-slate-400 mt-2 inline-flex items-center gap-1">
                                <i data-lucide="camera" class="w-3 h-3"></i> Dokumentasi CBT Sekolah
                            </span>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function downloadStandaloneApp() {
            const htmlContent = '<!DOCTYPE html>\n' + document.documentElement.outerHTML;
            const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'index.html';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('File aplikasi index.html berhasil diunduh (siap di-upload ke GitHub Pages)!', 'success');
        }

        function openInstallModal(platform) {
            const modal = document.getElementById('installPlatformModal');
            const title = document.getElementById('installPlatformTitle');
            const content = document.getElementById('installPlatformContent');

            if (platform === 'android') {
                title.innerText = 'Pasang Aplikasi CBT di HP / Tablet Android';
                content.innerHTML = `
                    <p>Ikuti langkah berikut untuk memasang aplikasi mandiri (tanpa bilah browser) di Android:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-slate-700 font-medium">
                        <li>Buka alamat Portal CBT ini menggunakan browser <strong>Google Chrome</strong> di HP Android.</li>
                        <li>Ketuk ikon <strong>titik tiga (⋮)</strong> di sudut kanan atas Chrome.</li>
                        <li>Pilih menu <strong>"Tambahkan ke Layar utama"</strong> (<em>Add to Home screen</em>) atau <strong>"Instal aplikasi"</strong>.</li>
                        <li>Ketuk <strong>Tambahkan / Instal</strong>. Ikon aplikasi CBT akan muncul di layar utama HP Anda.</li>
                    </ol>
                `;
            } else if (platform === 'ios') {
                title.innerText = 'Pasang Aplikasi CBT di iPhone & iPad (iOS)';
                content.innerHTML = `
                    <p>Untuk perangkat iPhone atau iPad, gunakan browser bawaan <strong>Safari</strong>:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-slate-700 font-medium">
                        <li>Buka link Portal CBT ini di aplikasi <strong>Safari</strong>.</li>
                        <li>Ketuk ikon <strong>Bagikan (Share — kotak dengan panah ke atas)</strong> di bagian bawah layar.</li>
                        <li>Gulir ke bawah dan pilih <strong>"Tambahkan ke Layar Utama"</strong> (<em>Add to Home Screen</em>).</li>
                        <li>Ketuk <strong>Tambah (Add)</strong> di pojok kanan atas.</li>
                    </ol>
                `;
            } else {
                title.innerText = 'Pasang / Jalankan di Komputer & GitHub Pages';
                content.innerHTML = `
                    <p>Tersedia dua cara penggunaan pada Komputer / Laptop &amp; Hosting GitHub:</p>
                    <ol class="list-decimal pl-5 space-y-2 text-slate-700 font-medium">
                        <li><strong>Hosting GitHub Pages:</strong> Upload file <code>index.html</code> ini ke repository GitHub Anda, lalu aktifkan GitHub Pages pada menu <em>Settings → Pages</em>.</li>
                        <li><strong>Instal sebagai Aplikasi Desktop:</strong> Buka URL GitHub Pages di <strong>Google Chrome / Microsoft Edge</strong>, lalu klik ikon <strong>Install App</strong> di ujung kanan bilah alamat URL.</li>
                    </ol>
                `;
            }

            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeInstallModal() {
            const modal = document.getElementById('installPlatformModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        // =================================================================
        // I. OTENTIKASI LOGIN (SISWA KE DASHBOARD, PENGAWAS, ADMIN)
        // =================================================================
        function openLoginModal(role = 'siswa') {
            const modal = document.getElementById('loginModal');
            const roleField = document.getElementById('loginRoleField');
            const badge = document.getElementById('loginModalBadge');
            const title = document.getElementById('loginModalTitle');
            const desc = document.getElementById('loginModalDesc');
            const userLabel = document.getElementById('loginUsernameLabel');
            const submitBtn = document.getElementById('loginSubmitButton');

            roleField.value = role;
            document.getElementById('loginUsernameField').value = '';
            document.getElementById('loginPasswordField').value = '';

            if (role === 'siswa') {
                badge.className = 'px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-bold uppercase';
                badge.innerText = 'Peserta Ujian';
                title.innerText = 'Login Dashboard Siswa';
                desc.innerText = 'Masuk ke Dashboard Siswa untuk cek profil, jadwal ujian, dan input token saat siap.';
                userLabel.innerText = 'NISN / Username Kartu Login';
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-sm transition shadow-lg mt-2';
                submitBtn.innerText = 'Masuk ke Dashboard Siswa';
            } else if (role === 'pengawas') {
                badge.className = 'px-3 py-1 rounded-full bg-amber-100 text-amber-800 text-xs font-bold uppercase';
                badge.innerText = 'Pengawas Ruang';
                title.innerText = 'Login Portal Pengawas';
                desc.innerText = 'Pantau siswa keluar/masuk, token sinkron, sirine pelanggaran, dan input soal.';
                userLabel.innerText = 'Username Pengawas';
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-amber-500 hover:bg-amber-600 text-white font-bold text-sm transition shadow-lg mt-2';
                submitBtn.innerText = 'Masuk sebagai Pengawas';
            } else {
                badge.className = 'px-3 py-1 rounded-full bg-slate-900 text-white text-xs font-bold uppercase';
                badge.innerText = 'Administrator';
                title.innerText = 'Login Administrator';
                desc.innerText = 'Kelola bank soal per mapel, assignment jadwal, akun siswa, dan audit log.';
                userLabel.innerText = 'Username Admin';
                submitBtn.className = 'w-full py-3.5 px-6 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-sm transition shadow-lg mt-2';
                submitBtn.innerText = 'Masuk Panel Admin';
            }

            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeLoginModal() {
            const modal = document.getElementById('loginModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function processLogin(e) {
            e.preventDefault();
            const role = document.getElementById('loginRoleField').value;
            const userIn = document.getElementById('loginUsernameField').value.trim();
            const passIn = document.getElementById('loginPasswordField').value.trim();
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };

            if (role === 'admin') {
                if (userIn === settings.adminUsername && passIn === settings.adminPassword) {
                    activeSession = { role: 'admin', user: { name: 'Administrator', username: userIn } };
                    recordSystemLog('admin', 'Administrator', userIn, 'Pusat Kendali', 'LOGIN', 'Administrator berhasil masuk ke sistem');
                    closeLoginModal();
                    showToast('Selamat datang, Administrator!', 'success');
                    navigateTo('admin');
                } else {
                    showToast('Username atau Password Administrator salah!', 'error');
                }
                return;
            }

            if (role === 'pengawas') {
                const proctors = getDB(STORAGE_KEYS.PROCTORS, []);
                const found = proctors.find(p => p.username.toLowerCase() === userIn.toLowerCase() && p.password === passIn);
                if (found) {
                    activeSession = { role: 'pengawas', user: found };
                    recordSystemLog('pengawas', found.name, found.username, found.room, 'LOGIN', `Pengawas (${found.room}) masuk ke Dashboard Pengawas`);
                    closeLoginModal();
                    showToast(`Selamat bertugas, ${found.name}!`, 'success');
                    navigateTo('pengawas');
                } else {
                    showToast('Username atau Password Pengawas tidak valid!', 'error');
                }
                return;
            }

            if (role === 'siswa') {
                const students = getDB(STORAGE_KEYS.STUDENTS, []);
                const foundStu = students.find(s =>
                    (s.username.toLowerCase() === userIn.toLowerCase() || s.nisn.toLowerCase() === userIn.toLowerCase()) &&
                    s.password === passIn
                );
                if (!foundStu) {
                    showToast('NISN/Username atau Password tidak cocok dengan Kartu Login!', 'error');
                    return;
                }

                activeSession = { role: 'siswa', user: foundStu };

                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                const prev = liveMap[foundStu.id] || {};
                if (prev.status !== 'Sedang Ujian' && prev.status !== 'Terkunci' && prev.status !== 'Selesai') {
                    liveMap[foundStu.id] = {
                        ...prev,
                        status: 'Di Dashboard Siswa',
                        loginTime: new Date().toLocaleTimeString('id-ID'),
                        isLocked: false,
                        remoteUnlocked: false
                    };
                    setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                }

                recordSystemLog('siswa', foundStu.name, `NISN: ${foundStu.nisn}`, `${foundStu.className} • ${foundStu.room}`, 'LOGIN', 'Siswa masuk ke Dashboard Peserta');
                closeLoginModal();
                showToast(`Selamat datang, ${foundStu.name}!`, 'success');
                navigateTo('siswa-dashboard');
            }
        }

        function logoutSession() {
            if (examSession.isRunning) {
                showToast('Selesaikan ujian terlebih dahulu sebelum keluar!', 'warning');
                return;
            }
            if (activeSession) {
                const u = activeSession.user;
                const idLabel = activeSession.role === 'siswa' ? `NISN: ${u.nisn}` : u.username;
                const roomLabel = activeSession.role === 'siswa' ? `${u.className} • ${u.room}` : (u.room || 'Admin');
                recordSystemLog(activeSession.role, u.name, idLabel, roomLabel, 'LOGOUT', `${activeSession.role.toUpperCase()} keluar dari aplikasi`);

                if (activeSession.role === 'siswa') {
                    const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                    if (liveMap[u.id] && liveMap[u.id].status === 'Di Dashboard Siswa') {
                        liveMap[u.id].status = 'Keluar Dashboard';
                        setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                    }
                }
            }
            activeSession = null;
            persistSessionState({ currentView: 'portal' });
            persistActiveExamState();
            showToast('Anda telah keluar dari sesi.', 'info');
            navigateTo('portal');
        }

        // =================================================================
        // J. DASHBOARD SISWA (PROFIL, PENUGASAN MAPEL, PANDUAN & TATIB)
        // =================================================================
        function switchStudentDashTab(tab) {
            currentStudentDashTab = tab;
            persistSessionState({ studentDashTab: tab });
            ['ujian', 'profil', 'panduan'].forEach(t => {
                const btn = document.getElementById(`stuTabBtn-${t}`);
                const sec = document.getElementById(`stuDashTab-${t}`);
                if (!btn || !sec) return;
                if (t === tab) {
                    btn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-blue-600 text-white transition flex items-center gap-2';
                    sec.classList.remove('hidden');
                } else {
                    btn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2';
                    sec.classList.add('hidden');
                }
            });
        }

        function renderStudentDashboard() {
            if (!activeSession || activeSession.role !== 'siswa') return;
            const stu = activeSession.user;
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const stuLive = liveMap[stu.id] || {};
            const completedPkgs = stuLive.completedPackages || {};

            document.getElementById('stuDashWelcomeName').innerText = `Selamat Datang, ${stu.name}`;
            document.getElementById('stuDashWelcomeSub').innerText = `NISN: ${stu.nisn} • Kelas: ${stu.className} • Ruangan: ${stu.room}`;

            document.getElementById('stuSideName').innerText = stu.name;
            document.getElementById('stuSideNisn').innerText = stu.nisn;
            document.getElementById('stuSideClass').innerText = stu.className;
            document.getElementById('stuSideRoom').innerText = stu.room;

            document.getElementById('stuCardSchoolName').innerText = settings.schoolName;
            document.getElementById('stuProfName').innerText = stu.name;
            document.getElementById('stuProfNisn').innerText = stu.nisn;
            document.getElementById('stuProfClass').innerText = stu.className;
            document.getElementById('stuProfRoom').innerText = stu.room;
            document.getElementById('stuProfUser').innerText = stu.username;
            document.getElementById('stuProfLoginTime').innerText = `Aktif sejak ${stuLive.loginTime || 'Sesi Ini'}`;

            const relevantPackages = packages.filter(pkg => {
                if (!pkg.targetGroup || pkg.targetGroup === 'ALL') return true;
                if (pkg.targetGroup === `class:${stu.className}`) return true;
                if (pkg.targetGroup === `room:${stu.room}`) return true;
                return false;
            });

            document.getElementById('stuAssignedCountBadge').innerText = `${relevantPackages.length} Mata Pelajaran`;

            const listEl = document.getElementById('studentAssignmentsList');
            if (relevantPackages.length === 0) {
                listEl.innerHTML = `
                    <div class="bg-white rounded-2xl border border-slate-200 p-8 text-center text-xs sm:text-sm text-slate-500">
                        Belum ada mata pelajaran yang dijadwalkan untuk kelas/ruangan Anda saat ini.
                    </div>
                `;
            } else {
                listEl.innerHTML = relevantPackages.map(pkg => {
                    const isDone = !!completedPkgs[pkg.id];
                    const isAssigned = pkg.status === 'ASSIGNED';
                    const qCount = (pkg.questions || []).length;

                    let statusBadge = '';
                    let actionBtn = '';

                    if (isDone) {
                        const res = completedPkgs[pkg.id];
                        statusBadge = `<span class="px-3 py-1 rounded-full bg-emerald-100 text-emerald-700 text-xs font-extrabold">SUDAH DIKERJAKAN</span>`;
                        actionBtn = `
                            <div class="flex flex-wrap items-center gap-2">
                                ${settings.showScoreToStudent ? `<span class="px-3 py-2 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-extrabold">Nilai: ${res.score}</span>` : ''}
                                <span class="px-3.5 py-2 rounded-xl bg-slate-100 text-slate-500 text-xs font-bold">Selesai (${res.finishedAt || ''})</span>
                            </div>
                        `;
                    } else if (isAssigned) {
                        statusBadge = `<span class="px-3 py-1 rounded-full bg-blue-100 text-blue-700 text-xs font-extrabold flex items-center gap-1.5"><span class="w-2 h-2 rounded-full bg-blue-600 animate-ping"></span> SIAP DIKERJAKAN (ASSIGNED)</span>`;
                        actionBtn = `
                            <button onclick="openStudentTokenModal('${pkg.id}')" class="px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition shadow-sm flex items-center gap-2">
                                <i data-lucide="play-circle" class="w-4 h-4"></i> Mulai Ujian &amp; Input Token
                            </button>
                        `;
                    } else {
                        statusBadge = `<span class="px-3 py-1 rounded-full bg-amber-100 text-amber-800 text-xs font-extrabold">MENUNGGU ASSIGN / TERJADWAL</span>`;
                        actionBtn = `
                            <button disabled class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-400 text-xs font-bold cursor-not-allowed flex items-center gap-1.5">
                                <i data-lucide="clock" class="w-3.5 h-3.5"></i> Belum Diaktifkan Admin
                            </button>
                        `;
                    }

                    return `
                        <div class="bg-white rounded-2xl border ${isAssigned && !isDone ? 'border-blue-300 shadow-sm' : 'border-slate-200'} p-5 transition space-y-4">
                            <div class="flex flex-wrap items-start justify-between gap-2">
                                <div>
                                    <span class="text-[11px] font-bold uppercase tracking-wider text-slate-400">${pkg.examTitle}</span>
                                    <h4 class="text-base sm:text-lg font-extrabold text-slate-900 mt-0.5">${pkg.subjectName}</h4>
                                </div>
                                ${statusBadge}
                            </div>

                            <div class="grid grid-cols-2 sm:grid-cols-3 gap-2.5 text-xs bg-slate-50 p-3 rounded-xl border border-slate-100">
                                <div>
                                    <span class="text-slate-400 block text-[10px] uppercase font-bold">Jadwal Hari/Tanggal</span>
                                    <span class="font-bold text-slate-800">${pkg.scheduleDate || '-'} • ${pkg.scheduleTime || '-'}</span>
                                </div>
                                <div>
                                    <span class="text-slate-400 block text-[10px] uppercase font-bold">Durasi &amp; Jumlah Soal</span>
                                    <span class="font-bold text-slate-800">${pkg.durationMinutes} Menit • ${qCount} Butir Soal</span>
                                </div>
                                <div class="col-span-2 sm:col-span-1">
                                    <span class="text-slate-400 block text-[10px] uppercase font-bold">Target Peserta</span>
                                    <span class="font-bold text-indigo-700">${pkg.targetGroup === 'ALL' ? 'Semua Kelas' : pkg.targetGroup.replace('class:', 'Kelas ').replace('room:', 'Ruang ')}</span>
                                </div>
                            </div>

                            <div class="flex flex-wrap items-center justify-between gap-3 pt-1">
                                <span class="text-[11px] text-slate-500">
                                    ${isAssigned && !isDone ? 'Siapkan Token Ujian dari Pengawas sebelum menekan Mulai.' : isDone ? 'Hasil ujian telah tersimpan di server.' : 'Menunggu jadwal dibuka oleh Administrator.'}
                                </span>
                                ${actionBtn}
                            </div>
                        </div>
                    `;
                }).join('');
            }

            const historyEl = document.getElementById('stuMyHistoryList');
            const doneKeys = Object.keys(completedPkgs);
            if (doneKeys.length === 0) {
                historyEl.innerHTML = `<div class="p-4 rounded-xl bg-slate-50 border border-slate-200 text-slate-500 text-center">Belum ada riwayat ujian yang diselesaikan.</div>`;
            } else {
                historyEl.innerHTML = doneKeys.map(k => {
                    const item = completedPkgs[k];
                    return `
                        <div class="p-3.5 rounded-xl bg-emerald-50/50 border border-emerald-200 flex items-center justify-between gap-2">
                            <div>
                                <div class="font-bold text-slate-900">${item.subjectName || 'Mata Pelajaran'}</div>
                                <div class="text-[11px] text-slate-500">Selesai: ${item.finishedAt || '-'} • Benar ${item.correctCount} / Salah ${item.wrongCount}</div>
                            </div>
                            ${settings.showScoreToStudent ? `<span class="px-3 py-1 rounded-lg bg-emerald-600 text-white font-extrabold text-xs shrink-0">Nilai: ${item.score}</span>` : `<span class="px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-700 font-bold text-[11px] shrink-0">Terkirim</span>`}
                        </div>
                    `;
                }).join('');
            }

            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function openStudentTokenModal(pkgId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg) return;

            if (pkg.status !== 'ASSIGNED') {
                showToast('Mata pelajaran ini belum di-assign oleh Administrator.', 'warning');
                return;
            }
            if (!pkg.questions || pkg.questions.length === 0) {
                showToast('Paket soal ini belum memiliki butir soal.', 'error');
                return;
            }

            document.getElementById('tokenModalPackageId').value = pkg.id;
            document.getElementById('tokenModalExamTitle').innerText = `${pkg.subjectName} — ${pkg.examTitle}`;
            document.getElementById('tokenModalExamMeta').innerText = `Durasi Pengerjaan: ${pkg.durationMinutes} Menit • Total: ${pkg.questions.length} Soal`;
            document.getElementById('studentExamTokenInput').value = '';

            const modal = document.getElementById('studentTokenEntryModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeStudentTokenModal() {
            const modal = document.getElementById('studentTokenEntryModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function confirmStartExamWithToken(e) {
            e.preventDefault();
            const pkgId = document.getElementById('tokenModalPackageId').value;
            const tokenIn = document.getElementById('studentExamTokenInput').value.trim().toUpperCase();
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };

            if (tokenIn !== settings.activeToken.toUpperCase()) {
                showToast('Token Ujian tidak valid! Periksa kembali token dari Pengawas.', 'error');
                return;
            }

            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg) return;

            closeStudentTokenModal();
            startStudentExamSession(pkg);
        }

        // =================================================================
        // K. MESIN UJIAN SISWA (4 TIPE SOAL + FOTO + AUTO-GRADE + SIRINE)
        // =================================================================
        function shuffleArray(arr) {
            const copy = [...arr];
            for (let i = copy.length - 1; i > 0; i--) {
                const j = Math.floor(Math.random() * (i + 1));
                [copy[i], copy[j]] = [copy[j], copy[i]];
            }
            return copy;
        }

        function startStudentExamSession(pkg) {
            const stu = activeSession.user;
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };

            let preparedQuestions = JSON.parse(JSON.stringify(pkg.questions));
            if (pkg.shuffleQuestions) {
                preparedQuestions = shuffleArray(preparedQuestions);
            }
            preparedQuestions = preparedQuestions.map(q => {
                const qType = q.type || 'pg';
                if (qType === 'pg') {
                    const keys = ['A', 'B', 'C', 'D', 'E'].filter(k => q.options && q.options[k] && q.options[k].trim() !== '');
                    const displayOrder = pkg.shuffleOptions ? shuffleArray(keys) : keys;
                    return { ...q, type: 'pg', displayOrder };
                } else if (qType === 'menjodohkan') {
                    const rightChoices = shuffleArray((q.pairs || []).map(p => p.right));
                    return { ...q, type: 'menjodohkan', shuffledRight: rightChoices };
                }
                return { ...q, type: qType };
            });

            examSession = {
                isRunning: true,
                isLocked: false,
                packageId: pkg.id,
                packageSubject: pkg.subjectName,
                packageTitle: pkg.examTitle,
                questions: preparedQuestions,
                answers: {},
                doubts: {},
                currentIdx: 0,
                remainingSeconds: (pkg.durationMinutes || 30) * 60,
                timerId: null,
                violations: 0
            };

            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const prev = liveMap[stu.id] || {};
            liveMap[stu.id] = {
                ...prev,
                status: 'Sedang Ujian',
                activePackageId: pkg.id,
                activeSubject: pkg.subjectName,
                answeredCount: 0,
                totalQuestions: preparedQuestions.length,
                violations: 0,
                isLocked: false,
                remoteUnlocked: false,
                startTime: new Date().toLocaleTimeString('id-ID')
            };
            setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

            recordSystemLog('siswa', stu.name, `NISN: ${stu.nisn}`, `${stu.className} • ${stu.room}`, 'EXAM_START', `Mulai mengerjakan ujian: ${pkg.subjectName}`);

            document.getElementById('studentExamTitle').innerText = `${pkg.subjectName} — ${pkg.examTitle}`;
            document.getElementById('studentExamName').innerText = stu.name;
            document.getElementById('studentExamNisn').innerText = stu.nisn;
            document.getElementById('studentExamRoom').innerText = `${stu.className} (${stu.room})`;
            document.getElementById('examTotalNum').innerText = preparedQuestions.length;

            // Kontrol Resize Font sesuai pengaturan Admin
            const fontCtrl = document.getElementById('examFontResizeControls');
            if (fontCtrl) {
                if (settings.stuAllowFontResize === false) fontCtrl.classList.add('hidden');
                else fontCtrl.classList.remove('hidden');
            }

            navigateTo('siswa');
            renderExamQuestion();
            renderExamNavigationGrid();
            startExamCountdown();
            persistActiveExamState();

            if (settings.requireFullscreen && document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
        }

        function startExamCountdown() {
            if (examSession.timerId) clearInterval(examSession.timerId);
            updateTimerDisplay();
            examSession.timerId = setInterval(() => {
                if (!examSession.isRunning) return;
                examSession.remainingSeconds--;
                updateTimerDisplay();
                if (examSession.remainingSeconds % 5 === 0) persistActiveExamState();
                if (examSession.remainingSeconds <= 0) {
                    clearInterval(examSession.timerId);
                    showToast('Waktu ujian telah habis! Jawaban dikirim otomatis.', 'warning');
                    finalizeStudentExam();
                }
            }, 1000);
        }

        function updateTimerDisplay() {
            const mins = Math.floor(Math.max(0, examSession.remainingSeconds) / 60);
            const secs = Math.max(0, examSession.remainingSeconds) % 60;
            const el = document.getElementById('studentTimerDisplay');
            if (el) el.innerText = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
        }

        function isQuestionAnswered(q) {
            const ans = examSession.answers[q.id];
            const qType = q.type || 'pg';
            if (ans === undefined || ans === null) return false;
            if (qType === 'pg') return String(ans).trim() !== '';
            if (qType === 'isian' || qType === 'uraian') return String(ans).trim().length > 0;
            if (qType === 'menjodohkan') {
                if (typeof ans !== 'object') return false;
                const totalPairs = (q.pairs || []).length;
                if (totalPairs === 0) return false;
                let filled = 0;
                for (let i = 0; i < totalPairs; i++) {
                    if (ans[i] && String(ans[i]).trim() !== '') filled++;
                }
                return filled === totalPairs;
            }
            return false;
        }

        function renderExamQuestion() {
            const q = examSession.questions[examSession.currentIdx];
            if (!q) return;

            const qType = q.type || 'pg';
            const typeNames = {
                pg: 'Pilihan Ganda',
                menjodohkan: 'Menjodohkan',
                isian: 'Isian Singkat',
                uraian: 'Uraian / Esai'
            };

            document.getElementById('examCurrentNum').innerText = examSession.currentIdx + 1;
            document.getElementById('examQuestionTypeBadge').innerText = typeNames[qType] || 'Pilihan Ganda';
            document.getElementById('examQuestionBody').innerText = q.question;

            // Tampilkan Foto Soal jika ada
            const imgBox = document.getElementById('examQuestionImageBox');
            const imgEl = document.getElementById('examQuestionImage');
            if (q.image && q.image.trim() !== '') {
                imgEl.src = q.image;
                imgBox.classList.remove('hidden');
            } else {
                imgEl.src = '';
                imgBox.classList.add('hidden');
            }

            const doubtBtn = document.getElementById('doubtToggleBtn');
            if (examSession.doubts[q.id]) {
                doubtBtn.className = 'px-3.5 py-1.5 rounded-xl bg-amber-400 text-slate-950 text-xs font-extrabold transition flex items-center gap-1.5 shadow-sm';
            } else {
                doubtBtn.className = 'px-3.5 py-1.5 rounded-xl border border-amber-300 bg-amber-50 hover:bg-amber-100 text-amber-800 text-xs font-bold transition flex items-center gap-1.5';
            }

            const workspaceEl = document.getElementById('examAnswerWorkspace');
            const labels = ['A', 'B', 'C', 'D', 'E'];

            if (qType === 'pg') {
                const selectedKey = examSession.answers[q.id];
                workspaceEl.innerHTML = (q.displayOrder || ['A', 'B', 'C', 'D']).map((origKey, idx) => {
                    const isSelected = selectedKey === origKey;
                    const displayLetter = labels[idx] || origKey;
                    return `
                        <button type="button" onclick="selectExamAnswer('${q.id}', '${origKey}')" class="w-full text-left p-4 rounded-2xl border-2 transition flex items-start gap-3.5 ${isSelected ? 'border-blue-600 bg-blue-50/70 shadow-sm' : 'border-slate-200 hover:border-blue-300 bg-white'}">
                            <span class="w-8 h-8 rounded-xl flex items-center justify-center font-extrabold text-xs shrink-0 ${isSelected ? 'bg-blue-600 text-white' : 'bg-slate-100 text-slate-700'}">
                                ${displayLetter}
                            </span>
                            <span class="text-xs sm:text-sm font-medium text-slate-800 pt-1 break-anywhere">${q.options[origKey]}</span>
                        </button>
                    `;
                }).join('');
            } else if (qType === 'menjodohkan') {
                const currentPairsAns = (typeof examSession.answers[q.id] === 'object' && examSession.answers[q.id]) ? examSession.answers[q.id] : {};
                const choices = q.shuffledRight || (q.pairs || []).map(p => p.right);
                workspaceEl.innerHTML = `
                    <div class="space-y-3">
                        <div class="text-xs text-slate-500 font-semibold">Pasangkan setiap pernyataan di sebelah kiri dengan pasangan yang paling tepat di sebelah kanan:</div>
                        ${(q.pairs || []).map((pair, pIdx) => {
                            const chosen = currentPairsAns[pIdx] || '';
                            return `
                                <div class="p-3.5 rounded-2xl border ${chosen ? 'border-blue-300 bg-blue-50/30' : 'border-slate-200 bg-slate-50/50'} grid grid-cols-1 md:grid-cols-12 gap-3 items-center">
                                    <div class="md:col-span-6 flex items-start gap-2.5">
                                        <span class="w-6 h-6 rounded-lg bg-slate-900 text-white font-bold text-xs flex items-center justify-center shrink-0">${pIdx + 1}</span>
                                        <span class="text-xs sm:text-sm font-semibold text-slate-800 pt-0.5 break-anywhere">${pair.left}</span>
                                    </div>
                                    <div class="md:col-span-6">
                                        <select onchange="selectMatchingAnswer('${q.id}', ${pIdx}, this.value)" class="w-full px-3.5 py-2.5 rounded-xl border-2 ${chosen ? 'border-blue-500 bg-white text-blue-950 font-bold' : 'border-slate-300 bg-white text-slate-700'} text-xs sm:text-sm focus:outline-none">
                                            <option value="">-- Pilih Pasangan Jawaban --</option>
                                            ${choices.map(c => `<option value="${escapeHtmlAttr(c)}" ${chosen === c ? 'selected' : ''}>${c}</option>`).join('')}
                                        </select>
                                    </div>
                                </div>
                            `;
                        }).join('')}
                    </div>
                `;
            } else if (qType === 'isian') {
                const val = examSession.answers[q.id] || '';
                workspaceEl.innerHTML = `
                    <div class="p-5 rounded-2xl border-2 border-emerald-200 bg-emerald-50/30 space-y-2">
                        <label class="block text-xs font-extrabold uppercase tracking-wider text-emerald-800">Jawaban Isian Singkat:</label>
                        <input type="text" value="${escapeHtmlAttr(val)}" oninput="setTextExamAnswer('${q.id}', this.value)" placeholder="Ketik jawaban singkat Anda dengan jelas di sini..." class="w-full px-4 py-3.5 rounded-xl border-2 border-emerald-400 bg-white text-sm sm:text-base font-bold text-slate-900 focus:outline-none focus:ring-4 focus:ring-emerald-500/20">
                        <p class="text-[11px] text-slate-500">Periksa kembali ejaan tulisan Anda sebelum berpindah ke nomor berikutnya.</p>
                    </div>
                `;
            } else if (qType === 'uraian') {
                const val = examSession.answers[q.id] || '';
                workspaceEl.innerHTML = `
                    <div class="p-5 rounded-2xl border-2 border-indigo-200 bg-indigo-50/30 space-y-2">
                        <div class="flex items-center justify-between">
                            <label class="block text-xs font-extrabold uppercase tracking-wider text-indigo-800">Jawaban Uraian / Esai:</label>
                            <span id="essayCharCount" class="text-[11px] font-semibold text-slate-500">${val.length} karakter</span>
                        </div>
                        <textarea rows="6" oninput="setTextExamAnswer('${q.id}', this.value); const c = document.getElementById('essayCharCount'); if(c) c.innerText = this.value.length + ' karakter';" placeholder="Tuliskan penjelasan atau uraian jawaban Anda secara lengkap dan runtut..." class="w-full p-4 rounded-xl border-2 border-indigo-300 bg-white text-xs sm:text-sm leading-relaxed text-slate-900 focus:outline-none focus:ring-4 focus:ring-indigo-500/20">${val}</textarea>
                    </div>
                `;
            }

            document.getElementById('btnPrevQuestion').disabled = examSession.currentIdx === 0;
            const nextBtn = document.getElementById('btnNextQuestion');
            if (examSession.currentIdx === examSession.questions.length - 1) {
                nextBtn.innerHTML = `Selesai Ujian <i data-lucide="check-circle" class="w-4 h-4"></i>`;
                nextBtn.className = 'px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2';
                nextBtn.onclick = confirmSubmitExam;
            } else {
                nextBtn.innerHTML = `Selanjutnya <i data-lucide="chevron-right" class="w-4 h-4"></i>`;
                nextBtn.className = 'px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white text-xs sm:text-sm font-bold transition flex items-center gap-2';
                nextBtn.onclick = goNextQuestion;
            }

            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function escapeHtmlAttr(str) {
            return String(str || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        }

        function updateLiveAnsweredProgress() {
            const answeredTotal = examSession.questions.filter(q => isQuestionAnswered(q)).length;
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            if (activeSession?.user && liveMap[activeSession.user.id]) {
                liveMap[activeSession.user.id].answeredCount = answeredTotal;
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
            }
            persistActiveExamState();
        }

        function selectExamAnswer(questionId, origOptionKey) {
            examSession.answers[questionId] = origOptionKey;
            updateLiveAnsweredProgress();
            renderExamQuestion();
            renderExamNavigationGrid();
        }

        function selectMatchingAnswer(questionId, pairIndex, chosenRightValue) {
            if (!examSession.answers[questionId] || typeof examSession.answers[questionId] !== 'object') {
                examSession.answers[questionId] = {};
            }
            examSession.answers[questionId][pairIndex] = chosenRightValue;
            updateLiveAnsweredProgress();
            renderExamNavigationGrid();
        }

        function setTextExamAnswer(questionId, val) {
            examSession.answers[questionId] = val;
            updateLiveAnsweredProgress();
            renderExamNavigationGrid();
        }

        function toggleDoubtQuestion() {
            const q = examSession.questions[examSession.currentIdx];
            if (!q) return;
            examSession.doubts[q.id] = !examSession.doubts[q.id];
            persistActiveExamState();
            renderExamQuestion();
            renderExamNavigationGrid();
        }

        function goPrevQuestion() {
            if (examSession.currentIdx > 0) {
                examSession.currentIdx--;
                persistActiveExamState();
                renderExamQuestion();
                renderExamNavigationGrid();
            }
        }

        function goNextQuestion() {
            if (examSession.currentIdx < examSession.questions.length - 1) {
                examSession.currentIdx++;
                persistActiveExamState();
                renderExamQuestion();
                renderExamNavigationGrid();
            }
        }

        function jumpToQuestion(idx) {
            examSession.currentIdx = idx;
            persistActiveExamState();
            renderExamQuestion();
            renderExamNavigationGrid();
        }

        function renderExamNavigationGrid() {
            const grid = document.getElementById('examNumberGrid');
            grid.innerHTML = examSession.questions.map((q, idx) => {
                const isCurrent = idx === examSession.currentIdx;
                const isDoubt = !!examSession.doubts[q.id];
                const isAnswered = isQuestionAnswered(q);

                let style = 'bg-slate-100 text-slate-700 border-slate-300';
                if (isDoubt) style = 'bg-amber-400 text-slate-950 border-amber-500';
                else if (isAnswered) style = 'bg-blue-600 text-white border-blue-700';

                return `
                    <button type="button" onclick="jumpToQuestion(${idx})" class="h-10 rounded-xl border font-bold text-xs transition flex items-center justify-center ${style} ${isCurrent ? 'ring-2 ring-slate-900 scale-105' : ''}">
                        ${idx + 1}
                    </button>
                `;
            }).join('');
        }

        function changeExamFontSize(size) {
            const body = document.getElementById('examQuestionBody');
            body.classList.remove('text-sm', 'text-base', 'text-lg', 'text-xl');
            if (size === 'sm') body.classList.add('text-sm');
            else if (size === 'lg') body.classList.add('text-lg');
            else body.classList.add('text-base');
        }

        function confirmSubmitExam() {
            const answeredCount = examSession.questions.filter(q => isQuestionAnswered(q)).length;
            const total = examSession.questions.length;
            const unanswered = total - answeredCount;
            const msg = unanswered > 0
                ? `Masih ada ${unanswered} soal yang belum terjawab lengkap. Apakah Anda yakin ingin mengakhiri dan mengirim ujian sekarang?`
                : 'Semua soal telah terjawab. Apakah Anda yakin ingin mengakhiri dan mengirim jawaban sekarang?';

            showConfirmDialog('Kirim & Selesaikan Ujian?', msg, () => {
                finalizeStudentExam();
            });
        }

        function calculateQuestionScore(q, ans) {
            const maxPts = Number(q.points || 10);
            const qType = q.type || 'pg';

            if (qType === 'pg') {
                return (ans && ans === q.answer) ? maxPts : 0;
            }
            if (qType === 'menjodohkan') {
                const pairs = q.pairs || [];
                if (pairs.length === 0 || !ans || typeof ans !== 'object') return 0;
                let correctPairs = 0;
                pairs.forEach((p, idx) => {
                    if (ans[idx] && ans[idx].trim().toLowerCase() === p.right.trim().toLowerCase()) {
                        correctPairs++;
                    }
                });
                return (correctPairs / pairs.length) * maxPts;
            }
            if (qType === 'isian') {
                if (!ans || String(ans).trim() === '') return 0;
                const cleanStudent = String(ans).trim().toLowerCase();
                const acceptedList = String(q.shortAnswer || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                return acceptedList.includes(cleanStudent) ? maxPts : 0;
            }
            if (qType === 'uraian') {
                const cleanStudent = String(ans || '').trim().toLowerCase();
                if (cleanStudent.length < 5) return 0;
                const keywords = String(q.essayRubric || '').split(',').map(k => k.trim().toLowerCase()).filter(Boolean);
                if (keywords.length === 0) {
                    return cleanStudent.length >= 15 ? maxPts : maxPts * 0.5;
                }
                let matched = 0;
                keywords.forEach(kw => {
                    if (cleanStudent.includes(kw)) matched++;
                });
                const ratio = Math.min(1, matched / Math.max(1, Math.ceil(keywords.length * 0.6)));
                return Math.max(cleanStudent.length >= 25 ? maxPts * 0.4 : 0, ratio * maxPts);
            }
            return 0;
        }

        function finalizeStudentExam() {
            if (!examSession.isRunning) return;
            clearInterval(examSession.timerId);
            examSession.isRunning = false;
            examSession.isLocked = false;
            persistActiveExamState();
            stopSirenAlarm();
            hideAntiCheatLockOverlay();

            let earnedPoints = 0;
            let totalPossiblePoints = 0;
            let correctCount = 0;
            let wrongCount = 0;

            examSession.questions.forEach(q => {
                const pts = Number(q.points || 10);
                totalPossiblePoints += pts;
                const got = calculateQuestionScore(q, examSession.answers[q.id]);
                earnedPoints += got;
                if (got >= pts * 0.6) correctCount++;
                else wrongCount++;
            });

            const finalScore = totalPossiblePoints > 0 ? Math.round((earnedPoints / totalPossiblePoints) * 100) : 0;
            const stu = activeSession.user;
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const prev = liveMap[stu.id] || {};
            const completedPackages = prev.completedPackages || {};

            completedPackages[examSession.packageId] = {
                packageId: examSession.packageId,
                subjectName: examSession.packageSubject,
                examTitle: examSession.packageTitle,
                score: finalScore,
                correctCount,
                wrongCount,
                violations: examSession.violations,
                finishedAt: new Date().toLocaleTimeString('id-ID')
            };

            liveMap[stu.id] = {
                ...prev,
                status: 'Selesai',
                activeSubject: examSession.packageSubject,
                answeredCount: examSession.questions.length,
                totalQuestions: examSession.questions.length,
                correctCount,
                wrongCount,
                score: finalScore,
                violations: examSession.violations,
                isLocked: false,
                remoteUnlocked: false,
                finishedAt: new Date().toLocaleTimeString('id-ID'),
                completedPackages
            };
            setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);

            recordSystemLog('siswa', stu.name, `NISN: ${stu.nisn}`, `${stu.className} • ${stu.room}`, 'EXAM_FINISH', `Selesai ujian ${examSession.packageSubject} (Nilai: ${finalScore})`);

            if (document.fullscreenElement) {
                document.exitFullscreen().catch(() => {});
            }

            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            const summaryBox = document.getElementById('studentScoreSummaryBox');
            if (settings.showScoreToStudent) {
                summaryBox.classList.remove('hidden');
                document.getElementById('studentFinalScoreText').innerText = finalScore;
                document.getElementById('studentFinalDetailText').innerText = `Tepat ${correctCount} • Kurang Tepat ${wrongCount} • Pelanggaran ${examSession.violations}x`;
            } else {
                summaryBox.classList.add('hidden');
            }

            const modal = document.getElementById('examFinishedModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeExamFinishedModal() {
            const modal = document.getElementById('examFinishedModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
            navigateTo('siswa-dashboard');
        }

        // =================================================================
        // L. PROTEKSI ANTI-CURANG + SIRINE ALARM + BUKA KUNCI PIN / JARAK JAUH
        // =================================================================
        function triggerAntiCheatViolation(reason) {
            if (!examSession.isRunning || examSession.isLocked) return;
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            if (!settings.detectTabSwitch) return;

            examSession.isLocked = true;
            examSession.violations += 1;

            const stu = activeSession.user;
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            liveMap[stu.id] = {
                ...(liveMap[stu.id] || {}),
                status: 'Terkunci',
                isLocked: true,
                remoteUnlocked: false,
                violations: examSession.violations,
                lastViolationReason: reason,
                lastViolationTime: new Date().toLocaleTimeString('id-ID')
            };
            setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
            persistActiveExamState();

            recordSystemLog('siswa', stu.name, `NISN: ${stu.nisn}`, `${stu.className} • ${stu.room}`, 'VIOLATION', `SIRINE & KUNCI LAYAR: ${reason} (Pelanggaran ke-${examSession.violations})`);

            startSirenAlarm();

            document.getElementById('lockReasonText').innerText = reason;
            document.getElementById('lockStudentName').innerText = stu.name;
            document.getElementById('lockStudentMeta').innerText = `${stu.nisn} / ${stu.className} (${stu.room})`;
            document.getElementById('lockExamSubject').innerText = examSession.packageSubject || '-';
            document.getElementById('lockViolationBadge').innerText = `${examSession.violations} Kali`;
            document.getElementById('unlockPinField').value = '';

            const overlay = document.getElementById('antiCheatLockOverlay');
            overlay.classList.remove('hidden');
            overlay.classList.add('flex');
        }

        function hideAntiCheatLockOverlay() {
            const overlay = document.getElementById('antiCheatLockOverlay');
            overlay.classList.add('hidden');
            overlay.classList.remove('flex');
        }

        function unlockExamScreen() {
            examSession.isLocked = false;
            stopSirenAlarm();
            hideAntiCheatLockOverlay();
            if (activeSession?.user) {
                const stu = activeSession.user;
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                if (liveMap[stu.id]) {
                    liveMap[stu.id].status = 'Sedang Ujian';
                    liveMap[stu.id].isLocked = false;
                    liveMap[stu.id].remoteUnlocked = false;
                    setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                }
                persistActiveExamState();
                recordSystemLog('siswa', stu.name, `NISN: ${stu.nisn}`, `${stu.className} • ${stu.room}`, 'UNLOCK', 'Sirine dimatikan & layar ujian dibuka kembali oleh Pengawas/Admin');
            }
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            if (settings.requireFullscreen && document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen().catch(() => {});
            }
            showToast('Sirine dimatikan & layar ujian dibuka kembali!', 'success');
        }

        function verifyPinUnlock() {
            const pinIn = document.getElementById('unlockPinField').value.trim();
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            if (pinIn === settings.unlockPin) {
                unlockExamScreen();
            } else {
                showToast('Kode PIN Pengawas salah! Sirine tetap aktif.', 'error');
            }
        }

        document.addEventListener('visibilitychange', () => {
            if (document.hidden && examSession.isRunning && !examSession.isLocked) {
                triggerAntiCheatViolation('Siswa meninggalkan tab atau meminimalkan aplikasi ujian!');
            }
        });

        window.addEventListener('blur', () => {
            if (examSession.isRunning && !examSession.isLocked) {
                triggerAntiCheatViolation('Layar ujian kehilangan fokus (membuka jendela/aplikasi lain)!');
            }
        });

        // =================================================================
        // M. FITUR MULTI-TIPE SOAL, FOTO JPG/PNG & PARSER WORD (.DOCX) / PDF
        // =================================================================
        function toggleQuestionTypeFields(prefix) {
            const typeEl = document.getElementById(`${prefix}QType`);
            if (!typeEl) return;
            const qType = typeEl.value;
            ['pg', 'menjodohkan', 'isian', 'uraian'].forEach(t => {
                const box = document.getElementById(`${prefix}Fields-${t}`);
                if (box) {
                    if (t === qType) box.classList.remove('hidden');
                    else box.classList.add('hidden');
                }
            });
            if (qType === 'menjodohkan') {
                const rowsContainer = document.getElementById(`${prefix}MatchingRows`);
                if (rowsContainer && rowsContainer.children.length === 0) {
                    addMatchingPairRow(`${prefix}MatchingRows`);
                    addMatchingPairRow(`${prefix}MatchingRows`);
                    addMatchingPairRow(`${prefix}MatchingRows`);
                }
            }
        }

        function addMatchingPairRow(containerId, leftVal = '', rightVal = '') {
            const container = document.getElementById(containerId);
            if (!container) return;
            const row = document.createElement('div');
            row.className = 'flex items-center gap-1.5 matching-pair-row';
            row.innerHTML = `
                <input type="text" placeholder="Pernyataan Kiri" value="${escapeHtmlAttr(leftVal)}" class="pair-left flex-1 min-w-0 px-2.5 py-1.5 rounded-lg border border-amber-300 text-xs bg-white">
                <span class="text-xs font-bold text-amber-700">=</span>
                <input type="text" placeholder="Pasangan Kanan" value="${escapeHtmlAttr(rightVal)}" class="pair-right flex-1 min-w-0 px-2.5 py-1.5 rounded-lg border border-amber-300 text-xs bg-white">
                <button type="button" onclick="this.parentElement.remove()" class="p-1.5 rounded-lg bg-red-100 text-red-600 hover:bg-red-200 text-xs font-bold shrink-0" title="Hapus baris">×</button>
            `;
            container.appendChild(row);
        }

        function handleQuestionImageSelect(event, prefix) {
            const file = event.target.files[0];
            if (!file) return;
            if (!['image/jpeg', 'image/png', 'image/jpg'].includes(file.type)) {
                showToast('Format foto harus JPG atau PNG!', 'error');
                event.target.value = '';
                return;
            }
            const reader = new FileReader();
            reader.onload = function (e) {
                const img = new Image();
                img.onload = function () {
                    // Kompres ringan agar tajam namun hemat penyimpanan & cloud sync
                    const maxW = 880;
                    let w = img.width;
                    let h = img.height;
                    if (w > maxW) {
                        h = Math.round((h * maxW) / w);
                        w = maxW;
                    }
                    const canvas = document.createElement('canvas');
                    canvas.width = w;
                    canvas.height = h;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, w, h);
                    const outputMime = file.type === 'image/png' ? 'image/png' : 'image/jpeg';
                    const compressedDataUrl = canvas.toDataURL(outputMime, 0.82);

                    document.getElementById(`${prefix}QImageData`).value = compressedDataUrl;
                    document.getElementById(`${prefix}QImagePreview`).src = compressedDataUrl;
                    document.getElementById(`${prefix}QImagePreviewBox`).classList.remove('hidden');
                    document.getElementById(`${prefix}QImgClearBtn`).classList.remove('hidden');
                    showToast('Foto soal JPG/PNG berhasil dilampirkan!', 'success');
                };
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        }

        function clearQuestionImageInput(prefix) {
            const fileIn = document.getElementById(`${prefix}QImageFile`);
            if (fileIn) fileIn.value = '';
            const dataIn = document.getElementById(`${prefix}QImageData`);
            if (dataIn) dataIn.value = '';
            const prevBox = document.getElementById(`${prefix}QImagePreviewBox`);
            if (prevBox) prevBox.classList.add('hidden');
            const clearBtn = document.getElementById(`${prefix}QImgClearBtn`);
            if (clearBtn) clearBtn.classList.add('hidden');
        }

        function downloadSampleQuestionFormat() {
            const sampleText =
`1. [PG] Komponen perangkat keras komputer yang berfungsi sebagai unit pemroses utama adalah...
A. RAM (Random Access Memory)
B. CPU (Central Processing Unit)
C. Solid State Drive (SSD)
D. Power Supply Unit
E. Network Interface Card
Kunci: B

2. [MENJODOHKAN] Jodohkan istilah teknologi jaringan berikut dengan fungsinya yang tepat!
1) HTTPS = Protokol web terenkripsi SSL/TLS
2) Firewall = Penyaring lalu lintas jaringan masuk dan keluar
3) Cloud Sync = Sinkronisasi data secara daring real-time

3. [ISIAN] Tanggal pembacaan teks Proklamasi Kemerdekaan Indonesia adalah...
Kunci: 17 Agustus 1945, 17-08-1945

4. [URAIAN] Jelaskan mengapa sikap jujur sangat penting diterapkan dalam pelaksanaan ujian berbasis komputer!
Kunci: kejujuran, integritas, mandiri, percaya diri
`;
            const blob = new Blob([sampleText], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'Contoh_Format_Soal_Lengkap_CBT.txt';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('Contoh format soal (PG, Menjodohkan, Isian, Uraian) berhasil diunduh!', 'info');
        }

        async function handleDocumentQuestionUpload(event, sourceRole) {
            const file = event.target.files[0];
            if (!file) return;

            const pkgId = sourceRole === 'admin' ? selectedAdminPkgId : selectedProctorPkgId;
            const statusBox = document.getElementById('uploadParseStatus');
            if (statusBox && sourceRole === 'admin') {
                statusBox.classList.remove('hidden');
                statusBox.innerText = `Memproses file "${file.name}"...`;
            }

            try {
                const fileName = file.name.toLowerCase();
                let extractedText = '';
                let extractedImagesByQuestion = {};

                if (fileName.endsWith('.docx')) {
                    if (typeof mammoth === 'undefined') {
                        throw new Error('Library pembaca Word (.docx) belum siap.');
                    }
                    const arrayBuffer = await file.arrayBuffer();
                    // Ekstrak HTML agar gambar JPG/PNG di dalam Word ikut terbaca
                    const htmlRes = await mammoth.convertToHtml({ arrayBuffer });
                    const rawRes = await mammoth.extractRawText({ arrayBuffer });
                    extractedText = rawRes.value || '';

                    // Petakan gambar tertanam ke nomor soal dari HTML hasil Mammoth
                    try {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(htmlRes.value || '', 'text/html');
                        let currentQNum = 0;
                        doc.body.querySelectorAll('p, div, table, img').forEach(node => {
                            const txt = (node.textContent || '').trim();
                            const numMatch = txt.match(/^(\d+)[\.\)]\s+/);
                            if (numMatch) {
                                currentQNum = parseInt(numMatch[1], 10);
                            }
                            if (node.tagName === 'IMG') {
                                const src = node.getAttribute('src') || '';
                                if (src.startsWith('data:image/') && currentQNum > 0 && !extractedImagesByQuestion[currentQNum]) {
                                    extractedImagesByQuestion[currentQNum] = src;
                                }
                            }
                        });
                    } catch (imgErr) {
                        console.warn('Ekstraksi gambar docx opsional:', imgErr);
                    }
                } else if (fileName.endsWith('.pdf')) {
                    if (typeof pdfjsLib === 'undefined') {
                        throw new Error('Library pembaca PDF belum siap.');
                    }
                    const arrayBuffer = await file.arrayBuffer();
                    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
                    let fullText = '';
                    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                        const page = await pdf.getPage(pageNum);
                        const content = await page.getTextContent();
                        const strings = content.items.map(item => item.str);
                        fullText += strings.join('\n') + '\n';
                    }
                    extractedText = fullText;
                } else {
                    extractedText = await file.text();
                }

                const parsedQuestions = parseQuestionsFromText(extractedText, extractedImagesByQuestion);
                if (parsedQuestions.length === 0) {
                    throw new Error('Tidak ditemukan pola soal yang sesuai. Unduh "Contoh Format" untuk melihat struktur penulisan.');
                }

                const packages = getDB(STORAGE_KEYS.PACKAGES, []);
                const pkgIdx = packages.findIndex(p => p.id === pkgId);
                if (pkgIdx !== -1) {
                    packages[pkgIdx].questions.push(...parsedQuestions);
                    setDB(STORAGE_KEYS.PACKAGES, packages);
                }

                if (statusBox && sourceRole === 'admin') {
                    statusBox.innerText = `Berhasil menambahkan ${parsedQuestions.length} soal dari "${file.name}"!`;
                }
                showToast(`Berhasil mengimpor ${parsedQuestions.length} soal dari ${file.name}!`, 'success');
                syncAllViews();
            } catch (err) {
                if (statusBox && sourceRole === 'admin') {
                    statusBox.innerText = `Gagal membaca file: ${err.message}`;
                }
                showToast(err.message || 'Gagal memproses dokumen soal.', 'error');
            } finally {
                event.target.value = '';
            }
        }

        function parseQuestionsFromText(rawText, embeddedImagesMap = {}) {
            const lines = rawText.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
            const results = [];
            let current = null;
            let currentQNum = 0;

            const finalizeCurrent = () => {
                if (!current || !current.question) return;
                if (embeddedImagesMap[currentQNum] && !current.image) {
                    current.image = embeddedImagesMap[currentQNum];
                }
                if (current.type === 'pg') {
                    if (current.options.A && current.options.B) {
                        results.push(current);
                    }
                } else if (current.type === 'menjodohkan') {
                    if (current.pairs && current.pairs.length >= 2) {
                        results.push(current);
                    }
                } else if (current.type === 'isian' || current.type === 'uraian') {
                    results.push(current);
                }
            };

            lines.forEach(line => {
                // Deteksi awal nomor soal: "1. [PG] ..." atau "1. ..."
                const qMatch = line.match(/^(\d+)[\.\)]\s+(?:\[(PG|MENJODOHKAN|ISIAN|URAIAN)\]\s*)?(.+)$/i);
                if (qMatch) {
                    finalizeCurrent();
                    currentQNum = parseInt(qMatch[1], 10);
                    const rawTag = (qMatch[2] || '').toUpperCase();
                    let detectedType = 'pg';
                    if (rawTag === 'MENJODOHKAN') detectedType = 'menjodohkan';
                    else if (rawTag === 'ISIAN') detectedType = 'isian';
                    else if (rawTag === 'URAIAN') detectedType = 'uraian';

                    current = {
                        id: 'q_doc_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
                        type: detectedType,
                        question: qMatch[3].trim(),
                        image: '',
                        options: { A: '', B: '', C: '', D: '', E: '' },
                        answer: 'A',
                        pairs: [],
                        shortAnswer: '',
                        essayRubric: '',
                        points: 10
                    };
                    return;
                }

                if (!current) return;

                // Deteksi tag [GAMBAR: url_atau_base64]
                const imgTagMatch = line.match(/^\[GAMBAR\s*:\s*(.+?)\]$/i);
                if (imgTagMatch) {
                    current.image = imgTagMatch[1].trim();
                    return;
                }

                // Deteksi pasangan menjodohkan: "1) Kiri = Kanan" atau "- Kiri = Kanan"
                if (current.type === 'menjodohkan') {
                    const pairMatch = line.match(/^(?:(?:\d+[\)\.]|[-•*])\s*)?(.+?)\s*=\s*(.+)$/);
                    if (pairMatch) {
                        current.pairs.push({ left: pairMatch[1].trim(), right: pairMatch[2].trim() });
                        return;
                    }
                }

                // Deteksi Opsi A-E untuk PG
                const optMatch = line.match(/^([A-Ea-e])[\.\)]\s+(.+)$/);
                if (optMatch && current.type === 'pg') {
                    const key = optMatch[1].toUpperCase();
                    current.options[key] = optMatch[2].trim();
                    return;
                }

                // Deteksi Kunci Jawaban
                const keyMatch = line.match(/^(?:Kunci|Jawaban|Answer)\s*[:\-]\s*(.+)$/i);
                if (keyMatch) {
                    const val = keyMatch[1].trim();
                    if (current.type === 'pg') {
                        const letter = val.charAt(0).toUpperCase();
                        if (['A', 'B', 'C', 'D', 'E'].includes(letter)) current.answer = letter;
                    } else if (current.type === 'isian') {
                        current.shortAnswer = val;
                    } else if (current.type === 'uraian') {
                        current.essayRubric = val;
                    }
                    return;
                }

                // Lanjutan baris pertanyaan jika belum masuk opsi/pasangan
                if (current.type === 'pg' && !current.options.A) {
                    current.question += ' ' + line;
                } else if (current.type === 'menjodohkan' && current.pairs.length === 0) {
                    current.question += ' ' + line;
                } else if ((current.type === 'isian' && !current.shortAnswer) || (current.type === 'uraian' && !current.essayRubric)) {
                    current.question += ' ' + line;
                }
            });

            finalizeCurrent();
            return results;
        }

        // =================================================================
        // N. MENU ADMINISTRATOR (7 MENU LENGKAP + SETELAN SEMUA DASHBOARD)
        // =================================================================
        function switchAdminMenu(menu) {
            currentAdminMenu = menu;
            persistSessionState({ adminMenu: menu });
            ['monitoring', 'soal', 'assign', 'logs', 'nilai', 'kartu', 'setelan'].forEach(m => {
                const btn = document.getElementById(`adminNav-${m}`);
                const sec = document.getElementById(`adminSection-${m}`);
                if (!btn || !sec) return;
                if (m === menu) {
                    btn.className = 'admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-indigo-600 text-white transition whitespace-nowrap';
                    sec.classList.remove('hidden');
                } else {
                    btn.className = 'admin-sidebar-btn flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition whitespace-nowrap';
                    sec.classList.add('hidden');
                }
            });
            syncAllViews();
        }

        function populateFilterDropdowns() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const classes = [...new Set(students.map(s => s.className).filter(Boolean))];
            const rooms = [...new Set(students.map(s => s.room).filter(Boolean))];

            const monRoom = document.getElementById('monRoomFilter');
            if (monRoom) {
                const cur = monRoom.value;
                monRoom.innerHTML = `<option value="">Semua Kelas &amp; Ruangan</option>` +
                    classes.map(c => `<option value="class:${c}">Kelas: ${c}</option>`).join('') +
                    rooms.map(r => `<option value="room:${r}">Ruangan: ${r}</option>`).join('');
                monRoom.value = cur;
            }

            const assignTarget = document.getElementById('assignTargetGroup');
            if (assignTarget) {
                const cur = assignTarget.value;
                assignTarget.innerHTML = `<option value="ALL">Semua Kelas &amp; Ruangan</option>` +
                    classes.map(c => `<option value="class:${c}">Khusus Kelas ${c}</option>`).join('') +
                    rooms.map(r => `<option value="room:${r}">Khusus ${r}</option>`).join('');
                if (cur) assignTarget.value = cur;
            }

            const nilaiClass = document.getElementById('nilaiClassFilter');
            if (nilaiClass) {
                const cur = nilaiClass.value;
                nilaiClass.innerHTML = `<option value="">Semua Kelas</option>` + classes.map(c => `<option value="${c}">${c}</option>`).join('');
                nilaiClass.value = cur;
            }

            const nilaiRoom = document.getElementById('nilaiRoomFilter');
            if (nilaiRoom) {
                const cur = nilaiRoom.value;
                nilaiRoom.innerHTML = `<option value="">Semua Ruangan</option>` + rooms.map(r => `<option value="${r}">${r}</option>`).join('');
                nilaiRoom.value = cur;
            }
        }

        // --- MENU 1: MONITORING REAL-TIME ---
        function renderAdminMonitoring() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const search = (document.getElementById('monSearchInput')?.value || '').toLowerCase();
            const roomFilter = document.getElementById('monRoomFilter')?.value || '';

            let activeCount = 0;
            let doneCount = 0;
            let lockedCount = 0;

            students.forEach(s => {
                const st = liveMap[s.id]?.status || 'Belum Masuk';
                if (st === 'Sedang Ujian' || st === 'Di Dashboard Siswa') activeCount++;
                else if (st === 'Selesai') doneCount++;
                else if (st === 'Terkunci' || liveMap[s.id]?.isLocked) lockedCount++;
            });

            document.getElementById('monStatTotal').innerText = students.length;
            document.getElementById('monStatActive').innerText = activeCount;
            document.getElementById('monStatDone').innerText = doneCount;
            document.getElementById('monStatLocked').innerText = lockedCount;

            const filtered = students.filter(s => {
                const matchSearch = s.name.toLowerCase().includes(search) || s.nisn.toLowerCase().includes(search);
                let matchFilter = true;
                if (roomFilter.startsWith('class:')) matchFilter = s.className === roomFilter.replace('class:', '');
                else if (roomFilter.startsWith('room:')) matchFilter = s.room === roomFilter.replace('room:', '');
                return matchSearch && matchFilter;
            });

            const tbody = document.getElementById('adminMonitoringTbody');
            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="py-8 text-center text-slate-400 text-xs">Tidak ada data peserta yang cocok.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(s => {
                const live = liveMap[s.id] || {};
                const status = live.status || 'Belum Masuk';
                const violations = live.violations || 0;
                const progressText = live.totalQuestions ? `${live.answeredCount || 0} / ${live.totalQuestions}` : '-';

                let badge = `<span class="px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 text-[11px] font-bold">Belum Masuk</span>`;
                if (status === 'Di Dashboard Siswa') badge = `<span class="px-2.5 py-1 rounded-full bg-indigo-100 text-indigo-700 text-[11px] font-bold">Di Dashboard</span>`;
                else if (status === 'Sedang Ujian') badge = `<span class="px-2.5 py-1 rounded-full bg-blue-100 text-blue-700 text-[11px] font-bold">Sedang Ujian</span>`;
                else if (status === 'Selesai') badge = `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">Selesai</span>`;
                else if (status === 'Terkunci' || live.isLocked) badge = `<span class="px-2.5 py-1 rounded-full bg-red-600 text-white text-[11px] font-extrabold animate-pulse">TERKUNCI + SIRINE</span>`;

                return `
                    <tr class="hover:bg-slate-50 transition ${live.isLocked ? 'bg-red-50/50' : ''}">
                        <td class="py-3 px-4">
                            <div class="font-bold text-slate-900">${s.name}</div>
                            <div class="text-[11px] font-mono text-slate-500">NISN: ${s.nisn}</div>
                        </td>
                        <td class="py-3 px-4 text-xs text-slate-700">${s.className} • <strong>${s.room}</strong></td>
                        <td class="py-3 px-4 text-xs font-semibold text-slate-700">${live.activeSubject || '-'}</td>
                        <td class="py-3 px-4">${badge}</td>
                        <td class="py-3 px-4 font-mono text-xs font-semibold">${progressText}</td>
                        <td class="py-3 px-4">
                            <span class="px-2 py-0.5 rounded-full text-xs font-bold ${violations > 0 ? 'bg-red-100 text-red-700' : 'bg-slate-100 text-slate-600'}">
                                ${violations}x
                            </span>
                        </td>
                        <td class="py-3 px-4 text-right space-x-1">
                            ${live.isLocked ? `
                                <button onclick="remoteUnlockStudent('${s.id}')" class="px-2.5 py-1.5 rounded-lg bg-amber-400 hover:bg-amber-500 text-slate-950 text-[11px] font-extrabold transition">
                                    Matikan Sirine &amp; Buka
                                </button>
                            ` : ''}
                            <button onclick="resetStudentExamSession('${s.id}')" class="px-2.5 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 text-[11px] font-bold transition">
                                Reset Sesi
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function remoteUnlockStudent(studentId) {
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            if (liveMap[studentId]) {
                liveMap[studentId].isLocked = false;
                liveMap[studentId].remoteUnlocked = true;
                liveMap[studentId].status = 'Sedang Ujian';
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                const students = getDB(STORAGE_KEYS.STUDENTS, []);
                const stu = students.find(x => x.id === studentId);
                if (stu) {
                    recordSystemLog(activeSession?.role || 'admin', activeSession?.user?.name || 'Petugas', '-', stu.room, 'UNLOCK', `Membuka kunci layar & mematikan sirine jarak jauh untuk ${stu.name}`);
                }
                showToast('Perintah matikan sirine & buka kunci layar dikirim ke perangkat siswa!', 'success');
                syncAllViews();
            }
        }

        function resetStudentExamSession(studentId) {
            showConfirmDialog('Reset Sesi Peserta?', 'Status ujian peserta akan direset sehingga dapat login dan mengerjakan ujian ulang.', () => {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                delete liveMap[studentId];
                setDB(STORAGE_KEYS.LIVE_STATUS, liveMap);
                showToast('Sesi peserta berhasil direset.', 'info');
                syncAllViews();
            });
        }

        // --- MENU 2: BANK SOAL MULTI-TIPE & FOTO JPG/PNG ---
        function renderAdminPackageControls() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const sel = document.getElementById('adminPackageSelector');
            if (!sel) return;
            if (packages.length > 0 && !packages.some(p => p.id === selectedAdminPkgId)) {
                selectedAdminPkgId = packages[0].id;
            }
            sel.innerHTML = packages.map(p => `
                <option value="${p.id}">${p.subjectName} (${(p.questions || []).length} Soal) — [${p.status === 'ASSIGNED' ? 'DI-ASSIGN' : 'MENUNGGU ASSIGN'}]</option>
            `).join('');
            sel.value = selectedAdminPkgId;
        }

        function selectAdminQuestionPackage(pkgId) {
            selectedAdminPkgId = pkgId;
            renderAdminQuestionList();
        }

        function createNewSubjectPackage(e) {
            e.preventDefault();
            const subName = document.getElementById('newPkgSubjectName').value.trim();
            if (!subName) return;
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const newPkg = {
                id: 'pkg_' + Date.now(),
                examTitle: settings.examTitle,
                subjectName: subName,
                scheduleDate: new Date().toISOString().split('T')[0],
                scheduleTime: '08:00 - 09:30 WITA',
                durationMinutes: 45,
                targetGroup: 'ALL',
                status: 'WAITING',
                shuffleQuestions: true,
                shuffleOptions: true,
                questions: []
            };
            packages.push(newPkg);
            setDB(STORAGE_KEYS.PACKAGES, packages);
            selectedAdminPkgId = newPkg.id;
            document.getElementById('newPkgSubjectName').value = '';
            showToast(`Paket mata pelajaran "${subName}" berhasil dibuat!`, 'success');
            syncAllViews();
        }

        function handleAddManualQuestion(e, sourceRole) {
            e.preventDefault();
            const pkgId = sourceRole === 'admin' ? selectedAdminPkgId : selectedProctorPkgId;
            const prefix = sourceRole === 'admin' ? 'admin' : 'proctor';

            const qType = document.getElementById(`${prefix}QType`).value;
            const text = document.getElementById(`${prefix}QText`).value.trim();
            const points = parseInt(document.getElementById(`${prefix}QPoints`).value, 10) || 10;
            const imageData = document.getElementById(`${prefix}QImageData`).value || '';

            const newQuestion = {
                id: 'q_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
                type: qType,
                question: text,
                image: imageData,
                points
            };

            if (qType === 'pg') {
                const optA = document.getElementById(`${prefix}QOptA`).value.trim();
                const optB = document.getElementById(`${prefix}QOptB`).value.trim();
                const optC = document.getElementById(`${prefix}QOptC`).value.trim();
                const optD = document.getElementById(`${prefix}QOptD`).value.trim();
                const optE = document.getElementById(`${prefix}QOptE`).value.trim();
                if (!optA || !optB) {
                    showToast('Minimal Opsi A dan Opsi B wajib diisi untuk Pilihan Ganda!', 'error');
                    return;
                }
                newQuestion.options = { A: optA, B: optB, C: optC, D: optD, E: optE };
                newQuestion.answer = document.getElementById(`${prefix}QAnswer`).value;
            } else if (qType === 'menjodohkan') {
                const rows = document.querySelectorAll(`#${prefix}MatchingRows .matching-pair-row`);
                const pairs = [];
                rows.forEach(r => {
                    const l = r.querySelector('.pair-left').value.trim();
                    const rt = r.querySelector('.pair-right').value.trim();
                    if (l && rt) pairs.push({ left: l, right: rt });
                });
                if (pairs.length < 2) {
                    showToast('Masukkan minimal 2 pasangan pernyataan untuk soal Menjodohkan!', 'error');
                    return;
                }
                newQuestion.pairs = pairs;
            } else if (qType === 'isian') {
                const shortAns = document.getElementById(`${prefix}QShortAnswer`).value.trim();
                if (!shortAns) {
                    showToast('Kunci jawaban Isian Singkat wajib diisi!', 'error');
                    return;
                }
                newQuestion.shortAnswer = shortAns;
            } else if (qType === 'uraian') {
                newQuestion.essayRubric = document.getElementById(`${prefix}QEssayRubric`).value.trim();
            }

            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkgIdx = packages.findIndex(p => p.id === pkgId);
            if (pkgIdx === -1) return;

            packages[pkgIdx].questions.push(newQuestion);
            setDB(STORAGE_KEYS.PACKAGES, packages);

            e.target.reset();
            clearQuestionImageInput(prefix);
            const rowsBox = document.getElementById(`${prefix}MatchingRows`);
            if (rowsBox) rowsBox.innerHTML = '';
            toggleQuestionTypeFields(prefix);

            showToast('Soal berhasil disimpan ke Bank Soal!', 'success');
            syncAllViews();
        }

        function renderQuestionPreviewCard(q, idx, pkgId, allowEditDelete = true) {
            const qType = q.type || 'pg';
            const typeBadgeMap = {
                pg: '<span class="px-2 py-0.5 rounded bg-blue-100 text-blue-700 text-[10px] font-bold">Pilihan Ganda</span>',
                menjodohkan: '<span class="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold">Menjodohkan</span>',
                isian: '<span class="px-2 py-0.5 rounded bg-emerald-100 text-emerald-700 text-[10px] font-bold">Isian Singkat</span>',
                uraian: '<span class="px-2 py-0.5 rounded bg-purple-100 text-purple-700 text-[10px] font-bold">Uraian</span>'
            };

            let detailHtml = '';
            if (qType === 'pg') {
                const keys = ['A', 'B', 'C', 'D', 'E'].filter(k => q.options && q.options[k]);
                detailHtml = `
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-1.5 text-xs">
                        ${keys.map(k => `
                            <div class="px-2.5 py-1.5 rounded-lg border ${k === q.answer ? 'bg-emerald-100 border-emerald-300 text-emerald-900 font-bold' : 'bg-white border-slate-200 text-slate-600'} break-anywhere">
                                <strong>${k}.</strong> ${q.options[k]}
                            </div>
                        `).join('')}
                    </div>
                `;
            } else if (qType === 'menjodohkan') {
                detailHtml = `
                    <div class="space-y-1 text-xs">
                        ${(q.pairs || []).map((p, i) => `
                            <div class="px-2.5 py-1.5 rounded-lg bg-amber-50 border border-amber-200 text-amber-950 flex flex-wrap items-center gap-2">
                                <span class="font-bold">${i + 1}. ${p.left}</span>
                                <span class="text-amber-600 font-extrabold">⇄</span>
                                <span class="font-semibold text-emerald-800">${p.right}</span>
                            </div>
                        `).join('')}
                    </div>
                `;
            } else if (qType === 'isian') {
                detailHtml = `
                    <div class="px-3 py-2 rounded-lg bg-emerald-50 border border-emerald-200 text-xs text-emerald-900">
                        <strong>Kunci Jawaban Isian:</strong> ${q.shortAnswer || '-'}
                    </div>
                `;
            } else if (qType === 'uraian') {
                detailHtml = `
                    <div class="px-3 py-2 rounded-lg bg-indigo-50 border border-indigo-200 text-xs text-indigo-900">
                        <strong>Pedoman / Kata Kunci Uraian:</strong> ${q.essayRubric || '(Dinilai berdasarkan kelengkapan uraian)'}
                    </div>
                `;
            }

            const imageHtml = q.image ? `
                <div class="pt-1">
                    <img src="${q.image}" alt="Lampiran Soal" onclick="openImageLightbox(this.src)" class="max-h-32 rounded-xl border border-slate-300 object-contain cursor-zoom-in bg-white p-1">
                </div>
            ` : '';

            return `
                <div class="p-4 rounded-2xl border border-slate-200 bg-slate-50/60 hover:bg-white transition space-y-2.5">
                    <div class="flex items-start justify-between gap-3">
                        <div class="text-xs sm:text-sm font-bold text-slate-900 break-anywhere">
                            <span class="text-indigo-600 mr-1">${idx + 1}.</span> ${q.question}
                            <div class="inline-flex flex-wrap items-center gap-1.5 ml-1 mt-1">
                                ${typeBadgeMap[qType] || ''}
                                <span class="px-2 py-0.5 rounded bg-slate-200 text-slate-700 text-[10px] font-semibold">${q.points || 10} Poin</span>
                            </div>
                        </div>
                        ${allowEditDelete ? `
                            <div class="flex items-center gap-1.5 shrink-0">
                                <button onclick="openEditQuestionModal('${pkgId}', '${q.id}')" class="px-2.5 py-1.5 rounded-lg bg-amber-100 hover:bg-amber-200 text-amber-800 text-xs font-bold transition flex items-center gap-1">
                                    <i data-lucide="edit-3" class="w-3.5 h-3.5"></i> Edit
                                </button>
                                <button onclick="deleteSingleQuestion('${pkgId}', '${q.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 transition" title="Hapus Soal">
                                    <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                                </button>
                            </div>
                        ` : ''}
                    </div>
                    ${imageHtml}
                    ${detailHtml}
                </div>
            `;
        }

        function renderAdminQuestionList() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === selectedAdminPkgId) || packages[0];
            const questions = pkg ? (pkg.questions || []) : [];

            document.getElementById('adminBankSoalCount').innerText = questions.length;
            const container = document.getElementById('adminSoalListContainer');

            if (!pkg || questions.length === 0) {
                container.innerHTML = `<div class="py-10 text-center text-slate-400 text-xs sm:text-sm">Paket soal ini masih kosong. Upload file Word (.docx) / PDF (.pdf) atau tambah soal secara manual.</div>`;
                return;
            }

            container.innerHTML = questions.map((q, idx) => renderQuestionPreviewCard(q, idx, pkg.id, true)).join('');
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        function openEditQuestionModal(pkgId, questionId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg) return;
            const q = (pkg.questions || []).find(item => item.id === questionId);
            if (!q) return;

            const qType = q.type || 'pg';
            document.getElementById('editQPkgId').value = pkgId;
            document.getElementById('editQId').value = q.id;
            document.getElementById('editQType').value = qType;
            document.getElementById('editQPoints').value = q.points || 10;
            document.getElementById('editQText').value = q.question;

            clearQuestionImageInput('edit');
            if (q.image) {
                document.getElementById('editQImageData').value = q.image;
                document.getElementById('editQImagePreview').src = q.image;
                document.getElementById('editQImagePreviewBox').classList.remove('hidden');
                document.getElementById('editQImgClearBtn').classList.remove('hidden');
            }

            toggleQuestionTypeFields('edit');

            if (qType === 'pg') {
                document.getElementById('editQOptA').value = q.options?.A || '';
                document.getElementById('editQOptB').value = q.options?.B || '';
                document.getElementById('editQOptC').value = q.options?.C || '';
                document.getElementById('editQOptD').value = q.options?.D || '';
                document.getElementById('editQOptE').value = q.options?.E || '';
                document.getElementById('editQAnswer').value = q.answer || 'A';
            } else if (qType === 'menjodohkan') {
                const box = document.getElementById('editMatchingRows');
                box.innerHTML = '';
                (q.pairs || []).forEach(p => addMatchingPairRow('editMatchingRows', p.left, p.right));
            } else if (qType === 'isian') {
                document.getElementById('editQShortAnswer').value = q.shortAnswer || '';
            } else if (qType === 'uraian') {
                document.getElementById('editQEssayRubric').value = q.essayRubric || '';
            }

            const modal = document.getElementById('editQuestionModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeEditQuestionModal() {
            const modal = document.getElementById('editQuestionModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function saveEditedQuestion(e) {
            e.preventDefault();
            const pkgId = document.getElementById('editQPkgId').value;
            const qId = document.getElementById('editQId').value;
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkgIdx = packages.findIndex(p => p.id === pkgId);
            if (pkgIdx === -1) return;

            const qIdx = packages[pkgIdx].questions.findIndex(item => item.id === qId);
            if (qIdx === -1) return;

            const qType = document.getElementById('editQType').value;
            const updatedQ = {
                id: qId,
                type: qType,
                question: document.getElementById('editQText').value.trim(),
                image: document.getElementById('editQImageData').value || '',
                points: parseInt(document.getElementById('editQPoints').value, 10) || 10
            };

            if (qType === 'pg') {
                updatedQ.options = {
                    A: document.getElementById('editQOptA').value.trim(),
                    B: document.getElementById('editQOptB').value.trim(),
                    C: document.getElementById('editQOptC').value.trim(),
                    D: document.getElementById('editQOptD').value.trim(),
                    E: document.getElementById('editQOptE').value.trim()
                };
                updatedQ.answer = document.getElementById('editQAnswer').value;
            } else if (qType === 'menjodohkan') {
                const rows = document.querySelectorAll('#editMatchingRows .matching-pair-row');
                const pairs = [];
                rows.forEach(r => {
                    const l = r.querySelector('.pair-left').value.trim();
                    const rt = r.querySelector('.pair-right').value.trim();
                    if (l && rt) pairs.push({ left: l, right: rt });
                });
                if (pairs.length < 2) {
                    showToast('Minimal 2 pasangan pernyataan wajib diisi!', 'error');
                    return;
                }
                updatedQ.pairs = pairs;
            } else if (qType === 'isian') {
                updatedQ.shortAnswer = document.getElementById('editQShortAnswer').value.trim();
            } else if (qType === 'uraian') {
                updatedQ.essayRubric = document.getElementById('editQEssayRubric').value.trim();
            }

            packages[pkgIdx].questions[qIdx] = updatedQ;
            setDB(STORAGE_KEYS.PACKAGES, packages);
            closeEditQuestionModal();
            showToast('Perubahan soal berhasil disimpan!', 'success');
            syncAllViews();
        }

        function deleteSingleQuestion(pkgId, questionId) {
            showConfirmDialog('Hapus Soal Ini?', 'Butir soal ini akan dihapus permanen dari paket.', () => {
                const packages = getDB(STORAGE_KEYS.PACKAGES, []);
                const pkgIdx = packages.findIndex(p => p.id === pkgId);
                if (pkgIdx !== -1) {
                    packages[pkgIdx].questions = packages[pkgIdx].questions.filter(q => q.id !== questionId);
                    setDB(STORAGE_KEYS.PACKAGES, packages);
                    showToast('Butir soal dihapus.', 'info');
                    syncAllViews();
                }
            });
        }

        function confirmClearPackageQuestions() {
            showConfirmDialog('Kosongkan Soal Paket Ini?', 'Semua butir soal pada paket mata pelajaran terpilih akan dihapus.', () => {
                const packages = getDB(STORAGE_KEYS.PACKAGES, []);
                const pkgIdx = packages.findIndex(p => p.id === selectedAdminPkgId);
                if (pkgIdx !== -1) {
                    packages[pkgIdx].questions = [];
                    setDB(STORAGE_KEYS.PACKAGES, packages);
                    showToast('Soal pada paket telah dikosongkan.', 'warning');
                    syncAllViews();
                }
            });
        }

        // --- MENU 3: ASSIGNMENT MATA PELAJARAN, JADWAL & TOKEN ---
        function loadSelectedPackageToAssignForm() {
            const pkgId = document.getElementById('assignPackageSelect').value;
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const pkg = packages.find(p => p.id === pkgId);
            if (!pkg) return;

            document.getElementById('assignExamTitle').value = pkg.examTitle || '';
            document.getElementById('assignSubject').value = pkg.subjectName || '';
            document.getElementById('assignScheduleDate').value = pkg.scheduleDate || '';
            document.getElementById('assignScheduleTime').value = pkg.scheduleTime || '';
            document.getElementById('assignDuration').value = pkg.durationMinutes || 45;
            document.getElementById('assignTargetGroup').value = pkg.targetGroup || 'ALL';
            document.getElementById('assignStatusSelect').value = pkg.status || 'ASSIGNED';
            document.getElementById('assignShuffleQuestions').checked = !!pkg.shuffleQuestions;
            document.getElementById('assignShuffleOptions').checked = !!pkg.shuffleOptions;
        }

        function savePackageAssignment(e) {
            e.preventDefault();
            const pkgId = document.getElementById('assignPackageSelect').value;
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const idx = packages.findIndex(p => p.id === pkgId);
            if (idx === -1) return;

            packages[idx].examTitle = document.getElementById('assignExamTitle').value.trim();
            packages[idx].subjectName = document.getElementById('assignSubject').value.trim();
            packages[idx].scheduleDate = document.getElementById('assignScheduleDate').value;
            packages[idx].scheduleTime = document.getElementById('assignScheduleTime').value.trim();
            packages[idx].durationMinutes = parseInt(document.getElementById('assignDuration').value, 10) || 45;
            packages[idx].targetGroup = document.getElementById('assignTargetGroup').value;
            packages[idx].status = document.getElementById('assignStatusSelect').value;
            packages[idx].shuffleQuestions = document.getElementById('assignShuffleQuestions').checked;
            packages[idx].shuffleOptions = document.getElementById('assignShuffleOptions').checked;

            setDB(STORAGE_KEYS.PACKAGES, packages);
            showToast(`Jadwal & Assignment "${packages[idx].subjectName}" tersinkron ke Dashboard Siswa!`, 'success');
            syncAllViews();
        }

        function togglePackageAssignStatus(pkgId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const idx = packages.findIndex(p => p.id === pkgId);
            if (idx === -1) return;
            packages[idx].status = packages[idx].status === 'ASSIGNED' ? 'WAITING' : 'ASSIGNED';
            setDB(STORAGE_KEYS.PACKAGES, packages);
            showToast(`Status "${packages[idx].subjectName}" diubah menjadi ${packages[idx].status === 'ASSIGNED' ? 'SIAP DIKERJAKAN (ASSIGNED)' : 'MENUNGGU ASSIGN'}.`, 'info');
            syncAllViews();
        }

        function deleteSubjectPackage(pkgId) {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            if (packages.length <= 1) {
                showToast('Minimal harus tersisa 1 paket mata pelajaran.', 'warning');
                return;
            }
            showConfirmDialog('Hapus Paket Mata Pelajaran?', 'Paket soal dan jadwal mata pelajaran ini akan dihapus.', () => {
                const updated = packages.filter(p => p.id !== pkgId);
                setDB(STORAGE_KEYS.PACKAGES, updated);
                selectedAdminPkgId = updated[0].id;
                selectedProctorPkgId = updated[0].id;
                showToast('Paket mata pelajaran dihapus.', 'info');
                syncAllViews();
            });
        }

        function renderAdminAssignmentsTable() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const sel = document.getElementById('assignPackageSelect');
            if (sel) {
                const cur = sel.value;
                sel.innerHTML = packages.map(p => `<option value="${p.id}">${p.subjectName} (${(p.questions || []).length} Soal)</option>`).join('');
                if (cur && packages.some(p => p.id === cur)) sel.value = cur;
                else if (packages.length > 0) {
                    sel.value = packages[0].id;
                    loadSelectedPackageToAssignForm();
                }
            }

            const tbody = document.getElementById('adminAssignmentsTableBody');
            if (!tbody) return;

            tbody.innerHTML = packages.map(p => {
                const isAssigned = p.status === 'ASSIGNED';
                return `
                    <tr class="hover:bg-slate-50 transition">
                        <td class="py-3 px-4">
                            <div class="font-bold text-slate-900">${p.subjectName}</div>
                            <div class="text-[11px] text-slate-500">${p.examTitle}</div>
                        </td>
                        <td class="py-3 px-4 text-xs text-slate-700">
                            <div><strong>${p.scheduleDate || '-'}</strong></div>
                            <div class="text-[11px] text-slate-500">${p.scheduleTime || '-'} (${p.durationMinutes} mnt)</div>
                        </td>
                        <td class="py-3 px-4 text-xs font-semibold text-indigo-700">
                            ${p.targetGroup === 'ALL' ? 'Semua Kelas' : p.targetGroup.replace('class:', 'Kelas ').replace('room:', 'Ruang ')}
                        </td>
                        <td class="py-3 px-4 font-mono text-xs font-bold">${(p.questions || []).length} Soal</td>
                        <td class="py-3 px-4">
                            ${isAssigned
                                ? `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">SIAP DIKERJAKAN</span>`
                                : `<span class="px-2.5 py-1 rounded-full bg-amber-100 text-amber-800 text-[11px] font-bold">MENUNGGU ASSIGN</span>`
                            }
                        </td>
                        <td class="py-3 px-4 text-right space-x-1">
                            <button onclick="togglePackageAssignStatus('${p.id}')" class="px-3 py-1.5 rounded-lg ${isAssigned ? 'bg-amber-100 hover:bg-amber-200 text-amber-800' : 'bg-emerald-600 hover:bg-emerald-700 text-white'} text-[11px] font-bold transition">
                                ${isAssigned ? 'Set Menunggu' : 'Assign Sekarang'}
                            </button>
                            <button onclick="deleteSubjectPackage('${p.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 transition" title="Hapus Paket">
                                <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function generateRandomToken() {
            const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
            let token = '';
            for (let i = 0; i < 6; i++) {
                token += chars.charAt(Math.floor(Math.random() * chars.length));
            }
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.activeToken = token;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            showToast(`Token baru "${token}" aktif & tersinkron otomatis ke Pengawas!`, 'success');
            syncAllViews();
        }

        function saveManualToken() {
            const val = document.getElementById('customTokenInput').value.trim().toUpperCase();
            if (val.length < 4) {
                showToast('Token minimal 4 karakter!', 'error');
                return;
            }
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            settings.activeToken = val;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            document.getElementById('customTokenInput').value = '';
            showToast(`Token "${val}" disimpan & disinkronkan ke Pengawas!`, 'success');
            syncAllViews();
        }

        // --- MENU 4: LOG MASUK/KELUAR & PELANGGARAN ---
        function renderAdminLogsTable() {
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            const search = (document.getElementById('logSearchInput')?.value || '').toLowerCase();
            const roleF = document.getElementById('logRoleFilter')?.value || '';
            const typeF = document.getElementById('logTypeFilter')?.value || '';

            const filtered = logs.filter(l => {
                const matchSearch = l.userName.toLowerCase().includes(search) ||
                    (l.userIdentifier || '').toLowerCase().includes(search) ||
                    l.detail.toLowerCase().includes(search);
                const matchRole = !roleF || l.role === roleF;
                const matchType = !typeF || l.category === typeF;
                return matchSearch && matchRole && matchType;
            });

            const tbody = document.getElementById('adminLogsTableBody');
            if (!tbody) return;

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="py-8 text-center text-slate-400 text-xs">Belum ada rekaman aktivitas sesuai filter.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(l => {
                const roleBadge = l.role === 'siswa'
                    ? `<span class="px-2 py-0.5 rounded bg-blue-100 text-blue-700 text-[10px] font-bold uppercase">Siswa</span>`
                    : l.role === 'pengawas'
                        ? `<span class="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold uppercase">Pengawas</span>`
                        : `<span class="px-2 py-0.5 rounded bg-slate-800 text-white text-[10px] font-bold uppercase">Admin</span>`;

                let catBadge = `<span class="px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-[10px] font-bold">${l.category}</span>`;
                if (l.category === 'LOGIN') catBadge = `<span class="px-2 py-0.5 rounded bg-emerald-100 text-emerald-700 text-[10px] font-bold">MASUK / LOGIN</span>`;
                else if (l.category === 'LOGOUT') catBadge = `<span class="px-2 py-0.5 rounded bg-slate-200 text-slate-700 text-[10px] font-bold">KELUAR / LOGOUT</span>`;
                else if (l.category === 'VIOLATION') catBadge = `<span class="px-2 py-0.5 rounded bg-red-600 text-white text-[10px] font-extrabold">PELANGGARAN + SIRINE</span>`;
                else if (l.category === 'EXAM_START') catBadge = `<span class="px-2 py-0.5 rounded bg-indigo-100 text-indigo-700 text-[10px] font-bold">MULAI UJIAN</span>`;
                else if (l.category === 'EXAM_FINISH') catBadge = `<span class="px-2 py-0.5 rounded bg-teal-100 text-teal-700 text-[10px] font-bold">SELESAI UJIAN</span>`;

                return `
                    <tr class="hover:bg-slate-50 transition ${l.category === 'VIOLATION' ? 'bg-red-50/40' : ''}">
                        <td class="py-2.5 px-4 font-mono text-[11px] text-slate-500 whitespace-nowrap">${l.timestamp}</td>
                        <td class="py-2.5 px-4">${roleBadge}</td>
                        <td class="py-2.5 px-4">
                            <div class="font-bold text-slate-900 text-xs">${l.userName}</div>
                            <div class="text-[10px] text-slate-500">${l.userIdentifier} • ${l.roomOrClass}</div>
                        </td>
                        <td class="py-2.5 px-4">${catBadge}</td>
                        <td class="py-2.5 px-4 text-xs text-slate-700">${l.detail}</td>
                    </tr>
                `;
            }).join('');
        }

        function exportActivityLogsExcel() {
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            if (logs.length === 0) {
                showToast('Belum ada data log untuk diunduh.', 'warning');
                return;
            }
            const rows = logs.map((l, i) => ({
                No: i + 1,
                Waktu: l.timestamp,
                Peran: l.role.toUpperCase(),
                Nama_Pengguna: l.userName,
                Identitas: l.userIdentifier,
                Kelas_Ruang: l.roomOrClass,
                Kategori: l.category,
                Keterangan: l.detail
            }));
            const ws = XLSX.utils.json_to_sheet(rows);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Audit_Log_CBT');
            XLSX.writeFile(wb, `Log_Aktivitas_CBT_${new Date().toISOString().slice(0, 10)}.xlsx`);
            showToast('File Excel Log Aktivitas berhasil diunduh!', 'success');
        }

        function confirmClearActivityLogs() {
            showConfirmDialog('Bersihkan Semua Log Aktivitas?', 'Seluruh catatan log masuk/keluar dan pelanggaran akan dikosongkan.', () => {
                setDB(STORAGE_KEYS.LOGS, []);
                showToast('Log aktivitas berhasil dibersihkan.', 'info');
                renderAdminLogsTable();
            });
        }

        // --- MENU 5: DATA & DOWNLOAD NILAI (.XLSX) ---
        function getFilteredStudentResults() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const search = (document.getElementById('nilaiSearchInput')?.value || '').toLowerCase();
            const classF = document.getElementById('nilaiClassFilter')?.value || '';
            const roomF = document.getElementById('nilaiRoomFilter')?.value || '';

            return students.filter(s => {
                const matchSearch = s.name.toLowerCase().includes(search) || s.nisn.toLowerCase().includes(search);
                const matchClass = !classF || s.className === classF;
                const matchRoom = !roomF || s.room === roomF;
                return matchSearch && matchClass && matchRoom;
            }).map(s => {
                const live = liveMap[s.id] || {};
                return {
                    ...s,
                    activeSubject: live.activeSubject || '-',
                    status: live.status || 'Belum Ujian',
                    correctCount: live.correctCount ?? '-',
                    wrongCount: live.wrongCount ?? '-',
                    score: live.score ?? '-',
                    violations: live.violations ?? 0
                };
            });
        }

        function renderAdminNilaiTable() {
            const rows = getFilteredStudentResults();
            const tbody = document.getElementById('adminNilaiTbody');

            if (rows.length === 0) {
                tbody.innerHTML = `<tr><td colspan="10" class="py-8 text-center text-slate-400 text-xs">Tidak ada data nilai peserta.</td></tr>`;
                return;
            }

            tbody.innerHTML = rows.map((r, i) => `
                <tr class="hover:bg-slate-50 transition">
                    <td class="py-3 px-3 text-slate-500">${i + 1}</td>
                    <td class="py-3 px-3 font-bold text-slate-900">${r.name}</td>
                    <td class="py-3 px-3 font-mono text-xs text-slate-600">${r.nisn}</td>
                    <td class="py-3 px-3 text-xs">${r.className} • <strong>${r.room}</strong></td>
                    <td class="py-3 px-3 text-xs font-semibold text-indigo-700">${r.activeSubject}</td>
                    <td class="py-3 px-3">
                        <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${r.status === 'Selesai' ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600'}">
                            ${r.status}
                        </span>
                    </td>
                    <td class="py-3 px-3 font-bold text-emerald-600">${r.correctCount}</td>
                    <td class="py-3 px-3 font-bold text-red-500">${r.wrongCount}</td>
                    <td class="py-3 px-3 font-extrabold text-blue-700 text-sm">${r.score}</td>
                    <td class="py-3 px-3 text-xs font-bold ${r.violations > 0 ? 'text-red-600' : 'text-slate-400'}">${r.violations}x</td>
                </tr>
            `).join('');
        }

        function downloadExcelResults() {
            if (typeof XLSX === 'undefined') {
                showToast('Library Excel belum termuat.', 'error');
                return;
            }
            const rows = getFilteredStudentResults();
            if (rows.length === 0) {
                showToast('Tidak ada data peserta untuk diunduh.', 'warning');
                return;
            }

            const excelData = rows.map((r, idx) => ({
                'No': idx + 1,
                'Nama Peserta': r.name,
                'NISN': r.nisn,
                'Kelas': r.className,
                'Ruangan': r.room,
                'Mata Pelajaran': r.activeSubject,
                'Status Ujian': r.status,
                'Jawaban Benar': r.correctCount,
                'Jawaban Salah': r.wrongCount,
                'Nilai Akhir': r.score,
                'Jumlah Pelanggaran': r.violations
            }));

            const worksheet = XLSX.utils.json_to_sheet(excelData);
            worksheet['!cols'] = [
                { wch: 5 }, { wch: 26 }, { wch: 15 }, { wch: 14 }, { wch: 14 },
                { wch: 28 }, { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 14 }, { wch: 18 }
            ];

            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap_Nilai_CBT');
            const dateStr = new Date().toISOString().slice(0, 10);
            XLSX.writeFile(workbook, `Rekap_Nilai_CBT_${dateStr}.xlsx`);
            showToast('File Excel (.XLSX) nilai ujian berhasil diunduh!', 'success');
        }

        // --- MENU 6: AKUN SISWA SIAP DIGUNAKAN & KARTU LOGIN ---
        function generateRandomStudentPassword() {
            return String(Math.floor(10000 + Math.random() * 90000));
        }

        function toggleShowStudentPasswords() {
            showStudentPasswords = !showStudentPasswords;
            renderAdminStudentAccountTable();
            renderAdminLoginCards();
        }

        function handleAddManualStudent(e) {
            e.preventDefault();
            const name = document.getElementById('kStuName').value.trim();
            const nisn = document.getElementById('kStuNisn').value.trim();
            const className = document.getElementById('kStuClass').value.trim();
            const room = document.getElementById('kStuRoom').value.trim();
            const customUser = document.getElementById('kStuUser').value.trim();
            const customPass = document.getElementById('kStuPass').value.trim();

            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            if (students.some(s => s.nisn === nisn)) {
                showToast(`NISN ${nisn} sudah terdaftar!`, 'error');
                return;
            }

            students.push({
                id: 'stu_' + Date.now(),
                name,
                nisn,
                className,
                room,
                username: customUser || nisn,
                password: customPass || generateRandomStudentPassword()
            });

            setDB(STORAGE_KEYS.STUDENTS, students);
            e.target.reset();
            showToast(`Akun peserta "${name}" siap digunakan & kartu dibuat!`, 'success');
            syncAllViews();
        }

        function downloadStudentExcelTemplate() {
            if (typeof XLSX === 'undefined') return;
            const sample = [
                { Nama: 'Rizky Pratama', NISN: '0081234510', Kelas: 'XII IPA 1', Ruangan: 'Ruang 01' },
                { Nama: 'Aulia Rahma', NISN: '0081234511', Kelas: 'XII IPS 2', Ruangan: 'Ruang 02' }
            ];
            const ws = XLSX.utils.json_to_sheet(sample);
            ws['!cols'] = [{ wch: 25 }, { wch: 16 }, { wch: 15 }, { wch: 15 }];
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Template_Peserta');
            XLSX.writeFile(wb, 'Template_Import_Akun_Siswa_CBT.xlsx');
            showToast('Template Excel peserta diunduh!', 'info');
        }

        function handleStudentExcelUpload(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheetName = workbook.SheetNames[0];
                    const jsonRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

                    if (jsonRows.length === 0) {
                        showToast('File Excel kosong!', 'error');
                        return;
                    }

                    const students = getDB(STORAGE_KEYS.STUDENTS, []);
                    let added = 0;

                    jsonRows.forEach((row, idx) => {
                        const nama = String(row.Nama || row.nama || row['Nama Lengkap'] || '').trim();
                        const nisn = String(row.NISN || row.nisn || '').trim();
                        const kelas = String(row.Kelas || row.kelas || 'Umum').trim();
                        const ruang = String(row.Ruangan || row.ruangan || row.Ruang || 'Ruang 01').trim();
                        const pass = String(row.Password || row.password || '').trim() || generateRandomStudentPassword();

                        if (nama && nisn && !students.some(s => s.nisn === nisn)) {
                            students.push({
                                id: 'stu_' + Date.now() + '_' + idx,
                                name: nama,
                                nisn,
                                className: kelas,
                                room: ruang,
                                username: nisn,
                                password: pass
                            });
                            added++;
                        }
                    });

                    setDB(STORAGE_KEYS.STUDENTS, students);
                    showToast(`Berhasil mengimpor ${added} akun siswa siap pakai!`, 'success');
                    syncAllViews();
                } catch (err) {
                    showToast('Gagal membaca file Excel: ' + err.message, 'error');
                } finally {
                    event.target.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function exportStudentAccountsExcel() {
            if (typeof XLSX === 'undefined') return;
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            if (students.length === 0) {
                showToast('Belum ada akun siswa.', 'warning');
                return;
            }
            const data = students.map((s, idx) => ({
                'No': idx + 1,
                'Nama Peserta': s.name,
                'NISN': s.nisn,
                'Kelas': s.className,
                'Ruangan': s.room,
                'Username Login': s.username,
                'Password Login': s.password
            }));
            const ws = XLSX.utils.json_to_sheet(data);
            ws['!cols'] = [{ wch: 5 }, { wch: 25 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 16 }];
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Akun_Login_Siswa');
            XLSX.writeFile(wb, `Daftar_Akun_Login_Siswa_${new Date().toISOString().slice(0, 10)}.xlsx`);
            showToast('Daftar akun & password siswa berhasil diunduh (.XLSX)!', 'success');
        }

        function openEditStudentModal(stuId) {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const stu = students.find(s => s.id === stuId);
            if (!stu) return;

            document.getElementById('editStuId').value = stu.id;
            document.getElementById('editStuName').value = stu.name;
            document.getElementById('editStuNisn').value = stu.nisn;
            document.getElementById('editStuClass').value = stu.className;
            document.getElementById('editStuRoom').value = stu.room;
            document.getElementById('editStuUser').value = stu.username;
            document.getElementById('editStuPass').value = stu.password;

            const modal = document.getElementById('editStudentModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeEditStudentModal() {
            const modal = document.getElementById('editStudentModal');
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        function generateRandomStudentPassInModal() {
            document.getElementById('editStuPass').value = generateRandomStudentPassword();
        }

        function saveEditedStudent(e) {
            e.preventDefault();
            const id = document.getElementById('editStuId').value;
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const idx = students.findIndex(s => s.id === id);
            if (idx === -1) return;

            students[idx] = {
                ...students[idx],
                name: document.getElementById('editStuName').value.trim(),
                nisn: document.getElementById('editStuNisn').value.trim(),
                className: document.getElementById('editStuClass').value.trim(),
                room: document.getElementById('editStuRoom').value.trim(),
                username: document.getElementById('editStuUser').value.trim(),
                password: document.getElementById('editStuPass').value.trim()
            };

            setDB(STORAGE_KEYS.STUDENTS, students);
            closeEditStudentModal();
            showToast('Akun & Password siswa berhasil diperbarui!', 'success');
            syncAllViews();
        }

        function deleteSingleStudent(stuId) {
            showConfirmDialog('Hapus Akun Siswa?', 'Akun dan kartu login peserta ini akan dihapus.', () => {
                const students = getDB(STORAGE_KEYS.STUDENTS, []).filter(s => s.id !== stuId);
                setDB(STORAGE_KEYS.STUDENTS, students);
                showToast('Akun siswa dihapus.', 'info');
                syncAllViews();
            });
        }

        function confirmClearAllStudents() {
            showConfirmDialog('Kosongkan Semua Akun Siswa?', 'Seluruh akun dan kartu login peserta akan dihapus.', () => {
                setDB(STORAGE_KEYS.STUDENTS, []);
                setDB(STORAGE_KEYS.LIVE_STATUS, {});
                showToast('Semua akun siswa dikosongkan.', 'warning');
                syncAllViews();
            });
        }

        function renderAdminStudentAccountTable() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const search = (document.getElementById('accountSearchInput')?.value || '').toLowerCase();
            document.getElementById('studentAccountCount').innerText = students.length;

            const visBtn = document.getElementById('btnTogglePassVis');
            if (visBtn) {
                visBtn.innerHTML = showStudentPasswords
                    ? `<i data-lucide="eye-off" class="w-3.5 h-3.5"></i> Sembunyikan Password`
                    : `<i data-lucide="eye" class="w-3.5 h-3.5"></i> Tampilkan Password`;
            }

            const filtered = students.filter(s =>
                s.name.toLowerCase().includes(search) ||
                s.nisn.toLowerCase().includes(search) ||
                s.className.toLowerCase().includes(search) ||
                s.room.toLowerCase().includes(search)
            );

            const tbody = document.getElementById('adminStudentAccountTbody');
            if (!tbody) return;

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="py-8 text-center text-slate-400">Tidak ada akun siswa ditemukan.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(s => `
                <tr class="hover:bg-slate-50 transition">
                    <td class="py-2.5 px-3">
                        <div class="font-bold text-slate-900">${s.name}</div>
                        <div class="text-[11px] font-mono text-slate-500">NISN: ${s.nisn}</div>
                    </td>
                    <td class="py-2.5 px-3 text-slate-700">${s.className} • <strong>${s.room}</strong></td>
                    <td class="py-2.5 px-3 font-mono font-bold text-slate-800">${s.username}</td>
                    <td class="py-2.5 px-3 font-mono font-extrabold text-blue-700">
                        ${showStudentPasswords ? s.password : '•••••'}
                    </td>
                    <td class="py-2.5 px-3 text-right space-x-1 whitespace-nowrap">
                        <button onclick="copyTextToClipboard('Username: ${s.username} | Password: ${s.password}', 'Akun ${s.name} disalin!')" class="p-1.5 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-600" title="Salin Akun">
                            <i data-lucide="copy" class="w-3.5 h-3.5"></i>
                        </button>
                        <button onclick="openEditStudentModal('${s.id}')" class="p-1.5 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-700" title="Edit / Reset Password">
                            <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
                        </button>
                        <button onclick="deleteSingleStudent('${s.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600" title="Hapus Akun">
                            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                        </button>
                    </td>
                </tr>
            `).join('');
        }

        function renderAdminLoginCards() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
            const container = document.getElementById('printCardsContainer');
            document.getElementById('cardTotalCount').innerText = students.length;

            if (students.length === 0) {
                container.innerHTML = `<div class="col-span-full py-10 text-center text-slate-400 text-xs">Belum ada kartu peserta. Tambahkan peserta secara manual atau import melalui Excel.</div>`;
                return;
            }

            container.innerHTML = students.map(s => `
                <div class="login-card-item bg-white border-2 border-slate-300 rounded-2xl p-4 shadow-sm flex flex-col justify-between space-y-3">
                    <div class="flex items-center justify-between border-b border-slate-200 pb-2.5 gap-2">
                        <div class="flex items-center gap-2.5 min-w-0">
                            <div class="w-9 h-9 rounded-xl bg-slate-900 text-white flex items-center justify-center font-extrabold text-xs shrink-0">CBT</div>
                            <div class="min-w-0">
                                <div class="text-[11px] font-extrabold uppercase tracking-wider text-slate-900">KARTU LOGIN PESERTA UJIAN</div>
                                <div class="text-[10px] text-slate-500 truncate">${settings.schoolName}</div>
                            </div>
                        </div>
                    </div>

                    <div class="grid grid-cols-2 gap-2 text-xs">
                        <div>
                            <span class="text-[10px] text-slate-400 block">Nama Lengkap</span>
                            <span class="font-bold text-slate-900 break-anywhere">${s.name}</span>
                        </div>
                        <div>
                            <span class="text-[10px] text-slate-400 block">NISN</span>
                            <span class="font-mono font-bold text-slate-800">${s.nisn}</span>
                        </div>
                        <div>
                            <span class="text-[10px] text-slate-400 block">Kelas</span>
                            <span class="font-semibold text-slate-800">${s.className}</span>
                        </div>
                        <div>
                            <span class="text-[10px] text-slate-400 block">Ruangan Ujian</span>
                            <span class="font-bold text-indigo-700">${s.room}</span>
                        </div>
                    </div>

                    <div class="bg-slate-100 rounded-xl p-2.5 grid grid-cols-2 gap-2 border border-slate-200">
                        <div>
                            <span class="text-[10px] text-slate-500 uppercase font-bold block">Username</span>
                            <span class="font-mono font-extrabold text-xs text-slate-900">${s.username}</span>
                        </div>
                        <div>
                            <span class="text-[10px] text-slate-500 uppercase font-bold block">Password</span>
                            <span class="font-mono font-extrabold text-xs text-blue-700">${s.password}</span>
                        </div>
                    </div>
                </div>
            `).join('');
        }

        // --- MENU 7: SETELAN KESELURUHAN & PENGATURAN SEMUA DASHBOARD ---
        function renderAdminSettingsForms() {
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            const proctors = getDB(STORAGE_KEYS.PROCTORS, []);

            document.getElementById('setSchoolName').value = settings.schoolName || '';
            document.getElementById('setSchoolSubtitle').value = settings.schoolSubtitle || '';
            document.getElementById('setAdminUsername').value = settings.adminUsername || 'admin';
            document.getElementById('setAdminPassword').value = settings.adminPassword || 'admin123';

            document.getElementById('setUnlockPin').value = settings.unlockPin || '123';
            document.getElementById('setSingleAppLock').checked = settings.singleAppLock !== false;
            document.getElementById('setRequireFullscreen').checked = !!settings.requireFullscreen;
            document.getElementById('setDetectTabSwitch').checked = !!settings.detectTabSwitch;
            document.getElementById('setEnableSiren').checked = settings.enableSiren !== false;
            document.getElementById('setPortalAnimatedBg').checked = settings.portalAnimatedBg !== false;

            // Pengaturan Dashboard Siswa & Pengawas
            document.getElementById('setStuBannerMsg').value = settings.stuBannerMsg || '';
            document.getElementById('setStuShowProfileTab').checked = settings.stuShowProfileTab !== false;
            document.getElementById('setStuShowGuideTab').checked = settings.stuShowGuideTab !== false;
            document.getElementById('setStuAllowFontResize').checked = settings.stuAllowFontResize !== false;
            document.getElementById('setShowScoreToStudent').checked = !!settings.showScoreToStudent;

            document.getElementById('setProcBannerMsg').value = settings.procBannerMsg || '';
            document.getElementById('setProcShowToken').checked = settings.procShowToken !== false;
            document.getElementById('setProcAllowQuestionInput').checked = settings.procAllowQuestionInput !== false;
            document.getElementById('setProcAllowRemoteUnlock').checked = settings.procAllowRemoteUnlock !== false;

            document.getElementById('setGasUrlDisplay').value = GAS_WEB_APP_URL;

            const procContainer = document.getElementById('proctorListContainer');
            procContainer.innerHTML = proctors.map(p => `
                <div class="flex items-center justify-between p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs">
                    <div>
                        <div class="font-bold text-slate-900">${p.name} <span class="text-amber-700">(${p.room})</span></div>
                        <div class="font-mono text-[11px] text-slate-500">User: ${p.username} • Pass: ${p.password}</div>
                    </div>
                    <button onclick="deleteProctor('${p.id}')" class="p-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-red-600">
                        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                    </button>
                </div>
            `).join('');
        }

        function saveAllDashboardSettings() {
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            settings.stuBannerMsg = document.getElementById('setStuBannerMsg').value.trim();
            settings.stuShowProfileTab = document.getElementById('setStuShowProfileTab').checked;
            settings.stuShowGuideTab = document.getElementById('setStuShowGuideTab').checked;
            settings.stuAllowFontResize = document.getElementById('setStuAllowFontResize').checked;
            settings.showScoreToStudent = document.getElementById('setShowScoreToStudent').checked;

            settings.procBannerMsg = document.getElementById('setProcBannerMsg').value.trim();
            settings.procShowToken = document.getElementById('setProcShowToken').checked;
            settings.procAllowQuestionInput = document.getElementById('setProcAllowQuestionInput').checked;
            settings.procAllowRemoteUnlock = document.getElementById('setProcAllowRemoteUnlock').checked;

            setDB(STORAGE_KEYS.SETTINGS, settings);
            showToast('Pengaturan Dashboard Siswa & Dashboard Pengawas berhasil disimpan & disinkronkan!', 'success');
            syncAllViews();
        }

        function saveSchoolIdentity() {
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            settings.schoolName = document.getElementById('setSchoolName').value.trim() || DEFAULT_SETTINGS.schoolName;
            settings.schoolSubtitle = document.getElementById('setSchoolSubtitle').value.trim() || DEFAULT_SETTINGS.schoolSubtitle;
            settings.adminUsername = document.getElementById('setAdminUsername').value.trim() || 'admin';
            settings.adminPassword = document.getElementById('setAdminPassword').value.trim() || 'admin123';
            setDB(STORAGE_KEYS.SETTINGS, settings);
            showToast('Identitas sekolah & akun Administrator diperbarui!', 'success');
            syncAllViews();
        }

        function saveSecuritySettings() {
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };
            settings.unlockPin = document.getElementById('setUnlockPin').value.trim() || '123';
            settings.singleAppLock = document.getElementById('setSingleAppLock').checked;
            settings.requireFullscreen = document.getElementById('setRequireFullscreen').checked;
            settings.detectTabSwitch = document.getElementById('setDetectTabSwitch').checked;
            settings.enableSiren = document.getElementById('setEnableSiren').checked;
            settings.portalAnimatedBg = document.getElementById('setPortalAnimatedBg').checked;
            setDB(STORAGE_KEYS.SETTINGS, settings);
            showToast('Setelan keamanan Single-App Lock, sirine & portal disimpan!', 'success');
            syncAllViews();
        }

        function handleAddProctor(e) {
            e.preventDefault();
            const proctors = getDB(STORAGE_KEYS.PROCTORS, []);
            proctors.push({
                id: 'proc_' + Date.now(),
                name: document.getElementById('procName').value.trim(),
                room: document.getElementById('procRoom').value.trim(),
                username: document.getElementById('procUser').value.trim(),
                password: document.getElementById('procPass').value.trim()
            });
            setDB(STORAGE_KEYS.PROCTORS, proctors);
            e.target.reset();
            showToast('Akun pengawas ruangan ditambahkan!', 'success');
            syncAllViews();
        }

        function deleteProctor(id) {
            const proctors = getDB(STORAGE_KEYS.PROCTORS, []).filter(p => p.id !== id);
            setDB(STORAGE_KEYS.PROCTORS, proctors);
            showToast('Akun pengawas dihapus.', 'info');
            syncAllViews();
        }

        function handleAddPortalNews(e) {
            e.preventDefault();
            const title = document.getElementById('newsTitleInput').value.trim();
            const category = document.getElementById('newsCategoryInput').value;
            const content = document.getElementById('newsContentInput').value.trim();
            if (!title || !content) return;

            const newsList = getDB(STORAGE_KEYS.NEWS, DEFAULT_NEWS);
            newsList.unshift({
                id: 'news_' + Date.now(),
                category,
                title,
                date: new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
                author: 'Administrator Sekolah',
                excerpt: content.length > 135 ? content.slice(0, 132) + '...' : content,
                content
            });
            setDB(STORAGE_KEYS.NEWS, newsList);
            e.target.reset();
            showToast('Berita pendidikan berhasil dipublikasikan ke Portal!', 'success');
            syncAllViews();
        }

        function handleAddGalleryItem(e) {
            e.preventDefault();
            const title = document.getElementById('galTitle').value.trim();
            const url = document.getElementById('galUrl').value.trim();
            const gallery = getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY);
            gallery.unshift({
                id: 'gal_' + Date.now(),
                title,
                url,
                svgType: 'lab'
            });
            setDB(STORAGE_KEYS.GALLERY, gallery);
            e.target.reset();
            showToast('Foto kegiatan ditambahkan ke Galeri Portal!', 'success');
            syncAllViews();
        }

        function exportDatabaseJSON() {
            const payload = {
                settings: getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS),
                packages: getDB(STORAGE_KEYS.PACKAGES, []),
                students: getDB(STORAGE_KEYS.STUDENTS, []),
                proctors: getDB(STORAGE_KEYS.PROCTORS, []),
                liveStatus: getDB(STORAGE_KEYS.LIVE_STATUS, {}),
                logs: getDB(STORAGE_KEYS.LOGS, []),
                gallery: getDB(STORAGE_KEYS.GALLERY, DEFAULT_GALLERY),
                news: getDB(STORAGE_KEYS.NEWS, DEFAULT_NEWS)
            };
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Backup_CBT_${new Date().toISOString().slice(0, 10)}.json`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            showToast('Cadangan database JSON diunduh!', 'success');
        }

        function importDatabaseJSON(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const parsed = JSON.parse(e.target.result);
                    if (parsed.settings) setDB(STORAGE_KEYS.SETTINGS, parsed.settings);
                    if (parsed.packages) setDB(STORAGE_KEYS.PACKAGES, parsed.packages);
                    if (parsed.students) setDB(STORAGE_KEYS.STUDENTS, parsed.students);
                    if (parsed.proctors) setDB(STORAGE_KEYS.PROCTORS, parsed.proctors);
                    if (parsed.liveStatus) setDB(STORAGE_KEYS.LIVE_STATUS, parsed.liveStatus);
                    if (parsed.logs) setDB(STORAGE_KEYS.LOGS, parsed.logs);
                    if (parsed.gallery) setDB(STORAGE_KEYS.GALLERY, parsed.gallery);
                    if (parsed.news) setDB(STORAGE_KEYS.NEWS, parsed.news);
                    showToast('Database berhasil dipulihkan dari file JSON!', 'success');
                    syncAllViews();
                } catch (err) {
                    showToast('File JSON tidak valid!', 'error');
                } finally {
                    event.target.value = '';
                }
            };
            reader.readAsText(file);
        }

        function confirmCleanStartEmpty() {
            showConfirmDialog('Mulai dari Kosong (Hapus Data Contoh)?', 'Semua data peserta contoh, soal contoh, dan log akan dikosongkan agar Anda dapat mengisi data sekolah asli.', () => {
                setDB(STORAGE_KEYS.STUDENTS, []);
                setDB(STORAGE_KEYS.LIVE_STATUS, {});
                setDB(STORAGE_KEYS.LOGS, []);
                const settings = getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS);
                setDB(STORAGE_KEYS.PACKAGES, [{
                    id: 'pkg_clean_1',
                    examTitle: settings.examTitle,
                    subjectName: 'Mata Pelajaran 1',
                    scheduleDate: new Date().toISOString().split('T')[0],
                    scheduleTime: '08:00 - 09:30 WITA',
                    durationMinutes: 60,
                    targetGroup: 'ALL',
                    status: 'WAITING',
                    shuffleQuestions: true,
                    shuffleOptions: true,
                    questions: []
                }]);
                selectedAdminPkgId = 'pkg_clean_1';
                selectedProctorPkgId = 'pkg_clean_1';
                showToast('Sistem siap digunakan dari kosong!', 'success');
                syncAllViews();
            });
        }

        function confirmFactoryReset() {
            showConfirmDialog('Reset Pengaturan Bawaan?', 'Seluruh data akan dikembalikan ke kondisi awal aplikasi.', () => {
                Object.values(STORAGE_KEYS).forEach(k => localStorage.removeItem(k));
                ensureLocalDatabaseDefaults();
                showToast('Aplikasi telah direset ke pengaturan awal.', 'warning');
                logoutSession();
            });
        }

        // =================================================================
        // O. DASHBOARD PENGAWAS (3 MENU LENGKAP & SINKRON)
        // =================================================================
        function switchProctorMenu(menu) {
            currentProctorMenu = menu;
            persistSessionState({ proctorMenu: menu });
            ['kehadiran', 'pelanggaran', 'soal'].forEach(m => {
                const btn = document.getElementById(`procMenuBtn-${m}`);
                const sec = document.getElementById(`proctorSection-${m}`);
                if (!btn || !sec) return;
                if (m === menu) {
                    btn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-amber-500 text-white transition flex items-center gap-2';
                    sec.classList.remove('hidden');
                } else {
                    btn.className = 'px-4 py-2.5 rounded-xl text-xs sm:text-sm font-semibold text-slate-600 hover:bg-slate-100 transition flex items-center gap-2';
                    sec.classList.add('hidden');
                }
            });
            syncAllViews();
        }

        function renderProctorMonitoring() {
            if (!activeSession || activeSession.role !== 'pengawas') return;
            const proctorRoom = activeSession.user.room;
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const logs = getDB(STORAGE_KEYS.LOGS, []);

            const roomStudents = students.filter(s =>
                proctorRoom.toLowerCase().includes('semua') || s.room.toLowerCase() === proctorRoom.toLowerCase()
            );

            const container = document.getElementById('proctorStudentCards');
            if (roomStudents.length === 0) {
                container.innerHTML = `<div class="p-6 text-center text-slate-400 text-xs border border-dashed border-slate-200 rounded-2xl">Belum ada peserta terdaftar di ${proctorRoom}.</div>`;
            } else {
                container.innerHTML = roomStudents.map(s => {
                    const live = liveMap[s.id] || {};
                    const st = live.status || 'Belum Masuk';
                    const isLocked = !!live.isLocked;

                    let badge = `<span class="px-2.5 py-1 rounded-full bg-slate-100 text-slate-600 text-[11px] font-bold">Belum Masuk</span>`;
                    if (st === 'Di Dashboard Siswa') badge = `<span class="px-2.5 py-1 rounded-full bg-indigo-100 text-indigo-700 text-[11px] font-bold">Di Dashboard</span>`;
                    else if (st === 'Sedang Ujian') badge = `<span class="px-2.5 py-1 rounded-full bg-blue-100 text-blue-700 text-[11px] font-bold">Sedang Ujian (${live.activeSubject || ''})</span>`;
                    else if (st === 'Selesai') badge = `<span class="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-700 text-[11px] font-bold">Selesai</span>`;
                    else if (isLocked) badge = `<span class="px-2.5 py-1 rounded-full bg-red-600 text-white text-[11px] font-extrabold animate-pulse">TERKUNCI + SIRINE</span>`;

                    return `
                        <div class="p-4 rounded-2xl border ${isLocked ? 'border-red-400 bg-red-50/70' : 'border-slate-200 bg-slate-50/40'} flex flex-wrap items-center justify-between gap-3">
                            <div>
                                <div class="font-bold text-slate-900 text-sm">${s.name}</div>
                                <div class="text-xs text-slate-500">NISN: ${s.nisn} • ${s.className} • <strong>${s.room}</strong></div>
                            </div>
                            <div class="flex items-center gap-2">
                                ${badge}
                                ${isLocked ? `
                                    <button onclick="remoteUnlockStudent('${s.id}')" class="px-3 py-1.5 rounded-xl bg-amber-400 hover:bg-amber-500 text-slate-950 text-xs font-extrabold transition shadow-sm">
                                        Matikan Sirine &amp; Buka
                                    </button>
                                ` : ''}
                            </div>
                        </div>
                    `;
                }).join('');
            }

            const studentLogs = logs.filter(l => l.role === 'siswa').slice(0, 15);
            const inOutEl = document.getElementById('proctorInOutLogList');
            if (studentLogs.length === 0) {
                inOutEl.innerHTML = `<div class="text-slate-400 text-center py-4">Belum ada log keluar/masuk siswa.</div>`;
            } else {
                inOutEl.innerHTML = studentLogs.map(l => `
                    <div class="p-2.5 rounded-xl border border-slate-100 bg-slate-50 space-y-0.5">
                        <div class="flex justify-between text-[10px] font-bold">
                            <span class="${l.category === 'VIOLATION' ? 'text-red-600' : l.category === 'LOGIN' ? 'text-emerald-600' : 'text-slate-600'}">${l.category}</span>
                            <span class="text-slate-400 font-mono">${l.timestamp}</span>
                        </div>
                        <div class="font-bold text-slate-800">${l.userName}</div>
                        <div class="text-slate-600">${l.detail}</div>
                    </div>
                `).join('');
            }
        }

        function renderProctorViolationsView() {
            const students = getDB(STORAGE_KEYS.STUDENTS, []);
            const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
            const logs = getDB(STORAGE_KEYS.LOGS, []);
            const settings = { ...DEFAULT_SETTINGS, ...getDB(STORAGE_KEYS.SETTINGS, DEFAULT_SETTINGS) };

            const lockedStudents = students.filter(s => liveMap[s.id]?.isLocked);
            const pill = document.getElementById('proctorViolationCountPill');
            if (pill) pill.innerText = lockedStudents.length;

            const lockedBox = document.getElementById('proctorActiveLockedBox');
            if (lockedStudents.length === 0) {
                lockedBox.innerHTML = `
                    <div class="bg-emerald-50 border border-emerald-200 rounded-2xl p-4 text-xs font-semibold text-emerald-800 flex items-center gap-2">
                        <i data-lucide="check-circle" class="w-4 h-4 text-emerald-600"></i> Tidak ada layar peserta yang sedang terkunci saat ini. Kondisi ruangan kondusif.
                    </div>
                `;
            } else {
                lockedBox.innerHTML = lockedStudents.map(s => {
                    const live = liveMap[s.id];
                    return `
                        <div class="bg-red-600 text-white rounded-2xl p-5 shadow-lg flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                            <div>
                                <span class="px-2.5 py-0.5 rounded-full bg-white/20 text-[10px] font-extrabold uppercase">SIRINE ALARM BERBUNYI</span>
                                <h4 class="text-lg font-extrabold mt-1">${s.name} (${s.className} • ${s.room})</h4>
                                <p class="text-xs text-red-100 mt-0.5">Alasan: ${live.lastViolationReason || 'Keluar dari aplikasi ujian'} • Pelanggaran ke-${live.violations || 1}</p>
                            </div>
                            ${settings.procAllowRemoteUnlock !== false ? `
                                <button onclick="remoteUnlockStudent('${s.id}')" class="px-5 py-3 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 font-extrabold text-xs sm:text-sm transition shadow self-start sm:self-center">
                                    Matikan Sirine &amp; Buka Kunci Layar
                                </button>
                            ` : `<span class="text-xs font-bold bg-white/15 px-3 py-2 rounded-xl">Gunakan PIN: ${settings.unlockPin}</span>`}
                        </div>
                    `;
                }).join('');
            }

            const vLogs = logs.filter(l => l.category === 'VIOLATION');
            const tbody = document.getElementById('proctorViolationLogsTbody');
            if (vLogs.length === 0) {
                tbody.innerHTML = `<tr><td colspan="3" class="py-6 text-center text-slate-400 text-xs">Belum ada catatan pelanggaran.</td></tr>`;
            } else {
                tbody.innerHTML = vLogs.map(v => `
                    <tr class="bg-red-50/30">
                        <td class="py-2.5 px-4 font-mono text-xs text-slate-500">${v.timestamp}</td>
                        <td class="py-2.5 px-4 font-bold text-slate-900 text-xs">${v.userName} (${v.userIdentifier})</td>
                        <td class="py-2.5 px-4 text-xs text-red-700 font-semibold">${v.detail}</td>
                    </tr>
                `).join('');
            }
        }

        function selectProctorQuestionPackage(pkgId) {
            selectedProctorPkgId = pkgId;
            renderProctorSoalView();
        }

        function renderProctorSoalView() {
            const packages = getDB(STORAGE_KEYS.PACKAGES, []);
            const sel = document.getElementById('proctorPackageSelector');
            if (!sel) return;
            if (packages.length > 0 && !packages.some(p => p.id === selectedProctorPkgId)) {
                selectedProctorPkgId = packages[0].id;
            }
            sel.innerHTML = packages.map(p => `<option value="${p.id}">${p.subjectName} (${(p.questions || []).length} Soal)</option>`).join('');
            sel.value = selectedProctorPkgId;

            const pkg = packages.find(p => p.id === selectedProctorPkgId) || packages[0];
            const questions = pkg ? (pkg.questions || []) : [];
            document.getElementById('proctorPkgSoalCount').innerText = questions.length;

            const listEl = document.getElementById('proctorSoalListContainer');
            if (!pkg || questions.length === 0) {
                listEl.innerHTML = `<div class="py-8 text-center text-slate-400 text-xs">Belum ada soal pada paket ini.</div>`;
            } else {
                listEl.innerHTML = questions.map((q, idx) => renderQuestionPreviewCard(q, idx, pkg.id, false)).join('');
            }
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        // =================================================================
        // P. BOOTSTRAP & SINKRONISASI REAL-TIME ANTAR TAB / CLOUD
        // =================================================================
        window.addEventListener('storage', (e) => {
            if (e.key === STORAGE_KEYS.LIVE_STATUS && examSession.isRunning && examSession.isLocked && activeSession?.user) {
                const liveMap = getDB(STORAGE_KEYS.LIVE_STATUS, {});
                if (liveMap[activeSession.user.id]?.remoteUnlocked) {
                    unlockExamScreen();
                }
            }
            syncAllViews();
        });

        window.addEventListener('beforeunload', (e) => {
            if (examSession && examSession.isRunning) {
                persistActiveExamState();
                e.preventDefault();
                e.returnValue = '';
            }
        });

        function restoreInterruptedExamSession(savedExam) {
            if (!savedExam || !savedExam.questions || savedExam.questions.length === 0) return false;
            const stu = activeSession.user;
            const elapsedSec = Math.max(0, Math.floor((Date.now() - (savedExam.savedAt || Date.now())) / 1000));
            const remaining = Math.max(30, (savedExam.remainingSeconds || 600) - elapsedSec);

            examSession = {
                isRunning: true,
                isLocked: !!savedExam.isLocked,
                packageId: savedExam.packageId,
                packageSubject: savedExam.packageSubject,
                packageTitle: savedExam.packageTitle,
                questions: savedExam.questions,
                answers: savedExam.answers || {},
                doubts: savedExam.doubts || {},
                currentIdx: savedExam.currentIdx || 0,
                remainingSeconds: remaining,
                timerId: null,
                violations: savedExam.violations || 0
            };

            document.getElementById('studentExamTitle').innerText = `${examSession.packageSubject} — ${examSession.packageTitle}`;
            document.getElementById('studentExamName').innerText = stu.name;
            document.getElementById('studentExamNisn').innerText = stu.nisn;
            document.getElementById('studentExamRoom').innerText = `${stu.className} (${stu.room})`;
            document.getElementById('examTotalNum').innerText = examSession.questions.length;

            navigateTo('siswa');
            renderExamQuestion();
            renderExamNavigationGrid();
            startExamCountdown();

            if (examSession.isLocked) {
                startSirenAlarm();
                document.getElementById('lockReasonText').innerText = 'Halaman ujian dimuat ulang saat layar dalam status terkunci!';
                document.getElementById('lockStudentName').innerText = stu.name;
                document.getElementById('lockStudentMeta').innerText = `${stu.nisn} / ${stu.className} (${stu.room})`;
                document.getElementById('lockExamSubject').innerText = examSession.packageSubject || '-';
                document.getElementById('lockViolationBadge').innerText = `${examSession.violations} Kali`;
                const overlay = document.getElementById('antiCheatLockOverlay');
                overlay.classList.remove('hidden');
                overlay.classList.add('flex');
            }
            return true;
        }

        document.addEventListener('DOMContentLoaded', async () => {
            ensureLocalDatabaseDefaults();
            initSingleAppSecurityLock();

            // Pastikan semua modal dalam kondisi tertutup rapi saat boot
            ['confirmModal', 'antiCheatLockOverlay', 'loginModal', 'studentTokenEntryModal', 'editStudentModal', 'editQuestionModal', 'installPlatformModal', 'examFinishedModal', 'newsReaderModal', 'imageLightboxModal', 'gasCodeModal'].forEach(id => {
                const el = document.getElementById(id);
                if (el) {
                    el.classList.add('hidden');
                    el.classList.remove('flex');
                }
            });

            await syncFromCloud(false);

            const savedSession = getDB(STORAGE_KEYS.ACTIVE_SESSION, null);
            const savedUi = getDB(STORAGE_KEYS.UI_STATE, {});
            const savedExam = getDB(STORAGE_KEYS.ACTIVE_EXAM, null);

            if (savedSession && savedSession.role && savedSession.user) {
                activeSession = savedSession;
                if (savedUi.adminMenu) currentAdminMenu = savedUi.adminMenu;
                if (savedUi.proctorMenu) currentProctorMenu = savedUi.proctorMenu;
                if (savedUi.studentDashTab) currentStudentDashTab = savedUi.studentDashTab;

                if (activeSession.role === 'siswa' && savedExam && savedExam.studentId === activeSession.user.id) {
                    if (!restoreInterruptedExamSession(savedExam)) {
                        navigateTo('siswa-dashboard');
                        switchStudentDashTab(currentStudentDashTab);
                    }
                } else if (activeSession.role === 'admin') {
                    navigateTo(savedUi.currentView === 'portal' ? 'portal' : 'admin');
                    switchAdminMenu(currentAdminMenu);
                } else if (activeSession.role === 'pengawas') {
                    navigateTo(savedUi.currentView === 'portal' ? 'portal' : 'pengawas');
                    switchProctorMenu(currentProctorMenu);
                } else {
                    navigateTo(savedUi.currentView === 'portal' ? 'portal' : 'siswa-dashboard');
                    switchStudentDashTab(currentStudentDashTab);
                }
            } else {
                navigateTo('portal');
                const params = new URLSearchParams(window.location.search);
                const roleParam = params.get('role');
                if (roleParam && ['siswa', 'pengawas', 'admin'].includes(roleParam)) {
                    openLoginModal(roleParam);
                }
            }

            // Sinkronisasi periodik ke Google Apps Script setiap 8 detik
            setInterval(() => {
                syncFromCloud(false);
            }, 8000);
        });
    </script>
</body>
</html>
