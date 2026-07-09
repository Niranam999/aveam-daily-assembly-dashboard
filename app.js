// AVEAM Assembly Progress Tracker Mockup Logic
// High-Fidelity Single-Screen Layout with Carousel

// Initial datasets
const INITIAL_OPERATORS = [
    "WANLOP CHANPHET",
    "SUWAT JANSIRI",
    "THANANCHAI THATHUMNUK",
    "NATTHAVUT TORUANG",
    "KESORN KETKAEW",
    "ASSAMA SINGSA",
    "SUMOLTHA BANCHUEN",
    "KAIRUNG LORTHONG",
    "THITIMA PANSONGNAM"
];

const INITIAL_UPDATES = [
    {
        timestamp: "02/07/2026 13:58",
        team_leader: "WANLOP CHANPHET",
        customer: "UIC",
        project_code: "SCRAP TAPE CUTTER DRIVE KIT",
        job_number: "47938",
        progress: 95,
        note: "Final Inspection and Packing เรียบร้อย"
    },
    {
        timestamp: "02/07/2026 11:30",
        team_leader: "THANANCHAI THATHUMNUK",
        customer: "Veeco",
        project_code: "ACBOX",
        job_number: "49077",
        progress: 95,
        note: "ติดตั้งสายไฟบนพาเนลย่อยสำเร็จ"
    },
    {
        timestamp: "02/07/2026 10:15",
        team_leader: "SUWAT JANSIRI",
        customer: "Veeco",
        project_code: "WF",
        job_number: "42510",
        progress: 75,
        note: "ประกอบชุด Air Gauge Sensor และ Lamp base"
    }
];

// App State
let isCloudMode = false;
let supabaseClient = null;
let googleSheetUrl = "";

const localProjects = JSON.parse(localStorage.getItem('aveam_projects'));
let projects = EXCEL_PROJECTS;

if (localProjects) {
    // Keep latest Excel structure but overlay saved progress
    projects = EXCEL_PROJECTS.map(excelProj => {
        const savedProj = localProjects.find(lp => lp.id === excelProj.id);
        if (savedProj) {
            excelProj.progress = savedProj.progress;
            excelProj.status = savedProj.status;
            excelProj.qty_done = savedProj.qty_done || 0;
            excelProj.sub_assemblies = excelProj.sub_assemblies.map(excelSub => {
                const savedSub = savedProj.sub_assemblies.find(ss => ss.pn === excelSub.pn);
                if (savedSub) {
                    excelSub.progress = savedSub.progress;
                }
                return excelSub;
            });
        }
        return excelProj;
    });
}
let updates = JSON.parse(localStorage.getItem('aveam_updates')) || INITIAL_UPDATES;
let activeFormProjectId = null;
let currentFilter = 'ALL';
let selectedProjectId = null;

// Carousel & Progress State
let carouselTimer = null;
let progressTimer = null;
let progressPercent = 0;
let isAutoRotating = true;
let pauseTimeout = null;

const CAROUSEL_DURATION = 15000; // 15 seconds
const PROGRESS_INTERVAL = 100;    // update bar every 100ms

// DOM Elements
const projectsGrid = document.getElementById('thumbnails-list');
const recentList = document.getElementById('recent-list');
const filterContainer = document.getElementById('customer-filters');
const kanbanBacklog = document.getElementById('kanban-backlog');
const kanbanAssigned = document.getElementById('kanban-assigned');
const kanbanInProgress = document.getElementById('kanban-inprogress');
const kanbanQA = document.getElementById('kanban-qa');
const kanbanCompleted = document.getElementById('kanban-completed');
const statOnTime = document.getElementById('stat-ontime');
const statAtRisk = document.getElementById('stat-atrisk');
const statDelayed = document.getElementById('stat-delayed');
const selectedJobPanel = document.getElementById('selected-job-panel');
const carouselBar = document.getElementById('carousel-bar');

// Mobile DOM Elements
const scannerPrompt = document.getElementById('scanner-prompt');
const mobileForm = document.getElementById('mobile-form');
const formProjectTitle = document.getElementById('form-project-title');
const formJobNo = document.getElementById('form-job-no');
const selectOperator = document.getElementById('select-operator');
const subSlidersContainer = document.getElementById('sub-sliders-container');
const inputQtyDone = document.getElementById('input-qty-done');
const selectIssue = document.getElementById('select-issue');
const textareaNote = document.getElementById('textarea-note');
const successScreen = document.getElementById('success-screen');

// Initialize Dashboard
async function init() {
    // Initialize Theme
    const savedTheme = localStorage.getItem('aveam_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeSelector(savedTheme);

    updateSystemTime();
    setInterval(updateSystemTime, 1000);
    
    // Try to load Supabase configuration for Cloud Mode
    try {
        const response = await fetch('supabase_config.json');
        if (response.ok) {
            const config = await response.json();
            const url = config.supabase_url;
            const key = config.supabase_key;
            googleSheetUrl = config.google_sheet_url || "";
            
            if (url && key) {
                const { createClient } = supabase;
                supabaseClient = createClient(url, key);
                
                // Load data from Supabase
                const { data: projData, error: projErr } = await supabaseClient
                    .from('projects')
                    .select('*')
                    .order('customer', { ascending: true })
                    .order('project_code', { ascending: true });
                    
                if (projErr) throw projErr;
                projects = projData || [];
                
                const { data: logData, error: logErr } = await supabaseClient
                    .from('updates')
                    .select('*')
                    .order('timestamp', { ascending: false })
                    .limit(10);
                    
                if (logErr) throw logErr;
                updates = logData || [];
                
                isCloudMode = true;
                console.log("Connected to Supabase Cloud Mode.");
                
                // Add indicator to title
                const headerTitle = document.querySelector('.header-title-area h1');
                if (headerTitle) {
                    headerTitle.innerHTML += ' <span class="company-badge" style="font-size: 0.65rem; padding: 0.15rem 0.4rem; vertical-align: middle; background: rgba(16, 185, 129, 0.15); border-color: rgba(16, 185, 129, 0.3); color: #10b981;">Cloud Mode</span>';
                }
                
                // Setup Real-time Listeners for instant TV dashboard reaction
                setupRealtimeListeners();
            }
        }
    } catch (e) {
        console.warn("Failed to initialize Cloud Mode, falling back to Local Demo Mode:", e);
    }
    
    // Fallback to Local Demo Mode indicator
    if (!isCloudMode) {
        const headerTitle = document.querySelector('.header-title-area h1');
        if (headerTitle) {
            headerTitle.innerHTML += ' <span class="company-badge" style="font-size: 0.65rem; padding: 0.15rem 0.4rem; vertical-align: middle;">Local Demo</span>';
        }
    }
    
    // Default select first project
    const filtered = getFilteredProjects();
    if (filtered.length > 0) {
        const stillExists = filtered.some(p => p.id === selectedProjectId);
        if (!stillExists) {
            selectedProjectId = filtered[0].id;
        }
    } else {
        selectedProjectId = null;
    }
    
    renderFilters();
    updateStats();
    renderProjects();
    renderRecentUpdates();
    populateOperators();
    
    // Start Carousel
    startCarousel();
}

function setupRealtimeListeners() {
    if (!supabaseClient) return;
    
    supabaseClient.channel('index-realtime-projects')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'projects' }, (payload) => {
            const eventType = payload.eventType;
            const newRecord = payload.new;
            const oldRecord = payload.old;
            
            if (eventType === 'INSERT') {
                projects.push(newRecord);
            } else if (eventType === 'UPDATE') {
                const idx = projects.findIndex(p => p.id === newRecord.id);
                if (idx !== -1) projects[idx] = newRecord;
            } else if (eventType === 'DELETE') {
                projects = projects.filter(p => p.id !== oldRecord.id);
            }
            
            updateStats();
            renderProjects();
            
            if (selectedProjectId && (eventType === 'UPDATE' && newRecord.id === selectedProjectId)) {
                renderDetailPanel(newRecord);
            }
        })
        .subscribe();
        
    supabaseClient.channel('index-realtime-logs')
        .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'updates' }, (payload) => {
            updates.unshift(payload.new);
            if (updates.length > 10) updates.pop();
            renderRecentUpdates();
        })
        .subscribe();
}

function updateSystemTime() {
    const now = new Date();
    document.getElementById('live-time').textContent = now.toLocaleTimeString('th-TH');
}

function renderFilters() {
    const customers = ['ALL', 'VEECO', 'COHU', 'UIC', 'ULC'];
    filterContainer.innerHTML = '';
    customers.forEach(c => {
        const btn = document.createElement('button');
        btn.className = `filter-btn ${c === currentFilter ? 'active' : ''}`;
        btn.textContent = c;
        btn.onclick = (e) => {
            e.stopPropagation();
            currentFilter = c;
            renderFilters();
            
            // Adjust selected project to first item in filtered list
            const filtered = getFilteredProjects();
            if (filtered.length > 0) {
                selectProject(filtered[0].id);
            } else {
                selectedProjectId = null;
                renderDetailPanel(null);
            }
            renderProjects();
            resetCarousel();
        };
        filterContainer.appendChild(btn);
    });
}

function getFilteredProjects() {
    if (currentFilter === 'ALL') return projects.filter(p => p.kanban_in_progress > 0);
    return projects.filter(p => p.customer.toUpperCase() === currentFilter && p.kanban_in_progress > 0);
}

function updateStats() {
    let backlog = 0;
    let assigned = 0;
    let inProgress = 0;
    let qa = 0;
    let completed = 0;
    
    projects.forEach(p => {
        backlog += p.kanban_backlog || 0;
        assigned += p.kanban_assigned || 0;
        inProgress += p.kanban_in_progress || 0;
        qa += p.kanban_qa || 0;
        completed += p.kanban_completed || 0;
    });
    
    if (kanbanBacklog) kanbanBacklog.textContent = backlog;
    if (kanbanAssigned) kanbanAssigned.textContent = assigned;
    if (kanbanInProgress) kanbanInProgress.textContent = inProgress;
    if (kanbanQA) kanbanQA.textContent = qa;
    if (kanbanCompleted) kanbanCompleted.textContent = completed;
    
    if (statOnTime) statOnTime.textContent = projects.filter(p => p.status === 'ontime').length;
    if (statAtRisk) statAtRisk.textContent = projects.filter(p => p.status === 'atrisk').length;
    if (statDelayed) statDelayed.textContent = projects.filter(p => p.status === 'delayed').length;
}

function populateOperators() {
    selectOperator.innerHTML = '<option value="" disabled selected>-- เลือกผู้รายงาน --</option>';
    INITIAL_OPERATORS.forEach(op => {
        const opt = document.createElement('option');
        opt.value = op;
        opt.textContent = op;
        selectOperator.appendChild(opt);
    });
}

function getProjectTotalHours(project) {
    if (!project) return 0;
    if (project.customer && project.customer.toUpperCase() === 'ULC') {
        if (project.sub_assemblies && project.sub_assemblies.length > 0) {
            let total = 0;
            project.sub_assemblies.forEach(sub => {
                total += (sub.qty || 0) * (sub.std_time || 0);
            });
            return total > 0 ? total : 361.5;
        }
        return 361.5;
    }
    return project.qty * project.est_hours;
}

// ==========================================================================
// RENDER DETAIL PANEL (TOP SELECTED JOB VIEW)
// ==========================================================================
function renderDetailPanel(project) {
    if (!project) {
        selectedJobPanel.innerHTML = `
            <div style="grid-column: 1 / span 4; display: flex; align-items: center; justify-content: center; height: 100%; color: var(--text-muted);">
                ไม่มีรายการจ๊อบในกลุ่มลูกค้าที่เลือก
            </div>
        `;
        return;
    }

    // Circumference of r=85 is 2 * PI * 85 = 534
    const circumference = 534;
    const offset = circumference - (project.progress / 100) * circumference;

    // Dynamic Layout Calculation based on sub-assembly count to prevent overflow (hide scrollbar)
    const numSubs = project.sub_assemblies.length;
    let containerGap = '1.25rem';
    let itemWidth = '24px';
    let trackWidth = '10px';
    let labelRotation = '45deg';
    let labelFontSize = '0.65rem';
    let valueFontSize = '0.65rem';

    if (numSubs > 18) {
        containerGap = '4px';
        itemWidth = '11px';
        trackWidth = '5px';
        labelRotation = '65deg';
        labelFontSize = '0.52rem';
        valueFontSize = '0.52rem';
    } else if (numSubs > 12) {
        containerGap = '8px';
        itemWidth = '16px';
        trackWidth = '7px';
        labelRotation = '55deg';
        labelFontSize = '0.6rem';
        valueFontSize = '0.6rem';
    } else if (numSubs > 6) {
        containerGap = '14px';
        itemWidth = '20px';
        trackWidth = '9px';
        labelRotation = '48deg';
        labelFontSize = '0.65rem';
        valueFontSize = '0.65rem';
    }

    // Build sub-assemblies vertical bars HTML
    let barsHtml = '';
    project.sub_assemblies.forEach(sub => {
        const barColor = getProgressColor(sub.progress);
        
        barsHtml += `
            <div class="vertical-bar-item" data-tooltip="${sub.pn}: ${sub.name} (${sub.progress}%)" style="width: ${itemWidth}; flex-shrink: 1;">
                <span class="vertical-bar-value" style="font-size: ${valueFontSize}; font-weight: 700; margin-bottom: 0.25rem; color: ${sub.progress > 0 ? barColor : 'var(--text-muted)'}">${sub.progress}%</span>
                <div class="vertical-bar-track" style="width: ${trackWidth};">
                    <div class="vertical-bar-fill" style="height: ${sub.progress}%; background-color: ${barColor}"></div>
                </div>
                <span class="vertical-bar-label" style="font-size: ${labelFontSize}; transform: rotate(${labelRotation});">${sub.pn}${sub.th_desc ? ' ' + sub.th_desc : ''}</span>
            </div>
        `;
    });

    // Product Image HTML rendering
    const imgPath = getProductImage(project);
    let productImageHtml = '';
    if (imgPath) {
        productImageHtml = `
            <div class="product-image-box">
                <img src="${imgPath}" alt="${project.project_code}" class="product-img">
            </div>
        `;
    }

    // Calculate FG percentage
    const qtyDone = project.qty_done || 0;
    const qtyTarget = project.qty || 1;
    const fgPercent = Math.min(100, Math.round((qtyDone / qtyTarget) * 100));

    selectedJobPanel.innerHTML = `
        <!-- Column 1: Info & Product Image & FG Status -->
        <div class="detail-left" style="display: flex; flex-direction: row; justify-content: space-between; align-items: stretch; border-right: 1px solid var(--card-border); padding-right: 1rem;">
            <!-- Left Info Block -->
            <div class="detail-left-main" style="display: flex; flex-direction: column; justify-content: space-between; flex: 1; min-width: 0; padding-right: 0.5rem;">
                <div>
                    <span class="thumb-client ${project.customer.toLowerCase()}">${project.customer}</span>
                    <h3 class="detail-project-name" style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-top: 0.25rem;">${project.project_code}</h3>
                    <p class="detail-job-number" style="font-size: 0.68rem; line-height: 1.45;">
                        Part Number: <strong>${project.part_number || '-'}</strong><br>
                        Job: <strong>${project.job_no}</strong>${project.mc_number && project.mc_number !== '-' ? ' (MC: <strong>' + project.mc_number + '</strong>)' : ''} | Qty: <strong>${project.qty}</strong><br>
                        Time: <strong>${Number(project.est_hours).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</strong> hrs/unit (Total: <strong>${getProjectTotalHours(project).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> hrs)
                    </p>
                    <div class="detail-desc-box" style="margin-top: 0.25rem;">
                        Description: <em>${project.description}</em>
                    </div>
                </div>
                ${productImageHtml}
            </div>
            
            <!-- Right FG Status Bar (Mockup Style) -->
            <div class="detail-fg-status" style="display: flex; flex-direction: column; align-items: center; justify-content: space-between; width: 68px; flex-shrink: 0; padding-left: 0.25rem; border-left: 1px solid var(--card-border);">
                <div style="text-align: center;">
                    <span style="display: block; font-size: 0.6rem; color: #fb923c; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">FG Status</span>
                    <span style="font-size: 0.85rem; font-weight: 800; color: var(--text-title);">${qtyDone}/${qtyTarget}</span>
                </div>
                <div class="fg-bar-track">
                    <div class="fg-bar-fill" style="height: ${fgPercent}%;"></div>
                </div>
                <span style="font-size: 0.68rem; font-weight: 700; color: #10b981;">${fgPercent}%</span>
            </div>
        </div>
        
        <!-- Column 2: Donut Progress Chart -->
        <div class="detail-middle">
            <div class="big-progress-circle-wrap ${project.progress === 100 ? 'celebrate-100' : ''}">
                <svg width="210" height="210">
                    <circle class="big-progress-bg" cx="105" cy="105" r="85"></circle>
                    <circle class="big-progress-bar" cx="105" cy="105" r="85" 
                        style="stroke-dasharray: ${circumference}; stroke-dashoffset: ${offset}; stroke: ${getProgressColor(project.progress)}"></circle>
                </svg>
                <span class="big-progress-percentage">${project.progress}%</span>
                ${project.progress === 100 ? '<canvas id="celebration-canvas" class="celebration-canvas" width="210" height="210"></canvas>' : ''}
            </div>
            <div class="detail-meta-text">
                <span class="stat-sub" style="font-size: 0.7rem;">ความคืบหน้ารวมเฉลี่ย</span>
                <span class="detail-lead-label">Lead: <span>${project.team_leader}</span></span>
                <span class="detail-lead-label" style="margin-top: 2px;">แม่ทีม: <span>${project.member_1}</span></span>
            </div>
        </div>
        
        <!-- Column 3: Sub-Assemblies Vertical Bar Chart -->
        <div class="detail-right">
            <div class="detail-subs-header">พาร์ทประกอบย่อยและสถานะรายตัว (${project.sub_assemblies.length})</div>
            <div class="detail-vertical-chart-container" style="gap: ${containerGap}; justify-content: ${numSubs > 6 ? 'space-between' : 'flex-start'};">
                ${barsHtml}
            </div>
        </div>

        <!-- Column 4: Far Right QR & Action -->
        <div class="detail-qr-side">
            <div class="qr-title-box">
                <strong>SCAN JOB TRAVELER</strong>
                สแกน QR Code นี้ด้วยมือถือพนักงานเพื่อบันทึก Man-Hours และอัปเดตสถานะการประกอบ
            </div>
            <div class="big-qr-box" title="สแกนเพื่อทำงานประกอบ">
                <img src="https://api.qrserver.com/v1/create-qr-code/?size=110x110&data=https://aveam.com/update?job=${project.job_no}" alt="Job QR">
            </div>
            <button class="submit-btn" onclick="triggerQRScan('${project.id}')" style="width: 100%; font-size: 0.72rem; height: 32px;">
                🔍 สแกนอัพเดทจ๊อบนี้ (Scan QR)
            </button>
        </div>
    `;

    // Trigger celebration if progress is 100%
    if (project.progress === 100) {
        setTimeout(initCelebration, 50);
    } else {
        stopCelebration();
    }
}

// ==========================================================================
// RENDER THUMBNAILS LIST (BOTTOM GRID)
// ==========================================================================
function renderProjects() {
    projectsGrid.innerHTML = '';
    const filtered = getFilteredProjects();
    
    if (filtered.length === 0) {
        projectsGrid.innerHTML = `
            <div style="grid-column: 1 / span 4; display: flex; align-items: center; justify-content: center; height: 100%; color: var(--text-muted);">
                ไม่มีจ๊อบที่สอดคล้องกับตัวกรอง
            </div>
        `;
        return;
    }
    
    filtered.forEach(p => {
        const isActive = p.id === selectedProjectId;
        
        // Circular progress parameters (radius = 16, circumference = 2 * PI * 16 = 100.5)
        const circumference = 100.5;
        const offset = circumference - (p.progress / 100) * circumference;
        
        const card = document.createElement('div');
        card.id = `card-${p.id}`;
        card.className = `thumbnail-card ${isActive ? 'active' : ''}`;
        card.onclick = () => {
            selectProject(p.id);
            pauseAutoRotation(); // Pause on user interaction
        };
        
        card.innerHTML = `
            <div class="thumb-info">
                <span class="thumb-client ${p.customer.toLowerCase()}">${p.customer}</span>
                <span class="thumb-code">${p.project_code}</span>
                <span class="thumb-job">Job: ${p.job_no}${p.mc_number && p.mc_number !== '-' ? ' (' + p.mc_number + ')' : ''}</span>
            </div>
            
            <div class="thumb-progress-wrap">
                <svg width="38" height="38">
                    <circle class="thumb-progress-bg" cx="19" cy="19" r="16"></circle>
                    <circle class="thumb-progress-bar" cx="19" cy="19" r="16" 
                        style="stroke-dasharray: ${circumference}; stroke-dashoffset: ${offset}; stroke: ${getProgressColor(p.progress)}"></circle>
                </svg>
                <span class="thumb-progress-percentage" style="color: var(--text-primary)">${p.progress}%</span>
            </div>
        `;
        
        projectsGrid.appendChild(card);
    });
    
    // Trigger detail rendering for currently selected
    const activeProject = projects.find(p => p.id === selectedProjectId);
    renderDetailPanel(activeProject);
}

// Select a project
function selectProject(projectId) {
    selectedProjectId = projectId;
    
    // Re-active card layout highlight
    const cards = projectsGrid.querySelectorAll('.thumbnail-card');
    cards.forEach(c => c.classList.remove('active'));
    
    const targetCard = document.getElementById(`card-${projectId}`);
    if (targetCard) {
        targetCard.classList.add('active');
    }
    
    const project = projects.find(p => p.id === projectId);
    renderDetailPanel(project);
}

// Render Recent Updates log
function renderRecentUpdates() {
    recentList.innerHTML = '';
    
    if (updates.length === 0) {
        recentList.innerHTML = '<div style="color: var(--text-muted); text-align: center; padding: 1rem; font-size: 0.75rem;">ไม่มีบันทึกประวัติย้อนหลัง</div>';
        return;
    }
    
    updates.slice(0, 10).forEach(u => {
        const div = document.createElement('div');
        div.className = 'recent-item';
        div.innerHTML = `
            <div class="recent-item-header">
                <span>${u.project_code} (Job: ${u.job_number})</span>
                <span style="color: var(--brand-color);">${u.progress}%</span>
            </div>
            <div class="recent-item-meta">
                โดย: <strong>${u.reporter || u.team_leader}</strong> ${u.qty_done ? `| FG: <strong>+${u.qty_done}</strong>` : ''} | ${u.timestamp}
            </div>
            <div class="recent-item-note">
                ${u.note || 'ไม่มีข้อมูลบันทึกเพิ่มเติม'}
            </div>
        `;
        recentList.appendChild(div);
    });
}

// ==========================================================================
// ROTATION / CAROUSEL TIMERS
// ==========================================================================
function startCarousel() {
    resetCarouselTimers();
    
    // Interval for progress fill
    progressPercent = 0;
    progressTimer = setInterval(() => {
        if (!isAutoRotating) return;
        
        progressPercent += (PROGRESS_INTERVAL / CAROUSEL_DURATION) * 100;
        if (progressPercent >= 100) {
            progressPercent = 0;
            rotateNextProject();
        }
        carouselBar.style.width = `${progressPercent}%`;
    }, PROGRESS_INTERVAL);
}

function rotateNextProject() {
    const filtered = getFilteredProjects();
    if (filtered.length <= 1) return;
    
    let currentIndex = filtered.findIndex(p => p.id === selectedProjectId);
    let nextIndex = (currentIndex + 1) % filtered.length;
    
    selectProject(filtered[nextIndex].id);
    
    // Scroll active card into view
    const targetCard = document.getElementById(`card-${filtered[nextIndex].id}`);
    if (targetCard) {
        targetCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}

function pauseAutoRotation() {
    isAutoRotating = false;
    carouselBar.style.width = '0%';
    
    // Clear existing timeout
    if (pauseTimeout) clearTimeout(pauseTimeout);
    
    // Resume auto-rotation after 60 seconds of user inactivity
    pauseTimeout = setTimeout(() => {
        isAutoRotating = true;
        progressPercent = 0;
    }, 60000);
}

function resetCarousel() {
    isAutoRotating = true;
    progressPercent = 0;
    if (pauseTimeout) clearTimeout(pauseTimeout);
}

function resetCarouselTimers() {
    if (progressTimer) clearInterval(progressTimer);
}

// ==========================================================================
// MOBILE UPDATE APP PORTAL INTEGRATION
// ==========================================================================
function triggerQRScan(projectId) {
    const project = projects.find(p => p.id === projectId);
    if (!project) return;
    
    activeFormProjectId = projectId;
    
    scannerPrompt.style.display = 'none';
    mobileForm.style.display = 'block';
    successScreen.style.display = 'none';
    
    formProjectTitle.textContent = `${project.customer} - ${project.project_code} (PN: ${project.part_number || '-'})`;
    formJobNo.textContent = project.job_no;
    
    // Set MC Number and Build Leader info in the mobile form
    const formMcNumber = document.getElementById('form-mc-number');
    const formBuildLeader = document.getElementById('form-build-leader');
    if (formMcNumber) formMcNumber.textContent = project.mc_number || '-';
    if (formBuildLeader) {
        const fullLeader = project.team_leader || '-';
        formBuildLeader.textContent = (fullLeader !== '-' && fullLeader !== 'Unassigned') ? fullLeader.trim().split(/\s+/)[0] : '-';
    }
    
    selectOperator.value = '';
    
    // Load sub-assembly progress sliders
    subSlidersContainer.innerHTML = '';
    project.sub_assemblies.forEach((sub, idx) => {
        const div = document.createElement('div');
        div.className = 'slider-group';
        div.innerHTML = `
            <div class="slider-meta">
                <span class="slider-label" title="${sub.name}">${sub.pn}: ${sub.name}</span>
                <span class="slider-val" id="slide-val-${idx}">${sub.progress}%</span>
            </div>
            <input type="range" class="slider-input" min="0" max="100" step="5" 
                value="${sub.progress}" 
                oninput="document.getElementById('slide-val-${idx}').textContent = this.value + '%'">
        `;
        subSlidersContainer.appendChild(div);
    });
    
    // Reset other input fields
    inputQtyDone.value = '';
    selectIssue.value = 'none';
    textareaNote.value = '';
    
    // Bounce main detail card to show loading connection
    selectedJobPanel.style.transform = 'scale(1.02)';
    setTimeout(() => selectedJobPanel.style.transform = 'none', 300);
}

function cancelMobileUpdate() {
    scannerPrompt.style.display = 'flex';
    mobileForm.style.display = 'none';
    successScreen.style.display = 'none';
    activeFormProjectId = null;
}

// Submit progress update from simulated phone
function submitProgressUpdate(event) {
    event.preventDefault();
    if (!activeFormProjectId) return;
    
    const project = projects.find(p => p.id === activeFormProjectId);
    if (!project) return;
    
    const operator = selectOperator.value;
    if (!operator) {
        alert("กรุณาเลือกชื่อผู้รายงาน");
        return;
    }
    
    // 1. Update sub-assemblies progress from sliders
    const sliderInputs = subSlidersContainer.querySelectorAll('.slider-input');
    let totalProgress = 0;
    const updatedSubs = project.sub_assemblies.map((sub, idx) => {
        const val = parseInt(sliderInputs[idx].value);
        totalProgress += val;
        return {
            ...sub,
            progress: val
        };
    });
    
    const avgProgress = Math.round(totalProgress / updatedSubs.length);
    
    // Update finished goods cumulative quantity
    const qtyDoneInput = document.getElementById('input-qty-done');
    const qtyDoneSubmitted = parseInt(qtyDoneInput.value) || 0;
    const newQtyDone = (project.qty_done || 0) + qtyDoneSubmitted;
    
    // Determine status
    const issueSelected = selectIssue.value;
    let newStatus = 'ontime';
    if (avgProgress >= 100) {
        newStatus = 'completed';
    } else if (issueSelected !== 'none') {
        newStatus = 'delayed';
    }
    
    const now = new Date();
    const isoTimestamp = now.toISOString();
    const formattedDate = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    
    let noteText = textareaNote.value.trim();
    if (issueSelected !== 'none') {
        noteText = `[ปัญหา: ${selectIssue.options[selectIssue.selectedIndex].text}] ${noteText}`;
    }
    
    if (isCloudMode) {
        // --- CLOUD MODE SAVE ---
        try {
            // Update project
            const { error: projErr } = await supabaseClient
                .from('projects')
                .update({
                    progress: avgProgress,
                    qty_done: newQtyDone,
                    status: newStatus,
                    sub_assemblies: updatedSubs,
                    updated_at: isoTimestamp
                })
                .eq('id', project.id);
                
            if (projErr) throw projErr;
            
            // Insert update log
            const { error: logErr } = await supabaseClient
                .from('updates')
                .insert({
                    timestamp: isoTimestamp,
                    reporter: operator,
                    customer: project.customer,
                    project_code: project.project_code,
                    job_number: project.job_no,
                    progress: avgProgress,
                    qty_done: qtyDoneSubmitted,
                    issue: issueSelected === 'none' ? 'normal' : selectIssue.options[selectIssue.selectedIndex].text,
                    note: noteText
                });
                
            if (logErr) throw logErr;
            
            // Send to Google Sheets Webhook
            if (googleSheetUrl) {
                const sheetPayload = {
                    timestamp: formattedDate,
                    reporter: operator,
                    customer: project.customer,
                    project_code: project.project_code,
                    job_number: project.job_no,
                    progress: avgProgress,
                    qty_done: qtyDoneSubmitted,
                    issue: issueSelected === 'none' ? '🟢 ปกติ' : `🔴 ปัญหา: ${selectIssue.options[selectIssue.selectedIndex].text}`,
                    note: noteText || 'ไม่มีบันทึกเพิ่มเติม'
                };
                
                fetch(googleSheetUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(sheetPayload)
                }).catch(err => console.warn("Google Sheet sync warning:", err));
            }
            
        } catch (err) {
            alert(`❌ บันทึกคลาวด์ล้มเหลว: ${err.message}`);
            return;
        }
    } else {
        // --- LOCAL DEMO MODE SAVE ---
        project.progress = avgProgress;
        project.qty_done = newQtyDone;
        project.status = newStatus;
        project.sub_assemblies = updatedSubs;
        
        const newUpdate = {
            timestamp: formattedDate,
            reporter: operator,
            customer: project.customer,
            project_code: project.project_code,
            job_number: project.job_no,
            progress: avgProgress,
            qty_done: qtyDoneSubmitted,
            note: noteText
        };
        
        updates.unshift(newUpdate);
        
        localStorage.setItem('aveam_projects', JSON.stringify(projects));
        localStorage.setItem('aveam_updates', JSON.stringify(updates));
        
        // Sync UI manually in Local Mode
        updateStats();
        renderProjects();
        renderRecentUpdates();
    }
    
    // Show success overlay on phone screen
    mobileForm.style.display = 'none';
    successScreen.style.display = 'flex';
    
    // Resume auto rotation on update success
    resetCarousel();
    
    setTimeout(() => {
        cancelMobileUpdate();
    }, 2200);
}

// Reset LocalStorage
function resetDemoData() {
    if (isCloudMode) {
        alert("⚠️ การล้างข้อมูลเดโมถูกปิดใช้งานในโหมดระบบคลาวด์ หากต้องการรีเซ็ตข้อมูลโครงการ โปรดอัปเดตข้อมูลใน Excel แล้วรันสคริปต์ Python Uploader หรือเข้าไปลบตารางใน Supabase");
        return;
    }
    if (confirm("คุณต้องการล้างข้อมูลและใช้ข้อมูลตั้งต้นจากตาราง Excel ใช่หรือไม่?")) {
        localStorage.removeItem('aveam_projects');
        localStorage.removeItem('aveam_updates');
        projects = JSON.parse(JSON.stringify(EXCEL_PROJECTS));
        updates = JSON.parse(JSON.stringify(INITIAL_UPDATES));
        
        if (projects.length > 0) {
            selectedProjectId = projects[0].id;
        } else {
            selectedProjectId = null;
        }
        
        cancelMobileUpdate();
        updateStats();
        renderProjects();
        renderRecentUpdates();
        resetCarousel();
    }
}

// Map projects to their high-fidelity images inside Asset/Products-RBG/
function getProductImage(project) {
    if (!project) return null;
    const code = project.project_code.toUpperCase();
    const customer = project.customer.toUpperCase();
    const pn = (project.part_number || '').toUpperCase();
    
    if (customer === 'VEECO') {
        if (code === 'ACBOX') return 'Asset/Products-RBG/Veeco-AC POWER BOX.png';
        if (code === 'WF') return 'Asset/Products-RBG/Veeco-WF.png';
        if (code.includes('AMP') || code.includes('AMPLIFIER')) return 'Asset/Products-RBG/Veeco-Air Gauge Amplifier.png';
        if (code.includes('SENSOR')) return 'Asset/Products-RBG/Veeco-Air Gauge Sensor.png';
    } else if (customer === 'UIC') {
        if (code.includes('FRAME') || code === 'LH & RH FRAME') return 'Asset/Products-RBG/UIC-LH and RH Frame.png';
        if (code.includes('BANK') || code.includes('FEEDER BANK')) return 'Asset/Products-RBG/UIC-Feeder Bank Assembly.png';
        if (code.includes('BASKET')) return 'Asset/Products-RBG/UIC-Feeder Basket.png';
        if (code.includes('TRANSFER') || code.includes('CART')) return 'Asset/Products-RBG/UIC-Feeder Transfer Cart.png';
        if (code.includes('STA') || code.includes('MATRIX') || code.includes('TRAY')) return 'Asset/Products-RBG/UIC-STA, MATRIX, TRAY.png';
        if (code.includes('BIN') || code.includes('SCRAP BIN')) return 'Asset/Products-RBG/UIC-Scrap Bin Assembly.png';
        if (code.includes('CUTTER') || code.includes('DRIVE') || code.includes('SCRAP TAPE')) return 'Asset/Products-RBG/UIC-Scrap Tape Cutter Drive Assy.png';
    } else if (customer === 'COHU') {
        if (code.includes('7001') || code.includes('JLP') || pn.endsWith('7001')) return 'Asset/Products-RBG/COHU-7001.png';
        if (code.includes('8001') || pn.endsWith('8001')) return 'Asset/Products-RBG/COHU-8001.png';
        if (code.includes('900') || code.includes('900X')) return 'Asset/Products-RBG/COHU-900X.png';
        if (code.includes('ZLINK') || code.includes('Z-LINK')) return 'Asset/Products-RBG/COHU-ZLINK.png';
    }
    return null;
}

// Dynamic Color Generator based on percentage (Red -> Orange -> Yellow -> Green)
function getProgressColor(percent) {
    // Hue ranges from 0 (Red) to 120 (Green)
    const hue = (percent / 100) * 120;
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
    
    if (currentTheme === 'dark') {
        // High vibrancy for dark background
        return `hsl(${hue}, 95%, 48%)`;
    } else {
        // Deeper shade with higher contrast for light/green backgrounds
        return `hsl(${hue}, 85%, 38%)`;
    }
}

// Theme Selector Logic
function setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('aveam_theme', theme);
    updateThemeSelector(theme);
    renderProjects(); // Trigger recalculation of colors for the new theme
}

function updateThemeSelector(theme) {
    const group = document.querySelector('.theme-selector-group');
    if (!group) return;
    
    // Remove active class from all buttons
    const buttons = group.querySelectorAll('.theme-select-btn');
    buttons.forEach(btn => btn.classList.remove('active'));
    
    // Add active class to selected theme button
    const activeBtn = group.querySelector(`.theme-select-btn.${theme}`);
    if (activeBtn) {
        activeBtn.classList.add('active');
    }
}

// ==========================================================================
// 100% COMPLETE CELEBRATION (SPARKLING STARS & FIREWORKS CANVAS ANIMATION)
// ==========================================================================
let celebrationAnimationId = null;

function initCelebration() {
    stopCelebration(); // Clean up any existing loop
    
    const canvas = document.getElementById('celebration-canvas');
    if (!canvas) return;
    
    const ctx = canvas.getContext('2d');
    const width = canvas.width;
    const height = canvas.height;
    const centerX = width / 2;
    const centerY = height / 2;
    
    let particles = [];
    let stars = [];
    
    function random(min, max) {
        return Math.random() * (max - min) + min;
    }
    
    function drawFourPointStar(c, x, y, size, rotation, opacity, color) {
        c.save();
        c.translate(x, y);
        c.rotate(rotation);
        c.globalAlpha = opacity;
        
        c.beginPath();
        c.moveTo(0, -size);
        c.quadraticCurveTo(0, 0, size, 0);
        c.quadraticCurveTo(0, 0, 0, size);
        c.quadraticCurveTo(0, 0, -size, 0);
        c.quadraticCurveTo(0, 0, 0, -size);
        c.closePath();
        
        c.fillStyle = color;
        c.shadowBlur = 6;
        c.shadowColor = color;
        c.fill();
        c.restore();
    }
    
    function spawnFirework() {
        const fx = centerX + random(-30, 30);
        const fy = centerY + random(-30, 30);
        const hue = random(0, 360);
        const numParticles = Math.floor(random(18, 30));
        
        for (let i = 0; i < numParticles; i++) {
            const angle = random(0, Math.PI * 2);
            const speed = random(1.5, 4.5);
            particles.push({
                x: fx,
                y: fy,
                vx: Math.cos(angle) * speed,
                vy: Math.sin(angle) * speed - 0.5,
                hue: hue + random(-25, 25),
                alpha: 1,
                decay: random(0.015, 0.035),
                size: random(1.5, 3.5),
                gravity: 0.06
            });
        }
    }
    
    function spawnStar() {
        const angle = random(0, Math.PI * 2);
        // radius between 35px and 95px from center to stay within/around the donut ring
        const radius = random(35, 95);
        stars.push({
            x: centerX + Math.cos(angle) * radius,
            y: centerY + Math.sin(angle) * radius,
            size: random(3.5, 7.5),
            rotation: random(0, Math.PI),
            rotationSpeed: random(-0.02, 0.02),
            opacity: 0,
            fadeSpeed: random(0.015, 0.035),
            phase: 'in',
            color: `hsl(${random(35, 65)}, 100%, ${random(75, 95)}%)` // golden/yellow glitter
        });
    }
    
    let lastFireworkTime = 0;
    let lastStarTime = 0;
    
    function animate(timestamp) {
        if (!document.getElementById('celebration-canvas')) {
            return; // Stop loop if canvas is gone
        }
        
        ctx.clearRect(0, 0, width, height);
        
        // 1. Spawning timers
        if (!lastFireworkTime) lastFireworkTime = timestamp;
        if (!lastStarTime) lastStarTime = timestamp;
        
        if (timestamp - lastFireworkTime > random(700, 1300)) {
            spawnFirework();
            lastFireworkTime = timestamp;
        }
        
        if (timestamp - lastStarTime > random(120, 250)) {
            spawnStar();
            lastStarTime = timestamp;
        }
        
        // 2. Update and draw stars
        for (let i = stars.length - 1; i >= 0; i--) {
            const s = stars[i];
            s.rotation += s.rotationSpeed;
            
            if (s.phase === 'in') {
                s.opacity += s.fadeSpeed;
                if (s.opacity >= 1) {
                    s.opacity = 1;
                    s.phase = 'out';
                }
            } else {
                s.opacity -= s.fadeSpeed;
            }
            
            if (s.opacity <= 0) {
                stars.splice(i, 1);
                continue;
            }
            
            drawFourPointStar(ctx, s.x, s.y, s.size, s.rotation, s.opacity, s.color);
        }
        
        // 3. Update and draw firework sparks
        for (let i = particles.length - 1; i >= 0; i--) {
            const p = particles[i];
            p.x += p.vx;
            p.y += p.vy;
            p.vy += p.gravity;
            p.alpha -= p.decay;
            
            if (p.alpha <= 0) {
                particles.splice(i, 1);
                continue;
            }
            
            ctx.save();
            ctx.globalAlpha = p.alpha;
            ctx.fillStyle = `hsl(${p.hue}, 100%, 65%)`;
            ctx.shadowBlur = 5;
            ctx.shadowColor = `hsl(${p.hue}, 100%, 55%)`;
            
            ctx.beginPath();
            ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
            ctx.fill();
            ctx.restore();
        }
        
        celebrationAnimationId = requestAnimationFrame(animate);
    }
    
    // Spawn initial items
    spawnFirework();
    for (let i = 0; i < 6; i++) {
        spawnStar();
        stars[i].opacity = random(0.1, 0.7);
        stars[i].phase = Math.random() > 0.5 ? 'in' : 'out';
    }
    
    celebrationAnimationId = requestAnimationFrame(animate);
}

function stopCelebration() {
    if (celebrationAnimationId) {
        cancelAnimationFrame(celebrationAnimationId);
        celebrationAnimationId = null;
    }
}

// Initialize on page load
window.onload = init;
