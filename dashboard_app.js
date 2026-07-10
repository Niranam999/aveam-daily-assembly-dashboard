// AVEAM Shop Floor Assembly Real-Time Dashboard Logic (Supabase Enabled)

// Constants
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

// App State
let projects = [];
let updates = [];
let currentFilter = 'ALL';
let selectedProjectId = null;
let supabaseClient = null;

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

// Initialize Dashboard
async function init() {
    // 1. Initialize Theme
    const savedTheme = localStorage.getItem('aveam_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeSelector(savedTheme);

    updateSystemTime();
    setInterval(updateSystemTime, 1000);
    
    // 2. Fetch Supabase configuration keys from server
    try {
        const response = await fetch('supabase_config.json');
        if (!response.ok) throw new Error('Cannot load supabase_config.json');
        const config = await response.json();
        
        const url = config.supabase_url;
        const key = config.supabase_key;
        
        if (!url || !key) throw new Error('Supabase URL or Key is missing in config file');
        
        // Initialize Supabase Client
        const { createClient } = supabase;
        supabaseClient = createClient(url, key);
        console.log("Supabase Client initialized successfully.");
        
        // Set Master QR Code image source dynamically based on current origin
        const currentPath = window.location.pathname;
        const dirPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
        const serverOrigin = `${window.location.protocol}//${window.location.host}${dirPath}`;
        const masterQrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=110x110&data=${encodeURIComponent(serverOrigin + '/update.html')}`;
        const masterQrImg = document.getElementById('master-qr-img');
        if (masterQrImg) {
            masterQrImg.src = masterQrUrl;
        }
        
        // 3. Load Initial Data
        await loadData();
        
        // 4. Setup Real-time Listeners
        setupRealtimeSubscriptions();
        
        // 5. Start Carousel
        startCarousel();
        
        // 6. Adjust scale for TV screen
        adjustDashboardScale();
        window.addEventListener('resize', adjustDashboardScale);
        
    } catch (err) {
        console.error("Dashboard Initialization Error:", err);
        selectedJobPanel.innerHTML = `
            <div style="grid-column: 1 / span 4; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--text-delayed); font-weight: 600;">
                <span>⚠️ ไม่สามารถเชื่อมต่อกับฐานข้อมูลคลาวด์ได้</span>
                <span style="font-size: 0.8rem; font-weight: 400; color: var(--text-muted); margin-top: 0.25rem;">โปรดตรวจสอบไฟล์ supabase_config.json หรือการตั้งค่าอินเทอร์เน็ต (${err.message})</span>
            </div>
        `;
    }
}

// Fetch projects and updates logs from Supabase
async function loadData() {
    if (!supabaseClient) return;
    
    try {
        // Fetch Projects (sort by customer)
        const { data: projData, error: projErr } = await supabaseClient
            .from('projects')
            .select('*')
            .order('customer', { ascending: true })
            .order('project_code', { ascending: true });
            
        if (projErr) throw projErr;
        projects = projData || [];
        
        // Fetch recent updates
        const { data: updateData, error: updateErr } = await supabaseClient
            .from('updates')
            .select('*')
            .order('timestamp', { ascending: false })
            .limit(10);
            
        if (updateErr) throw updateErr;
        updates = updateData || [];
        
        console.log(`Loaded ${projects.length} projects and ${updates.length} logs.`);
        
        // Select first project by default if none selected or if previous selected is gone
        const filtered = getFilteredProjects();
        if (filtered.length > 0) {
            const stillExists = filtered.some(p => p.id === selectedProjectId);
            if (!stillExists) {
                selectedProjectId = filtered[0].id;
            }
        } else {
            selectedProjectId = null;
        }
        
        // Sync UI
        renderFilters();
        updateStats();
        renderProjects();
        renderRecentUpdates();
        
    } catch (e) {
        console.error("Error loading Supabase data:", e);
    }
}

// Listen to database inserts and updates via WebSockets
function setupRealtimeSubscriptions() {
    if (!supabaseClient) return;
    
    // Subscribe to projects table changes
    supabaseClient.channel('realtime-projects')
        .on('postgres_changes', { event: '*', schema: 'public', table: 'projects' }, (payload) => {
            console.log('Project table change detected:', payload);
            
            const eventType = payload.eventType;
            const newRecord = payload.new;
            const oldRecord = payload.old;
            
            if (eventType === 'INSERT') {
                projects.push(newRecord);
            } else if (eventType === 'UPDATE') {
                const idx = projects.findIndex(p => p.id === newRecord.id);
                if (idx !== -1) {
                    projects[idx] = newRecord;
                }
            } else if (eventType === 'DELETE') {
                projects = projects.filter(p => p.id !== oldRecord.id);
            }
            
            // Re-render
            updateStats();
            renderProjects();
            
            // If the currently selected project was updated, re-render the detail panel
            if (selectedProjectId && (eventType === 'UPDATE' && newRecord.id === selectedProjectId)) {
                renderDetailPanel(newRecord);
            }
        })
        .subscribe();
        
    // Subscribe to updates table inserts (logs)
    supabaseClient.channel('realtime-logs')
        .on('postgres_changes', { event: 'INSERT', schema: 'public', table: 'updates' }, (payload) => {
            console.log('New update log detected:', payload);
            
            // Add new log to start
            updates.unshift(payload.new);
            if (updates.length > 10) {
                updates.pop(); // keep only 10
            }
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

function updateStats() {
    let backlog = 0;
    let assigned = 0;
    let inProgress = 0;
    let qa = 0;
    let completed = 0;
    
    projects.forEach(p => {
        if ((p.kanban_backlog || 0) > 0) backlog += 1;
        if ((p.kanban_assigned || 0) > 0) assigned += 1;
        if ((p.kanban_in_progress || 0) > 0) inProgress += 1;
        if ((p.kanban_qa || 0) > 0) qa += 1;
        if ((p.kanban_completed || 0) > 0) completed += 1;
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

    const circumference = 534;
    const offset = circumference - (project.progress / 100) * circumference;

    const numSubs = project.sub_assemblies ? project.sub_assemblies.length : 0;
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
    if (project.sub_assemblies) {
        project.sub_assemblies.forEach(sub => {
            const barColor = getProgressColor(sub.progress);
            
            // Extract a short English description (up to 20 chars) from sub.name
            let engDesc = '';
            if (sub.name) {
                const parts = sub.name.split(';');
                let cleaned = sub.name;
                if (parts.length > 1 && /^\d[0-9A-Z\-\.\/\s]*$/.test(parts[0].trim().toUpperCase())) {
                    cleaned = parts.slice(1).join(';').trim();
                }
                cleaned = cleaned.replace(/^[\s;]+/, '').trim();
                if (cleaned.length > 20) {
                    cleaned = cleaned.substring(0, 20) + '..';
                }
                engDesc = cleaned;
            }
            
            const fullLabel = `${sub.pn}${sub.th_desc ? ' ' + sub.th_desc : ''}${engDesc ? ' - ' + engDesc : ''}`;
            
            barsHtml += `
                <div class="vertical-bar-item" data-tooltip="${sub.pn}: ${sub.name} (${sub.progress}%)" style="width: ${itemWidth}; flex-shrink: 1;">
                    <span class="vertical-bar-value" style="font-size: ${valueFontSize}; font-weight: 700; margin-bottom: 0.25rem; color: ${sub.progress > 0 ? barColor : 'var(--text-muted)'}">${sub.progress}%</span>
                    <div class="vertical-bar-track" style="width: ${trackWidth};">
                        <div class="vertical-bar-fill" style="height: ${sub.progress}%; background-color: ${barColor}"></div>
                    </div>
                    <span class="vertical-bar-label" style="font-size: ${labelFontSize}; transform: rotate(${labelRotation});">${fullLabel}</span>
                </div>
            `;
        });
    }

    // Product Image Column HTML (Column 2)
    const imgPath = getProductImage(project);
    let productColHtml = '';
    if (imgPath) {
        productColHtml = `
            <!-- Column 2: Product Image -->
            <div class="detail-product-col" style="display: flex; flex-direction: column; align-items: center; border-right: 1px solid var(--card-border); padding-right: 1rem; height: 100%;">
                <span style="font-size: 0.8rem; font-weight: 700; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 0.25rem; flex-shrink: 0;">Product Picture</span>
                <div class="product-col-image-box" style="flex: 1; display: flex; align-items: center; justify-content: center; margin-top: 0.5rem; width: 100%; overflow: hidden; box-sizing: border-box;">
                    <img src="${imgPath}" alt="${project.project_code}" style="width: 100%; max-width: 100%; max-height: 220px; object-fit: contain; border-radius: 4px;">
                </div>
            </div>
        `;
    } else {
        productColHtml = `
            <!-- Column 2: Product Image Placeholder -->
            <div class="detail-product-col" style="display: flex; flex-direction: column; align-items: center; border-right: 1px solid var(--card-border); padding-right: 1rem; height: 100%;">
                <span style="font-size: 0.8rem; font-weight: 700; color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 0.25rem; flex-shrink: 0;">Product Picture</span>
                <div class="product-col-image-box placeholder-box" style="flex: 1; width: 100%; display: flex; flex-direction: column; align-items: center; justify-content: center; background: var(--circle-bg); border: 1px dashed var(--card-border); border-radius: 8px; margin-top: 0.5rem; width: 100%;">
                    <span style="font-size: 0.7rem; color: var(--text-muted);">ไม่มีรูปภาพสินค้า</span>
                </div>
            </div>
        `;
    }

    // Calculate FG percentage
    const qtyDone = project.qty_done || 0;
    const qtyTarget = project.qty || 1;
    const fgPercent = Math.min(100, Math.round((qtyDone / qtyTarget) * 100));

    // Dynamic Server IP for QR Code
    const currentPath = window.location.pathname;
    const dirPath = currentPath.substring(0, currentPath.lastIndexOf('/'));
    const serverOrigin = `${window.location.protocol}//${window.location.host}${dirPath}`;
    const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=110x110&data=${encodeURIComponent(serverOrigin + '/update.html?job=' + (project.job_no || project.id))}`;

    selectedJobPanel.innerHTML = `
        <!-- Column 1: Info & Horizontal FG Status -->
        <div class="detail-left" style="display: flex; flex-direction: column; justify-content: space-between; border-right: 1px solid var(--card-border); padding-right: 1rem; height: 100%;">
            <div>
                <span class="thumb-client ${project.customer.toLowerCase()}" style="font-size: 0.8rem; padding: 0.2rem 0.6rem;">${project.customer}</span>
                <h3 class="detail-project-name" style="white-space: nowrap; overflow: hidden; text-overflow: ellipsis; margin-top: 0.4rem;">${project.project_code}</h3>
                <p class="detail-job-number" style="line-height: 1.5; margin-top: 0.5rem;">
                    Part Number: <strong>${project.part_number || '-'}</strong><br>
                    ${project.job_no ? 'Job: <strong>' + project.job_no + '</strong>' : ''}${project.mc_number && project.mc_number !== '-' ? (project.job_no ? ' ' : '') + '(MC: <strong>' + project.mc_number + '</strong>)' : ''}${project.job_no || (project.mc_number && project.mc_number !== '-') ? '<br>' : ''}
                    Qty: <strong>${project.qty}</strong><br>
                    Time: <strong>${Number(project.est_hours).toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 })}</strong> hrs/unit (Total: <strong>${getProjectTotalHours(project).toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 1 })}</strong> hrs)
                </p>
                <div class="detail-desc-box" style="margin-top: 0.4rem;">
                    Description: <em>${project.description}</em>
                </div>
            </div>
            
            <!-- Bottom Horizontal FG Status Bar -->
            <div class="detail-fg-status-horizontal" style="display: flex; flex-direction: column; width: 100%; background: rgba(255,255,255,0.02); border-radius: 8px; padding: 0.5rem 0.75rem; margin-top: 0.5rem; box-sizing: border-box;">
                <div style="display: flex; justify-content: space-between; align-items: center;">
                    <span style="font-size: 0.75rem; color: #fb923c; font-weight: 700; text-transform: uppercase; letter-spacing: 0.75px;">FG Status</span>
                    <span style="font-size: 1.15rem; font-weight: 800; color: var(--text-title);">${qtyDone}/${qtyTarget}</span>
                </div>
                <div class="fg-bar-track-horizontal">
                    <div class="fg-bar-fill-horizontal" style="width: ${fgPercent}%;"></div>
                </div>
                <div style="display: flex; justify-content: flex-end;">
                    <span style="font-size: 0.85rem; font-weight: 700; color: #10b981;">${fgPercent}%</span>
                </div>
            </div>
        </div>
        
        ${productColHtml}
        
        <!-- Column 3: Donut Progress Chart -->
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
            <div class="detail-meta-text" style="margin-top: 1rem; gap: 0.25rem;">
                <span class="stat-sub" style="font-size: 0.95rem; font-weight: 600;">ความคืบหน้ารวมเฉลี่ย</span>
                <span class="detail-lead-label">Lead: <span>${project.team_leader}</span></span>
                <span class="detail-lead-label" style="margin-top: 2px;">แม่ทีม: <span>${project.member_1}</span></span>
            </div>
        </div>
        
        <!-- Column 4: Sub-Assemblies Vertical Bar Chart -->
        <div class="detail-right">
            <div class="detail-subs-header">พาร์ทประกอบย่อยและสถานะรายตัว (${numSubs})</div>
            <div class="detail-vertical-chart-container" style="gap: ${containerGap}; justify-content: ${numSubs > 6 ? 'space-between' : 'flex-start'};">
                ${barsHtml}
            </div>
        </div>

        <!-- Column 5: Far Right QR & Action -->
        <div class="detail-qr-side">
            <div class="qr-title-box">
                <strong>SCAN JOB TRAVELER</strong>
                สแกน QR Code นี้ด้วยมือถือพนักงานเพื่อบันทึก Man-Hours และอัปเดตสถานะการประกอบ
            </div>
            <div class="big-qr-box" title="สแกนเพื่อทำงานประกอบ">
                <img src="${qrUrl}" alt="Job QR">
            </div>
            <a class="submit-btn" href="${serverOrigin}/update.html?job=${project.job_no || project.id}" target="_blank" style="width: 100%; font-size: 0.72rem; height: 32px; display: flex; align-items: center; justify-content: center; text-decoration: none;">
                🔍 เปิดหน้าจอรายงานจ๊อบนี้
            </a>
        </div>
    `;

    // Trigger celebration if progress is 100%
    if (project.progress === 100) {
        setTimeout(initCelebration, 50);
    } else {
        stopCelebration();
    }
    
    // Ensure scaling fits after rendering
    if (typeof adjustDashboardScale === 'function') {
        adjustDashboardScale();
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
                <span class="thumb-job">${p.job_no ? 'Job: ' + p.job_no : ''}${p.mc_number && p.mc_number !== '-' ? (p.job_no ? ' ' : '') + '(' + p.mc_number + ')' : ''}</span>
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

function selectProject(projectId) {
    selectedProjectId = projectId;
    
    const cards = projectsGrid.querySelectorAll('.thumbnail-card');
    cards.forEach(c => c.classList.remove('active'));
    
    const targetCard = document.getElementById(`card-${projectId}`);
    if (targetCard) {
        targetCard.classList.add('active');
    }
    
    const project = projects.find(p => p.id === projectId);
    renderDetailPanel(project);
}

function renderRecentUpdates() {
    recentList.innerHTML = '';
    
    if (updates.length === 0) {
        recentList.innerHTML = '<div style="color: var(--text-muted); text-align: center; padding: 1rem; font-size: 0.75rem;">ไม่มีบันทึกประวัติย้อนหลัง</div>';
        return;
    }
    
    // Show only 3 items on mobile/tablet (width <= 1024px or height <= 600px), otherwise 10
    const isMobileOrTablet = window.innerWidth <= 1024 || window.innerHeight <= 600;
    const limit = isMobileOrTablet ? 3 : 10;
    
    updates.slice(0, limit).forEach(u => {
        // Format timestamp cleanly
        let cleanTime = u.timestamp;
        if (cleanTime && cleanTime.includes('T')) {
            const dt = new Date(cleanTime);
            cleanTime = `${String(dt.getDate()).padStart(2, '0')}/${String(dt.getMonth() + 1).padStart(2, '0')}/${dt.getFullYear()} ${String(dt.getHours()).padStart(2, '0')}:${String(dt.getMinutes()).padStart(2, '0')}`;
        }
        
        const div = document.createElement('div');
        div.className = 'recent-item';
        div.innerHTML = `
            <div class="recent-item-header">
                <span>${u.project_code} (Job: ${u.job_number})</span>
                <span style="color: var(--brand-color);">${u.progress}%</span>
            </div>
            <div class="recent-item-meta">
                โดย: <strong>${u.reporter}</strong> ${u.qty_done ? `| FG: <strong>+${u.qty_done}</strong>` : ''} | ${cleanTime}
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
    
    const targetCard = document.getElementById(`card-${filtered[nextIndex].id}`);
    if (targetCard) {
        targetCard.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
}

function pauseAutoRotation() {
    isAutoRotating = false;
    carouselBar.style.width = '0%';
    
    if (pauseTimeout) clearTimeout(pauseTimeout);
    
    pauseTimeout = setTimeout(() => {
        isAutoRotating = true;
        progressPercent = 0;
    }, 60000); // Resume auto-rotation after 60s of inactivity
}

function resetCarousel() {
    isAutoRotating = true;
    progressPercent = 0;
    if (pauseTimeout) clearTimeout(pauseTimeout);
}

function resetCarouselTimers() {
    if (progressTimer) clearInterval(progressTimer);
}

// Product images mapping
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
    } else if (customer === 'ULC') {
        return 'Asset/Products-RBG/ULC.png';
    }
    return null;
}

// Progress Color Generator
function getProgressColor(percent) {
    const hue = (percent / 100) * 120;
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
    if (currentTheme === 'dark') {
        return `hsl(${hue}, 95%, 48%)`;
    } else {
        return `hsl(${hue}, 85%, 38%)`;
    }
}

// Theme controller
function setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('aveam_theme', theme);
    updateThemeSelector(theme);
    renderProjects(); 
}

function updateThemeSelector(theme) {
    const group = document.querySelector('.theme-selector-group');
    if (!group) return;
    const buttons = group.querySelectorAll('.theme-select-btn');
    buttons.forEach(btn => btn.classList.remove('active'));
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
    stopCelebration();
    
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
            color: `hsl(${random(35, 65)}, 100%, ${random(75, 95)}%)`
        });
    }
    
    let lastFireworkTime = 0;
    let lastStarTime = 0;
    
    function animate(timestamp) {
        if (!document.getElementById('celebration-canvas')) {
            return; 
        }
        
        ctx.clearRect(0, 0, width, height);
        
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

// Start app
window.onload = init;

// Re-render logs dynamically on window resize or orientation changes
window.addEventListener('resize', renderRecentUpdates);

// Automatically scale dashboard to fit widescreen low-height displays (like TCL 32")
function adjustDashboardScale() {
    const container = document.querySelector('.app-container');
    if (!container) return;
    
    // Only apply scale transform in widescreen landscape viewports
    if (window.innerWidth > 900 && window.innerWidth > window.innerHeight) {
        const targetHeight = 810; // Target layout height in pixels
        const scale = Math.min(1.0, window.innerHeight / targetHeight);
        
        if (scale < 1.0) {
            container.style.transform = `scale(${scale})`;
            container.style.transformOrigin = 'top center';
            container.style.width = `${100 / scale}%`;
            container.style.height = `${100 / scale}%`;
            document.body.style.overflow = 'hidden';
            return;
        }
    }
    
    // Reset properties if screen is portrait/tablet
    container.style.transform = '';
    container.style.transformOrigin = '';
    container.style.width = '';
    container.style.height = '';
    document.body.style.overflow = '';
}

// Scale again 300ms after load to guarantee scaling is correct
window.addEventListener('load', () => setTimeout(adjustDashboardScale, 300));
