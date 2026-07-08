// AVEAM Mobile Assembly Report Portal Logic (Supabase Enabled)

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
let selectedProject = null;
let googleSheetUrl = "";
let supabaseClient = null;
let employeesList = [];

// DOM Elements
const loadingState = document.getElementById('loading-state');
const selectorState = document.getElementById('job-selector-state');
const mobileForm = document.getElementById('mobile-form');
const successScreen = document.getElementById('success-screen');

const formProjectTitle = document.getElementById('form-project-title');
const formJobNo = document.getElementById('form-job-no');
const formMcNumber = document.getElementById('form-mc-number');
const formBuildLeader = document.getElementById('form-build-leader');

const selectOperator = document.getElementById('select-operator');
const selectJobManual = document.getElementById('select-job-manual');
const subSlidersContainer = document.getElementById('sub-sliders-container');
const inputQtyDone = document.getElementById('input-qty-done');
const selectIssue = document.getElementById('select-issue');
const textareaNote = document.getElementById('textarea-note');

// Initialize Mobile App
async function init() {
    // 1. Fetch Supabase & Google configuration keys from server
    try {
        const response = await fetch('supabase_config.json');
        if (!response.ok) throw new Error('Cannot load supabase_config.json');
        const config = await response.json();
        
        const url = config.supabase_url;
        const key = config.supabase_key;
        googleSheetUrl = config.google_sheet_url || "";
        
        if (!url || !key) throw new Error('Supabase configuration is incomplete');
        
        // Initialize Supabase Client
        const { createClient } = supabase;
        supabaseClient = createClient(url, key);
        
        // 2. Fetch active projects list from Supabase
        const { data: projData, error: projErr } = await supabaseClient
            .from('projects')
            .select('*')
            .order('customer', { ascending: true });
            
        if (projErr) throw projErr;
        projects = projData || [];

        // Fetch employees list from Supabase
        const { data: empData, error: empErr } = await supabaseClient
            .from('employees')
            .select('*')
            .order('name', { ascending: true });
            
        if (empErr) throw empErr;
        employeesList = empData || [];
        populateOperators(employeesList);
        
        // Hide loading
        loadingState.style.display = 'none';
        
        // 3. Check for URL parameter '?job=JOB_NUMBER'
        const urlParams = new URLSearchParams(window.location.search);
        const jobParam = urlParams.get('job');
        
        if (jobParam) {
            // Try to match job
            const matched = projects.find(p => p.id === jobParam || p.job_no === jobParam || (p.job_no && p.job_no.includes(jobParam)));
            if (matched) {
                loadJobForm(matched);
            } else {
                alert(`⚠️ ไม่พบจ๊อบเลขที่ "${jobParam}" ที่กำลังผลิตอยู่ในระบบ`);
                showManualSelector();
            }
        } else {
            showManualSelector();
        }
        
    } catch (err) {
        console.error("Mobile Portal Initialization Error:", err);
        loadingState.innerHTML = `
            <div style="text-align: center; color: var(--text-delayed); padding: 2rem;">
                <span style="font-size: 1.25rem; font-weight: 700; display: block;">⚠️ เกิดข้อผิดพลาดในการเชื่อมต่อ</span>
                <span style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.5rem; display: block;">(${err.message})</span>
            </div>
        `;
    }
}

function populateOperators(list) {
    selectOperator.innerHTML = '<option value="" disabled selected>-- เลือกผู้รายงาน --</option>';
    list.forEach(emp => {
        const opt = document.createElement('option');
        opt.value = emp.id;
        opt.textContent = emp.name;
        selectOperator.appendChild(opt);
    });
}

function showManualSelector() {
    selectorState.style.display = 'block';
    mobileForm.style.display = 'none';
    successScreen.style.display = 'none';
    
    // Fill job dropdown select list
    selectJobManual.innerHTML = '<option value="" disabled selected>-- เลือกโครงการ/จ๊อบที่ต้องการอัปเดต --</option>';
    projects.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.id;
        opt.textContent = `[${p.customer}] ${p.project_code} (Job: ${p.job_no})`;
        selectJobManual.appendChild(opt);
    });
}

function handleManualJobSelect() {
    const selectedId = selectJobManual.value;
    const project = projects.find(p => p.id === selectedId);
    if (project) {
        loadJobForm(project);
    }
}

// Populate and show the reporting form
function loadJobForm(project) {
    selectedProject = project;
    
    selectorState.style.display = 'none';
    mobileForm.style.display = 'block';
    successScreen.style.display = 'none';
    
    // Set text labels
    formProjectTitle.textContent = `${project.customer} - ${project.project_code} (PN: ${project.part_number || '-'})`;
    formJobNo.textContent = project.job_no;
    formMcNumber.textContent = project.mc_number || '-';
    
    // Build Leader - show first name only
    const fullLeader = project.team_leader || '-';
    formBuildLeader.textContent = (fullLeader !== '-' && fullLeader !== 'Unassigned') ? fullLeader.trim().split(/\s+/)[0] : '-';
    
    // Populate operator selection if it was already selected previously
    selectOperator.value = '';
    
    // Populate sub-assembly progress sliders
    subSlidersContainer.innerHTML = '';
    const subs = project.sub_assemblies || [];
    subs.forEach((sub, idx) => {
        const div = document.createElement('div');
        div.className = 'slider-group';
        div.innerHTML = `
            <div class="slider-meta">
                <span class="slider-label" title="${sub.name}">${sub.pn}: ${sub.th_desc || sub.name}</span>
                <span class="slider-val" id="slide-val-${idx}">${sub.progress}%</span>
            </div>
            <input type="range" class="slider-input" min="0" max="100" step="5" 
                value="${sub.progress}" 
                oninput="document.getElementById('slide-val-${idx}').textContent = this.value + '%'">
        `;
        subSlidersContainer.appendChild(div);
    });
    
    // Reset standard form inputs
    inputQtyDone.value = '0';
    selectIssue.value = 'none';
    textareaNote.value = '';
}

// Handle Form Submission
async function submitProgressUpdate(event) {
    event.preventDefault();
    if (!selectedProject || !supabaseClient) return;
    
    const selectedEmpId = selectOperator.value;
    if (!selectedEmpId) {
        alert("กรุณาเลือกชื่อผู้รายงาน");
        return;
    }

    const enteredEmpId = document.getElementById('input-employee-id').value.trim();
    if (!enteredEmpId) {
        alert("กรุณากรอกรหัสพนักงาน");
        return;
    }

    if (selectedEmpId !== enteredEmpId) {
        alert("❌ รหัสพนักงานไม่ถูกต้อง กรุณาตรวจสอบและลองใหม่อีกครั้ง");
        return;
    }

    const operatorName = selectOperator.options[selectOperator.selectedIndex].text;
    
    const loadingBadge = document.createElement('div');
    loadingBadge.className = 'time-badge';
    loadingBadge.style.cssText = 'position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); z-index: 1000; border: none; background: rgba(15, 21, 36, 0.95); padding: 1.5rem;';
    loadingBadge.innerHTML = '<span class="live-dot" style="animation: pulse 1s infinite;"></span> กำลังบันทึกข้อมูลและส่งอีเมลแจ้งเตือน...';
    document.body.appendChild(loadingBadge);
    
    // 1. Calculate new progress based on sub-assembly sliders
    const sliderInputs = subSlidersContainer.querySelectorAll('.slider-input');
    let totalProgress = 0;
    const updatedSubs = selectedProject.sub_assemblies.map((sub, idx) => {
        const val = parseInt(sliderInputs[idx].value);
        totalProgress += val;
        return {
            ...sub,
            progress: val
        };
    });
    
    const avgProgress = Math.round(totalProgress / updatedSubs.length);
    
    // 2. Accumulate Finished Goods quantity
    const qtyDoneSubmitted = parseInt(inputQtyDone.value) || 0;
    const newQtyDone = (selectedProject.qty_done || 0) + qtyDoneSubmitted;
    
    // 3. Determine status
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
    
    try {
        // --- STEP A: Update projects table in Supabase ---
        const { error: projErr } = await supabaseClient
            .from('projects')
            .update({
                progress: avgProgress,
                qty_done: newQtyDone,
                status: newStatus,
                sub_assemblies: updatedSubs,
                updated_at: isoTimestamp
            })
            .eq('id', selectedProject.id);
            
        if (projErr) throw projErr;
        
        // --- STEP B: Insert log row in updates table in Supabase ---
        const { error: logErr } = await supabaseClient
            .from('updates')
            .insert({
                timestamp: isoTimestamp,
                reporter: operatorName,
                customer: selectedProject.customer,
                project_code: selectedProject.project_code,
                job_number: selectedProject.job_no,
                progress: avgProgress,
                qty_done: qtyDoneSubmitted,
                issue: issueSelected === 'none' ? 'normal' : selectIssue.options[selectIssue.selectedIndex].text,
                note: noteText
            });
            
        if (logErr) throw logErr;
        
        // --- STEP C: Send log to Google Sheets Webhook ---
        if (googleSheetUrl) {
            const sheetPayload = {
                timestamp: formattedDate,
                reporter: operatorName,
                customer: selectedProject.customer,
                project_code: selectedProject.project_code,
                job_number: selectedProject.job_no,
                progress: avgProgress,
                qty_done: qtyDoneSubmitted,
                issue: issueSelected === 'none' ? '🟢 ปกติ' : `🔴 ปัญหา: ${selectIssue.options[selectIssue.selectedIndex].text}`,
                note: noteText || 'ไม่มีบันทึกเพิ่มเติม'
            };
            
            // Post via fetch in no-cors mode to avoid CORS blocking since Google Apps Script accepts it
            fetch(googleSheetUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(sheetPayload)
            }).then(() => {
                console.log("Data synced successfully with Google Sheets.");
            }).catch(e => {
                console.warn("Google Sheets Sync Warning:", e);
            });
        }
        
        // Hide loading badge
        document.body.removeChild(loadingBadge);
        
        // Show success screen
        mobileForm.style.display = 'none';
        successScreen.style.display = 'flex';
        
    } catch (e) {
        document.body.removeChild(loadingBadge);
        alert(`❌ ไม่สามารถบันทึกข้อมูลได้: ${e.message}`);
        console.error("Supabase Save Error:", e);
    }
}

function resetFormState() {
    selectedProject = null;
    showManualSelector();
}

// Start app
window.onload = init;
