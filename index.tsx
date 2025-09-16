// --- Type Definitions for external libraries ---
declare const XLSX: any;
declare const Chart: any;
declare const ChartDataLabels: any;
declare const jspdf: { jsPDF: any };
declare const html2canvas: any;
// Note: jspdf-autotable is a plugin, so it attaches to the jsPDF instance.
// We'll access it via `doc.autoTable` where `doc` is a jsPDF instance.

// --- DOM Elements ---
const fileUpload = document.getElementById('file-upload') as HTMLInputElement;
const dashboardGrid = document.getElementById('dashboard-grid') as HTMLElement;
const lastUpdate = document.getElementById('last-update') as HTMLElement;
const placeholder = document.getElementById('placeholder') as HTMLElement;
const filterContainer = document.getElementById('filter-container') as HTMLElement;
const chartsContainer = document.getElementById('charts-container') as HTMLElement;
const applyFiltersBtn = document.getElementById('apply-filters-btn') as HTMLButtonElement;
const resetFiltersBtn = document.getElementById('reset-filters-btn') as HTMLButtonElement;
const totalFclDisplay = document.getElementById('total-fcl-display') as HTMLElement;
const totalFclCount = document.getElementById('total-fcl-count') as HTMLElement;
const planConfigContainer = document.getElementById('plan-config-container') as HTMLElement;
const dailyCapacityInput = document.getElementById('daily-capacity-input') as HTMLInputElement;
const bufferConfigContainer = document.getElementById('buffer-config-container') as HTMLElement;
const bufferCapacityInput = document.getElementById('buffer-capacity-input') as HTMLInputElement;
const bufferSearchInput = document.getElementById('buffer-search-input') as HTMLInputElement;


// View Tabs
const viewTabsContainer = document.getElementById('view-tabs-container') as HTMLElement;

// Filter Inputs
const statusFilter = document.getElementById('status-filter') as HTMLSelectElement;
const finalStatusFilter = document.getElementById('final-status-filter') as HTMLSelectElement;
const cargoTypeFilter = document.getElementById('cargo-type-filter') as HTMLSelectElement;
const poFilter = document.getElementById('po-filter') as HTMLSelectElement;
const vesselFilter = document.getElementById('vessel-filter') as HTMLSelectElement;
const vesselSearchInput = document.getElementById('vessel-search-input') as HTMLInputElement;
const shipownerFilter = document.getElementById('shipowner-filter') as HTMLSelectElement;
const shipownerSearchInput = document.getElementById('shipowner-search-input') as HTMLInputElement;
const loadingTypeFilter = document.getElementById('loading-type-filter') as HTMLSelectElement;

// Export Buttons
const exportPdfBtn = document.getElementById('export-pdf-btn') as HTMLButtonElement;
const translateBtn = document.getElementById('translate-btn') as HTMLButtonElement;
const translateBtnText = document.getElementById('translate-btn-text') as HTMLSpanElement;
const clearDataBtn = document.getElementById('clear-data-btn') as HTMLButtonElement;
const themeToggleBtn = document.getElementById('theme-toggle-btn') as HTMLButtonElement;


// Loading Overlay
const loadingOverlay = document.getElementById('loading-overlay') as HTMLElement;

// Modal Elements
const detailsModal = document.getElementById('details-modal') as HTMLElement;
const modalContent = document.getElementById('modal-content') as HTMLElement;
const modalHeaderContent = document.getElementById('modal-header-content') as HTMLElement;
const modalBody = document.getElementById('modal-body') as HTMLElement;
const modalCloseBtn = document.getElementById('modal-close-btn') as HTMLButtonElement;

// Chart Titles
const barChartTitle = document.getElementById('bar-chart-title') as HTMLElement;

// Summary Section
const summaryContainer = document.getElementById('summary-container') as HTMLElement;
const highRiskCount = document.getElementById('high-risk-count') as HTMLElement;
const mediumRiskCount = document.getElementById('medium-risk-count') as HTMLElement;
const lowRiskCount = document.getElementById('low-risk-count') as HTMLElement;
const noneRiskCount = document.getElementById('none-risk-count') as HTMLElement;

// --- Global State ---
type RiskCategory = 'high' | 'medium' | 'low' | 'none';
type View = 'vessel' | 'po' | 'warehouse' | 'plan' | 'buffer';
let originalData: any[] = [];
let filteredDataCache: any[] = [];
let charts: { [key: string]: any } = {};
let currentView: View = 'vessel';
const TODAY = new Date();
let activeModalItem: any = null;
let modalSortState: { column: string | null; direction: 'asc' | 'desc' } = { column: 'daysToDeadline', direction: 'asc' };
let activeRiskFilter: RiskCategory | null = null;
let currentLanguage: 'pt' | 'zh' = 'pt';
let deliveryPlanCache: { [date: string]: any[] } = {};
let containersInBuffer: Set<string> = new Set();
const STORAGE_KEY = 'containerDashboardData';
const THEME_STORAGE_KEY = 'containerDashboardTheme';

// --- Translation Dictionary ---
const translations = {
    pt: {
        main_title: "DASHBOARD DE CONTROLE DE CONTÊINERES",
        upload_prompt_initial: "Carregue um arquivo .xlsx para começar",
        upload_prompt_success: "Dados de",
        upload_prompt_updated: "Atualizado em:",
        upload_prompt_loaded_from_storage: "Dados carregados do armazenamento local de",
        total_containers_label: "Total de Contêineres (Visão Atual):",
        export_overview_btn: "Exportar Visão Geral",
        upload_btn: "Carregar XLSX",
        clear_data_btn: "Limpar Dados",
        filter_po: "Filtrar POs",
        filter_vessel: "Filtrar Navios",
        vessel_search_placeholder: "Pesquisar...",
        filter_status: "Status do Contêiner",
        filter_final_status: "Status Final",
        filter_cargo_type: "Tipo de Mercadoria",
        filter_shipowner: "Armador (Shipowner)",
        filter_loading_type: "Tipo de Carregamento",
        filter_btn: "Filtrar",
        clear_btn: "Limpar",
        view_vessel: "Por Navio",
        view_po: "Por PO",
        view_warehouse: "Por Armazém",
        view_plan: "Plano de Entrega",
        view_buffer: "Buffer",
        loading_text: "Processando...",
        summary_high: "Alto Risco (Atrasado)",
        summary_medium: "Médio Risco (Vence ≤ 7d)",
        summary_low: "Baixo Risco (> 7d)",
        summary_none: "Entregue",
        chart_status_title: "Distribuição de Status",
        chart_deadline_title: "Distribuição de Prazos",
        placeholder_title: "Aguardando arquivo...",
        placeholder_subtitle: "Selecione uma planilha para iniciar a análise.",
        containers_unit: "Contêiner(es)",
        view_by_vessel: "Total de Contêineres por Navio",
        view_by_po: "Total de Contêineres por PO",
        view_by_warehouse: "Total de Contêineres por Armazém",
        plan_no_pending: "Nenhum contêiner pendente",
        plan_no_pending_subtitle: "Não há contêineres que necessitem de planejamento com os filtros atuais.",
        plan_daily_schedule_for: "Plano de Entrega para",
        plan_export_pdf: "Exportar PDF",
        plan_export_excel: "Exportar Excel",
        plan_pending_containers: "contêineres pendentes",
        plan_daily_capacity: "Contêineres por Dia:",
        modal_delivered_status: "Entregue",
        buffer_capacity: "Capacidade do Buffer:",
        buffer_total_capacity: "Capacidade Total",
        buffer_in_buffer: "No Buffer",
        buffer_available_space: "Espaço Disponível",
        buffer_at_facility: "Na Fábrica",
        buffer_facility_list_title: "Contêineres na Fábrica",
        buffer_buffer_list_title: "Contêineres no Buffer",
        buffer_move_to: "Mover para Buffer",
        buffer_remove_from: "Remover do Buffer",
        buffer_is_full: "O buffer está cheio. Não é possível adicionar mais contêineres.",
        buffer_filter_container: "Filtrar Contêiner",
    },
    zh: {
        main_title: "集装箱控制仪表板",
        upload_prompt_initial: "上传一个 .xlsx 文件开始",
        upload_prompt_success: "数据来自",
        upload_prompt_updated: "更新于:",
        upload_prompt_loaded_from_storage: "从本地存储加载的数据",
        total_containers_label: "集装箱总数（当前视图）:",
        export_overview_btn: "导出概览",
        upload_btn: "上传 XLSX",
        clear_data_btn: "清除数据",
        filter_po: "筛选采购订单",
        filter_vessel: "筛选船只",
        vessel_search_placeholder: "搜索...",
        filter_status: "集装箱状态",
        filter_final_status: "最终状态",
        filter_cargo_type: "货物类型",
        filter_shipowner: "船东",
        filter_loading_type: "装载类型",
        filter_btn: "筛选",
        clear_btn: "清除",
        view_vessel: "按船只",
        view_po: "按采购订单",
        view_warehouse: "按仓库",
        view_plan: "交付计划",
        view_buffer: "缓冲区",
        loading_text: "处理中...",
        summary_high: "高风险（逾期）",
        summary_medium: "中风险（≤ 7天内到期）",
        summary_low: "低风险（> 7天）",
        summary_none: "已交付",
        chart_status_title: "状态分布",
        chart_deadline_title: "截止日期分布",
        placeholder_title: "等待文件...",
        placeholder_subtitle: "选择一个电子表格开始分析。",
        containers_unit: "个集装箱",
        view_by_vessel: "按船只的集装箱总数",
        view_by_po: "按采购订单的集装箱总数",
        view_by_warehouse: "按仓库的集装箱总数",
        plan_no_pending: "无待处理的集装箱",
        plan_no_pending_subtitle: "当前筛选条件下没有需要计划的集装箱。",
        plan_daily_schedule_for: "交付计划于",
        plan_export_pdf: "导出 PDF",
        plan_export_excel: "导出 Excel",
        plan_pending_containers: "个待处理的集装箱",
        plan_daily_capacity: "每日集装箱数量:",
        modal_delivered_status: "已交付",
        buffer_capacity: "缓冲区容量:",
        buffer_total_capacity: "总容量",
        buffer_in_buffer: "在缓冲区",
        buffer_available_space: "可用空间",
        buffer_at_facility: "在工厂",
        buffer_facility_list_title: "工厂中的集装箱",
        buffer_buffer_list_title: "缓冲区中的集装箱",
        buffer_move_to: "移至缓冲区",
        buffer_remove_from: "从缓冲区移除",
        buffer_is_full: "缓冲区已满。无法添加更多集装箱。",
        buffer_filter_container: "筛选集装箱",
    }
};

const t = (key: keyof typeof translations.pt, el?: HTMLElement) => {
    const text = translations[currentLanguage][key] || key;
    if (el) {
        if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
            (el as HTMLInputElement).placeholder = text;
        } else {
            el.innerHTML = text;
        }
    }
    return text;
};

// --- Main App Initialization ---
window.addEventListener('load', initializeApp);

function initializeApp() {
    if (typeof Chart === 'undefined' || typeof jspdf === 'undefined' || typeof html2canvas === 'undefined' || typeof XLSX === 'undefined') {
        setTimeout(initializeApp, 100); return;
    }
    Chart.register(ChartDataLabels);

    // Event Listeners
    fileUpload.addEventListener('change', handleFileUpload);
    applyFiltersBtn.addEventListener('click', applyFiltersAndRender);
    resetFiltersBtn.addEventListener('click', resetFiltersAndRender);
    vesselSearchInput.addEventListener('input', filterVesselOptions);
    shipownerSearchInput.addEventListener('input', filterShipownerOptions);
    viewTabsContainer.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        const button = target.closest('button');
        if (button?.id.startsWith('view-')) {
            const view = button.id.replace('view-', '').replace('-btn', '') as View;
            setView(view);
        }
    });
    exportPdfBtn.addEventListener('click', handlePdfExport);
    clearDataBtn.addEventListener('click', () => {
        if(confirm("Você tem certeza que deseja limpar os dados salvos?")) {
            localStorage.removeItem(STORAGE_KEY);
            resetUI();
            showToast("Dados salvos foram limpos.", "success");
        }
    });
    summaryContainer.addEventListener('click', handleSummaryCardClick);
    translateBtn.addEventListener('click', toggleLanguage);
    themeToggleBtn.addEventListener('click', toggleTheme);
    dashboardGrid.addEventListener('click', (e) => {
        if ((e.target as HTMLElement).closest('.export-plan-btn')) {
            handleDailyPlanExport(e);
        }
        if ((e.target as HTMLElement).closest('.buffer-action-btn')) {
            handleBufferAction(e);
        }
    });
    dailyCapacityInput.addEventListener('change', () => {
        if(currentView === 'plan') {
            renderDeliveryPlanView(filteredDataCache);
        }
    });
     bufferCapacityInput.addEventListener('change', () => {
        if (currentView === 'buffer') {
            renderBufferControlView(filteredDataCache);
        }
    });
    bufferSearchInput.addEventListener('input', () => {
        if (currentView === 'buffer') {
            renderBufferControlView(filteredDataCache);
        }
    });

    // Modal Listeners
    modalCloseBtn.addEventListener('click', closeModal);
    detailsModal.addEventListener('click', (e) => { if (e.target === detailsModal) closeModal(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !detailsModal.classList.contains('hidden')) closeModal(); });

    initializeTheme();
    loadDataFromStorage();
}

// --- Loading Indicator ---
function showLoading() { loadingOverlay.classList.remove('hidden'); }
function hideLoading() { loadingOverlay.classList.add('hidden'); }

// --- Toast Notifications ---
function showToast(message: string, type: 'success' | 'error' | 'warning' = 'success') {
    const toastContainer = document.getElementById('toast-container');
    if (!toastContainer) return;

    const toast = document.createElement('div');
    const icons = { success: 'fa-check-circle', error: 'fa-times-circle', warning: 'fa-exclamation-triangle' };
    const colors = { success: 'bg-green-500', error: 'bg-red-500', warning: 'bg-yellow-500' };
    toast.className = `toast ${colors[type]} text-white py-3 px-5 rounded-lg shadow-xl flex items-center mb-2`;
    toast.innerHTML = `<i class="fas ${icons[type]} mr-3"></i> <p>${message}</p>`;
    toastContainer.appendChild(toast);
    setTimeout(() => toast.remove(), 5000);
}

// --- Theme Management ---
function initializeTheme() {
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY) as 'light' | 'dark' | null;
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const theme = savedTheme || (prefersDark ? 'dark' : 'light');
    applyTheme(theme);
}

function toggleTheme() {
    const isDark = document.documentElement.classList.contains('dark');
    const newTheme = isDark ? 'light' : 'dark';
    localStorage.setItem(THEME_STORAGE_KEY, newTheme);
    applyTheme(newTheme);
}

function applyTheme(theme: 'light' | 'dark') {
    const themeToggleIcon = document.getElementById('theme-toggle-icon') as HTMLElement;
    if (theme === 'dark') {
        document.documentElement.classList.add('dark');
        themeToggleIcon.classList.remove('fa-moon');
        themeToggleIcon.classList.add('fa-sun');
    } else {
        document.documentElement.classList.remove('dark');
        themeToggleIcon.classList.remove('fa-sun');
        themeToggleIcon.classList.add('fa-moon');
    }

    const isDarkMode = theme === 'dark';
    Chart.defaults.color = isDarkMode ? '#cbd5e1' : '#64748b'; // slate-300 / slate-500
    Chart.defaults.borderColor = isDarkMode ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.1)';

    if (originalData.length > 0) {
        // Re-render charts with new theme colors
        applyFiltersAndRender();
    }
}


// --- Translation ---
function toggleLanguage() {
    currentLanguage = currentLanguage === 'pt' ? 'zh' : 'pt';
    translateBtnText.textContent = currentLanguage === 'pt' ? '中文' : 'Português';
    translateUI();
}

function translateUI() {
    document.querySelectorAll('[data-translate-key]').forEach(el => {
        const key = el.getAttribute('data-translate-key') as keyof typeof translations.pt;
        if (key) {
           t(key, el as HTMLElement);
        }
    });
    // Re-render dynamic content that needs translation
    if (originalData.length > 0) {
        applyFiltersAndRender();
        if (activeModalItem) {
            renderModalContent();
        }
    }
}

// --- Data Persistence ---
function loadDataFromStorage() {
    const savedStateJSON = localStorage.getItem(STORAGE_KEY);
    if (savedStateJSON) {
        try {
            const savedState = JSON.parse(savedStateJSON);
            if(savedState.data && Array.isArray(savedState.data) && savedState.timestamp) {
                originalData = savedState.data;
                containersInBuffer = new Set(savedState.containersInBuffer || []);
                populateFilters(originalData);
                applyFiltersAndRender();
                showDashboard();
                const loadedDate = new Date(savedState.timestamp);
                lastUpdate.textContent = `${t('upload_prompt_loaded_from_storage')} "${loadedDate.toLocaleDateString('pt-BR')}" | ${t('upload_prompt_updated')} ${loadedDate.toLocaleTimeString('pt-BR')}`;
                showToast("Dados carregados do armazenamento local.", "success");
            }
        } catch (e) {
            console.error("Failed to load data from storage", e);
            localStorage.removeItem(STORAGE_KEY);
        }
    }
}

function saveDataToStorage() {
    const stateToSave = {
        data: originalData,
        containersInBuffer: Array.from(containersInBuffer),
        timestamp: new Date().toISOString()
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(stateToSave));
}

// --- File Handling ---
function handleFileUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    const uploadLabel = document.querySelector('label[for="file-upload"] span');
    if (!file || !uploadLabel) return;

    const parentLabel = uploadLabel.parentElement as HTMLElement;
    parentLabel.classList.add('opacity-50', 'cursor-not-allowed');
    uploadLabel.innerHTML = `<i class="fas fa-spinner fa-spin mr-2"></i> ${t('loading_text')}`;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const workbook = XLSX.read(new Uint8Array(e.target!.result as ArrayBuffer), { type: 'array' });
            const sheetName = workbook.SheetNames.find((name: string) => name.toUpperCase().includes('PLANILHA1'));
            if (!sheetName || !workbook.Sheets[sheetName]) throw new Error(`Nenhuma planilha válida (ex: "Planilha1") foi encontrada.`);
            
            const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: false, defval: '' });
            if (!Array.isArray(rawData)) throw new Error("O formato dos dados da planilha é inválido.");
            if (rawData.length === 0) throw new Error("A planilha está vazia.");

            originalData = rawData.map((row: any) => {
                const normalizedRow: { [key: string]: any } = {};
                for (const key in row) {
                    if (Object.prototype.hasOwnProperty.call(row, key)) {
                        const normalizedKey = key.trim().toUpperCase().replace(/\s\s+/g, ' ');
                        normalizedRow[normalizedKey] = row[key];
                    }
                }
                return normalizedRow;
            });
            
            containersInBuffer = new Set(); // Reset buffer on new file upload
            saveDataToStorage();
            populateFilters(originalData);
            applyFiltersAndRender();
            showDashboard();

            lastUpdate.textContent = `${t('upload_prompt_success')} "${sheetName}" | ${t('upload_prompt_updated')} ${new Date().toLocaleString('pt-BR')}`;
            showToast('Dashboard carregado com sucesso!', 'success');
        } catch (err: any) {
            showToast(err.message || 'Erro ao processar arquivo.', 'error');
            resetUI();
        } finally {
            parentLabel.classList.remove('opacity-50', 'cursor-not-allowed');
            uploadLabel.innerHTML = t('upload_btn');
            fileUpload.value = '';
        }
    };
    reader.onerror = () => { showToast('Não foi possível ler o arquivo.', 'error'); resetUI(); };
    reader.readAsArrayBuffer(file);
}

// --- Data Processing & Filtering ---
function excelDateToJSDate(serial: any): Date | null {
    if (serial === null || serial === undefined || serial === '') return null;
    if (typeof serial === 'string') {
        if (/^\d{5}$/.test(serial)) { serial = parseInt(serial, 10); }
        else {
            const parts = serial.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
            if (parts) {
                const date = new Date(Date.UTC(Number(parts[3]), Number(parts[2]) - 1, Number(parts[1])));
                return isNaN(date.getTime()) ? null : date;
            }
            const fallbackDate = new Date(serial);
            return isNaN(fallbackDate.getTime()) ? null : fallbackDate;
        }
    }
    if (typeof serial === 'number' && serial >= 1) {
        const utc_days = Math.floor(serial - 25569);
        const date_info = new Date(utc_days * 86400 * 1000);
        return new Date(date_info.getTime() + (date_info.getTimezoneOffset() * 60 * 1000));
    }
    return null;
}

function getContainerRisk(container: any): RiskCategory {
    const isDelivered = (container['STATUS'] || '').toLowerCase().includes('entregue');
    if (isDelivered) return 'none';
    const deadline = excelDateToJSDate(container['DEADLINE RETURN CNTR']);
    if (deadline) {
        const daysToDeadline = Math.ceil((deadline.getTime() - TODAY.getTime()) / (1000 * 3600 * 24));
        if (daysToDeadline < 0) return 'high';
        if (daysToDeadline <= 7) return 'medium';
    }
    return 'low';
}

function getSelectedOptions(select: HTMLSelectElement) { return Array.from(select.selectedOptions).map(opt => opt.value); }

function filterVesselOptions() {
    const searchTerm = vesselSearchInput.value.toUpperCase();
    const options = vesselFilter.options;
    for (let i = 0; i < options.length; i++) {
        const option = options[i] as HTMLOptionElement;
        const optionText = option.textContent?.toUpperCase() || '';
        option.style.display = optionText.includes(searchTerm) ? '' : 'none';
    }
}

function filterShipownerOptions() {
    const searchTerm = shipownerSearchInput.value.toUpperCase();
    const options = shipownerFilter.options;
    for (let i = 0; i < options.length; i++) {
        const option = options[i] as HTMLOptionElement;
        const optionText = option.textContent?.toUpperCase() || '';
        option.style.display = optionText.includes(searchTerm) ? '' : 'none';
    }
}

function applyFiltersAndRender() {
    showLoading();
    activeRiskFilter = null; 
    updateSummaryCardStyles();

    setTimeout(() => {
        const selectedPOs = getSelectedOptions(poFilter).map(v => v.toUpperCase());
        const selectedVessels = getSelectedOptions(vesselFilter).map(v => v.toUpperCase());
        const selectedStatuses = getSelectedOptions(statusFilter).map(v => v.toUpperCase());
        const selectedFinalStatuses = getSelectedOptions(finalStatusFilter).map(v => v.toUpperCase());
        const selectedCargoTypes = getSelectedOptions(cargoTypeFilter).map(v => v.toUpperCase());
        const selectedShipowners = getSelectedOptions(shipownerFilter).map(v => v.toUpperCase());
        const selectedLoadingTypes = getSelectedOptions(loadingTypeFilter).map(v => v.toUpperCase());

        filteredDataCache = originalData.filter(row => 
            (selectedPOs.length === 0 || selectedPOs.includes(String(row['PO SAP'] || '').toUpperCase())) &&
            (selectedVessels.length === 0 || selectedVessels.includes(String(row['ARRIVAL VESSEL'] || '').toUpperCase())) &&
            (selectedStatuses.length === 0 || selectedStatuses.includes(String(row['STATUS CNTR WAREHOUSE'] || '').toUpperCase())) &&
            (selectedFinalStatuses.length === 0 || selectedFinalStatuses.includes(String(row['STATUS'] || '').toUpperCase())) &&
            (selectedCargoTypes.length === 0 || selectedCargoTypes.includes(String(row['TYPE OF CARGO'] || '').toUpperCase())) &&
            (selectedShipowners.length === 0 || selectedShipowners.includes(String(row['SHIPOWNER'] || '').toUpperCase())) &&
            (selectedLoadingTypes.length === 0 || selectedLoadingTypes.includes(String(row['LOADING TYPE'] || '').toUpperCase()))
        );
        
        setViewUI(currentView);
        
        totalFclCount.textContent = filteredDataCache.length.toString();
        hideLoading();
    }, 50);
}

function resetFiltersAndRender() {
    showLoading();
    setTimeout(() => {
        [poFilter, vesselFilter, statusFilter, finalStatusFilter, cargoTypeFilter, shipownerFilter, loadingTypeFilter].forEach(sel => sel.selectedIndex = -1);
        vesselSearchInput.value = '';
        filterVesselOptions();
        shipownerSearchInput.value = '';
        filterShipownerOptions();
        applyFiltersAndRender();
        hideLoading();
    }, 50);
}

function processDataForView(data: any[]) {
    const groupByKeyMap = { vessel: 'ARRIVAL VESSEL', po: 'PO SAP', warehouse: 'BONDED WAREHOUSE' };
    const groupByKey = groupByKeyMap[currentView as 'vessel' | 'po' | 'warehouse'];
    const defaultName = `Sem ${groupByKey}`.toUpperCase();
    
    const grouped = data.reduce((acc, row) => {
        const name = (row[groupByKey] ? String(row[groupByKey]).trim().toUpperCase() : defaultName) || defaultName;
        if (!acc[name]) acc[name] = [];
        acc[name].push(row);
        return acc;
    }, {} as Record<string, any[]>);

    return Object.entries(grouped).map(([name, containers]) => {
        const processed = containers.map(c => {
            const risk = getContainerRisk(c);
            const deadline = excelDateToJSDate(c['DEADLINE RETURN CNTR']);
            const daysToDeadline = deadline ? Math.ceil((deadline.getTime() - TODAY.getTime()) / (1000 * 3600 * 24)) : null;
            return { ...c, daysToDeadline, risk };
        });
        const hasHighRisk = processed.some(c => c.risk === 'high');
        const hasMediumRisk = processed.some(c => c.risk === 'medium');
        const isAllDelivered = processed.every(c => c.risk === 'none');
        let overallRisk: RiskCategory | 'low' = 'low';
        if(isAllDelivered) overallRisk = 'none';
        else if(hasHighRisk) overallRisk = 'high';
        else if(hasMediumRisk) overallRisk = 'medium';
        return { name, containers: processed, totalFCL: containers.length, overallRisk };
    }).sort((a, b) => {
        const riskOrder = { high: 0, medium: 1, low: 2, none: 3 };
        return riskOrder[a.overallRisk as keyof typeof riskOrder] - riskOrder[b.overallRisk as keyof typeof riskOrder];
    });
}

// --- View Switching ---
function setView(view: View) {
    if (currentView === view) return;
    if (currentView === 'buffer') {
        bufferSearchInput.value = '';
    }
    currentView = view;
    document.querySelectorAll('#view-tabs-container button').forEach(btn => btn.classList.remove('active'));
    document.getElementById(`view-${view}-btn`)?.classList.add('active');
    if (originalData.length > 0) setViewUI(view);
}

function setViewUI(view: View) {
    planConfigContainer.classList.add('hidden');
    bufferConfigContainer.classList.add('hidden');
    chartsContainer.classList.add('hidden');
    summaryContainer.classList.add('hidden');
    exportPdfBtn.classList.add('hidden');
    
    if (view === 'plan') {
        planConfigContainer.classList.remove('hidden');
        renderDeliveryPlanView(filteredDataCache);
    } else if (view === 'buffer') {
        bufferConfigContainer.classList.remove('hidden');
        renderBufferControlView(filteredDataCache);
    } else {
        chartsContainer.classList.remove('hidden');
        summaryContainer.classList.remove('hidden');
        exportPdfBtn.classList.remove('hidden');
        renderSummary(filteredDataCache);
        const processedData = processDataForView(filteredDataCache);
        renderDashboard(processedData);
        renderCharts(processedData, filteredDataCache);
    }
}

// --- UI Rendering ---
function renderSummary(filteredData: any[]) {
    const summaryCounts: Record<RiskCategory, number> = { high: 0, medium: 0, low: 0, none: 0 };
    for (const row of filteredData) {
        summaryCounts[getContainerRisk(row)]++;
    }
    highRiskCount.textContent = summaryCounts.high.toString();
    mediumRiskCount.textContent = summaryCounts.medium.toString();
    lowRiskCount.textContent = summaryCounts.low.toString();
    noneRiskCount.textContent = summaryCounts.none.toString();
}

function renderDashboard(dataToRender: any[]) {
    dashboardGrid.innerHTML = '';
    dashboardGrid.className = 'grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6';
    if (!Array.isArray(dataToRender) || dataToRender.length === 0) {
        placeholder.classList.remove('hidden');
        return;
    }
    placeholder.classList.add('hidden');
    dataToRender.forEach(item => {
        const card = createDashboardCard(item);
        card.addEventListener('click', () => openModal(item));
        dashboardGrid.appendChild(card);
    });
}

function createDashboardCard(item: any) {
    const card = document.createElement('div');
    card.className = `card risk-${item.overallRisk} bg-white dark:bg-slate-800`;
    card.innerHTML = `
        <div class="p-4">
            <h3 class="font-extrabold text-lg text-gray-800 dark:text-slate-100 truncate">${currentView === 'po' ? 'PO: ' : ''}${item.name}</h3>
            <p class="text-sm text-gray-500 dark:text-slate-400">${item.totalFCL} ${t('containers_unit')}</p>
        </div>`;
    return card;
}

// --- Delivery Plan ---
function renderDeliveryPlanView(data: any[]) {
    dashboardGrid.innerHTML = '';
    dashboardGrid.className = 'grid grid-cols-1 gap-6';

    const containersToSchedule = data
        .filter(c => getContainerRisk(c) !== 'none')
        .map(c => {
            const deadline = excelDateToJSDate(c['DEADLINE RETURN CNTR']);
            const daysToDeadline = deadline ? Math.ceil((deadline.getTime() - TODAY.getTime()) / (1000 * 3600 * 24)) : null;
            return { ...c, daysToDeadline, risk: getContainerRisk(c) };
        })
        .sort((a, b) => (a.daysToDeadline ?? Infinity) - (b.daysToDeadline ?? Infinity));

    if (containersToSchedule.length === 0) {
        dashboardGrid.innerHTML = `
            <div class="col-span-full text-center py-20 bg-white dark:bg-slate-800 rounded-lg shadow">
                <i class="fas fa-check-circle text-6xl text-green-400 mb-4"></i>
                <h2 class="text-2xl font-semibold text-gray-600 dark:text-slate-300">${t('plan_no_pending')}</h2>
                <p class="text-gray-400 dark:text-slate-500">${t('plan_no_pending_subtitle')}</p>
            </div>`;
        return;
    }

    deliveryPlanCache = {};
    const dailyCapacity = parseInt(dailyCapacityInput.value, 10) || 100;
    let scheduleDate = new Date(TODAY);
    for (const container of containersToSchedule) {
        let scheduled = false;
        while(!scheduled) {
            // Skip weekends
            while (scheduleDate.getDay() === 0 || scheduleDate.getDay() === 6) {
                scheduleDate.setDate(scheduleDate.getDate() + 1);
            }
            const dateString = scheduleDate.toISOString().split('T')[0];
            if (!deliveryPlanCache[dateString]) {
                deliveryPlanCache[dateString] = [];
            }
            if (deliveryPlanCache[dateString].length < dailyCapacity) {
                deliveryPlanCache[dateString].push(container);
                scheduled = true;
            } else {
                scheduleDate.setDate(scheduleDate.getDate() + 1);
            }
        }
    }
    
    let planHtml = '';
    const headers = ['#', 'Contêiner', 'Prazo Retorno', 'Dias Restantes', 'Navio', 'PO SAP', 'Armazém'];
    const tableHeader = headers.map(h => `<th class="px-2 py-2 text-left font-semibold text-gray-600 dark:text-slate-300 text-xs uppercase">${h}</th>`).join('');

    // FIX: Use Object.entries() to ensure `daySchedule` is correctly typed as an array. This resolves errors with `.map` and `.length` during iteration. The sort is maintained to keep output chronological.
    for (const [date, daySchedule] of Object.entries(deliveryPlanCache).sort((a, b) => a[0].localeCompare(b[0]))) {

        const formattedDate = new Date(date).toLocaleDateString('pt-BR', { timeZone: 'UTC', weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
        
        const tableRows = daySchedule.map((c: any, index: number) => {
            const daysText = c.daysToDeadline !== null ? `<span class="font-bold ${c.daysToDeadline < 0 ? 'text-red-600' : ''}">${c.daysToDeadline}</span>` : 'N/A';
            return `<tr class="border-b dark:border-slate-700 row-risk-${c.risk}">
                <td class="px-2 py-1.5 text-xs text-center">${index + 1}</td>
                <td class="px-2 py-1.5 text-xs font-semibold">${c['CNTRS ORIGINAL'] || ''}</td>
                <td class="px-2 py-1.5 text-xs">${formatDate(c['DEADLINE RETURN CNTR'])}</td>
                <td class="px-2 py-1.5 text-xs text-center">${daysText}</td>
                <td class="px-2 py-1.5 text-xs">${c['ARRIVAL VESSEL'] || ''}</td>
                <td class="px-2 py-1.5 text-xs">${c['PO SAP'] || ''}</td>
                <td class="px-2 py-1.5 text-xs">${c['BONDED WAREHOUSE'] || ''}</td>
            </tr>`;
        }).join('');

        planHtml += `
            <div class="bg-white dark:bg-slate-800 rounded-lg shadow-md overflow-hidden col-span-full">
                <div class="p-4 bg-gray-50 dark:bg-slate-700/50 border-b dark:border-slate-700 flex justify-between items-center">
                    <div>
                        <h3 class="text-lg font-bold text-gray-800 dark:text-slate-100">${t('plan_daily_schedule_for')} ${formattedDate}</h3>
                        <p class="text-sm text-gray-600 dark:text-slate-400">${daySchedule.length} ${t('containers_unit')}</p>
                    </div>
                    <div class="flex space-x-2">
                        <button class="export-plan-btn bg-red-600 text-white px-3 py-1.5 text-xs rounded-md shadow-sm hover:bg-red-700" data-date="${date}" data-format="pdf">
                            <i class="fas fa-file-pdf mr-1"></i> ${t('plan_export_pdf')}
                        </button>
                        <button class="export-plan-btn bg-green-600 text-white px-3 py-1.5 text-xs rounded-md shadow-sm hover:bg-green-700" data-date="${date}" data-format="excel">
                            <i class="fas fa-file-excel mr-1"></i> ${t('plan_export_excel')}
                        </button>
                    </div>
                </div>
                <div class="table-responsive p-4">
                    <table class="min-w-full text-sm">
                        <thead><tr>${tableHeader}</tr></thead>
                        <tbody>${tableRows}</tbody>
                    </table>
                </div>
            </div>`;
    }
    dashboardGrid.innerHTML = planHtml;
}

// --- Buffer Control ---
function renderBufferControlView(data: any[]) {
    dashboardGrid.innerHTML = '';
    dashboardGrid.className = 'grid grid-cols-1 gap-6';

    const searchTerm = bufferSearchInput.value.toUpperCase();
    const capacity = parseInt(bufferCapacityInput.value, 10) || 400;
    const inBufferCount = containersInBuffer.size;
    const availableSpace = capacity - inBufferCount;
    
    const allAtFacility = data.filter(c => !containersInBuffer.has(c['CNTRS ORIGINAL']));
    const allInBuffer = data.filter(c => containersInBuffer.has(c['CNTRS ORIGINAL']));
    
    const atFacilityContainers = allAtFacility.filter(c => 
        searchTerm === '' || String(c['CNTRS ORIGINAL'] || '').toUpperCase().includes(searchTerm)
    );
    const inBufferContainers = allInBuffer.filter(c => 
        searchTerm === '' || String(c['CNTRS ORIGINAL'] || '').toUpperCase().includes(searchTerm)
    );

    const statsHtml = `
    <div class="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-6 mb-6">
        <div class="bg-white dark:bg-slate-800 p-4 rounded-lg shadow flex items-center border-l-4 border-blue-500">
            <div class="p-3 rounded-full bg-blue-100 mr-4"><i class="fas fa-boxes text-blue-600 fa-lg"></i></div>
            <div><p class="text-sm font-medium text-gray-500 dark:text-slate-400 uppercase" data-translate-key="buffer_total_capacity">${t('buffer_total_capacity')}</p><p class="text-2xl font-bold text-gray-800 dark:text-slate-100">${capacity}</p></div>
        </div>
        <div class="bg-white dark:bg-slate-800 p-4 rounded-lg shadow flex items-center border-l-4 border-purple-500">
            <div class="p-3 rounded-full bg-purple-100 mr-4"><i class="fas fa-layer-group text-purple-600 fa-lg"></i></div>
            <div><p class="text-sm font-medium text-gray-500 dark:text-slate-400 uppercase" data-translate-key="buffer_in_buffer">${t('buffer_in_buffer')}</p><p class="text-2xl font-bold text-gray-800 dark:text-slate-100">${inBufferCount}</p></div>
        </div>
        <div class="bg-white dark:bg-slate-800 p-4 rounded-lg shadow flex items-center border-l-4 border-green-500">
            <div class="p-3 rounded-full bg-green-100 mr-4"><i class="fas fa-check-circle text-green-600 fa-lg"></i></div>
            <div><p class="text-sm font-medium text-gray-500 dark:text-slate-400 uppercase" data-translate-key="buffer_available_space">${t('buffer_available_space')}</p><p class="text-2xl font-bold text-gray-800 dark:text-slate-100">${availableSpace}</p></div>
        </div>
        <div class="bg-white dark:bg-slate-800 p-4 rounded-lg shadow flex items-center border-l-4 border-yellow-500">
            <div class="p-3 rounded-full bg-yellow-100 mr-4"><i class="fas fa-industry text-yellow-600 fa-lg"></i></div>
            <div><p class="text-sm font-medium text-gray-500 dark:text-slate-400 uppercase" data-translate-key="buffer_at_facility">${t('buffer_at_facility')}</p><p class="text-2xl font-bold text-gray-800 dark:text-slate-100">${allAtFacility.length}</p></div>
        </div>
    </div>`;

    const createTable = (titleKey: keyof typeof translations.pt, containers: any[], action: 'add' | 'remove') => {
        const headers = ['Contêiner', 'Navio', 'PO SAP', 'Ação'];
        const buttonText = action === 'add' ? t('buffer_move_to') : t('buffer_remove_from');
        const buttonClass = action === 'add' ? 'bg-blue-500 hover:bg-blue-600' : 'bg-red-500 hover:bg-red-600';

        const tableRows = containers.map(c => `
            <tr class="border-b dark:border-slate-700">
                <td class="p-2 text-xs font-semibold">${c['CNTRS ORIGINAL'] || ''}</td>
                <td class="p-2 text-xs">${c['ARRIVAL VESSEL'] || ''}</td>
                <td class="p-2 text-xs">${c['PO SAP'] || ''}</td>
                <td class="p-2 text-xs text-right">
                    <button class="buffer-action-btn text-white px-2 py-1 rounded text-xs ${buttonClass}" data-id="${c['CNTRS ORIGINAL']}" data-action="${action}">
                       ${buttonText}
                    </button>
                </td>
            </tr>
        `).join('');

        return `
            <div class="bg-white dark:bg-slate-800 p-4 rounded-lg shadow">
                <h3 class="text-xl font-bold mb-2 text-gray-700 dark:text-slate-200">${t(titleKey)}</h3>
                <div class="buffer-table-container">
                    <table class="min-w-full text-sm">
                        <thead class="bg-gray-50 dark:bg-slate-700"><tr>${headers.map(h => `<th class="p-2 text-left font-semibold text-gray-600 dark:text-slate-300 text-xs uppercase">${h}</th>`).join('')}</tr></thead>
                        <tbody>${tableRows}</tbody>
                    </table>
                </div>
            </div>`;
    };

    const tablesHtml = `
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
            ${createTable('buffer_facility_list_title', atFacilityContainers, 'add')}
            ${createTable('buffer_buffer_list_title', inBufferContainers, 'remove')}
        </div>
    `;

    dashboardGrid.innerHTML = statsHtml + tablesHtml;
}


// --- Chart Rendering ---
function renderCharts(processedData: any[], filteredData: any[]) {
    const viewKeyMap = { vessel: 'view_by_vessel', po: 'view_by_po', warehouse: 'view_by_warehouse' } as const;
    barChartTitle.textContent = t(viewKeyMap[currentView as keyof typeof viewKeyMap]);

    const updateChart = (chartId: string, type: any, data: any, options: any) => {
        const ctx = (document.getElementById(chartId) as HTMLCanvasElement)?.getContext('2d');
        if (!ctx) return;
        if (charts[chartId]) charts[chartId].destroy();
        charts[chartId] = new Chart(ctx, { type, data, options });
    };

    const safeProcessedData = Array.isArray(processedData) ? processedData : [];
    const barData = {
        labels: safeProcessedData.map(item => item.name),
        datasets: [{ label: 'Total de Contêineres', data: safeProcessedData.map(item => item.totalFCL), backgroundColor: 'rgba(59, 130, 246, 0.7)' }]
    };
    updateChart('main-bar-chart', 'bar', barData, { indexAxis: 'y', responsive: true, plugins: { legend: { display: false } } });
    
    const safeFilteredData = Array.isArray(filteredData) ? filteredData : [];
    const statusCounts = safeFilteredData.reduce((acc, row) => {
        const status = row['STATUS CNTR WAREHOUSE'] || 'Sem Status';
        acc[status] = (acc[status] || 0) + 1;
        return acc;
    }, {} as Record<string, number>);
    const pieData = {
        labels: Object.keys(statusCounts),
        datasets: [{ data: Object.values(statusCounts), backgroundColor: ['#ef4444', '#f97316', '#f59e0b', '#84cc16', '#22c55e', '#10b981', '#14b8a6', '#3b82f6', '#8b5cf6', '#a855f7'] }]
    };
    updateChart('status-pie-chart', 'pie', pieData, { responsive: true, plugins: { legend: { position: 'right' } } });

    const deadlineGroups = safeFilteredData.reduce((acc, row) => {
        if ((row['STATUS']||'').toLowerCase().includes('entregue')) return acc;
        const deadline = excelDateToJSDate(row['DEADLINE RETURN CNTR']);
        if (!deadline) return acc;
        const days = Math.ceil((deadline.getTime() - TODAY.getTime()) / (1000 * 3600 * 24));
        let group = '31+ dias';
        if (days < 0) group = 'Atrasado';
        else if (days <= 7) group = '0-7 dias';
        else if (days <= 15) group = '8-15 dias';
        else if (days <= 30) group = '16-30 dias';
        acc[group] = (acc[group] || 0) + 1;
        return acc;
    }, {} as Record<string, number>);
    const deadlineLabels = ['Atrasado', '0-7 dias', '8-15 dias', '16-30 dias', '31+ dias'];
    const deadlineData = {
        labels: deadlineLabels,
        datasets: [{ label: 'Nº de Contêineres', data: deadlineLabels.map(l => deadlineGroups[l] || 0), backgroundColor: ['#dc2626', '#f97316', '#facc15', '#84cc16', '#22c55e'] }]
    };
    updateChart('deadline-distribution-chart', 'bar', deadlineData, { responsive: true, plugins: { legend: { display: false } } });
}

// --- Helper Functions ---
function populateFilters(data: any[]) {
    const createOptions = (arr: string[]) => arr.map(item => `<option value="${item}">${item}</option>`).join('');
    const unique = (key: string) => {
        const seen = new Map<string, string>();
        for (const row of data) {
            const originalValue = String(row[key] || '').trim();
            if (originalValue) {
                const normalizedValue = originalValue.toUpperCase();
                if (!seen.has(normalizedValue)) seen.set(normalizedValue, originalValue);
            }
        }
        return [...seen.values()].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    };
    poFilter.innerHTML = createOptions(unique('PO SAP'));
    vesselFilter.innerHTML = createOptions(unique('ARRIVAL VESSEL'));
    statusFilter.innerHTML = createOptions(unique('STATUS CNTR WAREHOUSE'));
    finalStatusFilter.innerHTML = createOptions(unique('STATUS'));
    cargoTypeFilter.innerHTML = createOptions(unique('TYPE OF CARGO'));
    shipownerFilter.innerHTML = createOptions(unique('SHIPOWNER'));
    loadingTypeFilter.innerHTML = createOptions(unique('LOADING TYPE'));
}

function showDashboard() {
    filterContainer.classList.remove('hidden');
    viewTabsContainer.classList.remove('hidden');
    exportPdfBtn.classList.remove('hidden');
    clearDataBtn.classList.remove('hidden');
    totalFclDisplay.classList.remove('hidden');
}

function resetUI() {
    dashboardGrid.innerHTML = '';
    placeholder.classList.remove('hidden');
    filterContainer.classList.add('hidden');
    chartsContainer.classList.add('hidden');
    summaryContainer.classList.add('hidden');
    viewTabsContainer.classList.add('hidden');
    exportPdfBtn.classList.add('hidden');
    clearDataBtn.classList.add('hidden');
    totalFclDisplay.classList.add('hidden');
    planConfigContainer.classList.add('hidden');
    bufferConfigContainer.classList.add('hidden');
    originalData = [];
    containersInBuffer.clear();
    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};
    lastUpdate.textContent = t('upload_prompt_initial');
    setView('vessel');
}

function formatDate(dateString: any): string {
    const date = excelDateToJSDate(dateString);
    if (!date) return 'N/A';
    return date.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
}

// --- Event Handlers ---
function handleSummaryCardClick(event: Event) {
    const card = (event.target as HTMLElement).closest('.summary-card') as HTMLElement;
    if (!card || currentView === 'plan' || currentView === 'buffer') return;
    const risk = card.dataset.risk as RiskCategory;
    activeRiskFilter = activeRiskFilter === risk ? null : risk;
    updateSummaryCardStyles();
    const chartRawDataToRender = activeRiskFilter ? filteredDataCache.filter(row => getContainerRisk(row) === activeRiskFilter) : filteredDataCache;
    let gridDataToRender = processDataForView(filteredDataCache);
    if (activeRiskFilter) {
        gridDataToRender = gridDataToRender.map(group => {
            const filteredContainers = group.containers.filter((c: any) => c.risk === activeRiskFilter);
            return { ...group, containers: filteredContainers, totalFCL: filteredContainers.length };
        }).filter(group => group.totalFCL > 0);
    }
    renderDashboard(gridDataToRender);
    renderCharts(gridDataToRender, chartRawDataToRender);
}

function updateSummaryCardStyles() {
    summaryContainer.querySelectorAll('.summary-card').forEach(card => {
        const cardEl = card as HTMLElement;
        if (cardEl.dataset.risk === activeRiskFilter) cardEl.classList.add('active-summary-card');
        else cardEl.classList.remove('active-summary-card');
    });
}

function handleBufferAction(event: Event) {
    const button = (event.target as HTMLElement).closest('.buffer-action-btn') as HTMLElement;
    if (!button) return;

    const id = button.dataset.id;
    const action = button.dataset.action;
    if (!id || !action) return;
    
    if (action === 'add') {
        const capacity = parseInt(bufferCapacityInput.value, 10) || 400;
        if (containersInBuffer.size >= capacity) {
            showToast(t('buffer_is_full'), 'warning');
            return;
        }
        containersInBuffer.add(id);
    } else if (action === 'remove') {
        containersInBuffer.delete(id);
    }

    saveDataToStorage();
    renderBufferControlView(filteredDataCache);
}

// --- Exporting ---
async function handlePdfExport() {
    showLoading();
    const reportContent = document.getElementById('report-content');
    if (!reportContent) { hideLoading(); return; }
    try {
        const { jsPDF } = jspdf;
        const canvas = await html2canvas(reportContent, { scale: 2, useCORS: true });
        const imgData = canvas.toDataURL('image/png');
        const pdf = new jsPDF({ orientation: 'landscape', unit: 'px', format: [canvas.width, canvas.height] });
        pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
        pdf.save(`dashboard_export_${currentView}_${new Date().toISOString().split('T')[0]}.pdf`);
    } catch (error) {
        console.error("Error generating PDF:", error);
        showToast("Ocorreu um erro ao gerar o PDF.", "error");
    } finally { hideLoading(); }
}

function handleDailyPlanExport(event: Event) {
    const target = (event.target as HTMLElement).closest('.export-plan-btn') as HTMLElement;
    if (!target) return;

    const date = target.dataset.date;
    const format = target.dataset.format;
    const dailyData = date ? deliveryPlanCache[date] : [];

    if (dailyData.length === 0) {
        showToast('Nenhum dado para exportar.', 'warning');
        return;
    }

    if (format === 'pdf') {
        exportDailyPlanToPdf(dailyData, date!);
    } else if (format === 'excel') {
        exportDailyPlanToExcel(dailyData, date!);
    }
}

function exportDailyPlanToPdf(data: any[], date: string) {
    const { jsPDF } = jspdf;
    const doc = new jsPDF();
    const formattedDate = new Date(date).toLocaleDateString('pt-BR', { timeZone: 'UTC' });

    const head = [['#', 'Contêiner', 'Prazo Retorno', 'Dias Restantes', 'Navio', 'PO SAP', 'Armazém']];
    const body = data.map((c, i) => [
        i + 1,
        c['CNTRS ORIGINAL'] || '',
        formatDate(c['DEADLINE RETURN CNTR']),
        c.daysToDeadline ?? 'N/A',
        c['ARRIVAL VESSEL'] || '',
        c['PO SAP'] || '',
        c['BONDED WAREHOUSE'] || '',
    ]);

    doc.text(`${t('plan_daily_schedule_for')} ${formattedDate}`, 14, 15);
    (doc as any).autoTable({
        head,
        body,
        startY: 20,
        theme: 'grid',
        headStyles: { fillColor: [22, 160, 133] },
    });

    doc.save(`plano_entrega_${date}.pdf`);
}

function exportDailyPlanToExcel(data: any[], date: string) {
    const dataForSheet = data.map((c, i) => ({
        '#': i + 1,
        'Contêiner': c['CNTRS ORIGINAL'] || '',
        'Prazo Retorno': formatDate(c['DEADLINE RETURN CNTR']),
        'Dias Restantes': c.daysToDeadline ?? 'N/A',
        'Navio': c['ARRIVAL VESSEL'] || '',
        'PO SAP': c['PO SAP'] || '',
        'Armazém': c['BONDED WAREHOUSE'] || '',
    }));
    const worksheet = XLSX.utils.json_to_sheet(dataForSheet);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Plano de Entrega');
    XLSX.writeFile(workbook, `plano_entrega_${date}.xlsx`);
}

// --- Modal ---
function openModal(item: any) {
    activeModalItem = item;
    modalSortState = { column: 'daysToDeadline', direction: 'asc' };
    renderModalContent();
    detailsModal.classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
    setTimeout(() => detailsModal.classList.add('modal-open'), 10);
}

function closeModal() {
    detailsModal.classList.remove('modal-open');
     setTimeout(() => {
        detailsModal.classList.add('hidden');
        document.body.classList.remove('overflow-hidden');
        activeModalItem = null;
    }, 300);
}

function handleModalSort(event: Event) {
    const target = (event.target as HTMLElement).closest('th');
    if (!target || !target.dataset.key) return;
    const sortKey = target.dataset.key;
    if (modalSortState.column === sortKey) {
        modalSortState.direction = modalSortState.direction === 'asc' ? 'desc' : 'asc';
    } else {
        modalSortState.column = sortKey;
        modalSortState.direction = 'asc';
    }
    renderModalContent();
}

function renderModalContent() {
    if (!activeModalItem) return;

    const { name, totalFCL, containers } = activeModalItem;
    modalHeaderContent.innerHTML = `<h2 class="text-2xl font-bold text-gray-800 dark:text-slate-100">${currentView === 'po' ? 'PO: ' : ''}${name}</h2><p class="text-gray-600 dark:text-slate-400">${totalFCL} ${t('containers_unit')}</p>`;
    
    const deliveryDateKey = Object.keys(containers[0] || {}).find(key => key.toUpperCase().includes('DELIVERY DATE')) || 'DELIVERY DATE AT BYD';
    const headers = [
        { label: '#', key: null }, { label: 'Contêiner', key: 'CNTRS ORIGINAL' },
        { label: 'BL', key: 'BL' }, { label: 'PO SAP', key: 'PO SAP' },
        { label: 'Navio', key: 'ARRIVAL VESSEL' }, { label: 'Armazém', key: 'BONDED WAREHOUSE' },
        { label: 'Armador', key: 'SHIPOWNER' }, { label: 'Tipo Carga', key: 'LOADING TYPE' },
        { label: 'Status Armazém', key: 'STATUS CNTR WAREHOUSE' }, { label: 'Data Entrega', key: deliveryDateKey, type: 'date' },
        { label: 'Prazo Retorno', key: 'DEADLINE RETURN CNTR', type: 'date' }, { label: 'Dias Restantes', key: 'daysToDeadline', type: 'number' },
        { label: 'Status Final', key: 'STATUS' },
    ];

    const sortedContainers = [...containers];
    const { column, direction } = modalSortState;
    if (column) {
        sortedContainers.sort((a, b) => {
            const headerDef = headers.find(h => h.key === column);
            let valA = a[column]; let valB = b[column];
            if (headerDef?.type === 'date') {
                valA = excelDateToJSDate(valA)?.getTime() ?? (direction === 'asc' ? Infinity : -Infinity);
                valB = excelDateToJSDate(valB)?.getTime() ?? (direction === 'asc' ? Infinity : -Infinity);
            }
            if (valA === null || valA === undefined) return 1;
            if (valB === null || valB === undefined) return -1;
            if (valA < valB) return direction === 'asc' ? -1 : 1;
            if (valA > valB) return direction === 'asc' ? 1 : -1;
            return 0;
        });
    }

    const headerHtml = headers.map(h => {
        if (!h.key) return `<th class="px-2 py-2 text-left font-semibold text-gray-600 dark:text-slate-300 text-xs uppercase">${h.label}</th>`;
        const isSorted = modalSortState.column === h.key;
        const sortIcon = isSorted ? (modalSortState.direction === 'asc' ? 'fa-sort-up' : 'fa-sort-down') : 'fa-sort';
        return `<th class="px-2 py-2 text-left font-semibold text-gray-600 dark:text-slate-300 text-xs uppercase sortable-header ${isSorted ? 'sorted' : ''}" data-key="${h.key}">${h.label} <i class="fas ${sortIcon} sort-icon"></i></th>`;
    }).join('');

    const tableRows = sortedContainers.map((c: any, i: number) => {
        const daysText = c.risk === 'none' ? `<span class="font-semibold text-green-700">${t('modal_delivered_status')}</span>` : c.daysToDeadline !== null ? `<span class="font-bold ${c.daysToDeadline < 0 ? 'text-red-600' : ''}">${c.daysToDeadline}</span>` : 'N/A';
        return `<tr class="row-risk-${c.risk}">
            <td class="px-2 py-1.5 text-xs text-center">${i + 1}</td>
            <td class="px-2 py-1.5 text-xs font-semibold">${c['CNTRS ORIGINAL'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['BL'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['PO SAP'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['ARRIVAL VESSEL'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['BONDED WAREHOUSE'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['SHIPOWNER'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['LOADING TYPE'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${c['STATUS CNTR WAREHOUSE'] || ''}</td>
            <td class="px-2 py-1.5 text-xs">${formatDate(c[deliveryDateKey])}</td>
            <td class="px-2 py-1.5 text-xs">${formatDate(c['DEADLINE RETURN CNTR'])}</td>
            <td class="px-2 py-1.5 text-xs text-center">${daysText}</td>
            <td class="px-2 py-1.5 text-xs">${c['STATUS'] || ''}</td>
        </tr>`;
    }).join('');

    modalBody.innerHTML = `<div class="table-responsive"><table class="min-w-full text-sm">
        <thead class="bg-gray-100 dark:bg-slate-700"><tr class="border-b dark:border-slate-600">${headerHtml}</tr></thead>
        <tbody class="divide-y divide-gray-200 dark:divide-slate-700">${tableRows}</tbody>
    </table></div>`;
    modalBody.querySelector('thead')?.addEventListener('click', handleModalSort);
}