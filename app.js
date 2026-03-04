/* =============================================
   Fleet Maintenance Dashboard - Application Logic
   ============================================= */

// Data Management
const DB_KEY = 'fleet_maintenance_db';
const AUTOSAVE_KEY = 'fleet_autosave_enabled';
const SUGGESTIONS_KEY = 'fleet_suggestions';
let currentVehicleId = null;
let autoSaveEnabled = localStorage.getItem(AUTOSAVE_KEY) !== 'false'; // enabled by default

// Sorting state per table
const tableSortState = {
    dashboard: { column: null, direction: 'asc' },
    vehicle: { column: null, direction: 'asc' },
    allservices: { column: null, direction: 'asc' }
};

// Initialize the application
document.addEventListener('DOMContentLoaded', () => {
    initializeDB();
    updateCurrentDate();
    renderDashboard();
    setupNavigation();
    setDefaultServiceDate();
    updateAutoSaveToggle();
    populateAutocompleteSuggestions();
});

// Database Functions
function initializeDB() {
    // Try to load pre-loaded fleet data first
    if (typeof initializeWithPreloadedData === 'function') {
        initializeWithPreloadedData();
    }

    // Fallback: create empty database if nothing exists
    if (!localStorage.getItem(DB_KEY)) {
        const initialDB = {
            vehicles: [],
            lastUpdated: new Date().toISOString()
        };
        localStorage.setItem(DB_KEY, JSON.stringify(initialDB));
    }
}

function getDB() {
    return JSON.parse(localStorage.getItem(DB_KEY));
}

function saveDB(db) {
    db.lastUpdated = new Date().toISOString();
    localStorage.setItem(DB_KEY, JSON.stringify(db));
}

function generateId() {
    return 'id_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

// Auto-Save Functions
function autoSaveBackup(reason = 'change') {
    if (!autoSaveEnabled) return;

    // Data is already saved to localStorage via saveDB()
    // Just update the indicator to show save was successful
    try {
        console.log(`Data saved to local storage (${reason})`);
        updateAutoSaveIndicator(true);
    } catch (error) {
        console.error('Auto-save indicator update failed:', error);
        updateAutoSaveIndicator(false);
    }
}

function toggleAutoSave() {
    autoSaveEnabled = !autoSaveEnabled;
    localStorage.setItem(AUTOSAVE_KEY, autoSaveEnabled);
    updateAutoSaveToggle();
    showToast(autoSaveEnabled ? 'Auto-save enabled' : 'Auto-save disabled');
}

function updateAutoSaveToggle() {
    const toggle = document.getElementById('autosave-toggle');
    if (toggle) {
        toggle.classList.toggle('active', autoSaveEnabled);
        toggle.innerHTML = autoSaveEnabled
            ? '<i class="fas fa-save"></i> Auto-Save ON'
            : '<i class="fas fa-save"></i> Auto-Save OFF';
    }
}

function updateAutoSaveIndicator(success) {
    const indicator = document.getElementById('autosave-indicator');
    if (indicator) {
        indicator.style.display = 'block';
        indicator.className = 'autosave-indicator ' + (success ? 'success' : 'error');
        indicator.innerHTML = success
            ? '<i class="fas fa-check"></i> Saved'
            : '<i class="fas fa-times"></i> Failed';

        setTimeout(() => {
            indicator.style.display = 'none';
        }, 3000);
    }
}

// Autocomplete Suggestions Functions
function getSuggestions() {
    const stored = localStorage.getItem(SUGGESTIONS_KEY);
    if (stored) {
        return JSON.parse(stored);
    }
    return { work: [], provider: [], location: [] };
}

function saveSuggestions(suggestions) {
    localStorage.setItem(SUGGESTIONS_KEY, JSON.stringify(suggestions));
}

function addSuggestion(type, value) {
    if (!value || value.trim() === '') return;

    const suggestions = getSuggestions();
    const trimmedValue = value.trim();

    // Don't add duplicates (case insensitive check)
    const exists = suggestions[type].some(
        s => s.toLowerCase() === trimmedValue.toLowerCase()
    );

    if (!exists) {
        suggestions[type].push(trimmedValue);
        // Keep only last 500 unique suggestions per type
        if (suggestions[type].length > 500) {
            suggestions[type] = suggestions[type].slice(-500);
        }
        saveSuggestions(suggestions);
    }
}

function populateAutocompleteSuggestions() {
    const db = getDB();
    const suggestions = getSuggestions();

    // Collect unique values from database
    const workSet = new Set(suggestions.work);
    const providerSet = new Set(suggestions.provider);
    const locationSet = new Set(suggestions.location);

    db.vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            if (service.work) workSet.add(service.work);
            if (service.provider) providerSet.add(service.provider);
            if (service.location) locationSet.add(service.location);
        });
    });

    // Update datalists
    const workDatalist = document.getElementById('work-suggestions');
    const providerDatalist = document.getElementById('provider-suggestions');
    const locationDatalist = document.getElementById('location-suggestions');

    if (workDatalist) {
        // Sort by frequency and alphabetically
        const workArray = [...workSet].sort();
        workDatalist.innerHTML = workArray.map(w => `<option value="${escapeHtml(w)}">`).join('');
    }

    if (providerDatalist) {
        const providerArray = [...providerSet].sort();
        providerDatalist.innerHTML = providerArray.map(p => `<option value="${escapeHtml(p)}">`).join('');
    }

    if (locationDatalist) {
        const locationArray = [...locationSet].sort();
        locationDatalist.innerHTML = locationArray.map(l => `<option value="${escapeHtml(l)}">`).join('');
    }

    // Save merged suggestions back
    suggestions.work = [...workSet];
    suggestions.provider = [...providerSet];
    suggestions.location = [...locationSet];
    saveSuggestions(suggestions);
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function saveNewServiceSuggestions(work, provider, location) {
    addSuggestion('work', work);
    addSuggestion('provider', provider);
    addSuggestion('location', location);
    // Refresh datalists
    populateAutocompleteSuggestions();
}

// Navigation
function setupNavigation() {
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const page = item.dataset.page;
            showPage(page);
        });
    });
}

function showPage(pageName) {
    // Update navigation
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
        if (item.dataset.page === pageName) {
            item.classList.add('active');
        }
    });

    // Update page visibility
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });

    const pageElement = document.getElementById(`${pageName}-page`);
    if (pageElement) {
        pageElement.classList.add('active');
    }

    // Update page title
    const titles = {
        'dashboard': 'Dashboard',
        'vehicles': 'Vehicles',
        'vehicle-detail': 'Vehicle Details',
        'services': 'All Services',
        'reports': 'Reports & Analytics',
        'quicklog': 'Quick Log'
    };
    document.getElementById('page-title').textContent = titles[pageName] || 'Dashboard';

    // Render page content
    switch (pageName) {
        case 'dashboard':
            renderDashboard();
            break;
        case 'vehicles':
            renderVehiclesPage();
            break;
        case 'services':
            renderAllServicesPage();
            break;
        case 'reports':
            renderReportsPage();
            break;
    }
}

// Dashboard Services State
let dashboardServicesPage = 1;
const servicesPerPage = 50;
let filteredDashboardServices = [];

// Dashboard Rendering
function renderDashboard() {
    const db = getDB();
    const vehicles = db.vehicles;

    // Populate quick add vehicle dropdown
    populateQuickAddVehicles(vehicles);

    // Populate filters and render services
    populateDashboardFilters(vehicles);
    filterDashboardServices();

    // Render vehicles list
    const vehiclesList = document.getElementById('dashboard-vehicles-list');
    if (vehicles.length === 0) {
        vehiclesList.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-car"></i>
                <h3>No Vehicles Yet</h3>
                <p>Add your first vehicle to get started</p>
                <button class="btn-primary" onclick="openModal('add-vehicle-modal')">
                    <i class="fas fa-plus"></i> Add Vehicle
                </button>
            </div>
        `;
    } else {
        vehiclesList.innerHTML = vehicles.map(vehicle => {
            const latestMileage = getLatestMileage(vehicle);
            const lastService = vehicle.services.length > 0
                ? vehicle.services.sort((a, b) => new Date(b.date) - new Date(a.date))[0]
                : null;

            return `
                <div class="vehicle-item" onclick="openVehicleDetail('${vehicle.id}')">
                    <div class="vehicle-icon">
                        <i class="fas fa-truck"></i>
                    </div>
                    <div class="vehicle-info">
                        <div class="vehicle-reg">${vehicle.reg}</div>
                        <div class="vehicle-meta">
                            ${vehicle.make || ''} ${vehicle.model || ''} ${vehicle.year ? `(${vehicle.year})` : ''}
                            ${lastService ? `• Last: ${formatDate(lastService.date)}` : '• No services yet'}
                        </div>
                    </div>
                    <div class="vehicle-mileage">
                        <div class="value">${formatNumber(latestMileage)}</div>
                        <div class="label">miles</div>
                    </div>
                </div>
            `;
        }).join('');
    }

    // Render recent services
    const recentServicesList = document.getElementById('recent-services-list');
    const allServices = [];
    vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            allServices.push({ ...service, vehicleReg: vehicle.reg, vehicleId: vehicle.id });
        });
    });

    const recentServices = allServices
        .sort((a, b) => new Date(b.createdAt || b.date) - new Date(a.createdAt || a.date))
        .slice(0, 5);

    if (recentServices.length === 0) {
        recentServicesList.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-wrench"></i>
                <p>No services recorded yet</p>
            </div>
        `;
    } else {
        recentServicesList.innerHTML = recentServices.map(service => {
            const date = new Date(service.date);
            return `
                <div class="service-item">
                    <div class="service-date">
                        <div class="day">${date.getDate()}</div>
                        <div class="month">${date.toLocaleString('en', { month: 'short' })}</div>
                    </div>
                    <div class="service-info">
                        <div class="service-title">${truncateText(service.work, 50)}</div>
                        <div class="service-vehicle">${service.vehicleReg} • ${service.provider}</div>
                    </div>
                </div>
            `;
        }).join('');
    }
}

// Quick Add Service Functions
function populateQuickAddVehicles(vehicles) {
    const datalist = document.getElementById('vehicle-suggestions');
    if (!datalist) return;

    const sortedVehicles = [...vehicles].sort((a, b) => a.reg.localeCompare(b.reg));
    datalist.innerHTML = sortedVehicles.map(vehicle =>
        `<option value="${vehicle.reg}">`
    ).join('');

    // Set today's date as default
    const dateInput = document.getElementById('quick-date');
    if (dateInput && !dateInput.value) {
        dateInput.value = new Date().toISOString().split('T')[0];
    }
}

function quickAddService(event) {
    event.preventDefault();

    const vehicleReg = document.getElementById('quick-vehicle').value.trim().toUpperCase();
    const date = document.getElementById('quick-date').value;
    let work = document.getElementById('quick-work').value.trim();
    const part = document.getElementById('quick-part').value.trim();
    const provider = document.getElementById('quick-provider').value.trim() || 'Unknown';
    const mileage = document.getElementById('quick-mileage').value || 0;
    const location = document.getElementById('quick-location').value.trim() || '';

    if (!vehicleReg || !date || !work) {
        showToast('Please fill in Vehicle, Date and Work fields', 'error');
        return;
    }

    // Combine work with part if provided (e.g., "Tire - FP")
    if (part) {
        work = `${work} - ${part}`;
    }

    const db = getDB();
    // Find vehicle by registration (case insensitive)
    const vehicle = db.vehicles.find(v => v.reg.toUpperCase() === vehicleReg);

    if (!vehicle) {
        showToast(`Vehicle "${vehicleReg}" not found`, 'error');
        return;
    }

    const newService = {
        id: generateId(),
        date: date,
        work: work,
        duration: '',
        provider: provider,
        mileage: parseInt(mileage) || 0,
        location: location,
        cost: '',
        notes: '',
        createdAt: new Date().toISOString()
    };

    vehicle.services.push(newService);
    saveDB(db);

    // Save suggestions for autocomplete
    saveNewServiceSuggestions(work, provider, location);

    // Clear form except vehicle and date
    document.getElementById('quick-work').value = '';
    document.getElementById('quick-part').value = '';
    document.getElementById('quick-provider').value = '';
    document.getElementById('quick-mileage').value = '';
    document.getElementById('quick-location').value = '';

    showToast(`Service added to ${vehicle.reg}!`);
    autoSaveBackup('service_added');
    renderDashboard();
}

// Dashboard Services Filters
function populateDashboardFilters(vehicles) {
    const vehicleFilter = document.getElementById('dash-vehicle-filter');
    const workFilter = document.getElementById('dash-work-filter');
    const providerFilter = document.getElementById('dash-provider-filter');

    if (!vehicleFilter) return;

    // Save current values
    const currentVehicle = vehicleFilter.value;
    const currentWork = workFilter.value;
    const currentProvider = providerFilter.value;

    // Populate vehicles
    vehicleFilter.innerHTML = '<option value="all">All Vehicles</option>';
    const sortedVehicles = [...vehicles].sort((a, b) => a.reg.localeCompare(b.reg));
    sortedVehicles.forEach(vehicle => {
        vehicleFilter.innerHTML += `<option value="${vehicle.id}">${vehicle.reg}</option>`;
    });

    // Collect unique work types and providers
    const workTypes = new Set();
    const providers = new Set();

    vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            if (service.work) workTypes.add(service.work);
            if (service.provider) providers.add(service.provider);
        });
    });

    // Populate work types
    workFilter.innerHTML = '<option value="all">All Work Types</option>';
    [...workTypes].sort().forEach(work => {
        const shortWork = work.length > 30 ? work.substring(0, 30) + '...' : work;
        workFilter.innerHTML += `<option value="${work}">${shortWork}</option>`;
    });

    // Populate providers
    providerFilter.innerHTML = '<option value="all">All Providers</option>';
    [...providers].sort().forEach(provider => {
        const shortProvider = provider.length > 25 ? provider.substring(0, 25) + '...' : provider;
        providerFilter.innerHTML += `<option value="${provider}">${shortProvider}</option>`;
    });

    // Restore values
    if (currentVehicle) vehicleFilter.value = currentVehicle;
    if (currentWork) workFilter.value = currentWork;
    if (currentProvider) providerFilter.value = currentProvider;
}

function filterDashboardServices() {
    const db = getDB();
    const vehicleFilter = document.getElementById('dash-vehicle-filter')?.value || 'all';
    const workFilter = document.getElementById('dash-work-filter')?.value || 'all';
    const providerFilter = document.getElementById('dash-provider-filter')?.value || 'all';
    const dateFrom = document.getElementById('dash-date-from')?.value;
    const dateTo = document.getElementById('dash-date-to')?.value;

    // Collect all services
    let allServices = [];
    db.vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            allServices.push({
                ...service,
                vehicleReg: vehicle.reg,
                vehicleId: vehicle.id
            });
        });
    });

    // Apply filters
    if (vehicleFilter !== 'all') {
        allServices = allServices.filter(s => s.vehicleId === vehicleFilter);
    }
    if (workFilter !== 'all') {
        allServices = allServices.filter(s => s.work === workFilter);
    }
    if (providerFilter !== 'all') {
        allServices = allServices.filter(s => s.provider === providerFilter);
    }
    if (dateFrom) {
        allServices = allServices.filter(s => s.date >= dateFrom);
    }
    if (dateTo) {
        allServices = allServices.filter(s => s.date <= dateTo);
    }

    // Sort by date descending (newest first) unless user sorted by column
    const dashSortState = tableSortState.dashboard;
    if (dashSortState.column) {
        sortServiceArray(allServices, dashSortState.column, dashSortState.direction);
    } else {
        allServices.sort((a, b) => new Date(b.date) - new Date(a.date));
    }

    filteredDashboardServices = allServices;
    dashboardServicesPage = 1;
    renderDashboardServicesPage();
}

function renderDashboardServicesPage() {
    const servicesList = document.getElementById('dashboard-services-list');
    const servicesCount = document.getElementById('services-count');
    const pageInfo = document.getElementById('page-info');
    const prevBtn = document.getElementById('prev-page-btn');
    const nextBtn = document.getElementById('next-page-btn');

    if (!servicesList) return;

    const totalServices = filteredDashboardServices.length;
    const totalPages = Math.ceil(totalServices / servicesPerPage) || 1;
    const start = (dashboardServicesPage - 1) * servicesPerPage;
    const end = start + servicesPerPage;
    const pageServices = filteredDashboardServices.slice(start, end);

    // Update count and pagination info
    servicesCount.textContent = `${totalServices} services`;
    pageInfo.textContent = `Page ${dashboardServicesPage} of ${totalPages}`;
    prevBtn.disabled = dashboardServicesPage <= 1;
    nextBtn.disabled = dashboardServicesPage >= totalPages;

    if (pageServices.length === 0) {
        servicesList.innerHTML = `
            <tr>
                <td colspan="8" style="text-align: center; padding: 30px;">
                    <div class="empty-state">
                        <i class="fas fa-filter"></i>
                        <p>No services match the filters</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    servicesList.innerHTML = pageServices.map(service => `
        <tr>
            <td><strong>${service.vehicleReg}</strong></td>
            <td>${formatDate(service.date)}</td>
            <td title="${service.work}">${truncateText(service.work, 30)}</td>
            <td title="${service.provider}">${truncateText(service.provider, 20)}</td>
            <td>${service.mileage ? formatNumber(service.mileage) : '-'}</td>
            <td>${truncateText(service.location || '-', 15)}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-icon" onclick="editService('${service.vehicleId}', '${service.id}')" title="Edit">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn-icon danger" onclick="confirmDeleteService('${service.vehicleId}', '${service.id}')" title="Delete">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </td>
        </tr>
    `).join('');
}

function prevServicesPage() {
    if (dashboardServicesPage > 1) {
        dashboardServicesPage--;
        renderDashboardServicesPage();
    }
}

function nextServicesPage() {
    const totalPages = Math.ceil(filteredDashboardServices.length / servicesPerPage);
    if (dashboardServicesPage < totalPages) {
        dashboardServicesPage++;
        renderDashboardServicesPage();
    }
}

function clearDashboardFilters() {
    document.getElementById('dash-vehicle-filter').value = 'all';
    document.getElementById('dash-work-filter').value = 'all';
    document.getElementById('dash-provider-filter').value = 'all';
    document.getElementById('dash-date-from').value = '';
    document.getElementById('dash-date-to').value = '';
    filterDashboardServices();
}

// Column Sorting
function sortTable(thElement, tableId) {
    const column = thElement.dataset.sort;
    const state = tableSortState[tableId];

    // Toggle direction or set new column
    if (state.column === column) {
        state.direction = state.direction === 'asc' ? 'desc' : 'asc';
    } else {
        state.column = column;
        state.direction = 'asc';
    }

    // Update header icons
    const thead = thElement.closest('thead');
    thead.querySelectorAll('th.sortable').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        const icon = th.querySelector('i');
        if (icon) icon.className = 'fas fa-sort';
    });
    thElement.classList.add(state.direction === 'asc' ? 'sort-asc' : 'sort-desc');
    const icon = thElement.querySelector('i');
    if (icon) icon.className = state.direction === 'asc' ? 'fas fa-sort-up' : 'fas fa-sort-down';

    // Sort the data and re-render
    if (tableId === 'dashboard') {
        sortServiceArray(filteredDashboardServices, column, state.direction);
        dashboardServicesPage = 1;
        renderDashboardServicesPage();
    } else if (tableId === 'vehicle') {
        sortVehicleServices(column, state.direction);
    } else if (tableId === 'allservices') {
        sortAllServices(column, state.direction);
    }
}

function sortCompare(a, b, column, direction) {
    let valA, valB;

    if (column === 'date') {
        valA = new Date(a.date || 0);
        valB = new Date(b.date || 0);
    } else if (column === 'mileage') {
        valA = parseInt(a.mileage) || 0;
        valB = parseInt(b.mileage) || 0;
    } else {
        valA = (a[column] || '').toString().toLowerCase();
        valB = (b[column] || '').toString().toLowerCase();
    }

    let result;
    if (valA < valB) result = -1;
    else if (valA > valB) result = 1;
    else result = 0;

    return direction === 'desc' ? -result : result;
}

function sortServiceArray(arr, column, direction) {
    arr.sort((a, b) => sortCompare(a, b, column, direction));
}

function sortVehicleServices(column, direction) {
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === currentVehicleId);
    if (!vehicle) return;

    const filter = document.getElementById('service-filter').value;
    let services = [...vehicle.services];
    const now = new Date();
    if (filter === 'month') {
        services = services.filter(s => {
            const d = new Date(s.date);
            return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        });
    } else if (filter === 'year') {
        services = services.filter(s => new Date(s.date).getFullYear() === now.getFullYear());
    }

    sortServiceArray(services, column, direction);

    const servicesList = document.getElementById('vehicle-services-list');
    if (services.length === 0) return;

    servicesList.innerHTML = services.map(service => `
        <tr>
            <td>${formatDate(service.date)}</td>
            <td>${service.work}</td>
            <td>${service.provider}</td>
            <td>${formatNumber(service.mileage)}</td>
            <td>${service.location}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-icon" onclick="editService('${vehicle.id}', '${service.id}')" title="Edit">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn-icon danger" onclick="confirmDeleteService('${vehicle.id}', '${service.id}')" title="Delete">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </td>
        </tr>
    `).join('');
}

function sortAllServices(column, direction) {
    const db = getDB();
    const vehicleFilter = document.getElementById('vehicle-filter').value;
    const dateFrom = document.getElementById('date-from').value;
    const dateTo = document.getElementById('date-to').value;

    let allServices = [];
    db.vehicles.forEach(vehicle => {
        if (vehicleFilter === 'all' || vehicleFilter === vehicle.id) {
            vehicle.services.forEach(service => {
                allServices.push({ ...service, vehicleReg: vehicle.reg, vehicleId: vehicle.id });
            });
        }
    });

    if (dateFrom) allServices = allServices.filter(s => new Date(s.date) >= new Date(dateFrom));
    if (dateTo) allServices = allServices.filter(s => new Date(s.date) <= new Date(dateTo));

    sortServiceArray(allServices, column, direction);

    const servicesList = document.getElementById('all-services-list');
    if (allServices.length === 0) return;

    servicesList.innerHTML = allServices.map(service => `
        <tr>
            <td><strong>${service.vehicleReg}</strong></td>
            <td>${formatDate(service.date)}</td>
            <td>${service.work}</td>
            <td>${service.provider}</td>
            <td>${formatNumber(service.mileage)}</td>
            <td>${service.location}</td>
        </tr>
    `).join('');
}

// Vehicles Page
function renderVehiclesPage() {
    const db = getDB();
    const vehicles = db.vehicles;
    const vehiclesGrid = document.getElementById('vehicles-grid');

    // Update count
    const countEl = document.getElementById('vehicles-count');
    if (countEl) countEl.textContent = `(${vehicles.length})`;

    if (vehicles.length === 0) {
        vehiclesGrid.innerHTML = `
            <div class="empty-state" style="grid-column: 1/-1;">
                <i class="fas fa-car"></i>
                <h3>No Vehicles Yet</h3>
                <p>Add your first vehicle to start tracking maintenance</p>
                <button class="btn-primary" onclick="openModal('add-vehicle-modal')">
                    <i class="fas fa-plus"></i> Add Vehicle
                </button>
            </div>
        `;
        return;
    }

    vehiclesGrid.innerHTML = vehicles.map(vehicle => {
        const latestMileage = getLatestMileage(vehicle);
        const serviceCount = vehicle.services.length;
        const lastService = serviceCount > 0
            ? vehicle.services.sort((a, b) => new Date(b.date) - new Date(a.date))[0]
            : null;

        return `
            <div class="vehicle-card" onclick="openVehicleDetail('${vehicle.id}')">
                <div class="vehicle-card-header">
                    <div class="vehicle-card-icon">
                        <i class="fas fa-truck"></i>
                    </div>
                    <div class="vehicle-card-title">
                        <h3>${vehicle.reg}</h3>
                        <span>${vehicle.make || 'Unknown'} ${vehicle.model || ''}</span>
                    </div>
                </div>
                <div class="vehicle-card-stats">
                    <div class="vehicle-card-stat">
                        <div class="value">${formatNumber(latestMileage)}</div>
                        <div class="label">Current Mileage</div>
                    </div>
                    <div class="vehicle-card-stat">
                        <div class="value">${serviceCount}</div>
                        <div class="label">Total Services</div>
                    </div>
                </div>
                <div class="vehicle-card-footer">
                    <span class="last-service">
                        ${lastService ? `Last: ${formatDate(lastService.date)}` : 'No services yet'}
                    </span>
                    <div class="vehicle-card-actions" onclick="event.stopPropagation()">
                        <button class="btn-icon" onclick="openVehicleDetail('${vehicle.id}')" title="View Details">
                            <i class="fas fa-eye"></i>
                        </button>
                        <button class="btn-icon danger" onclick="confirmDeleteVehicle('${vehicle.id}')" title="Delete">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </div>
            </div>
        `;
    }).join('');
}

// Vehicle Detail Page
function openVehicleDetail(vehicleId) {
    currentVehicleId = vehicleId;
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === vehicleId);

    if (!vehicle) {
        showToast('Vehicle not found', 'error');
        return;
    }

    // Update header
    document.getElementById('vehicle-detail-reg').textContent = vehicle.reg;
    document.getElementById('vehicle-detail-type').textContent =
        `${vehicle.make || 'Unknown'} ${vehicle.model || ''} ${vehicle.year ? `(${vehicle.year})` : ''}`;

    // Render vehicle stats
    const latestMileage = getLatestMileage(vehicle);
    const totalServices = vehicle.services.length;
    const lastServiceDate = totalServices > 0
        ? formatDate(vehicle.services.sort((a, b) => new Date(b.date) - new Date(a.date))[0].date)
        : 'N/A';

    document.getElementById('vehicle-stats').innerHTML = `
        <div class="vehicle-stat-card">
            <i class="fas fa-road"></i>
            <span class="value">${formatNumber(latestMileage)}</span>
            <span class="label">Current Mileage</span>
        </div>
        <div class="vehicle-stat-card">
            <i class="fas fa-wrench"></i>
            <span class="value">${totalServices}</span>
            <span class="label">Total Services</span>
        </div>
        <div class="vehicle-stat-card">
            <i class="fas fa-calendar"></i>
            <span class="value">${lastServiceDate}</span>
            <span class="label">Last Service</span>
        </div>
        <div class="vehicle-stat-card">
            <i class="fas fa-info-circle"></i>
            <span class="value">${vehicle.color || 'N/A'}</span>
            <span class="label">Color</span>
        </div>
    `;

    // Render services
    renderVehicleServices(vehicle);

    // Render service intervals
    renderServiceIntervals(vehicle);

    showPage('vehicle-detail');
}

// Calculate and render service intervals
function renderServiceIntervals(vehicle) {
    const container = document.getElementById('service-intervals');

    if (vehicle.services.length < 2) {
        container.innerHTML = `
            <div class="no-intervals">
                <i class="fas fa-info-circle"></i>
                <p>Need at least 2 services to calculate intervals</p>
            </div>
        `;
        return;
    }

    // Service type patterns with icons
    const servicePatterns = [
        {
            name: 'Tyres / Tires',
            icon: 'fa-circle',
            patterns: ['tyre', 'tire', 'tyres', 'tires', 'wheel'],
            color: 'primary'
        },
        {
            name: 'Brake Pads',
            icon: 'fa-compact-disc',
            patterns: ['brake pad', 'pads', 'front pad', 'rear pad'],
            color: 'danger'
        },
        {
            name: 'Brake Discs',
            icon: 'fa-dot-circle',
            patterns: ['disc', 'discs', 'rotor'],
            color: 'warning'
        },
        {
            name: 'Oil Service',
            icon: 'fa-oil-can',
            patterns: ['oil', 'oil service', 'oil change'],
            color: 'success'
        },
        {
            name: 'Clutch',
            icon: 'fa-cog',
            patterns: ['clutch', 'flywheel'],
            color: 'purple'
        },
        {
            name: 'Mirrors',
            icon: 'fa-car-side',
            patterns: ['mirror', 'wing mirror', 'big mirror', 'mini mirror'],
            color: 'primary'
        },
        {
            name: 'Indicators',
            icon: 'fa-lightbulb',
            patterns: ['indicator', 'signal'],
            color: 'warning'
        },
        {
            name: 'CV Joint / Drive Shaft',
            icon: 'fa-cogs',
            patterns: ['cv joint', 'drive shaft', 'driveshaft', 'cv'],
            color: 'danger'
        },
        {
            name: 'Ball Joint / Bearing',
            icon: 'fa-circle-notch',
            patterns: ['ball joint', 'bearing', 'ball bearing'],
            color: 'primary'
        },
        {
            name: 'Bodywork',
            icon: 'fa-car-crash',
            patterns: ['bodywork', 'body', 'bumper', 'panel', 'dent'],
            color: 'warning'
        }
    ];

    // Sort services by date
    const sortedServices = [...vehicle.services].sort((a, b) => new Date(a.date) - new Date(b.date));

    // Group services by type
    const serviceGroups = {};

    sortedServices.forEach(service => {
        const workLower = service.work.toLowerCase();

        servicePatterns.forEach(pattern => {
            const matches = pattern.patterns.some(p => workLower.includes(p));
            if (matches) {
                if (!serviceGroups[pattern.name]) {
                    serviceGroups[pattern.name] = {
                        ...pattern,
                        services: []
                    };
                }
                serviceGroups[pattern.name].services.push(service);
            }
        });
    });

    // Calculate intervals for each group
    const intervalCards = [];

    Object.values(serviceGroups).forEach(group => {
        if (group.services.length < 1) return;

        const services = group.services;
        const intervals = [];

        // Calculate intervals between consecutive services
        for (let i = 1; i < services.length; i++) {
            const prev = services[i - 1];
            const curr = services[i];

            const prevDate = new Date(prev.date);
            const currDate = new Date(curr.date);
            const daysDiff = Math.round((currDate - prevDate) / (1000 * 60 * 60 * 24));

            let milesDiff = 0;
            if (prev.mileage && curr.mileage && curr.mileage > prev.mileage) {
                milesDiff = curr.mileage - prev.mileage;
            }

            intervals.push({
                days: daysDiff,
                miles: milesDiff
            });
        }

        // Calculate averages
        const avgDays = intervals.length > 0
            ? Math.round(intervals.reduce((sum, i) => sum + i.days, 0) / intervals.length)
            : 0;

        const milesWithData = intervals.filter(i => i.miles > 0);
        const avgMiles = milesWithData.length > 0
            ? Math.round(milesWithData.reduce((sum, i) => sum + i.miles, 0) / milesWithData.length)
            : 0;

        // Get last service info
        const lastService = services[services.length - 1];
        const lastServiceDate = new Date(lastService.date);
        const daysSinceLast = Math.round((new Date() - lastServiceDate) / (1000 * 60 * 60 * 24));

        // Calculate current mileage and miles since last service
        const currentMileage = getLatestMileage(vehicle);
        const milesSinceLast = lastService.mileage ? currentMileage - lastService.mileage : 0;

        // Determine status
        let status = 'success';
        let projection = null;

        if (avgDays > 0 || avgMiles > 0) {
            const daysProgress = avgDays > 0 ? (daysSinceLast / avgDays) * 100 : 0;
            const milesProgress = avgMiles > 0 ? (milesSinceLast / avgMiles) * 100 : 0;
            const progress = Math.max(daysProgress, milesProgress);

            if (progress >= 100) {
                status = 'danger';
                projection = 'Service may be overdue!';
            } else if (progress >= 75) {
                status = 'warning';
                const daysRemaining = avgDays > 0 ? Math.round(avgDays - daysSinceLast) : null;
                const milesRemaining = avgMiles > 0 ? Math.round(avgMiles - milesSinceLast) : null;

                if (daysRemaining && daysRemaining > 0) {
                    projection = `~${daysRemaining} days remaining`;
                } else if (milesRemaining && milesRemaining > 0) {
                    projection = `~${formatNumber(milesRemaining)} miles remaining`;
                }
            }
        }

        intervalCards.push({
            group,
            services,
            avgDays,
            avgMiles,
            lastService,
            daysSinceLast,
            milesSinceLast,
            status,
            projection,
            intervalsCount: intervals.length
        });
    });

    // Sort by status (danger first, then warning, then success)
    const statusOrder = { danger: 0, warning: 1, success: 2 };
    intervalCards.sort((a, b) => statusOrder[a.status] - statusOrder[b.status]);

    if (intervalCards.length === 0) {
        container.innerHTML = `
            <div class="no-intervals">
                <i class="fas fa-info-circle"></i>
                <p>No recurring services found to calculate intervals</p>
            </div>
        `;
        return;
    }

    container.innerHTML = intervalCards.map(card => `
        <div class="interval-card ${card.status}">
            <div class="interval-header">
                <div class="interval-type">
                    <i class="fas ${card.group.icon}"></i>
                    <span>${card.group.name}</span>
                </div>
                <span class="interval-count">${card.services.length} services</span>
            </div>
            <div class="interval-stats">
                <div class="interval-stat ${card.avgMiles > 0 ? 'highlight' : ''}">
                    <span class="value">${card.avgMiles > 0 ? formatNumber(card.avgMiles) : 'N/A'}</span>
                    <span class="label">Avg Miles</span>
                </div>
                <div class="interval-stat">
                    <span class="value">${card.avgDays > 0 ? formatDuration(card.avgDays) : 'N/A'}</span>
                    <span class="label">Avg Time</span>
                </div>
            </div>
            <div class="interval-details">
                <div class="last-service">
                    <span>Last service:</span>
                    <span>${formatDate(card.lastService.date)}</span>
                </div>
                <div class="last-service">
                    <span>Days since:</span>
                    <span>${card.daysSinceLast} days</span>
                </div>
                ${card.milesSinceLast > 0 ? `
                <div class="last-service">
                    <span>Miles since:</span>
                    <span>${formatNumber(card.milesSinceLast)} mi</span>
                </div>
                ` : ''}
                ${card.projection ? `
                <div class="projection ${card.status}">
                    <i class="fas ${card.status === 'danger' ? 'fa-exclamation-triangle' : 'fa-clock'}"></i>
                    <span>${card.projection}</span>
                </div>
                ` : ''}
            </div>
        </div>
    `).join('');
}

// Format duration in days to human readable
function formatDuration(days) {
    if (days < 30) {
        return `${days}d`;
    } else if (days < 365) {
        const months = Math.round(days / 30);
        return `${months}mo`;
    } else {
        const years = (days / 365).toFixed(1);
        return `${years}y`;
    }
}

function renderVehicleServices(vehicle, filter = 'all') {
    const servicesList = document.getElementById('vehicle-services-list');
    let services = [...vehicle.services].sort((a, b) => new Date(b.date) - new Date(a.date));

    // Apply filter
    const now = new Date();
    if (filter === 'month') {
        services = services.filter(s => {
            const d = new Date(s.date);
            return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        });
    } else if (filter === 'year') {
        services = services.filter(s => {
            const d = new Date(s.date);
            return d.getFullYear() === now.getFullYear();
        });
    }

    if (services.length === 0) {
        servicesList.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 40px;">
                    <div class="empty-state">
                        <i class="fas fa-clipboard-list"></i>
                        <p>No services recorded${filter !== 'all' ? ' for this period' : ''}</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    servicesList.innerHTML = services.map(service => `
        <tr>
            <td>${formatDate(service.date)}</td>
            <td>${service.work}</td>
            <td>${service.provider}</td>
            <td>${formatNumber(service.mileage)}</td>
            <td>${service.location}</td>
            <td>
                <div class="table-actions">
                    <button class="btn-icon" onclick="editService('${vehicle.id}', '${service.id}')" title="Edit">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn-icon danger" onclick="confirmDeleteService('${vehicle.id}', '${service.id}')" title="Delete">
                        <i class="fas fa-trash"></i>
                    </button>
                </div>
            </td>
        </tr>
    `).join('');
}

function filterServices() {
    const filter = document.getElementById('service-filter').value;
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === currentVehicleId);
    if (vehicle) {
        renderVehicleServices(vehicle, filter);
    }
}

// All Services Page
function renderAllServicesPage() {
    const db = getDB();

    // Populate vehicle filter
    const vehicleFilter = document.getElementById('vehicle-filter');
    vehicleFilter.innerHTML = '<option value="all">All Vehicles</option>';
    db.vehicles.forEach(vehicle => {
        vehicleFilter.innerHTML += `<option value="${vehicle.id}">${vehicle.reg}</option>`;
    });

    filterAllServices();
}

function filterAllServices() {
    const db = getDB();
    const vehicleFilter = document.getElementById('vehicle-filter').value;
    const dateFrom = document.getElementById('date-from').value;
    const dateTo = document.getElementById('date-to').value;

    let allServices = [];
    db.vehicles.forEach(vehicle => {
        if (vehicleFilter === 'all' || vehicleFilter === vehicle.id) {
            vehicle.services.forEach(service => {
                allServices.push({ ...service, vehicleReg: vehicle.reg, vehicleId: vehicle.id });
            });
        }
    });

    // Apply date filters
    if (dateFrom) {
        allServices = allServices.filter(s => new Date(s.date) >= new Date(dateFrom));
    }
    if (dateTo) {
        allServices = allServices.filter(s => new Date(s.date) <= new Date(dateTo));
    }

    // Sort by date descending (newest first) unless user sorted by column
    const sortState = tableSortState.allservices;
    if (sortState.column) {
        sortServiceArray(allServices, sortState.column, sortState.direction);
    } else {
        allServices.sort((a, b) => new Date(b.date) - new Date(a.date));
    }

    const servicesList = document.getElementById('all-services-list');

    // Update count
    const countEl = document.getElementById('all-services-count');
    if (countEl) countEl.textContent = `(${allServices.length})`;

    if (allServices.length === 0) {
        servicesList.innerHTML = `
            <tr>
                <td colspan="7" style="text-align: center; padding: 40px;">
                    <div class="empty-state">
                        <i class="fas fa-clipboard-list"></i>
                        <p>No services found</p>
                    </div>
                </td>
            </tr>
        `;
        return;
    }

    servicesList.innerHTML = allServices.map(service => `
        <tr>
            <td><strong>${service.vehicleReg}</strong></td>
            <td>${formatDate(service.date)}</td>
            <td>${service.work}</td>
            <td>${service.provider}</td>
            <td>${formatNumber(service.mileage)}</td>
            <td>${service.location}</td>
        </tr>
    `).join('');
}

// Quick Log - Parse WhatsApp-style messages into services
let parsedQuickLogEntries = [];

function resolvePartialReg(partial) {
    const db = getDB();
    const clean = partial.toUpperCase().replace(/\s+/g, '');

    // Direct full match
    const directMatch = db.vehicles.find(v => v.reg === clean);
    if (directMatch) return directMatch.reg;

    // Match by last 3 characters (e.g. "ZVA" -> "LD71ZVA")
    const suffixMatches = db.vehicles.filter(v => v.reg.endsWith(clean));
    if (suffixMatches.length === 1) return suffixMatches[0].reg;

    // Match by partial suffix (e.g. "1ZVA" or "71ZVA")
    const partialMatches = db.vehicles.filter(v => v.reg.includes(clean));
    if (partialMatches.length === 1) return partialMatches[0].reg;

    return null;
}

function parseQuickLogText(input) {
    const entries = [];
    const blocks = input.split(/(?=Date\s*:)/i).filter(b => b.trim());

    for (const block of blocks) {
        const dateMatch = block.match(/Date\s*:\s*(.+)/i);
        const regMatch = block.match(/Reg\s*:\s*(.+)/i);
        const posMatch = block.match(/Position\s*:\s*(.+)/i);
        const wearMatch = block.match(/(?:Wear\s*(?:or|\/)\s*Damage|Damage)\s*:\s*(.+)/i);
        const changedMatch = block.match(/Changed\s*by\s*:\s*(.+)/i);
        const mileageMatch = block.match(/Mileage\s*:\s*(\d+)/i);
        const workMatch = block.match(/(?:Work|Service|Description)\s*:\s*(.+)/i);
        const providerMatch = block.match(/Provider\s*:\s*(.+)/i);
        const locationMatch = block.match(/Location\s*:\s*(.+)/i);

        if (!dateMatch) continue;

        const rawDate = dateMatch[1].trim();
        let dateStr = null;

        let m = rawDate.match(/(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})/);
        if (m) {
            let [, day, month, year] = m;
            year = parseInt(year);
            if (year < 100) year += 2000;
            dateStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        } else {
            m = rawDate.match(/(\d{1,2})[.\/\-](\d{1,2})/);
            if (m) {
                const [, day, month] = m;
                const year = parseInt(month) <= 3 ? 2026 : 2025;
                dateStr = `${year}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
            }
        }

        if (!dateStr) continue;

        const rawReg = regMatch ? regMatch[1].trim().replace(/\*|<.*?>/g, '').trim() : '';
        const fullReg = rawReg ? resolvePartialReg(rawReg) : null;

        let work = '';
        if (workMatch) {
            work = workMatch[1].trim();
        } else if (posMatch) {
            const position = posMatch[1].trim();
            const wear = wearMatch ? wearMatch[1].trim() : '';
            work = `Tyre Change - Position: ${position}`;
            if (wear) work += ` | ${wear}`;
        }
        if (!work) work = 'Service';

        let provider = '';
        if (providerMatch) {
            provider = providerMatch[1].trim();
        } else if (changedMatch) {
            provider = changedMatch[1].trim().replace(/<.*?>/g, '').trim();
        }

        entries.push({
            date: dateStr,
            rawReg: rawReg,
            fullReg: fullReg,
            work: work,
            provider: provider,
            mileage: mileageMatch ? parseInt(mileageMatch[1]) : 0,
            location: locationMatch ? locationMatch[1].trim() : 'On Site',
            notes: wearMatch ? `Wear/Damage: ${wearMatch[1].trim()}` : ''
        });
    }

    return entries;
}

function parseQuickLog() {
    const input = document.getElementById('quicklog-input').value.trim();
    if (!input) {
        showToast('Please paste or type service messages first', 'error');
        return;
    }

    parsedQuickLogEntries = parseQuickLogText(input);

    if (parsedQuickLogEntries.length === 0) {
        showToast('No valid entries found. Check the format.', 'error');
        return;
    }

    // Show preview
    const preview = document.getElementById('quicklog-preview');
    const countEl = document.getElementById('quicklog-preview-count');
    const warningsEl = document.getElementById('quicklog-warnings');
    const listEl = document.getElementById('quicklog-preview-list');

    preview.style.display = 'block';
    countEl.textContent = parsedQuickLogEntries.length;

    const unresolvedCount = parsedQuickLogEntries.filter(e => !e.fullReg).length;
    if (unresolvedCount > 0) {
        warningsEl.innerHTML = `<i class="fas fa-exclamation-triangle"></i> ${unresolvedCount} unresolved registration(s)`;
    } else {
        warningsEl.textContent = '';
    }
    if (warnings.length > 0) {
        warningsEl.innerHTML += (warningsEl.textContent ? ' | ' : '') + warnings.join(', ');
    }

    listEl.innerHTML = parsedQuickLogEntries.map((entry, i) => `
        <tr>
            <td>
                ${entry.fullReg
                    ? `<strong>${entry.fullReg}</strong>`
                    : `<span class="quicklog-status-warn" title="Not found in fleet">${entry.rawReg || '?'} <i class="fas fa-question-circle"></i></span>`
                }
            </td>
            <td>${formatDate(entry.date)}</td>
            <td title="${entry.work}">${truncateText(entry.work, 35)}</td>
            <td>${entry.provider || '-'}</td>
            <td>${entry.mileage ? formatNumber(entry.mileage) : '-'}</td>
            <td>
                ${entry.fullReg
                    ? '<span class="quicklog-status-ok"><i class="fas fa-check"></i> Ready</span>'
                    : '<span class="quicklog-status-err"><i class="fas fa-times"></i> No match</span>'
                }
            </td>
        </tr>
    `).join('');

    showToast(`${parsedQuickLogEntries.length} entries parsed. Review and confirm.`, 'success');
}

function confirmQuickLog() {
    const validEntries = parsedQuickLogEntries.filter(e => e.fullReg);

    if (validEntries.length === 0) {
        showToast('No valid entries to save (all registrations unresolved)', 'error');
        return;
    }

    const msg = `Save ${validEntries.length} service(s) to the fleet?` +
        (validEntries.length < parsedQuickLogEntries.length
            ? `\n\n${parsedQuickLogEntries.length - validEntries.length} entries will be skipped (unresolved reg).`
            : '');

    if (!confirm(msg)) return;

    const db = getDB();
    let added = 0;

    for (const entry of validEntries) {
        const vehicle = db.vehicles.find(v => v.reg === entry.fullReg);
        if (!vehicle) continue;

        // Check for duplicate
        const isDup = vehicle.services.some(s =>
            s.date === entry.date && s.work === entry.work
        );
        if (isDup) continue;

        vehicle.services.push({
            id: generateId(),
            date: entry.date,
            work: entry.work,
            duration: '',
            provider: entry.provider,
            mileage: entry.mileage,
            location: entry.location,
            cost: '',
            notes: entry.notes,
            createdAt: new Date().toISOString()
        });

        if (entry.mileage > 0 && entry.mileage > (vehicle.initialMileage || 0)) {
            vehicle.initialMileage = entry.mileage;
        }

        added++;
    }

    saveDB(db);
    autoSaveBackup('quicklog');

    showToast(`${added} service(s) saved successfully!`, 'success');
    clearQuickLog();
    parsedQuickLogEntries = [];
}

function clearQuickLog() {
    document.getElementById('quicklog-input').value = '';
    document.getElementById('quicklog-preview').style.display = 'none';
    parsedQuickLogEntries = [];
}

function cancelQuickLog() {
    document.getElementById('quicklog-preview').style.display = 'none';
    parsedQuickLogEntries = [];
}

// Dashboard Quick Log (compact version)
function parseDashQuickLog() {
    const input = document.getElementById('dash-quicklog-input').value.trim();
    const statusEl = document.getElementById('dash-quicklog-status');

    if (!input) {
        statusEl.innerHTML = '<span style="color:var(--danger)">Paste messages first</span>';
        return;
    }

    // Parse using shared logic
    parsedQuickLogEntries = parseQuickLogText(input);

    const valid = parsedQuickLogEntries.filter(e => e.fullReg);
    const invalid = parsedQuickLogEntries.length - valid.length;

    if (valid.length === 0) {
        statusEl.innerHTML = '<span style="color:var(--danger)">No valid entries found</span>';
        return;
    }

    let msg = `${valid.length} service(s) found. Save them?`;
    if (invalid > 0) msg += `\n${invalid} skipped (unknown reg).`;

    if (!confirm(msg)) {
        statusEl.innerHTML = '<span style="color:var(--text-muted)">Cancelled</span>';
        return;
    }

    // Save directly
    const db = getDB();
    let added = 0;
    for (const entry of valid) {
        const vehicle = db.vehicles.find(v => v.reg === entry.fullReg);
        if (!vehicle) continue;
        const isDup = vehicle.services.some(s => s.date === entry.date && s.work === entry.work);
        if (isDup) continue;
        vehicle.services.push({
            id: generateId(),
            date: entry.date,
            work: entry.work,
            duration: '',
            provider: entry.provider,
            mileage: entry.mileage,
            location: entry.location,
            cost: '',
            notes: entry.notes,
            createdAt: new Date().toISOString()
        });
        if (entry.mileage > 0 && entry.mileage > (vehicle.initialMileage || 0)) {
            vehicle.initialMileage = entry.mileage;
        }
        added++;
    }

    saveDB(db);
    autoSaveBackup('quicklog');
    document.getElementById('dash-quicklog-input').value = '';
    statusEl.innerHTML = `<span style="color:var(--success)"><i class="fas fa-check"></i> ${added} saved!</span>`;
    parsedQuickLogEntries = [];
    renderDashboard();
}

// Reports Page
function renderReportsPage() {
    const db = getDB();

    // Services by Vehicle Chart
    const chartContainer = document.getElementById('services-by-vehicle-chart');
    if (db.vehicles.length === 0) {
        chartContainer.innerHTML = '<p style="color: var(--text-muted); text-align: center;">No data available</p>';
    } else {
        const maxServices = Math.max(...db.vehicles.map(v => v.services.length), 1);
        chartContainer.innerHTML = db.vehicles.map(vehicle => {
            const height = (vehicle.services.length / maxServices) * 100;
            return `
                <div class="chart-bar" style="height: ${Math.max(height, 5)}%;">
                    <div class="tooltip">${vehicle.reg}: ${vehicle.services.length} services</div>
                </div>
            `;
        }).join('');
    }

    // Services Over Time (last 12 months)
    const timeChart = document.getElementById('services-over-time-chart');
    const monthlyData = {};
    const now = new Date();

    // Initialize last 12 months
    for (let i = 11; i >= 0; i--) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        monthlyData[key] = 0;
    }

    // Count services per month
    db.vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            const d = new Date(service.date);
            const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
            if (monthlyData.hasOwnProperty(key)) {
                monthlyData[key]++;
            }
        });
    });

    const maxMonthly = Math.max(...Object.values(monthlyData), 1);
    timeChart.innerHTML = Object.entries(monthlyData).map(([month, count]) => {
        const height = (count / maxMonthly) * 100;
        const [year, m] = month.split('-');
        const monthName = new Date(year, parseInt(m) - 1).toLocaleString('en', { month: 'short' });
        return `
            <div class="chart-bar" style="height: ${Math.max(height, 5)}%;">
                <div class="tooltip">${monthName} ${year}: ${count} services</div>
            </div>
        `;
    }).join('');

    // Top Providers
    const providersContainer = document.getElementById('top-providers-list');
    const providerCounts = {};
    db.vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            if (service.provider) {
                providerCounts[service.provider] = (providerCounts[service.provider] || 0) + 1;
            }
        });
    });

    const topProviders = Object.entries(providerCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    if (topProviders.length === 0) {
        providersContainer.innerHTML = '<p style="color: var(--text-muted);">No providers recorded yet</p>';
    } else {
        providersContainer.innerHTML = topProviders.map(([name, count], index) => `
            <div class="provider-item">
                <div class="provider-rank">${index + 1}</div>
                <div class="provider-info">
                    <div class="provider-name">${name}</div>
                    <div class="provider-count">${count} service${count > 1 ? 's' : ''}</div>
                </div>
            </div>
        `).join('');
    }

    // Common Services
    const servicesContainer = document.getElementById('common-services-list');
    const serviceCounts = {};
    db.vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            // Extract keywords from service work
            const words = service.work.toLowerCase().split(/\s+/);
            const keywords = ['oil', 'tire', 'brake', 'filter', 'battery', 'inspection', 'repair', 'change', 'check', 'service', 'maintenance'];
            words.forEach(word => {
                if (keywords.some(k => word.includes(k))) {
                    const key = word.charAt(0).toUpperCase() + word.slice(1);
                    serviceCounts[key] = (serviceCounts[key] || 0) + 1;
                }
            });
        });
    });

    const commonServices = Object.entries(serviceCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    if (commonServices.length === 0) {
        servicesContainer.innerHTML = '<p style="color: var(--text-muted);">No services recorded yet</p>';
    } else {
        servicesContainer.innerHTML = commonServices.map(([name, count]) => `
            <div class="service-tag">${name}<span>${count}</span></div>
        `).join('');
    }
}

// Vehicle CRUD Operations
function addVehicle(event) {
    event.preventDefault();

    const db = getDB();
    const reg = document.getElementById('vehicle-reg').value.trim().toUpperCase();

    // Check for duplicate registration
    if (db.vehicles.some(v => v.reg.toUpperCase() === reg)) {
        showToast('A vehicle with this registration already exists', 'error');
        return;
    }

    const newVehicle = {
        id: generateId(),
        reg: reg,
        make: document.getElementById('vehicle-make').value.trim(),
        model: document.getElementById('vehicle-model').value.trim(),
        year: document.getElementById('vehicle-year').value,
        initialMileage: document.getElementById('vehicle-initial-mileage').value || 0,
        color: document.getElementById('vehicle-color').value.trim(),
        notes: document.getElementById('vehicle-notes').value.trim(),
        services: [],
        createdAt: new Date().toISOString()
    };

    db.vehicles.push(newVehicle);
    saveDB(db);

    closeModal('add-vehicle-modal');
    document.getElementById('add-vehicle-form').reset();
    showToast('Vehicle added successfully!');
    renderDashboard();
}

function openEditVehicleModal() {
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === currentVehicleId);

    if (!vehicle) {
        showToast('Vehicle not found', 'error');
        return;
    }

    // Populate form with current values
    document.getElementById('edit-vehicle-id').value = vehicle.id;
    document.getElementById('edit-vehicle-reg').value = vehicle.reg || '';
    document.getElementById('edit-vehicle-make').value = vehicle.make || '';
    document.getElementById('edit-vehicle-model').value = vehicle.model || '';
    document.getElementById('edit-vehicle-year').value = vehicle.year || '';
    document.getElementById('edit-vehicle-initial-mileage').value = vehicle.initialMileage || '';
    document.getElementById('edit-vehicle-color').value = vehicle.color || '';
    document.getElementById('edit-vehicle-notes').value = vehicle.notes || '';

    openModal('edit-vehicle-modal');
}

function updateVehicle(event) {
    event.preventDefault();

    const db = getDB();
    const vehicleId = document.getElementById('edit-vehicle-id').value;
    const vehicleIndex = db.vehicles.findIndex(v => v.id === vehicleId);

    if (vehicleIndex === -1) {
        showToast('Vehicle not found', 'error');
        return;
    }

    const newReg = document.getElementById('edit-vehicle-reg').value.trim().toUpperCase();
    const currentReg = db.vehicles[vehicleIndex].reg;

    // Check for duplicate registration (if changed)
    if (newReg !== currentReg && db.vehicles.some(v => v.reg.toUpperCase() === newReg)) {
        showToast('A vehicle with this registration already exists', 'error');
        return;
    }

    // Update vehicle
    db.vehicles[vehicleIndex] = {
        ...db.vehicles[vehicleIndex],
        reg: newReg,
        make: document.getElementById('edit-vehicle-make').value.trim(),
        model: document.getElementById('edit-vehicle-model').value.trim(),
        year: document.getElementById('edit-vehicle-year').value,
        initialMileage: document.getElementById('edit-vehicle-initial-mileage').value || 0,
        color: document.getElementById('edit-vehicle-color').value.trim(),
        notes: document.getElementById('edit-vehicle-notes').value.trim(),
        updatedAt: new Date().toISOString()
    };

    saveDB(db);
    closeModal('edit-vehicle-modal');
    showToast('Vehicle updated successfully!');
    openVehicleDetail(vehicleId);
}

function confirmDeleteVehicle(vehicleId) {
    document.getElementById('confirm-message').textContent =
        'Are you sure you want to delete this vehicle? All service history will be lost.';

    const confirmBtn = document.getElementById('confirm-action-btn');
    confirmBtn.onclick = () => deleteVehicle(vehicleId);

    openModal('confirm-modal');
}

function deleteVehicle(vehicleId) {
    const db = getDB();
    db.vehicles = db.vehicles.filter(v => v.id !== vehicleId);
    saveDB(db);

    closeModal('confirm-modal');
    showToast('Vehicle deleted successfully!');
    renderVehiclesPage();
    renderDashboard();
}

// Service CRUD Operations
function openAddServiceModal() {
    document.getElementById('service-vehicle-id').value = currentVehicleId;
    setDefaultServiceDate();
    openModal('add-service-modal');
}

function setDefaultServiceDate() {
    const today = new Date().toISOString().split('T')[0];
    document.getElementById('service-date').value = today;
}

function addService(event) {
    event.preventDefault();

    const db = getDB();
    const vehicleId = document.getElementById('service-vehicle-id').value;
    const vehicle = db.vehicles.find(v => v.id === vehicleId);

    if (!vehicle) {
        showToast('Vehicle not found', 'error');
        return;
    }

    const work = document.getElementById('service-work').value.trim();
    const provider = document.getElementById('service-provider').value.trim();
    const location = document.getElementById('service-location').value.trim();

    const newService = {
        id: generateId(),
        date: document.getElementById('service-date').value,
        work: work,
        duration: '',
        provider: provider,
        mileage: document.getElementById('service-mileage').value,
        location: location,
        cost: '',
        notes: document.getElementById('service-notes').value.trim(),
        createdAt: new Date().toISOString()
    };

    vehicle.services.push(newService);
    saveDB(db);

    // Save suggestions for autocomplete
    saveNewServiceSuggestions(work, provider, location);

    closeModal('add-service-modal');
    document.getElementById('add-service-form').reset();
    setDefaultServiceDate();
    showToast('Service added successfully!');
    autoSaveBackup('service_added');
    openVehicleDetail(vehicleId);
}

function editService(vehicleId, serviceId) {
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === vehicleId);
    const service = vehicle?.services.find(s => s.id === serviceId);

    if (!service) {
        showToast('Service not found', 'error');
        return;
    }

    // Populate form
    document.getElementById('edit-service-id').value = serviceId;
    document.getElementById('edit-service-vehicle-id').value = vehicleId;
    document.getElementById('edit-service-date').value = service.date;
    document.getElementById('edit-service-work').value = service.work;
    document.getElementById('edit-service-provider').value = service.provider;
    document.getElementById('edit-service-mileage').value = service.mileage;
    document.getElementById('edit-service-location').value = service.location;
    document.getElementById('edit-service-notes').value = service.notes || '';

    openModal('edit-service-modal');
}

function updateService(event) {
    event.preventDefault();

    const db = getDB();
    const vehicleId = document.getElementById('edit-service-vehicle-id').value;
    const serviceId = document.getElementById('edit-service-id').value;
    const vehicle = db.vehicles.find(v => v.id === vehicleId);
    const serviceIndex = vehicle?.services.findIndex(s => s.id === serviceId);

    if (serviceIndex === -1) {
        showToast('Service not found', 'error');
        return;
    }

    const work = document.getElementById('edit-service-work').value.trim();
    const provider = document.getElementById('edit-service-provider').value.trim();
    const location = document.getElementById('edit-service-location').value.trim();

    vehicle.services[serviceIndex] = {
        ...vehicle.services[serviceIndex],
        date: document.getElementById('edit-service-date').value,
        work: work,
        provider: provider,
        mileage: document.getElementById('edit-service-mileage').value,
        location: location,
        notes: document.getElementById('edit-service-notes').value.trim(),
        updatedAt: new Date().toISOString()
    };

    saveDB(db);

    // Save suggestions for autocomplete
    saveNewServiceSuggestions(work, provider, location);

    closeModal('edit-service-modal');
    showToast('Service updated successfully!');
    openVehicleDetail(vehicleId);
}

function confirmDeleteService(vehicleId, serviceId) {
    document.getElementById('confirm-message').textContent =
        'Are you sure you want to delete this service record?';

    const confirmBtn = document.getElementById('confirm-action-btn');
    confirmBtn.onclick = () => deleteService(vehicleId, serviceId);

    openModal('confirm-modal');
}

function deleteService(vehicleId, serviceId) {
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === vehicleId);

    if (vehicle) {
        vehicle.services = vehicle.services.filter(s => s.id !== serviceId);
        saveDB(db);
    }

    closeModal('confirm-modal');
    showToast('Service deleted successfully!');
    openVehicleDetail(vehicleId);
}

// Export/Import Functions
function exportAllData() {
    const db = getDB();
    const dataStr = JSON.stringify(db, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = `fleet_maintenance_backup_${formatDateForFile(new Date())}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('Data exported successfully!');
}

function exportAllDataExcel() {
    const db = getDB();

    // Create workbook
    const wb = XLSX.utils.book_new();

    // Helper function to split date
    const splitDate = (dateStr) => {
        if (!dateStr) return { monthDay: '', year: '' };
        const date = new Date(dateStr);
        const day = date.getDate();
        const month = date.toLocaleString('en', { month: 'short' });
        return {
            monthDay: `${day} ${month}`,
            year: date.getFullYear().toString()
        };
    };

    // Sheet 1: All Services
    const allServices = [];
    db.vehicles.forEach(vehicle => {
        vehicle.services.forEach(service => {
            const dateParts = splitDate(service.date);
            allServices.push({
                'Registration': vehicle.reg,
                'Make': vehicle.make || '',
                'Model': vehicle.model || '',
                'Day/Month': dateParts.monthDay,
                'Year': dateParts.year,
                'Work Done': service.work,
                'Provider': service.provider || '',
                'Mileage': service.mileage || '',
                'Location': service.location || '',
                'Notes': service.notes || ''
            });
        });
    });

    // Sort by date descending (using original date for sorting)
    allServices.sort((a, b) => {
        const dateA = new Date(`${a['Day/Month']} ${a['Year']}`);
        const dateB = new Date(`${b['Day/Month']} ${b['Year']}`);
        return dateB - dateA;
    });

    const wsServices = XLSX.utils.json_to_sheet(allServices);

    // Set column widths
    wsServices['!cols'] = [
        { wch: 12 },  // Registration
        { wch: 10 },  // Make
        { wch: 12 },  // Model
        { wch: 10 },  // Day/Month
        { wch: 6 },   // Year
        { wch: 40 },  // Work Done
        { wch: 30 },  // Provider
        { wch: 10 },  // Mileage
        { wch: 20 },  // Location
        { wch: 40 }   // Notes
    ];

    XLSX.utils.book_append_sheet(wb, wsServices, 'All Services');

    // Sheet 2: Vehicles Summary
    const vehiclesSummary = db.vehicles.map(vehicle => {
        const serviceCount = vehicle.services.length;
        const lastService = serviceCount > 0
            ? vehicle.services.sort((a, b) => new Date(b.date) - new Date(a.date))[0]
            : null;
        const lastDateParts = lastService ? splitDate(lastService.date) : { monthDay: '', year: '' };

        return {
            'Registration': vehicle.reg,
            'Make': vehicle.make || '',
            'Model': vehicle.model || '',
            'Vehicle Year': vehicle.year || '',
            'Color': vehicle.color || '',
            'Initial Mileage': vehicle.initialMileage || '',
            'Total Services': serviceCount,
            'Last Service Day/Month': lastDateParts.monthDay,
            'Last Service Year': lastDateParts.year,
            'Last Service Work': lastService ? lastService.work : '',
            'Notes': vehicle.notes || ''
        };
    });

    // Sort by registration
    vehiclesSummary.sort((a, b) => a.Registration.localeCompare(b.Registration));

    const wsVehicles = XLSX.utils.json_to_sheet(vehiclesSummary);

    wsVehicles['!cols'] = [
        { wch: 12 },  // Registration
        { wch: 10 },  // Make
        { wch: 12 },  // Model
        { wch: 12 },  // Vehicle Year
        { wch: 10 },  // Color
        { wch: 15 },  // Initial Mileage
        { wch: 14 },  // Total Services
        { wch: 18 },  // Last Service Day/Month
        { wch: 16 },  // Last Service Year
        { wch: 40 },  // Last Service Work
        { wch: 30 }   // Notes
    ];

    XLSX.utils.book_append_sheet(wb, wsVehicles, 'Vehicles Summary');

    // Sheet 3: Services by Vehicle (one sheet per vehicle)
    const allVehiclesSorted = [...db.vehicles]
        .sort((a, b) => a.reg.localeCompare(b.reg));

    allVehiclesSorted.forEach(vehicle => {
        const vehicleServices = vehicle.services
            .sort((a, b) => new Date(b.date) - new Date(a.date))
            .map(service => {
                const dateParts = splitDate(service.date);
                return {
                    'Day/Month': dateParts.monthDay,
                    'Year': dateParts.year,
                    'Work Done': service.work,
                    'Provider': service.provider || '',
                    'Mileage': service.mileage || '',
                    'Location': service.location || '',
                    'Notes': service.notes || ''
                };
            });

        const wsVehicle = XLSX.utils.json_to_sheet(vehicleServices);
        wsVehicle['!cols'] = [
            { wch: 10 },  // Day/Month
            { wch: 6 },   // Year
            { wch: 40 },  // Work Done
            { wch: 30 },  // Provider
            { wch: 10 },  // Mileage
            { wch: 20 },  // Location
            { wch: 40 }   // Notes
        ];

        // Sheet name max 31 chars
        const sheetName = vehicle.reg.substring(0, 31);
        XLSX.utils.book_append_sheet(wb, wsVehicle, sheetName);
    });

    // Generate and download file
    const fileName = `fleet_maintenance_${formatDateForFile(new Date())}.xlsx`;
    XLSX.writeFile(wb, fileName);

    showToast(`Excel exported: ${allServices.length} services, ${db.vehicles.length} vehicles`);
}

function exportVehicleData() {
    const db = getDB();
    const vehicle = db.vehicles.find(v => v.id === currentVehicleId);

    if (!vehicle) {
        showToast('Vehicle not found', 'error');
        return;
    }

    const dataStr = JSON.stringify(vehicle, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = `${vehicle.reg}_service_history_${formatDateForFile(new Date())}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('Vehicle data exported successfully!');
}

function importData(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const importedData = JSON.parse(e.target.result);

            // Validate structure
            if (!importedData.vehicles || !Array.isArray(importedData.vehicles)) {
                throw new Error('Invalid file format');
            }

            // Merge or replace data
            const currentDB = getDB();

            importedData.vehicles.forEach(importedVehicle => {
                const existingIndex = currentDB.vehicles.findIndex(
                    v => v.reg.toUpperCase() === importedVehicle.reg.toUpperCase()
                );

                if (existingIndex >= 0) {
                    // Merge services
                    const existingIds = new Set(currentDB.vehicles[existingIndex].services.map(s => s.id));
                    importedVehicle.services.forEach(service => {
                        if (!existingIds.has(service.id)) {
                            currentDB.vehicles[existingIndex].services.push(service);
                        }
                    });
                } else {
                    // Add new vehicle
                    currentDB.vehicles.push(importedVehicle);
                }
            });

            saveDB(currentDB);
            showToast('Data imported successfully!');
            autoSaveBackup('json_import');
            renderDashboard();
        } catch (error) {
            showToast('Error importing data: Invalid file format', 'error');
        }
    };
    reader.readAsText(file);
    event.target.value = ''; // Reset file input
}

// Excel Import Function
function importExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    showToast('Processing Excel file...', 'info');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });

            console.log('Sheets found:', workbook.SheetNames);

            const vehicles = {};
            const currentDB = getDB();

            // Helper functions
            const cleanReg = (reg) => {
                if (!reg) return null;
                reg = String(reg).trim().toUpperCase().replace(/\s+/g, '');
                return reg && reg !== 'NONE' && reg.length >= 7 ? reg : null;
            };

            const formatDate = (date) => {
                if (!date) return null;
                if (date instanceof Date) {
                    return date.toISOString().split('T')[0];
                }
                return null;
            };

            const isRegNumber = (val) => {
                if (!val || typeof val !== 'string') return false;
                const clean = val.trim().toUpperCase().replace(/\s+/g, '');
                return /^[A-Z]{2}\d{2}[A-Z]{3}$/.test(clean);
            };

            // 1. Process "Van list" sheet for vehicle info
            const vanListSheet = workbook.Sheets['Van list '] || workbook.Sheets['Van list'];
            if (vanListSheet) {
                const vanListData = XLSX.utils.sheet_to_json(vanListSheet, { header: 1 });
                vanListData.forEach((row, i) => {
                    if (i < 2 || !row[2]) return;
                    const reg = cleanReg(row[2]);
                    if (!reg) return;

                    if (!vehicles[reg]) {
                        vehicles[reg] = {
                            id: generateId(),
                            reg: reg,
                            make: String(row[1] || 'Ford').trim(),
                            model: 'Transit',
                            year: '2021',
                            initialMileage: 0,
                            color: 'White',
                            notes: row[0] ? `Provider: ${row[0]}` : '',
                            services: [],
                            createdAt: new Date().toISOString()
                        };
                    }
                });
            }

            // 2. Process individual vehicle sheets (most detailed)
            const vehicleSheets = workbook.SheetNames.filter(name =>
                /^[A-Z]{2}\d{2}[A-Z]{3}$/i.test(name.replace(/\s+/g, ''))
            );

            vehicleSheets.forEach(sheetName => {
                const reg = cleanReg(sheetName);
                if (!reg) return;

                const sheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                if (!vehicles[reg]) {
                    vehicles[reg] = {
                        id: generateId(),
                        reg: reg,
                        make: 'Ford',
                        model: 'Transit',
                        year: '2021',
                        initialMileage: 0,
                        color: 'White',
                        notes: '',
                        services: [],
                        createdAt: new Date().toISOString()
                    };
                }

                // Skip header rows, process from row 5 onwards
                rows.forEach((row, i) => {
                    if (i < 4 || !row[0]) return;

                    // Skip if first cell is a header
                    if (String(row[0]).toLowerCase().includes('date')) return;

                    const date = row[0] instanceof Date ? row[0] : null;
                    if (!date) return;

                    const description = String(row[3] || row[5] || 'Service').trim();
                    const miles = typeof row[4] === 'number' ? row[4] : 0;
                    const provider = String(row[1] || row[2] || 'In-house').trim();

                    vehicles[reg].services.push({
                        id: generateId(),
                        date: formatDate(date),
                        work: description,
                        duration: '',
                        provider: provider,
                        mileage: miles,
                        location: 'On Site',
                        cost: '',
                        notes: '',
                        createdAt: new Date().toISOString()
                    });

                    // Update max mileage
                    if (miles > vehicles[reg].initialMileage) {
                        vehicles[reg].initialMileage = miles;
                    }
                });
            });

            // 3. Process yearly sheets
            const yearlySheets = ['2022', '2023 ', '2023', '2024', '2025'];
            yearlySheets.forEach(yearSheet => {
                const sheet = workbook.Sheets[yearSheet];
                if (!sheet) return;

                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                rows.forEach((row, i) => {
                    if (i < 1 || !row) return;

                    // Find reg and date in row
                    let reg = null;
                    let date = null;
                    let work = '';
                    let location = 'Redruth';

                    row.forEach(cell => {
                        if (cell instanceof Date && !date) {
                            date = cell;
                        } else if (isRegNumber(String(cell))) {
                            reg = cleanReg(cell);
                        } else if (typeof cell === 'string') {
                            const cellUpper = cell.toUpperCase();
                            if (cellUpper.includes('REDRUTH') || cellUpper.includes('SITE') ||
                                cellUpper.includes('GARAGE') || cellUpper.includes('FORD')) {
                                location = cell;
                            }
                        }
                    });

                    if (!reg || !date) return;

                    // Get work description from last meaningful columns
                    for (let j = row.length - 1; j >= 0; j--) {
                        const val = row[j];
                        if (val && typeof val === 'string' && val.length > 3) {
                            const lower = val.toLowerCase();
                            if (!['yes', 'no', 'ok', 'none', 'ford', 'transit', 'redruth'].some(x => lower === x)) {
                                if (!lower.includes('ford transit') && !lower.includes('arval') && !lower.includes('leaseplan')) {
                                    work = val;
                                    break;
                                }
                            }
                        }
                    }

                    if (!work) work = 'Maintenance';

                    if (!vehicles[reg]) {
                        vehicles[reg] = {
                            id: generateId(),
                            reg: reg,
                            make: 'Ford',
                            model: 'Transit',
                            year: '2021',
                            initialMileage: 0,
                            color: 'White',
                            notes: '',
                            services: [],
                            createdAt: new Date().toISOString()
                        };
                    }

                    // Check for duplicate
                    const dateStr = formatDate(date);
                    const isDuplicate = vehicles[reg].services.some(
                        s => s.date === dateStr && s.work.toLowerCase() === work.toLowerCase()
                    );

                    if (!isDuplicate) {
                        vehicles[reg].services.push({
                            id: generateId(),
                            date: dateStr,
                            work: work,
                            duration: '',
                            provider: 'In-house',
                            mileage: 0,
                            location: location,
                            cost: '',
                            notes: '',
                            createdAt: new Date().toISOString()
                        });
                    }
                });
            });

            // Merge with existing data
            Object.values(vehicles).forEach(importedVehicle => {
                if (importedVehicle.services.length === 0) return; // Skip vehicles with no services

                const existingIndex = currentDB.vehicles.findIndex(
                    v => v.reg.toUpperCase() === importedVehicle.reg.toUpperCase()
                );

                if (existingIndex >= 0) {
                    // Merge services
                    const existing = currentDB.vehicles[existingIndex];
                    importedVehicle.services.forEach(service => {
                        const isDuplicate = existing.services.some(
                            s => s.date === service.date && s.work.toLowerCase() === service.work.toLowerCase()
                        );
                        if (!isDuplicate) {
                            existing.services.push(service);
                        }
                    });
                    // Update mileage if higher
                    if (importedVehicle.initialMileage > existing.initialMileage) {
                        existing.initialMileage = importedVehicle.initialMileage;
                    }
                } else {
                    currentDB.vehicles.push(importedVehicle);
                }
            });

            saveDB(currentDB);

            const totalVehicles = Object.keys(vehicles).length;
            const totalServices = Object.values(vehicles).reduce((sum, v) => sum + v.services.length, 0);

            showToast(`Imported ${totalVehicles} vehicles with ${totalServices} services!`);
            autoSaveBackup('excel_import');
            renderDashboard();

        } catch (error) {
            console.error('Excel import error:', error);
            showToast('Error importing Excel: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = '';
}

// VOR Data Import Function
function importVORData(event) {
    const file = event.target.files[0];
    if (!file) return;

    showToast('Processing VOR data file...', 'info');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet);

            console.log('VOR Data - Columns:', Object.keys(rows[0] || {}));
            console.log('VOR Data - Total rows:', rows.length);

            const vehicles = {};
            const currentDB = getDB();

            rows.forEach(row => {
                // Get Registration
                const reg = row['Registration'] ? String(row['Registration']).trim().toUpperCase() : null;
                if (!reg) return;

                // Create vehicle if not exists
                if (!vehicles[reg]) {
                    const make = row['Make'] ? String(row['Make']).trim() : 'Unknown';
                    const modelFull = row['Model'] ? String(row['Model']).trim() : '';
                    const model = modelFull.includes('TRNST') || modelFull.toUpperCase().includes('TRANSIT')
                        ? 'Transit'
                        : modelFull.substring(0, 30);

                    vehicles[reg] = {
                        id: generateId(),
                        reg: reg,
                        make: make,
                        model: model,
                        year: '2021',
                        initialMileage: 0,
                        color: 'White',
                        notes: row['Leasee'] ? `Leasee: ${row['Leasee']}` : '',
                        services: [],
                        createdAt: new Date().toISOString()
                    };
                }

                // Get End Date (service date)
                let serviceDate = null;
                const endDate = row['End Date'];
                if (endDate) {
                    if (endDate instanceof Date) {
                        serviceDate = endDate.toISOString().split('T')[0];
                    } else {
                        serviceDate = String(endDate).substring(0, 10);
                    }
                }
                if (!serviceDate) return;

                // Get Core VOR Reason (work done)
                let work = row['Core VOR Reason'] ? String(row['Core VOR Reason']).trim() : '';
                if (!work || work === 'NaN' || work === 'undefined') {
                    work = 'Maintenance';
                }

                // Get Supplier Name (provider)
                let provider = row['Supplier Name'] ? String(row['Supplier Name']).trim() : 'Unknown';
                if (!provider || provider === 'NaN' || provider === 'undefined') {
                    provider = 'Unknown';
                }

                // Get VOR duration
                const duration = row['Sum of VOR [Days : Hours : Minutes]']
                    ? String(row['Sum of VOR [Days : Hours : Minutes]']).trim()
                    : '';

                // Get VOR Summary as notes
                let notes = row['VOR Summary'] ? String(row['VOR Summary']).trim() : '';
                if (notes === 'NaN' || notes === 'undefined') notes = '';
                if (notes.length > 500) notes = notes.substring(0, 500) + '...';

                // Create service
                const service = {
                    id: generateId(),
                    date: serviceDate,
                    work: work,
                    duration: duration,
                    provider: provider,
                    mileage: 0,
                    location: 'VOR Service',
                    cost: '',
                    notes: notes,
                    createdAt: new Date().toISOString()
                };

                // Check for duplicates
                const isDuplicate = vehicles[reg].services.some(
                    s => s.date === serviceDate && s.work.toLowerCase() === work.toLowerCase()
                );

                if (!isDuplicate) {
                    vehicles[reg].services.push(service);
                }
            });

            // Merge with existing data
            let newVehicles = 0;
            let newServices = 0;
            let updatedVehicles = 0;

            Object.values(vehicles).forEach(importedVehicle => {
                if (importedVehicle.services.length === 0) return;

                const existingIndex = currentDB.vehicles.findIndex(
                    v => v.reg.toUpperCase() === importedVehicle.reg.toUpperCase()
                );

                if (existingIndex >= 0) {
                    // Merge services into existing vehicle
                    const existing = currentDB.vehicles[existingIndex];
                    let addedToExisting = 0;

                    importedVehicle.services.forEach(service => {
                        const isDuplicate = existing.services.some(
                            s => s.date === service.date && s.work.toLowerCase() === service.work.toLowerCase()
                        );
                        if (!isDuplicate) {
                            existing.services.push(service);
                            addedToExisting++;
                        }
                    });

                    if (addedToExisting > 0) {
                        updatedVehicles++;
                        newServices += addedToExisting;
                    }
                } else {
                    // Add new vehicle
                    currentDB.vehicles.push(importedVehicle);
                    newVehicles++;
                    newServices += importedVehicle.services.length;
                }
            });

            saveDB(currentDB);

            const totalVehicles = Object.keys(vehicles).length;
            showToast(`Imported: ${newVehicles} new vehicles, ${newServices} services. Updated ${updatedVehicles} existing vehicles.`);
            autoSaveBackup('vor_import');
            renderDashboard();

        } catch (error) {
            console.error('VOR import error:', error);
            showToast('Error importing VOR data: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = '';
}

// Search Function
function searchKeyPress(event) {
    if (event.key === 'Enter') {
        event.preventDefault();
        const query = document.getElementById('global-search').value.trim().toUpperCase();

        if (!query) return;

        const db = getDB();
        // Find exact match first
        const exactMatch = db.vehicles.find(v => v.reg.toUpperCase() === query);

        if (exactMatch) {
            // Clear search and open vehicle detail
            document.getElementById('global-search').value = '';
            openVehicleDetail(exactMatch.id);
        } else {
            // Find partial match
            const partialMatch = db.vehicles.find(v => v.reg.toUpperCase().includes(query));
            if (partialMatch) {
                document.getElementById('global-search').value = '';
                openVehicleDetail(partialMatch.id);
            } else {
                showToast(`Vehicle "${query}" not found`, 'error');
            }
        }
    }
}

function globalSearch(query) {
    query = query.toLowerCase().trim();
    if (!query) {
        renderDashboard();
        return;
    }

    const db = getDB();
    const results = db.vehicles.filter(vehicle => {
        if (vehicle.reg.toLowerCase().includes(query)) return true;
        if (vehicle.make?.toLowerCase().includes(query)) return true;
        if (vehicle.model?.toLowerCase().includes(query)) return true;
        return vehicle.services.some(s =>
            s.work.toLowerCase().includes(query) ||
            s.provider.toLowerCase().includes(query) ||
            s.location.toLowerCase().includes(query)
        );
    });

    const vehiclesList = document.getElementById('dashboard-vehicles-list');
    if (results.length === 0) {
        vehiclesList.innerHTML = `
            <div class="empty-state">
                <i class="fas fa-search"></i>
                <h3>No Results</h3>
                <p>No vehicles match your search</p>
            </div>
        `;
    } else {
        vehiclesList.innerHTML = results.map(vehicle => {
            const latestMileage = getLatestMileage(vehicle);
            return `
                <div class="vehicle-item" onclick="openVehicleDetail('${vehicle.id}')">
                    <div class="vehicle-icon">
                        <i class="fas fa-truck"></i>
                    </div>
                    <div class="vehicle-info">
                        <div class="vehicle-reg">${vehicle.reg}</div>
                        <div class="vehicle-meta">${vehicle.make || ''} ${vehicle.model || ''}</div>
                    </div>
                    <div class="vehicle-mileage">
                        <div class="value">${formatNumber(latestMileage)}</div>
                        <div class="label">miles</div>
                    </div>
                </div>
            `;
        }).join('');
    }
}

// Modal Functions
function openModal(modalId) {
    // Populate merge modal if opening it
    if (modalId === 'merge-vehicles-modal') {
        populateMergeModal();
    }
    document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
    document.getElementById(modalId).classList.remove('active');
}

// Merge Vehicles Functions
function populateMergeModal() {
    const db = getDB();
    const vehicles = [...db.vehicles].sort((a, b) => a.reg.localeCompare(b.reg));

    const targetSelect = document.getElementById('merge-target');
    targetSelect.innerHTML = '<option value="">Select target vehicle...</option>';

    vehicles.forEach(vehicle => {
        const serviceCount = vehicle.services.length;
        targetSelect.innerHTML += `<option value="${vehicle.id}">${vehicle.reg} (${serviceCount} services)</option>`;
    });

    // Clear source list and preview
    document.getElementById('merge-source-list').innerHTML = '<p style="color: var(--text-secondary); font-size: 13px;">Select a target vehicle first</p>';
    document.getElementById('merge-preview').style.display = 'none';
}

function updateMergeSourceList() {
    const db = getDB();
    const targetId = document.getElementById('merge-target').value;
    const sourceList = document.getElementById('merge-source-list');

    if (!targetId) {
        sourceList.innerHTML = '<p style="color: var(--text-secondary); font-size: 13px;">Select a target vehicle first</p>';
        document.getElementById('merge-preview').style.display = 'none';
        return;
    }

    const vehicles = [...db.vehicles]
        .filter(v => v.id !== targetId)
        .sort((a, b) => a.reg.localeCompare(b.reg));

    if (vehicles.length === 0) {
        sourceList.innerHTML = '<p style="color: var(--text-secondary); font-size: 13px;">No other vehicles to merge</p>';
        return;
    }

    sourceList.innerHTML = vehicles.map(vehicle => {
        const serviceCount = vehicle.services.length;
        return `
            <div class="merge-checkbox-item">
                <input type="checkbox" id="merge-${vehicle.id}" value="${vehicle.id}" onchange="updateMergePreview()">
                <label for="merge-${vehicle.id}">
                    <span>${vehicle.reg}</span>
                    <span class="service-count">${serviceCount} services</span>
                </label>
            </div>
        `;
    }).join('');

    updateMergePreview();
}

function updateMergePreview() {
    const db = getDB();
    const targetId = document.getElementById('merge-target').value;
    const preview = document.getElementById('merge-preview');
    const previewText = document.getElementById('merge-preview-text');

    if (!targetId) {
        preview.style.display = 'none';
        return;
    }

    const targetVehicle = db.vehicles.find(v => v.id === targetId);
    const selectedSources = Array.from(document.querySelectorAll('#merge-source-list input:checked'))
        .map(cb => cb.value);

    if (selectedSources.length === 0) {
        preview.style.display = 'none';
        return;
    }

    let totalServices = targetVehicle.services.length;
    const sourceRegs = [];

    selectedSources.forEach(sourceId => {
        const sourceVehicle = db.vehicles.find(v => v.id === sourceId);
        if (sourceVehicle) {
            totalServices += sourceVehicle.services.length;
            sourceRegs.push(sourceVehicle.reg);
        }
    });

    preview.style.display = 'block';
    previewText.innerHTML = `
        <strong>${sourceRegs.join(', ')}</strong> will be merged into <strong>${targetVehicle.reg}</strong><br>
        Total services after merge: <strong>${totalServices}</strong><br>
        <span style="color: var(--danger); font-size: 12px;">The merged vehicles will be deleted.</span>
    `;
}

function mergeVehicles() {
    const db = getDB();
    const targetId = document.getElementById('merge-target').value;

    if (!targetId) {
        showToast('Please select a target vehicle', 'error');
        return;
    }

    const selectedSources = Array.from(document.querySelectorAll('#merge-source-list input:checked'))
        .map(cb => cb.value);

    if (selectedSources.length === 0) {
        showToast('Please select at least one vehicle to merge', 'error');
        return;
    }

    const targetVehicle = db.vehicles.find(v => v.id === targetId);
    const sourceRegs = [];

    // Merge services from each source vehicle
    selectedSources.forEach(sourceId => {
        const sourceVehicle = db.vehicles.find(v => v.id === sourceId);
        if (sourceVehicle) {
            sourceRegs.push(sourceVehicle.reg);
            // Add all services from source to target
            sourceVehicle.services.forEach(service => {
                // Check for duplicate (same date and work)
                const isDuplicate = targetVehicle.services.some(
                    s => s.date === service.date && s.work.toLowerCase() === service.work.toLowerCase()
                );
                if (!isDuplicate) {
                    targetVehicle.services.push(service);
                }
            });
        }
    });

    // Remove merged vehicles
    db.vehicles = db.vehicles.filter(v => !selectedSources.includes(v.id));

    // Update mileage if needed (keep highest)
    const allMileages = targetVehicle.services
        .map(s => parseInt(s.mileage) || 0)
        .filter(m => m > 0);
    if (allMileages.length > 0) {
        const maxMileage = Math.max(...allMileages);
        if (maxMileage > (targetVehicle.initialMileage || 0)) {
            targetVehicle.initialMileage = maxMileage;
        }
    }

    saveDB(db);
    closeModal('merge-vehicles-modal');
    showToast(`Merged ${sourceRegs.join(', ')} into ${targetVehicle.reg}. Total: ${targetVehicle.services.length} services.`);
    autoSaveBackup('vehicles_merged');
    renderVehiclesPage();
    renderDashboard();
}

// Close modal when clicking outside
document.addEventListener('click', (e) => {
    if (e.target.classList.contains('modal')) {
        e.target.classList.remove('active');
    }
});

// Sidebar Toggle
function toggleSidebar() {
    const sidebar = document.querySelector('.sidebar');
    const mainContent = document.querySelector('.main-content');
    sidebar.classList.toggle('open');
    mainContent.classList.toggle('expanded');
}

// Utility Functions
function formatDate(dateString) {
    const date = new Date(dateString);
    return date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
}

function formatDateForFile(date) {
    return date.toISOString().split('T')[0];
}

function formatNumber(num) {
    if (!num) return '0';
    return parseInt(num).toLocaleString();
}

function truncateText(text, maxLength) {
    if (text.length <= maxLength) return text;
    return text.substr(0, maxLength) + '...';
}

function getLatestMileage(vehicle) {
    if (vehicle.services.length === 0) {
        return vehicle.initialMileage || 0;
    }
    const sorted = [...vehicle.services].sort((a, b) => new Date(b.date) - new Date(a.date));
    return sorted[0].mileage || vehicle.initialMileage || 0;
}

function updateCurrentDate() {
    const now = new Date();
    const options = { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' };
    document.getElementById('current-date').textContent = now.toLocaleDateString('en-GB', options);
}

// Toast Notification
function showToast(message, type = 'success') {
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toast-message');

    toast.className = 'toast show ' + type;
    toastMessage.textContent = message;

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Keyboard shortcuts
document.addEventListener('keydown', (e) => {
    // Escape to close modals
    if (e.key === 'Escape') {
        document.querySelectorAll('.modal.active').forEach(modal => {
            modal.classList.remove('active');
        });
    }

    // Ctrl/Cmd + N to add new vehicle
    if ((e.ctrlKey || e.metaKey) && e.key === 'n') {
        e.preventDefault();
        openModal('add-vehicle-modal');
    }
});
