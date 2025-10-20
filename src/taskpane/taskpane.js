// ============================================================================
// OUTLOOK MOVE-TO-FOLDER ADD-IN
// ============================================================================

// Configuration
const CONFIG = {
    CACHE_KEY: 'mailFoldersCache',
    RECENT_KEY: 'recentFolders',
    CACHE_TTL: 1000 * 60 * 60 * 6, // 6 hours
    RECENT_LIMIT: 8,
    DEBOUNCE_DELAY: 50,
    FUSE_THRESHOLD: 0.3
};

// Global state
let state = {
    folders: [],
    recentFolders: [],
    filteredFolders: [],
    selectedIndex: 0,
    fuse: null,
    accessToken: null,
    currentItemId: null,
    isLoading: false
};

// ============================================================================
// INITIALIZATION
// ============================================================================

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        initialize();
    }
});

async function initialize() {
    try {
        // Set up event listeners
        setupEventListeners();
        
        // Show loading state
        showState('loading');
        
        // Prefetch access token (warm up connection)
        await prefetchToken();
        
        // Get current email item ID
        await getCurrentItemId();
        
        // Load folders (from cache or API)
        await loadFolders();
        
        // Load recent folders
        loadRecentFolders();
        
        // Initialize Fuse.js for fuzzy search
        initializeFuzzySearch();
        
        // Show main interface
        showState('main');
        
        // Auto-focus search input
        document.getElementById('search-input').focus();
        
        // Display all folders initially
        renderFolders(state.folders);
        
    } catch (error) {
        console.error('Initialization error:', error);
        showError('Failed to initialize add-in: ' + error.message);
    }
}

// ============================================================================
// AUTHENTICATION & TOKEN MANAGEMENT
// ============================================================================

async function prefetchToken() {
    try {
        // Use REST API with Office.context.mailbox.getCallbackTokenAsync instead of SSO
        state.accessToken = await getAccessToken();
        
    } catch (error) {
        console.error('Token prefetch error:', error);
        // Non-fatal - we'll try again when needed
    }
}

async function getAccessToken() {
    if (state.accessToken) {
        return state.accessToken;
    }
    
    return new Promise((resolve, reject) => {
        // Try REST token first (works in most scenarios)
        Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                state.accessToken = result.value;
                resolve(result.value);
            } else {
                // Fallback: try regular callback token
                Office.context.mailbox.getCallbackTokenAsync((fallbackResult) => {
                    if (fallbackResult.status === Office.AsyncResultStatus.Succeeded) {
                        state.accessToken = fallbackResult.value;
                        resolve(fallbackResult.value);
                    } else {
                        reject(new Error('Failed to get access token: ' + fallbackResult.error.message));
                    }
                });
            }
        });
    });
}

// ============================================================================
// MICROSOFT GRAPH API
// ============================================================================

async function fetchFoldersFromGraph() {
    const token = await getAccessToken();
    const restUrl = Office.context.mailbox.restUrl;
    
    // Use Outlook REST API instead of Graph API (works with callback token)
    const response = await fetch(
        `${restUrl}/v2.0/me/mailfolders?$top=500&$select=Id,DisplayName,ParentFolderId`,
        {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        }
    );
    
    if (!response.ok) {
        throw new Error(`Outlook REST API error: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    // Convert REST API response format to match our expected format
    let folders = data.value.map(f => ({
        id: f.Id,
        displayName: f.DisplayName,
        parentFolderId: f.ParentFolderId
    }));
    
    // Recursively fetch child folders for each folder
    const allFolders = [];
    for (const folder of folders) {
        allFolders.push(folder);
        const childFolders = await fetchChildFolders(folder.id, token);
        allFolders.push(...childFolders);
    }
    
    return allFolders;
}

async function fetchChildFolders(parentId, token) {
    const restUrl = Office.context.mailbox.restUrl;
    
    const response = await fetch(
        `${restUrl}/v2.0/me/mailfolders/${parentId}/childfolders?$top=500&$select=Id,DisplayName,ParentFolderId`,
        {
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            }
        }
    );
    
    if (!response.ok) {
        return [];
    }
    
    const data = await response.json();
    const folders = data.value.map(f => ({
        id: f.Id,
        displayName: f.DisplayName,
        parentFolderId: f.ParentFolderId
    }));
    
    // Recursively fetch children of children
    const allFolders = [...folders];
    for (const folder of folders) {
        const childFolders = await fetchChildFolders(folder.id, token);
        allFolders.push(...childFolders);
    }
    
    return allFolders;
}

async function moveMessageToFolder(messageId, folderId) {
    const token = await getAccessToken();
    const restUrl = Office.context.mailbox.restUrl;
    
    const response = await fetch(
        `${restUrl}/v2.0/me/messages/${messageId}/move`,
        {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                DestinationId: folderId
            })
        }
    );
    
    if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error?.message || `Failed to move message: ${response.status}`);
    }
    
    return await response.json();
}

// ============================================================================
// FOLDER CACHING
// ============================================================================

async function loadFolders(force = false) {
    try {
        // Check cache first
        const cached = JSON.parse(localStorage.getItem(CONFIG.CACHE_KEY) || '{}');
        const isValid = cached.timestamp && (Date.now() - cached.timestamp < CONFIG.CACHE_TTL);
        
        if (!force && isValid && cached.folders && cached.folders.length > 0) {
            state.folders = cached.folders;
            return;
        }
        
        // Fetch from Graph API
        const folders = await fetchFoldersFromGraph();
        
        // Build folder hierarchy and paths
        const foldersWithPaths = buildFolderPaths(folders);
        
        // Cache the results
        localStorage.setItem(CONFIG.CACHE_KEY, JSON.stringify({
            folders: foldersWithPaths,
            timestamp: Date.now()
        }));
        
        state.folders = foldersWithPaths;
        
    } catch (error) {
        console.error('Error loading folders:', error);
        throw error;
    }
}

function buildFolderPaths(folders) {
    // Create a map for quick lookup
    const folderMap = new Map();
    folders.forEach(f => folderMap.set(f.id, f));
    
    // Build full paths
    return folders.map(folder => {
        const path = buildFolderPath(folder, folderMap);
        return {
            ...folder,
            path: path,
            searchText: `${folder.displayName} ${path}`.toLowerCase()
        };
    }).sort((a, b) => a.path.localeCompare(b.path));
}

function buildFolderPath(folder, folderMap, visited = new Set()) {
    if (visited.has(folder.id)) {
        return folder.displayName; // Prevent infinite loops
    }
    visited.add(folder.id);
    
    if (!folder.parentFolderId) {
        return folder.displayName;
    }
    
    const parent = folderMap.get(folder.parentFolderId);
    if (parent) {
        return buildFolderPath(parent, folderMap, visited) + ' / ' + folder.displayName;
    }
    
    return folder.displayName;
}

// ============================================================================
// RECENT FOLDERS
// ============================================================================

function loadRecentFolders() {
    try {
        const recent = JSON.parse(localStorage.getItem(CONFIG.RECENT_KEY) || '[]');
        
        // Filter to only include folders that still exist
        const folderIds = new Set(state.folders.map(f => f.id));
        state.recentFolders = recent.filter(f => folderIds.has(f.id));
        
    } catch (error) {
        console.error('Error loading recent folders:', error);
        state.recentFolders = [];
    }
}

function addToRecentFolders(folder) {
    // Remove if already exists
    state.recentFolders = state.recentFolders.filter(f => f.id !== folder.id);
    
    // Add to beginning
    state.recentFolders.unshift({
        id: folder.id,
        displayName: folder.displayName,
        path: folder.path
    });
    
    // Limit size
    state.recentFolders = state.recentFolders.slice(0, CONFIG.RECENT_LIMIT);
    
    // Save to localStorage
    localStorage.setItem(CONFIG.RECENT_KEY, JSON.stringify(state.recentFolders));
}

// ============================================================================
// FUZZY SEARCH
// ============================================================================

function initializeFuzzySearch() {
    const options = {
        keys: ['displayName', 'path'],
        threshold: CONFIG.FUSE_THRESHOLD,
        includeScore: true,
        ignoreLocation: true
    };
    
    state.fuse = new Fuse(state.folders, options);
}

let searchDebounceTimer = null;

function performSearch(query) {
    clearTimeout(searchDebounceTimer);
    
    searchDebounceTimer = setTimeout(() => {
        if (!query.trim()) {
            // Show all folders
            renderFolders(state.folders);
            return;
        }
        
        // Fuzzy search
        const results = state.fuse.search(query);
        const folders = results.map(r => r.item);
        
        renderFolders(folders);
        
    }, CONFIG.DEBOUNCE_DELAY);
}

// ============================================================================
// UI RENDERING
// ============================================================================

function showState(stateName) {
    const states = {
        loading: document.getElementById('loading-state'),
        main: document.getElementById('main-interface'),
        error: document.getElementById('error-state')
    };
    
    Object.values(states).forEach(el => el.classList.add('hidden'));
    states[stateName]?.classList.remove('hidden');
}

function renderFolders(folders) {
    state.filteredFolders = folders;
    state.selectedIndex = 0;
    
    const recentSection = document.getElementById('recent-section');
    const recentContainer = document.getElementById('recent-folders');
    const allFoldersContainer = document.getElementById('all-folders');
    const noResults = document.getElementById('no-results');
    
    // Clear containers
    recentContainer.innerHTML = '';
    allFoldersContainer.innerHTML = '';
    
    // Show/hide recent section
    const searchQuery = document.getElementById('search-input').value.trim();
    const shouldShowRecent = !searchQuery && state.recentFolders.length > 0;
    
    if (shouldShowRecent) {
        recentSection.classList.remove('hidden');
        state.recentFolders.forEach((folder, index) => {
            const el = createFolderElement(folder, index, true);
            recentContainer.appendChild(el);
        });
    } else {
        recentSection.classList.add('hidden');
    }
    
    // Render all folders
    if (folders.length === 0) {
        noResults.classList.remove('hidden');
        return;
    }
    
    noResults.classList.add('hidden');
    
    const startIndex = shouldShowRecent ? state.recentFolders.length : 0;
    folders.forEach((folder, index) => {
        const el = createFolderElement(folder, startIndex + index, false);
        allFoldersContainer.appendChild(el);
    });
    
    // Highlight first item
    highlightSelected();
}

function createFolderElement(folder, index, isRecent) {
    const div = document.createElement('div');
    div.className = 'folder-item px-4 py-3 cursor-pointer border-b border-gray-100 hover:bg-blue-50';
    div.dataset.index = index;
    div.dataset.folderId = folder.id;
    
    const nameEl = document.createElement('div');
    nameEl.className = 'font-medium text-sm text-gray-900';
    nameEl.textContent = folder.displayName;
    
    const pathEl = document.createElement('div');
    pathEl.className = 'text-xs text-gray-500 mt-1';
    pathEl.textContent = folder.path;
    
    div.appendChild(nameEl);
    div.appendChild(pathEl);
    
    // Click handler
    div.addEventListener('click', () => {
        state.selectedIndex = index;
        highlightSelected();
        moveToSelectedFolder();
    });
    
    return div;
}

function highlightSelected() {
    document.querySelectorAll('.folder-item').forEach((el, idx) => {
        if (parseInt(el.dataset.index) === state.selectedIndex) {
            el.classList.add('selected');
            // Scroll into view if needed
            el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        } else {
            el.classList.remove('selected');
        }
    });
}

function showStatus(message, type = 'info') {
    const statusEl = document.getElementById('status-message');
    statusEl.className = `px-4 py-3 text-sm ${
        type === 'success' ? 'bg-green-100 text-green-800' :
        type === 'error' ? 'bg-red-100 text-red-800' :
        'bg-blue-100 text-blue-800'
    }`;
    statusEl.textContent = message;
    statusEl.classList.remove('hidden');
    statusEl.classList.add('fade-in');
}

function hideStatus() {
    const statusEl = document.getElementById('status-message');
    statusEl.classList.add('fade-out');
    setTimeout(() => {
        statusEl.classList.add('hidden');
        statusEl.classList.remove('fade-out');
    }, 200);
}

function showError(message) {
    showState('error');
    document.getElementById('error-message').textContent = message;
}

// ============================================================================
// CURRENT EMAIL ITEM
// ============================================================================

async function getCurrentItemId() {
    try {
        // In Outlook, item.itemId is a property, not an async method
        const itemId = Office.context.mailbox.item.itemId;
        if (itemId) {
            state.currentItemId = itemId;
            return itemId;
        } else {
            throw new Error('No item ID available');
        }
    } catch (error) {
        console.error('Error getting item ID:', error);
        throw new Error('Failed to get current email ID');
    }
}

// ============================================================================
// MOVE ACTION
// ============================================================================

async function moveToSelectedFolder() {
    const selectedElements = document.querySelectorAll(`.folder-item[data-index="${state.selectedIndex}"]`);
    if (selectedElements.length === 0) return;
    
    const selectedElement = selectedElements[0];
    const folderId = selectedElement.dataset.folderId;
    
    // Find the folder object
    const folder = [...state.recentFolders, ...state.folders].find(f => f.id === folderId);
    if (!folder) return;
    
    try {
        // Step 1: Show immediate checkmark (50ms)
        showStatus(`✓ Moving to ${folder.displayName}...`, 'info');
        
        // Step 2: Execute move operation
        await moveMessageToFolder(state.currentItemId, folderId);
        
        // Step 3: Show success (150ms total)
        setTimeout(() => {
            showStatus(`✓ Moved to ${folder.displayName}`, 'success');
        }, 100);
        
        // Step 4: Add Office notification banner
        Office.context.mailbox.item.notificationMessages.addAsync(
            'moveSuccess',
            {
                type: 'informationalMessage',
                message: `📁 Moved to ${folder.displayName}`,
                icon: 'icon16',
                persistent: false
            }
        );
        
        // Step 5: Update recent folders
        addToRecentFolders(folder);
        
        // Step 6: Auto-close taskpane (300-350ms)
        setTimeout(() => {
            Office.context.ui.closeContainer();
        }, 250);
        
    } catch (error) {
        console.error('Move error:', error);
        showStatus(`Error: ${error.message}`, 'error');
        
        // Update notification banner with error
        Office.context.mailbox.item.notificationMessages.addAsync(
            'moveError',
            {
                type: 'errorMessage',
                message: `Failed to move email: ${error.message}`
            }
        );
    }
}

// ============================================================================
// EVENT LISTENERS
// ============================================================================

function setupEventListeners() {
    // Search input
    document.getElementById('search-input').addEventListener('input', (e) => {
        performSearch(e.target.value);
    });
    
    // Keyboard navigation
    document.getElementById('search-input').addEventListener('keydown', handleKeyboard);
    
    // Refresh button
    document.getElementById('refresh-btn').addEventListener('click', async () => {
        try {
            showStatus('Refreshing folders...', 'info');
            await loadFolders(true);
            initializeFuzzySearch();
            renderFolders(state.folders);
            hideStatus();
        } catch (error) {
            showStatus('Failed to refresh folders', 'error');
        }
    });
    
    // Retry button (error state)
    document.getElementById('retry-btn').addEventListener('click', () => {
        initialize();
    });
}

function handleKeyboard(e) {
    const totalItems = state.recentFolders.length + state.filteredFolders.length;
    
    switch(e.key) {
        case 'ArrowDown':
            e.preventDefault();
            state.selectedIndex = (state.selectedIndex + 1) % totalItems;
            highlightSelected();
            break;
            
        case 'ArrowUp':
            e.preventDefault();
            state.selectedIndex = (state.selectedIndex - 1 + totalItems) % totalItems;
            highlightSelected();
            break;
            
        case 'Enter':
            e.preventDefault();
            moveToSelectedFolder();
            break;
            
        case 'Escape':
            e.preventDefault();
            Office.context.ui.closeContainer();
            break;
    }
}


