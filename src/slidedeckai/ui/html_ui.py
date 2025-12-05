HTML_UI = """
<!DOCTYPE html>
<html>
<head>
    <title>SlideDeck AI - Advanced Report Generator</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            padding: 40px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        h1 {
            color: #2563eb;
            text-align: center;
            margin-bottom: 10px;
            font-size: 2.5em;
        }
        .subtitle {
            text-align: center;
            color: #666;
            margin-bottom: 30px;
            font-size: 1.1em;
        }
        .mode-section, .settings-section {
            margin: 25px 0;
            padding: 20px;
            background: #f9fafb;
            border-radius: 12px;
        }
        .mode-label, .settings-label {
            font-weight: 700;
            color: #374151;
            margin-bottom: 15px;
            font-size: 1.1em;
        }
        .mode-options {
            display: flex;
            gap: 15px;
        }
        .mode-card {
            flex: 1;
            padding: 20px;
            border: 2px solid #e5e7eb;
            border-radius: 10px;
            cursor: pointer;
            background: white;
            transition: all 0.3s;
        }
        .mode-card:hover {
            border-color: #2563eb;
            transform: translateY(-2px);
        }
        .mode-card.selected {
            border-color: #2563eb;
            background: #eff6ff;
        }
        .input-group {
            margin: 20px 0;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #333;
        }
        textarea, select, input[type="file"] {
            width: 100%;
            padding: 12px;
            border: 2px solid #e5e7eb;
            border-radius: 8px;
            font-size: 16px;
            font-family: inherit;
        }
        textarea {
            resize: vertical;
            min-height: 100px;
        }
        .btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 30px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            width: 100%;
            transition: transform 0.2s;
        }
        .btn:hover:not(:disabled) {
            transform: translateY(-2px);
        }
        .btn:disabled {
            opacity: 0.6;
            cursor: not-allowed;
        }
        .status {
            margin-top: 20px;
            padding: 15px;
            border-radius: 8px;
            display: none;
        }
        .status.show { display: block; }
        .status.loading { background: #dbeafe; color: #1e40af; }
        .status.success { background: #d1fae5; color: #065f46; }
        .status.error { background: #fee2e2; color: #991b1b; }
        .plan-review {
            display: none;
            margin-top: 30px;
            padding: 25px;
            background: #f9fafb;
            border-radius: 12px;
            border: 2px solid #e5e7eb;
        }
        .plan-review.show { display: block; }
        .plan-section {
            margin: 20px 0;
            padding: 15px;
            background: white;
            border-radius: 8px;
            border-left: 4px solid #2563eb;
        }
        .plan-section h3 {
            color: #1e40af;
            margin-bottom: 10px;
        }
        .query-list {
            list-style: none;
            padding: 10px 0;
        }
        .query-list li {
            padding: 8px;
            margin: 5px 0;
            background: #eff6ff;
            border-radius: 6px;
            font-size: 0.9em;
        }
        .action-buttons {
            display: flex;
            gap: 10px;
            margin-top: 20px;
        }
        .btn-approve {
            background: #10b981;
        }
        .btn-edit {
            background: #f59e0b;
        }
        .download-section {
            display: none;
            margin-top: 30px;
        }
        .download-section.show { display: block; }
        .download-buttons {
            display: flex;
            gap: 10px;
        }

        /* Preview & Chat Styles */
        .preview-container {
            display: none;
            margin-top: 30px;
            display: grid;
            grid-template-columns: 2fr 1fr;
            gap: 20px;
            background: #f9fafb;
            padding: 20px;
            border-radius: 12px;
            border: 1px solid #e5e7eb;
        }
        .preview-container.show { display: grid; }

        .slide-preview-area {
            background: white;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            padding: 20px;
            min-height: 400px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }

        .chat-area {
            background: white;
            border: 1px solid #d1d5db;
            border-radius: 8px;
            display: flex;
            flex-direction: column;
            height: 400px;
        }

        .chat-messages {
            flex: 1;
            padding: 15px;
            overflow-y: auto;
            background: #f9fafb;
        }

        .message {
            margin-bottom: 10px;
            padding: 8px 12px;
            border-radius: 8px;
            font-size: 0.9em;
            max-width: 85%;
        }
        .message.user {
            background: #eff6ff;
            color: #1e40af;
            align-self: flex-end;
            margin-left: auto;
        }
        .message.ai {
            background: #f3f4f6;
            color: #374151;
            align-self: flex-start;
        }

        .chat-input-area {
            padding: 10px;
            border-top: 1px solid #e5e7eb;
            display: flex;
            gap: 8px;
        }

        .slide-nav {
            display: flex;
            justify-content: space-between;
            margin-bottom: 15px;
            align-items: center;
        }

        .slide-card {
            border: 1px solid #e5e7eb;
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 6px;
        }
        .download-btn {
            flex: 1;
            padding: 12px;
            border: none;
            border-radius: 8px;
            font-weight: 600;
            cursor: pointer;
            color: white;
        }
        .btn-ppt { background: #d97706; }
        .btn-json { background: #059669; }
        .spinner {
            border: 3px solid #f3f4f6;
            border-top: 3px solid #2563eb;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
            display: none;
        }
        .spinner.show { display: block; }
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        .examples {
            margin-top: 25px;
            padding: 20px;
            background: #f9fafb;
            border-radius: 8px;
        }
        .example {
            padding: 12px;
            margin: 8px 0;
            background: white;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s;
        }
        .example:hover {
            background: #e5e7eb;
            transform: translateX(5px);
        }
        /* Settings Styles */
        /* Professional UI Updates */
        .settings-toggle {
            display: flex;
            justify-content: flex-end;
            margin-bottom: 20px;
        }
        .btn-settings {
            background: white;
            color: #4f46e5;
            border: 1px solid #e0e7ff;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: 600;
            border-radius: 30px;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        }
        .btn-settings:hover {
            border-color: #4f46e5;
            background: #f5f3ff;
            transform: translateY(-1px);
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .settings-section {
            background: white;
            border: 1px solid #e5e7eb;
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 30px;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
            animation: slideDown 0.4s cubic-bezier(0.16, 1, 0.3, 1);
        }
        .settings-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-top: 20px;
        }
        .settings-full {
            grid-column: 1 / -1;
        }
        .settings-label {
            font-size: 1.25rem;
            color: #111827;
            border-bottom: 2px solid #f3f4f6;
            padding-bottom: 12px;
            margin-bottom: 0;
        }
        input:focus, select:focus, textarea:focus {
            outline: none;
            border-color: #6366f1;
            box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.1);
        }
        .input-group label {
            text-transform: uppercase;
            letter-spacing: 0.05em;
            font-size: 0.75rem;
            color: #6b7280;
            margin-bottom: 6px;
        }
        @keyframes slideDown {
            from { opacity: 0; transform: translateY(-10px); }
            to { opacity: 1; transform: translateY(0); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="settings-toggle">
            <button onclick="toggleSettings()" class="btn btn-settings">
                ⚙️ Configure API & Model
            </button>
        </div>

        <h1>🚀 SlideDeck AI</h1>
        <p class="subtitle">Intelligent multi-agent system with review & approval workflow</p>
        
        <div id="settingsSection" class="settings-section" style="display: none;">
            <div class="settings-label">
                Configuration Settings
            </div>

            <div class="settings-grid">
                <div class="input-group" style="margin: 0;">
                    <label>LLM Provider</label>
                    <select id="llmProvider" onchange="updateModelOptions()">
                        <option value="oa">OpenAI</option>
                        <option value="an">Anthropic</option>
                        <option value="az">Azure OpenAI</option>
                        <option value="co">Cohere</option>
                        <option value="gg">Google Gemini</option>
                        <option value="ol">Ollama</option>
                        <option value="or">OpenRouter</option>
                        <option value="sn">SambaNova</option>
                        <option value="to">Together AI</option>
                    </select>
                </div>

                <div class="input-group" style="margin: 0;">
                    <label>Model</label>
                    <select id="llmModel">
                        <!-- Populated by JS -->
                    </select>
                </div>

                <div class="input-group settings-full" style="margin: 0;">
                    <label>API Key</label>
                    <input type="password" id="apiKey" placeholder="Enter API Key" style="width: 100%; padding: 12px; border: 2px solid #e5e7eb; border-radius: 8px; background: #fff;">
                </div>

                <div class="input-group settings-full" id="baseUrlGroup" style="display: none; margin: 0;">
                    <label>Base URL (Optional)</label>
                    <input type="text" id="apiBaseUrl" placeholder="https://..." style="width: 100%; padding: 12px; border: 2px solid #e5e7eb; border-radius: 8px;">
                </div>
            </div>
        </div>

        <div class="mode-section">
            <div class="mode-label">Search Mode</div>
            <div class="mode-options">
                <div class="mode-card selected" data-mode="normal" onclick="selectMode('normal')">
                    <h3>⚡ Normal</h3>
                    <p>3 queries/section • Fast generation</p>
                </div>
                <div class="mode-card" data-mode="deep" onclick="selectMode('deep')">
                    <h3>🔬 Deep</h3>
                    <p>5 queries/section • Comprehensive</p>
                </div>
            </div>
        </div>
        
        <div class="input-group">
            <label>Template Style</label>
            <select id="template">
                <option value="">Loading templates...</option>
            </select>
        </div>
        
        <div class="input-group">
            <label>Content Source</label>
            <div style="display: flex; gap: 20px; margin-bottom: 10px;">
                <label style="font-weight: normal;"><input type="radio" name="sourceType" value="search" checked onclick="toggleSource('search')"> Web Search</label>
                <label style="font-weight: normal;"><input type="radio" name="sourceType" value="file" onclick="toggleSource('file')"> Upload Files</label>
            </div>

            <div id="searchSource">
                <label>Research Query</label>
                <textarea id="query" placeholder="e.g., Tesla Q4 2024 financial performance and market position"></textarea>
            </div>

            <div id="fileSource" style="display: none;">
                <label>Upload Content Files (TXT, CSV, Excel)</label>
                <input type="file" id="contentFile" multiple accept=".txt,.csv,.xlsx,.xls">
                <small style="color: #666; margin-top: 5px; display: block;">Extracted content will be used instead of web search.</small>
                <label style="margin-top: 10px;">Topic / Subject</label>
                <input type="text" id="fileTopic" placeholder="Briefly describe the topic of the uploaded files" style="width: 100%; padding: 12px; border: 2px solid #e5e7eb; border-radius: 8px;">
            </div>
        </div>

        <div class="input-group">
             <label>Chart Data (Optional)</label>
             <input type="file" id="chartFile" accept=".png,.jpg,.jpeg,.csv,.xlsx,.xls">
             <small style="color: #666; margin-top: 5px; display: block;">Upload image, Excel, or CSV to generate charts based on data.</small>
        </div>
        
        <button class="btn" onclick="generatePlan()">🔍 Analyze & Create Plan</button>
        
        <div class="spinner" id="spinner"></div>
        <div class="status" id="status"></div>
        
        <div class="plan-review" id="planReview">
            <h2 style="margin-bottom: 20px; color: #1e40af;">📋 Research Plan Review</h2>
            <div id="planContent"></div>
            <div class="action-buttons">
                <button class="btn btn-approve" onclick="approvePlan()">✅ Approve & Generate Slides</button>
                <button class="btn btn-edit" onclick="editPlan()">✏️ Edit Plan</button>
            </div>
        </div>
        
        <div class="download-section" id="downloadSection">
            <h3 style="margin-bottom: 15px; color: #1e40af;">📥 Download Presentation</h3>
            <div class="download-buttons">
                <button class="download-btn btn-ppt" onclick="download('ppt')">📊 PowerPoint</button>
                <button class="download-btn btn-json" onclick="download('json')">📋 JSON</button>
            </div>
        </div>
        
        <div class="preview-container" id="previewContainer" style="display: none;">
            <div class="slide-preview-area">
                <div class="slide-nav">
                    <button class="btn" style="width: auto; padding: 5px 15px;" onclick="prevSlide()">◀</button>
                    <span id="slideCounter" style="font-weight: bold;">Slide 1 / 1</span>
                    <button class="btn" style="width: auto; padding: 5px 15px;" onclick="nextSlide()">▶</button>
                </div>
                <div id="slideContent" style="padding: 20px; border: 1px dashed #ccc; min-height: 300px;">
                    <!-- Slide content goes here -->
                    <h2 id="previewTitle" style="text-align: center; color: #333;">Slide Title</h2>
                    <ul id="previewBullets" style="margin-top: 20px;">
                        <li>Content loading...</li>
                    </ul>
                </div>
            </div>

            <div class="chat-area">
                <div style="padding: 10px; border-bottom: 1px solid #e5e7eb; font-weight: bold; color: #374151;">
                    💬 Refine Slide
                </div>
                <div class="chat-messages" id="chatMessages">
                    <div class="message ai">Select a slide and ask me to make changes!</div>
                </div>
                <div class="chat-input-area">
                    <input type="text" id="chatInput" placeholder="e.g., Make the title bolder..." style="flex: 1; padding: 8px; border: 1px solid #d1d5db; border-radius: 4px;">
                    <button onclick="sendChat()" class="btn" style="width: auto; padding: 8px 12px;">➤</button>
                </div>
            </div>
        </div>

        <div class="examples">
            <h3 style="margin-bottom: 15px;">💡 Example Queries</h3>
            <div class="example" onclick="setQuery('Apple Inc financial performance and market analysis Q4 2024')">
                🍎 Apple Inc financial performance and market analysis Q4 2024
            </div>
            <div class="example" onclick="setQuery('Global electric vehicle market trends and competitive landscape 2024')">
                🚗 Global electric vehicle market trends and competitive landscape 2024
            </div>
            <div class="example" onclick="setQuery('Artificial Intelligence in healthcare: applications, market size, and future outlook')">
                🏥 AI in healthcare: applications, market size, and future outlook
            </div>
        </div>
    </div>
    
    <script>
        let selectedMode = 'normal';
        let currentPlan = null;
        let reportId = null;
        let templateOptions = {};
        let planSectionsCollapsed = false;
        let validModels = {};
        let currentPreviewSlides = [];
        let currentSlideIndex = 0;

        // Valid models from backend config (simplified mapping for frontend)
        const MODEL_OPTIONS = {
            'an': ['claude-haiku-4-5'],
            'az': ['azure/open-ai'],
            'co': ['command-r-08-2024'],
            'gg': ['gemini-2.0-flash', 'gemini-2.0-flash-lite', 'gemini-2.5-flash', 'gemini-2.5-flash-lite'],
            'oa': ['gpt-4.1-mini', 'gpt-4.1-nano', 'gpt-5-nano'],
            'or': ['google/gemini-2.0-flash-001', 'openai/gpt-3.5-turbo'],
            'sn': ['DeepSeek-V3.1-Terminus', 'Llama-3.3-Swallow-70B-Instruct-v0.4'],
            'to': ['deepseek-ai/DeepSeek-V3', 'meta-llama/Llama-3.3-70B-Instruct-Turbo', 'meta-llama/Meta-Llama-3.1-8B-Instruct-Turbo-128K'],
            'ol': ['llama3'] // Example for ollama
        };

        function toggleSettings() {
            const el = document.getElementById('settingsSection');
            el.style.display = el.style.display === 'none' ? 'block' : 'none';
        }

        function updateModelOptions() {
            const provider = document.getElementById('llmProvider').value;
            const modelSelect = document.getElementById('llmModel');
            const baseUrlGroup = document.getElementById('baseUrlGroup');

            modelSelect.innerHTML = '';

            // Show Base URL for certain providers if needed (e.g. Azure, Ollama)
            if (provider === 'az' || provider === 'ol') {
                baseUrlGroup.style.display = 'block';
            } else {
                baseUrlGroup.style.display = 'none';
            }

            const models = MODEL_OPTIONS[provider] || [];
            models.forEach(m => {
                const opt = document.createElement('option');
                opt.value = `[${provider}]${m}`; // Match format in GlobalConfig
                opt.textContent = m;
                modelSelect.appendChild(opt);
            });

            // Trigger selection of first model
            if (models.length > 0) modelSelect.value = `[${provider}]${models[0]}`;
        }

        // Initialize models on load
        window.addEventListener('DOMContentLoaded', () => {
            updateModelOptions();
        });
        
        // Function to load templates from the backend
        async function loadTemplates() {
            console.log('🔄 Loading templates from /api/templates...');
            try {
                const response = await fetch('/api/templates');
                console.log('📡 Response status:', response.status);
                
                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }
                
                templateOptions = await response.json();
                console.log('✅ Templates received:', templateOptions);
                
                const templateSelect = document.getElementById('template');
                templateSelect.innerHTML = '';
                
                // Check if we got valid data
                if (!templateOptions || Object.keys(templateOptions).length === 0) {
                    throw new Error('No templates returned from server');
                }
                
                // Add options from the fetched data
                Object.keys(templateOptions).forEach(key => {
                    const option = document.createElement('option');
                    option.value = key;                 // value sent to backend
                    option.textContent = key;           // ✅ show the key, not caption
                    // if you want both:
                    // option.textContent = `${key} – ${templateOptions[key].caption}`;
                    templateSelect.appendChild(option);
                });
                
                // Set default selection
                templateSelect.value = Object.keys(templateOptions)[0];
                console.log('✅ Templates loaded successfully');
                
            } catch (error) {
                console.error('❌ Error loading templates:', error);
                // Fallback to hardcoded options
                const templateSelect = document.getElementById('template');
                    templateSelect.innerHTML = `
                        <option value="Basic">Basic</option>
                        <option value="Ion Boardroom">Ion Boardroom</option>
                        <option value="Minimalist Sales Pitch">Minimalist Sales Pitch</option>
                        <option value="Urban Monochrome">Urban Monochrome</option>
                        <option value="RRD Template">RRD Template</option>
                        <option value="WilliamsLea">WilliamsLea</option>
                    `;
                console.log('⚠️ Using fallback templates');
            }
        }
        
        // Call loadTemplates when the page loads
        window.addEventListener('DOMContentLoaded', () => {
            console.log('🚀 Page loaded, initializing...');
            loadTemplates();
        });
        
        function selectMode(mode) {
            selectedMode = mode;
            document.querySelectorAll('.mode-card').forEach(card => {
                card.classList.toggle('selected', card.dataset.mode === mode);
            });
        }
        
        function toggleSource(type) {
            if (type === 'search') {
                document.getElementById('searchSource').style.display = 'block';
                document.getElementById('fileSource').style.display = 'none';
            } else {
                document.getElementById('searchSource').style.display = 'none';
                document.getElementById('fileSource').style.display = 'block';
            }
        }

        function setQuery(text) {
            document.getElementById('query').value = text;
            // Ensure search mode is selected
            document.querySelector('input[name="sourceType"][value="search"]').click();
        }
        
        function showStatus(msg, type) {
            const status = document.getElementById('status');
            status.textContent = msg;
            status.className = 'status show ' + type;
        }
        
        async function generatePlan() {
            const sourceType = document.querySelector('input[name="sourceType"]:checked').value;
            let query = '';
            let formData = new FormData();
            
            const template = document.getElementById('template').value;
            formData.append('template', template);
            formData.append('search_mode', selectedMode);
            
            // Add Settings
            const provider = document.getElementById('llmProvider').value;
            const model = document.getElementById('llmModel').value;
            const apiKey = document.getElementById('apiKey').value;
            const apiBase = document.getElementById('apiBaseUrl').value;

            if (apiKey) formData.append('api_key', apiKey);
            if (model) formData.append('llm_model', model);
            if (apiBase) formData.append('api_base', apiBase);

            if (sourceType === 'search') {
                query = document.getElementById('query').value.trim();
                if (!query) {
                    showStatus('⚠️ Please enter a research query', 'error');
                    return;
                }
                formData.append('query', query);
            } else {
                const files = document.getElementById('contentFile').files;
                if (files.length === 0) {
                    showStatus('⚠️ Please upload at least one file', 'error');
                    return;
                }
                for (let i = 0; i < files.length; i++) {
                    formData.append('files', files[i]);
                }
                query = document.getElementById('fileTopic').value.trim();
                if (!query) {
                     showStatus('⚠️ Please enter a topic for the files', 'error');
                     return;
                }
                formData.append('query', query);
            }

            // Chart file
            const chartFile = document.getElementById('chartFile').files[0];
            if (chartFile) {
                formData.append('chart_file', chartFile);
            }

            document.getElementById('spinner').classList.add('show');
            document.getElementById('planReview').classList.remove('show');
            showStatus('🔍 Analyzing input and generating research plan...', 'loading');
            
            try {
                console.log('🚀 Sending request to /api/plan');
                
                const response = await fetch('/api/plan', {
                    method: 'POST',
                    body: formData // Send as FormData
                });
                
                console.log('📡 Response received');
                console.log('📡 Response status:', response.status);
                console.log('📡 Response ok:', response.ok);
                console.log('📡 Response headers:', [...response.headers.entries()]);
                
                if (!response.ok) {
                    const errorText = await response.text();
                    console.error('❌ Server error response:', errorText);
                    throw new Error('Plan generation failed: ' + errorText);
                }
                
                const responseText = await response.text();
                console.log('📦 Raw response text:', responseText);
                
                let responseData;
                try {
                    responseData = JSON.parse(responseText);
                } catch (parseError) {
                    console.error('❌ JSON parse error:', parseError);
                    console.error('❌ Failed to parse:', responseText.substring(0, 200));
                    throw new Error('Invalid JSON response from server');
                }
                
                console.log('📦 Parsed response:', responseData);
                console.log('📦 Response type:', typeof responseData);
                console.log('📦 Response keys:', Object.keys(responseData));
                console.log('📦 Has sections?:', 'sections' in responseData);
                console.log('📦 Sections value:', responseData.sections);
                console.log('📦 Sections type:', typeof responseData.sections);
                console.log('📦 Sections is array?:', Array.isArray(responseData.sections));
                console.log('📦 Sections length:', responseData.sections?.length);
                
                if (!responseData.sections) {
                    console.error('❌ responseData.sections is falsy:', responseData.sections);
                    throw new Error('Response missing sections field');
                }
                
                if (!Array.isArray(responseData.sections)) {
                    console.error('❌ responseData.sections is not an array:', typeof responseData.sections);
                    throw new Error('Response sections is not an array: ' + typeof responseData.sections);
                }
                
                if (responseData.sections.length === 0) {
                    console.warn('⚠️ responseData.sections is empty array');
                }
                
                currentPlan = responseData;
                console.log('✅ Set currentPlan:', currentPlan);
                
                displayPlan(currentPlan);
                
                document.getElementById('spinner').classList.remove('show');
                showStatus('✅ Research plan ready for review!', 'success');
                document.getElementById('planReview').classList.add('show');
                
            } catch (error) {
                document.getElementById('spinner').classList.remove('show');
                console.error('❌ Full error object:', error);
                console.error('❌ Error stack:', error.stack);
                showStatus('❌ Error: ' + error.message, 'error');
            }
        }
        
        function displayPlan(plan) {
            if (reportId === null) planSectionsCollapsed = false;
            const content = document.getElementById('planContent');
            
            // ✅ SAFETY CHECK: Ensure plan.sections exists and is an array
            if (!plan || !plan.sections || !Array.isArray(plan.sections)) {
                console.error('❌ Invalid plan structure:', plan);
                content.innerHTML = `
                    <div style="padding: 20px; background: #fee2e2; border-radius: 8px; color: #991b1b;">
                        <strong>⚠️ Error:</strong> Invalid plan structure received from server.
                        <br><small>Please try generating the plan again.</small>
                    </div>
                `;
                return;
            }
            
            console.log('✅ displayPlan called with valid plan:', plan);
            console.log('✅ Number of sections:', plan.sections.length);
            
            let html = `
                <div style="margin-bottom: 20px; padding: 15px; background: #eff6ff; border-radius: 8px;">
                    <div style="display: flex; justify-content: space-between; align-items: center;">
                        <div>
                            <h4>Report Type: ${plan.analysis?.report_type || 'N/A'}</h4>
                            <p><strong>Subject:</strong> ${plan.analysis?.core_subject || plan.query}</p>
                            <p><strong>Total Queries:</strong> ${plan.total_queries || 0}</p>
                            <p><strong>Template:</strong> ${plan.template || 'Default'}</p>
                        </div>
                        <button id="togglePlanBtn" onclick="togglePlanSections()" 
                                style="display: inline-block; background: #6b7280; color: white; border: none; padding: 8px 16px; border-radius: 6px; cursor: pointer; font-size: 13px;">
                            ▼ Collapse All
                        </button>
                    </div>
                </div>
                <div id="planSectionsContainer" style="display: block;">
            `;
        
            // ✅ FIX: Build sections HTML safely
            plan.sections.forEach((section, idx) => {
                console.log(`Processing section ${idx}:`, section.section_title);
                
                // Extract search queries from placeholder_specs
                let searchQueries = [];
                
                if (section.placeholder_specs && Array.isArray(section.placeholder_specs)) {
                    section.placeholder_specs.forEach(spec => {
                        if (spec.search_queries && Array.isArray(spec.search_queries)) {
                            searchQueries = searchQueries.concat(spec.search_queries);
                        }
                    });
                }
                
                console.log(`  Section ${idx} has ${searchQueries.length} search queries`);
                
                // ✅ Build query list HTML separately
                let queryListHtml = '';
                if (searchQueries.length > 0) {
                    queryListHtml = searchQueries.map(q => 
                        `<li>🔍 ${q.query || 'No query'}<br>
                        <small style="color: #6b7280;">Purpose: ${q.purpose || 'N/A'}</small></li>`
                    ).join('');
                } else {
                    queryListHtml = '<li style="color: #9ca3af;">No search queries defined</li>';
                }
                
                // ✅ Build placeholder specs HTML separately
                let placeholderSpecsHtml = '';
                if (section.placeholder_specs && section.placeholder_specs.length > 0) {
                    const specsHtml = section.placeholder_specs.map(spec => 
                        `<div style="background: #f9fafb; padding: 10px; margin: 5px 0; border-radius: 6px; border-left: 3px solid #3b82f6;">
                            <strong>Placeholder ${spec.placeholder_idx}</strong> (${spec.placeholder_type})<br>
                            <small style="color: #6b7280;">
                                Content Type: ${spec.content_type}<br>
                                Description: ${spec.content_description}
                            </small>
                        </div>`
                    ).join('');
                    
                    placeholderSpecsHtml = `
                        <details style="margin-top: 10px;">
                            <summary style="cursor: pointer; font-weight: 600; color: #374151;">
                                📋 Placeholder Specifications (${section.placeholder_specs.length})
                            </summary>
                            <div style="margin-top: 8px; padding-left: 10px;">
                                ${specsHtml}
                            </div>
                        </details>
                    `;
                }
                
                // ✅ Now safely build the section HTML
                html += `
                    <div class="plan-section" id="plan_section_${idx}">
                        <h3>${idx + 1}. ${section.section_title || 'Untitled Section'}</h3>
                        <p style="color: #6b7280; font-size: 0.9em; margin: 10px 0;">
                            ${section.section_purpose || 'No purpose specified'}
                        </p>
                        <p style="margin: 10px 0;">
                            <strong>Layout:</strong> ${section.layout_type || 'N/A'} (Index: ${section.layout_idx || 'N/A'})
                        </p>
                        
                        <details style="margin-top: 10px;">
                            <summary style="cursor: pointer; font-weight: 600; color: #374151;">
                                🔍 Search Queries (${searchQueries.length})
                            </summary>
                            <ul class="query-list" style="margin-top: 8px;">
                                ${queryListHtml}
                            </ul>
                        </details>
                        
                        ${placeholderSpecsHtml}
                    </div>
                `;
            });
        
            html += '</div>';
            
            console.log('✅ Setting innerHTML');
            content.innerHTML = html;
            
            console.log('✅ Setting collapsed state');
            setPlanSectionsCollapsed(planSectionsCollapsed);
            
            console.log('✅ displayPlan completed successfully');
        }
        
        function setPlanSectionsCollapsed(collapsed) {
            const container = document.getElementById('planSectionsContainer');
            const btn = document.getElementById('togglePlanBtn');
        
            if (!container || !btn) {
                console.warn('setPlanSectionsCollapsed: elements not found');
                return;
            }
        
            planSectionsCollapsed = collapsed;
        
            if (planSectionsCollapsed) {
                container.style.display = 'none';
                btn.textContent = '▶ Expand All';
            } else {
                container.style.display = 'block';
                btn.textContent = '▼ Collapse All';
            }
        }
        
        function togglePlanSections() {
            planSectionsCollapsed = !planSectionsCollapsed;
            setPlanSectionsCollapsed(planSectionsCollapsed);     // 🔥 force UI update after toggle
            const btn = document.getElementById('togglePlanBtn');
            if (btn) btn.blur();          // optional UX (prevents button staying focused)
        }

        function approvePlan() {
            console.log('🔍 Current plan:', currentPlan);
            if (!currentPlan || !currentPlan.plan_id) {
                console.error('❌ No plan_id found!', currentPlan);  // ✅ DEBUG
                showStatus('❌ No plan available to execute', 'error');
                return;
            }
            
            console.log('✅ Sending plan_id:', currentPlan.plan_id); 
            
            // Collapse sections BEFORE starting generation
            planSectionsCollapsed = true;
            setPlanSectionsCollapsed(planSectionsCollapsed);
            
            document.getElementById('spinner').classList.add('show');
            showStatus('🚀 Generating slides with SlideDeck AI...', 'loading');
            
            fetch('/api/execute', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({ 
                    plan_id: currentPlan.plan_id  // ✅ FIXED - send plan_id only
                })
            })
            .then(response => {
                if (!response.ok) {
                    return response.json().then(err => {
                        throw new Error(err.error || 'Slide generation failed');
                    });
                }
                return response.json();
            })
            .then(result => {
                reportId = result.report_id;
                document.getElementById('spinner').classList.remove('show');
                showStatus(`✅ Slides generated successfully! (${result.slides_generated} slides in ${result.execution_time})`, 'success');
                document.getElementById('downloadSection').classList.add('show');
                loadPreview(reportId);
            })
            .catch(error => {
                document.getElementById('spinner').classList.remove('show');
                showStatus('❌ Error: ' + error.message, 'error');
                console.error('Execution error:', error);
            });
        }
        
        function editPlan() {
            const content = document.getElementById('planContent');
            let html = '<h3 style="margin-bottom: 20px;">✏️ Edit Research Plan</h3>';
            
            // ✅ Safety check
            if (!currentPlan || !currentPlan.sections || !Array.isArray(currentPlan.sections)) {
                html += '<p style="color: #dc2626;">Error: Invalid plan structure. Cannot edit.</p>';
                content.innerHTML = html;
                return;
            }
            
            currentPlan.sections.forEach((section, idx) => {
                // ✅ Extract search queries from NEW FORMAT (placeholder_specs)
                let searchQueries = [];
                
                if (section.placeholder_specs && Array.isArray(section.placeholder_specs)) {
                    // Flatten all queries from all placeholder_specs
                    section.placeholder_specs.forEach(spec => {
                        if (spec.search_queries && Array.isArray(spec.search_queries)) {
                            searchQueries = searchQueries.concat(spec.search_queries);
                        }
                    });
                }
                
                // ✅ Store in OLD format for backward compatibility with edit UI
                section.search_queries = searchQueries;
                
                html += `
                    <div class="plan-section" id="section_${idx}" style="margin: 15px 0; position: relative;">
                        <button onclick="deleteSection(${idx})" 
                                style="position: absolute; top: 10px; right: 10px; background: #dc2626; color: white; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 12px;">
                            🗑️ Delete
                        </button>
                        
                        <label style="font-weight: 600; color: #374151;">Section ${idx + 1} Title:</label>
                        <input type="text" id="section_${idx}_title" value="${escapeHtml(section.section_title || '')}" 
                               style="width: 100%; padding: 8px; margin: 5px 0; border: 2px solid #e5e7eb; border-radius: 6px;">
                        
                        <label style="font-weight: 600; color: #374151; margin-top: 10px; display: block;">Purpose:</label>
                        <textarea id="section_${idx}_purpose" 
                                  style="width: 100%; padding: 8px; margin: 5px 0; border: 2px solid #e5e7eb; border-radius: 6px; min-height: 60px;">${escapeHtml(section.section_purpose || '')}</textarea>
                        
                        <label style="font-weight: 600; color: #374151; margin-top: 10px; display: block;">Layout Type:</label>
                        <input type="text" id="section_${idx}_layout" value="${escapeHtml(section.layout_type || 'single_column')}" 
                               style="width: 100%; padding: 8px; margin: 5px 0; border: 2px solid #e5e7eb; border-radius: 6px;"
                               placeholder="e.g., chart_layout, table_layout, double_column">
                        
                        <div style="margin-top: 10px;">
                            <label style="font-weight: 600; color: #374151;">Search Queries:</label>
                            <div id="queries_${idx}">
                                ${searchQueries.map((q, qIdx) => `
                                    <div class="query-item" id="query_item_${idx}_${qIdx}" style="background: #f3f4f6; padding: 10px; margin: 5px 0; border-radius: 6px; position: relative;">
                                        <button onclick="deleteQuery(${idx}, ${qIdx})" 
                                                style="position: absolute; top: 5px; right: 5px; background: #ef4444; color: white; border: none; padding: 3px 8px; border-radius: 3px; cursor: pointer; font-size: 11px;">
                                            ✕
                                        </button>
                                        <label style="font-size: 11px; color: #6b7280; display: block; margin-bottom: 3px;">Search Query:</label>
                                        <input type="text" id="query_${idx}_${qIdx}" value="${escapeHtml(q.query || '')}" 
                                               style="width: calc(100% - 30px); padding: 6px; margin: 2px 0; border: 1px solid #d1d5db; border-radius: 4px; font-size: 13px;">
                                        <label style="font-size: 11px; color: #6b7280; display: block; margin-bottom: 3px; margin-top: 5px;">Purpose:</label>
                                        <input type="text" id="query_purpose_${idx}_${qIdx}" value="${escapeHtml(q.purpose || '')}" 
                                               placeholder="Query purpose..."
                                               style="width: calc(100% - 30px); padding: 6px; margin: 2px 0; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px; color: #6b7280;">
                                        <label style="font-size: 11px; color: #6b7280; display: block; margin-bottom: 3px; margin-top: 5px;">Expected Source:</label>
                                        <select id="query_source_${idx}_${qIdx}" 
                                                style="width: calc(100% - 30px); padding: 6px; margin: 2px 0; border: 1px solid #d1d5db; border-radius: 4px; font-size: 12px;">
                                            <option value="research" ${q.expected_source_type === 'research' ? 'selected' : ''}>Research</option>
                                            <option value="news" ${q.expected_source_type === 'news' ? 'selected' : ''}>News</option>
                                            <option value="data" ${q.expected_source_type === 'data' ? 'selected' : ''}>Data</option>
                                            <option value="financial" ${q.expected_source_type === 'financial' ? 'selected' : ''}>Financial</option>
                                            <option value="expert" ${q.expected_source_type === 'expert' ? 'selected' : ''}>Expert</option>
                                        </select>
                                    </div>
                                `).join('')}
                            </div>
                            <button onclick="addQuery(${idx})" 
                                    style="background: #10b981; color: white; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 12px; margin-top: 5px;">
                                ➕ Add Query
                            </button>
                        </div>
                    </div>
                `;
            });
            
            html += `
                <div style="margin: 20px 0; text-align: center;">
                    <button onclick="addSection()" 
                            style="background: #2563eb; color: white; border: none; padding: 12px 24px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600;">
                        ➕ Add New Section
                    </button>
                </div>
                <div style="display: flex; gap: 10px; margin-top: 20px;">
                    <button class="btn" onclick="saveEdits()" style="flex: 1;">💾 Save Changes</button>
                    <button class="btn" onclick="cancelEdits()" style="flex: 1; background: #6b7280;">❌ Cancel</button>
                </div>
            `;
            
            content.innerHTML = html;
            showStatus('✏️ Editing mode active - update sections above', 'loading');
        }

        function updatePlanSectionsView() {
            const container = document.getElementById('planSectionsContainer');
            const btn = document.getElementById('togglePlanBtn');
        
            if (!container || !btn) {
                console.warn('updatePlanSectionsView: elements not found');
                return;
            }
        
            if (planSectionsCollapsed) {
                container.style.display = 'none';
                btn.textContent = '▶ Expand All';
            } else {
                container.style.display = 'block';
                btn.textContent = '▼ Collapse All';
            }
        }

        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }
        
        function addSection() {
            // Create a new section with default values
            const newSection = {
                section_title: "New Section",
                section_purpose: "Describe the purpose of this section",
                visualization_hint: "bullets",
                search_queries: [
                    {
                        query: "Enter search query here",
                        purpose: "what this query targets",
                        expected_source_type: "research"
                    }
                ]
            };
            
            // Add to current plan
            currentPlan.sections.push(newSection);
            
            // Refresh the edit view
            editPlan();
            
            showStatus('✅ New section added! Scroll down to edit it.', 'success');
            
            // Scroll to the new section
            setTimeout(() => {
                const newSectionId = `section_${currentPlan.sections.length - 1}`;
                const element = document.getElementById(newSectionId);
                if (element) {
                    element.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    element.style.border = '2px solid #2563eb';
                    setTimeout(() => {
                        element.style.border = '';
                    }, 2000);
                }
            }, 100);
        }
        
        function deleteSection(sectionIdx) {
            if (currentPlan.sections.length <= 1) {
                showStatus('⚠️ Cannot delete the last section!', 'error');
                return;
            }
            
            const sectionTitle = currentPlan.sections[sectionIdx].section_title;
            
            if (confirm(`Are you sure you want to delete "${sectionTitle}"?`)) {
                // Remove the section
                currentPlan.sections.splice(sectionIdx, 1);
                
                // Refresh the edit view
                editPlan();
                
                showStatus(`✅ Section "${sectionTitle}" deleted.`, 'success');
            }
        }
        
        function addQuery(sectionIdx) {
            const newQuery = {
                query: "Enter new search query",
                purpose: "what this query targets",
                expected_source_type: "research"
            };
            
            // Add to section's queries
            currentPlan.sections[sectionIdx].search_queries.push(newQuery);
            
            // Refresh the edit view
            editPlan();
            
            showStatus('✅ New query added to section!', 'success');
            
            // Auto-scroll to the new query
            setTimeout(() => {
                const queriesDiv = document.getElementById(`queries_${sectionIdx}`);
                if (queriesDiv) {
                    queriesDiv.lastElementChild.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                }
            }, 100);
        }
        
        function deleteQuery(sectionIdx, queryIdx) {
            const section = currentPlan.sections[sectionIdx];
            
            if (section.search_queries.length <= 1) {
                showStatus('⚠️ Each section must have at least one query!', 'error');
                return;
            }
            
            if (confirm('Delete this search query?')) {
                // Remove the query
                section.search_queries.splice(queryIdx, 1);
                
                // Refresh the edit view
                editPlan();
                
                showStatus('✅ Query deleted.', 'success');
            }
        }
        
        function cancelEdits() {
            // Revert the plan display to the current state (before editing)
            displayPlan(currentPlan);
            showStatus('📋 Edit cancelled. Plan is restored.', 'loading');
        }
        
        function saveEdits() {
            try {
                // Create a new array for updated sections
                const updatedSections = [];
                
                // Loop through current sections
                for (let idx = 0; idx < currentPlan.sections.length; idx++) {
                    const section = currentPlan.sections[idx];
                    
                    // Check if section still exists in DOM (not deleted)
                    const titleInput = document.getElementById(`section_${idx}_title`);
                    if (!titleInput) {
                        // Section was deleted, skip it
                        continue;
                    }
                    
                    // Update section details
                    section.section_title = titleInput.value.trim();
                    section.section_purpose = document.getElementById(`section_${idx}_purpose`).value.trim();
                    section.visualization_hint = document.getElementById(`section_${idx}_viz`).value;
                    
                    // Update queries for this section
                    const updatedQueries = [];
                    for (let qIdx = 0; qIdx < section.search_queries.length; qIdx++) {
                        const queryInput = document.getElementById(`query_${idx}_${qIdx}`);
                        const purposeInput = document.getElementById(`query_purpose_${idx}_${qIdx}`);
                        const sourceInput = document.getElementById(`query_source_${idx}_${qIdx}`);
                        
                        if (queryInput && purposeInput && sourceInput) {
                            updatedQueries.push({
                                query: queryInput.value.trim(),
                                purpose: purposeInput.value.trim(),
                                expected_source_type: sourceInput.value
                            });
                        }
                    }
                    
                    // Validate: must have at least one query
                    if (updatedQueries.length === 0) {
                        throw new Error(`Section "${section.section_title}" must have at least one search query.`);
                    }
                    
                    // Validate: fields must not be empty
                    if (!section.section_title || !section.section_purpose) {
                        throw new Error('Section title and purpose cannot be empty.');
                    }
                    
                    section.search_queries = updatedQueries;
                    updatedSections.push(section);
                }
                
                // Validate: must have at least one section
                if (updatedSections.length === 0) {
                    throw new Error('Plan must have at least one section.');
                }
                
                // Update the plan with cleaned sections
                currentPlan.sections = updatedSections;
                currentPlan.total_queries = currentPlan.sections.reduce(
                    (sum, section) => sum + section.search_queries.length, 
                    0
                );
                
                planSectionsCollapsed = true;
        
                // Refresh the display (it will use the collapsed state)
                displayPlan(currentPlan);
                
                showStatus('✅ Changes saved! Review and approve to generate report.', 'success');
                
            } catch (error) {
                showStatus('❌ Error saving changes: ' + error.message, 'error');
                console.error('Save error:', error);
            }
        }
        
        function download(format) {
            if (!reportId) {
                showStatus('❌ No report available to download', 'error');
                return;
            }
            
            showStatus(`📥 Preparing ${format.toUpperCase()} download...`, 'loading');
            
            fetch(`/api/download/${reportId}?format=${format}`)
            .then(response => {
                if (!response.ok) throw new Error('Download failed');
                
                if (format === 'json') {
                    return response.json().then(data => {
                        const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = `report_${reportId}.json`;
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        document.body.removeChild(a);
                    });
                } else {
                    return response.blob().then(blob => {
                        const url = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = url;
                        a.download = `report_${reportId}.pptx`;
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(url);
                        document.body.removeChild(a);
                    });
                }
            })
            .then(() => {
                showStatus(`✅ ${format.toUpperCase()} downloaded successfully!`, 'success');
            })
            .catch(error => {
                showStatus(`❌ Download failed: ${error.message}`, 'error');
            });
        }

        // Preview & Chat Functions
        function loadPreview(id) {
            fetch(`/api/preview/${id}`)
            .then(res => res.json())
            .then(data => {
                if(data.slides) {
                    currentPreviewSlides = data.slides;
                    currentSlideIndex = 0;
                    document.getElementById('previewContainer').style.display = 'grid';
                    renderSlide(0);
                }
            })
            .catch(err => console.error("Preview load failed", err));
        }

        function renderSlide(index) {
            if(!currentPreviewSlides || currentPreviewSlides.length === 0) return;
            const slide = currentPreviewSlides[index];
            document.getElementById('slideCounter').textContent = `Slide ${index + 1} / ${currentPreviewSlides.length}`;
            document.getElementById('previewTitle').textContent = slide.title || 'Untitled';

            const list = document.getElementById('previewBullets');
            list.innerHTML = '';

            if (slide.content && Array.isArray(slide.content)) {
                slide.content.forEach(item => {
                    const li = document.createElement('li');
                    li.textContent = item;
                    list.appendChild(li);
                });
            } else {
                list.innerHTML = '<li>(Visual Content)</li>';
            }
        }

        function prevSlide() {
            if(currentSlideIndex > 0) {
                currentSlideIndex--;
                renderSlide(currentSlideIndex);
            }
        }

        function nextSlide() {
            if(currentSlideIndex < currentPreviewSlides.length - 1) {
                currentSlideIndex++;
                renderSlide(currentSlideIndex);
            }
        }

        async function sendChat() {
            const input = document.getElementById('chatInput');
            const msg = input.value.trim();
            if(!msg || !reportId) return;

            const chatBox = document.getElementById('chatMessages');
            chatBox.innerHTML += `<div class="message user">${msg}</div>`;
            input.value = '';

            try {
                const res = await fetch('/api/chat', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        report_id: reportId,
                        slide_idx: currentSlideIndex,
                        instruction: msg
                    })
                });
                const data = await res.json();

                chatBox.innerHTML += `<div class="message ai">${data.message}</div>`;
                chatBox.scrollTop = chatBox.scrollHeight;

                // If demo, update content locally
                if(data.updated_content) {
                    currentPreviewSlides[currentSlideIndex].title = data.updated_content.title;
                    currentPreviewSlides[currentSlideIndex].content = data.updated_content.bullets;
                    renderSlide(currentSlideIndex);
                }

            } catch(e) {
                chatBox.innerHTML += `<div class="message ai" style="color:red;">Error: ${e.message}</div>`;
            }
        }
    </script>
</body>
</html>
"""
