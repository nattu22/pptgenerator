HTML_UI = """
<!DOCTYPE html>
<html>
<head>
    <title>SlideDeck AI - Advanced Report Generator</title>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        body { font-family: 'Inter', sans-serif; background-color: #f3f4f6; }
        .sidebar { width: 280px; height: 100vh; position: fixed; left: 0; top: 0; background: white; border-right: 1px solid #e5e7eb; display: flex; flex-direction: column; z-index: 50; }
        .main-content { margin-left: 280px; padding: 40px; min-height: 100vh; transition: margin-left 0.3s; }
        .nav-item { padding: 12px 20px; color: #4b5563; font-weight: 500; cursor: pointer; border-left: 4px solid transparent; transition: all 0.2s; display: flex; align-items: center; gap: 10px; }
        .nav-item:hover { background-color: #f9fafb; color: #111827; }
        .nav-item.active { background-color: #eff6ff; color: #2563eb; border-left-color: #2563eb; }
        .card { background: white; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border: 1px solid #e5e7eb; transition: transform 0.2s; }
        .step-container { display: none; }
        .step-container.active { display: block; animation: fadeIn 0.3s; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
        .spinner { border: 3px solid #f3f4f6; border-top: 3px solid #2563eb; border-radius: 50%; width: 24px; height: 24px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

        /* Preview Card Styles */
        .slide-preview {
            aspect-ratio: 16/9;
            background: white;
            border: 1px solid #e5e7eb;
            box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
            position: relative;
            overflow: hidden;
            border-radius: 8px;
            display: flex;
            flex-direction: column;
        }
        .slide-preview-header { background: #f8fafc; padding: 15px; border-bottom: 1px solid #e5e7eb; }
        .slide-preview-body { padding: 20px; flex: 1; overflow-y: auto; font-size: 0.9em; }
        .slide-preview-footer { padding: 10px; background: #f8fafc; border-top: 1px solid #e5e7eb; display: flex; justify-content: space-between; align-items: center; }

        /* Chat Styles */
        .chat-container { height: 500px; display: flex; flex-direction: column; border: 1px solid #e5e7eb; border-radius: 12px; background: white; }
        .chat-messages { flex: 1; padding: 20px; overflow-y: auto; background: #f9fafb; }
        .chat-input-area { padding: 15px; border-top: 1px solid #e5e7eb; background: white; border-radius: 0 0 12px 12px; }
        .message { margin-bottom: 15px; max-width: 80%; padding: 10px 15px; border-radius: 12px; font-size: 0.95em; line-height: 1.5; }
        .message.user { background: #2563eb; color: white; align-self: flex-end; margin-left: auto; border-bottom-right-radius: 4px; }
        .message.ai { background: white; border: 1px solid #e5e7eb; align-self: flex-start; margin-right: auto; border-bottom-left-radius: 4px; }

        .selected-slide { border: 3px solid #2563eb !important; box-shadow: 0 0 15px rgba(37, 99, 235, 0.2); }
    </style>
</head>
<body>

    <!-- Sidebar -->
    <div class="sidebar">
        <div class="p-6 border-b border-gray-200">
            <h1 class="text-2xl font-bold text-blue-600 flex items-center gap-2">
                <i class="fa-solid fa-layer-group"></i> SlideDeck AI
            </h1>
            <p class="text-xs text-gray-500 mt-1">Intelligent Presentation Agent</p>
        </div>
        <nav class="flex-1 py-4">
            <div class="nav-item active" onclick="switchTab('create')">
                <i class="fa-solid fa-wand-magic-sparkles w-5"></i> Create New
            </div>
            <div class="nav-item" onclick="switchTab('plan')" id="nav-plan" style="opacity: 0.5; pointer-events: none;">
                <i class="fa-solid fa-list-check w-5"></i> Research Plan
            </div>
            <div class="nav-item" onclick="switchTab('editor')" id="nav-editor" style="opacity: 0.5; pointer-events: none;">
                <i class="fa-solid fa-pen-to-square w-5"></i> Editor & Preview
            </div>
            <div class="nav-item" onclick="switchTab('download')" id="nav-download" style="opacity: 0.5; pointer-events: none;">
                <i class="fa-solid fa-download w-5"></i> Download
            </div>
        </nav>
        <div class="p-4 border-t border-gray-200">
            <div class="text-xs text-gray-400 text-center">v2.5.0 Production</div>
        </div>
    </div>

    <!-- Main Content -->
    <div class="main-content">
        
        <!-- Tab: Create -->
        <div id="tab-create" class="step-container active">
            <div class="max-w-4xl mx-auto">
                <h2 class="text-3xl font-bold text-gray-800 mb-2">Create New Presentation</h2>
                <p class="text-gray-500 mb-8">Configure your requirements and let AI handle the research and design.</p>

                <div class="card p-8 mb-8">
                    <!-- Report Type -->
                    <div class="mb-8">
                        <label class="block text-sm font-semibold text-gray-700 mb-3">Presentation Type</label>
                        <div class="grid grid-cols-4 gap-4">
                            <div class="mode-card cursor-pointer border-2 border-blue-500 bg-blue-50 rounded-lg p-4 transition hover:shadow-md text-center" onclick="selectType(this, 'sales')">
                                <div class="text-2xl mb-2">🚀</div>
                                <div class="font-bold text-blue-700">Sales Pitch</div>
                            </div>
                            <div class="mode-card cursor-pointer border-2 border-transparent bg-gray-50 rounded-lg p-4 transition hover:shadow-md text-center" onclick="selectType(this, 'executive')">
                                <div class="text-2xl mb-2">👔</div>
                                <div class="font-bold text-gray-700">Executive</div>
                            </div>
                            <div class="mode-card cursor-pointer border-2 border-transparent bg-gray-50 rounded-lg p-4 transition hover:shadow-md text-center" onclick="selectType(this, 'professional')">
                                <div class="text-2xl mb-2">💼</div>
                                <div class="font-bold text-gray-700">Professional</div>
                            </div>
                            <div class="mode-card cursor-pointer border-2 border-transparent bg-gray-50 rounded-lg p-4 transition hover:shadow-md text-center" onclick="selectType(this, 'report')">
                                <div class="text-2xl mb-2">📊</div>
                                <div class="font-bold text-gray-700">Report</div>
                            </div>
                        </div>
                    </div>

                    <!-- Template -->
                    <div class="mb-6">
                        <label class="block text-sm font-semibold text-gray-700 mb-2">Visual Template</label>
                        <select id="templateSelect" class="w-full p-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition">
                            <option>Loading templates...</option>
                        </select>
                    </div>

                    <!-- Topic -->
                    <div class="mb-6">
                        <label class="block text-sm font-semibold text-gray-700 mb-2">Topic / Query</label>
                        <textarea id="queryInput" class="w-full p-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none min-h-[100px]" placeholder="e.g. Analysis of AI trends in Healthcare for Q3 2024..."></textarea>
                    </div>

                    <!-- Mode -->
                    <div class="flex items-center gap-4 mb-8">
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="radio" name="mode" value="normal" checked class="w-4 h-4 text-blue-600">
                            <span class="text-gray-700">Normal Search (Fast)</span>
                        </label>
                        <label class="flex items-center gap-2 cursor-pointer">
                            <input type="radio" name="mode" value="deep" class="w-4 h-4 text-blue-600">
                            <span class="text-gray-700">Deep Research (Comprehensive)</span>
                        </label>
                    </div>

                    <button onclick="generatePlan()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-lg transition flex justify-center items-center gap-2">
                        <span id="btn-text-plan">Generate Plan</span>
                        <div id="spinner-plan" class="spinner" style="border-top-color: white; display: none;"></div>
                    </button>
                </div>
            </div>
        </div>

        <!-- Tab: Plan -->
        <div id="tab-plan" class="step-container">
            <div class="max-w-5xl mx-auto">
                <div class="flex justify-between items-center mb-6">
                    <h2 class="text-2xl font-bold text-gray-800">Review Plan</h2>
                    <div class="flex gap-3">
                        <button onclick="switchTab('create')" class="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded-lg">Back</button>
                        <button onclick="approvePlan()" class="px-6 py-2 bg-green-600 hover:bg-green-700 text-white font-bold rounded-lg flex items-center gap-2">
                            <span>Approve & Build</span>
                            <div id="spinner-build" class="spinner" style="border-top-color: white; border-width: 2px; width: 16px; height: 16px; display: none;"></div>
                        </button>
                    </div>
                </div>

                <div id="planContent" class="space-y-4">
                    <!-- Plan items injected here -->
                </div>
            </div>
        </div>

        <!-- Tab: Editor -->
        <div id="tab-editor" class="step-container">
            <div class="max-w-6xl mx-auto h-[calc(100vh-80px)] flex gap-6">
                <!-- Left: Slides List -->
                <div class="w-2/3 flex flex-col">
                    <div class="flex justify-between items-center mb-4">
                        <h2 class="text-2xl font-bold text-gray-800">Slide Editor</h2>
                        <button onclick="saveContent()" class="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm hover:bg-blue-700">
                            <i class="fa-solid fa-floppy-disk mr-2"></i> Save Changes
                        </button>
                    </div>

                    <div id="editor-slides-container" class="flex-1 overflow-y-auto pr-2 space-y-8 pb-20">
                        <!-- Slide Editors injected here -->
                    </div>
                </div>

                <!-- Right: Chat / Copilot -->
                <div class="w-1/3 flex flex-col">
                    <h3 class="text-lg font-bold text-gray-700 mb-4">Copilot</h3>
                    <div class="chat-container flex-1">
                        <div id="chat-messages" class="chat-messages flex flex-col gap-3">
                            <div class="message ai">
                                Hello! I can help you restructure slides. Try asking "Change slide 2 to a chart" or "Make slide 3 a comparison".
                            </div>
                        </div>
                        <div class="chat-input-area">
                            <div class="flex gap-2">
                                <input type="text" id="chat-input" class="flex-1 border border-gray-300 rounded-lg px-3 py-2 text-sm focus:outline-none focus:border-blue-500" placeholder="Type instructions...">
                                <button onclick="sendChat()" class="bg-blue-600 text-white px-3 py-2 rounded-lg hover:bg-blue-700">
                                    <i class="fa-solid fa-paper-plane"></i>
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tab: Download -->
        <div id="tab-download" class="step-container">
            <div class="max-w-2xl mx-auto text-center mt-20">
                <div class="w-20 h-20 bg-green-100 text-green-600 rounded-full flex items-center justify-center mx-auto mb-6 text-3xl">
                    <i class="fa-solid fa-check"></i>
                </div>
                <h2 class="text-3xl font-bold text-gray-800 mb-4">Presentation Ready!</h2>
                <p class="text-gray-600 mb-8">Your slides have been generated successfully. You can download the final PowerPoint file below.</p>

                <div class="grid grid-cols-2 gap-4">
                    <button onclick="downloadFile('ppt')" class="p-6 border-2 border-blue-100 hover:border-blue-500 bg-white rounded-xl transition group">
                        <div class="text-4xl text-orange-500 mb-3 group-hover:scale-110 transition">
                            <i class="fa-solid fa-file-powerpoint"></i>
                        </div>
                        <div class="font-bold text-gray-800">Download .PPTX</div>
                        <div class="text-xs text-gray-400 mt-1">Editable PowerPoint</div>
                    </button>

                    <button onclick="downloadFile('json')" class="p-6 border-2 border-gray-100 hover:border-gray-400 bg-white rounded-xl transition group">
                        <div class="text-4xl text-green-600 mb-3 group-hover:scale-110 transition">
                            <i class="fa-solid fa-file-code"></i>
                        </div>
                        <div class="font-bold text-gray-800">Download JSON</div>
                        <div class="text-xs text-gray-400 mt-1">Data Source</div>
                    </button>
                </div>

                <button onclick="switchTab('create')" class="mt-12 text-blue-600 hover:underline">Start New Project</button>
            </div>
        </div>

    </div>

    <!-- Toast Notification -->
    <div id="toast" class="fixed bottom-5 right-5 bg-gray-800 text-white px-6 py-3 rounded-lg shadow-lg transform translate-y-20 opacity-0 transition-all duration-300 z-50">
        Notification
    </div>

    <script>
        // State
        let currentPlan = null;
        let reportId = null;
        let selectedType = 'sales';
        let executionLog = [];
        let selectedSlideIndex = 0;

        // Initialization
        window.addEventListener('DOMContentLoaded', () => {
            loadTemplates();
        });

        // UI Helpers
        function switchTab(tabId) {
            document.querySelectorAll('.step-container').forEach(el => el.classList.remove('active'));
            document.getElementById(`tab-${tabId}`).classList.add('active');

            document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
            const navItem = document.getElementById(`nav-${tabId}`);
            if(navItem) navItem.classList.add('active');
            else if(tabId === 'create') document.querySelector('.nav-item').classList.add('active');
        }

        function showToast(msg, type='info') {
            const toast = document.getElementById('toast');
            toast.textContent = msg;
            toast.className = `fixed bottom-5 right-5 px-6 py-3 rounded-lg shadow-lg transform translate-y-0 opacity-100 transition-all duration-300 z-50 ${type === 'error' ? 'bg-red-600' : 'bg-gray-800'} text-white`;
            setTimeout(() => {
                toast.classList.add('translate-y-20', 'opacity-0');
            }, 3000);
        }

        function selectType(el, type) {
            document.querySelectorAll('.mode-card').forEach(c => {
                c.classList.remove('border-blue-500', 'bg-blue-50');
                c.classList.add('border-transparent', 'bg-gray-50');
                c.querySelector('.font-bold').classList.remove('text-blue-700');
                c.querySelector('.font-bold').classList.add('text-gray-700');
            });
            el.classList.remove('border-transparent', 'bg-gray-50');
            el.classList.add('border-blue-500', 'bg-blue-50');
            el.querySelector('.font-bold').classList.remove('text-gray-700');
            el.querySelector('.font-bold').classList.add('text-blue-700');
            selectedType = type;
        }

        async function loadTemplates() {
            try {
                const res = await fetch('/api/templates');
                const templates = await res.json();
                const select = document.getElementById('templateSelect');
                select.innerHTML = '';
                Object.keys(templates).forEach(key => {
                    const opt = document.createElement('option');
                    opt.value = key;
                    opt.textContent = key;
                    select.appendChild(opt);
                });
            } catch(e) {
                console.error(e);
            }
        }

        // --- PHASE 1: Plan ---

        async function generatePlan() {
            const query = document.getElementById('queryInput').value;
            if(!query) return showToast('Please enter a topic', 'error');

            document.getElementById('btn-text-plan').textContent = 'Analyzing...';
            document.getElementById('spinner-plan').style.display = 'block';

            try {
                const res = await fetch('/api/plan', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        query,
                        template: document.getElementById('templateSelect').value,
                        search_mode: document.querySelector('input[name="mode"]:checked').value,
                        report_type: selectedType
                    })
                });
                
                if(!res.ok) throw new Error(await res.text());
                
                currentPlan = await res.json();
                renderPlan(currentPlan);
                
                // Enable Plan Tab
                document.getElementById('nav-plan').style.opacity = '1';
                document.getElementById('nav-plan').style.pointerEvents = 'auto';
                switchTab('plan');

            } catch(e) {
                showToast('Error generating plan: ' + e.message, 'error');
            } finally {
                document.getElementById('btn-text-plan').textContent = 'Generate Plan';
                document.getElementById('spinner-plan').style.display = 'none';
            }
        }

        function renderPlan(plan) {
            const container = document.getElementById('planContent');
            container.innerHTML = plan.sections.map((sec, idx) => `
                <div class="card p-5 border-l-4 border-l-blue-500">
                    <div class="flex justify-between items-start">
                        <div>
                            <h3 class="font-bold text-lg text-gray-800">Slide ${idx+1}: ${sec.section_title}</h3>
                            <p class="text-sm text-gray-600 mt-1">${sec.section_purpose}</p>
                            <div class="mt-3 flex gap-2">
                                <span class="bg-blue-100 text-blue-700 text-xs px-2 py-1 rounded">Layout: ${sec.layout_type}</span>
                                <span class="bg-gray-100 text-gray-600 text-xs px-2 py-1 rounded">${sec.search_queries?.length || 0} Queries</span>
                            </div>
                        </div>
                    </div>
                </div>
            `).join('');
        }

        // --- PHASE 2: Execute ---

        async function approvePlan() {
            if(!currentPlan) return;

            document.getElementById('spinner-build').style.display = 'block';

            try {
                const res = await fetch('/api/execute', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ plan_id: currentPlan.plan_id })
                });

                if(!res.ok) throw new Error(await res.text());
                const data = await res.json();
                reportId = data.report_id;
                
                showToast('Presentation Generated!');
                
                // Enable Editor & Download
                document.getElementById('nav-editor').style.opacity = '1';
                document.getElementById('nav-editor').style.pointerEvents = 'auto';
                document.getElementById('nav-download').style.opacity = '1';
                document.getElementById('nav-download').style.pointerEvents = 'auto';
                
                loadEditorContent();
                switchTab('editor');

            } catch(e) {
                showToast('Execution error: ' + e.message, 'error');
            } finally {
                document.getElementById('spinner-build').style.display = 'none';
            }
        }

        // --- PHASE 3: Editor ---

        async function loadEditorContent() {
            try {
                const res = await fetch(`/api/report/${reportId}/content`);
                const data = await res.json();
                executionLog = data.execution_log;
                renderEditor();
            } catch(e) {
                console.error(e);
            }
        }

        function renderEditor() {
            const container = document.getElementById('editor-slides-container');
            container.innerHTML = executionLog.map((slide, idx) => {
                if(slide.status === 'failed') return '';
                
                // Find content placeholders based on active role
                let contentHtml = '';
                const mainPh = slide.placeholders.find(p => ['content', 'bullets', 'chart', 'table', 'kpi'].includes(p.role));
                
                if (mainPh) {
                    if (mainPh.role === 'chart') {
                        contentHtml = `<div class="p-4 bg-blue-50 text-blue-700 rounded text-center"><i class="fa-solid fa-chart-simple text-3xl mb-2"></i><br>Chart: ${mainPh.chart_data?.title || 'Data Visualization'}</div>`;
                    } else if (mainPh.role === 'table') {
                        contentHtml = `<div class="p-4 bg-green-50 text-green-700 rounded text-center"><i class="fa-solid fa-table text-3xl mb-2"></i><br>Table Data</div>`;
                    } else if (mainPh.role === 'kpi') {
                        contentHtml = `<div class="text-center"><div class="text-4xl font-bold text-blue-600">${mainPh.kpi_data?.value || '0'}</div><div class="text-gray-500">${mainPh.kpi_data?.label || 'Metric'}</div></div>`;
                    } else {
                        const bullets = mainPh.bullets || [];
                        contentHtml = `<ul class="list-disc pl-5 space-y-1 text-gray-600" data-field="bullets" data-slide="${idx}">` +
                                      bullets.map(b => `<li contenteditable="true" class="mb-1 p-1 hover:bg-gray-50 border border-transparent hover:border-gray-200 rounded">${b}</li>`).join('') +
                                      `</ul>`;
                    }
                }

                return `
                <div class="slide-preview cursor-pointer ${idx === selectedSlideIndex ? 'selected-slide' : ''}"
                     id="slide-editor-${idx}" onclick="selectSlide(${idx})">
                    <div class="slide-preview-header flex justify-between items-center">
                        <div class="font-bold text-gray-700">Slide ${slide.slide}: ${slide.title}</div>
                        <div class="text-xs text-gray-400">Layout: ${slide.layout_type}</div>
                    </div>
                    <div class="slide-preview-body">
                        <h1 contenteditable="true" class="text-2xl font-bold text-gray-800 mb-4 p-1 hover:bg-gray-50 border border-transparent hover:border-gray-200 rounded" data-field="title" data-slide="${idx}">${slide.title}</h1>
                        ${contentHtml}
                    </div>
                    <div class="slide-preview-footer">
                        <span class="text-xs text-gray-400">ID: ${reportId}</span>
                        ${idx === selectedSlideIndex ? '<span class="text-xs font-bold text-blue-600">Selected for Chat</span>' : ''}
                    </div>
                </div>
                `;
            }).join('');
        }

        function selectSlide(idx) {
            selectedSlideIndex = idx;
            renderEditor();
            document.getElementById('slide-editor-' + idx).scrollIntoView({ behavior: 'smooth', block: 'center' });
        }

        async function saveContent() {
            // Scrape content from DOM
            const updatedLog = JSON.parse(JSON.stringify(executionLog)); // Deep copy
            
            updatedLog.forEach((slide, idx) => {
                const dom = document.getElementById(`slide-editor-${idx}`);
                if(!dom) return;
                
                // Update Title
                const titleEl = dom.querySelector('[data-field="title"]');
                if(titleEl) slide.title = titleEl.innerText;
                
                // Update Bullets
                const bulletsContainer = dom.querySelector('[data-field="bullets"]');
                if(bulletsContainer) {
                    const bullets = Array.from(bulletsContainer.querySelectorAll('li')).map(li => li.innerText);

                    // Update in placeholders structure
                    const ph = slide.placeholders.find(p => p.role === 'content' || p.role === 'bullets');
                    if(ph) ph.bullets = bullets;
                }
            });
            
            try {
                const res = await fetch(`/api/report/${reportId}/update`, {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ execution_log: updatedLog })
                });
                if(res.ok) showToast('Content Saved & PPT Updated', 'success');
                else throw new Error('Save failed');
            } catch(e) {
                showToast(e.message, 'error');
            }
        }

        // --- PHASE 4: Chat ---

        async function sendChat() {
            const input = document.getElementById('chat-input');
            const msg = input.value.trim();
            if(!msg) return;

            // Add User Message
            const chatBox = document.getElementById('chat-messages');
            chatBox.innerHTML += `<div class="message user">${msg}</div>`;
            input.value = '';
            chatBox.scrollTop = chatBox.scrollHeight;

            try {
                const res = await fetch(`/api/report/${reportId}/chat`, {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({
                        message: msg,
                        slide_index: selectedSlideIndex
                    })
                });
                const data = await res.json();
                
                // Add AI Message
                chatBox.innerHTML += `<div class="message ai">${data.message}</div>`;
                chatBox.scrollTop = chatBox.scrollHeight;
                
                // Update specific slide content in the log and re-render
                if (data.updated_slide) {
                    executionLog[selectedSlideIndex] = data.updated_slide;
                    renderEditor();
                }
                
            } catch(e) {
                chatBox.innerHTML += `<div class="message ai text-red-500">Error: ${e.message}</div>`;
            }
        }

        // --- PHASE 5: Download ---

        function downloadFile(format) {
            window.location.href = `/api/download/${reportId}?format=${format}`;
        }

    </script>
</body>
</html>
"""
