// Gemini Excel Add-in - Task Pane JavaScript
// Complete AI Agent for Excel using Google Gemini API

let geminiConfig = {
    apiKey: localStorage.getItem('geminiApiKey') || '',
    model: localStorage.getItem('geminiModel') || 'gemini-1.5-flash',
    temperature: parseFloat(localStorage.getItem('geminiTemp') || '0.7'),
    apiEndpoint: 'https://generativelanguage.googleapis.com/v1beta/models'
};

// Initialize Office JavaScript API
Office.onReady((reason) => {
    if (reason === Office.MailboxEnums.InitializationReason.Completed) {
        console.log('Excel Add-in ready!');
        loadStoredSettings();
    }
});

// Tab switching functionality
function switchTab(tabName) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    
    document.getElementById(tabName).classList.add('active');
    event.target.classList.add('active');
}

// Load stored settings from localStorage
function loadStoredSettings() {
    const savedKey = localStorage.getItem('geminiApiKey');
    const savedModel = localStorage.getItem('geminiModel');
    const savedTemp = localStorage.getItem('geminiTemp');
    
    if (document.getElementById('geminiKey')) {
        document.getElementById('geminiKey').value = savedKey || '';
        document.getElementById('model').value = savedModel || 'gemini-1.5-flash';
        document.getElementById('temperature').value = savedTemp || '0.7';
    }
}

// Save settings
function saveSettings() {
    const apiKey = document.getElementById('geminiKey').value;
    const model = document.getElementById('model').value;
    const temperature = document.getElementById('temperature').value;
    
    if (!apiKey.trim()) {
        alert('Please enter a Gemini API Key');
        return;
    }
    
    localStorage.setItem('geminiApiKey', apiKey);
    localStorage.setItem('geminiModel', model);
    localStorage.setItem('geminiTemp', temperature);
    
    geminiConfig.apiKey = apiKey;
    geminiConfig.model = model;
    geminiConfig.temperature = parseFloat(temperature);
    
    alert('Settings saved successfully!');
}

// Get selected cells context
async function getSelectedCellsContext() {
    try {
        const range = await Excel.run(async (context) => {
            let range = context.application.getSelectedData(Excel.SelectionMode.normal);
            range.load('values,address');
            await context.sync();
            return { address: range.address, values: range.values };
        });
        return range;
    } catch (error) {
        console.log('No cells selected or error reading selection');
        return null;
    }
}

// Main query function
async function sendQuery() {
    const query = document.getElementById('userQuery').value.trim();
    const useSelection = document.getElementById('useSelection').checked;
    const loading = document.getElementById('loading');
    const responseBox = document.getElementById('responseBox');
    
    if (!query) {
        alert('Please enter a query');
        return;
    }
    
    if (!geminiConfig.apiKey) {
        alert('Please configure Gemini API key in Settings tab first');
        return;
    }
    
    loading.style.display = 'block';
    responseBox.style.display = 'none';
    
    try {
        let fullPrompt = query;
        
        if (useSelection) {
            const context = await getSelectedCellsContext();
            if (context) {
                fullPrompt = `Excel Data Context:\n${JSON.stringify(context, null, 2)}\n\nUser Query:\n${query}`;
            }
        }
        
        const response = await callGeminiAPI(fullPrompt);
        
        loading.style.display = 'none';
        responseBox.textContent = response;
        responseBox.style.display = 'block';
        
    } catch (error) {
        loading.style.display = 'none';
        responseBox.textContent = `Error: ${error.message}`;
        responseBox.style.display = 'block';
    }
}

// Call Gemini API
async function callGeminiAPI(prompt) {
    const url = `${geminiConfig.apiEndpoint}/${geminiConfig.model}:generateContent?key=${geminiConfig.apiKey}`;
    
    const payload = {
        contents: [
            {
                parts: [
                    {
                        text: prompt
                    }
                ]
            }
        ],
        generationConfig: {
            temperature: geminiConfig.temperature,
            maxOutputTokens: 2048
        }
    };
    
    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(payload)
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error?.message || 'API call failed');
        }
        
        const result = await response.json();
        return result.candidates[0].content.parts[0].text;
        
    } catch (error) {
        throw error;
    }
}

// Clear query
function clearQuery() {
    document.getElementById('userQuery').value = '';
    document.getElementById('responseBox').style.display = 'none';
}

// Excel Custom Functions - Can be called directly in Excel cells
// Note: These require Office Scripts or add-in API enablement

async function insertResponseToCell() {
    const response = document.getElementById('responseBox').textContent;
    if (!response) {
        alert('No response to insert');
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.application.getActiveWorksheet();
            const range = sheet.getSelectedRange();
            range.values = [[response]];
            await context.sync();
        });
        alert('Response inserted to selected cell!');
    } catch (error) {
        alert('Error inserting response: ' + error.message);
    }
}

// Batch processing
async function batchAnalyze() {
    try {
        await Excel.run(async (context) => {
            const worksheet = context.application.getActiveWorksheet();
            const range = worksheet.getSelectedRange();
            range.load('values,address');
            await context.sync();
            
            let results = [];
            for (let i = 0; i < range.values.length; i++) {
                const rowData = range.values[i];
                const prompt = `Analyze this data row: ${JSON.stringify(rowData)}`;
                const analysis = await callGeminiAPI(prompt);
                results.push([analysis]);
            }
            
            const nextColumn = worksheet.getRange(range.address).getOffsetRange(0, range.columnCount);
            nextColumn.values = results;
            await context.sync();
        });
    } catch (error) {
        alert('Batch analysis error: ' + error.message);
    }
}

// Advanced features placeholder
function showAdvancedOptions() {
    // Can be extended with more features
    alert('Advanced features coming soon!');
}

// Export results to new sheet
async function exportResults() {
    const response = document.getElementById('responseBox').textContent;
    if (!response) {
        alert('No results to export');
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const workbook = context.application.workbook;
            const sheet = workbook.worksheets.add('Gemini Results');
            sheet.getRange('A1').values = [['Gemini Analysis Results']];
            sheet.getRange('A2').values = [[response]];
            sheet.getRange('A1').format.bold = true;
            sheet.getRange('A1:A1').format.fill.color = '#667eea';
            await context.sync();
            workbook.worksheets.getActiveWorksheet().activate();
        });
        alert('Results exported to new sheet!');
    } catch (error) {
        alert('Export error: ' + error.message);
    }
}
