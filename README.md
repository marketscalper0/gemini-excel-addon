# Gemini Excel Add-in - AI-Powered Spreadsheet Agent

✨ **An intelligent Excel Add-in integrated with Google Gemini API for advanced data analysis, automation, and AI-powered insights directly in Excel**

## Features

### Core Capabilities
- 🤖 **Direct Gemini Integration**: Use any Gemini model (Pro, Flash, 2.0) for powerful AI operations
- 📊 **Data Analysis**: Automatically analyze spreadsheet data with AI insights
- 🧮 **Formula Generation**: Generate Excel formulas from natural language descriptions
- 🔍 **Pattern Recognition**: Find anomalies, trends, and patterns in your data
- ✅ **Data Validation**: Validate data quality and consistency across spreadsheets
- 💡 **Intelligent Suggestions**: Get AI-powered recommendations for data processing
- 🎯 **Tax Compliance Insights**: Specialized analysis for GST, ITR, and financial compliance

### Available Functions

Call these functions directly in Excel cells:

```excel
=GEMINI_ANALYZE(A1:A100)          # Get AI insights from data range
=GEMINI_CALC("sum formula")        # Generate formulas from text
=GEMINI_EXPLAIN(A1)               # Explain cell content
=GEMINI_VALIDATE(A1:A50)          # Validate data quality
=GEMINI_FIND(A1:A100, "anomalies") # Find patterns
=GEMINI_SUGGEST("context")        # Get AI suggestions
=GEMINI_SUMMARY(A1:A100)          # Create executive summary
=GEMINI_EXTRACT(A1, "emails")     # Extract entities
=GEMINI_CLASSIFY(A1, "sentiment")  # Classify data
=GEMINI_TRANSFORM(A1, "JSON to CSV") # Transform formats
=GEMINI_MATH("compound interest")  # Mathematical calculations
=GEMINI_TAX(A1:C100, "GST analysis") # Tax compliance analysis
```

## Installation

### Option 1: Local Development Setup

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/gemini-excel-addon.git
   cd gemini-excel-addon
   ```

2. **Get Gemini API Key**
   - Go to [Google AI Studio](https://aistudio.google.com/app/apikey)
   - Create a new API key
   - Keep it safe (you'll need this)

3. **Host the files**
   - Use HTTPS hosting (required for Excel Add-ins)
   - Upload `taskpane.html`, `taskpane.js`, and `functions.js` to your server
   - Update the `SourceLocation` in `manifest.xml` with your hosting URL

4. **Add the Add-in to Excel**
   - Open Excel
   - Go to **Insert** > **Get Add-ins** > **My Add-ins** > **Upload My Add-in**
   - Upload your `manifest.xml` file
   - The add-in will appear in your Excel toolbar

### Option 2: Using Microsoft AppSource (Production)

1. Create a Microsoft Partner Center account
2. Submit your add-in for certification
3. Once approved, it will be available on AppSource

## Quick Start

### 1. Configure API Key
1. Open the Excel Add-in (click the add-in icon in Excel ribbon)
2. Go to **Settings** tab
3. Paste your Gemini API Key
4. Select your preferred model (Gemini 1.5 Flash recommended)
5. Adjust Temperature (0-2) for creativity level
6. Click **Save Settings**

### 2. Use the Query Tab
1. Go to **Query** tab
2. Enter your question or request
3. Optionally select cells and check "Include selected cells in context"
4. Click **Ask Gemini**
5. View results and copy/insert as needed

### 3. Use Inline Functions
1. Simply type in any Excel cell: `=GEMINI_ANALYZE(A1:A100)`
2. Press Enter to execute
3. Function calls Gemini API and returns result

## Configuration

### API Key Security
- Keys are stored locally in browser localStorage
- Never transmitted to third-party servers
- Clear from Settings tab anytime

### Model Selection
- **Gemini 1.5 Pro**: Best for complex analysis (higher cost)
- **Gemini 1.5 Flash**: Balanced speed/cost (recommended)
- **Gemini 2.0 Flash**: Latest model with advanced capabilities

### Temperature Settings
- **0.0-0.5**: Deterministic, focused answers
- **0.5-1.0**: Balanced (recommended for most tasks)
- **1.0-2.0**: Creative, varied responses

## Use Cases

### For Chartered Accountants
- **ITR Analysis**: Analyze income tax returns, identify deductions, validate compliance
- **GST Compliance**: Analyze GST data, calculate tax liability, identify anomalies
- **Financial Analysis**: Generate insights from financial statements
- **Audit Reports**: Create summaries of audit findings

### For Data Analysts
- Data cleaning and validation
- Pattern detection and anomaly detection
- Report generation and summarization
- Data transformation and format conversion

### For Business Users
- Quick data insights without complex formulas
- Natural language data queries
- Automated analysis and reporting
- Compliance checking

## File Structure

```
gemini-excel-addon/
├── manifest.xml          # Excel Add-in configuration
├── taskpane.html         # Main UI interface
├── taskpane.js           # Core functionality & Gemini API integration
├── functions.js          # Custom Excel functions
└── README.md             # This file
```

## API Reference

### callGeminiAPI(prompt)
Internal function to call Gemini API

**Parameters:**
- `prompt` (string): The prompt to send to Gemini

**Returns:**
- Promise<string>: AI-generated response

### All Custom Functions
These are async functions that return AI-generated content:

- `GEMINI_ANALYZE(range)` - Data analysis
- `GEMINI_CALC(description)` - Formula generation
- `GEMINI_EXPLAIN(content)` - Content explanation
- `GEMINI_VALIDATE(range)` - Data validation
- `GEMINI_FIND(range, criteria)` - Pattern finding
- `GEMINI_SUGGEST(context)` - Suggestions
- `GEMINI_SUMMARY(range)` - Summarization
- `GEMINI_EXTRACT(text, type)` - Entity extraction
- `GEMINI_CLASSIFY(data, type)` - Data classification
- `GEMINI_TRANSFORM(data, type)` - Data transformation
- `GEMINI_MATH(calculation)` - Math calculations
- `GEMINI_TAX(range, type)` - Tax analysis

## Troubleshooting

### "API Key Invalid"
- Verify key is correctly copied from Google AI Studio
- Ensure no extra spaces
- Keys expire - regenerate if needed

### "CORS Error"
- Ensure manifest.xml has correct SourceLocation
- Verify hosting uses HTTPS
- Check firewall/security settings

### "Function Not Found"
- Ensure functions.js is loaded
- Functions are asynchronous - may take a moment
- Check browser console for errors

### "No Response"
- Check internet connection
- Verify API quota not exceeded
- Check Gemini API status

## Performance Tips

1. **For Large Datasets**: Use `GEMINI_SUMMARY` instead of `GEMINI_ANALYZE`
2. **Batch Operations**: Process data in chunks rather than entire spreadsheet
3. **Model Selection**: Use Flash for speed, Pro for complexity
4. **Temperature**: Keep low (0.5-0.7) for consistent results

## Limitations

- API responses limited to ~2048 tokens
- Real-time collaboration may have delays
- Large datasets (>1000 rows) may need pagination
- API rate limits apply based on Gemini tier

## Security & Privacy

- ✅ All data stays on your machine unless sent to Gemini API
- ✅ API key stored locally only
- ✅ No tracking or data collection
- ✅ HTTPS-only transmission
- ✅ Supports air-gapped environments with API proxy

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request

## License

MIT License - Free for personal and commercial use

## Support

For issues, questions, or suggestions:
- Open an issue on GitHub
- Check existing documentation
- Contact: [your-email@example.com]

## Roadmap

- [ ] Excel Online support
- [ ] Advanced caching for API efficiency
- [ ] Batch processing UI
- [ ] Custom function builder
- [ ] Multi-language support
- [ ] Integration with Microsoft Teams
- [ ] Advanced data visualization
- [ ] Scheduled analysis jobs

---

**Made with ❤️ for data professionals, accountants, and analysts**

*Powered by Google Gemini API*
