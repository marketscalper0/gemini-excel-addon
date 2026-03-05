# Deployment Guide - Gemini Excel Add-in

## Complete Step-by-Step Deployment Instructions

### Prerequisites
- Excel 2016 or later (Desktop)
- Excel Online (Web)
- Windows 10+, macOS 10.10+, or Mac OS X
- Node.js (v14+) for local development
- HTTPS hosting (required)

### Step 1: Get Your Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Click "Create API Key"
3. Select your Google Cloud project or create new one
4. Copy the API key - keep it safe and secure
5. Never share this key publicly

### Step 2: Host the Files

#### Option A: Using Netlify (Easiest)

1. Create account at [netlify.com](https://netlify.com)
2. Install Netlify CLI:
   ```bash
   npm install -g netlify-cli
   ```
3. Deploy:
   ```bash
   netlify deploy --prod --dir=.
   ```
4. Your hosting URL will be like: `https://your-site.netlify.app`

#### Option B: Using GitHub Pages

1. Fork this repository
2. Enable GitHub Pages in Settings
3. Your hosting URL will be: `https://username.github.io/gemini-excel-addon`

#### Option C: Using Your Own Server

1. Upload all files to your HTTPS server
2. Configure CORS headers:
   ```
   Access-Control-Allow-Origin: *
   Access-Control-Allow-Methods: GET, POST, OPTIONS
   ```

### Step 3: Update manifest.xml

1. Edit `manifest.xml`
2. Replace the `SourceLocation` URL:
   ```xml
   <SourceLocation DefaultValue="https://your-hosting-url.com/taskpane.html"/>
   ```
3. Ensure the URL uses HTTPS (not HTTP)

### Step 4: Add to Excel

#### Desktop Excel (Windows/Mac)

1. Open Excel
2. Go to **Insert** → **Get Add-ins** → **My Add-ins**
3. Click **Upload My Add-in**
4. Select your updated `manifest.xml` file
5. The add-in will appear in your ribbon

#### Excel Online

1. Open Excel Online (office.com)
2. Go to **Insert** → **Get Add-ins** → **My Add-ins**
3. Upload `manifest.xml`
4. Add-in will be available in the ribbon

### Step 5: Configure and Use

1. Click the Gemini AI Agent icon in Excel ribbon
2. Go to **Settings** tab
3. Paste your Gemini API key
4. Select model (Gemini 1.5 Flash recommended)
5. Click **Save Settings**
6. Start using in **Query** tab or as Excel functions

## Testing the Installation

### Test 1: Settings
1. Open the add-in
2. Go to Settings tab
3. Paste API key and save
4. Should show "Settings saved successfully!"

### Test 2: Simple Query
1. Go to Query tab
2. Type: "What is 2+2?"
3. Click "Ask Gemini"
4. Should return "4"

### Test 3: Excel Function
1. In any Excel cell, type: `=GEMINI_EXPLAIN("AI")`
2. Press Enter
3. Should get explanation of AI

## Troubleshooting

### "Unable to load add-in"
- Check manifest.xml syntax (XML must be valid)
- Verify SourceLocation uses HTTPS
- Check hosting URL is accessible
- Clear browser cache

### "CORS Error"
- Add CORS headers to your server
- Or use Netlify (handles CORS automatically)
- Check browser console for specific error

### "API Key not working"
- Verify key copied correctly (no extra spaces)
- Check key is active on Google AI Studio
- Try generating new key
- Ensure no rate limits exceeded

### "Functions not working"
- Refresh Excel (Ctrl+Shift+F9)
- Ensure functions.js is loaded (check console)
- Try reloading the add-in
- Check API usage quota

## Production Deployment

### For Enterprise Use

1. **Create Central Repository**
   ```
   https://your-company.com/excel-addins/gemini/
   ```

2. **Update manifest.xml with corporate URL**

3. **Submit for Admin Consent (Optional)**
   - Go to Microsoft Partner Center
   - Submit add-in for certification
   - Once approved, appears in company's App Catalog

4. **Distribute via Group Policy (Windows)**
   ```
   Computer Configuration → Administrative Templates → 
   Microsoft Office [Version] → User Configuration → 
   Application Settings → Microsoft Excel → Excel Addins
   ```

### Security Best Practices

1. **API Key Management**
   - Use environment variables for keys
   - Rotate keys regularly
   - Never commit keys to Git
   - Use Azure Key Vault for enterprise

2. **HTTPS Only**
   - Always use HTTPS
   - Use valid SSL certificates
   - Enable HSTS headers

3. **CORS Configuration**
   - Restrict to known domains
   - Use whitelist approach
   - Monitor for unauthorized access

4. **Rate Limiting**
   - Monitor API usage
   - Set up alerts for quota
   - Implement backoff strategy

## Monitoring and Maintenance

### Setup Monitoring

1. **Monitor API Quotas**
   - Check Google Cloud Console
   - Set up billing alerts
   - Monitor error rates

2. **User Analytics**
   - Track feature usage
   - Monitor performance metrics
   - Collect user feedback

### Regular Maintenance

1. **Update Dependencies**
   ```bash
   npm update
   ```

2. **Test After Updates**
   - Run all test cases
   - Verify Excel compatibility
   - Check API integration

3. **Version Management**
   - Update manifest version
   - Document changes
   - Maintain changelog

## Updates and Patching

### Applying Updates

1. Pull latest changes
   ```bash
   git pull origin main
   ```

2. Update hosted files
   ```bash
   netlify deploy --prod --dir=.
   ```

3. Users can reload:
   - Desktop: Close and reopen Excel
   - Online: Refresh browser tab

## Support and Troubleshooting

### Getting Help

1. **Check Documentation**
   - README.md
   - This file
   - Code comments

2. **Check Logs**
   - Browser Dev Tools (F12)
   - Console tab for errors
   - Network tab for API calls

3. **Report Issues**
   - Open GitHub issue
   - Include error message
   - Share manifest.xml (without sensitive data)

## Rollback Procedure

If issues occur:

1. Identify problem
2. Revert to previous version:
   ```bash
   git revert [commit-hash]
   git push origin main
   ```
3. Redeploy to hosting
4. Users reload add-in

## Performance Optimization

### Caching
- Browser caches API responses
- Consider implementing server-side caching
- Clear cache if issues arise

### Request Optimization
- Batch multiple requests when possible
- Use appropriate model sizes
- Monitor response times

---

**Questions?** Open an issue on GitHub or check the README.md
