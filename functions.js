// Custom Excel Functions powered by Gemini API
// These functions can be called directly in Excel cells

/**
 * Analyze data range with Gemini AI
 * Usage: =GEMINI_ANALYZE(A1:A100)
 */
async function GEMINI_ANALYZE(range) {
    try {
        const dataStr = Array.isArray(range) ? JSON.stringify(range) : String(range);
        const prompt = `Analyze this data and provide insights:\n${dataStr}`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Generate Excel formula from text description
 * Usage: =GEMINI_CALC("sum of column A divided by count")
 */
async function GEMINI_CALC(formulaDescription) {
    try {
        const prompt = `Generate an Excel formula for: ${formulaDescription}. Return only the formula.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Explain what a cell contains or calculate
 * Usage: =GEMINI_EXPLAIN(A1)
 */
async function GEMINI_EXPLAIN(cellContent) {
    try {
        const prompt = `Briefly explain what this means or does: ${cellContent}`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Validate data quality and consistency
 * Usage: =GEMINI_VALIDATE(A1:A50)
 */
async function GEMINI_VALIDATE(dataRange) {
    try {
        const dataStr = Array.isArray(dataRange) ? JSON.stringify(dataRange) : String(dataRange);
        const prompt = `Validate this data for quality issues:\n${dataStr}\nProvide a validation report.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Find patterns or specific items in data
 * Usage: =GEMINI_FIND(A1:A100, "anomalies")
 */
async function GEMINI_FIND(dataRange, searchCriteria) {
    try {
        const dataStr = Array.isArray(dataRange) ? JSON.stringify(dataRange) : String(dataRange);
        const prompt = `Find ${searchCriteria} in this data:\n${dataStr}\nReturn findings as list.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Get AI suggestions based on context
 * Usage: =GEMINI_SUGGEST("improve this data")
 */
async function GEMINI_SUGGEST(context) {
    try {
        const prompt = `Provide suggestions for: ${context}`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Generate summary from data range
 * Usage: =GEMINI_SUMMARY(A1:A100)
 */
async function GEMINI_SUMMARY(dataRange) {
    try {
        const dataStr = Array.isArray(dataRange) ? JSON.stringify(dataRange) : String(dataRange);
        const prompt = `Create a brief executive summary of this data:\n${dataStr}`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Extract entities from text
 * Usage: =GEMINI_EXTRACT(A1, "email addresses")
 */
async function GEMINI_EXTRACT(text, entityType) {
    try {
        const prompt = `Extract ${entityType} from this text: ${text}. Return as comma-separated list.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Classify data into categories
 * Usage: =GEMINI_CLASSIFY(A1, "sentiment")
 */
async function GEMINI_CLASSIFY(data, classification) {
    try {
        const prompt = `Classify this as ${classification}: ${data}. Return only the classification.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Transform data format
 * Usage: =GEMINI_TRANSFORM(A1, "JSON to CSV")
 */
async function GEMINI_TRANSFORM(data, transformation) {
    try {
        const prompt = `Transform the following ${transformation}: ${data}`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Mathematical and financial calculations
 * Usage: =GEMINI_MATH("compound interest 10000 5% 10 years")
 */
async function GEMINI_MATH(calculation) {
    try {
        const prompt = `Calculate: ${calculation}. Return only the numerical result.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}

/**
 * Generate insights for tax compliance
 * Usage: =GEMINI_TAX(A1:C100, "GST analysis")
 */
async function GEMINI_TAX(dataRange, analysisType) {
    try {
        const dataStr = Array.isArray(dataRange) ? JSON.stringify(dataRange) : String(dataRange);
        const prompt = `Perform ${analysisType} on this financial data:\n${dataStr}\nProvide tax compliance insights.`;
        return await callGeminiAPI(prompt);
    } catch (error) {
        return `#ERROR: ${error.message}`;
    }
}
