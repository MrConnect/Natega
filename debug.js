const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const https = require('https');

const httpsAgent = new https.Agent({ rejectUnauthorized: false });

(async () => {
    try {
        const r1 = await axios.get('https://studentportalsis.azu.edu.eg/Default.aspx', {
            httpsAgent, maxRedirects: 0, validateStatus: s => s < 400
        });

        let cook = (r1.headers['set-cookie'] || []).find(c => c.includes('ASP.NET_SessionId='));
        if (cook) cook = cook.split(';')[0];

        const r2 = await axios.get('https://studentportalsis.azu.edu.eg/Default.aspx', {
            httpsAgent, maxRedirects: 0, validateStatus: s => s < 400, headers: { Cookie: cook }
        });

        let $ = cheerio.load(r2.data);
        let vs = $('#__VIEWSTATE').val();
        let vsg = $('#__VIEWSTATEGENERATOR').val();
        let ev = $('#__EVENTVALIDATION').val();

        const form = new URLSearchParams();
        form.append('__EVENTTARGET', '');
        form.append('__EVENTARGUMENT', '');
        form.append('__VIEWSTATE', vs);
        form.append('__VIEWSTATEGENERATOR', vsg);
        form.append('__EVENTVALIDATION', ev);
        form.append('txtUsername', '120240002236');
        form.append('txtPassword', '30610011314071');
        form.append('btnEnter', 'دخول');

        const r3 = await axios.post('https://studentportalsis.azu.edu.eg/Default.aspx', form, {
            httpsAgent, maxRedirects: 0, validateStatus: s => s < 400, headers: { Cookie: cook }
        });

        let authCook = (r3.headers['set-cookie'] || []).find(c => c.includes('.ASPXAUTH'));
        if (authCook) authCook = authCook.split(';')[0];
        let finalCook = cook + '; ' + authCook;

        console.log('Login success, fetching results page...');

        const r4 = await axios.get('https://studentportalsis.azu.edu.eg/UI/StudentView/student_sem_work_Modular.aspx', {
            httpsAgent, headers: { Cookie: finalCook }
        });

        fs.writeFileSync('test_results_page.html', r4.data);
        console.log('Wrote initial results page to test_results_page.html');

        $ = cheerio.load(r4.data);
        vs = $('#__VIEWSTATE').val();
        vsg = $('#__VIEWSTATEGENERATOR').val();
        ev = $('#__EVENTVALIDATION').val();

        const resultsForm = new URLSearchParams();
        // Mimic search for 2024/2025 (15), Sem 1 (1), Round 1 (1)
        resultsForm.append('ctl00$ScriptManager1', 'ctl00$cntphmaster$UpdatePanel1|ctl00$cntphmaster$searchButton');
        resultsForm.append('__EVENTTARGET', '');
        resultsForm.append('__EVENTARGUMENT', '');
        resultsForm.append('__VIEWSTATE', vs);
        resultsForm.append('__VIEWSTATEGENERATOR', vsg);
        resultsForm.append('__EVENTVALIDATION', ev);
        resultsForm.append('ctl00$cntphmaster$ACadYearDropDownList', '15'); // 2024/2025
        resultsForm.append('ctl00$cntphmaster$semesterDropDownList', '1');  // Sem 1
        resultsForm.append('ctl00$cntphmaster$drpExamType', '1');           // Round 1
        resultsForm.append('__ASYNCPOST', 'true');
        resultsForm.append('ctl00$cntphmaster$searchButton', 'بحث');

        const r5 = await axios.post('https://studentportalsis.azu.edu.eg/UI/StudentView/student_sem_work_Modular.aspx', resultsForm, {
            httpsAgent,
            headers: {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Cookie': finalCook,
                'X-Requested-With': 'XMLHttpRequest',
                'X-MicrosoftAjax': 'Delta=true'
            }
        });

        fs.writeFileSync('test_results_response.html', r5.data);
        console.log('Wrote search results response to test_results_response.html');

    } catch (e) {
        console.error(e);
    }
})();
