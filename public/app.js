document.addEventListener('DOMContentLoaded', () => {

    // Check path to run appropriate logic
    const path = window.location.pathname;
    const isResultsPage = path.includes('results.html');

    if (isResultsPage) {
        initResultsPage();
    } else {
        initLoginPage();
    }

});

function getBaseUrl() {
    // Determine the base URL for the API calls
    return window.location.origin.includes('localhost')
        ? 'http://localhost:3000/api'
        : '/api'; // Assuming same-origin production setup
}

/* ======== LOGIN LOGIC ======== */
function initLoginPage() {
    // Redirect if already logged in (has token)
    if (localStorage.getItem('student_session')) {
        window.location.href = 'results.html';
        return;
    }

    const loginForm = document.getElementById('loginForm');
    const usernameInput = document.getElementById('username');
    const passwordInput = document.getElementById('password');
    const loginBtn = document.getElementById('loginBtn');
    const spinner = loginBtn.querySelector('.spinner');
    const btnText = loginBtn.querySelector('span');
    const btnIcon = loginBtn.querySelector('i');
    const messageEl = document.getElementById('loginMessage');

    if (loginForm) {
        loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const username = usernameInput.value.trim();
            const password = passwordInput.value.trim();

            if (!username || !password) return;

            // UI Loading State
            loginBtn.disabled = true;
            spinner.classList.remove('hide');
            btnText.style.opacity = '0';
            btnIcon.style.opacity = '0';

            messageEl.classList.add('hide');
            messageEl.className = 'message'; // reset classes

            try {
                const response = await fetch(`${getBaseUrl()}/login`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ username, password })
                });

                const data = await response.json();

                if (response.ok && data.success) {
                    // Save session token loosely in localStorage
                    // NOTE: In production, consider HttpOnly cookies if possible, 
                    // but since the original uses .ASPXAUTH, saving it here works for the proxy.
                    localStorage.setItem('student_session', data.sessionToken);
                    localStorage.setItem('student_id', username);
                    localStorage.setItem('student_national_id', data.nationalId);

                    messageEl.textContent = 'تم تسجيل الدخول بنجاح! جاري التحويل...';
                    messageEl.classList.add('success');
                    messageEl.classList.remove('hide');

                    setTimeout(() => {
                        window.location.href = 'results.html';
                    }, 1000);

                } else {
                    throw new Error(data.message || 'بيانات الدخول غير صحيحة');
                }

            } catch (error) {
                messageEl.textContent = error.message === 'Failed to fetch'
                    ? 'تعذر الاتصال بالخادم. تأكد من تشغيل الـ Backend'
                    : error.message;
                messageEl.classList.add('error');
                messageEl.classList.remove('hide');
            } finally {
                // UI Reset
                loginBtn.disabled = false;
                spinner.classList.add('hide');
                btnText.style.opacity = '1';
                btnIcon.style.opacity = '1';
            }
        });
    }
}


/* ======== RESULTS LOGIC ======== */
function initResultsPage() {
    const sessionToken = localStorage.getItem('student_session');

    // Redirect to login if no token
    if (!sessionToken) {
        window.location.href = 'index.html';
        return;
    }

    // Set initial info
    const studentId = localStorage.getItem('student_id');
    if (studentId) {
        document.getElementById('navStudentName').textContent = studentId;
    }

    // Handlers
    const logoutBtn = document.getElementById('logoutBtn');
    const resultsForm = document.getElementById('resultsForm');
    const searchBtn = document.getElementById('searchBtn');
    const spinner = searchBtn.querySelector('.spinner');

    const yearSelect = document.getElementById('yearSelect');
    const semesterSelect = document.getElementById('semesterSelect');
    const roundSelect = document.getElementById('roundSelect');

    const resultsSection = document.getElementById('resultsSection');
    const resultsTableBody = document.getElementById('resultsTableBody');
    const messageEl = document.getElementById('resultsMessage');

    const detailsBanner = document.getElementById('studentDetailsBanner');
    const detailsName = document.getElementById('studentDetailsName');

    // Logout
    logoutBtn.addEventListener('click', () => {
        localStorage.removeItem('student_session');
        localStorage.removeItem('student_id');
        localStorage.removeItem('student_national_id');
        window.location.href = 'index.html';
    });

    resultsForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const year = yearSelect.value;
        const semester = semesterSelect.value;
        const round = roundSelect.value;

        // UI Loading State
        searchBtn.disabled = true;
        spinner.classList.remove('hide');
        resultsSection.classList.add('hide');
        messageEl.classList.add('hide');
        messageEl.className = 'message';

        try {
            const response = await fetch(`${getBaseUrl()}/results`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    sessionToken,
                    year,
                    semester,
                    round
                })
            });

            const data = await response.json();

            if (response.ok && data.success) {

                if (data.data && data.data.length > 0) {
                    renderTable(data);

                    if (data.studentInfo && data.studentInfo.name) {
                        document.getElementById('studentDetailsName').textContent = data.studentInfo.name;
                        const studId = localStorage.getItem('student_id') || '-';
                        const natId = localStorage.getItem('student_national_id') || '-';
                        document.getElementById('studentDetailsId').innerHTML = `كود الطالب: <span style="color:var(--text-primary)">${studId}</span><br>الرقم القومي: <span style="color:var(--text-primary)">${natId}</span>`;
                        document.getElementById('studentDetailsFaculty').textContent = data.studentInfo.faculty || '';
                        document.getElementById('studentDetailsLevel').textContent = data.studentInfo.level || '';

                        document.getElementById('studentDetailsBanner').style.display = 'flex';
                    }

                    resultsSection.classList.remove('hide');
                } else {
                    messageEl.textContent = 'لا توجد نتائج مسجلة لهذا الفصل الدراسي حالياً.';
                    messageEl.classList.add('error'); // use error style for empty state warning
                    messageEl.classList.remove('hide');
                }

            } else {
                // If token expired
                if (response.status === 401) {
                    localStorage.removeItem('student_session');
                    alert('انتهت الجلسة. يرجى تسجيل الدخول مجدداً.');
                    window.location.href = 'index.html';
                } else {
                    throw new Error(data.message || 'حدث خطأ أثناء جلب النتائج');
                }
            }

        } catch (error) {
            messageEl.textContent = error.message;
            messageEl.classList.add('error');
            messageEl.classList.remove('hide');
        } finally {
            // UI Reset
            searchBtn.disabled = false;
            spinner.classList.add('hide');
        }
    });

    function getGradeBadge(gradeStr) {
        if (!gradeStr) return '-';
        const gradeMap = {
            'A+': 'grade-aplus',
            'A': 'grade-a',
            'A-': 'grade-aminus',
            'B+': 'grade-bplus',
            'B': 'grade-b',
            'C+': 'grade-cplus',
            'C': 'grade-c',
            'D': 'grade-d',
            'D-': 'grade-dminus',
            'F': 'grade-f',
            'P': 'grade-p',
            'NP': 'grade-np'
        };
        const g = gradeStr.trim().toUpperCase();
        if (gradeMap[g]) {
            return `<span class="grade-badge ${gradeMap[g]}">${g}</span>`;
        }
        return `<span>${gradeStr}</span>`;
    }

    function renderTable(dataPayload) {
        resultsTableBody.innerHTML = '';
        const results = dataPayload.data || [];

        results.forEach(item => {
            const tr = document.createElement('tr');

            // Format status color
            let statusHtml = `<span>${item.status}</span>`;
            if (item.status.includes('ناجح') || item.status.includes('مقبول') || item.status.includes('جيد')) {
                statusHtml = `<span class="status-pass">${item.status}</span>`;
            } else if (item.status.includes('راسب') || item.status.includes('ضعيف')) {
                statusHtml = `<span class="status-fail">${item.status}</span>`;
            }

            const finalMarkDisplay = item.finalMark || '-';
            const gradeDisplay = getGradeBadge(item.grade);

            tr.innerHTML = `
                <td>${item.subjectCode || '-'}</td>
                <td><strong>${item.subjectName || '-'}</strong></td>
                <td>${statusHtml}</td>
                <td>${finalMarkDisplay}</td>
                <td>${gradeDisplay}</td>
            `;
            resultsTableBody.appendChild(tr);
        });

        // Render Summary (GPA, Total Marks)
        const summaryContainer = document.querySelector('.summary-cards');
        if (summaryContainer) {
            summaryContainer.innerHTML = ''; // Clear previous summaries

            let summaryHTML = '';

            // Total Marks Sum
            if (dataPayload.totalMarksSum !== undefined) {
                summaryHTML += `
                <div style="flex: 1; min-width: 150px; background: rgba(0,0,0,0.2); border: 1px solid var(--surface-border); border-radius: 12px; text-align: center; padding: 15px;">
                    <h3 style="margin-bottom: 5px; color: var(--text-secondary); font-size: 1rem;">المجموع الكلي</h3>
                    <div style="font-size: 24px; font-weight: bold; color: var(--primary-color);">${dataPayload.totalMarksSum.toFixed(2)}</div>
                </div>`;
            }

            // GPAs
            if (dataPayload.studentInfo && (dataPayload.studentInfo.gpa1 || dataPayload.studentInfo.gpau1)) {
                summaryHTML += `
                 <div style="flex: 2; display: flex; gap: 15px; min-width: 300px;">
                    <div style="flex: 1; background: rgba(0,0,0,0.2); border: 1px solid var(--surface-border); border-radius: 12px; text-align: center; padding: 15px;">
                        <h3 style="margin-bottom: 5px; color: var(--text-secondary); font-size: 1rem;">GPA1</h3>
                        <div style="font-size: 20px; font-weight: bold; color: #eab308;">${dataPayload.studentInfo.gpa1 || '-'}</div>
                    </div>
                    <div style="flex: 1; background: rgba(0,0,0,0.2); border: 1px solid var(--surface-border); border-radius: 12px; text-align: center; padding: 15px;">
                        <h3 style="margin-bottom: 5px; color: var(--text-secondary); font-size: 1rem;">GPAU1</h3>
                        <div style="font-size: 20px; font-weight: bold; color: #eab308;">${dataPayload.studentInfo.gpau1 || '-'}</div>
                    </div>
                 </div>`;
            }

            summaryContainer.innerHTML = summaryHTML;
        }
    }
}
