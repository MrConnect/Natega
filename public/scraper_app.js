document.addEventListener('DOMContentLoaded', () => {
    const startBtn = document.getElementById('startBtn');
    const stopBtn = document.getElementById('stopBtn');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const logsBox = document.getElementById('logsBox');

    const statTotal = document.getElementById('statTotal');
    const statSuccess = document.getElementById('statSuccess');
    const statFailed = document.getElementById('statFailed');
    const statTimer = document.getElementById('statTimer');

    const activityPanel = document.getElementById('activityPanel');
    const activityTitle = document.getElementById('activityTitle');
    const actCurrentStudent = document.getElementById('actCurrentStudent');
    const actStartYear = document.getElementById('actStartYear');
    const actQuery = document.getElementById('actQuery');
    const actQueryProgress = document.getElementById('actQueryProgress');
    const queryProgressBar = document.getElementById('queryProgressBar');

    const downloadArea = document.getElementById('downloadArea');
    const downloadLinks = document.getElementById('downloadLinks');

    let eventSource = null;
    let timerInterval = null;
    let startTime = null;

    function addLog(message, type = 'info') {
        const div = document.createElement('div');
        div.className = `log-entry log-${type}`;
        const time = new Date().toLocaleTimeString('ar-EG');
        div.textContent = `[${time}] ${message}`;
        logsBox.appendChild(div);
        logsBox.scrollTop = logsBox.scrollHeight;
    }

    function startTimer() {
        startTime = Date.now();
        timerInterval = setInterval(() => {
            const elapsed = Math.floor((Date.now() - startTime) / 1000);
            const mins = String(Math.floor(elapsed / 60)).padStart(2, '0');
            const secs = String(elapsed % 60).padStart(2, '0');
            statTimer.textContent = `${mins}:${secs}`;
        }, 1000);
    }

    function stopTimer() {
        if (timerInterval) {
            clearInterval(timerInterval);
            timerInterval = null;
        }
    }

    function setScrapingUI(active) {
        if (active) {
            startBtn.disabled = true;
            startBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>جاري السحب...</span>';
            stopBtn.classList.add('visible');
            activityPanel.classList.add('active');
            downloadArea.classList.remove('visible');
            downloadLinks.innerHTML = '';
        } else {
            startBtn.disabled = false;
            startBtn.innerHTML = '<i class="fa-solid fa-play"></i> <span>بدء عملية السحب</span>';
            stopBtn.classList.remove('visible');
            activityPanel.classList.remove('active');
            stopBtn.disabled = false;
            stopBtn.innerHTML = '<i class="fa-solid fa-stop"></i> <span>إيقاف السحب</span>';
        }
    }

    // Start button
    startBtn.addEventListener('click', async () => {
        setScrapingUI(true);

        // Reset
        statTotal.textContent = '0';
        statSuccess.textContent = '0';
        statFailed.textContent = '0';
        statTimer.textContent = '00:00';
        progressBar.style.width = '0%';
        progressText.textContent = '0% (0 / 0)';
        logsBox.innerHTML = '';

        addLog('بدء عملية التحضير...', 'info');
        addLog('سيتم تحديد سنة بداية كل طالب من كوده وسحب النتائج لحد 2025/2026', 'info');
        addLog('المواد هتتوزع على 3 ملفات (ترم 1 / 2 / 3) بناءً على كود المادة', 'info');

        try {
            const response = await fetch('/api/start-scraping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });

            const result = await response.json();
            if (!result.success) {
                addLog('❌ فشل في بدء السحب: ' + result.message, 'error');
                setScrapingUI(false);
                return;
            }

            startTimer();

            // SSE
            eventSource = new EventSource('/api/scrape-stream');

            eventSource.onmessage = (event) => {
                const data = JSON.parse(event.data);

                switch (data.type) {
                    case 'init':
                        addLog(`📋 تم العثور على ${data.total} طالب (بعد تنظيف المتكرر)`, 'info');
                        statTotal.textContent = data.total;
                        break;

                    case 'student_start':
                        actCurrentStudent.textContent = data.student;
                        actStartYear.textContent = data.startYear;
                        activityTitle.textContent = `جاري معالجة الطالب ${data.index} / ${data.total}`;
                        actQuery.textContent = 'جاري التحضير...';
                        actQueryProgress.textContent = `0/${data.totalQueries}`;
                        queryProgressBar.style.width = '0%';
                        addLog(`👤 [${data.index}/${data.total}] بدء سحب الطالب ${data.student} — بدأ من ${data.startYear} (${data.totalQueries} استعلام)`, 'info');
                        break;

                    case 'query':
                        actQuery.textContent = data.queryLabel;
                        actQueryProgress.textContent = `${data.queryNum}/${data.totalQueries}`;
                        const qPercent = Math.round((data.queryNum / data.totalQueries) * 100);
                        queryProgressBar.style.width = `${qPercent}%`;
                        addLog(`  🔍 ${data.queryLabel}`, 'query');
                        break;

                    case 'progress': {
                        statSuccess.textContent = data.successCount;
                        statFailed.textContent = data.failCount;
                        const totalProcessed = data.successCount + data.failCount;
                        const percent = Math.round((totalProcessed / data.total) * 100);
                        progressBar.style.width = `${percent}%`;
                        progressText.textContent = `${percent}% (${totalProcessed} / ${data.total})`;

                        if (data.status === 'success') {
                            addLog(`✅ تم السحب بنجاح: ${data.student}`, 'success');
                        } else if (data.status === 'error') {
                            addLog(`❌ خطأ: ${data.student} — ${data.message}`, 'error');
                        }
                        break;
                    }

                    case 'info':
                        addLog(`ℹ️ ${data.message}`, 'warn');
                        break;

                    case 'complete':
                        stopTimer();
                        addLog('🎉 اكتملت عملية السحب!', 'success');

                        progressBar.style.width = '100%';
                        progressText.textContent = 'اكتمل 100%';

                        if (data.files && data.files.length > 0) {
                            downloadLinks.innerHTML = '';
                            data.files.forEach(file => {
                                const link = document.createElement('a');
                                link.href = `/api/download-excel?file=${encodeURIComponent(file.fileName)}`;
                                link.className = 'download-btn';
                                link.innerHTML = `<i class="fa-solid fa-file-excel"></i> ${file.label} (${file.count} طالب)`;
                                downloadLinks.appendChild(link);
                                addLog(`📁 ملف جاهز: ${file.fileName} — ${file.count} طالب`, 'success');
                            });
                            downloadArea.classList.add('visible');
                        } else {
                            addLog('⚠️ لم يتم إنشاء أي ملفات (لا نتائج)', 'warn');
                        }

                        eventSource.close();
                        setScrapingUI(false);
                        break;
                }
            };

            eventSource.onerror = () => {
                stopTimer();
                addLog('⚠️ انقطع الاتصال بالخادم.', 'error');
                if (eventSource) eventSource.close();
                setScrapingUI(false);
            };

        } catch (error) {
            stopTimer();
            console.error(error);
            addLog('❌ حدث خطأ في الاتصال بالخادم.', 'error');
            setScrapingUI(false);
        }
    });

    // Stop button
    stopBtn.addEventListener('click', async () => {
        stopBtn.disabled = true;
        stopBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>جاري الإيقاف...</span>';
        addLog('⏳ تم طلب إيقاف السحب — سيتوقف بعد الطالب الحالي...', 'warn');

        try {
            const response = await fetch('/api/stop-scraping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });
            const result = await response.json();
            if (result.success) {
                addLog('⏹️ ' + result.message, 'warn');
            } else {
                addLog('⚠️ ' + result.message, 'error');
            }
        } catch (e) {
            addLog('❌ فشل إرسال طلب الإيقاف', 'error');
        }
    });
});
