document.addEventListener('DOMContentLoaded', () => {
    const startBtn = document.getElementById('startBtn');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const logsBox = document.getElementById('logsBox');

    const statTotal = document.getElementById('statTotal');
    const statSuccess = document.getElementById('statSuccess');
    const statFailed = document.getElementById('statFailed');

    const downloadContainer = document.getElementById('downloadContainer');

    let eventSource = null;

    function addLog(message, type = 'info') {
        const div = document.createElement('div');
        div.className = `log-entry log-${type}`;

        const time = new Date().toLocaleTimeString('ar-EG');
        div.textContent = `[${time}] ${message}`;

        logsBox.appendChild(div);
        logsBox.scrollTop = logsBox.scrollHeight;
    }

    startBtn.addEventListener('click', async () => {
        // Disable button
        startBtn.disabled = true;
        startBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> <span>جاري تحضير السحب...</span>';
        downloadContainer.style.display = 'none';
        downloadContainer.innerHTML = '';

        // Reset stats
        statTotal.textContent = '0';
        statSuccess.textContent = '0';
        statFailed.textContent = '0';
        progressBar.style.width = '0%';
        progressText.textContent = '0% (0 / 0)';
        logsBox.innerHTML = '';
        addLog('بدء عملية التنظيف وتحضير السحب...', 'info');
        addLog('سيتم سحب 3 ترمات: Sem1 (2024/2025) + Sem2 (2024/2025) + Sem3 (2025/2026)', 'info');
        addLog('كل ترم يتم سحبه بالدور الأول والثاني (6 استعلامات لكل طالب)', 'info');

        try {
            // Start the process via POST
            const response = await fetch('/api/start-scraping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });

            const result = await response.json();
            if (!result.success) {
                addLog('فشل في بدء السحب: ' + result.message, 'error');
                resetBtn();
                return;
            }

            // Connect to SSE stream
            eventSource = new EventSource('/api/scrape-stream');

            eventSource.onmessage = (event) => {
                const data = JSON.parse(event.data);

                if (data.type === 'init') {
                    addLog(`تم العثور على ${data.total} طالب (بعد تنظيف المتكرر).`, 'info');
                    statTotal.textContent = data.total;
                }
                else if (data.type === 'progress') {
                    // Update stats
                    statSuccess.textContent = data.successCount;
                    statFailed.textContent = data.failCount;

                    const totalProcessed = data.successCount + data.failCount;
                    const percent = Math.round((totalProcessed / data.total) * 100);

                    progressBar.style.width = `${percent}%`;
                    progressText.textContent = `${percent}% (${totalProcessed} / ${data.total})`;

                    if (data.status === 'success') {
                        addLog(`تم السحب بنجاح: طالب [${data.student}]`, 'success');
                    } else if (data.status === 'error') {
                        addLog(`خطأ للطالب [${data.student}]: ${data.message}`, 'error');
                    }
                }
                else if (data.type === 'complete') {
                    addLog('✅ اكتملت عملية السحب بنجاح!', 'success');

                    progressBar.style.width = '100%';
                    progressText.textContent = 'اكتمل 100%';

                    // Show multiple download buttons
                    if (data.files && data.files.length > 0) {
                        downloadContainer.innerHTML = '';
                        data.files.forEach(file => {
                            const link = document.createElement('a');
                            link.href = `/api/download-excel?file=${encodeURIComponent(file.fileName)}`;
                            link.className = 'primary-btn';
                            link.style.cssText = 'background: #10b981; text-decoration: none; padding: 10px 20px; display: inline-flex; align-items: center; gap: 8px;';
                            link.innerHTML = `<i class="fa-solid fa-file-excel"></i> <span>${file.label} (${file.count} طالب)</span>`;
                            downloadContainer.appendChild(link);
                            addLog(`ملف جاهز: ${file.fileName} - ${file.count} طالب`, 'info');
                        });
                        downloadContainer.style.display = 'flex';
                    } else {
                        addLog('⚠️ لم يتم إنشاء أي ملفات (لا نتائج)', 'error');
                    }

                    eventSource.close();
                    resetBtn();
                }
            };

            eventSource.onerror = (err) => {
                console.error("SSE Error:", err);
                addLog('انقطع الاتصال بالخادم أثناء السحب.', 'error');
                if (eventSource) {
                    eventSource.close();
                }
                resetBtn();
            };

        } catch (error) {
            console.error(error);
            addLog('حدث خطأ في الاتصال بالخادم.', 'error');
            resetBtn();
        }
    });

    function resetBtn() {
        startBtn.disabled = false;
        startBtn.innerHTML = '<i class="fa-solid fa-play"></i> <span>بدء عملية السحب الآن</span>';
    }
});
