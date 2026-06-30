// ==UserScript==
// @name         千牛聊天记录导出器 Excel
// @namespace    https://tampermonkey.net/
// @version      5.5
// @description  仅在千牛聊天记录查询内层页面运行，按原始顺序逐条导出聊天记录为 Excel；一个客户一个工作表；不限制、不过滤营销信息
// @author       ChatGPT
// @match        *://myseller.taobao.com/*
// @match        *://qianniu.taobao.com/*
// @match        *://*.taobao.com/*
// @match        *://*.tmall.com/*
// @grant        none
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
    'use strict';

    /**
     * 只允许在聊天记录 iframe / 内层页面运行。
     * 外层页面不运行，避免出现两个窗口和乱点。
     */
    if (window.top === window.self) {
        return;
    }

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    const state = {
        running: false,
        stop: false,
        maxUsers: 20,
        waitAfterClick: 1500,
        scrollTimes: 2,
        scrollDelay: 700
    };

    const serviceNameKeywords = [
        '客服',
        '专卖店',
        '旗舰店',
        '店铺',
        '店',
        '售后',
        '售前',
        '接待',
        '导购',
        '小二'
    ];

    function getCleanText(el) {
        if (!el) return '';

        return (el.innerText || el.textContent || '')
            .replace(/\u00a0/g, ' ')
            .replace(/[ \t]+/g, ' ')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    }

    function isVisible(el) {
        if (!el) return false;

        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);

        return (
            rect.width > 5 &&
            rect.height > 5 &&
            style.display !== 'none' &&
            style.visibility !== 'hidden' &&
            style.opacity !== '0'
        );
    }

    function safeExcelText(text, maxLength = 32767) {
        if (text === undefined || text === null) return '';
        return String(text).slice(0, maxLength);
    }

    function getNowText() {
        const now = new Date();

        return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')} ` +
            `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
    }

    function isTimeLine(line) {
        const text = String(line || '').trim();

        return /^(\d{1,2}:\d{2}(:\d{2})?)$/.test(text);
    }

    function isDateLine(line) {
        const text = String(line || '').trim();

        return /^(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}(日)?)$/.test(text);
    }

    function isChatRecordPage() {
        const bodyText = document.body ? document.body.innerText || '' : '';

        return (
            bodyText.includes('聊天记录查询') ||
            bodyText.includes('查询结果') ||
            bodyText.includes('客服昵称') ||
            bodyText.includes('员工账号')
        );
    }

    function detectAccountType(sender, customerName) {
        const s = String(sender || '').trim();
        const customer = String(customerName || '').trim();

        if (!s) return '';

        if (customer && s === customer) {
            return '客户';
        }

        if (customer && s.includes(customer)) {
            return '客户';
        }

        if (serviceNameKeywords.some(k => s.includes(k))) {
            return '客服';
        }

        if (s.includes(':') || s.includes('：')) {
            return '客服';
        }

        return '客户';
    }

    function getCustomerListContainer() {
        let resultItems = Array.from(document.querySelectorAll('.results-list, [class*="results-list"]'))
            .filter(isVisible);

        if (resultItems.length) {
            const parentMap = new Map();

            resultItems.forEach(item => {
                let p = item.parentElement;
                let depth = 0;

                while (p && depth < 5) {
                    const rect = p.getBoundingClientRect();

                    if (
                        isVisible(p) &&
                        rect.width >= 80 &&
                        rect.width <= 380 &&
                        rect.height >= 100
                    ) {
                        parentMap.set(p, (parentMap.get(p) || 0) + 1);
                    }

                    p = p.parentElement;
                    depth++;
                }
            });

            const parents = Array.from(parentMap.entries())
                .sort((a, b) => b[1] - a[1])
                .map(item => item[0]);

            if (parents.length) return parents[0];
        }

        const candidates = Array.from(document.querySelectorAll('div'))
            .filter(el => {
                if (!isVisible(el)) return false;

                const rect = el.getBoundingClientRect();
                const text = getCleanText(el);

                if (!text) return false;

                return (
                    rect.left < window.innerWidth * 0.45 &&
                    rect.width >= 100 &&
                    rect.width <= 380 &&
                    rect.height >= 200 &&
                    text.length >= 10
                );
            })
            .sort((a, b) => {
                const ra = a.getBoundingClientRect();
                const rb = b.getBoundingClientRect();
                return rb.height - ra.height;
            });

        return candidates[0] || null;
    }

    function getCustomerItems() {
        const container = getCustomerListContainer();

        if (!container) return [];

        let items = Array.from(container.querySelectorAll('.results-list, [class*="results-list"]'))
            .filter(el => {
                if (!isVisible(el)) return false;

                const text = getCleanText(el);
                const rect = el.getBoundingClientRect();

                if (!text) return false;
                if (text.length > 200) return false;
                if (rect.height < 18 || rect.height > 90) return false;

                return true;
            });

        if (!items.length) {
            items = Array.from(container.children)
                .filter(el => {
                    if (!isVisible(el)) return false;

                    const text = getCleanText(el);
                    const rect = el.getBoundingClientRect();

                    if (!text) return false;
                    if (text.length > 200) return false;
                    if (rect.height < 18 || rect.height > 100) return false;

                    const badWords = [
                        '聊天记录查询',
                        '查询结果',
                        '客服昵称',
                        '员工账号',
                        '查询',
                        '开始导出',
                        '停止导出',
                        '最大导出客户数'
                    ];

                    return !badWords.some(word => text.includes(word));
                });
        }

        items = Array.from(new Set(items));

        items.sort((a, b) => {
            const ra = a.getBoundingClientRect();
            const rb = b.getBoundingClientRect();
            return ra.top - rb.top;
        });

        return items;
    }

    function getCustomerName(item, index) {
        if (!item) return `客户_${index + 1}`;

        const title = item.getAttribute('title');
        if (title && title.trim()) return title.trim();

        const aria = item.getAttribute('aria-label');
        if (aria && aria.trim()) return aria.trim();

        const text = getCleanText(item);
        if (!text) return `客户_${index + 1}`;

        const firstLine = text.split('\n').map(s => s.trim()).filter(Boolean)[0];

        return firstLine || `客户_${index + 1}`;
    }

    function getChatContainer() {
        const selectors = [
            '.message-list-right',
            '.message-list',
            '.message-container',
            '.chat-content',
            '.chat-list',
            '.conversation-content',
            '.dialog-content',
            '.im-message-list',
            '[class*="message"]',
            '[class*="chat"]',
            '[class*="conversation"]'
        ];

        for (const selector of selectors) {
            const matched = Array.from(document.querySelectorAll(selector))
                .filter(el => {
                    if (!isVisible(el)) return false;

                    const rect = el.getBoundingClientRect();
                    const text = getCleanText(el);

                    if (!text || text.length < 10) return false;

                    return (
                        rect.width >= 250 &&
                        rect.height >= 150 &&
                        rect.left > 80
                    );
                })
                .sort((a, b) => {
                    const ra = a.getBoundingClientRect();
                    const rb = b.getBoundingClientRect();
                    return (rb.width * rb.height) - (ra.width * ra.height);
                });

            if (matched.length) return matched[0];
        }

        const candidates = Array.from(document.querySelectorAll('div'))
            .filter(el => {
                if (!isVisible(el)) return false;

                const rect = el.getBoundingClientRect();
                const text = getCleanText(el);
                const style = window.getComputedStyle(el);

                if (!text || text.length < 20) return false;

                return (
                    rect.left > 180 &&
                    rect.width >= 300 &&
                    rect.height >= 180 &&
                    (
                        style.overflowY === 'auto' ||
                        style.overflowY === 'scroll' ||
                        el.scrollHeight > el.clientHeight + 80
                    )
                );
            })
            .sort((a, b) => {
                const ra = a.getBoundingClientRect();
                const rb = b.getBoundingClientRect();
                return (rb.width * rb.height) - (ra.width * ra.height);
            });

        return candidates[0] || null;
    }

    async function loadMoreHistory(container, times, delay) {
        if (!container) return;

        for (let i = 0; i < times; i++) {
            if (state.stop) return;

            try {
                container.scrollTop = 0;
                container.dispatchEvent(new Event('scroll', { bubbles: true }));
            } catch (e) {}

            await sleep(delay);
        }
    }

    /**
     * 核心解析逻辑：
     *
     * 目标原始格式：
     *
     * 20:57:11
     * 请在客户端查看原始聊天记录，网页版无法展示
     * 得力广州专卖店:小迪
     *
     * 20:57:13
     * http://item.taobao.com/item.htm?id=1011758423155
     * tb8946567709
     *
     * 每条消息拆为：
     * 发送时间 / 发送内容 / 发送账号
     */
    function parseMessagesFromContainerText(container, customerName) {
        const text = getCleanText(container);

        if (!text) return [];

        const rawLines = text
            .split('\n')
            .map(line => line.trim())
            .filter(Boolean)
            .filter(line => !isDateLine(line));

        const rows = [];
        let i = 0;

        while (i < rawLines.length) {
            const line = rawLines[i];

            if (!isTimeLine(line)) {
                i++;
                continue;
            }

            const sendTime = line;
            let j = i + 1;
            const block = [];

            while (j < rawLines.length && !isTimeLine(rawLines[j])) {
                block.push(rawLines[j]);
                j++;
            }

            /**
             * 正常 block：
             * [内容, 账号名]
             *
             * 如果内容多行：
             * [内容1, 内容2, 内容3, 账号名]
             *
             * 最后一行作为账号名，其余作为发送内容。
             */
            if (block.length >= 2) {
                const sender = block[block.length - 1];
                const content = block.slice(0, -1).join('\n').trim();

                const accountType = detectAccountType(sender, customerName);

                rows.push({
                    客户名: safeExcelText(customerName),
                    消息序号: rows.length + 1,
                    发送时间: safeExcelText(sendTime),
                    发送内容: safeExcelText(content),
                    发送账号: safeExcelText(sender),
                    账号类型: safeExcelText(accountType)
                });
            }

            /**
             * 如果 block 只有一行，说明页面结构可能异常。
             * 也保留，避免漏数据。
             */
            if (block.length === 1) {
                rows.push({
                    客户名: safeExcelText(customerName),
                    消息序号: rows.length + 1,
                    发送时间: safeExcelText(sendTime),
                    发送内容: safeExcelText(block[0]),
                    发送账号: '',
                    账号类型: ''
                });
            }

            i = j;
        }

        /**
         * 去重，避免页面重复节点导致重复导出。
         */
        const seen = new Set();

        return rows.filter(row => {
            const key = `${row.发送时间}|${row.发送内容}|${row.发送账号}`;

            if (seen.has(key)) return false;

            seen.add(key);
            return true;
        }).map((row, index) => {
            row.消息序号 = index + 1;
            return row;
        });
    }

    function log(text) {
        const box = document.querySelector('#qn-export-log');
        if (!box) return;

        const time = new Date().toLocaleTimeString();

        box.value += `[${time}] ${text}\n`;
        box.scrollTop = box.scrollHeight;
    }

    function updateProgress(current, total) {
        const fill = document.querySelector('#qn-export-progress-fill');
        const text = document.querySelector('#qn-export-progress-text');

        const pct = total ? Math.round((current / total) * 100) : 0;

        if (fill) fill.style.width = pct + '%';
        if (text) text.textContent = `${current}/${total} · ${pct}%`;
    }

    function safeSheetName(name) {
        let sheetName = String(name || '客户')
            .replace(/[\\\/\?\*\[\]\:]/g, '_')
            .replace(/\s+/g, ' ')
            .trim();

        if (!sheetName) sheetName = '客户';

        return sheetName.slice(0, 31);
    }

    function getUniqueSheetName(workbook, baseName) {
        let name = safeSheetName(baseName);
        let finalName = name;
        let index = 1;

        while (workbook.SheetNames.includes(finalName)) {
            const suffix = `_${index}`;
            finalName = name.slice(0, 31 - suffix.length) + suffix;
            index++;
        }

        return finalName;
    }

    function exportToExcel(customerDataList, summaryRows) {
        if (typeof XLSX === 'undefined') {
            throw new Error('Excel 导出库未加载成功，请刷新页面或检查网络。');
        }

        const workbook = XLSX.utils.book_new();

        const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
        summarySheet['!cols'] = [
            { wch: 24 },
            { wch: 12 },
            { wch: 14 },
            { wch: 18 },
            { wch: 40 },
            { wch: 22 }
        ];

        XLSX.utils.book_append_sheet(workbook, summarySheet, '导出汇总');

        customerDataList.forEach(customer => {
            const sheet = XLSX.utils.json_to_sheet(customer.rows);

            sheet['!cols'] = [
                { wch: 24 },
                { wch: 10 },
                { wch: 14 },
                { wch: 90 },
                { wch: 30 },
                { wch: 12 }
            ];

            const sheetName = getUniqueSheetName(workbook, customer.userName);

            XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
        });

        const now = new Date();

        const fileName =
            `千牛聊天记录_原始逐条_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_` +
            `${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}.xlsx`;

        XLSX.writeFile(workbook, fileName);
    }

    async function startExport() {
        if (state.running) return;

        state.running = true;
        state.stop = false;

        const startBtn = document.querySelector('#qn-export-start');
        const stopBtn = document.querySelector('#qn-export-stop');

        if (startBtn) {
            startBtn.disabled = true;
            startBtn.textContent = '导出中...';
        }

        if (stopBtn) {
            stopBtn.disabled = false;
        }

        const customerDataList = [];
        const summaryRows = [];

        try {
            state.maxUsers = parseInt(document.querySelector('#qn-max-users')?.value, 10) || 20;
            state.waitAfterClick = parseInt(document.querySelector('#qn-wait-after-click')?.value, 10) || 1500;
            state.scrollTimes = parseInt(document.querySelector('#qn-scroll-times')?.value, 10) || 2;
            state.scrollDelay = parseInt(document.querySelector('#qn-scroll-delay')?.value, 10) || 700;

            log('开始识别聊天记录查询页面...');

            if (!isChatRecordPage()) {
                throw new Error('当前内层页面不像聊天记录查询页面，请先进入“聊天记录查询”页面后再导出。');
            }

            log('开始识别客户列表...');

            let items = getCustomerItems();

            if (!items.length) {
                throw new Error('没有识别到客户列表。请确认查询结果左侧已有客户列表。');
            }

            items = items.slice(0, state.maxUsers);

            log(`识别到 ${items.length} 个客户，开始按原始逐条格式导出。`);

            updateProgress(0, items.length);

            for (let i = 0; i < items.length; i++) {
                if (state.stop) {
                    log('已停止导出，开始生成当前已抓取的数据。');
                    break;
                }

                const item = items[i];
                const userName = getCustomerName(item, i);
                const exportTime = getNowText();

                log(`正在处理 ${i + 1}/${items.length}：${userName}`);

                try {
                    item.scrollIntoView({
                        behavior: 'instant',
                        block: 'center'
                    });
                } catch (e) {}

                item.click();

                await sleep(state.waitAfterClick);

                const container = getChatContainer();

                if (!container) {
                    log(`失败：未找到聊天内容区域：${userName}`);

                    summaryRows.push({
                        客户名: safeExcelText(userName),
                        状态: '失败',
                        消息数: 0,
                        是否生成工作表: '否',
                        失败原因: '未找到聊天内容区域',
                        导出时间: exportTime
                    });

                    updateProgress(i + 1, items.length);
                    continue;
                }

                await loadMoreHistory(container, state.scrollTimes, state.scrollDelay);

                const rows = parseMessagesFromContainerText(container, userName);

                if (rows.length > 0) {
                    customerDataList.push({
                        userName,
                        rows
                    });
                }

                summaryRows.push({
                    客户名: safeExcelText(userName),
                    状态: '成功',
                    消息数: rows.length,
                    是否生成工作表: rows.length > 0 ? '是' : '否',
                    失败原因: rows.length > 0 ? '' : '无有效消息',
                    导出时间: exportTime
                });

                log(`完成：${userName}，导出 ${rows.length} 条消息。`);

                updateProgress(i + 1, items.length);

                await sleep(300);
            }

            if (!summaryRows.length) {
                throw new Error('没有抓取到任何数据。');
            }

            log('正在生成 Excel 文件...');
            exportToExcel(customerDataList, summaryRows);
            log('Excel 导出完成。');

        } catch (err) {
            console.error(err);
            alert('导出失败：' + err.message);
            log('导出失败：' + err.message);
        } finally {
            state.running = false;

            if (startBtn) {
                startBtn.disabled = false;
                startBtn.textContent = '开始导出 Excel';
            }

            if (stopBtn) {
                stopBtn.disabled = true;
            }
        }
    }

    function stopExport() {
        state.stop = true;
        log('正在停止，请等待当前客户处理完成...');
    }

    function createPanel() {
        if (document.querySelector('#qn-export-panel')) return;

        const panel = document.createElement('div');

        panel.id = 'qn-export-panel';
        panel.style.cssText = `
            position: fixed;
            right: 20px;
            bottom: 20px;
            width: 370px;
            background: #ffffff;
            border: 1px solid #ddd;
            border-radius: 10px;
            box-shadow: 0 6px 24px rgba(0,0,0,0.18);
            z-index: 999999999;
            font-size: 13px;
            color: #333;
            overflow: hidden;
            font-family: Arial, "Microsoft YaHei", sans-serif;
        `;

        panel.innerHTML = `
            <div id="qn-export-drag" style="
                background:#111827;
                color:#fff;
                padding:10px 12px;
                cursor:move;
                display:flex;
                justify-content:space-between;
                align-items:center;
                user-select:none;
            ">
                <strong>千牛聊天记录导出器 原始逐条版</strong>
                <button id="qn-export-minimize" style="
                    background:transparent;
                    color:#fff;
                    border:none;
                    cursor:pointer;
                    font-size:16px;
                ">－</button>
            </div>

            <div id="qn-export-body" style="padding:12px;">
                <div style="display:grid; grid-template-columns: 1fr 90px; gap:8px; align-items:center; margin-bottom:8px;">
                    <label>最大导出客户数</label>
                    <input id="qn-max-users" type="number" value="20" min="1" style="padding:5px; border:1px solid #ccc; border-radius:4px;">
                </div>

                <div style="display:grid; grid-template-columns: 1fr 90px; gap:8px; align-items:center; margin-bottom:8px;">
                    <label>切换等待时间 ms</label>
                    <input id="qn-wait-after-click" type="number" value="1500" min="300" style="padding:5px; border:1px solid #ccc; border-radius:4px;">
                </div>

                <div style="display:grid; grid-template-columns: 1fr 90px; gap:8px; align-items:center; margin-bottom:8px;">
                    <label>向上滚动次数</label>
                    <input id="qn-scroll-times" type="number" value="2" min="0" style="padding:5px; border:1px solid #ccc; border-radius:4px;">
                </div>

                <div style="display:grid; grid-template-columns: 1fr 90px; gap:8px; align-items:center; margin-bottom:10px;">
                    <label>滚动等待时间 ms</label>
                    <input id="qn-scroll-delay" type="number" value="700" min="100" style="padding:5px; border:1px solid #ccc; border-radius:4px;">
                </div>

                <div style="display:flex; gap:8px; margin-bottom:10px;">
                    <button id="qn-export-start" style="
                        flex:1;
                        background:#16a34a;
                        color:#fff;
                        border:none;
                        border-radius:6px;
                        padding:8px;
                        cursor:pointer;
                    ">开始导出 Excel</button>

                    <button id="qn-export-stop" disabled style="
                        flex:1;
                        background:#dc2626;
                        color:#fff;
                        border:none;
                        border-radius:6px;
                        padding:8px;
                        cursor:pointer;
                    ">停止导出</button>
                </div>

                <div style="height:8px; background:#eee; border-radius:999px; overflow:hidden; margin-bottom:6px;">
                    <div id="qn-export-progress-fill" style="
                        width:0%;
                        height:100%;
                        background:#16a34a;
                        transition:width .2s;
                    "></div>
                </div>

                <div id="qn-export-progress-text" style="text-align:center; color:#666; margin-bottom:8px;">0/0 · 0%</div>

                <textarea id="qn-export-log" readonly style="
                    width:100%;
                    height:130px;
                    resize:none;
                    border:1px solid #ddd;
                    border-radius:6px;
                    padding:6px;
                    box-sizing:border-box;
                    font-size:12px;
                    line-height:1.4;
                "></textarea>

                <div style="font-size:12px; color:#777; line-height:1.5; margin-top:8px;">
                    导出格式：客户名、消息序号、发送时间、发送内容、发送账号、账号类型。此版本不做营销过滤，页面展示什么就尽量导出什么。
                </div>
            </div>
        `;

        document.body.appendChild(panel);

        document.querySelector('#qn-export-start').addEventListener('click', startExport);
        document.querySelector('#qn-export-stop').addEventListener('click', stopExport);

        const minimizeBtn = document.querySelector('#qn-export-minimize');
        const body = document.querySelector('#qn-export-body');

        minimizeBtn.addEventListener('click', () => {
            if (body.style.display === 'none') {
                body.style.display = 'block';
                minimizeBtn.textContent = '－';
            } else {
                body.style.display = 'none';
                minimizeBtn.textContent = '+';
            }
        });

        makePanelDraggable(panel, document.querySelector('#qn-export-drag'));

        log('面板已加载。当前版本不做营销过滤，按原始消息逐条导出。');
    }

    function makePanelDraggable(panel, handle) {
        let startX = 0;
        let startY = 0;
        let startLeft = 0;
        let startTop = 0;
        let dragging = false;

        handle.addEventListener('mousedown', (e) => {
            dragging = true;

            const rect = panel.getBoundingClientRect();

            startX = e.clientX;
            startY = e.clientY;
            startLeft = rect.left;
            startTop = rect.top;

            panel.style.right = 'auto';
            panel.style.bottom = 'auto';
            panel.style.left = startLeft + 'px';
            panel.style.top = startTop + 'px';

            e.preventDefault();
        });

        document.addEventListener('mousemove', (e) => {
            if (!dragging) return;

            const dx = e.clientX - startX;
            const dy = e.clientY - startY;

            panel.style.left = startLeft + dx + 'px';
            panel.style.top = startTop + dy + 'px';
        });

        document.addEventListener('mouseup', () => {
            dragging = false;
        });
    }

    function init() {
        let retry = 0;

        const timer = setInterval(() => {
            retry++;

            if (document.body && isChatRecordPage()) {
                clearInterval(timer);
                createPanel();
            }

            if (retry >= 20) {
                clearInterval(timer);
            }
        }, 1000);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

})();
