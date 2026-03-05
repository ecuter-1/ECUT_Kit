// ==UserScript==
// @name         东华理工大学成绩增强系统
// @namespace    https://github.com/yourusername/ecut-grade-enhancer
// @version      4.1.0
// @description  成绩GPA计算、加权均分、筛选排序、Excel导出、全学期汇总，支持iPhone/iPad触控拖动
// @author       ECUT Grade Enhancer Contributors
// @license      MIT
// @match        *://172.20.130.13/jwglxt/cjcx/*
// @match        *://*.ecut.edu.cn/jwglxt/cjcx/*
// @match        *://ehall.ecut.edu.cn/jwglxt/cjcx/*
// @icon         https://www.ecut.edu.cn/favicon.ico
// @require      https://cdn.bootcdn.net/ajax/libs/jquery/3.6.4/jquery.min.js
// @require      https://cdn.bootcdn.net/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        GM_addStyle
// @grant        GM_registerMenuCommand
// @grant        GM_setValue
// @grant        GM_getValue
// @run-at       document-end
// @homepageURL  https://github.com/yourusername/ecut-grade-enhancer
// @supportURL   https://github.com/yourusername/ecut-grade-enhancer/issues
// ==/UserScript==

/* global XLSX, jQuery */
/* jshint esversion: 6 */

(function () {
    'use strict';

    var $ = window.jQuery || window.$;

    // ================= 配置常量 =================
    const CONFIG = {
        SCORE_MAP: {
            '优秀': 95, '优': 95,
            '良好': 85, '良': 85,
            '中等': 75, '中': 75,
            '及格': 65,
            '合格': 80, '通过': 80,
            '不及格': 50, '不合格': 50, '不通过': 50,
            '缓考': 0, '缺考': 0, '作弊': 0
        },
        // 页面条数设置
        PAGE_SIZE_DEFAULT: 1000,
        PAGE_SIZE_ALL: 1500,
        // 防抖延迟(ms)
        DEBOUNCE_SCAN: 800,
        DEBOUNCE_RESIZE: 300,
        // 轮询间隔(ms) — 仅在 MutationObserver 未覆盖时使用
        POLL_INTERVAL: 5000,

        CSS: `
            /* ===== 主容器 ===== */
            #zf-helper-app {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                             "Helvetica Neue", Arial, sans-serif;
                position: fixed;
                top: 80px;
                right: 20px;
                z-index: 2147483646;
                background: #fff;
                box-shadow: 0 12px 40px rgba(0,0,0,0.2);
                border-radius: 12px;
                font-size: 14px;
                color: #333;
                width: 850px;
                height: 600px;
                max-width: 95vw;
                max-height: 90vh;
                min-width: 350px;
                min-height: 200px;
                border: 1px solid #ebeef5;
                display: flex;
                flex-direction: column;
                transition: box-shadow 0.3s ease;
                user-select: none;
                -webkit-user-select: none;
                /* resize 需要 overflow:visible 在外层，内容用内层 overflow 控制 */
                resize: both;
                overflow: auto;
                -webkit-tap-highlight-color: transparent;
                touch-action: none;
            }

            /* 拖拽中取消过渡以防抖动 */
            #zf-helper-app.zf-dragging {
                transition: none;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            }

            /* ===== 响应式 ===== */
            @media only screen and (min-width: 1600px) {
                #zf-helper-app { width: 900px; height: 650px; }
            }
            @media only screen and (max-width: 1400px) {
                #zf-helper-app { width: 750px; height: 550px; }
            }
            @media only screen and (max-width: 1024px) {
                #zf-helper-app { width: 85vw; height: 70vh; max-width: 85vw; }
            }

            /* iPhone 竖屏 */
            @media only screen and (max-width: 480px) and (orientation: portrait) {
                #zf-helper-app {
                    top: 10px !important;
                    left: 2.5vw !important;
                    right: 2.5vw !important;
                    width: 95vw !important;
                    height: 82vh !important;
                    max-width: 95vw !important;
                    max-height: 90vh !important;
                    font-size: 13px;
                    resize: none;
                }
                .zf-dashboard {
                    grid-template-columns: repeat(3, 1fr) !important;
                    gap: 5px !important;
                    padding: 10px !important;
                    margin-bottom: 10px !important;
                }
                .zf-stat-val   { font-size: 14px !important; }
                .zf-stat-label { font-size: 10px !important; }
                .zf-table-wrapper { height: 45vh !important; min-height: 250px !important; }
                .zf-table th, .zf-table td { padding: 8px 6px !important; font-size: 12px !important; }
                .zf-btns { flex-wrap: wrap !important; }
                .zf-btn  { padding: 8px 12px !important; font-size: 12px !important; min-height: 40px !important; }
                #zf-helper-body { max-height: calc(82vh - 50px) !important; }

                #zf-helper-app.zf-minimized {
                    width: 220px !important;
                    height: auto !important;
                    min-height: auto !important;
                    padding: 0 !important;
                    top: 10px !important;
                    right: 10px !important;
                    left: auto !important;
                    touch-action: manipulation;
                }
                #zf-helper-app.zf-minimized #zf-helper-header { min-height: 44px; padding: 12px 15px; }
                #zf-helper-app.zf-minimized #zf-toggle-btn   { min-width: 60px; padding: 8px 12px; font-size: 12px; }
            }

            /* iPad */
            @media only screen and (min-width: 481px) and (max-width: 1024px) {
                #zf-helper-app {
                    top: 40px !important;
                    left: 2.5vw !important;
                    right: 2.5vw !important;
                    width: 95vw !important;
                    height: 75vh !important;
                    max-width: 95vw !important;
                    max-height: 85vh !important;
                }
                .zf-dashboard { grid-template-columns: repeat(3, 1fr) !important; gap: 8px !important; padding: 12px !important; }
                .zf-table th, .zf-table td { padding: 10px 8px !important; font-size: 13px !important; }
                .zf-btn { padding: 10px 14px !important; min-height: 44px !important; font-size: 13px !important; }
            }

            /* 横屏 */
            @media only screen and (max-height: 500px) and (orientation: landscape) {
                #zf-helper-app {
                    top: 5px !important;
                    height: 95vh !important;
                    width: 95vw !important;
                    max-height: 95vh !important;
                }
                .zf-table-wrapper { height: 55vh !important; }
                #zf-helper-body   { max-height: calc(95vh - 50px) !important; }
            }

            /* ===== 头部 ===== */
            #zf-helper-header {
                background: linear-gradient(135deg, #409eff, #096dd9);
                color: #fff;
                padding: 14px 18px;
                font-weight: bold;
                font-size: 16px;
                cursor: move;
                display: flex;
                justify-content: space-between;
                align-items: center;
                touch-action: none;
                -webkit-touch-callout: none;
                user-select: none;
                min-height: 50px;
                flex-shrink: 0;
                border-radius: 12px 12px 0 0;
            }
            #zf-helper-header.zf-dragging-header { cursor: grabbing; background: linear-gradient(135deg, #3080e0, #0850a0); }

            /* ===== 收起/展开按钮 ===== */
            #zf-toggle-btn {
                background: rgba(255,255,255,0.2);
                border: 1px solid rgba(255,255,255,0.4);
                color: white;
                padding: 8px 14px;
                border-radius: 6px;
                font-size: 13px;
                cursor: pointer;
                transition: background 0.2s;
                min-height: 36px;
                min-width: 70px;
                -webkit-appearance: none;
                touch-action: manipulation;
                flex-shrink: 0;
            }
            #zf-toggle-btn:hover, #zf-toggle-btn:active {
                background: rgba(255,255,255,0.35);
                border-color: rgba(255,255,255,0.6);
            }

            /* ===== Body ===== */
            #zf-helper-body {
                display: flex;
                flex-direction: column;
                overflow: hidden;
                flex: 1;
                min-height: 0;
            }

            .zf-content-area {
                padding: 18px;
                overflow-y: auto;
                flex: 1;
                display: flex;
                flex-direction: column;
                min-height: 0;
            }
            .zf-content-area::-webkit-scrollbar { width: 6px; }
            .zf-content-area::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.2); border-radius: 3px; }

            /* ===== 数据面板 ===== */
            .zf-dashboard {
                display: grid;
                grid-template-columns: repeat(6, 1fr);
                gap: 12px;
                margin-bottom: 18px;
                background: #f5f7fa;
                padding: 16px;
                border-radius: 10px;
                border: 1px solid #e4e7ed;
                flex-shrink: 0;
            }
            .zf-stat-item  { text-align: center; padding: 6px; }
            .zf-stat-val   { font-size: 18px; font-weight: bold; color: #1890ff; display: block; margin-bottom: 6px; line-height: 1.2; }
            .zf-stat-label { font-size: 12px; color: #606266; line-height: 1.2; }

            /* ===== 表格 ===== */
            .zf-table-wrapper {
                border: 1px solid #ebeef5;
                overflow-y: auto;
                flex: 1;
                min-height: 200px;
                border-radius: 8px;
                -webkit-overflow-scrolling: touch;
            }
            .zf-table-wrapper::-webkit-scrollbar { width: 6px; }
            .zf-table-wrapper::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.2); border-radius: 3px; }

            .zf-table {
                width: 100%;
                border-collapse: collapse;
                font-size: 14px;
                table-layout: fixed;
            }
            .zf-table th {
                background: #f5f7fa;
                position: sticky;
                top: 0;
                padding: 14px 10px;
                border-bottom: 2px solid #ebeef5;
                text-align: left;
                color: #606266;
                z-index: 1;
                font-weight: 600;
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                cursor: pointer;
                font-size: 14px;
            }
            .zf-table th:hover { background: #eef1f6; color: #409eff; }
            .zf-sort-icon { margin-left: 4px; font-size: 11px; color: #c0c4cc; display: inline-block; }

            .zf-table td {
                padding: 12px 10px;
                border-bottom: 1px solid #ebeef5;
                color: #606266;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                font-size: 14px;
            }
            .zf-table tr { transition: background-color 0.15s; }
            .zf-table tr:hover { background-color: #f0f9eb; cursor: pointer; }
            .zf-table tr:active { background-color: #e6f7d9 !important; }

            /* 移动端列宽 */
            @media only screen and (max-width: 768px) {
                .zf-table { table-layout: auto; min-width: 650px; }
                .zf-table-wrapper { overflow-x: auto; }
                .zf-table th:nth-child(1), .zf-table td:nth-child(1) { min-width: 180px; max-width: 220px; }
                .zf-table th:nth-child(2), .zf-table td:nth-child(2) { min-width: 70px;  max-width: 90px;  }
                .zf-table th:nth-child(3), .zf-table td:nth-child(3) { min-width: 60px;  max-width: 70px;  }
                .zf-table th:nth-child(4), .zf-table td:nth-child(4) { min-width: 70px;  max-width: 90px;  }
                .zf-table th:nth-child(5), .zf-table td:nth-child(5) { min-width: 70px;  max-width: 90px;  }
                .zf-table th:nth-child(6), .zf-table td:nth-child(6) { min-width: 90px;  max-width: 110px; }
            }

            /* 纯触屏设备取消 hover */
            @media (hover: none) and (pointer: coarse) {
                .zf-table tr:hover { background-color: inherit; }
                .zf-table th:hover { background-color: #f5f7fa; color: #606266; }
                .zf-btn:hover { opacity: 1; }
            }

            /* ===== 底部控制区 ===== */
            .zf-controls {
                display: flex;
                justify-content: space-between;
                align-items: center;
                padding: 16px 18px;
                background: #fff;
                border-top: 1px solid #ebeef5;
                flex-wrap: wrap;
                gap: 12px;
                flex-shrink: 0;
                box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
                min-height: 74px;
            }
            .zf-checkboxes {
                display: flex;
                gap: 18px;
                font-size: 14px;
                color: #606266;
                align-items: center;
                flex-wrap: wrap;
            }
            .zf-btns { display: flex; gap: 10px; flex-wrap: wrap; }

            .zf-btn {
                padding: 12px 18px;
                border: none;
                border-radius: 8px;
                cursor: pointer;
                color: #fff;
                font-weight: bold;
                font-size: 14px;
                transition: filter 0.15s, transform 0.1s, box-shadow 0.1s;
                white-space: nowrap;
                min-height: 46px;
                display: flex;
                align-items: center;
                justify-content: center;
                -webkit-appearance: none;
                user-select: none;
                touch-action: manipulation;
                box-shadow: 0 3px 6px rgba(0,0,0,0.1);
            }
            .zf-btn:hover  { filter: brightness(1.08); }
            .zf-btn:active { transform: translateY(1px); box-shadow: 0 1px 3px rgba(0,0,0,0.1); filter: brightness(0.95); }

            .btn-blue   { background: #409eff; }
            .btn-green  { background: #67c23a; }
            .btn-orange { background: #e6a23c; }
            .btn-orange.active { background: #d35400; }

            @media only screen and (max-width: 480px) {
                .zf-controls { padding: 12px 15px; min-height: 65px; }
                .zf-checkboxes { gap: 12px; font-size: 13px; }
                .zf-checkboxes label { display: flex; align-items: center; gap: 5px; }
                .zf-checkboxes input[type="checkbox"] { width: 18px; height: 18px; }
                .zf-btns { gap: 8px; }
                .zf-btn  { padding: 8px 12px; font-size: 12px; min-height: 40px; }
            }

            /* ===== 最小化 ===== */
            #zf-helper-app.zf-minimized {
                width: 280px !important;
                min-width: 280px !important;
                height: auto !important;
                max-height: none !important;
                min-height: auto !important;
                resize: none !important;
                overflow: hidden !important;
            }
            #zf-helper-app.zf-minimized #zf-helper-body  { display: none; }
            #zf-helper-app.zf-minimized #zf-helper-header { padding: 12px 15px; min-height: 44px; border-radius: 12px; }
            #zf-helper-app.zf-minimized #zf-toggle-btn   { min-width: 60px; padding: 6px 12px; font-size: 12px; }
            /* 最小化时隐藏 resize 指示角 */
            #zf-helper-app.zf-minimized::after { display: none; }

            /* resize 指示角（非最小化时显示） */
            #zf-helper-app::after {
                content: '';
                position: absolute;
                bottom: 0;
                right: 0;
                width: 15px;
                height: 15px;
                background: linear-gradient(135deg, transparent 50%, #409eff 50%);
                border-bottom-right-radius: 12px;
                pointer-events: none;
                opacity: 0.4;
                transition: opacity 0.2s;
            }
            #zf-helper-app:hover::after { opacity: 0.75; }

            /* ===== 详情模态框 ===== */
            #zf-detail-modal {
                position: fixed;
                inset: 0;
                background: rgba(0,0,0,0.65);
                z-index: 2147483647;
                display: none;
                justify-content: center;
                align-items: center;
                backdrop-filter: blur(4px);
                -webkit-backdrop-filter: blur(4px);
                padding: 15px;
            }
            .zf-modal-content {
                background: #fff;
                width: 680px;
                max-width: 95vw;
                max-height: 85vh;
                border-radius: 14px;
                box-shadow: 0 20px 60px rgba(0,0,0,0.4);
                display: flex;
                flex-direction: column;
                animation: zf-zoom-in 0.25s cubic-bezier(0.18, 0.89, 0.32, 1.28);
                overflow: hidden;
            }
            @media only screen and (max-width: 480px) {
                .zf-modal-content { max-height: 90vh; border-radius: 10px; width: 95vw; }
            }
            .zf-modal-header {
                padding: 18px 24px;
                border-bottom: 1px solid #ebeef5;
                font-weight: bold;
                display: flex;
                justify-content: space-between;
                align-items: center;
                font-size: 17px;
                background: linear-gradient(135deg, #409eff, #096dd9);
                color: white;
            }
            .zf-modal-close {
                cursor: pointer;
                font-size: 32px;
                color: white;
                padding: 5px 15px;
                min-width: 44px;
                min-height: 44px;
                display: flex;
                align-items: center;
                justify-content: center;
                border-radius: 6px;
                transition: background 0.2s;
            }
            .zf-modal-close:hover, .zf-modal-close:active { background: rgba(255,255,255,0.2); }

            .zf-modal-body {
                padding: 0;
                overflow-y: auto;
                background: #fcfcfc;
                -webkit-overflow-scrolling: touch;
                max-height: calc(85vh - 70px);
            }

            .zf-detail-table { width: 100%; border-collapse: collapse; font-size: 14px; }
            .zf-detail-table tr:nth-child(even) { background: #f9f9f9; }
            .zf-detail-table th {
                background: #f5f7fa;
                padding: 14px 20px;
                border-bottom: 1px solid #ebeef5;
                text-align: right;
                color: #606266;
                white-space: nowrap;
                font-size: 14px;
                font-weight: 600;
                width: 30%;
            }
            .zf-detail-table td {
                padding: 14px 20px;
                border-bottom: 1px solid #ebeef5;
                color: #303133;
                user-select: text;
                -webkit-user-select: text;
                word-break: break-word;
                font-size: 14px;
                width: 70%;
            }
            @media only screen and (max-width: 480px) {
                .zf-detail-table th, .zf-detail-table td { padding: 12px 16px; font-size: 13px; }
                .zf-detail-table th { width: 35%; }
                .zf-detail-table td { width: 65%; }
            }

            @keyframes zf-zoom-in {
                from { opacity: 0; transform: scale(0.95) translateY(-10px); }
                to   { opacity: 1; transform: scale(1)    translateY(0);     }
            }

            /* ===== 全局微调 ===== */
            #zf-helper-app * { -webkit-tap-highlight-color: transparent; }
            /* 防止 iOS 表单缩放 */
            #zf-helper-app input,
            #zf-helper-app select,
            #zf-helper-app textarea { font-size: 16px !important; }
        `
    };

    // ================= 状态管理 =================
    const state = {
        grades: [],
        allRowsData: [],
        headers: [],
        filters: { excludeElective: false, excludeFail: false },
        sort: { key: null, order: 'asc' },
        isMinimized: false,
        isAllSemester: false,
        originalQuery: { xn: '', xq: '' },
        isProcessing: false,
        isInitialized: false,
        isDragging: false,
        isModalOpen: false,
        studentName: '',
        currentSemester: '',
        // 计时器句柄，用于清理
        _scanDebounce: null,
        _pollTimer: null,
        _resizeObserver: null
    };

    // ================= 工具函数 =================

    /** 安全地将字符串转为浮点数，失败返回 0 */
    function parseNum(val) {
        if (!val) return 0;
        const s = parseFloat(String(val).replace(/&nbsp;/g, '').replace(/\s+/g, ''));
        return isNaN(s) ? 0 : s;
    }

    /** 将成绩字符串转为数值（优先数字，其次枚举映射，最后绩点推算） */
    function getScoreValue(scoreStr, point) {
        const clean = String(scoreStr).replace(/&nbsp;/g, '').trim();
        const num = parseFloat(clean);
        if (!isNaN(num)) return num;
        if (CONFIG.SCORE_MAP[clean] !== undefined) return CONFIG.SCORE_MAP[clean];
        const p = parseFloat(point);
        if (!isNaN(p) && p > 0) return Math.round(p * 10 + 50);
        return 0;
    }

    /** 判断课程是否为选修（含"选"或"公"且不含"必"） */
    function isElectiveCourse(natureName) {
        if (!natureName) return false;
        return /[选公]/.test(natureName) && !/必/.test(natureName);
    }

    /** 从页面获取学生姓名与当前学期 */
    function getStudentInfo() {
        const nameEl = document.querySelector('.user-name, .real-name, #realName, [name="realName"]');
        if (nameEl) {
            state.studentName = (nameEl.textContent || nameEl.value || '').trim();
        }

        const xnSel = document.getElementById('xnm');
        const xqSel = document.getElementById('xqm');
        if (xnSel && xqSel) {
            const xnText = xnSel.options[xnSel.selectedIndex]?.text || '';
            const xqText = xqSel.options[xqSel.selectedIndex]?.text || '';
            state.currentSemester = (xnText === '全部' || xqText === '全部')
                ? '全部学期'
                : `${xnText}${xqText}`;
        }
    }

    /** 生成合法的导出文件名 */
    function generateExportFileName() {
        getStudentInfo();
        let name = '';
        if (state.studentName) name += `${state.studentName}_`;
        if (state.currentSemester) name += `${state.currentSemester}_`;
        name += '教务系统成绩明细';
        // 去除文件名非法字符
        return name.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') || 'ECUT成绩分析';
    }

    // ================= 正方系统交互 =================

    /** 将指定 select 元素定位到文本匹配的选项，并触发 change 事件 */
    function setSelectByText(selectEl, targetText) {
        if (!selectEl) return false;
        for (let i = 0; i < selectEl.options.length; i++) {
            if (selectEl.options[i].text.trim() === targetText) {
                if (selectEl.selectedIndex === i) return false; // 已是目标值，无需操作
                selectEl.selectedIndex = i;
                selectEl.dispatchEvent(new Event('change', { bubbles: true }));
                $(selectEl).trigger('change');
                return true;
            }
        }
        console.warn(`[ECUT助手] 未找到下拉选项: ${targetText}`);
        return false;
    }

    /** 强制设置分页条数，优先操作 DOM，其次通过 jqGrid API */
    function forcePageSize(count) {
        // 方式一：直接操作 DOM select
        const sel = document.querySelector('select.ui-pg-selbox');
        if (sel && parseInt(sel.value) !== count) {
            if (![...sel.options].some(o => parseInt(o.value) === count)) {
                const opt = document.createElement('option');
                opt.value = count;
                opt.text = count;
                sel.appendChild(opt);
            }
            sel.value = count;
            sel.dispatchEvent(new Event('change', { bubbles: true }));
            $(sel).trigger('change');
            return true;
        }

        // 方式二：jqGrid API
        try {
            const grid = $('#cjcxlist');
            if (grid.length && grid[0].grid) {
                grid[0].grid.options.rowNum = count;
                grid[0].grid.populate();
                return true;
            }
        } catch (e) {
            /* ignore */
        }

        return false;
    }

    /**
     * 触发原生查询按钮，并在查询完成后自动扫描表格。
     * 依赖浏览器原生事件流，不依赖私有 API。
     */
    function triggerSearch() {
        if (state.isProcessing || state.isModalOpen) return;
        state.isProcessing = true;

        const targetSize = state.isAllSemester ? CONFIG.PAGE_SIZE_ALL : CONFIG.PAGE_SIZE_DEFAULT;
        const searchBtn = document.getElementById('search_go');
        if (!searchBtn) {
            state.isProcessing = false;
            return;
        }

        // 先确保条数正确，再触发查询
        forcePageSize(targetSize);
        setTimeout(() => {
            searchBtn.click();
            // 等待页面渲染完毕后再扫描
            setTimeout(() => {
                forcePageSize(targetSize); // 二次确认条数（某些系统会重置）
                if (scanTable()) update();
                state.isProcessing = false;
            }, 1800);
        }, 200);
    }

    // ================= 核心：扫描与计算 =================

    /** 扫描 jqGrid 表格，提取成绩数据到 state */
    function scanTable() {
        const headerEl = document.querySelector('.ui-jqgrid-htable');
        const bodyEl   = document.querySelector('#tabGrid');
        if (!headerEl || !bodyEl) return false;

        const rows = $(bodyEl).find('tr.jqgrow');
        if (!rows.length) return false;

        // 解析表头，定位关键列索引
        const headers = [];
        const col = { name: -1, nature: -1, credit: -1, score: -1, point: -1, creditPoint: -1 };

        $(headerEl).find('th').each(function (idx) {
            const txt = $(this).text().trim();
            headers.push(txt);
            if (txt.includes('课程名称'))                               col.name        = idx;
            if (txt === '课程性质' || txt === '性质')                   col.nature      = idx;
            if (txt.includes('学分') && !txt.includes('绩点'))          col.credit      = idx;
            if (txt === '成绩' || (col.score === -1 && txt.includes('成绩') && !txt.includes('备注'))) col.score = idx;
            if (txt === '绩点')                                         col.point       = idx;
            if (txt.includes('学分绩点'))                               col.creditPoint = idx;
        });

        if (col.score === -1) return false;

        state.headers     = headers;
        state.allRowsData = [];
        const gradeData   = [];

        rows.each(function (index) {
            const rowRaw = [];
            $(this).find('td').each(function () {
                rowRaw.push($(this).text().replace(/\s+/g, ' ').replace(/&nbsp;/g, '').trim());
            });
            state.allRowsData.push(rowRaw);

            const name = col.name !== -1 ? rowRaw[col.name] : '';
            if (!name) return;

            const nature      = col.nature      !== -1 ? rowRaw[col.nature]      : '必修';
            const credit      = parseNum(rowRaw[col.credit]);
            const scoreRaw    = rowRaw[col.score];
            const point       = parseNum(rowRaw[col.point]);
            const creditPoint = col.creditPoint !== -1
                ? parseNum(rowRaw[col.creditPoint])
                : parseFloat((credit * point).toFixed(2));
            const scoreVal    = getScoreValue(scoreRaw, point);

            gradeData.push({ index, name, nature, credit, point, creditPoint, scoreRaw, scoreVal });
        });

        state.grades = gradeData;
        return true;
    }

    /** 根据当前过滤器计算统计数据，返回结果对象或 null */
    function calculate() {
        if (!state.grades.length) return null;

        const res = { count: 0, totalCredit: 0, totalCreditPoint: 0, avgGPA: 0, avgArith: 0, avgWeight: 0, failCount: 0 };
        let sumScore = 0, sumScoreCredit = 0;

        state.grades.forEach(g => {
            if (state.filters.excludeElective && isElectiveCourse(g.nature)) return;
            if (state.filters.excludeFail     && g.point < 1.0)              return;
            res.count++;
            res.totalCredit      += g.credit;
            res.totalCreditPoint += g.creditPoint;
            sumScore             += g.scoreVal;
            sumScoreCredit       += g.scoreVal * g.credit;
            if (g.point < 1.0) res.failCount++;
        });

        if (res.totalCredit > 0) {
            res.avgGPA    = (res.totalCreditPoint / res.totalCredit).toFixed(4);
            res.avgWeight = (sumScoreCredit        / res.totalCredit).toFixed(2);
        }
        if (res.count > 0) {
            res.avgArith = (sumScore / res.count).toFixed(2);
        }
        res.totalCredit      = parseFloat(res.totalCredit.toFixed(1));
        res.totalCreditPoint = parseFloat(res.totalCreditPoint.toFixed(2));
        return res;
    }

    // ================= UI 渲染 =================

    function createUI() {
        if ($('#zf-helper-app').length) return;
        GM_addStyle(CONFIG.CSS);

        const savedPos  = GM_getValue('win_pos_v4',  null);
        const savedSize = GM_getValue('win_size_v4', null);

        const W = savedSize
            ? Math.max(350, Math.min(savedSize.width,  window.innerWidth  * 0.95))
            : 850;
        const H = savedSize
            ? Math.max(200, Math.min(savedSize.height, window.innerHeight * 0.9))
            : 600;

        const posStyle = savedPos
            ? `top:${savedPos.top}px; left:${savedPos.left}px;`
            : 'top:80px; right:20px;';

        $('body').append(`
            <div id="zf-helper-app" style="${posStyle} width:${W}px; height:${H}px;">
                <div id="zf-helper-header">
                    <span>📊 ECUT 成绩助手 v4.1</span>
                    <button id="zf-toggle-btn">收起</button>
                </div>
                <div id="zf-helper-body">
                    <div class="zf-content-area">
                        <div class="zf-dashboard">
                            <div class="zf-stat-item"><span class="zf-stat-val" id="d-gpa">-</span><span class="zf-stat-label">平均学分绩点</span></div>
                            <div class="zf-stat-item"><span class="zf-stat-val" id="d-credit">-</span><span class="zf-stat-label">总学分</span></div>
                            <div class="zf-stat-item"><span class="zf-stat-val" id="d-cp">-</span><span class="zf-stat-label">总学分绩点</span></div>
                            <div class="zf-stat-item"><span class="zf-stat-val" id="d-arith">-</span><span class="zf-stat-label">算术平均分</span></div>
                            <div class="zf-stat-item"><span class="zf-stat-val" id="d-weight">-</span><span class="zf-stat-label">加权平均分</span></div>
                            <div class="zf-stat-item"><span class="zf-stat-val" id="d-fail" style="color:red">-</span><span class="zf-stat-label">挂科数</span></div>
                        </div>
                        <div class="zf-table-wrapper">
                            <table class="zf-table">
                                <thead>
                                    <tr>
                                        <th data-sort="name">课程名称<span class="zf-sort-icon"></span></th>
                                        <th data-sort="nature">性质<span class="zf-sort-icon"></span></th>
                                        <th data-sort="credit">学分<span class="zf-sort-icon"></span></th>
                                        <th data-sort="scoreVal">成绩<span class="zf-sort-icon"></span></th>
                                        <th data-sort="point">绩点<span class="zf-sort-icon"></span></th>
                                        <th data-sort="creditPoint">学分绩点<span class="zf-sort-icon"></span></th>
                                    </tr>
                                </thead>
                                <tbody id="zf-tbody">
                                    <tr><td colspan="6" style="text-align:center;padding:20px;color:#999">等待数据加载...</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                    <div class="zf-controls">
                        <div class="zf-checkboxes">
                            <label><input type="checkbox" id="chk-ele"> 剔除选修</label>
                            <label><input type="checkbox" id="chk-fail"> 剔除挂科</label>
                        </div>
                        <div class="zf-btns">
                            <button class="zf-btn btn-orange" id="btn-all-term">📅 全部学期</button>
                            <button class="zf-btn btn-blue"   id="btn-run">🚀 重新分析</button>
                            <button class="zf-btn btn-green"  id="btn-export">📥 导出Excel</button>
                        </div>
                    </div>
                </div>
            </div>
            <div id="zf-detail-modal">
                <div class="zf-modal-content">
                    <div class="zf-modal-header">
                        <span>📝 课程详情</span>
                        <span class="zf-modal-close" id="zf-modal-close">×</span>
                    </div>
                    <div class="zf-modal-body" id="zf-detail-content"></div>
                </div>
            </div>
        `);

        bindEvents();
        initDraggable();
        initResizeObserver();
        state.isInitialized = true;

        // 尝试立即分析，若无数据则触发查询
        setTimeout(() => {
            getStudentInfo();
            if (scanTable()) {
                update();
            } else {
                forcePageSize(CONFIG.PAGE_SIZE_DEFAULT);
                setTimeout(triggerSearch, 600);
            }
        }, 500);
    }

    function renderTable() {
        const tbody = $('#zf-tbody');
        tbody.empty();

        if (!state.grades.length) {
            tbody.append('<tr><td colspan="6" style="text-align:center;padding:20px;color:#999">等待数据加载...</td></tr>');
            return;
        }

        let displayData = state.grades.filter(g => {
            if (state.filters.excludeElective && isElectiveCourse(g.nature)) return false;
            if (state.filters.excludeFail     && g.point < 1.0)              return false;
            return true;
        });

        if (!displayData.length) {
            tbody.append('<tr><td colspan="6" style="text-align:center;padding:20px;color:#999">没有找到匹配的课程</td></tr>');
            return;
        }

        if (state.sort.key) {
            displayData.sort((a, b) => {
                let vA = a[state.sort.key] ?? '';
                let vB = b[state.sort.key] ?? '';
                if (typeof vA === 'string') {
                    return state.sort.order === 'asc' ? vA.localeCompare(vB) : vB.localeCompare(vA);
                }
                return state.sort.order === 'asc' ? (vA > vB ? 1 : -1) : (vA < vB ? 1 : -1);
            });
        }

        const frag = document.createDocumentFragment();
        displayData.forEach(g => {
            const tr = document.createElement('tr');
            tr.dataset.idx = g.index;
            tr.innerHTML = `
                <td>${g.name        || ''}</td>
                <td>${g.nature      || ''}</td>
                <td>${g.credit      || ''}</td>
                <td>${g.scoreRaw    || ''}</td>
                <td>${g.point       || ''}</td>
                <td>${g.creditPoint || ''}</td>
            `;
            frag.appendChild(tr);
        });
        tbody[0].appendChild(frag);

        // 更新排序图标
        $('.zf-sort-icon').text('');
        if (state.sort.key) {
            $(`th[data-sort="${state.sort.key}"] .zf-sort-icon`).text(state.sort.order === 'asc' ? ' ▲' : ' ▼');
        }
    }

    function update() {
        const res = calculate();
        if (!res) {
            ['d-gpa','d-credit','d-cp','d-arith','d-weight','d-fail'].forEach(id => $(`#${id}`).text('-'));
        } else {
            $('#d-gpa').text(res.avgGPA);
            $('#d-credit').text(res.totalCredit);
            $('#d-cp').text(res.totalCreditPoint);
            $('#d-arith').text(res.avgArith);
            $('#d-weight').text(res.avgWeight);
            $('#d-fail').text(res.failCount);
        }
        renderTable();
    }

    // ================= 事件绑定 =================

    function bindEvents() {
        // 重新分析
        $(document).on('click', '#btn-run', function (e) {
            e.stopPropagation();
            if (scanTable()) update();
            else triggerSearch();
        });

        // 全部学期 / 恢复当前
        $(document).on('click', '#btn-all-term', function (e) {
            e.stopPropagation();
            e.preventDefault();
            if (state.isProcessing || state.isModalOpen) return;

            const xnSel = document.getElementById('xnm');
            const xqSel = document.getElementById('xqm');
            if (!xnSel || !xqSel) return;

            state.isAllSemester = !state.isAllSemester;

            if (state.isAllSemester) {
                state.originalQuery.xn = xnSel.options[xnSel.selectedIndex].text;
                state.originalQuery.xq = xqSel.options[xqSel.selectedIndex].text;
                setSelectByText(xnSel, '全部');
                setSelectByText(xqSel, '全部');
                $(this).addClass('active').text('📅 恢复当前');
            } else {
                setSelectByText(xnSel, state.originalQuery.xn);
                setSelectByText(xqSel, state.originalQuery.xq);
                $(this).removeClass('active').text('📅 全部学期');
            }

            setTimeout(() => {
                forcePageSize(state.isAllSemester ? CONFIG.PAGE_SIZE_ALL : CONFIG.PAGE_SIZE_DEFAULT);
                setTimeout(triggerSearch, 500);
            }, 300);
        });

        // 导出 Excel
        $(document).on('click', '#btn-export', function (e) {
            e.stopPropagation();
            if (!state.allRowsData.length) { alert('暂无数据可导出'); return; }
            const fileName = generateExportFileName();
            const ws = XLSX.utils.aoa_to_sheet([state.headers, ...state.allRowsData]);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, '成绩明细');
            XLSX.writeFile(wb, `${fileName}.xlsx`);
        });

        // 筛选复选框
        $(document).on('change', '#chk-ele, #chk-fail', function () {
            state.filters.excludeElective = $('#chk-ele').prop('checked');
            state.filters.excludeFail     = $('#chk-fail').prop('checked');
            update();
        });

        // 收起 / 展开
        $(document).on('click', '#zf-toggle-btn', function (e) {
            e.stopPropagation();
            e.preventDefault();
            const app = $('#zf-helper-app');
            state.isMinimized = !state.isMinimized;

            if (state.isMinimized) {
                const rect = app[0].getBoundingClientRect();
                GM_setValue('win_size_v4', { width: Math.round(rect.width), height: Math.round(rect.height) });
                app.addClass('zf-minimized');
                $(this).text('展开');
            } else {
                app.removeClass('zf-minimized');
                const saved = GM_getValue('win_size_v4', null);
                if (saved) {
                    app.css({
                        width:  Math.max(350, Math.min(saved.width,  window.innerWidth  * 0.95)) + 'px',
                        height: Math.max(200, Math.min(saved.height, window.innerHeight * 0.9))  + 'px'
                    });
                }
                $(this).text('收起');
            }
        });

        // 表头排序
        $(document).on('click', '.zf-table th[data-sort]', function (e) {
            e.stopPropagation();
            const key = $(this).data('sort');
            if (state.sort.key === key) {
                state.sort.order = state.sort.order === 'asc' ? 'desc' : 'asc';
            } else {
                state.sort.key   = key;
                state.sort.order = 'asc';
            }
            renderTable();
        });

        // 行点击 → 详情弹窗
        $(document).on('click', '#zf-tbody tr', function (e) {
            e.stopPropagation();
            const idx = parseInt($(this).data('idx'));
            if (!state.allRowsData[idx]) return;
            let html = '<table class="zf-detail-table"><tbody>';
            state.headers.forEach((h, i) => {
                html += `<tr><th>${h}</th><td>${state.allRowsData[idx][i] || '-'}</td></tr>`;
            });
            html += '</tbody></table>';
            $('#zf-detail-content').html(html);
            $('#zf-detail-modal').css('display', 'flex');
            state.isModalOpen = true;
        });

        // 关闭弹窗（按钮 + 背景 + ESC）
        function closeModal() {
            $('#zf-detail-modal').hide();
            state.isModalOpen = false;
        }
        $(document).on('click', '#zf-modal-close', function (e) { e.stopPropagation(); closeModal(); });
        $(document).on('click', '#zf-detail-modal', function (e) {
            if (e.target.id === 'zf-detail-modal') closeModal();
        });
        $(document).on('keydown', function (e) {
            if (e.key === 'Escape' && state.isModalOpen) closeModal();
        });
    }

    // ================= 拖拽（iPhone 优化） =================

    function initDraggable() {
        const el = document.getElementById('zf-helper-app');
        const hd = document.getElementById('zf-helper-header');
        if (!el || !hd) return;

        const isTouchDevice = 'ontouchstart' in window;

        let dragging = false;
        let startX, startY, originLeft, originTop, rafId;
        let curX, curY;

        const loop = () => {
            if (!dragging) return;
            const dx = curX - startX;
            const dy = curY - startY;
            const newLeft = Math.max(0, Math.min(window.innerWidth  - el.offsetWidth,  originLeft + dx));
            const newTop  = Math.max(0, Math.min(window.innerHeight - el.offsetHeight, originTop  + dy));
            el.style.left  = `${newLeft}px`;
            el.style.top   = `${newTop}px`;
            el.style.right = 'auto';
            rafId = requestAnimationFrame(loop);
        };

        const onStart = (e) => {
            // 按钮点击不触发拖拽
            if (e.target.closest && e.target.closest('button')) return;
            if (state.isModalOpen) return;
            if (e.cancelable) e.preventDefault();
            e.stopPropagation();

            dragging = true;
            state.isDragging = true;
            el.classList.add('zf-dragging');
            hd.classList.add('zf-dragging-header');
            el.style.transition = 'none';

            const pt   = e.touches ? e.touches[0] : e;
            startX = curX = pt.clientX;
            startY = curY = pt.clientY;
            const rect = el.getBoundingClientRect();
            originLeft = rect.left;
            originTop  = rect.top;

            rafId = requestAnimationFrame(loop);
        };

        const onMove = (e) => {
            if (!dragging) return;
            if (e.cancelable) e.preventDefault();
            const pt = e.touches ? e.touches[0] : e;
            curX = pt.clientX;
            curY = pt.clientY;
        };

        const onEnd = () => {
            if (!dragging) return;
            dragging = false;
            cancelAnimationFrame(rafId);
            el.classList.remove('zf-dragging');
            hd.classList.remove('zf-dragging-header');
            el.style.transition = '';

            const rect = el.getBoundingClientRect();
            GM_setValue('win_pos_v4', { top: Math.round(rect.top), left: Math.round(rect.left) });

            // 短暂延迟后解除标志，避免拖拽结束瞬间触发行点击
            setTimeout(() => { state.isDragging = false; }, 80);
        };

        // 鼠标
        hd.addEventListener('mousedown', onStart);
        window.addEventListener('mousemove',  onMove, { passive: false });
        window.addEventListener('mouseup',    onEnd);

        // 触摸
        hd.addEventListener('touchstart', onStart, { passive: false });
        // 最小化时整个窗口可拖
        if (isTouchDevice) {
            el.addEventListener('touchstart', (e) => {
                if (state.isMinimized) onStart(e);
            }, { passive: false });
        }
        window.addEventListener('touchmove',   onMove, { passive: false });
        window.addEventListener('touchend',    onEnd);
        window.addEventListener('touchcancel', onEnd);
    }

    // ================= 窗口大小持久化 =================

    function initResizeObserver() {
        const el = document.getElementById('zf-helper-app');
        if (!el || typeof ResizeObserver === 'undefined') return;

        let timeout;
        state._resizeObserver = new ResizeObserver(entries => {
            if (state.isMinimized) return;
            for (const entry of entries) {
                const { width, height } = entry.contentRect;
                clearTimeout(timeout);
                timeout = setTimeout(() => {
                    GM_setValue('win_size_v4', { width: Math.round(width), height: Math.round(height) });
                }, CONFIG.DEBOUNCE_RESIZE);
            }
        });
        state._resizeObserver.observe(el);
    }

    // ================= 自动监控：MutationObserver 驱动 =================

    /**
     * 仅在 MutationObserver 检测到 jqGrid 行变化时触发扫描。
     * 降频轮询作为备用保底，间隔设为 5s（仅页面无变化时触发）。
     */
    function startMonitor() {
        let lastRowCount = -1;

        const doScan = () => {
            if (state.isModalOpen || state.isProcessing) return;
            const rows = document.querySelectorAll('#tabGrid tr.jqgrow');
            if (rows.length && rows.length !== lastRowCount) {
                lastRowCount = rows.length;
                if (scanTable()) update();
            }
        };

        // MutationObserver：监听 jqGrid 行的增删
        const observer = new MutationObserver(mutations => {
            let changed = false;
            for (const m of mutations) {
                if (m.type !== 'childList') continue;
                for (const node of [...m.addedNodes, ...m.removedNodes]) {
                    if (node.nodeType === 1 &&
                        (node.classList?.contains('jqgrow') || node.querySelector?.('.jqgrow'))) {
                        changed = true;
                        break;
                    }
                }
                if (changed) break;
            }
            if (changed) {
                clearTimeout(state._scanDebounce);
                state._scanDebounce = setTimeout(doScan, CONFIG.DEBOUNCE_SCAN);
            }
        });

        observer.observe(document.body, { childList: true, subtree: true });

        // 备用低频轮询（页面静止时兜底）
        state._pollTimer = setInterval(() => {
            if (!state.isModalOpen && !state.isProcessing) doScan();
        }, CONFIG.POLL_INTERVAL);
    }

    // ================= 启动 =================

    let attempts = 0;
    const bootInterval = setInterval(() => {
        if (!document.body) return;
        createUI();
        if (document.getElementById('zf-helper-app')) {
            clearInterval(bootInterval);
            // 确保初始条数正确
            setTimeout(() => {
                forcePageSize(CONFIG.PAGE_SIZE_DEFAULT);
                if (!$('#tabGrid tr.jqgrow').length) {
                    setTimeout(triggerSearch, 600);
                }
            }, 800);
            startMonitor();
        } else if (++attempts >= 10) {
            clearInterval(bootInterval);
            console.error('[ECUT助手] 初始化失败，请刷新页面后重试');
        }
    }, 800);

})();
