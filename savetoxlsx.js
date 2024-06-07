// ==UserScript==
// @name         Appstoreconnect一键下载为xlsx付款和财务报告
// @namespace    https://github.com/JoanLeeo/appstoreconnect_payments_and_financial_reports_savetoxlsx
// @version      0.1
// @description  Appstoreconnect 一键下载付款报告为xlsx格式文件
// @author       mr.wendao
// @match        https://appstoreconnect.apple.com/itc/payments_and_financial_reports
// @require      https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.3.0/papaparse.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.3/xlsx.full.min.js
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    // 添加按钮
    function addButton() {
        var button = document.createElement('button');
        button.innerHTML = 'Download as XLSX';
        button.style.position = "fixed";
        button.style.top = "100px";
        button.style.right = "100px";
        button.style.zIndex = "99";
        button.addEventListener('click', downloadAndConvert);
        document.body.appendChild(button);
    }

    // 下载
    async function downloadAndConvert() {
        var link = document.querySelector('.rightside .actions.flexdist.linkable a');
        if (!link) return;

        var url = link.href;
        var response = await fetch(url);
        var csvData = await response.text();

        // 获取下载文件名字
        var urlParams = new URLSearchParams(link.search);
        var year = urlParams.get('year');
        var month = urlParams.get('month');
        // 保存文件名字
        var fileName = year + month + '.xlsx';

        convertCSVtoXLSX(csvData, fileName);
    }
    /// csv 转 xslx
    function convertCSVtoXLSX(csvData, fileName) {
        const parsedData = Papa.parse(csvData, { skipEmptyLines: false });
        if (parsedData.errors.length > 0) {
            console.error('Error parsing CSV:', parsedData.errors);
            return;
        }

        const worksheet = XLSX.utils.aoa_to_sheet(parsedData.data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
        const xlsxFile = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([xlsxFile], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        a.click();
        URL.revokeObjectURL(url);
    }
    // 添加按钮
    addButton();
})();
