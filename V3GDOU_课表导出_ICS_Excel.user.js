// ==UserScript==
// @name        广东海洋大学-正方教务系统课程表导出(日历+Excel+课刻+WakeUp)
// @namespace   http://tampermonkey.net/
// @version      1.6.0
// @description  适配广东海洋大学正方教务系统：课表页面一键导出 ICS 日历、普通 Excel、课刻专用 xlsx、WakeUp 课程表App专用 csv。
// @author      feldan / 31415926535x / 权爷
// @compatible  chrome
// @compatible  firefox
// @license     MIT
// @include     *://jw.gdou.edu.cn/*
// @include     *://wfw.gdou.edu.cn/*
// @run-at      document-start
// ==/UserScript==

var ClassScheduleToICSURL = "kbcx/xskbcx_cxXskbcxIndex.html";

var pageWin = window;
try{ if(typeof unsafeWindow !== "undefined" && unsafeWindow) pageWin = unsafeWindow; }catch(e){}

(function (){
    'use strict';
    var booted = false;
    function boot(){
        if(booted) return;
        booted = true;
        try{
            main();
        }catch(err){
            alert("课表导出脚本运行出错：" + (err && err.message ? err.message : err));
        }
    }
    if(document.readyState !== "loading"){
        boot();
    }else{
        try{ pageWin.addEventListener("load", boot); }catch(e){ boot(); }
        document.addEventListener("DOMContentLoaded", boot);
        setTimeout(boot, 5000);
    }
})();

function isKbPage(url){
    return url.indexOf(ClassScheduleToICSURL) != -1 || url.indexOf("xskbcx") != -1;
}

function main(){
    if(isKbPage(window.location.href)) ClassScheduleToICS();
}

function ClassScheduleToICS(){
    pageFullyLoaded();

    function pageFullyLoaded(){
        let div = document.getElementsByClassName("btn-toolbar pull-right")[0]
                  || document.getElementsByClassName("btn-toolbar")[0];
        if(!div || (div.offsetWidth === 0 && div.offsetHeight === 0)){
            div = document.createElement("div");
            div.className = "btn-toolbar pull-right";
            div.style.margin = "8px 0";
            let anchors = [document.querySelector(".panel-body"), document.querySelector(".content"), document.getElementById("content")];
            let anchor = null;
            for(let a of anchors){
                if(a && a.offsetWidth > 0){ anchor = a; break; }
            }
            if(!anchor) anchor = document.body;
            anchor.appendChild(div);
        }

        let btn = document.createElement("button");
        btn.className = "btn btn-default";
        btn.id = "exportbtn";
        let sp = document.createElement("span");
        sp.innerText = "导出ICS";
        sp.className = "bigger-120 glyphicon glyphicon-calendar";
        btn.append(sp);

        let xbtn = document.createElement("button");
        xbtn.className = "btn btn-default";
        xbtn.id = "exportexcelbtn";
        sp = document.createElement("span");
        sp.innerText = "导出Excel";
        sp.className = "bigger-120 glyphicon glyphicon-download-alt";
        xbtn.append(sp);

        let wbtn = document.createElement("button");
        wbtn.className = "btn btn-default";
        wbtn.id = "exportwidgetbtn";
        sp = document.createElement("span");
        sp.innerText = "导出课刻";
        sp.className = "bigger-120 glyphicon glyphicon-th";
        wbtn.append(sp);

        let ubtn = document.createElement("button");
        ubtn.className = "btn btn-default";
        ubtn.id = "exportwakeupbtn";
        sp = document.createElement("span");
        sp.innerText = "导出WakeUp";
        sp.className = "bigger-120 glyphicon glyphicon-phone";
        ubtn.append(sp);

        let dwnbtn = document.createElement("button");
        dwnbtn.className = "btn btn-default";
        sp = document.createElement("span");
        sp.innerText = "选择本学期第一个星期一:     ";
        sp.className = "bigger-120 glyphicon glyphicon-time";
        dwnbtn.appendChild(sp);
        let StartDate = document.createElement("input");
        StartDate.type = "date";
        StartDate.value = DEFAULT_FIRST_MONDAY;
        dwnbtn.appendChild(StartDate);

        div.appendChild(btn);
        div.appendChild(xbtn);
        div.appendChild(wbtn);
        div.appendChild(ubtn);
        div.appendChild(dwnbtn);

        btn.onclick = function(){
            startDate = StartDate.value;
            generateCalendar(parseCourses(parseTable()));
            alert("ICS日历文件已导出，可直接导入系统日历。");
        };
        xbtn.onclick = function(){
            startDate = StartDate.value;
            exportToExcel(parseCourses(parseTable()));
        };
        wbtn.onclick = function(){
            startDate = StartDate.value;
            exportToWidget(parseCourses(parseTable()));
        };
        ubtn.onclick = function(){
            startDate = StartDate.value;
            exportToWakeUp(parseCourses(parseTable()));
        };
    }

    var startDate;

    var DEFAULT_FIRST_MONDAY = "2026-08-31";

    var Week;
    (function (Week) {
        Week[Week["Monday"] = 1] = "Monday";
        Week[Week["TuesDay"] = 2] = "TuesDay";
        Week[Week["Wednesday"] = 3] = "Wednesday";
        Week[Week["ThursDay"] = 4] = "ThursDay";
        Week[Week["Friday"] = 5] = "Friday";
        Week[Week["Saturday"] = 6] = "Saturday";
        Week[Week["Sunday"] = 7] = "Sunday";
    })(Week || (Week = {}));

    function parseTable(){
        let week = new Array();
        let divs = new Array();

        let table = document.getElementById("kbgrid_table_0");
        if(!table){
            let tables = Array.from(document.querySelectorAll("table"));
            for(let t of tables){
                if(/星期[一二三四五六日]/.test(t.textContent || "")){
                    table = t;
                    break;
                }
            }
        }
        if(!table){
            alert("未找到课表表格，请确认当前页面为「学生课表查询」页面并已加载课表。");
            return {week: [], divs: []};
        }

        let tds = table.querySelectorAll("td");
        let idTds = 0;
        tds.forEach(element => {
            if(element.hasAttribute("id")){
                idTds++;
                if(element.hasChildNodes()){
                    let div = Array.from(element.getElementsByTagName("div"));
                    divs = divs.concat(div);
                    let wk = Week[element.getAttribute("id")[0]];
                    for(let i = 0; i < div.length; ++i){
                        week.push(wk);
                    }
                }
            }
        });

        if(week.length === 0){
            let rows = Array.from(table.querySelectorAll("tr"));
            let headerRow = null;
            for(let r of rows){
                let t = (r.textContent || "").trim();
                if(t.indexOf("星期一") !== -1 || /星期[一二三四五六日]/.test(t)){
                    headerRow = r;
                    break;
                }
            }
            if(headerRow){
                let cells = Array.from(headerRow.querySelectorAll("th, td"));
                let colWeek = {};
                cells.forEach((c, idx) => {
                    let t = (c.textContent || "").trim();
                    for(let w = 1; w <= 7; w++){
                        if(t === "星期" + "一二三四五六日"[w - 1]) colWeek[idx] = w;
                    }
                });
                rows.forEach(row => {
                    if(row === headerRow) return;
                    let c2 = Array.from(row.querySelectorAll("td, th"));
                    if(!c2.length) return;
                    let first = (c2[0].textContent || "").trim();
                    let secMatch = first.match(/第?\s*(\d{1,2})\s*节?/);
                    if(!secMatch) return;
                    let section = parseInt(secMatch[1]);
                    if(!(section >= 1 && section <= 24)) return;
                    for(let w = 1; w <= 7; w++){
                        let c = c2[colWeek[w]];
                        if(!c) continue;
                        let dv = Array.from(c.getElementsByTagName("div"));
                        if(dv.length){
                            dv.forEach(d => { divs.push(d); week.push(w); });
                        }
                        else if((c.textContent || "").trim() !== ""){
                            divs.push(c);
                            week.push(w);
                        }
                    }
                });
            }
        }

        return {week: week, divs: divs};
    }

    class Course{
        constructor(course){
            if(course){
                this.name = course.name;
                this.week = course.week;
                this.info = course.info;
                this.startTime = course.startTime;
                this.endTime = course.endTime;
                this.weekSegments = course.weekSegments;
                this.startWeek = course.startWeek;
                this.endWeek = course.endWeek;
                this.isSingleOrDouble = course.isSingleOrDouble;
                this.location = course.location;
                this.teacher = course.teacher;
            }
        }
    }

    function extractCourseName(div, depth){
        depth = depth || 0;
        if(depth > 1) return "";
        let name = "";
        let candidates = [];
        try{
            let spans = div.getElementsByTagName("span");
            if(spans.length){
                let fonts = spans[0].getElementsByTagName("font");
                if(fonts.length) candidates.push(fonts[0].innerText);
            }
        }catch(e){}
        try{
            let us = div.getElementsByTagName('u');
            if(us.length) candidates.push(us[0].innerText);
        }catch(e){}
        try{
            if(div.children.length) candidates.push(div.children[0].innerText);
        }catch(e){}
        try{
            if(div.innerText) candidates.push(div.innerText.split('\n\n')[0]);
        }catch(e){}
        for(let c of candidates){
            if(c && c.trim()){ name = c; break; }
        }
        if(!name){
            let td = div.closest ? div.closest('td') : null;
            if(td){
                let divs = td.querySelectorAll('div');
                for(let j = 0; j < divs.length; ++j){
                    if(divs[j] === div) continue;
                    let n = extractCourseName(divs[j], depth + 1);
                    if(n){ name = n; break; }
                }
            }
        }
        name = (name || "").replace(/^\s+/, '').replace(/\s+$/, '');
        name = name.replace(/☆/g, "").replace(/★/g, "").replace(/【调】/g, "").replace(/（调）/g, "");
        return name;
    }

    function parseWeekSegments(fullInfo){
        let res = [];
        if(!fullInfo) return res;
        let text = String(fullInfo);
        let regex = /第?\s*\d+(?:\s*[-—~]\s*\d+)?\s*周\s*(?:[\(（]\s*([单双])\s*[\)）])?/g;
        let match;
        let seen = {};
        while((match = regex.exec(text)) !== null){
            let seg = match[0];
            let core = seg.replace(/^\s*第\s*/, '').trim();
            let parity = match[1] === '单' ? 'odd' : (match[1] === '双' ? 'even' : 'all');
            let mRange = core.match(/^(\d+)\s*[-—~]\s*(\d+)\s*周/);
            if(mRange){
                let s = parseInt(mRange[1]), e = parseInt(mRange[2]);
                if(s >= 1 && e >= s && e <= 30){
                    let key = s + '-' + e + '-' + parity;
                    if(!seen[key]){ res.push({start:s, end:e, parity:parity}); seen[key] = 1; }
                }
                continue;
            }
            let mOne = core.match(/^(\d+)\s*周/);
            if(mOne){
                let s = parseInt(mOne[1]);
                if(s >= 1 && s <= 30){
                    let key = s + '-' + s + '-' + parity;
                    if(!seen[key]){ res.push({start:s, end:s, parity:parity}); seen[key] = 1; }
                }
            }
        }
        return res;
    }

    function isJunkName(s){
        return /^(重修|辅修|选修|必修|补修|线上|网课|慕课|停课|调课|考试|考查|考核|未安排|面授|讲授|讲课|学分|合计|公共|专业|实践|理论|备注)$/.test(s);
    }

    function washName(raw){
        if(!raw) return "";
        let t = String(raw).split(/考查|考核|考试|未安排|面授|讲授|讲课/)[0];
        t = t.replace(/[（(【\[][^）)】\]]*[）)】\]]/g, " ");
        t = t.replace(/(副教授|教授|讲师|助教|高级工程师|工程师|实验师|研究员|副研究员|外聘|兼职|专职|返聘|博士|硕士|老师|教师)/g, " ");
        let out = [];
        t.split(/[,，、\/;；\s]+/).forEach(function(p){
            p = p.trim().replace(/^无+/, "").replace(/无+$/, "").trim();
            if(!p) return;
            if((/^[一-龥·]{2,10}$/.test(p) || /^[A-Za-z][A-Za-z.\-]{1,20}$/.test(p)) && !isJunkName(p) && out.indexOf(p) < 0){
                out.push(p);
            }
        });
        return out.join(",");
    }

    function isRoleWord(s){
        return /^(副教授|教授|讲师|助教|外聘|兼职|专职|返聘|工程师|实验师|研究员|副研究员|博士|硕士|老师|教师)$/.test(s);
    }

    function parseTeacher(wholeText, courseName, rawText){
        let mTea = wholeText.match(/教师[:：]\s*([^，,\n]+)/);
        if(mTea){
            let n = washName(mTea[1]);
            if(n) return n;
        }
        let text2 = wholeText.replace(/[（(【\[][^）)】\]]*[）)】\]]/g, " ");
        let re = /([一-龥·]{2,10}(?:[,，、][一-龥·]{2,10})*)\s*无?\s*(?=考查|考核|考试)/g;
        let m, names = [];
        while((m = re.exec(text2)) !== null){
            let got = washName(m[1]);
            if(!got) continue;
            got.split(",").forEach(function(x){ if(x != courseName && names.indexOf(x) < 0) names.push(x); });
        }
        if(names.length < 2){
            let anchors = /考查|考核|考试/g, mm;
            while((mm = anchors.exec(wholeText)) !== null){
                let before = wholeText.slice(Math.max(0, mm.index - 60), mm.index);
                let tokens = before.split(/[\s,，、\/;；]+/).filter(Boolean);
                for(let i = tokens.length - 1; i >= 0; i--){
                    if(tokens[i] === courseName) break;
                    if(/(教学楼|实验楼|楼\s*\d|教室|球场|篮球场|体育馆|校区|中心|实验室)/.test(tokens[i])) break;
                    let n = washName(tokens[i]);
                    if(!n){
                        if(isRoleWord(tokens[i])) continue;
                        break;
                    }
                    n.split(",").forEach(function(x){ if(x != courseName && names.indexOf(x) < 0) names.push(x); });
                    if(names.length >= 4) break;
                }
                if(names.length) break;
            }
        }
        if(!names.length && rawText){
            let norm = String(courseName).replace(/[*$#☆★\s]+/g, "");
            let lines = String(rawText).split(/\r?\n/).map(function(s){ return s.trim(); }).filter(Boolean);
            for(let i = 0; i < lines.length; i++){
                let s = lines[i];
                if(s.length < 2 || s.length > 12) continue;
                if(/[\d（()）]/.test(s)) continue;                       // 节次/周次/学分/带括号的备注
                if(/^(考查|考核|考试|未安排|面授|讲授|讲课|学分|教师|暂无|重修|辅修|选修|必修|补修|线上|网课|慕课|停课|调课)/.test(s)) continue;
                if(/(楼|教室|球场|体育馆|校区|中心|实验室|操场)/.test(s)) continue;
                if(s === courseName || s.replace(/\s+/g, "") === norm) continue;
                let n = washName(s);
                if(!n) continue;
                if(n.split(",").length === 1 && /^[一-龥·]{2,10}$/.test(n)) return n;
            }
        }
        return names.join(",");
    }

    function parseCourses(data){
        var courses = new Array();
        for(let i = 0; i < data.divs.length; ++i){
            let course = new Course();
            course.week = data.week[i];
            course.name = extractCourseName(data.divs[i]);
            if(course.name.length == 0){
            }
            let rawText = data.divs[i].innerText || '';
            let wholeText = rawText.replace(/\s+/g, ' ').trim();
            course.info = wholeText;

            let mTime = wholeText.match(/[\(（](\d+)\s*[-—~]\s*(\d+)\s*节[\)）]/);
            if(mTime){
                course.startTime = parseInt(mTime[1]);
                course.endTime = parseInt(mTime[2]);
            } else {
                let m1 = wholeText.match(/[\(（](\d+)\s*节[\)）]/);
                if(m1){ course.startTime = course.endTime = parseInt(m1[1]); }
            }

            course.weekSegments = parseWeekSegments(wholeText);
            course.startWeek = course.weekSegments.map(s => s.start);
            course.endWeek = course.weekSegments.map(s => s.end);
            course.isSingleOrDouble = course.weekSegments.map(s => s.parity === 'all' ? 1 : 2);

            if(/^(考试|考查|考核|停课)/.test(course.name) ||
               (!course.startTime && !course.weekSegments.length && /未安排|考试|考查|考核/.test(wholeText))){
                continue;
            }

            let mLoc = wholeText.match(/(湖光校区|海滨校区|霞山校区|抚州校区)/);
            if(mLoc){
                let after = wholeText.substring(wholeText.indexOf(mLoc[0]) + mLoc[0].length).trim();
                let locWords = [];
                let parts = after.split(/\s+/);
                for(let p of parts){
                    if(!p) continue;
                    if(/校区|楼|场|馆|中心|教室|实验室|篮球场|操场|^[（(]/.test(p)){
                        locWords.push(p);
                    } else {
                        break;
                    }
                }
                course.location = locWords.join(" ");
            }
            if(!course.location){
                let mb = wholeText.match(/(明德楼|广学楼|钟海楼|博文楼|望海楼|篮球场|操场|教室|楼)[A-Za-z0-9(多)楼号]*/);
                if(mb) course.location = mb[0];
            }
            if(course.location){
                course.location = course.location
                    .replace(/^[一-龥]{1,2}(?=明德楼|广学楼|钟海楼|博文楼|望海楼|教学楼|篮球场|操场)/, "")
                    .split(/考查|考核|考试|未安排|面授|讲授|讲课/)[0]
                    .replace(/([)）]|\d)[一-龥]+$/, "$1")
                    .trim();
            }

            course.teacher = parseTeacher(wholeText, course.name, rawText);

            if(!course.startTime){ course.startTime = 1; course.endTime = 1; }
            if(!course.weekSegments || course.weekSegments.length === 0){
            }
            courses.push(course);
        }
        return courses;
    }

    var TIME = {
        "1":  { "name": "第1节",  "startTime": { "hour": '08', "minute": '10' } },
        "2":  { "name": "第2节",  "startTime": { "hour": '09', "minute": '00' } },
        "3":  { "name": "第3节",  "startTime": { "hour": '10', "minute": '15' } },
        "4":  { "name": "第4节",  "startTime": { "hour": '11', "minute": '05' } },
        "5":  { "name": "第5节",  "startTime": { "hour": '14', "minute": '30' } },
        "6":  { "name": "第6节",  "startTime": { "hour": '15', "minute": '20' } },
        "7":  { "name": "第7节",  "startTime": { "hour": '16', "minute": '30' } },
        "8":  { "name": "第8节",  "startTime": { "hour": '17', "minute": '20' } },
        "9":  { "name": "第9节",  "startTime": { "hour": '19', "minute": '30' } },
        "10": { "name": "第10节", "startTime": { "hour": '20', "minute": '25' } }
    };

    function getFixedLen(s, len){
        if(s.length < len){
            return getFixedLen("0" + s, len);
        }
        else if(s.length > len){
            return s.slice(0, len);
        }
        else{
            return s;
        }
    }

    function getTime(num, StartOrEnd) {
        if(num < 1) return "000000";
        var t = TIME[num] || TIME[1];
        var hour = t.startTime.hour;
        var minute = parseInt(t.startTime.minute) + StartOrEnd;
        if (minute >= 60) {
            minute = minute - 60;
            hour = parseInt(hour) + 1;
        }
        return "" + getFixedLen("" + hour, 2) + getFixedLen("" + minute, 2) + "00";
    }

    function formatDate(d){
        return d.getFullYear() + "-" + getFixedLen("" + (d.getMonth() + 1), 2) + "-" + getFixedLen("" + d.getDate(), 2);
    }

    function alignToMonday(dateStr){
        let d = dateStr instanceof Date ? new Date(dateStr) : new Date(dateStr.toString());
        if(isNaN(d.getTime())) d = new Date();
        let day = d.getDay();
        day = day === 0 ? 7 : day;
        d.setDate(d.getDate() - (day - 1));
        return d;
    }

    function getDate(num, wk){
        let date = alignToMonday(startDate);
        date.setDate(date.getDate() + (num - 1) * 7 + Week[wk] - 1);
        let res = "";
        res += getFixedLen(date.getUTCFullYear().toString(), 4);
        res += getFixedLen((date.getUTCMonth() + 1).toString(), 2);
        res += getFixedLen(date.getUTCDate().toString(), 2);
        res += "T";
        return res;
    }

    function weeksText(c){
        if(!c.weekSegments || !c.weekSegments.length) return "";
        return c.weekSegments.map(s => {
            let t = s.start === s.end ? s.start + "周" : s.start + "-" + s.end + "周";
            if(s.parity === 'odd') t += "(单)";
            else if(s.parity === 'even') t += "(双)";
            return t;
        }).join(", ");
    }

    function weekNum(c){
        return (typeof c.week === "number") ? c.week : Week[c.week];
    }

    function countOccurrences(startW, endW, interval){
        if(interval === 1) return endW - startW + 1;
        return Math.floor((endW - startW) / 2) + 1;
    }

    function generateCalendar(courses){
        let res = new ICS();
        courses.forEach(course => {
            for(let i = 0; i < course.isSingleOrDouble.length; ++i){
                let e = new ICSEvent(
                    "" + getDate(course.startWeek[i], course.week) + getTime(course.startTime, 0),
                    "" + getDate(course.startWeek[i], course.week) + getTime(course.endTime, 45),
                    "" + course.name,
                    "" + course.location,
                    "" + (course.teacher ? course.teacher + " " : "") + course.info
                );
                e.setRRULE("WEEKLY", res.Calendar.WKST,
                            "" + countOccurrences(course.startWeek[i], course.endWeek[i], course.isSingleOrDouble[i]),
                            "" + course.isSingleOrDouble[i],
                            "" + course.week.substr(0, 2).toUpperCase());
                res.pushEvent(e);
            }
        });
        res.pushCalendarEnd();
        res.exportIcs();
    }

    function loadScript(src){
        return new Promise(function(resolve, reject){
            let s = document.createElement("script");
            s.src = src;
            s.onload = resolve;
            s.onerror = function(){ reject(new Error("load fail: " + src)); };
            document.head.appendChild(s);
        });
    }

    function loadXLSXFromCDN(){
        let urls = [
            "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
            "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js",
            "https://cdn.bootcdn.net/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
            "https://cdn.staticfile.net/xlsx/0.18.5/xlsx.full.min.js"
        ];
        let i = 0;
        function tryNext(){
            if(i >= urls.length) return Promise.reject(new Error("all cdn failed"));
            let u = urls[i++];
            return loadScript(u).then(function(){
                let X = (typeof unsafeWindow !== "undefined" && unsafeWindow.XLSX) || window.XLSX;
                if(X) return X;
                return tryNext();
            }).catch(function(){ return tryNext(); });
        }
        return tryNext();
    }

    function exportToExcel(courses){
        if(!courses || courses.length === 0){
            alert("未解析到任何课程信息，请确认当前页面为「学生课表查询」页面并已加载课表。");
            return;
        }
        loadXLSXFromCDN().then(function(XLSX){
            try{
                buildAndSaveWorkbook(XLSX, courses);
                alert("Excel文件已导出：课程表.xlsx（含「课表」周视图与「课程明细」两个工作表）");
            }catch(err){
                fallbackToCSV(courses);
                alert("生成.xlsx失败，已降级导出CSV文件，用Excel/WPS打开即可。");
            }
        }).catch(function(){
            fallbackToCSV(courses);
            alert("未能加载Excel库（CDN不可达），已降级导出CSV文件，用Excel/WPS打开即可。");
        });
    }

    function buildAndSaveWorkbook(XLSX, courses){
        const dayNames = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"];

        let maxSection = 10;
        courses.forEach(c => { if(c.endTime > maxSection) maxSection = c.endTime; });

        let grid = {};
        let cellOwners = {};
        courses.forEach((c, ci) => {
            let col = weekNum(c);
            let parts = [c.name];
            if(c.location) parts.push(c.location + (c.teacher ? " " + c.teacher : ""));
            else if(c.teacher) parts.push(c.teacher);
            parts.push(weeksText(c));
            let content = parts.join("\n");
            for(let r = c.startTime; r <= c.endTime; r++){
                let key = r + "_" + col;
                grid[key] = grid[key] ? grid[key] + "\n\n" + content : content;
                (cellOwners[key] = cellOwners[key] || new Set()).add(ci);
            }
        });

        let aoa1 = [["节次"].concat(dayNames)];
        for(let r = 1; r <= maxSection; r++){
            let row = ["第" + r + "节"];
            for(let col = 1; col <= 7; col++){
                row.push(grid[r + "_" + col] || "");
            }
            aoa1.push(row);
        }
        let ws1 = XLSX.utils.aoa_to_sheet(aoa1);
        ws1["!cols"] = [{wch: 8}, {wch: 30}, {wch: 30}, {wch: 30}, {wch: 30}, {wch: 30}, {wch: 26}, {wch: 26}];

        let merges = [];
        courses.forEach((c, ci) => {
            let col = weekNum(c);
            if(c.endTime > c.startTime){
                let ok = true;
                for(let r = c.startTime; r <= c.endTime; r++){
                    let owners = cellOwners[r + "_" + col];
                    if(!owners || owners.size !== 1 || !owners.has(ci)){ ok = false; break; }
                }
                if(ok) merges.push({s: {r: c.startTime, c: col}, e: {r: c.endTime, c: col}});
            }
        });
        ws1["!merges"] = merges;

        let aoa2 = [["序号", "课程名称", "星期", "上课节次", "周次", "单双周", "上课地点", "任课教师", "原始信息"]];
        courses.forEach((c, i) => {
            let col = weekNum(c);
            let segTxts = c.weekSegments.map(s => s.start === s.end ? s.start + "周" : s.start + "-" + s.end + "周");
            let parityTxts = c.weekSegments.map(s => {
                if(s.parity === 'odd') return "单周";
                if(s.parity === 'even') return "双周";
                return "每周";
            });
            aoa2.push([
                i + 1,
                c.name,
                dayNames[col - 1],
                (c.startTime === c.endTime ? "第" + c.startTime + "节" : "第" + c.startTime + "-" + c.endTime + "节"),
                segTxts.join(", "),
                parityTxts.join(", "),
                c.location || "",
                c.teacher || "",
                c.info || ""
            ]);
        });
        let ws2 = XLSX.utils.aoa_to_sheet(aoa2);
        ws2["!cols"] = [{wch: 6}, {wch: 24}, {wch: 8}, {wch: 14}, {wch: 28}, {wch: 16}, {wch: 18}, {wch: 16}, {wch: 32}];

        let wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws1, "课表");
        XLSX.utils.book_append_sheet(wb, ws2, "课程明细");
        let data = XLSX.write(wb, {bookType: "xlsx", type: "array"});
        let blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
        let a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "课程表.xlsx";
        a.click();
        setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
    }

    function fallbackToCSV(courses){
        const dayNames = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"];
        let esc = v => '"' + String(v == null ? "" : v).replace(/"/g, '""') + '"';
        let lines = [["课程名称", "星期", "上课节次", "周次", "单双周", "上课地点", "任课教师"].join(",")];
        courses.forEach(c => {
            let col = weekNum(c);
            let segTxts = c.weekSegments.map(s => s.start === s.end ? s.start + "周" : s.start + "-" + s.end + "周");
            let parityTxts = c.weekSegments.map(s => s.parity === 'odd' ? "单周" : (s.parity === 'even' ? "双周" : "每周"));
            lines.push([
                esc(c.name),
                esc(dayNames[col - 1]),
                esc(c.startTime === c.endTime ? "第" + c.startTime + "节" : "第" + c.startTime + "-" + c.endTime + "节"),
                esc(segTxts.join(", ")),
                esc(parityTxts.join(", ")),
                esc(c.location || ""),
                esc(c.teacher || "")
            ].join(","));
        });
        let csv = "\ufeff" + lines.join("\r\n");
        let blob = new Blob([csv], {type: "text/csv;charset=utf-8"});
        let a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "课程表.csv";
        a.click();
        setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
    }

    function lessonTimeRows(){
        var rows = [["作息时间", "", ""]];
        for(var num = 1; num <= 10; num++){
            var t = TIME[num] || TIME[1];
            var startHour = parseInt(t.startTime.hour, 10);
            var startMinute = parseInt(t.startTime.minute, 10);
            var endMinute = startMinute + 45;
            var endHour = startHour;
            if(endMinute >= 60){ endMinute -= 60; endHour++; }
            rows.push([
                "第" + num + "节",
                getFixedLen("" + startHour, 2) + ":" + getFixedLen("" + startMinute, 2),
                getFixedLen("" + endHour, 2) + ":" + getFixedLen("" + endMinute, 2)
            ]);
        }
        return rows;
    }

    function widgetCellText(c){
        const day = weekNum(c);
        const sec = (c.startTime === c.endTime)
            ? "第" + c.startTime + "节"
            : "第" + c.startTime + "节-第" + c.endTime + "节";
        let clauses = [];
        if(!c.weekSegments || c.weekSegments.length === 0){
            clauses.push("周次待定,星期" + day + "," + sec);
        }
        else{
            for(let s of c.weekSegments){
                const w = s.start === s.end ? "第" + s.start + "周" : s.start + "-" + s.end + "周";
                const pw = w + (s.parity === 'odd' ? "(单)" : s.parity === 'even' ? "(双)" : "");
                clauses.push(pw + ",星期" + day + "," + sec);
            }
        }
        let lines = [c.name].concat(clauses);
        lines.push(c.location || "无");
        return lines.join("\n");
    }

    function buildWidgetWorkbook(XLSX, courses){
        const weekHeaders = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"];
        let maxSection = 10;
        courses.forEach(c => { if(c.endTime > maxSection) maxSection = c.endTime; });

        let used = {};
        let cells = {};
        courses.forEach(c => {
            let col = weekNum(c);
            let row = c.startTime;
            while((used[col] || (used[col] = new Set())).has(row)) row++;
            used[col].add(row);
            cells[row + "_" + col] = (cells[row + "_" + col]
                ? cells[row + "_" + col] + "\n\n"
                : "") + widgetCellText(c);
        });

        let aoa = [["节次"].concat(weekHeaders)];
        for(let r = 1; r <= maxSection; r++){
            let row = ["第" + r + "节"];
            for(let col = 1; col <= 7; col++){
                row.push(cells[r + "_" + col] || "");
            }
            aoa.push(row);
        }
        let ws = XLSX.utils.aoa_to_sheet(aoa);
        ws["!cols"] = [{wch: 8}, {wch: 34}, {wch: 34}, {wch: 34}, {wch: 34}, {wch: 34}, {wch: 30}, {wch: 30}, {wch: 10}, {wch: 10}, {wch: 10}];
        XLSX.utils.sheet_add_aoa(ws, lessonTimeRows(), { origin: "I1" });
        let wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "课表");
        let data = XLSX.write(wb, {bookType: "xlsx", type: "array"});
        let blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
        let a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "课刻-课程表.xlsx";
        a.click();
        setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
    }

    function exportToWidget(courses){
        if(!courses || courses.length === 0){
            alert("未解析到任何课程信息，请确认当前页面为「学生课表查询」页面并已加载课表。");
            return;
        }
        loadXLSXFromCDN().then(function(XLSX){
            try{
                buildWidgetWorkbook(XLSX, courses);
                alert("课刻格式Excel已导出：课刻-课程表.xlsx（含右侧作息时间表，导入课刻自动应用作息）。\n在课刻(desktop-course-widget)中点击「导入新课表」选择该文件即可。");
            }catch(err){
                alert("生成失败：" + (err && err.message ? err.message : err));
            }
        }).catch(function(){
            alert("未能加载Excel库（CDN不可达），无法生成课刻格式xlsx，请联网后重试。");
        });
    }

    function wakeupWeekText(c){
        if(!c.weekSegments || !c.weekSegments.length) return "";
        let parts = [];
        for(let s of c.weekSegments){
            let start = s.start, end = s.end;
            if(start === end){
                parts.push("" + start);
                continue;
            }
            let tag = "";
            if(s.parity === 'odd'){
                tag = "单";
                if(start % 2 === 0) start += 1;
            } else if(s.parity === 'even'){
                tag = "双";
                if(start % 2 === 1) start += 1;
            }
            if(start > end) continue;
            parts.push(start + "-" + end + tag);
        }
        return parts.join("、");
    }

    function wakeupLocationText(c){
        let location = (c.location || "无").trim();
        let teacher = (c.teacher || "").trim();
        if(!teacher) return location;
        if(location.endsWith(" " + teacher)) return location;
        if(location.endsWith(teacher)){
            location = location.slice(0, location.length - teacher.length).trim();
        }
        return (location ? location + " " : "") + teacher;
    }

    function exportToWakeUp(courses){
        if(!courses || courses.length === 0){
            alert("未解析到任何课程信息，请确认当前页面为「学生课表查询」页面并已加载课表。");
            return;
        }
        let esc = v => '"' + String(v == null ? "" : v).replace(/"/g, '""') + '"';
        let lines = [["课程名称", "星期", "开始节数", "结束节数", "老师", "地点", "周数"].join(",")];
        let noWeek = [];
        courses.forEach(c => {
            let wk = wakeupWeekText(c);
            if(!wk){
                noWeek.push(c.name);
                return;
            }
            lines.push([
                esc(c.name || "未命名课程"),
                weekNum(c),
                c.startTime,
                c.endTime,
                esc(c.teacher || "无"),
                esc(wakeupLocationText(c)),
                esc(wk)
            ].join(","));
        });
        let csv = "\ufeff" + lines.join("\r\n");
        let blob = new Blob([csv], {type: "text/csv;charset=utf-8"});
        let a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "WakeUp-课程表.csv";
        a.click();
        setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
        let msg = "WakeUp课程表CSV已导出：WakeUp-课程表.csv\n\n⚠千万别用Excel/WPS打开后再保存！\n这俩软件会把纯周数\"1-4\"自动显示成\"1月4日\"，看着不对，再一保存文件就真废了。想看内容就用记事本打开，或者直接发手机导入。\n\n导入方法：把文件发到手机(QQ/微信均可)，在聊天里点开文件 → 用其他应用打开 → 选WakeUp课程表；或在App里 我的 → 课程表管理 → 导入课程表(Excel导入) → 选该文件。\n\n⏰重要：CSV只有节次、不含上课时间，导入后时间显示是App默认作息，请在App里改一次就永久生效：\n主界面右上角三个点 → 课表设置 → 上课时间 → 取消勾选「每节课时长相同」，逐节改为海大作息：\n第1节08:10-08:55 第2节09:00-09:45 第3节10:15-11:00 第4节11:05-11:50 第5节14:30-15:15 第6节15:20-16:05 第7节16:30-17:15 第8节17:20-18:05 第9节19:30-20:15 第10节20:25-21:10\n再到「课表数据」里把学期开始日期设为开学第一周周一(当前学期2026-08-31)，周次才显示正确。";
        if(noWeek && noWeek.length){
            msg += "\n\n⚠以下课程周次未能解析，已跳过未写入(写空周数会导致整份导入失败)，请导入后在App里手动添加：\n" + noWeek.join("、");
        }
        alert(msg);
    }

}

var CRLF = "\n";
var SPACE = " ";
class ICS{

    Calendar;
    ics;
    res;

    constructor(){
        (function(Calendar){
            Calendar.PRODID = "-//31415926535x//ICalendar Exporter v1.0//CN";
            Calendar.VERSION = "2.0";
            Calendar.CALSCALE = "GREGORIAN";
            Calendar.TIMEZONE = "Asia/Shanghai";
            Calendar.ISVALARM = true;
            Calendar.VALARM = "-PT0D0H10M0S";
            Calendar.WKST = "SU";
        })(this.Calendar || (this.Calendar = {}));

        this.ics = new Array();
        this.ics.push("BEGIN:VCALENDAR");
        this.ics.push("VERSION:" + this.Calendar.VERSION);
        this.ics.push("PRODID:" + this.Calendar.PRODID);
        this.ics.push("CALSCALE:" + this.Calendar.CALSCALE);
    }

    pushEvent(e){
        this.ics.push("BEGIN:VEVENT");
        this.ics.push(e.getDTSTART());
        this.ics.push(e.getDTEND());
        if(e.isrrule == true)this.ics.push(e.getRRULE());
        this.ics.push(e.getSUMMARY());
        this.ics.push(e.getLOCATION());
        this.ics.push(e.getDESCRIPTION());
        if(this.Calendar.ISVALARM == true)this.pushAlarm();
        this.ics.push("END:VEVENT");
        this.ics.push(CRLF);
    }

    pushAlarm(){
        this.ics.push("BEGIN:VALARM");
        this.ics.push("ACTION:DISPLAY");
        this.ics.push("TRIGGER:" + this.Calendar.VALARM);
        this.ics.push("END:VALARM");
    }

    pushCalendarEnd(){
        this.ics.push("END:VCALENDAR");
    }

    getFixedIcs(){
        this.res = "";
        this.ics.forEach(line => {
            let bytes = unescape(encodeURIComponent(line)).length;
            if(bytes > 70){
                let cur = "", parts = [];
                for(let ch of line){
                    cur += ch;
                    if(unescape(encodeURIComponent(cur)).length >= 60){
                        parts.push(cur); cur = "";
                    }
                }
                if(cur) parts.push(cur);
                parts.forEach((part, idx) => {
                    this.res += (idx === 0 ? "" : SPACE) + part + CRLF;
                });
            }
            else{
                this.res += line + CRLF;
            }
        });
        return this.res;
    }

    exportIcs(){
        this.getFixedIcs();
        let link = window.URL.createObjectURL(new Blob([this.res], {
            type: "text/calendar;charset=utf-8"
        }));
        let a = document.createElement("a");
        a.setAttribute("href", link);
        a.setAttribute("download", "courses.ics");
        a.click();
    }
}
class ICSEvent{
    constructor(DTSTART, DTEND, SUMMARY, LOCATION, DESCRIPTION){
        this.DTSTART = DTSTART;
        this.DTEND = DTEND;
        this.SUMMARY = SUMMARY;
        this.LOCATION = LOCATION;
        this.DESCRIPTION = DESCRIPTION;
    }
    isrrule = false;
    RRULE;
    setRRULE(FREQ, WKST, COUNT, INTERVAL, BYDAY){
        this.isrrule = true;
        this.RRULE = "RRULE:FREQ=" + FREQ + ";WKST=" + WKST + ";COUNT=" + COUNT + ";INTERVAL=" + INTERVAL + ";BYDAY=" + BYDAY;
    }
    getRRULE(){
        return "" + this.RRULE;
    }
    getDTSTART(){
        return "DTSTART:" + this.DTSTART;
    }
    getDTEND(){
        return "DTEND:" + this.DTEND;
    }
    getSUMMARY(){
        return "SUMMARY:" + this.SUMMARY;
    }
    getLOCATION(){
        return "LOCATION:" + this.LOCATION;
    }
    getDESCRIPTION(){
        return "DESCRIPTION:" + this.DESCRIPTION;
    }
}

