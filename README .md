# 广东海洋大学 · 正方教务课表导出

> 适配广海教务系统的课表导出插件，把课表导成日历ics、Excel、课刻、WakeUp 四种格式。

[![License](https://img.shields.io/badge/License-MIT-2ea44f.svg)](LICENSE.txt)
[![Tampermonkey](https://img.shields.io/badge/Tampermonkey-4.x%2B-ff6f00.svg)](https://www.tampermonkey.net/)
[![Platform](https://img.shields.io/badge/Chrome%20%7C%20Edge%20%7C%20Firefox-supported-blue.svg)](#安装)

适用页面：`jw.gdou.edu.cn` 或 `wfw.gdou.edu.cn` 的 `kbcx/xskbcx_cxXskbcxIndex.html`（个人课表查询）。

---

## 四种导出

| 导出 | 文件 | 用在哪 |
| --- | --- | --- |
| **ICS 日历** | `*.ics` | 系统日历、Google Calendar、Outlook、手机日历；支持单双周与调课单周 |
| **普通 Excel** | `课程表.xlsx` | 「课表」周视图 + 「课程明细」两个工作表，跨节课程自动合并单元格 |
| **课刻** | `课刻-课程表.xlsx` | 课刻 `desktop-course-widget` 导入；右侧附 10 节作息表，导入后自动应用 |
| **WakeUp** | `*.csv` | WakeUp 课程表 App；七列固定格式，老师、地点直接可用 |

## 功能特性

- ✅ 四个导出按钮 + 「本学期第一个星期一」日期选择框，默认 `2026-08-31`，可自行修改
- ✅ 节次兼容 `(1-2节)`、`(3~4节)`、`(单节)`、`(1—2节)` 等写法
- ✅ 周次兼容 `1-5周,7-8周`、`第3周`、`1-3周(单),7-13周(单)`、`2-16周(双)`、`◎` 图标、`【调】` 前缀、全角括号
- ✅ 新版界面（td 无 id）自动兜底：按「星期X」表头列 + 行首节次定位课程
- ✅ 任课教师按教学班逐行读取：支持带 `*` / `$` 属性标记的教学班、`（外聘）` 备注、职称尾巴、少数民族长名（含 `·`）、多人任课
- ✅ 周次解析失败不臆造「1-20 周」，该课程留空，课刻侧安全跳过

## 作息时间表

广东海洋大学现行作息，脚本头部 `TIME` 常量中维护，可按学校情况修改。

| 节次 | 时间 | 节次 | 时间 |
| --- | --- | --- | --- |
| 第 1 节 | 08:10 – 08:55 | 第 6 节 | 15:20 – 16:05 |
| 第 2 节 | 09:00 – 09:45 | 第 7 节 | 16:30 – 17:15 |
| 第 3 节 | 10:15 – 11:00 | 第 8 节 | 17:20 – 18:05 |
| 第 4 节 | 11:05 – 11:50 | 第 9 节 | 19:30 – 20:15 |
| 第 5 节 | 14:30 – 15:15 | 第 10 节 | 20:25 – 21:10 |

## 安装

1. 浏览器安装 [Tampermonkey](https://www.tampermonkey.net/)（Chrome / Edge / Firefox 均可）。
2. 打开仓库里的 [`V3GDOU_课表导出_ICS_Excel.user.js`](V3GDOU_%E8%AF%BE%E8%A1%A8%E5%AF%BC%E5%87%BA_ICS_Excel.user.js)，点右上角 **Raw**，Tampermonkey 会弹出安装页，确认安装。
3. 登录教务系统，进入 **个人课表查询**，页面右上角出现四个导出按钮即安装成功。

> 从 Raw 地址安装后，Tampermonkey 会记住该地址并自动检查更新，发新版本只需改 `@version` 再推一次。

## 使用

1. **导出 ICS** —— 生成 `.ics`，导入任意日历软件。Google Calendar 需自行设置课程颜色。
2. **导出 Excel** —— 生成 `课程表.xlsx`，适合打印或核对。
3. **导出课刻** —— 生成 `课刻-课程表.xlsx`，在课刻中点「导入新课表」选择该文件，作息按右侧表格自动更新。
4. **导出 WakeUp** —— 生成 WakeUp 专用 csv；老师列为空时写「无」，地点列自动在末尾追加任课教师，导入后课表上显示「地点 老师」。

日期选择框用来确认本学期第一个星期一的日期，ICS 的周次计算以此为准。

## 解析说明

课程名后面的 `*`、`$`、`#` 是正方给教学班打的属性标记，不是课程名的一部分，脚本原样保留。

这类标记不同的班（例如 `人工智能基础*` 与 `人工智能基础$`）是两个教学班，教师、时间、地点都可能不同：

- 教师从**该教学班自己的课表块**逐行读取，不从别的教学班复制；
- 若该教学班教师行本身为空（常见于线上、网络课程），导出时写「无」，可到课程详情里核对。

## 常见问题

| 现象 | 说明 |
| --- | --- |
| 页面右上角没有按钮 | 确认当前是「学生课表查询」页面，且课表已加载；刚进页面时刷新一次即可 |
| 提示「未找到课表表格」 | 同上；新版界面已做兜底解析，仍失败请提 Issue 并附页面截图 |
| 某门课老师显示「无」 | 该教学班自己的教师行是空的，多为线上/网络课 |
| 周次显示异常 | 脚本不臆造周次，会留空；请附该课程原文提 Issue |
| 课刻导入后时间不对 | 确认导入的是「导出课刻」文件（带右侧作息表），并核对脚本头部 `TIME` |

## 项目结构

```
.
├── README.md
├── LICENSE.txt
└── V3GDOU_课表导出_ICS_Excel.user.js   # 油猴脚本，单文件无依赖
```

## 致谢

本项目基于以下开源项目修改适配：

- [Feldan/ZFClassTable2ICS](https://github.com/Feldan/ZFClassTable2ICS) —— 正方课表 → ICS 的参考实现
- [31415926535x/CollegeProjectBackup](https://github.com/31415926535x/CollegeProjectBackup) —— `ZhengfangClassScheduleToICS` 原始脚本
- [tyr1onX/desktop-course-widget](https://github.com/tyr1onX/desktop-course-widget) —— 课刻，其 xlsx 网格与作息表解析规则（`parse_cell` / `find_lesson_times`）

脚本头部保留原作者 `@author feldan / 31415926535x` 与 MIT 许可声明。

## 许可

[MIT](LICENSE.txt)

> （注：部分内容由AI 生成，本项目豆包工作参与协助，用的kimi k3.1pro）