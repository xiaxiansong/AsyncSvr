# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from cases import SHEETS, NON_MANUAL
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT_DIR = sys.argv[1]
FONT = "Arial"
thin = Side(style="thin", color="999999")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
HEAD_FILL = PatternFill("solid", start_color="1F4E78")
HEAD_FONT = Font(name=FONT, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT, size=10)
INPUT_FILL = PatternFill("solid", start_color="FFFF00")
WRAP = Alignment(wrap_text=True, vertical="top")

def num(lst):
    return "\n".join(f"{i}、{x}" for i, x in enumerate(lst, 1))

def header(ws, cols, widths, row=1):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font = HEAD_FONT; cell.fill = HEAD_FILL; cell.border = BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[cell.column_letter].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=3)

def body(cell):
    cell.font = BODY_FONT; cell.alignment = WRAP; cell.border = BORDER

wb = Workbook()
wb.remove(wb.active)

# ---------- 说明 ----------
ws = wb.create_sheet("说明")
lines = [
    ("洛书（屏幕感知）多机型适配 —— 真机手工测试用例", True),
    ("测试对象：swing_service 在洛书链路中的职责（接入鉴权、事件生命周期、参数下发、数据搬运、设备状态注入、结果分发、可靠性与打点）。测试方式：真机手工操作，同时抓日志核对。不评价算法识别准确率。", False),
    ("用例组织：公共用例（每个机型都要执行，建议再拿一台直板机作为基线）+ 三折叠 / 双折叠 / 阔折叠 / 小折叠 专项用例。手工无法覆盖的项见《非手工覆盖项》页。", False),
    ("执行顺序：每个机型先跑公共 L0，再跑本机型专项 L0，最后跑 L1/L2。L0=冒烟必过，L1=功能完整性，L2=边缘场景。", False),
    ("", False),
    ("一、入口约定（请填写一次，用例中统一使用以下简称）", True),
    ("【屏幕感知】（触发 901 ACTIVE_SCREEN_AWARENESS）入口：____________________", False),
    ("【游戏理解】（触发 903 GAME_OCR_UNDERSTAND）入口：____________________　测试游戏：____________", False),
    ("“触发屏幕感知 / 触发识别”指按业务方式让功能给出一次结果。", False),
    ("", False),
    ("二、每条用例的日志抓取流程", True),
    ("1、执行前：hdc shell hilog -r 清空缓存日志（必要时 hdc shell hilog -b D 打开 DEBUG 级别）", False),
    ("2、执行中：hdc shell hilog > 用例编号_机型.log 实时落盘（或使用公司内部抓日志工具）", False),
    ("3、执行后：按该用例的“日志核对点”列和《日志核对指南》页逐项检索；检查 /data/log/faultlog/ 是否有新增 crash", False),
    ("4、不通过时：在“备注/问题单”列记录现象、时间点、日志文件名和问题单号", False),
    ("", False),
    ("三、预期值", True),
    ("《预期值对照表》页的黄色单元格请按机型适配规格（FoldStatus 映射表、phyrotation.offset、rotation_correction、cameraId）先填好期望值，执行时填实测值，结论列自动判定。", False),
    ("注意：日志关键字、命令参数为建议写法，请以代码实际打印内容和本分支工具帮助为准。", False),
]
ws.column_dimensions["A"].width = 120
for i, (t, b) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t)
    c.font = Font(name=FONT, bold=b, size=14 if (b and i == 1) else (11 if b else 10))
    c.alignment = Alignment(wrap_text=True, vertical="top")

# ---------- 用例页 ----------
COLS = ["用例_编号", "用例_级别", "用例_名称", "用例_预置条件", "用例_测试步骤", "用例_预期结果", "日志核对点", "执行结果", "备注/问题单"]
WIDTHS = [14, 8, 32, 40, 48, 52, 40, 10, 20]
for sheet_name, cases in SHEETS:
    ws = wb.create_sheet(sheet_name)
    header(ws, COLS, WIDTHS)
    dv = DataValidation(type="list", formula1='"通过,不通过,阻塞,不涉及"', allow_blank=True)
    ws.add_data_validation(dv)
    for r, (cid, lvl, name, pre, steps, exp, log, _tags) in enumerate(cases, 2):
        vals = [cid, lvl, name, num(pre), num(steps), num(exp), num(log), None, None]
        for ci, v in enumerate(vals, 1):
            body(ws.cell(row=r, column=ci, value=v))
        dv.add(ws.cell(row=r, column=8))
    ws.auto_filter.ref = f"A1:I{len(cases)+1}"

# ---------- 日志核对指南 ----------
ws = wb.create_sheet("日志核对指南")
cols = ["职责", "每条用例要核对什么", "建议检索关键字（以代码实际打印为准）", "判定为问题的现象"]
header(ws, cols, [16, 50, 40, 50])
guide = [
    ("接入与鉴权", "权限校验结果；事件名→eventId（901/903）；domainId 校验", "permission / SUBSCRIBE_SWING_ABILITY / eventId / domainId", "有权限被拒；无权限放行；映射错 eventId；非法入参仍向 IGS 注册"),
    ("事件生命周期", "订阅/退订、引用计数、向 IGS 注册/注销次数；睡眠/S4 下是否注册；相机互斥白名单", "Subscribe / Unsubscribe / refCount / Register / Unregister / sleep / S4 / preload / whitelist", "引用计数为负或不归零；重复注册；退订后仍有结果；901/903 被降级为预加载；唤醒后未恢复"),
    ("注册与参数", "参数串与 SwingParameter.json 本机型配置一致；SetParam 命令码 1/79/73；EnableEvent", "SwingParameter / param / SetParam / EnableEvent", "取到默认值或其它机型值；SetParam 返回失败"),
    ("数据通道", "SendFileByFd 的 fd、byteSize、DataType→底层码（64/65/66/81/67/77）；fd 关闭", "SendFileByFd / DataType / fd / byteSize / fdsan / close", "映射错码；byteSize 不一致；fd 未关闭、fd 数量增长；fdsan 告警/abort"),
    ("设备状态注入（多机型核心）", "stateId=0 相机切换、1 亮灭屏、2 旋转、3 折叠；903 的 78 物理方向、80 折叠状态；初始注入值；切换时序", "stateId / FoldStatus / DisplayMode / rotation / phyrotation / rotation_correction / cameraId", "注入值与预期值表不符；出现 unknown/默认值；形态切换后未注入或注入积压；切屏误报灭屏；横竖方向 90/270 互换"),
    ("结果上行", "OnIgsEventDone → publishSwingEvent → 订阅者；901 不打印结果内容；ion_fd 生命周期", "OnIgsEventDone / publishSwingEvent / ion_fd", "结果串发给错误订阅者；901 结果内容出现在日志中；回调后访问已释放 buffer"),
    ("可靠性与打点", "OnServiceDied 重连；SensorHub 复位后参数重放；服务重启恢复；HiSysEvent 使用时长", "OnServiceDied / reconnect / replay / HiSysEvent", "未重连；重放的状态值是旧值；时长偏差大或字段缺失"),
    ("通用异常", "每条用例执行后检索", "fail / error / invalid / unknown / timeout / ret= 非 0", "非预期的错误日志；/data/log/faultlog 下新增 swing_service 相关 cppcrash"),
]
for r, row in enumerate(guide, 2):
    for ci, v in enumerate(row, 1):
        body(ws.cell(row=r, column=ci, value=v))
r = len(guide) + 3
ws.cell(row=r, column=1, value="常用命令（参考，以本分支工具帮助为准）").font = Font(name=FONT, bold=True)
cmds = [
    ("清空/抓取日志", "hdc shell hilog -r ；hdc shell hilog -b D ；hdc shell hilog | grep -iE \"swing|igs|hiai\""),
    ("查看机型参数", "hdc shell param get const.window.phyrotation.offset ；hdc shell param get const.dms.rotation_correction"),
    ("进程与 fd", "（调试版本）hdc shell pidof <swing_service 进程名> ；hdc shell ls /proc/<pid>/fd | wc -l"),
    ("内存", "hdc shell hidumper --mem <pid>"),
    ("打点", "hdc shell hisysevent -l（按事件域过滤）"),
    ("崩溃", "hdc shell ls /data/log/faultlog/faultlogger/ ；hdc file recv 拉取到电脑"),
    ("折叠/显示状态", "hidumper 查询 DisplayManager 的当前 FoldStatus / FoldDisplayMode"),
]
for i, (a, b) in enumerate(cmds, r + 1):
    body(ws.cell(row=i, column=1, value=a)); body(ws.cell(row=i, column=2, value=b))
    ws.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)

# ---------- 预期值对照表 ----------
ws = wb.create_sheet("预期值对照表")
ws.cell(row=1, column=1, value="说明：黄色列按机型适配规格填写期望值；执行时在“实测”列填写日志中的实际注入值；“结论”列自动判定（全部实测为空时不判定）。第 3 行为格式示例（非真实值）。").font = Font(name=FONT, italic=True)
ws.merge_cells("A1:O1")
cols = ["机型", "形态", "显示屏", "物理方向", "期望FoldStatus", "期望DisplayMode",
        "期望stateId=3/80", "实测stateId=3/80", "期望stateId=2", "实测stateId=2",
        "期望78", "实测78", "期望cameraId(stateId=0)", "实测cameraId", "结论"]
header(ws, cols, [10, 16, 10, 10, 16, 16, 14, 14, 12, 12, 10, 10, 18, 14, 10], row=2)
ws.freeze_panes = "E3"
devices = [
    ("三折叠", [("单屏态", "单屏"), ("双屏态", "双屏"), ("三屏态", "三屏"), ("第一铰链半开", "单屏/双屏"), ("第二铰链半开", "双屏/三屏")]),
    ("双折叠", [("折叠", "外屏"), ("展开", "内屏"), ("半折(书本)", "内屏"), ("半折(帐篷)", "外屏")]),
    ("阔折叠", [("折叠", "外屏"), ("展开", "内屏"), ("半折悬停", "内屏")]),
    ("小折叠", [("合盖", "外屏"), ("展开", "内屏"), ("自由悬停", "内屏")]),
]
rows = [("示例", "展开", "内屏", "90°", "EXPANDED", "FULL", "1", "1", "1", "1", "1", "3", "1", "1")]
for dev, states in devices:
    for st, scr in states:
        for d in ["0°", "90°", "180°", "270°"]:
            rows.append((dev, st, scr, d) + (None,) * 10)
for r, row in enumerate(rows, 3):
    for ci, v in enumerate(row, 1):
        c = ws.cell(row=r, column=ci, value=v); body(c)
        if ci in (5, 6, 7, 9, 11, 13) and r > 3:
            c.fill = INPUT_FILL
    if r == 3:
        for ci in range(1, 16):
            ws.cell(row=r, column=ci).font = Font(name=FONT, size=10, italic=True, color="808080")
    f = (f'=IF(COUNTA(H{r},J{r},L{r},N{r})=0,"",IF(AND(G{r}&""=H{r}&"",I{r}&""=J{r}&"",'
         f'K{r}&""=L{r}&"",M{r}&""=N{r}&""),"通过","不通过"))')
    body(ws.cell(row=r, column=15, value=f))

# ---------- 覆盖矩阵 ----------
ws = wb.create_sheet("覆盖矩阵")
cols = ["覆盖维度", "公共", "三折叠", "双折叠", "阔折叠", "小折叠"]
header(ws, cols, [36, 22, 22, 22, 22, 22])
DIMS = ["能力探测", "接入鉴权", "开关生命周期", "订阅方退出清理", "亮灭屏/睡眠", "相机", "参数下发",
        "数据通道", "游戏903", "日志隐私", "打点", "初始状态注入", "折叠映射", "半折", "旋转", "切屏",
        "外屏", "系统场景", "稳定性"]
for r, dim in enumerate(DIMS, 2):
    body(ws.cell(row=r, column=1, value=dim))
    for ci, (sheet_name, cases) in enumerate(SHEETS, 2):
        hit = [c[0] for c in cases if dim in c[7]]
        body(ws.cell(row=r, column=ci, value="、".join(hit) if hit else "—"))
unknown = {t for _, cs in SHEETS for c in cs for t in c[7]} - set(DIMS)
assert not unknown, unknown

# ---------- 非手工覆盖项 ----------
ws = wb.create_sheet("非手工覆盖项")
header(ws, ["序号", "测试点", "对应职责", "建议覆盖方式"], [8, 60, 16, 40])
for r, (pt, duty, how) in enumerate(NON_MANUAL, 2):
    for ci, v in enumerate((r - 1, pt, duty, how), 1):
        body(ws.cell(row=r, column=ci, value=v))

path = os.path.join(OUT_DIR, "洛书多机型适配测试用例.xlsx")
wb.save(path)
print(path)

# ---------- Markdown ----------
md = ["# 洛书（屏幕感知）多机型适配测试用例", "",
      "> 真机手工测试用例，与同目录 Excel 同源生成。Excel 中另含《说明》（入口约定、日志抓取流程）《日志核对指南》《预期值对照表》《覆盖矩阵》《非手工覆盖项》。", ""]
for sheet_name, cases in SHEETS:
    md.append(f"## {sheet_name}（{len(cases)} 条）\n")
    for cid, lvl, name, pre, steps, exp, log, _tags in cases:
        md.append(f"### {cid}【{lvl}】{name}\n")
        md.append("| 字段 | 内容 |\n|---|---|")
        br = lambda l: "<br>".join(f"{i}、{x}" for i, x in enumerate(l, 1)).replace("|", "\\|")
        md.append(f"| 用例_名称 | {name} |")
        md.append(f"| 用例_预置条件 | {br(pre)} |")
        md.append(f"| 用例_测试步骤 | {br(steps)} |")
        md.append(f"| 用例_预期结果 | {br(exp)} |")
        md.append(f"| 日志核对点 | {br(log)} |\n")
md.append("## 非手工覆盖项（建议开发自测/自动化）\n")
md.append("| 测试点 | 职责 | 建议方式 |\n|---|---|---|")
for pt, duty, how in NON_MANUAL:
    md.append(f"| {pt} | {duty} | {how} |")
open(os.path.join(OUT_DIR, "洛书多机型适配测试用例.md"), "w", encoding="utf-8").write("\n".join(md) + "\n")
