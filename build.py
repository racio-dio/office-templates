from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.utils import get_column_letter as L
import datetime as dt

BLUE="1F4E79"; LIGHT="DDEBF7"; GRAY="F2F2F2"
thin=Side(style="thin",color="BFBFBF"); B=Border(left=thin,right=thin,top=thin,bottom=thin)
F="微软雅黑"
def title(ws,text,ncol,sub=None):
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncol)
    c=ws.cell(1,1,text); c.font=Font(name=F,size=16,bold=True,color="FFFFFF")
    c.fill=PatternFill("solid",fgColor=BLUE); c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[1].height=32
    if sub:
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=ncol)
        s=ws.cell(2,1,sub); s.font=Font(name=F,size=9,italic=True,color="7F7F7F"); s.alignment=Alignment(horizontal="center")
def header(ws,row,cols,widths=None):
    for i,h in enumerate(cols,1):
        c=ws.cell(row,i,h); c.font=Font(name=F,bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="2E75B6")
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=B
    if widths:
        for i,w in enumerate(widths,1): ws.column_dimensions[L(i)].width=w
    ws.row_dimensions[row].height=24
def body(ws,r1,r2,c1,c2):
    for r in range(r1,r2+1):
        for c in range(c1,c2+1):
            x=ws.cell(r,c); x.border=B; x.font=Font(name=F,size=10); x.alignment=Alignment(horizontal="center",vertical="center")
            if r%2==0: x.fill=PatternFill("solid",fgColor=GRAY)

# 1 考勤
def kaoqin():
    wb=Workbook(); ws=wb.active; ws.title="考勤表"
    title(ws,"月度考勤统计表",40,"只需修改年份、月份和员工姓名；在日期格里选择 √ 出勤 / 迟 迟到 / 假 请假 / 旷 旷工 / 休 休息，右侧自动统计")
    ws["A3"]="年份"; ws["B3"]=2026; ws["C3"]="月份"; ws["D3"]=10
    for a in("A3","C3"): ws[a].font=Font(name=F,bold=True)
    for a in("B3","D3"): ws[a].fill=PatternFill("solid",fgColor="FFF2CC"); ws[a].border=B
    cols=["序号","姓名"]+[str(i) for i in range(1,32)]+["出勤","迟到","请假","旷工","休息","出勤率","备注"]
    header(ws,5,cols)
    ws.column_dimensions["A"].width=6; ws.column_dimensions["B"].width=10
    for i in range(3,34):
        col=L(i); ws.column_dimensions[col].width=4.2
        ws.cell(4,i,f'=IFERROR(MID("日一二三四五六",WEEKDAY(DATE($B$3,$D$3,{i-2})),1),"")')
        ws.cell(4,i).font=Font(name=F,size=8,color="7F7F7F"); ws.cell(4,i).alignment=Alignment(horizontal="center")
        ws.cell(5,i,f'=IF({i-2}<=DAY(EOMONTH(DATE($B$3,$D$3,1),0)),{i-2},"")')
    for i in range(34,41): ws.column_dimensions[L(i)].width=7
    ws.column_dimensions[L(40)].width=12
    dv=DataValidation(type="list",formula1='"√,迟,假,旷,休"',allow_blank=True); ws.add_data_validation(dv)
    names=["张三","李四","王五","赵六","钱七"]
    for k in range(30):
        r=6+k; ws.cell(r,1,k+1); 
        if k<5: ws.cell(r,2,names[k])
        rng=f"C{r}:AG{r}"; dv.add(rng)
        for j,s in enumerate(["√","迟","假","旷","休"]): ws.cell(r,34+j,f'=COUNTIF({rng},"{s}")')
        ws.cell(r,39,f'=IFERROR((AH{r}+AI{r})/(AH{r}+AI{r}+AJ{r}+AK{r}),"")'); ws.cell(r,39).number_format="0.0%"
    body(ws,6,35,1,40)
    for s,color in [("迟","FFE699"),("假","BDD7EE"),("旷","F8CBAD"),("休","E2EFDA")]:
        ws.conditional_formatting.add("C6:AG35",CellIsRule(operator="equal",formula=[f'"{s}"'],fill=PatternFill("solid",fgColor=color)))
    ws.conditional_formatting.add("C4:AG35",FormulaRule(formula=['OR(C$4="六",C$4="日")'],font=Font(color="C00000")))
    ws.freeze_panes="C6"
    wb.save("01_月度考勤统计表.xlsx")

# 2 工资条
def gongzi():
    wb=Workbook(); ws=wb.active; ws.title="工资明细"
    title(ws,"工资计算表（自动算个税）",13,"填写黄色区域；社保公积金按比例自动算，个税按 2026 年综合所得月度预扣简化计算（起征点 5000），仅供参考")
    ws["A3"]="社保个人比例"; ws["B3"]=0.105; ws["C3"]="公积金比例"; ws["D3"]=0.07
    for a in("B3","D3"): ws[a].number_format="0.0%"; ws[a].fill=PatternFill("solid",fgColor="FFF2CC")
    cols=["工号","姓名","部门","基本工资","绩效奖金","补贴","缺勤扣款","应发工资","社保","公积金","专项附加扣除","个税","实发工资"]
    header(ws,5,cols,[8,10,10,11,11,9,10,12,10,10,12,10,12])
    data=[("A001","张三","销售部",8000,2500,500,0,2000),("A002","李四","技术部",15000,3000,800,200,3000),("A003","王五","财务部",9000,1000,500,0,1000)]
    for k in range(20):
        r=6+k
        if k<3:
            for j,v in enumerate(data[k][:7]): ws.cell(r,1+j,v)
            ws.cell(r,11,data[k][7])
        ws.cell(r,8,f'=IF(B{r}="","",D{r}+E{r}+F{r}-G{r})')
        ws.cell(r,9,f'=IF(B{r}="","",ROUND(D{r}*$B$3,2))')
        ws.cell(r,10,f'=IF(B{r}="","",ROUND(D{r}*$D$3,2))')
        t=f'MAX(H{r}-I{r}-J{r}-K{r}-5000,0)'
        ws.cell(r,12,f'=IF(B{r}="","",ROUND(MAX({t}*{{0.03,0.1,0.2,0.25,0.3,0.35,0.45}}-{{0,210,1410,2660,4410,7160,15160}}),2))')
        ws.cell(r,13,f'=IF(B{r}="","",H{r}-I{r}-J{r}-L{r})')
        for c in range(4,14):
            ws.cell(r,c).number_format="#,##0.00"
            if c in(4,5,6,7,11): ws.cell(r,c).fill=PatternFill("solid",fgColor="FFF2CC")
    body(ws,6,25,1,13)
    for r in range(6,26):
        for c in (4,5,6,7,11): ws.cell(r,c).fill=PatternFill("solid",fgColor="FFF9E5")
    ws.cell(26,2,"合计").font=Font(name=F,bold=True)
    for c in range(8,14): ws.cell(26,c,f"=SUM({L(c)}6:{L(c)}25)").number_format="#,##0.00"
    ws.freeze_panes="C6"
    # 工资条
    s=wb.create_sheet("工资条（打印）")
    title(s,"工资条 — 在 A3 输入第几位员工，或整页打印自动生成全部",13)
    for i in range(20):
        hr=4+i*3
        for j,h in enumerate(cols,1):
            c=s.cell(hr,j,h); c.font=Font(name=F,bold=True,size=9,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="2E75B6"); c.border=B; c.alignment=Alignment(horizontal="center")
            d=s.cell(hr+1,j,f"=IF(工资明细!$B${6+i}=\"\",\"\",工资明细!{L(j)}{6+i})"); d.border=B; d.font=Font(name=F,size=9); d.alignment=Alignment(horizontal="center")
            if j>=4: d.number_format="#,##0.00"
    for j in range(1,14): s.column_dimensions[L(j)].width=11
    wb.save("02_工资计算与工资条.xlsx")

# 3 项目甘特图
def gantt():
    wb=Workbook(); ws=wb.active; ws.title="项目进度"
    title(ws,"项目进度甘特图",50,"修改开始日期和工期，右侧色块自动生成；进度填 0–100%，今天所在列自动标红")
    ws["A3"]="项目起始日"; ws["B3"]=dt.date(2026,10,12); ws["B3"].number_format="yyyy-mm-dd"; ws["B3"].fill=PatternFill("solid",fgColor="FFF2CC")
    cols=["序号","任务","负责人","开始日期","工期(天)","结束日期","进度","状态"]
    header(ws,5,cols,[6,22,10,12,9,12,8,9])
    tasks=[("需求调研","张三",0,5,1),("方案设计","李四",4,6,0.8),("开发实现","王五",9,15,0.4),("测试验收","赵六",22,6,0),("上线交付","张三",28,3,0)]
    for k in range(15):
        r=6+k; ws.cell(r,1,k+1)
        if k<5:
            t=tasks[k]; ws.cell(r,2,t[0]); ws.cell(r,3,t[1]); ws.cell(r,4,f"=$B$3+{t[2]}"); ws.cell(r,5,t[3]); ws.cell(r,7,t[4])
        ws.cell(r,4).number_format="mm-dd"; ws.cell(r,6,f'=IF(D{r}="","",D{r}+E{r}-1)'); ws.cell(r,6).number_format="mm-dd"
        ws.cell(r,7).number_format="0%"
        ws.cell(r,8,f'=IF(B{r}="","",IF(G{r}>=1,"已完成",IF(AND(TODAY()>F{r},G{r}<1),"延期",IF(G{r}>0,"进行中","未开始"))))')
    body(ws,6,20,1,8)
    for i in range(42):
        c=9+i; ws.column_dimensions[L(c)].width=3.2
        h=ws.cell(5,c,f"=$B$3+{i}"); h.number_format="d"; h.font=Font(name=F,size=8,bold=True,color="FFFFFF"); h.fill=PatternFill("solid",fgColor="2E75B6"); h.alignment=Alignment(horizontal="center")
        m=ws.cell(4,c,f'=IF(DAY(I$5)=1,MONTH(I$5)&"月","")' if i==0 else f'=IF(DAY({L(c)}$5)=1,MONTH({L(c)}$5)&"月","")'); m.font=Font(name=F,size=8)
        for r in range(6,21): ws.cell(r,c).border=Border(left=Side(style="hair",color="D9D9D9"),bottom=Side(style="hair",color="D9D9D9"))
    ws.cell(4,9,'=MONTH(I$5)&"月"')
    g="I6:AX20"
    ws.conditional_formatting.add(g,FormulaRule(formula=['AND(I$5>=$D6,I$5<$D6+ROUND($E6*$G6,0))'],fill=PatternFill("solid",fgColor="2E75B6")))
    ws.conditional_formatting.add(g,FormulaRule(formula=['AND(I$5>=$D6,I$5<=$F6)'],fill=PatternFill("solid",fgColor="BDD7EE")))
    ws.conditional_formatting.add("I5:AX20",FormulaRule(formula=['I$5=TODAY()'],border=Border(left=Side(style="medium",color="C00000"),right=Side(style="medium",color="C00000"))))
    ws.conditional_formatting.add("G6:G20",DataBarRule(start_type="num",start_value=0,end_type="num",end_value=1,color="70AD47"))
    ws.conditional_formatting.add("H6:H20",CellIsRule(operator="equal",formula=['"延期"'],font=Font(color="C00000",bold=True)))
    ws.freeze_panes="I6"
    wb.save("03_项目进度甘特图.xlsx")

# 4 记账
def jizhang():
    wb=Workbook(); ws=wb.active; ws.title="流水"
    title(ws,"个人/家庭记账本",7,"每笔一行，分类用下拉选择；「月度汇总」表自动统计并出图")
    header(ws,3,["日期","收/支","分类","金额","账户","备注","月份"],[12,8,12,12,12,24,10])
    cats=["工资","副业","理财","餐饮","交通","购物","住房","娱乐","医疗","教育","其他"]
    lst=wb.create_sheet("设置")
    lst["A1"]="分类"; 
    for i,c in enumerate(cats,2): lst.cell(i,1,c)
    lst["C1"]="账户"
    for i,c in enumerate(["微信","支付宝","银行卡","现金","信用卡"],2): lst.cell(i,3,c)
    dv1=DataValidation(type="list",formula1='"收入,支出"'); dv2=DataValidation(type="list",formula1="=设置!$A$2:$A$12"); dv3=DataValidation(type="list",formula1="=设置!$C$2:$C$6")
    for d in (dv1,dv2,dv3): ws.add_data_validation(d)
    dv1.add("B4:B503"); dv2.add("C4:C503"); dv3.add("E4:E503")
    sample=[(dt.date(2026,10,1),"收入","工资",12000,"银行卡","10月工资"),(dt.date(2026,10,2),"支出","住房",3500,"支付宝","房租"),(dt.date(2026,10,3),"支出","餐饮",68,"微信","午饭"),(dt.date(2026,10,5),"支出","交通",120,"支付宝","地铁充值"),(dt.date(2026,10,6),"支出","购物",459,"信用卡","衣服"),(dt.date(2026,10,8),"收入","副业",800,"微信","接单")]
    for k in range(500):
        r=4+k
        if k<len(sample):
            for j,v in enumerate(sample[k],1): ws.cell(r,j,v)
        ws.cell(r,1).number_format="yyyy-mm-dd"; ws.cell(r,4).number_format="#,##0.00"
        ws.cell(r,7,f'=IF(A{r}="","",TEXT(A{r},"yyyy-mm"))')
    body(ws,4,40,1,7)
    ws.conditional_formatting.add("A4:G503",FormulaRule(formula=['$B4="收入"'],font=Font(color="548235")))
    ws.freeze_panes="A4"
    s=wb.create_sheet("月度汇总",1)
    title(s,"月度汇总",4)
    s["A3"]="统计月份"; s["B3"]="2026-10"; s["B3"].fill=PatternFill("solid",fgColor="FFF2CC")
    s["A4"]="总收入"; s["B4"]='=SUMIFS(流水!D:D,流水!G:G,$B$3,流水!B:B,"收入")'
    s["A5"]="总支出"; s["B5"]='=SUMIFS(流水!D:D,流水!G:G,$B$3,流水!B:B,"支出")'
    s["A6"]="结余"; s["B6"]="=B4-B5"; s["A7"]="储蓄率"; s["B7"]='=IFERROR(B6/B4,0)'; s["B7"].number_format="0.0%"
    for a in("B4","B5","B6"): s[a].number_format="#,##0.00"
    header(s,9,["支出分类","金额","占比"],[14,14,10])
    for i,c in enumerate(cats[3:],10):
        s.cell(i,1,c); s.cell(i,2,f'=SUMIFS(流水!D:D,流水!G:G,$B$3,流水!C:C,A{i},流水!B:B,"支出")'); s.cell(i,2).number_format="#,##0.00"
        s.cell(i,3,f"=IFERROR(B{i}/$B$5,0)"); s.cell(i,3).number_format="0.0%"
    body(s,10,17,1,3)
    pie=PieChart(); pie.title="支出构成"; pie.add_data(Reference(s,min_col=2,min_row=9,max_row=17),titles_from_data=True); pie.set_categories(Reference(s,min_col=1,min_row=10,max_row=17)); pie.height=8; pie.width=12
    s.add_chart(pie,"E3")
    wb.save("04_个人家庭记账本.xlsx")

# 5 库存
def kucun():
    wb=Workbook(); ws=wb.active; ws.title="库存总览"
    title(ws,"进销存库存管理表",9,"在「出入库记录」登记每笔进出，总览自动算当前库存；低于安全库存自动标红")
    header(ws,3,["商品编码","商品名称","规格","单位","期初库存","累计入库","累计出库","当前库存","安全库存"],[11,16,10,6,10,10,10,10,10])
    items=[("P001","A4打印纸","70g","箱",20,10),("P002","中性笔","0.5mm","盒",50,20),("P003","文件夹","A4","个",100,30),("P004","订书钉","24/6","盒",30,10)]
    rec=wb.create_sheet("出入库记录")
    for k in range(100):
        r=4+k
        if k<4:
            for j,v in enumerate(items[k][:5],1): ws.cell(r,j,v)
            ws.cell(r,9,items[k][5])
        ws.cell(r,6,f'=IF(A{r}="","",SUMIFS(出入库记录!$E:$E,出入库记录!$B:$B,A{r},出入库记录!$D:$D,"入库"))')
        ws.cell(r,7,f'=IF(A{r}="","",SUMIFS(出入库记录!$E:$E,出入库记录!$B:$B,A{r},出入库记录!$D:$D,"出库"))')
        ws.cell(r,8,f'=IF(A{r}="","",E{r}+F{r}-G{r})')
    body(ws,4,30,1,9)
    ws.conditional_formatting.add("A4:I103",FormulaRule(formula=['AND($A4<>"",$H4<$I4)'],fill=PatternFill("solid",fgColor="F8CBAD"),font=Font(color="9C0006",bold=True)))
    ws.freeze_panes="C4"
    title(rec,"出入库记录",7)
    header(rec,3,["日期","商品编码","商品名称","类型","数量","经手人","备注"],[12,11,16,8,8,10,20])
    dv=DataValidation(type="list",formula1='"入库,出库"'); rec.add_data_validation(dv); dv.add("D4:D2003")
    dv2=DataValidation(type="list",formula1="=库存总览!$A$4:$A$103"); rec.add_data_validation(dv2); dv2.add("B4:B2003")
    smp=[(dt.date(2026,10,8),"P001","入库",5,"张三"),(dt.date(2026,10,9),"P002","出库",40,"李四"),(dt.date(2026,10,9),"P003","出库",20,"王五")]
    for k in range(2000):
        r=4+k
        if k<3:
            d,cde,t,q,p=smp[k]; rec.cell(r,1,d); rec.cell(r,2,cde); rec.cell(r,4,t); rec.cell(r,5,q); rec.cell(r,6,p)
        rec.cell(r,1).number_format="yyyy-mm-dd"
        if k<300: rec.cell(r,3,f'=IFERROR(INDEX(库存总览!$B:$B,MATCH(B{r},库存总览!$A:$A,0)),"")')
    body(rec,4,40,1,7)
    rec.freeze_panes="A4"
    wb.save("05_进销存库存管理.xlsx")

for f in (kaoqin,gongzi,gantt,jizhang,kucun): f()
print("ok")
