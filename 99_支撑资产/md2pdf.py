# -*- coding: utf-8 -*-
import os, re
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

# 字体：优先等线（DengXian），失败时回退到系统可用的中文字体。
# 若全部缺失，脚本会给出明确提示，而不是抛出难以理解的报错。
FONT_CANDIDATES=[
    # (常规, 粗体, 细体, 字体所在目录)  —— 等线（Windows 10+）
    ('Deng.ttf','Dengb.ttf','Dengl.ttf','C:/Windows/Fonts/'),
    # 微软雅黑（Windows）
    ('msyh.ttc','msyhbd.ttc','msyhl.ttc','C:/Windows/Fonts/'),
    # 苹方 / 华文黑体（macOS）
    ('PingFang.ttc','PingFang.ttc','PingFang.ttc','/System/Library/Fonts/'),
    # Noto Sans CJK（Linux）
    ('NotoSansCJK-Regular.ttc','NotoSansCJK-Bold.ttc','NotoSansCJK-Regular.ttc','/usr/share/fonts/opentype/noto/'),
    ('NotoSansSC-Regular.otf','NotoSansSC-Bold.otf','NotoSansSC-Regular.otf','/usr/share/fonts/opentype/noto/'),
]
def _register_fonts():
    for reg,bold,light,d in FONT_CANDIDATES:
        if all(os.path.exists(d+f) for f in (reg,bold,light)):
            pdfmetrics.registerFont(TTFont('Deng', d+reg))
            pdfmetrics.registerFont(TTFont('Deng-Bold', d+bold))
            pdfmetrics.registerFont(TTFont('Deng-Light', d+light))
            pdfmetrics.registerFontFamily('Deng', normal='Deng', bold='Deng-Bold', italic='Deng', boldItalic='Deng-Bold')
            return
    raise SystemExit('未找到可用的中文字体。请安装「等线」「微软雅黑」或 Noto Sans CJK，'
                     '或修改 md2pdf.py 顶部的 FONT_CANDIDATES 指向本机字体文件。')
_register_fonts()

NAVY=colors.HexColor('#152238'); NAVY2=colors.HexColor('#1F3864')
ORANGE=colors.HexColor('#F0A030'); ACCENT=colors.HexColor('#E8833A')
INK=colors.HexColor('#1A1A1A'); MUTED=colors.HexColor('#5A6472')
LINE=colors.HexColor('#D9DEE7'); BG_ALT=colors.HexColor('#F4F6FA'); HDR_BG=colors.HexColor('#1F3864')

PAGE_W,PAGE_H=A4; M=15*mm; AVAIL=PAGE_W-2*M

S={
 'ct':ParagraphStyle('ct',fontName='Deng-Bold',fontSize=30,leading=42,textColor=colors.white),
 'cs':ParagraphStyle('cs',fontName='Deng',fontSize=12.5,leading=20,textColor=colors.HexColor('#C9D4E5')),
 'ctag':ParagraphStyle('ctag',fontName='Deng',fontSize=9.5,leading=15,textColor=colors.HexColor('#8FA3BF')),
 'h1':ParagraphStyle('h1',fontName='Deng-Bold',fontSize=16,leading=22,spaceBefore=10,spaceAfter=6,textColor=NAVY),
 'h2':ParagraphStyle('h2',fontName='Deng-Bold',fontSize=13,leading=18,spaceBefore=8,spaceAfter=4,textColor=NAVY2),
 'h3':ParagraphStyle('h3',fontName='Deng-Bold',fontSize=11.5,leading=15,spaceBefore=6,spaceAfter=3,textColor=ACCENT),
 'h4':ParagraphStyle('h4',fontName='Deng-Bold',fontSize=10.5,leading=14,spaceBefore=4,spaceAfter=2,textColor=colors.HexColor('#3A4A6B')),
 'body':ParagraphStyle('body',fontName='Deng',fontSize=9.8,leading=15,textColor=INK),
 'code':ParagraphStyle('code',fontName='Deng',fontSize=8,leading=11.5,leftIndent=6,spaceBefore=2,spaceAfter=2,textColor=colors.HexColor('#33415C')),
 'idx':ParagraphStyle('idx',fontName='Deng',fontSize=9,leading=13.5,leftIndent=2,textColor=INK),
 'idxh':ParagraphStyle('idxh',fontName='Deng-Bold',fontSize=11,leading=16,spaceBefore=6,spaceAfter=2,textColor=NAVY),
 'sec':ParagraphStyle('sec',fontName='Deng-Bold',fontSize=13.5,leading=18,textColor=colors.white),
}

def esc(s): return s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
def norm(s): return s.replace('\u2212','-').replace('\u2013','-').replace('\u2014','-').replace('\u2011','-')
def inline(md):
    md=esc(norm(md)); md=re.sub(r'\*\*(.+?)\*\*',r'<b>\1</b>',md); md=re.sub(r'`(.+?)`',r'\1',md); return md
def vis_w(s): return sum(2 if ord(ch)>127 else 1 for ch in str(s))
def col_widths(rows):
    ncol=len(rows[0]) if rows else 1; maxl=[8]*ncol
    for r in rows:
        for j in range(ncol):
            c=r[j] if j<len(r) else ''; maxl[j]=max(maxl[j],min(vis_w(c),46))
    total=sum(maxl) or 1; widths=[max(AVAIL*l/total,26) for l in maxl]
    s=sum(widths)
    if s>AVAIL: widths=[w*AVAIL/s for w in widths]
    return widths
def parse_table_rows(lines):
    rows=[]
    for ln in lines:
        st=ln.strip()
        if st.startswith('|'): rows.append([c.strip() for c in st.strip('|').split('|')])
    return [r for r in rows if not all(re.fullmatch(r':?-{2,}:?',(c or '-')) for c in r)]

def render(md,story):
    lines=md.split('\n'); i=0; n=len(lines)
    while i<n:
        ln=lines[i].rstrip()
        if ln.startswith('```'):
            buf=[]; i+=1
            while i<n and not lines[i].startswith('```'):
                buf.append(lines[i]); i+=1
            i+=1
            code='\n'.join(buf)
            if code.strip():
                story.append(Paragraph(esc(norm(code)).replace('\n','<br/>'), S['code']))
                story.append(Spacer(1,3*mm))
            continue
        if ln.startswith('####'): story.append(Paragraph(inline(ln[4:].strip()),S['h4'])); i+=1; continue
        if ln.startswith('###'): story.append(Paragraph(inline(ln[3:].strip()),S['h3'])); i+=1; continue
        if ln.startswith('##'): story.append(Paragraph(inline(ln[2:].strip()),S['h2'])); i+=1; continue
        if ln.startswith('#'): story.append(Paragraph(inline(ln[1:].strip()),S['h1'])); i+=1; continue
        if ln.strip()=='---': story.append(Spacer(1,3*mm)); i+=1; continue
        if ln.strip().startswith('|'):
            block=[]
            while i<n and lines[i].strip().startswith('|'):
                block.append(lines[i]); i+=1
            rows=parse_table_rows(block)
            if rows:
                ncol=max(len(r) for r in rows); rows=[r+['']*(ncol-len(r)) for r in rows]
                data=[[Paragraph(inline(c),S['body']) for c in r] for r in rows]
                t=Table(data,colWidths=col_widths(rows),hAlign='LEFT',repeatRows=1)
                t.setStyle(TableStyle([
                    ('VALIGN',(0,0),(-1,-1),'TOP'),
                    ('BACKGROUND',(0,0),(-1,0),HDR_BG),('TEXTCOLOR',(0,0),(-1,0),colors.white),
                    ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,BG_ALT]),
                    ('GRID',(0,0),(-1,-1),0.4,LINE),
                    ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
                    ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
                ]))
                story.append(t); story.append(Spacer(1,4*mm)); continue
        if ln.strip().startswith('- '):
            buf=[]
            while i<n and lines[i].strip().startswith('- '):
                buf.append(lines[i].strip()[2:]); i+=1
            for item in buf: story.append(Paragraph('•&nbsp;'+inline(item),S['body']))
            story.append(Spacer(1,1.5*mm)); continue
        m=re.match(r'^(\d+)\.\s+(.*)',ln.strip())
        if m:
            buf=[]
            while i<n:
                mm2=re.match(r'^(\d+)\.\s+(.*)',lines[i].strip())
                if not mm2: break
                buf.append(mm2.group(2)); i+=1
            for k,item in enumerate(buf,1): story.append(Paragraph(f'<font color="#E8833A"><b>{k}.</b></font>&nbsp;'+inline(item),S['body']))
            story.append(Spacer(1,1.5*mm)); continue
        if ln.strip()=='': i+=1; continue
        story.append(Paragraph(inline(ln.strip()),S['body'])); i+=1

def secbar(text):
    t=Table([[Paragraph(text,S['sec'])]],colWidths=[AVAIL])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),NAVY2),('LEFTPADDING',(0,0),(-1,-1),10),
        ('RIGHTPADDING',(0,0),(-1,-1),10),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
        ('LINEBELOW',(0,0),(-1,-1),2.5,ORANGE),
    ]))
    return t

ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
out=os.path.join(ROOT,'Amazon_运营SOP_全库_v1.0.pdf')
groups=[
 ('00_总纲',['00-08_使用指南与交付说明.md','00-00_架构设计文档.md','00-01_知识库使用说明与维护规范.md','00-02_证据等级与真实性规范.md','00-03_术语表.md','00-04_数据来源与工具清单.md','00-05_运营日历总览.md','00-06_缺口审计报告.md','00-07_反向测试报告.md']),
]
for d in ['01_平台规则与账户','02_市场与选品','03_产品开发与供应链','04_Listing与内容','05_流量与广告','06_转化与客户体验','07_库存与物流','08_数据与诊断','09_利润与财务','10_风控与异常','11_运营节奏与复盘','12_团队与培训']:
    p=os.path.join(ROOT,d); groups.append((d,sorted([f for f in os.listdir(p) if f.endswith('.md')])))
asset=os.path.join(ROOT,'99_支撑资产'); af=[]
for fn in ['从0到1实操手册_通用案例.md','模板/运营工作日志模板.md','模板/SOP_17节模板.md','资料库/P0_资料研究卡片_001.md']:
    fp=os.path.join(asset,fn)
    if os.path.exists(fp): af.append(fn)
groups.append(('99_支撑资产',af))
files=[]
for d,fns in groups:
    p=os.path.join(ROOT,d) if d!='99_支撑资产' else asset
    for fn in fns:
        fp=os.path.join(p,fn)
        if os.path.exists(fp): files.append((d+' / '+fn.replace('.md',''),fp))

def footer(canvas,doc):
    pn=canvas.getPageNumber()
    if pn==1:
        canvas.setFillColor(NAVY); canvas.rect(0,0,PAGE_W,PAGE_H,fill=1,stroke=0)
        canvas.setFillColor(ORANGE); canvas.rect(0,PAGE_H*0.585,PAGE_W,2.2*mm,fill=1,stroke=0)
        canvas.setFillColor(NAVY2); canvas.rect(0,PAGE_H*0.15,PAGE_W,0.6*mm,fill=1,stroke=0)
    else:
        canvas.setStrokeColor(LINE); canvas.setLineWidth(0.6); canvas.line(M,12.5*mm,PAGE_W-M,12.5*mm)
        canvas.setFont('Deng',8); canvas.setFillColor(MUTED)
        canvas.drawString(M,9*mm,'Amazon 全流程运营 SOP 全库 · v1.0')
        canvas.drawRightString(PAGE_W-M,9*mm,'第 %d 页'%pn)

doc=SimpleDocTemplate(out,pagesize=A4,leftMargin=M,rightMargin=M,topMargin=16*mm,bottomMargin=16*mm,title='Amazon 全流程运营 SOP 全库',author='SOP 项目组')
story=[]
story.append(Spacer(1,38*mm))
story.append(Paragraph('Amazon 全流程运营',S['ct']))
story.append(Paragraph('运营 SOP 全库',S['ct']))
story.append(Spacer(1,8*mm))
story.append(Paragraph('一套可执行、可验证、可长期维护的 Amazon 运营操作系统',S['cs']))
story.append(Spacer(1,30*mm))
story.append(Paragraph('v1.0  ·  2026-08-27  ·  96 份文档',S['ctag']))
story.append(Spacer(1,3*mm))
story.append(Paragraph('82 份业务 SOP ｜ 8 份总纲 ｜ 反向测试报告 ｜ 通用案例 0→1 实操手册 ｜ 3 个 Excel 计算器',S['ctag']))
story.append(PageBreak())
story.append(Paragraph('文档目录',S['h1'])); story.append(Spacer(1,2*mm))
cur=None
for disp,fp in files:
    d=disp.split(' / ')[0]; fn=disp.split(' / ')[1]
    if d!=cur: cur=d; story.append(Paragraph(d,S['idxh']))
    story.append(Paragraph('·&nbsp;'+fn,S['idx']))
story.append(PageBreak())
for disp,fp in files:
    story.append(secbar(disp)); story.append(Spacer(1,4*mm))
    with open(fp,encoding='utf-8') as f: render(f.read(),story)
    story.append(PageBreak())
doc.build(story,onFirstPage=footer,onLaterPages=footer)
print('PDF OK:',out,'docs:',len(files))
