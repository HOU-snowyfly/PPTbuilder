"""Editable, colorful classroom comics and diagrams for the revised slides."""
import math
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement

INK='203354'; TEAL='008C82'; PURPLE='7952D5'; ORANGE='FFBE45'
CORAL='EB5969'; MINT='E4F8EF'; LILAC='F1EAFF'; CREAM='FFF2D3'
SKIN='FFD3B3'; WHITE='FFFFFF'; GREY='DDE5EF'
FONT='Noto Sans CJK SC'

class Canvas:
    def __init__(self,slide):self.s=slide
    def shape(self,kind,x,y,w,h,fill=WHITE,stroke=None,lw=1.5):
        o=self.s.shapes.add_shape(kind,Inches(x),Inches(y),Inches(w),Inches(h))
        o.fill.solid();o.fill.fore_color.rgb=RGBColor.from_string(fill)
        o._element.spPr.append(OxmlElement('a:effectLst'))
        for effect in o._element.xpath('.//a:effectRef'):effect.set('idx','0')
        if stroke:o.line.color.rgb=RGBColor.from_string(stroke);o.line.width=Pt(lw)
        else:o.line.fill.background()
        if kind==MSO_SHAPE.ROUNDED_RECTANGLE:o.adjustments[0]=.10
        return o
    def box(self,x,y,w,h,fill=WHITE,stroke=None):return self.shape(MSO_SHAPE.ROUNDED_RECTANGLE,x,y,w,h,fill,stroke)
    def oval(self,x,y,w,h,fill=WHITE,stroke=None):return self.shape(MSO_SHAPE.OVAL,x,y,w,h,fill,stroke)
    def line(self,x1,y1,x2,y2,c=INK,lw=2):
        o=self.s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
        o.line.color.rgb=RGBColor.from_string(c);o.line.width=Pt(lw)
        o._element.spPr.append(OxmlElement('a:effectLst'))
        for effect in o._element.xpath('.//a:effectRef'):effect.set('idx','0')
        return o
    def text(self,t,x,y,w,h,size=20,c=INK,bold=False,center=False):
        o=self.s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=o.text_frame
        tf.word_wrap=True;tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=Pt(0)
        tf.vertical_anchor=MSO_ANCHOR.MIDDLE
        for i,row in enumerate(t.split('\n')):
            p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text=row;p.line_spacing=1.05;p.space_after=Pt(0)
            p.alignment=PP_ALIGN.CENTER if center else PP_ALIGN.LEFT
            for r in p.runs:r.font.name=FONT;r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=RGBColor.from_string(c)
        return o
    def arrow(self,x,y,w=.45,c=CORAL):return self.shape(MSO_SHAPE.CHEVRON,x,y,w,.30,c)
    def badge(self,t,x,y,w,c=TEAL):self.box(x,y,w,.36,c);self.text(t,x+.07,y+.02,w-.14,.30,12,WHITE,True,True)
    def bubble(self,t,x,y,w,h=1.0,size=19):
        self.box(x,y,w,h,WHITE)
        tip=self.shape(MSO_SHAPE.ISOSCELES_TRIANGLE,x+.25,y+h-.01,.28,.24,WHITE);tip.rotation=180
        self.text(t,x+.16,y+.08,w-.32,h-.16,size,INK,True)
    def sparkle(self,x,y,w=.22,c=ORANGE):self.shape(MSO_SHAPE.STAR_4_POINT,x,y,w,w,c)
    def pill(self,x,y,w=1,h=.34,c=CORAL):
        self.box(x,y,w,h,c);self.box(x+.04,y+.04,w*.44,h-.08,WHITE)
    def medbox(self,x,y,w=1,h=1.2,c=TEAL,label='药盒'):
        self.box(x,y,w,h,WHITE,c);self.box(x+.05,y+.05,w-.10,.22,c)
        self.line(x+w*.5,y+.40,x+w*.5,y+.70,c,3);self.line(x+w*.35,y+.55,x+w*.65,y+.55,c,3)
        self.text(label,x+.05,y+h-.34,w-.10,.25,12,c,True,True)
    def student(self,x,y,k=.8,mood='happy',roommate=False):
        # Same student and hoodie across scenes; the roommate wears purple.
        def bx(a,b,w,h,c):return self.box(x+a*k,y+b*k,w*k,h*k,c)
        def ov(a,b,w,h,c):return self.oval(x+a*k,y+b*k,w*k,h*k,c)
        def ln(a,b,z,q,c=INK,lw=2):return self.line(x+a*k,y+b*k,x+z*k,y+q*k,c,lw)
        shirt=PURPLE if roommate else ORANGE
        bx(.04,1.04,1.53,1.03,shirt);bx(.62,.93,.32,.29,SKIN)
        ov(.21,.05,1.16,1.15,INK);ov(.24,.18,1.10,1.04,SKIN)
        ov(.10,.63,.22,.22,SKIN);ov(1.25,.63,.22,.22,SKIN)
        ov(.17,.02,.76,.37,INK);ov(.64,.10,.76,.28,INK)
        ov(.37,.79,.15,.13,'F49B98');ov(1.03,.79,.15,.13,'F49B98')
        for a in [.44,.94]:
            if mood=='sleepy':ln(a,.64,a+.15,.68,lw=2.6)
            elif mood=='thinking':ov(a,.63,.085,.105,INK)
            else:ln(a,.67,a+.07,.60,lw=2.5);ln(a+.07,.60,a+.14,.67,lw=2.5)
        ln(.80,.65,.83,.80,lw=1.4);ln(.83,.80,.77,.80,lw=1.4)
        if mood=='sleepy':ov(.73,.91,.15,.18,INK)
        elif mood=='thinking':ln(.66,.98,.87,.98,lw=2)
        else:ov(.64,.91,.33,.20,INK);ov(.62,.86,.37,.16,SKIN)
        ln(.48,1.14,.79,1.39,WHITE);ln(1.07,1.14,.79,1.39,WHITE)
        ln(.65,1.42,.65,1.66,WHITE,1.5);ln(.94,1.42,.94,1.66,WHITE,1.5)
        ov(1.10,1.59,.26,.27,SKIN)
        if mood=='sleepy':self.text('Z z',x+1.28*k,y-.11*k,.65*k,.38*k,19,PURPLE,True)
        if mood=='thinking':self.text('?',x+1.33*k,y-.10*k,.38*k,.45*k,28,PURPLE,True)
    def books(self,x,y,w=1.2):
        self.box(x,y,w,.20,TEAL);self.box(x+.08,y-.24,w,.20,PURPLE);self.box(x+.15,y-.48,w,.20,CORAL)
        for a,b in [(x+.09,y+.07),(x+.17,y-.17),(x+.24,y-.41)]:self.line(a,b,a+w-.28,b,WHITE,1.5)
    def brain(self,x,y,k=1):
        # Soft lobes and short curved-looking folds; icon, not anatomical scale.
        for a,b,w,h in [(0,.30,.65,.72),(.12,.05,.75,.76),(.58,0,.76,.80),(1.08,.22,.67,.77),(.60,.52,.77,.57)]:
            self.oval(x+a*k,y+b*k,w*k,h*k,'D8C4FF',PURPLE)
        for a,b,z,q in [(.31,.28,.48,.47),(.70,.24,.66,.57),(1.14,.38,1.30,.59),(.43,.76,.69,.85),(1.04,.80,1.18,.93)]:
            self.line(x+a*k,y+b*k,x+z*k,y+q*k,PURPLE,2)

def comic(c,d):
    exam=d['n']==15;fills=[CREAM,LILAC,MINT];colors=[TEAL,PURPLE,TEAL]
    for j,item in enumerate(d['items']):
        x=.80+j*4.04;c.box(x,2.10,3.80,3.98,fills[j]);c.badge(['01 情境','02 对话','03 悬念'][j],x+.23,2.29,1.20,colors[j])
        sentence=item.split('｜',1)[1];c.bubble(sentence,x+.23,2.91,3.34,1.04,19)
        c.student(x+.26,4.20,.70,'thinking' if j==2 else 'happy',j==1)
        if j==2:
            c.oval(x+2.05,4.42,.94,.94,WHITE,PURPLE);c.line(x+2.77,5.18,x+3.18,5.62,PURPLE,7)
            c.text('适用？' if exam else '成分？',x+2.14,4.69,.77,.30,16,PURPLE,True,True)
        elif exam and j==0:
            c.box(x+2.05,4.28,1.18,1.29,INK);c.box(x+2.14,4.39,1.00,.99,WHITE)
            c.text('快速\n提分？',x+2.20,4.56,.88,.62,19,CORAL,True,True);c.sparkle(x+3.27,4.15)
        else:c.medbox(x+2.14,4.37,1.05,1.26,PURPLE if j==1 else TEAL,'处方？' if exam else ('药盒B' if j==1 else '药盒A'))
        c.text(['先有一个需求','再来一个建议','判断停在这里'][j],x+.24,5.70,3.32,.24,14,colors[j],True)

def ingredient(c):
    for x,fill,name,sub in [(1.02,MINT,'药盒 A','单方退热止痛药'),(7.13,LILAC,'药盒 B','复方感冒药')]:
        c.box(x,2.12,5.20,3.19,fill);c.text(name,x+.27,2.31,2.15,.42,25,INK,True);c.text(sub,x+.27,2.80,4.6,.31,17)
        c.box(x+.24,3.31,4.72,.58,CORAL);c.text('对乙酰氨基酚',x+.34,3.38,4.52,.42,24,WHITE,True,True)
    c.text('只有这一种有效成分',1.30,4.34,4.6,.39,19,TEAL,True,True)
    for text,y,color in [('马来酸氯苯那敏',4.04,TEAL),('盐酸伪麻黄碱',4.61,PURPLE)]:
        c.box(7.37,y,4.72,.45,color);c.text(text,7.48,y+.04,4.50,.35,19,WHITE,True,True)
    c.line(5.98,3.60,6.43,3.60,CORAL,2.5);c.line(6.43,3.60,6.43,5.22,CORAL,2.5)
    c.line(7.37,3.60,6.90,3.60,CORAL,2.5);c.line(6.90,3.60,6.90,5.22,CORAL,2.5)
    c.line(6.43,5.22,6.67,5.40,CORAL,2.5);c.line(6.90,5.22,6.67,5.40,CORAL,2.5)
    c.box(4.30,5.40,4.75,.44,CREAM);c.text('同一成分，来自两盒',4.40,5.43,4.55,.36,22,CORAL,True,True)
    c.text('① 圈出重复　→　② 核规格与已服情况　→　③ 未确认，不自行叠加',1.1,5.97,11.1,.30,17,INK,True,True)

def timeline(c,d):
    for j,(time,label,color,fill) in enumerate([('18:00','已经服过',TEAL,MINT),('20:00','准备再加',CORAL,CREAM),('先补信息','核对后再决定',PURPLE,LILAC)]):
        x=.82+j*4.03;c.box(x,2.18,3.79,3.80,fill);c.badge(time,x+.24,2.38,1.43,color)
        if j<2:
            c.student(x+.33,3.08,.75,'happy' if j==0 else 'thinking');c.medbox(x+2.08,3.50,1.08,1.40,color,'已服' if j==0 else '待核')
        else:
            c.box(x+.41,3.10,2.96,1.68,WHITE)
            for k,t in enumerate(['药名＋成分','规格＋已服用量','时间＋其他用药']):c.text(t,x+.61,3.24+k*.45,2.55,.32,18,PURPLE,True)
        c.text(label,x+.24,5.11,3.31,.39,23,color,True,True)
        c.text(d['items'][j].split('｜',1)[1],x+.24,5.53,3.31,.34,14,INK,False,True)

def microbes(c):
    c.box(.85,2.08,5.63,4.02,MINT);c.box(6.83,2.08,5.63,4.02,LILAC)
    c.badge('细菌',1.12,2.31,1.03,TEAL);c.badge('病毒',7.10,2.31,1.03,PURPLE)
    # Bacterium: membrane, genetic material and dots representing ribosomes.
    c.oval(1.34,3.21,2.26,1.36,TEAL);c.oval(1.45,3.32,2.04,1.14,'B5EFDF')
    for j in range(12):
        a=j*math.tau/12;cx=2.47+1.14*math.cos(a);cy=3.89+.70*math.sin(a)
        c.line(cx,cy,cx+.22*math.cos(a),cy+.22*math.sin(a),TEAL,2)
    for j in range(5):c.line(1.85+j*.22,3.85+(j%2)*.15,2.07+j*.22,3.85+((j+1)%2)*.15,PURPLE,3)
    for a,b in [(1.72,3.58),(2.25,3.54),(2.97,3.66),(1.78,4.17),(2.61,4.23)]:c.oval(a,b,.11,.11,ORANGE)
    c.text('有细胞结构',3.81,3.48,2.27,.39,23,TEAL,True);c.text('在适宜条件下\n可自行繁殖',3.81,3.99,2.27,.61,17)
    c.text('某些细菌结构或过程\n可成为抗生素的靶点',1.15,5.11,4.99,.67,21,INK,True)
    # Virus: shell and genome, then a host cell rather than a free-standing cell.
    cx,cy=8.04,3.90
    for j in range(10):
        a=j*math.tau/10;x=cx+.65*math.cos(a);y=cy+.65*math.sin(a)
        c.line(cx+.43*math.cos(a),cy+.43*math.sin(a),x,y,PURPLE,2);c.oval(x-.06,y-.06,.12,.12,ORANGE)
    c.oval(cx-.48,cy-.48,.96,.96,PURPLE);c.oval(cx-.34,cy-.34,.68,.68,'D8C4FF')
    for j in range(3):c.line(cx-.18+j*.13,cy-.16,cx-.05+j*.13,cy+.15,CORAL,2)
    c.arrow(9.02,3.79,.40,PURPLE);c.oval(9.65,3.24,2.15,1.40,'DBF4ED',TEAL);c.oval(10.28,3.66,.61,.54,TEAL)
    c.text('宿主细胞',9.77,4.77,1.96,.30,17,TEAL,True,True)
    c.text('无细胞结构\n依赖宿主细胞复制',7.15,5.12,4.99,.66,21,INK,True)

def amr(c):
    c.text('2019年全球估计 · 2022年发表 · 单位：万人',.87,1.94,11.7,.34,16)
    for y,fill,color,title,body in [(2.55,MINT,TEAL,'相关死亡 · 495万人','实际耐药感染\n与“没有这类感染”比较'),(4.13,LILAC,PURPLE,'归因死亡 · 127万人','感染仍发生\n但假设细菌对药物敏感')]:
        c.box(.87,y,5.14,1.38,fill);c.text(title,1.13,y+.15,4.65,.39,23,color,True)
        c.text(body,1.13,y+.65,4.65,.56,17)
    x0,width=6.56,5.54
    for tick in range(0,701,100):
        x=x0+width*tick/700;c.line(x,2.73,x,5.36,GREY,.6);c.text(str(tick),x-.24,5.48,.48,.22,11,INK,False,True)
    for y,name,value,lo,hi,color in [(3.08,'相关',495,362,657,TEAL),(4.47,'归因',127,91.1,171,PURPLE)]:
        c.text(name,x0,y-.39,1.20,.28,15,color,True)
        c.box(x0,y,width*value/700,.40,color)
        yy=y+.20;lx=x0+width*lo/700;hx=x0+width*hi/700
        c.line(lx,yy,hx,yy,INK,2.2);c.line(lx,yy-.13,lx,yy+.13,INK,2.2);c.line(hx,yy-.13,hx,yy+.13,INK,2.2)
        c.oval(x0+width*value/700-.065,yy-.065,.13,.13,WHITE,INK)
        c.text(str(value),x0+width*value/700+.12,y-.39,.86,.31,20,color,True)
        c.text(f'95%不确定性区间：{lo:g}—{hi:g}',x0,y+.52,5.55,.27,13)
    c.line(x0,5.35,x0+width,5.35,INK,1)
    c.text('条形与圆点＝估计值；带端帽的细线＝95%不确定性区间',6.39,5.94,5.90,.23,11)
    c.text('两个比较问题，不能相加',1.13,5.76,4.65,.36,23,CORAL,True)

def histamine(c):
    c.box(.85,2.13,5.63,3.98,MINT);c.box(6.83,2.13,5.63,3.98,LILAC)
    c.text('鼻部：参与过敏反应',1.13,2.34,4.99,.44,25,TEAL,True)
    c.student(1.28,3.04,.62,'neutral');c.bubble('阿嚏！',2.61,3.11,1.48,.62,21)
    c.oval(4.65,3.15,.65,.65,ORANGE);c.text('组胺',5.36,3.28,.86,.32,16,INK,True,True)
    signal=c.arrow(4.78,3.92,.34,ORANGE);signal.rotation=90
    c.arrow(2.84,4.44,.39,TEAL);c.box(3.43,4.26,2.30,.63,TEAL);c.text('H1 受体',3.57,4.34,2.02,.45,23,WHITE,True,True)
    c.pill(1.30,4.43,.91,.30);c.text('抗组胺药',1.13,4.91,1.78,.31,15,CORAL,True)
    c.text('影响受体作用 → 缓解部分症状',1.13,5.48,4.99,.40,19,TEAL,True)
    c.text('脑内：也参与维持清醒',7.11,2.34,4.99,.44,25,PURPLE,True)
    c.text('血脑屏障',8.36,3.00,1.56,.30,16,PURPLE,True,True)
    for y in [3.43,4.08]:c.box(9.06,y,.16,.43,PURPLE)
    c.pill(7.32,3.79,1.04,.35);c.arrow(8.53,3.80,.46,CORAL);c.arrow(9.39,3.80,.42,CORAL)
    c.brain(10.11,3.34,.92);c.text('脑内 H1 信号',9.95,4.53,2.12,.30,15,PURPLE,True,True)
    c.text('部分药物可进入脑内',7.11,5.08,4.99,.36,22,PURPLE,True)
    c.text('清醒相关信号受影响 → 可能嗜睡',7.11,5.57,4.99,.34,17)

def allergy_scene(c):
    for x,fill,title in [(.85,MINT,'鼻子舒服了'),(6.83,LILAC,'上课却犯困？')]:
        c.box(x,2.12,5.63,3.43,fill);c.text(title,x+.27,2.32,4.99,.43,27,INK,True)
    c.student(1.24,3.05,.91);c.student(7.25,3.05,.91,'sleepy')
    c.bubble('终于不用一直擦鼻涕！',3.10,3.05,2.96,.89,18)
    c.bubble('鼻子好了，\n脑子想“下课”……',9.10,3.05,2.97,.89,18)
    c.badge('鼻部',3.10,4.17,.74);c.text('组胺参与部分过敏反应',3.10,4.63,2.99,.30,14)
    c.text('部分症状得到缓解',3.10,5.03,2.99,.31,17,TEAL,True)
    c.badge('脑内',9.10,4.17,.74,PURPLE);c.text('组胺信号参与维持清醒',9.10,4.63,2.99,.30,14)
    c.text('部分药物入脑后可能致困',9.10,5.03,2.99,.31,15,PURPLE,True)
    c.books(7.40,4.80,.90)
    c.box(.90,5.79,11.49,.40,CREAM);c.text('选药别只看症状：查嗜睡提示，说清上课与通勤需求',1.10,5.83,11.09,.30,20,INK,True,True)

def quiz(c,d):
    # All options share size, weight and spacing; no visual answer cue.
    for j,item in enumerate(d['items']):
        y=2.12+j*1.25;c.box(.84,y,9.03,1.03,'F4F6FA')
        c.badge(item[0],1.08,y+.33,.59,INK);c.text(item[2:],1.96,y+.17,7.57,.70,23,INK,False)
    c.box(10.14,2.12,2.32,3.55,'F4F6FA');c.student(10.57,3.54,.75,'thinking')
    c.text('给建议点赞' if d.get('question_mode')=='多选' else '先选哪一步？',10.28,2.40,2.04,.40,20,PURPLE,True,True)
    c.badge('可多选' if d.get('question_mode')=='多选' else '单选',10.69,2.98,1.21,PURPLE)
    c.text('想一想，再举手',10.28,5.25,2.04,.25,13,INK,False,True)

def draw_revised_slide(slide,d):
    c=Canvas(slide);n=d['n']
    if d['layout']=='quiz':quiz(c,d)
    elif n==4:allergy_scene(c)
    elif n in [6,15]:comic(c,d)
    elif n==7:ingredient(c)
    elif n==8:timeline(c,d)
    elif n==11:microbes(c)
    elif n==13:amr(c)
    elif n==17:histamine(c)
    else:return False
    return True
