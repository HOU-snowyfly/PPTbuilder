"""Reference-checked tissue illustrations; code only places art and editable labels."""
from pathlib import Path
from pptx.util import Inches
from visuals import Canvas, INK, TEAL, PURPLE, WHITE

MEDICAL_DIR=Path(__file__).parent/'assets'/'medical'
MEDICAL_ARTWORKS=[
 ('nasal_histamine.png','鼻黏膜、组胺与感觉神经作用对象','2026-10-13'),
 ('brain_bbb_h1.png','脑微血管、血脑屏障与脑内H1作用对象','2026-10-13'),
]
MEDICAL_ARTWORKS += [('bacterial_cell.png','细菌细胞结构示意','2026-10-13'),('rhinovirus.png','鼻病毒衣壳与RNA示意','2026-10-13')]
MEDICAL_USAGE=[(17,'nasal_histamine.png','FULL','鼻部症状缓解路径'),(17,'brain_bbb_h1.png','FULL','可能嗜睡的脑内路径'),(11,'bacterial_cell.png','FULL','细菌有细胞结构'),(11,'rhinovirus.png','FULL','鼻病毒无细胞结构')]

def picture(c,name,x,y,w):
    pic=c.s.shapes.add_picture(str(MEDICAL_DIR/name),Inches(x),Inches(y),width=Inches(w))
    pic.name={'nasal_histamine.png':'MEDICAL_NASAL','brain_bbb_h1.png':'MEDICAL_BBB','bacterial_cell.png':'MEDICAL_BACTERIA','rhinovirus.png':'MEDICAL_RHINOVIRUS'}[name]
    pic._element.xpath('.//p:cNvPr')[0].set('descr','独立渲染的组织层次医学科普插图，结构与解释按参考资料核对；受体标记为功能位置，不是原子结构模型')
    return pic

def label(c,t,x,y,w,point,color=INK):
    c.line(x+w/2,y+.17,point[0],point[1],color,1.2)
    c.box(x-.04,y-.015,w+.08,.36,WHITE)
    c.text(t,x,y,w,.32,18,color,True,True)

def histamine_medical(slide,d):
    if d['n']!=17:return False
    c=Canvas(slide)
    c.box(.85,2.13,5.63,4.05,'F0F8F6');c.box(6.83,2.13,5.63,4.05,'F5F1FA')
    c.text('鼻部：缓解部分过敏症状',1.11,2.30,5.11,.44,25,TEAL,True)
    c.text('脑内：可能影响清醒状态',7.09,2.30,5.11,.44,25,PURPLE,True)
    picture(c,'nasal_histamine.png',1.12,2.77,5.10)
    picture(c,'brain_bbb_h1.png',7.09,2.77,5.10)
    # Leaders attach to tissue locations, not to a guessed molecular receptor shape.
    label(c,'组胺',3.10,2.97,.88,(3.33,4.57),TEAL)
    label(c,'H1受体',4.53,4.98,1.22,(3.80,4.79),PURPLE)
    label(c,'药物分子',7.12,2.97,1.40,(8.82,3.61),INK)
    label(c,'血脑屏障',10.62,3.91,1.43,(10.12,4.34),PURPLE)
    label(c,'脑内H1受体',10.49,5.05,1.66,(10.29,5.15),PURPLE)
    c.box(1.03,5.48,5.27,.66,WHITE)
    c.text('抗组胺药降低H1相关信号\n可减轻鼻痒、喷嚏和流涕',1.20,5.54,4.93,.54,19,TEAL,True)
    c.box(7.01,5.48,5.27,.66,WHITE)
    c.text('部分药物可进入脑内\n影响H1相关信号，可能嗜睡',7.18,5.54,4.93,.54,19,PURPLE,True)
    return True


def microbes_medical(slide,d):
    if d['n']!=11:return False
    c=Canvas(slide)
    c.box(.85,2.13,5.63,4.05,'F0F8F6');c.box(6.83,2.13,5.63,4.05,'F5F1FA')
    c.text('细菌：具有细胞结构',1.11,2.30,5.11,.44,25,TEAL,True)
    c.text('鼻病毒：没有细胞结构',7.09,2.30,5.11,.44,25,PURPLE,True)
    picture(c,'bacterial_cell.png',1.12,2.77,5.10)
    picture(c,'rhinovirus.png',7.09,2.77,5.10)
    label(c,'DNA',2.19,3.01,.77,(3.53,4.88),INK)
    label(c,'衣壳',10.87,3.05,.87,(10.94,3.82),PURPLE)
    label(c,'RNA',11.02,4.83,.87,(10.26,4.54),INK)
    c.box(1.03,5.48,5.27,.66,WHITE)
    c.text('某些细菌结构或过程\n可成为抗生素的作用对象',1.20,5.54,4.93,.54,19,TEAL,True)
    c.box(7.01,5.48,5.27,.66,WHITE)
    c.text('依赖宿主细胞复制\n抗生素没有相应治疗作用',7.18,5.54,4.93,.54,19,PURPLE,True)
    return True
