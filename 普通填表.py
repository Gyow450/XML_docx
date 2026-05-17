import pandas as pd
from pandas import DataFrame
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Pt
from docx import Document
from PIL import Image
from io import BytesIO
from pathlib import Path
from datetime import datetime
import time
from src.LOG_DATA_STEEL import LOG_DICT
from src.interraction_terminal import set_argumments

def clean_and_save(doc_path, keyword="待删除段落"):
    """一轮清理：删除包含关键字的段落"""
    doc = Document(doc_path)
    
    for para in reversed(doc.paragraphs):
        # 删除条件：包含关键字 或 实质为空
        if keyword in para.text :
            para._element.getparent().remove(para._element)
    
    doc.save(doc_path)

def compress_image(path, max_width_px=1200, quality=85):
    """
    压缩图片并返回 BytesIO，供 InlineImage 直接使用
    """
    try:
        with Image.open(path) as img:
            # 如果尺寸过大，缩小
            if img.width > max_width_px:
                ratio = max_width_px / img.width
                new_size = (max_width_px, int(img.height * ratio))
                img = img.resize(new_size, Image.LANCZOS)
            
            # 压缩到内存 buffer
            buffer = BytesIO()
            if img.mode == 'RGBA':
                img.save(buffer, format='PNG')
            else:
                img.save(buffer, format='JPEG', quality=quality, optimize=True)
            buffer.seek(0)
            return buffer
    except Exception as e:
        print(f"压缩失败 {path}: {e}")
        return path  # 失败时返回原路径，让 InlineImage 自己处理

def make_data_in_list(tpl,df:DataFrame=None)->list:
    """编制填表的内容"""
    
    #   细节调整
    # df['探坑规格']=df['探坑规格（m）']
    # df['管道埋深']=df['管道埋深（m）'].astype(str)
    # df['检测日期']=df['检测日期'].dt.strftime('%Y年%m月%d日')
    # df['地表状况']=df['地形、地貌、地物描述']
    # df['近参比电位']=df['近参比电位（V，CSV）'].astype(str)+'V'
    # df['探坑坐标描述']='开挖点坐标（'+df['探坑坐标 X'].astype(str) +'，'+ df['探坑坐标 Y'].astype(str)+'）'
    # df['防腐层描述']='防腐层'+df['防腐层破损情况描述'].str.split('（').str[0]
    # df['管道描述']='管道'+df['管道本体腐蚀情况描述']
    # df['检验情况']=df[['探坑坐标描述','防腐层描述','管道描述']].values.tolist() 
    # df['检验结论']='防腐层外观评定为'+df['防腐层破损情况描述'].str.split('（').str[-1].str.replace('）','')
    # df['防腐层破损情况描述']=df['防腐层破损情况描述'].str.split('（').str[0]
    # cols_fcs=[f'FC1L{n}' for n in [0,3,6,9]]+[f'C1L{n}' for n in [0,3,6,9]]
    # df[cols_fcs]=df[cols_fcs].map(lambda x:f"{x:.2f}" if isinstance(x,float) else x)
    # df['探坑编号']=df['探坑编号'].fillna('1#')
    # df['环境条件']=df['环境条件'].fillna('晴')
    # #   开挖图片处理
    # for key,options in LOG_DICT['开挖勾选'].items():
    #     df[key]=df[key].apply(lambda x: ''.join([f"{option}（√）" if option in x.split(',') else f"{option}（ ）" for option in options ]))
    
    # df['竣工时间']=df['竣工时间'].apply(lambda x:x.strftime('%Y/%m/%d') if isinstance(x,datetime) else x)
    # df['投运时间']=df['投运时间'].apply(lambda x:x.strftime('%Y/%m/%d') if isinstance(x,datetime) else x)
    # drop_na_list=['起点','终点']
    # df[drop_na_list]=df[drop_na_list].fillna('/')
    # data_list=df.to_dict('records')
    figs=[]
    folder = Path(CONFIG['照片文件夹'])
    # for row in data_list:
    #     item=row.copy()
    #     for ex_name in ['.jpg','.jpeg','.png','.bmp','.gif']:
    #         path1:Path=(Path(CONFIG['照片文件夹'])/'防腐层图片'/row['自编号']).with_suffix(ex_name)
    #         if path1.exists():
    #             path_1=str(path1)
    #             break
    #     for ex_name in ['.jpg','.jpeg','.png','.bmp','.gif']:
    #         path2:Path=(Path(CONFIG['照片文件夹'])/'管道图片'/row['自编号']).with_suffix(ex_name)
    #         if path2.exists():
    #             path_2=str(path2)
    #             break
    #     # item['防腐层图片路径'] = str(path1)
    #     # item['管道图片路径'] = str(path2)
    #     item['防腐层图片'] = InlineImage(tpl, compress_image(path_1,800), width=Pt(120))  if path_1 else None
    #     item['管道图片'] = InlineImage(tpl, compress_image(path_2,800), width=Pt(120))  if path_2 else None
    #     items.append(item)
    for fig in folder.iterdir():
        figs.append({'路由图':InlineImage(tpl, compress_image(fig,800), width=Pt(600))})
    #   数据归总
    data_dict={'records':figs}
    # data_dict={'records':data_list}
    return data_dict

if __name__ == "__main__":
    # 读取excel的原始数据
    CONFIG=set_argumments([
        (2,'数据源文件','xlsx',r'E:\BaiduNetdiskDownload\管网\清单4-15.xlsx'),
        (2,'模板','docx',r'E:\BaiduNetdiskDownload\管网\合并过程\附件3：管道路由图.docx'),
        (0,'照片文件夹','',r'E:\BaiduNetdiskDownload\管网\路由图'),
        (0,'保存文件夹','',r'E:\BaiduNetdiskDownload\管网'),
    ])
    start = time.time()
    # df=pd.read_excel(Path(CONFIG['数据源文件']),sheet_name="Sheet5")
    # print(df.info())
    
    # 开启模板分析数据
    tpl = DocxTemplate(Path(CONFIG['模板']))
    data=make_data_in_list(tpl=tpl)
    
    # 填充并保存内容
    tpl.render(data)
    tpl.save(Path(CONFIG['保存文件夹'])/'raw.docx')
    clean_and_save(Path(CONFIG['保存文件夹'])/'raw.docx')
    print(f"\n总耗时: {time.time() - start:.2f} 秒")
    print('\nDone!')
