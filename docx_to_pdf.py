import sys
from pathlib import Path

from fastprogress import progress_bar
import win32com.client as win32

from src.interraction_terminal import set_local_setting

def docx_transform(input_dir:Path|str,output_dir:Path|str)->None:
    """将输入文件夹中的docx文件批量转为pdf文件，另存至输出文件夹内"""
    # base_set={
    #     (0,'数据源','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网\管网PE待审核'),
    #     (0,'保存目标','',r'E:\BaiduSyncdisk\成渝特检\模板文件与生成程序\记录、报告生成\PE管\1400管网\管网PE待审核'),
    # }
    # SET_DICT = interraction_terminal.set_argumments(base_set)
    # INPUT_DIR = Path(SET_DICT['数据源'])
    # OUTPUT_DIR = Path(SET_DICT['保存目标'])
    
    input_dir=Path(input_dir)
    output_dir=Path(output_dir)
    docx_list = [p for p in input_dir.glob('*.docx') if not p.name.startswith('~$') ]
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # 不显示 Word 窗口，加快处理速度
    word.DisplayAlerts = 0  # 关闭警告信息
    # 全局关闭拼写/语法检查
    word.Options.CheckSpellingAsYouType = False   # 关闭实时拼写检查
    word.Options.CheckGrammarAsYouType = False    # 关闭实时语法检查
    word.Options.ContextualSpeller = False        # 关闭上下文拼写检查（Word 2010+）
    for docx_path in progress_bar(docx_list[:]):
        doc=word.Documents.Open(str(docx_path))
        output_file = docx_path.with_suffix('.pdf').name
        output_path = output_dir/output_file
        # mid_name=name.split('.')[0]
        # output_file = f"{source_path}\\{int(mid_name):03d}.pdf"
        # 移动到文档的末端
        selection = word.Selection
        selection.EndKey(6)  # 6 表示 wdStory，即整个文档
        # 更新所有域（页码）
        doc.Fields.Update()
        doc.SaveAs2(str(output_path), FileFormat=17)  
        print(f"文档已保存为：{output_file}")
        doc.Close(SaveChanges=False)
    word.Quit()

if __name__ == '__main__':
    # 预置参数
    # 类型码：0-文件夹，1-保存文件，2-打开文件，3-布尔，4-字符串
    setting_dict={
        '源文件夹':[0,'',''],
        '输出文件夹':[0,'',''],
    }
    
    # 设置路径等参数，优先读取之前保存的参数
    script_name=Path(__file__).resolve().stem
    CONFIG=set_local_setting(script_name,setting_dict)
    
    if CONFIG is None:
        sys.exit(0)
    
    INPUT_DIR = Path(CONFIG['源文件夹'])
    OUTPUT_DIR = Path(CONFIG['保存目标'])
    docx_transform(INPUT_DIR,OUTPUT_DIR)
    