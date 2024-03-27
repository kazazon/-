import os
import pandas as pd
import hanlp

# 加载预训练模型
tokenizer = hanlp.load(hanlp.pretrained.tok.FINE_ELECTRA_SMALL_ZH)
tagger = hanlp.load(hanlp.pretrained.pos.CTB9_POS_ELECTRA_SMALL)
ner_tagger = hanlp.load(hanlp.pretrained.ner.MSRA_NER_ELECTRA_SMALL_ZH)
dep_parser = hanlp.load(hanlp.pretrained.dep.CTB9_UDC_ELECTRA_SMALL)

import pandas as pd

def process_txt(input_filename, output_excel_filename):
    with open(input_filename, "r", encoding="utf-8") as input_file:
        lines = input_file.readlines()

    data = []
    for i, line in enumerate(lines):
        original_text = line.strip()

        # 使用tokenizer进行分词
        tokens = tokenizer(original_text)

        # 使用tagger进行词性标注
        pos_tags = tagger(tokens)

        # 使用ner_tagger进行命名实体识别
        ner_tags = ner_tagger(tokens)

        # 使用dep_parser进行依存句法分析
        dep_relations = dep_parser(tokens)

        processed_text = " ".join([f"{token}({pos})" for token, pos in zip(tokens, pos_tags)])

        # 转换命名实体识别结果为字符串
        ner_tags_str = " ".join([f"{entity[0]}({entity[1]})" for entity in ner_tags])

        dependency_relation = " ".join([f"{rel.id}({rel.deprel}){idx}" for idx, rel in enumerate(dep_relations)])

        data.append([i + 1, original_text, processed_text, ner_tags_str, dependency_relation])

    df = pd.DataFrame(data, columns=['行号', '原始文本', '分词', '命名实体识别', '依存句法关系'])
    df.to_excel(output_excel_filename, index=False)
    print('数据已保存到 Excel 文件：', output_excel_filename)


# 输入TXT文件所在文件夹路径
txt_folder = '.'

# 递归地获取文件夹下的所有TXT文件
txt_files = []
for foldername, subfolders, filenames in os.walk(txt_folder):
    for filename in filenames:
        if filename.endswith('.txt'):
            txt_files.append(os.path.join(foldername, filename))

# 处理每个TXT文件
for idx, txt_file in enumerate(txt_files, start=1):
    # 构建完整的文件路径
    input_txt_file = txt_file
    # 生成输出Excel文件名
    output_excel_file = os.path.splitext(txt_file)[0] + '.xlsx'

    # 检查是否已存在相同文件名的Excel文件，如果不存在则运行主程序
    if not os.path.exists(output_excel_file):
        # 输出检测到的TXT文件名和输出的Excel文件名
        print(f"正在处理TXT文件 {idx}/{len(txt_files)}: {input_txt_file}")
        print(f"生成的Excel文件：{output_excel_file}")
        # 调用处理函数
        process_txt(input_txt_file, output_excel_file)
        print(f"TXT文件 {idx}/{len(txt_files)} 处理完成。\n")
    else:
        print(f"跳过TXT文件 {idx}/{len(txt_files)}: {input_txt_file}，因为相同文件名的Excel文件已存在。\n")

input("Press Enter to exit...")
