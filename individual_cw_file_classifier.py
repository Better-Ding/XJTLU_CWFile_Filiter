import pandas as pd
import os
import shutil

def read_excel_data(file_name):
    """读取Excel文件并获取用户输入的组号"""
    df = pd.read_excel(file_name)
    print("\n请输入要查询的组号（用空格分隔，例如：1 16 33 50 25）：")
    user_input = input().strip()
    group_numbers = user_input.split()
    
    # 转换为"Group "格式并筛选数据
    target_groups = [f'Group {num}' for num in group_numbers]
    filtered_data = df[df['Group'].isin(target_groups)]
    
    return filtered_data

def setup_output_folder(folder_name):
    """创建输出文件夹"""
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)

def find_pdf_files(root_dir, id_list, df_data):
    """查找并复制PDF文件，返回未找到文件的学生信息"""
    found_count = 0
    found_ids = set()
    
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith('.pdf'):
                for id_num in id_list:
                    if str(id_num) in file:
                        source_file = os.path.join(root, file)
                        dest_file = os.path.join(output_folder, file)
                        shutil.copy2(source_file, dest_file)
                        found_count += 1
                        found_ids.add(str(id_num))
    
    # 检查未找到的PDF文件
    missing_ids = set(str(id_num) for id_num in id_list) - found_ids
    if missing_ids:
        print("\n以下ID未找到对应的PDF文件(学生未提交或错误命名)：")
        for missing_id in missing_ids:
            student_info = df_data[df_data['ID number'] == int(missing_id)].iloc[0]
            print(student_info)
    
    return found_count

def main():
    # 全局变量
    global output_folder
    excel_file = 'CPT202_Group Forming_Result.xlsx'
    output_folder = "selected_pdfs"
    source_folder = "CPT202-2425-S2-Assignment 1 --- SoftwareRequirementReport-80948"
    
    # 读取数据并处理
    filtered_data = read_excel_data(excel_file)
    setup_output_folder(output_folder)
    
    # 获取ID列表并查找PDF
    id_list = filtered_data['ID number'].tolist()
    total_found = find_pdf_files(source_folder, id_list, filtered_data)
    
    # 输出统计信息
    print(f"\n总共: {len(filtered_data)}条数据")
    print(f"总共找到并复制了 {total_found} 个PDF文件")
    print(f"所有PDF文件已复制到 {output_folder} 文件夹")

if __name__ == "__main__":
    main()