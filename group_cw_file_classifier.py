from ntpath import dirname
import pandas as pd
import os
import shutil

def read_excel_data(file_name):
    """读取Excel文件并获取用户输入的组号"""
    df = pd.read_excel(file_name)
    default_groups = "1 16 33 50 25 22 5 10 49 32 17 9 23"
    
    print("\n选择操作：")
    print("1. 使用默认组号（1 16 33 50 25 22 5 10 49 32 17 9 23）")
    print("2. 手动输入组号")
    choice = input("请输入选择（1 或 2）：").strip()
    
    if choice == "1":
        group_numbers = default_groups.split()
    else:
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

def find_files(root_dir, id_list, df_data):
    """查找包含学生First Name的文件夹"""
    found_count = 0
    found_groups = set()
    
    # 获取每个ID对应的组号和First Name
    id_to_group = dict(zip(df_data['ID number'], df_data['Group']))
    id_to_name = dict(zip(df_data['ID number'], df_data['First name']))
    
    print("\n开始查找文件夹...")
    print("正在查找以下学生的提交：")
    for id_num in id_list:
        print(f"Group {id_to_group[id_num].split()[-1]}: {id_to_name[id_num]} (ID: {id_num})")
    
    print("\n开始搜索...")
    for root, dirs, files in os.walk(root_dir):
        for dir_name in dirs:
            # 尝试从文件夹名中匹配First Name
            for id_num in id_list:
                first_name = id_to_name[id_num]
                if dir_name.lower().split()[0] == first_name.lower():
                    group = id_to_group[id_num]
                    print(dir_name.lower())
                    print(first_name)
                    group_num = group.split()[-1]
                    if group not in found_groups:
                        found_groups.add(group)
                        source_path = os.path.join(root, dir_name)
                        dest_path = os.path.join(output_folder, f"{group}_{dir_name}")
                        # 如果目标文件夹已存在，先删除
                        if os.path.exists(dest_path):
                            shutil.rmtree(dest_path)
                        shutil.copytree(source_path, dest_path)
                        print(f"\n找到 Group {group_num} 的提交:")
                        print(f"- 学生: {first_name} (ID: {id_num})")
                        print(f"- 文件夹: {dir_name}")
                        print(f"- 路径: {source_path}")
                        found_count += 1
                        break

    # 检查哪些组未找到提交
    all_groups = set(df_data['Group'].unique())
    missing_groups = all_groups - found_groups
    if missing_groups:
        print("\n以下小组未找到提交：")
        for group in sorted(missing_groups):
            group_members = df_data[df_data['Group'] == group]
            print(f"\nGroup {group.split()[-1]}:")
            for _, member in group_members.iterrows():
                print(f"- {member['First name']} (ID: {member['ID number']})")
    
    print(f"\n总共找到 {found_count} 个小组的提交")
    return found_count

def main():
    # 全局变量
    global output_folder
    excel_file = 'CPT202_Group Forming_Result.xlsx'
    output_folder = "selected_files"
    source_folder = "CPT202-2425-S2-Assignment 2 --- Group Report-80949"
    
    # 读取数据并处理
    filtered_data = read_excel_data(excel_file)
    setup_output_folder(output_folder)

    # 获取ID列表并查找文件夹
    id_list = filtered_data['ID number'].tolist()
    find_files(source_folder, id_list, filtered_data)

if __name__ == "__main__":
    main()