import re


def fix_dbc_orphan_comments(input_file, output_file=None):
    """
    修复DBC文件中孤立的注释（引用不存在信号的注释）
    
    参数:
        input_file: 输入的DBC文件路径
        output_file: 输出的DBC文件路径（如果为None，则覆盖原文件）
    """
    if output_file is None:
        output_file = input_file
    
    print(f"正在修复DBC文件: {input_file}")
    
    # 读取文件内容
    with open(input_file, 'r', encoding='gbk') as f:
        lines = f.readlines()
    
    # 收集所有已定义的信号 {(can_id, signal_name): True}
    defined_signals = {}
    
    # 第一遍扫描：收集所有BO_和SG_定义
    current_bo_id = None
    for line in lines:
        line_stripped = line.strip()
        
        # 匹配报文定义 BO_ <ID> <NAME>: ...
        bo_match = re.match(r'^BO_\s+(\d+)\s+', line_stripped)
        if bo_match:
            current_bo_id = int(bo_match.group(1))
            continue
        
        # 匹配信号定义 SG_ <NAME> : ...
        # 使用search而不是match，因为strip后SG_可能在开头
        sg_match = re.search(r'^SG_\s+(\S+)\s+:', line_stripped)
        if sg_match and current_bo_id is not None:
            signal_name = sg_match.group(1)
            defined_signals[(current_bo_id, signal_name)] = True
    
    print(f"✓ 共找到 {len(defined_signals)} 个已定义的信号")
    
    # 第二遍扫描：检查并移除孤立的注释
    fixed_lines = []
    removed_comments = []
    current_bo_id = None
    
    for line_num, line in enumerate(lines, 1):
        line_stripped = line.strip()
        
        # 更新当前报文ID
        bo_match = re.match(r'^BO_\s+(\d+)\s+', line_stripped)
        if bo_match:
            current_bo_id = int(bo_match.group(1))
            fixed_lines.append(line)
            continue
        
        # 检查CM_ SG_注释
        cm_match = re.match(r'^CM_\s+SG_\s+(\d+)\s+(\S+)\s+"(.*)";', line_stripped)
        if cm_match:
            comment_can_id = int(cm_match.group(1))
            comment_signal_name = cm_match.group(2)
            comment_text = cm_match.group(3)
            
            # 检查该信号是否存在
            if (comment_can_id, comment_signal_name) not in defined_signals:
                removed_comments.append({
                    'line': line_num,
                    'can_id': comment_can_id,
                    'signal': comment_signal_name,
                    'text': comment_text
                })
                # 跳过这行（不添加到fixed_lines）
                continue
        
        fixed_lines.append(line)
    
    # 输出修复报告
    print(f"\n{'='*70}")
    print(f"修复报告")
    print(f"{'='*70}")
    
    if removed_comments:
        print(f"\n❌ 发现 {len(removed_comments)} 条孤立注释（已移除）:\n")
        for item in removed_comments:
            print(f"  第{item['line']:4d}行: CM_ SG_ {item['can_id']:5d} {item['signal']:30s} \"{item['text'][:50]}\"")
        
        print(f"\n⚠️  这些注释引用的信号在DBC文件中不存在，已被自动移除。")
        print(f"   建议检查Excel源数据，确保信号名称正确且未被过滤。\n")
    else:
        print(f"\n✅ 未发现孤立注释，文件无需修复。\n")
    
    # 写入修复后的文件
    with open(output_file, 'w', encoding='gbk', newline='') as f:
        f.writelines(fixed_lines)
    
    print(f"✓ 修复后的文件已保存: {output_file}")
    print(f"{'='*70}\n")
    
    return len(removed_comments)


if __name__ == "__main__":
    # 修复当前打开的DBC文件
    dbc_file = r"c:\Users\kkk\Desktop\4266DBC\output2.dbc"
    fix_dbc_orphan_comments(dbc_file)
