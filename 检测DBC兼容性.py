import os
import re


def check_dbc_compatibility(filepath):
    """
    检测DBC文件是否符合CANdb++的格式要求
    """
    print(f"\n{'='*70}")
    print(f"DBC文件兼容性检测报告")
    print(f"{'='*70}")
    print(f"文件路径: {filepath}")
    print(f"{'='*70}\n")
    
    # 检查文件是否存在
    if not os.path.exists(filepath):
        print("❌ 错误: 文件不存在")
        return False
    
    # 检查文件大小
    file_size = os.path.getsize(filepath)
    print(f"✓ 文件大小: {file_size} 字节 ({file_size/1024:.2f} KB)")
    
    if file_size == 0:
        print("❌ 错误: 文件为空")
        return False
    
    # 尝试读取文件内容
    try:
        encoding_used = 'utf-8'
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            print("⚠ UTF-8解码失败，尝试GBK编码...")
            encoding_used = 'gbk'
            with open(filepath, 'r', encoding='gbk') as f:
                content = f.read()
        
        print(f"✓ 文件编码: {encoding_used.upper()}")
        
    except Exception as e:
        print(f"❌ 错误: 无法读取文件 - {e}")
        return False
    
    lines = content.split('\n')
    print(f"✓ 文件总行数: {len(lines)}")
    
    issues = []      # 严重问题（会导致CANdb++无法打开）
    warnings = []    # 警告（可能影响编辑但不影响打开）
    info = []        # 提示信息
    
    # ==================== 1. 检查VERSION关键字 ====================
    has_version = any(line.strip().startswith('VERSION') for line in lines[:10])
    if has_version:
        print("\n✓ [必需] VERSION声明: 存在")
    else:
        issues.append("缺少VERSION声明")
        print("\n❌ [必需] VERSION声明: 缺失")
    
    # ==================== 2. 检查NS_（命名空间）声明 ====================
    has_ns = any(line.strip().startswith('NS_') for line in lines[:50])
    if has_ns:
        print("✓ [必需] NS_命名空间声明: 存在")
    else:
        issues.append("缺少NS_命名空间声明")
        print("❌ [必需] NS_命名空间声明: 缺失")
    
    # ==================== 3. 检查BS_（波特率设置）====================
    has_bs = any(line.strip().startswith('BS_:') for line in lines)
    if has_bs:
        print("✓ [必需] BS_波特率设置: 存在")
    else:
        warnings.append("缺少BS_波特率设置（非致命，但建议添加）")
        print("⚠ [建议] BS_波特率设置: 缺失")
    
    # ==================== 4. 检查BU_（节点定义）====================
    bu_lines = [line for line in lines if line.strip().startswith('BU_:')]
    if bu_lines:
        print(f"✓ [必需] BU_节点定义: 存在 ({len(bu_lines)} 个)")
        for bu_line in bu_lines:
            nodes = bu_line.replace('BU_:', '').strip()
            if nodes:
                node_list = nodes.split()
                print(f"   节点列表: {', '.join(node_list)} (共{len(node_list)}个节点)")
    else:
        issues.append("缺少BU_节点定义")
        print("❌ [必需] BU_节点定义: 缺失")
    
    # ==================== 5. 检查BO_（报文定义）====================
    bo_lines = [line for line in lines if re.match(r'^BO_\s+\d+', line.strip())]
    if bo_lines:
        print(f"\n✓ [必需] BO_报文定义: 存在 ({len(bo_lines)} 个报文)")
        
        # 检查报文ID重复
        bo_ids = []
        duplicate_ids = []
        valid_bo_count = 0
        invalid_bo_lines = []
        
        for bo_line in bo_lines:
            match = re.match(r'^BO_\s+(\d+)\s+(\S+):\s*(\d+)\s+(\S+)', bo_line.strip())
            if match:
                bo_id = int(match.group(1))
                bo_name = match.group(2)
                bo_length = int(match.group(3))
                bo_node = match.group(4)
                
                if bo_id in bo_ids:
                    duplicate_ids.append(bo_id)
                else:
                    bo_ids.append(bo_id)
                
                # 检查报文长度是否合理（通常1-8字节）
                if bo_length < 1 or bo_length > 64:
                    warnings.append(f"报文 {bo_name} (ID:{bo_id}) 长度异常: {bo_length} 字节")
                
                valid_bo_count += 1
            else:
                invalid_bo_lines.append(bo_line)
                issues.append(f"报文格式错误: {bo_line[:50]}")
        
        if duplicate_ids:
            issues.append(f"发现 {len(duplicate_ids)} 个重复的报文ID: {duplicate_ids}")
            print(f"❌ [严重] 发现 {len(duplicate_ids)} 个重复报文ID: {duplicate_ids}")
        else:
            print("✓ [必需] 报文ID唯一性: 通过")
        
        if invalid_bo_lines:
            print(f"❌ [严重] 发现 {len(invalid_bo_lines)} 个格式错误的报文定义")
        else:
            print(f"✓ [必需] 报文格式检查: 通过 ({valid_bo_count} 个有效报文)")
    else:
        warnings.append("文件中没有定义任何报文（BO_）")
        print("\n⚠ [警告] BO_报文定义: 未找到任何报文")
    
    # ==================== 6. 检查SG_（信号定义）====================
    sg_lines = [line for line in lines if re.match(r'^\s+SG_\s+', line.strip())]
    if sg_lines:
        print(f"✓ [必需] SG_信号定义: 存在 ({len(sg_lines)} 个信号)")
        
        # 检查信号格式
        invalid_sg_count = 0
        for sg_line in sg_lines:
            # SG_ 格式: SG_ <name> : <start_bit>|<length>@<byte_order>+ (<factor>,<offset>) [<min>|<max>] "<unit>" <receiver>
            if not re.match(r'^\s+SG_\s+\S+\s+:\s+\d+\|\d+@\d+[+-]\s+\([^)]+\)\s+\[[^\]]*\]\s+"[^"]*"\s+\S+', sg_line):
                invalid_sg_count += 1
        
        if invalid_sg_count > 0:
            warnings.append(f"发现 {invalid_sg_count} 个格式不规范的信号定义")
            print(f"⚠ [警告] 信号格式: {invalid_sg_count} 个信号格式不规范")
        else:
            print(f"✓ [必需] 信号格式检查: 通过")
    else:
        warnings.append("文件中没有定义任何信号（SG_）")
        print("⚠ [警告] SG_信号定义: 未找到任何信号")
    
    # ==================== 7. 检查CM_（注释）====================
    cm_lines = [line for line in lines if line.strip().startswith('CM_')]
    if cm_lines:
        print(f"✓ [可选] CM_注释: 存在 ({len(cm_lines)} 条注释)")
        
        # 检查注释格式是否有问题
        invalid_cm_count = 0
        for cm_line in cm_lines:
            # 检查是否有未闭合的引号
            quote_count = cm_line.count('"')
            if quote_count % 2 != 0:
                invalid_cm_count += 1
                issues.append(f"注释引号不匹配: {cm_line[:80]}")
        
        if invalid_cm_count > 0:
            print(f"❌ [严重] 发现 {invalid_cm_count} 条注释引号不匹配")
        else:
            print(f"✓ [可选] 注释格式检查: 通过")
    else:
        print("ℹ [可选] CM_注释: 未找到注释（不影响使用）")
    
    # ==================== 8. 检查VAL_（值表定义）====================
    val_lines = [line for line in lines if line.strip().startswith('VAL_')]
    if val_lines:
        print(f"✓ [可选] VAL_值表定义: 存在 ({len(val_lines)} 个值表)")
    else:
        print("ℹ [可选] VAL_值表定义: 未找到（不影响使用）")
    
    # ==================== 9. 检查特殊字符和编码问题 ====================
    chinese_chars = re.findall(r'[\u4e00-\u9fff]', content)
    if chinese_chars:
        print(f"\n✓ [提示] 中文内容: 包含 {len(chinese_chars)} 个中文字符")
        if encoding_used == 'utf-8':
            print("✓ [推荐] 编码方式: UTF-8（完美支持中文）")
        elif encoding_used == 'gbk':
            print("⚠ [注意] 编码方式: GBK（支持中文，但UTF-8更推荐）")
    else:
        print("\nℹ [提示] 中文内容: 未检测到中文字符")
    
    # ==================== 10. 检查行尾格式 ====================
    if '\r\n' in content:
        print("ℹ [提示] 行尾格式: Windows (CRLF)")
    elif '\r' in content:
        print("ℹ [提示] 行尾格式: Mac (CR)")
    else:
        print("ℹ [提示] 行尾格式: Unix/Linux (LF)")
    
    # ==================== 总结报告 ====================
    print(f"\n{'='*70}")
    print(f"检测报告总结")
    print(f"{'='*70}")
    
    if issues:
        print(f"\n❌ 发现 {len(issues)} 个严重问题:")
        for i, issue in enumerate(issues, 1):
            print(f"   {i}. {issue}")
        print("\n⚠️  CANdb++可能无法正常打开此文件！")
        print("   建议修复上述问题后再尝试打开。")
        can_open = False
    else:
        print("\n✅ 未发现严重问题")
        can_open = True
    
    if warnings:
        print(f"\n⚠ 发现 {len(warnings)} 个警告:")
        for i, warning in enumerate(warnings, 1):
            print(f"   {i}. {warning}")
        print("\n⚠️  CANdb++可以打开文件，但可能存在编辑问题。")
    
    if not issues and not warnings:
        print("\n🎉 文件格式完全符合CANdb++标准！")
        print("   可以正常打开和编辑。")
    
    print(f"\n{'='*70}")
    if can_open:
        print("最终结论: ✅ 可以被CANdb++正常打开")
        if warnings:
            print("           ⚠️  但建议修复警告项以获得更好的编辑体验")
    else:
        print("最终结论: ❌ 无法被CANdb++正常打开")
        print("           请先修复严重问题")
    print(f"{'='*70}\n")
    
    return can_open


if __name__ == "__main__":
    # 检测当前打开的DBC文件
    dbc_file = r"c:\Users\kkk\Desktop\4266DBC\output2.dbc"
    check_dbc_compatibility(dbc_file)
