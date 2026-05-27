import re

# 测试信号匹配
test_lines = [
    ' SG_ VCUDTC : 55|8@0+ (1,0) [0|0] "" Vector__XXX',
    'SG_ VCUDTC : 55|8@0+ (1,0) [0|0] "" Vector__XXX',
    ' BO_ 273 VCU_RPCU: 8 EVCAN',
]

for line in test_lines:
    print(f"测试行: {repr(line)}")
    
    # 尝试不同的正则
    match1 = re.match(r'^\s+SG_\s+(\S+)\s+:', line)
    print(f"  模式1 (^\\s+SG_): {match1.group(1) if match1 else 'None'}")
    
    match2 = re.match(r'^\s*SG_\s+(\S+)\s+:', line)
    print(f"  模式2 (^\\s*SG_): {match2.group(1) if match2 else 'None'}")
    
    match3 = re.search(r'SG_\s+(\S+)\s+:', line)
    print(f"  模式3 (search): {match3.group(1) if match3 else 'None'}")
    print()
