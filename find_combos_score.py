"""
双组别宠物精确覆盖计算器
-----------------------
找出所有恰好7只宠物的组合，使其覆盖全部组别且无重叠，并按星光值排序。

使用方法：
  python find_combos_score.py                   # 使用默认文件
  python find_combos_score.py 我的宠物列表.xlsx # 指定其他文件

Excel 格式要求：第一行为表头，之后每行：宠物名 | 组别1 | 组别2 | 星光值
"""

import sys
import os
import openpyxl

# ── 配置 ────────────────────────────────────────────────────────────────────
DEFAULT_INPUT  = "我的最佳.xlsx"
OUTPUT_FILE    = "results.txt"
# ────────────────────────────────────────────────────────────────────────────

def load_pets(path):
    """读取宠物数据：名称、组别1、组别2、星光值"""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    pets = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1] and row[2]:
            # 第四列（索引3）为星光值，默认0
            star_value = int(row[3]) if row[3] is not None else 0
            pets.append((str(row[0]).strip(), str(row[1]).strip(), str(row[2]).strip(), star_value))
    return pets

def find_all_covers(pets):
    """使用回溯算法找出所有精确覆盖组合"""
    all_groups = set(g for _, g1, g2, _ in pets for g in (g1, g2))
    pets_per_combo = len(all_groups) // 2          # 每只宠物覆盖2个组别
    results = []

    def backtrack(remaining, selected):
        if not remaining:
            results.append(list(selected))
            return
        needed = pets_per_combo - len(selected)
        if len(remaining) > needed * 2:            # 剩余槽位不够，剪枝
            return
        # 始终选字典序最小的未覆盖组别作为分支点，保证每个组合只被找到一次
        target = min(remaining)
        for name, g1, g2, _ in pets:
            if target in (g1, g2):
                other = g2 if target == g1 else g1
                if other in remaining:             # 精确覆盖：两个组别都还未使用
                    remaining.remove(target)
                    remaining.remove(other)
                    selected.append(name)
                    backtrack(remaining, selected)
                    selected.pop()
                    remaining.add(target)
                    remaining.add(other)

    backtrack(set(all_groups), [])
    return results, all_groups

def fmt_combo(combo, pet_groups):
    """格式化组合：显示宠物名(组别1+组别2)"""
    return ', '.join(f"{n}({pet_groups[n][0]}+{pet_groups[n][1]})" for n in combo)

def calc_combo_score(combo, pet_star):
    """计算组合的星光值总和"""
    return sum(pet_star[n] for n in combo)

def main():
    input_file = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT
    if not os.path.exists(input_file):
        print(f"文件不存在：{input_file}")
        sys.exit(1)

    print(f"读取文件：{input_file}")
    pets = load_pets(input_file)
    all_groups_count = len(set(g for _, g1, g2, _ in pets for g in (g1, g2)))
    print(f"宠物数：{len(pets)}  组别数：{all_groups_count}")

    results, all_groups = find_all_covers(pets)
    print(f"\n共找到 {len(results)} 种组合（每种恰好 {all_groups_count // 2} 只，覆盖全部 {all_groups_count} 个组别）")

    # 构建宠物信息字典：名称 -> (组别1, 组别2, 星光值)
    pet_groups = {name: (g1, g2) for name, g1, g2, _ in pets}
    pet_star = {name: star for name, _, _, star in pets}

    # 计算每种组合的星光值并排序（从高到低）
    combos_with_score = [(combo, calc_combo_score(combo, pet_star)) for combo in results]
    combos_with_score.sort(key=lambda x: x[1], reverse=True)

    output_path = os.path.join(os.path.dirname(input_file), OUTPUT_FILE)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(f"输入文件：{input_file}\n")
        f.write(f"组别（{all_groups_count}个）：{'、'.join(sorted(all_groups))}\n")
        f.write(f"\n共 {len(combos_with_score)} 种组合（按星光值从高到低排序）\n\n")
        for i, (combo, score) in enumerate(combos_with_score, 1):
            f.write(f"{i:5d}. 星光值={score:5d}  {fmt_combo(combo, pet_groups)}\n")
    print(f"结果已保存到：{output_path}\n")

    # 显示前10名
    print("Top 10 组合（按星光值排序）：")
    for i, (combo, score) in enumerate(combos_with_score[:10], 1):
        print(f"{i:3d}. 星光值={score:5d}  {fmt_combo(combo, pet_groups)}")
    if len(combos_with_score) > 10:
        print(f"  ... 共 {len(combos_with_score)} 条，完整列表见 {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
