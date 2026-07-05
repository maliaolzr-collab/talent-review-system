"""
Excel parser for talent review data.
Parses .xls files with the standard talent 9-box grid format.
Uses xlrd as primary reader, falls back to python-calamine for problematic files.
"""
import os
import xlrd
from datetime import datetime

try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False


# Column mapping for Sheet "02 Dept Personnel & Talent Grid"
# Row index 2 (0-based) contains headers; data starts at row 3
COLUMN_MAP = {
    0: 'seq_no',
    1: 'chinese_name',
    2: 'english_name',
    3: 'position_title',
    4: 'job_responsibility',
    5: 'job_level',
    6: 'age',
    7: 'education',
    8: 'graduation_institution',
    9: 'graduation_date',
    10: 'work_experience',
    11: 'entry_date',
    12: 'company_tenure',
    13: 'base_salary',
    14: 'performance_salary',
    15: 'total_salary',
    16: 'knowledge_skill_match',
    17: 'problem_solving_match',
    18: 'responsibility_match',
    19: 'person_position_score',
    20: 'annual_performance',
    21: 'learning_ability',
    22: 'thinking_ability',
    23: 'understanding_others',
    24: 'emotional_maturity',
    25: 'potential_score',
    26: 'performance_level',
    27: 'potential_level',
    28: 'grid_position',
    29: 'talent_pipeline',
    30: 'result_application',
    31: 'development_plan',
    32: 'seq_no_2',
    33: 'management_strategy',
}

# 9-box grid labels
GRID_INFO = {
    1: {'name': '问题员工', 'en': 'Problem Performer', 'perf': '低', 'pot': '低',
        'color': '#ff4d4f', 'strategy': '设定底线要求，明确改进期限，建立退出机制'},
    2: {'name': '差距员工', 'en': 'Gap Performer', 'perf': '低', 'pot': '中',
        'color': '#fa8c16', 'strategy': '当前绩效较差，可能尚未适应，明确提出改进要求'},
    3: {'name': '一般员工', 'en': 'Average Performer', 'perf': '中', 'pot': '低',
        'color': '#faad14', 'strategy': '重点提升绩效，保持原级或降级，减少管理职责'},
    4: {'name': '待发展者', 'en': 'Potential Developer', 'perf': '低', 'pot': '高',
        'color': '#fa8c16', 'strategy': '强化绩效改进，发掘潜力优势，重点帮扶培养'},
    5: {'name': '中坚力量', 'en': 'Core Contributor', 'perf': '中', 'pot': '中',
        'color': '#52c41a', 'strategy': '可依靠的稳定贡献者，重点开发培训，提高绩效'},
    6: {'name': '熟练员工', 'en': 'Skilled Performer', 'perf': '高', 'pot': '低',
        'color': '#1890ff', 'strategy': '可依靠的稳定贡献者，重点开发培训，激发潜力'},
    7: {'name': '潜力之星', 'en': 'Potential Star', 'perf': '中', 'pot': '高',
        'color': '#722ed1', 'strategy': '强化绩效辅导，指导职业规划，确保薪酬竞争力'},
    8: {'name': '绩效之星', 'en': 'Performance Star', 'perf': '高', 'pot': '中',
        'color': '#13c2c2', 'strategy': '进一步开发潜力，鼓励多担责，确保薪酬竞争力'},
    9: {'name': '超级明星', 'en': 'Superstar', 'perf': '高', 'pot': '高',
        'color': '#eb2f96', 'strategy': '新的挑战和机会，晋升+轮岗+高薪激励'},
}

PIPELINE_ORDER = ['第一梯队', '第二梯队', '第三梯队', '第四梯队', '第五梯队', '其他']


def _excel_date_to_str(val):
    """Convert value to date string."""
    if not val or val == '':
        return ''
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    try:
        if isinstance(val, (int, float)):
            dt = xlrd.xldate.xldate_as_datetime(val, 0)
            return dt.strftime('%Y-%m-%d')
        return str(val)
    except Exception:
        return str(val)


def _safe_float(val):
    """Convert to float, return None if empty/invalid."""
    if val == '' or val is None:
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def _safe_int(val):
    """Convert to int, return None if empty/invalid."""
    f = _safe_float(val)
    return int(f) if f is not None else None


def _clean_str(val):
    """Clean string value."""
    if val is None:
        return ''
    s = str(val).strip()
    return s


def _read_sheet_xlrd(filepath):
    """Read sheet using xlrd. Returns (rows_list, nrows, ncols)."""
    wb = xlrd.open_workbook(filepath)
    sheet_name = None
    for name in wb.sheet_names():
        if '02' in name or 'Personnel' in name:
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = wb.sheet_names()[1] if len(wb.sheet_names()) > 1 else wb.sheet_names()[0]
    sh = wb.sheet_by_name(sheet_name)
    rows = []
    for r in range(sh.nrows):
        row = [sh.cell_value(r, c) for c in range(sh.ncols)]
        rows.append(row)
    return rows, sh.nrows, sh.ncols


def _read_sheet_calamine(filepath):
    """Read sheet using python-calamine. Returns (rows_list, nrows, ncols)."""
    wb = CalamineWorkbook.from_path(filepath)
    sheet_name = None
    for name in wb.sheet_names:
        if '02' in name or 'Personnel' in name:
            sheet_name = name
            break
    if not sheet_name:
        names = wb.sheet_names
        sheet_name = names[1] if len(names) > 1 else names[0]
    sh = wb.get_sheet_by_name(sheet_name)
    rows = sh.to_python()
    ncols = max(len(r) for r in rows) if rows else 0
    return rows, len(rows), ncols


def parse_excel_file(filepath, dept_category, sub_dept_name):
    """
    Parse a single Excel file and return list of employee dicts.

    Args:
        filepath: Path to .xls file
        dept_category: Major department category (印染/服装/针织/辅料/总部)
        sub_dept_name: Sub-department name extracted from filename

    Returns:
        List of employee dicts
    """
    # Try xlrd first, fall back to calamine
    try:
        rows, nrows, ncols = _read_sheet_xlrd(filepath)
    except Exception:
        if HAS_CALAMINE:
            rows, nrows, ncols = _read_sheet_calamine(filepath)
        else:
            raise

    employees = []

    # Data starts at row 3 (index 3), headers at row 2
    for row_idx in range(3, nrows):
        row = rows[row_idx]
        # Pad row if needed
        while len(row) < 34:
            row.append('')

        # Check if row has a name (col 1)
        name_val = row[1]
        if not name_val or str(name_val).strip() == '':
            continue

        emp = {
            'dept_category': dept_category,
            'sub_dept': sub_dept_name,
        }

        for col_idx, field_name in COLUMN_MAP.items():
            val = row[col_idx] if col_idx < len(row) else ''

            if field_name in ('entry_date', 'graduation_date'):
                emp[field_name] = _excel_date_to_str(val)
            elif field_name == 'age':
                emp[field_name] = _safe_int(val)
            elif field_name in ('base_salary', 'performance_salary', 'total_salary',
                                'knowledge_skill_match', 'problem_solving_match',
                                'responsibility_match', 'person_position_score',
                                'potential_score'):
                emp[field_name] = _safe_float(val)
            elif field_name in ('learning_ability', 'thinking_ability',
                                'understanding_others', 'emotional_maturity',
                                'grid_position', 'seq_no', 'seq_no_2'):
                emp[field_name] = _safe_int(val)
            else:
                emp[field_name] = _clean_str(val)

        # Normalize annual_performance (A/B/C/S)
        perf = emp.get('annual_performance', '')
        if perf:
            emp['annual_performance'] = perf.strip().upper()

        # Normalize performance_level and potential_level
        for level_field in ('performance_level', 'potential_level'):
            val = emp.get(level_field, '')
            val_str = str(val)
            if '高' in val_str or 'High' in val_str:
                emp[level_field] = '高'
            elif '中' in val_str or 'Medium' in val_str or 'Growth' in val_str:
                emp[level_field] = '中'
            elif '低' in val_str or 'Low' in val_str or 'Limited' in val_str:
                emp[level_field] = '低'
            else:
                emp[level_field] = val_str.strip()

        # Normalize talent_pipeline
        pipeline = str(emp.get('talent_pipeline', '')).strip()
        matched = False
        for p in PIPELINE_ORDER:
            if p in pipeline:
                emp['talent_pipeline'] = p
                matched = True
                break
        if not matched:
            emp['talent_pipeline'] = '其他'

        # Add grid info
        grid_pos = emp.get('grid_position')
        if grid_pos and grid_pos in GRID_INFO:
            info = GRID_INFO[grid_pos]
            emp['grid_name'] = info['name']
            emp['grid_perf'] = info['perf']
            emp['grid_pot'] = info['pot']
            emp['grid_color'] = info['color']
        else:
            emp['grid_name'] = ''
            emp['grid_perf'] = ''
            emp['grid_pot'] = ''
            emp['grid_color'] = '#d9d9d9'

        employees.append(emp)

    return employees


def find_excel_files(base_dir):
    """
    Find all .xls files organized by department category.

    Returns:
        List of (filepath, dept_category, sub_dept_name) tuples
    """
    results = []
    dept_categories = ['印染', '服装', '针织', '辅料', '总部']

    for category in dept_categories:
        cat_dir = os.path.join(base_dir, category)
        if not os.path.isdir(cat_dir):
            continue
        for fname in sorted(os.listdir(cat_dir)):
            if fname.endswith('.xls') and not fname.startswith('.'):
                filepath = os.path.join(cat_dir, fname)
                sub_dept = os.path.splitext(fname)[0]
                results.append((filepath, category, sub_dept))

    return results


def parse_all_excel_files(base_dir):
    """
    Parse all Excel files in the directory structure.

    Returns:
        Tuple of (employees_list, departments_list)
    """
    files = find_excel_files(base_dir)
    all_employees = []
    departments = set()

    for filepath, category, sub_dept in files:
        try:
            emps = parse_excel_file(filepath, category, sub_dept)
            all_employees.extend(emps)
            departments.add((category, sub_dept))
        except Exception as e:
            print(f"Error parsing {filepath}: {e}")

    dept_list = sorted(departments, key=lambda x: (x[0], x[1]))
    return all_employees, dept_list


if __name__ == '__main__':
    base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    employees, depts = parse_all_excel_files(base)
    print(f"Total employees: {len(employees)}")
    print(f"Total sub-departments: {len(depts)}")
    for cat, sub in depts:
        count = sum(1 for e in employees if e['dept_category'] == cat and e['sub_dept'] == sub)
        print(f"  {cat}/{sub}: {count}人")
