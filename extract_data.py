"""
Extrae datos del Excel 'data/reporte.xlsx' y genera data.js
con todas las estadísticas preprocesadas para el dashboard.

Se ejecuta localmente o via GitHub Actions.
"""
import openpyxl
import json
import math
import os
from datetime import datetime

# Rutas relativas al directorio del script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(SCRIPT_DIR, 'data', 'reporte.xlsx')
OUTPUT_PATH = os.path.join(SCRIPT_DIR, 'data.js')

def extract_section(ws, section_name):
    """Extract student data from a worksheet."""
    students = []
    controls_with_data = []
    
    # Detect which controls have data
    for col_idx in range(5, 19):  # columns F(5) to S(18) => C1-C14
        has_data = False
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=col_idx+1, max_col=col_idx+1):
            if row[0].value is not None:
                has_data = True
                break
        if has_data:
            controls_with_data.append(col_idx - 5)
    
    num_controls = max(controls_with_data) + 1 if controls_with_data else 0
    control_labels = [f"C{i+1}" for i in range(num_controls)]
    
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=20):
        vals = [cell.value for cell in row]
        if vals[0] is None:
            continue
        
        scores = []
        for i in range(num_controls):
            v = vals[5 + i]
            scores.append(v if v is not None else 0)
        
        students.append({
            'name': f"{vals[1]} {vals[2]}, {vals[3]}".strip(),
            'id': str(vals[4]) if vals[4] else '',
            'scores': scores,
            'total': vals[19] if vals[19] is not None else sum(scores)
        })
    
    return {
        'name': section_name,
        'controlLabels': control_labels,
        'students': students,
        'numControls': num_controls
    }


def compute_stats(section):
    """Compute all statistics for a section."""
    num_controls = section['numControls']
    students = section['students']
    n_students = len(students)
    max_per_control = 5

    stats = {
        'controlLabels': section['controlLabels'],
        'numStudents': n_students,
        'maxPerControl': max_per_control,
    }

    mins = []
    maxs = []
    avgs = []
    avgs_responded = []
    responded_counts = []
    cumulative_mins = []
    cumulative_maxs = []
    cumulative_avgs = []
    cumulative_nota7 = []
    approval_pcts = []
    boxplot_data = []
    distributions = []

    cum_min = 0
    cum_max = 0
    cum_sum = 0

    for c in range(num_controls):
        scores = [s['scores'][c] for s in students]
        nonzero_scores = [s for s in scores if s > 0]

        c_min = min(scores)
        c_max = max(scores)
        c_avg = sum(scores) / len(scores) if scores else 0
        c_avg_resp = sum(nonzero_scores) / len(nonzero_scores) if nonzero_scores else 0
        c_responded = len(nonzero_scores)
        c_approval = sum(1 for s in scores if s >= 3) / len(scores) * 100 if scores else 0

        mins.append(round(c_min, 2))
        maxs.append(round(c_max, 2))
        avgs.append(round(c_avg, 2))
        avgs_responded.append(round(c_avg_resp, 2))
        responded_counts.append(c_responded)
        approval_pcts.append(round(c_approval, 1))

        cum_min += c_min
        cum_max += max_per_control
        cum_sum += c_avg
        cumulative_mins.append(round(cum_min, 2))
        cumulative_maxs.append(round(cum_max, 2))
        cumulative_avgs.append(round(cum_sum, 2))
        cumulative_nota7.append(round(cum_max * 0.8, 2))

        sorted_scores = sorted(scores)
        n = len(sorted_scores)

        def percentile(arr, p):
            k = (len(arr) - 1) * p
            f = math.floor(k)
            c_val = math.ceil(k)
            if f == c_val:
                return arr[int(k)]
            return arr[f] * (c_val - k) + arr[c_val] * (k - f)

        boxplot_data.append({
            'min': sorted_scores[0],
            'q1': round(percentile(sorted_scores, 0.25), 2),
            'median': round(percentile(sorted_scores, 0.5), 2),
            'q3': round(percentile(sorted_scores, 0.75), 2),
            'max': sorted_scores[-1],
            'mean': round(c_avg, 2)
        })

        dist = [0] * 6
        for s in scores:
            if 0 <= s <= 5:
                dist[int(s)] += 1
        distributions.append(dist)

    stats['mins'] = mins
    stats['maxs'] = maxs
    stats['avgs'] = avgs
    stats['avgsResponded'] = avgs_responded
    stats['respondedCounts'] = responded_counts
    stats['approvalPcts'] = approval_pcts
    stats['cumulativeMins'] = cumulative_mins
    stats['cumulativeMaxs'] = cumulative_maxs
    stats['cumulativeAvgs'] = cumulative_avgs
    stats['cumulativeNota7'] = cumulative_nota7
    stats['boxplot'] = boxplot_data
    stats['distributions'] = distributions

    stats['heatmap'] = {
        'names': [s['name'] for s in students],
        'scores': [s['scores'] for s in students],
        'totals': [s['total'] for s in students]
    }

    ranked = sorted(students, key=lambda s: s['total'], reverse=True)
    stats['ranking'] = {
        'names': [s['name'] for s in ranked[:15]],
        'totals': [s['total'] for s in ranked[:15]]
    }

    return stats


def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"ERROR: No se encontró el archivo Excel en: {EXCEL_PATH}")
        return

    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)

    section1 = extract_section(wb[wb.sheetnames[0]], 'Sección 1')
    section2 = extract_section(wb[wb.sheetnames[1]], 'Sección 2')

    stats1 = compute_stats(section1)
    stats2 = compute_stats(section2)

    combined = {
        'name': 'Ambas Secciones',
        'controlLabels': section1['controlLabels'],
        'students': section1['students'] + section2['students'],
        'numControls': section1['numControls']
    }
    stats_combined = compute_stats(combined)

    today = datetime.now().strftime('%Y-%m-%d')

    data = {
        'section1': stats1,
        'section2': stats2,
        'combined': stats_combined,
        'sectionNames': ['Sección 1', 'Sección 2', 'Ambas Secciones'],
        'maxPerControl': 5,
        'generatedAt': today
    }

    js_content = f"// Auto-generated data from Excel — {today}\nconst REPORT_DATA = {json.dumps(data, ensure_ascii=False, indent=2)};\n"

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(js_content)

    print("Data extracted successfully!")
    print(f"   Section 1: {len(section1['students'])} students, {section1['numControls']} controls")
    print(f"   Section 2: {len(section2['students'])} students, {section2['numControls']} controls")
    print(f"   Output: {OUTPUT_PATH}")


if __name__ == '__main__':
    main()
