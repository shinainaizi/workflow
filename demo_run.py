#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Demo Mode - No API Key required
"""

import os
import sys
import re
from pathlib import Path
from datetime import datetime
import docx
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

def safe_print(msg):
    """Safe print that handles encoding issues"""
    try:
        print(msg)
    except UnicodeEncodeError:
        # Fallback to ASCII
        safe_msg = msg.encode('ascii', 'ignore').decode('ascii')
        if safe_msg.strip():
            print(safe_msg)
        else:
            print("[Non-ASCII content]")


def extract_from_scheme(scheme_path):
    doc = docx.Document(scheme_path)
    full_text = '\n'.join([p.text for p in doc.paragraphs])
    
    data = {
        'activity_theme': '',
        'activity_date': '',
        'activity_time': '',
        'activity_location': '',
        'participants': '',
        'expected_participants': '',
        'actual_participants': '',
        'activity_goal': '',
        'activity_content': '',
        'activity_flow': [],
        'highlights': '',
        'problems': '',
        'improvements': '',
        'staff': 'Staff',
        'notes': ''
    }
    
    # Extract theme
    theme_match = re.search(r'(?:活动主题|主题)[\s]*[:：]\s*(.+)', full_text)
    if not theme_match:
        theme_match = re.search(r'(?:Activity Theme|Theme)[\s]*[:：]\s*(.+)', full_text)
    if theme_match:
        data['activity_theme'] = theme_match.group(1).strip()
    
    # Fallback: try to get from first line or title
    if not data['activity_theme']:
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        if lines:
            data['activity_theme'] = lines[0]
    
    # Extract date
    date_match = re.search(r'(\d{1,2})\s*[月/]\s*(\d{1,2})[日/]', full_text)
    if not date_match:
        date_match = re.search(r'(\d{1,2})[/-](\d{1,2})', full_text)
    if date_match:
        year = datetime.now().year
        data['activity_date'] = f"{year}.{date_match.group(1)}.{date_match.group(2)}"
    
    # Extract time
    time_match = re.search(r'(\d{1,2}:\d{2})\s*[-~]\s*(\d{1,2}:\d{2})', full_text)
    if time_match:
        data['activity_time'] = f"{time_match.group(1)}-{time_match.group(2)}"
    
    # Extract location
    location_match = re.search(r'(?:活动地点|地点)[\s]*[:：]\s*(.+)', full_text)
    if not location_match:
        location_match = re.search(r'(?:Activity Location|Location)[\s]*[:：]\s*(.+)', full_text)
    if location_match:
        data['activity_location'] = location_match.group(1).strip()
    
    # Extract participants
    participants_match = re.search(r'(?:参与对象|对象)[\s]*[:：]\s*(.+)', full_text)
    if not participants_match:
        participants_match = re.search(r'(?:Participants|Target)[\s]*[:：]\s*(.+)', full_text)
    if participants_match:
        data['participants'] = participants_match.group(1).strip()
    
    # Extract expected participants
    expected_match = re.search(r'(\d+)组\s*家庭', full_text)
    if not expected_match:
        expected_match = re.search(r'(\d+)\s*组', full_text)
    if not expected_match:
        expected_match = re.search(r'(\d+)', full_text)
    if expected_match:
        data['expected_participants'] = f"{expected_match.group(1)}组家庭"
    
    # Extract goal
    goal_match = re.search(r'(?:活动目的|目的)[\s]*[:：]\s*(.+)', full_text)
    if not goal_match:
        goal_match = re.search(r'(?:Activity Purpose|Purpose)[\s]*[:：]\s*(.+)', full_text)
    if goal_match:
        data['activity_goal'] = goal_match.group(1).strip()
    
    # Extract flow
    flow_items = re.findall(r'(?:环节[一二三四五六]|[\d]+、)[：:]\s*(.+)', full_text)
    if not flow_items:
        flow_items = re.findall(r'[\d]+[.、]\s*(.+)', full_text)
    data['activity_flow'] = flow_items[:5]
    
    # Extract notes
    notes_start = full_text.lower().find('notes')
    if notes_start == -1:
        notes_start = full_text.lower().find('precautions')
    if notes_start != -1:
        notes_section = full_text[notes_start:notes_start+500]
        note_items = re.findall(r'[bullet\-\d][.、]\s*(.+)', notes_section)
        data['notes'] = '\n'.join(note_items[:3])
    
    return data

def extract_from_notes(notes_path, base_data):
    """Extract additional info from notes"""
    with open(notes_path, 'r', encoding='utf-8') as f:
        notes_text = f.read()
    
    # Extract actual participants
    actual_match = re.search(r'(?:实际到场|实际|到场)\s*(\d+)[组个]', notes_text)
    if not actual_match:
        actual_match = re.search(r'(\d+)\s*组', notes_text)
    if actual_match:
        base_data['actual_participants'] = f"{actual_match.group(1)}组家庭"
    
    # Extract highlights
    highlights_match = re.search(r'(?:Highlights|亮点)[\s]*[:：]\s*([\s\S]+?)(?=\n\n|Problems|Issues)', notes_text)
    if highlights_match:
        highlights = highlights_match.group(1).strip()
        highlights = re.sub(r'^\s*[bullet\-\d][.、]\s*', '', highlights, flags=re.MULTILINE)
        base_data['highlights'] = highlights
    
    # Extract problems
    problems_match = re.search(r'(?:Problems|Issues|问题)[\s]*[:：]\s*([\s\S]+?)(?=\n\n|Improvements|Suggestions)', notes_text)
    if problems_match:
        problems = problems_match.group(1).strip()
        problems = re.sub(r'^\s*[bullet\-\d][.、]\s*', '', problems, flags=re.MULTILINE)
        base_data['problems'] = problems
    
    # Extract improvements
    improvements_match = re.search(r'(?:Improvements|Suggestions|建议)[\s]*[:：]\s*([\s\S]+?)(?=\n\n|Photos|Next)', notes_text)
    if improvements_match:
        improvements = improvements_match.group(1).strip()
        improvements = re.sub(r'^\s*[bullet\-\d][.、]\s*', '', improvements, flags=re.MULTILINE)
        base_data['improvements'] = improvements
    
    # Generate content description
    base_data['activity_content'] = generate_content_description(base_data)
    
    return base_data

def generate_content_description(data):
    """Generate activity content description"""
    lines = []
    
    # Opening
    date_str = data.get('activity_date', '').replace('.', 'month').replace('.', 'day')
    time_str = data.get('activity_time', '').split('-')[0] if data.get('activity_time') else ''
    lines.append(f"On {date_str} at {time_str},")
    lines.append(f"at {data.get('activity_location', 'Community')} held activity \"{data.get('activity_theme', '')}\",")
    
    actual = data.get('actual_participants', '')
    if actual:
        lines.append(f"attracted {actual} to participate.")
    else:
        lines.append("attracted community residents to participate actively.")
    
    lines.append("")
    
    # Activity content
    goal = data.get('activity_goal', '')
    if goal:
        lines.append(f"This activity aimed at {goal.split(',')[0]},")
    
    flow = data.get('activity_flow', [])
    if flow:
        lines.append(f"successively carried out {'、'.join(flow[:3])} and other sessions.")
    
    lines.append("The atmosphere was warm and joyful, gaining unanimous recognition from participants.")
    
    # Follow-up plan
    lines.append("")
    lines.append("The community will continue to launch diversified activities to enrich residents' cultural life.")
    
    return '\n'.join(lines)

def generate_ledger(data, output_dir='outputs/ledger'):
    """Generate activity ledger"""
    os.makedirs(output_dir, exist_ok=True)
    
    doc = docx.Document()
    
    # Page margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
    
    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('Community Activity Ledger')
    run.font.size = Pt(18)
    run.font.bold = True
    run.font.name = 'SimHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimHei')
    
    doc.add_paragraph()
    
    # Table
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    
    # Column widths
    table.columns[0].width = Cm(3)
    table.columns[1].width = Cm(12)
    
    # Row 1: Time | Location
    row1 = table.rows[0].cells
    row1[0].text = 'Time'
    row1[1].text = f"{data.get('activity_date', '')} {data.get('activity_time', '')}"
    
    # Row 2: Theme
    row2 = table.rows[1].cells
    row2[0].text = 'Theme'
    row2[1].text = data.get('activity_theme', '')
    
    # Row 3: Participants
    row3 = table.rows[2].cells
    row3[0].text = 'Participants'
    row3[1].text = data.get('participants', 'Community Residents')
    
    # Row 4: Content (merged)
    row4 = table.rows[3].cells
    row4[0].merge(row4[1])
    row4[0].text = 'Content:'
    
    content = data.get('activity_content', '')
    if content:
        content_para = row4[0].paragraphs[0]
        content_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        for i, line in enumerate(content.split('\n')):
            if i == 0:
                run = content_para.add_run(line)
            else:
                content_para.add_run('\n')
                run = content_para.add_run(line)
            run.font.size = Pt(12)
            run.font.name = 'SimSun'
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    
    # Row 5: Notes
    row5 = table.rows[4].cells
    row5[0].merge(row5[1])
    row5[0].text = 'Notes'
    
    notes_parts = []
    expected = data.get('expected_participants', '')
    actual = data.get('actual_participants', '')
    if expected or actual:
        notes_parts.append(f"Expected: {expected or 'N/A'}  Actual: {actual or 'N/A'}")
    
    if data.get('highlights'):
        highlights = data['highlights'].replace('\n', ' ')
        notes_parts.append(f"Highlights: {highlights[:100]}")
    
    if data.get('problems'):
        problems = data['problems'].replace('\n', ' ')
        notes_parts.append(f"Problems: {problems[:100]}")
    
    if data.get('improvements'):
        improvements = data['improvements'].replace('\n', ' ')
        notes_parts.append(f"Improvements: {improvements[:100]}")
    
    if notes_parts:
        notes_para = row5[0].paragraphs[0]
        notes_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for i, line in enumerate(notes_parts):
            if i == 0:
                run = notes_para.add_run(line)
            else:
                notes_para.add_run('\n')
                run = notes_para.add_run(line)
            run.font.size = Pt(10)
            run.font.name = 'SimSun'
            run.font.color.rgb = RGBColor(128, 128, 128)
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    
    # Format table
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in paragraph.runs:
                    run.font.size = Pt(12)
                    run.font.name = 'SimSun'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    
    # Save
    theme = data.get('activity_theme', 'Ledger').replace('/', '-').replace('\\', '-').replace(':', '-')
    filename = f"{theme}_{data.get('activity_date', datetime.now().strftime('%Y%m%d'))}.docx"
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)
    
    return output_path

def generate_report(activities, output_dir='outputs/report'):
    """Generate summary report"""
    os.makedirs(output_dir, exist_ok=True)
    
    doc = docx.Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(3.17)
        section.right_margin = Cm(3.17)
    
    # Title
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run('Community Activity Summary Report')
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.name = 'SimHei'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimHei')
    
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f'({datetime.now().strftime("%Y-%m")})')
    run.font.size = Pt(14)
    run.font.name = 'SimSun'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
    
    doc.add_paragraph()
    
    # Overview
    doc.add_heading('I. Overview', level=1)
    total = len(activities)
    total_participants = sum([int(re.search(r'\d+', str(a.get('actual_participants', '0'))).group()) if re.search(r'\d+', str(a.get('actual_participants', '0'))) else 0 for a in activities])
    overview = f"""During this period, {total} community activities were held, with approximately {total_participants} total participants.

Activities covered various types such as sensory training, parent-child interaction, and role-playing, serving infants, families, and community residents. All activities were carried out as planned with good atmosphere and positive feedback."""
    doc.add_paragraph(overview)
    
    # Activity details
    doc.add_heading('II. Activity Details', level=1)
    for i, activity in enumerate(activities, 1):
        doc.add_heading(f'{i}. {activity.get("activity_theme", "Unnamed")}', level=2)
        
        info_lines = [
            f'Time: {activity.get("activity_date", "")} {activity.get("activity_time", "")}',
            f'Location: {activity.get("activity_location", "")}',
            f'Participants: {activity.get("actual_participants", "Unknown")}',
            f'Target: {activity.get("participants", "")}'
        ]
        
        for line in info_lines:
            p = doc.add_paragraph(line)
            p.paragraph_format.left_indent = Cm(0.5)
        
        content = activity.get('activity_content', '')
        if content:
            p = doc.add_paragraph('Summary:')
            p.paragraph_format.left_indent = Cm(0.5)
            p = doc.add_paragraph(content[:200] + '...' if len(content) > 200 else content)
            p.paragraph_format.left_indent = Cm(0.5)
        
        if activity.get('highlights'):
            p = doc.add_paragraph(f'Highlights: {activity.get("highlights", "")[:150]}')
            p.paragraph_format.left_indent = Cm(0.5)
    
    # Review
    doc.add_heading('III. Review & Optimization', level=1)
    
    all_problems = []
    all_improvements = []
    for activity in activities:
        if activity.get('problems'):
            all_problems.append(activity['problems'])
        if activity.get('improvements'):
            all_improvements.append(activity['improvements'])
    
    if all_problems:
        doc.add_paragraph('(A) Existing Problems', style='Heading 3')
        for i, problem in enumerate(all_problems[:3], 1):
            p = doc.add_paragraph(f'{i}. {problem[:200]}')
            p.paragraph_format.left_indent = Cm(0.5)
    
    if all_improvements:
        doc.add_paragraph('(B) Improvement Directions', style='Heading 3')
        for i, improvement in enumerate(all_improvements[:3], 1):
            p = doc.add_paragraph(f'{i}. {improvement[:200]}')
            p.paragraph_format.left_indent = Cm(0.5)
    
    if not all_problems and not all_improvements:
        doc.add_paragraph('All activities were carried out smoothly. Continuous attention to participant feedback is recommended for ongoing optimization.')
    
    # Next steps
    doc.add_heading('IV. Next Steps', level=1)
    plans = """1. Continuously optimize activity processes to enhance participant experience
2. Enrich activity content based on resident feedback
3. Strengthen activity promotion to expand participation coverage
4. Improve activity documentation and establish activity archives"""
    doc.add_paragraph(plans)
    
    # Save
    filename = f'Summary_Report_{datetime.now().strftime("%Y%m%d")}.docx'
    report_path = os.path.join(output_dir, filename)
    doc.save(report_path)
    
    return report_path

def main():
    """Main flow"""
    # Write output to log file to avoid encoding issues
    log_path = 'run_log.txt'
    log_lines = []
    
    log_lines.append("=" * 60)
    log_lines.append("Community Event Ledger Automation - Demo Mode")
    log_lines.append("=" * 60)
    
    # Paths
    scheme_path = 'examples/input/scheme.docx'
    notes_path = 'examples/input/notes/回忆记录.txt'
    
    # Check files
    if not os.path.exists(scheme_path):
        log_lines.append(f"Error: Scheme not found {scheme_path}")
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(log_lines))
        sys.exit(1)
    
    # Step 1: Extract info
    log_lines.append("")
    log_lines.append("[1/4] Extracting info from scheme...")
    data = extract_from_scheme(scheme_path)
    log_lines.append(f"  OK: {data['activity_theme']}")
    
    # Step 2: Add notes
    if os.path.exists(notes_path):
        log_lines.append("")
        log_lines.append("[2/4] Adding notes info...")
        data = extract_from_notes(notes_path, data)
        log_lines.append(f"  OK: Actual participants {data['actual_participants']}")
    
    # Step 3: Generate ledger
    log_lines.append("")
    log_lines.append("[3/4] Generating ledger...")
    ledger_path = generate_ledger(data)
    log_lines.append(f"  OK: {ledger_path}")
    
    # Step 4: Generate report
    log_lines.append("")
    log_lines.append("[4/4] Generating report...")
    report_path = generate_report([data])
    log_lines.append(f"  OK: {report_path}")
    
    log_lines.append("")
    log_lines.append("=" * 60)
    log_lines.append("Done!")
    log_lines.append(f"Ledger: {ledger_path}")
    log_lines.append(f"Report: {report_path}")
    log_lines.append("=" * 60)
    
    # Write log
    with open(log_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(log_lines))
    
    # Also try to print
    for line in log_lines:
        safe_print(line)

if __name__ == '__main__':
    main()
