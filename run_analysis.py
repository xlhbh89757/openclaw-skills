# -*- coding: utf-8 -*-
import json, subprocess, os, sys
sys.stdout.reconfigure(encoding='utf-8')

with open(r'C:\Users\EDY\resume_risk_report\resumes_for_analysis.json', 'r', encoding='utf-8') as f:
    resumes = json.load(f)

input_json = json.dumps(resumes, ensure_ascii=False)
script = r'C:\Users\EDY\.qclaw\skills\resume-risk-assessor\scripts\analyze_resume.py'

result = subprocess.run(
    ['python', script, input_json],
    capture_output=True,
    cwd=r'C:\Users\EDY\resume_risk_report'
)

output = result.stdout.decode('gbk', errors='replace')
data = json.loads(output)

print('=== 评估结果汇总 ===')
print(f"总计: {data['count']} 份简历")
print(f"高风险: {data['high_risk']} | 中风险: {data['medium_risk']} | 低风险: {data['low_risk']}")
print()
print('=== 详细结果 ===')
risk_emoji_map = {'高': '🔴', '中': '🟡', '低': '🟢'}
for d in data['details']:
    emoji = risk_emoji_map.get(d['overall'], '⚪')
    print(f"{emoji} {d['name']} - {d['summary']}")

print()
print(f"报告已生成: C:\\Users\\EDY\\resume_risk_report\\resume_risk_report.xlsx")
