import sys
from victorious_christian_life import TOPICS

html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  @page { size: A4 portrait; margin: 15mm; }
  body { font-family: "Nirmala UI", "Segoe UI", Arial, sans-serif; font-size: 14pt; line-height: 1.4; color: #000; margin: 0; padding: 0; }
  h1 { text-align: center; margin: 0 0 5px 0; font-size: 24pt; color: #0A1128; }
  h2 { text-align: center; margin: 0 0 15px 0; font-size: 16pt; color: #555; font-weight: normal; }
  .grid { display: flex; flex-direction: column; gap: 12px; align-items: center; }
  .topic { border: 2px solid #aaa; padding: 10px 20px; border-radius: 8px; background-color: #fafafa; width: 80%; text-align: center; }
  .topic-title { font-weight: bold; font-size: 16pt; color: #92400E; }
  .ml-text { color: #1E3A8A; font-size: 15pt; display: block; margin-top: 5px; } 
</style>
</head>
<body>
  <h1>VICTORIOUS CHRISTIAN LIFE</h1>
  <h1 style="font-size:20pt; margin-bottom:8px; color:#1E3A8A;">വിജയകരമായ ക്രിസ്തീയ ജീവിതം</h1>
  <h2>ICPF Camp at Trivandrum 2026</h2>
  <div class="grid">
"""

for t in TOPICS:
    html += f'<div class="topic">\n'
    html += f'  <div class="topic-title">{t["num"]}. {t["en"]} <span class="ml-text">{t["ml"]}</span></div>\n'
    html += '</div>\n'

html += """
  </div>
</body>
</html>
"""

with open('printout.html', 'w', encoding='utf-8') as f:
    f.write(html)
