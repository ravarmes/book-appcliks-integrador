import glob, re
o = open('C:/Users/ravar/.gemini/antigravity/brain/055dec35-13fc-41ad-a34b-d73bf2b95695/fig_report.md', 'w', encoding='utf-8')
for f in sorted(glob.glob('sem*.html')):
    text = open(f, encoding='utf-8').read()
    o.write(f"## {f}\n")
    matches = re.finditer(r'<img\s+src=[\'"](.*?)[\'"][\s\S]*?(Figura \d+\.\d+)', text, re.IGNORECASE)
    for m in matches:
        o.write(f"- {m.group(2)} => {m.group(1)}\n")
    o.write("\n")
o.close()
