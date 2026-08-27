import re
h = open('e101_rebuilt.html', encoding='utf-8').read()
d = open('artifact_data.js', encoding='utf-8').read()
h = h.replace('<script src="__DATA__"></script>', '<script>' + d + '</script>')

parts = re.split(r'(<script>.*?</script>|<script src[^>]*></script>)', h, flags=re.S)
out = []
for p in parts:
    if p.startswith('<script'):
        out.append(re.sub(r'[^\x00-\x7f]', lambda m: '\\u%04x' % ord(m.group()), p))
    else:
        out.append(re.sub(r'[^\x00-\x7f]', lambda m: '&#x%x;' % ord(m.group()), p))
h = ''.join(out)
with open('e101_rebuilt_final.html', 'w', encoding='ascii') as f:
    f.write(h)
import os
print('final:', os.path.getsize('e101_rebuilt_final.html'), 'ascii-clean:', not re.search(r'[^\x00-\x7f]', h))
