import sys, zipfile, re, shutil
src = sys.argv[1]; tmp = src + ".tmp"
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
    for it in zi.infolist():
        data = zi.read(it.filename)
        if re.match(r"ppt/theme/theme\d+\.xml$", it.filename):
            s = data.decode("utf-8")
            s = s.replace('<a:cs typeface=""/>', '<a:cs typeface="Vazirmatn"/>')
            s = re.sub(r'<a:font script="Arab" typeface="[^"]*"/>', '<a:font script="Arab" typeface="Vazirmatn"/>', s)
            data = s.encode("utf-8")
        zo.writestr(it, data)
shutil.move(tmp, src)
print("patched", src)
