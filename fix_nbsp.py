# fix_nbsp.py
with open("parser.py", "r", encoding="utf-8") as f:
    text = f.read()

# заменяем все U+00A0 на обычный пробел
text = text.replace("\u00A0", " ")

with open("parser_fixed.py", "w", encoding="utf-8") as f:
    f.write(text)

print("Готово! Новый файл: parser_fixed.py")
