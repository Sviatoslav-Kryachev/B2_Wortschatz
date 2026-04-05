import re

with open('sheets_handler.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Удаляем slots=True из всех dataclass декораторов
content = re.sub(r'@dataclass\(slots=True\)', '@dataclass', content)

with open('sheets_handler.py', 'w', encoding='utf-8') as f:
    f.write(content)

print('Готово!')
