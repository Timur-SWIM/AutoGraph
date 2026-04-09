#!/usr/bin/env python3
"""Проверка кодировки TXT файла."""

file_path = 'NiPA-64-F-MD_1_18dBm_IMP_iPA-64-F_2026-03-03_17-42-42.txt'

with open(file_path, 'rb') as f:
    content = f.read()

print('Первые 200 байт:', content[:200])
print()

# Пробуем разные кодировки
for encoding in ['utf-8', 'cp1251', 'windows-1251', 'latin-1']:
    try:
        decoded = content.decode(encoding)
        print(f'{encoding}: Успешно')
        print(f'  Первые 5 строк:')
        lines = decoded.split('\n')[:5]
        for i, line in enumerate(lines, 1):
            print(f'    {i}: {line[:80]}')
        print()
    except Exception as e:
        print(f'{encoding}: Ошибка - {e}')
