#!/usr/bin/env python3
"""Проверка метаданных с правильной кодировкой."""

file_path = 'NiPA-64-F-MD_1_18dBm_IMP_iPA-64-F_2026-03-03_17-42-42.txt'

with open(file_path, 'rb') as f:
    content = f.read()

# Декодируем с правильной кодировкой
text = content.decode('cp1251')
lines = text.split('\n')

print('Метаданные из файла:')
for line in lines[:10]:
    if line.strip():
        parts = line.split('\t')
        if len(parts) >= 2:
            print(f'  {parts[0].strip()}: {parts[1].strip()}')

print('\nЗаголовок таблицы:')
for line in lines:
    if 'f, MHz' in line or line.startswith('№'):
        print(f'  {line.strip()}')
        break
