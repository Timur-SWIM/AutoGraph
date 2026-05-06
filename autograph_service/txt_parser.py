"""
Модуль для парсинга TXT файлов с измерительными данными.
Формат файла: NiPA-64-F-MD_1_18dBm_IMP_iPA-64-F_2026-03-03_17-42-42.txt
"""

import re
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class MeasurementData:
    """Класс для хранения измерительных данных."""
    metadata: Dict[str, str]
    headers: List[str]
    rows: List[List[str]]


class TxtFileParser:
    """Парсер для TXT файлов с измерительными данными."""
    
    def __init__(self):
        self.metadata: Dict[str, str] = {}
        self.headers: List[str] = []
        self.rows: List[List[str]] = []
    
    def parse_file(self, file_path: str) -> MeasurementData:
        """
        Парсит TXT файл и возвращает структурированные данные.
        
        Args:
            file_path: Путь к TXT файлу
            
        Returns:
            MeasurementData с метаданными, заголовками и строками данных
        """
        try:
            # Файлы в кодировке Windows-1251 (кириллица)
            with open(file_path, 'r', encoding='cp1251') as f:
                lines = f.readlines()
        except FileNotFoundError:
            raise FileNotFoundError(f"Файл не найден: {file_path}")
        except IOError as e:
            raise IOError(f"Ошибка чтения файла: {e}")
        
        self.metadata = {}
        self.headers = []
        self.rows = []
        
        self._parse_lines(lines)
        
        return MeasurementData(
            metadata=self.metadata,
            headers=self.headers,
            rows=self.rows
        )
    
    def _parse_lines(self, lines: List[str]) -> None:
        """Парсит все строки файла."""
        header_found = False
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            if not header_found:
                if self._is_metadata_line(line):
                    self._parse_metadata_line(line)
                elif 'f, MHz' in line or line.startswith('№'):
                    # Это строка заголовков
                    header_found = True
                    self._parse_header_line(line)
                else:
                    # Это строка данных (до заголовков - пропускаем)
                    pass
            else:
                # Это строка данных
                self._parse_data_line(line)
    
    def _is_metadata_line(self, line: str) -> bool:
        """Определяет, является ли строка метаданными."""
        # Строка заголовков таблицы содержит "f, MHz" или начинается с "№"
        # или первая колонка - это номер строки данных (число с точкой)
        if 'f, MHz' in line or line.startswith('№'):
            return False
        # Проверяем, является ли первая колонка числом (строка данных)
        parts = line.split('\t')
        if len(parts) >= 2:
            first_part = parts[0].strip().replace(',', '.')
            try:
                float(first_part)
                return False  # Это строка данных, а не метаданные
            except ValueError:
                pass
        return True
    
    def _parse_metadata_line(self, line: str) -> None:
        """Парсит строку метаданных."""
        parts = line.split('\t')
        if len(parts) >= 2:
            key = parts[0].strip()
            value = parts[1].strip()
            if key:
                self.metadata[key] = value
    
    def _parse_header_line(self, line: str) -> None:
        """Парсит строку заголовков таблицы."""
        # Заголовки разделены табуляцией
        # Формат: №\tf, MHz\tP, dBm\tIG, mA\tID, A\tGain, dB\tКПД, %\tPвых, W
        self.headers = [h.strip() for h in line.split('\t')]
    
    def _parse_data_line(self, line: str) -> None:
        """Парсит строку данных таблицы."""
        parts = [p.strip() for p in line.split('\t')]
        if len(parts) >= 2:  # Минимум 2 колонки
            self.rows.append(parts)
    
    def get_table_data(self) -> Tuple[List[str], List[List[str]]]:
        """
        Возвращает только табличные данные (заголовки и строки).
        
        Returns:
            Кортеж (headers, rows)
        """
        return self.headers, self.rows
