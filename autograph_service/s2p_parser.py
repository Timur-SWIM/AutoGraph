"""
Модуль для парсинга S2P файлов (Touchstone формат).
S2P файлы содержат S-параметры двухпортовых сетей.

Формат заголовка: # [частота] [параметр] [формат] R [сопротивление]
Пример: #  HZ   S   RI   R     50.00

Данные: частота re:S11 im:S11 re:S21 im:S21 re:S12 im:S12 re:S22 im:S22
"""

import re
import math
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass


@dataclass
class S2PData:
    """Класс для хранения данных S2P файла."""
    metadata: Dict[str, str]
    headers: List[str]
    rows: List[List[str]]


class S2PParser:
    """Парсер для S2P файлов (Touchstone формат)."""
    
    def __init__(self):
        self.metadata: Dict[str, str] = {}
        self.headers: List[str] = []
        self.rows: List[List[str]] = []
        self.frequency_unit: str = "GHZ"
        self.parameter_type: str = "S"
        self.data_format: str = "RI"  # RI = Real/Imaginary, MA = Magnitude/Angle, DB = dB/Angle
        self.resistance: float = 50.0
    
    def parse_file(self, file_path: str) -> S2PData:
        """
        Парсит S2P файл и возвращает структурированные данные.
        
        Args:
            file_path: Путь к S2P файлу
            
        Returns:
            S2PData с метаданными, заголовками и строками данных
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
        except FileNotFoundError:
            raise FileNotFoundError(f"Файл не найден: {file_path}")
        except IOError as e:
            raise IOError(f"Ошибка чтения файла: {e}")
        
        self.metadata = {}
        self.headers = []
        self.rows = []
        
        self._parse_lines(lines)
        
        return S2PData(
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
            
            # Пропускаем комментарии
            if line.startswith('!'):
                self._parse_comment(line)
                continue
            
            # Проверяем строку опций (начинается с #)
            if line.startswith('#'):
                self._parse_options(line)
                continue
            
            # Это строка данных
            self._parse_data_line(line)
    
    def _parse_comment(self, line: str) -> None:
        """Парсит строку комментария."""
        # Сохраняем комментарии в метаданные
        comment = line[1:].strip()  # Убираем '!'
        if 'comment' not in self.metadata:
            self.metadata['comment'] = comment
        else:
            self.metadata['comment'] += '; ' + comment
    
    def _parse_options(self, line: str) -> None:
        """Парсит строку опций (# GHZ S RI R 50)."""
        parts = line[1:].split()  # Убираем '#'
        
        if len(parts) >= 4:
            self.frequency_unit = parts[0].upper()
            self.parameter_type = parts[1].upper()
            self.data_format = parts[2].upper()
            
            # Ищем сопротивление
            for i, part in enumerate(parts):
                if part.upper() == 'R' and i + 1 < len(parts):
                    try:
                        self.resistance = float(parts[i + 1])
                    except ValueError:
                        pass
        
        self.metadata['frequency_unit'] = self.frequency_unit
        self.metadata['parameter_type'] = self.parameter_type
        self.metadata['data_format'] = self.data_format
        self.metadata['resistance'] = str(self.resistance)
    
    def _parse_data_line(self, line: str) -> None:
        """Парсит строку данных."""
        parts = line.split()
        
        if len(parts) >= 9:  # Частота + 8 значений S-параметров
            freq = parts[0]
            
            # Преобразуем частоту в ГГц
            freq_ghz = self._convert_frequency_to_ghz(float(freq))
            
            # Извлекаем S-параметры
            s11_re, s11_im = float(parts[1]), float(parts[2])
            s21_re, s21_im = float(parts[3]), float(parts[4])
            s12_re, s12_im = float(parts[5]), float(parts[6])
            s22_re, s22_im = float(parts[7]), float(parts[8])
            
            # Вычисляем магнитуды в дБ
            s11_db = self._complex_to_db(s11_re, s11_im)
            s21_db = self._complex_to_db(s21_re, s21_im)
            s12_db = self._complex_to_db(s12_re, s12_im)
            s22_db = self._complex_to_db(s22_re, s22_im)
            
            # Формируем строку данных для шаблона
            row = [
                str(freq_ghz),
                str(s11_db),
                str(s21_db),
                str(s12_db),
                str(s22_db)
            ]
            
            self.rows.append(row)
    
    def _convert_frequency_to_ghz(self, freq: float) -> float:
        """Преобразует частоту в ГГц."""
        unit = self.frequency_unit.upper()
        if unit == 'HZ':
            return freq / 1e9
        elif unit == 'KHZ':
            return freq / 1e6
        elif unit == 'MHZ':
            return freq / 1e3
        elif unit == 'GHZ':
            return freq
        else:
            return freq
    
    def _complex_to_db(self, real: float, imag: float) -> float:
        """Преобразует комплексное число в дБ (магнитуда)."""
        magnitude = math.sqrt(real**2 + imag**2)
        if magnitude > 0:
            return 20 * math.log10(magnitude)
        else:
            return -1000.0  # Очень маленькое значение для нулевой магнитуды
    
    def get_table_data(self) -> Tuple[List[str], List[List[str]]]:
        """
        Возвращает только табличные данные (заголовки и строки).
        
        Returns:
            Кортеж (headers, rows)
        """
        return self.headers, self.rows
    
    def get_headers_for_template(self) -> List[str]:
        """Возвращает заголовки для шаблона Excel."""
        return ['Freq, GHz', 'S11M', 'S21M', 'S12M', 'S22M']
