import pandas as pd
from pathlib import Path
import numpy as np

def excel_column_to_index(column_letter):
    """Конвертирует букву столбца Excel в индекс (0-based)."""
    column_letter = column_letter.upper()
    result = 0
    for char in column_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1  # 0-based индекс

def process_homework_scores(file_path, start_col='H', end_col='BB'):
    """
    Обрабатывает Excel-файл с оценками за домашние задания.
    Анализирует ДЗ от столбца H до столбца BB включительно.
    """
    try:
        print(f"Чтение файла: {file_path}")
        
        # Конвертируем буквы столбцов в индексы
        start_idx = excel_column_to_index(start_col)  # H → 7
        end_idx = excel_column_to_index(end_col)      # BB → 53
        
        print(f"Анализируем столбцы: {start_col}({start_idx+1}) до {end_col}({end_idx+1})")
        print(f"Всего столбцов для анализа: {end_idx - start_idx + 1}")
        
        # Читаем весь файл
        df = pd.read_excel(file_path, header=None)
        print(f"Размер файла: {df.shape[0]} строк, {df.shape[1]} столбцов")
        
        # Строка с максимальными баллами (7 строка Excel = индекс 6)
        max_scores_row = 6  # 7-я строка в Excel
        
        # Данные учеников с 8 строки Excel (индекс 7) по 34 строку (индекс 33)
        student_start_row = 7  # 8-я строка в Excel
        student_end_row = 33   # 34-я строка в Excel
        
        # Получаем максимальные баллы (столбцы H до BB включительно)
        max_scores_raw = df.iloc[max_scores_row, start_idx:end_idx+1]
        
        print(f"\nДиапазон столбцов: {start_col} до {end_col}")
        print(f"Индексы столбцов: {start_idx} до {end_idx}")
        print(f"Получено макс. баллов: {len(max_scores_raw)}")
        print(f"Первые 10 макс. баллов: {max_scores_raw.tolist()[:10]}")
        
        # Фильтруем ДЗ: игнорируем те, где максимальный балл = 1
        valid_columns = []
        valid_hw_names = []  # ДЗ-1, ДЗ-2, ДЗ-3 и т.д.
        valid_max_scores = []
        valid_excel_cols = []  # Сохраняем реальные буквы столбцов
        
        for i, max_score in enumerate(max_scores_raw):
            excel_col_idx = start_idx + i
            excel_col_letter = index_to_excel_column(excel_col_idx + 1)  # +1 для 1-based
            
            try:
                max_score_val = float(max_score) if pd.notna(max_score) else 0
                
                # Игнорируем ДЗ с максимальным баллом = 1
                if max_score_val != 1:
                    valid_columns.append(i)
                    hw_number = i + 1  # H → 1, I → 2, J → 3 и т.д.
                    valid_hw_names.append(f"ДЗ-{hw_number}")
                    valid_max_scores.append(max_score_val)
                    valid_excel_cols.append(excel_col_letter)
            except (ValueError, TypeError):
                continue
        
        print(f"\nДЗ после фильтрации (игнорируем макс. балл = 1): {len(valid_hw_names)}")
        print(f"Первые 10 ДЗ для анализа: {valid_hw_names[:10]}")
        print(f"Соответствующие столбцы Excel: {valid_excel_cols[:10]}")
        print(f"Первые 10 макс. баллов: {valid_max_scores[:10]}")
        
        if not valid_hw_names:
            print("Нет подходящих ДЗ для анализа!")
            return None, None, None, None
        
        # Собираем данные учеников (строки 8-34)
        results = []
        
        for row_idx in range(student_start_row, min(student_end_row + 1, len(df))):
            # Фамилия в столбце B (индекс 1)
            family_name = df.iloc[row_idx, 1]
            
            if pd.isna(family_name) or str(family_name).strip() == '':
                continue
            
            family_name = str(family_name).strip()
            
            # Собираем оценки ученика за каждое ДЗ
            student_scores = {}
            bonus_hw = []  # ДЗ, где ученик получил максимальный балл
            bonus_count = 0
            
            for i, hw_idx in enumerate(valid_columns):
                hw_name = valid_hw_names[i]
                max_score = valid_max_scores[i]
                excel_col = valid_excel_cols[i]
                col_idx = start_idx + hw_idx  # Реальный индекс столбца в DataFrame
                
                # Получаем оценку ученика
                score = df.iloc[row_idx, col_idx]
                
                try:
                    score_val = float(score) if pd.notna(score) else 0
                except:
                    score_val = 0
                
                student_scores[hw_name] = score_val
                
                # Проверяем, получил ли ученик максимальный балл за это ДЗ
                if score_val == max_score:
                    bonus_hw.append(hw_name)
                    bonus_count += 1
            
            # Считаем бонусные баллы (30 за каждое ДЗ с макс. баллом)
            bonus_points = bonus_count * 30
            
            # Преобразуем список ДЗ в строку
            if bonus_hw:
                bonus_hw_str = ', '.join(bonus_hw)
            else:
                bonus_hw_str = 'Нет'
            
            result = {
                'Фамилия': family_name,
                'ДЗ с макс. баллом': bonus_hw_str,
                'Количество таких ДЗ': bonus_count,
                'Начислено баллов': bonus_points,
                'Всего ДЗ': len(valid_hw_names)
            }
            
            # Добавляем информацию об оценках за первые 3 ДЗ для проверки
            for hw_name in valid_hw_names[:3]:
                result[f'{hw_name}_оценка'] = student_scores.get(hw_name, 0)
                result[f'{hw_name}_макс'] = valid_max_scores[valid_hw_names.index(hw_name)]
            
            results.append(result)
        
        print(f"\nОбработано учеников: {len(results)} (строки {student_start_row+1}-{student_end_row+1})")
        print(f"Диапазон столбцов: {start_col} до {end_col} ({len(max_scores_raw)} столбцов)")
        print(f"ДЗ для анализа: {len(valid_hw_names)} (игнорировано {len(max_scores_raw) - len(valid_hw_names)} с макс. баллом = 1)")
        
        if not results:
            print("Нет данных учеников для анализа!")
            return None, None, None, None
        
        results_df = pd.DataFrame(results)
        results_df = results_df.sort_values('Начислено баллов', ascending=False).reset_index(drop=True)
        
        return results_df, valid_hw_names, valid_max_scores, valid_excel_cols
        
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        import traceback
        traceback.print_exc()
        return None, None, None, None

def index_to_excel_column(index):
    """Конвертирует индекс (1-based) в букву столбца Excel."""
    result = ""
    while index > 0:
        index -= 1
        remainder = index % 26
        result = chr(65 + remainder) + result
        index = index // 26
    return result

def print_results(results_df, hw_names, max_scores, excel_cols):
    """
    Выводит результаты в удобном формате.
    """
    print("\n" + "=" * 100)
    print(f"МАКСИМАЛЬНЫЕ БАЛЛЫ ЗА ДЗ (столбцы {excel_cols[0]} до {excel_cols[-1]}):")
    print("Игнорируем ДЗ с макс. баллом = 1")
    print("-" * 100)
    
    # Показываем ДЗ группами по 8
    for i in range(0, len(hw_names), 8):
        group = hw_names[i:i+8]
        scores_group = max_scores[i:i+8]
        cols_group = excel_cols[i:i+8]
        
        for j, (hw, score, col) in enumerate(zip(group, scores_group, cols_group), i+1):
            print(f"{j:3}. {hw:<8} ({col:>3}): {score:4} баллов", end="   ")
            if (j - i) % 4 == 0:
                print()
        print()
    
    print("\n" + "=" * 100)
    print("РЕЗУЛЬТАТЫ УЧЕНИКОВ (ТОП-20):")
    print("-" * 100)
    print(f"{'№':<3} {'Фамилия':<20} {'ДЗ с макс. баллом':<40} {'Кол-во':<7} {'Баллы':<7}")
    print("-" * 100)
    
    # Показываем топ-20 учеников
    top_students = results_df.head(20)
    for i, (_, row) in enumerate(top_students.iterrows(), 1):
        hw_list = row['ДЗ с макс. баллом']
        
        if hw_list == 'Нет':
            hw_display = 'Нет'
        elif len(hw_list) <= 38:
            hw_display = hw_list
        else:
            # Показываем первые 4 ДЗ
            if ',' in hw_list:
                first_hw = hw_list.split(', ')[:4]
                hw_display = ', '.join(first_hw)
                total_hw = len(hw_list.split(', '))
                if total_hw > 4:
                    hw_display += f" (+ещё {total_hw - 4})"
            else:
                hw_display = hw_list[:35] + "..."
        
        print(f"{i:<3} {row['Фамилия']:<20} {hw_display:<40} {row['Количество таких ДЗ']:<7} {row['Начислено баллов']:<7}")
    
    print("-" * 100)
    
    total_students = len(results_df)
    total_bonus = results_df['Начислено баллов'].sum()
    students_with_bonus = (results_df['Начислено баллов'] > 0).sum()
    
    print(f"\nСТАТИСТИКА ПО ВСЕМ УЧЕНИКАМ ({total_students} чел.):")
    print(f"• Анализировали ДЗ: столбцы {excel_cols[0]}-{excel_cols[-1]} ({len(hw_names)} ДЗ)")
    print(f"• Учеников с бонусными баллами: {students_with_bonus} ({students_with_bonus/total_students*100:.1f}%)")
    print(f"• Сумма всех бонусных баллов: {total_bonus}")
    print(f"• Средний бонус на ученика: {total_bonus/total_students:.1f}")
    
    # Топ-5 учеников
    print(f"\nТОП-5 УЧЕНИКОВ:")
    top5 = results_df.head(5)
    for j, (_, student) in enumerate(top5.iterrows(), 1):
        print(f"{j}. {student['Фамилия']}: {student['Начислено баллов']} баллов "
              f"({student['Количество таких ДЗ']} ДЗ на макс. балл)")
    
    # Статистика по ДЗ
    print(f"\nСТАТИСТИКА ПО ДЗ:")
    print(f"• Всего ДЗ для анализа: {len(hw_names)}")
    if max_scores:
        print(f"• Диапазон макс. баллов: {min(max_scores)} - {max(max_scores)}")
        
        # Считаем распределение макс. баллов
        score_ranges = {
            '1 балл': 0,  # но мы их игнорируем
            '2-10 баллов': 0,
            '11-20 баллов': 0,
            '21-50 баллов': 0,
            '51-100 баллов': 0
        }
        
        for score in max_scores:
            if score == 1:
                score_ranges['1 балл'] += 1
            elif 2 <= score <= 10:
                score_ranges['2-10 баллов'] += 1
            elif 11 <= score <= 20:
                score_ranges['11-20 баллов'] += 1
            elif 21 <= score <= 50:
                score_ranges['21-50 баллов'] += 1
            elif score > 50:
                score_ranges['51-100 баллов'] += 1
        
        print(f"• Распределение макс. баллов:")
        for range_name, count in score_ranges.items():
            if count > 0:
                percentage = count / len(max_scores) * 100
                print(f"  {range_name}: {count} ДЗ ({percentage:.1f}%)")

def main():
    """
    Основная функция программы.
    """
    print("ПРОГРАММА ДЛЯ АНАЛИЗА ОЦЕНОК ЗА ДОМАШНИЕ ЗАДАНИЯ")
    print("=" * 80)
    print("ВАЖНО:")
    print("1. ДЗ с максимальным баллом = 1 игнорируются!")
    print("2. Анализируем ДЗ от столбца H до столбца BB включительно")
    print("3. Нумерация ДЗ: H → ДЗ-1, I → ДЗ-2, J → ДЗ-3 и т.д.")
    print("=" * 80)
    
    # Автоматически ищем файл
    file_paths = [
        "/Users/daniltotoev/Downloads/Новая таблица.xlsx",
        "Новая таблица.xlsx",
        "data/input/Новая таблица.xlsx"
    ]
    
    file_path = None
    for path in file_paths:
        if Path(path).exists():
            file_path = path
            print(f"\nНайден файл: {path}")
            break
    
    if not file_path:
        print("\nФайл не найден по стандартным путям.")
        user_path = input("Введите полный путь к файлу: ").strip()
        if Path(user_path).exists():
            file_path = user_path
        else:
            print(f"Файл '{user_path}' не найден!")
            return
    
    # Обработка файла с ограничением по столбцам
    results_df, hw_names, max_scores, excel_cols = process_homework_scores(
        file_path, 
        start_col='H', 
        end_col='BB'
    )
    
    if results_df is not None and len(results_df) > 0:
        print_results(results_df, hw_names, max_scores, excel_cols)
        
        # Сохранение результатов
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
        output_file = f'результаты_H_to_BB_{timestamp}.csv'
        results_df.to_csv(output_file, index=False, encoding='utf-8-sig')
        print(f"\nРезультаты сохранены в файл: {output_file}")
        
        # Создаем Excel файл
        excel_file = f'отчет_H_to_BB_{timestamp}.xlsx'
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            # Основные результаты
            results_df.to_excel(writer, sheet_name='Результаты', index=False)
            
            # Информация о ДЗ
            hw_info = pd.DataFrame({
                'Номер ДЗ': [name.replace('ДЗ-', '') for name in hw_names],
                'Название ДЗ': hw_names,
                'Столбец Excel': excel_cols,
                'Максимальный балл': max_scores,
                'Бонус за макс. балл': [30] * len(hw_names)
            })
            hw_info.to_excel(writer, sheet_name='Инфо о ДЗ', index=False)
            
            # Подробные данные
            detailed_df = results_df.copy()
            detailed_df['% ДЗ на макс'] = (detailed_df['Количество таких ДЗ'] / detailed_df['Всего ДЗ'] * 100).round(1)
            detailed_df['Рейтинг'] = range(1, len(detailed_df) + 1)
            detailed_df.to_excel(writer, sheet_name='Детали', index=False)
        
        print(f"Полный отчет сохранен в: {excel_file}")
        
        # Дополнительная информация для проверки
        print("\n" + "=" * 80)
        print("ПРОВЕРОЧНЫЕ ДАННЫЕ (первые 3 ученика):")
        print("=" * 80)
        
        for i in range(min(3, len(results_df))):
            student = results_df.iloc[i]
            print(f"\n{student['Фамилия']} (место #{i+1}):")
            print(f"  Бонусных баллов: {student['Начислено баллов']}")
            print(f"  ДЗ на макс. балл: {student['Количество таких ДЗ']} из {student['Всего ДЗ']}")
            
            # Показываем первые 3 ДЗ для проверки
            for hw in hw_names[:3]:
                score_key = f'{hw}_оценка'
                max_key = f'{hw}_макс'
                if score_key in student:
                    score = student[score_key]
                    max_val = student.get(max_key, '?')
                    status = "✓ МАКС" if score == max_val else f"{score}/{max_val}"
                    print(f"  {hw}: {status}")
        
        print("\n" + "=" * 80)
        print("ИНФОРМАЦИЯ О ДИАПАЗОНЕ СТОЛБЦОВ:")
        print("=" * 80)
        print(f"Первый столбец ДЗ: {excel_cols[0]} (ДЗ-1)")
        print(f"Последний столбец ДЗ: {excel_cols[-1]} (ДЗ-{len(hw_names)})")
        print(f"Всего проанализировано ДЗ: {len(hw_names)}")
        
    else:
        print("\nНе удалось получить результаты. Проверьте структуру файла.")

if __name__ == "__main__":
    main()