from fractions import Fraction
from docx import Document

# Исходная матрица выигрышей первого игрока.
A = [
    [22, 23, 23, 24, 9, 11],
    [27, 4, 4, 22, 9, 18],
    [2, 15, 28, 3, 28, 9],
    [3, 25, 11, 22, 20, 5],
    [22, 10, 9, 10, 28, 6]
]

# Настройка выбора уравнения, из которого выражается V (нумерация с 0 для сокращённой матрицы):
V_COL_INDEX_P1 = 0  # Из какого столбца (ограничения) выражать V для 1-го игрока (0 = 1-е уравнение)
V_ROW_INDEX_P2 = 0  # Из какой строки (ограничения) выражать V для 2-го игрока (0 = 1-е уравнение)


def print_table(table, row_labels, col_labels, step_title, crossed_columns=None):
    """Красивый вывод симплекс-таблицы Жордана в консоль."""
    crossed_columns = set(crossed_columns or [])
    print(f"\n{'=' * 75}")
    print(step_title)
    print(f"{'=' * 75}")
    header = (
            f"{'Базис':<8} | "
            + " | ".join(f"{str(c):>10}" for c in col_labels)
            + " |"
    )
    print(header)
    print("-" * len(header))
    for r_lbl, row in zip(row_labels, table):
        row_str = (
                f"{r_lbl:<8} | "
                + " | ".join(
            f"{('×' if j in crossed_columns else val):>10}"
            for j, val in enumerate(row)
        )
                + " |"
        )
        print(row_str)


def add_table_to_docx(
        document, table, row_labels, col_labels, step_title, crossed_columns=None
):
    """Добавляет одну симплекс-таблицу в DOCX-документ."""
    crossed_columns = set(crossed_columns or [])
    document.add_heading(step_title, level=2)
    docx_table = document.add_table(rows=1, cols=len(col_labels) + 1)
    docx_table.style = "Table Grid"

    header_cells = docx_table.rows[0].cells
    header_cells[0].text = "Базис"
    for j, label in enumerate(col_labels, start=1):
        header_cells[j].text = str(label)

    for row_label, row in zip(row_labels, table):
        cells = docx_table.add_row().cells
        cells[0].text = str(row_label)
        for j, value in enumerate(row, start=1):
            cells[j].text = (
                "×" if j - 1 in crossed_columns else str(value)
            )

    document.add_paragraph()


def format_linear_expression(coefficients, variables):
    """Форматирует линейное выражение из коэффициентов и переменных."""
    terms = []
    for coefficient, variable in zip(coefficients, variables):
        if coefficient == 0:
            continue
        sign = "-" if coefficient < 0 else "+"
        absolute = abs(coefficient)
        coefficient_text = "" if absolute == 1 else f"{absolute}·"
        term = f"{coefficient_text}{variable}"
        if not terms:
            terms.append(f"-{term}" if coefficient < 0 else term)
        else:
            terms.append(f" {sign} {term}")
    return "".join(terms) if terms else "0"


def build_equation_section(
        matrix,
        row_indices,
        col_indices,
        selected_index,
        player,
        original_row_count,
        original_column_count,
):
    """Формирует заголовок и строки системы уравнений."""
    lines = ["Исходная система:"]
    if player == 1:
        variables = [f"x_{index + 1}" for index in row_indices]
        slack_start = original_row_count
        for column_position, column_index in enumerate(col_indices):
            slack = f"x_{slack_start + column_position + 1}"
            expression = format_linear_expression(
                [matrix[row][column_index] for row in row_indices] + [-1],
                variables + [slack],
            )
            lines.append(f"{expression} = V")
        lines.append(
            f"{format_linear_expression([1] * len(variables), variables)} = 1"
        )
        lines.append(
            f"Выражаем V из уравнения для столбца B{selected_index + 1} "
            "и вычитаем его из остальных уравнений:"
        )
        selected_values = [matrix[row][selected_index] for row in row_indices]
        for column_index in col_indices:
            if column_index == selected_index:
                continue
            difference = [
                matrix[row][column_index] - selected_value
                for row, selected_value in zip(row_indices, selected_values)
            ]
            current_slack = f"x_{slack_start + column_index + 1}"
            selected_slack = f"x_{slack_start + selected_index + 1}"
            lines.append(
                f"{format_linear_expression(difference + [-1, 1], variables + [current_slack, selected_slack])}"
                f" = 0"
            )
    else:
        variables = [f"y_{index + 1}" for index in col_indices]
        slack_start = original_column_count
        for row_position, row_index in enumerate(row_indices):
            slack = f"y_{slack_start + row_position + 1}"
            expression = format_linear_expression(
                [matrix[row_index][column] for column in col_indices] + [1],
                variables + [slack],
            )
            lines.append(f"{expression} = V")
        lines.append(
            f"{format_linear_expression([1] * len(variables), variables)} = 1"
        )
        lines.append(
            f"Выражаем V из уравнения для строки A{selected_index + 1} "
            "и вычитаем его из остальных уравнений:"
        )
        selected_values = matrix[selected_index]
        for row_index in row_indices:
            if row_index == selected_index:
                continue
            difference = [
                matrix[row_index][column] - selected_values[column]
                for column in col_indices
            ]
            current_slack = f"y_{slack_start + row_indices.index(row_index) + 1}"
            selected_slack = f"y_{slack_start + row_indices.index(selected_index) + 1}"
            lines.append(
                f"{format_linear_expression(difference + [1, -1], variables + [current_slack, selected_slack])}"
                f" = 0"
            )
    return lines


def add_equations_to_docx(document, title, *args):
    """Добавляет исходную и преобразованную систему уравнений в DOCX."""
    document.add_heading(title, level=1)
    for line in build_equation_section(*args):
        document.add_paragraph(line)


def print_equations(title, *args):
    """Выводит исходную и преобразованную систему уравнений в терминал."""
    print(f"\n{title}")
    print("-" * len(title))
    for line in build_equation_section(*args):
        print(line)


def save_tables_to_docx(document, file_name="jordan_tables.docx"):
    """Сохраняет накопленные симплекс-таблицы в DOCX-файл."""
    try:
        document.save(file_name)
    except OSError as error:
        raise RuntimeError(f"Не удалось сохранить таблицы в '{file_name}': {error}") from error


def validate_matrix(matrix):
    """Проверяет, что матрица непустая и прямоугольная."""
    if not matrix or not matrix[0]:
        raise ValueError("Матрица выигрышей не должна быть пустой.")

    column_count = len(matrix[0])
    if any(len(row) != column_count for row in matrix):
        raise ValueError("Все строки матрицы выигрышей должны иметь одинаковую длину.")


def reduce_matrix(matrix):
    """Проверка седловой точки и итерационное сокращение доминируемых строк и столбцов."""
    validate_matrix(matrix)
    m_orig = len(matrix)
    n_orig = len(matrix[0])

    row_mins = [min(row) for row in matrix]
    col_maxs = [max(matrix[i][j] for i in range(m_orig)) for j in range(n_orig)]
    alpha = max(row_mins)
    beta = min(col_maxs)

    print(f"\n{'=' * 75}")
    print("ПРЕДВАРИТЕЛЬНЫЙ АНАЛИЗ И СОКРАЩЕНИЕ МАТРИЦЫ")
    print(f"{'=' * 75}")
    print(f"Нижняя цена игры α = {alpha}, верхняя цена игры β = {beta}")

    if alpha == beta:
        print(f"Найдена седловая точка! Цена игры V = {alpha} (решение в чистых стратегиях).")
        r_idx = row_mins.index(alpha)
        c_idx = col_maxs.index(beta)
        X_full = [Fraction(1 if i == r_idx else 0) for i in range(m_orig)]
        Y_full = [Fraction(1 if j == c_idx else 0) for j in range(n_orig)]
        return None, None, None, Fraction(alpha), X_full, Y_full

    active_rows = list(range(m_orig))
    active_cols = list(range(n_orig))
    changed = True
    any_removed = False

    while changed:
        changed = False
        # Проверка строк (1-й игрок: удаляем строку r1, если она <= r2 по всем активным столбцам)
        for r1 in list(active_rows):
            for r2 in list(active_rows):
                if r1 != r2 and r1 in active_rows and r2 in active_rows:
                    if all(matrix[r1][c] <= matrix[r2][c] for c in active_cols):
                        active_rows.remove(r1)
                        print(f"-> Строка A{r1 + 1} доминируется строкой A{r2 + 1} (удаляем A{r1 + 1}).")
                        changed = True
                        any_removed = True
                        break

        # Проверка столбцов (2-й игрок: удаляем столбец c1, если он >= c2 по всем активным строкам)
        for c1 in list(active_cols):
            for c2 in list(active_cols):
                if c1 != c2 and c1 in active_cols and c2 in active_cols:
                    if all(matrix[r][c1] >= matrix[r][c2] for r in active_rows):
                        active_cols.remove(c1)
                        print(f"-> Столбец B{c1 + 1} доминируется столбцом B{c2 + 1} (удаляем B{c1 + 1}).")
                        changed = True
                        any_removed = True
                        break

    if not any_removed:
        print("Доминируемых строк и столбцов нет. Матрица остаётся без изменений.")
    else:
        print(f"Матрица сокращена до размера {len(active_rows)}x{len(active_cols)}.")

    reduced = [[matrix[r][c] for c in active_cols] for r in active_rows]

    if len(reduced) == 1 and len(reduced[0]) == 1:
        val = Fraction(reduced[0][0])
        X_full = [Fraction(1 if i == active_rows[0] else 0) for i in range(m_orig)]
        Y_full = [Fraction(1 if j == active_cols[0] else 0) for j in range(n_orig)]
        print(f"После сокращения осталась матрица 1x1. Цена игры V = {val}.")
        return None, None, None, val, X_full, Y_full

    return reduced, active_rows, active_cols, None, None, None


def jordan_step(table, row_labels, col_labels, pivot_row, pivot_col):
    """Один шаг жордановых исключений по правилам из лекции."""
    rows = len(table)
    cols = len(table[0])
    pivot = table[pivot_row][pivot_col]

    new_table = [[Fraction(0) for _ in range(cols)] for _ in range(rows)]

    # Разрешающий элемент: 1 / a_ks.
    new_table[pivot_row][pivot_col] = Fraction(1) / pivot

    # Остальные значения разрешающей строки: a_kj / a_ks.
    for j in range(cols):
        if j != pivot_col:
            new_table[pivot_row][j] = table[pivot_row][j] / pivot

    # Остальные значения разрешающего столбца: -a_is / a_ks.
    for i in range(rows):
        if i != pivot_row:
            new_table[i][pivot_col] = -table[i][pivot_col] / pivot

    # Остальные элементы: (a_ij * a_ks - a_is * a_kj) / a_ks.
    for i in range(rows):
        for j in range(cols):
            if i != pivot_row and j != pivot_col:
                new_table[i][j] = (
                                          table[i][j] * pivot
                                          - table[i][pivot_col] * table[pivot_row][j]
                                  ) / pivot

    old_row_lbl = row_labels[pivot_row]
    old_col_lbl = col_labels[pivot_col]

    new_row_labels = list(row_labels)
    new_col_labels = list(col_labels)

    new_row_labels[pivot_row] = old_col_lbl.replace("-", "") + " ="

    if old_row_lbl.startswith("0"):
        new_col_labels[pivot_col] = "0"
    elif old_row_lbl.startswith("1"):
        new_col_labels[pivot_col] = "-1"
    else:
        new_col_labels[pivot_col] = "-" + old_row_lbl.replace(" =", "")

    return new_table, new_row_labels, new_col_labels


def remove_zero_columns(table, col_labels):
    """Удаляет столбцы, над которыми в предыдущей таблице появился '0'."""
    keep_indices = [j for j, lbl in enumerate(col_labels) if lbl != "0"]
    new_col_labels = [col_labels[j] for j in keep_indices]
    new_table = [[row[j] for j in keep_indices] for row in table]
    return new_table, new_col_labels


def extract_strategy_from_table(
    table, row_labels, col_labels, variable_prefix, variable_numbers, coefficient_sign=1
):
    """Восстанавливает вероятности переменных из итоговой таблицы."""
    result = []
    objective_row = table[-1]
    for variable_number in variable_numbers:
        column_label = f"-{variable_prefix}_{variable_number}"
        value = Fraction(0)
        for index, label in enumerate(col_labels):
            if label == column_label:
                value = coefficient_sign * objective_row[index]
                break
        result.append(value)
    return result


def solve_jordan_simplex(
        table, row_labels, col_labels, player_name, is_min=False, document=None
):
    """Полный цикл решения методом жордановых исключений до оптимальности V-строки."""
    step = 1
    crossed_columns = set()
    print(f"\n\n{'#' * 75}")
    print(f"РЕШЕНИЕ ДЛЯ: {player_name}")
    print(f"{'#' * 75}")
    print_table(
        table,
        row_labels,
        col_labels,
        f"Таблица {step} (Начальная)",
        crossed_columns=set(),
    )
    if document is not None:
        document.add_heading(player_name, level=1)
        add_table_to_docx(
            document,
            table,
            row_labels,
            col_labels,
            f"Таблица {step} (Начальная)",
            crossed_columns=set(),
        )

    while True:
        v_row = table[-1]
        pivot_col = -1

        if not is_min:
            # 1-й игрок (максимизация V): ищем отрицательный элемент в V-строке
            best_val = Fraction(0)
            for j in range(1, len(col_labels)):
                if col_labels[j] not in ("0", "-1") and v_row[j] < best_val:
                    best_val = v_row[j]
                    pivot_col = j
        else:
            # 2-й игрок (минимизация V): сначала вводим базис (пока не заменится '1 ='),
            # а затем уменьшаем V, выбирая столбцы с положительными оценками (кроме дополнительных y)
            has_one_row = any(r.startswith("1") for r in row_labels[:-1])
            if has_one_row:
                best_val = Fraction(0)
                for j in range(1, len(col_labels)):
                    if col_labels[j] not in ("0", "-1") and v_row[j] < best_val:
                        best_val = v_row[j]
                        pivot_col = j
            else:
                best_val = Fraction(0)
                for j in range(1, len(col_labels)):
                    if col_labels[j] not in ("0", "-1") and v_row[j] > best_val:
                        best_val = v_row[j]
                        pivot_col = j

        if pivot_col == -1:
            print("\n>>> Оптимальное решение достигнуто!")
            break

        pivot_row = -1
        min_ratio = None
        for i in range(len(table) - 1):
            elem = table[i][pivot_col]
            if elem > 0:
                ratio = table[i][0] / elem
                if min_ratio is None or ratio < min_ratio:
                    min_ratio = ratio
                    pivot_row = i

        if pivot_row == -1:
            print("Задача не имеет ограниченного решения!")
            break

        print(
            f"\n-> Шаг {step}: разрешающий столбец '{col_labels[pivot_col]}', "
            f"разрешающая строка '{row_labels[pivot_row]}' "
            f"(элемент = {table[pivot_row][pivot_col]})"
        )

        previous_col_labels = list(col_labels)
        table, row_labels, col_labels = jordan_step(
            table, row_labels, col_labels, pivot_row, pivot_col
        )
        step += 1
        # Столбец, получивший метку 0 на текущем шаге, ещё участвует
        # в расчёте и начинает отображаться крестиком только в следующей таблице.
        if step >= 3:
            crossed_columns.update(
                index
                for index, label in enumerate(previous_col_labels)
                if label == "0"
            )
        print_table(
            table,
            row_labels,
            col_labels,
            f"Таблица {step}",
            crossed_columns=crossed_columns,
        )
        if document is not None:
            add_table_to_docx(
                document,
                table,
                row_labels,
                col_labels,
                f"Таблица {step}",
                crossed_columns=crossed_columns,
            )

    return table, row_labels, col_labels


if __name__ == "__main__":
    validate_matrix(A)
    original_row_count = len(A)
    original_column_count = len(A[0])
    reduced_A, act_rows, act_cols, pure_V, X_opt, Y_opt = reduce_matrix(A)

    if reduced_A is not None:
        tables_document = Document()
        tables_document.add_heading("Таблицы метода Жордана", level=0)
        m, n = len(reduced_A), len(reduced_A[0])

        # Проверяем корректность выбранного индекса уравнения V
        v_col_p1 = V_COL_INDEX_P1 if 0 <= V_COL_INDEX_P1 < n else 0
        v_row_p2 = V_ROW_INDEX_P2 if 0 <= V_ROW_INDEX_P2 < m else 0

        # =====================================================================
        # 1. ПОСТРОЕНИЕ НАЧАЛЬНОЙ ТАБЛИЦЫ ДЛЯ ПЕРВОГО ИГРОКА (МАКСИМИЗАЦИЯ V)
        # =====================================================================
        table_p1 = []
        for j in range(n):
            if j == v_col_p1:
                continue
            row = [Fraction(0)]
            for i in range(m):
                row.append(Fraction(-(reduced_A[i][j] - reduced_A[i][v_col_p1])))
            slacks = [Fraction(0)] * n
            slacks[v_col_p1] = Fraction(-1)  # дополнительная переменная из уравнения V
            slacks[j] = Fraction(1)  # дополнительная переменная из текущего уравнения
            table_p1.append(row + slacks)

        # Строка нормировки 1 =
        table_p1.append([Fraction(1)] + [Fraction(1)] * m + [Fraction(0)] * n)

        # Целевая строка V =
        v_slacks_p1 = [Fraction(0)] * n
        v_slacks_p1[v_col_p1] = Fraction(1)
        v_p1 = (
                [Fraction(0)]
                + [Fraction(-reduced_A[i][v_col_p1]) for i in range(m)]
                + v_slacks_p1
        )
        table_p1.append(v_p1)

        rows_p1 = ["0 ="] * (n - 1) + ["1 =", "V ="]
        x_names = [f"x_{r + 1}" for r in act_rows] + [
            f"x_{original_row_count + k + 1}" for k in range(n)
        ]
        cols_p1 = ["1"] + [f"-{name}" for name in x_names]
        print_equations(
            "Система уравнений для первого игрока",
            reduced_A,
            list(range(m)),
            list(range(n)),
            v_col_p1,
            1,
            original_row_count,
            original_column_count,
        )
        add_equations_to_docx(
            tables_document,
            "Система уравнений для первого игрока",
            reduced_A,
            list(range(m)),
            list(range(n)),
            v_col_p1,
            1,
            original_row_count,
            original_column_count,
        )

        fin_t1, fin_r1, fin_c1 = solve_jordan_simplex(
            table_p1,
            rows_p1,
            cols_p1,
            f"ПЕРВЫЙ ИГРОК (V выражено из {v_col_p1 + 1}-го уравнения)",
            is_min=False,
            document=tables_document,
        )

        # =====================================================================
        # 2. ПОСТРОЕНИЕ НАЧАЛЬНОЙ ТАБЛИЦЫ ДЛЯ ВТОРОГО ИГРОКА (МИНИМИЗАЦИЯ V)
        # =====================================================================
        # Для 2-го игрока решаем задачу минимизации V по тому же критерию неотрицательности V-строки,
        # задавая целевую функцию W = -V -> max (итоговое V2 = -W_max)
        table_p2 = []
        for i in range(m):
            if i == v_row_p2:
                continue
            row = [Fraction(0)]
            for j in range(n):
                row.append(Fraction(-(reduced_A[i][j] - reduced_A[v_row_p2][j])))
            slacks = [Fraction(0)] * m
            slacks[v_row_p2] = Fraction(1)  # дополнительная переменная из уравнения V
            slacks[i] = Fraction(-1)  # дополнительная переменная из текущего уравнения
            table_p2.append(row + slacks)

        # Строка нормировки 1 =
        table_p2.append([Fraction(1)] + [Fraction(1)] * n + [Fraction(0)] * m)

        # Целевая строка V = (по образцу Таблицы 7 лекции)
        v_slacks_p2 = [Fraction(0)] * m
        v_slacks_p2[v_row_p2] = Fraction(-1)
        v_p2 = (
                [Fraction(0)]
                + [Fraction(-reduced_A[v_row_p2][j]) for j in range(n)]
                + v_slacks_p2
        )
        table_p2.append(v_p2)

        rows_p2 = ["0 ="] * (m - 1) + ["1 =", "V ="]
        y_names = [f"y_{c + 1}" for c in act_cols] + [
            f"y_{original_column_count + k + 1}" for k in range(m)
        ]
        cols_p2 = ["1"] + [f"-{name}" for name in y_names]
        print_equations(
            "Система уравнений для второго игрока",
            reduced_A,
            list(range(m)),
            list(range(n)),
            v_row_p2,
            2,
            original_row_count,
            original_column_count,
        )
        add_equations_to_docx(
            tables_document,
            "Система уравнений для второго игрока",
            reduced_A,
            list(range(m)),
            list(range(n)),
            v_row_p2,
            2,
            original_row_count,
            original_column_count,
        )

        fin_t2, fin_r2, fin_c2 = solve_jordan_simplex(
            table_p2,
            rows_p2,
            cols_p2,
            f"ВТОРОЙ ИГРОК (V выражено из {v_row_p2 + 1}-го уравнения)",
            is_min=True,
            document=tables_document,
        )

        save_tables_to_docx(tables_document)
        print("\nТаблицы сохранены в файл: jordan_tables.docx")

        # Восстанавливаем полные векторы стратегий (с нулями на месте удалённых)
        X_opt = [Fraction(0)] * original_row_count
        for r_lbl, row in zip(fin_r1[:-1], fin_t1[:-1]):
            for r_orig in act_rows:
                if r_lbl == f"x_{r_orig + 1} =":
                    X_opt[r_orig] = row[0]

        Y_opt = [Fraction(0)] * original_column_count
        for r_lbl, row in zip(fin_r2[:-1], fin_t2[:-1]):
            for c_orig in act_cols:
                if r_lbl == f"y_{c_orig + 1} =":
                    Y_opt[c_orig] = row[0]

        V1 = fin_t1[-1][0]
        V2 = fin_t2[-1][0]

        # Восстанавливаем двойственные стратегии с учётом того, что
        # дополнительные переменные могли стать базисными.
        Y_from_p1 = [Fraction(0)] * original_column_count
        y_active = extract_strategy_from_table(
            fin_t1,
            fin_r1,
            fin_c1,
            "x",
            [original_row_count + position + 1 for position in range(n)],
        )
        for position, c_orig in enumerate(act_cols):
            Y_from_p1[c_orig] = y_active[position]

        X_from_p2 = [Fraction(0)] * original_row_count
        x_active = extract_strategy_from_table(
            fin_t2,
            fin_r2,
            fin_c2,
            "y",
            [original_column_count + position + 1 for position in range(m)],
            coefficient_sign=-1,
        )
        for position, r_orig in enumerate(act_rows):
            X_from_p2[r_orig] = x_active[position]
    else:
        V1 = V2 = pure_V
        Y_from_p1 = list(Y_opt)
        X_from_p2 = list(X_opt)

    # =====================================================================
    # 3. ИТОГОВЫЙ ВЫВОД И ПРОВЕРКА СОВПАДЕНИЯ ВЫИГРЫШЕЙ
    # =====================================================================
    print("\n" + "=" * 75)
    print("ИТОГОВЫЕ РЕЗУЛЬТАТЫ И СВЕРКА")
    print("=" * 75)

    print(
        f"Оптимальная стратегия 1-го игрока X* = ({', '.join(str(x) for x in X_opt)})"
    )
    print(f"Сумма вероятностей X*: {sum(X_opt)}")
    print(
        f"Стратегия Y*, полученная из таблицы 1-го игрока = "
        f"({', '.join(str(y) for y in Y_from_p1)})"
    )
    print(f"Сумма вероятностей Y* из 1-й задачи: {sum(Y_from_p1)}")
    print(f"Цена игры из 1-й задачи V1 = {V1} (≈ {float(V1):.4f})\n")

    print(
        f"Оптимальная стратегия 2-го игрока Y* = ({', '.join(str(y) for y in Y_opt)})"
    )
    print(f"Сумма вероятностей Y*: {sum(Y_opt)}")
    print(
        f"Стратегия X*, полученная из таблицы 2-го игрока = "
        f"({', '.join(str(x) for x in X_from_p2)})"
    )
    print(f"Сумма вероятностей X* из 2-й задачи: {sum(X_from_p2)}")
    print(f"Цена игры из 2-й задачи V2 = {V2} (≈ {float(V2):.4f})\n")

    print(f"Выигрыши совпали (V1 == V2): {V1 == V2}")

("Первому игроку выгоднее всего чаще выбирать стратегии (A_1) и (A_4), "
 "а стратегии (A_2) и (A_3) использовать реже. Второму игроку следует в основном выбирать (B_1) и (B_4), "
 "почти не использовать (B_2) и (B_3), а (B_5) не выбирать вообще. "
 "При таких смешанных стратегиях первый игрок в среднем получает около 16,6 единицы выигрыша за игру. "
 "Второй игрок теряет столько же. "
 "Это означает, что найденные стратегии являются взаимовыгодными "
 "и не позволяют противнику улучшить свой результат односторонним изменением стратегии.")