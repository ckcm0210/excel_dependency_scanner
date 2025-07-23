import os
import re
import json
import formulas
import openpyxl
import hashlib
from openpyxl.worksheet.formula import ArrayFormula
import tkinter as tk
from tkinter import scrolledtext
from tkinter import ttk
import pythoncom
import win32com.client

working_path = r"C:\Users\user\Desktop\pytest\Formula Difference Analyzer"

left_scan_task = None
right_scan_task = None

def process_task_recursively(
    task,
    prefix="",
    current_path=None,
    parent_context=None,
    unique_nodes_for_report=None,
    final_dependency_map=None,
    trace_dependency_vine=None,
    working_path=None,
    display_mode="simple"
):
    if current_path is None:
        current_path = set()

    task_identifier = (task["file"], task["sheet"], task["cell"])
    
    if task_identifier in current_path:
        print(f"{prefix}📍 Circular reference to [{os.path.basename(task['file'])}]{task['sheet']}!{task['cell']} detected, stopping expansion.")
        return

    current_path.add(task_identifier)
    
    if unique_nodes_for_report is not None and task_identifier not in unique_nodes_for_report:
        unique_nodes_for_report.add(task_identifier)
        if final_dependency_map is not None:
            final_dependency_map.append(task)
    
    dependencies, is_formula, content, actual_value = trace_dependency_vine(task, working_path)

    if display_mode == "simple":
        if parent_context and task['file'] == parent_context['file']:
            if task['sheet'].lower() == parent_context['sheet'].lower():
                header = task['cell']
            else:
                header = f"{task['sheet']}!{task['cell']}"
        else:
            header = f"[{os.path.basename(task['file'])}]{task['sheet']}!{task['cell']}"
    elif display_mode == "detail":
        header = f"[{os.path.basename(task['file'])}]{task['sheet']}!{task['cell']}"
    elif display_mode == "fullpath":
        header = f"{task['file']}|{task['sheet']}!{task['cell']}"

    if not is_formula and content.startswith('['):
        print(f"{prefix}📍 {header}")
        print(f"{prefix.replace('📍', ' ' * len('📍'))}🔷 Characteristic: {content}")
    elif not is_formula:
        # 非公式儲存格：顯示標題和實際值
        if actual_value is not None:
            if isinstance(actual_value, str):
                value_display = f"'{actual_value}'"
            else:
                value_display = str(actual_value)
            print(f"{prefix}📍 {header}: {value_display}")
        else:
            print(f"{prefix}📍 {header}: {content}")
    else:
        # 公式儲存格：顯示標題、公式和計算結果
        print(f"{prefix}📍 {header}")
        symbol = "⚙️ Formula:"
        print(f"{prefix}{symbol} {content}")
        
        # 添加公式計算結果
        if actual_value is not None:
            if isinstance(actual_value, str):
                result_display = f"'{actual_value}'"
            else:
                result_display = str(actual_value)
            print(f"{prefix}📊 Result: {result_display}")
        else:
            print(f"{prefix}📊 Result: [Unable to calculate]")

    def sort_dependencies_by_formula_order(dependencies, formula):
        if not formula or not isinstance(formula, str) or not dependencies:
            return dependencies
        formula_upper = formula.upper()
        dep_positions = []
        for dep in dependencies:
            dep_cell = dep.get("cell", "")
            dep_sheet = dep.get("sheet", "")
            patterns = [
                re.escape(dep_cell),
                re.escape(f"{dep_sheet}!{dep_cell}"),
                re.escape(f"'{dep_sheet}'!{dep_cell}")
            ]
            min_pos = len(formula_upper)+1
            for pat in patterns:
                m = re.search(pat, formula_upper)
                if m:
                    min_pos = min(min_pos, m.start())
            dep_positions.append((min_pos, dep))
        dep_positions.sort(key=lambda x: x[0])
        return [d for pos, d in dep_positions]

    formula_for_order = None
    if is_formula and isinstance(content, str) and content.startswith("="):
        formula_for_order = content
    elif is_formula and isinstance(content, str):
        formula_for_order = content

    ordered_dependencies = sort_dependencies_by_formula_order(dependencies, formula_for_order)

    for i, dep_task in enumerate(ordered_dependencies):
        is_last = i == len(ordered_dependencies) - 1
        child_prefix = (prefix.replace("├─", "│    ").replace("└─", "     ")) + ("└─ " if is_last else "├─ ")
        process_task_recursively(
            dep_task,
            prefix=child_prefix,
            current_path=current_path.copy(),
            parent_context=task,
            unique_nodes_for_report=unique_nodes_for_report,
            final_dependency_map=final_dependency_map,
            trace_dependency_vine=trace_dependency_vine,
            working_path=working_path,
            display_mode=display_mode
        )

# 全局檔案快取
_file_cache = {}

def get_cached_workbook(file_path, data_only=False, use_resolved=False):
    """獲取快取的 workbook，避免重複載入同一檔案"""
    cache_key = f"{file_path}_{data_only}_{use_resolved}"
    
    if cache_key not in _file_cache:
        try:
            if use_resolved:
                from workbook_resolver import load_resolved_workbook
                _file_cache[cache_key] = load_resolved_workbook(file_path)
            else:
                _file_cache[cache_key] = openpyxl.load_workbook(filename=file_path, data_only=data_only)
        except Exception as e:
            print(f"Warning: Could not load {file_path}: {e}")
            return None
    
    return _file_cache[cache_key]

def clear_file_cache():
    """清理檔案快取"""
    for wb in _file_cache.values():
        try:
            wb.close()
        except:
            pass
    _file_cache.clear()

def trace_dependency_vine(task, working_path):
    target_file_path, target_sheet_name, target_cell_address = task["file"], task["sheet"], task["cell"]
    try:
        # 使用快取的 workbook
        wb_openpyxl = get_cached_workbook(target_file_path, data_only=False)
        if wb_openpyxl is None:
            return [], False, f"❌ Could not load file: {target_file_path}", None
        
        # formulas 庫只在需要時載入一次
        excel_model = formulas.ExcelModel().load(target_file_path)

        actual_sheet_name = next((s for s in wb_openpyxl.sheetnames if s.lower() == target_sheet_name.lower()), None)
        if not actual_sheet_name:
            raise ValueError(f"Worksheet '{target_sheet_name}' does not exist.")

        ws_openpyxl = wb_openpyxl[actual_sheet_name]
        cell_obj = ws_openpyxl[target_cell_address]

        if isinstance(cell_obj, tuple):
            rows = len(cell_obj)
            cols = len(cell_obj[0]) if rows > 0 else 0
            
            dimension_str = f"[{rows}R x {cols}C]"
            summary_str = ""

            total_sum = 0
            numeric_cells_count = 0
            error_cells_count = 0
            text_cells_count = 0
            hash_content_string = ""
            
            for row_of_cells in cell_obj:
                for cell in row_of_cells:
                    value = cell.value
                    if isinstance(value, (int, float)):
                        total_sum += value
                        numeric_cells_count += 1
                    elif isinstance(value, str):
                        if value.startswith('#'):
                            error_cells_count += 1
                        else:
                            text_cells_count += 1
                    if isinstance(value, ArrayFormula):
                        hash_content_string += "ArrayFormula||"
                    else:
                        hash_content_string += ("" if value is None else str(value)) + "||"
            
            sha256_hash = hashlib.sha256(hash_content_string.encode('utf-8')).hexdigest()
            hash_str = f" [Hash: {sha256_hash[:8]}...]"

            if numeric_cells_count > 0:
                summary_str = f" [Sum: {total_sum:,.2f}]".replace('.00', '')
            elif error_cells_count > 0:
                summary_str = f" [Errors: {error_cells_count}]"
            elif text_cells_count > 0:
                summary_str = " [Text]"

            display_content = f"{dimension_str}{summary_str}{hash_str}"
            return [], False, display_content

        cell_content = cell_obj.value
        is_formula = isinstance(cell_content, ArrayFormula) or (isinstance(cell_content, str) and cell_content.startswith('='))
        
        if not is_formula:
            if isinstance(cell_content, str):
                display_content = f"'{cell_content}'"
            else:
                display_content = str(cell_content)
        else:
            display_content = str(cell_content)

        normalized_parts = []

        target_key_lower = f"'[{os.path.basename(target_file_path).lower()}]{actual_sheet_name.lower()}'!{target_cell_address.lower()}"
        found_key = next((k for k in excel_model.cells if k.lower() == target_key_lower), None)
        if not found_key:
            simple_key = f"'{actual_sheet_name}'!{target_cell_address}"
            if simple_key in excel_model.cells:
                found_key = simple_key
        
        compiled_cell_object = excel_model.cells.get(found_key) if found_key else None

        if compiled_cell_object and hasattr(compiled_cell_object, 'inputs') and compiled_cell_object.inputs:
            raw_references = list(compiled_cell_object.inputs.keys())
            ref_pattern = re.compile(r"'(.*)\[(.*?)\](.*?)'!(.*)")
            for ref in raw_references:
                part = {}
                if match := ref_pattern.match(ref):
                    _, filename_part, sheetname_part, cell_address_part = match.groups()
                    absolute_path = os.path.join(working_path, filename_part)
                    part = {"file": absolute_path, "sheet": sheetname_part, "cell": cell_address_part}
                else:
                    sheetname_part, cell_address_part = ref.split('!')
                    part = {"file": target_file_path, "sheet": sheetname_part.strip("'"), "cell": cell_address_part}
                normalized_parts.append(part)

        if is_formula:
            raw_formula = str(cell_content.text) if isinstance(cell_content, ArrayFormula) else str(cell_content)
            reconstructed_formula = raw_formula

            if hasattr(wb_openpyxl, "_external_links") and wb_openpyxl._external_links and compiled_cell_object and hasattr(compiled_cell_object, 'inputs'):
                ref_pattern_for_map = re.compile(r".*\[(.*?)\]")
                external_filenames = sorted(list({ref_pattern_for_map.match(ref).group(1) for ref in compiled_cell_object.inputs if ref_pattern_for_map.match(ref)}))

                index_to_path_map = {
                    i + 1: os.path.join(working_path, filename)
                    for i, filename in enumerate(external_filenames)
                }

                def replacer(match):
                    placeholder_index = int(match.group(1))
                    formula_part = match.group(2)
                    full_path = index_to_path_map.get(placeholder_index)
                    if not full_path:
                        return match.group(0)
                    if '!' in formula_part:
                        sheet_name, cell_ref = formula_part.split('!', 1)
                        return f"'{os.path.dirname(full_path)}\\[{os.path.basename(full_path)}]{sheet_name}'!{cell_ref}"
                    else:
                        return f"'{os.path.dirname(full_path)}\\[{os.path.basename(full_path)}]{formula_part}'"

                reconstructed_formula = re.sub(r'\[(\d+)\]([^\]!]+(?:![\$A-Z0-9:]+)?)(?=[,)\s*+\-\/\^=<>:&]|$)', replacer, raw_formula)

            # 使用快取的 resolved workbook 來獲取解析後的公式顯示
            try:
                wb_resolved = get_cached_workbook(target_file_path, use_resolved=True)
                if wb_resolved:
                    ws_resolved = wb_resolved[actual_sheet_name]
                    cell_resolved = ws_resolved[target_cell_address]
                    resolved_value = cell_resolved.value
                    if isinstance(resolved_value, str):
                        display_content = resolved_value
                    else:
                        display_content = reconstructed_formula
                else:
                    display_content = reconstructed_formula
            except:
                display_content = reconstructed_formula
            
            if "INDIRECT" in raw_formula.upper():
                wb_data_only = None
                try:
                    match = re.search(r'INDIRECT\((.*)\)', raw_formula, re.IGNORECASE)
                    if match:
                        argument_str = match.group(1)
                        literals = re.findall(r'"(.*?)"', argument_str)
                        cell_refs = [ref for ref in re.split(r'"[^"]*"|&', argument_str) if ref]

                        wb_data_only = get_cached_workbook(target_file_path, data_only=True)
                        if wb_data_only:
                            ws_data_only = wb_data_only[actual_sheet_name]
                            evaluated_refs = [str(ws_data_only[cell.strip()].value) for cell in cell_refs]
                        else:
                            evaluated_refs = []
                        
                        final_target_str = ""
                        if len(literals) == 3 and len(evaluated_refs) == 2:
                                final_target_str = literals[0] + literals[1] + evaluated_refs[0] + literals[2] + evaluated_refs[1]
                        
                        if final_target_str:
                            ref_match = re.search(r"'?(.*\\\[(.*?)\])(.*?)'?!([A-Z0-9]+)", final_target_str, re.IGNORECASE)
                            if ref_match:
                                full_path_part, filename, sheet, cell = ref_match.groups()
                                dep_filepath = os.path.join(os.path.dirname(full_path_part), filename)
                                new_task = {"file": dep_filepath, "sheet": sheet, "cell": cell}
                                normalized_parts.insert(0, new_task)
                            else:
                                display_content += f" [Tracer Warning: Could not parse INDIRECT result '{final_target_str}']"
                except Exception as e:
                    display_content += f" [Tracer Warning: Could not resolve INDIRECT -> {e}]"
                finally:
                    if wb_data_only:
                        wb_data_only.close()

        # 獲取實際的儲存格值（用於顯示）
        actual_value = None
        try:
            wb_data_only = get_cached_workbook(target_file_path, data_only=True)
            if wb_data_only:
                ws_data_only = wb_data_only[actual_sheet_name]
                actual_value = ws_data_only[target_cell_address].value
        except:
            actual_value = None
        
        return normalized_parts, is_formula, display_content, actual_value

    except Exception as e:
        return [], False, f"❌ Error during analysis: {e}", None

def get_active_excel_info():
    pythoncom.CoInitialize()
    excel = win32com.client.GetObject(Class="Excel.Application")
    wb = excel.ActiveWorkbook
    ws = excel.ActiveSheet
    cell = excel.ActiveCell
    file_path = wb.FullName
    sheet_name = ws.Name
    cell_address = cell.Address.replace("$", "")
    return file_path, sheet_name, cell_address

def run_scan_and_show(text_widget, display_mode, summary_label_list=None, add_empty_lines=True, task=None, file_path=None, sheet_name=None, cell_address=None):
    if not task:
        return

    unique_nodes_for_report = set()
    final_dependency_map = []
    import io, sys
    buffer = io.StringIO()
    sys_stdout = sys.stdout
    sys.stdout = buffer
    process_task_recursively(
        task,
        unique_nodes_for_report=unique_nodes_for_report,
        final_dependency_map=final_dependency_map,
        trace_dependency_vine=trace_dependency_vine,
        working_path=os.path.dirname(task["file"]),
        display_mode=display_mode
    )
    sys.stdout = sys_stdout
    result = buffer.getvalue()
    
    # 清理檔案快取以釋放記憶體
    clear_file_cache()

    if add_empty_lines:
        spaced_result_lines = []
        for line in result.splitlines():
            if line.strip():
                spaced_result_lines.append(line)
                spaced_result_lines.append("")
        result = "\n".join(spaced_result_lines).strip()

    file_dir = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)
    ws = sheet_name
    cell = cell_address

    if summary_label_list:
        summary_label_list[0].config(text=f"File Path: {file_dir}")
        summary_label_list[1].config(text=f"Location: [{file_name}]'{ws}'!{cell}")
        summary_label_list[2].config(text="")
        summary_label_list[3].config(text="")

    text_widget.delete("1.0", tk.END)
    text_widget.insert(tk.END, result)
# --- 新增/覆蓋 patch：行號顯示要根據 widget 視窗顯示高度自動補到最底，resize 都即時更新 ---
    if text_widget == output_left:
        target_line_number_widget = line_number_left
    elif text_widget == output_right:
        target_line_number_widget = line_number_right
    else:
        target_line_number_widget = None

    def update_line_number_widget():
        if not target_line_number_widget:
            return
        text_widget.update_idletasks()
        target_line_number_widget.config(state="normal")
        target_line_number_widget.delete("1.0", tk.END)
        # 根據顯示高度計算總行數
        total_pixel_height = text_widget.winfo_height()
        try:
            dline = text_widget.dlineinfo("1.0")
            line_height = dline[3] if dline else 18
        except Exception:
            line_height = 18
        num_display_lines = max(int(total_pixel_height / line_height), 1)
        widget_total_lines = int(text_widget.index('end-1c').split('.')[0])
        line_count = max(num_display_lines, widget_total_lines)
        line_number_str = "\n".join(str(i+1) for i in range(line_count))
        target_line_number_widget.insert("1.0", line_number_str)
        target_line_number_widget.config(state="disabled")

    update_line_number_widget()

    def _sync_scroll(*args):
        text_widget.yview(*args)
        target_line_number_widget.yview(*args)

    text_widget.config(yscrollcommand=lambda *args: [target_line_number_widget.yview_moveto(args[0]), None])
    target_line_number_widget.config(yscrollcommand=lambda *args: [text_widget.yview_moveto(args[0]), None])
    text_widget.bind('<MouseWheel>', lambda e: (_sync_scroll('scroll', int(-1*(e.delta/120)), 'units'), 'break'))
    target_line_number_widget.bind('<MouseWheel>', lambda e: (_sync_scroll('scroll', int(-1*(e.delta/120)), 'units'), 'break'))

    def _on_resize(event):
        update_line_number_widget()
    text_widget.bind("<Configure>", _on_resize)
# --- 保持原有 code ---
    error_pattern = r"❌ Error during analysis: .*"
    for match in re.finditer(error_pattern, result):
        start_index = text_widget.search(re.escape(match.group(0)), "1.0", tk.END, regexp=True)
        if start_index:
            end_index = text_widget.index(f"{start_index}+{len(match.group(0))}c")
            text_widget.tag_add("error_highlight", start_index, end_index)

    circular_pattern = r"📍 Circular reference to \[.*?\]\w+!\w+ detected, stopping expansion\."
    for match in re.finditer(circular_pattern, result):
        start_index = text_widget.search(re.escape(match.group(0)), "1.0", tk.END, regexp=True)
        if start_index:
            end_index = text_widget.index(f"{start_index}+{len(match.group(0))}c")
            text_widget.tag_add("circular_ref", start_index, end_index)

    formula_full_line_pattern = r"⚙️ Formula: (=.*)"
    for match in re.finditer(formula_full_line_pattern, result):
        start_of_formula_content_in_line = match.start(1)
        
        full_match_start_tk_index = text_widget.search(re.escape(match.group(0)), "1.0", tk.END, regexp=True)
        if full_match_start_tk_index:
            formula_content_start_tk_index = text_widget.index(f"{full_match_start_tk_index}+{len('⚙️ Formula: ')}c")
            formula_content_end_tk_index = text_widget.index(f"{formula_content_start_tk_index}+{len(match.group(1))}c")
            text_widget.tag_add("formula_display", formula_content_start_tk_index, formula_content_end_tk_index)

            external_ref_in_formula_pattern = r"'(?:.*?\\)?\[.*?\][^']+'![\w$:]+"
            current_search_idx = formula_content_start_tk_index
            while True:
                ext_match_start = text_widget.search(external_ref_in_formula_pattern, current_search_idx, formula_content_end_tk_index, regexp=True)
                if not ext_match_start:
                    break
                ext_match_text = text_widget.get(ext_match_start, formula_content_end_tk_index)
                actual_ext_match = re.match(external_ref_in_formula_pattern, ext_match_text)
                if actual_ext_match:
                    ext_match_len = len(actual_ext_match.group(0))
                    ext_match_end = text_widget.index(f"{ext_match_start}+{ext_match_len}c")
                    text_widget.tag_add("external_ref", ext_match_start, ext_match_end)
                    current_search_idx = ext_match_end
                else:
                    break

    char_line_pattern = r"🔷 Characteristic: (.*)"
    for match in re.finditer(char_line_pattern, result):
        full_line_start_tk_index = text_widget.search(re.escape(match.group(0)), "1.0", tk.END, regexp=True)
        if full_line_start_tk_index:
            char_content_start_tk_index = text_widget.index(f"{full_line_start_tk_index}+{len('🔷 Characteristic: ')}c")
            char_content_end_tk_index = text_widget.index(f"{char_content_start_tk_index}+{len(match.group(1))}c")
            text_widget.tag_add("characteristic_info", char_content_start_tk_index, char_content_end_tk_index)

            char_text = match.group(1)
            offset_from_line_start = char_content_start_tk_index

            dimension_match = re.search(r"\[\d+R x \d+C\]", char_text)
            if dimension_match:
                dim_start_char = char_text.find(dimension_match.group(0))
                dim_end_char = dim_start_char + len(dimension_match.group(0))
                tk_dim_start = text_widget.index(f"{offset_from_line_start}+{dim_start_char}c")
                tk_dim_end = text_widget.index(f"{offset_from_line_start}+{dim_end_char}c")
                text_widget.tag_add("header_info", tk_dim_start, tk_dim_end)

            sum_match = re.search(r"\[Sum: [\d,.]+(?:\.\d+)?\]", char_text)
            if sum_match:
                sum_start_char = char_text.find(sum_match.group(0))
                sum_end_char = sum_start_char + len(sum_match.group(0))
                tk_sum_start = text_widget.index(f"{offset_from_line_start}+{sum_start_char}c")
                tk_sum_end = text_widget.index(f"{offset_from_line_start}+{sum_end_char}c")
                text_widget.tag_add("sum_info", tk_sum_start, tk_sum_end)

            hash_match = re.search(r"\[Hash: [0-9a-fA-F]{8}\.\.\.\]", char_text)
            if hash_match:
                hash_start_char = char_text.find(hash_match.group(0))
                hash_end_char = hash_start_char + len(hash_match.group(0))
                tk_hash_start = text_widget.index(f"{offset_from_line_start}+{hash_start_char}c")
                tk_hash_end = text_widget.index(f"{offset_from_line_start}+{hash_end_char}c")
                text_widget.tag_add("hash_info", tk_hash_start, tk_hash_end)
            
            text_error_match = re.search(r"\[Text\]|\[Errors: \d+\]", char_text)
# ... Original code ...
            if text_error_match:
                text_error_start_char = char_text.find(text_error_match.group(0))
                text_error_end_char = text_error_start_char + len(text_error_match.group(0))
                tk_text_error_start = text_widget.index(f"{offset_from_line_start}+{text_error_start_char}c")
                tk_text_error_end = text_widget.index(f"{offset_from_line_start}+{text_error_end_char}c")
                text_widget.tag_add("characteristic_info", tk_text_error_start, tk_text_error_end)

    lines_in_widget = result.splitlines()
    for i, line_content in enumerate(lines_in_widget):
        line_start_tk_index = f"{i+1}.0"

        if "📍" in line_content and \
           not line_content.strip().startswith(("⚙️ Formula:", "❌ Error during analysis:", "📍 Circular reference:", "🔷 Characteristic:")):
            
            parts = line_content.split('📍', 1)
            if len(parts) == 2:
                content_after_marker = parts[1]
                header_part = ""

                if ': ' in content_after_marker:
                    header_part = content_after_marker.split(': ', 1)[0].strip()
                else:
                    header_part = content_after_marker.strip()
                
                if header_part:
                    header_start_in_line = line_content.find(header_part, line_content.find('📍'))
                    
                    if header_start_in_line != -1:
                        header_end_in_line = header_start_in_line + len(header_part)
                        
                        tk_start_pos = text_widget.index(f"{line_start_tk_index}+{header_start_in_line}c")
                        tk_end_pos = text_widget.index(f"{line_start_tk_index}+{header_end_in_line}c")
                        
                        text_widget.tag_add("header_info", tk_start_pos, tk_end_pos)

    literal_value_pattern = r"'(?:[^']|'')*?'"
    current_index = "1.0"
    while True:
# ... Subsequent code ...
        start_pos = text_widget.search(literal_value_pattern, current_index, tk.END, regexp=True)
        if not start_pos:
            break
        
        matched_text_line_end = text_widget.get(start_pos, f"{start_pos} lineend")
        actual_match = re.match(literal_value_pattern, matched_text_line_end)
        
        if actual_match:
            end_pos = text_widget.index(f"{start_pos}+{len(actual_match.group(0))}c")
            
            line_content_at_start = text_widget.get(f"{start_pos} linestart", f"{start_pos} lineend")
            
            is_external_ref = re.search(r"'(?:.*?\\)?\[.*?\][^']+'![\w$:]+", line_content_at_start)
            is_formula_line = line_content_at_start.strip().startswith("⚙️ Formula:")
            is_error_line = "Error during analysis" in line_content_at_start
            is_circular_line = "Circular reference" in line_content_at_start
            is_characteristic_line = line_content_at_start.strip().startswith("🔷 Characteristic:")

            if not is_external_ref and not is_error_line and not is_circular_line and not is_characteristic_line:
                if is_formula_line:
                    text_widget.tag_add("literal_value", start_pos, end_pos)
                else:
                    if re.search(r":\s*$", text_widget.get(f"{start_pos} wordstart", start_pos)):
                        text_widget.tag_add("literal_value", start_pos, end_pos)
                    elif line_content_at_start.strip().endswith(actual_match.group(0)):
                         text_widget.tag_add("literal_value", start_pos, end_pos)

            current_index = end_pos
        else:
            current_index = text_widget.index(f"{start_pos}+1c")


def do_left_scan():
    global left_scan_task
    file_path, sheet_name, cell_address = get_active_excel_info()
    left_scan_task = {"file": file_path, "sheet": sheet_name, "cell": cell_address}
    refresh_left_result(file_path, sheet_name, cell_address)

def do_right_scan():
    global right_scan_task
    file_path, sheet_name, cell_address = get_active_excel_info()
    right_scan_task = {"file": file_path, "sheet": sheet_name, "cell": cell_address}
    refresh_right_result(file_path, sheet_name, cell_address)

def refresh_left_result(file_path, sheet_name, cell_address):
    if left_scan_task:
        run_scan_and_show(output_left, display_mode_left_var.get(), summary_left_labels, add_empty_lines_left_var.get(), left_scan_task, file_path, sheet_name, cell_address)

def refresh_right_result(file_path, sheet_name, cell_address):
    if right_scan_task:
        run_scan_and_show(output_right, display_mode_right_var.get(), summary_right_labels, add_empty_lines_right_var.get(), right_scan_task, file_path, sheet_name, cell_address)

root = tk.Tk()
root.title("Excel Dependency Scanner")
root.geometry("1600x900")

frame = tk.Frame(root)
frame.pack(fill="both", expand=True)

font_size_left_var = tk.IntVar(value=10)
font_style_left_var = tk.StringVar(value="Consolas")
def update_font_config_left():
    new_size = font_size_left_var.get()
    new_style = font_style_left_var.get()
    output_left.config(font=(new_style, new_size))
    output_left.tag_configure("error_highlight", font=(new_style, new_size, "bold"))
    output_left.tag_configure("circular_ref", font=(new_style, new_size, "italic"))
    output_left.tag_configure("header_info", font=(new_style, new_size, "bold"))
    output_left.tag_configure("literal_value", font=(new_style, new_size, "italic"))


font_size_right_var = tk.IntVar(value=10)
font_style_right_var = tk.StringVar(value="Consolas")
def update_font_config_right():
    new_size = font_size_right_var.get()
    new_style = font_style_right_var.get()
    output_right.config(font=(new_style, new_size))
    output_right.tag_configure("error_highlight", font=(new_style, new_size, "bold"))
    output_right.tag_configure("circular_ref", font=(new_style, new_size, "italic"))
    output_right.tag_configure("header_info", font=(new_style, new_size, "bold"))
    output_right.tag_configure("literal_value", font=(new_style, new_size, "italic"))


display_mode_left_var = tk.StringVar(value="simple")
display_mode_right_var = tk.StringVar(value="simple")

add_empty_lines_left_var = tk.BooleanVar(value=False)
add_empty_lines_right_var = tk.BooleanVar(value=False)

main_pane = ttk.PanedWindow(frame, orient=tk.HORIZONTAL)
main_pane.pack(fill="both", expand=True)

left_frame = tk.Frame(main_pane)
right_frame = tk.Frame(main_pane)

main_pane.add(left_frame, weight=1)
main_pane.add(right_frame, weight=1)

mode_left_frame = tk.Frame(left_frame)
mode_left_frame.pack(pady=5, anchor="w")
scan_btn_left = tk.Button(mode_left_frame, text="Scan Left", width=12, height=1, font=("Arial", 10, "bold"))
scan_btn_left.pack(side="left", padx=2)
tk.Label(mode_left_frame, text="Display Mode (Left):").pack(side="left")
tk.Radiobutton(mode_left_frame, text="Simple", variable=display_mode_left_var, value="simple", command=lambda: refresh_left_result(left_scan_task["file"], left_scan_task["sheet"], left_scan_task["cell"])).pack(side="left")
tk.Radiobutton(mode_left_frame, text="Detail", variable=display_mode_left_var, value="detail", command=lambda: refresh_left_result(left_scan_task["file"], left_scan_task["sheet"], left_scan_task["cell"])).pack(side="left")
tk.Radiobutton(mode_left_frame, text="Full Path", variable=display_mode_left_var, value="fullpath", command=lambda: refresh_left_result(left_scan_task["file"], left_scan_task["sheet"], left_scan_task["cell"])).pack(side="left")
tk.Checkbutton(mode_left_frame, text="Add Empty Lines", variable=add_empty_lines_left_var, command=lambda: refresh_left_result(left_scan_task["file"], left_scan_task["sheet"], left_scan_task["cell"])).pack(side="left", padx=5)

font_control_left_frame = tk.Frame(left_frame)
font_control_left_frame.pack(pady=2, anchor="w")
tk.Label(font_control_left_frame, text="Font Size:").pack(side="left")
tk.Spinbox(font_control_left_frame, from_=6, to=28, width=4, textvariable=font_size_left_var, command=lambda: update_font_config_left()).pack(side="left")
tk.Label(font_control_left_frame, text="Font Style:").pack(side="left")
font_style_options = ["Consolas", "Courier New", "Menlo", "Liberation Mono", "DejaVu Sans Mono"]
tk.OptionMenu(font_control_left_frame, font_style_left_var, *font_style_options, command=lambda _: update_font_config_left()).pack(side="left")


summary_left_frame = tk.Frame(left_frame)
summary_left_frame.pack(pady=2, anchor="w")
summary_left_labels = [tk.Label(summary_left_frame, text="", anchor="w", font=("Arial", 10, "bold")) for _ in range(4)]
for lab in summary_left_labels:
    lab.pack(anchor="w")

mode_right_frame = tk.Frame(right_frame)
mode_right_frame.pack(pady=5, anchor="w")
scan_btn_right = tk.Button(mode_right_frame, text="Scan Right", width=12, height=1, font=("Arial", 10, "bold"))
scan_btn_right.pack(side="left", padx=2)
tk.Label(mode_right_frame, text="Display Mode (Right):").pack(side="left")
tk.Radiobutton(mode_right_frame, text="Simple", variable=display_mode_right_var, value="simple", command=lambda: refresh_right_result(right_scan_task["file"], right_scan_task["sheet"], right_scan_task["cell"])).pack(side="left")
tk.Radiobutton(mode_right_frame, text="Detail", variable=display_mode_right_var, value="detail", command=lambda: refresh_right_result(right_scan_task["file"], right_scan_task["sheet"], right_scan_task["cell"])).pack(side="left")
tk.Radiobutton(mode_right_frame, text="Full Path", variable=display_mode_right_var, value="fullpath", command=lambda: refresh_right_result(right_scan_task["file"], right_scan_task["sheet"], right_scan_task["cell"])).pack(side="left")
tk.Checkbutton(mode_right_frame, text="Add Empty Lines", variable=add_empty_lines_right_var, command=lambda: refresh_right_result(right_scan_task["file"], right_scan_task["sheet"], right_scan_task["cell"])).pack(side="left", padx=5)

font_control_right_frame = tk.Frame(right_frame)
font_control_right_frame.pack(pady=2, anchor="w")
tk.Label(font_control_right_frame, text="Font Size:").pack(side="left")
tk.Spinbox(font_control_right_frame, from_=6, to=28, width=4, textvariable=font_size_right_var, command=lambda: update_font_config_right()).pack(side="left")
tk.Label(font_control_right_frame, text="Font Style:").pack(side="left")
tk.OptionMenu(font_control_right_frame, font_style_right_var, *font_style_options, command=lambda _: update_font_config_right()).pack(side="left")


summary_right_frame = tk.Frame(right_frame)
summary_right_frame.pack(pady=2, anchor="w")
summary_right_labels = [tk.Label(summary_right_frame, text="", anchor="w", font=("Arial", 10, "bold")) for _ in range(4)]
for lab in summary_right_labels:
    lab.pack(anchor="w")

line_number_left_frame = tk.Frame(left_frame)
line_number_left_frame.pack(side="left", fill="y")

line_number_left = tk.Text(line_number_left_frame, width=4, font=(font_style_left_var.get(), font_size_left_var.get()), state="disabled", bg="#f0f0f0", fg="gray")
line_number_left.pack(fill="y", expand=False, side="left")

output_left = scrolledtext.ScrolledText(left_frame, width=80, font=(font_style_left_var.get(), font_size_left_var.get()))
output_left.pack(side="left", expand=True, fill="both")
output_left.tag_configure("formula_display", foreground="blue")
output_left.tag_configure("external_ref", foreground="darkgreen")
output_left.tag_configure("error_highlight", foreground="red", font=(font_style_left_var.get(), font_size_left_var.get(), "bold"))
output_left.tag_configure("circular_ref", foreground="purple", font=(font_style_left_var.get(), font_size_left_var.get(), "italic"))
output_left.tag_configure("characteristic_info", foreground="magenta")
output_left.tag_configure("header_info", font=(font_style_left_var.get(), font_size_left_var.get(), "bold"))
output_left.tag_configure("literal_value", foreground="darkblue", font=(font_style_left_var.get(), font_size_left_var.get(), "italic"))
output_left.tag_configure("sum_info", foreground="orange")
output_left.tag_configure("hash_info", foreground="gray")


line_number_right_frame = tk.Frame(right_frame)
line_number_right_frame.pack(side="left", fill="y")

line_number_right = tk.Text(line_number_right_frame, width=4, font=(font_style_right_var.get(), font_size_right_var.get()), state="disabled", bg="#f0f0f0", fg="gray")
line_number_right.pack(fill="y", expand=False, side="left")

output_right = scrolledtext.ScrolledText(right_frame, width=80, font=(font_style_right_var.get(), font_size_right_var.get()))
output_right.pack(side="left", expand=True, fill="both")
output_right.tag_configure("formula_display", foreground="blue")
output_right.tag_configure("external_ref", foreground="darkgreen")
output_right.tag_configure("error_highlight", foreground="red", font=(font_style_right_var.get(), font_size_right_var.get(), "bold"))
output_right.tag_configure("circular_ref", foreground="purple", font=(font_style_right_var.get(), font_size_right_var.get(), "italic"))
output_right.tag_configure("characteristic_info", foreground="magenta")
output_right.tag_configure("header_info", font=(font_style_right_var.get(), font_size_right_var.get(), "bold"))
output_right.tag_configure("literal_value", foreground="darkblue", font=(font_style_right_var.get(), font_size_right_var.get(), "italic"))
output_right.tag_configure("sum_info", foreground="orange")
output_right.tag_configure("hash_info", foreground="gray")


scan_btn_left.config(command=do_left_scan)
scan_btn_right.config(command=do_right_scan)

update_font_config_left()
update_font_config_right()

root.mainloop()