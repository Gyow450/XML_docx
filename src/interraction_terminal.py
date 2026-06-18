"""互动窗口，得到参数设置，包括文件夹，文件名，其他控制参数等
类型码：0-文件夹，1-保存文件，2-打开文件，3-布尔，4-字符串"""
import json
from pathlib import Path
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox


# ---------------------------------------------------------------------------
# 本地配置持久化辅助函数
# ---------------------------------------------------------------------------
def _load_local_setting(path: Path = Path('local_setting.json')) -> dict:
    """读取本地配置，文件不存在或读取失败时返回空字典。"""
    if not path.exists():
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        # 配置损坏时自动备份，避免丢失用户数据
        backup = path.with_suffix('.json.bak')
        try:
            if backup.exists():
                backup.unlink()
            path.rename(backup)
        except OSError:
            backup = None
        msg = f'本地配置读取失败：{e}\n已使用默认配置启动。'
        if backup:
            msg += f'\n原文件已备份为：{backup}'
        messagebox.showwarning('配置读取失败', msg)
        return {}
    except OSError as e:
        messagebox.showwarning('配置读取失败', f'无法读取本地配置：{e}')
        return {}


def _save_local_setting(script_name: str, config: dict,
                        path: Path = Path('local_setting.json')) -> None:
    """保存配置，先读取再合并，避免覆盖其他脚本的配置。"""
    data = _load_local_setting(path)
    data[script_name] = config
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except OSError as e:
        messagebox.showwarning('配置保存失败', f'无法保存本地配置：{e}')


def _merge_with_defaults(setting_dict: dict[str, list],
                         saved: dict | None) -> list[tuple]:
    """
    以当前 setting_dict 为模板，用已保存的值覆盖默认值。
    这样可以保证：新增参数能出现、删除参数不会残留、类型码/扩展名始终跟随代码。
    """
    saved = saved or {}
    merged = []
    for key, value in setting_dict.items():
        type_code = value[0]
        ext_or_pattern = value[1] if len(value) > 1 else ''
        default = value[2] if len(value) > 2 else ''

        if (key in saved
                and isinstance(saved[key], (list, tuple))
                and len(saved[key]) >= 3):
            merged.append((type_code, key, ext_or_pattern, saved[key][2]))
        else:
            merged.append((type_code, key, ext_or_pattern, default))
    return merged


def _validate_string(entry: tk.Entry, pattern: str) -> bool:
    """使用传入的正则表达式验证输入内容，并给出视觉反馈。"""
    value = entry.get()
    if re.fullmatch(pattern, value):
        entry.config(bg='white')
        return True
    else:
        entry.config(bg='#ffdddd')
        return False


# ---------------------------------------------------------------------------
# 参数输入界面
# ---------------------------------------------------------------------------
def set_argumments(anything: list[tuple]) -> dict[str, str | bool] | None:
    """
    按照（特征码，关键字，扩展名|正则表达式，初始值）构成元组
    类型码：0-文件夹，1-保存文件，2-打开文件，3-布尔，4-字符串
    返回 None 表示用户取消。
    """
    def _guess_initialdir(key_word: str) -> str | None:
        """根据当前值猜测文件对话框的初始目录。"""
        current = any_var.get(key_word, tk.StringVar()).get()
        if not current:
            return None
        p = Path(current)
        if p.is_dir():
            return str(p)
        if p.parent.exists():
            return str(p.parent)
        return None

    def set_dir_value(key_word: str):
        selected = filedialog.askdirectory(
            title=f'选择{key_word}文件夹',
            initialdir=_guess_initialdir(key_word))
        if selected:
            any_var[key_word].set(selected)

    def set_savefile_value(key_word: str, ex_info: str):
        ex_names = [e.strip() for e in ex_info.split(',') if e.strip()]
        selected = filedialog.asksaveasfilename(
            title=f'选择{key_word}文件',
            defaultextension=f'.{ex_names[0]}' if ex_names else None,
            filetypes=[(f'{ex_name}文件', f'.{ex_name}') for ex_name in ex_names],
            initialdir=_guess_initialdir(key_word))
        if selected:
            any_var[key_word].set(selected)

    def set_openfile_value(key_word: str, ex_info: str):
        ex_names = [e.strip() for e in ex_info.split(',') if e.strip()]
        selected = filedialog.askopenfilename(
            title=f'选择{key_word}文件',
            defaultextension=f'.{ex_names[0]}' if ex_names else None,
            filetypes=[(f'{ex_name}文件', f'.{ex_name}') for ex_name in ex_names],
            initialdir=_guess_initialdir(key_word))
        if selected:
            any_var[key_word].set(selected)

    def validate_path(key_word: str, type_num: int, ex_info: str) -> bool:
        """校验文件/文件夹路径是否存在、扩展名是否匹配。"""
        value = any_var[key_word].get().strip()
        if not value:
            return True  # 空值由最终校验统一处理
        p = Path(value)
        if type_num == 0:
            return p.is_dir()
        if type_num in (1, 2):
            if not p.is_file():
                return False
            if ex_info:
                exts = [f'.{e.strip().lstrip(".").lower()}'
                        for e in ex_info.split(',') if e.strip()]
                if exts and not any(str(p).lower().endswith(ext) for ext in exts):
                    return False
        return True

    def update_path_style(key_word: str, type_num: int, ex_info: str):
        entry = entries.get(key_word)
        if entry is None:
            return
        if validate_path(key_word, type_num, ex_info):
            entry.config(bg='white')
        else:
            entry.config(bg='#ffdddd')

    def on_ok():
        # 最终校验：必填项不能为空、路径必须有效、字符串必须满足正则
        errors = []
        for temp_tuple in anything:
            type_num = temp_tuple[0]
            key_word = temp_tuple[1]
            ex_info = temp_tuple[2] if len(temp_tuple) > 2 else ''
            value = any_var[key_word].get().strip()

            if type_num in (0, 1, 2):
                if not value:
                    errors.append(f'请设置【{key_word}】')
                    continue
                if not validate_path(key_word, type_num, ex_info):
                    errors.append(f'【{key_word}】路径无效或扩展名不匹配：{value}')

            elif type_num == 4 and ex_info and value:
                if not re.fullmatch(ex_info, value):
                    errors.append(f'【{key_word}】格式不符合要求：{value}')

        if errors:
            messagebox.showerror('参数校验失败', '\n'.join(errors))
            return

        for name in final_dict.keys():
            final_dict[name] = any_var[name].get()
        root.destroy()

    def on_cancel():
        nonlocal cancelled
        cancelled = True
        root.destroy()

    final_dict: dict[str, str | bool] = {}
    any_var: dict[str, tk.StringVar | tk.BooleanVar] = {}
    entries: dict[str, tk.Entry] = {}
    cancelled = False

    root = tk.Tk()
    root.title('选择运行参数')

    i = -1  # row
    j = 0   # column
    for temp_tuple in anything:
        i += 1
        type_num = temp_tuple[0]
        key_word = temp_tuple[1]
        ex_info = temp_tuple[2] if len(temp_tuple) > 2 else ''
        variable_value = temp_tuple[3] if len(temp_tuple) > 3 else (
            False if isinstance(ex_info, bool) else '')

        if isinstance(ex_info, bool):
            final_dict[key_word] = variable_value
            any_var[key_word] = tk.BooleanVar(value=final_dict[key_word])
        else:
            final_dict[key_word] = variable_value
            any_var[key_word] = tk.StringVar(value=final_dict[key_word])

        if type_num == 0:
            tk.Label(root, text=f'选择{key_word}所在文件夹').grid(
                row=i, column=j, sticky='e')
            entry = tk.Entry(root, textvariable=any_var[key_word],
                             width=80, state='normal')
            entry.grid(row=i, column=j + 1)
            entries[key_word] = entry
            any_var[key_word].trace_add(
                'write',
                lambda *_, kw=key_word, tn=type_num, ei=ex_info:
                    update_path_style(kw, tn, ei))
            tk.Button(root, text='选择文件夹',
                      command=lambda kw=key_word: set_dir_value(kw)).grid(
                row=i, column=j + 2, padx=2)
        elif type_num == 1:
            tk.Label(root, text=f'选择{key_word}{ex_info}文件').grid(
                row=i, column=j, sticky='e')
            entry = tk.Entry(root, textvariable=any_var[key_word],
                             width=80, state='normal')
            entry.grid(row=i, column=j + 1)
            entries[key_word] = entry
            any_var[key_word].trace_add(
                'write',
                lambda *_, kw=key_word, tn=type_num, ei=ex_info:
                    update_path_style(kw, tn, ei))
            tk.Button(root, text='选择文件',
                      command=lambda kw=key_word, ei=ex_info:
                          set_savefile_value(kw, ei)).grid(
                row=i, column=j + 2, padx=2)
        elif type_num == 2:
            tk.Label(root, text=f'选择{key_word}{ex_info}文件').grid(
                row=i, column=j, sticky='e')
            entry = tk.Entry(root, textvariable=any_var[key_word],
                             width=80, state='normal')
            entry.grid(row=i, column=j + 1)
            entries[key_word] = entry
            any_var[key_word].trace_add(
                'write',
                lambda *_, kw=key_word, tn=type_num, ei=ex_info:
                    update_path_style(kw, tn, ei))
            tk.Button(root, text='选择文件',
                      command=lambda kw=key_word, ei=ex_info:
                          set_openfile_value(kw, ei)).grid(
                row=i, column=j + 2, padx=2)
        elif type_num == 3:
            tk.Checkbutton(root, text=key_word,
                           variable=any_var[key_word]).grid(
                row=i, column=j, sticky='w')
        else:
            tk.Label(root, text=f'输入{key_word}').grid(
                row=i, column=j, sticky='e')
            entry = tk.Entry(root, textvariable=any_var[key_word],
                             width=90, state='normal')
            entry.grid(row=i, column=j + 1)
            entries[key_word] = entry
            if ex_info:
                any_var[key_word].trace_add(
                    'write',
                    lambda *_, ent=entry, p=ex_info:
                        _validate_string(ent, p))

    i += 1
    tk.Button(root, text='确定', command=on_ok).grid(
        row=i, column=j, pady=5)
    tk.Button(root, text='取消', command=on_cancel).grid(
        row=i, column=j + 1, pady=5)

    # 窗口初始化后校验一次，让历史无效路径立即标红
    for temp_tuple in anything:
        type_num = temp_tuple[0]
        key_word = temp_tuple[1]
        ex_info = temp_tuple[2] if len(temp_tuple) > 2 else ''
        if type_num in (0, 1, 2) and key_word in entries:
            update_path_style(key_word, type_num, ex_info)

    root.mainloop()

    if cancelled:
        return None
    return final_dict


def set_local_setting(script_name: str,
                      setting_dict: dict[str, list]) -> dict[str, str | bool]:
    """
    设置本地参数，优先读取之前保存的参数。setting_dict 格式：
    {参数名: [类型码, 扩展名或正则表达式, 初始值]}
    类型码：0-文件夹，1-保存文件，2-打开文件，3-布尔，4-字符串
    """
    data = _load_local_setting()
    saved = data.get(script_name)

    merged = _merge_with_defaults(setting_dict, saved)
    f_config = set_argumments(merged)

    if f_config is None:
        # 保持原有行为：取消即退出
        sys.exit(0)

    # 保存时保持与旧格式一致：[类型码, 扩展名, 值]
    any_list = {
        key: [setting_dict[key][0],
              setting_dict[key][1] if len(setting_dict[key]) > 1 else '',
              value]
        for key, value in f_config.items()
    }
    _save_local_setting(script_name, any_list)

    return f_config


if __name__ == "__main__":
    a_list = [
        (0, '数据源', ''),
        (1, '输出', 'pdf,docx'),
        (3, '是否写入概述', False),
        (4, '数字参数', r'^(-?\d+(~(-?\d+))?(,|$))*$'),
    ]
    print(set_argumments(a_list))
