import os
import re
from datetime import datetime
from tkinter import Tk, filedialog

# 📅 日付取得
export_date = datetime.today().strftime("%Y-%m-%d")

def choose_base_folder():
    root = Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="VBAの親フォルダを選んでください")
    root.destroy()
    return folder

# 🎯 サブフォルダと拡張子
target_folders = [
    "com_clsAll", "acc_clsAll", "xl_clsAll",
    "com_modAll", "acc_modAll", "xl_modAll"
]
valid_exts = (".bas", ".cls")

# 🔁 共通パターン
func_pattern = re.compile(r"^\s*(Public|Private)?\s*(Function|Sub)\s+([A-Za-z0-9_]+)", re.IGNORECASE)
desc_pattern = re.compile(r"^\s*['‘’]?\s*(説明|機能)\s*[:：]?\s*(.+)", re.IGNORECASE)

# ① 📦 merge_modules機能

def merge_modules(base_folder):
    print("▶ モジュール統合を開始...")
    output_dir = os.path.join(base_folder, "txt_output")
    os.makedirs(output_dir, exist_ok=True)

    for folder in target_folders:
        folder_path = os.path.join(base_folder, folder)
        if not os.path.isdir(folder_path):
            continue

        files_data = []
        for filename in sorted(os.listdir(folder_path)):
            if not filename.endswith(valid_exts):
                continue
            path = os.path.join(folder_path, filename)
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except UnicodeDecodeError:
                with open(path, 'r', encoding='cp932') as f:
                    content = f.read()

            key = ''
            for line in content.splitlines():
                m = func_pattern.match(line)
                if m:
                    key = m.group(3)
                    break
            files_data.append((filename, key, content))

        files_data.sort(key=lambda x: (x[0].lower(), x[1]))

        output_filename = f"{folder}.txt"
        output_path = os.path.join(output_dir, output_filename)
        with open(output_path, 'w', encoding='utf-8') as out:
            # ヘッダにエクスポート日を追加
            out.write(f"'=== [ExportDate] {export_date} ===\n")
            for fn, _, content in files_data:
                out.write(f"'=== [File] {fn} ===\n")
                out.write(content)
                out.write("\n\n")
        print(f"✅ 出力完了：{output_filename}")

# ② 🧠 関数一覧出力

def generate_function_report(base_folder):
    print("▶ 関数一覧出力を開始...")
    output_dir = os.path.join(base_folder, "txt_output")
    os.makedirs(output_dir, exist_ok=True)

    output_filename = "function_list.txt"
    output_path = os.path.join(output_dir, output_filename)

    with open(output_path, 'w', encoding='utf-8') as out:
        out.write(f"'=== [ExportDate] {export_date} ===\n")
        out.write("所属フォルダ / 所属ファイル - ファイル種別 - 定義種別 関数名 : 説明\n")

        for folder in target_folders:
            folder_path = os.path.join(base_folder, folder)
            if not os.path.isdir(folder_path):
                continue

            for fname in sorted(os.listdir(folder_path)):
                if not fname.endswith(valid_exts):
                    continue

                path = os.path.join(folder_path, fname)
                try:
                    with open(path, 'r', encoding='utf-8') as f:
                        lines = f.readlines()
                except UnicodeDecodeError:
                    with open(path, 'r', encoding='cp932') as f:
                        lines = f.readlines()

                kind = 'クラス' if fname.endswith('.cls') else '標準モジュール'
                prev = []
                for line in lines:
                    m = func_pattern.match(line)
                    if m:
                        _, decl, name = m.groups()
                        desc = ''
                        for p in reversed(prev[-10:]):
                            d = desc_pattern.search(p)
                            if d:
                                desc = d.group(2).strip()
                                break
                        if not desc:
                            desc = '（なし）'
                        out.write(f"[{folder}] {fname} - {kind} - {decl} {name} : {desc}\n")
                    prev.append(line)

    print(f"✅ 一覧出力完了：{output_filename}")

# 🧭 メイン

def main():
    base = choose_base_folder()
    if not base:
        print("キャンセルされました")
        return

    print("\n🔧 実行する処理を選んでください：")
    print("1. モジュール統合（merge_modules）")
    print("2. 関数一覧出力（generate_function_report）")
    print("3. 両方実行")
    ch = input("番号を入力（1/2/3）: ")

    if ch == '1':
        merge_modules(base)
    elif ch == '2':
        generate_function_report(base)
    elif ch == '3':
        merge_modules(base)
        generate_function_report(base)
    else:
        print("⚠ 無効な入力でした")

if __name__ == '__main__':
    main()
