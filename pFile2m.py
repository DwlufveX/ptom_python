import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os

PTOM_EXE = "ptom.exe"  # 确保该 exe 在同一目录，或写绝对路径

def select_and_decrypt():
    files = filedialog.askopenfilenames(title="选择 P 文件", filetypes=[("P Files", "*.p")])
    if not files:
        return

    for pfile in files:
        abs_pfile = os.path.abspath(pfile)
        output_mfile = abs_pfile.rsplit('.', 1)[0] + '.m'

        print(f"输入文件: {abs_pfile}")
        print(f"输出文件: {output_mfile}")

        try:
            result = subprocess.run(
                [PTOM_EXE, abs_pfile, output_mfile],
                capture_output=True,
                text=True,
                encoding='gbk'
            )

            print(result.stdout)
            if result.stderr.strip():
                print("错误信息:", result.stderr)

        except Exception as e:
            messagebox.showerror("错误", f"处理失败：\n{e}")
            return

    messagebox.showinfo("完成", "全部 P 文件已转换为 M 文件")

def main():
    root = tk.Tk()
    root.title("P 文件批量解密工具")
    root.geometry("320x180")

    label = tk.Label(root, text="点击选择 .p 文件进行解密")
    label.pack(pady=20)

    button = tk.Button(root, text="选择 P 文件", command=select_and_decrypt)
    button.pack()

    root.mainloop()

if __name__ == "__main__":
    main()
