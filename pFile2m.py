import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import os
import sys
import re
import threading
PTOM_EXE = "ptom.exe"  # 确保该 exe 在同一目录，或写绝对路径

class Formatter:
    # control sequences
    ctrl_1line = re.compile(r'(\s*)(if|while|for|try)(\W\s*\S.*\W)((end|endif|endwhile|endfor);?)(\s+\S.*|\s*$)')
    fcnstart = re.compile(r'(\s*)(function|classdef)\s*(\W\s*\S.*|\s*$)')
    ctrlstart = re.compile(r'(\s*)(if|while|for|parfor|try|methods|properties|events|arguments|enumeration|spmd)\s*(\W\s*\S.*|\s*$)')
    ctrl_ignore = re.compile(r'(\s*)(import|clear|clearvars)(.*$)')
    ctrlstart_2 = re.compile(r'(\s*)(switch)\s*(\W\s*\S.*|\s*$)')
    ctrlcont = re.compile(r'(\s*)(elseif|else|case|otherwise|catch)\s*(\W\s*\S.*|\s*$)')
    ctrlend = re.compile(r'(\s*)((end|endfunction|endif|endwhile|endfor|endswitch);?)(\s+\S.*|\s*$)')
    linecomment = re.compile(r'(\s*)%.*$')
    ellipsis = re.compile(r'.*\.\.\..*$')
    blockcomment_open = re.compile(r'(\s*)%\{\s*$')
    blockcomment_close = re.compile(r'(\s*)%\}\s*$')
    block_close = re.compile(r'\s*[\)\]\}].*$')
    ignore_command = re.compile(r'.*formatter\s+ignore\s+(\d*).*$')

    # patterns
    p_string = re.compile(r'(.*?[\(\[\{,;=\+\-\*\/\|\&\s]|^)\s*(\'([^\']|\'\')+\')([\)\}\]\+\-\*\/=\|\&,;].*|\s+.*|$)')
    p_string_dq = re.compile(r'(.*?[\(\[\{,;=\+\-\*\/\|\&\s]|^)\s*(\"([^\"])*\")([\)\}\]\+\-\*\/=\|\&,;].*|\s+.*|$)')
    p_comment = re.compile(r'(.*\S|^)\s*(%.*)')
    p_blank = re.compile(r'^\s+$')
    p_num_sc = re.compile(r'(.*?\W|^)\s*(\d+\.?\d*)([eE][+-]?)(\d+)(.*)')
    p_num_R = re.compile(r'(.*?\W|^)\s*(\d+)\s*(\/)\s*(\d+)(.*)')
    p_incr = re.compile(r'(.*?\S|^)\s*(\+|\-)\s*(\+|\-)\s*([\)\]\},;].*|$)')
    p_sign = re.compile(r'(.*?[\(\[\{,;:=\*/\s]|^)\s*(\+|\-)(\w.*)')
    p_colon = re.compile(r'(.*?\S|^)\s*(:)\s*(\S.*|$)')
    p_ellipsis = re.compile(r'(.*?\S|^)\s*(\.\.\.)\s*(\S.*|$)')
    p_op_dot = re.compile(r'(.*?\S|^)\s*(\.)\s*(\+|\-|\*|/|\^)\s*(=)\s*(\S.*|$)')
    p_pow_dot = re.compile(r'(.*?\S|^)\s*(\.)\s*(\^)\s*(\S.*|$)')
    p_pow = re.compile(r'(.*?\S|^)\s*(\^)\s*(\S.*|$)')
    p_op_comb = re.compile(r'(.*?\S|^)\s*(\.|\+|\-|\*|\\|/|=|<|>|\||\&|!|~|\^)\s*(<|>|=|\+|\-|\*|/|\&|\|)\s*(\S.*|$)')
    p_not = re.compile(r'(.*?\S|^)\s*(!|~)\s*(\S.*|$)')
    p_op = re.compile(r'(.*?\S|^)\s*(\+|\-|\*|\\|/|=|!|~|<|>|\||\&)\s*(\S.*|$)')
    p_func = re.compile(r'(.*?\w)(\()\s*(\S.*|$)')
    p_open = re.compile(r'(.*?)(\(|\[|\{)\s*(\S.*|$)')
    p_close = re.compile(r'(.*?\S|^)\s*(\)|\]|\})(.*|$)')
    p_comma = re.compile(r'(.*?\S|^)\s*(,|;)\s*(\S.*|$)')
    p_multiws = re.compile(r'(.*?\S|^)(\s{2,})(\S.*|$)')

    def cellIndent(self, line, cellOpen:str, cellClose:str, indent):
        # clean line from strings and comments
        pattern = re.compile(fr'(\s*)((\S.*)?)(\{cellOpen}.*$)')
        line = self.cleanLineFromStringsAndComments(line)
        opened = line.count(cellOpen) - line.count(cellClose)
        if opened > 0:
            m = pattern.match(line)
            n = len(m.group(2))
            indent = (n+1) if self.matrixIndent else self.iwidth
        elif opened < 0:
            indent = 0
        return (opened, indent)

    def multilinematrix(self, line):
        tmp, self.matrix = self.cellIndent(line, '[', ']', self.matrix)
        return tmp

    def cellarray(self, line):
        tmp, self.cell = self.cellIndent(line, '{', '}', self.cell)
        return tmp

    # indentation
    ilvl = 0
    istep = []
    fstep = []
    iwidth = 0
    matrix = 0
    cell = 0
    isblockcomment = 0
    islinecomment = 0
    longline = 0
    continueline = 0
    iscomment = 0
    separateBlocks = False
    ignoreLines = 0

    def __init__(self, indentwidth, separateBlocks, indentMode, operatorSep, matrixIndent):
        self.iwidth = indentwidth
        self.separateBlocks = separateBlocks
        self.indentMode = indentMode
        self.operatorSep = operatorSep
        self.matrixIndent = matrixIndent

    def cleanLineFromStringsAndComments(self, line):
        split = self.extract_string_comment(line)
        if split:
            return self.cleanLineFromStringsAndComments(split[0]) + ' ' + \
                self.cleanLineFromStringsAndComments(split[2])
        else:
            return line

    # divide string into three parts by extracting and formatting certain
    # expressions

    def extract_string_comment(self, part):
        # string
        m = self.p_string.match(part)
        m2 = self.p_string_dq.match(part)
        # choose longer string to avoid extracting subexpressions
        if m2 and (not m or len(m.group(2)) < len(m2.group(2))):
            m = m2
        if m:
            return (m.group(1), m.group(2), m.group(4))

        # comment
        m = self.p_comment.match(part)
        if m:
            self.iscomment = 1
            return (m.group(1) + ' ',  m.group(2), '')

        return 0

    def extract(self, part):
        # whitespace only
        m = self.p_blank.match(part)
        if m:
            return ('', ' ', '')

        # string, comment
        stringOrComment = self.extract_string_comment(part)
        if stringOrComment:
            return stringOrComment

        # decimal number (e.g. 5.6E-3)
        m = self.p_num_sc.match(part)
        if m:
            return (m.group(1) + m.group(2), m.group(3), m.group(4) + m.group(5))

        # rational number (e.g. 1/4)
        m = self.p_num_R.match(part)
        if m:
            return (m.group(1) + m.group(2), m.group(3), m.group(4) + m.group(5))

        # incrementor (++ or --)
        m = self.p_incr.match(part)
        if m:
            return (m.group(1), m.group(2) + m.group(3), m.group(4))

        # signum (unary - or +)
        m = self.p_sign.match(part)
        if m:
            return (m.group(1), m.group(2), m.group(3))

        # colon
        m = self.p_colon.match(part)
        if m:
            return (m.group(1), m.group(2), m.group(3))

        # dot-operator-assignment (e.g. .+=)
        m = self.p_op_dot.match(part)
        if m:
            sep = ' ' if self.operatorSep > 0 else ''
            return (m.group(1) + sep, m.group(2) + m.group(3) + m.group(4), sep + m.group(5))

        # .power (.^)
        m = self.p_pow_dot.match(part)
        if m:
            sep = ' ' if self.operatorSep > 0.5 else ''
            return (m.group(1) + sep, m.group(2) + m.group(3), sep + m.group(4))

        # power (^)
        m = self.p_pow.match(part)
        if m:
            sep = ' ' if self.operatorSep > 0.5 else ''
            return (m.group(1) + sep, m.group(2), sep + m.group(3))

        # combined operator (e.g. +=, .+, etc.)
        m = self.p_op_comb.match(part)
        if m:
            # sep = ' ' if m.group(3) == '=' or self.operatorSep > 0 else ''
            sep = ' ' if self.operatorSep > 0 else ''
            return (m.group(1) + sep, m.group(2) + m.group(3), sep + m.group(4))

        # not (~ or !)
        m = self.p_not.match(part)
        if m:
            return (m.group(1) + ' ', m.group(2), m.group(3))

        # single operator (e.g. +, -, etc.)
        m = self.p_op.match(part)
        if m:
            # sep = ' ' if m.group(2) == '=' or self.operatorSep > 0 else ''
            sep = ' ' if self.operatorSep > 0 else ''
            return (m.group(1) + sep, m.group(2), sep + m.group(3))

        # function call
        m = self.p_func.match(part)
        if m:
            return (m.group(1), m.group(2), m.group(3))

        # parenthesis open
        m = self.p_open.match(part)
        if m:
            return (m.group(1), m.group(2), m.group(3))

        # parenthesis close
        m = self.p_close.match(part)
        if m:
            return (m.group(1), m.group(2), m.group(3))

        # comma/semicolon
        m = self.p_comma.match(part)
        if m:
            return (m.group(1), m.group(2), ' ' + m.group(3))

        # ellipsis
        m = self.p_ellipsis.match(part)
        if m:
            return (m.group(1) + ' ', m.group(2), ' ' + m.group(3))

        # multiple whitespace
        m = self.p_multiws.match(part)
        if m:
            return (m.group(1), ' ', m.group(3))

        return 0

    # recursively format string
    def format(self, part):
        m = self.extract(part)
        if m:
            return self.format(m[0]) + m[1] + self.format(m[2])
        return part

    # compute indentation
    def indent(self, addspaces=0):
        indnt = ((self.ilvl+self.continueline)*self.iwidth + addspaces)*' '
        return indnt

    # take care of indentation and call format(line)
    def formatLine(self, line):

        if (self.ignoreLines > 0):
            self.ignoreLines -= 1
            return (0, self.indent() + line.strip())

        # determine if linecomment
        if re.match(self.linecomment, line):
            self.islinecomment = 2
        else:
            # we also need to track whether the previous line was a commment
            self.islinecomment = max(0, self.islinecomment-1)

        # determine if blockcomment
        if re.match(self.blockcomment_open, line):
            self.isblockcomment = float("inf")
        elif re.match(self.blockcomment_close, line):
            self.isblockcomment = 1
        else:
            self.isblockcomment = max(0, self.isblockcomment-1)

        # find ellipsis
        self.iscomment = 0
        strippedline = self.cleanLineFromStringsAndComments(line)
        ellipsisInComment = self.islinecomment == 2 or self.isblockcomment
        if re.match(self.block_close, strippedline) or ellipsisInComment:
            self.continueline = 0
        else:
            self.continueline = self.longline
        if re.match(self.ellipsis, strippedline) and not ellipsisInComment:
            self.longline = 1
        else:
            self.longline = 0

        # find comments
        if self.isblockcomment:
            return(0, line.rstrip()) # don't modify indentation in block comments
        if self.islinecomment == 2:
            # check for ignore statement
            m = re.match(self.ignore_command, line)
            if m:
                if m.group(1) and int(m.group(1)) > 1:
                    self.ignoreLines =  int(m.group(1))
                else:
                    self.ignoreLines =  1
            return (0, self.indent() + line.strip())

        # find imports, clear, etc.
        m = re.match(self.ctrl_ignore, line)
        if m:
            return (0, self.indent() + line.strip())

        # find matrices
        tmp = self.matrix
        if self.multilinematrix(line) or tmp:
            return (0, self.indent(tmp) + self.format(line).strip())

        # find cell arrays
        tmp = self.cell
        if self.cellarray(line) or tmp:
            return (0, self.indent(tmp) + self.format(line).strip())

        # find control structures
        m = re.match(self.ctrl_1line, line)
        if m:
            return (0, self.indent() + m.group(2) + ' ' + self.format(m.group(3)).strip() + ' ' + m.group(4) + ' ' + self.format(m.group(6)).strip())

        m = re.match(self.fcnstart, line)
        if m:
            offset = self.indentMode
            self.fstep.append(1)
            if self.indentMode == -1:
                offset = int(len(self.fstep) > 1)
            return (offset, self.indent() + m.group(2) + ' ' + self.format(m.group(3)).strip())

        m = re.match(self.ctrlstart, line)
        if m:
            self.istep.append(1)
            return (1, self.indent() + m.group(2) + ' ' + self.format(m.group(3)).strip())

        m = re.match(self.ctrlstart_2, line)
        if m:
            self.istep.append(2)
            return (2, self.indent() + m.group(2) + ' ' + self.format(m.group(3)).strip())

        m = re.match(self.ctrlcont, line)
        if m:
            return (0, self.indent(-self.iwidth) + m.group(2) + ' ' + self.format(m.group(3)).strip())

        m = re.match(self.ctrlend, line)
        if m:
            if len(self.istep) > 0:
                step = self.istep.pop()
            elif len(self.fstep) > 0:
                step = self.fstep.pop()
            else:
                print('warning:There are more end-statements than blocks!', file=sys.stderr)
                print('continue...')
                step = 0
            return (-step, self.indent(-step*self.iwidth) + m.group(2) + ' ' + self.format(m.group(4)).strip())

        return (0, self.indent() + self.format(line).strip())

    # format file from line 'start' to line 'end'
    def formatFile(self, filename, start, end):
        # read lines from file
        wlines = rlines = []

        if filename == '-':
            with sys.stdin as f:
                rlines = f.readlines()[start-1:end]
        else:
            with open(filename, 'r', encoding='UTF-8') as f:
                rlines = f.readlines()[start-1:end]

        # take care of empty input
        if not rlines:
            rlines = ['']

        # get initial indent lvl
        p = r'(\s*)(.*)'
        m = re.match(p, rlines[0])
        if m:
            self.ilvl = len(m.group(1))//self.iwidth
            rlines[0] = m.group(2)

        blank = True
        for line in rlines:
            # remove additional newlines
            if re.match(r'^\s*$', line):
                if not blank:
                    blank = True
                    wlines.append('')
                continue

            # format line
            (offset, line) = self.formatLine(line)

            # adjust indent lvl
            self.ilvl = max(0, self.ilvl + offset)

            # add newline before block
            if (self.separateBlocks and offset > 0 and
                    not blank and not self.islinecomment):
                wlines.append('')

            # add formatted line
            wlines.append(line.rstrip())

            # add newline after block
            if self.separateBlocks and offset < 0:
                wlines.append('')
                blank = True
            else:
                blank = False

        # remove last line if blank
        while wlines and not wlines[-1]:
            wlines.pop()

        # take care of empty output
        if not wlines:
            wlines = ['']

        # write output
        for line in wlines:
            print(line)

# -------------------------------
# 格式化封装函数
# -------------------------------
def format_m_text(filename):
    """
    格式化M文件，返回格式化后的文本字符串
    """
    options = dict(
        startLine=1,
        endLine=None,
        indentWidth=4,
        separateBlocks=True,
        indentMode='all_functions',
        addSpaces='exclude_pow',
        matrixIndent='aligned'
    )

    indentModes = dict(all_functions=1, only_nested_functions=-1, classic=0)
    operatorSpaces = dict(all_operators=1, exclude_pow=0.5, no_spaces=0)
    matrixIndentation = dict(aligned=1, simple=0)

    indent = options['indentWidth']
    start = options['startLine']
    end = options['endLine']
    sep = options['separateBlocks']
    mode = indentModes.get(options['indentMode'], indentModes['all_functions'])
    opSp = operatorSpaces.get(options['addSpaces'], operatorSpaces['exclude_pow'])
    matInd = matrixIndentation.get(options['matrixIndent'], matrixIndentation['aligned'])

    formatter = Formatter(indent, sep, mode, opSp, matInd)

    with open(filename, 'r', encoding='utf-8') as f:
        rlines = f.readlines()[start-1:end]

    if not rlines:
        return ""

    p = r'(\s*)(.*)'
    m = re.match(p, rlines[0])
    if m:
        formatter.ilvl = len(m.group(1)) // formatter.iwidth
        rlines[0] = m.group(2)

    wlines = []
    blank = True
    for line in rlines:
        if re.match(r'^\s*$', line):
            if not blank:
                blank = True
                wlines.append('')
            continue

        offset, formatted = formatter.formatLine(line)
        formatter.ilvl = max(0, formatter.ilvl + offset)

        if (formatter.separateBlocks and offset > 0 and
                not blank and not formatter.islinecomment):
            wlines.append('')

        wlines.append(formatted.rstrip())

        if formatter.separateBlocks and offset < 0:
            wlines.append('')
            blank = True
        else:
            blank = False

    while wlines and not wlines[-1]:
        wlines.pop()

    return "\n".join(wlines) if wlines else ""

# -------------------------------
# 正则微调 spacing
# -------------------------------

def fix_spacing(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    new_lines = []
    for line in lines:
        line = line.rstrip()
        # 替换dottrans为 .'
        line = re.sub(r'(\w+|\))\s*\.?\s*dottrans\b', r"\1.'", line)
        # 共轭转置 .'
        # line = re.sub(r"\.\s*'", ".'", line)
        line = re.sub(r'\b([A-Za-z_]\w*)\s+\(', r'\1(', line)
        # 空格
        line = re.sub(r'(\w)\s*\.\s*(\w)', r'\1.\2', line)
        # 后接 ( 或 { 空格
        line = re.sub(r'(\w)\s*\.\s*(\(|\{)', r'\1.\2', line)
        # 变量和转置符空格
        line = re.sub(r"(\w)\s*'", r"\1'", line)
        # 变量和 { 索引空格
        line = re.sub(r'(\w)\s+\{', r'\1{', line)

        new_lines.append(line)

    with open(filepath, "w", encoding="utf-8") as f:
        f.write('\n'.join(new_lines))

# -------------------------------
# 格式化入口
# -------------------------------

def format_m_file(mfile_path):
    print("正在格式化：", mfile_path)
    try:
        formatted_text = format_m_text(mfile_path)
        if not formatted_text.strip():
            raise Exception("格式化后为空，请检查输入内容或格式化逻辑")

        with open(mfile_path, "w", encoding="utf-8") as f:
            f.write(formatted_text)

        fix_spacing(mfile_path)
        print("已格式化：", mfile_path)

    except Exception as e:
        print(f"格式化出错：{e}")
        messagebox.showerror("格式化出错", str(e))

# -------------------------------
# 解密+格式化逻辑
# -------------------------------

def process_files(files):
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

            format_m_file(output_mfile)

        except Exception as e:
            messagebox.showerror("错误", f"处理失败：\n{e}")
            return

    messagebox.showinfo("完成", "全部 P 文件已转换为 M 文件")


# -------------------------------
# 解密+格式化逻辑
# -------------------------------

def process_files(files):
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

            format_m_file(output_mfile)

        except Exception as e:
            messagebox.showerror("错误", f"处理失败：\n{e}")
            return

    messagebox.showinfo("完成", "全部 P 文件已转换为 M 文件")

# -------------------------------
# Tkinter界面逻辑
# -------------------------------

def select_and_decrypt():
    files = filedialog.askopenfilenames(title="选择 P 文件", filetypes=[("P Files", "*.p")])
    if not files:
        return

    t = threading.Thread(target=process_files, args=(files,))
    t.start()
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
