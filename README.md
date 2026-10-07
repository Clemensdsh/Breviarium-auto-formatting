# Breviarium Auto-Formatting

用 LuaLaTeX 排版拉丁文-中文双栏对照日课经的小工具：`psalter.sty` 提供圣咏、对经、圣歌、红字等排版命令（支持 Gregorio 乐谱），`main.py` 是一个 Tkinter 图形界面，用来拼装内容并生成、编译 `body.tex`。可以直接 `python main.py` 运行，或用 `pyinstaller BreviariumGen.spec` 打包成 exe；编译需要装有 LuaLaTeX 与 GregorioTeX 的 TeX Live / MiKTeX。手写 LaTeX 的用法参见 `examples/sample_body.tex`。早期成果见 [羅馬大日課-耶穌聖誕瞻禮（中拉對照）](https://clemensdsh.github.io/2025/12/23/pdftest/)。本仓库已不再维护。
