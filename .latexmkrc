# Latexmk 配置文件 for Psalter Generator
# 强制使用 LuaLaTeX 编译

$pdf_mode = 4;           # 使用 lualatex 生成 PDF
$lualatex = 'lualatex -interaction=nonstopmode -shell-escape %O %S';
$postscript_mode = $dvi_mode = 0;

# 清理临时文件
$clean_ext = 'aux log out toc lof lot fls fdb_latexmk synctex.gz';
