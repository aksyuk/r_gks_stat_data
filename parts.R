
# Old way to deal with cyrillics -----------------------------------------------

# 1. download html and deal with cyrillic ..................................
# temp files for encoding change
tmp_filename_src <- tempfile()
tmp_filename_res <- tempfile()
# download table
download(ref, tmp_filename_src, mode = "wb")
# read and write html file to change encoding
con <- file(tmp_filename_res, encoding = "UTF-8")
writeLines(iconv(readLines(tmp_filename_src), from = "CP1251", to = "UTF-8"),
           con, sep = "")
html_raw <- paste0(readLines(tmp_filename_res), collapse = "")
# remove temp files and close file connection
close(con)
unlink(tmp_filename_src)
unlink(tmp_filename_res)