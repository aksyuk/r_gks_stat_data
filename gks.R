
# TODO: add .doc parser from package "textreadr"
# TODO: how to get rid of a console dialog? Can we find a match between year 
#       books indexes and transform it to function arguments? 

# load packages ----------------------------------------------------------------
require("crayon")                 # colored console messages
require("RCurl")                  # loading URLs
require("XML")                    # parsing html
require("xml2")                   # parsing docx
require("tools")                  # file_ext()
library("broman")                 # hex to decimal convertion
library("BMS")                    # hex to binary convertion: hex2bin()

# init -------------------------------------------------------------------------
# url of the data host
glb_host_name <- "http://www.gks.ru"
# global object with scrapped data
glb_dataGKS <- NULL

# get all stat years and path to it's db
index_web_page <- 
    "https://www.gks.ru/folder/210/document/13204"
html <- getURL(index_web_page)
doc <- htmlTreeParse(html, useInternalNodes = T)
rootNode <- xmlRoot(doc)
xpath_pattern <- "//div[contains(text(), 'Социально-экономические показатели')]"
xpath_pattern <- "//div[@class='document-list__item-title']"

years_v <- xpathSApply(rootNode, xpath_pattern, xmlValue)
years_v[1]
x <- gsub('[\u008b]|[\u0081]|[\u0086]|[\u0087]|[\u008c]|[\u008d]|[\u0082]', 
          '', years_v[1])
x
Encoding(x) <- 'CP1251'
x

for (j in 1:length(stri_enc_list())) {
    loop.enc <- stri_enc_list()[[j]]
    
    Sys.sleep(3)
    cat(green(paste0(names(stri_enc_list())[j], "\r\n")))
    
    for (i in 1:length(loop.enc)) {
        x <- years_v
        Encoding(x) <- loop.enc[i]
        
        cat(yellow(paste0(loop.enc[i], "\r\n")))
        print(x[1])
    }
}


years_v[1]


Encoding(years_v) <- 'cp1251'
as_utf8(years_v)[1]
years_v[1]

iconv(years_v, from = Encoding(years_v)[1], to = "CP1252", sub = "byte")
years_v


years_v <- years_v[nchar(years_v) == 4]
db_urls <- xpathSApply(rootNode, paste0(xpath_pattern, "/a"), xmlGetAttr, "href")

# index of databases by year
glb_years <- data.frame(years_v, db_paths = db_urls[1:length(years_v)],
                        stringsAsFactors = F)

# remove temp objects
rm(html, doc, rootNode, xpath_pattern, years_v, db_urls)


# functions --------------------------------------------------------------------

# vectorize seq()
# source: https://stackoverflow.com/questions/15917233/elegant-way-to-vectorize-seq
seq2 <- Vectorize(seq.default, vectorize.args = c("from", "to"))

# returns string w/o leading or trailing whitespace
# source: https://stackoverflow.com/questions/2261079/how-to-trim-leading-and-trailing-whitespace
trim <- function(x) gsub("^\\s+|\\s+$", "", x)

# count repeat values for merged table cells
indexRep <- function(x, index) {
    res <- NULL
    frst <- 1
    for (i in 1:length(index)) {
        index_sample <- x[frst:(frst + index[i] - 1)]
        res <- c(res, sum(index_sample))
        frst <- frst + index[i]
    }
    return(res)
}

# convert cyrillic strings to prorer encoding ==================================
toLocalEncoding <- function(x, sep = ",", quote = TRUE, encoding = "utf-8") {
    rawcsv <- tempfile()
    write.csv(x, file = rawcsv)
    result <- read.csv(rawcsv, encoding = "UTF-8")
    unlink(rawcsv)
    result
    }

# parse table header from html table
parseTableHeader <- function(html_table) {
    # find max rowspan
    max_rowspan <- as.numeric(max(unlist(xmlSApply(html_table, 
                            function(x) {xpathSApply(x, "./tr/td",
                                                     xmlGetAttr, 
                                                     name = "rowspan")}))))
    # # number of cells in data row
    # n_col <- 
    #     length(xmlSApply(html_table, 
    #                      function(x) {xpathSApply(x, paste0("//table/tr[",
    #                                                         max_rowspan + 1,
    #                                                         "]/td"), xmlValue)}))
    
    # extract header rows to list
    header_lines <- lapply(as.list(1:max_rowspan), function(x) {
        xmlSApply(html_table, 
                  function(y) {xpathSApply(y, paste0("//table/tr[", x, "]/td"), 
                                           xmlValue)})
    })
    header_lines[[1]][1, 1] <- "obs"
    header_lines <- sapply(header_lines, function(x) as.vector(t(x)))
    
    # extract colspans to list
    header_colspans <- lapply(as.list(1:max_rowspan), function(x) {
        xmlSApply(html_table, 
                  function(y) {xpathSApply(y, paste0("//table/tr[", x, "]/td"), 
                                           xmlGetAttr, name = "colspan")})
    })
    header_colspans <- sapply(header_colspans, function(x) {
        sapply(x[, 1], function(y) ifelse(is.null(y), NA, as.numeric(y)))
        })
    
    # combine header text and colspans into one list
    header_dfs <- as.list(1:length(header_lines))
    for (i in 1:length(header_lines)) {
        header_dfs[[i]] <- data.frame(header = header_lines[[i]],
                                      colspan = header_colspans[[i]],
                                      stringsAsFactors = F)
    }
    
    # return header lines with colspan information in list
    return(header_dfs)
}

# add repeat numbers for merged header cells, using colspans
addRepeatsToHeaderList <- function(lines_df_list) {
    # add repeat numbers to each line of header (based on colspans)
    lines_df_list <- lapply(lines_df_list, function(x) {
        x$rep <- x$colspan
        x$rep[is.na(x$rep)] <- 1
        x
    })
    
    # increase repeat numbers considering colspans in parent header lines
    for (i_header in seq(length(lines_df_list) - 1, 1, -1)) {
        if (!all(lines_df_list[[i_header]]$rep != 1)) {
            which_neq_1 <- which(lines_df_list[[i_header]]$rep != 1)
            lines_df_list[[i_header]]$rep[which_neq_1] <- 
                indexRep(lines_df_list[[i_header + 1]]$rep, 
                         lines_df_list[[i_header]]$rep[lines_df_list[[i_header]]$rep != 1])
        }
    }
    
    # return list, where each element (data.frame) has additional cilumn with
    #  repeat numbers
    return(lines_df_list)
}

# construct table header from merged cells
mergeTableHeader <- function(header_list) {
    # this is base (bottom) header line
    df_header <- header_list[[length(header_list)]]$header
    
    # loop header lines from (n-1) to 1 and concat general header
    for (i_list in seq(length(header_list) - 1, 1, -1)) {
        if (sum(!is.na(header_list[[i_list]]$colspan))) {
            parent_line <- rep(header_list[[i_list]]$header, 
                               header_list[[i_list]]$rep)
            
            # matrix of two columns: 1st - positions of non-NA colspan cells,
            #  2nd - value of colspan for each position
            index_which_add <- 
                cbind(which(!is.na(header_list[[i_list]]$colspan)), 
                      header_list[[i_list]]$rep[!is.na(header_list[[i_list]]$colspan)])
            
            # correct starting positions considering colspan values
            index_which_add[-1, 1] <- 
                index_which_add[-1, 1] + 
                cumsum(index_which_add[-nrow(index_which_add), 2] - 1) 
            # calc positions of repeating header cells and convert into vector
            index_which_add <- 
                unlist(seq2(index_which_add[, 1], 
                            index_which_add[, 1] + index_which_add[, 2] - 1, by = 1))
            
            # repeat header cells using index_which_add vector
            parent_line[index_which_add] <- 
                paste(parent_line[index_which_add], df_header, sep = ".")
            
            # save result
            df_header <- parent_line
        }
    }
    
    # remove "-" symbols
    df_header <- gsub("-", "", df_header)
    
    # return header as a vector
    return(df_header)
}

# scrape data from html table ==================================================
getTableFromHtm <- function(ref) {
    # 1. download html and deal with cyrillic ..................................
    # temp files for encoding change
    tmp_filename_src <- tempfile()
    tmp_filename_res <- tempfile()
    # download table
    download.file(ref, tmp_filename_src)
    # read and write html file to change encoding
    con <- file(tmp_filename_res, encoding = "UTF-8")
    writeLines(iconv(readLines(tmp_filename_src), from = "CP1251", to = "UTF8"),
               con, sep = "")
    html_raw <- paste0(readLines(tmp_filename_res), collapse = "")
    # remove temp files and close file connection
    close(con)
    unlink(tmp_filename_src)
    unlink(tmp_filename_res)
    
    # 2. clean downloaded html from markup noise and parse .....................
    # # delete newlines and <b>, <br> tags
    # html.raw <- gsub("(\r\n|<BR>|</BR>|<B>|</B>)", "", html.raw)
    # fix " symbol
    html_raw <- gsub('\"', "'", html_raw)
    # fix undreakable space symbol
    html_raw <- gsub("&nbsp;", " ", html_raw)
    # remove <sup> tags and keep footnote number
    html_raw <- gsub("<SUP>(.{1,3})</SUP>", "[#\\1]", html_raw)
    # parse as html
    html_nodes <- htmlParse(html_raw, encoding = "UTF-8")
    
    # 3. process tables ........................................................
    count_tables <- xpathSApply(html_nodes, "//table", xmlValue)
    if (!is.null(count_tables)) {
        count_tables <- length(count_tables)
        } else {
        count_tables <- 0
        stop("Нет таблицы в источнике")
        }
    
    if (count_tables > 1) {
        message(paste0("На странице обнаружено ", count_tables, "таблиц"))
        i_table <- as.numeric(readline(prompt = "Введите номер таблицы (Esc для выхода): "))
        if (is.na(i_table)) {
            stop("Номер таблицы не является числом")
        } else {
            if (i_table < 1 | i_table > count_tables | i_table != floor(i_table)) {
                stop("Введён неверный номер таблицы")
            }
        }
        } else {
            i_table <- 1
        }
    
    cat(green(paste0("Обрабатываю таблицу ", i_table, " из ", 
                     count_tables, "\r\n")))
    
    # parse table from html: get list of cells in rows
    html_this_table <- getNodeSet(html_nodes, 
                                  paste0("//table[", i_table, "]"))
    list_this_table <- 
        sapply(getNodeSet(html_nodes, 
                          paste0("//table[", i_table, "]/tr")),
               function(x) {xpathSApply(x, "./td", xmlValue)})
    
    
    # дальше разбор таблицы, по идее общий для doc, html, docx .................
    
    # how much cells in a row depends of the row type         
    row_len <- sapply(list_this_table, length)
    
    # data cells are in the middle (fortunately, unmerged)
    df_data <- do.call("rbind", 
                       list_this_table[which(row_len == max(row_len))])
    
    # first short rows with various number of cells are headers with merged cells
    if (which(row_len == max(row_len))[1] > 1) {
        colnames(df_data) <- 
            mergeTableHeader(addRepeatsToHeaderList(parseTableHeader(html_this_table)))
    } else {
        df_data[1, 1] <- "obs"
        df_data[, 1] <- gsub("-", "", df_data[, 1])
        colnames(df_data) <- df_data[1, ]
        df_data <- df_data[-1, ]
    }
    
    # fix row/column headers
    df_data[1, ] <- gsub("-", "", df_data[1, ])
    colnames(df_data) <- gsub("\\.+", "\\.", make.names(colnames(df_data)))
    
    # last row with one cell only is a footnote
    if (row_len[length(row_len)] == 1) {
        # read footnotes
        footnote_text <- gsub("^_*", "", list_this_table[[length(row_len)]])
        footnote_text <- unlist(strsplit(gsub(")]", ")", footnote_text, 
                                              fixed = T),
                                         "\\[#"))
        footnote_text = footnote_text[footnote_text != ""]
        df_footnote <- 
            data.frame(number = gsub("(^[[:digit:]]{1,3})\\).*$", "\\1", 
                                     footnote_text), 
                       text = trim(footnote_text), stringsAsFactors = F)
        
        # link footnotes to data cells
        footnote_positions <- which(apply(df_data, 2, function(x) {
            grepl("\\[#([[:digit:]]{1,3})\\).*$", x)
        }), arr.ind = T)
        
        footnote_index <- gsub("^.*\\[#([[:digit:]]{1,3}\\)).*$", "\\1", 
                               df_data[footnote_positions])
        footnote_positions <- as.data.frame(footnote_positions)
        footnote_positions$footnote_index <- footnote_index
    }
    
    # transform data to data.frame
    df_data <- apply(df_data, 2, function(x) {gsub(",", ".", x)})
    df_data <- data.frame(df_data, stringsAsFactors = F)
    df_data[, -1] <- apply(df_data[, -1], 2, as.numeric)
    
    cat(green("Сохраняю данные\r\n"))
    
    # assign data to global object
    assign("glb_dataGKS", list(URL = ref,
                               download_time = paste0("file downloaded at: ",
                                                      Sys.time()),
                               data = df_data,
                               footnote_positions = footnote_positions,
                               footnote_legend = df_footnote), 
           .GlobalEnv)
    }


# find sequence position in bytes string
# little endian encoding means than bytes go in reverse order
findSeqPos <- function(seq_vec, txt_vec) {
    
    i_pos <- 1
    patt <- seq_vec[i_pos]
    res_pos <- which(txt_vec == patt)
    
    while (i_pos < length(seq_vec)) {
        i_pos <- i_pos + 1
        
        patt <- seq_vec[i_pos]
        patt_pos <- which(txt_vec == patt)
        res_pos <- patt_pos[patt_pos %in% (res_pos + i_pos - 1)] - i_pos + 1
    }
    return(res_pos)
}


# get tables from doc ==========================================================
getTableFromDoc <- function(word_doc) {
    
    #! debug
    word_doc <- my_url
    
    # create temp links for docx file
    tmpf <- tempfile()
    
    # download to temp destfile
    download.file(word_doc, tmpf)
    
    # read doc file as binary to vector, where each element is a BYTE, 
    #  coded with hex number
    doc_bin <- file(tmpf, "rb")
    doc_bin <- readBin(doc_bin, raw(), n = file.info(tmpf)$size, 
                       endian = "little")
    
    # remove link to temp file
    unlink(tmpf)
    
    # find FIB in binary doc (starts with 0xA5EC, 1472 bytes length)
    # source: http://b2xtranslator.sourceforge.net/howtos/How_to_retrieve_text_from_a_binary_doc_file.pdf
    FIB_base_pos <- findSeqPos(rev(c("a5", "ec")), doc_bin)
    # FIB base length is 32 bytes
    FIB_base <- doc_bin[FIB_base_pos:(FIB_base_pos + 32 - 1)]
    
    # extract nFIB, 2 bytes after FIB start
    nFIB <- paste0("0x", paste0(rev(doc_bin[(FIB_base_pos + 2):(FIB_base_pos + 3)]), 
                                collapse = ""))
    nFibs <- c(0x00c1, 0x00d9, 0x0101, 0x010c, 0x0112)
    cbRgFcLcbs <- c("005d", "006c", "0088", "00a4", "00b7")
    cbRgFcLcb <- cbRgFcLcbs[nFibs == as.numeric(nFIB)]
    
    # extract FibRgFcLcb block of FIB
    cbRgFcLcb_pos <- findSeqPos(c(substr(cbRgFcLcb, 3, 4), 
                                  substr(cbRgFcLcb, 1, 2)), doc_bin)
    cbRgFcLcb_pos <- cbRgFcLcb_pos[cbRgFcLcb_pos > (FIB_base_pos + 32)][1]
    
    #! debug
    doc_bin[cbRgFcLcb_pos:(cbRgFcLcb_pos + 1)]
    tmp <- doc_bin[(cbRgFcLcb_pos + 2):(cbRgFcLcb_pos + 2 + (68 - 1) * 4 + 4 - 1)]
    tmp <- matrix(tmp, length(tmp) / 4, 4, byrow = T)
    tmp
    
    # here we find piece table position and length in xTable stream
    fcClx_bytes_pos <- (cbRgFcLcb_pos + 2 + (67 - 1) * 4)
    fcClx <- as.numeric(paste0("0x", paste0(rev(doc_bin[fcClx_bytes_pos:(fcClx_bytes_pos + 4 - 1)]),
                                            collapse = "")))
    lcbClx_bytes_pos <- (cbRgFcLcb_pos + 2 + (68 - 1) * 4)
    lcbClx <- as.numeric(paste0("0x", paste0(rev(doc_bin[lcbClx_bytes_pos:(lcbClx_bytes_pos + 4 - 1)]),
                                             collapse = "")))

    # extract FibRgLw97 block of FIB, preceed by cslw = 0x0016
    cslw_pos <- findSeqPos(rev(c("00", "16")), doc_bin)
    cslw_pos <- cslw_pos[cslw_pos > (FIB_base_pos + 32)][1]
    FibRgLw97 <- doc_bin[(cslw_pos + 2):(cslw_pos + 2 + 88 - 1)]
    
    #! debug
    doc_bin[cslw_pos:(cslw_pos + 1)]
    
    # here we find a lot of counts for doc element (just in case): 
    #   1. number of CPs in the main document
    ccpText <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[13:16]), 
                                              collapse = "")))
    #   2. number of CPs in the footnote subdocument
    ccpFtn <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[17:20]), 
                                             collapse = "")))
    #   3. number of CPs in the header subdocument
    ccpHdd <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[21:24]), 
                                             collapse = "")))
    #   4. number of CPs in the comment subdocument
    ccpAtn <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[29:32]), 
                                             collapse = "")))
    #   5. number of CPs in the endnote subdocument
    ccpEdn <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[33:36]), 
                                             collapse = "")))
    #   6. number of CPs in the textbox subdocument of the main document
    ccpTxbx <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[37:40]), 
                                             collapse = "")))
    #   7. number of CPs in the textbox subdocument of the header
    ccpHdrTxbx <- as.numeric(paste0("0x", paste0(rev(FibRgLw97[41:44]), 
                                              collapse = "")))
    
    # searching for a piece table
    # 16 bitsat position 0x000A sets a name of file stream
    pos <- 0x000A + 1
    if (hex2bin(paste0("0x", FIB_base[pos]))[7] == 1) {
        table_stream_name <- "1Table"
    } else {
        table_stream_name <- "0Table"
    }
    
    table_stream_name_bytes <- 
        sapply(seq(1, nchar(table_stream_name)), 
               function(x) format(as.hexmode(as.character(charToRaw(substr(table_stream_name, x, x)))), width = 4))
   
    table_stream_name_bytes <- c(sapply(table_stream_name_bytes, function(x) {
        x <- c(substr(x, 3, 4), substr(x, 1, 2))
    }))
    
    xTable_pos <- findSeqPos(table_stream_name_bytes, doc_bin)
    
    # read table file stream
    xTable <- doc_bin[xTable_pos:length(doc_bin)]
    
    # read Clx
    xTable[fcClx:fcClx + 4]
    
    
    i_pos <- 1
    pos_of_02 <- which(xTable[i_pos:table_len] == 2)[1]
    while (!is.na(pos_of_02)) {
        xTable_part <- xTable[pos_of_02:table_len]
        as.numeric(paste0("0x", paste0(rev(xTable_part[2:5]), collapse = "")))
        # i_pos <- 
    }

    
    
    cell_end <- c("12", "f0")
    sec_symb <- grep(c(cell_end[1]), FIB)
    fst_symb <- grep(c(cell_end[2]), FIB)
    paste0(FIB[fst_symb[fst_symb %in% (sec_symb - 1)]],
           FIB[sec_symb[sec_symb %in% (fst_symb + 1)]])
    
    # find FIB base in binary doc (starts with 0xA5EC, 1472 bytes length)
    a5_pos <- which(doc_bin == "a5")
    ec_pos <- which(doc_bin == "ec")
    FIB_base_pos <- ec_pos[ec_pos %in% (a5_pos - 1)]
    FIB_base <- doc_bin[FIB_base_pos:(FIB_base_pos + 1472 + 1)]
    # little endian encoding means than bytes go in reverse order
    FIB_base <- FIB_base[seq(length(FIB_base), 1, -1)]
    
    FIB_base[which(FIB_base == "a5"):(which(FIB_base == "a5") - 1)]
    
    
    
    
    

    
    doc_bin[]
    
    
    
    
        
}


# get tables from docx =========================================================
getTableFromDocx <- function(word_doc) {
    # create temp links for docx file
    tmpd <- tempdir()
    tmpf <- tempfile(tmpdir = tmpd, fileext = ".zip")
    
    # copy & unzip docx to temp destfile
    file.copy(word_doc, tmpf)
    unzip(tmpf, exdir = sprintf("%s/docdata", tmpd))
    
    # parse docx file
    doc <- read_xml(sprintf("%s/docdata/word/document.xml", tmpd))
    
    # remove temp files
    unlink(tmpf)
    unlink(sprintf("%s/docdata", tmpd), recursive = TRUE)
    
    # docx namespace
    ns <- xml_ns(doc)
    # list of all tables
    tbls <- xml_find_all(doc, ".//w:tbl", ns = ns)
    
    lapply(tbls, function(tbl) {
        cells <- xml_find_all(tbl, "./w:tr/w:tc", ns = ns)
        rows <- xml_find_all(tbl, "./w:tr", ns = ns)
        dat <- data.frame(matrix(xml_text(cells), 
                                 ncol = (length(cells)/length(rows)), 
                                 byrow = TRUE), stringsAsFactors = FALSE)
        colnames(dat) <- dat[1, ]
        dat <- dat[-1, ]
        rownames(dat) <- NULL
        assign("glb_dataGKS", dat, .GlobalEnv)
        })
    }

# load doc with stat data and convert containing table into dataframe ==========
loadGKSData <- function(ref){
    ext <- file_ext(ref)
    switch(ext,
           doc = {
               # ref url points to docx file
               print("Неподдерживаемый тип источника")
               # cat(green("Читаю данные из файла .doc...\n"))
               # getTableFromDoc(ref)
           },
           # ref url points to docx file
           docx = {
               cat(green("Читаю данные из файла .docx...\n"))
               getTableFromDocx(ref)
           },
           # ref url points to html file
           html = {
               cat(green("Читаю данные с веб-страницы...\n"))
               getTableFromHtm(ref)
           },
           htm = {
               cat(green("Читаю данные с веб-страницы...\n"))
               getTableFromHtm(ref)
           },
           # other extention
           stop(paste0("Неизвестный формат файла: ", ext))
    )
    }

# dialog with user and return reference to needed stat doc =====================
getGKSDataRef <- function() {
    params <- "/?List&Id="
    id <- -1
    
    # ask user to choose a year
    year <- readline(prompt = paste0("Введите год от ", glb_years$years_v[1],
                                    " до ", tail(glb_years$years_v, n = 1), 
                                    " (Esc для выхода): "))
    current_selection_string <- year
    cat(green(paste0(current_selection_string, "\n")))
    db_path <- glb_years[glb_years$years_v == year, "db_paths"]
    
    # go through tree until we got a link to doc instead of id in stat db
    while (TRUE) {
        url <- paste(db_path, params, id, sep = "")
        xml <- xmlTreeParse(url, useInternalNodes = T)
        names <- xpathSApply(xml, "//name", xmlValue)
        Encoding(names) <- "UTF-8" 
        refs <- xpathSApply(xml, "//ref", xmlValue)
        
        # options to choose
        for (i in 1:length(names)) {
            message(paste0(i, ". ", names[i]))
            }
        
        # ask user to select an option number
        num <- as.numeric(readline(prompt = "Введите номер раздела (Esc для выхода): "))
        current_selection_string <- paste0(current_selection_string, " -> ", 
                                           names[num])
        cat(green(paste0(current_selection_string, "\n")))
        ref <- refs[as.numeric(num)]
        
        # got a database id
        if (substr(ref, 1, 1) != "?") {
            # preface has an artifact in url
            ref <- gsub("[/]%3Cextid%3E[/]%3Cstoragepath%3E::[|]", "", ref)
            ref <- paste0(glb_host_name, ref)
            cat(blue(paste0(ref, "\n")))
            return(ref)
        }
        
        id <- substr(ref, 2, nchar(ref))
        }
    }
