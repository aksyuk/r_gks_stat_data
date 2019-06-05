
# TODO: add .doc parser from package "textreadr"
# TODO: how to get rid of a console dialog? Can we find a match between year 
#       books indexes and transform it to function arguments? 

# load packages ----------------------------------------------------------------
require("crayon")
require("RCurl")
require("XML")
require("xml2")
require("tools")
# require("textreadr")

# init -------------------------------------------------------------------------
# url of the data host
glb_host_name <- "http://www.gks.ru"
# global object with scrapped data
glb_dataGKS <- NULL

# get all stat years and path to it's db
index_web_page <- 
    "http://www.gks.ru/wps/wcm/connect/rosstat_main/rosstat/ru/statistics/publications/catalog/doc_1138623506156"
html <- getURL(index_web_page)
doc <- htmlTreeParse(html, useInternalNodes = T)
rootNode <- xmlRoot(doc)
xpath_pattern <- "//tbody/tr/td"
years_v <- xpathSApply(rootNode, xpath_pattern, xmlValue)
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


# get tables from doc ==========================================================
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
