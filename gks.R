
# load packages ----------------------------------------------------------------
require("RCurl")
require("XML")
require("xml2")
require("tools")
# require("textreadr")

# init -------------------------------------------------------------------------
# url of the data host
host_name <- "http://www.gks.ru"
# global object with scrapped data
dataGKS <- NULL

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
years <- data.frame(years_v, db_paths = db_urls[1:length(years_v)])

# remove temp objects
rm(html, doc, rootNode, xpath_pattern, years_v, db_urls)

# functions --------------------------------------------------------------------
# convert cyrillic strings to prorer incoding
toLocalEncoding <- function(x, sep = ",", quote = TRUE, encoding = "utf-8") {
    rawcsv <- tempfile()
    write.csv(x, file = rawcsv)
    result <- read.csv(rawcsv, encoding = "UTF-8")
    unlink(rawcsv)
    result
    }

# scrape data from html table
getTableFromHtm <- function(ref) {
    url <- paste(host_name, ref, sep = "")
    doc <- htmlParse(url, encoding = "Windows-1251")
    if (length(xpathSApply(doc,"//table", xmlValue)) == 0) {
        print("Нет таблицы в источнике")
        return()
        }
    
    assign("dataGKS", readHTMLTable(doc, trim = TRUE, which = 1,
                                    stringsAsFactors = FALSE, 
                                    as.data.frame = TRUE), .GlobalEnv)
    names(dataGKS) <<- gsub("[\r\n]", "", names(dataGKS))
    dataGKS[, 1] <<- gsub("[\r\n]", "", dataGKS[, 1])
    dataGKS <<- toLocalEncoding(dataGKS)
}

# get tables from doc
getTableFromDoc <- function(word_doc) {
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
        assign("dataGKS", dat, .GlobalEnv)
        })
    }

# load doc with stat data and convert containing table into dataframe
loadGKSData <- function(ref){
    ext <- file_ext(ref)
    if (ext == "doc") {
        print("Неподдерживаемый тип источника")
        #getTableFromDoc(ref)
        }
    if (ext == "docx") {
        getTableFromDoc(ref)
    } else if (ext == "htm") {
      getTableFromHtm(ref)
    } else {
      print(paste0(ext, ": неподдерживаемый тип источника")) 
    }
}

# dialog with user and return reference to needed stat doc
getGKSDataRef <- function() {
    params <- "/?List&Id="
    id <- -1
    year <- readline(prompt = paste("Введите год от", years$years_v[1],
                                    "до", tail(years$years_v, n = 1), " "))
    db_path <- years$db_paths[years_v == year]
    # go through tree until we got a link to doc instead of id in stat db
    while (TRUE) {
        url <- paste(db_path, params, id, sep = "")
        xml <- xmlTreeParse(url, useInternalNodes = T)
        names <- xpathSApply(xml, "//name", xmlValue)
        Encoding(names) <- "UTF-8" 
        refs <- xpathSApply(xml, "//ref", xmlValue)
        
        for (i in 1:length(names)) {
            print(paste(i, names[i]))
            }
        
        num <- readline(prompt = "Введите номер ")
        ref <- refs[as.numeric(num)]
        
        if (substr(ref, 1, 1) != "?")
            return(ref)
        id <- substr(ref, 2, nchar(ref))
        }
    }
