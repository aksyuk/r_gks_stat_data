

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
    FIB <- doc_bin[FIB_base_pos:(FIB_base_pos + 1472 - 1)]
    FIB[c(hex2dec(0x01A2), hex2dec(0x01A2) - 1)]
    FIB[c(hex2dec(0x01A6), hex2dec(0x01A6) - 1)]
    FIB[c(hex2dec(0x000A), hex2dec(0x000A) - 1)]
    
    # Table1 or Table0? extract nFIB, 2 bytes after FIB start
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