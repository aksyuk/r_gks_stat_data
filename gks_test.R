
# size of all odjects in memory
# source: https://www.r-bloggers.com/size-of-each-object-in-rs-workspace/
for (thing in ls()) {
    message(thing); print(object.size(get(thing)), units = 'auto')
    }

# find first, second table -----------------------------------------------------
exmpl <- "<TABLE><TR><TD>first table</TD><TD>is here</TD></TR></TABLE><TABLE><TR><TD>second table</TD><TD>comes next</TD></TR></TABLE>"
html.exmpl <- htmlParse(exmpl, encoding = "UTF-8")
htlm.this.table <- getNodeSet(html.exmpl, "//table[1]")
htlm.this.table
htlm.this.table <- getNodeSet(html.exmpl, "//table[2]")
htlm.this.table

# find max rowspan with places of values ---------------------------------------
all_rowspans <- xmlSApply(html_table, 
                          function(x) {xpathSApply(x, "./tr",
                                                   xmlChildren)})
try <- sapply(all_rowspans, 
              function(x) sapply(x, 
                                 function(y) xmlGetAttr(y,
                                                        name = "rowspan")))
try[which(sapply(try, function(x) prod(sapply(x, function(y) is.null(y)))) == 1)] <- NULL


# merge header cells using colspans information --------------------------------
test_header <- list(letters[1:10], 1:8, c(rep("#", 2), rep("$", 3)))
test_colspan <- list(c(NA, NA, 4, NA, NA, NA, 2, NA, 2, NA),
                     c(NA, NA, NA, 2, 3, NA, NA, NA), rep(NA, 5))
lst <- as.list(1:length(test_header))
for (i in 1:length(test_header)) {
    lst[[i]] <- data.frame(header = test_header[[i]], 
                           colspan = test_colspan[[i]],
                           stringsAsFactors = F)
}
lst 
addRepeatsToHeaderList(lst)
mergeTableHeader(addRepeatsToHeaderList(lst))


# test 01: get data from html table --------------------------------------------
# select 2013, 2, 1 
my_url <- "http://www.gks.ru/bgd/regl/B13_14p/IssWWW.exe/Stg/d1/01-01.htm"
# my_url <- getGKSDataRef()
loadGKSData(my_url)
str(glb_dataGKS)
head(glb_dataGKS$data)
tail(glb_dataGKS$data)
str(glb_dataGKS$data)

my_url <- "http://www.gks.ru/bgd/regl/B13_14p/IssWWW.exe/Stg/d1/01-01.htm"
ref <- my_url

# test 01: get data from html table --------------------------------------------
#  
my_url <- "http://www.gks.ru/bgd/regl/B13_14p/IssWWW.exe/Stg/d1/04-02.htm"
# ref <- my_url
glb_dataGKS <- loadGKSData(my_url)
str(glb_dataGKS)
head(glb_dataGKS$data)
tail(glb_dataGKS$data)
str(glb_dataGKS$data)

# test 02: get data from doc table ---------------------------------------------
#
my_url <- "http://www.gks.ru/bgd/regl/B15_14p/IssWWW.exe/Stg/d02/10-01.doc"
my_url <- "http://www.gks.ru/bgd/regl/B17_14p/IssWWW.exe/Stg/d01/04-05.doc"
glb_dataGKS <- loadGKSData(my_url)

