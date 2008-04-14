setwd("./output")

n <- 10
x <- matrix(rnorm(n*n),ncol=n)

rownames(x) <- rep(letters,length=n)
colnames(x) <- rep(letters,length=n)

write.xls(x,"matrixFromR.xls","matrix")
write.xls(x,"matrixFromR.no.colnms.xls","matrix",writeColNms=F)
write.xls(x,"matrix.no.rownms.xls","matrix",writeRowNms=F)
write.xls(x,"matrix.no.nms.xls","matrix",writeColNms=F,writeRowNms=F)

write.xls(x[,1],"vectorFromR.xls","vector")
write.xls(x[,1],"vectorFromR.no.colnms.xls","vector",writeColNms=F)
write.xls(x[,1],"vectorFromR.no.rownms.xls","vector",writeRowNms=F)
write.xls(x[,1],"vectorFromR.no.nms.xls","vector",writeColNms=F,writeRowNms=F)

# char data test
y <- matrix(rep(letters,each=26),ncol=26)
rownames(y) <- rep(letters,length=26)
colnames(y) <- rep(letters,length=26)
write.xls(y,"char.data.xls","alphabet")


# logical data test
z <- matrix(as.logical(round(runif(100),0)),ncol=n)
rownames(z) <- rep(letters,length=n)
colnames(z) <- rep(letters,length=n)

write.xls(z,"logical.data.xls","myLogic")
write.xls(z[,1],"logical.data.vector.xls","myLogicVector")

# test for writing data frames
x.df <- as.data.frame(x)
write.xls(x.df,"data.frame.xls","myDF")

# test list writing

data.list <- list()
data.list[["double"]] <- x
data.list[["char"]] <- y
data.list[["logical"]] <- z
data.list[["dataframe"]] <- x.df

write.xls(data.list,"listFromR.xls")


# test for reading xls data

read.xls <- function(filename, worksheetNames=NA,colnames=T,rownames=T) {
	.Call("xls2R", filename, worksheetNames, colnames, rownames)
}


#from.xls <-read.xls("matrixFromR.xls")

#from.xls

#q()
