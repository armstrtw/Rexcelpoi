
# memory leak test
N <- 10000

nc <- 4
nr <- 1000
x <- matrix(rnorm(nc*nr),ncol=nc,nrow=nr)
rownames(x) <- rep(letters,length=nr)
colnames(x) <- rep(letters,length=nc)


# char data test
y <- matrix(rep(letters,each=26),ncol=26)
rownames(y) <- rep(letters,length=26)
colnames(y) <- rep(letters,length=26)

# logical data test
z <- matrix(as.logical(round(runif(nc*nr),0)),ncol=nc)
rownames(z) <- rep(letters,length=nr)
colnames(z) <- rep(letters,length=nc)

# test for writing data frames
x.df <- as.data.frame(x)

data.list <- list()
data.list[["double"]] <- x
data.list[["char"]] <- y
data.list[["logical"]] <- z
data.list[["dataframe"]] <- x.df


for(i in 1:N) {
cat(i,"\n")

# write a matrix
write.xls(x,"matrixFromR.xls","matrix")
write.xls(x,"matrixFromR.no.colnms.xls","matrix",writeColNms=FALSE)
write.xls(x,"matrix.no.rownms.xls","matrix",writeRowNms=FALSE)
write.xls(x,"matrix.no.nms.xls","matrix",writeColNms=FALSE,writeRowNms=FALSE)

# add some formats
write.xls(x,"matrixFromR_formatted.xls","matrix",formats=rep("0.0",nc))
write.xls(x,"matrixFromR_formatted2.xls","matrix",formats=rep(c("0.0","0"),nc/2))
write.xls(x,"matrixFromR_formatted3.xls","matrix",formats=rep("#,##0.0;[Red]#-##0.0;\"\"",nc))

write.xls(x[,1],"vectorFromR.xls","vector")
write.xls(x[,1],"vectorFromR.no.colnms.xls","vector",writeColNms=FALSE)
write.xls(x[,1],"vectorFromR.no.rownms.xls","vector",writeRowNms=FALSE)
write.xls(x[,1],"vectorFromR.no.nms.xls","vector",writeColNms=FALSE,writeRowNms=FALSE)

write.xls(y,"char.data.xls","alphabet")

write.xls(z,"logical.data.xls","myLogic")
write.xls(z[,1],"logical.data.vector.xls","myLogicVector")

write.xls(x.df,"data.frame.xls","myDF")

write.xls(data.list,"listFromR.xls")
}

